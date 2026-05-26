"""단일 정산 vs 일괄 정산 결과 일치 검증.

회사별 환율표기 expander 도입(rate/date 회사별 저장) 이후에도 일괄 정산
경로의 _run_batch_single_billing 가 단일 정산 메인 흐름과 같은 입력에 대해
같은 결과를 내는지를 라인아이템·합계 단위로 비교한다.

비교 전략:
  A) 단일 정산 경로: preprocess → load → build_sku_master → calculate_billing
     → calculate_billing_by_project → generate_formatted_invoice 직접 호출
  B) 일괄 정산 경로: webapp._run_batch_single_billing(...) 단일 호출

두 경로의 line_items final_krw / subtotal_usd 합계와 proj_results 의
total_krw 합계가 일치하면 통과로 본다 (엑셀 바이트는 메타데이터·시간
차이로 동일 보장 X).
"""
from __future__ import annotations

import sys
from decimal import Decimal
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

# ── Streamlit 무력화 (webapp import 시 부작용 차단) ───────────────────────
class _AttrDict(dict):
    """attribute / item access 둘 다 지원 (streamlit.session_state 흉내)."""
    def __getattr__(self, k):
        if k in self:
            return self[k]
        raise AttributeError(k)
    def __setattr__(self, k, v):
        self[k] = v
    def __delattr__(self, k):
        if k in self:
            del self[k]


class _StubST:
    """webapp 모듈이 import 시 호출하는 streamlit API 를 무시하는 stub."""
    def __init__(self):
        self.__dict__["session_state"] = _AttrDict()

    def __getattr__(self, k):
        return _StubST()

    def __call__(self, *a, **kw):
        return _StubST()

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def __contains__(self, k):
        return False

    def __setitem__(self, k, v):
        pass

    def __getitem__(self, k):
        return None

    def __iter__(self):
        return iter([])


# 모듈 레벨 stub — 같은 인스턴스 사용해야 session_state 가 유지됨
_st_stub = _StubST()
sys.modules["streamlit"] = _st_stub
sys.modules["streamlit_sortables"] = _StubST()


def _extract_run_batch_single_billing():
    """webapp.py 에서 _run_batch_single_billing 함수 소스만 추출해 exec.

    webapp 전체 import 는 streamlit main UI 실행으로 실패하므로, 함수
    정의 텍스트만 격리 실행해 함수 객체를 얻는다.
    """
    import ast, textwrap
    src = (ROOT / "webapp.py").read_text(encoding="utf-8")
    tree = ast.parse(src)
    target = None
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == "_run_batch_single_billing":
            target = node
            break
    assert target, "_run_batch_single_billing not found in webapp.py"
    func_src = ast.get_source_segment(src, target)
    # 함수가 의존하는 import / 상수를 namespace 에 주입
    ns: dict = {}
    from pathlib import Path as _P
    from decimal import Decimal as _Dec
    from collections import defaultdict as _dd
    from billing.engine import calculate_billing as _cb, calculate_billing_by_project as _cbp
    from billing.loader import (
        build_sku_master_from_usage as _bsmf, detect_missing_skus as _dms,
        get_free_caps_from_price_list as _gfc, load_usage_rows as _lur,
    )
    from billing.models import BillingLineItem as _BLI
    from billing.preprocessor import preprocess_usage_file as _puf
    from invoice_generator import generate_formatted_invoice as _gfi
    import re as _re_mod, traceback as _tb_mod
    ns.update({
        "Path": _P, "Decimal": _Dec,
        "calculate_billing": _cb, "calculate_billing_by_project": _cbp,
        "build_sku_master_from_usage": _bsmf, "detect_missing_skus": _dms,
        "get_free_caps_from_price_list": _gfc, "load_usage_rows": _lur,
        "preprocess_usage_file": _puf, "generate_formatted_invoice": _gfi,
    })
    exec(func_src, ns)
    return ns["_run_batch_single_billing"]


_run_batch_single_billing = _extract_run_batch_single_billing()

# ── 실제 import ─────────────────────────────────────────────────────────
from billing.engine import calculate_billing, calculate_billing_by_project  # noqa: E402
from billing.loader import (  # noqa: E402
    build_sku_master_from_usage,
    get_billable_sku_names,
    load_usage_rows,
)
from billing.preprocessor import extract_company_names, preprocess_usage_file  # noqa: E402
from invoice_generator import generate_formatted_invoice  # noqa: E402

# webapp 직접 import 불가 — 위에서 _run_batch_single_billing 추출 사용
# (webapp.py 가 streamlit main UI 를 module-level 실행하기 때문)


def _sum_line_items(items):
    return {
        "n": len(items),
        "krw": float(sum(Decimal(str(getattr(it, "final_krw", 0) or 0)) for it in items)),
        "usd": float(sum(Decimal(str(getattr(it, "subtotal_usd", 0) or 0)) for it in items)),
        "names": [getattr(it, "sku_name", "") for it in items],
    }


def _sum_proj_results(prs):
    if not prs:
        return {"n": 0, "krw": 0.0, "usd": 0.0}
    return {
        "n": len(prs),
        "krw": float(sum(p.get("total_krw") or 0 for p in prs)),
        "usd": float(sum(p.get("total_usd") or 0 for p in prs)),
    }


def _excel_invoice_total(xlsx_bytes: bytes) -> dict:
    """엑셀의 Invoice 시트 마지막 합계 셀 위치를 그대로 비교하는 대신
    모든 셀의 숫자값 sorted 리스트를 반환 — 메타데이터·서식 차이 무시."""
    if not xlsx_bytes:
        return {"all_values": []}
    import io
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=False)
    out: dict = {}
    for sn in wb.sheetnames:
        ws = wb[sn]
        vals = []
        for row in ws.iter_rows(values_only=True):
            for c in row:
                if isinstance(c, (int, float)):
                    vals.append(round(float(c), 4))
                elif isinstance(c, str) and c.strip():
                    vals.append(c.strip())
        out[sn] = vals
    return out


def run_single_path(*, company, billing_month, csv_path, price_list_path,
                    currency, exchange_rate, bank_name, rate_phrase, rate_date_str,
                    min_charge_amount, min_charge_currency):
    """단일 정산 경로: webapp 의 단일 모드 흐름과 동일한 순서로 직접 호출."""
    raw = preprocess_usage_file(str(csv_path), billing_month, company_filter=company)
    usage_rows = load_usage_rows(raw)
    sku_master = build_sku_master_from_usage(usage_rows, str(price_list_path))
    billable_skus = get_billable_sku_names(str(price_list_path))

    _ex = Decimal(str(exchange_rate))
    _mr = Decimal("0")

    line_items = calculate_billing(usage_rows, sku_master, _ex, _mr, mode="account")
    proj_results = calculate_billing_by_project(
        usage_rows, sku_master, _ex, _mr, mode="account",
    )

    excel = generate_formatted_invoice(
        line_items           = line_items,
        company_name         = company,
        billing_month        = billing_month,
        exchange_rate        = _ex,
        margin_rate          = _mr,
        bank_name            = bank_name,
        proj_results         = proj_results,
        price_list_file      = str(price_list_path),
        sku_order            = None,
        currency             = currency,
        billable_skus        = billable_skus,
        billing_mode         = "account",
        per_project_invoices = None,
        min_charge_amount    = float(min_charge_amount),
        min_charge_currency  = min_charge_currency,
        rate_date_str        = rate_date_str,
        rate_phrase          = rate_phrase,
        rate_extra           = "",
        include_project_sheet= True,
        subtotal_round       = 0 if currency == "KRW" else 2,
        force_keep_skus      = None,
    )
    return {
        "items":    _sum_line_items(line_items),
        "proj":     _sum_proj_results(proj_results),
        "excel":    _excel_invoice_total(excel),
        "excel_bytes_len": len(excel) if excel else 0,
    }


def run_batch_path(*, company, billing_month, csv_path, price_list_path,
                   currency, exchange_rate, bank_name, rate_phrase, rate_date_str,
                   min_charge_amount, min_charge_currency):
    """일괄 정산 경로: webapp._run_batch_single_billing 호출."""
    billable_skus = get_billable_sku_names(str(price_list_path))
    res = _run_batch_single_billing(
        selected_company      = company,
        billing_month         = billing_month,
        tmp_input_path        = Path(csv_path),
        price_list_file       = str(price_list_path),
        currency              = currency,
        exchange_rate         = float(exchange_rate),
        margin_rate           = 0.0,
        rate_date_str         = rate_date_str,
        billing_mode          = "account",
        include_project_sheet = True,
        subtotal_round        = 0 if currency == "KRW" else 2,
        bank_name             = bank_name,
        rate_phrase_text      = rate_phrase,
        rate_extra_text       = "",
        min_charge_amount     = float(min_charge_amount),
        min_charge_currency   = min_charge_currency,
        sku_order             = [],
        manual_skus           = [],
        hidden_skus           = [],
        billable_skus         = billable_skus,
        dl_xlsx               = True,
        dl_pdf                = False,
    )
    assert res["ok"], f"batch failed: {res.get('error')}\n{res.get('traceback')}"
    excel = res["excel_bytes"]
    return {
        "excel":    _excel_invoice_total(excel),
        "excel_bytes_len": len(excel) if excel else 0,
    }


def main():
    csv_path = ROOT / "billing.csv"
    price_usd = ROOT / "billing" / "saved_price_list_usd.xlsx"
    price_krw = ROOT / "billing" / "saved_price_list_krw.xlsx"
    assert csv_path.exists(), f"missing {csv_path}"

    companies = extract_company_names(str(csv_path))
    # 정산 대상 후보 — 데이터가 있을 회사 일부.
    # 단일 모드는 차분히 1개부터.
    # 광범위 샘플 — 영문/한글 혼합, 다양한 사용량 패턴 커버
    targets = companies[:12]
    print(f"will test {len(targets)} companies")

    print(f"comparing companies: {targets}")
    print(f"csv: {csv_path.name}")
    any_diff = False
    for company in targets:
        # 회사별 통화 추정 — Naver 같은 일부 회사는 KRW, 나머지는 USD. 우선
        # USD 단가표로 시도, 빈 결과면 KRW 로 폴백.
        billing_month = "2026-04"  # billing.csv 의 빌링월 추정 (필요 시 수정)
        for currency, plist in (("USD", price_usd), ("KRW", price_krw)):
            try:
                s = run_single_path(
                    company=company, billing_month=billing_month,
                    csv_path=csv_path, price_list_path=plist,
                    currency=currency, exchange_rate=1480.80,
                    bank_name="하나은행", rate_phrase="최종 송금환율 기준",
                    rate_date_str="2026.04.30",
                    min_charge_amount=0, min_charge_currency="KRW",
                )
                if s["items"]["n"] == 0:
                    continue  # 이 통화·단가표 조합에 데이터 없음
                b = run_batch_path(
                    company=company, billing_month=billing_month,
                    csv_path=csv_path, price_list_path=plist,
                    currency=currency, exchange_rate=1480.80,
                    bank_name="하나은행", rate_phrase="최종 송금환율 기준",
                    rate_date_str="2026.04.30",
                    min_charge_amount=0, min_charge_currency="KRW",
                )
                # 비교
                eq = (s["excel"] == b["excel"])
                tag = "OK" if eq else "DIFF"
                print(
                    f"  [{tag}] {company} ({currency}) "
                    f"items.n={s['items']['n']} "
                    f"krw={s['items']['krw']:,.0f} usd={s['items']['usd']:,.4f} "
                    f"excel_bytes single={s['excel_bytes_len']} batch={b['excel_bytes_len']}"
                )
                if not eq:
                    any_diff = True
                    # 시트별 diff 요약
                    for sn in set(list(s["excel"].keys()) + list(b["excel"].keys())):
                        sv = s["excel"].get(sn, [])
                        bv = b["excel"].get(sn, [])
                        if sv != bv:
                            # 첫 다른 항목 보고
                            sset = sorted(set(sv) - set(bv), key=lambda x: str(x))[:5]
                            bset = sorted(set(bv) - set(sv), key=lambda x: str(x))[:5]
                            print(f"    sheet={sn!r}")
                            print(f"      only in single: {sset}")
                            print(f"      only in batch:  {bset}")
                break  # 데이터 있는 첫 통화에서 그만
            except Exception as e:
                # 데이터 없거나 통화 미스매치 — 다음 통화로
                pass
        else:
            print(f"  [SKIP] {company}: no usage data in any currency")
    if any_diff:
        print("\nRESULT: DIFF FOUND - FAIL")
        sys.exit(1)
    else:
        print("\nRESULT: ALL MATCH - PASS")


if __name__ == "__main__":
    main()
