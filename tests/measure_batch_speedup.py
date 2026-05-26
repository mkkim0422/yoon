"""Price List 캐싱 적용 후 일괄 정산 속도 측정.

목적: 캐시 효과 정량 확인.
방식:
  1) 캐시 비운 상태에서 12개사 정산 → 소요시간 측정
  2) 같은 12개사를 다시 정산 (모든 호출이 캐시 hit) → 소요시간 측정
  3) 회사 1개당 평균/총합 비교
"""
from __future__ import annotations

import sys
import time
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


class _AttrDict(dict):
    def __getattr__(self, k):
        if k in self: return self[k]
        raise AttributeError(k)
    def __setattr__(self, k, v): self[k] = v


class _StubST:
    def __init__(self): self.__dict__["session_state"] = _AttrDict()
    def __getattr__(self, k): return _StubST()
    def __call__(self, *a, **kw): return _StubST()
    def __enter__(self): return self
    def __exit__(self, *a): return False
    def __contains__(self, k): return False
    def __setitem__(self, k, v): pass
    def __getitem__(self, k): return None
    def __iter__(self): return iter([])


sys.modules["streamlit"] = _StubST()
sys.modules["streamlit_sortables"] = _StubST()

from billing.engine import calculate_billing, calculate_billing_by_project
from billing.loader import (
    build_sku_master_from_usage, get_billable_sku_names, load_usage_rows,
    detect_missing_skus, clear_price_list_cache,
)
from billing.preprocessor import extract_company_names, preprocess_usage_file
from invoice_generator import generate_formatted_invoice


def gen_one(company, billing_month, csv_path, plist, currency, rate):
    raw = preprocess_usage_file(str(csv_path), billing_month, company_filter=company)
    usage = load_usage_rows(raw)
    master = build_sku_master_from_usage(usage, str(plist))
    billable = get_billable_sku_names(str(plist))
    detect_missing_skus(usage, master)
    ex = Decimal(str(rate)); mr = Decimal("1.0")
    items = calculate_billing(usage, master, ex, mr, mode="account")
    proj = calculate_billing_by_project(usage, master, ex, mr, mode="account")
    return generate_formatted_invoice(
        line_items=items, company_name=company, billing_month=billing_month,
        exchange_rate=ex, margin_rate=mr, bank_name="하나은행", proj_results=proj,
        price_list_file=str(plist), sku_order=None, currency=currency,
        billable_skus=billable, billing_mode="account",
        per_project_invoices=None, min_charge_amount=0.0,
        min_charge_currency="KRW", rate_date_str="2026.04.30",
        rate_phrase="최종 송금환율 기준", rate_extra="", include_project_sheet=True,
        subtotal_round=0 if currency == "KRW" else 2, force_keep_skus=None,
    )


def run_pass(label, targets, billing_month, csv, plist_usd, plist_krw):
    print(f"\n=== {label} ===")
    t_start = time.time()
    per_company = []
    for c in targets:
        # 통화 결정
        currency = "USD"; plist = plist_usd
        try:
            raw = preprocess_usage_file(str(csv), billing_month, company_filter=c)
            if not raw:
                continue
        except Exception:
            continue
        # 빈 결과면 KRW 시도
        t0 = time.time()
        try:
            _ = gen_one(c, billing_month, csv, plist, currency, 1480.80)
        except Exception:
            currency = "KRW"; plist = plist_krw
            _ = gen_one(c, billing_month, csv, plist, currency, 1480.80)
        dt = time.time() - t0
        per_company.append((c, dt))
        print(f"  {c}: {dt:.3f}초")
    total = time.time() - t_start
    avg = total / len(per_company) if per_company else 0
    print(f"  --- 합계 {total:.3f}초 / 평균 {avg:.3f}초/사 ({len(per_company)}개사) ---")
    return total, avg, per_company


def main():
    csv = ROOT / "billing.csv"
    plist_usd = ROOT / "billing" / "saved_price_list_usd.xlsx"
    plist_krw = ROOT / "billing" / "saved_price_list_krw.xlsx"
    billing_month = "2026-04"
    companies = extract_company_names(str(csv))
    targets = []
    for c in companies[:30]:
        try:
            raw = preprocess_usage_file(str(csv), billing_month, company_filter=c)
            if raw:
                targets.append(c)
        except Exception:
            pass
        if len(targets) >= 12:
            break

    # PASS A: 캐시 비운 첫 실행 (콜드)
    clear_price_list_cache()
    t1, avg1, _ = run_pass("PASS A: 콜드(캐시 비움)", targets, billing_month, csv, plist_usd, plist_krw)

    # PASS B: 두 번째 실행 (워크북 캐시 hit)
    t2, avg2, _ = run_pass("PASS B: 웜(캐시 hit)", targets, billing_month, csv, plist_usd, plist_krw)

    print("\n=== 요약 ===")
    print(f"PASS A (콜드): {t1:.3f}초, 평균 {avg1:.3f}초/사")
    print(f"PASS B (웜):   {t2:.3f}초, 평균 {avg2:.3f}초/사")
    diff = t1 - t2
    pct = (diff / t1 * 100) if t1 > 0 else 0
    print(f"절감: {diff:.3f}초 ({pct:.1f}%)")


if __name__ == "__main__":
    main()
