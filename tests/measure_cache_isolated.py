"""Price List 캐시 효과를 분리해 측정.

방법: 각 회사 정산 시작 직전에 캐시를 비우는 시나리오(NO CACHE) vs
회사 첫 번째 정산 후 캐시가 채워진 상태로 계속 정산(WITH CACHE).
generate_formatted_invoice 내 _copy_price_list_sheet 와
build_sku_master_from_usage 내 get_sku_tiers 의 워크북 파싱이
얼마나 절감되는지 직접 비교.
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


def gen_one(company, billing_month, csv_path, plist):
    raw = preprocess_usage_file(str(csv_path), billing_month, company_filter=company)
    if not raw:
        return None
    usage = load_usage_rows(raw)
    master = build_sku_master_from_usage(usage, str(plist))
    billable = get_billable_sku_names(str(plist))
    detect_missing_skus(usage, master)
    ex = Decimal("1480.80"); mr = Decimal("1.0")
    items = calculate_billing(usage, master, ex, mr, mode="account")
    proj = calculate_billing_by_project(usage, master, ex, mr, mode="account")
    return generate_formatted_invoice(
        line_items=items, company_name=company, billing_month=billing_month,
        exchange_rate=ex, margin_rate=mr, bank_name="하나은행", proj_results=proj,
        price_list_file=str(plist), sku_order=None, currency="USD",
        billable_skus=billable, billing_mode="account",
        per_project_invoices=None, min_charge_amount=0.0,
        min_charge_currency="KRW", rate_date_str="2026.04.30",
        rate_phrase="최종 송금환율 기준", rate_extra="", include_project_sheet=True,
        subtotal_round=2, force_keep_skus=None,
    )


def main():
    csv = ROOT / "billing.csv"
    plist = ROOT / "billing" / "saved_price_list_usd.xlsx"
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
        if len(targets) >= 8:
            break

    # WARM 모듈/openpyxl — 1 회 정산 후 측정 시작 (콜드 스타트 노이즈 제거)
    clear_price_list_cache()
    _ = gen_one(targets[0], billing_month, csv, plist)

    # PASS A: 매 회사 직전에 캐시를 비워 캐시 효과 OFF
    print("=== PASS A: 캐시 OFF (매 회사 직전 clear) ===")
    t1 = time.time()
    for c in targets:
        clear_price_list_cache()
        t0 = time.time()
        _ = gen_one(c, billing_month, csv, plist)
        print(f"  {c}: {time.time()-t0:.3f}초")
    a_total = time.time() - t1
    a_avg = a_total / len(targets)
    print(f"  --- 합계 {a_total:.3f}초 / 평균 {a_avg:.3f}초/사 ---")

    # PASS B: 시작 시 한 번만 비우고 이후 캐시 hit
    print("\n=== PASS B: 캐시 ON (시작 시 1회 clear) ===")
    clear_price_list_cache()
    t1 = time.time()
    for c in targets:
        t0 = time.time()
        _ = gen_one(c, billing_month, csv, plist)
        print(f"  {c}: {time.time()-t0:.3f}초")
    b_total = time.time() - t1
    b_avg = b_total / len(targets)
    print(f"  --- 합계 {b_total:.3f}초 / 평균 {b_avg:.3f}초/사 ---")

    print("\n=== 요약 ===")
    print(f"PASS A (캐시 OFF): {a_total:.3f}초, 평균 {a_avg:.3f}초/사")
    print(f"PASS B (캐시 ON):  {b_total:.3f}초, 평균 {b_avg:.3f}초/사")
    diff = a_total - b_total
    pct = (diff / a_total * 100) if a_total > 0 else 0
    pct_per = (a_avg - b_avg) / a_avg * 100 if a_avg > 0 else 0
    print(f"절감: 총 {diff:.3f}초 ({pct:.1f}%), 회사당 {a_avg-b_avg:.3f}초 ({pct_per:.1f}%)")


if __name__ == "__main__":
    main()
