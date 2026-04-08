"""
main.py  -  구글 빌링 정산 자동화 시스템 진입점 (Entry Point)

파이프라인:
  preprocess_usage_file          (전처리: CSV/Excel → 정제된 dict 리스트)
      ↓
  load_sku_master / load_usage_rows  (모델 로드: dict → dataclass)
      ↓
  calculate_billing              (과금 엔진: Waterfall 계산)
      ↓
  Excel 저장                     (pandas to_excel, Raw Data only)

실행 예시:
  python main.py billing.csv --billing-month 2026-03 --exchange-rate 1427.87
  python main.py billing.csv -m 2026-03 -e 1427.87 -r 1.12
"""
from __future__ import annotations

import argparse
from decimal import Decimal
from pathlib import Path
from typing import Any

import pandas as pd

from billing.engine import calculate_billing
from billing.loader import load_sku_master, load_usage_rows
from billing.models import BillingLineItem
from billing.preprocessor import preprocess_usage_file


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def generate_invoice_excel(
    input_file_path:  str | Path,
    output_file_path: str | Path,
    exchange_rate:    float | Decimal,
    sku_master_rows:  list[dict[str, Any]],
    billing_month:    str,
    margin_rate:      float | Decimal = 1.12,
    company_filter:   str | None = None,
) -> list[BillingLineItem]:
    """
    사용고지서 파일을 읽어 정산 계산 후 Excel 파일로 저장한다.

    Args:
        input_file_path:  원본 사용고지서 경로 (.csv / .xlsx / .xls)
        output_file_path: 결과 Excel 저장 경로 (.xlsx)
        exchange_rate:    해당 월 USD → KRW 환율
        sku_master_rows:  SKU 마스터 데이터 (DB 또는 Mock)
                          load_sku_master()가 요구하는 dict 리스트 형식
        billing_month:    정산 월 'YYYY-MM'
        margin_rate:      마진율 (기본 1.12 = 12%)

    Returns:
        계산 완료된 BillingLineItem 리스트
    """
    exchange_rate = Decimal(str(exchange_rate))
    margin_rate   = Decimal(str(margin_rate))

    # ── Step 1: 전처리 ────────────────────────────────────────────────────
    raw_rows = preprocess_usage_file(
        input_file_path, billing_month, company_filter=company_filter
    )

    # ── Step 2: 모델 로드 ─────────────────────────────────────────────────
    sku_master = load_sku_master(sku_master_rows)
    usage_rows = load_usage_rows(raw_rows)

    # ── Step 3: 과금 계산 ─────────────────────────────────────────────────
    line_items = calculate_billing(usage_rows, sku_master, exchange_rate, margin_rate)

    # ── Step 4: Excel 저장 ────────────────────────────────────────────────
    _export_excel(line_items, Path(output_file_path))

    return line_items


# ─────────────────────────────────────────────────────────────────────────────
# 내부: Excel 저장 (Raw Data, 서식 없음)
# ─────────────────────────────────────────────────────────────────────────────

def _export_excel(items: list[BillingLineItem], output_path: Path) -> None:
    """BillingLineItem 리스트를 pandas로 xlsx에 저장 (Raw Data only)."""

    # ── Invoice Summary 시트 ──────────────────────────────────────────────
    summary_rows = []
    for item in items:
        summary_rows.append(
            {
                "정산월":       item.billing_month,
                "프로젝트ID":  item.project_id,
                "SKU_ID":      item.sku_id,
                "SKU명":       item.sku_name,
                "총사용량":    item.total_usage,
                "무료차감":    item.free_cap_applied,
                "청구대상":    item.billable_usage,
                "소계_USD":    float(item.subtotal_usd),
                "환율":        float(item.exchange_rate),
                "마진율":      float(item.margin_rate),
                "최종_KRW":   float(item.final_krw),
            }
        )
    df_summary = pd.DataFrame(summary_rows)

    # ── Tier Breakdown 시트 ───────────────────────────────────────────────
    breakdown_rows = []
    for item in items:
        for tb in item.tier_breakdown:
            breakdown_rows.append(
                {
                    "프로젝트ID":  item.project_id,
                    "SKU_ID":      item.sku_id,
                    "구간번호":    tb.tier_number,
                    "구간사용량":  tb.usage_in_tier,
                    "구간단가_CPM": float(tb.tier_cpm),
                    "구간금액_USD": float(tb.amount_usd),
                }
            )
    df_breakdown = pd.DataFrame(breakdown_rows)

    # ── Excel 저장 ────────────────────────────────────────────────────────
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_summary.to_excel(writer,   sheet_name="Invoice Summary", index=False)
        df_breakdown.to_excel(writer, sheet_name="Tier Breakdown",  index=False)


# ─────────────────────────────────────────────────────────────────────────────
# SKU 마스터 (실제 단가표 기준, Decimal 사용)
# tier_limit: 누적 상한 건수 (마지막 구간은 None → 무제한)
# ─────────────────────────────────────────────────────────────────────────────

def _make_sku_rows(
    sku_id: str,
    sku_name: str,
    category: str,
    free_usage_cap: int,
    tiers: list[tuple[int | None, str]],   # [(tier_limit, cpm_str), ...]
) -> list[dict[str, Any]]:
    """SKU 1개의 모든 구간 행을 생성하는 헬퍼."""
    base = {
        "sku_id": sku_id, "sku_name": sku_name,
        "is_billable": True, "category": category,
        "free_usage_cap": free_usage_cap,
    }
    return [
        {**base, "tier_number": i + 1, "tier_limit": lim, "tier_cpm": Decimal(cpm)}
        for i, (lim, cpm) in enumerate(tiers)
    ]


SKU_MASTER_ROWS: list[dict[str, Any]] = [
    # ── Dynamic Maps (Google SKU ID: FAF4-3B2D-51B2) ──────────────────────
    *_make_sku_rows(
        sku_id="FAF4-3B2D-51B2", sku_name="Dynamic Maps",
        category="Maps", free_usage_cap=10_000,
        tiers=[
            (100_000,   "7.00"),
            (500_000,   "5.60"),
            (1_000_000, "4.20"),
            (5_000_000, "2.10"),
            (None,      "0.53"),
        ],
    ),
    # ── Geocoding (Google SKU ID: BAC8-4E68-E261) ─────────────────────────
    *_make_sku_rows(
        sku_id="BAC8-4E68-E261", sku_name="Geocoding",
        category="Maps", free_usage_cap=10_000,
        tiers=[
            (100_000,   "5.00"),
            (500_000,   "4.00"),
            (1_000_000, "3.00"),
            (5_000_000, "1.50"),
            (None,      "0.38"),
        ],
    ),
    # ── Places Details (Google SKU ID: FC5C-DF28-543F) ────────────────────
    *_make_sku_rows(
        sku_id="FC5C-DF28-543F", sku_name="Places Details",
        category="Places", free_usage_cap=5_000,
        tiers=[
            (100_000,   "17.00"),
            (500_000,   "13.60"),
            (1_000_000, "10.20"),
            (5_000_000,  "5.10"),
            (None,       "1.28"),
        ],
    ),
    # ── Autocomplete - Per Request (Google SKU ID: 7384-2DE4-D388) ────────
    *_make_sku_rows(
        sku_id="7384-2DE4-D388", sku_name="Autocomplete - Per Request",
        category="Places", free_usage_cap=10_000,
        tiers=[
            (100_000,   "2.83"),
            (500_000,   "2.27"),
            (1_000_000, "1.70"),
            (5_000_000, "0.85"),
            (None,      "0.21"),
        ],
    ),
    # ── 세금 (비과금) ─────────────────────────────────────────────────────
    {
        "sku_id": "tax.vat", "sku_name": "세금",
        "is_billable": False, "category": "Tax", "free_usage_cap": 0,
        "tier_number": None, "tier_limit": None, "tier_cpm": None,
    },
]


# ─────────────────────────────────────────────────────────────────────────────
# CLI 진입점
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="구글 Maps 사용고지서 → 정산 Excel 자동 생성",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "input_file",
        help="원본 사용고지서 파일 경로 (.csv / .xlsx)\n예: billing.csv",
    )
    parser.add_argument(
        "-m", "--billing-month",
        required=True,
        metavar="YYYY-MM",
        help="정산 월  (예: 2026-03)",
    )
    parser.add_argument(
        "-e", "--exchange-rate",
        required=True,
        type=float,
        metavar="RATE",
        help="USD → KRW 환율  (예: 1427.87)",
    )
    parser.add_argument(
        "-r", "--margin-rate",
        type=float,
        default=1.12,
        metavar="RATE",
        help="마진율  (기본값: 1.12)",
    )
    parser.add_argument(
        "-o", "--output",
        default="real_invoice_output.xlsx",
        metavar="FILE",
        help="결과 Excel 파일명  (기본값: real_invoice_output.xlsx)",
    )

    args = parser.parse_args()

    INPUT_FILE  = Path(args.input_file)
    OUTPUT_FILE = Path(__file__).parent / args.output

    results = generate_invoice_excel(
        input_file_path  = INPUT_FILE,
        output_file_path = OUTPUT_FILE,
        exchange_rate    = args.exchange_rate,
        sku_master_rows  = SKU_MASTER_ROWS,
        billing_month    = args.billing_month,
        margin_rate      = args.margin_rate,
    )

    print(f"\n{'='*55}")
    print(f"  정산 완료 - {len(results)}건")
    print(f"{'='*55}")
    for item in results:
        print(f"\n[{item.project_id}] {item.sku_name}")
        print(f"  총사용량  : {item.total_usage:>10,} 건")
        print(f"  무료차감  : {item.free_cap_applied:>10,} 건")
        print(f"  청구대상  : {item.billable_usage:>10,} 건")
        for tb in item.tier_breakdown:
            print(f"  T{tb.tier_number}: {tb.usage_in_tier:>8,}건 x ${tb.tier_cpm}/1000"
                  f" = ${tb.amount_usd:,.2f}")
        print(f"  소계(USD) : ${item.subtotal_usd:>10,.4f}")
        print(f"  최종(KRW) :  {item.final_krw:>10,} 원")

    print(f"\n  Excel 저장 완료 -> {OUTPUT_FILE.name}")
    print(f"{'='*55}\n")
