"""
gen_invoice_from_raw.py
Google 원본 빌링 CSV를 파싱해 coupang 인보이스를 생성합니다.
"""
from __future__ import annotations

import sys
import io
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import pandas as pd

BASE_DIR = Path(__file__).parent
RAW_CSV  = Path("c:/Users/User/Downloads/{Google} billing raw data.csv")
OUTPUT   = BASE_DIR / "invoice_coupang_final.xlsx"

MARGIN_RATE = Decimal("1.0")   # 실제 Google 청구액 그대로 사용

# ── 고객 청구용 환율 (하나은행 최종 송금환율) — 매월 업데이트 ────────────────────
# Google CSV의 환율(비용 역산용)과 달리, 이 값이 인보이스 금액 계산에 사용됩니다.
BILLING_EXCHANGE_RATE = Decimal("1525.30")

# ── GMP API SKU 화이트리스트 (SKU ID → 표시 이름) ──────────────────────────
GMP_SKU_WHITELIST: dict[str, str] = {
    # GMP 마스터 14개
    "FAF4-3B2D-51B2": "Dynamic Maps",
    "28A8-3EB4-4595": "Directions",
    "BAC8-4E68-E261": "Geocoding",
    "4ED6-464A-2AFC": "Places Details",
    "75D4-C522-326B": "Basic Data",
    "F095-CD01-81B2": "Contact Data",
    "D63D-5CC5-302A": "Atmosphere Data",
    "E95A-86C7-7F47": "Places - Text Search",
    "6B23-8A17-D29D": "Places - Nearby Search",
    "44A2-D839-A3AC": "Find Place",
    "7384-2DE4-D388": "Autocomplete - Per Request",
    "B52C-8320-6DC5": "Autocomplete without Places Details",
    "FC4B-1880-63EF": "Query Autocomplete",
    "C1B6-FF9D-7700": "Distance Matrix",
    # 원본 데이터에 나타나는 실제 SKU ID (신규 API)
    "FC5C-DF28-543F": "Places Details",
    "DC67-3188-2294": "Find Place",
    "C4B1-8805-63EF": "Query Autocomplete",
    "B43B-2A59-C153": "Elevation",
    "3C2D-B525-2E5F": "Static Maps",
    "E9C9-934B-DDB1": "Autocomplete (included with Places Details)",
}


def main() -> None:
    # ── 0. CSV 메타데이터(1~8행) 파싱 → 동적 값 추출 ─────────────────────────
    meta = pd.read_csv(RAW_CSV, encoding="utf-8-sig", header=None, nrows=8)
    # 메타 구조: col0=키, col1=값
    #   Row1: 인보이스 날짜  Row6: 환율
    meta_dict = {str(row[0]).strip(): str(row[1]).strip()
                 for _, row in meta.iterrows() if pd.notna(row[1])}

    INVOICE_DATE  = meta_dict.get("인보이스 날짜", "")
    EXCHANGE_RATE = Decimal(
        meta_dict.get("환율", "1525.30").replace(",", "")
    )

    # ── 1. 원본 CSV 파싱 ──────────────────────────────────────────────────────
    df = pd.read_csv(RAW_CSV, encoding="utf-8-sig", skiprows=8)

    # ── 2. coupang + GMP SKU 필터 (usage + RESELLER_MARGIN 포함, 순비용 계산) ──
    # RESELLER_MARGIN 행은 크레딧 유형='RESELLER_MARGIN'으로 표시되며
    # 음수 비용을 가짐 → 사용 비용과 합산하면 순비용(net cost)이 됨
    mask = (
        df["결제 계정 이름"].astype(str).str.lower().str.contains("coupang", na=False)
        & (df["크레딧 유형"].isna() | (df["크레딧 유형"] == "RESELLER_MARGIN"))
        & df["SKU ID"].isin(GMP_SKU_WHITELIST.keys())
    )
    coupang = df[mask].copy()

    # 결제 계정 이름 동적 추출
    COMPANY = coupang["결제 계정 이름"].dropna().iloc[0] if len(coupang) > 0 else "Unknown"
    # 청구 월: 인보이스 날짜(메타데이터)의 연월을 기준으로 사용 (가장 신뢰도 높음)
    BILLING_MONTH = INVOICE_DATE[:7] if INVOICE_DATE else "2026-03"

    print(f"[1/3] 필터 결과: {len(coupang)}행 (coupang GMP, usage+RESELLER_MARGIN)")
    print(f"      결제 계정: {COMPANY} | 청구 월: {BILLING_MONTH} | 인보이스 날짜: {INVOICE_DATE}")
    print(f"      Google 고지서 환율: {EXCHANGE_RATE} | 고객 청구 환율(송금환율): {BILLING_EXCHANGE_RATE}")

    # ── 3. 사용량·비용 수치 변환 ──────────────────────────────────────────────
    coupang["_usage"] = (
        pd.to_numeric(
            coupang["사용량"].astype(str).str.replace(",", ""), errors="coerce"
        ).fillna(0)
    )
    # 반올림되지 않은 비용(₩) 우선 사용, 없으면 비용(₩) 사용
    _krw_col = next(
        (c for c in ["반올림되지 않은 비용(₩)", "비용(₩)"] if c in coupang.columns),
        "비용(₩)",
    )
    coupang["_krw"] = (
        pd.to_numeric(
            coupang[_krw_col].astype(str).str.replace(",", ""), errors="coerce"
        ).fillna(0)
    )

    # 원본 USD 비용 컬럼 탐지 (Google CSV 컬럼명 후보)
    _usd_col = next(
        (c for c in ["비용($)", "비용(USD)", "Cost ($)", "Cost (USD)"] if c in coupang.columns),
        None,
    )
    if _usd_col:
        coupang["_cost_usd"] = pd.to_numeric(
            coupang[_usd_col].astype(str).str.replace(",", "").str.replace("$", ""),
            errors="coerce",
        ).fillna(0)
    else:
        # Google CSV의 환율(Google이 KRW 산출에 사용한 환율)로 USD 역산
        # ※ BILLING_EXCHANGE_RATE(고객 청구용)와 다를 수 있으므로 Google 환율 사용
        coupang["_cost_usd"] = coupang["_krw"] / float(EXCHANGE_RATE)

    # 공시 단가 컬럼 탐지
    _price_col = next(
        (c for c in ["단가", "단가($)", "Price per usage unit", "Price"] if c in coupang.columns),
        None,
    )
    if _price_col:
        coupang["_unit_price"] = pd.to_numeric(
            coupang[_price_col].astype(str).str.replace(",", "").str.replace("$", ""),
            errors="coerce",
        ).fillna(0)
    else:
        coupang["_unit_price"] = 0

    # SKU ID가 여러 개 있어도 표시 이름으로 통합 (e.g. Places Details)
    coupang["_display_name"] = coupang["SKU ID"].map(GMP_SKU_WHITELIST)

    # ── 3a. SKU별 전체 집계 (Invoice 시트용) ──────────────────────────────────
    agg = (
        coupang.groupby("_display_name")
        .agg(total_usage=("_usage", "sum"), total_krw=("_krw", "sum"))
        .reset_index()
        .rename(columns={"_display_name": "sku_name"})
    )

    print(f"\n{'SKU 이름':45} {'Usage':>15} {'KRW':>15} {'USD':>12}")
    print("-" * 92)
    for _, row in agg.iterrows():
        usd = row["total_krw"] / float(BILLING_EXCHANGE_RATE)
        print(
            f'{row["sku_name"]:45} {int(row["total_usage"]):>15,} '
            f'{int(row["total_krw"]):>15,} {usd:>12,.2f}'
        )
    total_krw = agg["total_krw"].sum()
    total_usd = total_krw / float(BILLING_EXCHANGE_RATE)
    print("-" * 92)
    print(f'{"합계":45} {"":>15} {int(total_krw):>15,} {total_usd:>12,.2f}')

    # ── 3b. 프로젝트 × SKU별 집계 (Project 시트용) ───────────────────────────
    # 프로젝트 컬럼명 탐지 (한/영 대응)
    _proj_name_col = next(
        (c for c in ["프로젝트 이름", "Project name", "Project Name"] if c in coupang.columns),
        None,
    )
    _proj_id_col = next(
        (c for c in ["프로젝트 ID", "Project ID"] if c in coupang.columns),
        None,
    )

    proj_results = None
    if _proj_name_col and _proj_id_col:
        _pagg = (
            coupang.groupby([_proj_id_col, _proj_name_col, "_display_name"])
            .agg(
                p_usage=("_usage", "sum"),
                p_krw=("_krw", "sum"),
                p_cost_usd=("_cost_usd", "sum"),
                p_unit_price=("_unit_price", "first"),  # 단가: 합산 금지, 첫 값 사용
            )
            .reset_index()
        )

        proj_results = []
        for proj_id in sorted(_pagg[_proj_id_col].unique()):
            _rows = _pagg[_pagg[_proj_id_col] == proj_id]
            proj_name = _rows[_proj_name_col].iloc[0]

            skus: dict = {}
            p_total_usd = Decimal("0")
            p_total_krw = Decimal("0")
            for _, r in _rows.iterrows():
                _usage      = int(r["p_usage"])
                # ── 단일 소스 원칙: p_krw(순비용 KRW) → BILLING_EXCHANGE_RATE로 USD 환산
                # p_cost_usd는 RESELLER_MARGIN 행의 USD가 누락된 경우 과다 계상됨 → 사용 금지
                _p_krw_dec  = Decimal(str(round(float(r["p_krw"]), 2)))
                _usd        = (_p_krw_dec / BILLING_EXCHANGE_RATE).quantize(Decimal("0.01"), ROUND_HALF_UP)
                _krw        = (_usd * BILLING_EXCHANGE_RATE).quantize(Decimal("1"), ROUND_HALF_UP)
                _unit_price = float(r["p_unit_price"]) if float(r["p_unit_price"]) > 0 else None
                skus[r["_display_name"]] = {
                    "usage":        _usage,
                    "subtotal_usd": _usd,
                    "final_krw":    _krw,
                    "unit_price":   _unit_price,
                }
                p_total_usd += _usd
                p_total_krw += _krw

            proj_results.append({
                "proj_id":   proj_id,
                "proj_name": proj_name,
                "skus":      skus,
                "total_usd": p_total_usd,
                "total_krw": p_total_krw,
            })
        print(f"\n[2/3] 프로젝트 수: {len(proj_results)}개")
        for pr in proj_results:
            print(f"      {pr['proj_name']:35s}  KRW={int(pr['total_krw']):>15,}")
    else:
        print("\n[2/3] '프로젝트 이름'/'프로젝트 ID' 컬럼 없음 → Project 시트 생략")

    # ── 4. BillingLineItem 객체 생성 ──────────────────────────────────────────
    from billing.models import BillingLineItem, TierBreakdown

    line_items: list[BillingLineItem] = []
    for _, row in agg.iterrows():
        usage    = int(row["total_usage"])
        krw      = Decimal(str(round(float(row["total_krw"]), 4)))
        usd      = (krw / BILLING_EXCHANGE_RATE).quantize(Decimal("0.01"), ROUND_HALF_UP)

        # 단가(CPM) 계산: usage가 0이면 0으로 처리
        cpm = (usd / usage * 1000).quantize(Decimal("0.0001"), ROUND_HALF_UP) if usage > 0 else Decimal("0")

        tb = TierBreakdown(
            tier_number=1,
            usage_in_tier=usage,
            tier_cpm=cpm,
            amount_usd=usd,
        )

        li = BillingLineItem(
            billing_month   = BILLING_MONTH,
            project_id      = "coupang",
            project_name    = "coupang",
            sku_id          = row["sku_name"],   # 표시 이름을 ID 대신 사용
            sku_name        = row["sku_name"],
            total_usage     = usage,
            free_usage_cap  = 0,
            free_cap_applied= 0,
            billable_usage  = usage,
            tier_breakdown  = [tb],
            subtotal_usd    = usd,
            exchange_rate   = BILLING_EXCHANGE_RATE,
            margin_rate     = MARGIN_RATE,
            final_krw       = krw,
        )
        line_items.append(li)

    # ── 5. 인보이스 생성 (Invoice 시트 + Project 시트) ───────────────────────
    print(f"\n[3/3] 인보이스 생성: {OUTPUT.name}")
    from invoice_generator import generate_formatted_invoice

    generate_formatted_invoice(
        line_items    = line_items,
        company_name  = COMPANY,
        billing_month = BILLING_MONTH,
        exchange_rate = BILLING_EXCHANGE_RATE,
        margin_rate   = MARGIN_RATE,
        invoice_date  = INVOICE_DATE,
        output_path   = OUTPUT,
        proj_results  = proj_results,   # ← Project 시트 생성
    )

    print(f"\n완료 -> {OUTPUT}")
    print(f"  항목 수: {len(line_items)}건")
    for li in sorted(line_items, key=lambda x: x.sku_name):
        print(f"  - {li.sku_name}: ${float(li.subtotal_usd):,.2f}  (₩{int(li.final_krw):,})")
    print(f"\n  합계 KRW: ₩{int(total_krw):,}")


if __name__ == "__main__":
    main()
