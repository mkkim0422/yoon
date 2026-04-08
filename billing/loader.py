import hashlib, io, re
from decimal import Decimal
from typing import Any
from billing.models import Sku, SkuTier, UsageRow

_UNLIMITED_CAP = 999_999_999

# 💡 [검토완료] 사용자님 이미지(image_dc16bf.png) 기반 14개 진짜 ID 고정
GMP_MASTER_LIST = {
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
    "C1B6-FF9D-7700": "Distance Matrix"
}

def load_sku_master(rows: list[dict[str, Any]]) -> dict[str, Sku]:
    sku_map = {}
    for row in rows:
        sid = row["sku_id"]
        if sid not in sku_map:
            sku_map[sid] = Sku(sku_id=sid, sku_name=row["sku_name"], is_billable=row["is_billable"], 
                               category=row["category"], free_usage_cap=row["free_usage_cap"], tiers=[])
        if row.get("tier_number") is not None:
            sku_map[sid].tiers.append(SkuTier(tier_number=int(row["tier_number"]), 
                                             tier_limit=row["tier_limit"], 
                                             tier_cpm=Decimal(str(row["tier_cpm"]))))
    for s in sku_map.values(): s.tiers.sort(key=lambda t: t.tier_number)
    return sku_map

def parse_gmp_price_excel(file_bytes: bytes) -> list[dict]:
    import pandas as pd
    all_sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None, dtype=str)
    result_map = {}

    # 1. 먼저 14개 SKU의 기본 틀(단가 0)을 무조건 만듭니다.
    for sid, name in GMP_MASTER_LIST.items():
        f_cap = 100000 if "Maps" in name or "Geocoding" in name else 0
        limits = [100000, 500000, 1000000, 5000000, None]
        for t_num in range(1, 6):
            result_map[(sid, t_num)] = {
                "sku_id": sid, "sku_name": name, "is_billable": True, "category": "GMP",
                "free_usage_cap": f_cap, "tier_number": t_num, 
                "tier_limit": limits[t_num-1], "tier_cpm": "0"
            }

    # 2. 엑셀을 뒤져서 단가($)가 보이면 업데이트합니다.
    for df in all_sheets.values():
        df = df.fillna("").astype(str)
        for i in range(df.shape[0]):
            row_text = " ".join([str(v) for v in df.iloc[i]]).lower()
            for sid, name in GMP_MASTER_LIST.items():
                if name.lower() in row_text:
                    # 해당 행에서 숫자(단가)처럼 보이는 것들을 다 뽑습니다.
                    prices = [re.sub(r'[$,\s]', '', v) for v in df.iloc[i] if '$' in v or ('.' in v and re.search(r'\d', v))]
                    for idx, p_val in enumerate(prices[:5]):
                        if (sid, idx+1) in result_map:
                            result_map[(sid, idx+1)]["tier_cpm"] = p_val
    
    return list(result_map.values())

def load_usage_rows(rows: list[dict[str, Any]]) -> list[UsageRow]:
    return [UsageRow(billing_month=r["billing_month"], project_id=r["project_id"], 
                     project_name=r.get("project_name", r["project_id"]), 
                     sku_id=r["sku_id"], usage_amount=int(r["usage_amount"])) for r in rows]

def load_exchange_rate(row: dict[str, Any]) -> Decimal:
    return Decimal(str(row["usd_to_krw"]))