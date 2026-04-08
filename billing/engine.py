from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict
from billing.models import BillingLineItem, TierBreakdown

def calculate_billing(usage_rows, sku_master, exchange_rate, margin_rate=Decimal("1.12")):
    # 1. '쿠팡' 키워드가 들어간 데이터만 집계
    usage_map = defaultdict(int)
    target_proj = {"id": "coupang-all", "name": "Coupang Project"}
    
    for row in usage_rows:
        if row.project_name and "coupang" in row.project_name.lower():
            usage_map[row.sku_id] += row.usage_amount
            target_proj = {"id": row.project_id, "name": row.project_name}

    results = []
    # 2. 💡 [강력 권고] sku_master(14개)를 기준으로 루프를 돕니다.
    # 사용량이 0이어도 sku_master에 들어있으면 무조건 결과에 포함됩니다.
    for sku_id, sku in sku_master.items():
        total_usage = usage_map.get(sku_id, 0)
        free_cap = sku.free_usage_cap
        billable_usage = max(0, total_usage - free_cap)
        
        tier_breakdown, subtotal_usd = _apply_waterfall(billable_usage, sku)
        final_krw = (subtotal_usd * exchange_rate * margin_rate).quantize(Decimal("1"), ROUND_HALF_UP)
        
        results.append(BillingLineItem(
            billing_month="2026-03", project_id=target_proj["id"], project_name=target_proj["name"],
            sku_id=sku_id, sku_name=sku.sku_name, total_usage=total_usage,
            free_cap_applied=min(total_usage, free_cap), billable_usage=billable_usage,
            tier_breakdown=tier_breakdown, subtotal_usd=subtotal_usd,
            exchange_rate=exchange_rate, margin_rate=margin_rate, final_krw=final_krw
        ))
    return results

def _apply_waterfall(billable_usage, sku):
    remaining = billable_usage
    breakdown, subtotal_usd, cum_lower = [], Decimal("0"), 0
    # 모든 티어(5개 구간)를 강제 생성 (사용량 0원 표시용)
    for tier in sorted(sku.tiers, key=lambda t: t.tier_number):
        if tier.tier_limit is None: usage_in_tier = max(0, remaining)
        else:
            capacity = tier.tier_limit - cum_lower
            usage_in_tier = max(0, min(remaining, capacity))
            cum_lower = tier.tier_limit
        
        amt = (Decimal(usage_in_tier) / Decimal("1000")) * tier.tier_cpm
        breakdown.append(TierBreakdown(tier_number=tier.tier_number, usage_in_tier=usage_in_tier, 
                                     tier_cpm=tier.tier_cpm, amount_usd=amt))
        subtotal_usd += amt
        remaining -= usage_in_tier
    return breakdown, subtotal_usd

def summarize_by_project(line_items):
    totals = defaultdict(Decimal)
    for it in line_items: totals[it.project_id] += it.final_krw
    return dict(totals)