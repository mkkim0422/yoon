from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict
from billing.models import BillingLineItem, TierBreakdown


def calculate_billing(usage_rows, sku_master, exchange_rate, margin_rate=Decimal("1.12")):
    """전체 합산 청구 계산 → Sheet 1 Invoice용"""
    usage_map = defaultdict(int)
    for row in usage_rows:
        usage_map[row.sku_id] += row.usage_amount

    results = []
    for sku_id, sku in sku_master.items():
        total_usage = usage_map.get(sku_id, 0)
        if total_usage == 0:
            continue  # 미사용 SKU 제외
        free_cap = sku.free_usage_cap
        free_cap_applied = min(total_usage, free_cap)
        billable_usage = max(0, total_usage - free_cap)
        tier_breakdown, subtotal_usd = _apply_waterfall(billable_usage, sku)
        final_krw = (subtotal_usd * exchange_rate * margin_rate).quantize(Decimal("1"), ROUND_HALF_UP)
        results.append(BillingLineItem(
            billing_month="", project_id="", project_name="",
            sku_id=sku_id, sku_name=sku.sku_name, total_usage=total_usage,
            free_usage_cap=free_cap, free_cap_applied=free_cap_applied, billable_usage=billable_usage,
            tier_breakdown=tier_breakdown, subtotal_usd=subtotal_usd,
            exchange_rate=exchange_rate, margin_rate=margin_rate, final_krw=final_krw,
        ))
    return results


def calculate_billing_by_project(usage_rows, sku_master, exchange_rate, margin_rate=Decimal("1.12")):
    """프로젝트별 청구 계산 → Sheet 2 Project 요약용
    반환: [{'proj_id', 'proj_name', 'skus': {sku_name: {usage, subtotal_usd, final_krw}},
            'total_usd', 'total_krw'}, ...]
    """
    proj_usage = defaultdict(lambda: defaultdict(int))
    proj_names = {}
    for row in usage_rows:
        proj_usage[row.project_id][row.sku_id] += row.usage_amount
        proj_names[row.project_id] = row.project_name

    results = []
    for proj_id in sorted(proj_names.keys()):
        proj_name = proj_names[proj_id]
        skus = {}
        total_usd = Decimal("0")
        total_krw = Decimal("0")
        for sku_id, sku in sku_master.items():
            usage = proj_usage[proj_id].get(sku_id, 0)
            free_cap = sku.free_usage_cap
            billable = max(0, usage - free_cap)
            _, subtotal = _apply_waterfall(billable, sku)
            final_krw = (subtotal * exchange_rate * margin_rate).quantize(Decimal("1"), ROUND_HALF_UP)
            total_usd += subtotal
            total_krw += final_krw
            # usage==0 이어도 포함 (pivot 테이블 컬럼 일관성 유지)
            skus[sku.sku_name] = {"usage": usage, "subtotal_usd": subtotal, "final_krw": final_krw}
        results.append({
            "proj_id": proj_id,
            "proj_name": proj_name,
            "skus": skus,
            "total_usd": total_usd,
            "total_krw": total_krw,
        })
    return results


def _apply_waterfall(billable_usage, sku):
    remaining = billable_usage
    breakdown, subtotal_usd, cum_lower = [], Decimal("0"), 0
    for tier in sorted(sku.tiers, key=lambda t: t.tier_number):
        if tier.tier_limit is None:
            usage_in_tier = max(0, remaining)
        else:
            capacity = tier.tier_limit - cum_lower
            usage_in_tier = max(0, min(remaining, capacity))
            cum_lower = tier.tier_limit
        amt = (Decimal(usage_in_tier) / Decimal("1000")) * tier.tier_cpm
        breakdown.append(TierBreakdown(
            tier_number=tier.tier_number, usage_in_tier=usage_in_tier,
            tier_cpm=tier.tier_cpm, amount_usd=amt,
        ))
        subtotal_usd += amt
        remaining -= usage_in_tier
    return breakdown, subtotal_usd
