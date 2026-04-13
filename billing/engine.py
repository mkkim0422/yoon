from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict
from billing.models import BillingLineItem, TierBreakdown


def calculate_billing(usage_rows, sku_master, exchange_rate, margin_rate=Decimal("1.0")):  # noqa: ARG001
    """전체 합산 청구 계산 → Sheet 1 Invoice용"""
    usage_map   = defaultdict(int)
    cost_krw_map = defaultdict(Decimal)
    for row in usage_rows:
        usage_map[row.sku_id] += row.usage_amount
        if row.cost_krw is not None:
            cost_krw_map[row.sku_id] += row.cost_krw

    results = []
    for sku_id, sku in sku_master.items():
        total_usage = usage_map.get(sku_id, 0)
        if total_usage == 0:
            continue  # 미사용 SKU 제외
        free_cap = sku.free_usage_cap
        free_cap_applied = min(total_usage, free_cap)
        billable_usage = max(0, total_usage - free_cap)

        # 실제 KRW 비용이 있으면 그걸로 USD 역산, 없으면 waterfall 계산
        # tier_breakdown은 항상 waterfall로 산출 (인보이스 표시용)
        if cost_krw_map.get(sku_id, Decimal("0")) != Decimal("0"):
            actual_krw = cost_krw_map[sku_id]
            subtotal_usd = (actual_krw / exchange_rate).quantize(Decimal("0.0001"), ROUND_HALF_UP)
            tier_breakdown, _ = _apply_waterfall(billable_usage, sku)
        else:
            tier_breakdown, subtotal_usd = _apply_waterfall(billable_usage, sku)

        final_krw = (subtotal_usd * exchange_rate).quantize(Decimal("1"), ROUND_HALF_UP)
        results.append(BillingLineItem(
            billing_month="", project_id="", project_name="",
            sku_id=sku_id, sku_name=sku.sku_name, total_usage=total_usage,
            free_usage_cap=free_cap, free_cap_applied=free_cap_applied, billable_usage=billable_usage,
            tier_breakdown=tier_breakdown, subtotal_usd=subtotal_usd,
            exchange_rate=exchange_rate, margin_rate=Decimal("1.0"), final_krw=final_krw,
        ))
    return results


def calculate_billing_by_project(usage_rows, sku_master, exchange_rate, margin_rate=Decimal("1.0")):
    """프로젝트별 청구 계산 → Sheet 2 Project 요약용
    반환: [{'proj_id', 'proj_name', 'skus': {sku_name: {usage, subtotal_usd, final_krw}},
            'total_usd', 'total_krw'}, ...]
    """
    proj_usage    = defaultdict(lambda: defaultdict(int))
    proj_cost_krw = defaultdict(lambda: defaultdict(Decimal))
    proj_price    = defaultdict(lambda: defaultdict(lambda: None))
    proj_names    = {}
    for row in usage_rows:
        proj_usage[row.project_id][row.sku_id] += row.usage_amount
        proj_names[row.project_id] = row.project_name
        if row.cost_krw is not None:
            proj_cost_krw[row.project_id][row.sku_id] += row.cost_krw
        if row.unit_price is not None and proj_price[row.project_id][row.sku_id] is None:
            proj_price[row.project_id][row.sku_id] = row.unit_price

    results = []
    for proj_id in sorted(proj_names.keys()):
        proj_name = proj_names[proj_id]
        skus = {}
        total_usd = Decimal("0")
        total_krw = Decimal("0")

        # 실제 청구된 sku_id 집합 (usage > 0 or cost > 0)
        relevant_sku_ids = set(proj_usage[proj_id].keys()) | set(proj_cost_krw[proj_id].keys())

        for sku_id, sku in sku_master.items():
            usage = proj_usage[proj_id].get(sku_id, 0)
            actual_krw = proj_cost_krw[proj_id].get(sku_id, Decimal("0"))

            if actual_krw != Decimal("0"):
                # 실제 KRW 비용으로 USD 역산
                subtotal = (actual_krw / exchange_rate).quantize(Decimal("0.0001"), ROUND_HALF_UP)
                # 단가: CSV 단가 컬럼 → 없으면 cost_usd / usage
                _unit_price = proj_price[proj_id].get(sku_id)
                if _unit_price is None and usage > 0:
                    _unit_price = float(subtotal / Decimal(str(usage)))
            else:
                # 실제 비용 없음 → waterfall 계산 (fallback)
                free_cap = sku.free_usage_cap
                billable = max(0, usage - free_cap)
                _, subtotal = _apply_waterfall(billable, sku)
                _sorted_tiers = sorted(sku.tiers, key=lambda t: t.tier_number)
                _unit_price = (
                    float(_sorted_tiers[0].tier_cpm) / 1000
                    if _sorted_tiers and billable > 0 else None
                )

            final_krw = (subtotal * exchange_rate).quantize(Decimal("1"), ROUND_HALF_UP)
            total_usd += subtotal
            total_krw += final_krw
            # usage==0 이어도 포함 (pivot 테이블 컬럼 일관성 유지)
            skus[sku.sku_name] = {
                "usage":        usage,
                "subtotal_usd": subtotal,
                "final_krw":    final_krw,
                "unit_price":   _unit_price,
            }
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
