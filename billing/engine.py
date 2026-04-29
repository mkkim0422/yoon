from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict
from billing.models import BillingLineItem, TierBreakdown


def calculate_billing(usage_rows, sku_master, exchange_rate,
                      margin_rate=Decimal("1.0"),  # noqa: ARG001
                      mode: str = "account",
                      free_cap_override: dict[str, int] | None = None):
    """전체 합산 청구 계산 → Sheet 1 Invoice용.

    mode:
      "account"     — 결제계정(회사) 단위 waterfall. SKU 별 usage 를 모두
                      합친 뒤 tier 적용 (Google 실제 청구 방식, 기본).
      "per_project" — 프로젝트별로 독립 waterfall 을 돌린 뒤, 각 SKU 의
                      tier 별 수량/금액을 프로젝트에 걸쳐 **합산**.

    free_cap_override: {sku_id: adjusted_free_cap} — 계정 무료 제공량을
      프로젝트 간 공유하기 위한 비례 배분값. per_project 모드로 각 프로젝트
      별 Invoice 시트를 생성할 때, 각 프로젝트는 자기 usage 비율만큼만
      무료량을 받도록 호출부에서 미리 계산해 넣는다. 지정된 SKU 는 master
      의 free_usage_cap 대신 이 값이 적용된다.

    master 에 없는 SKU 도 CSV 실제 KRW 비용으로 청구 항목에 포함한다.
    """
    if mode == "per_project":
        return _calculate_billing_per_project(
            usage_rows, sku_master, exchange_rate
        )

    # ── "account" 모드 (기본) ──────────────────────────────────────────
    usage_map    = defaultdict(int)
    cost_krw_map = defaultdict(Decimal)
    sku_names    = {}   # sku_id → CSV sku_name (master 미등록 SKU 이름 복구용)
    for row in usage_rows:
        usage_map[row.sku_id] += row.usage_amount
        if row.cost_krw is not None:
            cost_krw_map[row.sku_id] += row.cost_krw
        if getattr(row, "sku_name", None) and row.sku_id not in sku_names:
            sku_names[row.sku_id] = row.sku_name

    # 실제 청구 대상 SKU = 사용량 발생 OR 비용 발생한 SKU (master 여부 무관)
    relevant_sku_ids = set(usage_map.keys()) | set(cost_krw_map.keys())

    results = []
    for sku_id in relevant_sku_ids:
        total_usage = usage_map.get(sku_id, 0)
        actual_krw  = cost_krw_map.get(sku_id, Decimal("0"))
        if total_usage == 0 and actual_krw == Decimal("0"):
            continue

        sku = sku_master.get(sku_id)
        if sku is not None:
            sku_name = sku.sku_name
            free_cap = sku.free_usage_cap
        else:
            # master 미등록 SKU → CSV sku_name 사용, 무료 한도 0, tier 없음
            sku_name = sku_names.get(sku_id) or sku_id
            free_cap = 0

        # 호출부가 (프로젝트 무료량 비례 배분 등) 무료 한도를 재지정하면 사용
        if free_cap_override is not None and sku_id in free_cap_override:
            free_cap = int(free_cap_override[sku_id])

        free_cap_applied = min(total_usage, free_cap)
        billable_usage   = max(0, total_usage - free_cap)

        # Invoice 엑셀은 SUMIF(Price List) tier 단가 × billable 로 계산한다.
        # Python 도 동일해야 Project 시트 합계 = Invoice 합계 가 된다.
        # → master(= master_data + Price List 보완본) 에 tier 정보가 있으면
        #    waterfall 우선. 없을 때만 actual_krw(CSV cost_krw) fallback.
        if sku is not None:
            tier_breakdown, subtotal_usd = _apply_waterfall(billable_usage, sku)
        elif actual_krw != Decimal("0"):
            subtotal_usd = (actual_krw / exchange_rate).quantize(Decimal("0.0001"), ROUND_HALF_UP)
            tier_breakdown = []
        else:
            tier_breakdown, subtotal_usd = [], Decimal("0")

        final_krw = (subtotal_usd * exchange_rate).quantize(Decimal("1"), ROUND_HALF_UP)
        results.append(BillingLineItem(
            billing_month="", project_id="", project_name="",
            sku_id=sku_id, sku_name=sku_name, total_usage=total_usage,
            free_usage_cap=free_cap, free_cap_applied=free_cap_applied, billable_usage=billable_usage,
            tier_breakdown=tier_breakdown, subtotal_usd=subtotal_usd,
            exchange_rate=exchange_rate, margin_rate=Decimal("1.0"), final_krw=final_krw,
        ))
    return results


def _calculate_billing_per_project(usage_rows, sku_master, exchange_rate):
    """프로젝트별 독립 waterfall → SKU 단위로 tier/amount 합산.

    account 모드와 LineItem 인터페이스는 동일. 차이점:
    - 각 tier_breakdown.usage_in_tier/amount_usd 는 "프로젝트별 waterfall
      결과의 tier별 합" 으로, tier_limit 한도를 초과할 수 있다.
    - subtotal_usd = tier_breakdown amount 합 (프로젝트별 합과 동일).
    """
    # (project_id, sku_id) → usage, krw, sku_name
    proj_sku_usage = defaultdict(lambda: defaultdict(int))
    proj_sku_krw   = defaultdict(lambda: defaultdict(Decimal))
    sku_names      = {}
    for row in usage_rows:
        proj_sku_usage[row.project_id][row.sku_id] += row.usage_amount
        if row.cost_krw is not None:
            proj_sku_krw[row.project_id][row.sku_id] += row.cost_krw
        if getattr(row, "sku_name", None) and row.sku_id not in sku_names:
            sku_names[row.sku_id] = row.sku_name

    # SKU 별 집계 — total_usage/subtotal 은 더하기, tier_breakdown 은 tier_number 별로 합산
    agg_usage: dict[str, int] = defaultdict(int)
    agg_actual_krw: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    agg_subtotal: dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
    # tier_number → {"usage_in_tier": int, "tier_cpm": Decimal, "amount_usd": Decimal}
    agg_tiers: dict[str, dict[int, dict]] = defaultdict(dict)

    for proj_id, sku_usage_map in proj_sku_usage.items():
        for sku_id, usage in sku_usage_map.items():
            actual_krw = proj_sku_krw[proj_id].get(sku_id, Decimal("0"))
            sku = sku_master.get(sku_id)
            free_cap = sku.free_usage_cap if sku else 0
            billable = max(0, usage - free_cap)

            agg_usage[sku_id] += usage
            agg_actual_krw[sku_id] += actual_krw

            # master/Price List 에 tier 데이터가 있으면 waterfall 우선
            # (Invoice 엑셀 SUMIF 결과와 일치). 없을 때만 actual_krw fallback.
            if sku is not None:
                breakdown, subtotal_usd = _apply_waterfall(billable, sku)
                agg_subtotal[sku_id] += subtotal_usd
                _merge_tier_breakdown(agg_tiers[sku_id], breakdown)
            elif actual_krw != Decimal("0"):
                subtotal_usd = (actual_krw / exchange_rate).quantize(
                    Decimal("0.0001"), ROUND_HALF_UP
                )
                agg_subtotal[sku_id] += subtotal_usd
            # master 도 없고 비용도 없으면 skip (subtotal=0 유지)

    results = []
    for sku_id, total_usage in agg_usage.items():
        subtotal = agg_subtotal[sku_id]
        actual_krw = agg_actual_krw[sku_id]
        if total_usage == 0 and actual_krw == Decimal("0"):
            continue
        sku = sku_master.get(sku_id)
        sku_name = sku.sku_name if sku else (sku_names.get(sku_id) or sku_id)
        free_cap = sku.free_usage_cap if sku else 0
        free_cap_applied = min(total_usage, free_cap)
        billable_usage   = max(0, total_usage - free_cap)

        # tier_breakdown: tier_number 오름차순 TierBreakdown 리스트로 변환
        tier_breakdown = [
            TierBreakdown(
                tier_number=tn,
                usage_in_tier=t["usage_in_tier"],
                tier_cpm=t["tier_cpm"],
                amount_usd=t["amount_usd"],
            )
            for tn, t in sorted(agg_tiers.get(sku_id, {}).items())
        ]
        final_krw = (subtotal * exchange_rate).quantize(Decimal("1"), ROUND_HALF_UP)
        results.append(BillingLineItem(
            billing_month="", project_id="", project_name="",
            sku_id=sku_id, sku_name=sku_name, total_usage=total_usage,
            free_usage_cap=free_cap, free_cap_applied=free_cap_applied,
            billable_usage=billable_usage, tier_breakdown=tier_breakdown,
            subtotal_usd=subtotal, exchange_rate=exchange_rate,
            margin_rate=Decimal("1.0"), final_krw=final_krw,
        ))
    return results


def _merge_tier_breakdown(acc: dict[int, dict], breakdown) -> None:
    """tier_number 별로 usage_in_tier / amount_usd 를 누적."""
    for b in breakdown:
        slot = acc.setdefault(b.tier_number, {
            "usage_in_tier": 0, "tier_cpm": b.tier_cpm,
            "amount_usd": Decimal("0"),
        })
        slot["usage_in_tier"] += b.usage_in_tier
        slot["amount_usd"]    += b.amount_usd
        # tier_cpm 은 동일해야 함 (같은 SKU 같은 tier)


def calculate_billing_by_project(usage_rows, sku_master, exchange_rate,
                                  margin_rate=Decimal("1.0"),  # noqa: ARG001
                                  mode: str = "per_project",
                                  proj_sku_free_cap: dict[str, dict[str, int]] | None = None):
    """프로젝트별 청구 계산 → Sheet 2 Project 요약용.

    mode:
      "per_project" — 각 프로젝트가 자기 usage 만으로 독립 waterfall (기본).
      "account"     — 회사 통합 waterfall 로 SKU 별 총액을 구한 뒤, 각
                      프로젝트에 그 SKU 의 **usage 비율만큼 분배**.

    proj_sku_free_cap: {proj_id: {sku_id: adjusted_free_cap}} — per_project 모드
      전용. 계정 무료 제공량이 프로젝트들 간 공유되도록 호출부가 usage 비율로
      분할해 전달하면, 각 프로젝트가 자기 몫만큼만 무료 한도를 받는다.
      (기본값 None → 기존 동작 유지: 각 프로젝트가 전체 free cap 을 그대로 사용)

    master 에 없는 SKU 는 CSV 실제 KRW 비용으로 프로젝트 시트에 포함.
    반환: [{'proj_id', 'proj_name', 'skus': {sku_name: {...}}, 'total_usd', 'total_krw'}, ...]
    """
    proj_usage    = defaultdict(lambda: defaultdict(int))
    proj_cost_krw = defaultdict(lambda: defaultdict(Decimal))
    proj_price    = defaultdict(lambda: defaultdict(lambda: None))
    proj_names    = {}
    sku_names     = {}   # sku_id → CSV sku_name
    for row in usage_rows:
        proj_usage[row.project_id][row.sku_id] += row.usage_amount
        proj_names[row.project_id] = row.project_name
        if row.cost_krw is not None:
            proj_cost_krw[row.project_id][row.sku_id] += row.cost_krw
        if row.unit_price is not None and proj_price[row.project_id][row.sku_id] is None:
            proj_price[row.project_id][row.sku_id] = row.unit_price
        if getattr(row, "sku_name", None) and row.sku_id not in sku_names:
            sku_names[row.sku_id] = row.sku_name

    # ── account 모드 전용: SKU 별 전체 합계로 waterfall 1 회 → 프로젝트에 분배 ──
    # 각 SKU 의 "account waterfall subtotal(USD)" 과 "전체 usage" 를 구해 두고,
    # 나중에 각 프로젝트 subtotal 을 (proj_usage / total_usage) 비율로 분배한다.
    acct_sku_subtotal: dict[str, Decimal] = {}
    acct_sku_total_usage: dict[str, int]  = {}
    if mode == "account":
        sku_total_usage: dict[str, int]  = defaultdict(int)
        sku_total_krw:   dict[str, Decimal] = defaultdict(lambda: Decimal("0"))
        for _pid, _m in proj_usage.items():
            for _sid, _u in _m.items():
                sku_total_usage[_sid] += _u
        for _pid, _m in proj_cost_krw.items():
            for _sid, _k in _m.items():
                sku_total_krw[_sid] += _k
        for _sid in set(sku_total_usage.keys()) | set(sku_total_krw.keys()):
            _total_usage = sku_total_usage[_sid]
            _total_krw   = sku_total_krw[_sid]
            _sku         = sku_master.get(_sid)
            _free_cap    = _sku.free_usage_cap if _sku else 0
            _billable    = max(0, _total_usage - _free_cap)
            if _total_krw != Decimal("0"):
                _sub = (_total_krw / exchange_rate).quantize(
                    Decimal("0.0001"), ROUND_HALF_UP
                )
            elif _sku is not None:
                _, _sub = _apply_waterfall(_billable, _sku)
            else:
                _sub = Decimal("0")
            acct_sku_subtotal[_sid] = _sub
            acct_sku_total_usage[_sid] = _total_usage

    results = []
    for proj_id in sorted(proj_names.keys()):
        proj_name = proj_names[proj_id]
        skus = {}
        total_usd = Decimal("0")
        total_krw = Decimal("0")

        # 이 프로젝트가 실제 사용했거나 비용이 발생한 SKU만 (master 여부 무관)
        relevant_sku_ids = set(proj_usage[proj_id].keys()) | set(proj_cost_krw[proj_id].keys())

        for sku_id in relevant_sku_ids:
            usage      = proj_usage[proj_id].get(sku_id, 0)
            actual_krw = proj_cost_krw[proj_id].get(sku_id, Decimal("0"))
            sku        = sku_master.get(sku_id)

            if sku is not None:
                sku_display_name = sku.sku_name
                free_cap         = sku.free_usage_cap
            else:
                sku_display_name = sku_names.get(sku_id) or sku_id
                free_cap         = 0

            # 계정 무료 제공량 공유 (per_project 모드): 호출부가 사전 분배한
            # 값으로 덮어써 각 프로젝트가 자기 몫만 무료로 받게 한다.
            if (proj_sku_free_cap is not None
                    and proj_id in proj_sku_free_cap
                    and sku_id in proj_sku_free_cap[proj_id]):
                free_cap = int(proj_sku_free_cap[proj_id][sku_id])

            if mode == "account":
                # account waterfall 결과를 usage 비율로 분배
                _acct_sub   = acct_sku_subtotal.get(sku_id, Decimal("0"))
                _acct_usage = acct_sku_total_usage.get(sku_id, 0)
                if _acct_usage > 0:
                    _ratio   = Decimal(usage) / Decimal(_acct_usage)
                    subtotal = (_acct_sub * _ratio).quantize(
                        Decimal("0.0001"), ROUND_HALF_UP
                    )
                else:
                    subtotal = Decimal("0")
                _unit_price = proj_price[proj_id].get(sku_id)
                if _unit_price is None and usage > 0 and subtotal > 0:
                    _unit_price = float(subtotal / Decimal(str(usage)))
            elif sku is not None:
                # master(+Price List 보완) 에 tier 데이터 존재 → waterfall.
                # Invoice 엑셀 SUMIF 결과와 일치시키려면 actual_krw 가 있어도
                # waterfall 을 우선.
                billable = max(0, usage - free_cap)
                _, subtotal = _apply_waterfall(billable, sku)
                _sorted_tiers = sorted(sku.tiers, key=lambda t: t.tier_number)
                _unit_price = (
                    float(_sorted_tiers[0].tier_cpm) / 1000
                    if _sorted_tiers and billable > 0 else None
                )
            elif actual_krw != Decimal("0"):
                # master/Price List 모두 미등록 → actual_krw fallback
                subtotal = (actual_krw / exchange_rate).quantize(Decimal("0.0001"), ROUND_HALF_UP)
                _unit_price = proj_price[proj_id].get(sku_id)
                if _unit_price is None and usage > 0:
                    _unit_price = float(subtotal / Decimal(str(usage)))
            else:
                # master에도 없고 비용도 없음 (거의 발생 안 함) → 0원
                subtotal    = Decimal("0")
                _unit_price = None

            final_krw = (subtotal * exchange_rate).quantize(Decimal("1"), ROUND_HALF_UP)
            total_usd += subtotal
            total_krw += final_krw
            skus[sku_display_name] = {
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
