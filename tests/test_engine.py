"""
Waterfall 과금 엔진 단위 테스트

실행: python -m pytest tests/ -v
"""
from decimal import Decimal

import pytest

from billing.engine import calculate_billing, summarize_by_project, _apply_waterfall
from billing.models import Sku, SkuTier, UsageRow


# ── 공통 픽스처 ────────────────────────────────────────────────────────────

def make_sku(
    sku_id: str = "maps.directions",
    free_usage_cap: int = 40_000,
    tiers: list[tuple] | None = None,
    is_billable: bool = True,
) -> Sku:
    """테스트용 SKU 생성 헬퍼. tiers = [(limit, cpm), ...] (limit=None → 무제한)"""
    if tiers is None:
        tiers = [(100_000, "5.000000"), (500_000, "4.000000"), (None, "3.000000")]

    return Sku(
        sku_id=sku_id,
        sku_name="Test SKU",
        is_billable=is_billable,
        category="Essentials",
        free_usage_cap=free_usage_cap,
        tiers=[
            SkuTier(tier_number=i + 1, tier_limit=lim, tier_cpm=Decimal(cpm))
            for i, (lim, cpm) in enumerate(tiers)
        ],
    )


EXCHANGE_RATE = Decimal("1350")
MARGIN_RATE = Decimal("1.12")


# ── Step 1: is_billable 필터 ──────────────────────────────────────────────

class TestBillableFilter:
    def test_non_billable_sku_excluded(self):
        sku_master = {
            "maps.directions": make_sku(is_billable=True),
            "tax.vat": make_sku(sku_id="tax.vat", is_billable=False),
        }
        rows = [
            UsageRow("2024-03", "proj-A", "maps.directions", 50_000),
            UsageRow("2024-03", "proj-A", "tax.vat", 999_999),  # 제외 대상
        ]
        results = calculate_billing(rows, sku_master, EXCHANGE_RATE)
        sku_ids = {r.sku_id for r in results}
        assert "tax.vat" not in sku_ids
        assert "maps.directions" in sku_ids

    def test_all_non_billable_returns_empty(self):
        sku_master = {"credit.google": make_sku(sku_id="credit.google", is_billable=False)}
        rows = [UsageRow("2024-03", "proj-A", "credit.google", 100_000)]
        results = calculate_billing(rows, sku_master, EXCHANGE_RATE)
        assert results == []


# ── Step 2: 사용량 합산 ────────────────────────────────────────────────────

class TestUsageAggregation:
    def test_multiple_rows_same_project_sku_are_summed(self):
        sku_master = {"maps.directions": make_sku(free_usage_cap=0)}
        rows = [
            UsageRow("2024-03", "proj-A", "maps.directions", 30_000),
            UsageRow("2024-03", "proj-A", "maps.directions", 20_000),
        ]
        results = calculate_billing(rows, sku_master, EXCHANGE_RATE)
        assert len(results) == 1
        assert results[0].total_usage == 50_000

    def test_different_projects_are_separate_line_items(self):
        sku_master = {"maps.directions": make_sku(free_usage_cap=0)}
        rows = [
            UsageRow("2024-03", "proj-A", "maps.directions", 50_000),
            UsageRow("2024-03", "proj-B", "maps.directions", 80_000),
        ]
        results = calculate_billing(rows, sku_master, EXCHANGE_RATE)
        assert len(results) == 2


# ── Step 3: 무료 사용량 차감 ──────────────────────────────────────────────

class TestFreeCapDeduction:
    def test_usage_within_free_cap_is_zero_billable(self):
        sku_master = {"maps.directions": make_sku(free_usage_cap=40_000)}
        rows = [UsageRow("2024-03", "proj-A", "maps.directions", 30_000)]
        results = calculate_billing(rows, sku_master, EXCHANGE_RATE)
        item = results[0]
        assert item.free_cap_applied == 30_000
        assert item.billable_usage == 0
        assert item.subtotal_usd == Decimal("0")
        assert item.final_krw == Decimal("0")

    def test_free_cap_partially_deducted(self):
        sku_master = {"maps.directions": make_sku(free_usage_cap=40_000)}
        rows = [UsageRow("2024-03", "proj-A", "maps.directions", 50_000)]
        results = calculate_billing(rows, sku_master, EXCHANGE_RATE)
        item = results[0]
        assert item.free_cap_applied == 40_000
        assert item.billable_usage == 10_000

    def test_free_cap_zero_means_no_deduction(self):
        sku_master = {"maps.directions": make_sku(free_usage_cap=0)}
        rows = [UsageRow("2024-03", "proj-A", "maps.directions", 50_000)]
        results = calculate_billing(rows, sku_master, EXCHANGE_RATE)
        assert results[0].billable_usage == 50_000


# ── Step 4: Waterfall 구간 과금 ──────────────────────────────────────────

class TestWaterfallTiers:
    """
    테스트 기준 SKU (free_cap=0):
      Tier 1: 0 ~ 100,000건  @ $5.00/1000
      Tier 2: 100,001 ~ 500,000건  @ $4.00/1000
      Tier 3: 500,001건~  @ $3.00/1000
    """

    def setup_method(self):
        self.sku = make_sku(free_usage_cap=0)

    def test_usage_only_in_tier1(self):
        breakdown, total = _apply_waterfall(50_000, self.sku)
        assert len(breakdown) == 1
        assert breakdown[0].tier_number == 1
        assert breakdown[0].usage_in_tier == 50_000
        # 50000 / 1000 * 5.0 = 250.0
        assert total == Decimal("250.0")

    def test_usage_spanning_tier1_and_tier2(self):
        # 150,000건: 100,000건(T1) + 50,000건(T2)
        breakdown, total = _apply_waterfall(150_000, self.sku)
        assert len(breakdown) == 2
        assert breakdown[0].usage_in_tier == 100_000  # Tier1 전량 소진
        assert breakdown[1].usage_in_tier == 50_000
        # T1: 100000/1000*5 = 500, T2: 50000/1000*4 = 200 → 700
        assert total == Decimal("700.0")

    def test_usage_spanning_all_three_tiers(self):
        # 600,000건: 100,000(T1) + 400,000(T2) + 100,000(T3)
        breakdown, total = _apply_waterfall(600_000, self.sku)
        assert len(breakdown) == 3
        assert breakdown[2].usage_in_tier == 100_000
        # T1: 500, T2: 1600, T3: 300 → 2400
        assert total == Decimal("2400.0")

    def test_exactly_at_tier_boundary(self):
        # 정확히 100,000건 = Tier1 상한
        breakdown, total = _apply_waterfall(100_000, self.sku)
        assert len(breakdown) == 1
        assert total == Decimal("500.0")

    def test_zero_billable_usage_returns_empty(self):
        breakdown, total = _apply_waterfall(0, self.sku)
        assert breakdown == []
        assert total == Decimal("0")


# ── Step 5: 최종 원화 환산 ────────────────────────────────────────────────

class TestFinalKrw:
    def test_krw_calculation(self):
        """50,000건 (free_cap=0) → T1: 50000/1000*5=$250 → $250*1350*1.12=378,000원"""
        sku_master = {"maps.directions": make_sku(free_usage_cap=0)}
        rows = [UsageRow("2024-03", "proj-A", "maps.directions", 50_000)]
        results = calculate_billing(rows, sku_master, EXCHANGE_RATE, MARGIN_RATE)
        item = results[0]
        assert item.subtotal_usd == Decimal("250")
        expected_krw = (Decimal("250") * EXCHANGE_RATE * MARGIN_RATE).quantize(Decimal("1"))
        assert item.final_krw == expected_krw

    def test_custom_exchange_rate_and_margin(self):
        sku_master = {"maps.directions": make_sku(free_usage_cap=0)}
        rows = [UsageRow("2024-03", "proj-A", "maps.directions", 100_000)]
        results = calculate_billing(
            rows, sku_master,
            exchange_rate=Decimal("1400"),
            margin_rate=Decimal("1.15"),
        )
        # 100000/1000*5=$500 → 500*1400*1.15=805,000
        assert results[0].final_krw == Decimal("805000")


# ── 프로젝트별 합계 ────────────────────────────────────────────────────────

class TestSummarizeByProject:
    def test_multiple_skus_per_project_summed(self):
        sku_master = {
            "maps.directions": make_sku(sku_id="maps.directions", free_usage_cap=0),
            "maps.geocoding": make_sku(sku_id="maps.geocoding", free_usage_cap=0),
        }
        rows = [
            UsageRow("2024-03", "proj-A", "maps.directions", 100_000),
            UsageRow("2024-03", "proj-A", "maps.geocoding", 100_000),
        ]
        items = calculate_billing(rows, sku_master, EXCHANGE_RATE)
        totals = summarize_by_project(items)
        # 각각 $500 → 합계 $1000 → 1000*1350*1.12 = 1,512,000
        assert totals["proj-A"] == Decimal("1512000")
