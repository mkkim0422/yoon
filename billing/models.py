"""
데이터 모델 정의 - DB 레코드를 Python에서 다루기 위한 dataclass
모든 금액/단가 필드는 Decimal 타입 강제
"""
from dataclasses import dataclass, field
from decimal import Decimal
from typing import Optional


@dataclass
class SkuTier:
    """SKU 구간 단가 정보 (sku_tiers 테이블 1행)"""
    tier_number: int
    tier_limit: Optional[int]   # None이면 해당 구간 상한 없음 (마지막 구간)
    tier_cpm: Decimal           # 1,000건당 단가 (USD)


@dataclass
class Sku:
    """SKU 마스터 정보 (skus + sku_tiers JOIN)"""
    sku_id: str
    sku_name: str
    is_billable: bool
    category: str
    free_usage_cap: int         # 월별 무료 제공량 (건수)
    tiers: list[SkuTier] = field(default_factory=list)

    def __post_init__(self):
        # tier_number 오름차순 정렬 보장 (Waterfall 순서)
        self.tiers.sort(key=lambda t: t.tier_number)


@dataclass
class UsageRow:
    """원시 사용량 1행"""
    billing_month: str          # 'YYYY-MM'
    project_id: str
    project_name: str
    sku_id: str
    usage_amount: int
    cost_krw: Optional[Decimal] = None   # 실제 청구 KRW 비용 (CSV에서 읽음)
    unit_price: Optional[float] = None   # 공시 단가 (CSV 단가 컬럼 또는 None)


@dataclass
class TierBreakdown:
    """구간별 과금 내역 (디버깅/인보이스 세부내역용)"""
    tier_number: int
    usage_in_tier: int          # 해당 구간에서 소비된 사용량
    tier_cpm: Decimal
    amount_usd: Decimal         # usage_in_tier / 1000 * tier_cpm


@dataclass
class BillingLineItem:
    """프로젝트 + SKU 단위 정산 결과 1행"""
    billing_month: str
    project_id: str
    project_name: str
    sku_id: str
    sku_name: str
    total_usage: int
    free_usage_cap: int         # SKU 고정 무료 한도 (표기용)
    free_cap_applied: int       # 실제 차감된 무료 사용량
    billable_usage: int
    tier_breakdown: list[TierBreakdown]
    subtotal_usd: Decimal
    exchange_rate: Decimal
    margin_rate: Decimal        # 기본값 Decimal('1.12')
    final_krw: Decimal
