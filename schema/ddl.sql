-- ============================================================
-- Google Maps API 월별 정산 시스템 - DB 스키마
-- ============================================================

-- SKU 마스터 테이블
-- 각 Google Maps API 제품의 기본 정보 및 무료 제공량 정의
CREATE TABLE skus (
    sku_id          VARCHAR(100)    PRIMARY KEY,
    sku_name        VARCHAR(255)    NOT NULL,
    is_billable     BOOLEAN         NOT NULL DEFAULT TRUE,
    category        VARCHAR(100)    NOT NULL,           -- 예: 'Essentials', 'Pro', 'Enterprise'
    free_usage_cap  INTEGER         NOT NULL DEFAULT 0, -- 월별 프로젝트당 무료 제공량 (건수)
    created_at      TIMESTAMPTZ     NOT NULL DEFAULT NOW(),
    updated_at      TIMESTAMPTZ     NOT NULL DEFAULT NOW()
);

COMMENT ON TABLE skus IS 'Google Maps API SKU 마스터 - 제품 정보 및 무료 제공량';
COMMENT ON COLUMN skus.sku_id IS '고유 식별자 (예: maps.directions, maps.staticmap)';
COMMENT ON COLUMN skus.is_billable IS 'FALSE이면 세금/크레딧 등 과금 제외 항목';
COMMENT ON COLUMN skus.free_usage_cap IS '월별 프로젝트당 무료 차감 건수 (0이면 무료 제공 없음)';


-- SKU 구간 단가 테이블 (N구간 확장 가능)
-- 하나의 SKU에 대해 여러 구간을 정의. tier_number 오름차순으로 Waterfall 적용
CREATE TABLE sku_tiers (
    id              SERIAL          PRIMARY KEY,
    sku_id          VARCHAR(100)    NOT NULL REFERENCES skus(sku_id) ON DELETE CASCADE,
    tier_number     SMALLINT        NOT NULL CHECK (tier_number >= 1),
    tier_limit      INTEGER,                            -- NULL이면 해당 구간 상한 없음 (마지막 구간)
    tier_cpm        NUMERIC(12, 6)  NOT NULL,           -- 1,000건당 단가 ($), Decimal 정밀도 확보
    UNIQUE (sku_id, tier_number)
);

COMMENT ON TABLE sku_tiers IS 'SKU별 N구간 Waterfall 단가표';
COMMENT ON COLUMN sku_tiers.tier_limit IS '해당 구간의 누적 상한 건수. NULL = 무제한 (최종 구간)';
COMMENT ON COLUMN sku_tiers.tier_cpm IS '1,000건당 단가 (USD). NUMERIC으로 부동소수점 오차 방지';


-- 월별 원시 사용량 입력 테이블
-- Admin에서 고객사별 프로젝트/SKU 데이터를 업로드
CREATE TABLE usage_raw (
    id              BIGSERIAL       PRIMARY KEY,
    billing_month   CHAR(7)         NOT NULL,           -- 'YYYY-MM' 형식
    project_id      VARCHAR(100)    NOT NULL,           -- 고객사 프로젝트 식별자
    sku_id          VARCHAR(100)    NOT NULL REFERENCES skus(sku_id),
    usage_amount    BIGINT          NOT NULL DEFAULT 0, -- 해당 월 해당 SKU 사용 건수
    UNIQUE (billing_month, project_id, sku_id)
);

COMMENT ON TABLE usage_raw IS '고객사별 월별 SKU 원시 사용량';


-- 월별 환율 테이블
CREATE TABLE exchange_rates (
    billing_month   CHAR(7)         PRIMARY KEY,        -- 'YYYY-MM'
    usd_to_krw      NUMERIC(10, 4)  NOT NULL            -- USD → KRW 환율
);

COMMENT ON TABLE exchange_rates IS '월별 적용 환율 (USD → KRW)';


-- 정산 결과 테이블 (계산 후 저장)
CREATE TABLE billing_results (
    id                  BIGSERIAL       PRIMARY KEY,
    billing_month       CHAR(7)         NOT NULL,
    project_id          VARCHAR(100)    NOT NULL,
    sku_id              VARCHAR(100)    NOT NULL REFERENCES skus(sku_id),
    total_usage         BIGINT          NOT NULL,
    free_cap_applied    INTEGER         NOT NULL,
    billable_usage      BIGINT          NOT NULL,
    subtotal_usd        NUMERIC(18, 6)  NOT NULL,       -- 세전 달러 금액
    exchange_rate       NUMERIC(10, 4)  NOT NULL,
    margin_rate         NUMERIC(6, 4)   NOT NULL,       -- 기본값 1.12
    final_krw           NUMERIC(18, 2)  NOT NULL,       -- 최종 원화 금액
    calculated_at       TIMESTAMPTZ     NOT NULL DEFAULT NOW(),
    UNIQUE (billing_month, project_id, sku_id)
);

COMMENT ON TABLE billing_results IS '월별 확정 정산 결과';


-- ============================================================
-- 샘플 마스터 데이터
-- ============================================================

INSERT INTO skus (sku_id, sku_name, is_billable, category, free_usage_cap) VALUES
    ('maps.directions',        'Directions API',             TRUE,  'Essentials', 40000),
    ('maps.staticmap',         'Maps Static API',            TRUE,  'Essentials', 100000),
    ('maps.geocoding',         'Geocoding API',              TRUE,  'Essentials', 40000),
    ('maps.places.details',    'Place Details API',          TRUE,  'Pro',        5000),
    ('maps.places.search',     'Places Nearby Search API',   TRUE,  'Pro',        5000),
    ('credit.google',          'Google Credit',              FALSE, 'Credit',     0),
    ('tax.vat',                'VAT',                        FALSE, 'Tax',        0);

-- Directions API 구간 단가 (3구간)
INSERT INTO sku_tiers (sku_id, tier_number, tier_limit, tier_cpm) VALUES
    ('maps.directions', 1, 100000,  5.000000),   -- 0 ~ 100,000건: $5.00/1000
    ('maps.directions', 2, 500000,  4.000000),   -- 100,001 ~ 500,000건: $4.00/1000
    ('maps.directions', 3, NULL,    3.000000);   -- 500,001건~: $3.00/1000

-- Maps Static API 구간 단가
INSERT INTO sku_tiers (sku_id, tier_number, tier_limit, tier_cpm) VALUES
    ('maps.staticmap', 1, 100000,  2.000000),
    ('maps.staticmap', 2, NULL,    1.600000);

-- Geocoding API 구간 단가
INSERT INTO sku_tiers (sku_id, tier_number, tier_limit, tier_cpm) VALUES
    ('maps.geocoding', 1, 100000,  5.000000),
    ('maps.geocoding', 2, 500000,  4.000000),
    ('maps.geocoding', 3, NULL,    3.000000);

-- Place Details API 구간 단가
INSERT INTO sku_tiers (sku_id, tier_number, tier_limit, tier_cpm) VALUES
    ('maps.places.details', 1, 100000,  17.000000),
    ('maps.places.details', 2, NULL,    13.000000);

-- Places Nearby Search API 구간 단가
INSERT INTO sku_tiers (sku_id, tier_number, tier_limit, tier_cpm) VALUES
    ('maps.places.search', 1, 100000,  32.000000),
    ('maps.places.search', 2, NULL,    24.000000);
