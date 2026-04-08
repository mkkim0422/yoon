"""
전처리 모듈(preprocessor.py) 단위 테스트

실행: python -m pytest tests/test_preprocessor.py -v
"""
import io
import textwrap
from pathlib import Path

import pandas as pd
import pytest

from billing.preprocessor import preprocess_usage_file, COLUMN_MAP


# ── 픽스처: 임시 파일 생성 헬퍼 ───────────────────────────────────────────

def make_csv(tmp_path: Path, content: str, filename: str = "test.csv") -> Path:
    """content 문자열을 CSV 파일로 저장 후 경로 반환."""
    p = tmp_path / filename
    p.write_text(textwrap.dedent(content).strip(), encoding="utf-8-sig")
    return p


def make_excel(tmp_path: Path, rows: list[dict], filename: str = "test.xlsx") -> Path:
    """rows 리스트를 Excel 파일로 저장 후 경로 반환."""
    p = tmp_path / filename
    pd.DataFrame(rows).to_excel(p, index=False)
    return p


# ── 공통 Mock 데이터 ───────────────────────────────────────────────────────

VALID_CSV_CONTENT = """\
프로젝트 ID,SKU ID,SKU 설명,비용 유형,사용량
popo-01,maps.dynamic,Dynamic Maps,사용량,150000
popo-01,tax.vat,세금,TAX,1
popo-01,maps.dynamic,Dynamic Maps,RESELLER_MARGIN,999
popo-02,places.details,Places Details,사용량,3000
popo-02,places.details,Places Details,사용량,2000
"""

BILLING_MONTH = "2025-03"


# ── Step 1: 파일 읽기 ─────────────────────────────────────────────────────

class TestFileReading:
    def test_read_csv(self, tmp_path):
        p = make_csv(tmp_path, VALID_CSV_CONTENT)
        result = preprocess_usage_file(p, BILLING_MONTH)
        assert isinstance(result, list)
        assert len(result) > 0

    def test_read_excel(self, tmp_path):
        rows = [
            {"프로젝트 ID": "popo-01", "SKU ID": "maps.dynamic",
             "SKU 설명": "Dynamic Maps", "비용 유형": "사용량", "사용량": "50000"},
        ]
        p = make_excel(tmp_path, rows)
        result = preprocess_usage_file(p, BILLING_MONTH)
        assert len(result) == 1

    def test_file_not_found_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            preprocess_usage_file(tmp_path / "no_such_file.csv", BILLING_MONTH)

    def test_unsupported_extension_raises(self, tmp_path):
        p = tmp_path / "data.txt"
        p.write_text("dummy")
        with pytest.raises(ValueError, match="지원하지 않는 파일 형식"):
            preprocess_usage_file(p, BILLING_MONTH)

    def test_missing_required_column_raises(self, tmp_path):
        bad_csv = "프로젝트 ID,SKU ID,SKU 설명,사용량\npopo-01,maps.dynamic,Dynamic Maps,1000\n"
        p = make_csv(tmp_path, bad_csv)
        with pytest.raises(KeyError, match="비용 유형"):
            preprocess_usage_file(p, BILLING_MONTH)


# ── Step 2: 필터링 ────────────────────────────────────────────────────────

class TestFiltering:
    def setup_method(self):
        """각 테스트에서 공통으로 쓸 CSV 내용 (in-memory)."""
        self._content = VALID_CSV_CONTENT

    def test_reseller_margin_dropped(self, tmp_path):
        """RESELLER_MARGIN 행은 결과에서 제외되어야 한다."""
        p = make_csv(tmp_path, self._content)
        result = preprocess_usage_file(p, BILLING_MONTH)
        for r in result:
            assert r["sku_id"] != "maps.dynamic" or r["usage_amount"] != 999

    def test_tax_sku_dropped(self, tmp_path):
        """'세금' SKU 설명 행은 결과에서 제외되어야 한다."""
        p = make_csv(tmp_path, self._content)
        result = preprocess_usage_file(p, BILLING_MONTH)
        sku_ids = {r["sku_id"] for r in result}
        assert "tax.vat" not in sku_ids

    def test_custom_drop_cost_types(self, tmp_path):
        """drop_cost_types 커스텀 지정 시 해당 비용 유형 드랍."""
        csv = """\
프로젝트 ID,SKU ID,SKU 설명,비용 유형,사용량
popo-01,maps.dynamic,Dynamic Maps,USAGE,10000
popo-01,maps.dynamic,Dynamic Maps,CREDIT,5000
"""
        p = make_csv(tmp_path, csv)
        result = preprocess_usage_file(p, BILLING_MONTH, drop_cost_types=["CREDIT"])
        assert len(result) == 1
        assert result[0]["usage_amount"] == 10000

    def test_custom_drop_sku_names(self, tmp_path):
        """drop_sku_names 커스텀 지정 시 해당 SKU 드랍."""
        csv = """\
프로젝트 ID,SKU ID,SKU 설명,비용 유형,사용량
popo-01,maps.dynamic,Dynamic Maps,사용량,10000
popo-01,promo.xyz,프로모션 크레딧,사용량,9999
"""
        p = make_csv(tmp_path, csv)
        result = preprocess_usage_file(p, BILLING_MONTH, drop_sku_names=["프로모션 크레딧"])
        assert len(result) == 1
        assert result[0]["sku_id"] == "maps.dynamic"


# ── Step 3: 결측치 & 타입 처리 ───────────────────────────────────────────

class TestDataCleaning:
    def test_nan_usage_treated_as_zero(self, tmp_path):
        """'사용량'이 비어있는 행은 0으로 처리."""
        csv = """\
프로젝트 ID,SKU ID,SKU 설명,비용 유형,사용량
popo-01,maps.dynamic,Dynamic Maps,사용량,
"""
        p = make_csv(tmp_path, csv)
        result = preprocess_usage_file(p, BILLING_MONTH)
        assert result[0]["usage_amount"] == 0

    def test_usage_is_int(self, tmp_path):
        """usage_amount는 항상 int 타입이어야 한다."""
        p = make_csv(tmp_path, VALID_CSV_CONTENT)
        result = preprocess_usage_file(p, BILLING_MONTH)
        for r in result:
            assert isinstance(r["usage_amount"], int)

    def test_float_usage_truncated_to_int(self, tmp_path):
        """소수점 사용량(예: 100.9)은 정수로 변환된다."""
        csv = """\
프로젝트 ID,SKU ID,SKU 설명,비용 유형,사용량
popo-01,maps.dynamic,Dynamic Maps,사용량,100.9
"""
        p = make_csv(tmp_path, csv)
        result = preprocess_usage_file(p, BILLING_MONTH)
        assert result[0]["usage_amount"] == 100


# ── Step 4: Groupby + Sum ────────────────────────────────────────────────

class TestGrouping:
    def test_same_project_sku_rows_are_summed(self, tmp_path):
        """동일 project_id + sku_id 행은 합산되어 1건으로 나와야 한다."""
        p = make_csv(tmp_path, VALID_CSV_CONTENT)
        result = preprocess_usage_file(p, BILLING_MONTH)
        popo02 = next(r for r in result if r["project_id"] == "popo-02")
        # 3000 + 2000 = 5000
        assert popo02["usage_amount"] == 5000

    def test_different_projects_are_separate(self, tmp_path):
        """다른 project_id는 별개 항목으로 남아야 한다."""
        p = make_csv(tmp_path, VALID_CSV_CONTENT)
        result = preprocess_usage_file(p, BILLING_MONTH)
        project_ids = {r["project_id"] for r in result}
        assert "popo-01" in project_ids
        assert "popo-02" in project_ids

    def test_different_skus_same_project_are_separate(self, tmp_path):
        """같은 프로젝트라도 sku_id가 다르면 별개 항목."""
        csv = """\
프로젝트 ID,SKU ID,SKU 설명,비용 유형,사용량
popo-01,maps.dynamic,Dynamic Maps,사용량,10000
popo-01,places.details,Places Details,사용량,5000
"""
        p = make_csv(tmp_path, csv)
        result = preprocess_usage_file(p, BILLING_MONTH)
        assert len(result) == 2


# ── Step 5: 반환 형식 ─────────────────────────────────────────────────────

class TestOutputFormat:
    def test_required_keys_present(self, tmp_path):
        """반환 딕셔너리에 load_usage_rows가 요구하는 4개 키가 모두 있어야 한다."""
        p = make_csv(tmp_path, VALID_CSV_CONTENT)
        result = preprocess_usage_file(p, BILLING_MONTH)
        for r in result:
            assert "billing_month" in r
            assert "project_id"    in r
            assert "sku_id"        in r
            assert "usage_amount"  in r

    def test_billing_month_injected(self, tmp_path):
        """billing_month는 함수 인자로 받은 값이 주입되어야 한다."""
        p = make_csv(tmp_path, VALID_CSV_CONTENT)
        result = preprocess_usage_file(p, "2025-06")
        for r in result:
            assert r["billing_month"] == "2025-06"

    def test_end_to_end_popo01_dynamic_maps(self, tmp_path):
        """
        전체 파이프라인 통합:
        popo-01 / maps.dynamic 사용량 150,000건
        (RESELLER_MARGIN 999건은 드랍, 세금 1건 드랍)
        """
        p = make_csv(tmp_path, VALID_CSV_CONTENT)
        result = preprocess_usage_file(p, BILLING_MONTH)
        popo01 = next(r for r in result if r["project_id"] == "popo-01")
        assert popo01["sku_id"]        == "maps.dynamic"
        assert popo01["usage_amount"]  == 150_000
        assert popo01["billing_month"] == BILLING_MONTH
