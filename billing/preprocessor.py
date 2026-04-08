"""
구글 Maps 플랫폼 사용고지서 전처리 모듈

[처리 흐름]
Step 1. CSV / Excel 파일 읽기 (확장자 자동 감지)
Step 2. 불필요한 행 드랍  - '비용 유형' == RESELLER_MARGIN
                          - 'SKU 설명'  == 세금 (또는 커스텀 목록)
Step 3. 결측치 처리        - '사용량' NaN → 0, 정수 변환
Step 4. Groupby + Sum     - ['프로젝트 ID', 'SKU ID', 'SKU 설명'] 기준 사용량 합산
Step 5. 딕셔너리 리스트 반환 → load_usage_rows() 바로 주입 가능한 형태
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd


# ── 원본 컬럼명 → 내부 키 매핑 ────────────────────────────────────────────
# 실제 고지서 컬럼명이 달라질 경우 이 딕셔너리만 수정하면 됨
COLUMN_MAP: dict[str, str] = {
    "프로젝트 ID":   "project_id",
    "프로젝트 이름": "project_name",
    "SKU ID":        "sku_id",
    "SKU 설명":      "sku_name",
    "비용 유형":     "cost_type",
    "사용량":        "usage_amount",
}

# 헤더 행 탐지 기준: 이 문자열 중 하나가 포함된 행을 헤더로 간주
# (한글 고지서 / 영문 고지서 모두 대응)
HEADER_ANCHORS: tuple[str, ...] = ("프로젝트 ID", "Project ID")

# Step 2 기본 드랍 기준
DEFAULT_DROP_COST_TYPES: list[str] = ["RESELLER_MARGIN"]
DEFAULT_DROP_SKU_NAMES:  list[str] = ["세금"]

# 결제 계정 이름 컬럼 (company filter 용)
COMPANY_COL: str = "결제 계정 이름"


def extract_company_names(
    file_path: str | Path,
    encoding: str = "utf-8-sig",
) -> list[str]:
    """
    파일에서 '결제 계정 이름' 컬럼의 유니크 값을 정렬된 리스트로 반환.

    해당 컬럼이 없거나 파일을 읽을 수 없으면 빈 리스트를 반환.
    """
    try:
        df = _read_file(Path(file_path), encoding)
    except Exception:
        return []
    if COMPANY_COL not in df.columns:
        return []
    return sorted(
        v for v in df[COMPANY_COL].dropna().unique().tolist() if str(v).strip()
    )


def preprocess_usage_file(
    file_path: str | Path,
    billing_month: str,
    *,
    drop_cost_types: list[str] | None = None,
    drop_sku_names:  list[str] | None = None,
    encoding: str = "utf-8-sig",   # BOM 포함 UTF-8 대응 (구글 CSV 기본값)
    company_filter: str | None = None,  # 특정 결제 계정만 필터링
) -> list[dict[str, Any]]:
    """
    구글 Maps 플랫폼 사용고지서(CSV / Excel)를 읽어 정제한 뒤
    load_usage_rows()에 바로 주입할 수 있는 딕셔너리 리스트를 반환한다.

    Args:
        file_path:        CSV 또는 Excel 파일 경로
        billing_month:    정산 월 'YYYY-MM' (파일에 없으므로 외부에서 주입)
        drop_cost_types:  드랍할 '비용 유형' 값 목록 (None → 기본값 사용)
        drop_sku_names:   드랍할 'SKU 설명' 값 목록 (None → 기본값 사용)
        encoding:         CSV 인코딩 (Excel은 무시됨)

    Returns:
        [{"billing_month": ..., "project_id": ..., "sku_id": ..., "usage_amount": ...}, ...]

    Raises:
        FileNotFoundError: 파일이 존재하지 않을 때
        KeyError:          필수 컬럼이 파일에 없을 때 (누락 컬럼명 포함)
        ValueError:        지원하지 않는 파일 확장자일 때
    """
    file_path = Path(file_path)

    # ── Step 1: 파일 읽기 ────────────────────────────────────────────────
    df = _read_file(file_path, encoding)

    # 결제 계정 필터링 (company_filter 지정 시, 해당 계정 행만 남김)
    if company_filter is not None and COMPANY_COL in df.columns:
        df = df[df[COMPANY_COL] == company_filter].copy()

    # 필수 컬럼 존재 여부 검증
    _validate_columns(df, required=list(COLUMN_MAP.keys()))

    # 내부 키로 리네임 (분석 편의)
    df = df.rename(columns=COLUMN_MAP)

    # ── Step 2: 불필요한 행 드랍 ─────────────────────────────────────────
    drop_ctypes = drop_cost_types if drop_cost_types is not None else DEFAULT_DROP_COST_TYPES
    drop_snames = drop_sku_names  if drop_sku_names  is not None else DEFAULT_DROP_SKU_NAMES

    mask_cost = df["cost_type"].isin(drop_ctypes)
    mask_sku  = df["sku_name"].isin(drop_snames)
    df = df[~(mask_cost | mask_sku)].copy()

    # ── Step 3: 결측치 & 타입 처리 ───────────────────────────────────────
    df["usage_amount"] = (
        df["usage_amount"]
        .fillna("")                              # NaN → 빈 문자열
        .astype(str)
        .str.strip()
        .str.replace(",", "", regex=False)       # "3,909" → "3909"
        .replace("", "0")                        # 빈 문자열 → "0"
        .astype(float)                           # 소수점 문자열 대비 float 경유
        .astype(int)                             # 최종 정수
    )

    # ── Step 4: Groupby + Sum ─────────────────────────────────────────────
    grouped = (
        df.groupby(["project_id", "project_name", "sku_id", "sku_name"], as_index=False)
          .agg(usage_amount=("usage_amount", "sum"))
    )

    # ── Step 5: 딕셔너리 리스트 변환 ─────────────────────────────────────
    records: list[dict[str, Any]] = []
    for row in grouped.itertuples(index=False):
        records.append(
            {
                "billing_month": billing_month,
                "project_id":    row.project_id,
                "project_name":  row.project_name,
                "sku_id":        row.sku_id,
                "usage_amount":  int(row.usage_amount),
            }
        )

    return records


# ── 내부 헬퍼 ─────────────────────────────────────────────────────────────

def _find_header_row_csv(file_path: Path, encoding: str) -> int:
    """
    CSV를 한 줄씩 읽어 HEADER_ANCHORS 중 하나가 포함된 첫 행의 인덱스(0-based)를 반환.

    이 인덱스를 pd.read_csv(skiprows=N)에 전달하면
    메타데이터 줄 수에 관계없이 올바른 헤더 행을 찾아낸다.

    Raises:
        ValueError: 파일 전체를 읽었으나 헤더 행을 찾지 못한 경우
    """
    with open(file_path, encoding=encoding, errors="replace") as f:
        for idx, line in enumerate(f):
            if any(anchor in line for anchor in HEADER_ANCHORS):
                return idx
    raise ValueError(
        f"헤더 행을 찾을 수 없습니다. "
        f"{list(HEADER_ANCHORS)} 중 하나가 포함된 행이 없습니다: {file_path}"
    )


def _find_header_row_excel(file_path: Path) -> int:
    """
    Excel을 헤더 없이 전체 로드한 뒤 HEADER_ANCHORS를 포함하는 행 인덱스를 반환.

    Raises:
        ValueError: 헤더 행을 찾지 못한 경우
    """
    df_raw = pd.read_excel(file_path, header=None, dtype=str)
    for row_idx, row in df_raw.iterrows():
        if any(
            anchor in str(cell)
            for cell in row
            for anchor in HEADER_ANCHORS
        ):
            return int(row_idx)
    raise ValueError(
        f"헤더 행을 찾을 수 없습니다. "
        f"{list(HEADER_ANCHORS)} 중 하나가 포함된 행이 없습니다: {file_path}"
    )


def _read_file(file_path: Path, encoding: str) -> pd.DataFrame:
    """확장자 기반으로 CSV / Excel 읽기 (동적 헤더 탐지 포함)."""
    if not file_path.exists():
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

    suffix = file_path.suffix.lower()

    if suffix == ".csv":
        header_row = _find_header_row_csv(file_path, encoding)
        return pd.read_csv(
            file_path,
            skiprows=header_row,
            encoding=encoding,
            dtype=str,
        )

    if suffix in {".xlsx", ".xls"}:
        header_row = _find_header_row_excel(file_path)
        return pd.read_excel(
            file_path,
            skiprows=header_row,
            dtype=str,
        )

    raise ValueError(
        f"지원하지 않는 파일 형식입니다: '{suffix}' "
        "(지원 형식: .csv / .xlsx / .xls)"
    )


def _validate_columns(df: pd.DataFrame, required: list[str]) -> None:
    """필수 컬럼 누락 시 KeyError 발생."""
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise KeyError(
            f"필수 컬럼이 파일에 없습니다: {missing}\n"
            f"파일의 실제 컬럼: {list(df.columns)}"
        )
