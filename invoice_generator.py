"""
invoice_generator.py — t1 레이아웃 완전 복제 인보이스 생성기

레이아웃:
  Col A   : 빈 여백(margin) — 너비 10 (~1cm)
  Col B-I : 실제 인보이스 내용 (8개 컬럼)

  Row 1  : 이미지 영역 (tt1 @ B1, tt2 @ I1) — 행 높이 수동 지정
  Row 2  : 빈 스페이서
  Row 3  : Invoice Date
  Row 4  : Billing Account Name
  Row 5  : Term of Use
  Row 6  : 테이블 헤더 (API / Usage / Free Usage / Subtotal / 할인구간 / 수량 / 단가 / Amount)
  Row 7+ : 데이터 (API 행, 구간 분해, 소계)
  ...    : 합계 / 환율 / 청구금액(KRW)
  ...    : (부가세 별도)
  ...    : tt3 이미지

사용법:
    python invoice_generator.py          # billing.csv 자동 처리 (기본값)
    python invoice_generator.py --company coupang --billing-month 2026-03
"""
from __future__ import annotations

import argparse
import calendar
import io
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side,
)
from openpyxl.utils import get_column_letter

# 이미지 삽입 지원 여부 동적 확인
try:
    from openpyxl.drawing.image import Image as XLImage
    _HAS_IMAGE = True
except Exception:
    _HAS_IMAGE = False

# ── 프로젝트 내 모듈 ──────────────────────────────────────────────────────────
from billing.engine import calculate_billing
from billing.loader import load_sku_master, load_usage_rows
from billing.preprocessor import preprocess_usage_file

BASE_DIR   = Path(__file__).parent
ASSETS_DIR = BASE_DIR / "assets"
MASTER_CSV = BASE_DIR / "billing" / "master_data.csv"

# ── SKU 마스터 (main.py 와 동일한 단가표) ────────────────────────────────────
from main import SKU_MASTER_ROWS


# ─────────────────────────────────────────────────────────────────────────────
# 레이아웃 상수
# ─────────────────────────────────────────────────────────────────────────────
# A열 = 여백(margin), 실제 컨텐츠는 B열(=2)부터 I열(=9)까지
C1 = 2   # 첫 번째 컨텐츠 열 (B)
CN = 9   # 마지막 컨텐츠 열 (I)


# ─────────────────────────────────────────────────────────────────────────────
# 스타일 상수
# ─────────────────────────────────────────────────────────────────────────────
_THIN = Side(style="thin")
_BORDER_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=10, name="맑은 고딕") -> Font:
    return Font(bold=bold, color=color, size=size, name=name)

def _align(h="center", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

# 색상
C_DARK   = "595959"   # 테이블 헤더 / 청구금액 행 배경
C_SUB    = "F2F2F2"   # 소계 행 배경
C_WHITE  = "FFFFFF"


# ─────────────────────────────────────────────────────────────────────────────
# 메인 공개 함수
# ─────────────────────────────────────────────────────────────────────────────
def generate_formatted_invoice(
    line_items: list,
    company_name: str,
    billing_month: str,          # "YYYY-MM"
    exchange_rate: Decimal,
    margin_rate: Decimal,
    invoice_date: str | None = None,
    output_path: str | Path | None = None,
    bank_name: str = "하나은행",
) -> bytes | None:
    """
    t1 레이아웃과 동일한 인보이스 Excel 을 생성한다.

    output_path 가 주어지면 파일로 저장하고 None 을 반환.
    주어지지 않으면 bytes 를 반환.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    _set_white_background(ws)
    _set_column_widths(ws)
    _write_image_area(ws)
    _write_invoice_info(ws, company_name, billing_month, invoice_date)
    _write_table_header(ws)
    last_data_row = _write_data_rows(ws, line_items)
    bottom_row = _write_summary_rows(ws, line_items, exchange_rate, margin_rate, last_data_row, billing_month, bank_name)
    _write_vat_note(ws, bottom_row)
    _write_bottom_image(ws, bottom_row + 2)
    _set_freeze_pane(ws)

    if output_path:
        wb.save(str(output_path))
        return None
    else:
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 시트 전체 흰색 배경 + 그리드 숨기기
# ─────────────────────────────────────────────────────────────────────────────
def _set_white_background(ws) -> None:
    """그리드라인 숨기기 + 전체 셀 흰색 배경."""
    ws.sheet_view.showGridLines = False
    white = _fill(C_WHITE)
    for row in ws.iter_rows(min_row=1, max_row=250, min_col=1, max_col=12):
        for cell in row:
            cell.fill = white


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 열 너비
# ─────────────────────────────────────────────────────────────────────────────
def _set_column_widths(ws) -> None:
    """
    A : 빈 여백 (~1cm)
    B : API 이름 (넓게)
    C-H : 수치 컬럼 (균일)
    I : Amount — tt2 이미지 기준이므로 충분히 넓게.
    """
    ws.column_dimensions["A"].width = 10   # 여백
    ws.column_dimensions["B"].width = 32   # API 이름
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["I"].width = 18


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: Row 1 — 이미지 영역
# ─────────────────────────────────────────────────────────────────────────────
def _write_image_area(ws) -> None:
    """
    Row 1 : tt1 (B1, 400px) + tt2 (I1, 150px)
    Row 2 : 빈 스페이서
    이미지 파일이 없으면 셀에 텍스트 플레이스홀더를 표시.
    두 이미지 모두 상단 여백 약 0.3cm(12px) 확보.
    """
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.units import pixels_to_EMU

    tt1_path = ASSETS_DIR / "tt1.png"
    tt2_path = ASSETS_DIR / "tt2.png"

    TOP_OFFSET_PX = 12   # 상단 여백 약 0.3cm

    if _HAS_IMAGE and tt1_path.exists():
        img1 = XLImage(str(tt1_path))
        orig_w, orig_h = img1.width, img1.height
        target_w = 400
        img1_h = int(orig_h * target_w / orig_w) if orig_w else target_w
        img1.width  = target_w
        img1.height = img1_h
        ws.row_dimensions[1].height = int(img1_h * 0.75) + 10
        ws.row_dimensions[2].height = 10  # 스페이서

        # B열(index=1), row=0 (1행), 상단 offset 적용
        marker1 = AnchorMarker(col=1, colOff=0, row=0,
                               rowOff=pixels_to_EMU(TOP_OFFSET_PX))
        size1 = XDRPositiveSize2D(pixels_to_EMU(target_w), pixels_to_EMU(img1_h))
        img1.anchor = OneCellAnchor(_from=marker1, ext=size1)
        ws.add_image(img1)
    else:
        ws.row_dimensions[1].height = 110
        ws.row_dimensions[2].height = 10
        ws["B1"] = "[tt1: SPH 회사 로고]"
        ws["B1"].font      = _font(bold=True, size=9, color="888888")
        ws["B1"].alignment = _align("left", "center")

    if _HAS_IMAGE and tt2_path.exists():
        img2 = XLImage(str(tt2_path))
        orig_w, orig_h = img2.width, img2.height
        target_w2 = 150
        img2_h = int(orig_h * target_w2 / orig_w) if orig_w else target_w2
        img2.width  = target_w2
        img2.height = img2_h

        # I열(index=8), row=0 (1행), 상단 offset 적용
        marker2 = AnchorMarker(col=8, colOff=0, row=0,
                               rowOff=pixels_to_EMU(TOP_OFFSET_PX))
        size2 = XDRPositiveSize2D(pixels_to_EMU(target_w2), pixels_to_EMU(img2_h))
        img2.anchor = OneCellAnchor(_from=marker2, ext=size2)
        ws.add_image(img2)
    else:
        ws["I1"] = "[tt2: Google Maps Platform Invoice]"
        ws["I1"].font      = _font(bold=True, size=9, color="888888")
        ws["I1"].alignment = _align("center", "center")


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: Row 3-5 — 인보이스 메타데이터
# ─────────────────────────────────────────────────────────────────────────────
def _write_invoice_info(ws, company_name: str, billing_month: str,
                         invoice_date: str | None) -> None:
    """
    Row 3 : Invoice Date        — B열: 레이블, C열: 값 (병합 없음)
    Row 4 : Billing Account Name
    Row 5 : Term of Use
    배경: 흰색, 테두리: 없음, 정렬: 왼쪽
    """
    if invoice_date is None:
        invoice_date = date.today().strftime("%Y-%m-%d")

    year, month = int(billing_month[:4]), int(billing_month[5:7])
    last_day = calendar.monthrange(year, month)[1]
    term_start = f"{billing_month}-01"
    term_end   = f"{billing_month}-{last_day:02d}"

    # (행번호, B열 레이블, C열 값)
    info_rows = [
        (3, "Invoice Date",         f":  {invoice_date}"),
        (4, "Billing Account Name", f":  {company_name}"),
        (5, "Term of Use",          f":  {term_start}  ~  {term_end}"),
    ]

    _no_border = Border()   # 테두리 없음

    for row_num, label, value in info_rows:
        ws.row_dimensions[row_num].height = 20

        # B열: 레이블 (왼쪽 정렬)
        cell_b = ws.cell(row=row_num, column=C1, value=label)
        cell_b.font      = _font(bold=False, size=10)
        cell_b.fill      = _fill(C_WHITE)
        cell_b.alignment = _align("left", "center")
        cell_b.border    = _no_border

        # C열: 값 (왼쪽 정렬)
        cell_c = ws.cell(row=row_num, column=C1 + 1, value=value)
        cell_c.font      = _font(bold=False, size=10)
        cell_c.fill      = _fill(C_WHITE)
        cell_c.alignment = _align("left", "center")
        cell_c.border    = _no_border

        # D~I열: 흰색 배경, 테두리 없음
        for c in range(C1 + 2, CN + 1):
            cell = ws.cell(row=row_num, column=c)
            cell.fill   = _fill(C_WHITE)
            cell.border = _no_border


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: Row 6 — 테이블 헤더
# ─────────────────────────────────────────────────────────────────────────────
TABLE_HEADER_ROW = 6
HEADERS = ["API", "Usage", "Free Usage", "Subtotal",
           "할인 구간", "수량", "단가", "Amount"]

def _write_table_header(ws) -> None:
    ws.row_dimensions[TABLE_HEADER_ROW].height = 22
    for idx, label in enumerate(HEADERS):
        col = C1 + idx   # B=2, C=3, ..., I=9
        cell = ws.cell(row=TABLE_HEADER_ROW, column=col, value=label)
        cell.font      = _font(bold=True, color=C_WHITE, size=10)
        cell.fill      = _fill(C_DARK)
        cell.alignment = _align("center", "center")
        cell.border    = _BORDER_ALL


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: Row 7+ — 데이터 행
# ─────────────────────────────────────────────────────────────────────────────
TIER_LABELS = {1: "0-100K", 2: "~500K", 3: "~1M", 4: "~5M", 5: "~10M"}
DATA_START_ROW = TABLE_HEADER_ROW + 1   # = 7

# 컬럼 오프셋 매핑 (0-based → actual column number)
#  0=API, 1=Usage, 2=FreeUsage, 3=Subtotal, 4=할인구간, 5=수량, 6=단가, 7=Amount
def _col(idx: int) -> int:
    """0-based 컬럼 인덱스 → 실제 열 번호 (B=2 시작)"""
    return C1 + idx

def _write_data_rows(ws, line_items: list) -> int:
    """
    데이터를 기록하고 마지막으로 사용된 행 번호를 반환.
    """
    curr = DATA_START_ROW

    for item in sorted(line_items, key=lambda x: x.sku_name):
        n = len(item.tier_breakdown)
        billable_disp = item.billable_usage if item.billable_usage > 0 else "-"

        # ── 왼쪽 4열 (API, Usage, Free Usage, Subtotal) — 구간 수만큼 병합 ──
        if n > 1:
            _merge_write(ws, curr, curr + n - 1, _col(0), item.sku_name,
                         fmt_data="center", wrap=True)
            _merge_write(ws, curr, curr + n - 1, _col(1), item.total_usage,
                         fmt_data="number")
            _merge_write(ws, curr, curr + n - 1, _col(2),
                         f"-{item.free_usage_cap:,}", fmt_data="center")
            _merge_write(ws, curr, curr + n - 1, _col(3), billable_disp,
                         fmt_data="number" if billable_disp != "-" else "center")
        elif n == 1:
            _cell_write(ws, curr, _col(0), item.sku_name,  h="center", wrap=True)
            _cell_write(ws, curr, _col(1), item.total_usage, h="right", num="#,##0")
            _cell_write(ws, curr, _col(2), f"-{item.free_usage_cap:,}", h="center")
            _cell_write(ws, curr, _col(3), billable_disp,
                        h="right" if billable_disp != "-" else "center",
                        num="#,##0" if billable_disp != "-" else "General")
        else:
            # tier 없음 — 1행
            _cell_write(ws, curr, _col(0), item.sku_name,  h="center", wrap=True)
            _cell_write(ws, curr, _col(1), item.total_usage, h="right", num="#,##0")
            _cell_write(ws, curr, _col(2), f"-{item.free_usage_cap:,}", h="center")
            _cell_write(ws, curr, _col(3), "-", h="center")
            _cell_write(ws, curr, _col(4), "-", h="center")
            _cell_write(ws, curr, _col(5), "-", h="center")
            _cell_write(ws, curr, _col(6), "-", h="center")
            _cell_write(ws, curr, _col(7), 0.0, h="right", num="$#,##0.00")
            curr += 1

        # ── 오른쪽 4열 (할인구간, 수량, 단가, Amount) — 구간별 ──────────────
        for i, tb in enumerate(item.tier_breakdown):
            label = TIER_LABELS.get(tb.tier_number, f"T{tb.tier_number}")
            usage = tb.usage_in_tier if tb.usage_in_tier > 0 else "-"
            _cell_write(ws, curr + i, _col(4), label, h="center")
            _cell_write(ws, curr + i, _col(5), usage,
                        h="right" if usage != "-" else "center",
                        num="#,##0" if usage != "-" else "General")
            _cell_write(ws, curr + i, _col(6), float(tb.tier_cpm), h="right",
                        num="$#,##0.00")
            _cell_write(ws, curr + i, _col(7), float(tb.amount_usd), h="right",
                        num="$#,##0.00")

        if n > 0:
            curr += n

        # ── 소계 행 ──────────────────────────────────────────────────────────
        ws.merge_cells(
            start_row=curr, start_column=_col(0),
            end_row=curr,   end_column=_col(4)
        )
        cell = ws.cell(row=curr, column=_col(0), value="소계")
        cell.font      = _font(bold=True)
        cell.fill      = _fill(C_SUB)
        cell.alignment = _align("center", "center")
        cell.border    = _BORDER_ALL
        _apply_border_merged(ws, curr, _col(0), curr, _col(4))

        sub_usage = item.billable_usage if item.billable_usage > 0 else "-"
        _cell_write(ws, curr, _col(5), sub_usage,
                    h="right" if sub_usage != "-" else "center",
                    num="#,##0" if sub_usage != "-" else "General",
                    bold=True, bg=C_SUB)
        _cell_write(ws, curr, _col(6), "",    h="center", bold=True, bg=C_SUB)
        _cell_write(ws, curr, _col(7), float(item.subtotal_usd), h="right",
                    num="$#,##0.00", bold=True, bg=C_SUB)
        curr += 1

    return curr - 1   # 마지막 기록 행


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 합계 / 환율 / 청구금액 행
# ─────────────────────────────────────────────────────────────────────────────
def _write_summary_rows(ws, line_items, exchange_rate, margin_rate,
                         last_data_row: int,
                         billing_month: str = "2026-01",
                         bank_name: str = "하나은행") -> int:
    """합계 ~ 청구금액 3행을 기록하고 청구금액 행 번호를 반환."""
    t_usd = sum((item.subtotal_usd for item in line_items), Decimal("0"))
    t_usd_rounded = t_usd.quantize(Decimal("1"), ROUND_HALF_UP)
    t_krw = (t_usd_rounded * exchange_rate * margin_rate
             ).quantize(Decimal("1"), ROUND_HALF_UP)

    # 정산월 말일 계산 (YYYY.MM.DD 형식)
    year, month = int(billing_month[:4]), int(billing_month[5:7])
    last_day = calendar.monthrange(year, month)[1]
    last_day_str = f"{year}.{month:02d}.{last_day:02d}"

    r = last_data_row + 1

    # ── 합계(USD) — 좌측 정렬 ─────────────────────────────────────────────
    ws.merge_cells(start_row=r, start_column=_col(0), end_row=r, end_column=_col(6))
    cell = ws.cell(row=r, column=_col(0), value="합        계(USD)")
    cell.font      = _font(bold=True)
    cell.alignment = _align("left", "center")
    cell.border    = _BORDER_ALL
    _apply_border_merged(ws, r, _col(0), r, _col(6))
    _cell_write(ws, r, _col(7), float(t_usd_rounded), h="right", num="$#,##0.00", bold=True)

    # ── 환율 — 동적 문구 + 좌측 정렬 ────────────────────────────────────
    r += 1
    ws.merge_cells(start_row=r, start_column=_col(0), end_row=r, end_column=_col(6))
    cell = ws.cell(row=r, column=_col(0),
                   value=f"환        율({bank_name} {last_day_str} 최종 송금환율 기준)")
    cell.font      = _font(bold=True)
    cell.alignment = _align("left", "center")
    cell.border    = _BORDER_ALL
    _apply_border_merged(ws, r, _col(0), r, _col(6))
    _cell_write(ws, r, _col(7), float(exchange_rate), h="right",
                num="$#,##0.00", bold=True)

    # ── 청구 금액(KRW) ──────────────────────────────────────────────────────
    r += 1
    ws.merge_cells(start_row=r, start_column=_col(0), end_row=r, end_column=_col(6))
    cell = ws.cell(row=r, column=_col(0), value="청 구 금 액(KRW)")
    cell.font      = _font(bold=True, color=C_WHITE)
    cell.fill      = _fill(C_DARK)
    cell.alignment = _align("center", "center")
    cell.border    = _BORDER_ALL
    _apply_border_merged(ws, r, _col(0), r, _col(6), bold=True, color=C_WHITE, bg=C_DARK)

    cell8 = ws.cell(row=r, column=_col(7), value=int(t_krw))
    cell8.font         = _font(bold=True, color=C_WHITE)
    cell8.fill         = _fill(C_DARK)
    cell8.alignment    = _align("right", "center")
    cell8.number_format = "#,##0"
    cell8.border       = _BORDER_ALL

    return r   # 청구금액 행 번호


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: (부가세 별도) 텍스트
# ─────────────────────────────────────────────────────────────────────────────
def _write_vat_note(ws, krw_row: int) -> None:
    r = krw_row + 1
    ws.row_dimensions[r].height = 16
    ws.merge_cells(start_row=r, start_column=C1, end_row=r, end_column=CN)
    cell = ws.cell(row=r, column=C1, value="(부가세 별도)")
    cell.font      = _font(size=9, color="595959")
    cell.alignment = _align("right", "center")


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 하단 tt3 이미지
# ─────────────────────────────────────────────────────────────────────────────
def _write_bottom_image(ws, start_row: int) -> None:
    tt3_path = ASSETS_DIR / "tt3.png"

    if _HAS_IMAGE and tt3_path.exists():
        from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
        from openpyxl.drawing.xdr import XDRPositiveSize2D
        from openpyxl.utils.units import pixels_to_EMU

        img3 = XLImage(str(tt3_path))
        orig_w, orig_h = img3.width, img3.height
        target_w = 400
        img3_h = int(orig_h * target_w / orig_w) if orig_w else target_w
        img3.width  = target_w
        img3.height = img3_h

        # 행 높이: 이미지 높이 + 여유
        ws.row_dimensions[start_row].height = int(img3_h * 0.75) + 10

        # 위치 계산: A=10, B=32, C-H=14×6, I=18 (단위: 열너비 문자 수 × 7px/char)
        # tt3 오른쪽 끝 = 표 오른쪽 끝(I열 끝) = "(부가세 별도)" 닫는괄호 위치와 일치
        col_widths_px = [10*7, 32*7, 14*7, 14*7, 14*7, 14*7, 14*7, 14*7, 18*7]
        # col index:        A=0   B=1   C=2   D=3   E=4   F=5   G=6   H=7   I=8
        table_right_px  = sum(col_widths_px)                  # 1008 px
        extra_right_px  = 0                                    # 표 우측 끝에 딱 맞춤
        img_left_px     = (table_right_px - target_w) + extra_right_px  # 608 px

        anchor_col        = 0
        anchor_col_off_px = 0
        cumul = 0
        for i, w in enumerate(col_widths_px):
            if cumul + w > img_left_px:
                anchor_col        = i
                anchor_col_off_px = img_left_px - cumul
                break
            cumul += w

        marker = AnchorMarker(
            col=anchor_col,
            colOff=pixels_to_EMU(anchor_col_off_px),
            row=start_row - 1,
            rowOff=0,
        )
        size = XDRPositiveSize2D(pixels_to_EMU(target_w), pixels_to_EMU(img3_h))
        img3.anchor = OneCellAnchor(_from=marker, ext=size)
        ws.add_image(img3)
    else:
        ws.row_dimensions[start_row].height = 80
        ws.merge_cells(
            start_row=start_row, start_column=C1,
            end_row=start_row,   end_column=CN
        )
        cell = ws.cell(row=start_row, column=C1,
                       value="[tt3: Google Cloud Premier Partner 배지]")
        cell.font      = _font(size=9, color="888888")
        cell.alignment = _align("right", "center")


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 틀 고정 해제
# ─────────────────────────────────────────────────────────────────────────────
def _set_freeze_pane(ws) -> None:
    """틀 고정 해제."""
    ws.freeze_panes = None


# ─────────────────────────────────────────────────────────────────────────────
# 저수준 셀 작성 헬퍼
# ─────────────────────────────────────────────────────────────────────────────
def _cell_write(ws, row: int, col: int, value,
                h="center", v="center", wrap=False,
                num="General", bold=False,
                color="000000", bg=C_WHITE) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font           = _font(bold=bold, color=color)
    cell.fill           = _fill(bg)
    cell.alignment      = _align(h, v, wrap)
    cell.number_format  = num
    cell.border         = _BORDER_ALL
    ws.row_dimensions[row].height = 18


def _merge_write(ws, r1: int, r2: int, col: int, value,
                 fmt_data="center", wrap=False,
                 bold=False, color="000000", bg=C_WHITE) -> None:
    """단일 열 병합 (r1~r2, col 고정)."""
    ws.merge_cells(
        start_row=r1, start_column=col,
        end_row=r2,   end_column=col
    )
    cell = ws.cell(row=r1, column=col, value=value)
    h = "right" if fmt_data == "number" else "center"
    num = "#,##0" if fmt_data == "number" else "General"
    cell.font           = _font(bold=bold, color=color)
    cell.fill           = _fill(bg)
    cell.alignment      = _align(h, "center", wrap)
    cell.number_format  = num
    cell.border         = _BORDER_ALL
    _apply_border_merged(ws, r1, col, r2, col,
                         bold=bold, color=color, bg=bg)


def _apply_border_merged(ws, r1: int, c1: int, r2: int, c2: int,
                          bold=False, color="000000", bg=C_WHITE) -> None:
    """병합된 셀 영역의 모든 셀에 테두리·스타일 적용 (top-left 제외 나머지 채우기)."""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = _BORDER_ALL
            if r != r1 or c != c1:    # top-left 은 이미 값/스타일 설정됨
                cell.fill = _fill(bg)
                cell.font = _font(bold=bold, color=color)


# ─────────────────────────────────────────────────────────────────────────────
# CLI 진입점
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="t1 레이아웃 인보이스 생성기",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("input_file", nargs="?",
                        default="billing.csv",
                        help="사용고지서 파일 (기본: billing.csv)")
    parser.add_argument("-m", "--billing-month", default="2026-03",
                        metavar="YYYY-MM")
    parser.add_argument("-e", "--exchange-rate", type=float,
                        default=1427.87, metavar="RATE")
    parser.add_argument("-r", "--margin-rate", type=float,
                        default=1.12, metavar="RATE")
    parser.add_argument("-c", "--company", default=None,
                        help="특정 회사만 필터링 (예: coupang)")
    parser.add_argument("-o", "--output", default="invoice_output.xlsx",
                        metavar="FILE")
    parser.add_argument("--invoice-date", default=None,
                        help="인보이스 날짜 YYYY-MM-DD (기본: 오늘)")
    args = parser.parse_args()

    INPUT_FILE   = Path(args.input_file)
    OUTPUT_FILE  = BASE_DIR / args.output
    EXCHANGE     = Decimal(str(args.exchange_rate))
    MARGIN       = Decimal(str(args.margin_rate))

    print(f"[1/3] 데이터 전처리: {INPUT_FILE.name}")
    raw_rows = preprocess_usage_file(
        INPUT_FILE, args.billing_month,
        company_filter=args.company,
    )

    print(f"[2/3] 과금 계산 ({len(raw_rows)} 행)")
    sku_master = load_sku_master(SKU_MASTER_ROWS)
    usage_rows = load_usage_rows(raw_rows)
    line_items = calculate_billing(usage_rows, sku_master, EXCHANGE, MARGIN)

    company_display = args.company or "All Companies"
    print(f"[3/3] 인보이스 생성: {OUTPUT_FILE.name}")
    generate_formatted_invoice(
        line_items     = line_items,
        company_name   = company_display,
        billing_month  = args.billing_month,
        exchange_rate  = EXCHANGE,
        margin_rate    = MARGIN,
        invoice_date   = args.invoice_date,
        output_path    = OUTPUT_FILE,
    )

    print(f"\n완료 -> {OUTPUT_FILE}")
    print(f"  항목 수: {len(line_items)}건")
    for it in line_items:
        print(f"  - {it.sku_name}: ${float(it.subtotal_usd):,.2f}")
