"""
invoice_generator.py — Invoice 시트 생성기

레이아웃:
  Col A   : 빈 여백(margin)
  Col B-I : 실제 인보이스 내용 (8개 컬럼)

  Row 1  : 이미지 영역 (tt1 @ B1, tt2 @ I1)
  Row 2  : 빈 스페이서
  Row 3  : Invoice Date
  Row 4  : Billing Account Name
  Row 5  : Term of Use
  Row 6  : 테이블 헤더
  Row 7+ : 데이터 행
  ...    : 합계 / 환율 / 청구금액(KRW)
  ...    : (부가세 별도)
  ...    : tt3 이미지

사용법:
    python invoice_generator.py
    python invoice_generator.py --company coupang --billing-month 2026-03
"""
from __future__ import annotations

import argparse
import calendar
import io
from copy import copy
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

try:
    from openpyxl.drawing.image import Image as XLImage
    _HAS_IMAGE = True
except Exception:
    _HAS_IMAGE = False

from billing.engine import calculate_billing
from billing.loader import load_sku_master, load_usage_rows
from billing.preprocessor import preprocess_usage_file
from project_sheet import write_project_sheet

BASE_DIR   = Path(__file__).parent
ASSETS_DIR = BASE_DIR / "assets"
MASTER_CSV = BASE_DIR / "billing" / "master_data.csv"

from main import SKU_MASTER_ROWS


# ─────────────────────────────────────────────────────────────────────────────
# 레이아웃 상수
# ─────────────────────────────────────────────────────────────────────────────
C1 = 2   # 첫 번째 컨텐츠 열 (B)
CN = 9   # 마지막 컨텐츠 열 (I)


# ─────────────────────────────────────────────────────────────────────────────
# 스타일 상수
# ─────────────────────────────────────────────────────────────────────────────
_THIN       = Side(style="thin")
_BORDER_ALL = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=10, name="맑은 고딕") -> Font:
    return Font(bold=bold, color=color, size=size, name=name)

def _align(h="center", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

C_DARK  = "595959"
C_SUB   = "F2F2F2"
C_WHITE = "FFFFFF"


# ─────────────────────────────────────────────────────────────────────────────
# 메인 공개 함수
# ─────────────────────────────────────────────────────────────────────────────
def generate_formatted_invoice(
    line_items: list,
    company_name: str,
    billing_month: str,
    exchange_rate: Decimal,
    margin_rate: Decimal,
    invoice_date: str | None = None,
    output_path: str | Path | None = None,
    bank_name: str = "하나은행",
    proj_results: list | None = None,
    price_list_file=None,
) -> bytes | None:
    """
    Invoice 시트를 생성한다.
    proj_results 전달 시 Project 시트(2번)도 추가.
    price_list_file 전달 시 GMP Price List 시트(3번)도 추가.
    output_path 지정 시 파일 저장 후 None 반환, 아니면 bytes 반환.
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
    bottom_row = _write_summary_rows(
        ws, line_items, exchange_rate, margin_rate,
        last_data_row, billing_month, bank_name,
    )
    _write_vat_note(ws, bottom_row)
    _write_bottom_image(ws, bottom_row + 2)
    _set_freeze_pane(ws)

    # 2번 시트: Project (가로 확장형 — project_sheet.write_project_sheet)
    if proj_results:
        write_project_sheet(
            wb, proj_results, company_name, billing_month,
            exchange_rate, margin_rate, invoice_date, bank_name,
            line_items=line_items,
        )

    # 3번 시트: GMP Price List (원본 엑셀 As-is 복제)
    if price_list_file is not None:
        _copy_price_list_sheet(wb, price_list_file)

    if output_path:
        wb.save(str(output_path))
        return None

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 시트 전체 흰색 배경
# ─────────────────────────────────────────────────────────────────────────────
def _set_white_background(ws) -> None:
    ws.sheet_view.showGridLines = False
    white = _fill(C_WHITE)
    for row in ws.iter_rows(min_row=1, max_row=250, min_col=1, max_col=12):
        for cell in row:
            cell.fill = white


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 열 너비
# ─────────────────────────────────────────────────────────────────────────────
def _set_column_widths(ws) -> None:
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 32
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["I"].width = 18


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: Row 1 — 이미지 영역
# ─────────────────────────────────────────────────────────────────────────────
def _write_image_area(ws) -> None:
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.units import pixels_to_EMU

    tt1_path = ASSETS_DIR / "tt1.png"
    tt2_path = ASSETS_DIR / "tt2.png"
    TOP_OFFSET_PX = 12

    if _HAS_IMAGE and tt1_path.exists():
        img1 = XLImage(str(tt1_path))
        orig_w, orig_h = img1.width, img1.height
        target_w = 400
        img1_h = int(orig_h * target_w / orig_w) if orig_w else target_w
        img1.width  = target_w
        img1.height = img1_h
        ws.row_dimensions[1].height = int(img1_h * 0.75) + 10
        ws.row_dimensions[2].height = 10
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
    if invoice_date is None:
        invoice_date = date.today().strftime("%Y-%m-%d")

    year, month = int(billing_month[:4]), int(billing_month[5:7])
    last_day   = calendar.monthrange(year, month)[1]
    term_start = f"{billing_month}-01"
    term_end   = f"{billing_month}-{last_day:02d}"

    info_rows = [
        (3, "Invoice Date",         f":  {invoice_date}"),
        (4, "Billing Account Name", f":  {company_name}"),
        (5, "Term of Use",          f":  {term_start}  ~  {term_end}"),
    ]
    _no_border = Border()

    for row_num, label, value in info_rows:
        ws.row_dimensions[row_num].height = 20

        cell_b = ws.cell(row=row_num, column=C1, value=label)
        cell_b.font      = _font(bold=False, size=10)
        cell_b.fill      = _fill(C_WHITE)
        cell_b.alignment = _align("left", "center")
        cell_b.border    = _no_border

        cell_c = ws.cell(row=row_num, column=C1 + 1, value=value)
        cell_c.font      = _font(bold=False, size=10)
        cell_c.fill      = _fill(C_WHITE)
        cell_c.alignment = _align("left", "center")
        cell_c.border    = _no_border

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
        col  = C1 + idx
        cell = ws.cell(row=TABLE_HEADER_ROW, column=col, value=label)
        cell.font      = _font(bold=True, color=C_WHITE, size=10)
        cell.fill      = _fill(C_DARK)
        cell.alignment = _align("center", "center")
        cell.border    = _BORDER_ALL


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: Row 7+ — 데이터 행
# ─────────────────────────────────────────────────────────────────────────────
TIER_LABELS    = {1: "0-100K", 2: "~500K", 3: "~1M", 4: "~5M", 5: "~10M"}
DATA_START_ROW = TABLE_HEADER_ROW + 1   # = 7

def _col(idx: int) -> int:
    return C1 + idx

def _write_data_rows(ws, line_items: list) -> int:
    curr = DATA_START_ROW

    for item in sorted(line_items, key=lambda x: x.sku_name):
        n = len(item.tier_breakdown)
        billable_disp = item.billable_usage if item.billable_usage > 0 else "-"

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
            # n==0: SKU에 정의된 티어 없음 (엣지케이스) — 단일 행
            _cell_write(ws, curr, _col(0), item.sku_name,  h="center", wrap=True)
            _cell_write(ws, curr, _col(1), item.total_usage, h="right", num="#,##0")
            _cell_write(ws, curr, _col(2), f"-{item.free_usage_cap:,}", h="center")
            _cell_write(ws, curr, _col(3), billable_disp,
                        h="right" if billable_disp != "-" else "center",
                        num="#,##0" if billable_disp != "-" else "General")
            _cell_write(ws, curr, _col(4), "-", h="center")
            _cell_write(ws, curr, _col(5), "-", h="center")
            _cell_write(ws, curr, _col(6), "-", h="center")
            _cell_write(ws, curr, _col(7), float(item.subtotal_usd),
                        h="right", num="$#,##0.00")
            curr += 1

        for i, tb in enumerate(item.tier_breakdown):
            label = TIER_LABELS.get(tb.tier_number, f"T{tb.tier_number}")
            usage = tb.usage_in_tier if tb.usage_in_tier > 0 else "-"
            _cell_write(ws, curr + i, _col(4), label, h="center")
            _cell_write(ws, curr + i, _col(5), usage,
                        h="right" if usage != "-" else "center",
                        num="#,##0" if usage != "-" else "General")
            _cell_write(ws, curr + i, _col(6), float(tb.tier_cpm),
                        h="right", num="$#,##0.00")
            _cell_write(ws, curr + i, _col(7), float(tb.amount_usd),
                        h="right", num="$#,##0.00")

        curr += n

        # ── 소계 행: 병합 전 앵커(_col(0)) 포함 전체 범위 테두리 먼저 ────────
        for c in range(_col(0), _col(4) + 1):
            ws.cell(row=curr, column=c).fill   = _fill(C_SUB)
            ws.cell(row=curr, column=c).border = _BORDER_ALL

        ws.merge_cells(
            start_row=curr, start_column=_col(0),
            end_row=curr,   end_column=_col(4),
        )
        cell = ws.cell(row=curr, column=_col(0), value="소계")
        cell.font      = _font(bold=True)
        cell.fill      = _fill(C_SUB)
        cell.alignment = _align("center", "center")
        cell.border    = _BORDER_ALL

        # 소계 = 실제 렌더된 tier rows의 합 (engine subtotal_usd와 불일치 방지)
        tier_sum = sum(tb.amount_usd for tb in item.tier_breakdown)
        sub_usage = item.billable_usage if item.billable_usage > 0 else "-"
        _cell_write(ws, curr, _col(5), sub_usage,
                    h="right" if sub_usage != "-" else "center",
                    num="#,##0" if sub_usage != "-" else "General",
                    bold=True, bg=C_SUB)
        _cell_write(ws, curr, _col(6), "",   h="center", bold=True, bg=C_SUB)
        _cell_write(ws, curr, _col(7), float(tier_sum),
                    h="right", num="$#,##0.00", bold=True, bg=C_SUB)
        curr += 1

    return curr - 1


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 합계 / 환율 / 청구금액 행
# ─────────────────────────────────────────────────────────────────────────────
def _write_summary_rows(ws, line_items, exchange_rate, margin_rate,
                        last_data_row: int,
                        billing_month: str = "2026-01",
                        bank_name: str = "하나은행") -> int:
    # 합계 = 렌더된 tier rows의 합 (소계와 일관성 유지)
    t_usd = sum(
        (sum(tb.amount_usd for tb in item.tier_breakdown)
         .quantize(Decimal("1"), ROUND_HALF_UP)
         for item in line_items),
        Decimal("0"),
    )
    t_krw = (t_usd * exchange_rate * margin_rate
             ).quantize(Decimal("1"), ROUND_HALF_UP)

    year, month = int(billing_month[:4]), int(billing_month[5:7])
    last_day     = calendar.monthrange(year, month)[1]
    last_day_str = f"{year}.{month:02d}.{last_day:02d}"

    r = last_data_row + 1

    # ── 합계(USD) — 병합 전 앵커(_col(0)) 포함 전체 범위 테두리 먼저 ────────
    for c in range(_col(0), _col(6) + 1):
        ws.cell(row=r, column=c).fill   = _fill(C_WHITE)
        ws.cell(row=r, column=c).border = _BORDER_ALL

    ws.merge_cells(start_row=r, start_column=_col(0),
                   end_row=r,   end_column=_col(6))
    cell = ws.cell(row=r, column=_col(0), value="합        계(USD)")
    cell.font      = _font(bold=True)
    cell.fill      = _fill(C_WHITE)
    cell.alignment = _align("left", "center")
    cell.border    = _BORDER_ALL
    _cell_write(ws, r, _col(7), float(t_usd),
                h="right", num="$#,##0", bold=True)

    # ── 환율 — 병합 전 앵커(_col(0)) 포함 전체 범위 테두리 먼저 ────────────
    r += 1
    for c in range(_col(0), _col(6) + 1):
        ws.cell(row=r, column=c).fill   = _fill(C_WHITE)
        ws.cell(row=r, column=c).border = _BORDER_ALL

    ws.merge_cells(start_row=r, start_column=_col(0),
                   end_row=r,   end_column=_col(6))
    cell = ws.cell(row=r, column=_col(0),
                   value=f"환        율({bank_name} {last_day_str} 최종 송금환율 기준)")
    cell.font      = _font(bold=True)
    cell.fill      = _fill(C_WHITE)
    cell.alignment = _align("left", "center")
    cell.border    = _BORDER_ALL
    _cell_write(ws, r, _col(7), float(exchange_rate),
                h="right", num="₩#,##0.00", bold=True)

    # ── 청구금액(KRW) — 병합 전 앵커(_col(0)) 포함 전체 범위 테두리 먼저 ───
    r += 1
    for c in range(_col(0), _col(6) + 1):
        ws.cell(row=r, column=c).fill   = _fill(C_DARK)
        ws.cell(row=r, column=c).border = _BORDER_ALL

    ws.merge_cells(start_row=r, start_column=_col(0),
                   end_row=r,   end_column=_col(6))
    cell = ws.cell(row=r, column=_col(0), value="청 구 금 액(KRW)")
    cell.font      = _font(bold=True, color=C_WHITE)
    cell.fill      = _fill(C_DARK)
    cell.alignment = _align("center", "center")
    cell.border    = _BORDER_ALL

    cell8 = ws.cell(row=r, column=_col(7), value=int(t_krw))
    cell8.font          = _font(bold=True, color=C_WHITE)
    cell8.fill          = _fill(C_DARK)
    cell8.alignment     = _align("right", "center")
    cell8.number_format = "₩#,##0"
    cell8.border        = _BORDER_ALL

    return r


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: (부가세 별도)
# ─────────────────────────────────────────────────────────────────────────────
def _write_vat_note(ws, krw_row: int) -> None:
    r = krw_row + 1
    ws.row_dimensions[r].height = 16

    # 병합 전 비앵커 셀 스타일 먼저
    for c in range(C1 + 1, CN + 1):
        ws.cell(row=r, column=c).fill   = _fill(C_WHITE)
        ws.cell(row=r, column=c).border = Border()

    ws.merge_cells(start_row=r, start_column=C1, end_row=r, end_column=CN)
    cell = ws.cell(row=r, column=C1, value="(부가세 별도)")
    cell.font      = _font(size=9, color="595959")
    cell.fill      = _fill(C_WHITE)
    cell.alignment = _align("right", "center")
    cell.border    = Border()


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
        ws.row_dimensions[start_row].height = int(img3_h * 0.75) + 10

        col_widths_px   = [10*7, 32*7, 14*7, 14*7, 14*7, 14*7, 14*7, 14*7, 18*7]
        table_right_px  = sum(col_widths_px)
        img_left_px     = table_right_px - target_w

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

        # 병합 전 비앵커 셀 스타일 먼저
        for c in range(C1 + 1, CN + 1):
            ws.cell(row=start_row, column=c).fill   = _fill(C_WHITE)
            ws.cell(row=start_row, column=c).border = Border()

        ws.merge_cells(
            start_row=start_row, start_column=C1,
            end_row=start_row,   end_column=CN,
        )
        cell = ws.cell(row=start_row, column=C1,
                       value="[tt3: Google Cloud Premier Partner 배지]")
        cell.font      = _font(size=9, color="888888")
        cell.fill      = _fill(C_WHITE)
        cell.alignment = _align("right", "center")
        cell.border    = Border()


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: 틀 고정 해제
# ─────────────────────────────────────────────────────────────────────────────
def _set_freeze_pane(ws) -> None:
    ws.freeze_panes = None


# ─────────────────────────────────────────────────────────────────────────────
# 저수준 셀 작성 헬퍼
# ─────────────────────────────────────────────────────────────────────────────
def _cell_write(ws, row: int, col: int, value,
                h="center", v="center", wrap=False,
                num="General", bold=False,
                color="000000", bg=C_WHITE) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font          = _font(bold=bold, color=color)
    cell.fill          = _fill(bg)
    cell.alignment     = _align(h, v, wrap)
    cell.number_format = num
    cell.border        = _BORDER_ALL
    ws.row_dimensions[row].height = 18


def _merge_write(ws, r1: int, r2: int, col: int, value,
                 fmt_data="center", wrap=False,
                 bold=False, color="000000", bg=C_WHITE) -> None:
    """단일 열 세로 병합 (r1~r2, col 고정).
    병합 전에 범위 내 모든 셀(r1~r2 전체)에 테두리를 먼저 입혀 왼쪽 선 유실 방지."""
    h   = "right" if fmt_data == "number" else "center"
    num = "#,##0" if fmt_data == "number" else "General"

    # ① 병합 전: r1 포함 전체 범위(r1~r2)에 스타일 먼저 적용
    for r in range(r1, r2 + 1):
        cell = ws.cell(row=r, column=col)
        cell.fill   = _fill(bg)
        cell.font   = _font(bold=bold, color=color)
        cell.border = _BORDER_ALL

    # ② 병합 실행
    ws.merge_cells(start_row=r1, start_column=col,
                   end_row=r2,   end_column=col)

    # ③ 앵커 셀(r1)에 값·정렬 추가 — 테두리는 ①에서 확정됨
    cell = ws.cell(row=r1, column=col, value=value)
    cell.font          = _font(bold=bold, color=color)
    cell.fill          = _fill(bg)
    cell.alignment     = _align(h, "center", wrap)
    cell.number_format = num
    cell.border        = _BORDER_ALL


# ─────────────────────────────────────────────────────────────────────────────
# 내부 헬퍼: GMP Price List 시트 복제 (MergedCell 완전 방어)
# ─────────────────────────────────────────────────────────────────────────────
def _copy_price_list_sheet(wb: Workbook, price_list_file) -> None:
    """원본 엑셀 첫 번째 시트를 'GMP Price List' 탭으로 복제."""
    try:
        if hasattr(price_list_file, "read"):
            price_list_file.seek(0)
            raw = price_list_file.read()
            src_wb = load_workbook(io.BytesIO(raw), data_only=True)
        else:
            src_wb = load_workbook(str(price_list_file), data_only=True)
    except Exception:
        return

    src_ws = src_wb.worksheets[0]
    dst_ws = wb.create_sheet("GMP Price List")
    dst_ws.sheet_view.showGridLines = False

    # 탭 색상
    if src_ws.sheet_properties.tabColor:
        dst_ws.sheet_properties.tabColor = copy(src_ws.sheet_properties.tabColor)

    from openpyxl.utils import get_column_letter

    # 시트 기본값
    default_col_w = src_ws.sheet_format.defaultColWidth or 8.43
    if src_ws.sheet_format.defaultColWidth:
        dst_ws.sheet_format.defaultColWidth = src_ws.sheet_format.defaultColWidth
    if src_ws.sheet_format.defaultRowHeight:
        dst_ws.sheet_format.defaultRowHeight = src_ws.sheet_format.defaultRowHeight

    # 명시적 열 너비 복사
    explicit_cols = set()
    for col_letter, col_dim in src_ws.column_dimensions.items():
        dst_col = dst_ws.column_dimensions[col_letter]
        if col_dim.width and col_dim.width > 0:
            dst_col.width = col_dim.width * 1.2
            explicit_cols.add(col_letter)
        dst_col.hidden = col_dim.hidden

    # column_dimensions에 없는 열(기본 너비 사용) → 셀 내용 길이로 자동 산정
    auto_w: dict[str, int] = {}
    for row in src_ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell) or cell.value is None:
                continue
            cl = get_column_letter(cell.column)
            if cl not in explicit_cols:
                auto_w[cl] = max(auto_w.get(cl, 0), len(str(cell.value)))
    for cl, w in auto_w.items():
        dst_ws.column_dimensions[cl].width = max((w + 4) * 1.3, default_col_w)

    # I열 추가 확장
    if dst_ws.column_dimensions["I"].width:
        dst_ws.column_dimensions["I"].width *= 1.25

    # 행 높이
    for row_idx, row_dim in src_ws.row_dimensions.items():
        dst_row = dst_ws.row_dimensions[row_idx]
        dst_row.height = row_dim.height
        dst_row.hidden = row_dim.hidden

    # 병합 범위 선이식
    for merged_range in list(src_ws.merged_cells.ranges):
        try:
            dst_ws.merge_cells(str(merged_range))
        except Exception:
            pass

    # 셀 값 + 스타일 복사
    for row in src_ws.iter_rows():
        for src_cell in row:
            if isinstance(src_cell, MergedCell):
                continue

            dst_cell = dst_ws.cell(row=src_cell.row, column=src_cell.column)

            if isinstance(dst_cell, MergedCell):
                continue

            dst_cell.value = src_cell.value

            if src_cell.has_style:
                dst_cell.font          = copy(src_cell.font)
                dst_cell.fill          = copy(src_cell.fill)
                dst_cell.border        = copy(src_cell.border)
                dst_cell.alignment     = copy(src_cell.alignment)
                dst_cell.number_format = src_cell.number_format
                if src_cell.protection:
                    dst_cell.protection = copy(src_cell.protection)

    # "COST PER THOUSAND (CPM)" 셀 border 누락 보정
    _thin = Side(style="thin")
    _border_fix = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
    for row in dst_ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            if cell.value and "COST PER THOUSAND" in str(cell.value).upper():
                cell.border = _border_fix


# ─────────────────────────────────────────────────────────────────────────────
# CLI 진입점
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="t1 레이아웃 인보이스 생성기",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument("input_file", nargs="?", default="billing.csv",
                        help="사용고지서 파일 (기본: billing.csv)")
    parser.add_argument("-m", "--billing-month", default="2026-03", metavar="YYYY-MM")
    parser.add_argument("-e", "--exchange-rate", type=float, default=1427.87, metavar="RATE")
    parser.add_argument("-r", "--margin-rate",   type=float, default=1.12,    metavar="RATE")
    parser.add_argument("-c", "--company",        default=None,
                        help="특정 회사만 필터링 (예: coupang)")
    parser.add_argument("-o", "--output",         default="invoice_output.xlsx", metavar="FILE")
    parser.add_argument("--invoice-date",         default=None,
                        help="인보이스 날짜 YYYY-MM-DD (기본: 오늘)")
    args = parser.parse_args()

    INPUT_FILE  = Path(args.input_file)
    OUTPUT_FILE = BASE_DIR / args.output
    EXCHANGE    = Decimal(str(args.exchange_rate))
    MARGIN      = Decimal(str(args.margin_rate))

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
        line_items    = line_items,
        company_name  = company_display,
        billing_month = args.billing_month,
        exchange_rate = EXCHANGE,
        margin_rate   = MARGIN,
        invoice_date  = args.invoice_date,
        output_path   = OUTPUT_FILE,
    )

    print(f"\n완료 -> {OUTPUT_FILE}")
    print(f"  항목 수: {len(line_items)}건")
    for it in line_items:
        print(f"  - {it.sku_name}: ${float(it.subtotal_usd):,.2f}")
