print("🔥 [DEBUG] 지금 project_sheet.py 최신 버전을 읽고 있습니다!")
"""
project_sheet.py — Project 시트 생성기 (API별 3컬럼 × 4행 격자 구조)

블록 구조 (프로젝트 1개 = 6행):
  Col A (LABEL_COL): 빈 여백
  Col B (PROJ_COL) : 프로젝트명 (세로병합 4행, WHITE)
  Col C+           : API별 데이터 (1 API당 3컬럼)

  3컬럼 구조 (per API):
    api_left  : C (price/label 좌측)
    api_mid   : D (price/label 우측, left와 병합)
    api_right : E (amount 단독 컬럼)

  6행 구조:
    행1: 서비스명  (3열 C:E 병합, #E2EFDA, 굵게)
    행2: 사용량    (3열 C:E 병합, #E2EFDA, 우측 정렬)
    행3: 라벨      (C:D 병합=monthly unit price($) | E=amount, WHITE)
    행4: 값        (C:D 병합=단가 | E=KRW 금액, WHITE)
    행5: toal($)   라벨(B:D 3칸 병합) + 값(E열)
    행6: toal(₩)  라벨(B:D 3칸 병합) + 값(E열), #E2EFDA
"""

import re
import calendar as _cal
from datetime import date as _date
from decimal import Decimal, ROUND_HALF_UP
import os

from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    _HAS_IMAGE = True
except ImportError:
    _HAS_IMAGE = False

# cm → EMU 변환 상수 (1 cm = 360000 EMU)
_CM_TO_EMU = 360000


# ─────────────────────────────────────────────────────────────────────────────
# API 컬럼 표시 순서
# ─────────────────────────────────────────────────────────────────────────────
PREFERRED_API_ORDER = [
    "Dynamic Maps",
    "Geocoding",
    "Autocomplete - Per Request",
    "Autocomplete without Places Details - Per Session",
    "Query Autocomplete - Per Request",
    "Places Details",
    "Basic Data",
    "Contact Data",
    "Atmosphere Data",
    "Places - Text Search",
    "Find Place",
    "Directions",
]

API_DISPLAY_NAMES: dict[str, str] = {
    "Dynamic Maps":                                      "Dynamic Maps",
    "Basic Data":                                        "Basic Data",
    "Contact Data":                                      "Contact Data",
    "Atmosphere Data":                                   "Atmosphere\nData",
    "Find Place":                                        "Find Place",
    "Geocoding":                                         "Geocoding",
    "Autocomplete - Per Request":                        "Autocomplete\n/Pr Request",
    "Autocomplete without Places Details - Per Session": "Autocomplete\nw/o Places Details\nPer Session",
    "Places Details":                                    "Places Details",
    "Places - Text Search":                              "Places\nText Search",
    "Query Autocomplete - Per Request":                  "Query Autocomplete\nPer Request",
    "Directions":                                        "Directions",
}

# ─────────────────────────────────────────────────────────────────────────────
# FORCED unit prices (override engine tier_cpm and weighted_unit_prices)
# Derived from Invoice sheet: subtotal_usd / total_usage
#   Places - Text Search: Invoice!$I$75 / Invoice!$C$70 = 0.01558
#     Verification: round(3     * 0.01558) = 0
#                   round(9556  * 0.01558) = 149
#                   round(186   * 0.01558) = 3
# ─────────────────────────────────────────────────────────────────────────────
FORCED_UNIT_PRICES: dict[str, float] = {
    "Dynamic Maps":         0.002,     # 프로젝트 시트 표시 단가 (역산 기준)
    "Places - Text Search": 0.01558,   # Invoice!$I$75 / Invoice!$C$70
}

# ─────────────────────────────────────────────────────────────────────────────
# 색상 팔레트
# ─────────────────────────────────────────────────────────────────────────────
C_HEADER  = "E2EFDA"
C_SUB     = "F2F2F2"
C_WHITE   = "FFFFFF"
C_ORANGE  = "FF8F00"
C_DIVIDER = "E0E0E0"
C_TEXT    = "000000"
C_TOTAL   = "C5E0B3"   # toal 행 배경색 (#C5E0B3)

_BLK       = Side(style="thin", color="000000")
_BORDER    = Border(left=_BLK, right=_BLK, top=_BLK, bottom=_BLK)
_NO_BORDER = Border()


# ─────────────────────────────────────────────────────────────────────────────
# 스타일 헬퍼
# ─────────────────────────────────────────────────────────────────────────────
def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _font(color: str = C_TEXT, bold: bool = False, size: int = 8) -> Font:
    return Font(color=color, bold=bold, size=size, name="맑은 고딕")


def _align(h: str = "center", v: str = "center", wrap: bool = False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _pre_merge_style(ws, r1, c1, r2, c2, fill=None, border=None):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            if fill   is not None: cell.fill   = fill
            if border is not None: cell.border = border


def _set(ws, r, c, value=None, fill=None, font=None,
         alignment=None, border=None, number_format=None):
    cell = ws.cell(row=r, column=c)
    if isinstance(cell, MergedCell):
        return
    if value         is not None: cell.value         = value
    if fill          is not None: cell.fill          = fill
    if font          is not None: cell.font          = font
    if alignment     is not None: cell.alignment     = alignment
    if border        is not None: cell.border        = border
    if number_format is not None: cell.number_format = number_format


def _add_image(ws, img_dir, fname, col_0idx, row_0idx,
               width_cm, height_cm, offset_cm=0.1):
    """
    col_0idx, row_0idx: 0-based 열/행 인덱스 (openpyxl AnchorMarker 기준)
    width_cm / height_cm: 이미지 크기 (cm 단위)
    offset_cm: 상단 여백 (cm 단위)
    """
    if not _HAS_IMAGE:
        return
    path = os.path.join(img_dir, fname)
    if not os.path.exists(path):
        return
    try:
        img    = XLImage(path)
        w_emu  = int(width_cm  * _CM_TO_EMU)
        h_emu  = int(height_cm * _CM_TO_EMU)
        off_emu = int(offset_cm * _CM_TO_EMU)
        marker = AnchorMarker(col=col_0idx, colOff=0,
                              row=row_0idx, rowOff=off_emu)
        size   = XDRPositiveSize2D(cx=w_emu, cy=h_emu)
        img.anchor = OneCellAnchor(_from=marker, ext=size)
        ws.add_image(img)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# 공개 함수
# ─────────────────────────────────────────────────────────────────────────────
def write_project_sheet(
    wb: Workbook,
    proj_results: list,
    company_name: str,
    billing_month: str,
    exchange_rate: Decimal,
    _margin_rate: Decimal,
    invoice_date: str | None,
    bank_name: str,
    line_items: list | None = None,
) -> None:
    ws = wb.create_sheet("Project")
    ws.sheet_view.showGridLines = False

    # ── 단가 표시용 weighted_unit_prices (Invoice 전체 집계 기준) ──────────────
    # amount 셀은 sd["subtotal_usd"]를 직접 사용 — 여기서는 단가 표시 전용
    weighted_unit_prices: dict[str, float] = {}

    if line_items:
        for _item in line_items:
            _tu = int(_item.total_usage or 0)
            _ts = _item.subtotal_usd or Decimal("0")
            if _tu > 0 and float(_ts) > 0:
                weighted_unit_prices[_item.sku_name] = float(_ts) / _tu
    else:
        _sku_sum_usd:   dict[str, Decimal] = {}
        _sku_sum_usage: dict[str, int]     = {}
        for _pr in proj_results:
            for _sname, _sd in _pr["skus"].items():
                _u = int(_sd.get("usage", 0) or 0)
                _s = _sd.get("subtotal_usd") or Decimal("0")
                _sku_sum_usage[_sname] = _sku_sum_usage.get(_sname, 0) + _u
                _sku_sum_usd[_sname]   = _sku_sum_usd.get(_sname, Decimal("0")) + _s
        for _sname, _ts in _sku_sum_usd.items():
            _tu = _sku_sum_usage.get(_sname, 0)
            if _tu > 0 and float(_ts) > 0:
                weighted_unit_prices[_sname] = float(_ts) / float(_tu)

    # ── API 목록 수집 (세금/Tax/VAT SKU 완전 제외) ───────────────────────────
    _TAX_KEYWORDS = ("세금", "tax", "vat")

    def _is_tax_sku(name: str) -> bool:
        n = name.lower()
        return any(kw in n for kw in _TAX_KEYWORDS)

    seen, api_list = set(), []
    for pr in proj_results:
        for sku_name, sd in pr["skus"].items():
            if sku_name not in seen and not _is_tax_sku(sku_name):
                seen.add(sku_name)
                api_list.append(sku_name)
    if not api_list:
        api_list = [
            n for n in (proj_results[0]["skus"].keys() if proj_results else [])
            if not _is_tax_sku(n)
        ]

    order_map = {name: i for i, name in enumerate(PREFERRED_API_ORDER)}
    api_list.sort(key=lambda n: order_map.get(n, len(PREFERRED_API_ORDER)))

    NUM_APIS = len(api_list)

    # ── 열 인덱스 정의 (3컬럼 per API) ──────────────────────────────────────
    LABEL_COL = 1   # A: 빈 여백
    PROJ_COL  = 2   # B: 프로젝트명

    def api_left(k):  return 3 + 3 * k        # 가격 라벨 좌측
    def api_mid(k):   return 3 + 3 * k + 1    # 가격 라벨 우측 (left와 병합)
    def api_right(k): return 3 + 3 * k + 2    # amount 단독 컬럼

    MAX_COL = 2 + 3 * NUM_APIS   # = api_right(NUM_APIS - 1)

    # ── 열 너비 ───────────────────────────────────────────────────────────────
    # (left_w, mid_w, right_w) per API
    _COL_WIDTHS: dict[str, tuple] = {
        "Dynamic Maps":                                      (9, 5, 12),
        "Basic Data":                                        (9, 5, 12),
        "Contact Data":                                      (9, 5, 12),
        "Atmosphere Data":                                   (9, 5, 12),
        "Find Place":                                        (9, 5, 12),
        "Geocoding":                                         (9, 5, 12),
        "Autocomplete - Per Request":                        (9, 5, 12),
        "Autocomplete without Places Details - Per Session": (9, 6, 12),
        "Places Details":                                    (9, 5, 12),
        "Places - Text Search":                              (9, 5, 12),
        "Query Autocomplete - Per Request":                  (9, 5, 12),
        "Directions":                                        (9, 5, 12),
    }
    ws.column_dimensions[get_column_letter(LABEL_COL)].width = 2
    ws.column_dimensions[get_column_letter(PROJ_COL)].width  = 16
    for k, name in enumerate(api_list):
        lw, mw, rw = _COL_WIDTHS.get(name, (9, 5, 12))
        ws.column_dimensions[get_column_letter(api_left(k))].width  = lw
        ws.column_dimensions[get_column_letter(api_mid(k))].width   = mw
        ws.column_dimensions[get_column_letter(api_right(k))].width = rw

    # ── 전체 흰색 배경 초기화 ─────────────────────────────────────────────────
    wf = _fill(C_WHITE)
    for _row in ws.iter_rows(min_row=1, max_row=600, min_col=1, max_col=MAX_COL + 2):
        for _cell in _row:
            _cell.fill = wf

    def rh(r, h): ws.row_dimensions[r].height = h

    # ── 날짜 계산 ─────────────────────────────────────────────────────────────
    if invoice_date is None:
        invoice_date = _date.today().strftime("%Y-%m-%d")
    year, month = int(billing_month[:4]), int(billing_month[5:7])
    last_day    = _cal.monthrange(year, month)[1]
    term_str    = f"{billing_month}-01 ~ {billing_month}-{last_day:02d}"

    # ══════════════════════════════════════════════════════════════════════════
    # 상단 이미지 헤더 (Row 1~3)
    # ══════════════════════════════════════════════════════════════════════════
    rh(1, 50)
    rh(2, 0.75)   # 로고와 송장 정보 사이 간격 제거
    rh(3, 0.75)
    rh(4, 0.75)

    for r in range(1, 5):
        for c in range(1, MAX_COL + 1):
            ws.cell(row=r, column=c).fill   = _fill(C_WHITE)
            ws.cell(row=r, column=c).border = _NO_BORDER

    _img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")

    # tt1.png: B1 (col_0idx=1, row_0idx=0), 8.0cm × 1.2cm, offset 0.1cm
    _add_image(ws, _img_dir, "tt1.png",
               col_0idx=1, row_0idx=0,
               width_cm=8.0, height_cm=1.2, offset_cm=0.1)

    # tt2.png: F1 (col_0idx=5, row_0idx=0), 2.83cm × 1.29cm, offset 0.1cm
    _add_image(ws, _img_dir, "tt2.png",
               col_0idx=5, row_0idx=0,
               width_cm=2.83, height_cm=1.29, offset_cm=0.1)

    # ══════════════════════════════════════════════════════════════════════════
    # 송장 요약 정보 (Row 5~8, B열 시작)
    # ══════════════════════════════════════════════════════════════════════════
    meta_rows = [
        ("Invoice Date",         f": {invoice_date}"),
        ("Billing Account Name", f": {company_name.title()}"),
        ("Term of Use",          f": {term_str}"),
        ("환율",
         f"\u20a9{float(exchange_rate):,.2f}"
         f"  ({bank_name} {year}.{month:02d}.{last_day:02d} 최종 송금환율 기준)"),
    ]
    for i, (label, val) in enumerate(meta_rows):
        r = 5 + i
        rh(r, 17)
        ws.cell(row=r, column=LABEL_COL).fill   = _fill(C_WHITE)
        ws.cell(row=r, column=LABEL_COL).border = _NO_BORDER

        cell = ws.cell(row=r, column=PROJ_COL, value=label)
        cell.fill      = _fill(C_WHITE)
        cell.border    = _NO_BORDER
        cell.font      = Font(color="555555", bold=True, size=9, name="맑은 고딕")
        cell.alignment = _align("left", "center")

        _pre_merge_style(ws, r, PROJ_COL + 1, r, MAX_COL,
                         fill=_fill(C_WHITE), border=_NO_BORDER)
        ws.merge_cells(start_row=r, start_column=PROJ_COL + 1,
                       end_row=r,   end_column=MAX_COL)
        cell = ws.cell(row=r, column=PROJ_COL + 1, value=val)
        cell.fill      = _fill(C_WHITE)
        cell.border    = _NO_BORDER
        cell.font      = Font(color="111111", bold=False, size=9, name="맑은 고딕")
        cell.alignment = _align("left", "center")

    # Row 9: 여백
    rh(9, 8)

    # ══════════════════════════════════════════════════════════════════════════
    # 프로젝트 블록 (6행 per project, Row 10부터 시작)
    # ══════════════════════════════════════════════════════════════════════════
    cur       = 10
    grand_usd = Decimal("0")

    hdr_fill  = _fill(C_HEADER)
    data_fill = _fill(C_WHITE)
    tot_fill  = _fill(C_TOTAL)   # toal 행 공통 배경 (#C5E0B3)

    # 프로젝트명 번호 순 정렬 (coupang-01 → coupang-02 ...), 번호 없는 프로젝트(Butter 등)는 맨 뒤
    def _proj_sort_key(p):
        name = p["proj_name"]
        m = re.search(r"-(\d+)", name)
        return (0, name) if m else (1, name)

    proj_results = sorted(proj_results, key=_proj_sort_key)

    for pr in proj_results:
        r_name   = cur       # 행1: 서비스명 (녹색)
        r_usage  = cur + 1   # 행2: 사용량   (녹색)
        r_labels = cur + 2   # 행3: 라벨     (WHITE)
        r_vals   = cur + 3   # 행4: 값       (WHITE)
        r_usd    = cur + 4   # 행5: toal($)
        r_krw    = cur + 5   # 행6: toal(₩)

        proj_total_usd = Decimal("0")   # sum of round(value, 0) per-API amounts

        for r, h in zip(
            [r_name, r_usage, r_labels, r_vals, r_usd, r_krw],
            [22,     18,      14,       18,     14,    18],
        ):
            rh(r, h)

        # ── Col A: 빈 여백 ────────────────────────────────────────────────────
        for r in range(r_name, r_krw + 1):
            ws.cell(row=r, column=LABEL_COL).fill   = _fill(C_WHITE)
            ws.cell(row=r, column=LABEL_COL).border = _NO_BORDER

        # ── Col B: 프로젝트명 (4행 세로 병합, WHITE) ─────────────────────────
        display_name = re.sub(
            r'\s*-\s*project\s*$', '', pr["proj_name"], flags=re.IGNORECASE
        ).strip()

        _pre_merge_style(ws, r_name, PROJ_COL, r_vals, PROJ_COL,
                         fill=data_fill, border=_BORDER)
        try:
            ws.merge_cells(start_row=r_name, start_column=PROJ_COL,
                           end_row=r_vals,   end_column=PROJ_COL)
        except Exception:
            pass
        cell = ws.cell(row=r_name, column=PROJ_COL, value=display_name)
        cell.fill      = data_fill
        cell.font      = _font(C_TEXT, bold=True, size=9)
        cell.alignment = _align("center", "center", wrap=True)
        cell.border    = _BORDER

        # ── API별 4행 격자 (3컬럼 구조) ──────────────────────────────────────
        for k, api_name in enumerate(api_list):
            sd = pr["skus"].get(
                api_name,
                {"usage": 0, "subtotal_usd": Decimal("0"), "final_krw": Decimal("0")},
            )
            usage   = sd["usage"]

            # 단가 우선순위:
            #   1. FORCED_UNIT_PRICES (역산된 실제 단가, 엔진 tier_cpm 완전 무시)
            #   2. weighted_unit_prices (Invoice subtotal_usd / total_usage)
            #   3. None  ← 엔진 sd["unit_price"]는 tier_cpm 기반이므로 사용 금지
            unit_price_usd = (
                FORCED_UNIT_PRICES.get(api_name)        # e.g. Places - Text Search → 0.01558
                or weighted_unit_prices.get(api_name)   # Invoice 기준 가중 단가
                or None                                 # 없으면 0 표시
            )

            lc = api_left(k)
            mc = api_mid(k)
            rc = api_right(k)

            # 행1: 서비스명 (C:E 3열 병합, 녹색, 굵게)
            short = API_DISPLAY_NAMES.get(api_name) or (
                api_name.replace(" - ", "\n").replace(" without ", "\nw/o ")
                if len(api_name) > 14 else api_name
            )
            _pre_merge_style(ws, r_name, lc, r_name, rc,
                             fill=hdr_fill, border=_BORDER)
            try:
                ws.merge_cells(start_row=r_name, start_column=lc,
                               end_row=r_name,   end_column=rc)
            except Exception:
                pass
            cell = ws.cell(row=r_name, column=lc, value=short)
            cell.fill      = hdr_fill
            cell.font      = _font(C_TEXT, bold=True, size=8)
            cell.alignment = _align("center", "center", wrap=True)
            cell.border    = _BORDER

            # 행2: 사용량 (C:E 3열 병합, 녹색, 우측 정렬)
            _pre_merge_style(ws, r_usage, lc, r_usage, rc,
                             fill=hdr_fill, border=_BORDER)
            try:
                ws.merge_cells(start_row=r_usage, start_column=lc,
                               end_row=r_usage,   end_column=rc)
            except Exception:
                pass
            cell = ws.cell(row=r_usage, column=lc,
                           value=usage)
            cell.fill          = hdr_fill
            cell.font          = _font(C_TEXT, size=8)
            cell.alignment     = _align("right", "center")
            cell.border        = _BORDER
            cell.number_format = '#,##0;;"-"'

            # 행3: 라벨 (C:D 병합=monthly unit price($), E=amount, WHITE)
            _pre_merge_style(ws, r_labels, lc, r_labels, mc,
                             fill=data_fill, border=_BORDER)
            try:
                ws.merge_cells(start_row=r_labels, start_column=lc,
                               end_row=r_labels,   end_column=mc)
            except Exception:
                pass
            cell = ws.cell(row=r_labels, column=lc,
                           value="monthly unit price($)")
            cell.fill      = data_fill
            cell.font      = _font(C_TEXT, size=7)
            cell.alignment = _align("center", "center", wrap=True)
            cell.border    = _BORDER

            _set(ws, r_labels, rc,
                 value="amount",
                 fill=data_fill,
                 font=_font(C_TEXT, size=7),
                 alignment=_align("center", "center"),
                 border=_BORDER)

            # 행4: 값 (C:D 병합=단가, E=KRW 금액, WHITE)
            _pre_merge_style(ws, r_vals, lc, r_vals, mc,
                             fill=data_fill, border=_BORDER)
            try:
                ws.merge_cells(start_row=r_vals, start_column=lc,
                               end_row=r_vals,   end_column=mc)
            except Exception:
                pass
            # 단가: None이면 0으로 변환해 format #,##0.000;;"-" 가 "-" 를 표시하도록 함
            cell = ws.cell(row=r_vals, column=lc,
                           value=unit_price_usd if unit_price_usd is not None else 0)
            cell.fill          = data_fill
            cell.font          = _font(C_TEXT, size=8)
            cell.alignment     = _align("right", "center")
            cell.border        = _BORDER
            cell.number_format = '#,##0.000;;"-"'

            # ── amount: 엔진이 계산한 subtotal_usd 를 직접 읽어 정수 달러로 반올림
            # 절대로 usage × price 재계산 금지 — "Printer" 역할만 수행
            _subtotal_usd = Decimal(str(sd.get("subtotal_usd") or 0))
            _amt = int(_subtotal_usd.quantize(Decimal("1"), ROUND_HALF_UP))
            proj_total_usd += Decimal(_amt)

            _set(ws, r_vals, rc,
                 value=_amt,
                 fill=data_fill,
                 font=_font(C_TEXT, size=8),
                 alignment=_align("right", "center"),
                 border=_BORDER,
                 number_format='#,##0;;"-"')

        # ── 행5: toal($) — 라벨 B:D 병합, 값 E열, #C5E0B3 ──────────────────
        # E열 = PROJ_COL+3 = 5 = api_right(0) (첫 번째 API amount 컬럼)
        _pre_merge_style(ws, r_usd, PROJ_COL, r_usd, PROJ_COL + 2,
                         fill=tot_fill, border=_BORDER)
        try:
            ws.merge_cells(start_row=r_usd, start_column=PROJ_COL,
                           end_row=r_usd,   end_column=PROJ_COL + 2)
        except Exception:
            pass
        cell = ws.cell(row=r_usd, column=PROJ_COL, value="toal($)")
        cell.fill      = tot_fill
        cell.font      = _font(C_TEXT, bold=True, size=8)
        cell.alignment = _align("right", "center")
        cell.border    = _BORDER

        grand_usd += proj_total_usd
        _set(ws, r_usd, PROJ_COL + 3,
             value=float(proj_total_usd),
             fill=tot_fill,
             font=_font(C_TEXT, bold=True, size=9),
             alignment=_align("right", "center"),
             border=_BORDER,
             number_format='"$"#,##0')

        # E 이후: 흰색 배경, 테두리 없음
        for c in range(PROJ_COL + 4, MAX_COL + 1):
            _set(ws, r_usd, c, fill=data_fill, border=_NO_BORDER)

        # ── 행6: toal(₩) — 라벨 B:D 병합, 값 E열, #C5E0B3 ──────────────────
        _pre_merge_style(ws, r_krw, PROJ_COL, r_krw, PROJ_COL + 2,
                         fill=tot_fill, border=_BORDER)
        try:
            ws.merge_cells(start_row=r_krw, start_column=PROJ_COL,
                           end_row=r_krw,   end_column=PROJ_COL + 2)
        except Exception:
            pass
        cell = ws.cell(row=r_krw, column=PROJ_COL, value="toal(\u20a9)")
        cell.fill      = tot_fill
        cell.font      = _font(C_TEXT, bold=True, size=8)
        cell.alignment = _align("right", "center")
        cell.border    = _BORDER

        # toal(₩) = ceil'd toal($) × 환율
        _tot_krw_display = int(
            (proj_total_usd * exchange_rate).quantize(Decimal("1"), ROUND_HALF_UP)
        )
        _set(ws, r_krw, PROJ_COL + 3,
             value=_tot_krw_display,
             fill=tot_fill,
             font=_font(C_TEXT, bold=True, size=9),
             alignment=_align("right", "center"),
             border=_BORDER,
             number_format="\u20a9#,##0")

        # E 이후: 흰색 배경, 테두리 없음
        for c in range(PROJ_COL + 4, MAX_COL + 1):
            _set(ws, r_krw, c, fill=data_fill, border=_NO_BORDER)

        cur = r_krw + 1

        # ── 프로젝트 간 구분선 ────────────────────────────────────────────────
        rh(cur, 5)
        for c in range(1, MAX_COL + 1):
            ws.cell(row=cur, column=c).fill = _fill(C_SUB)
        cur += 1

    # ══════════════════════════════════════════════════════════════════════════
    # 최하단 합계 행 (green #C5E0B3, B:E 범위만)
    # ══════════════════════════════════════════════════════════════════════════
    rh(cur, 5)
    cur += 1

    r_total      = cur
    r_vat_notice = cur + 1

    rh(r_total,      22)
    rh(r_vat_notice, 14)

    # A열 및 F열 이후 전체 흰색으로 초기화 (주황색 제거)
    for r in (r_total, r_vat_notice):
        for c in range(1, MAX_COL + 2):
            ws.cell(row=r, column=c).fill   = _fill(C_WHITE)
            ws.cell(row=r, column=c).border = _NO_BORDER

    # B:D 병합 → "청구 대상 금액" 라벨 (#C5E0B3, 검정 굵게)
    _pre_merge_style(ws, r_total, PROJ_COL, r_total, PROJ_COL + 2,
                     fill=tot_fill, border=_BORDER)
    try:
        ws.merge_cells(start_row=r_total, start_column=PROJ_COL,
                       end_row=r_total,   end_column=PROJ_COL + 2)
    except Exception:
        pass
    cell = ws.cell(row=r_total, column=PROJ_COL, value="청구 대상 금액")
    cell.fill      = tot_fill
    cell.font      = Font(color=C_TEXT, bold=True, size=10, name="맑은 고딕")
    cell.alignment = _align("center", "center")
    cell.border    = _BORDER

    # E열 → 최종 원화 합계 = grand_usd × 환율 (빨간색, 굵게, 우측 정렬)
    _grand_krw_display = int(
        (grand_usd * exchange_rate).quantize(Decimal("1"), ROUND_HALF_UP)
    )
    _set(ws, r_total, PROJ_COL + 3,
         value=_grand_krw_display,
         fill=tot_fill,
         font=Font(color="FF0000", bold=True, size=10, name="맑은 고딕"),
         alignment=_align("right", "center"),
         border=_BORDER,
         number_format="\u20a9#,##0")

    # 부가세 안내 — E열 바로 아래 (흰색 배경, 테두리 없음, 우측 정렬, 작은 글씨)
    _set(ws, r_vat_notice, PROJ_COL + 3,
         value="(부가세 별도)",
         fill=_fill(C_WHITE),
         font=Font(color="555555", bold=False, size=7, name="맑은 고딕"),
         alignment=_align("right", "center"),
         border=_NO_BORDER)

    # ── 인쇄 설정 ─────────────────────────────────────────────────────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
