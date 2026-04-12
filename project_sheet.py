print("🔥 [DEBUG] 지금 project_sheet.py 최신 버전을 읽고 있습니다!")
"""
project_sheet.py — Project 시트 생성기 (가로 와이드 그리드)

블록 구조 (프로젝트 1개 = 3행 + USD소계 + 구분선):
  Col A (LABEL_COL): 행 레이블 (사용량 / monthly unit price / amount)
  Col B (PROJ_COL) : 프로젝트명 (세로병합 3행)
  Col C+           : API별 데이터 (1 col per API)
  TOTAL_COL        : 합계(KRW)
"""

import calendar as _cal
from datetime import date as _date
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────────────────────────────────────
# API 컬럼 표시 순서 (detail.png 기준)
# ─────────────────────────────────────────────────────────────────────────────
PREFERRED_API_ORDER = [
    "Dynamic Maps",
    "Basic Data",
    "Contact Data",
    "Atmosphere Data",
    "Find Place",
    "Geocoding",
    "Autocomplete - Per Request",
    "Autocomplete without Places Details - Per Session",
    "Places Details",
    "Places - Text Search",
    "Query Autocomplete - Per Request",
    "Directions",
]

# 컬럼 헤더 표시명 (detail.png 헤더 텍스트와 일치)
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
# 색상 팔레트
# ─────────────────────────────────────────────────────────────────────────────
C_HEADER  = "E2EFDA"   # 헤더 / amount 행 배경 (연한 녹색)
C_SUB     = "F2F2F2"   # monthly unit price 행 배경 (연회색)
C_WHITE   = "FFFFFF"
C_ORANGE  = "FF8F00"   # 합계 대상 금액
C_BLUE    = "1A73E8"
C_GCLOUD  = "4285F4"
C_DIVIDER = "E0E0E0"
C_TEXT    = "000000"

_BLK  = Side(style="thin", color="000000")
_GRAY = Side(style="thin", color="CCCCCC")
_BORDER      = Border(left=_BLK,  right=_BLK,  top=_BLK,  bottom=_BLK)
_BORDER_GRAY = Border(left=_GRAY, right=_GRAY, top=_GRAY, bottom=_GRAY)
_NO_BORDER   = Border()


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
) -> None:
    ws = wb.create_sheet("Project")
    ws.sheet_view.showGridLines = False

    # ── API 목록 수집 (usage > 0 우선, fallback 전체) ─────────────────────────
    seen, api_list = set(), []
    for pr in proj_results:
        for sku_name, sd in pr["skus"].items():
            if sku_name not in seen and sd["usage"] > 0:
                seen.add(sku_name)
                api_list.append(sku_name)
    if not api_list:
        api_list = list(proj_results[0]["skus"].keys()) if proj_results else []

    # PREFERRED_API_ORDER 순서대로 정렬 (미등록 SKU는 뒤에 추가)
    order_map = {name: i for i, name in enumerate(PREFERRED_API_ORDER)}
    api_list.sort(key=lambda n: order_map.get(n, len(PREFERRED_API_ORDER)))

    NUM_APIS = len(api_list)

    # ── 열 인덱스 정의 ────────────────────────────────────────────────────────
    # Col 1 (A): 행 레이블 (사용량 / monthly unit price / amount)
    # Col 2 (B): 프로젝트명 (세로 병합 3행)
    # Col 3+   : API별 데이터 (1 col per API)
    # TOTAL_COL: 합계
    LABEL_COL = 1
    PROJ_COL  = 2
    def api_col(k): return 3 + k
    TOTAL_COL = 3 + NUM_APIS
    MAX_COL   = TOTAL_COL

    # ── 열 너비 ───────────────────────────────────────────────────────────────
    # 컬럼별 맞춤 너비 (detail.png 기준)
    _COL_WIDTHS: dict[str, float] = {
        "Dynamic Maps":                                      13,
        "Basic Data":                                        10,
        "Contact Data":                                      11,
        "Atmosphere Data":                                   11,
        "Find Place":                                        10,
        "Geocoding":                                         10,
        "Autocomplete - Per Request":                        13,
        "Autocomplete without Places Details - Per Session": 16,
        "Places Details":                                    12,
        "Places - Text Search":                              12,
        "Query Autocomplete - Per Request":                  14,
        "Directions":                                        11,
    }
    ws.column_dimensions[get_column_letter(LABEL_COL)].width = 15
    ws.column_dimensions[get_column_letter(PROJ_COL)].width  = 20
    for k, name in enumerate(api_list):
        ws.column_dimensions[get_column_letter(api_col(k))].width = _COL_WIDTHS.get(name, 13)
    ws.column_dimensions[get_column_letter(TOTAL_COL)].width = 18

    # ── 전체 흰색 배경 초기화 ─────────────────────────────────────────────────
    wf = _fill(C_WHITE)
    for _row in ws.iter_rows(min_row=1, max_row=600, min_col=1, max_col=MAX_COL + 2):
        for _cell in _row:
            _cell.fill = wf

    def rh(r, h): ws.row_dimensions[r].height = h

    # ── 날짜/기간 계산 ────────────────────────────────────────────────────────
    if invoice_date is None:
        invoice_date = _date.today().strftime("%Y-%m-%d")
    year, month = int(billing_month[:4]), int(billing_month[5:7])
    last_day    = _cal.monthrange(year, month)[1]
    term_str    = f"{billing_month}-01  ~  {billing_month}-{last_day:02d}"

    # ══════════════════════════════════════════════════════════════════════════
    # 상단 메타 헤더 (Row 1~7)
    # ══════════════════════════════════════════════════════════════════════════
    rh(1, 36)
    _pre_merge_style(ws, 1, 1, 1, 4, fill=_fill(C_WHITE), border=_NO_BORDER)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    cell = ws.cell(row=1, column=1, value="\U0001F4CD  Google Maps Platform")
    cell.fill      = _fill(C_WHITE)
    cell.font      = Font(color="43A047", bold=True, size=13, name="맑은 고딕")
    cell.alignment = _align("left", "center")
    cell.border    = _NO_BORDER

    _pre_merge_style(ws, 1, 5, 1, 7, fill=_fill(C_BLUE), border=_NO_BORDER)
    ws.merge_cells(start_row=1, start_column=5, end_row=1, end_column=7)
    cell = ws.cell(row=1, column=5, value="Invoice")
    cell.fill      = _fill(C_BLUE)
    cell.font      = Font(color="FFFFFF", bold=True, size=11, name="맑은 고딕")
    cell.alignment = _align("center", "center")
    cell.border    = _NO_BORDER

    # Google Cloud 로고 — MergedCell 방어 코드
    try:
        cell = ws.cell(row=1, column=MAX_COL, value="Google Cloud")
        cell.fill      = _fill(C_WHITE)
        cell.font      = Font(color=C_GCLOUD, bold=True, size=11, name="맑은 고딕")
        cell.alignment = _align("right", "center")
        cell.border    = _NO_BORDER
    except Exception:
        pass

    # Row 2: 구분선
    rh(2, 3)
    for c in range(1, MAX_COL + 1):
        ws.cell(row=2, column=c).fill = _fill(C_DIVIDER)

    # Row 3~6: 메타 정보
    meta_rows = [
        ("Invoice Date",         invoice_date),
        ("Billing Account Name", company_name),
        ("Term of Use",          term_str),
        ("환율",
         f"\u20a9{float(exchange_rate):,.2f}"
         f"  ({bank_name} {year}.{month:02d}.{last_day:02d} 최종 송금환율 기준)"),
    ]
    for i, (label, val) in enumerate(meta_rows):
        r = 3 + i
        rh(r, 17)
        cell = ws.cell(row=r, column=1, value=label)
        cell.fill = _fill(C_WHITE); cell.border = _NO_BORDER
        cell.font = Font(color="555555", bold=True, size=9, name="맑은 고딕")
        cell.alignment = _align("left", "center")

        _pre_merge_style(ws, r, 2, r, MAX_COL, fill=_fill(C_WHITE), border=_NO_BORDER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=MAX_COL)
        cell = ws.cell(row=r, column=2, value=val)
        cell.fill = _fill(C_WHITE); cell.border = _NO_BORDER
        cell.font = Font(color="111111", bold=False, size=9, name="맑은 고딕")
        cell.alignment = _align("left", "center")

    rh(7, 8)

    # ══════════════════════════════════════════════════════════════════════════
    # Row 8: 컬럼 헤더 행 (API 서비스명)
    # ══════════════════════════════════════════════════════════════════════════
    cur = 8
    rh(cur, 42)   # 3줄 헤더 텍스트 수용을 위해 높이 확대
    hdr_fill = _fill(C_HEADER)

    # Col A (LABEL_COL): 빈 헤더
    _set(ws, cur, LABEL_COL,
         fill=hdr_fill,
         border=_BORDER)

    # Col B (PROJ_COL): "프로젝트"
    _set(ws, cur, PROJ_COL,
         value="프로젝트",
         fill=hdr_fill,
         font=_font(C_TEXT, bold=True, size=9),
         alignment=_align("center", "center"),
         border=_BORDER)

    # API 이름 헤더 — #E2EFDA, 검정, 굵게, thin 테두리
    for k, api_name in enumerate(api_list):
        # API_DISPLAY_NAMES에 등록된 이름 우선 사용, 없으면 자동 축약
        short = API_DISPLAY_NAMES.get(api_name) or (
            api_name.replace(" - ", "\n").replace(" without ", "\nw/o ")
            if len(api_name) > 14 else api_name
        )
        _set(ws, cur, api_col(k),
             value=short,
             fill=hdr_fill,
             font=_font(C_TEXT, bold=True, size=8),
             alignment=_align("center", "center", wrap=True),
             border=_BORDER)

    # TOTAL 헤더
    _set(ws, cur, TOTAL_COL,
         value="TOTAL",
         fill=hdr_fill,
         font=_font(C_TEXT, bold=True, size=9),
         alignment=_align("center", "center"),
         border=_BORDER)

    cur += 1

    # ══════════════════════════════════════════════════════════════════════════
    # 프로젝트 블록 (3행 per project)
    # ══════════════════════════════════════════════════════════════════════════
    grand_usd = Decimal("0")
    grand_krw = Decimal("0")

    data_fill = _fill(C_WHITE)   # row 1: 사용량 — 흰색
    sub_fill  = _fill(C_SUB)     # row 2: monthly unit price — #F2F2F2
    amt_fill  = _fill(C_HEADER)  # row 3: amount — #E2EFDA

    LABELS       = ["사용량", "monthly unit price", "amount"]
    LABEL_FILLS  = [data_fill, sub_fill, amt_fill]
    ROW_HEIGHTS  = [18, 16, 18]

    for pr in proj_results:
        rA = cur          # 사용량 (usage)
        rB = cur + 1      # monthly unit price
        rC = cur + 2      # amount (금액, KRW)

        tot_usd    = pr["total_usd"]
        tot_krw    = pr["total_krw"]
        grand_usd += tot_usd
        grand_krw += tot_krw

        for r, h in zip([rA, rB, rC], ROW_HEIGHTS):
            rh(r, h)

        # ── 행 레이블 (Col A, 행별 배경색) ───────────────────────────────────
        for r, label, lf in zip([rA, rB, rC], LABELS, LABEL_FILLS):
            _set(ws, r, LABEL_COL,
                 value=label,
                 fill=lf,
                 font=_font(C_TEXT, bold=True, size=8),
                 alignment=_align("center", "center"),
                 border=_BORDER)

        # ── 프로젝트명 셀 (Col B, 세로 병합 3행) ─────────────────────────────
        _pre_merge_style(ws, rA, PROJ_COL, rC, PROJ_COL,
                         fill=hdr_fill, border=_BORDER)
        ws.merge_cells(start_row=rA, start_column=PROJ_COL,
                       end_row=rC,   end_column=PROJ_COL)
        cell = ws.cell(row=rA, column=PROJ_COL, value=pr["proj_name"])
        cell.fill      = hdr_fill
        cell.font      = _font(C_TEXT, bold=True, size=9)
        cell.alignment = _align("center", "center", wrap=True)
        cell.border    = _BORDER

        # ── API별 데이터 ──────────────────────────────────────────────────────
        for k, api_name in enumerate(api_list):
            sd = pr["skus"].get(
                api_name,
                {"usage": 0, "subtotal_usd": Decimal("0"), "final_krw": Decimal("0")},
            )
            usage     = sd["usage"]
            final_krw = sd["final_krw"]

            unit_price_krw = (
                float(final_krw) / usage * 1000
                if usage > 0 and final_krw > 0 else None
            )

            ac = api_col(k)

            # rA: 사용량 — 흰색, 우측 정렬, #,##0
            _set(ws, rA, ac,
                 value=usage if usage > 0 else None,
                 fill=data_fill,
                 font=_font(C_TEXT, size=8),
                 alignment=_align("right", "center"),
                 border=_BORDER,
                 number_format="#,##0")

            # rB: monthly unit price — #F2F2F2, 우측 정렬
            _set(ws, rB, ac,
                 value=unit_price_krw,
                 fill=sub_fill,
                 font=_font(C_TEXT, size=8),
                 alignment=_align("right", "center"),
                 border=_BORDER,
                 number_format="\u20a9#,##0.0")

            # rC: amount — #E2EFDA, 우측 정렬, 소수점 없음
            _set(ws, rC, ac,
                 value=int(final_krw) if final_krw else None,
                 fill=amt_fill,
                 font=_font(C_TEXT, size=8),
                 alignment=_align("right", "center"),
                 border=_BORDER,
                 number_format="\u20a9#,##0")

        # ── TOTAL 열 ──────────────────────────────────────────────────────────
        _set(ws, rA, TOTAL_COL, fill=data_fill, border=_BORDER)
        _set(ws, rB, TOTAL_COL, fill=sub_fill,  border=_BORDER)
        _set(ws, rC, TOTAL_COL,
             value=int(tot_krw),
             fill=amt_fill,
             font=_font(C_TEXT, bold=True, size=9),
             alignment=_align("right", "center"),
             border=_BORDER,
             number_format="\u20a9#,##0")

        cur = rC + 1

        # ── USD 소계 행 (compact) ─────────────────────────────────────────────
        rh(cur, 14)
        _pre_merge_style(ws, cur, LABEL_COL, cur, PROJ_COL,
                         fill=sub_fill, border=_BORDER)
        ws.merge_cells(start_row=cur, start_column=LABEL_COL,
                       end_row=cur,   end_column=PROJ_COL)
        cell = ws.cell(row=cur, column=LABEL_COL, value="합계($)")
        cell.fill      = sub_fill
        cell.font      = _font(C_TEXT, bold=True, size=7)
        cell.alignment = _align("right", "center")
        cell.border    = _BORDER

        for k in range(NUM_APIS):
            sd = pr["skus"].get(
                api_list[k],
                {"usage": 0, "subtotal_usd": Decimal("0"), "final_krw": Decimal("0")},
            )
            sub_usd = sd["subtotal_usd"]
            _set(ws, cur, api_col(k),
                 value=float(sub_usd) if sub_usd else None,
                 fill=sub_fill,
                 font=_font(C_TEXT, size=7),
                 alignment=_align("right", "center"),
                 border=_BORDER,
                 number_format='"$"#,##0.0000')

        _set(ws, cur, TOTAL_COL,
             value=float(tot_usd),
             fill=sub_fill,
             font=_font(C_TEXT, bold=True, size=8),
             alignment=_align("right", "center"),
             border=_BORDER,
             number_format='"$"#,##0.00')

        cur += 1

        # ── 프로젝트 간 구분선 (테두리 없는 얇은 행) ──────────────────────────
        rh(cur, 5)
        for c in range(1, MAX_COL + 1):
            ws.cell(row=cur, column=c).fill = _fill(C_SUB)
        cur += 1

    # ══════════════════════════════════════════════════════════════════════════
    # 최하단 합계 행 (orange)
    # ══════════════════════════════════════════════════════════════════════════
    rh(cur, 5)
    cur += 1

    orange_fill = _fill(C_ORANGE)

    for r in (cur, cur + 1):
        rh(r, 22)
        for c in range(1, MAX_COL + 1):
            ws.cell(row=r, column=c).fill = orange_fill

    _pre_merge_style(ws, cur, LABEL_COL, cur + 1, TOTAL_COL - 1,
                     fill=orange_fill, border=_BORDER)
    ws.merge_cells(start_row=cur,   start_column=LABEL_COL,
                   end_row=cur + 1, end_column=TOTAL_COL - 1)
    cell = ws.cell(row=cur, column=LABEL_COL, value="합계 대상 금액")
    cell.fill      = orange_fill
    cell.font      = Font(color="FFFFFF", bold=True, size=10, name="맑은 고딕")
    cell.alignment = _align("right", "center")
    cell.border    = _BORDER

    _set(ws, cur, TOTAL_COL,
         value=float(grand_usd),
         fill=orange_fill,
         font=Font(color="FFFFFF", bold=True, size=10, name="맑은 고딕"),
         alignment=_align("right", "center"),
         border=_BORDER,
         number_format='"$"#,##0.00')

    _set(ws, cur + 1, TOTAL_COL,
         value=int(grand_krw),
         fill=orange_fill,
         font=Font(color="FFFFFF", bold=True, size=11, name="맑은 고딕"),
         alignment=_align("right", "center"),
         border=_BORDER,
         number_format="\u20a9#,##0")

    # ── 인쇄 설정 ─────────────────────────────────────────────────────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
