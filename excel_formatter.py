import io
from decimal import Decimal
import xlsxwriter


def create_report_excel(line_items, company_name, billing_month, exchange_rate,
                        margin_rate, sku_master_rows, proj_results=None):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {"in_memory": True})

    _write_invoice_sheet(wb, line_items, company_name, billing_month, exchange_rate, margin_rate)
    if proj_results:
        _write_project_sheet(wb, proj_results)

    wb.close()
    output.seek(0)
    return output.read()


# ─────────────────────────────────────────────────────────────────────────────
# Sheet 1 : Invoice
# ─────────────────────────────────────────────────────────────────────────────
def _write_invoice_sheet(wb, line_items, company_name, billing_month, exchange_rate, margin_rate):
    ws = wb.add_worksheet("Invoice")

    # ── 포맷 ──────────────────────────────────────────────────────────────────
    H  = wb.add_format({"bold": True, "align": "center", "valign": "vcenter",
                        "bg_color": "#707070", "font_color": "#FFFFFF", "border": 1})
    HI = wb.add_format({"bold": True, "align": "left", "valign": "vcenter",
                        "bg_color": "#E8F4F6", "border": 1})
    B  = wb.add_format({"align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
    N  = wb.add_format({"align": "center", "num_format": "#,##0", "border": 1})
    D  = wb.add_format({"align": "right",  "num_format": "$#,##0.00", "border": 1})
    S  = wb.add_format({"align": "center", "bg_color": "#F2F2F2", "border": 1, "bold": True})
    SD = wb.add_format({"align": "right",  "bg_color": "#F2F2F2", "border": 1,
                        "bold": True, "num_format": "$#,##0.00"})
    K  = wb.add_format({"bold": True, "align": "right", "bg_color": "#707070",
                        "font_color": "#FFFFFF", "border": 1})
    KN = wb.add_format({"bold": True, "align": "right", "bg_color": "#707070",
                        "font_color": "#FFFFFF", "border": 1, "num_format": "#,##0"})

    # ── 컬럼 폭 ──────────────────────────────────────────────────────────────
    ws.set_column("A:A", 35)
    ws.set_column("B:H", 13)

    # ── 인보이스 정보 헤더 (2행) ─────────────────────────────────────────────
    ws.merge_range(0, 0, 0, 7, f"Billing Account Name : {company_name}", HI)
    ws.merge_range(1, 0, 1, 7, f"Term of Use : {billing_month}", HI)

    # ── 테이블 헤더 행 ────────────────────────────────────────────────────────
    ws.write_row(2, 0, ["API", "Usage", "Free Usage", "Subtotal",
                        "할인 구간", "수량", "단가", "Amount"], H)

    tier_labels = {1: "0-100K", 2: "~500K", 3: "~1M", 4: "~5M", 5: "~10M"}
    curr = 3   # 0,1: 인보이스 헤더, 2: 테이블 헤더
    t_usd = Decimal("0")

    for item in sorted(line_items, key=lambda x: x.sku_name):
        # total_usage==0인 항목은 출력 제외 (미사용 API)
        if item.total_usage == 0:
            continue

        n = len(item.tier_breakdown)
        if n == 0:
            # tier 정의가 없는 SKU (is_billable=False 등) — 1행으로 출력
            ws.write(curr, 0, item.sku_name, B)
            ws.write(curr, 1, item.total_usage, N)
            ws.write(curr, 2, f"-{item.free_cap_applied:,}", N)
            ws.write(curr, 3, item.billable_usage, N)
            ws.write(curr, 4, "-", B)
            ws.write(curr, 5, "-", N)
            ws.write(curr, 6, "-", B)
            ws.write(curr, 7, 0.0, D)
            curr += 1
        else:
            if n > 1:
                ws.merge_range(curr, 0, curr + n - 1, 0, item.sku_name, B)
                ws.merge_range(curr, 1, curr + n - 1, 1, item.total_usage, N)
                ws.merge_range(curr, 2, curr + n - 1, 2,
                               f"-{item.free_cap_applied:,}", N)
                ws.merge_range(curr, 3, curr + n - 1, 3, item.billable_usage, N)
            else:
                ws.write(curr, 0, item.sku_name, B)
                ws.write(curr, 1, item.total_usage, N)
                ws.write(curr, 2, f"-{item.free_cap_applied:,}", N)
                ws.write(curr, 3, item.billable_usage, N)

            for i, tb in enumerate(item.tier_breakdown):
                ws.write(curr + i, 4,
                         tier_labels.get(tb.tier_number, f"T{tb.tier_number}"), B)
                ws.write(curr + i, 5,
                         tb.usage_in_tier if tb.usage_in_tier > 0 else "-", N)
                ws.write(curr + i, 6, float(tb.tier_cpm), D)
                ws.write(curr + i, 7, float(tb.amount_usd), D)
            curr += n

        # 소계 행
        ws.merge_range(curr, 0, curr, 4, "소계", S)
        ws.write(curr, 5, item.billable_usage, S)
        ws.write(curr, 6, "", S)
        ws.write(curr, 7, float(item.subtotal_usd), SD)
        t_usd += item.subtotal_usd
        curr += 1

    # ── 합계 / 환율 / 청구금액 ────────────────────────────────────────────────
    t_krw = (t_usd * Decimal(str(exchange_rate)) * Decimal(str(margin_rate))
             ).quantize(Decimal("1"))

    ws.merge_range(curr,     0, curr,     6, "합        계(USD)", B)
    ws.write      (curr,     7, float(t_usd), D)
    ws.merge_range(curr + 1, 0, curr + 1, 6,
                   f"환        율(적용 환율: {exchange_rate:,.2f})", B)
    ws.write      (curr + 1, 7, float(exchange_rate), D)
    ws.merge_range(curr + 2, 0, curr + 2, 6, "청 구 금 액(KRW)", K)
    ws.write      (curr + 2, 7, float(t_krw), KN)


# ─────────────────────────────────────────────────────────────────────────────
# Sheet 2 : Project (프로젝트별 요약 피벗 테이블)
# ─────────────────────────────────────────────────────────────────────────────
def _write_project_sheet(wb, proj_results):
    ws = wb.add_worksheet("Project")

    # 사용량이 1건 이상인 API만 컬럼으로 표시
    all_apis = set()
    for p in proj_results:
        for api, d in p["skus"].items():
            if d["usage"] > 0:
                all_apis.add(api)
    api_list = sorted(all_apis)

    # ── 포맷 ──────────────────────────────────────────────────────────────────
    DK = wb.add_format({"bold": True, "align": "center", "valign": "vcenter",
                        "bg_color": "#707070", "font_color": "#FFFFFF",
                        "border": 1, "text_wrap": True})
    SH = wb.add_format({"bold": True, "align": "center", "valign": "vcenter",
                        "bg_color": "#A9C5CD", "border": 1, "text_wrap": True})
    PJ = wb.add_format({"align": "left", "valign": "vcenter",
                        "border": 1, "text_wrap": True})
    N  = wb.add_format({"align": "right", "num_format": "#,##0", "border": 1})
    D  = wb.add_format({"align": "right", "num_format": "$#,##0.00", "border": 1})
    GN = wb.add_format({"align": "right", "num_format": "#,##0",
                        "bg_color": "#D6EAD0", "border": 1})
    GD = wb.add_format({"align": "right", "num_format": "$#,##0.00",
                        "bg_color": "#D6EAD0", "border": 1})
    TH = wb.add_format({"bold": True, "align": "right", "num_format": "#,##0",
                        "bg_color": "#4CAF50", "font_color": "#FFFFFF", "border": 1})
    TD = wb.add_format({"bold": True, "align": "right", "num_format": "$#,##0.00",
                        "bg_color": "#4CAF50", "font_color": "#FFFFFF", "border": 1})
    TL = wb.add_format({"bold": True, "align": "center",
                        "bg_color": "#4CAF50", "font_color": "#FFFFFF", "border": 1})

    # ── 컬럼 폭 설정 ─────────────────────────────────────────────────────────
    ws.set_column(0, 0, 28)          # 프로젝트명
    n_apis = len(api_list)
    # 각 API → 2 cols (사용량, 금액KRW), 마지막에 합계USD/합계KRW
    ws.set_column(1, 2 * n_apis + 2, 11)

    # ── Row 0 : API 그룹 헤더 (2열 병합) ────────────────────────────────────
    ws.write(0, 0, "프로젝트", DK)
    for idx, api in enumerate(api_list):
        col = 1 + idx * 2
        ws.merge_range(0, col, 0, col + 1, api, DK)
    last_col = 1 + n_apis * 2
    ws.write(0, last_col,     "합계(USD)", DK)
    ws.write(0, last_col + 1, "합계(KRW)", DK)

    # ── Row 1 : 서브 헤더 ────────────────────────────────────────────────────
    ws.write(1, 0, "", SH)
    for idx in range(n_apis):
        col = 1 + idx * 2
        ws.write(1, col,     "사용량",    SH)
        ws.write(1, col + 1, "금액(KRW)", SH)
    ws.write(1, last_col,     "", SH)
    ws.write(1, last_col + 1, "", SH)

    # ── 데이터 행 (프로젝트 1개 = 1행) ────────────────────────────────────────
    row = 2
    grand_usd = Decimal("0")
    grand_krw = Decimal("0")
    col_usage_sum = {api: 0 for api in api_list}
    col_krw_sum   = {api: Decimal("0") for api in api_list}

    for p in proj_results:
        # 해당 프로젝트에 사용량이 전혀 없으면 행 생략
        if all(p["skus"].get(api, {}).get("usage", 0) == 0 for api in api_list):
            continue

        ws.write(row, 0, p["proj_name"], PJ)
        for idx, api in enumerate(api_list):
            col = 1 + idx * 2
            d = p["skus"].get(api, {"usage": 0, "final_krw": Decimal("0")})
            usage = d["usage"]
            krw   = d["final_krw"]
            ws.write(row, col,     usage if usage > 0 else "-", N)
            ws.write(row, col + 1, float(krw) if krw > 0 else "-", GN)
            col_usage_sum[api] += usage
            col_krw_sum[api]   += krw

        ws.write(row, last_col,     float(p["total_usd"]), D)
        ws.write(row, last_col + 1, float(p["total_krw"]), GD)
        grand_usd += p["total_usd"]
        grand_krw += p["total_krw"]
        row += 1

    # ── 합계 행 ──────────────────────────────────────────────────────────────
    ws.write(row, 0, "합  계", TL)
    for idx, api in enumerate(api_list):
        col = 1 + idx * 2
        ws.write(row, col,     col_usage_sum[api], TH)
        ws.write(row, col + 1, float(col_krw_sum[api]), TH)
    ws.write(row, last_col,     float(grand_usd), TD)
    ws.write(row, last_col + 1, float(grand_krw), TH)
