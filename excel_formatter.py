import io, xlsxwriter
from decimal import Decimal

def create_report_excel(line_items, company_name, billing_month, exchange_rate, margin_rate, sku_master_rows):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet("Invoice")
    
    # 스타일
    H = wb.add_format({'bold':True, 'align':'center', 'bg_color':'#707070', 'font_color':'#FFFFFF', 'border':1})
    B = wb.add_format({'align':'center', 'valign':'vcenter', 'border':1, 'text_wrap':True})
    N = wb.add_format({'align':'center', 'num_format':'#,##0', 'border':1})
    D = wb.add_format({'align':'right', 'num_format':'$#,##0.00', 'border':1})
    S = wb.add_format({'align':'center', 'bg_color':'#F2F2F2', 'border':1, 'bold':True})
    K = wb.add_format({'bold':True, 'align':'right', 'bg_color':'#707070', 'font_color':'#FFFFFF', 'border':1})

    ws.write_row(0, 0, ['API', 'Usage', 'Free Usage', 'Subtotal', '할인 구간', '수량', '단가', 'Amount'], H)
    ws.set_column('A:A', 35); ws.set_column('B:H', 13)

    tier_labels = {1:"0-100K", 2:"~500K", 3:"~1M", 4:"~5M", 5:"~10M"}
    curr = 1
    t_usd, t_krw = Decimal("0"), Decimal("0")

    for item in sorted(line_items, key=lambda x: x.sku_name):
        n = len(item.tier_breakdown)
        if n == 0: continue

        # 병합 겹침 방지: n > 1일 때만 병합
        if n > 1:
            ws.merge_range(curr, 0, curr+n-1, 0, item.sku_name, B)
            ws.merge_range(curr, 1, curr+n-1, 1, item.total_usage, N)
            ws.merge_range(curr, 2, curr+n-1, 2, f"-{item.free_cap_applied:,}", N)
            ws.merge_range(curr, 3, curr+n-1, 3, item.billable_usage, N)
        else:
            ws.write(curr, 0, item.sku_name, B); ws.write(curr, 1, item.total_usage, N)
            ws.write(curr, 2, f"-{item.free_cap_applied:,}", N); ws.write(curr, 3, item.billable_usage, N)

        for i, tb in enumerate(item.tier_breakdown):
            ws.write(curr+i, 4, tier_labels.get(tb.tier_number, f"T{tb.tier_number}"), B)
            ws.write(curr+i, 5, tb.usage_in_tier if tb.usage_in_tier > 0 else "-", N)
            ws.write(curr+i, 6, float(tb.tier_cpm), D)
            ws.write(curr+i, 7, float(tb.amount_usd), D)
        
        curr += n
        ws.merge_range(curr, 0, curr, 4, "소계", S)
        ws.write(curr, 5, item.billable_usage, S)
        ws.write(curr, 6, "", S)
        ws.write(curr, 7, float(item.subtotal_usd), S)
        t_usd += item.subtotal_usd; t_krw += item.final_krw
        curr += 1

    ws.merge_range(curr, 0, curr, 6, "합        계(USD)", B); ws.write(curr, 7, float(t_usd), D)
    ws.merge_range(curr+1, 0, curr+1, 6, f"환        율(적용 환율: {exchange_rate:,.2f})", B); ws.write(curr+1, 7, float(exchange_rate), D)
    ws.merge_range(curr+2, 0, curr+2, 6, "청 구 금 액(KRW)", K); ws.write(curr+2, 7, float(t_krw), K)

    wb.close(); output.seek(0)
    return output.read()