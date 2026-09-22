from __future__ import annotations

import json
import os
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / 'outputs' / '20260920-new-source-parent-model'
MODEL = OUT_DIR / 'model.json'
OUT = Path(os.environ.get('PARENT_XLSX_OUT', str(OUT_DIR / '户外地垫市场分析-SPEC3-父体口径可回勾.xlsx')))

data = json.loads(MODEL.read_text(encoding='utf-8'))
months = data['months']
parents = data['parentRows']
raw_rows = data['rawRows']
candidates = data['candidateRows']
summaries = data['summaries']

OUT.parent.mkdir(parents=True, exist_ok=True)
wb = xlsxwriter.Workbook(str(OUT), {'constant_memory': False, 'nan_inf_to_errors': True, 'strings_to_urls': False})
wb.set_properties({
    'title': '户外地垫市场分析 SPEC3 父体口径',
    'subject': '2024.08—2026.07 新数据源父体与有效子体销售额分析',
    'author': 'Codex',
    'comments': '公式列保留缓存结果，主源只读。'
})

fmt = {
    'title': wb.add_format({'bold': True, 'font_size': 16, 'font_color': '#FFFFFF', 'bg_color': '#1F2937'}),
    'h1': wb.add_format({'bold': True, 'font_size': 13, 'font_color': '#FFFFFF', 'bg_color': '#334155'}),
    'header': wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#475569', 'border': 1, 'text_wrap': True, 'valign': 'vcenter'}),
    'subheader': wb.add_format({'bold': True, 'bg_color': '#E2E8F0', 'border': 1, 'text_wrap': True}),
    'text': wb.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'}),
    'number': wb.add_format({'border': 1, 'num_format': '#,##0.00'}),
    'integer': wb.add_format({'border': 1, 'num_format': '#,##0'}),
    'percent': wb.add_format({'border': 1, 'num_format': '0.0%'}),
    'money': wb.add_format({'border': 1, 'num_format': '$#,##0.00'}),
    'formula': wb.add_format({'border': 1, 'bg_color': '#F8FAFC', 'num_format': '#,##0.00'}),
    'formula_int': wb.add_format({'border': 1, 'bg_color': '#F8FAFC', 'num_format': '#,##0'}),
    'formula_pct': wb.add_format({'border': 1, 'bg_color': '#F8FAFC', 'num_format': '0.0%'}),
    'note': wb.add_format({'text_wrap': True, 'valign': 'top', 'bg_color': '#FFF7ED', 'border': 1}),
    'blank': wb.add_format({'border': 1}),
}
for name,style in fmt.items():
    style.set_font_name('Arial')
    if name not in ('title','h1'):style.set_font_size(11)
    if name not in ('header','title','h1'):
        style.set_border(0);style.set_font_color('#334155')
    style.set_valign('vcenter')
fmt['formula'].set_num_format('#,##0.00')

def growth(current, previous):
    return current / previous - 1 if isinstance(current, (int, float)) and isinstance(previous, (int, float)) and previous > 0 else None

def safe(v):
    if v is None:
        return ''
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return v

def write_sheet_header(ws, headers, widths=None):
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 2)
    ws.autofilter(0, 0, 0, len(headers) - 1)
    ws.set_row(0, 44)
    ws.set_default_row(23)
    for c, h in enumerate(headers):
        ws.write(0, c, h, fmt['header'])
        if widths and c < len(widths):
            ws.set_column(c, c, widths[c])

def formula(ws, row, col, expression, cached, kind='number'):
    ws.write_formula(row, col, expression, fmt.get(kind, fmt['formula']), safe(cached))

def qsheet(name):
    return "'" + name.replace("'", "''") + "'"

# 00 overview
ws = wb.add_worksheet('00_数据总览')
ws.set_tab_color('#0F766E')
overview = [
    ('SPEC版本', data['metadata']['specVersion']),
    ('批次ID', data['metadata']['batchId']),
    ('数据源', data['metadata']['sourceDir']),
    ('覆盖月份', f"{data['metadata']['firstMonth'][:4]}.{data['metadata']['firstMonth'][4:]}—{data['metadata']['lastMonth'][:4]}.{data['metadata']['lastMonth'][4:]}"),
    ('源文件数', data['metadata']['fileCount']),
    ('原始行数', data['metadata']['rawRowCount']),
    ('BSR候选子体行数', data['metadata']['candidateRowCount']),
    ('父体月度组数', data['metadata']['parentMonthCount']),
    ('同月重复ASIN组数', data['metadata']['duplicateAsinGroups']),
    ('父体销量冲突组数', data['metadata']['conflictParentCount']),
]
write_sheet_header(ws, ['字段', '值', '说明'], [26, 45, 80])
for r, (k, v) in enumerate(overview, 1):
    ws.write(r, 0, k, fmt['text']); ws.write(r, 1, safe(v), fmt['text']); ws.write(r, 2, '批次元数据或源审计结果', fmt['text'])
notes = [
    '处理顺序：24个月新源 → 名称/排名逐项对应Outdoor Rugs BSR 1—100 → 候选子体明细 → 按父ASIN分别聚合。',
    '父体销量为清洗估计：严格多数>50%取众数，否则取中位数；所有多值父体均进入诊断。',
    '父体有效子体销售额只累加同时有子体销量和子体销售额的子体行；月销售额($)仅作源字段回勾，不参与父体销售额合计。',
    '空白不等于0。PARTIAL/NONE覆盖状态必须与销售额同时展示。',
    '本批次不计算利润，不使用旧主源，不用参考表数字替换新源结果。'
]
ws.write(12, 0, '规则摘要', fmt['h1']); ws.merge_range(12, 0, 12, 2, '规则摘要', fmt['h1'])
for i, n in enumerate(notes, 13):
    ws.merge_range(i, 0, i, 2, '• ' + n, fmt['note']); ws.set_row(i, 36)
for i,name in enumerate(['01_整体市场月度','02_PP市场月度','03_高客单价月度','04_Genimo品牌月度','05_年度YOY','06_BSR分层','07_父体冲突诊断','09_勾稽检查','91_BSR候选子体明细','92_父体月度汇总'],20):
    ws.write_url(i,0,f"internal:'{name}'!A1",fmt['text'],name)
ws.merge_range('A32:C33','人工验算：91页按月份、父体键筛选；92页每项汇总都有公式。灰底为计算列。修改91页数值后可重算，新增/删除行或修改父体键、月份、分类须重新构建。',fmt['note'])

# 90 raw input
ws = wb.add_worksheet('90_原始输入')
raw_headers = ['月份', '源文件', '源Sheet', '源行号'] + list(data['rawRows'][0]['raw'].keys()) + ['Outdoor Rugs排名（解析）','候选筛选状态']
write_sheet_header(ws, raw_headers, [10, 32, 24, 9] + [20] * len(data['rawRows'][0]['raw']) + [20,26])
for r, item in enumerate(raw_rows, 1):
    values = [item['month'], item['file'], item['sheet'], item['sourceRow']] + [safe(item['raw'].get(h)) for h in data['rawRows'][0]['raw'].keys()] + [safe(item['rank']),item['bsrStatus']]
    for c, v in enumerate(values): ws.write(r, c, v, fmt['text'] if isinstance(v, str) else fmt['number'])

# 91 candidate child detail
ws = wb.add_worksheet('91_BSR候选子体明细')
child_headers = ['月份', '父体键', '父ASIN', 'ASIN', '源行号', '小类BSR', '商品标题', '品牌', '月销量', '月销售额($)', '子体销量', '子体销售额($)', '价格($)', 'PP候选行', 'Genimo候选行', 'BSR合格公式', '子体有效公式', '部分缺失公式', '源文件', '源Sheet','月销量出现次数','月销量首次出现','隐含售价($)','与标价差异率']
write_sheet_header(ws, child_headers, [10, 18, 16, 16, 9, 10, 48, 18, 13, 16, 13, 17, 11, 12, 14, 12, 13, 13, 32, 24])
groups={}
for i,item in enumerate(candidates,2): groups.setdefault((item['month'],item['parentKey']),[]).append(i)
for r, item in enumerate(candidates, 1):
    vals = [item['month'], item['parentKey'], item['parent'], item['asin'], item['sourceRow'], item['rank'], item['title'], item['brand'], item['sales'], item['sourceRevenue'], item['childSales'], item['childRevenue'], item['price'], 1 if item['pp'] else 0, 1 if item['genimo'] else 0]
    for c, v in enumerate(vals):
        if v is None: ws.write_blank(r, c, None, fmt['blank'])
        elif isinstance(v, (int, float)): ws.write(r, c, v, fmt['number'])
        else: ws.write(r, c, v, fmt['text'])
    excel_row = r + 1
    formula(ws, r, 15, f'=IF(AND(F{excel_row}>=1,F{excel_row}<=100),1,0)', 1, 'formula_int')
    valid = 1 if item.get('childSales') is not None and item.get('childRevenue') is not None and item['childSales'] >= 0 and item['childRevenue'] >= 0 and not (item['childSales'] == 0 and item['childRevenue'] > 0) else 0
    formula(ws, r, 16, f'=IF(AND(ISNUMBER(K{excel_row}),ISNUMBER(L{excel_row}),K{excel_row}>=0,L{excel_row}>=0,OR(K{excel_row}>0,L{excel_row}=0)),1,0)', valid, 'formula_int')
    partial = 1 if ((item.get('childSales') is not None) ^ (item.get('childRevenue') is not None)) else 0
    formula(ws, r, 17, f'=IF(OR(ISNUMBER(K{excel_row}),ISNUMBER(L{excel_row})),IF(Q{excel_row}=1,0,1),0)', partial, 'formula_int')
    ws.write(r, 18, item['file'], fmt['text']); ws.write(r, 19, item['sheet'], fmt['text'])
    indices=groups[(item['month'],item['parentKey'])]; first,last=indices[0],indices[-1]
    peers=candidates[first-2:last-1]
    freq=sum(x['sales']==item['sales'] for x in peers) if item['sales'] is not None else 0
    formula(ws,r,20,f'=IF(ISNUMBER(I{excel_row}),COUNTIFS($I${first}:$I${last},I{excel_row}),0)',freq,'formula_int')
    first_value=int(item['sales'] is not None and not any(x['sales']==item['sales'] for x in candidates[first-2:r-1]))
    formula(ws,r,21,f'=IF(ISNUMBER(I{excel_row}),IF(COUNTIFS($I${first}:I{excel_row},I{excel_row})=1,1,0),0)',first_value,'formula_int')
    implicit=item['sourceRevenue']/item['sales'] if item['sales'] and item['sourceRevenue'] is not None else None
    formula(ws,r,22,f'=IF(AND(ISNUMBER(I{excel_row}),I{excel_row}>0,ISNUMBER(J{excel_row})),J{excel_row}/I{excel_row},"")',implicit)
    formula(ws,r,23,f'=IF(AND(ISNUMBER(W{excel_row}),ISNUMBER(M{excel_row}),M{excel_row}>0),W{excel_row}/M{excel_row}-1,"")',implicit/item['price']-1 if implicit is not None and item['price'] and item['price']>0 else None,'formula_pct')
ws.set_column(20,23,18)

# 92 parent aggregation
ws = wb.add_worksheet('92_父体月度汇总')
parent_headers = ['月份', '父体键', '父ASIN', '市场分类', 'Genimo', 'BSR中位数', 'BSR最小值', 'BSR粗档', 'BSR细档', '父体月销量估计', '月销量众数', '众数次数', '众数覆盖率', '月销量中位数', '不同值数', '销量状态', '候选子体行数', '有效子体行数', '有效子体销量', '有效子体销售额($)', '子体覆盖率', '子体状态', '子体加权均价($)', '原始月销售额中位数($)', '原始月销量中位数', 'PP命中率', '源行号摘要','源销量最小值','源销量最大值','销量多值冲突','有效父体销量行数','PP混合提示']
write_sheet_header(ws, parent_headers, [10, 20, 16, 14, 10, 11, 11, 12, 12, 14, 14, 12, 12, 14, 11, 16, 14, 14, 14, 19, 12, 12, 16, 20, 16, 12, 40])
detail_last = len(candidates) + 1
for r, p in enumerate(parents, 1):
    market = 'PP市场' if p['pp'] else '高客单价市场'
    if p['genimo']:
        market = 'PP市场' if p['pp'] else '高客单价市场'
    values = [p['month'], p['parentKey'], p['parent'], market, '是' if p['genimo'] else '否', p['rankMedian'], p['rankMin'], next((b['name'] for b in data['bands'] if b['key'] == p['band']), ''), next((b['name'] for b in data['fineBands'] if b['key'] == p['fineBand']), ''), p['parentSales'], p['salesMode'], p['salesModeCount'], p['salesModeShare'], p['salesMedian'], p['salesUnique'], p['salesStatus']]
    for c, v in enumerate(values): ws.write(r, c, safe(v), fmt['text'] if isinstance(v, str) else (fmt['percent'] if c == 12 else fmt['number']))
    er = r + 1
    indices=groups[(p['month'],p['parentKey'])]; start,end=indices[0],indices[-1]
    def dr(col): return f"'91_BSR候选子体明细'!${col}${start}:${col}${end}"
    formula(ws,r,5,f'=MEDIAN({dr("F")})',p['rankMedian'])
    formula(ws,r,6,f'=MIN({dr("F")})',p['rankMin'])
    formula(ws,r,7,f'=IF(F{er}<=20,"头部 1—20",IF(F{er}<=50,"中部 21—50","尾部 51—100"))',next(b['name'] for b in data['bands'] if b['key']==p['band']),'text')
    formula(ws,r,8,f'=IF(F{er}<=5,"1—5",IF(F{er}<=10,"6—10",IF(F{er}<=20,"11—20",IF(F{er}<=50,"21—50","51—100"))))',next(b['name'] for b in data['fineBands'] if b['key']==p['fineBand']),'text')
    formula(ws,r,9,f'=IF(AE{er}=0,"",IF(M{er}>0.5,K{er},N{er}))',p['parentSales'])
    formula(ws,r,10,f'=IF(AE{er}=0,"",IF(MAX({dr("U")})=1,MIN({dr("I")}),MODE({dr("I")})))',p['salesMode'])
    formula(ws,r,11,f'=MAX({dr("U")})',p['salesModeCount'],'formula_int')
    formula(ws,r,12,f'=IF(AE{er}>0,L{er}/AE{er},"")',p['salesModeShare'],'formula_pct')
    formula(ws,r,13,f'=IF(AE{er}>0,MEDIAN({dr("I")}),"")',p['salesMedian'])
    formula(ws,r,14,f'=SUM({dr("V")})',p['salesUnique'],'formula_int')
    formula(ws,r,15,f'=IF(AE{er}=0,"NO_VALUE",IF(O{er}=1,"CONSISTENT",IF(M{er}>0.5,"MODE","MEDIAN_CONFLICT")))',p['salesStatus'],'text')
    # Formula columns based on the child detail sheet. Cached values come from the independent JS model.
    month_ref = f'$A{er}'; parent_ref = f'$B{er}'
    formula(ws, r, 16, f'=COUNTIFS(\'91_BSR候选子体明细\'!$A$2:$A${detail_last},{month_ref},\'91_BSR候选子体明细\'!$B$2:$B${detail_last},{parent_ref})', p['candidateRows'], 'formula_int')
    formula(ws, r, 17, f'=COUNTIFS(\'91_BSR候选子体明细\'!$A$2:$A${detail_last},{month_ref},\'91_BSR候选子体明细\'!$B$2:$B${detail_last},{parent_ref},\'91_BSR候选子体明细\'!$Q$2:$Q${detail_last},1)', p['childValidRows'], 'formula_int')
    formula(ws,r,18,f'=IF(R{er}=0,"",SUMIFS({dr("K")},{dr("Q")},1))',p['childSales'],'formula')
    formula(ws,r,19,f'=IF(R{er}=0,"",SUMIFS({dr("L")},{dr("Q")},1))',p['childRevenue'],'formula')
    formula(ws, r, 20, f'=IF(Q{er}>0,R{er}/Q{er},"")', p['childCoverage'], 'formula_pct')
    formula(ws,r,21,f'=IF(R{er}=0,"NONE",IF(R{er}=Q{er},"FULL","PARTIAL"))',p['childStatus'],'text')
    formula(ws, r, 22, f'=IF(AND(ISNUMBER(S{er}),S{er}>0),T{er}/S{er},"")', p['childAsp'], 'formula')
    ws.write(r, 23, safe(p['sourceRevenueMedian']), fmt['number']); ws.write(r, 24, safe(p['sourceSalesMedian']), fmt['number'])
    ws.write(r, 25, safe(p['ppRatio']), fmt['percent']); ws.write(r, 26, ','.join(str(x) for x in p['sourceRows']), fmt['text'])
    formula(ws,r,25,f'=IF(COUNTA({dr("G")})>0,SUM({dr("N")})/COUNTA({dr("G")}),"")',p['ppRatio'],'formula_pct')
    formula(ws,r,3,f'=IF(Z{er}>=0.5,"PP市场","高客单价市场")',market,'text')
    formula(ws,r,4,f'=IF(COUNTA({dr("H")})=0,"否",IF(SUM({dr("O")})/COUNTA({dr("H")})>=0.5,"是","否"))','是' if p['genimo'] else '否','text')
    formula(ws,r,23,f'=IF(COUNT({dr("J")})>0,MEDIAN({dr("J")}),"")',p['sourceRevenueMedian'])
    formula(ws,r,24,f'=IF(ISNUMBER(N{er}),N{er},"")',p['salesMedian'])
    formula(ws,r,27,f'=IF(AE{er}>0,MIN({dr("I")}),"")',p['salesMin'])
    formula(ws,r,28,f'=IF(AE{er}>0,MAX({dr("I")}),"")',p['salesMax'])
    formula(ws,r,29,f'=IF(O{er}>1,1,0)',int(p['salesConflict']),'formula_int')
    formula(ws,r,30,f'=COUNT({dr("I")})',p['salesValidRows'],'formula_int')
    formula(ws,r,31,f'=IF(AND(Z{er}>0,Z{er}<1),"混合标题，暂按多数归类","")','混合标题，暂按多数归类' if p['ppRatio'] is not None and 0<p['ppRatio']<1 else '', 'text')
ws.set_column(27,30,16); ws.set_column(31,31,32)

def write_monthly_sheet(name, title, key):
    ws = wb.add_worksheet(name)
    headers = ['月份', '父ASIN数', 'BSR候选子体行数', '父体月销量估计', '有效子体销量', '有效子体销售额($)', '子体覆盖率', '销量MOM（上月）', '销售额MOM（上月）', '销量YOY（去年同月）', '销售额YOY（去年同月）', '月销量多值父ASIN数', '多值率', '有效销售额父体数', '父体纳入率','子体加权均价($)','销售额状态','取中位数父体数','无销量父体数','有效子体行数','源排名未对应行数','比较范围提示']
    write_sheet_header(ws, headers, [10, 11, 16, 15, 16, 20, 12, 14, 14, 14, 14, 18, 12, 18, 14,18,14,18,16,16,18,34])
    rows = summaries[key]
    parent_last = len(parents) + 1
    for i, row in enumerate(rows, 1):
        ws.write(i, 0, row['month'], fmt['text'])
        er = i + 1
        # 92_父体月度汇总 is the single formula source for the monthly business pages.
        criteria = [f"'92_父体月度汇总'!$A$2:$A${parent_last},A{er}"]
        if key == 'pp': criteria += [f"'92_父体月度汇总'!$D$2:$D${parent_last},\"PP市场\""]
        elif key == 'high': criteria += [f"'92_父体月度汇总'!$D$2:$D${parent_last},\"高客单价市场\""]
        elif key in ['genimo','genimoPP']:
            criteria += [f"'92_父体月度汇总'!$E$2:$E${parent_last},\"是\""]
            if key=='genimoPP': criteria += [f"'92_父体月度汇总'!$D$2:$D${parent_last},\"PP市场\""]
        crit = ','.join(criteria)
        count_formula = f'=COUNTIFS({crit})'
        candidate_formula = f'=SUMIFS(\'92_父体月度汇总\'!$Q$2:$Q${parent_last},{crit})'
        parent_sales_formula = f'=IF(B{er}>S{er},SUMIFS(\'92_父体月度汇总\'!$J$2:$J${parent_last},{crit}),"")'
        child_sales_formula = f'=IF(N{er}>0,SUMIFS(\'92_父体月度汇总\'!$S$2:$S${parent_last},{crit}),"")'
        child_revenue_formula = f'=IF(N{er}>0,SUMIFS(\'92_父体月度汇总\'!$T$2:$T${parent_last},{crit}),"")'
        valid_rows_formula = f'=SUMIFS(\'92_父体月度汇总\'!$R$2:$R${parent_last},{crit})'
        conflict_formula = f'=COUNTIFS({crit},\'92_父体月度汇总\'!$AD$2:$AD${parent_last},1)'
        revenue_parent_formula = f'=COUNTIFS({crit},\'92_父体月度汇总\'!$V$2:$V${parent_last},"<>NONE")'
        formula(ws, i, 1, count_formula, row['parentCount'], 'formula_int')
        formula(ws, i, 2, candidate_formula, row['candidateRows'], 'formula_int')
        formula(ws, i, 3, parent_sales_formula, row['parentSales'], 'formula')
        formula(ws, i, 4, child_sales_formula, row['childSales'], 'formula')
        formula(ws, i, 5, child_revenue_formula, row['childRevenue'], 'formula')
        formula(ws, i, 6, f'=IFERROR(SUMIFS(\'92_父体月度汇总\'!$R$2:$R${parent_last},{crit})/SUMIFS(\'92_父体月度汇总\'!$Q$2:$Q${parent_last},{crit}),"")', row['childCoverage'], 'formula_pct')
        # MOM and YOY reference the formula-driven current/prior rows, using blank when the denominator is not available.
        def gf(col,prior): return f'=IF(AND(ISNUMBER({col}{er}),ISNUMBER({col}{prior}),{col}{prior}>0),{col}{er}/{col}{prior}-1,"")' if prior else '=""'
        formula(ws, i, 7, gf('D',er-1 if i>1 else None), row['momParentSales'], 'formula_pct')
        formula(ws, i, 8, gf('F',er-1 if i>1 else None), row['momChildRevenue'], 'formula_pct')
        prior_year_row = next((j + 2 for j, x in enumerate(rows) if x['month'] == f"{int(row['month'][:4])-1}{row['month'][4:]}"), None)
        formula(ws, i, 9, gf('D',prior_year_row), row['yoyParentSales'], 'formula_pct')
        formula(ws, i, 10, gf('F',prior_year_row), row['yoyChildRevenue'], 'formula_pct')
        formula(ws, i, 11, conflict_formula, row['conflictParents'], 'formula_int')
        formula(ws, i, 12, f'=IFERROR(L{er}/B{er},"")', row['conflictRate'], 'formula_pct')
        formula(ws, i, 13, revenue_parent_formula, row['validRevenueParents'], 'formula_int')
        formula(ws, i, 14, f'=IFERROR(N{er}/B{er},"")', row['revenueParentCoverage'], 'formula_pct')
        formula(ws,i,15,f'=IF(AND(ISNUMBER(E{er}),E{er}>0),F{er}/E{er},"")',row['childAsp'])
        formula(ws,i,16,f'=IF(T{er}=0,"NONE",IF(T{er}=C{er},"FULL","PARTIAL"))',row['revenueStatus'],'text')
        formula(ws,i,17,f'=COUNTIFS({crit},\'92_父体月度汇总\'!$P$2:$P${parent_last},"MEDIAN_CONFLICT")',row['medianParents'],'formula_int')
        formula(ws,i,18,f'=COUNTIFS({crit},\'92_父体月度汇总\'!$P$2:$P${parent_last},"NO_VALUE")',row['missingSalesParents'],'formula_int')
        formula(ws,i,19,valid_rows_formula,row['validChildRows'],'formula_int')
        ws.write(i,20,row['ambiguousBsrRows'],fmt['integer'])
        limitations=[]
        if row['momScopeLimited']: limitations.append('MOM含范围缺失月')
        if row['yoyScopeLimited']: limitations.append('YOY含范围缺失月')
        ws.write(i,21,'；'.join(limitations) or '对应关系完整',fmt['text'])
    ws.write(len(rows)+3, 0, title, fmt['h1']); ws.merge_range(len(rows)+3, 0, len(rows)+3, 4, title, fmt['h1'])
    ws.write(len(rows)+4, 0, '销售额口径', fmt['subheader']); ws.merge_range(len(rows)+4, 1, len(rows)+4, 14, '有效子体销售额：仅累加同时存在子体销量和子体销售额的子体行；月销售额($)不在父体下累加。', fmt['note'])
    ws.set_column(15,19,19)
    ws.set_row(len(rows)+4,35)
    for c,title,position in [(3,'父体月销量估计','A32'),(5,'有效子体销售额（USD）','I32'),(7,'销量MOM（相对上月）','A49'),(15,'有效子体加权均价（USD）','I49')]:
        chart=wb.add_chart({'type':'column' if c in (3,5) else 'line'})
        chart.add_series({'name':title,'categories':[name,1,0,len(rows),0],'values':[name,1,c,len(rows),c],'line':{'color':'#2563EB'},'fill':{'color':'#2563EB'}})
        chart.set_title({'name':title}); chart.set_legend({'none':True}); chart.set_size({'width':660,'height':320})
        chart.set_y_axis({'num_format':'0%' if c==7 else '#,##0','min':0} if c in (3,5) else {'num_format':'0%' if c==7 else '#,##0.00'})
        ws.insert_chart(position,chart)
    return ws

write_monthly_sheet('01_整体市场月度', '第一部分｜整体市场', 'overall')
write_monthly_sheet('02_PP市场月度', '第二部分｜PP市场', 'pp')
write_monthly_sheet('03_高客单价月度', '第三部分｜高客单价市场', 'high')
write_monthly_sheet('04_Genimo品牌月度', '第四部分｜Genimo品牌', 'genimo')
write_monthly_sheet('04B_Genimo_PP月度', 'Genimo在PP市场', 'genimoPP')

# 05 matched-month annual comparison: both sides use exactly the same months.
monthly_names={'overall':'01_整体市场月度','pp':'02_PP市场月度','high':'03_高客单价月度','genimo':'04_Genimo品牌月度','genimoPP':'04B_Genimo_PP月度'}
scope_names={'overall':'整体市场','pp':'PP市场','high':'高客单价市场','genimo':'Genimo品牌','genimoPP':'Genimo PP市场'}
ws=wb.add_worksheet('05_年度YOY')
write_sheet_header(ws,['市场范围','年份','本期共同月份','基期共同月份','本期父体销量估计','基期父体销量估计','本期有效子体销售额($)','基期有效子体销售额($)','销量YOY','销售额YOY','本期子体覆盖率','基期子体覆盖率','期间状态','数据范围提示'],[20,10,46,46,22,22,24,24,16,16,20,20,28,38])
row_idx=1
for key, rows in data['annual'].items():
    sh=monthly_names[key]
    for item in rows:
        er=row_idx+1
        for c,v in enumerate([scope_names[key],item['year'],','.join(item['months']),','.join(item['priorMonths'])]): ws.write(row_idx,c,v,fmt['text'])
        def period_sum(col,ms):
            refs=[f"'{sh}'!{col}{months.index(m)+2}" for m in ms]
            return f'IF(COUNT({",".join(refs)})>0,SUM({",".join(refs)}),"")' if refs else '""'
        for col,source,period,keyname in [(4,'D','months','parentSales'),(5,'D','priorMonths','priorSales'),(6,'F','months','childRevenue'),(7,'F','priorMonths','priorRevenue')]:
            formula(ws,row_idx,col,'='+period_sum(source,item[period]),item[keyname])
        for c,a,b,k in [(8,'E','F','yoyParentSales'),(9,'G','H','yoyChildRevenue')]:
            formula(ws,row_idx,c,f'=IF(AND(ISNUMBER({a}{er}),ISNUMBER({b}{er}),{b}{er}>0),{a}{er}/{b}{er}-1,"")',item[k],'formula_pct')
        for c,period,k in [(10,'months','childCoverage'),(11,'priorMonths','priorCoverage')]:
            ms=item[period]
            expression='='+period_sum('T',ms)+'/'+period_sum('C',ms) if ms else '=""'
            # Parentheses preserve guarded subtotals as expressions.
            expression=f'=IF(({period_sum("C",ms)})>0,({period_sum("T",ms)})/({period_sum("C",ms)}),"")' if ms else '=""'
            formula(ws,row_idx,c,expression,item[k],'formula_pct')
        ws.write(row_idx,12,{'NO_BASE':'无同比基期','MATCHED_MONTHS':'仅共同月份，非全年','FULL_YEAR':'完整年度'}[item['status']],fmt['text'])
        ws.write(row_idx,13,'共同月份含类目排名未对应行，结果低估' if item['scopeLimited'] else 'Outdoor Rugs范围完整',fmt['text'])
        row_idx+=1

# 06 tier business formulas use the same parent rows, with no gaps at fractional medians.
ws=wb.add_worksheet('06_BSR分层')
write_sheet_header(ws,['市场范围','月份','分层','父ASIN数','父体销量估计','有效子体销量','有效子体销售额($)','子体覆盖率','销量MOM（上月）','销售额MOM（上月）','销量YOY（去年同月）','销售额YOY（去年同月）','候选子体行数','有效子体行数','有销量父体数'],[20,12,20,14,20,20,22,16,20,20,22,22,20,20,20])
row_idx=1
parent_last=len(parents)+1
def pr(col): return f"'92_父体月度汇总'!${col}$2:${col}${parent_last}"
for key,band_groups in data['tiers'].items():
    for band in band_groups:
        first_er=row_idx+1
        for index,cur in enumerate(band['rows']):
            er=row_idx+1
            for c,v in enumerate([scope_names[key],cur['month'],band['name']]):ws.write(row_idx,c,v,fmt['text'])
            criteria=[f'{pr("A")},B{er}',f'{pr("H" if band["type"]=="coarse" else "I")},C{er}']
            if key in ['pp','high','genimoPP']:criteria.append(f'{pr("D")},"'+('高客单价市场' if key=='high' else 'PP市场')+'"')
            if key in ['genimo','genimoPP']:criteria.append(f'{pr("E")},"是"')
            crit=','.join(criteria)
            formula(ws,row_idx,3,f'=COUNTIFS({crit})',cur['parentCount'],'formula_int')
            for c,source,k,guard in [(4,'J','parentSales','O'),(5,'S','childSales','N'),(6,'T','childRevenue','N')]:
                formula(ws,row_idx,c,f'=IF({guard}{er}>0,SUMIFS({pr(source)},{crit}),"")',cur[k])
            for c,source,k in [(12,'Q','candidateRows'),(13,'R','validChildRows')]:formula(ws,row_idx,c,f'=SUMIFS({pr(source)},{crit})',cur[k],'formula_int')
            formula(ws,row_idx,14,f'=COUNTIFS({crit},{pr("P")},"<>NO_VALUE")',cur['salesParents'],'formula_int')
            formula(ws,row_idx,7,f'=IF(M{er}>0,N{er}/M{er},"")',cur['childCoverage'],'formula_pct')
            for c,col,offset,k in [(8,'E',1,'momParentSales'),(9,'G',1,'momChildRevenue'),(10,'E',12,'yoyParentSales'),(11,'G',12,'yoyChildRevenue')]:
                prior=er-offset
                formula(ws,row_idx,c,f'=IF(AND(ISNUMBER({col}{er}),ISNUMBER({col}{prior}),{col}{prior}>0),{col}{er}/{col}{prior}-1,"")' if prior>=first_er else '=""',cur[k],'formula_pct')
            row_idx+=1

# 07 diagnostics
ws = wb.add_worksheet('07_父体冲突诊断')
headers = ['月份', '父ASIN', '父体键', '候选子体行数', '月销量众数', '众数次数', '众数覆盖率', '月销量中位数', '月销量不同值数', '最小BSR', 'BSR中位数', '有效子体数', '子体覆盖率', '有效子体销售额($)', '源行号摘要']
write_sheet_header(ws, headers, [10, 16, 22, 15, 15, 11, 13, 15, 14, 10, 12, 13, 13, 20, 46])
for r, p in enumerate(data['conflictParents'], 1):
    vals = [p['month'], p['parent'], p['parentKey'], p['candidateRows'], p['salesMode'], p['salesModeCount'], p['salesModeShare'], p['salesMedian'], p['salesUnique'], p['rankMin'], p['rankMedian'], p['childValidRows'], p['childCoverage'], p['childRevenue'], ','.join(str(x) for x in p['sourceRows'])]
    for c, v in enumerate(vals): ws.write(r, c, safe(v), fmt['percent'] if c in (6,12) else (fmt['number'] if c in (3,4,5,7,8,9,10,11,13) else fmt['text']))
ws.write(len(data['conflictParents']) + 3, 0, '诊断说明', fmt['h1']); ws.merge_range(len(data['conflictParents']) + 3, 0, len(data['conflictParents']) + 3, 6, '诊断说明', fmt['h1'])
ws.merge_range(len(data['conflictParents']) + 4, 0, len(data['conflictParents']) + 6, 6, '全部多值父体都列出；主值仍为估计。严格多数>50%取众数，否则取中位数。最小BSR值仅作审计。', fmt['note'])
parent_index={(p['month'],p['parentKey']):i+2 for i,p in enumerate(parents)}
for r,p in enumerate(data['conflictParents'],1):
    source=parent_index[(p['month'],p['parentKey'])]
    for c,col in [(3,'Q'),(4,'K'),(5,'L'),(6,'M'),(7,'N'),(8,'O'),(9,'G'),(10,'F'),(11,'R'),(12,'U'),(13,'T')]:
        existing=[p['candidateRows'],p['salesMode'],p['salesModeCount'],p['salesModeShare'],p['salesMedian'],p['salesUnique'],p['rankMin'],p['rankMedian'],p['childValidRows'],p['childCoverage'],p['childRevenue']][c-3]
        formula(ws,r,c,f'=IF(ISNUMBER(\'92_父体月度汇总\'!{col}{source}),\'92_父体月度汇总\'!{col}{source},"")',existing,'formula_pct' if c in [6,12] else 'formula')

# 08 strategy
ws = wb.add_worksheet('08_2027策略')
write_sheet_header(ws, ['项目', '内容', '使用说明'], [28, 85, 55])
strategy = [
    ('策略使用范围', '基于候选父体估计销量、有效子体销售额和覆盖率。', '不输出缺乏依据的固定链接数量。'),
    ('数据可执行限制', '子体销售额在部分月份覆盖率较低；父体月销量存在冲突；报告应优先把趋势和份额作为方向证据。', '不是利润推导，不补齐缺失销售额。'),
    ('PP市场策略', '比较PP父体销量、有效子体销售额、覆盖率和BSR中位数趋势；连续两个月改善且覆盖率稳定时再扩大测试。', '以实际数据回算后写入报告。'),
    ('高客单价策略', '与PP使用完全相同的父体口径；不能仅依据行级月销售额判断增长。', '关注父体月销量和有效子体销售额的共同方向。'),
    ('Genimo策略', '使用Genimo父体份额、PP内份额、BSR分层和冲突/覆盖率状态确定链接测试优先级。', '不使用利润字段；不把进留退当成上新下架。'),
    ('花型/尺寸/颜色', '本批次只保留源字段，不在覆盖不足时虚构趋势。', '需要可靠提取规则后再扩展。')
]
for r, row in enumerate(strategy, 1):
    for c, v in enumerate(row): ws.write(r, c, v, fmt['text'])
for r in range(1,7):ws.set_row(r,48)
for c,label in enumerate(['观察指标','2026.07','2025.07','份额变化（百分点）','2027建议及条件']):ws.write(9,c,label,fmt['header'])
ws.set_column(1,2,34);ws.set_column(3,3,24);ws.set_column(4,4,72)
for r,(label,nk,dk,action) in enumerate([
    ('Genimo父体销量整体份额','genimo','overall','先维护已有父体销量份额，按头中尾表现选择增量测试位置。'),
    ('Genimo父体销量PP份额','genimoPP','pp','优先验证PP内的份额变化；按父体而不是子体条数规划链接。'),
    ('PP父体销量整体份额','pp','overall','PP与高客单价试验预算按销量趋势复核；不直接用缺失销售额推断市场空间。')],10):
    ws.write(r,0,label,fmt['text']); ws.write(r,4,action,fmt['text']); ws.set_row(r,44)
    refs={'overall':'01_整体市场月度','pp':'02_PP市场月度','genimo':'04_Genimo品牌月度','genimoPP':'04B_Genimo_PP月度'}
    values=[]
    for c,idx in [(1,23),(2,11)]:
        a=summaries[nk][idx]['parentSales'];b=summaries[dk][idx]['parentSales'];value=a/b if a is not None and b and b>0 else None;values.append(value)
        er=idx+2
        formula(ws,r,c,f'=IF(AND(ISNUMBER(\'{refs[nk]}\'!D{er}),ISNUMBER(\'{refs[dk]}\'!D{er}),\'{refs[dk]}\'!D{er}>0),\'{refs[nk]}\'!D{er}/\'{refs[dk]}\'!D{er},"")',value,'formula_pct')
    formula(ws,r,3,f'=IF(AND(ISNUMBER(B{r+1}),ISNUMBER(C{r+1})),B{r+1}-C{r+1},"")',values[0]-values[1] if all(x is not None for x in values) else None,'formula_pct')

# 94 rules
ws = wb.add_worksheet('94_规则说明')
write_sheet_header(ws, ['规则项', '当前规则', '人工核验方式'], [30, 100, 70])
rules = [
    ('数据范围', '仅使用 data/raw/0920new 内2024.08—2026.07的24个文件。', '检查00_数据总览和源文件清单。'),
    ('BSR', '小类目与排名按换行逐项对应，只取Outdoor Rugs的1—100排名。', '90页核对原始类目/BSR，91页查看合格排名。'),
    ('父体月销量', '父体下月销量严格多数>50%取众数，否则取中位数；全部多值情况均标记。', '92页核对众数、中位数、不同值数和状态。'),
    ('父体销售额', '有效子体销售额是同时存在子体销量和子体销售额的子体行销售额之和。', '92页核对有效行数、销售额、覆盖率。'),
    ('月销售额($)', '保留为源字段回勾和隐含平均价诊断，不在父体下累加。', '91页W/X列隐含售价与标价差异率。'),
    ('空白', '没有有效子体时销售额留空；FULL仅表示候选行都有值，不代表采集了全部变体。', '91页公式和92页状态。'),
    ('PP/高客单价', '父体候选子体标题按plastic命中多数归类，两个市场互补。', '按父ASIN抽查标题命中。'),
    ('Genimo', '父体候选子体品牌按Genimo多数归类。', '按父ASIN抽查品牌。'),
    ('MOM/YOY', 'MOM=本月/上月−1；YOY=本月/去年同月−1。', '01—05页核对期间和分母。'),
    ('利润', '不计算利润；FBA、毛利率、Coupon不参与利润推导。', '检查08页和HTML正文。')
]
rules += [
    ('年度期间','只用当年和上年都存在的相同月份；2025年比8—12月，2026年比1—7月。','05页的分子分母均列出。'),
    ('均价','有效子体销售额÷同批有效子体销量，不除以父体估计销量。','92页W=T/S；业务月度页P=F/E。'),
    ('BSR中位数边界','中位数可为小数；粗档(0,20]、(20,50]、(50,100]，细档同理。','20.5进入中部；5.5进入6—10档。'),
    ('父体范围限制','候选父体销量包含其父体全部变体估计；有效子体销售额仅涵盖已采集且BSR合格的有效子体。','两指标分别展示，不互除推算市场ASP。'),
    ('分类混合','整词plastic按有效标题多数分配；50%平局暂归PP并标注。高客单价是非plastic补集名称，不保证售价更高。','92页PP命中率和混合提示。'),
    ('重算边界','91页数值编辑联动92和月度/年度/分层页；原始90快照、分类键、结构及HTML须重新构建。','HTML为发布快照，不会跟随本地Excel编辑。')
    ,('2025.08源限制','354行只有一个小类名称但有两个BSR，无法可靠对应Outdoor Rugs排名，已排除。','90页筛选202508和AMBIGUOUS_CATEGORY_RANK；93页看排除数。')
]
for r, row in enumerate(rules, 1):
    for c, v in enumerate(row): ws.write(r, c, v, fmt['text'])

ws=wb.add_worksheet('09_勾稽检查')
write_sheet_header(ws,['月份','父体数分区差','父体销量分区差','有效子体销售额分区差','粗档父体数差','细档父体数差','Genimo销量整体份额','Genimo销量PP份额','Genimo有效销售额整体份额','Genimo有效销售额PP份额','结果'],[12,18,22,26,20,20,24,24,27,27,16])
for i,month in enumerate(months,1):
    er=i+1; ws.write(i,0,month,fmt['text'])
    for c,col in [(1,'B'),(2,'D'),(3,'F')]:
        formula(ws,i,c,f"='01_整体市场月度'!{col}{er}-'02_PP市场月度'!{col}{er}-'03_高客单价月度'!{col}{er}",0)
    tier_last=1+sum(len(b['rows']) for bands in data['tiers'].values() for b in bands)
    def tier_count(labels):
        return '+'.join(f'SUMIFS(\'06_BSR分层\'!$D$2:$D${tier_last},\'06_BSR分层\'!$A$2:$A${tier_last},"整体市场",\'06_BSR分层\'!$B$2:$B${tier_last},A{er},\'06_BSR分层\'!$C$2:$C${tier_last},"{label}")' for label in labels)
    # Coarse and fine tier labels overlap at 21—50/51—100 only in display? Coarse labels include prefixes.
    for c,labels in [(4,[b['name'] for b in data['bands']]),(5,[b['name'] for b in data['fineBands']])]:
        formula(ws,i,c,f'={tier_count(labels)}-\'01_整体市场月度\'!B{er}',0)
    for c,num,den,metric,field in [(6,'04_Genimo品牌月度','01_整体市场月度','D','parentSales'),(7,'04B_Genimo_PP月度','02_PP市场月度','D','parentSales'),(8,'04_Genimo品牌月度','01_整体市场月度','F','childRevenue'),(9,'04B_Genimo_PP月度','02_PP市场月度','F','childRevenue')]:
        nk='genimoPP' if c in [7,9] else 'genimo'; dk='pp' if c in [7,9] else 'overall'
        a=summaries[nk][i-1][field]; b=summaries[dk][i-1][field]
        formula(ws,i,c,f'=IF(AND(ISNUMBER(\'{num}\'!{metric}{er}),ISNUMBER(\'{den}\'!{metric}{er}),\'{den}\'!{metric}{er}>0),\'{num}\'!{metric}{er}/\'{den}\'!{metric}{er},"")',a/b if a is not None and b is not None and b>0 else None,'formula_pct')
    formula(ws,i,10,f'=IF(AND(ABS(B{er})<0.000001,ABS(C{er})<0.000001,ABS(D{er})<0.000001,ABS(E{er})<0.000001,ABS(F{er})<0.000001),"一致","请复核")','一致','text')

ws=wb.add_worksheet('93_来源登记')
write_sheet_header(ws,['月份','源文件','业务Sheet','原始行数','BSR候选行数','SHA-256','非Outdoor Rugs','名称排名不对应','无效排名','超过100'],[12,60,26,16,18,72,20,20,18,18])
for i,s in enumerate(data['sourceFiles'],1):
    for c,k in enumerate(['month','file','sheet','rawRows','candidateRows','sha256']):ws.write(i,c,s[k],fmt['text'] if c not in [3,4] else fmt['integer'])
    for c,k in enumerate(['OTHER_CATEGORY','AMBIGUOUS_CATEGORY_RANK','INVALID_BSR','OUTSIDE_TOP100'],6):ws.write(i,c,s['exclusions'][k],fmt['integer'])

# Filters cover the data body, not just the header. Outputs precede the supporting detail.
limits={'90_原始输入':len(raw_rows),'91_BSR候选子体明细':len(candidates),'92_父体月度汇总':len(parents),'05_年度YOY':15,'06_BSR分层':tier_last-1,'07_父体冲突诊断':len(data['conflictParents']),'09_勾稽检查':24,'93_来源登记':24}
for sheet in wb.worksheets():
    name=sheet.get_name()
    limit=24 if name in monthly_names.values() else limits.get(name,sheet.dim_rowmax)
    if name!='00_数据总览':sheet.autofilter(0,0,limit,sheet.dim_colmax)
    sheet.set_zoom(85)
    sheet.set_landscape(); sheet.fit_to_pages(1,0); sheet.repeat_rows(0)
    if name.startswith(('01','02','03','04')):sheet.set_tab_color('#2563EB')
    if name in ['91_BSR候选子体明细','92_父体月度汇总']:sheet.set_tab_color('#64748B')
sheet_order=['00_数据总览','01_整体市场月度','02_PP市场月度','03_高客单价月度','04_Genimo品牌月度','04B_Genimo_PP月度','05_年度YOY','06_BSR分层','07_父体冲突诊断','08_2027策略','09_勾稽检查','90_原始输入','91_BSR候选子体明细','92_父体月度汇总','93_来源登记','94_规则说明']
order_index={name:i for i,name in enumerate(sheet_order)}
wb.worksheets_objs.sort(key=lambda sheet:order_index[sheet.get_name()])

wb.close()
print(json.dumps({'output': str(OUT), 'bytes': OUT.stat().st_size, 'sheets': len(wb.worksheets_objs)}, ensure_ascii=False))
