from __future__ import annotations

import json
import os
from pathlib import Path

import xlsxwriter

ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = ROOT / 'outputs' / '20260920-new-source-parent-model'
MODEL = json.loads((OUT_DIR / 'model-spec37.json').read_text(encoding='utf-8'))
OUT = Path(os.environ.get('SPEC37_XLSX_OUT', str(OUT_DIR / '户外地垫市场分析-SPEC3.7-父体口径.xlsx')))
OUT.parent.mkdir(parents=True, exist_ok=True)

wb = xlsxwriter.Workbook(str(OUT), {'constant_memory': False, 'strings_to_urls': False, 'nan_inf_to_errors': True})
wb.set_properties({'title': '户外地垫市场分析 SPEC 3.7', 'subject': '父体Q/T平均与BSR月份+BSR值平均', 'author': 'Codex'})

fmt = {
    'title': wb.add_format({'bold': True, 'font_size': 15, 'font_color': '#FFFFFF', 'bg_color': '#1F2937'}),
    'header': wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#334155', 'border': 1, 'text_wrap': True, 'valign': 'vcenter'}),
    'section': wb.add_format({'bold': True, 'bg_color': '#E2E8F0', 'border': 1}),
    'text': wb.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'}),
    'num': wb.add_format({'border': 1, 'num_format': '#,##0.00'}),
    'money': wb.add_format({'border': 1, 'num_format': '$#,##0.00'}),
    'int': wb.add_format({'border': 1, 'num_format': '#,##0'}),
    'ratio': wb.add_format({'border': 1, 'num_format': '0.00x'}),
    'formula': wb.add_format({'border': 1, 'bg_color': '#F8FAFC', 'num_format': '#,##0.00'}),
    'formula_money': wb.add_format({'border': 1, 'bg_color': '#F8FAFC', 'num_format': '$#,##0.00'}),
    'formula_ratio': wb.add_format({'border': 1, 'bg_color': '#F8FAFC', 'num_format': '0.00x'}),
    'formula_int': wb.add_format({'border': 1, 'bg_color': '#F8FAFC', 'num_format': '#,##0'}),
    'note': wb.add_format({'text_wrap': True, 'valign': 'top', 'bg_color': '#FFF7ED', 'border': 1}),
}

def safe(v):
    return '' if v is None else v

def write_table(name, headers, rows, widths=None):
    ws = wb.add_worksheet(name)
    ws.hide_gridlines(2)
    ws.freeze_panes(1, 0)
    ws.set_row(0, 38)
    for c, h in enumerate(headers):
        ws.write(0, c, h, fmt['header'])
        if widths and c < len(widths):
            ws.set_column(c, c, widths[c])
    for r, row in enumerate(rows, 1):
        for c, value in enumerate(row):
            if value is None:
                ws.write_blank(r, c, None, fmt['text'])
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                ws.write_number(r, c, value, fmt['num'])
            else:
                ws.write(r, c, value, fmt['text'])
    ws.autofilter(0, 0, max(len(rows), 1), len(headers) - 1)
    return ws

def set_fmt(ws, col, rows, style):
    if rows:
        ws.set_column(col, col, None, fmt[style])

def formula(ws, row, col, expression, cached, style='formula'):
    ws.write_formula(row, col, expression, fmt[style], safe(cached))

def month_label(m):
    return f'{m[:4]}.{m[4:]}'

def qsheet(s):
    return "'" + s.replace("'", "''") + "'"

months = MODEL['months']
parent_rows = MODEL['parentRows']
cand_rows = MODEL['candidateRows']
parent_end = len(parent_rows) + 1
cand_end = len(cand_rows) + 1

# 00_数据总览
ws = wb.add_worksheet('00_数据总览'); ws.hide_gridlines(2)
ws.set_column('A:A', 24); ws.set_column('B:B', 44); ws.set_column('C:C', 78)
ws.merge_range('A1:C1', '户外地垫市场分析 · SPEC 3.7', fmt['title'])
overview = [
    ('SPEC版本', MODEL['metadata']['specVersion'], '当前执行规范'),
    ('批次ID', MODEL['metadata']['batchId'], '模型、Excel、HTML共用批次'),
    ('数据源', MODEL['metadata']['sourceDir'], '24个按月源文件'),
    ('覆盖月份', f"{month_label(MODEL['metadata']['firstMonth'])}—{month_label(MODEL['metadata']['lastMonth'])}", '共24个月'),
    ('原始行数', MODEL['metadata']['rawRowCount'], '源文件业务行总数'),
    ('BSR候选行数', MODEL['metadata']['candidateRowCount'], '任一小类BSR 1—100有效排名的源行'),
    ('BSR观察记录数', MODEL['metadata']['bsrObservationCount'], '多值BSR按每个有效排名生成记录'),
    ('父体月份组数', MODEL['metadata']['parentMonthCount'], '月份+父ASIN'),
    ('同月重复ASIN组数', MODEL['metadata']['duplicateAsinGroups'], '应为0才能发布'),
    ('Q冲突父体组数', MODEL['metadata']['qConflictParentCount'], '同组Q有效值不唯一'),
    ('T冲突父体组数', MODEL['metadata']['tConflictParentCount'], '同组T有效值不唯一'),
]
for c, h in enumerate(['字段', '值', '说明']): ws.write(2, c, h, fmt['header'])
for r, row in enumerate(overview, 3):
    for c, value in enumerate(row): ws.write(r, c, safe(value), fmt['text'])
ws.merge_range('A16:C16', '共同规则', fmt['section'])
notes = [
    '候选池只看小类BSR是否有1—100有效排名，小类目名称只保留回勾。',
    '父体主指标按月份+父ASIN，对Q列月销量和T列月销售额分别取有效非负值平均。',
    'BSR独立按月份+BSR值汇总，忽略ASIN；同月同BSR的Q/T取平均，再按1—20、21—50、51—100三档。',
    'MOM=本月÷去年同月；YOY=本年与上一年共同覆盖月份汇总相除，均不减1。',
    '子体销量和子体销售额只作为覆盖与诊断指标；本批次不计算利润。',
]
for i, note in enumerate(notes, 16): ws.merge_range(i, 0, i, 2, note, fmt['note'])

# 90 raw input
raw_headers = ['月份', '源文件', '源Sheet', '源行号', 'ASIN', '品牌', '商品标题', '父ASIN', '小类目', '小类BSR', '月销量', '月销售额($)', '子体销量', '子体销售额($)', '价格($)', 'BSR解析值', '合格状态']
raw_rows = []
for r in MODEL['rawRows']:
    raw_rows.append([r['month'], r['file'], r['sheet'], r['sourceRow'], r['asin'], r['brand'], r['title'], r['parent'], r['category'], r['sourceBsr'], r['sales'], r['sourceRevenue'], r['childSales'], r['childRevenue'], r['price'], '\n'.join(map(str, r['parsedRanks'])), r['bsrStatus']])
raw_ws = write_table('90_原始输入', raw_headers, raw_rows, [10, 34, 24, 9, 16, 18, 46, 16, 28, 18, 13, 17, 13, 17, 12, 16, 16])
for col in [10, 12]: raw_ws.set_column(col, col, 14, fmt['num'])
for col in [11, 13, 14]: raw_ws.set_column(col, col, 17, fmt['money'])

# 91 candidate detail
cand_headers = ['月份', '父体键', '父ASIN', 'ASIN', '源行号', '小类BSR解析值', '商品标题', '品牌', '月销量Q', '月销售额T($)', '子体销量', '子体销售额($)', '价格($)', 'PP候选行', 'Genimo候选行', 'Q有效', 'T有效', '子体有效', '源文件', '源Sheet']
cand_out = []
for r in cand_rows:
    valid_child = all(isinstance(r.get(k), (int, float)) and r.get(k) >= 0 for k in ['childSales', 'childRevenue']) and not (r.get('childSales') == 0 and r.get('childRevenue', 0) > 0)
    cand_out.append([r['month'], r['parentKey'], r['parent'], r['asin'], r['sourceRow'], '\n'.join(map(str, r['eligibleRanks'])), r['title'], r['brand'], r['sales'], r['sourceRevenue'], r['childSales'], r['childRevenue'], r['price'], None if r['pp'] is None else int(r['pp']), None if r['genimo'] is None else int(r['genimo']), int(isinstance(r.get('sales'), (int, float)) and r['sales'] >= 0), int(isinstance(r.get('sourceRevenue'), (int, float)) and r['sourceRevenue'] >= 0), int(valid_child), r['file'], r['sheet']])
cand_ws = write_table('91_BSR候选子体明细', cand_headers, cand_out, [10, 20, 16, 16, 9, 14, 46, 18, 13, 17, 13, 17, 12, 12, 14, 10, 10, 10, 34, 24])
for col in [8, 10]: cand_ws.set_column(col, col, 14, fmt['num'])
for col in [9, 11, 12]: cand_ws.set_column(col, col, 17, fmt['money'])

# 92 parent aggregation
parent_headers = ['月份', '父体键', '父ASIN', '市场分类', 'Genimo', '候选行数', 'Q有效行数', 'T有效行数', '月销量Q平均', '月销售额T平均($)', '平均标价($)', '加权成交均价($)', 'Q覆盖率', 'T覆盖率', 'Q不同值数', 'T不同值数', '子体销量诊断', '子体销售额诊断($)', '子体覆盖率', '子体状态', 'PP命中率', '品牌命中率', '源行号摘要']
parent_out = []
for p in parent_rows:
    parent_out.append([p['month'], p['parentKey'], p['parent'], 'PP市场' if p['pp'] else '高客单价市场', '是' if p['genimo'] else '否', p['candidateRows'], p['qValidRows'], p['tValidRows'], p['qAvg'], p['tAvg'], p['avgPrice'], p['weightedPrice'], p['qValidRows'] / p['candidateRows'], p['tValidRows'] / p['candidateRows'], p['qUnique'], p['tUnique'], p['childSales'], p['childRevenue'], p['childCoverage'], p['childStatus'], p['ppRatio'], p['genimoRatio'], ','.join(map(str, p['sourceRows']))])
parent_ws = write_table('92_父体月度汇总', parent_headers, parent_out, [10, 20, 16, 16, 10, 11, 11, 11, 15, 18, 15, 18, 12, 12, 12, 12, 16, 18, 12, 13, 12, 12, 42])
for col in [8, 9, 10, 16]: parent_ws.set_column(col, col, 17, fmt['num'])
for col in [9, 10, 11, 17]: parent_ws.set_column(col, col, 18, fmt['money'])
for col in [12, 13, 18, 20, 21]: parent_ws.set_column(col, col, 12, fmt['ratio'])
for i, p in enumerate(parent_rows, 1):
    er = i + 1
    formula(parent_ws, i, 8, f'=IFERROR(AVERAGEIFS({qsheet("91_BSR候选子体明细")}!$I$2:$I${cand_end},{qsheet("91_BSR候选子体明细")}!$A$2:$A${cand_end},A{er},{qsheet("91_BSR候选子体明细")}!$B$2:$B${cand_end},B{er}),"")', p['qAvg'], 'formula')
    formula(parent_ws, i, 9, f'=IFERROR(AVERAGEIFS({qsheet("91_BSR候选子体明细")}!$J$2:$J${cand_end},{qsheet("91_BSR候选子体明细")}!$A$2:$A${cand_end},A{er},{qsheet("91_BSR候选子体明细")}!$B$2:$B${cand_end},B{er}),"")', p['tAvg'], 'formula_money')
    formula(parent_ws, i, 10, f'=IFERROR(AVERAGEIFS({qsheet("91_BSR候选子体明细")}!$M$2:$M${cand_end},{qsheet("91_BSR候选子体明细")}!$A$2:$A${cand_end},A{er},{qsheet("91_BSR候选子体明细")}!$B$2:$B${cand_end},B{er}),"")', p['avgPrice'], 'formula_money')
    formula(parent_ws, i, 11, f'=IFERROR(J{er}/I{er},"")', p['weightedPrice'], 'formula_money')

# 01—04B monthly output. Core figures are formula-linked to 92; cached values come from model.
monthly_headers = ['月份', '月销量', '月销售额($)', '平均标价($)', '加权成交均价($)', '销量MOM', '销量YOY', '销售额MOM', '销售额YOY']
sheet_names = {'overall': '01_整体市场月度', 'pp': '02_PP市场月度', 'high': '03_高客单价市场月度', 'genimo': '04_Genimo品牌月度', 'genimoPP': '04B_Genimo_PP月度'}
def criteria(key):
    if key == 'overall': return ''
    if key == 'pp': return f',{qsheet("92_父体月度汇总")}!$D$2:$D${parent_end},"PP市场"'
    if key == 'high': return f',{qsheet("92_父体月度汇总")}!$D$2:$D${parent_end},"高客单价市场"'
    if key == 'genimo': return f',{qsheet("92_父体月度汇总")}!$E$2:$E${parent_end},"是"'
    return f',{qsheet("92_父体月度汇总")}!$D$2:$D${parent_end},"PP市场",{qsheet("92_父体月度汇总")}!$E$2:$E${parent_end},"是"'
for key, sheet_name in sheet_names.items():
    s = write_table(sheet_name, monthly_headers, [[m, None, None, None, None, None, None, None, None] for m in months], [12, 16, 18, 16, 18, 12, 12, 14, 14])
    rows = MODEL['summaries'][key]
    crit = criteria(key)
    for i, (m, rmodel) in enumerate(zip(months, rows), 1):
        er = i + 1
        formula(s, i, 1, f'=SUMIFS({qsheet("92_父体月度汇总")}!$I$2:$I${parent_end},{qsheet("92_父体月度汇总")}!$A$2:$A${parent_end},A{er}{crit})', rmodel['sales'], 'formula')
        formula(s, i, 2, f'=SUMIFS({qsheet("92_父体月度汇总")}!$J$2:$J${parent_end},{qsheet("92_父体月度汇总")}!$A$2:$A${parent_end},A{er}{crit})', rmodel['revenue'], 'formula_money')
        formula(s, i, 3, f'=IFERROR(AVERAGEIFS({qsheet("92_父体月度汇总")}!$K$2:$K${parent_end},{qsheet("92_父体月度汇总")}!$A$2:$A${parent_end},A{er}{crit}),"")', rmodel['avgPrice'], 'formula_money')
        formula(s, i, 4, f'=IFERROR(C{er}/B{er},"")', rmodel['weightedPrice'], 'formula_money')
        prior = f'{int(m[:4]) - 1}{m[4:]}'
        pi = months.index(prior) if prior in months else None
        formula(s, i, 5, f'=IFERROR(B{er}/B{pi + 2},"")' if pi is not None else '=""', rmodel['momSales'], 'formula_ratio')
        formula(s, i, 7, f'=IFERROR(C{er}/C{pi + 2},"")' if pi is not None else '=""', rmodel['momRevenue'], 'formula_ratio')
        year = int(m[:4])
        if year == 2025:
            current_start, prior_start = '202508', '202408'
        elif year == 2026:
            current_start, prior_start = '202601', '202501'
        else:
            current_start = prior_start = None
        if current_start:
            prior_end = f'{year - 1}{m[4:]}'
            yoy_sales = f'=IFERROR(SUMIFS($B$2:$B$25,$A$2:$A$25,">={current_start}",$A$2:$A$25,"<={m}")/SUMIFS($B$2:$B$25,$A$2:$A$25,">={prior_start}",$A$2:$A$25,"<={prior_end}"),"")'
            yoy_revenue = f'=IFERROR(SUMIFS($C$2:$C$25,$A$2:$A$25,">={current_start}",$A$2:$A$25,"<={m}")/SUMIFS($C$2:$C$25,$A$2:$A$25,">={prior_start}",$A$2:$A$25,"<={prior_end}"),"")'
        else:
            yoy_sales = yoy_revenue = '=""'
        formula(s, i, 6, yoy_sales, rmodel['yoySales'], 'formula_ratio')
        formula(s, i, 8, yoy_revenue, rmodel['yoyRevenue'], 'formula_ratio')

# 05 annual
annual_out = []
for key in ['overall', 'pp', 'high', 'genimo']:
    for r in MODEL['annual'][key]:
        annual_out.append([names[key] if False else {'overall':'整体市场','pp':'PP市场','high':'高客单价市场','genimo':'Genimo品牌'}[key], r['year'], ', '.join(map(month_label, r['months'])) or '—', ', '.join(map(month_label, r['priorMonths'])) or '—', r['sales'], r['priorSales'], r['revenue'], r['priorRevenue'], r['yoySales'], r['yoyRevenue'], r['weightedPrice'], r['status']])
annual_ws = write_table('05_年度YOY', ['范围', '年份', '本期共同月份', '基期共同月份', '本期销量', '基期销量', '本期销售额($)', '基期销售额($)', '销量YOY', '销售额YOY', '本期加权成交均价($)', '期间状态'], annual_out, [18, 10, 34, 34, 16, 16, 18, 18, 12, 14, 20, 16])
for col in [4, 5]: annual_ws.set_column(col, col, 16, fmt['num'])
for col in [6, 7, 10]: annual_ws.set_column(col, col, 18, fmt['money'])
for col in [8, 9]: annual_ws.set_column(col, col, 14, fmt['ratio'])

# 06 BSR tiers
bsr_out = []
for key in ['overall', 'pp', 'high', 'genimo', 'genimoPP']:
    for band in ['head', 'middle', 'tail']:
        for r in MODEL['bsr'][key]['tiers'][band]:
            bsr_out.append([{'overall':'整体市场','pp':'PP市场','high':'高客单价市场','genimo':'Genimo品牌','genimoPP':'Genimo PP市场'}[key], r['month'], r['bsrTier'], r['sales'], r['revenue'], r['avgPrice'], r['weightedPrice'], r['bsrValueCount'], r['observationCount'], r['momSales'], r['yoySales'], r['momRevenue'], r['yoyRevenue']])
bsr_ws = write_table('06_BSR分层', ['范围', '月份', 'BSR档位', '月销量', '月销售额($)', '平均标价($)', '加权成交均价($)', 'BSR值组数', 'BSR观察记录数', '销量MOM', '销量YOY', '销售额MOM', '销售额YOY'], bsr_out, [18, 12, 16, 16, 18, 16, 18, 12, 16, 12, 12, 14, 14])
for col in [3, 7, 8]: bsr_ws.set_column(col, col, 16, fmt['int'])
for col in [4, 5, 6]: bsr_ws.set_column(col, col, 18, fmt['money'])
for col in [9, 10, 11, 12]: bsr_ws.set_column(col, col, 14, fmt['ratio'])

# 07—09
conflict_out = [[p['month'], p['parentKey'], p['candidateRows'], p['qValidRows'], p['qMin'], p['qMax'], p['qUnique'], p['tValidRows'], p['tMin'], p['tMax'], p['tUnique'], 'Q多值' if p['qConflict'] else '', 'T多值' if p['tConflict'] else '', p['ppStatus'], p['genimoStatus']] for p in parent_rows if p['qConflict'] or p['tConflict'] or p['ppStatus'] == 'TIE' or p['genimoStatus'] == 'TIE']
write_table('07_父体冲突诊断', ['月份', '父体键', '候选行数', 'Q有效行数', 'Q最小值', 'Q最大值', 'Q不同值数', 'T有效行数', 'T最小值', 'T最大值', 'T不同值数', 'Q状态', 'T状态', 'PP状态', 'Genimo状态'], conflict_out, [10, 20, 12, 12, 14, 14, 12, 12, 16, 16, 12, 12, 12, 12, 14])
write_table('08_2027策略', ['范围', '指标', '结果', '2027建议'], [
    ['整体市场', '月度MOM/年度YOY', '见01、05页', '按相同父体和BSR口径持续追踪；比较倍数低于1时先复核覆盖与源值冲突。'],
    ['PP市场', 'BSR三档', '见06页', '按头部、中部、尾部的销量和销售额比较倍数选择小批测款档位。'],
    ['高客单价市场', 'BSR三档', '见06页', '单独观察非PP补集的销量、销售额和均价，不用PP趋势代替。'],
    ['Genimo', '整体/PP份额', '见04、04B页', '以整体份额和PP内份额为基线，连续追踪后再安排链接扩展。'],
], [18, 22, 18, 76])
checks = []
for i, m in enumerate(months):
    a, p, h = MODEL['summaries']['overall'][i], MODEL['summaries']['pp'][i], MODEL['summaries']['high'][i]
    checks.append([m, a['sales'], p['sales'], h['sales'], (p['sales'] or 0) + (h['sales'] or 0), (a['sales'] or 0) - (p['sales'] or 0) - (h['sales'] or 0), a['revenue'], p['revenue'], h['revenue'], (a['revenue'] or 0) - (p['revenue'] or 0) - (h['revenue'] or 0)])
check_ws = write_table('09_勾稽检查', ['月份', '整体销量', 'PP销量', '高客单价销量', 'PP+高客单价销量', '销量差额', '整体销售额($)', 'PP销售额($)', '高客单价销售额($)', '销售额差额'], checks, [12, 16, 16, 18, 22, 14, 18, 16, 22, 16])

# 93—94
write_table('93_来源登记', ['月份', '源文件', '业务Sheet', '原始行数', '候选行数', 'BSR观察记录数', '无效BSR行', '100外BSR行', 'SHA-256'], [[s['month'], s['file'], s['sheet'], s['rawRows'], s['candidateRows'], s['eligibleObservationRows'], s['exclusions']['INVALID_BSR'], s['exclusions']['OUTSIDE_TOP100'], s['sha256']] for s in MODEL['sourceFiles']], [10, 38, 24, 12, 12, 16, 14, 14, 68])
write_table('94_规则说明', ['规则项', '当前规则'], [
    ['候选池', '小类目不参与筛选；小类BSR拆分后任一有效1—100排名入池，包含1和100，多值全部保留。'],
    ['父体分组', '月份+父ASIN；父ASIN缺失时回退ASIN，再回退源行ID。'],
    ['父体销量', '同一月份+父ASIN内，Q列月销量有效非负值取平均。'],
    ['父体销售额', '同一月份+父ASIN内，T列月销售额($)有效非负值取平均。'],
    ['BSR汇总', '月份+BSR值，忽略ASIN；同月同BSR的Q/T取平均，再按1—20、21—50、51—100三档。'],
    ['PP/高客单价', '父体候选标题中的plastic命中率≥50%归PP，其他归高客单价。'],
    ['Genimo', '父体有效品牌中Genimo占多数归Genimo；整体份额与PP内份额分别计算。'],
    ['MOM', '本月÷去年同月，不减1。'],
    ['YOY', '本年与上一年共同覆盖月份汇总相除，不减1；2025比较8—12月，2026比较1—7月。'],
    ['均价', '加权成交均价=销售额÷销量；空白不补零。'],
    ['利润', '不计算利润；毛利率、FBA、Coupon仅保留源字段回勾。'],
], [20, 110])

# Keep the reader-facing outputs first, followed by source and audit tabs.
sheet_order = ['00_数据总览', '01_整体市场月度', '02_PP市场月度', '03_高客单价市场月度', '04_Genimo品牌月度', '04B_Genimo_PP月度', '05_年度YOY', '06_BSR分层', '07_父体冲突诊断', '08_2027策略', '09_勾稽检查', '90_原始输入', '91_BSR候选子体明细', '92_父体月度汇总', '93_来源登记', '94_规则说明']
wb.worksheets_objs.sort(key=lambda ws: sheet_order.index(ws.name))

wb.close()
print(json.dumps({'status': 'BUILT', 'file': str(OUT), 'bytes': OUT.stat().st_size, 'sheets': 16, 'candidateRows': len(cand_rows), 'parentRows': len(parent_rows)}, ensure_ascii=False, indent=2))
