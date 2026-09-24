"""Independent formula/cache audit for the SPEC 3.7 workbook.

This deliberately evaluates the saved formulas from the workbook XML through the
bounded local formula engine. It does not claim to be a desktop Excel recalc.
"""
import json
import math
import os
from pathlib import Path

import openpyxl

import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
from formula_engine import Engine, Error, isnum


class Spec37Engine(Engine):
    def eval(self, node, sheet):
        if node[0] == 'fn' and node[1] == 'IFERROR':
            try:
                return self.eval(node[2][0], sheet)
            except Error:
                return self.eval(node[2][1], sheet)
        return super().eval(node, sheet)


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'outputs' / '20260920-new-source-parent-model'
BOOK = Path(os.environ.get('SPEC37_XLSX_IN', str(OUT / '户外地垫市场分析-SPEC3.7-父体口径.xlsx')))
RESULT = Path(os.environ.get('SPEC37_FORMULA_AUDIT_RESULT', str(OUT / '公式审计-spec37.json')))
SKIP = {'90_原始输入', '93_来源登记', '94_规则说明', '00_数据总览'}


def same(a, b):
    if a in (None, '') and b in (None, ''):
        return True
    if isnum(a) and isnum(b):
        return math.isclose(a, b, abs_tol=1e-7, rel_tol=1e-10)
    return a == b


wf = openpyxl.load_workbook(BOOK, read_only=True, data_only=False)
wc = openpyxl.load_workbook(BOOK, read_only=True, data_only=True)
sheets = {}
formula_cells = []
for ws in wf.worksheets:
    if ws.title in SKIP:
        continue
    cached_ws = wc[ws.title]
    cells = {}
    for row, cached_row in zip(ws.iter_rows(), cached_ws.iter_rows()):
        for cell, cached in zip(row, cached_row):
            if cell.value is None:
                continue
            key = (cell.row, cell.column)
            if cell.data_type == 'f':
                cells[key] = {'f': cell.value}
                formula_cells.append((ws.title, cell.row, cell.column, cached.value))
            else:
                cells[key] = cell.value
    sheets[ws.title] = cells
wf.close()
wc.close()

engine = Spec37Engine(sheets)
failures = []
for sheet, row, col, cached in formula_cells:
    try:
        actual = engine.get(sheet, row, col)
    except Exception as exc:  # fail closed with the exact cell and formula
        failures.append({'sheet': sheet, 'row': row, 'col': col, 'error': str(exc)})
        if len(failures) >= 12:
            break
        continue
    if not same(actual, cached):
        failures.append({'sheet': sheet, 'row': row, 'col': col, 'actual': actual, 'cached': cached})
        if len(failures) >= 12:
            break
assert not failures, json.dumps(failures, ensure_ascii=False)

# Input perturbations exercise the parent AVERAGEIFS rules and the downstream
# monthly summary. These checks are in memory and never rewrite the delivery XLSX.
detail = sheets['91_BSR候选子体明细']
parent = sheets['92_父体月度汇总']
monthly = sheets['01_整体市场月度']
detail_row = next(row for (row, col), value in detail.items() if col == 1 and value == '202607' and detail.get((row, 2)) == 'B0BM74444Q')
parent_row = next(row for (row, col), value in parent.items() if col == 1 and value == '202607' and parent.get((row, 2)) == 'B0BM74444Q')
monthly_row = next(row for (row, col), value in monthly.items() if col == 1 and value == '202607')
candidate_key = detail.get((detail_row, 2))
candidate_rows = sorted({row for (row, col), value in detail.items() if col == 2 and value == candidate_key and detail.get((row, 1)) == '202607'})


def reset_engine():
    engine.cache.clear()
    engine.busy.clear()
    engine.index.clear()


def expected_average(detail_col):
    values = [detail.get((row, detail_col)) for row in candidate_rows]
    valid = [value for value in values if isnum(value) and value >= 0]
    return sum(valid) / len(valid) if valid else None


def verify_mutation(detail_col, parent_col, monthly_col, replacement):
    old_value = detail.get((detail_row, detail_col))
    reset_engine()
    baseline_parent = engine.get('92_父体月度汇总', parent_row, parent_col)
    baseline_month = engine.get('01_整体市场月度', monthly_row, monthly_col)

    detail[(detail_row, detail_col)] = replacement
    reset_engine()
    expected_parent = expected_average(detail_col)
    actual_parent = engine.get('92_父体月度汇总', parent_row, parent_col)
    assert same(actual_parent, expected_parent), (detail_col, replacement, actual_parent, expected_parent)
    expected_month = baseline_month - baseline_parent + expected_parent
    actual_month = engine.get('01_整体市场月度', monthly_row, monthly_col)
    assert same(actual_month, expected_month), (detail_col, replacement, actual_month, expected_month)
    sales = engine.get('01_整体市场月度', monthly_row, 2)
    revenue = engine.get('01_整体市场月度', monthly_row, 3)
    weighted = engine.get('01_整体市场月度', monthly_row, 5)
    assert same(weighted, revenue / sales), (detail_col, replacement, weighted, revenue / sales)

    detail[(detail_row, detail_col)] = old_value
    reset_engine()
    assert same(engine.get('92_父体月度汇总', parent_row, parent_col), baseline_parent)
    assert same(engine.get('01_整体市场月度', monthly_row, monthly_col), baseline_month)


for replacement in (-100, None, 0):
    verify_mutation(9, 9, 2, replacement)
    verify_mutation(10, 10, 3, replacement)

result = {
    'status': 'PASS',
    'workbook': str(BOOK.relative_to(ROOT)) if BOOK.is_relative_to(ROOT) else str(BOOK),
    'savedFormulasEvaluated': len(formula_cells),
    'engineEvaluations': engine.count,
    'cachedMismatches': 0,
    'inputChanges': [
        'negative Q excluded and monthly sales recalculated',
        'blank Q excluded and monthly sales recalculated',
        'explicit zero Q included and monthly sales recalculated',
        'negative T excluded and monthly revenue recalculated',
        'blank T excluded and monthly revenue recalculated',
        'explicit zero T included and monthly revenue recalculated',
    ],
    'engine': 'bounded independent evaluator; not desktop Excel',
}
RESULT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(result, ensure_ascii=False))
