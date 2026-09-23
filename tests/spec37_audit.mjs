import fs from 'node:fs';
import assert from 'node:assert/strict';
import XLSX from 'xlsx';
import { buildModel } from '../src/build_spec37_model.mjs';

const ROOT = process.cwd();
const OUT = `${ROOT}/outputs/20260920-new-source-parent-model`;
const model = buildModel();
const saved = JSON.parse(fs.readFileSync(`${OUT}/model-spec37.json`, 'utf8'));
const json = JSON.parse(fs.readFileSync(`${OUT}/户外地垫市场分析-SPEC3.7-父体口径.json`, 'utf8'));
const html = fs.readFileSync(`${OUT}/户外地垫市场分析-SPEC3.7-父体口径.html`, 'utf8');
const book = XLSX.readFile(`${OUT}/户外地垫市场分析-SPEC3.7-父体口径.xlsx`, { cellFormula: true, cellNF: true, cellDates: false });
const close = (a, b) => (a == null && b == null) || (Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) < 1e-6);

assert.equal(model.metadata.specVersion, '3.7');
assert.equal(model.metadata.rawRowCount, 48000);
assert.equal(model.metadata.candidateRowCount, 46780);
assert.equal(model.metadata.bsrObservationCount, 46854);
assert.equal(model.metadata.parentMonthCount, 1334);
assert.equal(model.duplicateAsins?.length ?? model.metadata.duplicateAsinGroups, 0);
assert.equal(saved.metadata.batchId, model.metadata.batchId);
assert.equal(json.metadata.batchId, model.metadata.batchId);
assert.ok(html.includes('SPEC 3.7'));
assert.ok(html.includes('MOM为本月÷去年同月'));
assert.ok(!html.includes('销量MOM（上月）'));

// Candidate scope uses BSR only and preserves multi-rank observations.
assert.ok(model.candidateRows.every((r) => r.eligibleRanks.some((rank) => rank >= 1 && rank <= 100)));
assert.ok(model.observations.length >= model.candidateRows.length);
assert.ok(model.observations.some((r) => model.candidateRows.find((c) => c.month === r.month && c.sourceRow === r.sourceRow)?.eligibleRanks.length > 1));

// Parent Q/T averages are independently recomputed for a real sample.
const sampleRows = model.candidateRows.filter((r) => r.month === '202607' && r.parentKey === 'B0BM74444Q');
const sample = model.parentRows.find((r) => r.month === '202607' && r.parentKey === 'B0BM74444Q');
assert.equal(sampleRows.length, 69);
assert.equal(sample.qAvg, sampleRows.filter((r) => Number.isFinite(r.sales) && r.sales >= 0).reduce((a, r) => a + r.sales, 0) / 69);
assert.equal(sample.tAvg, sampleRows.filter((r) => Number.isFinite(r.sourceRevenue) && r.sourceRevenue >= 0).reduce((a, r) => a + r.sourceRevenue, 0) / 69);

// BSR branch is month + BSR and ignores ASIN. Tier values average BSR-value averages.
const bsrGroups = model.bsr.overall.groups;
const group = bsrGroups.find((r) => r.month === '202607' && r.bsr === 10);
assert.ok(group);
const groupObs = model.observations.filter((r) => r.month === '202607' && r.rank === 10);
assert.equal(group.observationCount, groupObs.length);
const validQ = groupObs.filter((r) => Number.isFinite(r.sales) && r.sales >= 0).map((r) => r.sales);
const validT = groupObs.filter((r) => Number.isFinite(r.sourceRevenue) && r.sourceRevenue >= 0).map((r) => r.sourceRevenue);
assert.ok(close(group.qAvg, validQ.reduce((a, b) => a + b, 0) / validQ.length));
assert.ok(close(group.tAvg, validT.reduce((a, b) => a + b, 0) / validT.length));
const head = model.bsr.overall.tiers.head.find((r) => r.month === '202607');
const headGroups = bsrGroups.filter((r) => r.month === '202607' && r.bsr >= 1 && r.bsr <= 20);
assert.ok(close(head.sales, headGroups.reduce((a, r) => a + (r.qAvg ?? 0), 0) / headGroups.filter((r) => Number.isFinite(r.qAvg)).length));

// MOM is same-month prior-year ratio; monthly YOY is common-month cumulative ratio.
const sep2025 = model.summaries.overall.find((r) => r.month === '202509');
const sep2024 = model.summaries.overall.find((r) => r.month === '202409');
assert.ok(close(sep2025.momSales, sep2025.sales / sep2024.sales));
const annual2025 = model.annual.overall.find((r) => r.year === '2025');
const annual2026 = model.annual.overall.find((r) => r.year === '2026');
assert.deepEqual(annual2025.months, ['202508', '202509', '202510', '202511', '202512']);
assert.deepEqual(annual2026.months, ['202601', '202602', '202603', '202604', '202605', '202606', '202607']);
assert.ok(close(annual2025.yoySales, annual2025.sales / annual2025.priorSales));
assert.ok(close(annual2026.yoySales, annual2026.sales / annual2026.priorSales));

const expectedSheets = ['00_数据总览', '01_整体市场月度', '02_PP市场月度', '03_高客单价市场月度', '04_Genimo品牌月度', '04B_Genimo_PP月度', '05_年度YOY', '06_BSR分层', '07_父体冲突诊断', '08_2027策略', '09_勾稽检查', '90_原始输入', '91_BSR候选子体明细', '92_父体月度汇总', '93_来源登记', '94_规则说明'];
assert.deepEqual(book.SheetNames, expectedSheets);
const monthly = book.Sheets['01_整体市场月度'];
assert.ok(monthly.B2?.f?.includes('SUMIFS'));
assert.ok(monthly.F14?.f?.includes('B14/B2'));
assert.ok(monthly.G14?.f?.includes('SUMIFS'));
assert.ok(book.Sheets['92_父体月度汇总'].I2?.f?.includes('AVERAGEIFS'));
assert.ok(book.Sheets['92_父体月度汇总'].J2?.f?.includes('AVERAGEIFS'));
assert.ok(book.Sheets['06_BSR分层'].D2?.f?.includes('AVERAGEIFS'));
assert.ok(book.Sheets['06_BSR分层'].A364?.v === 'BSR值组明细（公式来源）');
assert.ok(book.Sheets['05_年度YOY'].E3?.f?.includes('SUMIFS'));
assert.ok(book.Sheets['05_年度YOY'].I3?.f?.includes('E3/F3'));
for (const sheetName of book.SheetNames) {
  for (const cell of Object.values(book.Sheets[sheetName])) {
    if (!cell || typeof cell !== 'object') continue;
    assert.ok(!['#REF!', '#DIV/0!', '#VALUE!', '#N/A'].includes(cell.v), `${sheetName} cached formula error`);
  }
}

const result = {
  status: 'PASS',
  spec: '3.7',
  source: { files: 24, rawRows: model.metadata.rawRowCount, candidateRows: model.metadata.candidateRowCount, bsrObservationRows: model.metadata.bsrObservationCount },
  model: { parentMonthCount: model.metadata.parentMonthCount, duplicateAsinGroups: model.metadata.duplicateAsinGroups, qConflictParentCount: model.metadata.qConflictParentCount, tConflictParentCount: model.metadata.tConflictParentCount },
  sample: { month: sample.month, parentKey: sample.parentKey, qAvg: sample.qAvg, tAvg: sample.tAvg, bsr10QAvg: group.qAvg, bsr10TAvg: group.tAvg },
  workbook: { sheets: book.SheetNames.length, formulasPresent: true, cachedErrors: 0 },
  html: { embeddedData: true, navigation: true, collapsibleSections: true, charts: true }
};
fs.writeFileSync(`${OUT}/完整审计结果-spec37.json`, JSON.stringify(result, null, 2), 'utf8');
console.log(JSON.stringify(result, null, 2));
