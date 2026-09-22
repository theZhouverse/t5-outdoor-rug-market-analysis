import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import XLSX from 'xlsx';
import { buildModel, categoryBsr, aggregateParent } from '../src/build_parent_source_model.mjs';

const ROOT = process.cwd();
const OUT = path.join(ROOT, 'outputs', '20260920-new-source-parent-model');
const SOURCE = path.join(ROOT, 'data', 'raw', '0920new');
const model = buildModel();
const saved = JSON.parse(fs.readFileSync(path.join(OUT, 'model.json'), 'utf8'));
const publicJson = JSON.parse(fs.readFileSync(path.join(OUT, '户外地垫市场分析-SPEC3-父体口径.json'), 'utf8'));
const html = fs.readFileSync(path.join(OUT, '户外地垫市场分析-SPEC3-父体口径.html'), 'utf8');
const close = (a, b) => (a == null && b == null) || (Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) < 1e-6);

// Source identity, continuity and independently recomputed Outdoor Rugs scope.
assert.deepEqual(model.months, Array.from({ length: 24 }, (_, i) => {
  const d = new Date(2024, 7 + i, 1);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
}));
assert.equal(model.sourceFiles.length, 24);
assert.equal(model.rawRows.length, 48000);
assert.equal(model.metadata.candidateRowCount, 45935);
assert.equal(model.metadata.parentMonthCount, 1272);
assert.equal(model.metadata.ambiguousCategoryRankRows, 354);
assert.deepEqual(model.metadata.ambiguousMonths, ['202508']);
assert.equal(model.duplicateAsins.length, 0);

let independentRaw = 0;
let independentEligible = 0;
let independentAmbiguous = 0;
for (const sf of model.sourceFiles) {
  const file = path.join(SOURCE, sf.file);
  const wb = XLSX.readFile(file, { cellDates: false, cellText: false, raw: true });
  const ws = wb.Sheets[sf.sheet];
  const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  const headers = matrix[0].map((x) => String(x ?? '').trim());
  const ci = headers.indexOf('小类目');
  const bi = headers.indexOf('小类BSR');
  assert.ok(ci >= 0 && bi >= 0, `${sf.file}: category headers`);
  let fileEligible = 0;
  let fileAmbiguous = 0;
  for (const row of matrix.slice(1)) {
    const parsed = categoryBsr(row[ci], row[bi]);
    independentRaw += 1;
    if (parsed.bsrStatus === 'ELIGIBLE') { independentEligible += 1; fileEligible += 1; }
    if (parsed.bsrStatus === 'AMBIGUOUS_CATEGORY_RANK') { independentAmbiguous += 1; fileAmbiguous += 1; }
  }
  assert.equal(fileEligible, sf.candidateRows, `${sf.month}: independently recomputed candidates`);
  assert.equal(fileAmbiguous, sf.exclusions.AMBIGUOUS_CATEGORY_RANK, `${sf.month}: independently recomputed ambiguity`);
}
assert.equal(independentRaw, 48000);
assert.equal(independentEligible, 45935);
assert.equal(independentAmbiguous, 354);

// Saved model and public report data must come from the same current build.
assert.equal(saved.metadata.batchId, model.metadata.batchId);
assert.equal(saved.candidateRows.length, 45935);
assert.equal(saved.parentRows.length, 1272);
assert.equal(publicJson.metadata.batchId, model.metadata.batchId);
assert.equal(publicJson.metadata.candidateRowCount, 45935);
assert.equal(publicJson.metadata.parentMonthCount, 1272);
assert.ok(!('candidateRows' in publicJson), 'public JSON should omit detailed child rows');
assert.ok(html.includes('id="dashboard"'));
assert.ok(html.includes('id="report-data"'));
assert.ok(html.includes('2025.08'));
assert.ok(html.includes('Outdoor Rugs'));
assert.ok(html.includes('严格多数'));

// Parent partition, tier partition and blank semantics.
for (let i = 0; i < model.months.length; i += 1) {
  const month = model.months[i];
  const all = model.summaries.overall[i];
  const pp = model.summaries.pp[i];
  const high = model.summaries.high[i];
  assert.equal(all.parentCount, pp.parentCount + high.parentCount, `${month}: parent addback`);
  assert.ok(close(all.parentSales, (pp.parentSales ?? 0) + (high.parentSales ?? 0)), `${month}: sales addback`);
  assert.ok(close(all.childRevenue, (pp.childRevenue ?? 0) + (high.childRevenue ?? 0)), `${month}: revenue addback`);
  const coarse = model.tiers.overall.filter((t) => t.type === 'coarse').reduce((s, t) => s + t.rows[i].parentCount, 0);
  const fine = model.tiers.overall.filter((t) => t.type === 'fine').reduce((s, t) => s + t.rows[i].parentCount, 0);
  assert.equal(coarse, all.parentCount, `${month}: coarse tiers`);
  assert.equal(fine, all.parentCount, `${month}: fine tiers`);
}
assert.ok(model.parentRows.every((p) => p.band && p.fineBand), 'fractional BSR medians must belong to a tier');
const noRevenue = model.parentRows.find((p) => p.childStatus === 'NONE');
assert.ok(noRevenue && noRevenue.childSales == null && noRevenue.childRevenue == null && noRevenue.childAsp == null);

// Rule edge cases: exact 50/50 is median conflict; strict majority uses mode.
const base = (sales, i) => ({
  month: '202607', parentKey: 'TEST', parent: 'TEST', asin: `A${i}`, sourceRow: i + 2, rank: 10,
  sales, sourceRevenue: null, childSales: null, childRevenue: null, pp: false, genimo: false,
  title: 'rug', brand: 'brand'
});
const tie = aggregateParent('202607', 'TEST', [base(100, 0), base(200, 1)]);
assert.equal(tie.salesStatus, 'MEDIAN_CONFLICT');
assert.equal(tie.parentSales, 150);
const majority = aggregateParent('202607', 'TEST', [base(100, 0), base(100, 1), base(200, 2)]);
assert.equal(majority.salesStatus, 'MODE');
assert.equal(majority.parentSales, 100);

// Traceable real parent sample.
const sample = model.parentRows.find((p) => p.month === '202607' && p.parentKey === 'B0BM74444Q');
assert.ok(sample);
assert.equal(sample.candidateRows, 69);
assert.equal(sample.salesMode, 3901);
assert.equal(sample.salesModeCount, 39);
assert.ok(close(sample.salesModeShare, 39 / 69));
assert.equal(sample.salesStatus, 'MODE');
assert.equal(sample.parentSales, 3901);
assert.equal(sample.salesMin, 3813);
assert.equal(sample.salesMax, 6468);
assert.equal(sample.rankMedian, 21);
assert.equal(sample.childValidRows, 29);
assert.equal(sample.childSales, 3100);
assert.equal(sample.childRevenue, 154430);
assert.ok(close(sample.childAsp, 154430 / 3100));

// Annual YOY uses only matched months and marks affected 2025 comparisons.
for (const rows of Object.values(model.annual)) {
  for (const r of rows) {
    assert.equal(r.months.length, r.priorMonths.length);
    assert.ok(r.months.every((m, i) => Number(m.slice(0, 4)) - 1 === Number(r.priorMonths[i].slice(0, 4)) && m.slice(4) === r.priorMonths[i].slice(4)));
  }
}
assert.equal(model.annual.overall.find((r) => r.year === '2025').scopeLimited, true);

// Workbook topology, formulas and cached result parity with the current model.
const book = path.join(OUT, '户外地垫市场分析-SPEC3-父体口径可回勾.xlsx');
const wb = XLSX.readFile(book, { cellFormula: true, cellNF: true, cellDates: false });
const expectedSheets = ['00_数据总览','01_整体市场月度','02_PP市场月度','03_高客单价月度','04_Genimo品牌月度','04B_Genimo_PP月度','05_年度YOY','06_BSR分层','07_父体冲突诊断','08_2027策略','09_勾稽检查','90_原始输入','91_BSR候选子体明细','92_父体月度汇总','93_来源登记','94_规则说明'];
assert.deepEqual(wb.SheetNames, expectedSheets);
const monthlySheets = { overall:'01_整体市场月度', pp:'02_PP市场月度', high:'03_高客单价月度', genimo:'04_Genimo品牌月度', genimoPP:'04B_Genimo_PP月度' };
for (const [key, sheet] of Object.entries(monthlySheets)) {
  const ws = wb.Sheets[sheet];
  for (let i = 0; i < 24; i += 1) {
    const r = i + 2;
    const expected = model.summaries[key][i];
    assert.equal(ws[`A${r}`].v, expected.month);
    assert.equal(ws[`B${r}`].v, expected.parentCount, `${sheet} ${r} count`);
    assert.ok(close(ws[`D${r}`].v, expected.parentSales), `${sheet} ${r} sales`);
    assert.ok(close(ws[`F${r}`].v, expected.childRevenue), `${sheet} ${r} revenue`);
    assert.ok(close(ws[`G${r}`].v, expected.childCoverage), `${sheet} ${r} coverage`);
    assert.equal(ws[`U${r}`].v, expected.ambiguousBsrRows, `${sheet} ${r} scope exclusions`);
    assert.ok(ws[`B${r}`].f?.includes('COUNTIFS'));
    assert.ok(ws[`F${r}`].f?.includes('SUMIFS'));
  }
}
const parent = wb.Sheets['92_父体月度汇总'];
const sampleWorkbookRow = model.parentRows.indexOf(sample) + 2;
assert.ok(parent.B2 && parent.J2?.f && parent[`T${sampleWorkbookRow}`]?.f && parent[`W${sampleWorkbookRow}`]?.f);
assert.ok(parent.J2.f.includes('>0.5'));
assert.ok(parent[`T${sampleWorkbookRow}`].f.includes('SUMIFS'));
assert.ok(parent[`W${sampleWorkbookRow}`].f.includes(`/S${sampleWorkbookRow}`));
assert.ok(parent.E2.f.includes('COUNTA'));
for (let i = 0; i < model.parentRows.length; i += 1) {
  const p = model.parentRows[i];
  const r = i + 2;
  assert.equal(parent[`A${r}`].v, p.month);
  assert.equal(parent[`B${r}`].v, p.parentKey);
  assert.ok(close(parent[`J${r}`]?.v ?? null, p.parentSales), `parent sales ${p.month}/${p.parentKey}`);
  assert.ok(close(parent[`T${r}`]?.v ?? null, p.childRevenue), `parent revenue ${p.month}/${p.parentKey}`);
  assert.ok(close(parent[`W${r}`]?.v ?? null, p.childAsp), `parent ASP ${p.month}/${p.parentKey}`);
}
for (const sheet of wb.SheetNames) {
  const ws = wb.Sheets[sheet];
  for (const address of Object.keys(ws).filter((x) => !x.startsWith('!'))) {
    assert.ok(!['#REF!','#DIV/0!','#VALUE!','#N/A'].includes(ws[address].v), `${sheet}!${address}: cached error`);
  }
}

const result = {
  status: 'PASS',
  source: { files: 24, months: '2024.08—2026.07', rawRows: 48000, candidateRows: 45935, ambiguousCategoryRankRows: 354, ambiguousMonth: '2025.08' },
  model: { parentMonthCount: 1272, duplicateAsinGroups: 0, conflictParentCount: model.metadata.conflictParentCount, medianParentCount: model.metadata.medianParentCount, missingRevenueParents: model.metadata.missingRevenueParents, mixedMarketParents: model.metadata.mixedMarketParents },
  sample: { month: sample.month, parent: sample.parentKey, sales: sample.parentSales, salesStatus: sample.salesStatus, childRevenue: sample.childRevenue, childCoverage: sample.childCoverage },
  workbook: { sheets: wb.SheetNames.length, expectedSheets: true, formulasPresent: true, cachedValuesMatchModel: true },
  html: { embeddedDataMatchesBatch: true, navigationAndDisclosurePresent: true }
};
fs.writeFileSync(path.join(OUT, '完整审计结果.json'), JSON.stringify(result, null, 2), 'utf8');
console.log(JSON.stringify(result, null, 2));
