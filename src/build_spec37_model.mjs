import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import XLSX from 'xlsx';
import { pathToFileURL } from 'node:url';

const ROOT = process.cwd();
const SOURCE_DIR = path.join(ROOT, 'data', 'raw', '0920new');
const OUT_DIR = path.join(ROOT, 'outputs', '20260920-new-source-parent-model');
const MONTH_RE = /^Competitor-US-(\d{4})\.(\d{2})-/i;
const SOURCE_FILES = fs.readdirSync(SOURCE_DIR)
  .filter((f) => /^Competitor-US-\d{4}\.\d{2}-.*\.xlsx$/i.test(f))
  .sort();

const HEADER_NAMES = [
  'ASIN', 'SKU', '详细参数', '品牌', '品牌链接', '商品标题', '商品详情页链接', '商品主图', '父ASIN',
  '类目路径', '大类目', '大类BSR', '大类BSR增长数', '大类BSR增长率', '小类目', '小类BSR',
  '月销量', '销量环比增长率', '销量同比增长率', '月销售额($)', '子体销量', '子体销售额($)', '变体数',
  '价格($)', 'prime价格($)', 'Coupon', 'Q&A', '评分数', '月新增评分数', '评分', '留评率', 'FBA($)',
  '毛利率', '评级', '上架时间', '上架天数', '配送方式', '买家运费($)', 'LQS', '卖家数', 'BuyBox卖家',
  'BuyBox类型', '卖家所属地', '卖家信息', '卖家首页', 'Best Seller标识', "Amazon's Choice", 'New Release标识',
  'A+页面', '视频介绍', 'SP广告', '品牌故事', '品牌广告', '秒杀', 'AC关键词', '商品重量', '商品重量（单位换算）',
  '商品尺寸', '商品尺寸（单位换算）', '包装重量', '包装重量（单位换算）', '包装尺寸', '包装尺寸（单位换算）', '包装尺寸分段', '标签'
];
const SELECTED_FIELDS = ['ASIN', '品牌', '商品标题', '父ASIN', '小类目', '小类BSR', '月销量', '月销售额($)', '子体销量', '子体销售额($)', '价格($)'];
const BANDS = [
  { key: 'head', name: '头部 1—20', lo: 1, hi: 20 },
  { key: 'middle', name: '中部 21—50', lo: 21, hi: 50 },
  { key: 'tail', name: '尾部 51—100', lo: 51, hi: 100 }
];

const text = (v) => v == null ? '' : String(v).replace(/\u00a0/g, ' ').replace(/\r?\n/g, ' ').trim();
const number = (v) => {
  const s = text(v);
  if (!s || /^[-—–]$/.test(s) || /^(n\/a|na|null|undefined)$/i.test(s)) return null;
  const cleaned = s.replace(/[$,%￥,\s]/g, '').replace(/^\((.*)\)$/, '-$1');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
};
const valid = (v) => Number.isFinite(v) && v >= 0;
const sum = (values) => { const xs = values.filter(Number.isFinite); return xs.length ? xs.reduce((a, b) => a + b, 0) : null; };
const avg = (values) => { const xs = values.filter(Number.isFinite); return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null; };
const ratio = (a, b) => Number.isFinite(a) && Number.isFinite(b) && b > 0 ? a / b : null;
const sha256 = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
const parseMonth = (file) => { const m = file.match(MONTH_RE); return m ? `${m[1]}${m[2]}` : null; };
const parseRank = (v) => {
  const s = text(v).replace(/,/g, '');
  if (!/^\d+(?:\.0+)?$/.test(s)) return null;
  const n = Number(s);
  return Number.isInteger(n) && n > 0 ? n : null;
};
const parseRanks = (v) => [...new Set(String(v ?? '').split(/\r?\n/).map(parseRank).filter(Number.isInteger))];
const tierFor = (rank) => BANDS.find((b) => rank >= b.lo && rank <= b.hi)?.key ?? null;
const isPlastic = (v) => /\bplastic\b/i.test(text(v));
const isGenimo = (v) => text(v).toLowerCase() === 'genimo';

function categoryByMajority(rows, field) {
  const values = rows.map((r) => r[field]).filter((v) => v !== null);
  const yes = values.filter(Boolean).length;
  const no = values.length - yes;
  if (!values.length) return { value: false, status: 'NO_TEXT', yes, no, ratio: null };
  if (yes === no) return { value: yes > 0, status: 'TIE', yes, no, ratio: yes / values.length };
  return { value: yes > no, status: 'MAJORITY', yes, no, ratio: yes / values.length };
}

function readSourceFile(file) {
  const month = parseMonth(file);
  if (!month) throw new Error(`无法解析月份: ${file}`);
  const filePath = path.join(SOURCE_DIR, file);
  const wb = XLSX.readFile(filePath, { cellDates: false, cellNF: true, cellText: false });
  const sheets = wb.SheetNames.filter((s) => /^Competitor-US-\d{6}$/i.test(s));
  if (sheets.length !== 1) throw new Error(`${file}业务Sheet数量不是1: ${sheets.join(',')}`);
  const sheet = sheets[0];
  if (sheet !== `Competitor-US-${month}`) throw new Error(`${file} Sheet月份不一致: ${sheet} != Competitor-US-${month}`);
  const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null, raw: true });
  const headers = (matrix[0] || []).map(text);
  if (headers.join('|') !== HEADER_NAMES.join('|')) throw new Error(`${file}表头顺序或数量变化`);
  const idx = Object.fromEntries(HEADER_NAMES.map((h) => [h, headers.indexOf(h)]));
  const rows = [];
  for (let i = 1; i < matrix.length; i += 1) {
    const cells = matrix[i] || [];
    const get = (h) => cells[idx[h]];
    const rawBsr = get('小类BSR');
    const sourceBsr = text(rawBsr);
    const parsedRanks = parseRanks(rawBsr);
    const eligibleRanks = parsedRanks.filter((r) => r >= 1 && r <= 100);
    const bsrStatus = eligibleRanks.length ? 'ELIGIBLE' : parsedRanks.length ? 'OUTSIDE_TOP100' : 'INVALID_BSR';
    const row = {
      month, file, sheet, sourceRow: i + 1,
      asin: text(get('ASIN')), parent: text(get('父ASIN')), parentKey: text(get('父ASIN')) || text(get('ASIN')) || `row:${i + 1}`,
      brand: text(get('品牌')), title: text(get('商品标题')), category: text(get('小类目')), sourceBsr,
      parsedRanks, eligibleRanks, bsrStatus, multiBsr: parsedRanks.length > 1,
      sales: number(get('月销量')), sourceRevenue: number(get('月销售额($)')),
      childSales: number(get('子体销量')), childRevenue: number(get('子体销售额($)')), price: number(get('价格($)')),
      pp: text(get('商品标题') ?? '') ? isPlastic(get('商品标题')) : null,
      genimo: text(get('品牌') ?? '') ? isGenimo(get('品牌')) : null,
      raw: Object.fromEntries(SELECTED_FIELDS.map((h) => [h, cells[idx[h]] ?? null]))
    };
    rows.push(row);
  }
  return {
    month, file, sheet, sha256: sha256(filePath), rows,
    rawRows: rows.length,
    candidateRows: rows.filter((r) => r.eligibleRanks.length > 0).length,
    eligibleObservationRows: rows.reduce((n, r) => n + r.eligibleRanks.length, 0),
    exclusions: {
      INVALID_BSR: rows.filter((r) => r.bsrStatus === 'INVALID_BSR').length,
      OUTSIDE_TOP100: rows.filter((r) => r.bsrStatus === 'OUTSIDE_TOP100').length
    }
  };
}

function aggregateParent(month, parentKey, rows) {
  const q = rows.map((r) => r.sales).filter(valid);
  const t = rows.map((r) => r.sourceRevenue).filter(valid);
  const price = rows.map((r) => r.price).filter(valid);
  const childRows = rows.filter((r) => valid(r.childSales) && valid(r.childRevenue) && !(r.childSales === 0 && r.childRevenue > 0));
  const pp = categoryByMajority(rows, 'pp');
  const genimo = categoryByMajority(rows, 'genimo');
  return {
    month, parentKey, parent: rows.find((r) => r.parent)?.parent || '',
    asinCount: new Set(rows.map((r) => r.asin).filter(Boolean)).size, candidateRows: rows.length,
    qAvg: avg(q), tAvg: avg(t), avgPrice: avg(price), weightedPrice: ratio(avg(t), avg(q)),
    qValidRows: q.length, tValidRows: t.length, qMin: q.length ? Math.min(...q) : null, qMax: q.length ? Math.max(...q) : null,
    tMin: t.length ? Math.min(...t) : null, tMax: t.length ? Math.max(...t) : null,
    qUnique: new Set(q).size, tUnique: new Set(t).size,
    qStatus: q.length === 0 ? 'NO_VALUE' : q.length === rows.length ? 'FULL' : 'PARTIAL',
    tStatus: t.length === 0 ? 'NO_VALUE' : t.length === rows.length ? 'FULL' : 'PARTIAL',
    qConflict: new Set(q).size > 1, tConflict: new Set(t).size > 1,
    childValidRows: childRows.length, childSales: sum(childRows.map((r) => r.childSales)), childRevenue: sum(childRows.map((r) => r.childRevenue)),
    childCoverage: ratio(childRows.length, rows.length), childStatus: childRows.length === 0 ? 'NONE' : childRows.length === rows.length ? 'FULL' : 'PARTIAL',
    childAsp: ratio(sum(childRows.map((r) => r.childRevenue)), sum(childRows.map((r) => r.childSales))),
    pp: pp.value, ppStatus: pp.status, ppRatio: pp.ratio, ppYes: pp.yes, ppNo: pp.no,
    genimo: genimo.value, genimoStatus: genimo.status, genimoRatio: genimo.ratio,
    sourceRows: rows.map((r) => r.sourceRow), asins: rows.map((r) => r.asin).filter(Boolean)
  };
}

function addComparisons(rows) {
  const byMonth = new Map(rows.map((r) => [r.month, r]));
  for (const r of rows) {
    const priorYear = byMonth.get(`${Number(r.month.slice(0, 4)) - 1}${r.month.slice(4)}`);
    r.momSales = ratio(r.sales, priorYear?.sales);
    r.momRevenue = ratio(r.revenue, priorYear?.revenue);
    const year = Number(r.month.slice(0, 4));
    const monthNumber = Number(r.month.slice(4));
    const currentYtd = rows.filter((x) => Number(x.month.slice(0, 4)) === year && Number(x.month.slice(4)) <= monthNumber && byMonth.has(`${year - 1}${x.month.slice(4)}`));
    const previousYtd = currentYtd.map((x) => byMonth.get(`${year - 1}${x.month.slice(4)}`));
    r.yoySales = ratio(sum(currentYtd.map((x) => x.sales)), sum(previousYtd.map((x) => x.sales)));
    r.yoyRevenue = ratio(sum(currentYtd.map((x) => x.revenue)), sum(previousYtd.map((x) => x.revenue)));
  }
  return rows;
}

function monthSummary(month, parents, filter = () => true) {
  const list = parents.filter(filter);
  const sales = sum(list.map((p) => p.qAvg));
  const revenue = sum(list.map((p) => p.tAvg));
  const childSales = sum(list.map((p) => p.childSales));
  const childRevenue = sum(list.map((p) => p.childRevenue));
  const qParents = list.filter((p) => Number.isFinite(p.qAvg));
  const tParents = list.filter((p) => Number.isFinite(p.tAvg));
  const validChildRows = sum(list.map((p) => p.childValidRows)) ?? 0;
  const candidateRows = sum(list.map((p) => p.candidateRows)) ?? 0;
  return {
    month, parentCount: list.length, candidateRows, sales, revenue, parentSales: sales, parentRevenue: revenue,
    avgPrice: avg(list.map((p) => p.avgPrice)), weightedPrice: ratio(revenue, sales),
    validSalesParents: qParents.length, validRevenueParents: tParents.length,
    qCoverage: ratio(qParents.length, list.length), tCoverage: ratio(tParents.length, list.length),
    childSales, childRevenue, childCoverage: ratio(validChildRows, candidateRows), validChildRows,
    revenueStatus: validChildRows === 0 ? 'NONE' : validChildRows === candidateRows ? 'FULL' : 'PARTIAL',
    qConflictParents: list.filter((p) => p.qConflict).length, tConflictParents: list.filter((p) => p.tConflict).length,
    missingSalesParents: list.length - qParents.length, missingRevenueParents: list.length - tParents.length,
    mixedMarketParents: list.filter((p) => p.ppStatus === 'TIE' || (p.ppRatio > 0 && p.ppRatio < 1)).length
  };
}

function annualRows(rows) {
  const years = [...new Set(rows.map((r) => r.month.slice(0, 4)))].sort();
  return years.map((year) => {
    const available = rows.filter((r) => r.month.startsWith(year));
    const common = available.filter((r) => rows.some((p) => p.month === `${Number(year) - 1}${r.month.slice(4)}`));
    const previous = common.map((r) => rows.find((p) => p.month === `${Number(year) - 1}${r.month.slice(4)}`));
    const sales = sum(common.map((r) => r.sales));
    const revenue = sum(common.map((r) => r.revenue));
    const priorSales = sum(previous.map((r) => r.sales));
    const priorRevenue = sum(previous.map((r) => r.revenue));
    return {
      year, availableMonths: available.map((r) => r.month), months: common.map((r) => r.month), priorMonths: previous.map((r) => r.month),
      sales, revenue, parentSales: sales, parentRevenue: revenue, priorSales, priorRevenue,
      yoySales: ratio(sales, priorSales), yoyRevenue: ratio(revenue, priorRevenue), yoyParentSales: ratio(sales, priorSales), yoyParentRevenue: ratio(revenue, priorRevenue),
      avgPrice: avg(common.map((r) => r.avgPrice)), weightedPrice: ratio(revenue, sales),
      status: !common.length ? 'NO_BASE' : common.length === 12 ? 'FULL_YEAR' : 'MATCHED_MONTHS'
    };
  });
}

function buildBsrGroups(observations, parentsByKey, rowFilter = () => true) {
  const grouped = new Map();
  for (const obs of observations) {
    const parent = parentsByKey.get(`${obs.month}|${obs.parentKey}`);
    if (!parent || !rowFilter(parent)) continue;
    const key = `${obs.month}|${obs.rank}`;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push({ ...obs, parent });
  }
  const rows = [];
  for (const [key, list] of grouped) {
    const [month, rankText] = key.split('|');
    const q = list.map((r) => r.sales).filter(valid);
    const t = list.map((r) => r.sourceRevenue).filter(valid);
    rows.push({
      month, bsr: Number(rankText), tier: tierFor(Number(rankText)), observationCount: list.length,
      qAvg: avg(q), tAvg: avg(t), avgPrice: avg(list.map((r) => r.price).filter(valid)), weightedPrice: ratio(avg(t), avg(q)),
      qValidRows: q.length, tValidRows: t.length, qCoverage: ratio(q.length, list.length), tCoverage: ratio(t.length, list.length),
      ppObservations: list.filter((r) => r.parent.pp).length, genimoObservations: list.filter((r) => r.parent.genimo).length,
      parentKeys: [...new Set(list.map((r) => r.parentKey))]
    });
  }
  return rows.sort((a, b) => a.month.localeCompare(b.month) || a.bsr - b.bsr);
}

function buildBsrTiers(groupRows) {
  const groups = groupRows;
  const months = [...new Set(groups.map((r) => r.month))].sort();
  const tiers = Object.fromEntries(BANDS.map((b) => [b.key, addComparisons(months.map((month) => {
    const rows = groups.filter((r) => r.month === month && r.bsr >= b.lo && r.bsr <= b.hi);
    const sales = avg(rows.map((r) => r.qAvg));
    const revenue = avg(rows.map((r) => r.tAvg));
    return {
      month, bsrTier: b.name, tier: b.key, bsrValueCount: rows.length, observationCount: sum(rows.map((r) => r.observationCount)) ?? 0,
      sales, revenue, qAvg: sales, tAvg: revenue, avgPrice: avg(rows.map((r) => r.avgPrice)), weightedPrice: ratio(revenue, sales),
      validQGroups: rows.filter((r) => Number.isFinite(r.qAvg)).length, validTGroups: rows.filter((r) => Number.isFinite(r.tAvg)).length
    };
  }))]));
  return { groups, tiers };
}

export function buildModel() {
  if (SOURCE_FILES.length !== 24) throw new Error(`新源文件数量应为24，实际为${SOURCE_FILES.length}`);
  const sourceFiles = SOURCE_FILES.map(readSourceFile);
  const months = sourceFiles.map((s) => s.month).sort();
  const expected = Array.from({ length: 24 }, (_, i) => { const d = new Date(2024, 7 + i, 1); return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`; });
  if (months.join(',') !== expected.join(',')) throw new Error('月份必须唯一且连续覆盖202408—202607');
  const rawRows = sourceFiles.flatMap((s) => s.rows);
  const candidateRows = rawRows.filter((r) => r.eligibleRanks.length > 0).sort((a, b) => {
    const salesOrder = (Number.isFinite(a.sales) ? a.sales : Number.POSITIVE_INFINITY) - (Number.isFinite(b.sales) ? b.sales : Number.POSITIVE_INFINITY);
    return a.month.localeCompare(b.month) || a.parentKey.localeCompare(b.parentKey) || salesOrder || a.asin.localeCompare(b.asin) || a.sourceRow - b.sourceRow;
  });
  const observations = candidateRows.flatMap((r) => r.eligibleRanks.map((rank) => ({
    month: r.month, parentKey: r.parentKey, sourceRow: r.sourceRow, asin: r.asin, rank,
    sales: r.sales, sourceRevenue: r.sourceRevenue, price: r.price
  })));
  const byParent = new Map();
  for (const row of candidateRows) {
    const key = `${row.month}|${row.parentKey}`;
    if (!byParent.has(key)) byParent.set(key, []);
    byParent.get(key).push(row);
  }
  const parentRows = [...byParent.entries()].map(([key, rows]) => {
    const [month, ...rest] = key.split('|');
    return aggregateParent(month, rest.join('|'), rows);
  }).sort((a, b) => a.month.localeCompare(b.month) || a.parentKey.localeCompare(b.parentKey));
  const parentByKey = new Map(parentRows.map((p) => [`${p.month}|${p.parentKey}`, p]));
  const parentByMonth = new Map();
  for (const p of parentRows) { if (!parentByMonth.has(p.month)) parentByMonth.set(p.month, []); parentByMonth.get(p.month).push(p); }
  const categories = {
    overall: { name: '整体市场', filter: () => true },
    pp: { name: 'PP市场', filter: (p) => p.pp },
    high: { name: '高客单价市场', filter: (p) => !p.pp },
    genimo: { name: 'Genimo品牌', filter: (p) => p.genimo },
    genimoPP: { name: 'Genimo PP市场', filter: (p) => p.genimo && p.pp }
  };
  const summaries = {};
  const annual = {};
  for (const [key, c] of Object.entries(categories)) {
    summaries[key] = addComparisons(months.map((month) => monthSummary(month, parentByMonth.get(month) || [], c.filter)));
    annual[key] = annualRows(summaries[key]);
  }
  // BSR is an independent branch: filter observations by the target business scope
  // first, then average within month + BSR value. This prevents a mixed group from
  // leaking observations from another market merely because it shares the same rank.
  const bsr = Object.fromEntries(Object.entries(categories).map(([key, c]) => [key, buildBsrTiers(buildBsrGroups(observations, parentByKey, c.filter))]));
  const duplicateMap = new Map();
  for (const row of rawRows) if (row.asin) {
    const key = `${row.month}|${row.asin}`;
    if (!duplicateMap.has(key)) duplicateMap.set(key, []);
    duplicateMap.get(key).push(row.sourceRow);
  }
  const duplicateAsins = [...duplicateMap.entries()].filter(([, rows]) => rows.length > 1).map(([key, rows]) => ({ key, rows }));
  if (duplicateAsins.length) throw new Error(`同月重复ASIN未解决，禁止发布：${duplicateAsins.length}组`);
  for (const [i, month] of months.entries()) {
    const all = summaries.overall[i]; const pp = summaries.pp[i]; const high = summaries.high[i];
    for (const field of ['sales', 'revenue']) {
      const a = all[field], p = pp[field], h = high[field];
      if (Number.isFinite(a) && Number.isFinite(p) && Number.isFinite(h) && Math.abs(a - p - h) > 1e-6) throw new Error(`${month}整体${field}回加失败`);
    }
  }
  const candidateCount = candidateRows.length;
  const observationCount = observations.length;
  const sourceHash = crypto.createHash('sha256').update(sourceFiles.map((s) => s.sha256).join('')).digest('hex').slice(0, 12);
  const metadata = {
    specVersion: '3.7', batchId: `spec37-${months[0]}-${months.at(-1)}-${sourceHash}`, generatedAt: new Date().toISOString(),
    sourceDir: 'data/raw/0920new', months, firstMonth: months[0], lastMonth: months.at(-1), fileCount: sourceFiles.length,
    rawRowCount: rawRows.length, candidateRowCount: candidateCount, bsrObservationCount: observationCount, parentMonthCount: parentRows.length,
    analysisOrder: '24个月新源 -> 小类BSR任一有效1—100入池 -> 父体月份+父ASIN的Q/T平均 -> 四部分；BSR支线按月份+BSR值的Q/T平均 -> 三档',
    candidateRule: '小类目不参与筛选；小类BSR拆分后任一有效排名1—100入池；多值排名全部保留',
    parentRule: '同一月份+父ASIN内，Q列月销量和T列月销售额分别对有效非负值取算术平均',
    bsrRule: '同一月份+BSR值内，忽略ASIN，对Q/T分别取平均；头部1—20，中部21—50，尾部51—100；档位为BSR值组平均的平均',
    momRule: '本月/去年同月', yoyRule: '本年与上一年共同覆盖月份汇总/上一年对应共同覆盖月份汇总',
    annualCoverageRule: '2025对2024仅8—12月；2026对2025仅1—7月；不是完整自然年度',
    sourceExclusions: { invalidBsr: sourceFiles.reduce((n, s) => n + s.exclusions.INVALID_BSR, 0), outsideTop100: sourceFiles.reduce((n, s) => n + s.exclusions.OUTSIDE_TOP100, 0) },
    duplicateAsinGroups: duplicateAsins.length,
    mixedMarketParents: parentRows.filter((p) => p.ppStatus === 'TIE' || (p.ppRatio > 0 && p.ppRatio < 1)).length,
    qConflictParentCount: parentRows.filter((p) => p.qConflict).length, tConflictParentCount: parentRows.filter((p) => p.tConflict).length,
    missingQParents: parentRows.filter((p) => p.qAvg == null).length, missingTParents: parentRows.filter((p) => p.tAvg == null).length
  };
  return { metadata, months, sourceFiles, rawRows, candidateRows, observations, parentRows, summaries, annual, bsr, duplicateAsins, bands: BANDS };
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const model = buildModel();
  fs.mkdirSync(OUT_DIR, { recursive: true });
  fs.writeFileSync(path.join(OUT_DIR, 'model-spec37.json'), JSON.stringify(model, null, 2), 'utf8');
  fs.writeFileSync(path.join(OUT_DIR, 'metadata-spec37.json'), JSON.stringify(model.metadata, null, 2), 'utf8');
  console.log(JSON.stringify({ output: OUT_DIR, metadata: model.metadata, sample: model.parentRows.find((p) => p.month === '202607' && p.parentKey === 'B0BM74444Q') ?? null }, null, 2));
}

export { OUT_DIR, parseRanks, tierFor, aggregateParent, monthSummary, annualRows, ratio };
