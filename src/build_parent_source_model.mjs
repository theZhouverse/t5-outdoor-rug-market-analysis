import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import XLSX from 'xlsx';
import { pathToFileURL } from 'node:url';

const ROOT = process.cwd();
const SOURCE_DIR = path.join(ROOT, 'data', 'raw', '0920new');
const OUT_DIR = path.join(ROOT, 'outputs', '20260920-new-source-parent-model');
const SOURCE_FILES = fs.readdirSync(SOURCE_DIR)
  .filter((f) => /^Competitor-US-\d{4}\.\d{2}-.*\.xlsx$/i.test(f))
  .sort();

const MONTH_RE = /^Competitor-US-(\d{4})\.(\d{2})-/i;
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
const AUDIT_FIELDS = [
  'ASIN', 'SKU', '商品标题', '品牌', '父ASIN', '小类目', '小类BSR', '月销量', '销量环比增长率', '销量同比增长率',
  '月销售额($)', '子体销量', '子体销售额($)', '变体数', '价格($)', 'prime价格($)', 'Coupon', 'FBA($)', '毛利率', '标签'
];

const BANDS = [
  { key: 'head', name: '头部 1—20', lo: 1, hi: 20 },
  { key: 'middle', name: '中部 21—50', lo: 21, hi: 50 },
  { key: 'tail', name: '尾部 51—100', lo: 51, hi: 100 }
];
const FINE_BANDS = [
  { key: '1_5', name: '1—5', lo: 1, hi: 5 },
  { key: '6_10', name: '6—10', lo: 6, hi: 10 },
  { key: '11_20', name: '11—20', lo: 11, hi: 20 },
  { key: '21_50', name: '21—50', lo: 21, hi: 50 },
  { key: '51_100', name: '51—100', lo: 51, hi: 100 }
];

function text(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/\u00a0/g, ' ').replace(/\r?\n/g, ' ').trim();
}

function number(v) {
  const s = text(v);
  if (!s || /^[-—–]$/.test(s) || /^(n\/a|na|null|undefined)$/i.test(s)) return null;
  const cleaned = s.replace(/[$,%￥,\s]/g, '').replace(/^\((.*)\)$/, '-$1');
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function parseBsr(v) {
  const s = text(v).replace(/,/g, '');
  if (!/^\d+(?:\.0+)?$/.test(s)) return null;
  const n = Number(s);
  return Number.isInteger(n) && n > 0 ? n : null;
}
function categoryBsr(category, bsr) {
  const names=String(category??'').split(/\r?\n/).map(text);
  const ranks=String(bsr??'').split(/\r?\n/).map(text);
  const indexes=names.map((x,i)=>x.toLowerCase()==='outdoor rugs'?i:-1).filter(i=>i>=0);
  if(indexes.length===0) return {rank:null,bsrStatus:'OTHER_CATEGORY'};
  if(indexes.length!==1||names.length!==ranks.length) return {rank:null,bsrStatus:'AMBIGUOUS_CATEGORY_RANK'};
  const rank=parseBsr(ranks[indexes[0]]);
  return {rank,bsrStatus:rank===null?'INVALID_BSR':rank<=100?'ELIGIBLE':'OUTSIDE_TOP100'};
}
const numeric = (v) => Number.isFinite(v) && v >= 0;
const sumKnown = (rows, key) => {
  const vals = rows.map(r => r[key]).filter(Number.isFinite);
  return vals.length ? vals.reduce((a,b) => a+b, 0) : null;
};
// Fractional medians fall into the next rank band, e.g. 20.5 belongs to (20,50].
const tier = (rank, bands) => rank == null ? null : bands.find(b => rank > b.lo - 1 && rank <= b.hi)?.key || null;

function median(values) {
  const a = values.filter(Number.isFinite).slice().sort((x, y) => x - y);
  if (!a.length) return null;
  const i = Math.floor(a.length / 2);
  return a.length % 2 ? a[i] : (a[i - 1] + a[i]) / 2;
}

function mode(values) {
  const counts = new Map();
  for (const v of values.filter(Number.isFinite)) counts.set(v, (counts.get(v) || 0) + 1);
  if (!counts.size) return { value: null, count: 0, unique: 0 };
  const sorted = [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0] - b[0]);
  return { value: sorted[0][0], count: sorted[0][1], unique: counts.size };
}

function ratio(a, b) {
  return Number.isFinite(a) && Number.isFinite(b) && b > 0 ? a / b : null;
}

function growth(current, previous) {
  return Number.isFinite(current) && Number.isFinite(previous) && previous > 0 ? current / previous - 1 : null;
}

function sha256(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

function parseMonth(file) {
  const m = file.match(MONTH_RE);
  return m ? `${m[1]}${m[2]}` : null;
}

function normalisedHeaders(row) {
  return row.map((v) => text(v));
}

function indexOf(headers, name) {
  const wanted = text(name).toLowerCase().replace(/[（）()\s]/g, '');
  return headers.findIndex((h) => text(h).toLowerCase().replace(/[（）()\s]/g, '') === wanted);
}

function isPlastic(v) { return /\bplastic\b/i.test(text(v)); }
function isGenimo(v) { return text(v).toLowerCase() === 'genimo'; }

function categoryByMajority(rows, field) {
  const values = rows.map((r) => field === 'pp' ? r.pp : r.genimo).filter((v) => v !== null);
  const trueCount = values.filter(Boolean).length;
  const falseCount = values.length - trueCount;
  if (!values.length) return { value: false, status: 'NO_TEXT', trueCount: 0, falseCount: 0, ratio: null };
  if (trueCount === falseCount) return { value: trueCount > 0, status: 'TIE', trueCount, falseCount, ratio: trueCount / values.length };
  return { value: trueCount > falseCount, status: 'MAJORITY', trueCount, falseCount, ratio: trueCount / values.length };
}

function aggregateParent(month, parentKey, rows) {
  const salesValues = rows.map((r) => r.sales).filter(numeric);
  const salesMode = mode(salesValues);
  const salesMedian = median(salesValues);
  const modeShare = ratio(salesMode.count, salesValues.length);
  const parentSales = salesMode.value !== null && modeShare > 0.5 ? salesMode.value : salesMedian;
  const salesStatus = salesMode.value === null ? 'NO_VALUE' : salesMode.unique === 1 ? 'CONSISTENT' : modeShare > 0.5 ? 'MODE' : 'MEDIAN_CONFLICT';
  const validChildren = rows.filter((r) => numeric(r.childSales) && numeric(r.childRevenue) && !(r.childSales === 0 && r.childRevenue > 0));
  const childSales = sumKnown(validChildren, 'childSales');
  const childRevenue = sumKnown(validChildren, 'childRevenue');
  const impliedPrices = rows.map((r) => Number.isFinite(r.sales) && r.sales > 0 && Number.isFinite(r.sourceRevenue) ? r.sourceRevenue / r.sales : null).filter(Number.isFinite);
  const childAsp = childSales > 0 ? childRevenue / childSales : null;
  const rankValues = rows.map((r) => r.rank).filter(Number.isFinite);
  const rankMedian = median(rankValues);
  const rankMin = rankValues.length ? Math.min(...rankValues) : null;
  const pp = categoryByMajority(rows, 'pp');
  const genimo = categoryByMajority(rows, 'genimo');
  const title = rows.find((r) => text(r.title))?.title || '';
  const brand = rows.find((r) => text(r.brand))?.brand || '';
  const sourceRevenueValues = rows.map((r) => r.sourceRevenue).filter(Number.isFinite);
  const sourceSalesValues = rows.map((r) => r.sales).filter(Number.isFinite);
  return {
    month, parentKey, parent: rows.find((r) => text(r.parent))?.parent || '',
    asinCount: new Set(rows.map((r) => r.asin).filter(Boolean)).size,
    candidateRows: rows.length, rankMin, rankMedian, bsrForTier: rankMedian,
    band: tier(rankMedian, BANDS), fineBand: tier(rankMedian, FINE_BANDS),
    parentSales, salesMode: salesMode.value, salesModeCount: salesMode.count, salesModeShare: modeShare,
    salesMedian, salesUnique: salesMode.unique, salesStatus,
    salesMin: salesValues.length ? Math.min(...salesValues) : null,
    salesMax: salesValues.length ? Math.max(...salesValues) : null,
    salesConflict: salesMode.unique > 1, salesValidRows: salesValues.length,
    sourceRevenueMedian: median(sourceRevenueValues), sourceSalesMedian: median(sourceSalesValues),
    childValidRows: validChildren.length, childSales, childRevenue, childCoverage: ratio(validChildren.length, rows.length), childAsp,
    childStatus: validChildren.length === 0 ? 'NONE' : validChildren.length === rows.length ? 'FULL' : 'PARTIAL',
    pp: pp.value, ppStatus: pp.status, ppRatio: pp.ratio, ppTrueRows: pp.trueCount, ppFalseRows: pp.falseCount,
    genimo: genimo.value, genimoStatus: genimo.status, genimoRatio: genimo.ratio,
    title, brand,
    sourceRows: rows.map((r) => r.sourceRow),
    asins: rows.map((r) => r.asin).filter(Boolean)
  };
}

function monthSummary(month, parents, category = null) {
  const list = category ? parents.filter(category) : parents;
  const candidateRows = list.reduce((s, r) => s + r.candidateRows, 0);
  const validChildRows = list.reduce((s, r) => s + r.childValidRows, 0);
  const parentSales = sumKnown(list, 'parentSales');
  const childSales = sumKnown(list, 'childSales');
  const childRevenue = sumKnown(list, 'childRevenue');
  const salesParents = list.filter((r) => Number.isFinite(r.parentSales)).length;
  const conflictParents = list.filter((r) => r.salesConflict).length;
  const validRevenueParents = list.filter((r) => r.childStatus !== 'NONE').length;
  return {
    month, parentCount: list.length, candidateRows, salesParents, parentSales,
    validChildRows, childSales, childRevenue, childCoverage: ratio(validChildRows, candidateRows),
    validRevenueParents, revenueParentCoverage: ratio(validRevenueParents, list.length),
    conflictParents, conflictRate: ratio(conflictParents, list.length),
    medianParents: list.filter(r => r.salesStatus === 'MEDIAN_CONFLICT').length,
    missingSalesParents: list.length - salesParents,
    parentSalesLow: sumKnown(list, 'salesMin'), parentSalesHigh: sumKnown(list, 'salesMax'),
    revenueStatus: validChildRows === 0 ? 'NONE' : validChildRows === candidateRows ? 'FULL' : 'PARTIAL',
    childAsp: childSales > 0 ? childRevenue / childSales : null
  };
}

function annualRows(rows) {
  return [...new Set(rows.map(r=>r.month.slice(0,4)))].map(year=>{
    const available=rows.filter(r=>r.month.startsWith(year));
    const current=available.filter(r=>rows.some(p=>p.month===`${Number(year)-1}${r.month.slice(4)}`));
    const previous=current.map(r=>rows.find(p=>p.month===`${Number(year)-1}${r.month.slice(4)}`));
    const sales=sumKnown(current,'parentSales'), revenue=sumKnown(current,'childRevenue');
    const priorSales=sumKnown(previous,'parentSales'), priorRevenue=sumKnown(previous,'childRevenue');
    return {year, availableMonths:available.map(r=>r.month), months:current.map(r=>r.month), priorMonths:previous.map(r=>r.month),
      parentSales:sales, childRevenue:revenue, priorSales, priorRevenue,
      yoyParentSales:growth(sales,priorSales), yoyChildRevenue:growth(revenue,priorRevenue),
      childCoverage:ratio(current.reduce((s,r)=>s+r.validChildRows,0),current.reduce((s,r)=>s+r.candidateRows,0)),
      priorCoverage:ratio(previous.reduce((s,r)=>s+r.validChildRows,0),previous.reduce((s,r)=>s+r.candidateRows,0)),
      status:!current.length?'NO_BASE':current.length===12?'FULL_YEAR':'MATCHED_MONTHS'};
  });
}

function addComparisons(rows) {
  const byMonth = new Map(rows.map((r) => [r.month, r]));
  for (const r of rows) {
    const y = `${Number(r.month.slice(0, 4)) - 1}${r.month.slice(4)}`;
    const prev = byMonth.get(`${Number(r.month) - 1}`) || null;
    const prevMonth = (() => {
      const d = new Date(Number(r.month.slice(0, 4)), Number(r.month.slice(4)) - 1, 1);
      d.setMonth(d.getMonth() - 1);
      return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
    })();
    const prior = byMonth.get(prevMonth);
    const priorYear = byMonth.get(y);
    r.momParentSales = growth(r.parentSales, prior?.parentSales);
    r.momChildRevenue = growth(r.childRevenue, prior?.childRevenue);
    r.yoyParentSales = growth(r.parentSales, priorYear?.parentSales);
    r.yoyChildRevenue = growth(r.childRevenue, priorYear?.childRevenue);
  }
  return rows;
}

function buildModel() {
  if (SOURCE_FILES.length !== 24) throw new Error(`新源文件数量应为24，实际为${SOURCE_FILES.length}`);
  const months = [];
  const sourceFiles = [];
  const rawRows = [];
  const candidateRows = [];
  const sheetStats = [];
  for (const file of SOURCE_FILES) {
    const month = parseMonth(file);
    if (!month) throw new Error(`无法解析月份: ${file}`);
    const filePath = path.join(SOURCE_DIR, file);
    const wb = XLSX.readFile(filePath, { cellDates: false, cellNF: true, cellText: false });
    const businessSheets = wb.SheetNames.filter((s) => /^Competitor-US-\d{6}$/i.test(s));
    if (businessSheets.length !== 1) throw new Error(`${file}业务Sheet数量不是1: ${businessSheets.join(',')}`);
    const sheetName = businessSheets[0];
    const expectedSheet = `Competitor-US-${month}`;
    if (sheetName !== expectedSheet) throw new Error(`${file} Sheet月份不一致: ${sheetName} != ${expectedSheet}`);
    const ws = wb.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
    const headers = normalisedHeaders(matrix[0] || []);
    for (const h of HEADER_NAMES) if (!headers.includes(h)) throw new Error(`${file}缺少表头: ${h}`);
    if(headers.join('|')!==HEADER_NAMES.join('|')) throw new Error(`${file}表头顺序或数量变化`);
    const idx = Object.fromEntries(HEADER_NAMES.map((h) => [h, headers.indexOf(h)]));
    const monthRows = [];
    for (let i = 1; i < matrix.length; i += 1) {
      const cells = matrix[i] || [];
      const get = (h) => cells[idx[h]];
      const asin = text(get('ASIN'));
      const parent = text(get('父ASIN'));
      const sku = text(get('SKU'));
      const title = text(get('商品标题'));
      const brand = text(get('品牌'));
      const {rank,bsrStatus} = categoryBsr(get('小类目'),get('小类BSR'));
      const row = {
        month, file, sheet: sheetName, sourceRow: i + 1, asin, parent, sku, title, brand,
        parentKey: parent || asin || `row:${i + 1}`,
        category: text(get('小类目')), sourceBsr: text(get('小类BSR')), rank, bsrStatus,
        sales: number(get('月销量')), sourceRevenue: number(get('月销售额($)')),
        childSales: number(get('子体销量')), childRevenue: number(get('子体销售额($)')),
        price: number(get('价格($)')), sourceMOM: number(get('销量环比增长率')), sourceYOY: number(get('销量同比增长率')),
        pp: title ? isPlastic(title) : null, genimo: brand ? isGenimo(brand) : null,
        raw: Object.fromEntries(AUDIT_FIELDS.map((h) => [h, cells[idx[h]] ?? null]))
      };
      monthRows.push(row); rawRows.push(row);
      if (rank !== null && rank >= 1 && rank <= 100) candidateRows.push(row);
    }
    months.push(month);
    const exclusions=Object.fromEntries(['OTHER_CATEGORY','AMBIGUOUS_CATEGORY_RANK','INVALID_BSR','OUTSIDE_TOP100'].map(k=>[k,monthRows.filter(r=>r.bsrStatus===k).length]));
    sourceFiles.push({ file, month, sheet: sheetName, sha256: sha256(filePath), rawRows: monthRows.length, candidateRows: monthRows.filter((r) => r.rank !== null && r.rank >= 1 && r.rank <= 100).length,exclusions });
    sheetStats.push({ month, file, sheet: sheetName, rawRows: monthRows.length, candidateRows: monthRows.filter((r) => r.rank !== null && r.rank >= 1 && r.rank <= 100).length, headerCount: headers.length });
  }
  const monthsSorted = [...months].sort();
  const expectedMonths=Array.from({length:24},(_,i)=>{const d=new Date(2024,7+i,1);return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}`;});
  if (monthsSorted.join(',') !== expectedMonths.join(',')) throw new Error('月份必须唯一且连续覆盖202408—202607');
  candidateRows.sort((a,b)=>a.month.localeCompare(b.month)||a.parentKey.localeCompare(b.parentKey)||(a.sales??-1)-(b.sales??-1)||a.asin.localeCompare(b.asin)||a.sourceRow-b.sourceRow);
  const candidatesByMonth = new Map();
  for (const r of candidateRows) {
    const key = `${r.month}|${r.parentKey}`;
    if (!candidatesByMonth.has(key)) candidatesByMonth.set(key, []);
    candidatesByMonth.get(key).push(r);
  }
  const parentRows = [...candidatesByMonth.entries()].map(([key, rows]) => aggregateParent(rows[0].month, key.split('|').slice(1).join('|'), rows));
  const parentByMonth = new Map();
  for (const p of parentRows) { if (!parentByMonth.has(p.month)) parentByMonth.set(p.month, []); parentByMonth.get(p.month).push(p); }
  const categories = {
    overall: { key: 'overall', name: '整体市场', filter: () => true },
    pp: { key: 'pp', name: 'PP市场', filter: (p) => p.pp },
    high: { key: 'high', name: '高客单价市场', filter: (p) => !p.pp },
    genimo: { key: 'genimo', name: 'Genimo品牌', filter: (p) => p.genimo },
    genimoPP: { key: 'genimoPP', name: 'Genimo PP市场', filter: (p) => p.genimo && p.pp }
  };
  const summaries = {};
  for (const c of Object.values(categories)) {
    const rows = monthsSorted.map((month) => monthSummary(month, parentByMonth.get(month) || [], c.key === 'overall' ? null : c.filter));
    summaries[c.key] = addComparisons(rows);
  }
  const annual=Object.fromEntries(Object.keys(categories).map(key=>[key,annualRows(summaries[key])]));
  const ambiguousByMonth=new Map(sourceFiles.map(f=>[f.month,f.exclusions.AMBIGUOUS_CATEGORY_RANK]));
  for(const rows of Object.values(summaries)) for(const [i,r] of rows.entries()) {
    r.ambiguousBsrRows=ambiguousByMonth.get(r.month)||0;
    r.scopeStatus=r.ambiguousBsrRows?'类目与排名无法一一对应，候选范围缺失':'Outdoor Rugs排名已一一对应';
    r.momScopeLimited=Boolean(r.ambiguousBsrRows||(i>0&&ambiguousByMonth.get(rows[i-1].month)));
    r.yoyScopeLimited=Boolean(r.ambiguousBsrRows||ambiguousByMonth.get(`${Number(r.month.slice(0,4))-1}${r.month.slice(4)}`));
  }
  for(const rows of Object.values(annual)) for(const r of rows) {
    r.scopeLimited=[...r.months,...r.priorMonths].some(m=>(ambiguousByMonth.get(m)||0)>0);
  }
  const tiers={};
  for(const [key,c] of Object.entries(categories)) {
    tiers[key]=[...BANDS.map(b=>({...b,type:'coarse'})),...FINE_BANDS.map(b=>({...b,type:'fine'}))].map(b=>({
      key:b.key,name:b.name,type:b.type,rows:addComparisons(monthsSorted.map(month=>monthSummary(month,(parentByMonth.get(month)||[]).filter(p=>c.filter(p)&&(b.type==='coarse'?p.band:p.fineBand)===b.key))))
    }));
  }
  const overall = summaries.overall;
  for (let i = 0; i < monthsSorted.length; i += 1) {
    const pp = summaries.pp[i]; const high = summaries.high[i]; const all = overall[i];
    if (all.parentCount !== pp.parentCount + high.parentCount || Math.abs(all.parentSales - pp.parentSales - high.parentSales) > 1e-6 || Math.abs(all.childRevenue - pp.childRevenue - high.childRevenue) > 1e-6) {
      throw new Error(`整体/PP/高客单价回加失败: ${monthsSorted[i]}`);
    }
  }
  const duplicateGroups = new Map();
  for (const r of candidateRows) {
    if (!r.asin) continue;
    const key = `${r.month}|${r.asin}`;
    if (!duplicateGroups.has(key)) duplicateGroups.set(key, []);
    duplicateGroups.get(key).push(r.sourceRow);
  }
  const duplicateAsins = [...duplicateGroups.entries()]
    .filter(([, rows]) => rows.length > 1)
    .map(([key, rows]) => { const [month, asin] = key.split('|'); return { month, asin, rows }; });
  if(duplicateAsins.length) throw new Error(`同月重复ASIN未解决，禁止发布：${duplicateAsins.length}组`);
  const conflictParents = parentRows.filter((p) => p.salesConflict).sort((a, b) => a.month.localeCompare(b.month) || a.parentKey.localeCompare(b.parentKey));
  const metadata = {
    specVersion: '3.1-parent-metrics',
    batchId: `parent-v31-${monthsSorted[0]}-${monthsSorted.at(-1)}-${crypto.createHash('sha256').update(sourceFiles.map((s) => s.sha256).join('')).digest('hex').slice(0, 12)}`,
    generatedAt: new Date().toISOString(),
    sourceDir: 'data/raw/0920new',
    targetCategory:'Outdoor Rugs',
    months: monthsSorted,
    firstMonth: monthsSorted[0], lastMonth: monthsSorted.at(-1), fileCount: sourceFiles.length,
    rawRowCount: rawRows.length, candidateRowCount: candidateRows.length, parentMonthCount: parentRows.length,
    analysisOrder: '24个月新源 -> 类目与排名逐项对应，取Outdoor Rugs BSR 1-100 inclusive -> 保留子体明细 -> 父ASIN指标分别聚合 -> PP/高客单价/Genimo',
    parentSalesRule: '父ASIN下月销量全部一致取一致值；严格多数>50%取众数；否则取中位数并标记MEDIAN_CONFLICT。所有多值父体均标记冲突，主值为清洗估计',
    revenueRule: '父体有效子体销售额=同时具有子体销量和子体销售额的子体行销售额之和；月销售额($)只作源字段回勾，不累加',
    blankRule: '空白不等于0；子体销量和子体销售额同时有效才进入有效子体统计；部分字段缺失标记PARTIAL',
    classificationRule: '父体候选子体标题命中plastic多数归PP，否则归高客单价；Genimo按父体品牌多数',
    momRule: '本月/上月-1', yoyRule: '本月/去年同月-1',
    duplicateAsinGroups: duplicateAsins.length, conflictParentCount: conflictParents.length,
    medianParentCount:parentRows.filter(p=>p.salesStatus==='MEDIAN_CONFLICT').length,
    missingRevenueParents:parentRows.filter(p=>p.childStatus==='NONE').length,
    mixedMarketParents:parentRows.filter(p=>p.ppRatio>0&&p.ppRatio<1).length,
    ambiguousCategoryRankRows:sourceFiles.reduce((s,f)=>s+f.exclusions.AMBIGUOUS_CATEGORY_RANK,0),
    ambiguousMonths:sourceFiles.filter(f=>f.exclusions.AMBIGUOUS_CATEGORY_RANK>0).map(f=>f.month)
  };
  return { metadata, months: monthsSorted, sourceFiles, sheetStats, rawRows, candidateRows, parentRows, summaries, annual, tiers, duplicateAsins, conflictParents, bands: BANDS, fineBands: FINE_BANDS };
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
const model = buildModel();
fs.mkdirSync(OUT_DIR, { recursive: true });
const serialisableModel = { ...model, candidateRows: model.candidateRows.map(({ raw, ...row }) => row) };
fs.writeFileSync(path.join(OUT_DIR, 'model.json'), JSON.stringify(serialisableModel, null, 2), 'utf8');
fs.writeFileSync(path.join(OUT_DIR, 'metadata.json'), JSON.stringify(model.metadata, null, 2), 'utf8');
console.log(JSON.stringify({
  output: OUT_DIR, metadata: model.metadata,
  monthly: model.summaries.overall.map((r) => ({ month: r.month, parentCount: r.parentCount, candidateRows: r.candidateRows, parentSales: r.parentSales, childRevenue: r.childRevenue, childCoverage: r.childCoverage, conflictParents: r.conflictParents })),
  sample: model.parentRows.find((r) => r.month === '202607' && r.parentKey === 'B0BM74444Q') || null,
  duplicateAsinGroups: model.duplicateAsins.length
}, null, 2));
}

export { buildModel, OUT_DIR, aggregateParent, monthSummary, annualRows, parseBsr, categoryBsr, growth };
