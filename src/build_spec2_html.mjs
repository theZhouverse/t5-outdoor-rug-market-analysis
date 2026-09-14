import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { pathToFileURL } from 'node:url';
import XLSX from 'xlsx';

const ROOT = process.cwd();
const SOURCE = path.join(ROOT, 'data', 'raw', '地垫-卖家精灵市场数据.xlsx');
const OUTPUT = path.join(ROOT, '交付', '户外地垫市场分析报告-优化版.html');
const SOURCE_HASH = fs.existsSync(SOURCE) ? crypto.createHash('sha256').update(fs.readFileSync(SOURCE)).digest('hex') : '';
const UI_CSS_PATH = path.join(ROOT, 'src', 'original_ui.css');
const UI_CSS = fs.existsSync(UI_CSS_PATH) ? fs.readFileSync(UI_CSS_PATH, 'utf8') : '';

const MONTHS = [];
const BANDS = [
  { key: 'head', name: '头部 1-20', lo: 1, hi: 20 },
  { key: 'middle', name: '中部 21-50', lo: 21, hi: 50 },
  { key: 'tail', name: '尾部 51-100', lo: 51, hi: 100 }
];
const FINE_BANDS = [
  { key: '1_5', name: '1-5', lo: 1, hi: 5 },
  { key: '6_10', name: '6-10', lo: 6, hi: 10 },
  { key: '11_20', name: '11-20', lo: 11, hi: 20 },
  { key: '21_50', name: '21-50', lo: 21, hi: 50 },
  { key: '51_100', name: '51-100', lo: 51, hi: 100 }
];

function normaliseMonth(value) {
  const s = clean(value);
  const compact = s.match(/^(\d{4})(\d{2})$/);
  if (compact) return compact[1] + compact[2];
  const dotted = s.match(/^(\d{4})[.．](\d{1,2})$/);
  if (dotted) return dotted[1] + String(Number(dotted[2])).padStart(2, '0');
  return null;
}

function clean(value) {
  if (value === null || value === undefined) return '';
  const s = String(value).replace(/\u00a0/g, ' ').replace(/\r?\n/g, ' ').trim();
  if (!s || /^[-—–]$/.test(s) || /^(n\/a|na|null|undefined)$/i.test(s)) return '';
  return s;
}

function decodeXml(value) {
  return String(value)
    .replace(/<[^>]*>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g,(_,n)=>String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi,(_,n)=>String.fromCodePoint(parseInt(n,16)));
}

function display(cell) {
  if (!cell) return '';
  // Rich text is the actual cell content; hyperlink display can contain only
  // its first run (100 Rugs.com titles in the supplied source).
  if (cell.r !== null && cell.r !== undefined) {
    const rich=decodeXml(cell.r);
    if(clean(rich)) return rich;
  }
  if (cell.v !== null && cell.v !== undefined && clean(cell.v) !== '') return String(cell.v);
  if (cell.l && cell.l.display !== null && cell.l.display !== undefined && clean(cell.l.display) !== '') {
    return String(cell.l.display);
  }
  if (cell.r !== null && cell.r !== undefined && clean(cell.r) !== '') return decodeXml(cell.r);
  if (cell.w !== null && cell.w !== undefined && clean(cell.w) !== '') return String(cell.w);
  return '';
}

function normalHeader(value) {
  return clean(value).toLowerCase().replace(/\s/g, '').replace(/[（）()]/g, '');
}

function findColumn(headers, names) {
  const wanted = names.map(normalHeader);
  for (const name of wanted) {
    const idx = headers.findIndex((h) => normalHeader(h) === name);
    if (idx >= 0) return idx;
  }
  return -1;
}

function parseNumber(value) {
  const s = clean(value);
  if (!s) return null;
  const negative = /^\(.*\)$/.test(s);
  const pct = s.includes('%');
  const normalized = s.replace(/[,$￥\s]/g, '').replace(/^\((.*)\)$/, '$1').replace(/%$/, '');
  const n = Number(normalized);
  if (!Number.isFinite(n)) return null;
  return negative ? -n : n;
}

function parseBsr(value) {
  const s = clean(value).replace(/,/g, '');
  if (!s) return null;
  // Accept integer ranks and integer-looking N.0 text. Do not split a
  // non-integer decimal such as 1.5 into the false ranks 1 and 5.
  const matches = s.match(/(?<![\d.\-−])[-−]?\s*\d+(?:\.0+)?(?![\d.])/g) || [];
  const values = matches.map((v) => Number(v.replace(/\s/g,'').replace('−','-'))).filter((v) => Number.isInteger(v) && v > 0);
  return values.length ? Math.min(...values) : null;
}

function isPlasticTitle(value) {
  return /\bplastic\b/i.test(clean(value));
}

function isGenimo(value) {
  return clean(value).toLowerCase() === 'genimo';
}

function sourceId(row) {
  return row.month + '#' + row.sourceRow;
}

function listingKey(row) {
  // The current business scope keeps every BSR-qualified source row.  Use the
  // immutable month/source-row identity so even repeated ASINs remain separate
  // analysis units; ASIN and parent ASIN are audit attributes only.
  return 'source:' + sourceId(row);
}

// SellerSprite exports are not completely uniform: older monthly sheets have
// a report title in row 1 and the actual header in row 2, while newer sheets
// start with the header in row 1. Detect the header by its stable ASIN field
// instead of assuming a fixed row number.
function findHeaderRow(ws, range) {
  const lastProbe = Math.min(range.e.r, range.s.r + 8);
  for (let r = range.s.r; r <= lastProbe; r += 1) {
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      headers.push(display(ws[XLSX.utils.encode_cell({ r, c })]));
    }
    const asin = findColumn(headers, ['ASIN']);
    const title = findColumn(headers, ['商品标题', '标题']);
    const brand = findColumn(headers, ['品牌']);
    if (asin >= 0 && (title >= 0 || brand >= 0)) return r;
  }
  return range.s.r + 1;
}

function readRawRows() {
  if (!fs.existsSync(SOURCE)) throw new Error('找不到原始主源: ' + SOURCE);
  const wb = XLSX.readFile(SOURCE, {
    cellFormula: true,
    cellStyles: true,
    cellNF: true,
    cellText: false
  });
  const sheets = wb.SheetNames.map((name) => ({ original: String(name), month: normaliseMonth(name) }))
    .filter((item) => item.month)
    .sort((a, b) => a.month.localeCompare(b.month));
  if (!sheets.length) throw new Error('原始工作簿没有找到 YYYYMM 月度明细子表');
  const allRows = [];
  const sheetStats = [];
  for (const sheet of sheets) {
    const month = sheet.month;
    MONTHS.push(month);
    const ws = wb.Sheets[sheet.original];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const headerRow = findHeaderRow(ws, range);
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      headers.push(display(ws[XLSX.utils.encode_cell({ r: headerRow, c })]));
    }
    const idx = {
      asin: findColumn(headers, ['ASIN']),
      sku: findColumn(headers, ['SKU']),
      brand: findColumn(headers, ['品牌']),
      title: findColumn(headers, ['商品标题', '标题']),
      parent: findColumn(headers, ['父ASIN', '父体ASIN']),
      smallBsr: findColumn(headers, ['小类BSR', '小类目BSR', '小类排名']),
      sales: findColumn(headers, ['月销量', '月销售量']),
      revenue: findColumn(headers, ['月销售额($)', '月销售额']),
      price: findColumn(headers, ['价格($)', '价格']),
      fba: findColumn(headers, ['FBA($)', 'FBA']),
      margin: findColumn(headers, ['毛利率']),
      coupon: findColumn(headers, ['Coupon'])
    };
    let rowCount = 0;
    let rankCount = 0;
    for (let r = headerRow + 1; r <= range.e.r; r += 1) {
      const values = [];
      let hasValue = false;
      for (let c = range.s.c; c <= range.e.c; c += 1) {
        const value = display(ws[XLSX.utils.encode_cell({ r, c })]);
        values.push(value);
        if (clean(value)) hasValue = true;
      }
      if (!hasValue) continue;
      const row = {
        month,
        sourceSheet: sheet.original,
        sourceRow: r + 1,
        asin: clean(values[idx.asin]),
        sku: clean(values[idx.sku]),
        brand: clean(values[idx.brand]),
        title: clean(values[idx.title]),
        parent: clean(values[idx.parent]),
        sourceBsr: clean(values[idx.smallBsr]),
        rank: parseBsr(values[idx.smallBsr]),
        bsrStatus: clean(values[idx.smallBsr]) ? (parseBsr(values[idx.smallBsr]) === null ? '无效' : '有效') : '缺失',
        sales: parseNumber(values[idx.sales]),
        revenue: parseNumber(values[idx.revenue]),
        price: parseNumber(values[idx.price]),
        fba: parseNumber(values[idx.fba]),
        margin: clean(values[idx.margin]),
        coupon: clean(values[idx.coupon]),
        plastic: isPlasticTitle(values[idx.title]),
        genimo: isGenimo(values[idx.brand])
      };
      // Some exports append a human-readable note below the data table in
      // the ASIN column (for example, a duplicate-removal explanation). Keep
      // legitimate rows whose ASIN is present but all metrics are blank, yet
      // drop note-only rows that have no product text or numeric field.
      const hasBusinessText = Boolean(row.asin || row.parent || row.title || row.brand || row.sku);
      const hasDataMetric = [row.rank, row.sales, row.revenue, row.price, row.fba]
        .some((value) => Number.isFinite(value));
      const plausibleAsin = /^[A-Z0-9]{8,15}$/i.test(row.asin);
      if ((!hasBusinessText && !hasDataMetric) || (!plausibleAsin && !row.parent && !row.title && !row.brand && !row.sku && !hasDataMetric)) continue;
      allRows.push(row);
      rowCount += 1;
      if (row.rank !== null) rankCount += 1;
    }
    sheetStats.push({ month, rowCount, rankCount, headerRow: headerRow + 1 });
  }
  return { rows: allRows, sheetStats };
}

function representativeCompare(a, b) {
  const ar = a.rank === null ? Number.POSITIVE_INFINITY : a.rank;
  const br = b.rank === null ? Number.POSITIVE_INFINITY : b.rank;
  if (ar !== br) return ar - br;
  const ac = (Number.isFinite(a.sales) ? 1 : 0) + (Number.isFinite(a.revenue) ? 1 : 0);
  const bc = (Number.isFinite(b.sales) ? 1 : 0) + (Number.isFinite(b.revenue) ? 1 : 0);
  if (ac !== bc) return bc - ac;
  if (Number.isFinite(a.price) !== Number.isFinite(b.price)) return Number.isFinite(b.price) - Number.isFinite(a.price);
  return a.sourceRow - b.sourceRow;
}

/**
 * Legacy helper retained for historical audits.  The current business model
 * deliberately does not call it: after BSR 1..100 filtering every candidate
 * source row is retained as one analysis unit.
 */
function dedup(rows) {
  const groups = new Map();
  for (const row of rows) {
    if (!row.parent) continue;
    const key = row.month + '|' + listingKey(row);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  const result = [];
  for (const [key, items] of groups.entries()) {
    const rep = items.slice().sort(representativeCompare)[0];
    result.push({
      ...rep,
      familyKey: key,
      // Classification is intentionally based on the representative row
      // selected from the already BSR-filtered candidate group.
      plastic: isPlasticTitle(rep.title),
      genimo: isGenimo(rep.brand),
      rawRows: items.length,
      sourceRows: items.map(sourceId)
    });
  }
  return result.sort((a, b) => (a.month + a.familyKey).localeCompare(b.month + b.familyKey));
}

function percent(current, previous) {
  if (!Number.isFinite(current) || !Number.isFinite(previous) || previous <= 0) return null;
  return (current - previous) / Math.abs(previous);
}

function sum(rows, field) {
  const values = rows.map((r) => r[field]).filter(Number.isFinite);
  return values.length ? values.reduce((a, b) => a + b, 0) : null;
}

function avg(rows, field) {
  const values = rows.map((r) => r[field]).filter(Number.isFinite);
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : null;
}

function summarise(rows) {
  const salesRows = rows.filter((r) => Number.isFinite(r.sales));
  const revenueRows = rows.filter((r) => Number.isFinite(r.revenue));
  const paired = rows.filter((r) => Number.isFinite(r.sales) && Number.isFinite(r.revenue) && r.sales > 0);
  const sales = sum(rows, 'sales');
  const revenue = sum(rows, 'revenue');
  const pairedSales = sum(paired, 'sales');
  const pairedRevenue = sum(paired, 'revenue');
  return {
    count: rows.length,
    sales,
    revenue,
    avgPrice: avg(rows, 'price'),
    pairedAsp: pairedSales && pairedRevenue !== null ? pairedRevenue / pairedSales : null,
    salesCoverage: rows.length ? salesRows.length / rows.length : null,
    revenueCoverage: rows.length ? revenueRows.length / rows.length : null,
    missingSales: rows.length - salesRows.length,
    missingRevenue: rows.length - revenueRows.length,
    missingPrice: rows.length - rows.filter((r) => Number.isFinite(r.price)).length
  };
}

function previousMonth(month) {
  const y = Number(month.slice(0, 4));
  const m = Number(month.slice(4, 6));
  const date = new Date(Date.UTC(y, m - 2, 1));
  return String(date.getUTCFullYear()) + String(date.getUTCMonth() + 1).padStart(2, '0');
}

function previousYear(month) {
  return String(Number(month.slice(0, 4)) - 1) + month.slice(4, 6);
}

function enrichTrend(months, rows) {
  const byMonth = new Map();
  for (const month of months) byMonth.set(month, []);
  for (const row of rows) {
    if (!byMonth.has(row.month)) byMonth.set(row.month, []);
    byMonth.get(row.month).push(row);
  }
  const base = new Map();
  for (const month of months) base.set(month, summarise(byMonth.get(month) || []));
  return months.map((month) => {
    const current = base.get(month);
    // Project reporting convention: monthly MOM means the current month
    // versus the same month in the prior year. Annual comparisons are shown
    // separately as YOY in the annual tables.
    const mom = base.get(previousYear(month)) || {};
    return {
      month,
      ...current,
      comparisonMonth: previousYear(month),
      comparisonCount: mom.count ?? null,
      comparisonSales: mom.sales ?? null,
      comparisonRevenue: mom.revenue ?? null,
      quality: current.count === 0 ? '无样本' : ((current.salesCoverage < .95 || current.revenueCoverage < .95 || Math.abs(percent(current.count, mom.count) ?? 0) > .3 || Math.abs(percent(current.count, base.get(previousMonth(month))?.count) ?? 0) > .3) ? '可比性受限' : '覆盖可用'),
      momSales: percent(current.sales, mom.sales),
      momRevenue: percent(current.revenue, mom.revenue),
      momAvgPrice: percent(current.avgPrice, mom.avgPrice),
      momPairedAsp: percent(current.pairedAsp, mom.pairedAsp)
    };
  });
}

function annualRows(months, rows) {
  const byYear = new Map();
  for (const row of rows) {
    const year = row.month.slice(0, 4);
    if (!byYear.has(year)) byYear.set(year, []);
    byYear.get(year).push(row);
  }
  const years = [...new Set(months.map(m => m.slice(0, 4)))].sort();
  const summaries = years.map((year) => ({ year, ...summarise(byYear.get(year) || []) }));
  return summaries.map((current, i) => {
    const previousYear = String(Number(current.year) - 1);
    const currentMonths = new Set(months.filter(m => m.startsWith(current.year)).map(m => m.slice(4)));
    const previousMonths = new Set(months.filter(m => m.startsWith(previousYear)).map(m => m.slice(4)));
    const commonMonths = Array.from(currentMonths).filter((month) => previousMonths.has(month)).sort();
    const currentComparable = commonMonths.length
      ? summarise((byYear.get(current.year) || []).filter((row) => commonMonths.includes(row.month.slice(4, 6))))
      : {};
    const previousComparable = commonMonths.length
      ? summarise((byYear.get(previousYear) || []).filter((row) => commonMonths.includes(row.month.slice(4, 6))))
      : {};
    const yoyPeriod = previousYear && commonMonths.length
      ? current.year + ' [' + commonMonths.join(',') + '] 对 ' + previousYear + ' [' + commonMonths.join(',') + ']'
      : null;
    return {
      ...current,
      coverage: [...currentMonths].join(','),
      commonMonths,
      currentComparable,
      previousComparable,
      yoySales: percent(currentComparable.sales, previousComparable.sales),
      yoyRevenue: percent(currentComparable.revenue, previousComparable.revenue),
      yoyPeriod
    };
  });
}

function filterBand(rows, band) {
  return rows.filter((r) => r.rank !== null && r.rank >= band.lo && r.rank <= band.hi);
}

function countRawRows(rows) {
  return rows.reduce((total, row) => total + (Number.isFinite(row.rawRows) ? row.rawRows : 1), 0);
}

function candidateMonthly(rows) {
  return MONTHS.map((month) => {
    const items = rows.filter((row) => row.month === month);
    return { month, count: items.length, rawCount: countRawRows(items) };
  });
}

function buildCategory(title, fullRows, topRows, candidateRows = []) {
  const tierRows = {};
  for (const band of BANDS) tierRows[band.key] = filterBand(topRows, band);
  const fineRows = {};
  for (const band of FINE_BANDS) fineRows[band.key] = filterBand(topRows, band);
  return {
    title,
    fullRows,
    topRows,
    topCandidateMonthly: candidateMonthly(candidateRows),
    monthly: enrichTrend(MONTHS, fullRows),
    topMonthly: enrichTrend(MONTHS, topRows),
    annual: annualRows(MONTHS, fullRows),
    topAnnual: annualRows(MONTHS, topRows),
    tiers: Object.fromEntries(BANDS.map((band) => [band.key, enrichTrend(MONTHS, tierRows[band.key])])),
    fine: Object.fromEntries(FINE_BANDS.map((band) => [band.key, enrichTrend(MONTHS, fineRows[band.key])])),
    annualBands: Object.fromEntries([...BANDS, ...FINE_BANDS].map(band => [band.key, annualRows(MONTHS, filterBand(topRows, band))]))
  };
}

function esc(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function fmt(value, digits = 0) {
  if (!Number.isFinite(value)) return '—';
  return value.toLocaleString('zh-CN', { maximumFractionDigits: digits, minimumFractionDigits: digits });
}

function fmtPct(value) {
  return Number.isFinite(value) ? (value * 100).toFixed(1) + '%' : '—';
}

function lastNonEmpty(rows) {
  return rows.slice().reverse().find((r) => r.count > 0) || rows[rows.length - 1];
}

function chartMonthLabel(month) {
  return String(month).slice(0, 4) + '.' + Number(String(month).slice(4, 6));
}

function chartValueLabel(value, percentMode = false, digits = 2) {
  return percentMode ? fmtPct(value) : fmt(value, digits);
}

function svgBarChart(title, rows, field, color, digits) {
  const values = rows.map((r) => Number.isFinite(r[field]) ? r[field] : 0);
  const max = Math.max(...values, 0);
  const width = 900;
  const height = 260;
  const left = 54;
  const bottom = 42;
  const plotH = 170;
  const plotW = width - left - 18;
  const barW = rows.length ? Math.max(2, plotW / rows.length - 2) : 2;
  const chartMax = max > 0 ? max : 1;
  let body = '<div class="chart interactive-chart" data-chart-kind="bar"><div class="chart-head"><h4>' + esc(title) + '</h4><span class="chart-hint">悬停或按 Tab 聚焦</span></div><div class="chart-stage"><svg viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="' + esc(title) + '" data-interactive-chart="bar">';
  for (const ratio of [0, 0.25, 0.5, 0.75, 1]) {
    const y = height - bottom - ratio * plotH;
    body += '<line class="chart-grid-line" x1="' + left + '" y1="' + y.toFixed(1) + '" x2="' + width + '" y2="' + y.toFixed(1) + '"/>';
    body += '<text x="8" y="' + (y + 4).toFixed(1) + '" class="axis-label">' + esc(fmt(chartMax * ratio, digits)) + '</text>';
  }
  rows.forEach((row, i) => {
    if (!Number.isFinite(row[field])) return;
    const value = values[i];
    const h = max > 0 ? (value / chartMax) * plotH : 0;
    const x = left + i * (plotW / Math.max(rows.length, 1)) + 1;
    const y = height - bottom - h;
    const label = chartMonthLabel(row.month);
    const valueLabel = chartValueLabel(row[field], false, digits);
    const tooltip = label + '  ·  ' + valueLabel + (field==='revenue'?' 美元':' 件') + (row.quality?' · '+row.quality:'');
    const cx = x + barW / 2;
    body += '<rect class="chart-mark chart-bar" data-chart-point data-chart-index="' + i + '" data-x="' + cx.toFixed(1) + '" data-label="' + esc(label) + '" data-tooltip="' + esc(tooltip) + '" tabindex="0" role="img" aria-label="' + esc(tooltip) + '" x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" width="' + barW.toFixed(1) + '" height="' + h.toFixed(1) + '" rx="2" fill="' + color + '"><title>' + esc(tooltip) + '</title></rect>';
    if (i % Math.max(1, Math.ceil(rows.length / 10)) === 0) {
      body += '<text x="' + cx.toFixed(1) + '" y="' + (height - 17) + '" text-anchor="middle" class="axis-label">' + esc(label) + '</text>';
    }
  });
  body += '</svg><div class="chart-hoverline" aria-hidden="true"></div><div class="chart-tooltip" role="tooltip" hidden></div></div><div class="chart-readout" aria-live="polite">悬停或聚焦任意月份查看精确值</div></div>';
  return body;
}

function svgLineChart(title, rows, series, percentMode = false) {
  const width = 900;
  const height = 280;
  const left = 54;
  const right = 18;
  const top = 28;
  const bottom = 44;
  const plotW = width - left - right;
  const plotH = height - top - bottom;
  const values = [];
  for (const item of series) for (const row of rows) if (Number.isFinite(row[item.field])) values.push(row[item.field]);
  let min = values.length ? Math.min(...values) : 0;
  let max = values.length ? Math.max(...values) : 1;
  if (min === max) { min -= 1; max += 1; }
  const scaleX = (i) => left + (rows.length <= 1 ? 0 : i * plotW / (rows.length - 1));
  const scaleY = (v) => top + (max - v) * plotH / (max - min);
  let body = '<div class="chart interactive-chart" data-chart-kind="line"><div class="chart-head"><h4>' + esc(title) + '</h4><span class="chart-hint">悬停或按 Tab 聚焦</span></div><div class="chart-stage"><svg viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="' + esc(title) + '" data-interactive-chart="line">';
  for (const ratio of [0, 0.25, 0.5, 0.75, 1]) {
    const y = top + ratio * plotH;
    const value = max - ratio * (max - min);
    body += '<line class="chart-grid-line" x1="' + left + '" y1="' + y.toFixed(1) + '" x2="' + (width - right) + '" y2="' + y.toFixed(1) + '"/>';
    body += '<text x="8" y="' + (y + 4).toFixed(1) + '" class="axis-label">' + esc(chartValueLabel(value, percentMode)) + '</text>';
  }
  if (min < 0 && max > 0) {
    const y0 = scaleY(0);
    body += '<line class="chart-zero-line" x1="' + left + '" y1="' + y0.toFixed(1) + '" x2="' + (width - right) + '" y2="' + y0.toFixed(1) + '"/>';
  }
  body += '<line class="chart-axis-line" x1="' + left + '" y1="' + top + '" x2="' + left + '" y2="' + (height - bottom) + '"/>';
  for (const item of series) {
    const points = rows.map((row, i) => Number.isFinite(row[item.field]) ? scaleX(i).toFixed(1) + ',' + scaleY(row[item.field]).toFixed(1) : null).filter(Boolean);
    if (points.length > 1) body += '<polyline class="chart-line" fill="none" stroke="' + item.color + '" stroke-width="2.5" points="' + points.join(' ') + '"/>';
    rows.forEach((row, i) => {
      if (Number.isFinite(row[item.field])) {
        const label = chartMonthLabel(row.month);
        const valueLabel = chartValueLabel(row[item.field], percentMode);
        body += '<circle class="chart-point" data-chart-index="' + i + '" data-x="' + scaleX(i).toFixed(1) + '" cx="' + scaleX(i).toFixed(1) + '" cy="' + scaleY(row[item.field]).toFixed(1) + '" r="3.4" fill="' + item.color + '"><title>' + esc(label + '  ·  ' + item.name + '：' + valueLabel) + '</title></circle>';
      }
    });
  }
  const step = Math.max(1, Math.ceil(rows.length / 10));
  rows.forEach((row, i) => {
    if (i % step === 0) body += '<text x="' + scaleX(i).toFixed(1) + '" y="' + (height - 17) + '" text-anchor="middle" class="axis-label">' + esc(chartMonthLabel(row.month)) + '</text>';
  });
  series.forEach((item, i) => {
    const x = left + i * 155;
    body += '<rect x="' + x + '" y="8" width="12" height="12" fill="' + item.color + '"/><text x="' + (x + 17) + '" y="18" class="legend-label">' + esc(item.name) + '</text>';
  });
  rows.forEach((row, i) => {
    const label = chartMonthLabel(row.month);
    const tooltip = [label].concat(series.map((item) => item.name + '：' + chartValueLabel(row[item.field], percentMode))).join('\n');
    const x = scaleX(i);
    const slot = plotW / Math.max(rows.length, 1);
    const hitX = i === 0 ? left : x - slot / 2;
    const hitWidth = i === 0 || i === rows.length - 1 ? slot / 2 : slot;
    body += '<rect class="chart-hit" data-chart-point data-chart-index="' + i + '" data-x="' + x.toFixed(1) + '" data-label="' + esc(label) + '" data-tooltip="' + esc(tooltip) + '" tabindex="0" role="img" aria-label="' + esc(tooltip.replace(/\n/g, '；')) + '" x="' + hitX.toFixed(1) + '" y="' + top + '" width="' + Math.max(8, hitWidth).toFixed(1) + '" height="' + plotH + '"></rect>';
  });
  body += '</svg><div class="chart-hoverline" aria-hidden="true"></div><div class="chart-tooltip" role="tooltip" hidden></div></div><div class="chart-readout" aria-live="polite">悬停或聚焦任意月份查看精确值</div></div>';
  return body;
}

function table(headers, rows) {
  let out = '<div class="table-wrap"><table><thead><tr>';
  for (const h of headers) out += '<th>' + esc(h) + '</th>';
  out += '</tr></thead><tbody>';
  for (const row of rows) {
    out += '<tr>';
    for (const cell of row) out += '<td>' + (cell && cell.__html ? cell.value : esc(cell)) + '</td>';
    out += '</tr>';
  }
  return out + '</tbody></table></div>';
}

// A report subsection is a native <details> element so readers can collapse
// long tables/charts without losing the section's stable anchor. The summary
// remains the visible navigation label and the body contains the existing
// generated artifact.
function subsection(id, label, body, open = true) {
  return '<details id="' + esc(id) + '" class="subsection"' + (open ? ' open' : '') + '><summary>' + esc(label) + '</summary>' + body + '</details>';
}

// Small, dependency-free viewer runtime for the generated SVG charts. It
// follows Archify's interaction pattern: semantic data attributes on SVG
// marks, a single stateful hover/focus layer, and keyboard-accessible targets.
function chartRuntimeScript() {
  return String.raw`(function(){
var root=document.documentElement,button=document.getElementById("theme-toggle"),saved="dark";
try{saved=localStorage.getItem("market-report-theme")||"dark"}catch(e){}
root.dataset.theme=saved;button.textContent=saved==="dark"?"☀":"☾";
button.addEventListener("click",function(){var next=root.dataset.theme==="dark"?"light":"dark";root.dataset.theme=next;button.textContent=next==="dark"?"☀":"☾";try{localStorage.setItem("market-report-theme",next)}catch(e){}});
var links=[].slice.call(document.querySelectorAll(".sidebar nav a")),groups=[].slice.call(document.querySelectorAll(".nav-group")),targets=[].slice.call(document.querySelectorAll(".subsection[id],section[id]"));
function openTarget(id){var el=document.getElementById(id);if(!el)return;var details=el.closest("details");if(details)details.open=true;var section=el.closest("section");if(section){var group=document.querySelector('.nav-group[data-target="'+section.id+'"]');if(group)group.open=true}el.scrollIntoView({behavior:"smooth",block:"start"})}
links.forEach(function(a){a.addEventListener("click",function(){openTarget(a.getAttribute("href").slice(1))})});
function setActive(el){var id=el.id;links.forEach(function(a){a.classList.toggle("is-active",a.getAttribute("href")==="#"+id)});var section=el.closest("section[id]"),sectionId=section?section.id:(el.matches("section[id]")?el.id:null);groups.forEach(function(group){group.classList.toggle("is-active",!!sectionId&&group.dataset.target===sectionId)})}
if("IntersectionObserver" in window){var observer=new IntersectionObserver(function(entries){entries.forEach(function(entry){if(entry.isIntersecting)setActive(entry.target)})},{rootMargin:"-18% 0px -68% 0px"});targets.forEach(function(el){observer.observe(el)})}
function clearChart(chart){var tip=chart.querySelector(".chart-tooltip"),line=chart.querySelector(".chart-hoverline"),readout=chart.querySelector(".chart-readout");if(tip)tip.hidden=true;if(line)line.classList.remove("is-visible");chart.querySelectorAll("[data-chart-index]").forEach(function(mark){mark.classList.remove("is-active")});if(readout)readout.textContent="悬停或聚焦任意月份查看精确值"}
function showChart(chart,mark){var tip=chart.querySelector(".chart-tooltip"),line=chart.querySelector(".chart-hoverline"),readout=chart.querySelector(".chart-readout"),stage=chart.querySelector(".chart-stage"),svg=chart.querySelector("svg");if(!tip||!stage||!svg)return;var index=mark.getAttribute("data-chart-index"),text=mark.getAttribute("data-tooltip")||"",x=Number(mark.getAttribute("data-x"));chart.querySelectorAll("[data-chart-index]").forEach(function(item){item.classList.toggle("is-active",item.getAttribute("data-chart-index")===index)});tip.textContent=text;tip.hidden=false;var vb=svg.viewBox.baseVal,sr=svg.getBoundingClientRect(),wr=stage.getBoundingClientRect(),localX=(x-vb.x)/vb.width*sr.width+sr.left-wr.left;if(line){line.style.left=localX+"px";line.classList.add("is-visible")}var maxLeft=Math.max(8,stage.clientWidth-(tip.offsetWidth||230)-8);tip.style.left=Math.min(maxLeft,Math.max(8,localX+12))+"px";tip.style.top="8px";if(readout)readout.textContent=text.replace(/\n/g,"  ·  ")}
document.querySelectorAll(".interactive-chart").forEach(function(chart){chart.addEventListener("pointerover",function(event){var mark=event.target.closest("[data-chart-point]");if(mark&&chart.contains(mark))showChart(chart,mark)});chart.addEventListener("focusin",function(event){var mark=event.target.closest("[data-chart-point]");if(mark&&chart.contains(mark))showChart(chart,mark)});chart.addEventListener("pointerleave",function(){if(!chart.querySelector("[data-chart-point]:focus"))clearChart(chart)});chart.addEventListener("focusout",function(event){if(!chart.contains(event.relatedTarget))clearChart(chart)})});
})();`;
}

function trendTable(data, topData, candidateData) {
  const rows = data.map((r, i) => [
    r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
    fmt(r.count),
    fmt(r.sales),
    fmt(r.revenue, 2),
    fmt(r.avgPrice, 2),
    fmt(r.pairedAsp, 2),
    fmtPct(r.momSales),
    fmtPct(r.momRevenue),
    fmt(candidateData[i] ? candidateData[i].rawCount : null),
    fmt(topData[i] ? topData[i].count : null),
    fmt(r.count-r.missingSales),fmt(r.count-r.missingRevenue),fmt(r.count-r.missingPrice),
    fmt(r.comparisonCount),r.quality
  ]);
  return table(['月份', 'BSR候选行数', '销量', '销售额($)', '平均标价($)', '加权成交均价($)', '销量MOM', '销售额MOM', 'BSR候选原始行数', '候选父ASIN数','销量有效数','金额有效数','价格有效数','去年同月候选行数','可比性'], rows);
}

function topTrendTable(data, candidateData) {
  const rows = data.map((r, i) => [
    r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
    fmt(r.count),
    fmt(r.sales),
    fmt(r.revenue, 2),
    fmt(r.avgPrice, 2),
    fmt(r.pairedAsp, 2),
    fmtPct(r.momSales),
    fmtPct(r.momRevenue),
    fmt(candidateData[i] ? candidateData[i].rawCount : null),fmt(candidateData[i]?.count),fmt(candidateData[i] ? Math.max(0,candidateData[i].count-r.count) : null),r.count<100?'少于100个样本':r.count>100?'超过100；并列或多小类':'100个样本',r.quality
  ]);
  return table(['月份', 'BSR候选行数', '销量', '销售额($)', '平均标价($)', '加权成交均价($)', '销量MOM', '销售额MOM', '候选原始行数','候选父ASIN数','重复父ASIN行数','样本状态','可比性'], rows);
}

function annualTable(data, topData, prefix = '整体') {
  const rows = data.map((r, i) => {
    const top = topData[i] || {};
    return [
      r.year,
      fmt(r.count),
      fmt(r.sales),
      fmt(r.revenue, 2),
      fmt(r.avgPrice, 2),
      fmtPct(r.yoySales),
      fmtPct(r.yoyRevenue),
      fmt(top.count),
      fmt(top.sales),
      fmt(top.revenue, 2),
      fmtPct(top.yoySales),
      fmtPct(top.yoyRevenue),r.coverage,r.yoyPeriod || '无共同基期',fmt(r.currentComparable?.sales),fmt(r.previousComparable?.sales),fmt(r.currentComparable?.revenue,2),fmt(r.previousComparable?.revenue,2)
    ];
  });
  return table(['年份', prefix + 'BSR候选行月次', prefix + '销量', prefix + '销售额($)', prefix + '平均标价($)', prefix + '销量YOY', prefix + '销售额YOY', 'BSR候选行月次（回勾）', 'BSR候选行销量（回勾）', 'BSR候选行销售额($)（回勾）', 'BSR候选行销量YOY', 'BSR候选行销售额YOY','本期覆盖月份','YOY比较期间','同比本期销量','同比基期销量','同比本期金额($)','同比基期金额($)'], rows);
}

function tierTable(data, title) {
  const rows = data.map((r) => [
    r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
    fmt(r.count),
    fmt(r.sales),
    fmt(r.revenue, 2),
    fmtPct(r.momSales),
    fmtPct(r.momRevenue)
  ]);
  return '<h4>' + esc(title) + '</h4>' + table(['月份', '商品数', '销量', '销售额($)', '销量MOM', '销售额MOM'], rows);
}

function narrative(category, overall, pp, nonpp) {
  const latest = lastNonEmpty(category.monthly);
  const top = category.topMonthly.find((r) => r.month === latest.month) || {};
  const pieces = [];
  pieces.push('截至 ' + latest.month.slice(0, 4) + '.' + Number(latest.month.slice(4, 6)) + '，' + category.title + '的BSR候选行池共有 ' + fmt(latest.count) + ' 条Listing月次，销量 ' + fmt(latest.sales) + '，销售额 $' + fmt(latest.revenue, 2) + '，平均标价 $' + fmt(latest.avgPrice, 2) + '。');
  pieces.push('与去年同月相比（MOM），销量' + (latest.momSales === null ? '缺少可比月份' : (latest.momSales >= 0 ? '增长 ' : '下降 ') + fmtPct(Math.abs(latest.momSales))) + '，销售额' + (latest.momRevenue === null ? '缺少可比月份' : (latest.momRevenue >= 0 ? '增长 ' : '下降 ') + fmtPct(Math.abs(latest.momRevenue)) ) + '。');
  pieces.push('该分析池由原始行先筛 BSR 1—100（包含100），再保留每条候选行，不按父ASIN去重；本月保留 ' + fmt(top.count) + ' 条候选Listing行，父ASIN仅用于重复诊断，源候选行 ' + fmt(category.topCandidateMonthly.find((r) => r.month === latest.month)?.rawCount) + ' 行。');
  if (category.title === '整体市场') {
    const p = lastNonEmpty(pp.monthly);
    const n = lastNonEmpty(nonpp.monthly);
    pieces.push('整体市场按完整单词 plastic 划分为 PP 与高客单价市场，两者并集构成整体市场；最新月 PP 销量为 ' + fmt(p.sales) + '，高客单价市场销量为 ' + fmt(n.sales) + '，请结合覆盖率和缺失值复核方向。');
  }
  pieces.push('原始销量覆盖率 ' + fmtPct(latest.salesCoverage) + '，销售额覆盖率 ' + fmtPct(latest.revenueCoverage) + '；空值保留为空，不把缺失当作零。');
  pieces.push('可比性：'+latest.quality+'；本月 '+fmt(latest.count)+' 个Listing，去年同月 '+fmt(latest.comparisonCount)+' 个。样本变化超过30%或销售字段覆盖不足95%会提示复核，阈值是审查政策，不是市场事实。');
  pieces.push('配对成交均价 $'+fmt(latest.pairedAsp,2)+'，去年同月变化 '+fmtPct(latest.momPairedAsp)+'；该价格来自同月有有效销量及金额的样本，不能据此认定同款涨价。');
  const year=category.annual.at(-1);
  pieces.push('最近年度共同周期 '+(year.yoyPeriod || '无共同基期')+'：销量 '+fmt(year.previousComparable?.sales)+' → '+fmt(year.currentComparable?.sales)+'（'+fmtPct(year.yoySales)+'）；金额 $'+fmt(year.previousComparable?.revenue,2)+' → $'+fmt(year.currentComparable?.revenue,2)+'（'+fmtPct(year.yoyRevenue)+'）。未完结年度不与去年全年直接比较。');
  const efficiencies=BANDS.map(b=>{const x=category.annualBands[b.key].at(-1);return {b,x,eff:x.count>0&&Number.isFinite(x.sales)?x.sales/x.count:null};});
  for(const {b,x,eff} of efficiencies) pieces.push('核心期 '+year.year+' ['+year.coverage+'] '+b.name+'：销量 '+fmt(x.sales)+'，'+fmt(x.count)+' Listing月次，单Listing月次销量 '+fmt(eff,2)+'；同周期销量YOY '+fmtPct(x.yoySales)+'。');
  const best=efficiencies.filter(x=>Number.isFinite(x.eff)).sort((a,b)=>b.eff-a.eff)[0];
  pieces.push(best?'建议：优先检查 '+best.b.name+' 现有链接的销售效率和榜单保留情况，再小批验证候选产品；高效率不证明容易进入该排名。若连续两个月金额、份额与BSR同步改善再考虑增加链接；可比性受限时先核验源样本。':'建议：没有足够分层样本，不给出确定的扩量层级。');
  if(category.marketStrategy) pieces.push(...category.marketStrategy);
  return pieces;
}

function annualBandTable(category){
  return table(['年份','层级','Listing月次','销量','销售额($)','标价($)','成交均价($)','销量YOY','销售额YOY','覆盖月份','比较期间','本期可比销量','基期可比销量','本期可比金额($)','基期可比金额($)'], [...BANDS,...FINE_BANDS].flatMap(b=>category.annualBands[b.key].map(r=>[r.year,b.name,fmt(r.count),fmt(r.sales),fmt(r.revenue,2),fmt(r.avgPrice,2),fmt(r.pairedAsp,2),fmtPct(r.yoySales),fmtPct(r.yoyRevenue),r.coverage,r.yoyPeriod || '无共同基期',fmt(r.currentComparable?.sales),fmt(r.previousComparable?.sales),fmt(r.currentComparable?.revenue,2),fmt(r.previousComparable?.revenue,2)])));
}

function compositionTable(category){
  return '<h4>BSR候选行池市场组成（PP + 高客单价 = 整体）</h4>'+table(['月份','整体候选行数','PP候选行数','高客单价候选行数','整体销量','PP销量','高客单价销量','整体金额($)','PP金额($)','高客单价金额($)'],category.composition.map(r=>[r.month,...['count','sales','revenue'].flatMap(k=>['overall','pp','high'].map(s=>fmt(r[s][k],k==='revenue'?2:0)))]));
}

function marketSection(id, number, category, overall, pp, nonpp) {
  const sectionNo = number === '第一部分' ? '1' : number === '第二部分' ? '2' : '3';
  const latest = lastNonEmpty(category.monthly);
  const topLatest = category.topMonthly.find((r) => r.month === latest.month) || {};
  const fineCombined = MONTHS.map((month, i) => {
    const row = { month };
    for (const band of FINE_BANDS) row['mom_' + band.key] = category.fine[band.key][i] ? category.fine[band.key][i].momSales : null;
    return row;
  });
  const tierCombined = MONTHS.map((month, i) => {
    const row = { month };
    for (const band of BANDS) row['mom_' + band.key] = category.tiers[band.key][i] ? category.tiers[band.key][i].momSales : null;
    return row;
  });
  let out = '<section id="' + id + '"><h2>' + number + '｜' + esc(category.title) + '</h2>';
  out += '<p class="lead">本部分按“BSR候选行月度 → 年度YOY → 趋势图 → 头部/中部/尾部与五档 → 文字分析”展开。分析池先筛 BSR 1—100（含100），保留每条候选行，不按父ASIN去重；月度 MOM 为本月与去年同月比较，年度 YOY 为本年与上一年度共同月份比较。</p>';
  out += '<p class="scope-notice">最新月：'+esc(latest.quality)+'。这是一份原始导出样本的分析；样本变化可能影响增长率，表格保留两期数量与覆盖率。</p>';
  out += '<div class="cards"><div class="card"><b>最新月整体销量</b><strong>' + fmt(latest.sales) + '</strong><small>' + esc(latest.month) + '</small></div>';
  out += '<div class="card"><b>最新月整体销售额</b><strong>$' + fmt(latest.revenue, 2) + '</strong><small>原始月销售额合计</small></div>';
  out += '<div class="card"><b>最新月 BSR候选行销量</b><strong>' + fmt(topLatest.sales) + '</strong><small>' + fmt(topLatest.count) + ' 条候选行</small></div>';
  out += '<div class="card"><b>最新月平均标价</b><strong>$' + fmt(latest.avgPrice, 2) + '</strong><small>按有价格记录算术平均</small></div></div>';
  out += subsection(id + '-1-1', sectionNo + '.1 月度汇总与可回勾数据', trendTable(category.monthly, category.topMonthly, category.topCandidateMonthly));
  out += subsection(id + '-1-2', sectionNo + '.2 BSR候选行月度明细', topTrendTable(category.topMonthly, category.topCandidateMonthly)+(category.composition?compositionTable(category):''));
  out += subsection(id + '-1-3', sectionNo + '.3 年度 YOY 汇总', annualTable(category.annual, category.topAnnual));
  let charts = '<div class="charts">' + svgBarChart(category.title + '月销量', category.monthly, 'sales', '#2563eb', 0);
  charts += svgBarChart(category.title + '月销售额', category.monthly, 'revenue', '#0f766e', 2);
  charts += svgLineChart(category.title + '平均标价趋势', category.monthly, [{ name: '平均标价($)', field: 'avgPrice', color: '#7c3aed' }]);
  charts += svgLineChart(category.title + '销量MOM（去年同月）', category.monthly, [{ name: '销量MOM', field: 'momSales', color: '#2563eb' }], true) + '</div>';
  out += subsection(id + '-1-4', sectionNo + '.4 市场趋势图', charts);
  let rankCharts = '<p>BSR候选行池先按可解析的小类BSR筛选 1—100（含100），保留每条候选行，不按父ASIN去重；五档互斥为 1-5、6-10、11-20、21-50、51-100。MOM统一比较去年同月。</p><div class="charts">';
  rankCharts += svgLineChart(category.title + ' BSR候选行五档销量MOM', fineCombined, FINE_BANDS.map((band, i) => ({ name: band.name, field: 'mom_' + band.key, color: ['#1d4ed8', '#0891b2', '#059669', '#f59e0b', '#dc2626'][i] })), true);
  rankCharts += svgLineChart(category.title + ' BSR候选行头/中/尾销量MOM', tierCombined, BANDS.map((band, i) => ({ name: band.name, field: 'mom_' + band.key, color: ['#2563eb', '#059669', '#f97316'][i] })), true) + '</div>';
  out += subsection(id + '-1-5', sectionNo + '.5 BSR候选行五档与头中尾 MOM', rankCharts);
  let tierBody = '';
  for (const band of [...BANDS,...FINE_BANDS]) tierBody += tierTable((category.tiers[band.key] || category.fine[band.key]), band.name);
  tierBody += '<h4>年度同周期分层</h4>'+annualBandTable(category);
  out += subsection(id + '-1-6', sectionNo + '.6 头部 / 中部 / 尾部明细', tierBody);
  const analysisItems = narrative(category, overall, pp, nonpp).map((paragraph) => '<li>' + esc(paragraph) + '</li>').join('');
  out += subsection(id + '-1-7', sectionNo + '.7 数据驱动文字分析', '<div class="analysis"><ul class="analysis-list">' + analysisItems + '</ul></div>');
  out += '</section>';
  return out;
}

function genimoSection(category, overall, pp, genimoPP) {
  const rows = category.monthly.map((r, i) => {
    const base = overall.monthly[i] || {};
    const top = category.topMonthly[i] || {};
    return [
      r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
      fmt(r.count),
      fmt(r.sales),
      fmt(r.revenue, 2),
      fmt(r.avgPrice, 2),
      fmtPct(Number.isFinite(r.sales) && Number.isFinite(base.sales) && base.sales !== 0 ? r.sales / base.sales : null),
      fmtPct(Number.isFinite(r.revenue) && Number.isFinite(base.revenue) && base.revenue !== 0 ? r.revenue / base.revenue : null),
      fmtPct(r.momSales),
      fmtPct(r.momRevenue),
      fmt(top.count)
    ];
  });
  let out = '<section id="genimo"><h2>第四部分｜GENIMO 品牌分析</h2>';
  out += '<p class="lead">GENIMO 是整体市场中的品牌视角。本部分从同一 BSR 1—100 候选行池中按品牌筛选，给出 GENIMO 在整体市场和 PP 市场中的份额、BSR分层、进留退与基于原始数据的 2027 年行动建议。</p>';
  out += subsection('genimo-4-1', '4.1 月度表现与整体市场份额', table(['月份', 'GENIMO候选行数', '销量', '销售额($)', '平均标价($)', '销量占整体', '销售额占整体', '销量MOM', '销售额MOM', 'BSR候选行数'], rows));
  const ppRows = genimoPP.monthly.map((r, i) => {
    const base = pp.monthly[i] || {};
    return [
      r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
      fmt(r.count),
      fmt(r.sales),
      fmt(r.revenue, 2),
      fmt(r.avgPrice, 2),
      fmtPct(Number.isFinite(r.sales) && Number.isFinite(base.sales) && base.sales !== 0 ? r.sales / base.sales : null),
      fmtPct(Number.isFinite(r.revenue) && Number.isFinite(base.revenue) && base.revenue !== 0 ? r.revenue / base.revenue : null),
      fmtPct(r.momSales),
      fmtPct(r.momRevenue)
    ];
  });
  let ppBody = '<p>PP市场份额分母为PP候选行池；GENIMO PP来自同一 BSR 1—100 候选行池中的品牌与PP交集。</p>';
  ppBody += table(['月份', 'GENIMO PP Listing数', 'PP内销量', 'PP内销售额($)', '平均标价($)', '销量占PP', '销售额占PP', '销量MOM', '销售额MOM'], ppRows);
  ppBody += '<h4>GENIMO PP年度YOY</h4>' + annualTable(genimoPP.annual, genimoPP.topAnnual, 'GENIMO PP');
  ppBody += '<div class="charts">' + svgBarChart('GENIMO PP 月销量', genimoPP.monthly, 'sales', '#be185d', 0) + svgBarChart('GENIMO PP 月销售额', genimoPP.monthly, 'revenue', '#9d174d', 2) + '</div>';
  out += subsection('genimo-4-2', '4.2 GENIMO 在 PP 市场中的表现', ppBody);
  let genimoCharts = '<div class="charts">' + svgBarChart('GENIMO 月销量', category.monthly, 'sales', '#9333ea', 0);
  genimoCharts += svgBarChart('GENIMO 月销售额', category.monthly, 'revenue', '#c026d3', 2);
  genimoCharts += svgLineChart('GENIMO平均标价',category.monthly,[{name:'平均标价($)',field:'avgPrice',color:'#9333ea'}]);
  genimoCharts += svgLineChart('GENIMO 销量MOM（去年同月）', category.monthly, [{ name: '销量MOM', field: 'momSales', color: '#2563eb' }], true) + '</div>';
  out += subsection('genimo-4-3', '4.3 GENIMO 趋势图', genimoCharts);
  let genimoTierBody = '<h4>GENIMO BSR候选行明细</h4>'+topTrendTable(category.topMonthly,category.topCandidateMonthly);
  for (const band of [...BANDS,...FINE_BANDS]) genimoTierBody += tierTable(category.tiers[band.key] || category.fine[band.key], 'GENIMO ' + band.name);
  for(const bs of [BANDS,FINE_BANDS]){const combined=MONTHS.map((month,i)=>({month,...Object.fromEntries(bs.map(b=>[b.key,(category.tiers[b.key]||category.fine[b.key])[i].momSales]))}));genimoTierBody+=svgLineChart('GENIMO整体榜内分层销量MOM',combined,bs.map((b,i)=>({name:b.name,field:b.key,color:['#2563eb','#059669','#f97316','#9333ea','#dc2626'][i]})),true);}
  out += subsection('genimo-4-4', '4.4 GENIMO BSR 头中尾与五档', genimoTierBody);
  out += subsection('genimo-4-5', '4.5 GENIMO 年度 YOY 与分层', annualTable(category.annual, category.topAnnual)+annualBandTable(category));
  const latest = lastNonEmpty(category.monthly);
  const latestTop = category.topMonthly.find((r) => r.month === latest.month) || {};
  const advice = [
    '先稳定能够进入 BSR 1—100 的链接：最新月 GENIMO 候选行池有 ' + fmt(latestTop.count) + ' 条候选行，销量 ' + fmt(latestTop.sales) + '。产品测试和链接补充应按 1-20、21-50、51-100 三个层级分别记录，不用单一总排名替代层级判断。',
    '用整体市场份额和 PP/高客单价市场拆分制定资源优先级：当 GENIMO 在某一市场的销量或销售额占比连续上升时，优先补充该市场相同字段完整且 BSR 可追踪的链接；当份额下降时，先核对缺失销量、销售额、价格和 BSR 的覆盖率，再决定是否调整产品组合。',
    '建立月度复盘表：销量、销售额、平均标价、销量MOM、销售额MOM、BSR候选行数量、头中尾占比和候选行标题 plastic 标记必须同批次留痕。报告只使用原始工作簿已有字段，不推导利润、成本、广告或库存指标。'
  ].map((item) => '<li>' + esc(item) + '</li>').join('');
  const evidence=narrative(category,overall,pp,null).map(t=>'<li>'+esc(t)+'</li>').join('');
  const planRows=BANDS.map(b=>{const r=category.annualBands[b.key].at(-1);const total=category.topAnnual.at(-1);return [b.name,fmt(r.sales),fmt(r.count),fmt(r.count>0?r.sales/r.count:null,2),fmtPct(total.sales>0?r.sales/total.sales:null),fmt(r.avgPrice,2),'待输入总链接数'];});
  out += subsection('genimo-4-6', '4.6 2027 年建议（仅基于当前原始字段）', '<div class="analysis"><ul class="analysis-list">' + evidence + advice + '<li>新增链接总量在XLSX 08页B3输入，默认空；按上表最近年度核心期品牌榜内销量权重分配，前两档向下取整、尾档取余。是资源分配情景，不是销量预测。</li><li>花型、颜色和尺寸尚未可靠提取，不能据此建议具体花型；不推导利润、广告或库存。</li></ul></div>'+table(['目标BSR层级','核心期销量','Listing月次','单Listing月次销量','分配权重','历史标价($)','新增链接数'],planRows));
  out += subsection('genimo-4-7','4.7 GENIMO BSR候选行进留退', '<p>比较相邻自然月候选行池内的Listing键集合；进入/退出不等于上新/下架。父ASIN只作为审计字段。</p>'+table(['月份','基期月份','当前','基期','进入','保留','退出'],category.movements.map(r=>[r.month,r.previous,...['current','prior','entered','retained','exited'].map(k=>fmt(r[k]))]))+table(['月份','Listing键','基期BSR','本期BSR','状态'],category.movements.flatMap(r=>r.details.map(x=>[r.month,x.key,fmt(x.previousRank),fmt(x.currentRank),x.state]))));
  out += '</section>';
  return out;
}

function rawFieldCoverage(rows) {
  const fields = [
    ['商品标题', 'title'],
    ['品牌', 'brand'],
    ['父ASIN', 'parent'],
    ['小类BSR', 'rank'],
    ['月销量', 'sales'],
    ['月销售额($)', 'revenue'],
    ['价格($)', 'price']
  ];
  const result = fields.map(([name, field]) => {
    // Text columns (title/brand/parent) must be counted by non-empty text;
    // numeric columns use the parsed numeric value. Treating every field as
    // numeric makes valid text coverage appear as 0.0%.
    const present = rows.filter((row) => field === 'rank'
      ? row.rank !== null
      : clean(row[field]) !== '').length;
    return [name, fmt(present), fmtPct(rows.length ? present / rows.length : null), '参与对应的市场统计或筛选'];
  });
  return table(['原始字段', '非空/有效行数', '覆盖率', '处理方式'], result);
}

function buildHtml(raw, categories) {
  const rawCount = raw.rows.length;
  const ranked = raw.rows.filter((r) => r.rank !== null).length;
  const topCandidates = raw.rows.filter((r) => r.rank !== null && r.rank >= 1 && r.rank <= 100).length;
  const sourceMonthsWithData = raw.sheetStats.filter((item) => item.rowCount > 0).length;
  const detectedHeaderRows = Array.from(new Set(raw.sheetStats.map((item) => item.headerRow))).sort((a, b) => a - b);
  const auditMonth = '202509';
  const auditIndex = MONTHS.indexOf(auditMonth);
  const auditOverall = auditIndex >= 0 ? categories.overall.topCandidateMonthly[auditIndex] : null;
  const auditPp = auditIndex >= 0 ? categories.pp.topCandidateMonthly[auditIndex] : null;
  const auditNonpp = auditIndex >= 0 ? categories.nonpp.topCandidateMonthly[auditIndex] : null;
  const auditTitlePpRows = raw.rows.filter((r) => r.month === auditMonth && r.rank !== null && r.rank >= 1 && r.rank <= 100 && isPlasticTitle(r.title)).length;
  const generated = new Date().toISOString().replace('T', ' ').replace(/\.\d+Z$/, ' UTC');
  const compatibilityCss = '.notice,.meta{padding:14px 16px;margin:14px 0;border:1px solid var(--line);background:var(--surface-soft);color:var(--muted)}.notice{border-left:2px solid var(--warning)}.notice b,.meta b{color:var(--text)}.cards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;margin:22px 0}.card{min-height:120px;padding:18px;border:1px solid var(--line);background:var(--card-bg);box-shadow:var(--shadow)}.card b,.card small{display:block;color:var(--muted);font-size:14px}.card strong{display:block;margin:13px 0 7px;font-family:var(--display);font-size:32px;font-weight:400}.charts{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.chart{min-width:0;margin:0;padding:14px;border:1px solid var(--line);background:var(--surface-soft);overflow:auto}.chart h4{margin:0;color:var(--text)}.chart svg{display:block;width:100%;height:auto;min-width:390px}.axis-label,.legend-label{font-size:10px;fill:var(--muted)}.analysis{padding:15px;border:1px solid var(--line);background:var(--surface-soft)}.analysis p{margin:8px 0;color:var(--muted)}footer{margin-top:36px;padding-top:16px;border-top:1px solid var(--line);color:var(--muted);font-size:13px}@media(max-width:980px){.cards{grid-template-columns:repeat(2,minmax(0,1fr))}.charts{grid-template-columns:1fr}}@media(max-width:560px){.cards{grid-template-columns:1fr}}';
  const navGroup = (target, icon, label, items) => {
    const links = items.map(([href, text]) => '<a href="#' + esc(href) + '">' + esc(text) + '</a>').join('');
    return '<details class="nav-group" data-target="' + esc(target) + '" open><summary><span class="nav-icon">' + esc(icon) + '</span><span>' + esc(label) + '</span><span class="nav-chevron" aria-hidden="true">⌄</span></summary><div class="nav-subitems">' + links + '</div></details>';
  };
  const nav = navGroup('dashboard', '总', '数据总览', [
    ['dashboard-0-1', '0.1 范围与关键指标'],
    ['dashboard-0-2', '0.2 数据口径与来源'],
    ['dashboard-0-3', '0.3 原始字段覆盖']
  ]) + navGroup('overall', '整', '整体市场', [
    ['overall-1-1', '1.1 月度汇总与可回勾数据'],
    ['overall-1-2', '1.2 BSR候选行月度明细'],
    ['overall-1-3', '1.3 年度 YOY 汇总'],
    ['overall-1-4', '1.4 市场趋势图'],
    ['overall-1-5', '1.5 BSR候选行五档与头中尾 MOM'],
    ['overall-1-6', '1.6 头部 / 中部 / 尾部明细'],
    ['overall-1-7', '1.7 数据驱动文字分析']
  ]) + navGroup('pp', 'PP', 'PP市场', [
    ['pp-1-1', '2.1 月度汇总与可回勾数据'],
    ['pp-1-2', '2.2 BSR候选行月度明细'],
    ['pp-1-3', '2.3 年度 YOY 汇总'],
    ['pp-1-4', '2.4 市场趋势图'],
    ['pp-1-5', '2.5 BSR候选行五档与头中尾 MOM'],
    ['pp-1-6', '2.6 头部 / 中部 / 尾部明细'],
    ['pp-1-7', '2.7 数据驱动文字分析']
  ]) + navGroup('nonpp', '高', '高客单价市场', [
    ['nonpp-1-1', '3.1 月度汇总与可回勾数据'],
    ['nonpp-1-2', '3.2 BSR候选行月度明细'],
    ['nonpp-1-3', '3.3 年度 YOY 汇总'],
    ['nonpp-1-4', '3.4 市场趋势图'],
    ['nonpp-1-5', '3.5 BSR候选行五档与头中尾 MOM'],
    ['nonpp-1-6', '3.6 头部 / 中部 / 尾部明细'],
    ['nonpp-1-7', '3.7 数据驱动文字分析']
  ]) + navGroup('genimo', 'G', 'GENIMO品牌', [
    ['genimo-4-1', '4.1 月度表现与整体市场份额'],
    ['genimo-4-2', '4.2 GENIMO 在 PP 市场中的表现'],
    ['genimo-4-3', '4.3 GENIMO 趋势图'],
    ['genimo-4-4', '4.4 GENIMO BSR 头中尾与五档'],
    ['genimo-4-5', '4.5 GENIMO 年度 YOY'],
    ['genimo-4-6', '4.6 2027 年建议'],
    ['genimo-4-7', '4.7 品牌进留退']
  ]);
  let html = '<!doctype html><html lang="zh-CN" data-theme="dark"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>户外地垫市场洞察 · Outdoor Rug Intelligence</title><style>' + UI_CSS + compatibilityCss + '</style></head><body><div class="app-shell">';
  html += '<aside class="sidebar"><div class="brand"><div class="brand-mark"></div><div><strong>Market Intelligence</strong><small>户外地垫市场分析系统</small></div></div><div class="local-badge"><i></i> DATA · VERIFIED</div><span class="nav-label">市场分析</span><nav>' + nav + '</nav><div class="sidebar-footer"><span>可比数据范围</span><strong>' + esc(MONTHS[0]) + ' — ' + esc(MONTHS[MONTHS.length - 1]) + '</strong><small>' + fmt(MONTHS.length) + '个月 · 原始主源直读</small></div></aside>';
  html += '<div class="workspace"><header class="topbar"><div><span class="eyebrow">MARKET OVERVIEW · SPEC 2.1</span><h1>户外地垫市场分析</h1><p>统一 BSR 候选行池 · 整体市场、PP市场、高客单价市场与 GENIMO 品牌 · 销量、销售额、均价 · MOM / YOY</p></div><div class="top-actions"><span class="privacy-chip"><i></i> 源数据只读</span><button class="theme-button" id="theme-toggle" aria-label="切换主题">☀</button></div></header><main class="content">';
  html += '<section id="dashboard" class="overview-section"><div class="section-kicker">报告导航 · 0</div><h2>数据总览</h2><p class="lead">先确认数据范围、统计顺序和候选行口径，再进入四个分析部分。每个小节都可以点击左侧导航定位，也可以点击标题前的小三角收起。</p>';
  html += subsection('dashboard-0-1', '0.1 分析范围与关键指标', '<div class="scope-notice"><span>◎</span><div><b>分析范围：</b>原始工作簿覆盖 ' + MONTHS.length + ' 个月；每月先筛 BSR 1—100（包含100），保留每条候选源行，不按父ASIN去重。整体市场由 PP 与高客单价补集构成，GENIMO 作为同一候选行池中的品牌视角。</div></div><div class="metrics-grid"><article class="metric-card"><span class="metric-label">原始有效行</span><strong class="metric-value">' + fmt(rawCount) + '</strong><span class="metric-note">逐月明细读取</span></article><article class="metric-card"><span class="metric-label">BSR候选原始行</span><strong class="metric-value">' + fmt(topCandidates) + '</strong><span class="metric-note">BSR 1—100（含100）</span></article><article class="metric-card"><span class="metric-label">BSR候选行月次</span><strong class="metric-value">' + fmt(categories.overall.topRows.length) + '</strong><span class="metric-note">每条候选行直接计入</span></article><article class="metric-card"><span class="metric-label">GENIMO候选行月次</span><strong class="metric-value">' + fmt(categories.genimo.fullRows.length) + '</strong><span class="metric-note">候选行品牌精确匹配</span></article></div>');
  const scopeBullets = [
    '处理顺序：BSR 1—100（含100）→ 保留每条候选源行 → 候选行标题/品牌分类；父ASIN只作审计字段。',
    '整体市场 = PP市场 ∪ 高客单价市场，两者是同一BSR候选行池中的互补分区。',
    'PP市场使用候选行商品标题中的完整单词 plastic 匹配；高客单价市场为未匹配 plastic 的补集。',
    '月度 MOM 比较本月与去年同月；年度 YOY 比较本年与上一年度，未完结年度按实际覆盖范围标记。',
    '数据源：' + esc(SOURCE) + '；SHA-256：' + esc(SOURCE_HASH) + '。',
    '可解析小类BSR ' + fmt(ranked) + ' 行，BSR 1—100（包含100）的候选 ' + fmt(topCandidates) + ' 行；月度子表有数据 ' + fmt(sourceMonthsWithData) + '/' + fmt(raw.sheetStats.length) + '，识别表头行 ' + esc(detectedHeaderRows.join('、')) + '。',
    auditOverall ? '抽查 2025.09：BSR 1—100 候选 ' + fmt(auditOverall.rawCount) + ' 行，全部保留为 ' + fmt(auditOverall.count) + ' 条分析行；按候选行标题分类，PP为 ' + fmt(auditPp.count) + ' 条，高客单价为 ' + fmt(auditNonpp.count) + ' 条，合计回到 ' + fmt(auditOverall.count) + ' 条。源表逐行标题含 plastic 的 ' + fmt(auditTitlePpRows) + ' 行与PP候选行一致。' : '',
    '缺失值保留为空，不当作零；原始字段只用于市场统计、筛选和回勾。'
  ].map((item) => '<li>' + item + '</li>').join('');
  const statUnitBullets = [
    '统计单元：每条BSR合格候选源行一个Listing月次；ASIN和源行号用于回勾，父ASIN保留作重复诊断。',
    '代表行：当前口径不选代表行、不合并父ASIN；91页每条候选行直接回勾90页对应源行。',
    '原始字段覆盖表：仅做只读回勾，不参与任何利润推导。'
  ].map((item) => '<li>' + item + '</li>').join('');
  const scopeBody = '<div class="scope-notice"><span>◎</span><div><b>口径与来源：</b><ul class="analysis-list compact-list">' + scopeBullets + '</ul></div></div><div class="meta"><b>统计单元与回勾方式</b><ul class="analysis-list compact-list">' + statUnitBullets + '</ul></div>';
  html += subsection('dashboard-0-2', '0.2 数据口径与来源', scopeBody);
  html += subsection('dashboard-0-3', '0.3 原始字段覆盖（只读）', rawFieldCoverage(raw.rows));
  html += '</section>';
  html += marketSection('overall', '第一部分', categories.overall, categories.overall, categories.pp, categories.nonpp);
  html += marketSection('pp', '第二部分', categories.pp, categories.overall, categories.pp, categories.nonpp);
  html += marketSection('nonpp', '第三部分', categories.nonpp, categories.overall, categories.pp, categories.nonpp);
  html += genimoSection(categories.genimo, categories.overall, categories.pp, categories.genimoPP);
  html += '<footer>报告由 src/build_spec2_html.mjs 从原始工作簿生成。市场总览与数据口径已合并；左侧导航可展开到每个小节，正文小节可用小三角折叠。图表支持悬停与键盘聚焦查看精确值。源表发生更换时，请先确认文件身份、月份覆盖和列名，再重新生成。</footer></main></div></div><script>' + chartRuntimeScript() + '</script></body></html>';
  return html;
}

function addMetric(a, b) {
  if (Number.isFinite(a) && Number.isFinite(b)) return a + b;
  if (Number.isFinite(a)) return a;
  if (Number.isFinite(b)) return b;
  return null;
}

function metricEqual(a, b) {
  if (!Number.isFinite(a) && !Number.isFinite(b)) return true;
  return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) < 1e-6;
}

function buildDataset() {
  MONTHS.length = 0;
  const raw = readRawRows();
  // The sole business pool starts with the numeric BSR predicate.  The
  // confirmed current rule retains every BSR-qualified source row and does
  // not deduplicate by parent ASIN.  Parent ASIN remains available for audit
  // and variant-family diagnostics only.
  const topCandidates = raw.rows.filter((r) => r.rank !== null && r.rank >= 1 && r.rank <= 100);
  const missingParentCandidates = topCandidates.filter((r) => !r.parent);
  const mainPool = topCandidates.map((row) => ({
    ...row,
    familyKey: row.month + '|' + listingKey(row),
    rawRows: 1,
    sourceRows: [sourceId(row)]
  }));
  // Keep historical property names as aliases for downstream builders.  They
  // now all point to the same BSR-filtered row pool; there is no second
  // unfiltered market branch and no parent-ASIN collapse.
  const fullDedup = mainPool;
  const topCandidateDedup = mainPool;
  const topDedup = mainPool;
  const ppRows = mainPool.filter((r) => r.plastic);
  const nonppRows = mainPool.filter((r) => !r.plastic);
  const genimoRows = mainPool.filter((r) => r.genimo);
  const genimoPPRows = mainPool.filter((r) => r.genimo && r.plastic);
  const topCandidatesFor = (rows) => rows;
  const ppTopCandidates = topCandidatesFor(ppRows);
  const nonppTopCandidates = topCandidatesFor(nonppRows);
  const genimoTopCandidates = topCandidatesFor(genimoRows);
  const genimoPPTopCandidates = topCandidatesFor(genimoPPRows);
  const categories = {
    // Candidate audit rows are the retained source rows.  rawRows remains 1
    // per row so the table can show the exact source-row count.
    overall: buildCategory('整体市场', mainPool, mainPool, topCandidateDedup),
    pp: buildCategory('PP市场', ppRows, ppRows, ppRows),
    nonpp: buildCategory('高客单价市场', nonppRows, nonppRows, nonppRows),
    genimo: buildCategory('GENIMO品牌', genimoRows, genimoRows, genimoRows),
    genimoPP: buildCategory('GENIMO PP市场', genimoPPRows, genimoPPRows, genimoPPRows),
    // Compatibility aliases used by the workbook's composition section. All
    // aliases are subsets of the same mainPool and therefore cannot re-open
    // an unfiltered branch.
    globalPP: buildCategory('整体市场中的PP', ppRows, ppRows, ppRows),
    globalHigh: buildCategory('整体市场中的高客单价', nonppRows, nonppRows, nonppRows),
    genimoIndependent: buildCategory('GENIMO品牌（同一BSR候选行池回勾）', genimoRows, genimoRows, genimoRows)
  };
  for (let i = 0; i < MONTHS.length; i += 1) {
    const overallMonth = categories.overall.monthly[i];
    const ppMonth = categories.pp.monthly[i];
    const nonppMonth = categories.nonpp.monthly[i];
    const fullSales = addMetric(ppMonth.sales, nonppMonth.sales);
    const fullRevenue = addMetric(ppMonth.revenue, nonppMonth.revenue);
    if (overallMonth.count !== ppMonth.count + nonppMonth.count ||
        !metricEqual(overallMonth.sales, fullSales) ||
        !metricEqual(overallMonth.revenue, fullRevenue)) {
      throw new Error('PP/高客单价市场互补校验失败: ' + MONTHS[i]);
    }
  }
  if (!raw.rows.some((row) => row.rank === 100)) throw new Error('原始主源没有检测到 BSR=100，无法确认包含100');
  const movements = MONTHS.map(month => {
    const previous = previousMonth(month);
    const currentRows = categories.genimo.topRows.filter(r => r.month === month);
    const priorRows = categories.genimo.topRows.filter(r => r.month === previous);
    const before = new Map(priorRows.map(r => [listingKey(r), r]));
    const after = new Map(currentRows.map(r => [listingKey(r), r]));
    const available = MONTHS.includes(previous);
    const details = [...new Set([...before.keys(), ...after.keys()])].sort().map(key => ({key, previousRank: before.get(key)?.rank ?? null, currentRank: after.get(key)?.rank ?? null, state: !available ? '无基期' : before.has(key) ? (after.has(key) ? '保留' : '退出') : '进入'}));
    return {month, previous, available, current:after.size, prior:available ? before.size : null, entered:available ? details.filter(r => r.state === '进入').length : null, retained:available ? details.filter(r => r.state === '保留').length : null, exited:available ? details.filter(r => r.state === '退出').length : null, details};
  });
  const metadata={
    specVersion:'2.1',
    sourceSha256:SOURCE_HASH,
    generatedAt:new Date().toISOString(),
    batchId:'spec21-'+SOURCE_HASH.slice(0,12),
    latestMonth:MONTHS.at(-1),
    analysisOrder:'BSR 1-100 inclusive -> retain every candidate source row (no parent ASIN dedup) -> row title/brand classification',
    bsrCandidateRows:topCandidates.length,
    bsrCandidateParentRows:new Set(topCandidates.filter((r) => r.parent).map((r) => r.parent)).size,
    analysisPoolRows:mainPool.length,
    analysisPoolListingMonths:mainPool.length,
    missingParentCandidateRows:missingParentCandidates.length,
    analysisUnit:'BSR-qualified source row identified by month/sourceRow; ASIN and parent ASIN are audit-only',
    parentAsinDedup:'not applied',
    classificationRule:'candidate-row title complete-word plastic; candidate-row brand exact GENIMO',
    overallCompositionRule:'overall = PP + high-price complement within BSR candidate rows'
  };
  categories.genimo.movements=movements;
  categories.genimo.independentTopMonthly=categories.genimoIndependent.topMonthly;
  categories.overall.composition=MONTHS.map((month,i)=>({month,overall:categories.overall.topMonthly[i],pp:categories.globalPP.topMonthly[i],high:categories.globalHigh.topMonthly[i]}));
  categories.overall.metadata=metadata;
  const py=categories.pp.annual.at(-1),hy=categories.nonpp.annual.at(-1),gy=categories.genimo.annual.at(-1),gpy=categories.genimoPP.annual.at(-1);
  const share=(a,b)=>Number.isFinite(a)&&Number.isFinite(b)&&b>0?a/b:null;
  const target=Number.isFinite(py.yoyRevenue)&&Number.isFinite(hy.yoyRevenue)?(py.yoyRevenue>=hy.yoyRevenue?'PP':'高客单价'):null;
  categories.genimo.marketStrategy=[
    '市场选择证据：'+py.yoyPeriod+'，PP金额YOY '+fmtPct(py.yoyRevenue)+'，高客单价金额YOY '+fmtPct(hy.yoyRevenue)+'；GENIMO在PP内的金额份额 '+fmtPct(share(gpy.previousComparable.revenue,py.previousComparable.revenue))+' → '+fmtPct(share(gpy.currentComparable.revenue,py.currentComparable.revenue))+'。',
    target?'2027建议：先核验 '+target+' 市场，再作为小批测款的优先候选，其同周期金额变化相对更强；现有PP品牌盘应单独跟踪份额，不能因为另一个市场相对更强就直接迁移全部链接。确认样本可比后，用连续两个月份额、金额与BSR变化决定追加或收缩。':'2027建议：比较期证据不足，暂不指定优先市场。'
  ];
  for(const cat of Object.values(categories)) {
    for(const collection of [cat.monthly,cat.topMonthly,...Object.values(cat.tiers),...Object.values(cat.fine)]) for(const r of collection){
      const current=raw.sheetStats.find(s=>s.month===r.month)?.rowCount;
      const priorYear=raw.sheetStats.find(s=>s.month===previousYear(r.month))?.rowCount;
      const priorMonth=raw.sheetStats.find(s=>s.month===previousMonth(r.month))?.rowCount;
      r.rawCount=current;r.rawPreviousYearCount=priorYear ?? null;
      if(r.count && (Math.abs(percent(current,priorYear) ?? 0)>.3 || Math.abs(percent(current,priorMonth) ?? 0)>.3)) r.quality='可比性受限';
    }
  }
  return { raw, fullDedup, mainPool, topCandidates, topCandidateDedup, topDedup, categories, movements, metadata };
}

function runHtmlBuild() {
  const dataset = buildDataset();
  fs.mkdirSync(path.dirname(OUTPUT), { recursive: true });
  fs.writeFileSync(OUTPUT, buildHtml(dataset.raw, dataset.categories), 'utf8');
  console.log(JSON.stringify({
    output: OUTPUT,
    months: MONTHS.length,
    rawRows: dataset.raw.rows.length,
    sourceMonthsWithData: dataset.raw.sheetStats.filter((item) => item.rowCount > 0).length,
    sourceMonthsEmpty: dataset.raw.sheetStats.filter((item) => item.rowCount === 0).map((item) => item.month),
    detectedHeaderRows: Array.from(new Set(dataset.raw.sheetStats.map((item) => item.headerRow))).sort((a, b) => a - b),
    sourceSha256: SOURCE_HASH,
    fullDedup: dataset.fullDedup.length,
    top100CandidateRows: dataset.topCandidates.length,
    top100CandidateDedup: dataset.topCandidateDedup.length,
    top100Dedup: dataset.topDedup.length,
    categories: Object.fromEntries(Object.entries(dataset.categories).map(([k, v]) => [k, { full: v.fullRows.length, top100: v.topRows.length }]))
  }, null, 2));
}

export {
  BANDS,
  FINE_BANDS,
  MONTHS,
  SOURCE,
  SOURCE_HASH,
  buildCategory,
  buildDataset,
  buildHtml,
  annualRows,
  representativeCompare,
  isPlasticTitle,
  summarise,
  narrative,
  clean,
  dedup,
  display,
  fmt,
  fmtPct,
  normaliseMonth,
  parseBsr,
  parseNumber,
  previousMonth,
  previousYear
};

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const {buildDelivery}=await import('./build_delivery.mjs');
  await buildDelivery();
}
