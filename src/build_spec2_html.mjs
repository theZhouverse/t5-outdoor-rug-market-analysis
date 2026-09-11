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
    .replace(/&#39;/g, "'");
}

function display(cell) {
  if (!cell) return '';
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
  const matches = s.match(/\d+(?:\.0+)?/g) || [];
  const values = matches.map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0);
  return values.length ? Math.min(...values) : null;
}

function isPlasticTitle(value) {
  return /(?:^|[^a-z])plastic(?:[^a-z]|$)/i.test(clean(value));
}

function isGenimo(value) {
  return clean(value).toLowerCase() === 'genimo';
}

function sourceId(row) {
  return row.month + '#' + row.sourceRow;
}

function listingKey(row) {
  return row.parent || row.asin || 'source:' + sourceId(row);
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
        sourceRow: r + 1,
        asin: clean(values[idx.asin]),
        sku: clean(values[idx.sku]),
        brand: clean(values[idx.brand]),
        title: clean(values[idx.title]),
        parent: clean(values[idx.parent]),
        rank: parseBsr(values[idx.smallBsr]),
        sales: parseNumber(values[idx.sales]),
        revenue: parseNumber(values[idx.revenue]),
        price: parseNumber(values[idx.price]),
        fba: parseNumber(values[idx.fba]),
        margin: clean(values[idx.margin]),
        coupon: clean(values[idx.coupon])
      };
      if (!row.asin && !row.parent && !row.title && !row.brand && row.rank === null) continue;
      allRows.push(row);
      rowCount += 1;
      if (row.rank !== null) rankCount += 1;
    }
    sheetStats.push({ month, rowCount, rankCount, headerRow: headerRow + 1 });
  }
  return { rows: allRows, sheetStats };
}

function buildFamilyFlags(rows) {
  const flags = new Map();
  for (const row of rows) {
    const key = row.month + '|' + listingKey(row);
    const current = flags.get(key) || { plastic: false, genimo: false };
    current.plastic = current.plastic || isPlasticTitle(row.title);
    current.genimo = current.genimo || isGenimo(row.brand);
    flags.set(key, current);
  }
  return flags;
}

function representativeCompare(a, b) {
  const ar = a.rank === null ? Number.POSITIVE_INFINITY : a.rank;
  const br = b.rank === null ? Number.POSITIVE_INFINITY : b.rank;
  if (ar !== br) return ar - br;
  const ac = (Number.isFinite(a.sales) ? 1 : 0) + (Number.isFinite(a.revenue) ? 1 : 0);
  const bc = (Number.isFinite(b.sales) ? 1 : 0) + (Number.isFinite(b.revenue) ? 1 : 0);
  if (ac !== bc) return bc - ac;
  if (Number.isFinite(a.price) !== Number.isFinite(b.price)) return Number.isFinite(b.price) - Number.isFinite(a.price);
  return sourceId(a).localeCompare(sourceId(b));
}

function dedup(rows, flags) {
  const groups = new Map();
  for (const row of rows) {
    const key = row.month + '|' + listingKey(row);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  const result = [];
  for (const [key, items] of groups.entries()) {
    const rep = items.slice().sort(representativeCompare)[0];
    const family = flags.get(key) || { plastic: false, genimo: false };
    result.push({
      ...rep,
      familyKey: key,
      plastic: family.plastic,
      genimo: family.genimo,
      rawRows: items.length,
      sourceRows: items.map(sourceId)
    });
  }
  return result.sort((a, b) => (a.month + a.familyKey).localeCompare(b.month + b.familyKey));
}

function percent(current, previous) {
  if (!Number.isFinite(current) || !Number.isFinite(previous) || previous === 0) return null;
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
  const years = Array.from(byYear.keys()).sort();
  const summaries = years.map((year) => ({ year, ...summarise(byYear.get(year)) }));
  return summaries.map((current, i) => {
    const previous = summaries[i - 1] || {};
    const previousYear = previous.year;
    const currentMonths = new Set(byYear.get(current.year).map((row) => row.month.slice(4, 6)));
    const previousMonths = previousYear ? new Set(byYear.get(previousYear).map((row) => row.month.slice(4, 6))) : new Set();
    const commonMonths = Array.from(currentMonths).filter((month) => previousMonths.has(month)).sort();
    const currentComparable = commonMonths.length
      ? summarise(byYear.get(current.year).filter((row) => commonMonths.includes(row.month.slice(4, 6))))
      : {};
    const previousComparable = commonMonths.length
      ? summarise(byYear.get(previousYear).filter((row) => commonMonths.includes(row.month.slice(4, 6))))
      : {};
    const yoyPeriod = previousYear && commonMonths.length
      ? previousYear + '.' + Number(commonMonths[0]) + '—' + current.year + '.' + Number(commonMonths[commonMonths.length - 1])
      : null;
    return {
      ...current,
      yoySales: percent(currentComparable.sales, previousComparable.sales),
      yoyRevenue: percent(currentComparable.revenue, previousComparable.revenue),
      yoyPeriod
    };
  });
}

function filterBand(rows, band) {
  return rows.filter((r) => r.rank !== null && r.rank >= band.lo && r.rank <= band.hi);
}

function buildCategory(title, fullRows, topRows) {
  const tierRows = {};
  for (const band of BANDS) tierRows[band.key] = filterBand(topRows, band);
  const fineRows = {};
  for (const band of FINE_BANDS) fineRows[band.key] = filterBand(topRows, band);
  return {
    title,
    fullRows,
    topRows,
    monthly: enrichTrend(MONTHS, fullRows),
    topMonthly: enrichTrend(MONTHS, topRows),
    annual: annualRows(MONTHS, fullRows),
    topAnnual: annualRows(MONTHS, topRows),
    tiers: Object.fromEntries(BANDS.map((band) => [band.key, enrichTrend(MONTHS, tierRows[band.key])])),
    fine: Object.fromEntries(FINE_BANDS.map((band) => [band.key, enrichTrend(MONTHS, fineRows[band.key])]))
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
  let body = '<div class="chart"><h4>' + esc(title) + '</h4><svg viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="' + esc(title) + '">';
  body += '<line x1="' + left + '" y1="' + (height - bottom) + '" x2="' + width + '" y2="' + (height - bottom) + '" stroke="#94a3b8"/>';
  rows.forEach((row, i) => {
    const value = values[i];
    const h = max > 0 ? (value / max) * plotH : 0;
    const x = left + i * (plotW / Math.max(rows.length, 1)) + 1;
    const y = height - bottom - h;
    body += '<rect x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" width="' + barW.toFixed(1) + '" height="' + h.toFixed(1) + '" rx="2" fill="' + color + '"><title>' + esc(row.month + '：' + fmt(row[field], digits)) + '</title></rect>';
    if (i % Math.max(1, Math.ceil(rows.length / 10)) === 0) {
      body += '<text x="' + (x + barW / 2).toFixed(1) + '" y="' + (height - 17) + '" text-anchor="middle" class="axis-label">' + esc(row.month.slice(0, 4) + '.' + Number(row.month.slice(4, 6))) + '</text>';
    }
  });
  body += '<text x="8" y="18" class="axis-label">' + esc(max > 0 ? fmt(max, digits) : '—') + '</text></svg></div>';
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
  let body = '<div class="chart"><h4>' + esc(title) + '</h4><svg viewBox="0 0 ' + width + ' ' + height + '" role="img" aria-label="' + esc(title) + '">';
  if (min < 0 && max > 0) {
    const y0 = scaleY(0);
    body += '<line x1="' + left + '" y1="' + y0.toFixed(1) + '" x2="' + (width - right) + '" y2="' + y0.toFixed(1) + '" stroke="#cbd5e1" stroke-dasharray="4 4"/>';
  }
  body += '<line x1="' + left + '" y1="' + top + '" x2="' + left + '" y2="' + (height - bottom) + '" stroke="#94a3b8"/>';
  for (const item of series) {
    const points = rows.map((row, i) => Number.isFinite(row[item.field]) ? scaleX(i).toFixed(1) + ',' + scaleY(row[item.field]).toFixed(1) : null).filter(Boolean);
    if (points.length > 1) body += '<polyline fill="none" stroke="' + item.color + '" stroke-width="2.5" points="' + points.join(' ') + '"/>';
    rows.forEach((row, i) => {
      if (Number.isFinite(row[item.field])) {
        body += '<circle cx="' + scaleX(i).toFixed(1) + '" cy="' + scaleY(row[item.field]).toFixed(1) + '" r="2.8" fill="' + item.color + '"><title>' + esc(row.month + '：' + (percentMode ? fmtPct(row[item.field]) : fmt(row[item.field], 2))) + '</title></circle>';
      }
    });
  }
  const step = Math.max(1, Math.ceil(rows.length / 10));
  rows.forEach((row, i) => {
    if (i % step === 0) body += '<text x="' + scaleX(i).toFixed(1) + '" y="' + (height - 17) + '" text-anchor="middle" class="axis-label">' + esc(row.month.slice(0, 4) + '.' + Number(row.month.slice(4, 6))) + '</text>';
  });
  series.forEach((item, i) => {
    const x = left + i * 155;
    body += '<rect x="' + x + '" y="8" width="12" height="12" fill="' + item.color + '"/><text x="' + (x + 17) + '" y="18" class="legend-label">' + esc(item.name) + '</text>';
  });
  body += '<text x="8" y="' + (top + 5) + '" class="axis-label">' + esc(percentMode ? fmtPct(max) : fmt(max, 0)) + '</text>';
  body += '<text x="8" y="' + (height - bottom) + '" class="axis-label">' + esc(percentMode ? fmtPct(min) : fmt(min, 0)) + '</text>';
  body += '</svg></div>';
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

function trendTable(data, topData) {
  const rows = data.map((r, i) => [
    r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
    fmt(r.count),
    fmt(r.sales),
    fmt(r.revenue, 2),
    fmt(r.avgPrice, 2),
    fmt(r.pairedAsp, 2),
    fmtPct(r.momSales),
    fmtPct(r.momRevenue),
    fmt(topData[i] ? topData[i].count : null)
  ]);
  return table(['月份', '整体商品数', '销量', '销售额($)', '平均标价($)', '加权成交均价($)', '销量MOM', '销售额MOM', 'BSR前100商品数'], rows);
}

function topTrendTable(data) {
  const rows = data.map((r) => [
    r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
    fmt(r.count),
    fmt(r.sales),
    fmt(r.revenue, 2),
    fmt(r.avgPrice, 2),
    fmt(r.pairedAsp, 2),
    fmtPct(r.momSales),
    fmtPct(r.momRevenue)
  ]);
  return table(['月份', 'Top100商品数', '销量', '销售额($)', '平均标价($)', '加权成交均价($)', '销量MOM', '销售额MOM'], rows);
}

function annualTable(data, topData) {
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
      fmtPct(top.yoyRevenue)
    ];
  });
  return table(['年份', '整体商品数', '整体销量', '整体销售额($)', '整体平均标价($)', '整体销量YOY', '整体销售额YOY', 'Top100商品数', 'Top100销量', 'Top100销售额($)', 'Top100销量YOY', 'Top100销售额YOY'], rows);
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
  pieces.push('截至 ' + latest.month.slice(0, 4) + '.' + Number(latest.month.slice(4, 6)) + '，' + category.title + '整体盘共有 ' + fmt(latest.count) + ' 个去重 Listing，销量 ' + fmt(latest.sales) + '，销售额 $' + fmt(latest.revenue, 2) + '，平均标价 $' + fmt(latest.avgPrice, 2) + '。');
  pieces.push('与去年同月相比（MOM），销量' + (latest.momSales === null ? '缺少可比月份' : (latest.momSales >= 0 ? '增长 ' : '下降 ') + fmtPct(Math.abs(latest.momSales))) + '，销售额' + (latest.momRevenue === null ? '缺少可比月份' : (latest.momRevenue >= 0 ? '增长 ' : '下降 ') + fmtPct(Math.abs(latest.momRevenue)) ) + '。');
  pieces.push('BSR 1—100（包含100）在该月覆盖 ' + fmt(top.count) + ' 个 Listing，销量 ' + fmt(top.sales) + '，销售额 $' + fmt(top.revenue, 2) + '；Top100与整体盘的差异用于判断头部集中度。');
  if (category.title === '整体市场') {
    const p = lastNonEmpty(pp.monthly);
    const n = lastNonEmpty(nonpp.monthly);
    pieces.push('整体市场按完整单词 plastic 划分为 PP 与高客单价市场，两者并集构成整体市场；最新月 PP 销量为 ' + fmt(p.sales) + '，高客单价市场销量为 ' + fmt(n.sales) + '，请结合覆盖率和缺失值复核方向。');
  }
  pieces.push('原始销量覆盖率 ' + fmtPct(latest.salesCoverage) + '，销售额覆盖率 ' + fmtPct(latest.revenueCoverage) + '；空值保留为空，不把缺失当作零。');
  return pieces;
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
  out += '<p class="lead">本部分按“整体盘 → BSR Top100 → 头部/中部/尾部 → 五档 → 文字分析”展开。月度 MOM 为本月与去年同月比较，年度 YOY 为本年与上一年度比较。</p>';
  out += '<div class="cards"><div class="card"><b>最新月整体销量</b><strong>' + fmt(latest.sales) + '</strong><small>' + esc(latest.month) + '</small></div>';
  out += '<div class="card"><b>最新月整体销售额</b><strong>$' + fmt(latest.revenue, 2) + '</strong><small>原始月销售额合计</small></div>';
  out += '<div class="card"><b>最新月 Top100 销量</b><strong>' + fmt(topLatest.sales) + '</strong><small>' + fmt(topLatest.count) + ' 个 Listing</small></div>';
  out += '<div class="card"><b>最新月平均标价</b><strong>$' + fmt(latest.avgPrice, 2) + '</strong><small>按有价格记录算术平均</small></div></div>';
  out += subsection(id + '-1-1', sectionNo + '.1 月度汇总与可回勾数据', trendTable(category.monthly, category.topMonthly));
  out += subsection(id + '-1-2', sectionNo + '.2 BSR Top100 月度明细', topTrendTable(category.topMonthly));
  out += subsection(id + '-1-3', sectionNo + '.3 年度 YOY 汇总', annualTable(category.annual, category.topAnnual));
  let charts = '<div class="charts">' + svgBarChart(category.title + '月销量', category.monthly, 'sales', '#2563eb', 0);
  charts += svgBarChart(category.title + '月销售额', category.monthly, 'revenue', '#0f766e', 2);
  charts += svgLineChart(category.title + '平均标价趋势', category.monthly, [{ name: '平均标价($)', field: 'avgPrice', color: '#7c3aed' }]);
  charts += svgLineChart(category.title + '销量MOM（去年同月）', category.monthly, [{ name: '销量MOM', field: 'momSales', color: '#2563eb' }], true) + '</div>';
  out += subsection(id + '-1-4', sectionNo + '.4 市场趋势图', charts);
  let rankCharts = '<p>Top100 先按可解析的小类BSR筛选 1—100（含100），再按父ASIN优先、否则ASIN去重；五档互斥为 1-5、6-10、11-20、21-50、51-100。MOM统一比较去年同月。</p><div class="charts">';
  rankCharts += svgLineChart(category.title + ' Top100 五档销量MOM', fineCombined, FINE_BANDS.map((band, i) => ({ name: band.name, field: 'mom_' + band.key, color: ['#1d4ed8', '#0891b2', '#059669', '#f59e0b', '#dc2626'][i] })), true);
  rankCharts += svgLineChart(category.title + ' Top100 头/中/尾销量MOM', tierCombined, BANDS.map((band, i) => ({ name: band.name, field: 'mom_' + band.key, color: ['#2563eb', '#059669', '#f97316'][i] })), true) + '</div>';
  out += subsection(id + '-1-5', sectionNo + '.5 BSR Top100 五档与头中尾 MOM', rankCharts);
  let tierBody = '';
  for (const band of BANDS) tierBody += tierTable(category.tiers[band.key], band.name);
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
      fmtPct(percent(r.sales, base.sales)),
      fmtPct(percent(r.revenue, base.revenue)),
      fmtPct(r.momSales),
      fmtPct(r.momRevenue),
      fmt(top.count)
    ];
  });
  let out = '<section id="genimo"><h2>第四部分｜GENIMO 品牌分析</h2>';
  out += '<p class="lead">GENIMO 是整体市场中的品牌视角。本部分同时给出 GENIMO 在整体市场的份额、BSR Top100表现、头中尾分层与基于原始数据的 2027 年行动建议。</p>';
  out += subsection('genimo-4-1', '4.1 月度表现与整体市场份额', table(['月份', 'GENIMO Listing数', '销量', '销售额($)', '平均标价($)', '销量占整体', '销售额占整体', '销量MOM', '销售额MOM', 'Top100 Listing数'], rows));
  const ppRows = genimoPP.monthly.map((r, i) => {
    const base = pp.monthly[i] || {};
    return [
      r.month.slice(0, 4) + '.' + Number(r.month.slice(4, 6)),
      fmt(r.count),
      fmt(r.sales),
      fmt(r.revenue, 2),
      fmt(r.avgPrice, 2),
      fmtPct(percent(r.sales, base.sales)),
      fmtPct(percent(r.revenue, base.revenue)),
      fmtPct(r.momSales),
      fmtPct(r.momRevenue)
    ];
  });
  let ppBody = '<p>PP市场份额分母为PP整体盘；GENIMO在PP中的统计仍沿用父ASIN优先/ASIN去重和Top100先筛选规则。</p>';
  ppBody += table(['月份', 'GENIMO PP Listing数', 'PP内销量', 'PP内销售额($)', '平均标价($)', '销量占PP', '销售额占PP', '销量MOM', '销售额MOM'], ppRows);
  ppBody += '<div class="charts">' + svgBarChart('GENIMO PP 月销量', genimoPP.monthly, 'sales', '#be185d', 0) + svgBarChart('GENIMO PP 月销售额', genimoPP.monthly, 'revenue', '#9d174d', 2) + '</div>';
  out += subsection('genimo-4-2', '4.2 GENIMO 在 PP 市场中的表现', ppBody);
  let genimoCharts = '<div class="charts">' + svgBarChart('GENIMO 月销量', category.monthly, 'sales', '#9333ea', 0);
  genimoCharts += svgBarChart('GENIMO 月销售额', category.monthly, 'revenue', '#c026d3', 2);
  genimoCharts += svgLineChart('GENIMO 销量MOM（去年同月）', category.monthly, [{ name: '销量MOM', field: 'momSales', color: '#2563eb' }], true) + '</div>';
  out += subsection('genimo-4-3', '4.3 GENIMO 趋势图', genimoCharts);
  let genimoTierBody = '';
  for (const band of BANDS) genimoTierBody += tierTable(category.tiers[band.key], 'GENIMO ' + band.name);
  out += subsection('genimo-4-4', '4.4 GENIMO BSR 头中尾与五档', genimoTierBody);
  out += subsection('genimo-4-5', '4.5 GENIMO 年度 YOY', annualTable(category.annual, category.topAnnual));
  const latest = lastNonEmpty(category.monthly);
  const latestTop = category.topMonthly.find((r) => r.month === latest.month) || {};
  const advice = [
    '先稳定能够进入 BSR 1—100 的链接：最新月 GENIMO Top100 有 ' + fmt(latestTop.count) + ' 个 Listing，销量 ' + fmt(latestTop.sales) + '。产品测试和链接补充应按 1-20、21-50、51-100 三个层级分别记录，不用单一总排名替代层级判断。',
    '用整体市场份额和 PP/高客单价市场拆分制定资源优先级：当 GENIMO 在某一市场的销量或销售额占比连续上升时，优先补充该市场相同字段完整且 BSR 可追踪的链接；当份额下降时，先核对缺失销量、销售额、价格和 BSR 的覆盖率，再决定是否调整产品组合。',
    '建立月度复盘表：销量、销售额、平均标价、销量MOM、销售额MOM、Top100数量、头中尾占比和标题中 plastic 标记必须同批次留痕。报告只使用原始工作簿已有字段，不推导利润、成本、广告或库存指标。'
  ].map((item) => '<li>' + esc(item) + '</li>').join('');
  out += subsection('genimo-4-6', '4.6 2027 年建议（仅基于当前原始字段）', '<div class="analysis"><ul class="analysis-list">' + advice + '</ul></div>');
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
    ['overall-1-2', '1.2 BSR Top100 月度明细'],
    ['overall-1-3', '1.3 年度 YOY 汇总'],
    ['overall-1-4', '1.4 市场趋势图'],
    ['overall-1-5', '1.5 BSR Top100 五档与头中尾 MOM'],
    ['overall-1-6', '1.6 头部 / 中部 / 尾部明细'],
    ['overall-1-7', '1.7 数据驱动文字分析']
  ]) + navGroup('pp', 'PP', 'PP市场', [
    ['pp-1-1', '2.1 月度汇总与可回勾数据'],
    ['pp-1-2', '2.2 BSR Top100 月度明细'],
    ['pp-1-3', '2.3 年度 YOY 汇总'],
    ['pp-1-4', '2.4 市场趋势图'],
    ['pp-1-5', '2.5 BSR Top100 五档与头中尾 MOM'],
    ['pp-1-6', '2.6 头部 / 中部 / 尾部明细'],
    ['pp-1-7', '2.7 数据驱动文字分析']
  ]) + navGroup('nonpp', '高', '高客单价市场', [
    ['nonpp-1-1', '3.1 月度汇总与可回勾数据'],
    ['nonpp-1-2', '3.2 BSR Top100 月度明细'],
    ['nonpp-1-3', '3.3 年度 YOY 汇总'],
    ['nonpp-1-4', '3.4 市场趋势图'],
    ['nonpp-1-5', '3.5 BSR Top100 五档与头中尾 MOM'],
    ['nonpp-1-6', '3.6 头部 / 中部 / 尾部明细'],
    ['nonpp-1-7', '3.7 数据驱动文字分析']
  ]) + navGroup('genimo', 'G', 'GENIMO品牌', [
    ['genimo-4-1', '4.1 月度表现与整体市场份额'],
    ['genimo-4-2', '4.2 GENIMO 在 PP 市场中的表现'],
    ['genimo-4-3', '4.3 GENIMO 趋势图'],
    ['genimo-4-4', '4.4 GENIMO BSR 头中尾与五档'],
    ['genimo-4-5', '4.5 GENIMO 年度 YOY'],
    ['genimo-4-6', '4.6 2027 年建议']
  ]);
  let html = '<!doctype html><html lang="zh-CN" data-theme="dark"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>户外地垫市场洞察 · Outdoor Rug Intelligence</title><style>' + UI_CSS + compatibilityCss + '</style></head><body><div class="app-shell">';
  html += '<aside class="sidebar"><div class="brand"><div class="brand-mark"></div><div><strong>Market Intelligence</strong><small>户外地垫市场分析系统</small></div></div><div class="local-badge"><i></i> DATA · VERIFIED</div><span class="nav-label">市场分析</span><nav>' + nav + '</nav><div class="sidebar-footer"><span>可比数据范围</span><strong>' + esc(MONTHS[0]) + ' — ' + esc(MONTHS[MONTHS.length - 1]) + '</strong><small>' + fmt(MONTHS.length) + '个月 · 原始主源直读</small></div></aside>';
  html += '<div class="workspace"><header class="topbar"><div><span class="eyebrow">MARKET OVERVIEW · SPEC 2.0</span><h1>户外地垫市场分析</h1><p>整体市场、PP市场、高客单价市场与 GENIMO 品牌 · 销量、销售额、均价 · BSR Top100 · MOM / YOY</p></div><div class="top-actions"><span class="privacy-chip"><i></i> 源数据只读</span><button class="theme-button" id="theme-toggle" aria-label="切换主题">☀</button></div></header><main class="content">';
  html += '<section id="dashboard" class="overview-section"><div class="section-kicker">报告导航 · 0</div><h2>市场总览｜范围与数据口径</h2><p class="lead">先确认数据范围、市场划分和统计单元，再进入四个分析部分。下面每个小节都可以点击左侧导航展开，也可以点击标题前的小三角收起。</p>';
  html += subsection('dashboard-0-1', '0.1 分析范围与关键指标', '<div class="scope-notice"><span>◎</span><div><b>分析范围：</b>原始工作簿覆盖 ' + MONTHS.length + ' 个月；整体市场由 PP 与高客单价市场构成，GENIMO 作为品牌视角单独分析。BSR Top100 使用 1—100（包含100）的独立池。</div></div><div class="metrics-grid"><article class="metric-card"><span class="metric-label">原始有效行</span><strong class="metric-value">' + fmt(rawCount) + '</strong><span class="metric-note">逐月明细读取</span></article><article class="metric-card"><span class="metric-label">整体盘去重 Listing</span><strong class="metric-value">' + fmt(categories.overall.fullRows.length) + '</strong><span class="metric-note">父ASIN优先 / ASIN兜底</span></article><article class="metric-card"><span class="metric-label">BSR Top100 去重</span><strong class="metric-value">' + fmt(categories.overall.topRows.length) + '</strong><span class="metric-note">小类BSR 1—100 含100</span></article><article class="metric-card"><span class="metric-label">GENIMO Listing</span><strong class="metric-value">' + fmt(categories.genimo.fullRows.length) + '</strong><span class="metric-note">品牌整体视角</span></article></div>');
  html += subsection('dashboard-0-2', '0.2 数据口径与来源', '<div class="scope-notice"><span>◎</span><div><b>口径与来源：</b>整体市场 = PP市场 ∪ 高客单价市场；PP 使用商品标题完整单词 plastic 匹配，高客单价市场为补集。月度 MOM 比较去年同月，年度 YOY 比较上一年度（未完结年度按实际覆盖范围标记）。数据源：' + esc(SOURCE) + '；SHA-256：' + esc(SOURCE_HASH) + '；可解析小类BSR ' + fmt(ranked) + ' 行，Top100候选 ' + fmt(topCandidates) + ' 行；月度子表有数据 ' + fmt(sourceMonthsWithData) + '/' + fmt(raw.sheetStats.length) + '，识别表头行 ' + esc(detectedHeaderRows.join('、')) + '。缺失值不当作零。</div></div><div class="meta"><b>统计单元与代表行</b><br>独立 Listing（父ASIN优先，否则ASIN；均无则保留源行）；代表行按最小可解析BSR、销量/销售额完整度、价格完整度、源行ID确定。原始字段覆盖表仅做只读回勾，不参与任何利润推导。</div>');
  html += subsection('dashboard-0-3', '0.3 原始字段覆盖（只读）', rawFieldCoverage(raw.rows));
  html += '</section>';
  html += marketSection('overall', '第一部分', categories.overall, categories.overall, categories.pp, categories.nonpp);
  html += marketSection('pp', '第二部分', categories.pp, categories.overall, categories.pp, categories.nonpp);
  html += marketSection('nonpp', '第三部分', categories.nonpp, categories.overall, categories.pp, categories.nonpp);
  html += genimoSection(categories.genimo, categories.overall, categories.pp, categories.genimoPP);
  html += '<footer>报告由 src/build_spec2_html.mjs 从原始工作簿生成。市场总览与数据口径已合并；左侧导航可展开到每个小节，正文小节可用小三角折叠。源表发生更换时，请先确认文件身份、月份覆盖和列名，再重新生成。</footer></main></div></div><script>(function(){var root=document.documentElement,button=document.getElementById("theme-toggle"),saved="dark";try{saved=localStorage.getItem("market-report-theme")||"dark"}catch(e){}root.dataset.theme=saved;button.textContent=saved==="dark"?"☀":"☾";button.addEventListener("click",function(){var next=root.dataset.theme==="dark"?"light":"dark";root.dataset.theme=next;button.textContent=next==="dark"?"☀":"☾";try{localStorage.setItem("market-report-theme",next)}catch(e){}});var links=[].slice.call(document.querySelectorAll(".sidebar nav a"));var groups=[].slice.call(document.querySelectorAll(".nav-group"));var targets=[].slice.call(document.querySelectorAll(".subsection[id],section[id]"));function openTarget(id){var el=document.getElementById(id);if(!el)return;var details=el.closest("details");if(details)details.open=true;var section=el.closest("section");if(section){var group=document.querySelector(\'.nav-group[data-target="\'+section.id+\'"]\');if(group)group.open=true}el.scrollIntoView({behavior:"smooth",block:"start"})}links.forEach(function(a){a.addEventListener("click",function(){openTarget(a.getAttribute("href").slice(1))})});function setActive(el){var id=el.id;links.forEach(function(a){a.classList.toggle("is-active",a.getAttribute("href")==="#"+id)});var section=el.closest("section[id]");var sectionId=section?section.id:(el.matches("section[id]")?el.id:null);groups.forEach(function(group){group.classList.toggle("is-active",!!sectionId&&group.dataset.target===sectionId)})}if("IntersectionObserver" in window){var observer=new IntersectionObserver(function(entries){entries.forEach(function(entry){if(entry.isIntersecting)setActive(entry.target)})},{rootMargin:"-18% 0px -68% 0px"});targets.forEach(function(el){observer.observe(el)})}})();</script></body></html>';
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
  const raw = readRawRows();
  const flags = buildFamilyFlags(raw.rows);
  const fullDedup = dedup(raw.rows, flags);
  const topCandidates = raw.rows.filter((r) => r.rank !== null && r.rank >= 1 && r.rank <= 100);
  // Keep a deterministic maximum of 100 listings per month and scope. A
  // source export can contain tied BSR values, so filtering rank <= 100 alone
  // can produce more than 100 listings. Rank is the primary order; the stable
  // family key and source row make ties reproducible for manual rechecks.
  const selectTop100 = (rows) => {
    const byMonth = new Map();
    for (const row of rows) {
      if (row.rank === null || row.rank < 1 || row.rank > 100) continue;
      if (!byMonth.has(row.month)) byMonth.set(row.month, []);
      byMonth.get(row.month).push(row);
    }
    return MONTHS.flatMap((month) => (byMonth.get(month) || [])
      .slice()
      .sort((a, b) => a.rank - b.rank || a.familyKey.localeCompare(b.familyKey) || a.sourceRow - b.sourceRow)
      .slice(0, 100));
  };
  const topDedup = selectTop100(fullDedup);
  const ppRows = fullDedup.filter((r) => r.plastic);
  const nonppRows = fullDedup.filter((r) => !r.plastic);
  const genimoRows = fullDedup.filter((r) => r.genimo);
  const genimoPPRows = fullDedup.filter((r) => r.genimo && r.plastic);
  const categories = {
    overall: buildCategory('整体市场', fullDedup, topDedup),
    pp: buildCategory('PP市场', ppRows, selectTop100(ppRows)),
    nonpp: buildCategory('高客单价市场', nonppRows, selectTop100(nonppRows)),
    genimo: buildCategory('GENIMO品牌', genimoRows, selectTop100(genimoRows)),
    genimoPP: buildCategory('GENIMO PP市场', genimoPPRows, selectTop100(genimoPPRows))
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
  return { raw, fullDedup, topCandidates, topDedup, categories };
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

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) runHtmlBuild();
