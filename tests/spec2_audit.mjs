import assert from 'node:assert/strict';
import { buildDataset, parseBsr } from '../src/build_spec2_html.mjs';

// Parser boundary regressions from SPEC 2.0.
assert.equal(parseBsr('591.0'), 591);
assert.equal(parseBsr('1.5'), null);
assert.equal(parseBsr('Category A #45; Category B #120'), 45);

const dataset = buildDataset();
const month = '202509';
const raw = dataset.raw.rows.filter((row) => row.month === month);
const candidates = raw.filter((row) => Number.isInteger(row.rank) && row.rank >= 1 && row.rank <= 100);
const top = dataset.categories.overall.topMonthly.find((row) => row.month === month);
const candidateAudit = dataset.categories.overall.topCandidateMonthly.find((row) => row.month === month);
const ppCandidateAudit = dataset.categories.pp.topCandidateMonthly.find((row) => row.month === month);
const nonppCandidateAudit = dataset.categories.nonpp.topCandidateMonthly.find((row) => row.month === month);

assert.equal(raw.length, 1976);
assert.equal(candidates.length, 114);
assert.equal(candidateAudit.rawCount, 114);
assert.equal(candidateAudit.count, 114);
assert.equal(top.count, 114);
assert.equal(ppCandidateAudit.rawCount + nonppCandidateAudit.rawCount, 114);
assert.equal(ppCandidateAudit.count + nonppCandidateAudit.count, 114);
assert.equal(raw.filter((row) => row.rank >= 1 && row.rank <= 100 && row.plastic).length, 65);

console.log(JSON.stringify({
  month,
  rawRows: raw.length,
  bsrCandidateRows: candidates.length,
  bsrCandidateListingRows: top.count,
  ppCandidateRowsByListingScope: ppCandidateAudit.rawCount,
  nonppCandidateRowsByListingScope: nonppCandidateAudit.rawCount,
  titlePlasticCandidateRows: raw.filter((row) => row.rank >= 1 && row.rank <= 100 && row.plastic).length
}, null, 2));
