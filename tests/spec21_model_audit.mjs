import assert from 'node:assert/strict';
import fs from 'node:fs';
import {parseBsr,isPlasticTitle,representativeCompare,annualRows,display} from '../src/build_spec2_html.mjs';
assert.equal(display({v:'',r:'<t></t><r><t>Rugs.com</t></r><r><t> Aruba &#39; rug</t></r>',l:{display:'Rugs.com'}}),"Rugs.com Aruba ' rug");
for(const [v,result] of [['100',100],['100.0',100],['-1',null],['− 5',null],['1.5',null],['-1.5',null],['#-10; #20',20],['591.0',591],['#1,000; #99',99],['0',null]])assert.equal(parseBsr(v),result,v);
for(const [v,result] of [['plastic rug',true],['PLASTIC',true],['plastics',false],['plastic123',false],['plastic_rug',false],['plasticity',false],['(plastic)',true]])assert.equal(isPlasticTitle(v),result,v);
const a={rank:5,sales:10,revenue:20,price:2,sourceRow:9},b={...a,sourceRow:100};assert.ok(representativeCompare(a,b)<0);
assert.equal(annualRows(['202401','202601'],[{month:'202401',sales:10,revenue:20,price:2},{month:'202601',sales:20,revenue:40,price:2}])[1].yoySales,null,'No skipping a missing year');
const rs=annualRows(['202401','202403','202501','202502','202503'],[{month:'202401',sales:10,revenue:20,price:2},{month:'202403',sales:10,revenue:20,price:2},{month:'202501',sales:20,revenue:40,price:2},{month:'202502',sales:999,revenue:999,price:2},{month:'202503',sales:20,revenue:40,price:2}]);assert.equal(rs[1].yoySales,1);assert.deepEqual(rs[1].commonMonths,['01','03']);
const d=JSON.parse(fs.readFileSync('tmp/formula_market_builder/spec2_dataset.json','utf8'));
assert.equal(d.raw.length,73810);assert.equal(d.months.length,50);assert.equal(new Set(d.months).size,50);
assert.equal(d.mainPool.length,5190);
assert.equal(d.dedup.length,d.mainPool.length);
assert.ok(d.dedup.every(r=>Number.isInteger(r.rank)&&r.rank>=1&&r.rank<=100&&r.parent),'business pool must be BSR-filtered parent ASIN rows');
for(const month of d.months){
  const sets=Object.fromEntries(['overall','pp','nonpp','genimo','genimoPP'].map(scope=>[scope,new Set(d.categories[scope].topRows.filter(r=>r.month===month).map(r=>r.familyKey))]));
  assert.equal(sets.overall.size,sets.pp.size+sets.nonpp.size,month+' PP/high complement');
  for(const key of sets.pp)assert.ok(sets.overall.has(key)&&!sets.nonpp.has(key),month+' PP subset');
  for(const key of sets.nonpp)assert.ok(sets.overall.has(key)&&!sets.pp.has(key),month+' high subset');
  for(const key of sets.genimoPP)assert.ok(sets.genimo.has(key)&&sets.pp.has(key),month+' GENIMO PP intersection');
}
assert.equal(d.categories.genimo.topMonthly.find(r=>r.month==='202505').count,7);assert.equal(d.categories.genimoIndependent.topMonthly.find(r=>r.month==='202505').count,7);
for(const mv of d.movements)if(mv.available){assert.equal(mv.current,mv.entered+mv.retained);assert.equal(mv.prior,mv.exited+mv.retained);}
for(const [scope,cat] of Object.entries(d.categories))for(const r of cat.monthly){assert.ok(r.count||r.sales===null);if(r.count===0)assert.equal(r.revenue,null);}
console.log('SPEC21_MODEL_BOUNDARIES_OK');
