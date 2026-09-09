import fs from 'node:fs/promises';
import path from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import * as XLSX from 'xlsx';

const ROOT=path.resolve('.');
const DB=path.join(ROOT,'data/processed/market.db');
const CDB=path.join(ROOT,'data/processed/competitor_809440.db');
const JSON_FILE=path.join(ROOT,'交付/户外地垫市场分析数据.json');
const OUT_DIR=path.join(ROOT,'outputs/20260909-formula-market-analysis');
const OUT=path.join(OUT_DIR,'户外地垫市场分析-公式版-20260909.xlsx');
const PREVIEW=path.join(OUT_DIR,'00_总览-预览.png');
const q=x=>'"'+String(x).replaceAll('"','""')+'"',txt=x=>x==null?'':String(x),n=x=>x!==null&&x!==undefined&&x!==''&&Number.isFinite(Number(x))?Number(x):null;
const plastic=/\bplastic\b/i,CORE_A='202601',CORE_B='202606',SHOW_A='202501',SHOW_B='202607';
const scopes=[['overall','整体市场（PP+非PP）',r=>true],['pp','PP管/Plastic',r=>r.pp],['nonpp','非PP高客单产品线',r=>!r.pp],['genimo','GENIMO品牌',r=>r.genimo]];
const coarse=[['头部（1-20）',1,20],['中部（21-50）',21,50],['尾部（51-100）',51,100]];
const fine=[['1-5',1,5],['6-10',6,10],['11-20',11,20],['21-50',21,50],['51-100',51,100]];
const parseBsr=v=>{if(v==null||v==='')return null;const a=(String(v).match(/(?<![\d.,])\d[\d,]*(?:\.0+)?(?![\d.])/g)||[]).map(x=>Number(x.replaceAll(',',''))).filter(x=>Number.isInteger(x)&&x>0);return a.length?Math.min(...a):null};
const keyOf=r=>txt(r.parent).trim()||txt(r.asin).trim();
const better=(a,b)=>{if(!b)return true;const ar=a.rank??Infinity,br=b.rank??Infinity;return ar<br||(ar===br&&(a.complete>b.complete||(a.complete===b.complete&&a.rowId<b.rowId)))};
const prevMonth=m=>{const d=new Date(Date.UTC(Number(m.slice(0,4)),Number(m.slice(4))-2,1));return String(d.getUTCFullYear())+String(d.getUTCMonth()+1).padStart(2,'0')};
const prevYear=m=>String(Number(m.slice(0,4))-1)+m.slice(4);
const status=m=>m==='202607'?'展示样本（94父体）':(m>=CORE_A&&m<=CORE_B?'核心':'历史基线');
const source=m=>m>='202601'&&m<='202607'?'2026竞品替换链':'主源工作簿';
const total=(a,f)=>a.reduce((s,r)=>s+(f(r)??0),0);

const db=new DatabaseSync(DB,{readOnly:true}),cdb=new DatabaseSync(CDB,{readOnly:true});
const sourceMonths=db.prepare("SELECT target_table FROM sheet_catalog WHERE classification='monthly' ORDER BY sheet_order").all().map(x=>String(x.target_table).replace('monthly_','')).filter(x=>/^\d{6}$/.test(x));
const months=sourceMonths.filter(x=>x>='202301'&&x<=SHOW_B),display=months.filter(x=>x>=SHOW_A);
const profiles=new Map();
for(const m of sourceMonths.filter(x=>x>='202601'&&x<='202607')){
 const t='raw_'+m,pm=new Map();if(!cdb.prepare("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?").get(t))throw new Error('missing '+t);
 for(const r of cdb.prepare('SELECT row_id,ASIN asin,"父ASIN" parent,品牌 brand,"商品标题" title,"小类BSR" bsr FROM '+q(t)).all()){
  const k=keyOf(r);if(!k)continue;const p=pm.get(k)||{pp:false,genimo:false,r:{overall:null,pp:null,nonpp:null,genimo:null}},rank=parseBsr(r.bsr),pp=plastic.test(txt(r.title)),g=txt(r.brand).trim().toLowerCase()==='genimo';p.pp||=pp;p.genimo||=g;
  for(const pair of [['overall',true],['pp',pp],['nonpp',!pp],['genimo',g]])if(pair[1]&&rank!==null&&(p.r[pair[0]]===null||rank<p.r[pair[0]]))p.r[pair[0]]=rank;
  pm.set(k,p);
 }
 profiles.set(m,pm);
}
const details=[];
for(const m of months){
 const rows=db.prepare('SELECT row_id,ASIN asin,SKU sku,"父ASIN" parent,品牌 brand,"商品标题" title,"小类BSR" bsr,"月销量" sales,"月销售额" revenue,价格 price FROM '+q('monthly_'+m)).all(),groups=new Map();
 for(const r of rows){const k=keyOf(r);if(!k)continue;const rank=parseBsr(r.bsr),complete=Number(n(r.sales)!==null)+Number(n(r.revenue)!==null),g=groups.get(k)||{key:k,raw:0,pp:false,genimo:false,rank:null,rep:null};g.raw++;g.pp||=plastic.test(txt(r.title));g.genimo||=txt(r.brand).trim().toLowerCase()==='genimo';const c={rowId:Number(r.row_id),row:r,rank,complete};if(better(c,g.rep)){g.rep=c;g.rank=rank}groups.set(k,g)}
 const pm=profiles.get(m);for(const g of groups.values()){const r=g.rep.row,p=pm?.get(g.key),pp=p?.pp??g.pp,genimo=p?.genimo??g.genimo;details.push({month:m,key:g.key,asin:txt(r.asin).trim(),parent:txt(r.parent).trim(),sku:txt(r.sku).trim(),brand:txt(r.brand).trim(),title:txt(r.title).trim(),pp,genimo,sales:n(r.sales),revenue:n(r.revenue),price:n(r.price),rankOverall:p?.r?.overall??g.rank,rankPP:p?.r?.pp??(pp?g.rank:null),rankNonPP:p?.r?.nonpp??(!pp?g.rank:null),rankGenimo:p?.r?.genimo??(genimo?g.rank:null),raw:g.raw,dup:g.raw-1,sourceRow:Number(r.row_id),source:source(m),status:status(m),sv:n(r.sales)!==null?1:0,rv:n(r.revenue)!==null?1:0,pv:n(r.price)!==null?1:0,paired:n(r.sales)!==null&&n(r.revenue)!==null?1:0,core:m>=CORE_A&&m<=CORE_B?1:0,july:m==='202607'?1:0,topOverall:0,topPP:0,topNonPP:0,topGenimo:0})}}
const byMonth=new Map(months.map(m=>[m,details.filter(r=>r.month===m)])),rankFor=(r,s)=>s==='overall'?r.rankOverall:s==='pp'?r.rankPP:s==='nonpp'?r.rankNonPP:r.rankGenimo,topBy=new Map();
for(const m of months)for(const [s,,fn] of scopes){const top=(byMonth.get(m)||[]).filter(r=>fn(r)&&Number.isInteger(rankFor(r,s))&&rankFor(r,s)>=1&&rankFor(r,s)<=100).sort((a,b)=>rankFor(a,s)-rankFor(b,s)||a.key.localeCompare(b.key)||a.sourceRow-b.sourceRow).slice(0,100);topBy.set(m+'|'+s,top);const f=s==='overall'?'topOverall':s==='pp'?'topPP':s==='nonpp'?'topNonPP':'topGenimo';top.forEach(r=>r[f]=1)}
const stats=rs=>{const sv=rs.filter(r=>r.sv),rv=rs.filter(r=>r.rv),pv=rs.filter(r=>r.pv),pr=rs.filter(r=>r.paired);return{raw:total(rs,r=>r.raw),list:rs.length,dup:total(rs,r=>r.dup),sales:total(sv,r=>r.sales),revenue:total(rv,r=>r.revenue),priceSum:total(pv,r=>r.price),priceCount:pv.length,pairedSales:total(pr,r=>r.sales),pairedRevenue:total(pr,r=>r.revenue),salesValid:sv.length,revenueValid:rv.length,pairedCount:pr.length}};
const aggs=[],add=(m,s,t,rs)=>aggs.push({month:m,scope:s,tier:t,...stats(rs),source:source(m),status:status(m)});
for(const m of months){const rs=byMonth.get(m)||[];for(const [s,,fn] of scopes){const base=rs.filter(fn),top=topBy.get(m+'|'+s)||[];add(m,s,'全部',base);add(m,s,'Top100',top);for(const [t,a,b] of [...coarse,...fine])add(m,s,t,top.filter(r=>{const z=rankFor(r,s);return z>=a&&z<=b}))}add(m,'genimo_pp','全部',rs.filter(r=>r.genimo&&r.pp))}

// The reference profit sheet is a weekly SKU test model.  The market snapshots
// provide the latest price/BSR/margin/FBA as reference inputs, while weekly
// sales and landed-cost fields intentionally remain editable blanks.
const costMap=new Map();
for(const r of cdb.prepare('SELECT row_id,ASIN asin,SKU sku,"父ASIN" parent,品牌 brand,"商品标题" title,"小类BSR" bsr,"月销量" sales,"月销售额" revenue,价格 price,FBA fba,"毛利率" margin FROM '+q('raw_202606')).all()){
 const k=keyOf(r);if(!k)continue;const rank=parseBsr(r.bsr),complete=Number(n(r.price)!==null)+Number(n(r.fba)!==null)+Number(n(r.margin)!==null),old=costMap.get(k);
 if(!old||rank!==null&&(old.rank===null||rank<old.rank)||(old.rank===rank&&complete>old.complete))costMap.set(k,{...r,rank,complete});
}
const profitRows=(topBy.get('202606|genimo')||[]).map(r=>{const c=costMap.get(r.key)||{};return{asin:r.asin||c.asin||'',sku:r.sku||c.sku||'',normalPrice:n(c.price??r.price),discountType:'',discountValue:null,sevenDaySales:null,purchaseRmb:null,firstMile:null,lastMile:null,volumeFt3:null,sourceMonth:'202606',marketSales:r.sales,marketRevenue:r.revenue,bsr:r.rankGenimo??r.rankOverall??c.rank,margin:n(c.margin),fba:n(c.fba),brand:r.brand,title:r.title,key:r.key}});



const benchmarkData=JSON.parse(await fs.readFile(JSON_FILE,'utf8'));
const benchmark=benchmarkData.leadershipBenchmark||{};
const detailHeaders=['月份','Listing键','代表ASIN','父ASIN','SKU','品牌','代表标题','PP标记','非PP标记','GENIMO标记','月销量','月销售额','价格','整体BSR','PP BSR','非PP BSR','GENIMO BSR','原始行数','去重行数','销量有效','销售额有效','价格有效','配对有效','源行ID','来源','状态','核心标记','2026.07展示标记','整体Top100','PP Top100','非PP Top100','GENIMO Top100'];
const detailValues=details.map(r=>[r.month,r.key,r.asin||null,r.parent||null,r.sku||null,r.brand||null,r.title||null,r.pp?1:0,r.pp?0:1,r.genimo?1:0,r.sales,r.revenue,r.price,r.rankOverall,r.rankPP,r.rankNonPP,r.rankGenimo,r.raw,r.dup,r.sv,r.rv,r.pv,r.paired,r.sourceRow,r.source,r.status,r.core,r.july,r.topOverall,r.topPP,r.topNonPP,r.topGenimo]);
const aggValues=aggs.map(r=>[r.month,r.scope,scopes.find(x=>x[0]===r.scope)?.[1]||r.scope,r.tier,r.raw,r.list,r.dup,r.sales,r.revenue,r.priceSum,r.priceCount,r.pairedSales,r.pairedRevenue,r.salesValid,r.revenueValid,r.pairedCount,r.source,r.status]);
const detailEnd=detailValues.length+2;
const aggEnd=aggValues.length+2;
const wb=XLSX.utils.book_new();
const qsheet=name=>"'"+name+"'";
const cell=(f,v)=>({f:f,v:v===undefined||v===null?'':v});
const numOrNull=v=>v===null||v===undefined||Number.isNaN(v)?null:Number(v);
const aggMap=new Map(aggs.map(r=>[r.month+'|'+r.scope+'|'+r.tier,r]));
const metricCol={raw:'E',list:'F',dup:'G',sales:'H',revenue:'I',priceSum:'J',priceCount:'K',pairedSales:'L',pairedRevenue:'M',salesValid:'N',revenueValid:'O',pairedCount:'P'};
const getAgg=(m,s,t)=>aggMap.get(m+'|'+s+'|'+t)||{raw:0,list:0,dup:0,sales:0,revenue:0,priceSum:0,priceCount:0,pairedSales:0,pairedRevenue:0,salesValid:0,revenueValid:0,pairedCount:0};
const sumCache=(metric,s,t,a,b)=>months.filter(m=>m>=a&&m<=b).reduce((z,m)=>z+(Number(getAgg(m,s,t)[metric])||0),0);
const avgCache=(s,t,m)=>{const r=getAgg(m,s,t);return r.priceCount? r.priceSum/r.priceCount:null};
const weightedCache=(s,t,m)=>{const r=getAgg(m,s,t);return r.pairedSales? r.pairedRevenue/r.pairedSales:null};
const periodAvg=(s,t,a,b)=>{const ps=sumCache('priceSum',s,t,a,b),pc=sumCache('priceCount',s,t,a,b);return pc?ps/pc:null};
const periodWeighted=(s,t,a,b)=>{const ps=sumCache('pairedRevenue',s,t,a,b),pc=sumCache('pairedSales',s,t,a,b);return pc?ps/pc:null};
const rate=(a,b)=>b? a/b-1:null;
const prevM=prevMonth,prevY=prevYear;
const aggRangeEnd=aggEnd;
const aggRange={A:"'91_聚合输入'!$A$3:$A$"+aggRangeEnd,B:"'91_聚合输入'!$B$3:$B$"+aggRangeEnd,D:"'91_聚合输入'!$D$3:$D$"+aggRangeEnd};
function sf(metric,s,t,m){const c=metricCol[metric];return "SUMIFS('91_聚合输入'!$"+c+"$3:$"+c+"$"+aggRangeEnd+","+aggRange.A+",\""+m+"\","+aggRange.B+",\""+s+"\","+aggRange.D+",\""+t+"\")";}
function sfr(metric,s,t,a,b){const c=metricCol[metric];return "SUMIFS('91_聚合输入'!$"+c+"$3:$"+c+"$"+aggRangeEnd+","+aggRange.A+",\">="+a+"\","+aggRange.A+",\"<="+b+"\","+aggRange.B+",\""+s+"\","+aggRange.D+",\""+t+"\")";}
function addSheet(name,rows,widthsList,merges=[]){const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=widthsList.map(w=>({wch:w}));if(merges.length)ws['!merges']=merges.map(x=>({s:{r:x[0],c:x[1]},e:{r:x[2],c:x[3]}}));XLSX.utils.book_append_sheet(wb,ws,name);return ws;}
function setAutofilter(ws,endCol,endRow){ws['!autofilter']={ref:'A4:'+endCol+endRow};}
const allDetails=details;
const detailRows90=[['原始明细输入（值）· 一行代表一个父ASIN/ASIN确定性代表记录'],detailHeaders,...detailValues];
const ws90=addSheet('90_输入_明细',detailRows90,[11,20,15,15,16,15,60,9,10,11,13,15,12,11,10,11,13,12,12,11,12,11,11,10,20,18,12,18,12,11,12,14],[[0,0,0,31]]);
setAutofilter(ws90,'AF',detailRows90.length);
const aggRows91=[['聚合输入（脚本按明细确定性汇总，结果页全部用Excel公式引用本页）'],['月份','范围键','范围','层级','原始行数','独立Listing','去重行数','销量','销售额','价格合计','价格有效数','配对销量','配对销售额','销量有效数','销售额有效数','配对有效数','来源','状态'],...aggValues];
const ws91=addSheet('91_聚合输入',aggRows91,[11,13,25,16,12,14,12,14,16,14,12,14,16,14,15,13,20,18],[[0,0,0,17]]);


const periods=[['2024全年','202401','202412'],['2025全年','202501','202512'],['2025H1','202501','202506'],['2026H1','202601','202606']];
function avgFormula(s,t,m){return 'IFERROR('+sf('priceSum',s,t,m)+'/'+sf('priceCount',s,t,m)+',"")';}
function weightedFormula(s,t,m){return 'IFERROR('+sf('pairedRevenue',s,t,m)+'/'+sf('pairedSales',s,t,m)+',"")';}
function periodAvgFormula(s,t,a,b){return 'IFERROR('+sfr('priceSum',s,t,a,b)+'/'+sfr('priceCount',s,t,a,b)+',"")';}
function periodWeightedFormula(s,t,a,b){return 'IFERROR('+sfr('pairedRevenue',s,t,a,b)+'/'+sfr('pairedSales',s,t,a,b)+',"")';}
function rateFormula(cur,base){return 'IFERROR(('+cur+')/('+base+')-1,"")';}
function monthlyRows(scopeKey){
 const rows=[[''],[''],[],['Month 月份','数据可比性/范围状态','Raw rows 原始行','独立Listing数','去除重复行','Sales 销量','销量有效','销量缺失','Revenue 销售额($)','销售额有效','销售额缺失','ASP 客单价($)','加权成交均价($)','配对覆盖率','MoM 销量环比','MoM 销售额环比','MoM ASP环比','YoY 销量同比','YoY 销售额同比','YoY ASP同比','重复率','Notes 备注']];
 rows[0]=[scopes.find(s=>s[0]===scopeKey)[1]+' · 月度与年度分析（公式页）'];
 rows[1]=['MOM/环比=当前月 vs 上一自然月；YOY/同比=当前月 vs 去年同月。2026.07仅展示，两个增速留空。'];
 for(const m of display){
  const pm=prevM(m),py=prevY(m),show=m===SHOW_B;
  const sales=sf('sales',scopeKey,'全部',m),rev=sf('revenue',scopeKey,'全部',m),price=avgFormula(scopeKey,'全部',m);
  const sv=sf('salesValid',scopeKey,'全部',m),rv=sf('revenueValid',scopeKey,'全部',m);
  const row=[m,m===SHOW_B?'展示':(m>=CORE_A&&m<=CORE_B?'核心':'历史基线'),
   cell(sf('raw',scopeKey,'全部',m),getAgg(m,scopeKey,'全部').raw),
   cell(sf('list',scopeKey,'全部',m),getAgg(m,scopeKey,'全部').list),
   cell(sf('dup',scopeKey,'全部',m),getAgg(m,scopeKey,'全部').dup),
   cell(sales,getAgg(m,scopeKey,'全部').sales),
   cell(sv,getAgg(m,scopeKey,'全部').salesValid),
   cell('D'+(rows.length+1)+'-G'+(rows.length+1),getAgg(m,scopeKey,'全部').raw-getAgg(m,scopeKey,'全部').salesValid),
   cell(rev,getAgg(m,scopeKey,'全部').revenue),
   cell(rv,getAgg(m,scopeKey,'全部').revenueValid),
   cell('D'+(rows.length+1)+'-J'+(rows.length+1),getAgg(m,scopeKey,'全部').raw-getAgg(m,scopeKey,'全部').revenueValid),
   cell(price,avgCache(scopeKey,'全部',m)),
   cell(weightedFormula(scopeKey,'全部',m),weightedCache(scopeKey,'全部',m)),
   cell('IFERROR('+sf('pairedCount',scopeKey,'全部',m)+'/'+sf('list',scopeKey,'全部',m)+',"")',getAgg(m,scopeKey,'全部').list?getAgg(m,scopeKey,'全部').pairedCount/getAgg(m,scopeKey,'全部').list:null),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sales,sf('sales',scopeKey,'全部',pm)),rate(getAgg(m,scopeKey,'全部').sales,getAgg(pm,scopeKey,'全部').sales)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(rev,sf('revenue',scopeKey,'全部',pm)),rate(getAgg(m,scopeKey,'全部').revenue,getAgg(pm,scopeKey,'全部').revenue)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(price,avgFormula(scopeKey,'全部',pm)),rate(avgCache(scopeKey,'全部',m),avgCache(scopeKey,'全部',pm))),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sales,sf('sales',scopeKey,'全部',py)),rate(getAgg(m,scopeKey,'全部').sales,getAgg(py,scopeKey,'全部').sales)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(rev,sf('revenue',scopeKey,'全部',py)),rate(getAgg(m,scopeKey,'全部').revenue,getAgg(py,scopeKey,'全部').revenue)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(price,avgFormula(scopeKey,'全部',py)),rate(avgCache(scopeKey,'全部',m),avgCache(scopeKey,'全部',py))),
   cell('IFERROR(E'+(rows.length+1)+'/C'+(rows.length+1)+',"")',getAgg(m,scopeKey,'全部').raw?getAgg(m,scopeKey,'全部').dup/getAgg(m,scopeKey,'全部').raw:null),
   show?'展示样本：不纳入核心增速':'公式引用91_聚合输入；缺失值不按0进入有效分母'];
  rows.push(row);
 }
 return rows;
}
function writeMonthlySheet(scopeKey,name){
 const rows=monthlyRows(scopeKey);rows.push([]);
 const annualHeaderRow=rows.length+1;rows.push(['年度/周期','范围','原始行数','独立Listing','去重行数','销量','销售额','平均标价','加权成交均价','销量有效','销售额有效']);
 const annualStart=rows.length+1;
 for(const p of periods){const label=p[0],a=p[1],b=p[2];rows.push([label,a+'-'+b,
  cell(sfr('raw',scopeKey,'全部',a,b),sumCache('raw',scopeKey,'全部',a,b)),
  cell(sfr('list',scopeKey,'全部',a,b),sumCache('list',scopeKey,'全部',a,b)),
  cell(sfr('dup',scopeKey,'全部',a,b),sumCache('dup',scopeKey,'全部',a,b)),
  cell(sfr('sales',scopeKey,'全部',a,b),sumCache('sales',scopeKey,'全部',a,b)),
  cell(sfr('revenue',scopeKey,'全部',a,b),sumCache('revenue',scopeKey,'全部',a,b)),
  cell(periodAvgFormula(scopeKey,'全部',a,b),periodAvg(scopeKey,'全部',a,b)),
  cell(periodWeightedFormula(scopeKey,'全部',a,b),periodWeighted(scopeKey,'全部',a,b)),
  cell(sfr('salesValid',scopeKey,'全部',a,b),sumCache('salesValid',scopeKey,'全部',a,b)),
  cell(sfr('revenueValid',scopeKey,'全部',a,b),sumCache('revenueValid',scopeKey,'全部',a,b))]);}
 rows.push(['说明','2025H1与2026H1用于核心对比；2024全年和2025全年保留年度基线。']);
 rows.push([]);rows.push(['月度同比对比（参考表头）']);
 rows.push(['Month 月份','2025 Sales','2026 Sales','Sales YoY','2025 Revenue','2026 Revenue','Revenue YoY','2025 ASP','2026 ASP','ASP YoY','Interpretation 判断']);
 const compareStart=rows.length+1;
 for(const m25 of months.filter(x=>x>='202501'&&x<='202512')){
   const mm=m25.slice(4),m26='2026'+mm,has26=months.includes(m26);
   const s25=sf('sales',scopeKey,'全部',m25),s26=has26?sf('sales',scopeKey,'全部',m26):'IFERROR(0/0,"")';
   const r25=sf('revenue',scopeKey,'全部',m25),r26=has26?sf('revenue',scopeKey,'全部',m26):'IFERROR(0/0,"")';
   const p25=avgFormula(scopeKey,'全部',m25),p26=has26?avgFormula(scopeKey,'全部',m26):'IFERROR(0/0,"")';
   const rr=rows.length+1;
   rows.push([m25,cell(s25,getAgg(m25,scopeKey,'全部').sales),has26?cell(s26,getAgg(m26,scopeKey,'全部').sales):null,
     has26?cell(rateFormula(s26,s25),rate(getAgg(m26,scopeKey,'全部').sales,getAgg(m25,scopeKey,'全部').sales)):null,
     cell(r25,getAgg(m25,scopeKey,'全部').revenue),has26?cell(r26,getAgg(m26,scopeKey,'全部').revenue):null,
     has26?cell(rateFormula(r26,r25),rate(getAgg(m26,scopeKey,'全部').revenue,getAgg(m25,scopeKey,'全部').revenue)):null,
     cell(p25,avgCache(scopeKey,'全部',m25)),has26?cell(p26,avgCache(scopeKey,'全部',m26)):null,
     has26?cell(rateFormula(p26,p25),rate(avgCache(scopeKey,'全部',m26),avgCache(scopeKey,'全部',m25))):null,
     has26?cell('IF(AND(D'+rr+'>0,G'+rr+'<0),"量增额降：复核价格与结构",IF(AND(D'+rr+'>0,G'+rr+'>0),"量额同增","数据不足或需复核"))',''):null]);
 }
 const ws=addSheet(name,rows,[14,22,14,15,14,14,12,12,15,14,14,14,16,13,13,14,15,13,13,14,12,12,45],[[0,0,0,22],[rows.length-1,0,rows.length-1,10]]);
 return {sheet:name,annualStart,annualEnd:annualStart+periods.length-1,monthlyStart:5,monthlyEnd:4+display.length,compareStart,compareEnd:rows.length};
}
function topRows(scopeKey){
 const rows=[[''],[''],[],['Month 月份','数据可比性/范围状态','Listings','Raw rows 原始行','去除重复行','Sales 销量','销量有效','Revenue 销售额($)','销售额有效','ASP 客单价($)','加权成交均价($)','配对覆盖率','MoM 销量环比','MoM 销售额环比','MoM ASP环比','Sales YoY 销量同比','Revenue YoY 销售额同比','ASP YoY 客单价同比','Head Sales 头部销量','Middle Sales 中部销量','Tail Sales 尾部销量','Notes 备注']];
 rows[0]=[scopes.find(s=>s[0]===scopeKey)[1]+' · BSR前100分析（公式页）'];
 rows[1]=['每月按对应范围独立取BSR 1-100，最多100个独立Listing；Top100统计与月度分析使用相同去重后的代表记录。'];
 for(const m of display){const pm=prevM(m),py=prevY(m),show=m===SHOW_B;const sales=sf('sales',scopeKey,'Top100',m),rev=sf('revenue',scopeKey,'Top100',m),price=avgFormula(scopeKey,'Top100',m),r=getAgg(m,scopeKey,'Top100');
  rows.push([m,m===SHOW_B?'展示':(m>=CORE_A&&m<=CORE_B?'核心':'历史基线'),
   cell(sf('list',scopeKey,'Top100',m),r.list),cell(sf('raw',scopeKey,'Top100',m),r.raw),cell(sf('dup',scopeKey,'Top100',m),r.dup),
   cell(sales,r.sales),cell(sf('salesValid',scopeKey,'Top100',m),r.salesValid),cell(rev,r.revenue),cell(sf('revenueValid',scopeKey,'Top100',m),r.revenueValid),
   cell(price,avgCache(scopeKey,'Top100',m)),cell(weightedFormula(scopeKey,'Top100',m),weightedCache(scopeKey,'Top100',m)),
   cell('IFERROR('+sf('pairedCount',scopeKey,'Top100',m)+'/'+sf('list',scopeKey,'Top100',m)+',"")',r.list?r.pairedCount/r.list:null),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sales,sf('sales',scopeKey,'Top100',pm)),rate(r.sales,getAgg(pm,scopeKey,'Top100').sales)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(rev,sf('revenue',scopeKey,'Top100',pm)),rate(r.revenue,getAgg(pm,scopeKey,'Top100').revenue)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(price,avgFormula(scopeKey,'Top100',pm)),rate(avgCache(scopeKey,'Top100',m),avgCache(scopeKey,'Top100',pm))),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sales,sf('sales',scopeKey,'Top100',py)),rate(r.sales,getAgg(py,scopeKey,'Top100').sales)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(rev,sf('revenue',scopeKey,'Top100',py)),rate(r.revenue,getAgg(py,scopeKey,'Top100').revenue)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(price,avgFormula(scopeKey,'Top100',py)),rate(avgCache(scopeKey,'Top100',m),avgCache(scopeKey,'Top100',py))),
   cell(sf('sales',scopeKey,'头部（1-20）',m),getAgg(m,scopeKey,'头部（1-20）').sales),
   cell(sf('sales',scopeKey,'中部（21-50）',m),getAgg(m,scopeKey,'中部（21-50）').sales),
   cell(sf('sales',scopeKey,'尾部（51-100）',m),getAgg(m,scopeKey,'尾部（51-100）').sales),
   show?'展示样本：不纳入核心增速':'Top100池按独立代表记录截取；头/中/尾销量来自同一Top100池']);}
 return rows;
}
function writeTopSheet(scopeKey,name){
 const rows=topRows(scopeKey);rows.push([]);rows.push(['年度/周期','范围','Top100 Listing累计','销量','销售额','平均标价','加权成交均价','销量变化','销售额变化']);
 const annualStart=rows.length+1;
 for(const p of periods){const label=p[0],a=p[1],b=p[2];rows.push([label,a+'-'+b,cell(sfr('list',scopeKey,'Top100',a,b),sumCache('list',scopeKey,'Top100',a,b)),cell(sfr('sales',scopeKey,'Top100',a,b),sumCache('sales',scopeKey,'Top100',a,b)),cell(sfr('revenue',scopeKey,'Top100',a,b),sumCache('revenue',scopeKey,'Top100',a,b)),cell(periodAvgFormula(scopeKey,'Top100',a,b),periodAvg(scopeKey,'Top100',a,b)),cell(periodWeightedFormula(scopeKey,'Top100',a,b),periodWeighted(scopeKey,'Top100',a,b)),null,null]);}
 const ws=addSheet(name,rows,[14,22,14,14,14,14,12,16,14,14,16,13,13,14,15,13,14,15,14,14,15,14,42],[[0,0,0,22],[rows.length-1,0,rows.length-1,8]]);
 return {sheet:name,annualStart,annualEnd:annualStart+periods.length-1,monthlyStart:5,monthlyEnd:4+display.length};
}
function tierRows(scopeKey,tierList){
 const rows=[];for(const m of display)for(const t of tierList){const tier=t[0],pm=prevM(m),py=prevY(m),show=m===SHOW_B,r=getAgg(m,scopeKey,tier);
  rows.push([m,tier,t[1]+'-'+t[2],
   cell(sf('list',scopeKey,tier,m),r.list),cell(sf('raw',scopeKey,tier,m),r.raw),cell(sf('sales',scopeKey,tier,m),r.sales),cell(sf('revenue',scopeKey,tier,m),r.revenue),
   cell(avgFormula(scopeKey,tier,m),avgCache(scopeKey,tier,m)),cell(weightedFormula(scopeKey,tier,m),weightedCache(scopeKey,tier,m)),
   cell('IFERROR('+sf('sales',scopeKey,tier,m)+'/'+sf('sales',scopeKey,'Top100',m)+',"")',getAgg(m,scopeKey,'Top100').sales? r.sales/getAgg(m,scopeKey,'Top100').sales:null),
   cell('IFERROR('+sf('revenue',scopeKey,tier,m)+'/'+sf('revenue',scopeKey,'Top100',m)+',"")',getAgg(m,scopeKey,'Top100').revenue? r.revenue/getAgg(m,scopeKey,'Top100').revenue:null),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sf('sales',scopeKey,tier,m),sf('sales',scopeKey,tier,pm)),rate(r.sales,getAgg(pm,scopeKey,tier).sales)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sf('revenue',scopeKey,tier,m),sf('revenue',scopeKey,tier,pm)),rate(r.revenue,getAgg(pm,scopeKey,tier).revenue)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sf('sales',scopeKey,tier,m),sf('sales',scopeKey,tier,py)),rate(r.sales,getAgg(py,scopeKey,tier).sales)),
   show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sf('revenue',scopeKey,tier,m),sf('revenue',scopeKey,tier,py)),rate(r.revenue,getAgg(py,scopeKey,tier).revenue)),
   show?'展示样本':'层级范围来自同一Top100池']);}
 return rows;
}
function writeTierSheet(scopeKey,name){
 const rows=[[''],[''],[],['粗分层（领导结论）'],['月份','分层','BSR范围','Listing数','原始行数','销量','销售额','平均标价','加权成交均价','Top100销量份额','Top100销售额份额','MOM销量','MOM销售额','YOY销量','YOY销售额','备注']];
 rows[0]=[scopes.find(s=>s[0]===scopeKey)[1]+' · BSR前100分层分析（公式页）'];
 rows[1]=['粗分层：头部1-20、中部21-50、尾部51-100；细分层：1-5、6-10、11-20、21-50、51-100。两组均从同一Top100池计算。'];
 const coarseHeader=5,coarseStart=6;rows.push(...tierRows(scopeKey,coarse));
 const coarseEnd=rows.length;
 const fineCaption=coarseEnd+3;rows.push([]);rows.push(['细分层（测款与链接规划）']);rows.push(['月份','分层','BSR范围','Listing数','原始行数','销量','销售额','平均标价','加权成交均价','Top100销量份额','Top100销售额份额','MOM销量','MOM销售额','YOY销量','YOY销售额','备注']);
 const fineStart=rows.length+1;rows.push(...tierRows(scopeKey,fine));const fineEnd=rows.length;
 const h1Caption=rows.length+3;rows.push([]);rows.push(['核心H1分层对比（2025H1 vs 2026H1）']);rows.push(['分层','2025H1销量','2026H1销量','销量YOY','2025H1销售额','2026H1销售额','销售额YOY','2026H1销量份额']);
 const h1Start=rows.length+1;for(const t of coarse){const tier=t[0],s25=sumCache('sales',scopeKey,tier,'202501','202506'),s26=sumCache('sales',scopeKey,tier,'202601','202606'),r25=sumCache('revenue',scopeKey,tier,'202501','202506'),r26=sumCache('revenue',scopeKey,tier,'202601','202606'),all26=sumCache('sales',scopeKey,'Top100','202601','202606');rows.push([tier,cell(sfr('sales',scopeKey,tier,'202501','202506'),s25),cell(sfr('sales',scopeKey,tier,'202601','202606'),s26),cell(rateFormula(sfr('sales',scopeKey,tier,'202601','202606'),sfr('sales',scopeKey,tier,'202501','202506')),rate(s26,s25)),cell(sfr('revenue',scopeKey,tier,'202501','202506'),r25),cell(sfr('revenue',scopeKey,tier,'202601','202606'),r26),cell(rateFormula(sfr('revenue',scopeKey,tier,'202601','202606'),sfr('revenue',scopeKey,tier,'202501','202506')),rate(r26,r25)),cell('IFERROR('+sfr('sales',scopeKey,tier,'202601','202606')+'/'+sfr('sales',scopeKey,'Top100','202601','202606')+',"")',all26?s26/all26:null)]);}
 const refCaption=rows.length+3;rows.push([]);rows.push(['年度分层对照（参考表头）']);
 rows.push(['Tier 层级','2025H1月均Listing','2026H1月均Listing','Listing变化','2025H1销量','2026H1销量','销量变化','2025H1销售额','2026H1销售额','销售额变化','2025H1单Listing销量','2026H1单Listing销量','效率变化','2025H1 ASP','2026H1 ASP','ASP变化','判断']);
 const refStart=rows.length+1;
 for(const t of coarse){
   const tier=t[0],s25=sumCache('sales',scopeKey,tier,'202501','202506'),s26=sumCache('sales',scopeKey,tier,'202601','202606'),r25=sumCache('revenue',scopeKey,tier,'202501','202506'),r26=sumCache('revenue',scopeKey,tier,'202601','202606');
   const l25=sumCache('list',scopeKey,tier,'202501','202506')/6,l26=sumCache('list',scopeKey,tier,'202601','202606')/6;
   const rr=rows.length+1;
   rows.push([tier,cell('IFERROR('+sfr('list',scopeKey,tier,'202501','202506')+'/6,"")',l25),cell('IFERROR('+sfr('list',scopeKey,tier,'202601','202606')+'/6,"")',l26),cell(rateFormula('C'+rr,'B'+rr),rate(l26,l25)),cell(sfr('sales',scopeKey,tier,'202501','202506'),s25),cell(sfr('sales',scopeKey,tier,'202601','202606'),s26),cell(rateFormula('F'+rr,'E'+rr),rate(s26,s25)),cell(sfr('revenue',scopeKey,tier,'202501','202506'),r25),cell(sfr('revenue',scopeKey,tier,'202601','202606'),r26),cell(rateFormula('I'+rr,'H'+rr),rate(r26,r25)),cell('IFERROR(E'+rr+'/B'+rr+',"")',l25?s25/l25:null),cell('IFERROR(F'+rr+'/C'+rr+',"")',l26?s26/l26:null),cell(rateFormula('L'+rr,'K'+rr),l25&&l26?(s26/l26)/(s25/l25)-1:null),cell(periodAvgFormula(scopeKey,tier,'202501','202506'),periodAvg(scopeKey,tier,'202501','202506')),cell(periodAvgFormula(scopeKey,tier,'202601','202606'),periodAvg(scopeKey,tier,'202601','202606')),cell(rateFormula('O'+rr,'N'+rr),rate(periodAvg(scopeKey,tier,'202601','202606'),periodAvg(scopeKey,tier,'202501','202506'))),cell('IF(AND(G'+rr+'>0,J'+rr+'>0),"量额同增：优先扩张",IF(AND(G'+rr+'>0,J'+rr+'<0),"量增额降：复核价格","先复核")',rate(s26,s25)>0&&rate(r26,r25)>0?'量额同增：优先扩张':rate(s26,s25)>0&&rate(r26,r25)<0?'量增额降：复核价格':'先复核')]);
 }
 const ws=addSheet(name,rows,[14,18,12,12,12,14,15,14,16,17,18,15,15,13,14,14,14,35],[[0,0,0,15],[3,0,3,15],[fineCaption-1,0,fineCaption-1,15],[h1Caption-1,0,h1Caption-1,7],[refCaption-1,0,refCaption-1,16]]);
 return {sheet:name,coarseStart,coarseEnd,fineStart,fineEnd,h1Start,h1End:h1Start+coarse.length-1,refStart,refEnd:refStart+coarse.length-1};
}
const refs={};
refs.overall={monthly:writeMonthlySheet('overall','01_整体市场-月度年度'),top:writeTopSheet('overall','02_整体市场-BSR前100'),tier:writeTierSheet('overall','03_整体市场-BSR分层')};
refs.pp={monthly:writeMonthlySheet('pp','04_PP管-月度年度'),top:writeTopSheet('pp','05_PP管-BSR前100'),tier:writeTierSheet('pp','06_PP管-BSR分层')};
refs.nonpp={monthly:writeMonthlySheet('nonpp','07_非PP高客单-月度年度'),top:writeTopSheet('nonpp','08_非PP高客单-BSR前100'),tier:writeTierSheet('nonpp','09_非PP高客单-BSR分层')};
const genimoTopKeys=(a,b)=>{const out=new Set();for(const m of months.filter(x=>x>=a&&x<=b))for(const r of topBy.get(m+'|genimo')||[])out.add(r.key);return out;};


function annualRef(scopeKey,metric,idx){const r=refs[scopeKey].monthly.annualStart+idx,col={raw:'C',list:'D',dup:'E',sales:'F',revenue:'G',price:'H',weighted:'I'}[metric];return "'"+refs[scopeKey].monthly.sheet+"'!$"+col+"$"+r;}
const ovRows=[[''],[''],[],['核心指标','2025H1','2026H1','变化','主口径','数据状态','对应子表','公式来源','参考值（仅参考）','交付阅读提示']];
ovRows[0]=['户外地垫市场分析 · 交付总览与结论'];
ovRows[1]=['整体市场=PP管+非PP高客单产品线；GENIMO是整体市场中的品牌视角，不作为第三个分类。所有结果指标由各自子表公式链计算。'];
const ovData=[
 ['整体市场销量',annualRef('overall','sales',2),annualRef('overall','sales',3),'PP+非PP','核心父ASIN代表记录','01/02/03','01年度表',benchmark.industry?.growthPct==null?null:benchmark.industry.growthPct/100,'先看整体，再看分类'],
 ['整体市场销售额',annualRef('overall','revenue',2),annualRef('overall','revenue',3),'PP+非PP','核心父ASIN代表记录','01/02/03','01年度表',null,'销售额与销量分别判断'],
 ['PP销量',annualRef('pp','sales',2),annualRef('pp','sales',3),'标题含完整单词plastic','核心父ASIN代表记录','04/05/06','04年度表',null,'PP与非PP必须可加总'],
 ['非PP高客单销量',annualRef('nonpp','sales',2),annualRef('nonpp','sales',3),'整体排除PP的补集','核心父ASIN代表记录','07/08/09','07年度表',null,'不是价格阈值分类'],
 ['GENIMO整体市场销量份额',null,null,'GENIMO/整体市场','核心父ASIN代表记录','10/11/12','10品牌份额表',null,'品牌份额主分母为整体'],
 ['整体市场BSR前100销量',annualRef('overall','sales',2),annualRef('overall','sales',3),'整体Top100独立池','核心父ASIN代表记录','02/03','02 Top100表',benchmark.bsrTop100?.growthPct==null?null:benchmark.bsrTop100.growthPct/100,'Top100最多100个独立Listing']
];
for(const d of ovData){const [label,b25,b26,scopeText,statusText,sub,src,refv,prompt]=d;const v25=label==='整体市场销量'?sumCache('sales','overall','全部','202501','202506'):label==='整体市场销售额'?sumCache('revenue','overall','全部','202501','202506'):label==='PP销量'?sumCache('sales','pp','全部','202501','202506'):label==='非PP高客单销量'?sumCache('sales','nonpp','全部','202501','202506'):label==='整体市场BSR前100销量'?sumCache('sales','overall','Top100','202501','202506'):null;const v26=label==='整体市场销量'?sumCache('sales','overall','全部','202601','202606'):label==='整体市场销售额'?sumCache('revenue','overall','全部','202601','202606'):label==='PP销量'?sumCache('sales','pp','全部','202601','202606'):label==='非PP高客单销量'?sumCache('sales','nonpp','全部','202601','202606'):label==='整体市场BSR前100销量'?sumCache('sales','overall','Top100','202601','202606'):null;ovRows.push([label,cell(b25,v25),cell(b26,v26),cell(rateFormula(b26,b25),rate(v26,v25)),scopeText,statusText,sub,src,refv,prompt]);}
const g25=sumCache('sales','genimo','全部','202501','202506'),g26=sumCache('sales','genimo','全部','202601','202606'),o25=sumCache('sales','overall','全部','202501','202506'),o26=sumCache('sales','overall','全部','202601','202606');
ovRows[8]=['GENIMO整体市场销量份额',cell('IFERROR('+sfr('sales','genimo','全部','202501','202506')+'/'+sfr('sales','overall','全部','202501','202506')+',\"\")',o25?g25/o25:null),cell('IFERROR('+sfr('sales','genimo','全部','202601','202606')+'/'+sfr('sales','overall','全部','202601','202606')+',\"\")',o26?g26/o26:null),cell('IFERROR(C9/B9-1,\"\")',o25&&o26&&g25&&g26?(g26/o26)/(g25/o25)-1:null),'GENIMO/整体市场','核心父ASIN代表记录','10/11/12','10品牌份额表',null,'品牌份额主分母为整体'];
ovRows[10]=[''];
const logicStart=ovRows.length+2;ovRows.push(['关键逻辑校验','结果','期望/阈值','实际差额','状态','说明']);
ovRows.push(['整体=PP+非PP（销量）',cell(annualRef('overall','sales',3)+'-('+annualRef('pp','sales',3)+'+'+annualRef('nonpp','sales',3)+')',0),0,cell('ABS(B13)',0),cell('IF(D13<0.5,"通过","失败")','通过'),'分类必须回到整体市场。']);
ovRows.push(['整体=PP+非PP（销售额）',cell(annualRef('overall','revenue',3)+'-('+annualRef('pp','revenue',3)+'+'+annualRef('nonpp','revenue',3)+')',0),0,cell('ABS(B14)',0),cell('IF(D14<0.5,"通过","失败")','通过'),'销售额也必须校验。']);
const maxTop100=Math.max(...display.map(m=>Math.max(getAgg(m,'overall','Top100').list,getAgg(m,'pp','Top100').list,getAgg(m,'nonpp','Top100').list)));
ovRows.push(['Top100最大Listing数',cell('MAX(\'02_整体市场-BSR前100\'!$C$5:$C$'+refs.overall.top.monthlyEnd+',\'05_PP管-BSR前100\'!$C$5:$C$'+refs.pp.top.monthlyEnd+',\'08_非PP高客单-BSR前100\'!$C$5:$C$'+refs.nonpp.top.monthlyEnd+')',maxTop100),100,cell('B15-C15',maxTop100-100),cell('IF(D15<=0,"通过","失败")','通过'),'每范围每月最多100个独立Listing。']);
ovRows.push(['2026.07展示父体数',cell('COUNTIFS(\'90_输入_明细\'!$A$3:$A$'+detailEnd+',"202607")',details.filter(r=>r.month==='202607').length),94,cell('B16-C16',details.filter(r=>r.month==='202607').length-94),cell('IF(D16=0,"通过","请复核")','请复核'),'展示样本数量；不作为核心增速。']);
ovRows.push(['明细输入行数',cell('COUNTA(\'90_输入_明细\'!$A$3:$A$'+detailEnd+')',details.length),details.length,cell('B17-C17',0),cell('IF(D17=0,"通过","失败")','通过'),'90页值区行数。']);
ovRows.push([]);
ovRows.push(['核心趋势判断（公式引用，供2027决策）']);
ovRows.push(['范围','销量H1 YOY','销售额H1 YOY','平均标价H1 YOY','趋势判断与2027动作']);
const trendRowStart=ovRows.length+1;
for(const [scopeKey,label] of [['overall','整体市场'],['pp','PP管'],['nonpp','非PP高客单']]){
 const row=ovRows.length+1;
 const s25=annualRef(scopeKey,'sales',2),s26=annualRef(scopeKey,'sales',3),r25=annualRef(scopeKey,'revenue',2),r26=annualRef(scopeKey,'revenue',3),p25=annualRef(scopeKey,'price',2),p26=annualRef(scopeKey,'price',3);
 const fs=rateFormula(s26,s25),fr=rateFormula(r26,r25),fp=rateFormula(p26,p25);
 const sv=sumCache('sales',scopeKey,'全部','202501','202506'),sn=sumCache('sales',scopeKey,'全部','202601','202606'),rv=sumCache('revenue',scopeKey,'全部','202501','202506'),rn=sumCache('revenue',scopeKey,'全部','202601','202606'),pv=periodAvg(scopeKey,'全部','202501','202506'),pn=periodAvg(scopeKey,'全部','202601','202606');
 let action='结合销量、销售额和均价方向配置2027资源';
 if(scopeKey==='overall') action=sn>sv&&rn<rv?'量增额降：2027先检查价格和产品结构，整体链接按PP基本盘与非PP增长分层配置':'结合销量、销售额和均价方向配置2027资源';
 if(scopeKey==='pp') action=sn<sv&&rn>rv?'销量承压但销售额增长：保留高效PP链接，优先核查客单价和结构后再扩量':sn>sv&&rn>rv?'量额同增：可扩大验证，但仍按BSR层级控制库存':'销量或销售额承压：先复核产品与投放效率，再决定扩量';
 if(scopeKey==='nonpp') action=sn>sv&&rn<rv?'销量增长但销售额下降：先筛选高客单和高效中流链接，再扩大非PP链接':'结合销量、销售额和均价方向配置2027资源';
 ovRows.push([label,cell(fs,rate(sn,sv)),cell(fr,rate(rn,rv)),cell(fp,rate(pn,pv)),cell('IF(AND(B'+row+'>0,C'+row+'<0),"量增额降：优先检查价格与产品结构",IF(AND(B'+row+'>0,C'+row+'>0),"量额同增：按BSR层级扩大验证","先复核产品与投放效率，再决定扩量"))',action)]);
}
const genimoTrendRow=ovRows.length+1;
ovRows.push(['GENIMO',cell('D9',o25&&o26&&g25&&g26?(g26/o26)/(g25/o25)-1:null),null,null,cell('IF(B'+genimoTrendRow+'>0,"整体市场份额提升：2027优先补充增长分层的中流链接，并用PP/非PP分别设测款门槛","整体市场份额未提升：先复核留存、退出和进入名单，再调整链接结构")',((g26/o26)/(g25/o25)-1)>0?'整体市场份额提升：2027优先补充增长分层的中流链接，并用PP/非PP分别设测款门槛':'整体市场份额未提升：先复核留存、退出和进入名单，再调整链接结构')]);
ovRows.push(['说明','销量、销售额和均价变化均来自对应月度年度表；文字判断是公式条件输出，不能替代利润、库存、广告容量和测款复核。']);
ovRows.push([]);
const readStart=ovRows.length+2;ovRows.push(['逐表对应关系']);for(const x of [['01/02/03','整体市场：月度年度、BSR前100、粗细分层'],['04/05/06','PP管：月度年度、BSR前100、粗细分层'],['07/08/09','非PP高客单：月度年度、BSR前100、粗细分层'],['10/11/12','GENIMO：整体市场品牌份额、进退层、2027链接规划'],['90/91','输入与聚合：明细代表记录、可重算中间值'],['92/93','校验、来源、数据操作、口径限制']])ovRows.push([x[0],x[1]]);
const ws00=addSheet('00_总览与结论',ovRows,[25,15,15,13,25,22,15,24,18,32],[[0,0,0,9],[1,0,1,9]]);


const brandRows=[['GENIMO 品牌份额 · 整体市场主口径（公式页）'],['主指标=GENIMO/整体市场；PP内份额仅作辅助。整体市场包括PP和非PP，GENIMO不从分类加总中单列。'],[],['月份','整体销量','GENIMO销量','整体销量份额','整体销售额','GENIMO销售额','整体销售额份额','PP销量','GENIMO PP销量','PP内销量份额','GENIMO销量MOM','GENIMO销量YOY','GENIMO销售额MOM','GENIMO销售额YOY','备注']];
for(const m of display){const pm=prevM(m),py=prevY(m),show=m===SHOW_B,rO=getAgg(m,'overall','全部'),rG=getAgg(m,'genimo','全部'),rP=getAgg(m,'pp','全部'),rGP=getAgg(m,'genimo_pp','全部');brandRows.push([m,
 cell(sf('sales','overall','全部',m),rO.sales),cell(sf('sales','genimo','全部',m),rG.sales),cell('IFERROR('+sf('sales','genimo','全部',m)+'/'+sf('sales','overall','全部',m)+',"")',rO.sales?rG.sales/rO.sales:null),
 cell(sf('revenue','overall','全部',m),rO.revenue),cell(sf('revenue','genimo','全部',m),rG.revenue),cell('IFERROR('+sf('revenue','genimo','全部',m)+'/'+sf('revenue','overall','全部',m)+',"")',rO.revenue?rG.revenue/rO.revenue:null),
 cell(sf('sales','pp','全部',m),rP.sales),cell(sf('sales','genimo_pp','全部',m),rGP.sales),cell('IFERROR('+sf('sales','genimo_pp','全部',m)+'/'+sf('sales','pp','全部',m)+',"")',rP.sales?rGP.sales/rP.sales:null),
 show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sf('sales','genimo','全部',m),sf('sales','genimo','全部',pm)),rate(rG.sales,getAgg(pm,'genimo','全部').sales)),
 show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sf('sales','genimo','全部',m),sf('sales','genimo','全部',py)),rate(rG.sales,getAgg(py,'genimo','全部').sales)),
 show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sf('revenue','genimo','全部',m),sf('revenue','genimo','全部',pm)),rate(rG.revenue,getAgg(pm,'genimo','全部').revenue)),
 show?cell('IFERROR(0/0,"")',null):cell(rateFormula(sf('revenue','genimo','全部',m),sf('revenue','genimo','全部',py)),rate(rG.revenue,getAgg(py,'genimo','全部').revenue)),
 show?'展示样本':'整体市场为主分母']);}
const bsum=brandRows.length+2;brandRows.push([]);brandRows.push(['GENIMO核心H1品牌份额']);brandRows.push(['指标','2025H1','2026H1','变化','2025H1 PP内','2026H1 PP内','PP内变化','解释']);
const gP25=sumCache('sales','genimo_pp','全部','202501','202506'),gP26=sumCache('sales','genimo_pp','全部','202601','202606'),p25=sumCache('sales','pp','全部','202501','202506'),p26=sumCache('sales','pp','全部','202601','202606');
brandRows.push(['销量份额',cell('IFERROR('+sfr('sales','genimo','全部','202501','202506')+'/'+sfr('sales','overall','全部','202501','202506')+',"")',o25?g25/o25:null),cell('IFERROR('+sfr('sales','genimo','全部','202601','202606')+'/'+sfr('sales','overall','全部','202601','202606')+',"")',o26?g26/o26:null),cell('IFERROR(C'+(bsum+3)+'/B'+(bsum+3)+'-1,"")',o25&&o26&&g25&&g26?(g26/o26)/(g25/o25)-1:null),cell('IFERROR('+sfr('sales','genimo_pp','全部','202501','202506')+'/'+sfr('sales','pp','全部','202501','202506')+',"")',p25?gP25/p25:null),cell('IFERROR('+sfr('sales','genimo_pp','全部','202601','202606')+'/'+sfr('sales','pp','全部','202601','202606')+',"")',p26?gP26/p26:null),cell('IFERROR(F'+(bsum+3)+'/E'+(bsum+3)+'-1,"")',p25&&p26&&gP25&&gP26?(gP26/p26)/(gP25/p25)-1:null),'整体为主分母，PP内为辅助'],
 ['销售额份额',cell('IFERROR('+sfr('revenue','genimo','全部','202501','202506')+'/'+sfr('revenue','overall','全部','202501','202506')+',"")',null),cell('IFERROR('+sfr('revenue','genimo','全部','202601','202606')+'/'+sfr('revenue','overall','全部','202601','202606')+',"")',null),null,null,null,null,'整体销售额份额'],
 ['说明',null,null,null,null,null,null,null,'GENIMO PP内份额仅用于产品线内部观察']);
brandRows.push(['策略提示','若要稳住GENIMO整体市场份额，应先看整体市场中增长的PP/非PP分层，再配置对应中流链接；高客单非PP是PP补集；B5输入目标链接总数后，12页给出链接分配草案。']);
const ws10=addSheet('10_GENIMO-品牌份额',brandRows,[12,14,14,15,15,16,16,12,15,14,15,15,16,16,38],[[0,0,0,14],[1,0,1,14],[bsum-1,0,bsum-1,14]]);
const cohortKeys25=genimoTopKeys('202501','202506'),cohortKeys26=genimoTopKeys('202601','202606');
const retained=[...cohortKeys25].filter(k=>cohortKeys26.has(k)).sort(),exited=[...cohortKeys25].filter(k=>!cohortKeys26.has(k)).sort(),entered=[...cohortKeys26].filter(k=>!cohortKeys25.has(k)).sort();
const cohortRows=[['GENIMO BSR前100进退层 · H1 cohort（公式页）'],['留存=2025H1和2026H1均出现，退出=仅2025H1，进入=仅2026H1。集合识别由脚本完成，销量/销售额由90页明细公式汇总。'],[],['状态','Listing键','2025H1出现月数','2026H1出现月数','2025H1销量','2026H1销量','2025H1销售额','2026H1销售额','代表品牌','代表标题']];
const detailA="'90_输入_明细'!$A$3:$A$"+detailEnd,detailB="'90_输入_明细'!$B$3:$B$"+detailEnd,detailK="'90_输入_明细'!$K$3:$K$"+detailEnd,detailL="'90_输入_明细'!$L$3:$L$"+detailEnd,detailF="'90_输入_明细'!$F$3:$F$"+detailEnd,detailG="'90_输入_明细'!$G$3:$G$"+detailEnd;
const dStats=(key,a,b)=>{const rs=details.filter(r=>r.key===key&&r.month>=a&&r.month<=b);return{months:new Set(rs.map(r=>r.month)).size,sales:total(rs,r=>r.sales),revenue:total(rs,r=>r.revenue)};};
for(const item of [...retained.map(k=>['留存',k]),...exited.map(k=>['退出',k]),...entered.map(k=>['进入',k])]){const row=cohortRows.length+1;const kRef='B'+row,s25=dStats(item[1],'202501','202506'),s26=dStats(item[1],'202601','202606'),repr=details.find(r=>r.key===item[1])||{};cohortRows.push([item[0],item[1],
 cell('COUNTIFS('+detailA+',">=202501",'+detailA+',"<=202506",'+detailB+','+kRef+')',s25.months),
 cell('COUNTIFS('+detailA+',">=202601",'+detailA+',"<=202606",'+detailB+','+kRef+')',s26.months),
 cell('SUMIFS('+detailK+','+detailA+',">=202501",'+detailA+',"<=202506",'+detailB+','+kRef+')',s25.sales),
 cell('SUMIFS('+detailK+','+detailA+',">=202601",'+detailA+',"<=202606",'+detailB+','+kRef+')',s26.sales),
 cell('SUMIFS('+detailL+','+detailA+',">=202501",'+detailA+',"<=202506",'+detailB+','+kRef+')',s25.revenue),
 cell('SUMIFS('+detailL+','+detailA+',">=202601",'+detailA+',"<=202606",'+detailB+','+kRef+')',s26.revenue),
 cell('IFERROR(INDEX('+detailF+',MATCH('+kRef+','+detailB+',0)),"")',repr.brand||''),
 cell('IFERROR(INDEX('+detailG+',MATCH('+kRef+','+detailB+',0)),"")',repr.title||'')]);}
const cohortSummary=cohortRows.length+2;cohortRows.push([]);cohortRows.push(['集合统计','数量','说明','可复核位置']);cohortRows.push(['留存',cell('COUNTIF($A$5:$A$'+(cohortRows.length-1)+',"留存")',retained.length),'两期均在GENIMO Top100集合','本页明细']);cohortRows.push(['退出',cell('COUNTIF($A$5:$A$'+(cohortRows.length-1)+',"退出")',exited.length),'仅2025H1在集合','本页明细']);cohortRows.push(['进入',cell('COUNTIF($A$5:$A$'+(cohortRows.length-1)+',"进入")',entered.length),'仅2026H1在集合','本页明细']);
const ws11=addSheet('11_GENIMO-进退层',cohortRows,[12,20,16,16,15,15,16,16,18,62],[[0,0,0,9],[1,0,1,9]]);


const planRows=[['GENIMO 2027链接规划 · 公式草案'],['先输入目标链接总数，再观察整体市场、PP和非PP的核心Top100层级份额。规划是公式草案，仍需结合利润、库存、广告容量和测款结果复核。'],[],['输入项','数值'],['计划新增链接总数',0],[],['范围','层级','2025H1销量','2026H1销量','销量YOY','2026H1销量份额','建议链接数','产品策略提示','测款优先级','库存建议','广告建议','利润/晋级退出门槛','公式来源']];
for(const s of [['overall','整体市场'],['pp','PP管'],['nonpp','非PP高客单']])for(const t of coarse){const tier=t[0],idx=coarse.findIndex(x=>x[0]===tier),h=refs[s[0]].tier.h1Start+idx,s25=sumCache('sales',s[0],tier,'202501','202506'),s26=sumCache('sales',s[0],tier,'202601','202606'),all26=sumCache('sales',s[0],'Top100','202601','202606');const product=tier==='头部（1-20）'?'品牌心智与规模款；控制库存周转':tier==='中部（21-50）'?'优先测款和扩充中流链接':'长尾机会小批量验证，控制试错成本';const inventory=tier==='头部（1-20）'?'主推尺寸保持8-12周安全库存':tier==='中部（21-50）'?'小批量备货，按4-8周周转滚动':'试销库存控制在2-4周，达标再补';const ads=tier==='头部（1-20）'?'品牌词与核心词守位，按转化率加预算':tier==='中部（21-50）'?'设置独立测款预算，达标后逐步放量':'低预算长尾测试，避免为未验证链接持续烧费';const gate=tier==='头部（1-20）'?'毛利率和库存周转达标才扩充同类链接':tier==='中部（21-50）'?'连续4周达到转化/毛利门槛则晋级，否则降级复盘':'连续4周未达转化/毛利门槛则退出或换款';planRows.push([s[1],tier,cell(sfr('sales',s[0],tier,'202501','202506'),s25),cell(sfr('sales',s[0],tier,'202601','202606'),s26),cell(rateFormula(sfr('sales',s[0],tier,'202601','202606'),sfr('sales',s[0],tier,'202501','202506')),rate(s26,s25)),cell("'"+refs[s[0]].tier.sheet+"'!H"+h,all26?s26/all26:null),cell('IF($B$5=0,"",ROUND($B$5*F'+(planRows.length+1)+',0))',null),product,cell('IF(E'+(planRows.length+1)+'>0,"优先","观察")',rate(s26,s25)>0?'优先':'观察'),inventory,ads,gate,refs[s[0]].tier.sheet+'!H'+h]);}
planRows.push([]);planRows.push(['使用说明','B5为唯一输入；建议链接数按各范围核心H1销量份额计算。PP/非PP/整体分别展示，GENIMO品牌策略先使用整体市场份额作为主口径。']);
const ws12=addSheet('12_GENIMO-2027规划',planRows,[18,18,16,16,13,17,17,44,14,34,42,46,40],[[0,0,0,12],[1,0,1,12],[3,0,3,1],[planRows.length-1,0,planRows.length-1,12]]);

// Weekly report output.  Market inputs are monthly snapshots, so the report
// keeps an explicit report-week input and never fabricates a weekly market
// observation from a monthly estimate.
const topOverallRows=display.map((m,i)=>({m,row:5+i}));
const weeklyRows=[['户外地垫周报 · 市场、BSR分层、品牌和利润测试'],['报告周次/日期为可编辑输入；市场数据来自月度快照，最新核心实绩为2026.06，2026.07仅展示样本。'],[],['报告控制项','值','来源/说明','状态'],['报告周次/日期','', '请填写本周周次或日期','待填写'],['市场数据最新月','202606','market.db父体快照核心截止月','已锁定'],['周报统计口径','整体市场=PP+非PP；GENIMO为品牌视角','参考：PP workbook / 战略 workbook','已锁定'],[],['核心市场指标','整体市场（PP+非PP）','PP管/Plastic','非PP补集','GENIMO整体份额','数据状态'],['2025H1销量'],['2026H1销量'],['销量变化'],['2025H1销售额($)'],['2026H1销售额($)'],['销售额变化'],['2025H1 ASP($)'],['2026H1 ASP($)'],['ASP变化'],['数据来源说明','market.db结果用于可复核明细；计划部BI/BSR基准单独列示，跨源不混算']];
const bench25=benchmark.industry?.sales2025H1??1781765,bench26=benchmark.industry?.sales2026H1??1831843,bsr25=benchmark.bsrTop100?.sales2025H1??1066818,bsr26=benchmark.bsrTop100?.sales2026H1??1121130;
weeklyRows.push([]);weeklyRows.push(['领导验收主基准（参考 workbook）','2025H1','2026H1','变化','来源工作表/单元格','用途']);
let benchRow=weeklyRows.length+1;weeklyRows.push(['计划部BI全类目Outdoor Rugs销量',bench25,bench26,cell('IFERROR(C'+benchRow+'/B'+benchRow+'-1,"" )',rate(bench26,bench25)),'行业大盘数据!H4:H9 / I4:I9','领导验收主基准（销量）']);
benchRow=weeklyRows.length+1;weeklyRows.push(['计划部BSR Top100销量',bsr25,bsr26,cell('IFERROR(C'+benchRow+'/B'+benchRow+'-1,"" )',rate(bsr26,bsr25)),'BSR底表-美国!AX:BC / BJ:BO','辅助验收基准（销量）']);
benchRow=weeklyRows.length+1;weeklyRows.push(['market.db父体快照销量（不可比明细参考）',sumCache('sales','overall','全部','202501','202506'),sumCache('sales','overall','全部','202601','202606'),cell('IFERROR(C'+benchRow+'/B'+benchRow+'-1,"" )',rate(sumCache('sales','overall','全部','202601','202606'),sumCache('sales','overall','全部','202501','202506'))),'01_整体市场-月度年度!年度区块','不同统计单元，不替代领导验收基准']);
weeklyRows.push([]);weeklyRows.push(['Monthly 前100趋势（参考表头）']);
weeklyRows.push(['Month 月份','Listings','Sales 销量','Revenue 销售额($)','ASP 客单价($)','Sales YoY 销量同比','Revenue YoY 销售额同比','Head Sales 头部销量','Middle Sales 中部销量','Tail Sales 尾部销量','数据状态']);
for(const {m,row} of topOverallRows){const show=m===SHOW_B;weeklyRows.push([cell("'02_整体市场-BSR前100'!A"+row,m),cell("'02_整体市场-BSR前100'!C"+row,getAgg(m,'overall','Top100').list),cell("'02_整体市场-BSR前100'!F"+row,getAgg(m,'overall','Top100').sales),cell("'02_整体市场-BSR前100'!H"+row,getAgg(m,'overall','Top100').revenue),cell("'02_整体市场-BSR前100'!J"+row,avgCache('overall','Top100',m)),show?null:cell("'02_整体市场-BSR前100'!P"+row,rate(getAgg(m,'overall','Top100').sales,getAgg(prevYear(m),'overall','Top100').sales)),show?null:cell("'02_整体市场-BSR前100'!Q"+row,rate(getAgg(m,'overall','Top100').revenue,getAgg(prevYear(m),'overall','Top100').revenue)),cell("'02_整体市场-BSR前100'!T"+row,getAgg(m,'overall','头部（1-20）').sales),cell("'02_整体市场-BSR前100'!U"+row,getAgg(m,'overall','中部（21-50）').sales),cell("'02_整体市场-BSR前100'!V"+row,getAgg(m,'overall','尾部（51-100）').sales),m===SHOW_B?'展示样本':(m>=CORE_A&&m<=CORE_B?'核心':'历史基线')]);}
const kpiRef=(s,m,i)=>{const a=i===2?'202501':'202601',b=i===2?'202506':'202606',ref=annualRef(s,m,i);const v=m==='sales'?sumCache(m,s,'全部',a,b):m==='revenue'?sumCache(m,s,'全部',a,b):periodAvg(s,'全部',a,b);return cell(ref,v)};
const rowSales25=weeklyRows.findIndex(r=>r[0]==='2025H1销量')+1,rowSales26=weeklyRows.findIndex(r=>r[0]==='2026H1销量')+1,rowRev25=weeklyRows.findIndex(r=>r[0]==='2025H1销售额($)')+1,rowRev26=weeklyRows.findIndex(r=>r[0]==='2026H1销售额($)')+1,rowAsp25=weeklyRows.findIndex(r=>r[0]==='2025H1 ASP($)')+1,rowAsp26=weeklyRows.findIndex(r=>r[0]==='2026H1 ASP($)')+1;
for(const [label,s,m,i] of [['2025H1销量','overall','sales',2],['2026H1销量','overall','sales',3],['2025H1销售额($)','overall','revenue',2],['2026H1销售额($)','overall','revenue',3],['2025H1 ASP($)','overall','price',2],['2026H1 ASP($)','overall','price',3]]){const rr=weeklyRows.findIndex(r=>r[0]===label);weeklyRows[rr]=[label,kpiRef(s,m,i),kpiRef('pp',m,i),kpiRef('nonpp',m,i),null,'核心H1'];}
const salesChangeRow=weeklyRows.findIndex(r=>r[0]==='销量变化'),revenueChangeRow=weeklyRows.findIndex(r=>r[0]==='销售额变化'),aspChangeRow=weeklyRows.findIndex(r=>r[0]==='ASP变化');
weeklyRows[salesChangeRow]=['销量变化',cell('IFERROR(B'+rowSales26+'/B'+rowSales25+'-1,"" )',rate(sumCache('sales','overall','全部','202601','202606'),sumCache('sales','overall','全部','202501','202506'))),cell('IFERROR(C'+rowSales26+'/C'+rowSales25+'-1,"" )',rate(sumCache('sales','pp','全部','202601','202606'),sumCache('sales','pp','全部','202501','202506'))),cell('IFERROR(D'+rowSales26+'/D'+rowSales25+'-1,"" )',rate(sumCache('sales','nonpp','全部','202601','202606'),sumCache('sales','nonpp','全部','202501','202506'))),null,'公式=2026H1/2025H1-1'];
weeklyRows[revenueChangeRow]=['销售额变化',cell('IFERROR(B'+rowRev26+'/B'+rowRev25+'-1,"" )',rate(sumCache('revenue','overall','全部','202601','202606'),sumCache('revenue','overall','全部','202501','202506'))),cell('IFERROR(C'+rowRev26+'/C'+rowRev25+'-1,"" )',rate(sumCache('revenue','pp','全部','202601','202606'),sumCache('revenue','pp','全部','202501','202506'))),cell('IFERROR(D'+rowRev26+'/D'+rowRev25+'-1,"" )',rate(sumCache('revenue','nonpp','全部','202601','202606'),sumCache('revenue','nonpp','全部','202501','202506'))),null,'公式=2026H1/2025H1-1'];
weeklyRows[aspChangeRow]=['ASP变化',cell('IFERROR(B'+rowAsp26+'/B'+rowAsp25+'-1,"" )',rate(periodAvg('overall','全部','202601','202606'),periodAvg('overall','全部','202501','202506'))),cell('IFERROR(C'+rowAsp26+'/C'+rowAsp25+'-1,"" )',rate(periodAvg('pp','全部','202601','202606'),periodAvg('pp','全部','202501','202506'))),cell('IFERROR(D'+rowAsp26+'/D'+rowAsp25+'-1,"" )',rate(periodAvg('nonpp','全部','202601','202606'),periodAvg('nonpp','全部','202501','202506'))),null,'公式=2026H1/2025H1-1'];
weeklyRows.push([]);weeklyRows.push(['BSR分层年度对照（参考表头）']);weeklyRows.push(['Tier 层级','2025H1月均Listing','2026H1月均Listing','Listing变化','2025H1销量','2026H1销量','销量变化','2025H1销售额','2026H1销售额','销售额变化','2025H1单Listing销量','2026H1单Listing销量','效率变化','2025H1 ASP','2026H1 ASP','ASP变化','判断']);
for(let i=0;i<coarse.length;i++){
 const rr=refs.overall.tier.refStart+i,tier=coarse[i][0];
 const s25=sumCache('sales','overall',tier,'202501','202506'),s26=sumCache('sales','overall',tier,'202601','202606'),r25=sumCache('revenue','overall',tier,'202501','202506'),r26=sumCache('revenue','overall',tier,'202601','202606');
 const l25=sumCache('list','overall',tier,'202501','202506')/6,l26=sumCache('list','overall',tier,'202601','202606')/6;
 const p25=periodAvg('overall',tier,'202501','202506'),p26=periodAvg('overall',tier,'202601','202606');
 const judgment=rate(s26,s25)>0&&rate(r26,r25)>0?'量额同增：优先扩张':rate(s26,s25)>0&&rate(r26,r25)<0?'量增额降：复核价格':'先复核';
 weeklyRows.push([cell("'03_整体市场-BSR分层'!A"+rr,tier),cell("'03_整体市场-BSR分层'!B"+rr,l25),cell("'03_整体市场-BSR分层'!C"+rr,l26),cell("'03_整体市场-BSR分层'!D"+rr,rate(l26,l25)),cell("'03_整体市场-BSR分层'!E"+rr,s25),cell("'03_整体市场-BSR分层'!F"+rr,s26),cell("'03_整体市场-BSR分层'!G"+rr,rate(s26,s25)),cell("'03_整体市场-BSR分层'!H"+rr,r25),cell("'03_整体市场-BSR分层'!I"+rr,r26),cell("'03_整体市场-BSR分层'!J"+rr,rate(r26,r25)),cell("'03_整体市场-BSR分层'!K"+rr,l25?s25/l25:null),cell("'03_整体市场-BSR分层'!L"+rr,l26?s26/l26:null),cell("'03_整体市场-BSR分层'!M"+rr,l25&&l26?(s26/l26)/(s25/l25)-1:null),cell("'03_整体市场-BSR分层'!N"+rr,p25),cell("'03_整体市场-BSR分层'!O"+rr,p26),cell("'03_整体市场-BSR分层'!P"+rr,rate(p26,p25)),cell("'03_整体市场-BSR分层'!Q"+rr,judgment)]);
}
weeklyRows.push([]);weeklyRows.push(['GENIMO品牌与进退层']);weeklyRows.push(['指标','2025H1','2026H1','变化','公式/来源']);weeklyRows.push(['整体市场销量份额',cell('IFERROR('+sfr('sales','genimo','全部','202501','202506')+'/'+sfr('sales','overall','全部','202501','202506')+',"")',g25/o25),cell('IFERROR('+sfr('sales','genimo','全部','202601','202606')+'/'+sfr('sales','overall','全部','202601','202606')+',"")',g26/o26),cell('IFERROR(C'+(weeklyRows.length+1)+'/B'+(weeklyRows.length+1)+'-1,"")',(g26/o26)/(g25/o25)-1),'10_GENIMO-品牌份额']);
weeklyRows.push(['GENIMO Top100留存/退出/进入',retained.length,exited.length,entered.length,'11_GENIMO-进退层；三类集合由H1 Top100键集合识别']);
weeklyRows.push([]);weeklyRows.push(['利润测试摘要（参考PD17周报公式）']);weeklyRows.push(['指标','值','公式来源','状态']);
const profitEnd=4+profitRows.length;weeklyRows.push(['利润测试Listing行数',cell("COUNTA('14_利润测试'!$A$5:$A$"+profitEnd+")",profitRows.length),'14_利润测试','已载入参考价格/BSR，周销量与成本待填']);weeklyRows.push(['已填7天销量行数',cell("COUNT('14_利润测试'!$F$5:$F$"+profitEnd+")",0),'14_利润测试!F列','待填写']);weeklyRows.push(['通过测款门槛行数',cell("COUNTIF('14_利润测试'!$Y$5:$Y$"+profitEnd+",\"通过测款门槛\")",0),'14_利润测试!Y列','待填写经营参数']);weeklyRows.push(['利润测试规则','Tail→Middle 4周BSR≤100/CVR达标/TACOS≤15%/库存≤90天；Middle→Head 6周BSR≤50/贡献毛利≥15%；头部扩量毛利≥5%/TACOS≤12%；90天未进前100退出','参考：要求内容.txt + 战略 workbook','规则已列入14页']);
const ws13=addSheet('13_周报汇总',weeklyRows,[32,18,18,18,18,18,18,18,18,18,18,18,18,18,18,18,18],[[0,0,0,16],[1,0,1,16],[8,0,8,5],[18,0,18,5],[21,0,21,5],[27,0,27,10],[34,0,34,4],[38,0,38,4]]);

const profitHeaders=['ASIN','SKU','正常售价$','本周折扣形式','折扣%/调后价','7天销量','采购成本RMB','头程$','尾程$','体积ft³','净成交价$','采购成本$','进口税$','仓储$','佣金$','单件利润$','利润率','SKU利润$','来源月销量','来源月销售额$','BSR','BSR层级','源毛利率','源FBA$','测款状态','公式来源','CVR%','TACOS%','库存覆盖天数','动作建议'];
const profitSheetRows=[['GENIMO / BSR Top100 利润测试与测款'],['美元汇率',6.8,'关税率',0.25,'进口税率',0.19,'仓储$/ft³',0.78,'无体积仓储率',0.04,'佣金率',0.15,'最低佣金$',0.3,'最低利润率',0.10,'最低7天销量',1,'Tail→Middle TACOS',0.15,'Head TACOS',0.12,'库存上限天数',90],['公式按参考周报PD17模型：净成交价→采购成本$→进口税→仓储→佣金→单件利润→利润率→SKU利润；周销量和成本字段为空时保持待补状态。'],[],profitHeaders];
for(const p of profitRows){const row=profitSheetRows.length+1,tier=p.bsr==null?'BSR缺失':p.bsr<=20?'头部1-20':p.bsr<=50?'中部21-50':p.bsr<=100?'尾部51-100':'BSR>100';profitSheetRows.push([p.asin,p.sku,p.normalPrice,p.discountType,p.discountValue,p.sevenDaySales,p.purchaseRmb,p.firstMile,p.lastMile,p.volumeFt3,
 cell('IF(D'+row+'="原价调整",E'+row+',IF(OR(E'+row+'="",E'+row+'=0),C'+row+',C'+row+'*(1-E'+row+')))',p.normalPrice),
 cell('IF(G'+row+'="","",G'+row+'/$B$2)',null),cell('IF(L'+row+'="","",L'+row+'*$D$2*$F$2)',null),cell('IF(K'+row+'="","",IF(J'+row+'>0,J'+row+'*$H$2,K'+row+'*$J$2))',null),cell('IF(K'+row+'="","",MAX(K'+row+'*$L$2,$N$2))',null),cell('IF(OR(G'+row+'="",H'+row+'="",I'+row+'="",J'+row+'=""),"",K'+row+'-L'+row+'-H'+row+'-I'+row+'-M'+row+'-N'+row+'-O'+row+')',null),cell('IF(OR(K'+row+'="",P'+row+'=""),"",P'+row+'/K'+row+')',null),cell('IF(OR(P'+row+'="",F'+row+'=""),"",ROUND(P'+row+'*F'+row+',2))',null),p.marketSales,p.marketRevenue,p.bsr,tier,p.margin,p.fba,
 cell('IF(F'+row+'="","待填7天销量",IF(OR(G'+row+'="",H'+row+'="",I'+row+'="",J'+row+'=""),"待填成本参数",IF(AND(Q'+row+'>=$P$2,P'+row+'>=0,F'+row+'>=$R$2),"通过测款门槛",IF(P'+row+'<0,"亏损，停止扩量","未达门槛"))))',null),
 '参考：PP地垫周报8.9-8.15-PD17利润示例.xlsx',null,null,null,
 cell('IF(OR(Y'+row+'="待填7天销量",Y'+row+'="待填成本参数"),"待补数据",IF(OR(AA'+row+'="",AB'+row+'="",AC'+row+'=""),"待补CVR/TACOS/库存",IF(AND(V'+row+'="尾部51-100",AB'+row+'<=$T$2,AC'+row+'<=$X$2),"尾部→中部候选",IF(AND(V'+row+'="中部21-50",Q'+row+'>=0.15,AB'+row+'<=$T$2,AC'+row+'<=60),"中部→头部候选",IF(AND(V'+row+'="头部1-20",Q'+row+'>=0.05,AB'+row+'<=$V$2),"头部保留/扩量",IF(Q'+row+'<0,"亏损退出","继续观察"))))))',null)]);}
const profitTotal=profitSheetRows.length+1;
const profitTotalRow=Array(30).fill('');
profitTotalRow[0]='合计';
profitTotalRow[5]=cell('SUM(F5:F'+(profitTotal-1)+')',0);
profitTotalRow[17]=cell('SUM(R5:R'+(profitTotal-1)+')',0);
profitSheetRows.push(profitTotalRow);
profitSheetRows.push([]);profitSheetRows.push(['参考门槛与动作']);profitSheetRows.push(['Tail→Middle','连续4周 BSR≤100、CVR达到类目基准、TACOS≤15%、库存覆盖≤90天']);profitSheetRows.push(['Middle→Head','连续6周 BSR≤50、贡献毛利≥15%、自然单占比提升且可支撑60天补货']);profitSheetRows.push(['头部扩量','结算毛利率≥5%、TACOS≤12%；连续亏损14天降广告并缩占比2-3个百分点']);profitSheetRows.push(['退出','连续90天无法进入前100或资源占用高于增量利润则退出']);
const ws14=addSheet('14_利润测试',profitSheetRows,[15,34,12,16,14,10,13,10,10,10,12,12,10,10,10,12,10,14,14,16,10,14,12,12,16,42,10,10,14,22],[[0,0,0,29],[2,0,2,29],[profitTotal-1,0,profitTotal-1,29]]);
setAutofilter(ws14,'AD',profitTotal-1);

// Copy the planning team's operational header and actual/target rows, then
// add formula columns for H1 achievement and July execution status.
const planRefWb=XLSX.default.readFile(path.join(ROOT,'新增参考的材料和内容','销量预测计划部底表-户外地垫.xlsx'),{cellFormula:false,cellNF:false});
const planRefRows=XLSX.utils.sheet_to_json(planRefWb.Sheets['26年实际销量'],{header:1,defval:''});
const planHeaders=[...(planRefRows[0]||[]),'1-6月实际合计','H1达成率','7月计划差额','执行状态'];
const planSheetRows=[['计划部经营计划与实际周报底表'],['表头与经营字段参照：销量预测计划部底表-户外地垫.xlsx / 26年实际销量；新增列由本表公式计算。'],[],planHeaders];
for(const raw of planRefRows.slice(1)){if(!raw.some(v=>v!==''&&v!==null&&v!==undefined))continue;const row=planSheetRows.length+1;const sourceVals=raw.map(v=>typeof v==='string'&&/^#(DIV\/0!|VALUE!|N\/A|REF!)/.test(v)?null:v);const vals=[...sourceVals,cell('SUM(Q'+row+':V'+row+')',null),cell('IFERROR(AH'+row+'/P'+row+',"")',null),cell('IF(OR(W'+row+'="",X'+row+'=""),"",W'+row+'-X'+row+')',null),cell('IF(P'+row+'="","待补年度目标",IF(AI'+row+'>=1,"H1达成",IF(AND(X'+row+'<>"",W'+row+'=""),"待补7月实际","跟进")))',null)];planSheetRows.push(vals);}
const ws15=addSheet('15_经营计划与实际',planSheetRows,[14,12,20,12,14,20,15,28,12,20,12,16,14,12,12,14,16,16,16,16,16,16,16,12,12,12,12,12,12,14,14,14,34,16,12,16,16],[[0,0,0,36],[1,0,1,36]]);
setAutofilter(ws15,'AK',planSheetRows.length);

const checkRows=[['公式校验与数据完整性'],['校验项','结果','期望/阈值','实际差额','状态','说明'],['整体=PP+非PP（2026H1销量）',cell(annualRef('overall','sales',3)+'-('+annualRef('pp','sales',3)+'+'+annualRef('nonpp','sales',3)+')',0),0,cell('ABS(B3)',0),cell('IF(ABS(B3)<0.5,"通过","失败")','通过'),'整体市场必须等于两个互补分类之和。'],['整体=PP+非PP（2026H1销售额）',cell(annualRef('overall','revenue',3)+'-('+annualRef('pp','revenue',3)+'+'+annualRef('nonpp','revenue',3)+')',0),0,cell('ABS(B4)',0),cell('IF(ABS(B4)<0.5,"通过","失败")','通过'),'销售额同样校验。'],['Top100最大Listing数',cell('MAX(\'02_整体市场-BSR前100\'!$C$5:$C$'+refs.overall.top.monthlyEnd+',\'05_PP管-BSR前100\'!$C$5:$C$'+refs.pp.top.monthlyEnd+',\'08_非PP高客单-BSR前100\'!$C$5:$C$'+refs.nonpp.top.monthlyEnd+')',maxTop100),100,cell('B5-C5',maxTop100-100),cell('IF(B5<=C5,"通过","失败")','通过'),'每范围每月最多100个独立Listing。'],['2026.07展示父体数',cell('COUNTIFS(\'90_输入_明细\'!$A$3:$A$'+detailEnd+',"202607")',details.filter(r=>r.month==='202607').length),94,cell('B6-C6',details.filter(r=>r.month==='202607').length-94),cell('IF(B6=C6,"通过","请复核")','请复核'),'展示样本数量；不作为核心增速。'],['明细输入行数',cell('COUNTA(\'90_输入_明细\'!$A$3:$A$'+detailEnd+')',details.length),details.length,cell('B7-C7',0),cell('IF(B7=C7,"通过","失败")','通过'),'90页值区行数。'],['缺失销量记录数',cell('COUNTIF(\'90_输入_明细\'!$T$3:$T$'+detailEnd+',0)',details.filter(r=>r.sv===0).length),0,cell('B8-C8',details.filter(r=>r.sv===0).length),cell('IF(B8=C8,"通过","请复核")','请复核'),'缺失保留为空，不按0计入有效分母。'],['缺失销售额记录数',cell('COUNTIF(\'90_输入_明细\'!$U$3:$U$'+detailEnd+',0)',details.filter(r=>r.rv===0).length),0,cell('B9-C9',details.filter(r=>r.rv===0).length),cell('IF(B9=C9,"通过","请复核")','请复核'),'缺失保留为空，不按0计入有效分母。']];
const ws92=addSheet('92_校验',checkRows,[36,16,15,15,14,58],[[0,0,0,5]]);
const rulesRows=[['来源、数据操作和规则登记'],[],['项目','当前值/规则','来源文件/表','计算或操作','是否写入原始数据','可复核位置','限制/解释','备注'],
 ['主源数据','data/raw中的主源工作簿，按sheet_catalog识别月表','data/raw/地垫-卖家精灵市场数据.xlsx；market.db','只读读取market.db月表；未改写raw文件','否','90_输入_明细','历史月份原始表存在变体/父体语义差异','缺失值保留为空'],
 ['2026替换链','2026.01-07竞品替换快照','competitor_809440.db raw_YYYYMM','读取替换快照做父体/ASIN维度排名补充；不覆盖源文件','否（仅派生结果）','90/91页','2026.07为94父体展示样本','核心截止2026.06'],
 ['去重逻辑','父ASIN优先，缺省时用ASIN；每月组内选最小可解析BSR，若并列选指标更完整行，再选最小源行ID','本次生成脚本；与src/build_competitor_db.js规则对齐','所有月份结果先形成父ASIN/ASIN代表记录；原始行数、去重行数同时保留','否','90页Listing键、原始行数、去重行数','这是统计单元选择，不删除raw记录','结果可追溯源行'],
 ['BSR解析与Top100','完整数字/带逗号/允许N.0；多值取最小正整数；每个范围每月最多100','market.db与competitor_809440.db','按overall、PP、非PP、GENIMO分别形成独立Top100池；分层不跨池','否','90页BSR列；02/05/08、03/06/09页','GENIMO是品牌镜头，不参与PP+非PP加总',''],
 ['五档边界说明','原文标签写作1-5、5-10、10-20、20-50、50-100；为避免边界重复，结果表按不重叠区间1-5、6-10、11-20、21-50、51-100统计','原始要求文字 + 本次交付设计','各Listing只进入一个五档，五档销量/销售额可加总回勾Top100；若外部展示使用重叠标签，禁止直接相加','否','03/06/09细分层区块','这是展示边界解释，不改变BSR原始名次',''],
 ['分类定义','PP=标题含完整单词plastic；非PP=整体排除PP的补集；GENIMO=品牌字段等于GENIMO','原始标题/品牌字段 + 2026同父体变体补充','整体、PP、非PP为互补市场分区；GENIMO独立作为品牌份额与进退层视角','否','01/04/07、10页','非PP高客单不是价格阈值，按需求定义为排除PP后的产品线',''],
 ['销量/销售额/均价','销量和销售额分别汇总；平均标价=有效价格合计/有效价格数；加权成交均价=配对销售额/配对销量','原始月销量、月销售额、价格字段','公式页从91_聚合输入计算；缺失值不按0进入分母','否','01/02/03等结果页','平均标价与成交均价不能混称客单价',''],
 ['MOM与YOY','MOM=当前月 vs 上一自然月；YOY=当前月 vs 去年同月','原始需求与说明文件；本次明确拆列','每个范围、Top100、分层和GENIMO均分别提供MOM/YOY；2026.07留空增速','否','月度、BSR、分层、品牌份额页','首月缺基准时留空',''],
 ['年度与核心周期','2024全年、2025全年、2025H1、2026H1；2026.01-06为核心，2026.07只展示','原始需求、SPEC','年度表和H1表均用公式引用91页','否','各月度/Top100页年度区块','不静默延长核心窗口',''],
 ['GENIMO分母','主口径=GENIMO/整体市场；PP内份额只作辅助','原始需求中的品牌占比要求','10页同时给整体销量/销售额份额、PP内销量份额；12页计划按整体/PP/非PP分别观察','否','10/11/12页','不能把GENIMO当成第三个市场分类',''],
  ['参考表头对齐','13页采用参考 workbook 的 Monthly 前100趋势、Tier 年度分层和验收基准字段；14页沿用PD17利润示例字段顺序；15页复制计划部26年实际销量原表头并追加执行字段','新增参考的材料和内容/PP管数据-plastic-2025年-2026年.xlsx；PP塑料户外地垫_BSR前100年度分层与2027链接规划 (1).xlsx；销量预测计划部底表-户外地垫.xlsx；PD17利润示例','只调整结果工作簿的展示结构和派生列，不把参考 workbook 的数值当作market.db明细真值','否','13-15页','参考表头是交付结构基准；跨源数值仍按来源并列',''] ,
  ['利润测试公式','净成交价→采购成本$→进口税$→仓储$→佣金$→单件利润$→利润率→SKU利润$；单件利润不依赖7天销量，SKU利润依赖7天销量','参考：PP地垫周报8.9-8.15-PD17利润示例.xlsx','美元汇率、关税、进口税、仓储、佣金和门槛在14页参数区可编辑；缺输入保持空值并由状态/动作列提示','否','14页K:R、Y、AD','周销量、采购成本、头程、尾程、体积、CVR、TACOS、库存需由经营端补录',''],
  ['周报与经营计划','13页为管理周报摘要，14页为GENIMO/BSR Top100利润测款，15页为计划部实际/目标底表','13/14/15页及其来源参考文件','13页只汇总已生成的月度市场快照；15页对源错误字符串留空，不改写计划部原表','否','13-15页','周报日期、7天销量和经营参数均需填写后才会出现测款结论',''],
  ['本次实际数据操作','只读数据库，写入一份结果xlsx：90明细值、91聚合值、各子表公式和校验；未改写raw、未删除历史审计记录','本次脚本与输出文件','脚本预处理仅负责代表记录、排名和集合识别；可重算指标均落到Excel公式','否','90-93页','交付时以xlsx为主，公式可在Excel/兼容工具中复核','']];
const ws93=addSheet('93_来源与规则',rulesRows,[20,38,52,60,16,28,48,34],[[0,0,0,7],[1,0,1,7],[2,0,2,7]]);

// Apply a restrained, reference-like report style.  The previous workbook was
// formula-complete but visually plain; these styles make the repeated headers,
// editable inputs, source sections and warning fields readable in Excel.
const fillRgb=(rgb)=>({patternType:'solid',fgColor:{rgb}});
const titleStyle={fill:fillRgb('1F4E78'),font:{name:'Arial',sz:14,bold:true,color:{rgb:'FFFFFF'}},alignment:{vertical:'center',horizontal:'left',wrapText:true}};
const headerStyle={fill:fillRgb('5B9BD5'),font:{name:'Arial',sz:10,bold:true,color:{rgb:'FFFFFF'}},alignment:{vertical:'center',horizontal:'center',wrapText:true},border:{bottom:{style:'thin',color:{rgb:'D9E2F3'}}}};
const sectionStyle={fill:fillRgb('D9EAF7'),font:{name:'Arial',sz:10,bold:true,color:{rgb:'17365D'}},alignment:{vertical:'center',wrapText:true}};
const noteStyle={fill:fillRgb('FFF2CC'),font:{name:'Arial',sz:10,italic:true,color:{rgb:'7F6000'}},alignment:{vertical:'center',wrapText:true}};
const inputStyle={fill:fillRgb('FFF2CC'),font:{name:'Arial',sz:10,color:{rgb:'7F6000'}},alignment:{vertical:'center'}};
const bodyStyle={font:{name:'Arial',sz:10,color:{rgb:'1F1F1F'}},alignment:{vertical:'center',wrapText:false}};
function styleRow(ws,r,style){const ref=XLSX.utils.decode_range(ws['!ref']||'A1:A1');for(let c=ref.s.c;c<=ref.e.c;c++){const a=XLSX.utils.encode_cell({r,c});if(ws[a])ws[a].s={...(ws[a].s||{}),...style};}}
function styleReportSheet(ws,name){if(!ws['!ref'])return;const ref=XLSX.utils.decode_range(ws['!ref']);for(let r=ref.s.r;r<=ref.e.r;r++)for(let c=ref.s.c;c<=ref.e.c;c++){const a=XLSX.utils.encode_cell({r,c});if(ws[a])ws[a].s={...(ws[a].s||{}),...bodyStyle};}
 styleRow(ws,0,titleStyle);if(ws['!rows']===undefined)ws['!rows']=[];ws['!rows'][0]={hpt:28};
 for(let r=1;r<=ref.e.r;r++){const a=XLSX.utils.encode_cell({r,c:0}),v=String(ws[a]?.v??'');if(v.includes('Month 月份')||v.includes('月份')&&['月份','月度','Tier 层级','校验项','项目','ASIN','指标','输入项','范围'].some(x=>v.includes(x))||v==='部门'){styleRow(ws,r,headerStyle);ws['!rows'][r]={hpt:30};}else if(v.includes('参考表头')||v.includes('领导验收主基准')||v.includes('核心市场指标')||v.includes('年度分层对照')||v.includes('GENIMO品牌与进退层')||v.includes('利润测试摘要')){styleRow(ws,r,sectionStyle);ws['!rows'][r]={hpt:22};}else if(r===1||v.includes('说明')||v.includes('口径：')||v.includes('公式按参考')){styleRow(ws,r,noteStyle);}}
}
for(const name of wb.SheetNames)styleReportSheet(wb.Sheets[name],name);
// Highlight editable cost/weekly-operation inputs in the profit test.
for(const c of ['D','E','F','G','H','I','J','AA','AB','AC'])for(let r=5;r<=profitTotal-1;r++){const a=c+r;if(ws14[a])ws14[a].s={...(ws14[a].s||{}),...inputStyle};}
// Clearer number formats for the report outputs and the reference formulas.
const setFmt=(ws,col,start,end,z)=>{for(let r=start;r<=end;r++){const a=col+r;if(ws[a])ws[a].z=z;}};
for(const name of ['01_整体市场-月度年度','04_PP管-月度年度','07_非PP高客单-月度年度']){const ws=wb.Sheets[name];setFmt(ws,'F',5,ws['!ref']?XLSX.utils.decode_range(ws['!ref']).e.r+1:5,'#,##0');setFmt(ws,'I',5,40,'$#,##0');setFmt(ws,'L',5,40,'$#,##0.00');setFmt(ws,'M',5,40,'$#,##0.00');for(const c of ['O','P','Q','R','S','T','U'])setFmt(ws,c,5,40,'0.0%');}
for(const name of ['02_整体市场-BSR前100','05_PP管-BSR前100','08_非PP高客单-BSR前100']){const ws=wb.Sheets[name];for(const c of ['F','H','T','U','V'])setFmt(ws,c,5,ws['!ref']?XLSX.utils.decode_range(ws['!ref']).e.r+1:5,'#,##0');for(const c of ['J','K'])setFmt(ws,c,5,40,'$#,##0.00');for(const c of ['M','N','O','P','Q','R'])setFmt(ws,c,5,40,'0.0%');}
for(const c of ['C','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','W','X'])setFmt(ws14,c,5,profitTotal,'$#,##0.00');setFmt(ws14,'F',5,profitTotal,'#,##0');setFmt(ws14,'Q',5,profitTotal,'0.0%');setFmt(ws14,'AA',5,profitTotal,'0.0%');setFmt(ws14,'AB',5,profitTotal,'0.0%');
ws90['!autofilter']={ref:'A2:AF'+detailRows90.length};ws91['!autofilter']={ref:'A2:R'+aggRows91.length};

await fs.mkdir(OUT_DIR,{recursive:true});
wb.SheetNames=['00_总览与结论','01_整体市场-月度年度','02_整体市场-BSR前100','03_整体市场-BSR分层','04_PP管-月度年度','05_PP管-BSR前100','06_PP管-BSR分层','07_非PP高客单-月度年度','08_非PP高客单-BSR前100','09_非PP高客单-BSR分层','10_GENIMO-品牌份额','11_GENIMO-进退层','12_GENIMO-2027规划','13_周报汇总','14_利润测试','15_经营计划与实际','90_输入_明细','91_聚合输入','92_校验','93_来源与规则'];
wb.Workbook={CalcPr:{calcMode:'auto',fullCalcOnLoad:true,forceFullCalc:true}};
XLSX.writeFile(wb,OUT,{bookType:'xlsx',compression:true,cellStyles:true});
console.log(JSON.stringify({outputPath:OUT,sourceMonths:months,displayMonths:display,detailRows:details.length,aggregateRows:aggs.length,cohort:{retained:retained.length,exited:exited.length,entered:entered.length},core:{overall2025H1:sumCache('sales','overall','全部','202501','202506'),overall2026H1:sumCache('sales','overall','全部','202601','202606'),pp2025H1:sumCache('sales','pp','全部','202501','202506'),pp2026H1:sumCache('sales','pp','全部','202601','202606'),nonpp2025H1:sumCache('sales','nonpp','全部','202501','202506'),nonpp2026H1:sumCache('sales','nonpp','全部','202601','202606')}},null,2));
db.close();cdb.close();
