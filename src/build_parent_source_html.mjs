import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import {chartMarkup} from './parent_chart.mjs';
const ROOT=process.cwd(),OUT=path.join(ROOT,'outputs','20260920-new-source-parent-model');
const model=JSON.parse(fs.readFileSync(path.join(OUT,'model.json'),'utf8'));
const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const n=v=>v==null?'—':Number(v).toLocaleString('en-US',{maximumFractionDigits:2});
const pct=v=>v==null?'—':(v*100).toFixed(1)+'%';
const month=v=>v.slice(0,4)+'.'+v.slice(4);
const bullets=items=>'<ul class="analysis-list">'+items.map(i=>'<li>'+i+'</li>').join('')+'</ul>';
const sub=(id,title,body,open=true)=>`<details id="${id}" class="subsection" ${open?'open':''}><summary>${esc(title)}</summary><div class="subsection-body">${body}</div></details>`;
const table=(heads,rows,id='',months=[])=>`<div class="table-wrap"><table ${id?`id="${id}"`:''}><thead><tr>${heads.map(h=>'<th>'+esc(h)+'</th>').join('')}</tr></thead><tbody>${rows.map((r,i)=>`<tr ${months[i]?`data-month="${months[i]}"`:''}>${r.map(v=>'<td>'+esc(v)+'</td>').join('')}</tr>`).join('')}</tbody></table></div>`;
const names={overall:'整体市场',pp:'PP市场',high:'高客单价市场',genimo:'Genimo品牌',genimoPP:'Genimo PP市场'};
const state={NO_BASE:'无同比基期',MATCHED_MONTHS:'仅共同月份，非全年',FULL_YEAR:'完整年度'};
const charts=[];
function chart(scope,suffix,title,rows,series,unit,type='line'){
  const config={id:scope+'-chart-'+suffix,scope,title,rows,series,unit,type};charts.push(config);
  return `<div class="chart-card"><div class="chart-heading"><b>${esc(title)}</b><span>${esc(unit)}</span></div><div class="chart-legend">${series.map((s,i)=>`<button data-chart="${config.id}" data-series="${i}" aria-pressed="true" style="--color:${s.color}">${esc(s.label)}</button>`).join('')}</div><div id="${config.id}" class="chart-svg">${chartMarkup(config)}</div><div class="chart-readout" aria-live="polite">悬停或用 Tab 聚焦月份，方向键切换；点击图例显示或隐藏曲线</div></div>`;
}
function annualTable(key){return table(['年份','本期共同月份','基期共同月份','本期父体销量估计','基期父体销量估计','本期有效子体销售额($)','基期有效子体销售额($)','销量YOY','销售额YOY','本期覆盖率','基期覆盖率','期间状态','数据范围提示'],model.annual[key].map(r=>[r.year,r.months.map(month).join(', ')||'—',r.priorMonths.map(month).join(', ')||'—',n(r.parentSales),n(r.priorSales),n(r.childRevenue),n(r.priorRevenue),pct(r.yoyParentSales),pct(r.yoyChildRevenue),pct(r.childCoverage),pct(r.priorCoverage),state[r.status],r.scopeLimited?'共同月份含类目排名未对应行，结果低估':'Outdoor Rugs范围完整']));}
function narrative(key){
  const rows=model.summaries[key],last=rows.at(-1),prior=rows.at(-2),year=rows.at(-13);
  const best=model.tiers[key].filter(b=>b.type==='coarse').map(b=>({name:b.name,r:b.rows.at(-1)})).sort((a,b)=>(b.r.parentSales??0)-(a.r.parentSales??0))[0];
  return bullets([
    `${month(last.month)}，${names[key]}纳入 <b>${n(last.parentCount)} 个父体</b>、${n(last.candidateRows)} 条候选子体行。父体月销量估计 <b>${n(last.parentSales)}</b>，有效子体销售额 $${n(last.childRevenue)}。`,
    `销量相对上月 ${pct(last.momParentSales)}，相对去年同月 ${pct(last.yoyParentSales)}。其中 ${n(last.conflictParents)} 个父体有多个源销量值，${n(last.medianParents)} 个使用中位数，${n(last.missingSalesParents)} 个缺少销量。`,
    `有效子体销售额相对上月 ${pct(last.momChildRevenue)}，相对去年同月 ${pct(last.yoyChildRevenue)}。候选行覆盖率由上月 ${pct(prior.childCoverage)}、去年同月 ${pct(year.childCoverage)} 变为本月 ${pct(last.childCoverage)}，金额变化同时包含数据覆盖变化。`,
    `有效子体加权均价 $${n(last.childAsp)}，使用同批有效子体销售额÷有效子体销量。父体估计销量与有效子体销售额覆盖范围不同，不能互除推算市场成交均价。`,
    `按父体BSR中位数分层，${best.name}的估计销量最高，为 ${n(best.r.parentSales)}，有效子体覆盖率为 ${pct(best.r.childCoverage)}。该结果描述当前候选父体，不能直接当作未来测试转化率。`,
    '扩大测试前，先回查多值父体和混合标题，再连续跟踪相同口径的销量、份额与覆盖率。'
  ]);
}
const share=(x,y)=>x!=null&&y>0?x/y:null;
function shares(){return table(['月份','Genimo销量整体份额','Genimo销量PP份额','Genimo有效销售额整体份额','Genimo有效销售额PP份额','Genimo覆盖率','整体覆盖率'],model.months.map((m,i)=>{const g=model.summaries.genimo[i],gp=model.summaries.genimoPP[i],a=model.summaries.overall[i],p=model.summaries.pp[i];return [month(m),pct(share(g.parentSales,a.parentSales)),pct(share(gp.parentSales,p.parentSales)),pct(share(g.childRevenue,a.childRevenue)),pct(share(gp.childRevenue,p.childRevenue)),pct(g.childCoverage),pct(a.childCoverage)];}),'',model.months);}
function strategy(){
  const g=model.summaries.genimo.at(-1),a=model.summaries.overall.at(-1),gp=model.summaries.genimoPP.at(-1),p=model.summaries.pp.at(-1),oldg=model.summaries.genimo.at(-13),olda=model.summaries.overall.at(-13);
  return bullets([
    `<b>份额维护：</b>2026.07 Genimo父体估计销量整体份额 ${pct(share(g.parentSales,a.parentSales))}，去年同月 ${pct(share(oldg.parentSales,olda.parentSales))}；PP内份额 ${pct(share(gp.parentSales,p.parentSales))}。将这些份额作为2027复核基线，每月按相同父体范围检查。`,
    `<b>测试方向：</b>PP父体销量YOY ${pct(p.yoyParentSales)}，高客单价市场 ${pct(model.summaries.high.at(-1).yoyParentSales)}。优先复核销量走势较强、源值冲突较少的档位，小批测试后连续两个月检查销量、份额和覆盖率，再决定是否扩量。`,
    '<b>链接数量：</b>数据不含品牌资源、测试转化和预算，无法给出可靠的固定头中尾链接数量。按五档表现形成测试优先级，不将子体数量当成父体链接数。',
    '<b>花型和尺寸：</b>源SKU与标题保留在Excel供人工抽查；尚无完整验证过的分类，不输出未经验证的花型优劣结论。'
  ]);
}
const labels=['月度汇总与可回勾数据','父体明细与人工核验','共同月份年度YOY','月度趋势与变化','BSR头中尾与五档','覆盖率与源值冲突','数据分析与建议'];
const nav=['overall','pp','high','genimo'].map((k,i)=>`<details class="nav-group" open><summary>第${i+1}部分 · ${names[k]}</summary><nav>${labels.map((x,j)=>`<a href="#${k}-${j+1}">${i+1}.${j+1} ${x}</a>`).join('')}</nav></details>`).join('');
const baseCss=fs.readFileSync(path.join(ROOT,'src','original_ui.css'),'utf8');
const extraCss=`.topbar h1{font-size:clamp(30px,3vw,46px)}.market-section h2,.overview-section h2{font-size:32px}.nav-group{padding:8px 0;border-bottom:1px solid var(--line)}.nav-group summary{font-size:14px;color:var(--text);cursor:pointer}.sidebar .nav-group a{padding:5px 8px;font-size:12px}.subsection{margin:20px 0;border-top:1px solid var(--line);scroll-margin-top:20px}.subsection summary{font-size:19px;padding:14px 0;cursor:pointer}.subsection-body{padding:0 2px 16px}.analysis-list{padding-left:22px}.analysis-list li{margin:8px 0}.lead,.hint{color:var(--muted)}.hint{font-size:14px}.controls{display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin:12px 0}.controls select,.controls input,.controls button,.download,.chart-legend button{font:inherit;font-size:14px;padding:8px 12px;color:var(--text);background:var(--surface);border:1px solid var(--line-strong);border-radius:6px}.controls input{min-width:240px}.download{display:inline-block;text-decoration:none}.table-wrap{max-height:560px;overflow:auto;border:1px solid var(--line);margin-top:12px}.table-wrap th{position:sticky;top:0;z-index:2;background:var(--surface);white-space:normal;min-width:110px}.table-wrap td{white-space:nowrap;font-size:14px;font-variant-numeric:tabular-nums}.table-wrap tr:hover td{background:var(--surface-strong)}.chart-card{margin:20px 0;padding:16px;border:1px solid var(--line)}.chart-heading{display:flex;justify-content:space-between;gap:12px}.chart-heading span{color:var(--muted)}.chart-svg{overflow:auto}.chart-svg svg{display:block;width:100%;min-width:620px}.chart-svg text{fill:var(--muted);font:12px Arial}.chart-svg .grid{stroke:var(--line)}.chart-hit:focus,.chart-hit:hover{fill:rgba(148,163,184,.12);outline:none}.chart-readout{min-height:40px;font-size:13px;color:var(--muted);padding:8px 0}.chart-legend{display:flex;flex-wrap:wrap;gap:8px;margin:12px 0}.chart-legend button{cursor:pointer;font-size:12px}.chart-legend button:before{content:'';display:inline-block;width:9px;height:9px;background:var(--color);margin-right:6px}.chart-legend button[aria-pressed=false]{opacity:.4}.metrics-grid{grid-template-columns:repeat(4,minmax(0,1fr))}.metric-card .metric-value{font-size:30px}@media(max-width:850px){.sidebar{position:relative;width:100%;max-height:290px}.workspace{margin-left:0}.metrics-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.topbar{padding:22px;flex-wrap:wrap}.content{padding:18px}.top-actions{flex-wrap:wrap}}`;
let html=`<!doctype html><html lang="zh-CN" data-theme="dark"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>户外地垫市场分析 · SPEC 3.1</title><style>${baseCss}\n${extraCss}</style></head><body><div class="app-shell"><aside class="sidebar"><div class="brand"><div class="brand-mark"></div><div><strong>Market Intelligence</strong><small>户外地垫 · 父体口径</small></div></div><nav><a href="#dashboard">数据总览</a>${nav}</nav><div class="sidebar-footer">2024.08 — 2026.07<br>24个月 · SPEC 3.1</div></aside><div class="workspace"><header class="topbar"><div><span class="eyebrow">MARKET ANALYSIS · SPEC 3.1</span><h1>户外地垫市场分析</h1><p>父体月销量估计 · 有效子体销售额 · 可回勾数据</p></div><div class="top-actions"><a class="download" href="./户外地垫市场分析-SPEC3-父体口径可回勾.xlsx" download>下载公式Excel</a><button id="theme-toggle" class="theme-button" aria-label="切换主题">☀</button></div></header><main class="content"><section id="dashboard" class="overview-section"><h2>数据总览</h2><p class="lead">以BSR前100候选父体为样本。父体估计销量与已披露子体销售额分别统计。</p><div class="metrics-grid">${[['源数据月份',24],['BSR候选子体行',model.metadata.candidateRowCount],['父体月份组合',model.metadata.parentMonthCount],['源销量多值组合',model.metadata.conflictParentCount]].map(([k,v])=>`<article class="metric-card"><span class="metric-label">${k}</span><strong class="metric-value">${n(v)}</strong></article>`).join('')}</div><div class="controls"><button id="expand-all">展开全部小节</button><button id="collapse-all">收起全部小节</button></div>`;
html+=sub('dashboard-1','0.1 指标和计算范围',bullets([
  '唯一主源为24个按月文件，2024.08—2026.07。小类名与排名按原始换行逐项对应，只取Outdoor Rugs的BSR 1—100（含100），再按月份+父ASIN聚合。',
  '<b>已知源限制：</b>2025.08有354行仅一个小类名称却有两个BSR，无法可靠对应Outdoor Rugs排名，保留在原始页并从主统计排除；涉及该月的比较均标记范围受限。',
  '<b>父体月销量估计：</b>源值全部一致取一致值；严格多数超过50%取众数；否则取中位数。所有多值组合均标记，50%平局不任意选较小值。',
  '<b>有效子体销售额：</b>只累加同时有有效子体销量和子体销售额的行。没有有效行时留空，明确为0的源值保留。',
  '<b>候选范围：</b>父体销量估计涉及该父体全部变体；子体销售额只涵盖已采集、BSR合格且两字段有效的子体。FULL只代表候选行字段完整。',
  '<b>分类：</b>标题整词plastic命中率≥50%暂归PP，其余归高客单价；混合与平局列入核查。高客单价是补集名称，不保证实际价格高于PP。',
  '<b>比较：</b>MOM为相对上月，YOY为相对去年同月；年度只比较共同月份。2025年比8—12月，2026年比1—7月。',
  '<b>限制：</b>不将缺失填0，不累加原始父体月销售额，不计算利润。'
]));
html+=sub('dashboard-2','0.2 来源与质量',table(['月份','原始行','BSR候选行','父体数','有效子体行','子体覆盖率','多值父体','无多数取中位数'],model.summaries.overall.map((r,i)=>[month(r.month),model.sourceFiles[i].rawRows,r.candidateRows,r.parentCount,r.validChildRows,pct(r.childCoverage),r.conflictParents,r.medianParents]))+`<p class="hint">批次 ${esc(model.metadata.batchId)}。源文件SHA-256见Excel的93_来源登记。覆盖率是有效行/候选行，不是完整市场的覆盖率。</p>`);
const sample=model.parentRows.find(p=>p.month==='202607'&&p.parentKey==='B0BM74444Q');
html+=sub('dashboard-3','0.3 人工验算示例',table(['月份','父ASIN','候选行','众数','出现次数','占比','有效子体行','有效子体销量','有效子体销售额($)','子体均价($)'],[[month(sample.month),sample.parentKey,sample.candidateRows,sample.salesMode,sample.salesModeCount,pct(sample.salesModeShare),sample.childValidRows,sample.childSales,sample.childRevenue,n(sample.childAsp)]])+bullets(['在91页筛选月份202607、父体键B0BM74444Q，查看月销量分布。','在92页查看众数、覆盖率、取值状态和有效子体SUMIFS。','加权均价=154,430÷3,100=$49.8161…；不是子体销量÷有效行数。','Excel修改后只联动工作簿公式；网页与JSON需重新构建才能更新。']))+'</section>';
for(const [idx,key] of ['overall','pp','high','genimo'].entries()){
  const p=idx+1,rows=model.summaries[key];
  html+=`<section id="${key}" class="market-section"><div class="section-kicker">第${p}部分</div><h2>${names[key]}</h2><div class="controls"><label>表格和图表年份 <select id="${key}-year"><option value="all">全部24个月</option>${['2024','2025','2026'].map(y=>`<option>${y}</option>`).join('')}</select></label></div>`;
  html+=sub(`${key}-1`,`${p}.1 月度汇总与可回勾数据`,table(['月份','父体数','父体月销量估计','有效子体销量','有效子体销售额($)','子体均价($)','子体覆盖率','销量MOM（上月）','销售额MOM（上月）','销量YOY（去年同月）','销售额YOY（去年同月）','多值父体','取中位数父体','无销量父体','范围提示'],rows.map(r=>[month(r.month),r.parentCount,n(r.parentSales),n(r.childSales),n(r.childRevenue),n(r.childAsp),pct(r.childCoverage),pct(r.momParentSales),pct(r.momChildRevenue),pct(r.yoyParentSales),pct(r.yoyChildRevenue),r.conflictParents,r.medianParents,r.missingSalesParents,r.ambiguousBsrRows?`${r.ambiguousBsrRows}行类目排名未对应`:'完整']),'',model.months));
  html+=sub(`${key}-2`,`${p}.2 父体明细与人工核验`,`<div class="controls"><label>月份 <select id="${key}-month">${model.months.map(m=>`<option value="${m}" ${m===model.metadata.lastMonth?'selected':''}>${month(m)}</option>`).join('')}</select></label><input id="${key}-search" placeholder="输入父ASIN或子体ASIN" aria-label="${names[key]} ASIN检索"><span id="${key}-parent-count"></span></div>`+table(['父ASIN','BSR中位数','父体销量估计','取值方法','源销量最小—最大','有效子体销量','有效子体销售额($)','子体覆盖率','子体均价($)','金额状态','PP命中率','分类提示'],[],key+'-parents')+'<p class="hint">源销量最小—最大用于查看估计敏感度，并非统计置信区间。完整源行与公式请下载Excel。</p>');
  html+=sub(`${key}-3`,`${p}.3 共同月份年度YOY`,annualTable(key));
  html+=sub(`${key}-4`,`${p}.4 月度趋势与变化`,[
    chart(key,'sales','父体月销量估计',rows,[{key:'parentSales',label:'父体销量估计',color:'#38bdf8'}],'件','bar'),
    chart(key,'revenue','有效子体销售额',rows,[{key:'childRevenue',label:'有效子体销售额',color:'#e4ad58'}],'USD','bar'),
    chart(key,'mom','月度变化（相对上月）',rows,[{key:'momParentSales',label:'销量MOM',color:'#38bdf8'},{key:'momChildRevenue',label:'有效销售额MOM',color:'#e4ad58'}],'%'),
    chart(key,'price','有效子体加权均价',rows,[{key:'childAsp',label:'有效子体均价',color:'#a7a5ff'}],'USD')
  ].join(''));
  const fine=model.tiers[key].filter(b=>b.type==='fine'),colors=['#38bdf8','#6ee7b7','#a7a5ff','#e4ad58','#fda4af'];
  const tierChartRows=model.months.map((m,i)=>({month:m,...Object.fromEntries(fine.map(b=>[b.key,b.rows[i].momParentSales]))}));
  const trs=model.tiers[key].flatMap(b=>b.rows.map(r=>({band:b.name,...r})));
  html+=sub(`${key}-5`,`${p}.5 BSR头中尾与五档`,chart(key,'tiers','五档父体销量MOM',tierChartRows,fine.map((b,i)=>({key:b.key,label:b.name,color:colors[i]})),'%')+table(['分层','月份','父体数','父体销量估计','有效子体销售额($)','子体覆盖率','销量MOM','销量YOY'],trs.map(r=>[r.band,month(r.month),r.parentCount,n(r.parentSales),n(r.childRevenue),pct(r.childCoverage),pct(r.momParentSales),pct(r.yoyParentSales)]),'',trs.map(r=>r.month))+'<p class="hint">父体BSR使用候选子体中位数；小数边界连续，例如20.5属于中部。三档与五档分别回加整体，不可跨两套分层重复累加。</p>',false);
  html+=sub(`${key}-6`,`${p}.6 覆盖率与源值冲突`,chart(key,'coverage','有效子体覆盖率与父体多值率',rows,[{key:'childCoverage',label:'有效子体覆盖率',color:'#6ee7b7'},{key:'conflictRate',label:'父体多值率',color:'#fda4af'}],'%')+table(['月份','候选子体行','有效子体行','子体覆盖率','销售额状态','源销量多值父体','无多数取中位数','无销量父体'],rows.map(r=>[month(r.month),r.candidateRows,r.validChildRows,pct(r.childCoverage),r.revenueStatus,r.conflictParents,r.medianParents,r.missingSalesParents]),'',model.months),false);
  html+=sub(`${key}-7`,`${p}.7 数据分析与建议`,narrative(key)+(key==='genimo'?'<h3>Genimo在整体与PP市场的表现</h3>'+shares()+'<h3>2027年战略建议</h3>'+strategy():''))+'</section>';
}
const publicData={metadata:model.metadata,months:model.months,sourceFiles:model.sourceFiles,sheetStats:model.sheetStats,summaries:model.summaries,annual:model.annual,tiers:model.tiers,parentRows:model.parentRows};
const encode=data=>JSON.stringify(data).replace(/</g,'\\u003c');
html+=`<footer>SPEC 3.1 · ${esc(model.metadata.batchId)} · Excel、网页与JSON来自同一批次。结果反映卖家精灵源数据和明确的清洗估计，不代表已验证的实际成交。</footer></main></div></div><script id="report-data" type="application/json">${encode(publicData)}</script><script id="chart-data" type="application/json">${encode(charts)}</script><script>${chartMarkup.toString()}\n${fs.readFileSync(path.join(ROOT,'src','parent_report_client.js'),'utf8')}</script></body></html>`;
const base=path.join(OUT,'户外地垫市场分析-SPEC3-父体口径');
fs.writeFileSync(base+'.html',html);fs.writeFileSync(base+'.json',JSON.stringify(publicData,null,2));
fs.writeFileSync(base+'.md',`# 户外地垫市场分析 SPEC 3.1\n\n批次：${model.metadata.batchId}\n\n${['overall','pp','high','genimo'].map(k=>'## '+names[k]+'\n\n'+narrative(k).replace(/<li>/g,'\n- ').replace(/<[^>]+>/g,'')).join('\n\n')}\n\n## Genimo 2027建议\n${strategy().replace(/<li>/g,'\n- ').replace(/<[^>]+>/g,'')}\n`);
const files=[base+'.html',base+'.json',base+'.md',path.join(OUT,'户外地垫市场分析-SPEC3-父体口径可回勾.xlsx')];
fs.writeFileSync(path.join(OUT,'构建清单.json'),JSON.stringify({...model.metadata,files:files.map(file=>({path:path.relative(ROOT,file),sha256:crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex'),bytes:fs.statSync(file).size}))},null,2));
console.log(JSON.stringify({status:'BUILT',batchId:model.metadata.batchId,charts:charts.length,htmlBytes:Buffer.byteLength(html)}));
