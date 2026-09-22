const reportData=JSON.parse(document.getElementById('report-data').textContent);
window.reportData=reportData;
const charts=JSON.parse(document.getElementById('chart-data').textContent);
const escapeCell=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const num=v=>v==null?'—':Number(v).toLocaleString('en-US',{maximumFractionDigits:2});
const pc=v=>v==null?'—':(v*100).toFixed(1)+'%';
const scopeFilter=(p,key)=>key==='overall'||key==='pp'&&p.pp||key==='high'&&!p.pp||key==='genimo'&&p.genimo;
const statusName={CONSISTENT:'源值一致',MODE:'多值，取多数值',MEDIAN_CONFLICT:'无多数，取中位数',NO_VALUE:'无销量',FULL:'候选行完整',PARTIAL:'部分子体',NONE:'无有效子体'};
function drawCharts(){
  for(const config of charts){
    const year=document.getElementById(config.scope+'-year').value;
    const copy={...config,rows:config.rows.filter(r=>year==='all'||r.month.startsWith(year))};
    document.getElementById(config.id).innerHTML=chartMarkup(copy);
  }
}
function drawParents(scope){
  const month=document.getElementById(scope+'-month').value,search=document.getElementById(scope+'-search').value.trim().toUpperCase();
  const rows=reportData.parentRows.filter(p=>p.month===month&&scopeFilter(p,scope)&&(!search||p.parentKey.toUpperCase().includes(search)||p.asins.some(a=>a.toUpperCase().includes(search))));
  const table=document.getElementById(scope+'-parents');
  table.querySelector('tbody').innerHTML=rows.length?rows.map(p=>'<tr>'+[p.parent||p.parentKey,num(p.rankMedian),num(p.parentSales),statusName[p.salesStatus],num(p.salesMin)+' — '+num(p.salesMax),num(p.childSales),num(p.childRevenue),pc(p.childCoverage),num(p.childAsp),statusName[p.childStatus],pc(p.ppRatio),p.ppStatus==='TIE'?'平局暂归PP':p.ppRatio>0&&p.ppRatio<1?'标题混合':''].map(v=>'<td>'+escapeCell(v)+'</td>').join('')+'</tr>').join(''):'<tr><td colspan="12">该条件下没有父体记录</td></tr>';
  document.getElementById(scope+'-parent-count').textContent=rows.length+' 个父体';
}
for(const scope of ['overall','pp','high','genimo']){
  document.getElementById(scope+'-month').addEventListener('change',()=>drawParents(scope));
  document.getElementById(scope+'-search').addEventListener('input',()=>drawParents(scope));
  document.getElementById(scope+'-year').addEventListener('change',function(){
    document.querySelectorAll('#'+scope+' [data-month]').forEach(row=>row.hidden=this.value!=='all'&&!row.dataset.month.startsWith(this.value));
    drawCharts();
  });
  drawParents(scope);
}
document.querySelectorAll('[data-series]').forEach(button=>button.addEventListener('click',()=>{
  const config=charts.find(c=>c.id===button.dataset.chart),series=config.series[Number(button.dataset.series)];series.hidden=!series.hidden;
  button.setAttribute('aria-pressed',String(!series.hidden));drawCharts();
}));
document.addEventListener('pointerover',showPoint);document.addEventListener('focusin',showPoint);
function showPoint(event){const point=event.target.closest('.chart-hit');if(!point)return;point.closest('.chart-card').querySelector('.chart-readout').textContent=point.dataset.tooltip;}
document.addEventListener('keydown',event=>{if(event.target.matches('.chart-hit')&&['ArrowLeft','ArrowRight'].includes(event.key)){const points=[...event.target.closest('svg').querySelectorAll('.chart-hit')];const index=points.indexOf(event.target)+(event.key==='ArrowRight'?1:-1);points[index]?.focus();event.preventDefault();}});
document.getElementById('theme-toggle').addEventListener('click',()=>{const root=document.documentElement;root.dataset.theme=root.dataset.theme==='dark'?'light':'dark';});
document.getElementById('collapse-all').addEventListener('click',()=>document.querySelectorAll('main details.subsection').forEach(d=>d.open=false));
document.getElementById('expand-all').addEventListener('click',()=>document.querySelectorAll('main details.subsection').forEach(d=>d.open=true));
document.querySelectorAll('.sidebar a').forEach(a=>a.addEventListener('click',()=>{const target=document.querySelector(a.getAttribute('href'));if(target?.matches('details'))target.open=true;document.querySelectorAll('.sidebar a').forEach(x=>x.classList.toggle('is-active',x===a));}));
