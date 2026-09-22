// Shared SVG renderer: the browser and exported HTML use the same units and missing-value rules.
export function chartMarkup(config) {
  const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const format=v=>v==null?'无数据':config.unit==='%'?`${(v*100).toFixed(1)}%`:(config.unit==='USD'?'$':'')+Number(v).toLocaleString('en-US',{maximumFractionDigits:2});
  const rows=config.rows,series=config.series.filter(s=>!s.hidden), W=960,H=300,L=90,R=22,T=22,B=48;
  const vals=rows.flatMap(r=>series.map(s=>r[s.key])).filter(Number.isFinite);
  let min=Math.min(0,...vals),max=Math.max(0,...vals);if(max===min)max=min+1;
  const y=v=>T+(max-v)/(max-min)*(H-T-B), step=(W-L-R)/Math.max(rows.length,1),x=i=>L+step*(i+.5);
  let svg=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(config.title)}">`;
  for(let i=0;i<=4;i++){const v=min+(max-min)*i/4;svg+=`<line x1="${L}" x2="${W-R}" y1="${y(v)}" y2="${y(v)}" class="grid"/><text x="${L-10}" y="${y(v)+4}" text-anchor="end">${esc(format(v))}</text>`;}
  for(const s of series){
    if(config.type==='bar') rows.forEach((r,i)=>{const v=r[s.key];if(!Number.isFinite(v))return;svg+=`<rect fill="${s.color}" x="${x(i)-step*.32}" y="${Math.min(y(v),y(0))}" width="${step*.64}" height="${Math.max(1,Math.abs(y(v)-y(0)))}" rx="2"/>`;});
    else {let d='';rows.forEach((r,i)=>{const v=r[s.key];if(!Number.isFinite(v)){d+=' ';return;}const prior=i>0&&Number.isFinite(rows[i-1][s.key]);d+=`${prior?'L':'M'}${x(i)},${y(v)} `;});svg+=`<path d="${d}" fill="none" stroke="${s.color}" stroke-width="2.5"/>`;}
  }
  rows.forEach((r,i)=>{
    if(i%Math.max(1,Math.ceil(rows.length/8))===0||i===rows.length-1)svg+=`<text x="${x(i)}" y="${H-17}" text-anchor="middle">${r.month.slice(0,4)}.${r.month.slice(4)}</text>`;
    const tip=[r.month.slice(0,4)+'.'+r.month.slice(4),...series.map(s=>s.label+': '+format(r[s.key])),r.childCoverage!=null?'有效子体覆盖率: '+(r.childCoverage*100).toFixed(1)+'%':''].filter(Boolean).join(' · ');
    svg+=`<rect class="chart-hit" x="${L+i*step}" y="${T}" width="${step}" height="${H-T-B}" fill="transparent" tabindex="0" data-tooltip="${esc(tip)}" aria-label="${esc(tip)}"><title>${esc(tip)}</title></rect>`;
  });
  return svg+'</svg>';
}
