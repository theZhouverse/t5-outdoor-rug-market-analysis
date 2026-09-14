"""Prepare an auditable formula plan; XLSX serialization is owned by artifact-tool."""
import json, math, os, collections
from pathlib import Path
ROOT=Path(__file__).resolve().parent.parent
d=json.loads((ROOT/'tmp/formula_market_builder/spec2_dataset.json').read_text(encoding='utf-8'))
raw=d['raw']; detail=d['dedup']; months=d['months']; cats=d['categories']; bands=d['bands']
out=open(ROOT/'tmp/formula_market_builder/workbook_plan.jsonl','w',encoding='utf-8')
def emit(x): out.write(json.dumps(x,ensure_ascii=False,separators=(',',':'))+'\n')
def col(n):
 s=''
 while n: n,k=divmod(n-1,26);s=chr(65+k)+s
 return s
def finite(v): return isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v)
UNSET=object()
def F(f,v=UNSET): return {'f':'='+f.lstrip('='),**({'v':v} if v is not UNSET else {})}
def ref(s,c,r): return "'%s'!%s%d"%(s,c,r)
def sheet(name,title,headers,note='',widths=None):
 emit({'sheet':name,'title':title,'headers':headers,'note':note,'widths':widths or {}})
def row(n,values): emit({'row':n,'values':values})
def rate(a,b,v=UNSET,minus=True):return F(f'IF(AND(ISNUMBER({a}),ISNUMBER({b}),{b}>0),{a}/{b}'+('-1' if minus else '')+',"")',v)
def total(refs,v=UNSET):return F('SUM('+','.join(refs)+')',v)
def ratio(a,b):return a/b if finite(a) and finite(b) and b>0 else None
meta=d['metadata']; names=['00_总览与校验','01_整体市场月度','02_PP市场月度','03_高客单价市场月度','04_GENIMO品牌月度','05_年度YOY','06_BSR分层','07_Top100组成与进退','08_2027规划与分析','90_原始输入','91_BSR候选明细','92_聚合输入','93_数据校验','94_规则说明']
emit({'metadata':meta,'names':names})
def analysis_key(r):
 # No parent-ASIN or ASIN deduplication is applied in the current scope. Use
 # the immutable source row identity so every candidate row stays separate.
 return 'source:'+r['month']+'#'+str(r['sourceRow'])
rawmap={}; rawbounds={}; members=collections.defaultdict(list)
for i,r in enumerate(raw,5):
 m=r['month'];key=analysis_key(r)
 rawmap[i]=r
 # Every BSR-qualified source row is one business analysis unit.  Parent ASIN
 # is not used to collapse rows; it remains available for audit only.
 if finite(r.get('rank')) and 1 <= r['rank'] <= 100:
  members[(m,key)].append(i)
 rawbounds.setdefault(m,[i,i])[1]=i
detmap={}; detbounds={}
for i,r in enumerate(detail,5):detmap[(r['month'],r['familyKey'].split('|',1)[1])]=i;detbounds.setdefault(r['month'],[i,i])[1]=i
def priority(r):return (r['rank'] if finite(r['rank']) else 10000000)*10000000+(2-int(finite(r['sales']))-int(finite(r['revenue'])))*100000+int(not finite(r['price']))*10000+r['sourceRow']
assert max((r['rank'] or 0) for r in raw)<10000000
assert max(r['sourceRow'] for r in raw)<10000
sheet('90_原始输入','原始输入与可修订数值',['月份','源行号','ASIN','SKU','品牌','商品标题','父ASIN','原始小类BSR','数值BSR（可修订）','导入BSR状态','月销量','月销售额($)','价格($)','FBA原始','毛利率原始','Coupon原始','标题PP标记','品牌GENIMO标记','有效BSR≤100','来源子表','代表优先级分值'],
 '来源：地垫-卖家精灵市场数据.xlsx；SHA256 '+meta['sourceSha256']+'。I/K/L/M为数值输入；业务池先按1—100（含100）筛选并保留每条候选源行。父ASIN和ASIN仅作回勾和重复诊断，月份+源行号是行级Listing键；新增行、键或分类变化需重新导入。分值仅保留审计用途。',{'F':60,'H':30,'T':18,'U':25})
for i,r in rawmap.items():
 vals=[r['month'],r['sourceRow'],r['asin'],r['sku'],r['brand'],r['title'],r['parent'],r['sourceBsr'],r['rank'],r['bsrStatus'],r['sales'],r['revenue'],r['price'],r['fba'],r['margin'],r['coupon'],int(r['plastic']),int(r['genimo'])]
 vals += [F(f'IF(AND(ISNUMBER(I{i}),I{i}>=1,I{i}<=100,I{i}=INT(I{i})),1,0)',int(finite(r['rank']) and r['rank']<=100)),r['sourceSheet'],F(f'IF(AND(ISNUMBER(I{i}),I{i}>0,I{i}=INT(I{i})),I{i},10000000)*10000000+(2-IF(ISNUMBER(K{i}),1,0)-IF(ISNUMBER(L{i}),1,0))*100000+IF(ISNUMBER(M{i}),0,10000)+B{i}',priority(r))]
 row(i,vals)
sheet('91_BSR候选明细','BSR候选行与各池公式',['月份','Listing键','ASIN','父ASIN','SKU','品牌','商品标题','商品标题PP','品牌GENIMO','小类BSR','月销量','月销售额($)','价格($)','销量有效','销售额有效','价格有效','配对有效','BSR候选','候选行数','来源子表','源行号','90源行索引','整体候选序','整体入池','PP候选序','PP入池','高客单价候选序','高客单价入池','GENIMO候选序','GENIMO入池','配对销量','配对销售额','整体池内PP','整体池内高客单价','整体池内GENIMO','整体池内GENIMO PP','成员90行号','混合品牌组'],
 '仅包含 BSR 1—100（含100）的候选源行；不按父ASIN折叠，V列直接回勾对应90源行，H/I从该行标题/品牌回勾，W—AJ均从90输入重算。',{'B':22,'G':60,'AK':50})
topsets={k:set((r['month'],r['familyKey'].split('|',1)[1]) for r in v['topRows']) for k,v in cats.items()}
for i,r in enumerate(detail,5):
 m=r['month'];key=r['familyKey'].split('|',1)[1]; ids=members[(m,key)];a,b=rawbounds[m];lo,hi=detbounds[m]; rid=next(j for j in ids if rawmap[j]['sourceRow']==r['sourceRow']); matched=min(ids,key=lambda j:priority(rawmap[j]));assert rid==matched
 # MIN accepts at most 255 arguments: fold large membership sets in blocks.
 args=[ref('90_原始输入','U',j) for j in ids]
 while len(args)>200:args=['MIN('+','.join(args[k:k+200])+')' for k in range(0,len(args),200)]
 lookup=F('MATCH(MIN('+','.join(args)+f"),'90_原始输入'!$U${a}:$U${b},0)+{a-1}",rid) if len(ids)>1 else F(f"ROW('90_原始输入'!A{rid})",rid)
 def pull(c,val,numeric=False):
  q=f"INDEX('90_原始输入'!${c}$5:${c}${len(raw)+4},$V{i}-4)"
  return F(f'IF(ISNUMBER({q}),{q},"")' if numeric else f'IF({q}="","",{q})',val)
 vals=[m,key,pull('C',r['asin']),pull('G',r['parent']),pull('D',r['sku']),pull('E',r['brand']),pull('F',r['title']),pull('Q',int(r['plastic']),True),pull('R',int(r['genimo']),True),pull('I',r['rank'],True),pull('K',r['sales'],True),pull('L',r['revenue'],True),pull('M',r['price'],True)]
 vals += [F(f'IF(ISNUMBER({c}{i}),1,0)',int(finite(r[field]))) for c,field in [('K','sales'),('L','revenue'),('M','price')]]
 vals += [F(f'IF(AND(N{i}=1,O{i}=1,K{i}>0),1,0)',int(finite(r['sales']) and finite(r['revenue']) and r['sales']>0)),F(f'IF(AND(ISNUMBER(J{i}),J{i}>=1,J{i}<=100,J{i}=INT(J{i})),1,0)',int(finite(r['rank']) and r['rank']<=100)),len(ids),pull('T',r['sourceSheet']),pull('B',r['sourceRow'],True),lookup]
 def order(extra,eligible):
  cond=''
  for c,v in extra:cond+=f',${c}${lo}:${c}${hi},{v}'
  prior=''
  for c,v in extra:prior+=f',${c}${lo}:{c}{i},{v}'
  return F(f'IF(AND(R{i}=1,{eligible}),COUNTIFS($J${lo}:$J${hi},"<"&J{i},$R${lo}:$R${hi},1{cond})+COUNTIFS($J${lo}:J{i},J{i},$R${lo}:R{i},1{prior}),0)')
 vals += [order([],'TRUE'),F(f'IF(AND(R{i}=1),1,0)',int((m,key) in topsets['overall'])),order([('H',1)],f'H{i}=1'),F(f'IF(AND(R{i}=1,H{i}=1),1,0)',int((m,key) in topsets['pp'])),order([('H',0)],f'H{i}=0'),F(f'IF(AND(R{i}=1,H{i}=0),1,0)',int((m,key) in topsets['nonpp'])),order([('I',1)],f'I{i}=1'),F(f'IF(AND(R{i}=1,I{i}=1),1,0)',int((m,key) in topsets['genimoIndependent'])),F(f'IF(Q{i}=1,K{i},0)',r['sales'] if finite(r['sales']) and finite(r['revenue']) and r['sales']>0 else 0),F(f'IF(Q{i}=1,L{i},0)',r['revenue'] if finite(r['sales']) and finite(r['revenue']) and r['sales']>0 else 0),F(f'X{i}*H{i}',int((m,key) in topsets['globalPP'])),F(f'X{i}*(1-H{i})',int((m,key) in topsets['globalHigh'])),F(f'X{i}*I{i}',int((m,key) in topsets['genimo'])),F(f'X{i}*I{i}*H{i}',int((m,key) in topsets['genimoPP'])),','.join(map(str,ids)),int(len({rawmap[j]['brand'].lower() for j in ids})>1)]
 row(i,vals)
aggmap={}; cache={}; aggr=[]
levels=[('全部',None),('Top100',None)]+[(b['name'],b) for b in bands]
for scope,cat in cats.items():
 for mi,m in enumerate(months):
  for level,band in levels:
   r=cat['monthly'][mi] if level=='全部' else cat['topMonthly'][mi] if level=='Top100' else (cat['tiers'] if band['key'] in cat['tiers'] else cat['fine'])[band['key']][mi]
   n=len(aggr)+5;aggmap[(scope,m,level)]=n;cache[(scope,m,level)]=r;aggr.append((scope,m,level,band,r))
sheet('92_聚合输入','月度范围与层级公式汇总',['月份','范围键','范围','层级','候选行数','销量','销售额($)','价格合计($)','价格有效数','配对销量','配对销售额($)','销量有效数','销售额有效数','状态','平均标价($)','加权成交均价($)','候选行数回勾','截去数','原始候选行数'],
 '从91 BSR候选行各月份有界范围COUNTIFS/SUMIFS计算。无有效数值留空，真实0保留。整体=PP+高客单价；头中尾和五档均从同一BSR候选行池分层。')
for n,(scope,m,level,band,z) in enumerate(aggr,5):
 lo,hi=detbounds[m]
 cond=[]
 # Every aggregation starts with Q/R's BSR candidate flag.  The remaining
 # conditions are row-level title/brand classifications from 91, so an edited
 # BSR or candidate input propagates through every scope.
 full={'overall':[('R',1)],'pp':[('R',1),('H',1)],'nonpp':[('R',1),('H',0)],'genimo':[('R',1),('I',1)],'genimoPP':[('R',1),('I',1),('H',1)],'genimoIndependent':[('R',1),('I',1)],'globalPP':[('R',1),('AG',1)],'globalHigh':[('R',1),('AH',1)]}
 tc={'overall':'X','pp':'Z','nonpp':'AB','genimo':'AI','genimoPP':'AJ','globalPP':'AG','globalHigh':'AH','genimoIndependent':'AD'}
 cond=full.get(scope,[]) if level=='全部' else [(tc[scope],1)]
 def crange(c):return f"'91_BSR候选明细'!${c}${lo}:${c}${hi}"
 args=','.join(crange(c)+','+str(v) for c,v in cond)
 if band: args+=(',' if args else '')+crange('J')+',">='+str(band['lo'])+'",'+crange('J')+',"<='+str(band['hi'])+'"'
 if not args:args=crange('A')+',A'+str(n)
 def sums(c):return 'SUMIFS('+crange(c)+','+args+')'
 def cs(c):return F(sums(c))
 vals=[m,scope,cats[scope]['title'],level,F('COUNTIFS('+args+')',z['count']),F(f'IF(L{n}=0,"",{sums("K")})',z['sales']),F(f'IF(M{n}=0,"",{sums("L")})',z['revenue']),cs('M'),cs('P'),cs('AE'),cs('AF'),F(sums('N'),z['count']-z['missingSales']),F(sums('O'),z['count']-z['missingRevenue']),F(f'IF(E{n}=0,"无样本",IF(OR(L{n}/E{n}<0.95,M{n}/E{n}<0.95),"指标缺失","覆盖可用"))'),rate(f'H{n}',f'I{n}',z['avgPrice'],False),rate(f'K{n}',f'J{n}',z['pairedAsp'],False)]
 candargs=[('R',1)]+full.get(scope,[])
 ca=','.join(crange(c)+','+str(v) for c,v in candargs)
 vals += [F('COUNTIFS('+ca+')') if level=='Top100' else None,F(f'MAX(0,Q{n}-E{n})') if level=='Top100' else None]
 # Source candidate counts use every row that passed BSR 1..100, including
 # rows without a parent ASIN. Classification is taken directly from that
 # candidate row; no family representative or parent collapse is involved.
 ids=[j for j,rr in rawmap.items() if level=='Top100' and rr['month']==m and finite(rr.get('rank')) and 1 <= rr['rank'] <= 100 and ((scope=='overall') or (scope in ['pp','globalPP'] and rr['plastic']) or (scope in ['nonpp','globalHigh'] and not rr['plastic']) or (scope in ['genimo','genimoIndependent'] and rr['genimo']) or (scope=='genimoPP' and rr['genimo'] and rr['plastic']))]
 refs=[ref('90_原始输入','S',j) for j in ids]
 while len(refs)>200:refs=['SUM('+','.join(refs[k:k+200])+')' for k in range(0,len(refs),200)]
 vals += [total(refs) if refs else None];row(n,vals)
def ar(scope,m,level,c):return ref('92_聚合输入',c,aggmap[(scope,m,level)])
def link(scope,m,level,c,field=None):return F(ar(scope,m,level,c),cache[(scope,m,level)].get(field) if field else UNSET)
def guarded(q,v=UNSET):return F(f'IF(ISNUMBER({q}),{q},"")',v)
def linkn(scope,m,level,c,field=None):return guarded(ar(scope,m,level,c),cache[(scope,m,level)].get(field) if field else UNSET)
def py(m):return str(int(m[:4])-1)+m[4:]
def quality(scope,m,level):
 n=aggmap[(scope,m,level)];cur=ar(scope,m,level,'E');base=ar(scope,py(m),level,'E') if py(m) in months else '0';idx=months.index(m);prev=ar(scope,months[idx-1],level,'E') if idx else '0'
 rawchecks=[]
 for pm in [py(m),months[idx-1] if idx else None]:
  if pm in rawbounds:
   a,b=rawbounds[m];pa,pb=rawbounds[pm];rawchecks.append(f"ABS(COUNTA('90_原始输入'!A{a}:A{b})-COUNTA('90_原始输入'!A{pa}:A{pb}))>0.3*COUNTA('90_原始输入'!A{pa}:A{pb})")
 rawcondition=(','+','.join(rawchecks)) if rawchecks else ''
 return F(f'IF({cur}=0,"无样本",IF(OR(IF({base}>0,ABS({cur}-{base})>0.3*{base},FALSE),IF({prev}>0,ABS({cur}-{prev})>0.3*{prev},FALSE),{ar(scope,m,level,"L")}/{cur}<0.95,{ar(scope,m,level,"M")}/{cur}<0.95{rawcondition}),"可比性受限","覆盖可用"))',cache[(scope,m,level)]['quality'])
def periodvals(scope,m,level):
 z=cache[(scope,m,level)];v=[m,link(scope,m,level,'E','count'),linkn(scope,m,level,'F','sales'),linkn(scope,m,level,'G','revenue'),linkn(scope,m,level,'O','avgPrice'),linkn(scope,m,level,'P','pairedAsp')]
 for c,k in [('F','momSales'),('G','momRevenue')]:v.append(rate(ar(scope,m,level,c),ar(scope,py(m),level,c),z[k]) if py(m) in months else F('""',None))
 return v
commonheaders=['月份','Listing数','销量','销售额($)','平均标价($)','加权成交均价($)','销量MOM','销售额MOM']
for scope,name in [('overall',names[1]),('pp',names[2]),('nonpp',names[3])]:
 sheet(name,cats[scope]['title']+' 月度汇总',commonheaders+['销量有效数','销售额有效数','价格有效数','BSR候选行数','BSR候选销量','BSR候选销售额($)','BSR候选标价($)','BSR候选成交均价($)','BSR候选销量MOM','BSR候选销售额MOM','可比性','BSR候选原始行数','BSR候选行数（回勾）','截去数','BSR候选样本状态','去年同月候选行数','去年同月销量','去年同月销售额($)'],
 'MOM=本月/去年同月−1；非自然月环比。空白=无样本/无有效值/无正数基期。整体=PP+高客单价；原始销量为源估算值。业务池先筛1—100并保留每条候选行，不按父ASIN去重。')
 for n,m in enumerate(months,5):
  v=periodvals(scope,m,'全部');v += [link(scope,m,'全部',c) for c in ['L','M','I']];v+=periodvals(scope,m,'Top100')[1:];v +=[quality(scope,m,'全部')]+[link(scope,m,'Top100',c) for c in ['S','Q','R']]+[F(f'IF(L{n}<100,"少于100个样本",IF(L{n}>100,"超过100；并列或多小类","100个样本"))')]
  v +=[linkn(scope,py(m),'全部',c) if py(m) in months else None for c in ['E','F','G']];row(n,v)
sheet(names[4],'GENIMO 品牌整体、PP和BSR候选行池表现',commonheaders+['销量份额','销售额份额','BSR候选行数','可比性'],
 'GENIMO只从同一BSR 1—100候选行池筛选；整体份额分母整体候选行池，PP份额分母PP候选行池。所有MOM对去年同月。')
for block,(scope,level,label,base) in enumerate([('genimo','全部','4.1 GENIMO整体表现（BSR候选行池）','overall'),('genimoPP','全部','4.2 GENIMO PP表现（BSR候选行池）','pp'),('genimo','Top100','4.4 GENIMO BSR候选行池表现','overall'),('genimoIndependent','Top100','4.4 GENIMO BSR候选行池回勾','overall')]):
 start=5+block*(len(months)+4)
 if block:row(start-2,[label]);row(start-1,commonheaders+['销量份额','销售额份额','BSR候选行数','可比性'])
 for n,m in enumerate(months,start):
  z=cache[(scope,m,level)];b=cache[(base,m,level)];v=periodvals(scope,m,level)+[rate(ar(scope,m,level,'F'),ar(base,m,level,'F'),ratio(z['sales'],b['sales']),False),rate(ar(scope,m,level,'G'),ar(base,m,level,'G'),ratio(z['revenue'],b['revenue']),False),link('genimoPP' if scope=='genimoPP' else 'genimo',m,'Top100','E','count'),quality(scope,m,level)];row(n,v)
years=sorted(set(m[:4] for m in months))
annual_headers=['年份','范围','层级','本期覆盖月份','YOY比较期间','Listing月次','销量','销售额($)','平均标价($)','成交均价($)','销量YOY','销售额YOY','同比本期销量','同比基期销量','同比本期金额','同比基期金额','样本/基期状态']
def annual_values(scope,level,y,band=None):
 cat=cats[scope]; z=next(a for a in (cat['annual'] if level=='全部' else cat['topAnnual'] if level=='Top100' else cat['annualBands'][band['key']]) if a['year']==y);ym=[m for m in months if m.startswith(y)];cm=[y+s for s in z['commonMonths']];pm=[str(int(y)-1)+s for s in z['commonMonths']]
 def summ(c,ms):return 'SUM('+','.join(ar(scope,m,level,c) for m in ms)+')' if ms else '0'
 def sums(c,ms,valid=None):return F(f'IF({summ(valid,ms)}=0,"",{summ(c,ms)})' if valid else summ(c,ms))
 vals=[y,cat['title'],level,z['coverage'],z['yoyPeriod'] or '无共同基期',sums('E',ym),sums('F',ym,'L'),sums('G',ym,'M'),rate(summ('H',ym),summ('I',ym),z['avgPrice'],False),rate(summ('K',ym),summ('J',ym),z['pairedAsp'],False)]
 vals += [F(f'IF(AND({summ(valid,cm)}>0,{summ(valid,pm)}>0,{summ(c,pm)}>0),{summ(c,cm)}/{summ(c,pm)}-1,"")',z[k]) if cm else F('""',None) for c,k,valid in [('F','yoySales','L'),('G','yoyRevenue','M')]]
 vals += [sums(c,ms,valid) for c,ms,valid in [('F',cm,'L'),('F',pm,'L'),('G',cm,'M'),('G',pm,'M')]]
 return vals
sheet(names[5],'年度总量与共同月份YOY',annual_headers,'年度数为Listing月次；本期总量与YOY分子可不同。只有相邻年份共同覆盖月份参与YOY；分子和分母在M—P列。')
n=5
for scope in ['overall','pp','nonpp','genimo','genimoPP']:
 for level in ['全部','Top100']:
  for y in years:
   v=annual_values(scope,level,y);v+=[F(f'IF(F{n}=0,"无样本",IF(OR(N{n}="",P{n}="",N{n}<=0,P{n}<=0),"无有效正数基期","同周期；检查样本可比性"))')];row(n,v);n+=1
sheet(names[6],'BSR候选行池粗细档 月度与年度明细',['月份','范围','层级','BSR下限','BSR上限','候选行数','销量','销售额($)','平均标价($)','销量MOM','销售额MOM','可比性'],
 '粗档与细档均从同一BSR 1—100候选行池回勾；不按父ASIN去重，GENIMO使用候选行品牌。月度之后接年度分层，按相邻年共同月份计算。')
n=5
for scope in ['overall','pp','nonpp','genimo']:
 for band in bands:
  level=band['name']
  for m in months:
   v=periodvals(scope,m,level);row(n,[m,cats[scope]['title'],level,band['lo'],band['hi']]+v[1:5]+v[6:]+[quality(scope,m,level)]);n+=1
n+=3;row(n,annual_headers);n+=1
for scope in ['overall','pp','nonpp','genimo']:
 for band in bands:
  for y in years:row(n,annual_values(scope,band['name'],y,band));n+=1
sheet(names[7],'BSR候选行池组成与GENIMO进留退',['月份','整体候选行数','PP候选行数','高客单价候选行数','整体销量','PP销量','高客单价销量','整体金额','PP金额','高客单价金额','GENIMO当前候选行数','GENIMO基期候选行数','进入','保留','退出','基期月份'],
 '组成对同一BSR候选行池分区；整体=PP+高客单价。进入/保留/退出比较相邻自然月候选行池内GENIMO ASIN键集合，不代表上新或下架。详细Listing键见下方。')
# Include all imported keys in either month so numerical BSR edits can change membership.
move_rows=[];move_ranges={};n=5+len(months)+4
for m in months:
 previous=str(int(m[:4])-1)+'12' if m[4:]=='01' else m[:4]+str(int(m[4:])-1).zfill(2)
 keys=sorted(set(k for mm,k in detmap if mm in [m,previous] and detail[detmap[(mm,k)]-5]['genimo']))
 start=n
 for key in keys:move_rows.append((n,m,previous,key,detmap.get((m,key)),detmap.get((previous,key))));n+=1
 move_ranges[m]=(start,n-1)
for n,m in enumerate(months,5):
 lo,hi=move_ranges[m];v=[m]+[link(scope,m,'Top100',c) if c=='E' else linkn(scope,m,'Top100',c) for c in ['E','F','G'] for scope in ['overall','globalPP','globalHigh']]
 v += [link('genimo',m,'Top100','E')];move=next(x for x in d['movements'] if x['month']==m)
 v += [link('genimo',move['previous'],'Top100','E') if move['available'] else None]
 v +=[F(f'COUNTIFS($D${lo}:$D${hi},"{state}")',move[field]) if move['available'] else None for state,field in [('进入','entered'),('保留','retained'),('退出','exited')]];v+=[move['previous']];row(n,v)
row(5+len(months)+2,['月份','Listing键','基期月份','进留退','本期榜内','基期榜内','本期BSR','基期BSR'])
for n,m,prev,key,cur,prior in move_rows:
 current=ref('91_BSR候选明细','AI',cur) if cur else '0';before=ref('91_BSR候选明细','AI',prior) if prior else '0'
 state=F(f'IF(AND(E{n}=1,F{n}=1),"保留",IF(E{n}=1,"进入",IF(F{n}=1,"退出","未入榜")))') if prev in months else '无基期'
 row(n,[m,key,prev,state,F(current),F(before),guarded(ref('91_BSR候选明细','J',cur)) if cur else None,guarded(ref('91_BSR候选明细','J',prior)) if prior else None])
sheet(names[8],'2027链接规划与四部分分析',['层级','目标BSR','核心期销量','Listing月次','每Listing月次销量','销量权重','目标新增链接数','历史平均标价($)','动作与条件'],
 '建议情景，不是预测实绩。B3输入非负整数总链接数；默认空。按最近核心期GENIMO榜内层级销量权重分配，前两档向下取整，尾档取余。可比性受限先核验源。',{'I':85})
row(3,['目标新增链接总数',None,'核心期',','.join(m for m in months if m.startswith(months[-1][:4]))])
ym=[m for m in months if m.startswith(months[-1][:4])]
for n,band in enumerate(bands[:3],5):
 def ss(c):return 'SUM('+','.join(ar('genimo',m,band['name'],c) for m in ym)+')'
 alloc=F(f'IF(AND(ISNUMBER($B$3),$B$3>=0,$B$3=INT($B$3),ISNUMBER(F{n})),INT($B$3*F{n}),"")') if n<7 else F('IF(AND(ISNUMBER($B$3),$B$3>=0,$B$3=INT($B$3),ISNUMBER(F7)),$B$3-SUM(G5:G6),"")')
 row(n,[band['name'],str(band['lo'])+'—'+str(band['hi']),F(ss('F')),F(ss('E')),rate(f'C{n}',f'D{n}',minus=False),rate(f'C{n}','SUM($C$5:$C$7)',minus=False),alloc,rate(ss('H'),ss('I'),minus=False),['守住核心链接；连续两个月销售额和份额下降则复核字段覆盖与BSR。','小批测款；连续两个月份额与销售额均增长且BSR改善再考虑追加。','探索备选；连续两个月无销量或退出榜单则重新评估，不自动判定下架。'][n-5]])
row(9,['分配合计',None,total(['C5:C7']),total(['D5:D7']),None,total(['F5:F7']),F('IF(ISNUMBER(B3),SUM(G5:G7),"")'),None,'花型、颜色、尺寸未可靠提取；不据此编造产品趋势。'])
n=12
for scope in ['overall','pp','nonpp','genimo']:
 cat=cats[scope];latest=cat['monthly'][-1];annual=cat['annual'][-1]
 row(n,[cat['title'],'截至 '+months[-1],None,None,None,None,None,None,'MOM=去年同月；年度YOY='+str(annual['yoyPeriod'])]);n+=1
 row(n,['最新销量',linkn(scope,months[-1],'全部','F'),'最新销售额',linkn(scope,months[-1],'全部','G'),'加权成交均价',linkn(scope,months[-1],'全部','P'),None,None,'方向、幅度需结合销量覆盖和样本数量；快照均价不证明同款涨价。']);n+=1
 row(n,['行动依据',quality(scope,months[-1],'全部'),None,None,None,None,None,None,'建议：先比较同周期销售额和单Listing效率，覆盖受限先核验；市场增长不直接等于品牌机会。']);n+=2
row(30,['2027市场选择证据','共同月份',None,None,None,None,None,None,'相对变化更强仅用于选择核验与小批测款候选；不自动转移全部链接。'])
row(31,['候选市场','可比销量','可比金额($)','金额YOY','GENIMO本期金额份额','GENIMO基期金额份额','份额变化','可比性','建议'])
genimo_year_row=5+3*len(years)*2+len(years)-1;genimo_pp_year_row=5+4*len(years)*2+len(years)-1
for nr,scope,baseannual in [(32,'pp',5+len(years)*2+len(years)-1),(33,'nonpp',5+2*len(years)*2+len(years)-1)]:
 def an(c,r):return ref('05_年度YOY',c,r)
 def brandrev(c):return an(c,genimo_pp_year_row) if scope=='pp' else f'({an(c,genimo_year_row)}-{an(c,genimo_pp_year_row)})'
 year=months[-1][:4];pm=[str(int(year)-1)+m[4:] for m in ym if str(int(year)-1)+m[4:] in months]
 def countdiff(ms):return 'SUM('+','.join(ar('genimo',m,'全部','E') for m in ms)+')-SUM('+','.join(ar('genimoPP',m,'全部','E') for m in ms)+')'
 def brandshare(c,ms):return rate(brandrev(c),an(c,baseannual),minus=False) if scope=='pp' else F(f'IF(AND(({countdiff(ms)})>0,{an(c,baseannual)}>0),{brandrev(c)}/{an(c,baseannual)},"")')
 row(nr,[cats[scope]['title'],guarded(an('M',baseannual)),guarded(an('O',baseannual)),guarded(an('L',baseannual)),brandshare('O',ym),brandshare('P',pm),F(f'IF(AND(ISNUMBER(E{nr}),ISNUMBER(F{nr})),E{nr}-F{nr},"")'),quality(scope,months[-1],'全部'),'先核验样本；跟踪连续两个月金额、份额与BSR，再决定扩充。'])
row(35,['优先核验/小测市场',F('IF(AND(ISNUMBER(D32),ISNUMBER(D33)),IF(D32>=D33,A32,A33),"证据不足")'),None,None,None,None,None,None,'依据：同周期销售额变化相对更强；保留PP现有品牌盘，先用实测验证，再分配链接。'])
sheet(names[0],'户外地垫市场分析 SPEC 2.1',['项目','值','用途/边界','对应工作表'], '一份主源、一个BSR候选行池、四部分业务；先筛BSR 1—100后保留每条候选行，不按父ASIN去重，整体=PP+高客单价，GENIMO是品牌视角。数值修订可在表内重算；结构变更重新导入。',{'A':28,'B':65,'C':75,'D':30})
for n,v in enumerate([['版本','2.1','当前规范','94_规则说明'],['主源SHA256',meta['sourceSha256'],'核验原始文件未改','90_原始输入'],['批次',meta['batchId'],meta['generatedAt'],'所有交付文件'],['截止月份',months[-1],'未完整年度只比较共同月份','05_年度YOY'],['原始行数',F(f"COUNTA('90_原始输入'!A5:A{len(raw)+4})",len(raw)),'原始业务行','90_原始输入'],['BSR候选行月次',F(f"COUNTA('91_BSR候选明细'!A5:A{len(detail)+4})",len(detail)),'BSR 1—100后保留每条候选源行，不按父ASIN去重','91_BSR候选明细'],['整体市场','01 / 05 / 06 / 07','BSR候选行池整体、分层、组成','01_整体市场月度'],['PP市场','02 / 05 / 06','候选行标题命中plastic','02_PP市场月度'],['高客单价市场','03 / 05 / 06','PP补集，无额外价格门槛','03_高客单价市场月度'],['GENIMO','04 / 05 / 06 / 07 / 08','同一BSR候选行池的品牌份额、表现、进留退、2027计划','04_GENIMO品牌月度']],5):row(n,v)
sheet(names[12],'终端校验',['月份/项目','校验项','差额/结果','状态'], '只读观察，不给任何业务公式提供输入。数值比较差额，文本等值比较。')
n=5
row(n,['主源','原始业务行数',F(f"COUNTA('90_原始输入'!A5:A{len(raw)+4})-{len(raw)}",0),F('IF(C5=0,"通过","失败")','通过')]);n+=1
row(n,['边界','BSR100存在',F(f'IF(COUNTIFS(\'91_BSR候选明细\'!J5:J{len(detail)+4},100)>0,"通过","失败")','通过'),F('IF(C6="通过","通过","失败")','通过')]);n+=1
for m in months:
 for c,label in [('E','数量'),('F','销量'),('G','金额')]:
  for base,parts,pool in [('overall',['pp','nonpp'],'全部'),('overall',['globalPP','globalHigh'],'Top100')]:
   f=ar(base,m,pool,c)+'-'+ar(parts[0],m,pool,c)+'-'+ar(parts[1],m,pool,c);row(n,[m,pool+label+'分区回勾',F(f,0),F(f'IF(ABS(C{n})<0.000001,"通过","失败")','通过')]);n+=1
 for scope in ['overall','pp','nonpp','genimo']:
  for bandset in [bands[:3],bands[3:]]:
   for c in ['E','F','G']:
    base=ar(scope,m,'Top100',c);bs='SUM('+','.join(ar(scope,m,b['name'],c) for b in bandset)+')';f=f'IF({base}="",IF({bs}=0,0,"缺失不一致"),{base}-{bs})';row(n,[m,scope+'分层'+c,F(f,0),F(f'IF(ISNUMBER(C{n}),IF(ABS(C{n})<0.000001,"通过","失败"),"失败")','通过')]);n+=1
 for lhs,rhs in [('K','M'),('L','O')]:
  mr=months.index(m)+5
  if d['movements'][mr-5]['available']:
   row(n,[m,'GENIMO进留退回勾',F(f"'07_Top100组成与进退'!{lhs}{mr}-'07_Top100组成与进退'!{rhs}{mr}-'07_Top100组成与进退'!N{mr}",0),F(f'IF(C{n}=0,"通过","失败")','通过')]);n+=1
sheet(names[13],'规则与人工复核',['规则项','说明'], '人工复核按90输入→91候选行→92聚合→01—08业务结果→93校验。',{'A':30,'B':130})
rules=[('唯一主源','地垫-卖家精灵市场数据.xlsx；原文件只读。'),('分析顺序','原始行先筛BSR 1—100（含100），保留每条候选源行，不按月份+父ASIN去重，之后按候选行标题/品牌分类。'),('输入边界','90 I/K/L/M数字可改；新增/删除行、标题/品牌分类、父ASIN键、月份需重新导入。父ASIN只作审计字段，ASIN和源行号用于回勾。'),('分析单位','每条满足1≤BSR≤100的源行是一个Listing候选月次；同一父ASIN下的多个子体行分别计入，不能把重复行合并。'),('PP分类','候选行商品标题匹配完整单词plastic；命中为PP，未命中为高客单价补集，不设置额外价格门槛。'),('整体市场','同一BSR候选行池中PP+高客单价；数量、销量、销售额逐月必须回加一致。'),('GENIMO','同一BSR候选行池按候选行品牌精确匹配；GENIMO PP为品牌与PP交集。'),('BSR候选行池','符合1≤BSR≤100的候选全部保留，不再额外截取100或去重；头中尾和五档均从该池分层。'),('缺失与零','无样本/无有效值留空，真实0保留；增长率基期≤0不可用。'),('时间','月度MOM对去年同月；年度YOY对相邻年度共同月份，查看05 M—P分子分母。'),('均价','标价=有效价格总和/有效数；成交均价=销量>0且金额有效记录的金额合计/销量合计。'),('质量','样本比去年同月/上月变化>30%或销售字段覆盖<95%提示可比性受限；阈值为核验政策。'),('2027规划','08 B3默认空；输入非负整数按核心期GENIMO候选行池层级销量权重分配；前两档向下取整尾档取余。'),('源真实性','卖家精灵估算不等于真实成交；程序核验统计正确性，不能保证供应方估算真实性。'),('不计算利润','毛利率/FBA/Coupon仅保留原始证据，不进入分析/成本/利润推导。'),('文本与报告','HTML/JSON/Markdown为同批次静态报告，修改XLSX输入后需重新构建才同步外部文件。')]
for n,v in enumerate(rules,5):row(n,list(v))
out.close();print(json.dumps({'rows':len(raw),'detail':len(detail),'aggregate':len(aggr),'plan':str(ROOT/'tmp/formula_market_builder/workbook_plan.jsonl')},ensure_ascii=False))
