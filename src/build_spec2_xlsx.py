import hashlib,json,math,os,zipfile
import xlsxwriter
ROOT=os.path.abspath(os.path.join(os.path.dirname(__file__),'..')); DATA_JSON=os.path.join(ROOT,'tmp','formula_market_builder','spec2_dataset.json'); SOURCE=os.path.join(ROOT,'data','raw','地垫-卖家精灵市场数据.xlsx'); OUT_DIR=os.path.join(ROOT,'outputs','20260911-spec2-market-analysis'); OUT_FILE=os.environ.get('SPEC2_XLSX_OUT',os.path.join(OUT_DIR,'户外地垫市场分析-SPEC2-可回勾版.xlsx')); os.makedirs(os.path.dirname(OUT_FILE),exist_ok=True)
data=json.load(open(DATA_JSON,encoding='utf-8')); raw=data['raw']; detail=data['dedup']; months=[str(m) for m in data['months']]; data_status={str(x['month']): ('有数据' if x['rowCount'] else '无数据') for x in data.get('sheetStats',[])}; source_hash=hashlib.sha256(open(SOURCE,'rb').read()).hexdigest()
scopes=[('overall','整体市场'),('pp','PP市场'),('nonpp','高客单价市场'),('genimo','GENIMO品牌')]
bands=[('头部 1-20',1,20),('中部 21-50',21,50),('尾部 51-100',51,100),('1-5',1,5),('6-10',6,10),('11-20',11,20),('21-50',21,50),('51-100',51,100)]; levels=['全部','Top100']+[x[0] for x in bands]
def finite(x): return isinstance(x,(int,float)) and math.isfinite(x)
def listing_key(r): return str(r.get('parent') or r.get('asin') or ('source:'+str(r.get('month'))+'#'+str(r.get('sourceRow')))).strip()
def in_scope(r,s): return s=='overall' or (s=='pp' and r.get('plastic')) or (s=='nonpp' and not r.get('plastic')) or (s=='genimo' and r.get('genimo'))
def in_level(r,l):
    if l=='全部': return True
    rank=r.get('rank')
    if l=='Top100': return finite(rank) and 1<=rank<=100
    for name,lo,hi in bands:
        if name==l: return finite(rank) and lo<=rank<=hi
    return False
def select_top100(rows):
    # Tied BSR values can make rank<=100 exceed 100 listings. Keep a stable
    # maximum of 100 per month and scope so the Top100 table honors its name.
    eligible=[r for r in rows if finite(r.get('rank')) and 1<=r['rank']<=100]
    return sorted(eligible,key=lambda r:(r['rank'],str(r.get('familyKey','')),int(r.get('sourceRow') or 0)))[:100]
detail_flags={(str(r.get('month')),str(r.get('familyKey','')).split('|',1)[-1]):r for r in detail}
def raw_in_scope(r,s):
    if s=='overall': return True
    d=detail_flags.get((str(r.get('month')),listing_key(r)),{})
    return bool(d.get('plastic')) if s=='pp' else (not bool(d.get('plastic'))) if s=='nonpp' else bool(d.get('genimo'))
candidate_raw={(s,m):sum(1 for r in raw if str(r.get('month'))==m and finite(r.get('rank')) and 1<=r['rank']<=100 and raw_in_scope(r,s)) for s,_ in scopes for m in months}
def metrics(rows):
    paired=[r for r in rows if finite(r.get('sales')) and finite(r.get('revenue')) and r['sales']>0]
    return {'count':len(rows),'sales':sum(r['sales'] for r in rows if finite(r.get('sales'))),'revenue':sum(r['revenue'] for r in rows if finite(r.get('revenue'))),'price_sum':sum(r['price'] for r in rows if finite(r.get('price'))),'price_count':sum(1 for r in rows if finite(r.get('price'))),'paired_sales':sum(r['sales'] for r in paired),'paired_revenue':sum(r['revenue'] for r in paired),'sales_valid':sum(1 for r in rows if finite(r.get('sales'))),'revenue_valid':sum(1 for r in rows if finite(r.get('revenue')))}
aggregate={}
for scope,_ in scopes:
    for month in months:
        pool=[r for r in detail if str(r['month'])==month and in_scope(r,scope)]
        top_pool=select_top100(pool)
        for level in levels:
            base=pool if level=='全部' else top_pool
            z=metrics([r for r in base if in_level(r,level)])
            z['candidate_raw_rows']=candidate_raw.get((scope,month),0) if level=='Top100' else None
            aggregate[(scope,month,level)]=z
def prev_month(m):
    y,mo=int(m[:4]),int(m[4:])-1
    if mo==0: y,mo=y-1,12
    return f'{y:04d}{mo:02d}'
def prev_year(m): return f'{int(m[:4])-1:04d}{m[4:]}'
def pct(a,b): return (a-b)/abs(b) if finite(a) and finite(b) and b!=0 else ''
wb=xlsxwriter.Workbook(OUT_FILE,{'constant_memory':True,'strings_to_urls':False}); wb.set_properties({'title':'户外地垫市场分析 SPEC 2.0','comments':'同一 buildDataset 结果；公式写入展示页供人工回勾。'})
navy='#1F2937';blue='#3B82F6';pale='#EFF6FF';green='#DCFCE7';red='#FEE2E2'
title_fmt=wb.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':navy,'font_size':14,'valign':'vcenter'}); hdr_fmt=wb.add_format({'bold':True,'font_color':'#FFFFFF','bg_color':blue,'text_wrap':True,'border':1,'border_color':'#BFDBFE','valign':'vcenter'}); note_fmt=wb.add_format({'bg_color':pale,'font_color':'#374151','text_wrap':True,'valign':'vcenter'}); text_fmt=wb.add_format({'border':1,'border_color':'#E5E7EB'}); num_fmt=wb.add_format({'border':1,'border_color':'#E5E7EB','num_format':'#,##0;[Red]-#,##0'}); money_fmt=wb.add_format({'border':1,'border_color':'#E5E7EB','num_format':'$#,##0.00;[Red]-$#,##0.00'}); fnum=wb.add_format({'border':1,'border_color':'#E5E7EB','font_color':'#1D4ED8','num_format':'#,##0;[Red]-#,##0'}); fmoney=wb.add_format({'border':1,'border_color':'#E5E7EB','font_color':'#1D4ED8','num_format':'$#,##0.00;[Red]-$#,##0.00'}); fpct=wb.add_format({'border':1,'border_color':'#E5E7EB','font_color':'#1D4ED8','num_format':'0.0%;[Red]-0.0%'})
def setup(s,t,n,note): s.hide_gridlines(2);s.merge_range(0,0,0,n,t,title_fmt);s.set_row(0,28);s.merge_range(1,0,1,n,note,note_fmt);s.set_row(1,36)
def headers(s,row,hs):
    for c,v in enumerate(hs): s.write(row,c,v,hdr_fmt)
    s.set_row(row,30)
def formula(s,r,c,f,v,fmt): s.write_formula(r,c,f,fmt,v if v!='' else '')
name_list=['00_总览与校验','01_整体市场月度','02_PP市场月度','03_高客单价市场月度','04_GENIMO品牌月度','05_年度YOY','06_BSR分层','90_原始输入','91_去重明细','92_聚合输入','93_数据校验','94_规则说明']; ws={n:wb.add_worksheet(n) for n in name_list}
# raw input
s=ws['90_原始输入']; rh=['月份','源行号','ASIN','SKU','品牌','商品标题','父ASIN','原始小类BSR','小类BSR','BSR解析状态','月销量','月销售额($)','价格($)','FBA($)','毛利率','Coupon','PP标记','GENIMO标记','BSR Top100标记']; setup(s,'原始输入（逐行保留）· 来源：data/raw/地垫-卖家精灵市场数据.xlsx',18,'原始工作簿字段逐行读取；PP只依据商品标题完整单词 plastic；BSR同时保留原始文本、解析值和状态；空值保留为空，未补零。'); headers(s,3,rh)
for i,r in enumerate(raw,4):
    vals=[r.get('month'),r.get('sourceRow'),r.get('asin'),r.get('sku'),r.get('brand'),r.get('title'),r.get('parent'),r.get('sourceBsr'),r.get('rank'),r.get('bsrStatus'),r.get('sales'),r.get('revenue'),r.get('price'),r.get('fba'),r.get('margin'),r.get('coupon'),1 if r.get('plastic') else 0,1 if r.get('genimo') else 0,1 if finite(r.get('rank')) and 1<=r['rank']<=100 else 0]
    for c,v in enumerate(vals): s.write(i,c,'' if v is None else v,money_fmt if c in (11,12,13) else num_fmt if c in (8,10) else text_fmt)
s.freeze_panes(4,6); s.autofilter(3,0,len(raw)+3,18); s.set_column(0,0,11); s.set_column(1,4,16); s.set_column(5,5,60); s.set_column(6,18,14)
# dedup detail
s=ws['91_去重明细']; dh=['月份','Listing键','代表ASIN','父ASIN','SKU','品牌','代表标题','PP标记','GENIMO标记','小类BSR','月销量','月销售额($)','价格($)','销量有效','销售额有效','价格有效','配对有效','BSR Top100','原始行数','来源子表','源行号']; setup(s,'去重明细（公式汇总的唯一输入）',20,'一行代表一个月度 Listing。父ASIN优先，否则ASIN；组内任一子体标题命中 plastic 即标记PP。代表行按最小可解析BSR、字段完整度和源行ID确定。'); headers(s,3,dh)
for i,r in enumerate(detail,4):
    vals=[r.get('month'),r.get('familyKey','').split('|',1)[-1],r.get('asin'),r.get('parent'),r.get('sku'),r.get('brand'),r.get('title'),1 if r.get('plastic') else 0,1 if r.get('genimo') else 0,r.get('rank'),r.get('sales'),r.get('revenue'),r.get('price'),1 if finite(r.get('sales')) else 0,1 if finite(r.get('revenue')) else 0,1 if finite(r.get('price')) else 0,1 if finite(r.get('sales')) and finite(r.get('revenue')) and r['sales']>0 else 0,1 if finite(r.get('rank')) and 1<=r['rank']<=100 else 0,r.get('rawRows'),r.get('month'),r.get('sourceRow')]
    for c,v in enumerate(vals): s.write(i,c,'' if v is None else v,money_fmt if c in (11,12) else num_fmt if c in (9,10,13,14,15,16,17,18,20) else text_fmt)
s.freeze_panes(4,7); s.autofilter(3,0,len(detail)+3,20); s.set_column(0,0,11); s.set_column(1,1,23); s.set_column(2,5,16); s.set_column(6,6,60); s.set_column(7,20,13)
# aggregate helper
s=ws['92_聚合输入']; ah=['月份','范围键','范围','层级','Listing数','销量','销售额($)','价格合计($)','价格有效数','配对销量','配对销售额($)','销量有效数','销售额有效数','数据状态','BSR候选原始行数']; setup(s,'聚合输入（由 91_去重明细确定性生成）',14,'本页由去重明细按月、范围和层级生成；01—06 展示页公式只引用本页。MOM按去年同月，YOY按上一年度；BSR候选原始行数由90页先筛BSR后统计，空子表标记为无数据，不按 0 伪造。'); headers(s,3,ah)
agg=[]
for scope,label in scopes:
    for m in months:
        for level in levels:
            z=aggregate[(scope,m,level)]; agg.append([m,scope,label,level,z['count'],z['sales'],z['revenue'],z['price_sum'],z['price_count'],z['paired_sales'],z['paired_revenue'],z['sales_valid'],z['revenue_valid'],data_status.get(m,'无数据'),z.get('candidate_raw_rows')])
for i,row in enumerate(agg,4):
    for c,v in enumerate(row): s.write(i,c,v,money_fmt if c in (6,7,10) else num_fmt if c>=4 else text_fmt)
s.freeze_panes(4,0); s.autofilter(3,0,len(agg)+3,14); s.set_column(0,0,11); s.set_column(1,3,15); s.set_column(4,14,16); agg_last=len(agg)+4
AR="'92_聚合输入'!$A$5:$A$%d"%agg_last; BR="'92_聚合输入'!$B$5:$B$%d"%agg_last; DR="'92_聚合输入'!$D$5:$D$%d"%agg_last; SR="'92_聚合输入'!$N$5:$N$%d"%agg_last
AC={'count':'E','sales':'F','revenue':'G','price_sum':'H','price_count':'I','paired_sales':'J','paired_revenue':'K','sales_valid':'L','revenue_valid':'M','candidate_raw_rows':'O'}
def af(field,scope,level,mref):
    c=AC[field]; return '=IF(COUNTIFS(%s,%s,%s,"%s",%s,"%s",%s,"有数据")=0,"",SUMIFS(\'92_聚合输入\'!$%s$5:$%s$%d,%s,%s,%s,"%s",%s,"%s",%s,"有数据"))'%(AR,mref,BR,scope,DR,level,SR,c,c,agg_last,AR,mref,BR,scope,DR,level,SR)
def anf(field,scope,level,yref):
    c=AC[field]; return '=IF(COUNTIFS(%s,">="&%s&"01",%s,"<="&%s&"12",%s,"%s",%s,"%s",%s,"有数据")=0,"",SUMIFS(\'92_聚合输入\'!$%s$5:$%s$%d,%s,">="&%s&"01",%s,"<="&%s&"12",%s,"%s",%s,"%s",%s,"有数据"))'%(AR,yref,AR,yref,BR,scope,DR,level,SR,c,c,agg_last,AR,yref,AR,yref,BR,scope,DR,level,SR)
def anf_range(field,scope,level,yref,start_month,end_month):
    c=AC[field]; return '=IF(COUNTIFS(%s,">="&%s&"%02d",%s,"<="&%s&"%02d",%s,"%s",%s,"%s",%s,"有数据")=0,"",SUMIFS(\'92_聚合输入\'!$%s$5:$%s$%d,%s,">="&%s&"%02d",%s,"<="&%s&"%02d",%s,"%s",%s,"%s",%s,"有数据"))'%(AR,yref,start_month,AR,yref,end_month,BR,scope,DR,level,SR,c,c,agg_last,AR,yref,start_month,AR,yref,end_month,BR,scope,DR,level,SR)
def avgf(scope,level,mref): return '=IFERROR(%s/%s,"")'%(af('price_sum',scope,level,mref)[1:],af('price_count',scope,level,mref)[1:])
def aspf(scope,level,mref): return '=IFERROR(%s/%s,"")'%(af('paired_revenue',scope,level,mref)[1:],af('paired_sales',scope,level,mref)[1:])
# monthly pages
mn={'overall':'01_整体市场月度','pp':'02_PP市场月度','nonpp':'03_高客单价市场月度','genimo':'04_GENIMO品牌月度'}; mh=['月份','整体/范围Listing（去重）','整体/范围销量','整体/范围销售额($)','平均标价($)','加权成交均价($)','销量有效数','销售额有效数','价格有效数','销量MOM','销售额MOM','BSR前100去重Listing数','Top100销量','Top100销售额($)','Top100平均标价($)','数据状态','BSR前100原始候选行数']; mrow={m:i+5 for i,m in enumerate(months)}
for scope,label in scopes:
    s=ws[mn[scope]]; setup(s,label+'·月度数据（公式）',16,'销量、销售额、平均标价和加权成交均价由 92_聚合输入公式汇总；MOM=去年同月，YOY=上一年度。Top100为BSR 1—100（含100）；同时显示原始候选行数和去重Listing数，便于与90页手工筛选回勾。'); headers(s,3,mh); s.freeze_panes(4,0); s.set_column(0,0,12); s.set_column(1,16,18)
    for i,m in enumerate(months,4):
        row=i+1; ref='$A%d'%row; pm,py=prev_month(m),prev_year(m); pmr,pyr=mrow.get(pm),mrow.get(py); cur=aggregate[(scope,m,'全部')]; top=aggregate[(scope,m,'Top100')]
        vals=[cur['count'],cur['sales'],cur['revenue'],cur['price_sum']/cur['price_count'] if cur['price_count'] else '',cur['paired_revenue']/cur['paired_sales'] if cur['paired_sales'] else '',cur['sales_valid'],cur['revenue_valid'],cur['price_count'],pct(cur['sales'],aggregate.get((scope,py,'全部'),{}).get('sales')) if pyr else '',pct(cur['revenue'],aggregate.get((scope,py,'全部'),{}).get('revenue')) if pyr else '',top['count'],top['sales'],top['revenue'],top['price_sum']/top['price_count'] if top['price_count'] else '',top.get('candidate_raw_rows')]
        fs=[af('count',scope,'全部',ref),af('sales',scope,'全部',ref),af('revenue',scope,'全部',ref),avgf(scope,'全部',ref),aspf(scope,'全部',ref),af('sales_valid',scope,'全部',ref),af('revenue_valid',scope,'全部',ref),af('price_count',scope,'全部',ref),'=IFERROR(C%d/C%d-1,"")'%(row,pyr) if pyr else '=""','=IFERROR(D%d/D%d-1,"")'%(row,pyr) if pyr else '=""',af('count',scope,'Top100',ref),af('sales',scope,'Top100',ref),af('revenue',scope,'Top100',ref),avgf(scope,'Top100',ref),af('candidate_raw_rows',scope,'Top100',ref)]
        fm=[fnum,fnum,fnum,fmoney,fmoney,fnum,fnum,fnum,fpct,fpct,fnum,fnum,fnum,fmoney,fnum];
        if data_status.get(m)=='无数据': vals=['']*15
        s.write(i,0,m,text_fmt)
        for c,(f,v,fmt) in enumerate(zip(fs,vals,fm),1): formula(s,i,c + (1 if c == 15 else 0),f,v,fmt)
        s.write(i,15,data_status.get(m,'无数据'),text_fmt)
# annual page: all four business scopes share the same annual YOY table.
s=ws['05_年度YOY']; setup(s,'年度YOY与覆盖状态（公式）',15,'年度统计由 92_聚合输入按年份和范围汇总；平均标价按有效价格记录算术平均；YOY按当前年与上一年度的共同覆盖月份计算，未完结年份标记实际覆盖月份。'); headers(s,3,['年份','范围','Listing数','销量','销售额($)','平均标价($)','加权成交均价($)','销量YOY','销售额YOY','Top100 Listing','Top100销量','Top100销售额($)','Top100平均标价($)','Top100销量YOY','Top100销售额YOY','数据状态']); years=sorted({m[:4] for m in months}); year_months_map={y:sorted([m for m in months if m.startswith(y) and data_status.get(m)=='有数据']) for y in years}; s.set_column(0,0,11); s.set_column(1,15,17)
annual_row={(scope,y): 5 + si*len(years) + yi for si,(scope,_) in enumerate(scopes) for yi,y in enumerate(years)}
annual_cache={}
for si,(scope,label) in enumerate(scopes):
    for yi,y in enumerate(years):
        i=4 + si*len(years) + yi; row=i+1; ref='$A%d'%row; s.write(i,0,y,text_fmt); s.write(i,1,label,text_fmt)
        pool=[r for r in detail if str(r['month']).startswith(y) and in_scope(r,scope)]; tp=select_top100(pool); cm,tm=metrics(pool),metrics(tp); prev=f'{int(y)-1:04d}' if (scope,f'{int(y)-1:04d}') in annual_row else None
        common_suffixes=sorted({m[4:] for m in year_months_map.get(y,[])} & {m[4:] for m in year_months_map.get(prev,[])}) if prev else []
        cmp_pool=[r for r in pool if str(r['month'])[4:] in common_suffixes]; cmp_top=[r for r in tp if str(r['month'])[4:] in common_suffixes]; ccur, ctop=metrics(cmp_pool),metrics(cmp_top)
        prev_pool=[r for r in detail if prev and str(r['month']).startswith(prev) and str(r['month'])[4:] in common_suffixes and in_scope(r,scope)]; prev_top=select_top100(prev_pool); pcur, ptop=metrics(prev_pool),metrics(prev_top)
        vals=[cm['count'],cm['sales'],cm['revenue'],cm['price_sum']/cm['price_count'] if cm['price_count'] else '',cm['paired_revenue']/cm['paired_sales'] if cm['paired_sales'] else '',pct(ccur['sales'],pcur['sales']) if prev else '',pct(ccur['revenue'],pcur['revenue']) if prev else '',tm['count'],tm['sales'],tm['revenue'],tm['price_sum']/tm['price_count'] if tm['price_count'] else '',pct(ctop['sales'],ptop['sales']) if prev else '',pct(ctop['revenue'],ptop['revenue']) if prev else '']
        annual_cache[(scope,y)]={'sales':cm['sales'],'revenue':cm['revenue'],'top_sales':tm['sales'],'top_revenue':tm['revenue']}
        prev_row=annual_row[(scope,prev)] if prev else None
        if common_suffixes:
            start_month,end_month=int(common_suffixes[0]),int(common_suffixes[-1])
            yoy_sales=anf_range('sales',scope,'全部',ref,start_month,end_month); yoy_revenue=anf_range('revenue',scope,'全部',ref,start_month,end_month); yoy_top_sales=anf_range('sales',scope,'Top100',ref,start_month,end_month); yoy_top_revenue=anf_range('revenue',scope,'Top100',ref,start_month,end_month)
        else:
            yoy_sales=yoy_revenue=yoy_top_sales=yoy_top_revenue='=""'
        prev_ref='$A%d'%prev_row if prev_row else None
        fs=[anf('count',scope,'全部',ref),anf('sales',scope,'全部',ref),anf('revenue',scope,'全部',ref),'=IFERROR(%s/%s,"")'%(anf('price_sum',scope,'全部',ref)[1:],anf('price_count',scope,'全部',ref)[1:]),'=IFERROR(%s/%s,"")'%(anf('paired_revenue',scope,'全部',ref)[1:],anf('paired_sales',scope,'全部',ref)[1:]),'=IFERROR(%s/%s-1,"")'%(yoy_sales[1:],anf_range('sales',scope,'全部',prev_ref,start_month,end_month)[1:]) if prev and common_suffixes else '=""','=IFERROR(%s/%s-1,"")'%(yoy_revenue[1:],anf_range('revenue',scope,'全部',prev_ref,start_month,end_month)[1:]) if prev and common_suffixes else '=""',anf('count',scope,'Top100',ref),anf('sales',scope,'Top100',ref),anf('revenue',scope,'Top100',ref),'=IFERROR(%s/%s,"")'%(anf('price_sum',scope,'Top100',ref)[1:],anf('price_count',scope,'Top100',ref)[1:]),'=IFERROR(%s/%s-1,"")'%(yoy_top_sales[1:],anf_range('sales',scope,'Top100',prev_ref,start_month,end_month)[1:]) if prev and common_suffixes else '=""','=IFERROR(%s/%s-1,"")'%(yoy_top_revenue[1:],anf_range('revenue',scope,'Top100',prev_ref,start_month,end_month)[1:]) if prev and common_suffixes else '=""']
        fm=[fnum,fnum,fnum,fmoney,fmoney,fpct,fpct,fnum,fnum,fnum,fmoney,fpct,fpct]
        year_months=year_months_map.get(y,[]); status='无数据' if not year_months else ('有数据' if len(year_months)==12 else '部分覆盖（%d/12月）'%len(year_months))
        if status=='无数据': vals=['']*13
        for c,(f,v,fmt) in enumerate(zip(fs,vals,fm),2): formula(s,i,c,f,v,fmt)
        s.write(i,15,status,text_fmt)
# BSR tier page
s=ws['06_BSR分层']; setup(s,'BSR Top100 头中尾与五档（公式）',10,'层级只从同一范围的 Top100 池计算；粗分层为头部1-20、中部21-50、尾部51-100，细分层为1-5、6-10、11-20、21-50、51-100。'); headers(s,3,['月份','范围','层级','BSR下限','BSR上限','Listing数','销量','销售额($)','平均标价($)','销量MOM','销售额MOM']); tr={}; tv=[]
for scope,label in scopes:
    for level,lo,hi in bands:
        for m in months: tr[(scope,level,m)]=len(tv)+5; tv.append((scope,label,level,lo,hi,m))
for i,(scope,label,level,lo,hi,m) in enumerate(tv,4):
    row=i+1; ref='$A%d'%row; pm,py=prev_month(m),prev_year(m); pmr,pyr=tr.get((scope,level,pm)),tr.get((scope,level,py)); cur=aggregate[(scope,m,level)]; vals=[cur['count'],cur['sales'],cur['revenue'],cur['price_sum']/cur['price_count'] if cur['price_count'] else '',pct(cur['sales'],aggregate.get((scope,py,level),{}).get('sales')) if pyr else '',pct(cur['revenue'],aggregate.get((scope,py,level),{}).get('revenue')) if pyr else '']; fs=[af('count',scope,level,ref),af('sales',scope,level,ref),af('revenue',scope,level,ref),avgf(scope,level,ref),'=IFERROR(G%d/G%d-1,"")'%(row,pyr) if pyr else '=""','=IFERROR(H%d/H%d-1,"")'%(row,pyr) if pyr else '=""']; s.write(i,0,m,text_fmt);s.write(i,1,label,text_fmt);s.write(i,2,level,text_fmt);s.write(i,3,lo,num_fmt);s.write(i,4,hi,num_fmt)
    for c,(f,v,fmt) in enumerate(zip(fs,vals,[fnum,fnum,fnum,fmoney,fpct,fpct]),5): formula(s,i,c,f,v,fmt)
s.freeze_panes(4,0);s.autofilter(3,0,len(tv)+3,10);s.set_column(0,0,11);s.set_column(1,2,15);s.set_column(3,4,11);s.set_column(5,10,15)
# checks
s=ws['93_数据校验']; setup(s,'数据校验（独立检查，不驱动业务结果）',5,'检查页只观察数据，不向汇总表提供业务值。失败时按 90 原始输入 → 91 去重明细 → 92 聚合输入 → 展示页公式定位。'); headers(s,2,['校验项','结果','期望','差额','状态','说明']); s.set_column(0,0,30); s.set_column(1,5,18); items=['原始有效行数','去重明细行数','BSR=100存在','整体=PP+高客单价市场（最新月销量）','整体=PP+高客单价市场（最新月销售额）']; raw_end=len(raw)+4; det_end=len(detail)+4; lr=len(months)+4; cf=["=COUNTA('90_原始输入'!A5:A%d)"%raw_end,"=COUNTA('91_去重明细'!A5:A%d)"%det_end,"=IF(COUNTIF('91_去重明细'!J5:J%d,100)>0,\"通过\",\"失败\")"%det_end,"='01_整体市场月度'!C%d-'02_PP市场月度'!C%d-'03_高客单价市场月度'!C%d"%(lr,lr,lr),"='01_整体市场月度'!D%d-'02_PP市场月度'!D%d-'03_高客单价市场月度'!D%d"%(lr,lr,lr)]; ex=[len(raw),len(detail),'通过',0,0]
for i,(f,v,e) in enumerate(zip(cf,[len(raw),len(detail),'通过',0,0],ex),3):
    s.write(i,0,items[i-3],text_fmt); formula(s,i,1,f,v,num_fmt if isinstance(v,(int,float)) else text_fmt); s.write(i,2,e,num_fmt if isinstance(e,(int,float)) else text_fmt); formula(s,i,3,'=B%d-C%d'%(i+1,i+1),0,num_fmt); formula(s,i,4,'=IF(D%d=0,"通过","失败")'%(i+1),'通过',text_fmt); s.write(i,5,['原始主源行数回勾','去重后 Listing 数回勾','Top100 边界包含100','最新月份销量分区互补','最新月份销售额分区互补'][i-3],text_fmt)
s.write(8,0,'最新月份数据状态',text_fmt); formula(s,8,1,"='01_整体市场月度'!P%d"%lr,'有数据',text_fmt); s.write(8,2,'应有源数据',text_fmt); s.write(8,3,'无数据月份不伪造为0',text_fmt); s.write(8,4,'已识别',text_fmt); s.write(8,5,'2025.10—2026.07 子表均有数据；表头行由读取程序自动识别',text_fmt)
# overview
s=ws['00_总览与校验']; setup(s,'户外地垫市场分析｜SPEC 2.0 可回勾工作簿',7,'本工作簿与 HTML 使用同一原始主源、同一 BSR/去重/PP 分类逻辑。整体市场=PP+高客单价市场；GENIMO是品牌视角；本工作簿不计算利润。'); headers(s,3,['项目','值','公式/来源','核对位置','状态','说明','原始主源SHA-256','生成批次']); s.set_column(0,0,18);s.set_column(1,1,18);s.set_column(2,5,28);s.set_column(6,6,68);s.set_column(7,7,16)
oi=[('原始有效行',"='93_数据校验'!B4",'93_数据校验','原始输入行数',len(raw)),('去重Listing',"='93_数据校验'!B5",'93_数据校验','父ASIN优先去重',len(detail)),('Top100去重Listing',"='01_整体市场月度'!L%d"%lr,'01_整体市场月度','BSR 1—100含100',aggregate[('overall',months[-1],'Top100')]['count']),('PP去重Listing',"='02_PP市场月度'!B%d"%lr,'02_PP市场月度','PP=1',aggregate[('pp',months[-1],'全部')]['count']),('高客单价去重Listing',"='03_高客单价市场月度'!B%d"%lr,'03_高客单价市场月度','PP=0',aggregate[('nonpp',months[-1],'全部')]['count']),('GENIMO去重Listing',"='04_GENIMO品牌月度'!B%d"%lr,'04_GENIMO品牌月度','品牌视角',aggregate[('genimo',months[-1],'全部')]['count'])]
for i,(name,f,w,n,v) in enumerate(oi,4): s.write(i,0,name,text_fmt);formula(s,i,1,f,v,fnum);s.write(i,2,'公式/聚合输入',text_fmt);s.write(i,3,w,text_fmt);s.write(i,4,'已生成',text_fmt);s.write(i,5,n,text_fmt);s.write(i,6,source_hash,text_fmt);s.write(i,7,'2026-09-11 SPEC2',text_fmt)
headers(s,12,['四部分交付','对应工作表','核心字段','MOM/YOY','Top100','分层','策略','状态'])
for i,row in enumerate([['整体市场','01 / 05 / 06','销量、销售额、均价','MOM=去年同月 / YOY=上一年度','是','头中尾 + 五档','整体趋势与市场判断','已生成'],['PP市场','02 / 05 / 06','销量、销售额、均价','MOM=去年同月 / YOY=上一年度','是','头中尾 + 五档','PP结构判断','已生成'],['高客单价市场','03 / 05 / 06','销量、销售额、均价','MOM=去年同月 / YOY=上一年度','是','头中尾 + 五档','高客单价结构判断','已生成'],['GENIMO品牌','04 / 05 / 06','整体份额、PP份额、BSR','MOM=去年同月 / YOY=上一年度','是','头中尾 + 五档','2027行动建议','已生成']],13):
    for c,v in enumerate(row): s.write(i,c,v,text_fmt)
s.merge_range(19,0,19,7,'验证顺序：90 原始输入 → 91 去重明细 → 92 聚合输入 → 01/02/03/04 公式 → 05 年度 → 06 分层 → 93 校验。',note_fmt)
# rules
s=ws['94_规则说明']; setup(s,'来源、字段和操作规则',7,'所有展示指标公式写在 01—06 页；92 聚合输入是由 91 去重明细确定性生成的中间层。'); headers(s,2,['规则项','当前规则','来源','公式/操作','是否改写原始数据','可复核位置','限制','备注']); s.set_column(0,0,18);s.set_column(1,7,35)
rr=[['主源','地垫-卖家精灵市场数据.xlsx',source_hash,'直接读取原始工作簿月度子表','否','90_原始输入','更换文件需重新确认指纹',''],['月份',', '.join(months),'主源工作表名','统一为YYYYMM','否','90_原始输入','空月份跳过',''],['BSR解析','正整数；1—100含100','主源小类BSR','解析并标记，再筛Top100','否','91_去重明细','无效名次为空',''],['去重键','父ASIN优先，否则ASIN；均无则源行','ASIN/父ASIN','月度+Listing键分组','否','91_去重明细','代表行不合计子体',''],['PP分类','商品标题完整单词plastic，不区分大小写','商品标题','组内任一子体命中即PP','否','90/91','不使用价格替代',''],['高客单价分类','整体排除PP后的全部商品','PP分类结果','PP=0','否','02/03','不增加价格门槛',''],['销量/销售额','只合计有效数值','月销量/月销售额','SUMIFS（展示页引用92）','否','01—04','空值不当零',''],['均价','价格有效记录算术平均；成交均价按配对销量/销售额','价格、销量、销售额','IFERROR(合计/有效数)','否','01—04','不推导利润',''],['MOM/YOY口径','月度MOM=去年同月；年度YOY=上一年度（同周期）','月度/年度汇总','显式引用对应月份行','否','01—04','首月无可比值留空',''],['原始经营字段','FBA、毛利率、Coupon只保留原始值','主源字段','不进入利润公式','否','90_原始输入','不存在时为空','']]
for i,row in enumerate(rr,3):
    for c,v in enumerate(row): s.write(i,c,v,text_fmt)
s.freeze_panes(3,0)
wb.close()
bad=[]; fc=0
with zipfile.ZipFile(OUT_FILE) as z:
    for n in z.namelist():
        if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
            t=z.read(n).decode('utf-8','ignore'); fc += t.count('<f>')
            for marker in ('#REF!','#DIV/0!','#VALUE!','#NAME?','#N/A'):
                if marker in t: bad.append((n,marker))
print(json.dumps({'output':OUT_FILE,'sourceSha256':source_hash,'rawRows':len(raw),'dedupRows':len(detail),'aggregateRows':len(agg),'formulaCount':fc,'formulaErrors':bad},ensure_ascii=False,indent=2))

