"""Monthly campus-card screening. Integer cents; one source transaction per row.
No spreadsheet/UI dependencies: this engine also powers the command-line workflow.
"""
from __future__ import annotations
import calendar, collections, datetime as dt, hashlib, json, re, sqlite3, zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from contextlib import contextmanager

@contextmanager
def database(path):
    connection=sqlite3.connect(path)
    try:
        with connection:yield connection
    finally:connection.close()

VERSION = '3.1.0-web'
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
CATEGORIES = {'学生食堂餐饮类':'餐饮','学生食堂餐饮类（餐饮）':'餐饮','餐饮':'餐饮',
 '学生食堂饮品水果类':'饮品','学生食堂饮品水果类（饮品）':'饮品','饮品':'饮品',
 '商服及浴池相关类':'商服','商服及浴池相关类（商服）':'商服','商服':'商服'}
DEFAULT_RULES = dict(meal_count=50,meal_amount=750,drink_count=10,drink_amount=50,
 service_count=10,service_amount=200,missing_zero=False,deduplicate=False,
 daily_basis='active',extra_drink_average=True)

def text(v):
    return '' if v is None else str(v).strip()

def iter_xlsx(path):
    """Read stored cell values, never execute formulas/macros or trust sheet dimensions."""
    with zipfile.ZipFile(path) as z:
        if sum(x.file_size for x in z.infolist()) > 2_000_000_000:
            raise ValueError('工作簿解压后超过2GB，请按月份拆分。')
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as stream:
                for _,e in ET.iterparse(stream,events=('end',)):
                    if e.tag==NS+'si':
                        shared.append(''.join(t.text or '' for t in e.iter(NS+'t')));e.clear()
        rels={e.attrib['Id']:e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        root=ET.fromstring(z.read('xl/workbook.xml'))
        prop=root.find(NS+'workbookPr')
        epoch1904=prop is not None and prop.attrib.get('date1904') in ('1','true')
        for sh in root.find(NS+'sheets'):
            rid=sh.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
            target=rels[rid]
            target=target.lstrip('/') if target.startswith('/') else str(PurePosixPath('xl')/target)
            with z.open(target) as stream:
                tree=ET.iterparse(stream,events=('start','end'));_,doc=next(tree)
                sheetdata=None
                for event,e in tree:
                    if event=='start' and e.tag==NS+'sheetData':sheetdata=e
                    if event!='end' or e.tag!=NS+'row':continue
                    values={}
                    for c in e:
                        if c.tag!=NS+'c':continue
                        letters=re.match(r'[A-Z]+',c.attrib.get('r','A1')).group()
                        col=0
                        for ch in letters:col=col*26+ord(ch)-64
                        typ=c.attrib.get('t');v=c.find(NS+'v');value=v.text if v is not None else None
                        if typ=='s' and value is not None:value=shared[int(value)]
                        elif typ=='inlineStr':value=''.join(t.text or '' for t in c.iter(NS+'t'))
                        elif typ=='e':value='#EXCEL_ERROR:'+str(value)
                        values[col-1]=value
                    if any(text(v) for v in values.values()):
                        yield sh.attrib['name'],int(e.attrib.get('r',0)),[values.get(i) for i in range(max(values)+1)],epoch1904
                    e.clear()
                    if sheetdata is not None:sheetdata.clear()

def parse_date(value,epoch1904=False):
    s=text(value)
    if not s:return None
    m=re.match(r'^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})',s)
    if m:
        try:return dt.date(*map(int,m.groups()))
        except ValueError:return None
    try:
        n=float(s)
        if 1<n<150000:return (dt.datetime(1904,1,1) if epoch1904 else dt.datetime(1899,12,30))+dt.timedelta(days=int(n))
    except ValueError:pass
    return None

def cents(value):
    try:
        v=Decimal(text(value))
        if not v.is_finite() or v*100!=(v*100).to_integral_value():raise ValueError()
        if abs(v)>Decimal('1000000000'):raise ValueError()
        return int(v*100)
    except (InvalidOperation,ValueError):raise ValueError('金额为空、非数字或精度超过分')

def ingest(path,db_path,progress=lambda **kw:None):
    path,db_path=Path(path),Path(db_path)
    if db_path.exists():raise ValueError('缓存路径已存在，请使用新的处理目录。')
    conn=sqlite3.connect(db_path)
    conn.execute('PRAGMA journal_mode=OFF');conn.execute('PRAGMA synchronous=OFF')
    conn.executescript('CREATE TABLE tx(month TEXT,day TEXT,card TEXT,account TEXT,category TEXT,raw_category TEXT,merchant TEXT,amount INTEGER,duplicate INTEGER,source_sheet TEXT,source_row INTEGER); CREATE TABLE errors(sheet TEXT,rownum INTEGER,reason TEXT,month TEXT);')
    headers={};seen=set();counts=collections.Counter();months=collections.Counter();cats=collections.Counter();merchants=collections.Counter();types=collections.Counter();batch=[];errors=[];mismatch=0;first=None;last=None
    digest=hashlib.sha256()
    with path.open('rb') as src:
        for b in iter(lambda:src.read(1024*1024),b''):digest.update(b)
    try:
        for sheet,rownum,row,epoch in iter_xlsx(path):
            if sheet not in headers:
                normalized=[text(v) for v in row]
                if '一卡通号' not in normalized or '交易额' not in normalized:
                    if rownum<=20:continue
                    raise ValueError(f'{sheet}前20行未找到一卡通号和交易额表头。')
                if len(normalized)!=len(set(normalized)):raise ValueError(f'{sheet}有重复列名，请整理表头。')
                if not ({'消费日期','入帐时间'} & set(normalized)) or '消费场所类别' not in normalized:raise ValueError(f'{sheet}缺少消费日期/入帐时间或消费场所类别。三类筛选不能从POS号猜分类。')
                headers[sheet]=normalized;continue
            h=headers[sheet];d=dict(zip(h,row));counts['input_rows']+=1
            date=parse_date(d.get('消费日期'),epoch) or parse_date(d.get('入帐时间'),epoch)
            month=date.strftime('%Y-%m') if date else None
            card=text(d.get('一卡通号'));category=text(d.get('消费场所类别'));merchant=text(d.get('对方账号（商户名称）',d.get('商户名称','')))
            try:
                if not date:raise ValueError('无有效消费日期')
                if not card or card.lower() in ('nan','none','null'):raise ValueError('一卡通号缺失')
                amount=cents(d.get('交易额'))
                stated=text(d.get('消费月份'))
                mt=re.search(r'(\d{1,2})\s*月',stated)
                if mt and int(mt.group(1))!=date.month:mismatch+=1
                date2=parse_date(d.get('入帐时间'),epoch)
                if date2 and date2.strftime('%Y-%m-%d')!=date.strftime('%Y-%m-%d'):counts['date_mismatch']+=1
            except ValueError as e:
                errors.append((sheet,rownum,str(e),month));counts['invalid_rows']+=1;continue
            key=hashlib.sha256(json.dumps([text(v) for v in row],ensure_ascii=False,separators=(',',':')).encode()).digest()
            duplicate=int(key in seen);seen.add(key)
            counts['duplicate_rows']+=duplicate;months[month]+=1;cats[(month,category)]+=1;types[(month,text(d.get('消费类型')))]+=1
            if amount<0:merchants[(month,category,merchant)]+=1
            day=date.strftime('%Y-%m-%d');first=min(first or day,day);last=max(last or day,day)
            batch.append((month,day,card,text(d.get('帐号')),CATEGORIES.get(category,''),category,merchant,amount,duplicate,sheet,rownum))
            if len(batch)>=10000:
                conn.executemany('INSERT INTO tx VALUES (?,?,?,?,?,?,?,?,?,?,?)',batch);batch=[]
                progress(stage='读取流水',rows=counts['input_rows'])
        if not headers:raise ValueError('没有找到可读取的消费流水表。')
        conn.executemany('INSERT INTO tx VALUES (?,?,?,?,?,?,?,?,?,?,?)',batch)
        conn.executemany('INSERT INTO errors VALUES (?,?,?,?)',errors)
        conn.execute('CREATE INDEX tx_month ON tx(month)')
        conn.execute('CREATE INDEX tx_card ON tx(card)')
        meta={'version':VERSION,'source_name':path.name,'source_sha256':digest.hexdigest(),'counts':dict(counts),'months':dict(sorted(months.items())),'date_start':first,'date_end':last,'month_label_mismatch':mismatch,'categories':[{'month':m,'category':c,'rows':v} for (m,c),v in sorted(cats.items())],'merchants':[{'month':m,'category':c,'merchant':n,'rows':v} for (m,c,n),v in sorted(merchants.items())],'types':[{'month':m,'type':t,'rows':v} for (m,t),v in sorted(types.items())]}
        conn.execute('CREATE TABLE meta(payload TEXT)');conn.execute('INSERT INTO meta VALUES (?)',(json.dumps(meta,ensure_ascii=False),));conn.commit()
        return meta
    finally:conn.close()

def read_meta(db):
    with database(db) as c:return json.loads(c.execute('SELECT payload FROM meta').fetchone()[0])

def read_roster(path):
    mapping={};headers={}
    for sheet,r,row,_ in iter_xlsx(path):
        if sheet not in headers:
            if '一卡通号' in row:headers[sheet]=[text(v) for v in row]
            continue
        d=dict(zip(headers[sheet],row));card=text(d.get('一卡通号'))
        if not card:continue
        data=[text(d.get(k)) for k in ['学号','姓名','学院','年级']]
        if card in mapping and mapping[card]!=data:raise ValueError(f'学生名册存在同一卡号信息冲突（{sheet}第{r}行），请核对。')
        mapping[card]=data
    if not headers:raise ValueError('学生名册必须有一卡通号列。')
    return mapping

def evaluate(meal,drink,service,rules):
    return [meal['count']>rules['meal_count'],meal['cents']<cents(rules['meal_amount']),
            drink['count']<rules['drink_count'],drink['cents']<cents(rules['drink_amount']),
            service['count']<rules['service_count'],service['cents']<cents(rules['service_amount'])]

def analyze(db,month,rules=None,roster=None,category_mapping=None,supermarkets=None):
    cfg=DEFAULT_RULES|dict(rules or {});cfg['missing_zero']=False;roster=roster or {};meta=read_meta(db)
    if month not in meta['months']:raise ValueError('所选月份没有数据。')
    for k in ['meal_count','drink_count','service_count']:
        if isinstance(cfg[k],bool) or int(cfg[k])!=cfg[k] or cfg[k]<0:raise ValueError('次数阈值必须是非负整数。')
    for k in ['meal_amount','drink_amount','service_amount']:
        if cents(cfg[k])<0:raise ValueError('金额阈值不能为负。')
    if cfg['daily_basis'] not in ('active','calendar'):raise ValueError('日均口径无效。')
    mapping=CATEGORIES|dict(category_mapping or {})
    if any(v not in ('餐饮','饮品','商服','忽略') for v in mapping.values()):raise ValueError('分类映射只能为餐饮、饮品、商服或忽略。')
    supermarket_set=set(supermarkets) if supermarkets is not None else None
    con=sqlite3.connect(db)
    invalid=con.execute('SELECT COUNT(*) FROM errors WHERE month=? OR month IS NULL',(month,)).fetchone()[0]
    if invalid:
        con.close();raise ValueError(f'有{invalid}条所选月或日期不明的无效记录，请先在数据检查中核实并修正源表，不能静默筛选。')
    where='month=?'+(' AND duplicate=0' if cfg['deduplicate'] else '')
    unknown=con.execute(f'SELECT raw_category,COUNT(*) FROM tx WHERE {where} AND amount<0 GROUP BY raw_category',(month,)).fetchall()
    missing=[(k,n) for k,n in unknown if k not in mapping]
    if missing:
        con.close();raise ValueError('请先设置未识别分类：'+', '.join(f'{k or "空分类"}（{n}笔）' for k,n in missing))
    # Query by original classification and merchant; distinct days are reconciled across merchants below.
    rows=con.execute(f'SELECT card,raw_category,merchant,COUNT(*),-SUM(amount) FROM tx WHERE {where} AND amount<0 GROUP BY card,raw_category,merchant',(month,)).fetchall()
    def empty():return {'count':0,'cents':0,'days':set(),'super_cents':0}
    people={};ignored_count=ignored_cents=0;selected_super=set()
    for card,cat,merchant,n,total in rows:
        category=mapping[cat]
        if category=='忽略':ignored_count+=n;ignored_cents+=total;continue
        if card not in people:people[card]={c:empty() for c in ('餐饮','饮品','商服')}
        g=people[card][category];g['count']+=n;g['cents']+=total
        issuper=merchant in supermarket_set if supermarket_set is not None else '超市' in merchant
        if category=='商服' and issuper:g['super_cents']+=total;selected_super.add(merchant)
    for card,rawcat,day in con.execute(f'SELECT DISTINCT card,raw_category,day FROM tx WHERE {where} AND amount<0',(month,)):
        cat=mapping[rawcat]
        if cat!='忽略':people[card][cat]['days'].add(day)
    signs=con.execute('SELECT CASE WHEN amount<0 THEN \'支出\' WHEN amount>0 THEN \'正向\' ELSE \'零金额\' END,COUNT(*),SUM(amount) FROM tx WHERE month=? GROUP BY 1',(month,)).fetchall()
    dups=con.execute('SELECT COUNT(*) FROM tx WHERE month=? AND duplicate=1',(month,)).fetchone()[0]
    dates=con.execute('SELECT MIN(day),MAX(day),COUNT(DISTINCT day) FROM tx WHERE month=?',(month,)).fetchone()
    con.close()
    year,mon=map(int,month.split('-'));calendar_days=calendar.monthrange(year,mon)[1]
    tabs={'Sheet1':[],'Sheet2':[],'Sheet3':[],'Sheet4-汇总表':[]};audit=[]
    for card,groups in people.items():
        identity=[card]+roster.get(card,['','','',''])
        meal,drink,service=[groups[k] for k in ['餐饮','饮品','商服']]
        for g in groups.values():
            g['days']=len(g['days']);g['amount']=g['cents']/100
            g['average']=g['amount']/g['count'] if g['count'] else None
            denominator=g['days'] if cfg['daily_basis']=='active' else calendar_days
            g['daily']=g['amount']/denominator if denominator else None
        if meal['count']>0:tabs['Sheet1'].append(identity+[meal['count'],meal['amount'],meal['average'],meal['days'],meal['daily']])
        dr=identity+[drink['count'],drink['amount'],drink['days'],drink['daily']]
        if cfg['extra_drink_average']:dr.append(drink['average'])
        if drink['count']>0:tabs['Sheet2'].append(dr)
        if service['count']>0:tabs['Sheet3'].append(identity+[service['amount'],service['count'],service['super_cents']/100])
        checks=evaluate(meal,drink,service,cfg)
        coverage=all(g['count']>0 for g in groups.values())
        passed=all(checks) and coverage
        if passed:tabs['Sheet4-汇总表'].append(identity+[meal['count'],meal['amount'],drink['count'],drink['amount'],service['amount'],service['super_cents']/100])
        audit.append({'一卡通号':card,'餐饮次数':meal['count'],'餐饮金额':meal['amount'],'饮品次数':drink['count'],'饮品金额':drink['amount'],'商服次数':service['count'],'商服金额':service['amount'],
                      '餐饮次数达标':checks[0],'餐饮金额达标':checks[1],'饮品次数达标':checks[2],'饮品金额达标':checks[3],'商服次数达标':checks[4],'商服金额达标':checks[5],
                      '类别记录条件满足':coverage,'无记录类别':'、'.join(k for k,g in groups.items() if not g['count']),'入选':passed})
    tabs['Sheet1'].sort(key=lambda r:(-r[5],r[0]));tabs['Sheet2'].sort(key=lambda r:(r[5],r[0]));tabs['Sheet3'].sort(key=lambda r:(r[5],r[0]));tabs['Sheet4-汇总表'].sort(key=lambda r:(-r[5],r[0]));audit.sort(key=lambda r:(not r['入选'],-r['餐饮次数'],r['一卡通号']))
    totals={k:{'count':sum(g[k]['count'] for g in people.values()),'cents':sum(g[k]['cents'] for g in people.values()),'cards_with_records':sum(g[k]['count']>0 for g in people.values())} for k in ['餐饮','饮品','商服']}
    summary={'version':VERSION,'month':month,'rules':cfg,'card_count':len(people),'selected_count':len(tabs['Sheet4-汇总表']),'category_totals':totals,'source_name':meta['source_name'],'source_sha256':meta['source_sha256'],
             'range':dates,'calendar_days':calendar_days,'signs':signs,'duplicates':dups,'duplicate_policy':'删除完全重复行' if cfg['deduplicate'] else '保留原始交易行',
             'ignored_count':ignored_count,'ignored_cents':ignored_cents,'roster_matches':sum(k in roster for k in people),'supermarket_merchants':sorted(selected_super),'supermarket_basis':'手动商户清单' if supermarket_set is not None else '商户名称含“超市”',
             'category_mapping':mapping,'three_table_cards':sum(all(g[c]['count']>0 for c in ('餐饮','饮品','商服')) for g in people.values()),'partial_month':dates[0]!=f'{month}-01' or dates[1]!=f'{month}-{calendar_days:02d}','generated_at':dt.datetime.now().astimezone().isoformat(timespec='seconds')}
    return {'summary':summary,'sheets':tabs,'audit':audit}

def transaction_trace(db,month,card,limit=500):
    with database(db) as c:
        c.row_factory=sqlite3.Row
        total=c.execute('SELECT COUNT(*) FROM tx WHERE month=? AND card=?',(month,card)).fetchone()[0]
        data=[dict(r) for r in c.execute('SELECT day,raw_category,merchant,amount,duplicate,source_sheet,source_row FROM tx WHERE month=? AND card=? ORDER BY day,source_row LIMIT ?',(month,card,limit))]
    return {'total':total,'rows':data}
