"""Browser-worker API. All input, SQLite cache and results live in this worker's memory."""
import json,shutil,zipfile,calendar
from pathlib import Path
from urllib.parse import urlparse,parse_qs
from engine import ingest,analyze,read_roster,transaction_trace,database,DEFAULT_RULES
from exporter import export_result
BASE=Path('/workspace');BASE.mkdir(exist_ok=True)
meta=None;result=None;roster={};serial=0
def dispatch(url,payload,filename,progress):
    global meta,result,roster,serial
    u=urlparse(url);q={k:v[0] for k,v in parse_qs(u.query).items()}
    if u.path=='/api/bootstrap':return {'defaults':DEFAULT_RULES,'last':None}
    if u.path=='/api/upload':
        result=None;meta=None;roster={}
        db=BASE/'cache.sqlite'
        if db.exists():db.unlink()
        try:meta=ingest(BASE/'incoming.xlsx',db,progress)
        finally:(BASE/'incoming.xlsx').unlink(missing_ok=True)
        meta['source_name']=filename
        meta['month_checks']={}
        with database(db) as c:
            for month in meta['months']:
                dates=c.execute('SELECT MIN(day),MAX(day) FROM tx WHERE month=?',(month,)).fetchone()
                y,m=map(int,month.split('-'));last=calendar.monthrange(y,m)[1]
                rows=c.execute('SELECT raw_category,COUNT(*),COUNT(DISTINCT card),-SUM(amount) FROM tx WHERE month=? AND amount<0 GROUP BY raw_category',(month,)).fetchall()
                meta['month_checks'][month]={'range':dates,'partial':dates[0]!=month+'-01' or dates[1]!=f'{month}-{last:02d}','categories':[{'category':cat,'count':n,'cards':cards,'cents':amount} for cat,n,cards,amount in rows]}
        batch=BASE/'全部月份结果.zip'
        batch.unlink(missing_ok=True)
        with database(db) as c:c.execute('UPDATE meta SET payload=?',(json.dumps(meta,ensure_ascii=False),))
        return {'session':'browser','meta':meta}
    if u.path=='/api/roster':
        try:loaded=read_roster(BASE/'incoming.xlsx')
        finally:(BASE/'incoming.xlsx').unlink(missing_ok=True)
        roster=loaded;return {'count':len(roster)}
    if u.path=='/api/analyze':
        if meta is None:raise ValueError('请先导入消费记录')
        progress(stage='按月份汇总与筛选')
        r=analyze(BASE/'cache.sqlite',payload['month'],payload.get('rules'),roster,payload.get('mapping'),payload.get('supermarkets'))
        progress(stage='生成Excel与完整结果包')
        out=BASE/'result'
        if out.exists():shutil.rmtree(out)
        export_result(r,out);result=r;serial+=1
        return {'session':'browser','run':str(serial),'summary':result['summary']}
    if u.path=='/api/analyze-all':
        if meta is None:raise ValueError('请先导入消费记录')
        final=BASE/'全部月份结果.zip';final.unlink(missing_ok=True)
        stage=BASE/'batch'
        if stage.exists():shutil.rmtree(stage)
        stage.mkdir()
        summaries=[]
        try:
            with zipfile.ZipFile(stage/'all.zip','w',zipfile.ZIP_DEFLATED) as z:
                for month in sorted(meta['months']):
                    progress(stage=f'正在整理 {month} 并生成四表')
                    r=analyze(BASE/'cache.sqlite',month,payload.get('rules'),roster,payload.get('mapping'),payload.get('supermarkets'))
                    out=stage/month;export_result(r,out);summaries.append(r['summary'])
                    for f in out.iterdir():
                        if f.suffix!='.zip':z.write(f,month+'/'+f.name)
                    shutil.rmtree(out)
                z.writestr('各月份核对.json',json.dumps(summaries,ensure_ascii=False,indent=2))
            (stage/'all.zip').replace(final)
        finally:shutil.rmtree(stage)
        return {'months':len(summaries),'selected':{s['month']:s['selected_count'] for s in summaries}}
    if u.path=='/api/result':
        if result is None:raise ValueError('请先生成结果')
        tab=q.get('tab','summary')
        if tab=='summary':return {'summary':result['summary']}
        rows=result['audit'] if tab=='依据' else result['sheets'][tab]
        if tab=='Sheet2' and not result['summary']['rules']['extra_drink_average']:rows=[r+[r[6]/r[5] if r[5] else None] for r in rows]
        search=q.get('search','').strip().lower()
        if search:rows=[r for r in rows if search in str(r.get('一卡通号','') if isinstance(r,dict) else r[:5]).lower()]
        page=max(1,int(q.get('page',1)));start=(page-1)*50
        return {'rows':rows[start:start+50],'total':len(rows),'page':page}
    if u.path=='/api/trace':return transaction_trace(BASE/'cache.sqlite',q['month'],q['card'])
    raise ValueError('未知操作')
def download_path(kind):
    if kind=='all':
        path=BASE/'全部月份结果.zip'
        if not path.exists():raise ValueError('请先生成全部月份结果')
        return str(path)
    if result is None:raise ValueError('请先生成结果')
    return str(BASE/'result'/('完整结果包.zip' if kind=='zip' else result['summary']['month']+'_数据整理表.xlsx'))
