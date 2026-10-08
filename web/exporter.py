"""Redistributable Excel writer; numeric results and cached formulas share engine values."""
import json,zipfile,shutil
from pathlib import Path
import xlsxwriter
from result_notes import write_csv,write_explanation

ROOT=Path(__file__).resolve().parent
def export_result(result,out):
    out=Path(out);out.mkdir(parents=True,exist_ok=True)
    headers=json.loads((ROOT/'template_headers.json').read_text(encoding='utf-8'))
    rules=result['summary']['rules']
    dest=out/(result['summary']['month']+'_数据整理表.xlsx')
    with xlsxwriter.Workbook(str(dest),{'constant_memory':True,'tmpdir':str(out),'strings_to_urls':False,'strings_to_formulas':False}) as book:
        base={'font_name':'Microsoft YaHei','font_size':10,'valign':'vcenter'}
        plain=book.add_format(base)
        text=book.add_format({**base,'num_format':'@'})
        number=book.add_format({**base,'num_format':'0'})
        money=book.add_format({**base,'num_format':'#,##0.00'})
        head=[book.add_format({**base,'bold':True,'text_wrap':True,'align':'center','border':1,'border_color':'#CBD5E1','bg_color':c}) for c in ['#FFFFFF','#FFFF00','#DDEBF7']]
        for name,rows in result['sheets'].items():
            names=list(headers[name]);extra=name=='Sheet2' and rules['extra_drink_average']
            if extra:names.append('每次平均消费额度')
            s=book.add_worksheet(name);s.freeze_panes(1,0);s.set_default_row(21);s.set_row(0,62)
            for i,width in enumerate([19,18,12,22,13]+[19]*(len(names)-5)):s.set_column(i,i,width)
            for i,label in enumerate(names):
                color=1 if (name=='Sheet1' and i>=8) or (name=='Sheet2' and i in (7,8)) or (name=='Sheet4-汇总表' and i in (5,6,9,10)) else 2 if name=='Sheet4-汇总表' and i in (7,8) else 0
                s.write_string(0,i,label,head[color])
            moneycols={'Sheet1':{6,7,9},'Sheet2':{6,8,9},'Sheet3':{5,7},'Sheet4-汇总表':{6,8,9,10}}[name]
            for ri,row in enumerate(rows,1):
                excel=ri+1
                for ci,value in enumerate(row):
                    fmt=text if ci<5 else money if ci in moneycols else number
                    formula=None
                    if (name=='Sheet1' and ci==7) or (name=='Sheet2' and ci==9):formula=f'IF(F{excel}=0,"",G{excel}/F{excel})'
                    if (name=='Sheet1' and ci==9) or (name=='Sheet2' and ci==8):
                        day='I' if name=='Sheet1' else 'H'
                        formula=f'G{excel}/{result["summary"]["calendar_days"]}' if rules['daily_basis']=='calendar' else f'IF({day}{excel}=0,"",G{excel}/{day}{excel})'
                    if formula:s.write_formula(ri,ci,'='+formula,fmt,'' if value is None else value)
                    elif value is None or value=='':s.write_blank(ri,ci,None,fmt)
                    elif isinstance(value,str):s.write_string(ri,ci,value,fmt)
                    else:s.write_number(ri,ci,value,fmt)
            s.autofilter(0,0,len(rows),len(names)-1)
    write_csv(out/'逐卡筛选依据.csv',result['audit'],['一卡通号','餐饮次数','餐饮金额','饮品次数','饮品金额','商服次数','商服金额','餐饮次数达标','餐饮金额达标','饮品次数达标','饮品金额达标','商服次数达标','商服金额达标','类别记录条件满足','无记录类别','入选'])
    (out/'运行参数与核对.json').write_text(json.dumps(result['summary'],ensure_ascii=False,indent=2),encoding='utf-8')
    write_explanation(result,out)
    with zipfile.ZipFile(out/'完整结果包.zip','w',zipfile.ZIP_DEFLATED) as z:
        for f in [dest,out/'逐卡筛选依据.csv',out/'运行参数与核对.json',out/'结果说明.md']:z.write(f,f.name)
    return dest
