import csv
from engine import VERSION

def csv_safe(value):
    if isinstance(value,str) and value.startswith(('=','+','-','@')):return "'"+value
    return value

def write_csv(path,rows,fields=None):
    fields=fields or (list(rows[0]) if rows else [])
    with path.open('w',newline='',encoding='utf-8-sig') as f:
        w=csv.writer(f);w.writerow(fields)
        w.writerows([[csv_safe(row.get(k,'')) for k in fields] for row in rows])

def write_explanation(result,out):
    s=result['summary'];r=s['rules'];cat=s['category_totals']
    lines=[f'# {s["month"]} 消费筛选结果说明','',f'版本：{VERSION}；生成时间：{s["generated_at"]}',f'来源：{s["source_name"]}',f'来源SHA256：{s["source_sha256"]}','',f'统计卡号：{s["card_count"]}；同时满足条件：{s["selected_count"]}。','',
      f'- 餐饮：次数 > {r["meal_count"]} 且金额 < {r["meal_amount"]} 元。',f'- 饮品：次数 < {r["drink_count"]} 且金额 < {r["drink_amount"]} 元。',f'- 商服：次数 < {r["service_count"]} 且金额 < {r["service_amount"]} 元。','- 六个条件必须同时满足；等于边界不入选。','',
      '每条负金额交易记一次消费，金额按绝对值以整数分累加。正向、零金额交易单独报告，不当作消费或充值能力评分。',
      '日期优先采用消费日期，缺失时采用入帐时间；按真实年月分别计算，不跨月合计。',
      '每次平均额度 = 类别消费总额 / 类别交易笔数；消费天数 = 该类别有消费的不同日期数。',
      '日均额度分母：'+('该类别实际消费天数。' if r['daily_basis']=='active' else '当月自然日天数。'),
      '无记录类别：'+('按0次、0元参与筛选；无消费时笔均和活跃日均留空。' if r['missing_zero'] else '三类都必须有消费记录才可入选。'),
      f'重复记录：{s["duplicates"]}条；策略：{s["duplicate_policy"]}。没有交易流水号，完全相同不一定代表重复导出。',
      f'学生名册匹配：{s["roster_matches"]}/{s["card_count"]}。未匹配身份字段留空，帐号不是学号。',
      '超市子项口径：'+s['supermarket_basis']+'；计入商户：'+('、'.join(s['supermarket_merchants']) or '无。0仅表示本次上传记录没有识别到超市交易，不代表没有超市消费。'),'',
      '|类别|次数|总额（元）|有记录卡号|','|---|---:|---:|---:|']
    lines += [f'|{k}|{v["count"]}|{v["cents"]/100:.2f}|{v["cards_with_records"]}|' for k,v in cat.items()]
    lines += ['', f'实际覆盖：{s["range"][0]}至{s["range"][1]}。'+('未覆盖整月日期范围，仍按原月度阈值筛选，未折算。' if s['partial_month'] else '日期范围覆盖月初至月末，但仍需确认原始业务数据完整。'), 'Excel保留模板四个表单、原列名和顺序。'+('饮品表末尾追加“每次平均消费额度”，以满足新增要求。' if r['extra_drink_average'] else '饮品笔均在网页明细中查看。'),
      '三类表分别只列实际有该类支出的卡号，不补零行。按一卡通号比对三表共有记录，六项条件同时满足才进入汇总；缺失类别在逐卡依据中标明。餐饮按次数降序，饮品按次数升序，商服按金额升序；同值按卡号排序。',
      '汇总表仅包含六个条件同时满足者。模板未设商服次数列，完整次数及六个布尔判断保存在“逐卡筛选依据.csv”。',
      'Excel是一次运行的结果快照。改流水、名册或阈值后，应在系统重新生成，不能只修改Excel就认为汇总名单会自动重新筛选。',
      '本系统执行用户指定的消费筛选条件，不输出贫困概率，也不作正式经济困难认定。']
    (out/'结果说明.md').write_text('\n'.join(lines),encoding='utf-8')