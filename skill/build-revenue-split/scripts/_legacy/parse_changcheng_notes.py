# -*- coding: utf-8 -*-
"""解析长城各年报『销售/管理/研发费用』附注 -> 结构化 note_data(亿元)。
分类: 工资薪金->labor, 股份支付费用->share, 折旧与摊销->DA, 售后服务费->售后, 其余->third-party。
"""
import re, json, os

TXT = "/tmp/gw_{}.txt"
YEARS = list(range(2014, 2026))
NUM = re.compile(r'([\d,]+\.\d+|-)')

def parse_note(fy, kw):
    """返回 {label: current_value_元} 与 合计_元; 取合并报表(首次出现)。"""
    path = TXT.format(fy)
    if not os.path.exists(path):
        return None, None
    lines = open(path, encoding='utf-8').read().splitlines()
    start = None
    for i, ln in enumerate(lines):
        if re.match(rf'^\s*\d+、\s*{kw}\s*$', ln):
            start = i; break
    if start is None:
        return None, None
    items = {}; total = None
    for ln in lines[start+1:start+40]:
        s = ln.strip()
        if not s or '人民币元' in s or '年度' in s or '项目' in s or '经审计' in s or '已重述' in s:
            continue
        if re.match(r'^\d+、', s):   # 下一个附注
            break
        m = NUM.search(s)
        if not m:
            continue
        label = s[:m.start()].strip()
        label = re.sub(r'\(注.*?\)|（注.*?）', '', label).strip()
        val = m.group(1)
        v = 0.0 if val == '-' else float(val.replace(',', ''))
        if not label:
            continue
        if label == '合计':
            total = v; break
        items[label] = v
    return items, total

LABOR = '工资薪金'
SHARE = '股份支付费用'
DA = '折旧与摊销'
AFTER = '售后服务费'

def classify(items):
    """-> dict(third-party labels->值), labor, share, da, after (元)"""
    tp = {}; labor = share = da = after = 0.0
    for lab, v in items.items():
        if lab == LABOR: labor += v
        elif lab == SHARE: share += v
        elif lab == DA: da += v
        elif lab == AFTER: after += v
        else: tp[lab] = v
    return tp, labor, share, da, after

# 长城三表 原始披露 销售/管理/研发费用 (百万元) 校验基准
T = {
 '销售':{2025:11273.11,2024:7830.34,2023:8285.38,2022:5876.16,2021:5192.18,2020:4103.39,2019:3896.67,2018:4575.2,2017:4406.4,2016:3175.42,2015:2841.57,2014:2084.76},
 '管理':{2025:4746.82,2024:4756.38,2023:4735.15,2022:4893.45,2021:4043.07,2020:2552.67,2019:1955.45,2018:1676.3,2017:4963.04,2016:4574.7,2015:4030.6,2014:3822.34},
 '研发':{2025:10432.34,2024:9284.37,2023:8054.25,2022:6445.16,2021:4489.57,2020:3067.48,2019:2716.22,2018:1743.38,2017:None,2016:None,2015:None,2014:None},
}

result = {}
print(f"{'FY':6}{'sec':4}{'合计(附注)':>14}{'三表(百万)':>12}{'差':>8}")
for fy in YEARS:
    rec = {'sales':{}, 'mgmt':{}, 'rd':{}}
    for sec, kw in [('sales','销售费用'),('mgmt','管理费用'),('rd','研发费用')]:
        items, total = parse_note(fy, kw)
        if items is None:
            print(f"{fy:<6}{sec:4}{'(无附注)':>14}")
            rec[sec] = None
            continue
        tp, labor, share, da, after = classify(items)
        rec[sec] = {'tp':tp,'labor':labor,'share':share,'da':da,'after':after,'total':total}
        base = T[{'sales':'销售','mgmt':'管理','rd':'研发'}[sec]][fy]
        diff = '' if base is None else f"{total/1e6-base:+.2f}"
        print(f"{fy:<6}{sec:4}{total/1e6:>14.2f}{(base or 0):>12.2f}{diff:>8}")
    result[fy] = rec

json.dump(result, open('/tmp/gw_notes.json','w'), ensure_ascii=False)
print("\nsaved /tmp/gw_notes.json")
