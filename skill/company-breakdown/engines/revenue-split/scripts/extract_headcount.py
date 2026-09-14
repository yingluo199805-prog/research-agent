# -*- coding: utf-8 -*-
"""通用"专业构成"人数提取器。映射: 制造=生产+其他, 销售=销售, 管理=财务+行政+管理, 研发=研发优先否则技术。
返回 {fy:{制造人数,销售人数,管理人数,研发人数}}。同时回报合计与"在职员工合计"以判断是否群体口径。"""
import re, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common

CATS = ['生产人员', '销售人员', '技术人员', '研发人员', '财务人员', '行政人员', '管理人员', '其他人员',
        '生产及质量人员', '生产及技术人员']

def grab(line):
    m = re.findall(r'[\d,]{2,}', line)
    return int(m[-1].replace(',', '')) if m else None

def extract(prefix, years):
    out = {}
    for fy in years:
        path = common.txt_path(prefix, fy)
        try:
            lines = open(path, encoding='utf-8').read().splitlines()
        except FileNotFoundError:
            continue
        vals = {}; total_emp = None
        for i, ln in enumerate(lines):
            if ('在职员工的数量合计' in ln or '在职员工人数合计' in ln) and total_emp is None:
                total_emp = grab(ln)
            if '专业构成' in ln or '專業構成' in ln:
                for ln2 in lines[i:i+26]:
                    for c in CATS:
                        if c in ln2 and c not in vals:
                            v = grab(ln2)
                            if v is not None:
                                vals[c] = v
                if vals:
                    break
        if not vals:
            continue
        prod = vals.get('生产人员', 0) + vals.get('生产及质量人员', 0) + vals.get('生产及技术人员', 0)
        oth = vals.get('其他人员', 0)
        sale = vals.get('销售人员', 0)
        rd = vals.get('研发人员', 0) or vals.get('技术人员', 0)
        mgmt = vals.get('财务人员', 0) + vals.get('行政人员', 0) + vals.get('管理人员', 0)
        sub = sum(vals.values())
        out[fy] = {'制造人数': prod + oth, '销售人数': sale, '管理人数': mgmt, '研发人数': rd,
                   '_构成合计': sub, '_在职合计': total_emp}
    return out

def clean(prefix, years):
    """仅返回群体口径年份的 {fy:{制造人数,销售人数,管理人数,研发人数}}; 母公司口径(构成<<在职)或全0年份剔除。"""
    raw = extract(prefix, years)
    out = {}
    for fy, d in raw.items():
        emp, sub = d['_在职合计'], d['_构成合计']
        if emp and sub and sub / emp < 0.5:   # 母公司口径, 与合并薪酬不可比
            continue
        hc = {k: d[k] for k in ('制造人数', '销售人数', '管理人数', '研发人数')}
        if sum(hc.values()) == 0:
            continue
        out[fy] = hc
    return out


if __name__ == '__main__':
    import sys
    specs = {'ca': range(2015,2026), 'sl': range(2016,2026), 'jh': range(2015,2026),
             'gz': range(2015,2026), 'sq': range(2015,2025)}
    for p in (sys.argv[1:] or specs):
        rng = specs[p]
        r = extract(p, rng)
        print(f"\n=== {p} ===")
        for fy, d in r.items():
            grp = '群体' if d['_在职合计'] and d['_构成合计'] and abs(d['_构成合计']-d['_在职合计'])/d['_在职合计'] < 0.2 else f"构成{d['_构成合计']}/在职{d['_在职合计']}"
            print(f"  {fy}: 制造{d['制造人数']} 销售{d['销售人数']} 管理{d['管理人数']} 研发{d['研发人数']}  [{grp}]")
