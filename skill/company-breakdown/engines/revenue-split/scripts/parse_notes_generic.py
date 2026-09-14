# -*- coding: utf-8 -*-
"""通用 A股费用附注解析器。兼容注释编号 "N、" 与 "(中文数字)"。
分类用包含匹配: 工资/薪酬/福利->labor, 股份支付->share, 折旧/摊销->DA, 售后/三包/质保->after, 其余->third-party。
"""
import re, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common

NUM = re.compile(r'(-?\d{1,3}(?:,\d{3})+(?:\.\d+)?|-?\d+\.\d+|-)')  # 千分位整数/小数/短横
CN = '一二三四五六七八九十百两〇零'
_LEAD = rf'(?:\d+\s*[、.．]|[(（][{CN}\d]+[)）])'   # "40、" / "46." / "(五十一)" / "(64)"
HDR = lambda kw: re.compile(rf'^\s*{_LEAD}\s*{kw}\s*$')
_NEXTNOTE = re.compile(rf'^\s*{_LEAD}\s*\S')

def parse_note(prefix, fy, kw):
    path = common.txt_path(prefix, fy)
    if not os.path.exists(path):
        return None, None
    lines = open(path, encoding='utf-8').read().splitlines()
    pat = HDR(kw)
    start = next((i for i, ln in enumerate(lines) if pat.match(ln)), None)
    if start is None:
        return None, None
    BOILER = ('人民币', '年度', '项目', '经审计', '重述', '续', '单位：', '财务报表',
              '合并财务', '母公司', '股份有限公司', '附注', '页')
    items, total = {}, None
    for ln in lines[start+1:start+80]:
        s = ln.strip()
        if not s or any(b in s for b in BOILER):
            continue
        if _NEXTNOTE.match(ln) and not pat.match(ln):
            break
        m = NUM.search(s)
        if not m:
            continue
        label = re.sub(r'[(（]注.*?[)）]|[(（][一二三四五六七八九十]+[)）]', '', s[:m.start()]).strip()
        if not label:
            continue
        v = 0.0 if m.group(1) == '-' else float(m.group(1).replace(',', ''))
        if label in ('合计', '小计'):
            total = v
            break
        items[label] = v
    if total is None and items:
        total = sum(items.values())
    return items, total

def classify(items):
    tp, labor, share, da, after = {}, 0.0, 0.0, 0.0, 0.0
    for lab, v in items.items():
        if '股份支付' in lab:
            share += v
        elif ('工资' in lab) or ('薪酬' in lab) or ('福利' in lab) or ('薪金' in lab):
            labor += v
        elif ('折旧' in lab) or ('摊销' in lab):
            da += v
        elif ('售后' in lab) or ('三包' in lab) or ('质保' in lab) or ('质量保证' in lab):
            after += v
        else:
            tp[lab] = v
    return tp, labor, share, da, after

def parse_company(prefix, years):
    """-> {fy: {'sales':{tp,labor,share,da,after,total} or None, 'mgmt':.., 'rd':..}}"""
    out = {}
    for fy in years:
        rec = {}
        for sec, kw in [('sales', '销售费用'), ('mgmt', '管理费用'), ('rd', '研发费用')]:
            items, total = parse_note(prefix, fy, kw)
            if items is None:
                rec[sec] = None
            else:
                tp, labor, share, da, after = classify(items)
                rec[sec] = {'tp': tp, 'labor': labor, 'share': share, 'da': da, 'after': after, 'total': total}
        out[fy] = rec
    return out
