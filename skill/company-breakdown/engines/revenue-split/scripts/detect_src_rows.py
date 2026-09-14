# -*- coding: utf-8 -*-
"""按段定位三表 src_rows（标准化数据端格式）。返回 dict 供 run_company 用。"""
import re, openpyxl

def norm(s):
    s = str(s)
    s = re.sub(r'^[一二三四五六七八九十]+、', '', s)
    s = re.sub(r'(其中|加|减)：', '', s)
    return s.replace('—', '').replace(' ', '').strip()

def detect(path, sheet, sections):
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    ks = sorted(sections.values())
    pl_s = sections['一、利润表']
    bs_s = sections['二、资产负债表']
    cf_s = sections['三、现金流量表']
    end = ws.max_row + 1
    def find(seg_start, seg_end, want, exact=True, contains=False):
        for r in range(seg_start, seg_end):
            lab = ws.cell(r, 1).value
            if lab is None:
                continue
            n = norm(lab)
            if (exact and n == want) or (contains and want in n) or (not exact and not contains and n.startswith(want)):
                return r
        return None
    R = {}
    # 利润表段
    pl = (pl_s, bs_s)
    R['营业总收入'] = find(*pl, '营业总收入')
    R['营业收入'] = find(*pl, '营业收入')
    R['营业成本'] = find(*pl, '营业成本')
    for lab in ['利息收入', '利息支出', '手续费及佣金支出', '税金及附加', '销售费用', '管理费用',
                '研发费用', '财务费用', '公允价值变动收益', '投资收益', '资产处置收益',
                '资产减值损失', '信用减值损失', '其他收益', '营业外收入', '营业外支出', '所得税费用']:
        rr = find(*pl, lab)
        if rr:
            R[lab] = rr
    R['净利润母公司'] = find(*pl, '归属于母公司所有者的净利润')
    R['少数股东损益'] = find(*pl, '少数股东损益')
    # 资产负债表段
    R['应付职工薪酬'] = find(bs_s, cf_s, '应付职工薪酬')
    # 现金流量表段(含补充资料)
    R['支付职工现金'] = find(cf_s, end, '支付给职工以及为职工支付的现金')
    R['固定资产折旧'] = find(cf_s, end, '固定资产折旧', exact=False)
    R['投资性房地产折旧'] = find(cf_s, end, '投资性房地产折旧')
    R['使用权资产折旧'] = find(cf_s, end, '使用权资产折旧', exact=False)
    R['无形资产摊销'] = find(cf_s, end, '无形资产摊销')
    R['长期待摊费用摊销'] = find(cf_s, end, '长期待摊费用摊销')
    return {k: v for k, v in R.items() if v is not None}

if __name__ == '__main__':
    import sys, json
    path, sheet = sys.argv[1], sys.argv[2]
    sections = json.loads(sys.argv[3])
    R = detect(path, sheet, sections)
    for k, v in R.items():
        print(f"  '{k}':{v},")
