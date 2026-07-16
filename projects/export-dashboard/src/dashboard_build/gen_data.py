# -*- coding: utf-8 -*-
"""Build aggregated.json from 海关出口数据 xlsx > 海关CV sheet (长表, 海关口径).

口径说明：
- 2023/2024：年度粒度（行里 月==年）
- 2025+   ：月度粒度（行里 月∈1..12），年度总量=12 个月汇总
- 不完整年份（如 2026 M1-2）：年度总量 = 已有月份汇总；同比 = 同年 M1-N vs 前一年 M1-N

元数据（meta）驱动前端：YRS / monthKeys / monthLabels / partialYear / partialMonths 全部从数据自动推。
"""
import os, re, json, pandas as pd, warnings
import openpyxl
warnings.filterwarnings('ignore')

HERE = os.path.dirname(os.path.abspath(__file__))

# ══ 月更只需改这一行 ══
DATA_FILE = os.environ.get('DASHBOARD_DATA_FILE', '海关出口数据-260612_working.xlsx')
# ═══════════════════

xlsx = os.path.join(HERE, DATA_FILE)

def read_haiguan_cv(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['海关CV']
    ws.reset_dimensions()
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        rows.append(list(row[:11]))
    wb.close()
    return pd.DataFrame(rows, columns=['year','month','country','region','subregion','brand','model','energy','qty','series','oem'])

df = read_haiguan_cv(xlsx)

# dataFileDate can be provided explicitly or parsed from the filename:
# 260414 -> '2026-04-14'
DATA_FILE_DATE = os.environ.get('DASHBOARD_DATA_FILE_DATE')
if not DATA_FILE_DATE:
    _m = re.search(r'(\d{2})(\d{2})(\d{2})', DATA_FILE)
    DATA_FILE_DATE = f'20{_m.group(1)}-{_m.group(2)}-{_m.group(3)}' if _m else ''

# 区域/能源 清洗
def remap_region(r, c):
    if r == '欧洲':
        return '俄罗斯' if c == '俄罗斯' else '欧洲（非俄）'
    return r
df['region2'] = df.apply(lambda r: remap_region(r['region'], r['country']), axis=1)
df['energy_big'] = df['energy'].apply(lambda e: '燃油车' if e == '燃油车' else '新能源')
df = df[df['oem'].astype(str).str.strip().str.lower() != 'na'].copy()
df['year'] = df['year'].astype(int)
df['month'] = df['month'].astype(int)
df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0).astype(int)

# ── 序/枚举 ──
R_ORD = ['亚洲','欧洲（非俄）','俄罗斯','北美洲','南美洲','非洲','大洋洲']
S_ORD = ['中系','日系','欧系','美系','韩系']
ENERGIES = ['纯电动','插电混动','增程式','燃油车']
ENERGY_FILTERS = ['all','新能源'] + ENERGIES

# ── 自动推导年份 + 月份 + 不完整年份 ──
# 规则：如果某年的所有数据都 月==年 → 纯年度粒度；否则按 month∈1..12 判断是否 12 月齐全
years_in_data = sorted(df['year'].unique().astype(int).tolist())
# 分类：哪些年是年度粒度（月==年），哪些是月度粒度
def year_granularity(y):
    months = sorted(df[df['year']==y]['month'].unique().astype(int).tolist())
    if months == [y]:
        return 'annual'
    return 'monthly'

YR_GRAN = {y: year_granularity(y) for y in years_in_data}
YR_INTS = years_in_data
YRS = [str(y) for y in YR_INTS]

# latest_year / latest_month（只在月度粒度的年份上推）
monthly_years = [y for y,g in YR_GRAN.items() if g=='monthly']
if monthly_years:
    latest_year = max(monthly_years)
    latest_month = int(df[(df['year']==latest_year) & (df['month']!=latest_year) & (df['qty']>0)]['month'].max())
else:
    latest_year = max(years_in_data)
    latest_month = 12

# partial year：最新年份若 month 数 < 12 就算不完整
lm_months = sorted(df[(df['year']==latest_year) & (df['month']!=latest_year)]['month'].unique().astype(int).tolist())
if lm_months and lm_months != list(range(1,13)):
    PARTIAL_YEAR = str(latest_year)
    PARTIAL_MONTHS = lm_months
else:
    PARTIAL_YEAR = None
    PARTIAL_MONTHS = []

# 月度 keys（从第一个月度年份的 1 月起到 latest_month）
mk_start_year = monthly_years[0] if monthly_years else latest_year
recent_months = []
y, m = mk_start_year, 1
while (y, m) <= (latest_year, latest_month):
    recent_months.append((y, m))
    m += 1
    if m > 12: m = 1; y += 1
MK = [f'{y}-{m}' for y,m in recent_months]
ML = {f'{y}-{m}': f'{y}/{m}月' for y,m in recent_months}

print(f'Years: {YRS} (granularity: {YR_GRAN})')
print(f'Latest: {latest_year}/{latest_month}, partial: {PARTIAL_YEAR} M{PARTIAL_MONTHS}')
print(f'Month keys: {MK[:3]}...{MK[-3:]}')

# ══════════════════════════════════════════════
# 长表原生 helpers
# ══════════════════════════════════════════════
def aqty(sub, year):
    """年度销量：年度粒度用 month==year 行；月度粒度用同 year 下全部 month∈1..12 行。"""
    y = int(year)
    if YR_GRAN.get(y) == 'annual':
        s = sub[(sub['year']==y) & (sub['month']==y)]
    else:
        s = sub[(sub['year']==y) & (sub['month']!=y)]
    return int(s['qty'].sum())

def mqty(sub, ym):
    """指定(年,月)销量（月度粒度）。"""
    y, mo = ym
    s = sub[(sub['year']==y) & (sub['month']==mo)]
    return int(s['qty'].sum())

def arank(sub, group_col, year):
    """按 group_col 分组的 year 年销量降序 Series。"""
    y = int(year)
    if YR_GRAN.get(y) == 'annual':
        s = sub[(sub['year']==y) & (sub['month']==y)]
    else:
        s = sub[(sub['year']==y) & (sub['month']!=y)]
    return s.groupby(group_col)['qty'].sum().sort_values(ascending=False)

def yoy(cur, prev):
    if prev and prev > 0:
        return round((cur/prev - 1)*100, 1)
    return None

def add_yoy(d, yrs=YRS):
    """完整年份 YoY；不完整年份跳过（由 add_partial_yoy_from_monthly 补）。"""
    for i in range(1, len(yrs)):
        if yrs[i] == PARTIAL_YEAR:
            continue
        v = yoy(d.get(yrs[i], 0), d.get(yrs[i-1], 0))
        if v is not None:
            d[yrs[i] + '_yoy'] = v
    return d

def add_partial_yoy_from_monthly(d, sub, filters=None):
    """不完整年份同比：latest_year 前 N 月 vs latest_year-1 同 N 月。"""
    if not PARTIAL_YEAR:
        return d
    s = sub
    if filters:
        for col, val in filters.items():
            s = s[s[col] == val]
    cur = int(s[(s['year']==latest_year) & (s['month'].isin(PARTIAL_MONTHS))]['qty'].sum())
    prev = int(s[(s['year']==latest_year-1) & (s['month'].isin(PARTIAL_MONTHS))]['qty'].sum())
    v = yoy(cur, prev)
    if v is not None:
        d[PARTIAL_YEAR + '_yoy'] = v
    return d

def filter_energy(src, eng):
    """按能源筛选。支持 'all' / '新能源' / 其他具体能源类型。"""
    if eng == 'all': return src
    if eng == '新能源': return src[src['energy_big']=='新能源']
    return src[src['energy']==eng]

def build_ann(sub):
    """构造年度 dict: {年: qty, 年_yoy: yoy}。"""
    d = {}
    for y in YRS: d[y] = aqty(sub, y)
    add_yoy(d)
    add_partial_yoy_from_monthly(d, sub, {})
    return d

def build_mon(sub, with_yoy=True):
    """构造月度 dict: {'YYYY-M': qty, 'YYYY-M_yoy': yoy（可选）}。"""
    d = {}
    for ym in recent_months:
        mk = f'{ym[0]}-{ym[1]}'
        d[mk] = mqty(sub, ym)
        if with_yoy:
            prev = mqty(sub, (ym[0]-1, ym[1]))
            v = yoy(d[mk], prev)
            if v is not None:
                d[mk + '_yoy'] = v
    return d

def has_stock_code(name):
    return bool(re.search(r'[AHU][S.]?\d{3,6}|US\.\w+', str(name)))

agg = {}

# 清洗过 + 带 region/series 的数据（用于主聚合，排除异常行）
clean = df.dropna(subset=['energy','region','series']).copy()

# ══════════════════════════════════════════════
# 1. 总量 + 分系别 + 分动力 + 分区域（板块 0）
# ══════════════════════════════════════════════
agg['total'] = {'中国乘用车总出口': build_ann(clean)}

agg['bySeries'] = {s: build_ann(clean[clean['series']==s]) for s in S_ORD}

agg['byEnergyBig'] = {eb: build_ann(clean[clean['energy_big']==eb]) for eb in ['燃油车','新能源']}

agg['byEnergySub'] = {e: build_ann(clean[clean['energy']==e]) for e in ['纯电动','插电混动','增程式']}

agg['byRegion'] = {r: build_ann(clean[clean['region2']==r]) for r in R_ORD}

# 月度（板块 0）
monthly = {}
monthly['total'] = {'中国乘用车总出口': build_mon(clean)}
monthly['bySeries'] = {s: build_mon(clean[clean['series']==s]) for s in S_ORD}
monthly['byEnergyBig'] = {eb: build_mon(clean[clean['energy_big']==eb]) for eb in ['燃油车','新能源']}
monthly['byEnergySub'] = {e: build_mon(clean[clean['energy']==e]) for e in ['纯电动','插电混动','增程式']}
monthly['byRegion'] = {r: build_mon(clean[clean['region2']==r]) for r in R_ORD}

# ══════════════════════════════════════════════
# 2. 重点车企（带股票代码）板块 0
# ══════════════════════════════════════════════
RANK_YR = latest_year  # 按最新完整/部分年排序
coded_rank_series = arank(clean, 'oem', RANK_YR)
coded_oems_all = [o for o in coded_rank_series.index if has_stock_code(str(o))]

def build_oem_annual(oem_list):
    data = {o: build_ann(clean[clean['oem']==o]) for o in oem_list}
    return {'order': oem_list, 'data': data}

def build_oem_monthly(oem_list, with_yoy=True):
    data = {o: build_mon(clean[clean['oem']==o], with_yoy=with_yoy) for o in oem_list}
    return {'order': oem_list, 'data': data}

agg['top12Oem'] = build_oem_annual(coded_oems_all)
monthly['top12Oem'] = build_oem_monthly(coded_oems_all)
agg['codedTop15'] = coded_oems_all[:15]
agg['codedAll'] = coded_oems_all

# OEM → series 映射
agg['oemSeries'] = clean.groupby('oem')['series'].first().to_dict()

# ══════════════════════════════════════════════
# 3. 分车企 × 动力类型 （板块 0 / 1）
# ══════════════════════════════════════════════
oem_by_eng, oem_by_eng_m = {}, {}
for eng in ENERGY_FILTERS:
    sub = filter_energy(clean, eng)
    order = arank(sub, 'oem', RANK_YR).index.tolist()
    oem_by_eng[eng] = {'order': order,
                       'data': {o: build_ann(sub[sub['oem']==o]) for o in order}}
    oem_by_eng_m[eng] = {'order': order,
                         'data': {o: build_mon(sub[sub['oem']==o]) for o in order}}
agg['oemByEnergy'] = oem_by_eng
monthly['oemByEnergy'] = oem_by_eng_m

# ══════════════════════════════════════════════
# 4. 板块 1 分系别 × 区域/能源/国家
# ══════════════════════════════════════════════
s1sr, s1sr_m = {}, {}
for s in S_ORD:
    sd = clean[clean['series']==s]
    s1sr[s] = {r: build_ann(sd[sd['region2']==r]) for r in R_ORD}
    s1sr_m[s] = {r: build_mon(sd[sd['region2']==r]) for r in R_ORD}
agg['s1BySeriesRegion'] = s1sr
monthly['s1BySeriesRegion'] = s1sr_m

s1ser, s1ser_m = {}, {}
for s in S_ORD:
    sd = clean[clean['series']==s]
    eng_d, eng_dm = {}, {}
    for eng in ENERGY_FILTERS[1:]:  # skip 'all'
        sub = filter_energy(sd, eng)
        reg, reg_m = {}, {}
        for r in R_ORD:
            sub_r = sub[sub['region2']==r]
            ad = build_ann(sub_r)
            md = build_mon(sub_r)
            if any(ad.get(y,0)>0 for y in YRS): reg[r] = ad
            if any(md.get(mk,0)>0 for mk in MK): reg_m[r] = md
        eng_d[eng] = reg
        eng_dm[eng] = reg_m
    s1ser[s] = eng_d
    s1ser_m[s] = eng_dm
agg['s1BySeriesEnergyRegion'] = s1ser
monthly['s1BySeriesEnergyRegion'] = s1ser_m

s1src, s1src_m = {}, {}
for s in S_ORD:
    sd = clean[clean['series']==s]
    rc, rc_m = {}, {}
    for r in R_ORD + ['全部区域']:
        rsub = sd[sd['region2']==r] if r != '全部区域' else sd
        top_c = arank(rsub, 'country', RANK_YR).head(20).index.tolist()
        cd = {c: build_ann(rsub[rsub['country']==c]) for c in top_c}
        cd = {k: v for k, v in cd.items() if any(v.get(y,0)>0 for y in YRS)}
        cd_m = {c: build_mon(rsub[rsub['country']==c]) for c in top_c}
        cd_m = {k: v for k, v in cd_m.items() if any(v.get(mk,0)>0 for mk in MK)}
        if cd: rc[r] = cd
        if cd_m: rc_m[r] = cd_m
    s1src[s] = rc
    s1src_m[s] = rc_m
agg['s1BySeriesRegionCountry'] = s1src
monthly['s1BySeriesRegionCountry'] = s1src_m

s1serc, s1serc_m = {}, {}
for s in S_ORD:
    sd = clean[clean['series']==s]
    eng_d, eng_dm = {}, {}
    for eng in ENERGY_FILTERS[1:]:
        sub = filter_energy(sd, eng)
        rc, rc_m = {}, {}
        for r in R_ORD + ['全部区域']:
            rsub = sub[sub['region2']==r] if r != '全部区域' else sub
            top_c = arank(rsub, 'country', RANK_YR).head(20).index.tolist()
            cd = {c: build_ann(rsub[rsub['country']==c]) for c in top_c}
            cd = {k: v for k, v in cd.items() if any(v.get(y,0)>0 for y in YRS)}
            cd_m = {c: build_mon(rsub[rsub['country']==c]) for c in top_c}
            cd_m = {k: v for k, v in cd_m.items() if any(v.get(mk,0)>0 for mk in MK)}
            if cd: rc[r] = cd
            if cd_m: rc_m[r] = cd_m
        eng_d[eng] = rc
        eng_dm[eng] = rc_m
    s1serc[s] = eng_d
    s1serc_m[s] = eng_dm
agg['s1BySeriesEnergyRegionCountry'] = s1serc
monthly['s1BySeriesEnergyRegionCountry'] = s1serc_m

# regionByEnergy[energy][region] / regionCountry[region][country] / countryByEnergyRegion[region][energy][country]
reg_by_eng, reg_by_eng_m = {}, {}
for eng in ENERGY_FILTERS[1:]:  # skip 'all'
    sub = filter_energy(clean, eng)
    reg_by_eng[eng] = {r: build_ann(sub[sub['region2']==r]) for r in R_ORD}
    reg_by_eng_m[eng] = {r: build_mon(sub[sub['region2']==r]) for r in R_ORD}
agg['regionByEnergy'] = reg_by_eng
monthly['regionByEnergy'] = reg_by_eng_m

reg_country, reg_country_m = {}, {}
for r in R_ORD + ['全部区域']:
    rsub = clean[clean['region2']==r] if r != '全部区域' else clean
    top_c = arank(rsub, 'country', RANK_YR).head(20).index.tolist()
    reg_country[r] = {c: build_ann(rsub[rsub['country']==c]) for c in top_c}
    reg_country_m[r] = {c: build_mon(rsub[rsub['country']==c]) for c in top_c}
agg['regionCountry'] = reg_country
monthly['regionCountry'] = reg_country_m

country_by_eng_r, country_by_eng_r_m = {}, {}
for r in R_ORD + ['全部区域']:
    rsub = clean[clean['region2']==r] if r != '全部区域' else clean
    ed, ed_m = {}, {}
    for eng in ENERGY_FILTERS[1:]:
        sub = filter_energy(rsub, eng)
        top_c = arank(sub, 'country', RANK_YR).head(15).index.tolist()
        ad = {c: build_ann(sub[sub['country']==c]) for c in top_c}
        ad = {k: v for k, v in ad.items() if any(v.get(y,0)>0 for y in YRS)}
        md = {c: build_mon(sub[sub['country']==c]) for c in top_c}
        md = {k: v for k, v in md.items() if any(v.get(mk,0)>0 for mk in MK)}
        ed[eng] = ad
        ed_m[eng] = md
    country_by_eng_r[r] = ed
    country_by_eng_r_m[r] = ed_m
agg['countryByEnergyRegion'] = country_by_eng_r
monthly['countryByEnergyRegion'] = country_by_eng_r_m

# ══════════════════════════════════════════════
# 5. 板块 2 分车企视角
# ══════════════════════════════════════════════
all_oem_rank = arank(clean, 'oem', RANK_YR)
all_oem_list = all_oem_rank.index.tolist()
coded_sec2 = [o for o in all_oem_list if has_stock_code(str(o))]
uncoded_sec2 = [o for o in all_oem_list if not has_stock_code(str(o))]

s2_annual, s2_monthly = {}, {}
for oem in all_oem_list:
    od = clean[clean['oem']==oem]
    # byEnergyBig/Sub
    eng_big = {eb: build_ann(od[od['energy_big']==eb]) for eb in ['燃油车','新能源']}
    eng_sub = {e: build_ann(od[od['energy']==e]) for e in ['纯电动','插电混动','增程式']}
    # byRegion
    by_reg = {}
    for r in R_ORD:
        rd = od[od['region2']==r]
        ad = build_ann(rd)
        if any(ad.get(y,0)>0 for y in YRS): by_reg[r] = ad
    # byRegionEnergy
    by_reg_eng = {}
    for eng in ENERGY_FILTERS:
        sub = filter_energy(od, eng)
        d = {}
        for r in R_ORD:
            rd = sub[sub['region2']==r]
            ad = build_ann(rd)
            if any(ad.get(y,0)>0 for y in YRS): d[r] = ad
        by_reg_eng[eng] = d
    # regionCountry (含全部区域)
    reg_cnt = {}
    for r in ['全部区域'] + R_ORD:
        rd = od[od['region2']==r] if r != '全部区域' else od
        top_c = arank(rd, 'country', RANK_YR).head(30 if r=='全部区域' else 20).index.tolist()
        cd = {c: build_ann(rd[rd['country']==c]) for c in top_c}
        cd = {k: v for k, v in cd.items() if any(v.get(y,0)>0 for y in YRS)}
        if cd: reg_cnt[r] = cd
    # modelByRegionCountry (含全部区域)
    model_rc = {}
    for r in ['全部区域'] + R_ORD:
        rd = od[od['region2']==r] if r != '全部区域' else od
        countries = arank(rd, 'country', RANK_YR).index.tolist()
        rcm = {}
        for c in countries:
            cd = rd[rd['country']==c]
            top_models = arank(cd, 'model', RANK_YR).head(15).index.tolist()
            md = {m: build_ann(cd[cd['model']==m]) for m in top_models}
            md = {k: v for k, v in md.items() if any(v.get(y,0)>0 for y in YRS)}
            if md: rcm[c] = md
        if rcm: model_rc[r] = rcm

    s2_annual[oem] = {
        'byEnergyBig': eng_big, 'byEnergySub': eng_sub,
        'byRegion': by_reg, 'byRegionEnergy': by_reg_eng,
        'regionCountry': reg_cnt, 'modelByRegionCountry': model_rc
    }

    # 月度
    eng_big_m = {eb: build_mon(od[od['energy_big']==eb]) for eb in ['燃油车','新能源']}
    eng_sub_m = {e: build_mon(od[od['energy']==e]) for e in ['纯电动','插电混动','增程式']}
    by_reg_m = {}
    for r in R_ORD:
        md = build_mon(od[od['region2']==r])
        if any(md.get(mk,0)>0 for mk in MK): by_reg_m[r] = md
    by_reg_eng_m = {}
    for eng in ENERGY_FILTERS:
        sub = filter_energy(od, eng)
        d = {}
        for r in R_ORD:
            md = build_mon(sub[sub['region2']==r])
            if any(md.get(mk,0)>0 for mk in MK): d[r] = md
        by_reg_eng_m[eng] = d
    reg_cnt_m = {}
    for r in ['全部区域'] + R_ORD:
        rd = od[od['region2']==r] if r != '全部区域' else od
        countries = arank(rd, 'country', RANK_YR).index.tolist()
        cd = {c: build_mon(rd[rd['country']==c]) for c in countries}
        cd = {k: v for k, v in cd.items() if any(v.get(mk,0)>0 for mk in MK)}
        if cd: reg_cnt_m[r] = cd
    model_rc_m = {}
    for r in ['全部区域'] + R_ORD:
        rd = od[od['region2']==r] if r != '全部区域' else od
        countries = arank(rd, 'country', RANK_YR).index.tolist()
        rcm = {}
        for c in countries:
            cd = rd[rd['country']==c]
            top_models = arank(cd, 'model', RANK_YR).head(15).index.tolist()
            md = {m: build_mon(cd[cd['model']==m]) for m in top_models}
            md = {k: v for k, v in md.items() if any(v.get(mk,0)>0 for mk in MK)}
            if md: rcm[c] = md
        if rcm: model_rc_m[r] = rcm
    s2_monthly[oem] = {
        'byEnergyBig': eng_big_m, 'byEnergySub': eng_sub_m,
        'byRegion': by_reg_m, 'byRegionEnergy': by_reg_eng_m,
        'regionCountry': reg_cnt_m, 'modelByRegionCountry': model_rc_m
    }

# ── 品牌级数据（仅 coded OEMs）──
oem_brands = {}
brand_annual, brand_monthly = {}, {}
for oem in coded_sec2:
    osub = clean[clean['oem']==oem]
    brand_rank = arank(osub, 'brand', RANK_YR)
    brands = [{'name': b, 'sales': int(s)} for b, s in brand_rank.items() if s > 0]
    oem_brands[oem] = brands
    for bi in brands:
        bname = bi['name']
        bsub = osub[osub['brand']==bname]
        if bsub.empty: continue
        # 年度
        eng_big = {eb: build_ann(bsub[bsub['energy_big']==eb]) for eb in ['燃油车','新能源']}
        eng_sub = {e: build_ann(bsub[bsub['energy']==e]) for e in ['纯电动','插电混动','增程式']}
        by_reg = {}
        for r in R_ORD:
            ad = build_ann(bsub[bsub['region2']==r])
            if any(ad.get(y,0)>0 for y in YRS): by_reg[r] = ad
        by_reg_eng = {}
        for eng in ENERGY_FILTERS:
            sub = filter_energy(bsub, eng)
            d = {}
            for r in R_ORD:
                ad = build_ann(sub[sub['region2']==r])
                if any(ad.get(y,0)>0 for y in YRS): d[r] = ad
            by_reg_eng[eng] = d
        reg_cnt = {}
        for r in ['全部区域'] + R_ORD:
            rd = bsub[bsub['region2']==r] if r != '全部区域' else bsub
            top_c = arank(rd, 'country', RANK_YR).index.tolist()
            cd = {c: build_ann(rd[rd['country']==c]) for c in top_c}
            cd = {k: v for k, v in cd.items() if any(v.get(y,0)>0 for y in YRS)}
            if cd: reg_cnt[r] = cd
        model_rc = {}
        for r in ['全部区域'] + R_ORD:
            rd = bsub[bsub['region2']==r] if r != '全部区域' else bsub
            countries = arank(rd, 'country', RANK_YR).index.tolist()
            rcm = {}
            for c in countries:
                cd = rd[rd['country']==c]
                top_models = arank(cd, 'model', RANK_YR).head(15).index.tolist()
                md = {m: build_ann(cd[cd['model']==m]) for m in top_models}
                md = {k: v for k, v in md.items() if any(v.get(y,0)>0 for y in YRS)}
                if md: rcm[c] = md
            if rcm: model_rc[r] = rcm
        brand_annual[bname] = {
            'byEnergyBig': eng_big, 'byEnergySub': eng_sub,
            'byRegion': by_reg, 'byRegionEnergy': by_reg_eng,
            'regionCountry': reg_cnt, 'modelByRegionCountry': model_rc
        }
        # 月度
        eng_big_m = {eb: build_mon(bsub[bsub['energy_big']==eb]) for eb in ['燃油车','新能源']}
        eng_sub_m = {e: build_mon(bsub[bsub['energy']==e]) for e in ['纯电动','插电混动','增程式']}
        by_reg_m = {}
        for r in R_ORD:
            md = build_mon(bsub[bsub['region2']==r])
            if any(md.get(mk,0)>0 for mk in MK): by_reg_m[r] = md
        by_reg_eng_m = {}
        for eng in ENERGY_FILTERS:
            sub = filter_energy(bsub, eng)
            d = {}
            for r in R_ORD:
                md = build_mon(sub[sub['region2']==r])
                if any(md.get(mk,0)>0 for mk in MK): d[r] = md
            by_reg_eng_m[eng] = d
        reg_cnt_m = {}
        for r in ['全部区域'] + R_ORD:
            rd = bsub[bsub['region2']==r] if r != '全部区域' else bsub
            countries = arank(rd, 'country', RANK_YR).index.tolist()
            cd = {c: build_mon(rd[rd['country']==c]) for c in countries}
            cd = {k: v for k, v in cd.items() if any(v.get(mk,0)>0 for mk in MK)}
            if cd: reg_cnt_m[r] = cd
        model_rc_m = {}
        for r in ['全部区域'] + R_ORD:
            rd = bsub[bsub['region2']==r] if r != '全部区域' else bsub
            countries = arank(rd, 'country', RANK_YR).index.tolist()
            rcm = {}
            for c in countries:
                cd = rd[rd['country']==c]
                top_models = arank(cd, 'model', RANK_YR).head(15).index.tolist()
                md = {m: build_mon(cd[cd['model']==m]) for m in top_models}
                md = {k: v for k, v in md.items() if any(v.get(mk,0)>0 for mk in MK)}
                if md: rcm[c] = md
            if rcm: model_rc_m[r] = rcm
        brand_monthly[bname] = {
            'byEnergyBig': eng_big_m, 'byEnergySub': eng_sub_m,
            'byRegion': by_reg_m, 'byRegionEnergy': by_reg_eng_m,
            'regionCountry': reg_cnt_m, 'modelByRegionCountry': model_rc_m
        }

agg['sec2'] = {
    'allOems': all_oem_list, 'codedOems': coded_sec2, 'uncodedOems': uncoded_sec2,
    'top15': coded_sec2[:15],
    'annual': s2_annual, 'monthly': s2_monthly,
    'brandAnnual': brand_annual, 'brandMonthly': brand_monthly
}
agg['oemBrands'] = oem_brands

# 车型 → 主要动力类型映射
model_energy_map = {}
me = clean.groupby(['model','energy'])['qty'].sum().reset_index()
for model in me['model'].unique():
    sub = me[me['model']==model]
    model_energy_map[model] = sub.sort_values('qty', ascending=False).iloc[0]['energy']
agg['modelEnergy'] = model_energy_map

# 月度数据挂到 monthly 子键
agg['monthly'] = monthly

# ══════════════════════════════════════════════
# 6. 元数据
# ══════════════════════════════════════════════
agg['meta'] = {
    'years': YRS,
    'partialYear': PARTIAL_YEAR,
    'partialMonths': PARTIAL_MONTHS,
    'monthKeys': MK,
    'monthLabels': ML,
    'latestYear': latest_year,
    'latestMonth': latest_month,
    'dataFileDate': DATA_FILE_DATE,
    'regionOrder': R_ORD,
    'seriesOrder': S_ORD,
    'energies': ENERGIES,
}

# ══════════════════════════════════════════════
# 7. Save
# ══════════════════════════════════════════════
out_path = os.path.join(HERE, 'aggregated.json')
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(agg, f, ensure_ascii=False, separators=(',', ':'))

sz = len(open(out_path, encoding='utf-8').read())
print(f'Done. File size: {sz/1024:.0f} KB')
print(f'top15 (板块0): {agg["top12Oem"]["order"][:5]} ... ({len(coded_oems_all)} total)')
print(f'sec2 top15: {agg["sec2"]["top15"][:5]}')
print(f'sec2 total OEMs: {len(all_oem_list)}')
print(f'YRS: {YRS}, MK: {MK[0]}..{MK[-1]} ({len(MK)} months)')
