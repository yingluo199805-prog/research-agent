# -*- coding: utf-8 -*-
"""Generate aggregated_mkls.json from 海关出口数据 xlsx > mkls-CV sheet（长表原生）.

数据口径：
- mkls-CV sheet 已按"轻型车 + 排除中国"清洗为长表（年/月/国家/集团/品牌/车型/系别/车企划分/数量/...）
- 2020-2023：年度粒度（行里 年==月）
- 2024+  ：月度粒度（行里 年!=月，月 ∈ 1..12）
所有聚合直接 groupby 长表字段，不再 pivot 为宽表。
"""
import os, json, re, pandas as pd, warnings
warnings.filterwarnings('ignore')

HERE = os.path.dirname(os.path.abspath(__file__))

# ══ 月更只需改这一行 ══
DATA_FILE = os.environ.get('DASHBOARD_DATA_FILE', '海关出口数据-260612_working.xlsx')
# ═══════════════════

xlsx = os.path.join(HERE, DATA_FILE)
# MKLS 数据本轮单独更新到 2026-04-28（raw=MarkLines_sales_data_cn 260428.xlsx，
# 已通过 mkls-data-cleaning skill 替换 25M7-26M3 至最新版本）
# 主文件名仍为 260419，其他 Tab(中汽协/海关) 仍按文件名解析 → 2026-04-19
DATA_FILE_DATE = os.environ.get('DASHBOARD_DATA_FILE_DATE', '2026-06-27')
df = pd.read_excel(xlsx, sheet_name='mkls-CV')
# 保险过滤（sheet 本身已过滤，以防手工编辑引入异常）
df = df[(df['标准车种']=='轻型车') & (df['国家/地区']!='中国')].copy()
print(f'轻型车（海外）: {len(df)} long rows')

# ── 列名规整为下游惯用名 ──
df.rename(columns={
    '年': 'year', '月': 'month',
    '新增区域划分': 'region', '大洲': 'continent',
    '集团': 'group', '整车厂/品牌': 'oem',
    '标准动力类型': 'energy', '标准车型': 'model',
    '系别': 'series', '国家/地区': 'country',
    '车企划分': 'listed_company', '数量': 'qty',
}, inplace=True)

# 关键列类型
df['year'] = pd.to_numeric(df['year'], errors='coerce').astype('Int64')
df['month'] = pd.to_numeric(df['month'], errors='coerce').astype('Int64')
df['qty'] = pd.to_numeric(df['qty'], errors='coerce').fillna(0).astype(int)
df = df.dropna(subset=['year','month']).copy()
df['year'] = df['year'].astype(int)
df['month'] = df['month'].astype(int)

# 动力类型映射（混动→常规混动）
ENERGY_MAP = {'ICE':'燃油车','EV':'纯电动','PHV':'插电混动','HV':'常规混动'}
df['energy_cn'] = df['energy'].map(ENERGY_MAP).fillna('其他')

# 欧洲拆分为"欧洲（非俄）"和"俄罗斯"
def remap_continent(row):
    if row['continent'] == '欧洲':
        return '俄罗斯' if row['country'] == '俄罗斯' else '欧洲（非俄）'
    return row['continent']
df['continent'] = df.apply(remap_continent, axis=1)

# 系别合并：只保留中日韩欧美，其余归为"其他"
KEEP_SERIES = {'中系','日系','韩系','欧系','美系'}
df['series'] = df['series'].apply(lambda s: s if s in KEEP_SERIES else '其他')

# 新能源 = 纯电 + 插混（不含常规混动）
def is_nev(e): return e in ('纯电动','插电混动')
df['energy_big'] = df['energy_cn'].apply(lambda e: '新能源' if is_nev(e) else e)

# ── 序/枚举 ──
R_ORD = ['北美','欧洲','日韩','拉美','东南亚','南亚','中亚','俄罗斯','澳新','中东','非洲']
C_ORD = ['亚洲','欧洲（非俄）','俄罗斯','北美洲','南美洲','大洋洲','非洲']
S_ORD = ['中系','日系','欧系','美系','韩系','其他']
ENERGIES = ['纯电动','插电混动','常规混动','燃油车']
ENERGY_FILTERS = ['all','新能源'] + ENERGIES

# ── 年份 + 不完整月份（自动从数据推导）──
YR_INTS = sorted(df['year'].unique().astype(int).tolist())
YRS = [str(y) for y in YR_INTS]
latest_year = max(YR_INTS)
# 在 latest_year 的月度行中（year!=month）找到最大有数据的月
_lm = df[(df['year']==latest_year) & (df['month']!=latest_year) & (df['qty']>0)]['month']
latest_month = int(_lm.max()) if len(_lm) else 0
PARTIAL_YEAR = str(latest_year)
PARTIAL_MONTHS = list(range(1, latest_month+1))
print(f'Partial year: {PARTIAL_YEAR}, months: {PARTIAL_MONTHS}')

# ── 月度 keys（2025/1 起） ──
recent_months = []
y, m = 2025, 1
while (y, m) <= (latest_year, latest_month):
    recent_months.append((y, m))
    m += 1
    if m > 12: m = 1; y += 1
MK = [f'{y}-{m}' for y,m in recent_months]
ML = {f'{y}-{m}': f'{y}/{m}月' for y,m in recent_months}

# ═══════════════════════════════════════════
# 长表原生 helpers
# ═══════════════════════════════════════════
def aqty(sub, year):
    """年度销量（长表原生）：
    - year <= 2023：取 year==month 的年度行（这段数据只有年度粒度）
    - year >= 2024：取同 year 下全部月度行（month∈1..12）"""
    y = int(year)
    if y <= 2023:
        s = sub[(sub['year']==y) & (sub['month']==y)]
    else:
        s = sub[(sub['year']==y) & (sub['month']!=y)]
    return int(s['qty'].sum())

def mqty(sub, ym):
    """指定(年,月)销量（长表原生）。"""
    y, m = ym
    s = sub[(sub['year']==y) & (sub['month']==m)]
    return int(s['qty'].sum())

def arank(sub, group_col, year):
    """按 group_col 分组后的 year 年销量降序 Series。"""
    y = int(year)
    if y <= 2023:
        s = sub[(sub['year']==y) & (sub['month']==y)]
    else:
        s = sub[(sub['year']==y) & (sub['month']!=y)]
    return s.groupby(group_col)['qty'].sum().sort_values(ascending=False)

def yoy(cur, prev):
    if prev and prev > 0:
        return round((cur/prev - 1)*100, 1)
    return None

def add_yoy(d, yrs=YRS):
    """完整年份 YoY（PARTIAL_YEAR 跳过，由 add_partial_yoy_from_monthly 补）。"""
    for i in range(1, len(yrs)):
        if yrs[i] == PARTIAL_YEAR:
            continue
        v = yoy(d.get(yrs[i], 0), d.get(yrs[i-1], 0))
        if v is not None:
            d[yrs[i] + '_yoy'] = v
    return d

def add_partial_yoy_from_monthly(d, sub, filters=None):
    """不完整年份同比：2026 前 N 月 vs 2025 同 N 月。"""
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

def annual_val(sub, year_str, filters=None):
    s = sub
    if filters:
        for col, val in filters.items():
            s = s[s[col] == val]
    return aqty(s, year_str)

def monthly_val(sub, ym, filters=None):
    s = sub
    if filters:
        for col, val in filters.items():
            s = s[s[col] == val]
    return mqty(s, ym)

agg = {}

# ═══ 1. 总量 ═══
total_annual = {}
for y in YRS:
    total_annual[y] = aqty(df, y)
add_yoy(total_annual)
add_partial_yoy_from_monthly(total_annual, df)
agg['total'] = {'海外轻型车终端实销': total_annual}

total_monthly = {}
for ym in recent_months:
    mk = f'{ym[0]}-{ym[1]}'
    total_monthly[mk] = mqty(df, ym)
    prev_val = mqty(df, (ym[0]-1, ym[1]))
    v = yoy(total_monthly[mk], prev_val)
    if v is not None: total_monthly[mk+'_yoy'] = v
agg['monthly_total'] = {'海外轻型车终端实销': total_monthly}

# ═══ 2. 分系别 ═══
all_series = df['series'].dropna().unique().tolist()
series_list = [s for s in S_ORD if s in all_series] + [s for s in all_series if s not in S_ORD]

series_annual = {}
for s in series_list:
    d = {}
    for y in YRS: d[y] = aqty(df[df['series']==s], y)
    add_yoy(d)
    add_partial_yoy_from_monthly(d, df, {'series': s})
    series_annual[s] = d
agg['bySeries'] = series_annual

series_monthly = {}
for s in series_list:
    d = {}
    ssub = df[df['series']==s]
    for ym in recent_months:
        mk = f'{ym[0]}-{ym[1]}'
        d[mk] = mqty(ssub, ym)
        prev_val = mqty(ssub, (ym[0]-1, ym[1]))
        v = yoy(d[mk], prev_val)
        if v is not None: d[mk+'_yoy'] = v
    series_monthly[s] = d
agg['monthly_bySeries'] = series_monthly

# ═══ 3. 分动力类型 ═══
energy_annual = {}
for e in ENERGIES:
    d = {}
    esub = df[df['energy_cn']==e]
    for y in YRS: d[y] = aqty(esub, y)
    add_yoy(d)
    add_partial_yoy_from_monthly(d, esub, {})
    energy_annual[e] = d
# 新能源汇总（纯电+插混）
nev_sub = df[df['energy_cn'].isin(['纯电动','插电混动'])]
nev = {}
for y in YRS: nev[y] = aqty(nev_sub, y)
add_yoy(nev)
add_partial_yoy_from_monthly(nev, nev_sub, {})
energy_annual['新能源'] = nev
agg['byEnergy'] = energy_annual

energy_monthly = {}
for e in ENERGIES:
    d = {}
    esub = df[df['energy_cn']==e]
    for ym in recent_months:
        mk = f'{ym[0]}-{ym[1]}'
        d[mk] = mqty(esub, ym)
        prev_val = mqty(esub, (ym[0]-1, ym[1]))
        v = yoy(d[mk], prev_val)
        if v is not None: d[mk+'_yoy'] = v
    energy_monthly[e] = d
# 新能源月度汇总
nev_m = {}
for ym in recent_months:
    mk = f'{ym[0]}-{ym[1]}'
    nev_m[mk] = mqty(nev_sub, ym)
    prev_val = mqty(nev_sub, (ym[0]-1, ym[1]))
    v = yoy(nev_m[mk], prev_val)
    if v is not None: nev_m[mk+'_yoy'] = v
energy_monthly['新能源'] = nev_m
agg['monthly_byEnergy'] = energy_monthly

# ═══ 3b. 分动力类型 × 大洲/国家（层级展示格式）═══
ENG_DISPLAY_KEYS = ['新能源','纯电动','插电混动','常规混动','燃油车']

def build_energy_display(sub):
    """构建层级展示格式的动力类型年度数据（含新能源=纯电+插混）"""
    result = {}
    for e in ENERGIES:
        d = {}
        esub = sub[sub['energy_cn']==e]
        for y in YRS: d[y] = aqty(esub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, esub, {})
        if any(d.get(y,0)>0 for y in YRS): result[e] = d
    # 新能源汇总
    nev_d = {}
    nev_s = sub[sub['energy_cn'].isin(['纯电动','插电混动'])]
    for y in YRS: nev_d[y] = aqty(nev_s, y)
    add_yoy(nev_d)
    add_partial_yoy_from_monthly(nev_d, nev_s, {})
    if any(nev_d.get(y,0)>0 for y in YRS): result['新能源'] = nev_d
    return result

# energyByContinent[continent] = {energy_display_key: {year: qty}}
energy_by_cont = {}
for cont in C_ORD + ['全部区域']:
    csub = df[df['continent']==cont] if cont != '全部区域' else df
    energy_by_cont[cont] = build_energy_display(csub)
agg['energyByContinent'] = energy_by_cont

# energyByCountry[continent][country] = {energy_display_key: {year: qty}}
# countryListByContinent[continent] = [country1, ...]（按2025降序）
energy_by_country = {}
country_list_by_cont = {}
for cont in C_ORD + ['全部区域']:
    csub = df[df['continent']==cont] if cont != '全部区域' else df
    top_countries = arank(csub, 'country', 2025).head(30).index.tolist()
    country_list_by_cont[cont] = top_countries
    cd = {}
    for country in top_countries:
        ctsub = csub[csub['country']==country]
        ed = build_energy_display(ctsub)
        if ed: cd[country] = ed
    energy_by_country[cont] = cd
agg['energyByCountry'] = energy_by_country
agg['countryListByContinent'] = country_list_by_cont

# ═══ 4. 分大洲 + 分区域 ═══
continent_annual = {}
for c in C_ORD:
    d = {}
    csub = df[df['continent']==c]
    for y in YRS: d[y] = aqty(csub, y)
    add_yoy(d)
    add_partial_yoy_from_monthly(d, csub, {})
    if any(d.get(y,0)>0 for y in YRS): continent_annual[c] = d
agg['byContinent'] = continent_annual

region_annual = {}
for r in R_ORD:
    d = {}
    rsub = df[df['region']==r]
    for y in YRS: d[y] = aqty(rsub, y)
    add_yoy(d)
    add_partial_yoy_from_monthly(d, rsub, {})
    if any(d.get(y,0)>0 for y in YRS): region_annual[r] = d
agg['byRegion'] = region_annual

continent_monthly = {}
for c in C_ORD:
    d = {}
    csub = df[df['continent']==c]
    for ym in recent_months:
        mk = f'{ym[0]}-{ym[1]}'
        d[mk] = mqty(csub, ym)
        prev_val = mqty(csub, (ym[0]-1, ym[1]))
        v = yoy(d[mk], prev_val)
        if v is not None: d[mk+'_yoy'] = v
    if any(d.get(mk,0)>0 for mk in MK): continent_monthly[c] = d
agg['monthly_byContinent'] = continent_monthly

region_monthly = {}
for r in R_ORD:
    d = {}
    rsub = df[df['region']==r]
    for ym in recent_months:
        mk = f'{ym[0]}-{ym[1]}'
        d[mk] = mqty(rsub, ym)
        prev_val = mqty(rsub, (ym[0]-1, ym[1]))
        v = yoy(d[mk], prev_val)
        if v is not None: d[mk+'_yoy'] = v
    if any(d.get(mk,0)>0 for mk in MK): region_monthly[r] = d
agg['monthly_byRegion'] = region_monthly

# ═══ 5. 分国家（全部区域 + 各大洲 + 各区域）═══
def build_country_data(src_df, group_col, group_val):
    sub = src_df[src_df[group_col]==group_val] if group_val else src_df
    countries = arank(sub, 'country', int(YRS[-2])).index.tolist()
    annual_d = {}
    for c in countries:
        csub = sub[sub['country']==c]
        d = {}
        for y in YRS: d[y] = aqty(csub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, csub, {})
        if any(d.get(y,0)>0 for y in YRS): annual_d[c] = d
    monthly_d = {}
    for c in countries:
        csub = sub[sub['country']==c]
        d = {}
        for ym in recent_months:
            mk = f'{ym[0]}-{ym[1]}'
            d[mk] = mqty(csub, ym)
            prev_val = mqty(csub, (ym[0]-1, ym[1]))
            v = yoy(d[mk], prev_val)
            if v is not None: d[mk+'_yoy'] = v
        if any(d.get(mk,0)>0 for mk in MK): monthly_d[c] = d
    return annual_d, monthly_d

a_all, m_all = build_country_data(df, None, None)
country_annual = {'全部区域': a_all}
country_monthly = {'全部区域': m_all}
for c in C_ORD:
    a, m = build_country_data(df, 'continent', c)
    if a: country_annual[c] = a
    if m: country_monthly[c] = m
for r in R_ORD:
    a, m = build_country_data(df, 'region', r)
    if a: country_annual[r] = a
    if m: country_monthly[r] = m
agg['countryByRegion'] = country_annual
agg['monthly_countryByRegion'] = country_monthly

# ═══ 6. 重点车企（中系为主，按2025排序）═══
cn_groups = arank(df[df['series']=='中系'], 'group', 2025).index.tolist()
all_groups = arank(df, 'group', 2025).index.tolist()

group_annual = {}
for g in all_groups:
    d = {}
    gsub = df[df['group']==g]
    for y in YRS: d[y] = aqty(gsub, y)
    add_yoy(d)
    add_partial_yoy_from_monthly(d, gsub, {})
    group_annual[g] = d

group_monthly = {}
for g in all_groups:
    d = {}
    gsub = df[df['group']==g]
    for ym in recent_months:
        mk = f'{ym[0]}-{ym[1]}'
        d[mk] = mqty(gsub, ym)
        prev_val = mqty(gsub, (ym[0]-1, ym[1]))
        v = yoy(d[mk], prev_val)
        if v is not None: d[mk+'_yoy'] = v
    group_monthly[g] = d

group_series = df.groupby('group')['series'].first().to_dict()

agg['groupAnnual'] = {'order': all_groups, 'data': group_annual}
agg['groupMonthly'] = {'order': all_groups, 'data': group_monthly}
agg['cnTop15'] = cn_groups[:15]
agg['cnAll'] = cn_groups
agg['groupSeries'] = group_series

# ═══ 7. 分车企视角（按上市公司角度 listed_company） ═══
_code_pat = re.compile(r'[AH]\d{4,6}|US\.')
lc_rank_series = arank(df, 'listed_company', 2025)
all_listed = lc_rank_series.index.tolist()
coded_oems = [o for o in all_listed if _code_pat.search(str(o)) and lc_rank_series[o] > 0]
uncoded_oems = [o for o in all_listed if o not in coded_oems and lc_rank_series[o] > 0][:30]
sec2_oem_list = coded_oems + uncoded_oems
print(f'Sec2: {len(coded_oems)} coded OEMs, {len(uncoded_oems)} uncoded OEMs, total {len(sec2_oem_list)}')

# modelEnergy 映射：每个model取最常见的energy_cn
_me = df.groupby('model')['energy_cn'].agg(lambda x: x.value_counts().idxmax())
model_energy_map = _me.to_dict()

sec2_annual = {}
sec2_monthly = {}

for g in sec2_oem_list:
    gsub = df[df['listed_company']==g]
    if gsub.empty:
        continue
    # 动力（年度）
    eng_d = {}
    for e in ENERGIES:
        d = {}
        esub = gsub[gsub['energy_cn']==e]
        for y in YRS: d[y] = aqty(esub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, esub, {})
        eng_d[e] = d
    # 大洲
    reg_d = {}
    for c in C_ORD:
        d = {}
        csub = gsub[gsub['continent']==c]
        for y in YRS: d[y] = aqty(csub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, csub, {})
        if any(d.get(y,0)>0 for y in YRS): reg_d[c] = d
    # 细分区域
    subreg_d = {}
    for r in R_ORD:
        d = {}
        rsub = gsub[gsub['region']==r]
        for y in YRS: d[y] = aqty(rsub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, rsub, {})
        if any(d.get(y,0)>0 for y in YRS): subreg_d[r] = d
    # 国家（全部+各大洲）
    country_d = {}
    top_c = arank(gsub, 'country', 2025).head(30).index.tolist()
    cd = {}
    for c in top_c:
        d = {}
        csub = gsub[gsub['country']==c]
        for y in YRS: d[y] = aqty(csub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, csub, {})
        if any(d.get(y,0)>0 for y in YRS): cd[c] = d
    country_d['全部区域'] = cd
    for cont in C_ORD:
        csub = gsub[gsub['continent']==cont]
        top_cc = arank(csub, 'country', 2025).head(20).index.tolist()
        ccd = {}
        for c in top_cc:
            d = {}
            cc2 = csub[csub['country']==c]
            for y in YRS: d[y] = aqty(cc2, y)
            add_yoy(d)
            add_partial_yoy_from_monthly(d, cc2, {})
            if any(d.get(y,0)>0 for y in YRS): ccd[c] = d
        if ccd: country_d[cont] = ccd
    # 车型
    top_models = arank(gsub, 'model', 2025).head(30).index.tolist()
    model_d = {}
    for mdl in top_models:
        d = {}
        msub = gsub[gsub['model']==mdl]
        for y in YRS: d[y] = aqty(msub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, msub, {})
        if any(d.get(y,0)>0 for y in YRS): model_d[mdl] = d

    # byRegionEnergy: {energy: {continent: {year: qty}}}
    bre_d = {}
    for e in ENERGIES:
        esub = gsub[gsub['energy_cn']==e]
        cont_d = {}
        for c in C_ORD:
            d = {}
            cc2 = esub[esub['continent']==c]
            for y in YRS: d[y] = aqty(cc2, y)
            add_yoy(d)
            add_partial_yoy_from_monthly(d, cc2, {})
            if any(d.get(y,0)>0 for y in YRS): cont_d[c] = d
        if cont_d: bre_d[e] = cont_d
    # 新能源
    nev_s = gsub[gsub['energy_cn'].isin(['纯电动','插电混动'])]
    nev_cont_d = {}
    for c in C_ORD:
        d = {}
        cc2 = nev_s[nev_s['continent']==c]
        for y in YRS: d[y] = aqty(cc2, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, cc2, {})
        if any(d.get(y,0)>0 for y in YRS): nev_cont_d[c] = d
    if nev_cont_d: bre_d['新能源'] = nev_cont_d

    # modelByRegionCountry: {continent: {country: {model: {year: qty}}}}
    mbrc_d = {}
    # 全部区域
    mbrc_all = {}
    for c in top_c:
        csub_m = gsub[gsub['country']==c]
        cm_rank = arank(csub_m, 'model', 2025).head(20)
        cm_d = {}
        for mdl in cm_rank.index:
            d = {}
            mm = csub_m[csub_m['model']==mdl]
            for y in YRS: d[y] = aqty(mm, y)
            add_yoy(d)
            add_partial_yoy_from_monthly(d, mm, {})
            if any(d.get(y,0)>0 for y in YRS): cm_d[mdl] = d
        if cm_d: mbrc_all[c] = cm_d
    mbrc_d['全部区域'] = mbrc_all
    for cont in C_ORD:
        csub = gsub[gsub['continent']==cont]
        cont_countries = arank(csub, 'country', 2025).head(15).index.tolist()
        cont_cm = {}
        for c in cont_countries:
            ccsub = csub[csub['country']==c]
            cm_rank2 = arank(ccsub, 'model', 2025).head(15)
            cm_d2 = {}
            for mdl in cm_rank2.index:
                d = {}
                mm = ccsub[ccsub['model']==mdl]
                for y in YRS: d[y] = aqty(mm, y)
                add_yoy(d)
                add_partial_yoy_from_monthly(d, mm, {})
                if any(d.get(y,0)>0 for y in YRS): cm_d2[mdl] = d
            if cm_d2: cont_cm[c] = cm_d2
        if cont_cm: mbrc_d[cont] = cont_cm

    sec2_annual[g] = {
        'byEnergy': eng_d, 'byContinent': reg_d, 'byRegion': subreg_d,
        'countryByRegion': country_d, 'models': model_d,
        'byRegionEnergy': bre_d, 'modelByRegionCountry': mbrc_d
    }

    # 月度
    eng_m = {}
    for e in ENERGIES:
        d = {}
        esub = gsub[gsub['energy_cn']==e]
        for ym in recent_months:
            mk = f'{ym[0]}-{ym[1]}'
            d[mk] = mqty(esub, ym)
        eng_m[e] = d
    reg_m = {}
    for c in C_ORD:
        d = {}
        csub2 = gsub[gsub['continent']==c]
        for ym in recent_months:
            mk = f'{ym[0]}-{ym[1]}'
            d[mk] = mqty(csub2, ym)
        if any(d.get(mk,0)>0 for mk in MK): reg_m[c] = d
    model_m = {}
    for mdl in top_models:
        d = {}
        msub = gsub[gsub['model']==mdl]
        for ym in recent_months:
            mk = f'{ym[0]}-{ym[1]}'
            d[mk] = mqty(msub, ym)
        if any(d.get(mk,0)>0 for mk in MK): model_m[mdl] = d
    country_m = {}
    for cont in C_ORD + ['全部区域']:
        csub2 = gsub[gsub['continent']==cont] if cont != '全部区域' else gsub
        top_cc = arank(csub2, 'country', 2025).head(20).index.tolist()
        ccd = {}
        for c in top_cc:
            d = {}
            ccsub = csub2[csub2['country']==c]
            for ym in recent_months:
                mk = f'{ym[0]}-{ym[1]}'
                d[mk] = mqty(ccsub, ym)
                prev_val = mqty(ccsub, (ym[0]-1, ym[1]))
                v = yoy(d[mk], prev_val)
                if v is not None: d[mk+'_yoy'] = v
            if any(d.get(mk,0)>0 for mk in MK): ccd[c] = d
        if ccd: country_m[cont] = ccd
    sec2_monthly[g] = {'byEnergy': eng_m, 'byContinent': reg_m, 'models': model_m, 'countryByRegion': country_m}

# ── 品牌级数据（仅 coded OEMs）──
oem_brands = {}
brand_annual = {}
brand_monthly = {}
for lc in coded_oems:
    gsub = df[df['listed_company']==lc]
    brand_rank = arank(gsub, 'oem', 2025)
    brands = [{'name': b, 'sales': int(s)} for b, s in brand_rank.items() if s > 0]
    oem_brands[lc] = brands
    for bi in brands:
        bname = bi['name']
        bsub = gsub[gsub['oem']==bname]
        if bsub.empty: continue
        # 年度
        eng_d = {}
        for e in ENERGIES:
            d = {}
            esub = bsub[bsub['energy_cn']==e]
            for y in YRS: d[y] = aqty(esub, y)
            add_yoy(d)
            add_partial_yoy_from_monthly(d, esub, {})
            eng_d[e] = d
        reg_d = {}
        for c in C_ORD:
            d = {}
            csub = bsub[bsub['continent']==c]
            for y in YRS: d[y] = aqty(csub, y)
            add_yoy(d)
            add_partial_yoy_from_monthly(d, csub, {})
            if any(d.get(y,0)>0 for y in YRS): reg_d[c] = d
        country_d = {}
        for cont in ['全部区域'] + C_ORD:
            csub3 = bsub[bsub['continent']==cont] if cont != '全部区域' else bsub
            top_cc = arank(csub3, 'country', 2025).head(15).index.tolist()
            ccd = {}
            for c in top_cc:
                d = {}
                cc2 = csub3[csub3['country']==c]
                for y in YRS: d[y] = aqty(cc2, y)
                add_yoy(d)
                add_partial_yoy_from_monthly(d, cc2, {})
                if any(d.get(y,0)>0 for y in YRS): ccd[c] = d
            if ccd: country_d[cont] = ccd
        top_models_b = arank(bsub, 'model', 2025).head(15).index.tolist()
        model_d = {}
        for mdl in top_models_b:
            d = {}
            mm = bsub[bsub['model']==mdl]
            for y in YRS: d[y] = aqty(mm, y)
            add_yoy(d)
            add_partial_yoy_from_monthly(d, mm, {})
            if any(d.get(y,0)>0 for y in YRS): model_d[mdl] = d
        brand_annual[bname] = {'byEnergy': eng_d, 'byContinent': reg_d, 'countryByRegion': country_d, 'models': model_d}
        # 月度
        eng_m = {}
        for e in ENERGIES:
            d = {}
            esub = bsub[bsub['energy_cn']==e]
            for ym in recent_months:
                mk = f'{ym[0]}-{ym[1]}'
                d[mk] = mqty(esub, ym)
            eng_m[e] = d
        reg_m = {}
        for c in C_ORD:
            d = {}
            csub2 = bsub[bsub['continent']==c]
            for ym in recent_months:
                mk = f'{ym[0]}-{ym[1]}'
                d[mk] = mqty(csub2, ym)
            if any(d.get(mk,0)>0 for mk in MK): reg_m[c] = d
        model_m = {}
        for mdl in top_models_b:
            d = {}
            msub = bsub[bsub['model']==mdl]
            for ym in recent_months:
                mk = f'{ym[0]}-{ym[1]}'
                d[mk] = mqty(msub, ym)
            if any(d.get(mk,0)>0 for mk in MK): model_m[mdl] = d
        country_m = {}
        for cont in ['全部区域'] + C_ORD:
            csub2 = bsub[bsub['continent']==cont] if cont != '全部区域' else bsub
            top_cc = arank(csub2, 'country', 2025).head(15).index.tolist()
            ccd = {}
            for c in top_cc:
                d = {}
                ccsub = csub2[csub2['country']==c]
                for ym in recent_months:
                    mk = f'{ym[0]}-{ym[1]}'
                    d[mk] = mqty(ccsub, ym)
                if any(d.get(mk,0)>0 for mk in MK): ccd[c] = d
            if ccd: country_m[cont] = ccd
        brand_monthly[bname] = {'byEnergy': eng_m, 'byContinent': reg_m, 'models': model_m, 'countryByRegion': country_m}

agg['sec2'] = {
    'annual': sec2_annual,
    'monthly': sec2_monthly,
    'codedOems': coded_oems,
    'uncodedOems': uncoded_oems,
    'brandAnnual': brand_annual,
    'brandMonthly': brand_monthly,
}
agg['oemBrands'] = oem_brands
agg['modelEnergy'] = model_energy_map

# ═══ 8. 板块1 分系别×动力×区域/国家 交叉数据 ═══
def filter_energy_mkls(src, eng):
    if eng == 'all': return src
    if eng == '新能源': return src[src['energy_cn'].isin(['纯电动','插电混动'])]
    return src[src['energy_cn']==eng]

def build_annual_dict(sub, group_col, group_vals):
    """按 group_col 聚合年度数据"""
    result = {}
    for g in group_vals:
        d = {}
        gsub = sub[sub[group_col]==g] if group_col else sub
        for y in YRS: d[y] = aqty(gsub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, gsub, {})
        if any(d.get(y,0)>0 for y in YRS): result[g] = d
    return result

def build_monthly_dict(sub, group_col, group_vals):
    """按 group_col 聚合月度数据"""
    result = {}
    for g in group_vals:
        gsub = sub[sub[group_col]==g] if group_col else sub
        d = {}
        for ym in recent_months:
            mk = f'{ym[0]}-{ym[1]}'
            d[mk] = mqty(gsub, ym)
            prev_val = mqty(gsub, (ym[0]-1, ym[1]))
            v = yoy(d[mk], prev_val)
            if v is not None: d[mk+'_yoy'] = v
        if any(d.get(mk,0)>0 for mk in MK): result[g] = d
    return result

# s1BySeriesRegion[series][continent] = {year:qty}
s1sr, s1sr_m = {}, {}
for s in S_ORD:
    sd = df[df['series']==s]
    s1sr[s] = build_annual_dict(sd, 'continent', C_ORD)
    s1sr_m[s] = build_monthly_dict(sd, 'continent', C_ORD)
    # 全部区域
    all_d = {}
    for y in YRS: all_d[y] = aqty(sd, y)
    add_yoy(all_d)
    s1sr[s]['全部区域'] = all_d
agg['s1BySeriesRegion'] = s1sr
agg['monthly_s1BySeriesRegion'] = s1sr_m

# s1BySeriesEnergyRegion[series][energy][continent]
s1ser, s1ser_m = {}, {}
for s in S_ORD:
    sd = df[df['series']==s]
    eng_d, eng_dm = {}, {}
    for eng in ENERGY_FILTERS[1:]:
        sub = filter_energy_mkls(sd, eng)
        eng_d[eng] = build_annual_dict(sub, 'continent', C_ORD)
        eng_dm[eng] = build_monthly_dict(sub, 'continent', C_ORD)
    s1ser[s] = eng_d
    s1ser_m[s] = eng_dm
agg['s1BySeriesEnergyRegion'] = s1ser
agg['monthly_s1BySeriesEnergyRegion'] = s1ser_m

# s1BySeriesRegionCountry[series][continent][country]
s1src, s1src_m = {}, {}
for s in S_ORD:
    sd = df[df['series']==s]
    rc, rc_m = {}, {}
    for cont in C_ORD + ['全部区域']:
        csub = sd[sd['continent']==cont] if cont != '全部区域' else sd
        top_c = arank(csub, 'country', 2025).head(15).index.tolist()
        rc[cont] = build_annual_dict(csub, 'country', top_c)
        rc_m[cont] = build_monthly_dict(csub, 'country', top_c)
    s1src[s] = rc
    s1src_m[s] = rc_m
agg['s1BySeriesRegionCountry'] = s1src
agg['monthly_s1BySeriesRegionCountry'] = s1src_m

# s1BySeriesEnergyRegionCountry[series][energy][continent][country]
s1serc, s1serc_m = {}, {}
for s in S_ORD:
    sd = df[df['series']==s]
    eng_d, eng_dm = {}, {}
    for eng in ENERGY_FILTERS[1:]:
        sub = filter_energy_mkls(sd, eng)
        rc, rc_m = {}, {}
        for cont in C_ORD + ['全部区域']:
            csub = sub[sub['continent']==cont] if cont != '全部区域' else sub
            top_c = arank(csub, 'country', 2025).head(15).index.tolist()
            if top_c:
                rc[cont] = build_annual_dict(csub, 'country', top_c)
                rc_m[cont] = build_monthly_dict(csub, 'country', top_c)
        eng_d[eng] = rc
        eng_dm[eng] = rc_m
    s1serc[s] = eng_d
    s1serc_m[s] = eng_dm
agg['s1BySeriesEnergyRegionCountry'] = s1serc
agg['monthly_s1BySeriesEnergyRegionCountry'] = s1serc_m

# regionByEnergy[energy][continent]（全系别）
reg_by_eng, reg_by_eng_m = {}, {}
for eng in ENERGY_FILTERS[1:]:
    sub = filter_energy_mkls(df, eng)
    reg_by_eng[eng] = build_annual_dict(sub, 'continent', C_ORD)
    reg_by_eng_m[eng] = build_monthly_dict(sub, 'continent', C_ORD)
agg['regionByEnergy'] = reg_by_eng
agg['monthly_regionByEnergy'] = reg_by_eng_m

# regionCountry（全系别全动力）[continent/全部区域][country]
rc_all, rc_all_m = {}, {}
for cont in C_ORD + ['全部区域']:
    csub = df[df['continent']==cont] if cont != '全部区域' else df
    top_c = arank(csub, 'country', 2025).head(20).index.tolist()
    rc_all[cont] = build_annual_dict(csub, 'country', top_c)
    rc_all_m[cont] = build_monthly_dict(csub, 'country', top_c)
agg['regionCountry'] = rc_all
agg['monthly_regionCountry'] = rc_all_m

# countryByEnergyRegion[continent][energy][country]
cber, cber_m = {}, {}
for cont in C_ORD + ['全部区域']:
    csub = df[df['continent']==cont] if cont != '全部区域' else df
    eng_d, eng_dm = {}, {}
    for eng in ENERGY_FILTERS[1:]:
        sub = filter_energy_mkls(csub, eng)
        top_c = arank(sub, 'country', 2025).head(15).index.tolist()
        if top_c:
            eng_d[eng] = build_annual_dict(sub, 'country', top_c)
            eng_dm[eng] = build_monthly_dict(sub, 'country', top_c)
    cber[cont] = eng_d
    cber_m[cont] = eng_dm
agg['countryByEnergyRegion'] = cber
agg['monthly_countryByEnergyRegion'] = cber_m

# oemByEnergy[energy_key] = {order:[oem...], data:{oem:{year:qty}}}
oem_by_eng, oem_by_eng_m = {}, {}
for eng in ENERGY_FILTERS:
    sub = filter_energy_mkls(df, eng)
    order = arank(sub, 'listed_company', 2025).index.tolist()
    data_a, data_m = {}, {}
    for g in order:
        gsub = sub[sub['listed_company']==g]
        d = {}
        for y in YRS: d[y] = aqty(gsub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, gsub, {})
        data_a[g] = d
        dm = {}
        for ym in recent_months:
            mk = f'{ym[0]}-{ym[1]}'
            dm[mk] = mqty(gsub, ym)
            prev_val = mqty(gsub, (ym[0]-1, ym[1]))
            v = yoy(dm[mk], prev_val)
            if v is not None: dm[mk+'_yoy'] = v
        data_m[g] = dm
    oem_by_eng[eng] = {'order': order, 'data': data_a}
    oem_by_eng_m[eng] = {'order': order, 'data': data_m}
agg['oemByEnergy'] = oem_by_eng
agg['monthly_oemByEnergy'] = oem_by_eng_m

# oemByEnergyCont[energy_key][continent] / oemByEnergyCountry[energy_key][continent][country]
def build_oem_data(sub_df):
    """按listed_company聚合年度+月度OEM数据"""
    rank = arank(sub_df, 'listed_company', 2025)
    order = rank[rank>0].index.tolist()
    data_a, data_m = {}, {}
    for g in order:
        gsub = sub_df[sub_df['listed_company']==g]
        d = {}
        for y in YRS: d[y] = aqty(gsub, y)
        add_yoy(d)
        add_partial_yoy_from_monthly(d, gsub, {})
        data_a[g] = d
        dm = {}
        for ym in recent_months:
            mk = f'{ym[0]}-{ym[1]}'
            dm[mk] = mqty(gsub, ym)
            prev_val = mqty(gsub, (ym[0]-1, ym[1]))
            v = yoy(dm[mk], prev_val)
            if v is not None: dm[mk+'_yoy'] = v
        data_m[g] = dm
    return {'order': order, 'data': data_a}, {'order': order, 'data': data_m}

oem_by_eng_cont, oem_by_eng_cont_m = {}, {}
oem_by_eng_country, oem_by_eng_country_m = {}, {}
for eng in ENERGY_FILTERS:
    sub = filter_energy_mkls(df, eng)
    cont_a, cont_m = {}, {}
    country_a, country_m = {}, {}
    for cont in C_ORD + ['全部区域']:
        csub = sub[sub['continent']==cont] if cont != '全部区域' else sub
        a, m = build_oem_data(csub)
        cont_a[cont] = a
        cont_m[cont] = m
        top_countries = arank(csub, 'country', 2025).head(20).index.tolist()
        cd_a, cd_m = {}, {}
        for c in top_countries:
            ctsub = csub[csub['country']==c]
            ca, cm = build_oem_data(ctsub)
            if ca['order']: cd_a[c] = ca
            if cm['order']: cd_m[c] = cm
        country_a[cont] = cd_a
        country_m[cont] = cd_m
    oem_by_eng_cont[eng] = cont_a
    oem_by_eng_cont_m[eng] = cont_m
    oem_by_eng_country[eng] = country_a
    oem_by_eng_country_m[eng] = country_m
agg['oemByEnergyCont'] = oem_by_eng_cont
agg['monthly_oemByEnergyCont'] = oem_by_eng_cont_m
agg['oemByEnergyCountry'] = oem_by_eng_country
agg['monthly_oemByEnergyCountry'] = oem_by_eng_country_m

# oemSeries映射（listed_company → 系别）
agg['oemSeries'] = df.groupby('listed_company')['series'].first().to_dict()

# ═══ 元数据 ═══
agg['meta'] = {
    'years': YRS,
    'monthKeys': MK,
    'monthLabels': ML,
    'latestYear': latest_year,
    'latestMonth': latest_month,
    'partialYear': PARTIAL_YEAR,
    'partialMonths': PARTIAL_MONTHS,
    'dataFileDate': DATA_FILE_DATE,
    'continentOrder': C_ORD,
    'regionOrder': R_ORD,
    'seriesList': series_list,
    'energyList': ENERGIES,
}

# ═══ Save ═══
out_path = os.path.join(HERE, 'aggregated_mkls.json')
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(agg, f, ensure_ascii=False, separators=(',', ':'))

sz = len(open(out_path, encoding='utf-8').read())
print(f'Done. File size: {sz/1024:.0f} KB')
print(f'Years: {YRS}, Latest: {latest_year}/{latest_month}')
print(f'CN groups top5: {cn_groups[:5]}')
print(f'Month keys: {MK[:3]}...{MK[-3:]}')
