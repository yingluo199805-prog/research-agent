# -*- coding: utf-8 -*-
"""Generate aggregated_caam.json from 中汽协CV sheet."""
import os, json, re, pandas as pd, warnings
warnings.filterwarnings('ignore')

HERE = os.path.dirname(os.path.abspath(__file__))

# ══ 月更只需改这一行 ══
DATA_FILE = os.environ.get('DASHBOARD_DATA_FILE', '海关出口数据-260612_working.xlsx')
# ═══════════════════

xlsx = os.path.join(HERE, DATA_FILE)
# 中汽协 数据本轮未更新（仍为 260419 raw 结果），主文件改名到 260507 是因海关数据更新
# 因此 CAAM dataFileDate 锁定为 2026-04-19；下次中汽协 raw 更新时改回从文件名解析
DATA_FILE_DATE = os.environ.get('DASHBOARD_DATA_FILE_DATE', '2026-06-27')
df = pd.read_excel(xlsx, sheet_name='中汽协CV ')
df.columns = ['year','month_seq','month_label','group','oem_old','brand','model','qty','series','oem']

# 清理
df = df[df['oem'].notna()]
df = df[df['qty'].notna()]
df['qty'] = df['qty'].astype(int)
# 标记无具体车型名称的行（NaN/空白），后续汇总为"其他"
df['model'] = df['model'].fillna('其他').astype(str).str.strip()
df.loc[df['model']=='', 'model'] = '其他'
df.loc[df['model'].str.lower()=='nan', 'model'] = '其他'

# ── 常量 ──
YRS = [str(y) for y in range(2020, 2027)]  # 2020-2026
S_ORD = ['中系','日系','欧系','美系','韩系']

def has_stock_code(name):
    return bool(re.search(r'[AHU][S.]?\d{3,6}|US\.\w+', str(name)))

def yoy(cur, prev):
    if prev and prev > 0:
        return round((cur/prev - 1)*100, 1)
    return None

def add_yoy(d, yrs=YRS):
    for i in range(1, len(yrs)):
        v = yoy(d.get(yrs[i],0), d.get(yrs[i-1],0))
        if v is not None: d[yrs[i]+'_yoy'] = v
    return d

# ── 年度数据：每年汇总 ──
# 2026年只有M1-2，同比需对齐到2025年M1-2
# 先确定2026年有哪些月份
latest_year = int(df['year'].max())
latest_months_in_year = sorted(df[df['year']==latest_year]['month_seq'].unique().tolist())
PARTIAL_YEAR = str(latest_year) if latest_months_in_year != list(range(1,13)) else None
PARTIAL_MONTHS = latest_months_in_year if PARTIAL_YEAR else []
print(f'Partial year: {PARTIAL_YEAR}, months: {PARTIAL_MONTHS}')

annual = df.groupby(['year','oem','series','model'])['qty'].sum().reset_index()

def add_yoy_annual(d, yrs=YRS):
    """计算年度同比，对不完整年份用同期可比数据"""
    for i in range(1, len(yrs)):
        cur_y = yrs[i]
        prev_y = yrs[i-1]
        # 如果当前年是不完整年份，已经在数据中做了同期截断，直接计算
        v = yoy(d.get(cur_y,0), d.get(prev_y,0))
        if v is not None: d[cur_y+'_yoy'] = v
    return d

# ── 月度数据：从2025年1月起至最新月（含同比）──
max_year = int(df['year'].max())
max_month = int(df[df['year']==max_year]['month_seq'].max())
# 从2025年1月到最新月
recent_months = []
y, m = 2025, 1
while (y, m) <= (max_year, max_month):
    recent_months.append((y, m))
    m += 1
    if m > 12:
        m = 1
        y += 1
print(f'Latest data: {max_year}/{max_month}')
print(f'Recent 12 months: {recent_months}')

# 同比月份（前一年同月）
yoy_months = [(y-1, m) for y, m in recent_months]

monthly_df = df[df.apply(lambda r: (int(r['year']), int(r['month_seq'])) in recent_months, axis=1)].copy()
yoy_df = df[df.apply(lambda r: (int(r['year']), int(r['month_seq'])) in yoy_months, axis=1)].copy()

# 月度key: "2026-1", "2026-2" etc
monthly_df['mk'] = monthly_df['year'].astype(str) + '-' + monthly_df['month_seq'].astype(str)
yoy_df['mk'] = (yoy_df['year']+1).astype(str) + '-' + yoy_df['month_seq'].astype(str)  # 对齐到今年

MK = [f'{y}-{m}' for y,m in recent_months]
ML = {}
for y,m in recent_months:
    ML[f'{y}-{m}'] = f'{y}/{m}月'

agg = {}

# ═══ helper: 年度聚合（对不完整年份用同期可比口径）═══
# 从原始df做年度聚合（保留month_seq信息）
def annual_qty_comparable(src_df, year_str, filters=None):
    """获取年度数据。不完整年份=仅有月份的数据（全年数据不可用）。
    其他年份正常取全年数据。"""
    yi = int(year_str)
    sub = src_df.copy()
    if filters:
        for col, val in filters.items():
            sub = sub[sub[col]==val]
    if PARTIAL_YEAR and year_str == PARTIAL_YEAR:
        # 不完整年份：只取已有月份
        return int(sub[(sub['year']==yi)&(sub['month_seq'].isin(PARTIAL_MONTHS))]['qty'].sum())
    else:
        # 完整年份：取全年
        return int(sub[sub['year']==yi]['qty'].sum())

def add_yoy_annual(d, yrs=YRS):
    """计算年度同比。对不完整年份(2026)，同比基数用前一年同期月份"""
    for i in range(1, len(yrs)):
        cur_y = yrs[i]
        prev_y = yrs[i-1]
        if PARTIAL_YEAR and cur_y == PARTIAL_YEAR:
            # 不完整年份的同比：需要用前一年同期月份作为基数
            # 但 d[prev_y] 是全年数据，需要单独计算同期
            # 这里跳过，在外部单独处理
            pass
        else:
            v = yoy(d.get(cur_y,0), d.get(prev_y,0))
            if v is not None: d[cur_y+'_yoy'] = v
    return d

def add_partial_yoy(d, src_df, filters=None):
    """为不完整年份(如2026M1-2)计算同比：用前一年同期月份数据做基数"""
    if not PARTIAL_YEAR: return d
    prev_y = str(int(PARTIAL_YEAR) - 1)
    sub = src_df.copy()
    if filters:
        for col, val in filters.items():
            sub = sub[sub[col]==val]
    prev_comparable = int(sub[(sub['year']==int(prev_y))&(sub['month_seq'].isin(PARTIAL_MONTHS))]['qty'].sum())
    v = yoy(d.get(PARTIAL_YEAR,0), prev_comparable)
    if v is not None: d[PARTIAL_YEAR+'_yoy'] = v
    return d

# ═══ 1. 总量（年度 + 月度） ═══
# 注意：annual_qty_comparable 需要有 month_seq 列，所以用 df（原始数据）
total_annual = {}
for y in YRS:
    total_annual[y] = annual_qty_comparable(df, y)
add_yoy_annual(total_annual)
add_partial_yoy(total_annual, df)
agg['total'] = {'中汽协乘用车总出口': total_annual}

total_monthly = {}
for mk in MK:
    total_monthly[mk] = int(monthly_df[monthly_df['mk']==mk]['qty'].sum())
# 月度同比
for mk in MK:
    cur = total_monthly.get(mk, 0)
    prev = int(yoy_df[yoy_df['mk']==mk]['qty'].sum()) if len(yoy_df[yoy_df['mk']==mk]) > 0 else 0
    v = yoy(cur, prev)
    if v is not None: total_monthly[mk+'_yoy'] = v
agg['monthly_total'] = {'中汽协乘用车总出口': total_monthly}

# ═══ 2. 分系别（年度 + 月度） ═══
series_annual = {}
for s in S_ORD:
    d = {}
    for y in YRS:
        d[y] = annual_qty_comparable(df, y, {'series': s})
    add_yoy_annual(d)
    add_partial_yoy(d, df, {'series': s})
    series_annual[s] = d
agg['bySeries'] = series_annual

series_monthly = {}
for s in S_ORD:
    d = {}
    sub = monthly_df[monthly_df['series']==s]
    sub_yoy = yoy_df[yoy_df['series']==s]
    for mk in MK:
        d[mk] = int(sub[sub['mk']==mk]['qty'].sum())
    for mk in MK:
        cur = d.get(mk,0)
        prev = int(sub_yoy[sub_yoy['mk']==mk]['qty'].sum()) if len(sub_yoy[sub_yoy['mk']==mk])>0 else 0
        v = yoy(cur, prev)
        if v is not None: d[mk+'_yoy'] = v
    series_monthly[s] = d
agg['monthly_bySeries'] = series_monthly

# ═══ 3. 重点车企（年度 + 月度）═══
# 按最新完整年排序
latest_full_year = max_year if max_month == 12 else max_year - 1
rank_year = str(max(int(YRS[-1]), latest_full_year))
# 按2026年以来累计销量排序
rank_yr = latest_year  # 2026
oem_rank = df[df['year']==rank_yr].groupby('oem')['qty'].sum().sort_values(ascending=False)
all_oems = oem_rank.index.tolist()
coded_oems = [o for o in all_oems if has_stock_code(o)]
uncoded_oems = [o for o in all_oems if not has_stock_code(o)]
coded_top15 = coded_oems[:15]

oem_annual = {}
for oem in all_oems:
    d = {}
    for y in YRS:
        d[y] = annual_qty_comparable(df, y, {'oem': oem})
    add_yoy_annual(d)
    add_partial_yoy(d, df, {'oem': oem})
    oem_annual[oem] = d

oem_monthly = {}
for oem in all_oems:
    d = {}
    sub = monthly_df[monthly_df['oem']==oem]
    sub_yoy = yoy_df[yoy_df['oem']==oem]
    for mk in MK:
        d[mk] = int(sub[sub['mk']==mk]['qty'].sum())
    for mk in MK:
        cur = d.get(mk,0)
        prev = int(sub_yoy[sub_yoy['mk']==mk]['qty'].sum()) if len(sub_yoy[sub_yoy['mk']==mk])>0 else 0
        v = yoy(cur, prev)
        if v is not None: d[mk+'_yoy'] = v
    oem_monthly[oem] = d

agg['oemAnnual'] = {'order': all_oems, 'data': oem_annual}
agg['oemMonthly'] = {'order': all_oems, 'data': oem_monthly}
agg['codedTop15'] = coded_top15
agg['codedAll'] = coded_oems
agg['uncodedOems'] = uncoded_oems
agg['oemSeries'] = annual.groupby('oem')['series'].first().to_dict()

# ═══ 4. 分车企-分车型（年度 + 月度）═══
oem_model_annual = {}
for oem in all_oems:
    sub = annual[annual['oem']==oem]
    model_rank = sub[sub['year']==latest_year].groupby('model')['qty'].sum().sort_values(ascending=False)
    top_models = model_rank.head(20).index.tolist()
    md = {}
    oem_df = df[df['oem']==oem]
    for model in top_models:
        d = {}
        for y in YRS:
            d[y] = annual_qty_comparable(oem_df, y, {'model': model})
        add_yoy_annual(d)
        add_partial_yoy(d, oem_df, {'model': model})
        if any(d.get(y,0)>0 for y in YRS): md[model] = d
    oem_model_annual[oem] = md

oem_model_monthly = {}
for oem in all_oems:
    sub = monthly_df[monthly_df['oem']==oem]
    sub_yoy_o = yoy_df[yoy_df['oem']==oem]
    model_rank = sub.groupby('model')['qty'].sum().sort_values(ascending=False)
    top_models = model_rank.head(20).index.tolist()
    md = {}
    for model in top_models:
        d = {}
        msub = sub[sub['model']==model]
        msub_yoy = sub_yoy_o[sub_yoy_o['model']==model]
        for mk in MK:
            d[mk] = int(msub[msub['mk']==mk]['qty'].sum())
        for mk in MK:
            cur = d.get(mk,0)
            prev = int(msub_yoy[msub_yoy['mk']==mk]['qty'].sum()) if len(msub_yoy[msub_yoy['mk']==mk])>0 else 0
            v = yoy(cur, prev)
            if v is not None: d[mk+'_yoy'] = v
        if any(d.get(mk,0)>0 for mk in MK): md[model] = d
    oem_model_monthly[oem] = md

agg['oemModelAnnual'] = oem_model_annual
agg['oemModelMonthly'] = oem_model_monthly

# ═══ 5. 品牌级数据（仅 coded OEMs）═══
oem_brands = {}
brand_model_annual = {}
brand_model_monthly = {}
for oem in coded_oems:
    osub = df[df['oem']==oem]
    brand_rank = osub[osub['year']==latest_year].groupby('brand')['qty'].sum().sort_values(ascending=False)
    brands = [{'name': b, 'sales': int(s)} for b, s in brand_rank.items() if s > 0]
    oem_brands[oem] = brands
    for bi in brands:
        bname = bi['name']
        bdf = osub[osub['brand']==bname]
        if bdf.empty: continue
        # 年度 model data
        model_rank_b = bdf[bdf['year']==latest_year].groupby('model')['qty'].sum().sort_values(ascending=False)
        top_models_b = model_rank_b.head(20).index.tolist()
        md = {}
        for model in top_models_b:
            d = {}
            for y in YRS:
                d[y] = annual_qty_comparable(bdf, y, {'model': model})
            add_yoy_annual(d)
            add_partial_yoy(d, bdf, {'model': model})
            if any(d.get(y,0)>0 for y in YRS): md[model] = d
        brand_model_annual[bname] = md
        # 月度 model data
        bsub_m = monthly_df[monthly_df['oem']==oem]
        bsub_m = bsub_m[bsub_m['brand']==bname]
        bsub_yoy = yoy_df[yoy_df['oem']==oem]
        bsub_yoy = bsub_yoy[bsub_yoy['brand']==bname]
        model_rank_bm = bsub_m.groupby('model')['qty'].sum().sort_values(ascending=False)
        top_models_bm = model_rank_bm.head(20).index.tolist()
        mdm = {}
        for model in top_models_bm:
            d = {}
            msub = bsub_m[bsub_m['model']==model]
            msub_yoy = bsub_yoy[bsub_yoy['model']==model]
            for mk in MK:
                d[mk] = int(msub[msub['mk']==mk]['qty'].sum())
            for mk in MK:
                cur = d.get(mk,0)
                prev = int(msub_yoy[msub_yoy['mk']==mk]['qty'].sum()) if len(msub_yoy[msub_yoy['mk']==mk])>0 else 0
                v = yoy(cur, prev)
                if v is not None: d[mk+'_yoy'] = v
            if any(d.get(mk,0)>0 for mk in MK): mdm[model] = d
        brand_model_monthly[bname] = mdm

agg['oemBrands'] = oem_brands
agg['brandModelAnnual'] = brand_model_annual
agg['brandModelMonthly'] = brand_model_monthly

# ═══ 元数据 ═══
agg['meta'] = {
    'years': YRS,
    'monthKeys': MK,
    'monthLabels': ML,
    'latestYear': max_year,
    'latestMonth': max_month,
    'partialYear': PARTIAL_YEAR,
    'partialMonths': PARTIAL_MONTHS,
    'dataFileDate': DATA_FILE_DATE,
}

# ═══ Save ═══
with open('aggregated_caam.json', 'w', encoding='utf-8') as f:
    json.dump(agg, f, ensure_ascii=False, separators=(',',':'))

sz = len(open('aggregated_caam.json', encoding='utf-8').read())
print(f'Done. File size: {sz/1024:.0f} KB')
print(f'Years: {YRS}')
print(f'OEMs total: {len(all_oems)}, coded: {len(coded_oems)}')
print(f'codedTop15: {coded_top15}')
print(f'Month keys: {MK}')
