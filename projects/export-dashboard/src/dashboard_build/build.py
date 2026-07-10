"""Build the export dashboard HTML with embedded aggregated data."""
import json, os

HERE = os.path.dirname(os.path.abspath(__file__))
data = json.load(open(os.path.join(HERE, 'aggregated.json'), encoding='utf-8'))
DATA_JSON = json.dumps(data, ensure_ascii=False, separators=(',', ':'))

# 中汽协数据
caam_path = os.path.join(HERE, 'aggregated_caam.json')
if os.path.exists(caam_path):
    caam_data = json.load(open(caam_path, encoding='utf-8'))
    CAAM_JSON = json.dumps(caam_data, ensure_ascii=False, separators=(',', ':'))
else:
    CAAM_JSON = 'null'

# Marklines数据
mkls_path = os.path.join(HERE, 'aggregated_mkls.json')
if os.path.exists(mkls_path):
    mkls_data = json.load(open(mkls_path, encoding='utf-8'))
    MKLS_JSON = json.dumps(mkls_data, ensure_ascii=False, separators=(',', ':'))
else:
    MKLS_JSON = 'null'

HTML = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>中国乘用车海外数据看板</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
:root{
  --bg:#0A1628;--card:#0F2044;--hover:#132952;--input:#1a3060;
  --bd:#1e3a6e;--gold:#C9A84C;--blue:#4A90E2;--teal:#5BC4A0;
  --t1:#F0F4FF;--t2:#8A9DC0;--t3:#4A6080;
  --pos:#E85A5A;--neg:#4CAF82;
}
body{background:var(--bg);color:var(--t1);
  font-family:"PingFang SC","Microsoft YaHei","Noto Sans SC",sans-serif;
  font-size:14px;min-height:100vh}
.hd{background:linear-gradient(135deg,#0F2044,#132952);border-bottom:1px solid var(--bd);
  padding:0 32px;height:60px;display:flex;align-items:center;justify-content:space-between;
  position:sticky;top:38px;z-index:100}
/* Data source tabs */
.src-tabs{display:flex;gap:0;background:#0D1B32;border-bottom:1px solid var(--bd);
  position:sticky;top:0;z-index:101;padding:0 32px}
.src-tab{padding:9px 28px;font-size:14px;font-weight:600;color:var(--t2);
  cursor:pointer;border-bottom:2px solid transparent;transition:all .2s;user-select:none}
.src-tab.on{color:var(--gold);border-bottom-color:var(--gold)}
.src-tab:hover:not(.on){color:var(--t1)}
.src-placeholder{display:flex;align-items:center;justify-content:center;
  min-height:60vh;font-size:18px;color:var(--t2);flex-direction:column;gap:12px}
.data-footnote{background:var(--hover);border:1px solid var(--bd);border-radius:10px;
  margin:24px 28px 32px;padding:16px 20px;font-size:12px;color:var(--t2);line-height:1.8}
.data-footnote .fn-title{font-size:14px;font-weight:700;color:var(--gold);margin-bottom:8px}
.data-footnote .fn-item{margin-bottom:4px}
.data-footnote .fn-label{color:var(--t1);font-weight:600}
.hd-title{font-size:19px;font-weight:700;letter-spacing:1px}
.hd-title .acc{color:var(--gold)}
.hd-meta{font-size:12px;color:var(--t2)}
.main{padding:20px 28px;max-width:1600px;margin:0 auto}
/* Section */
.sec{background:var(--card);border:1px solid var(--bd);border-radius:10px;margin-bottom:18px;overflow:hidden}
.sec-hd{padding:12px 18px;border-bottom:1px solid var(--bd);display:flex;align-items:center;justify-content:space-between;
  background:linear-gradient(90deg,var(--hover) 0%,transparent 70%)}
.sec-title{font-size:18px;font-weight:700;display:flex;align-items:center;gap:8px}
.sec-bar{width:3px;height:14px;background:var(--gold);border-radius:2px;flex-shrink:0}
.sec-sub{font-size:11px;color:var(--t2)}
.sec-body{padding:14px 16px 18px}
/* Chart wrap */
.cw{background:var(--bg);border:1px solid var(--bd);border-radius:8px;padding:12px 10px 8px;position:relative}
.cw-lbl{font-size:17px;color:var(--gold);margin-bottom:4px;letter-spacing:.3px;font-weight:700}
.cw-note{font-size:12px;color:#C9D6EC;margin-top:2px;font-style:italic}
/* Grids */
.g2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.g12{display:grid;grid-template-columns:1fr 2fr;gap:14px}
.g13{display:grid;grid-template-columns:1fr 3fr;gap:14px}
.mt14{margin-top:14px}
/* Tab buttons */
.tab-bar{display:flex;gap:2px;background:var(--input);border-radius:6px;padding:2px;margin-left:16px}
.tab-btn{padding:5px 16px;border-radius:4px;border:none;background:transparent;
  color:var(--t2);font-size:13px;font-weight:700;cursor:pointer;transition:all .18s;font-family:inherit}
.tab-btn.on{background:var(--blue);color:#fff;font-weight:600}
.tab-btn:hover:not(.on){color:var(--t1)}
/* Year filter */
.fbar{background:var(--hover);border:1px solid var(--bd);border-radius:8px;
  padding:10px 14px;margin-bottom:14px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.flbl{font-size:16px;color:var(--t1);margin-right:2px;white-space:nowrap;font-weight:700}
.ychk{display:inline-flex;align-items:center;gap:5px;cursor:pointer;font-size:13px;font-weight:700;color:#C9D6EC;
  padding:5px 13px;border-radius:5px;border:1px solid var(--bd);background:transparent;transition:all .18s;user-select:none;
  opacity:.85}
.ychk.on{color:#fff;opacity:1;background:var(--blue);border-color:var(--blue)}
.ychk input{display:none}
.ychk.y23{border-color:rgba(74,144,226,.4)}.ychk.y24{border-color:rgba(201,168,76,.4)}.ychk.y25{border-color:rgba(91,196,160,.4)}
.ychk.y23.on{background:#4A90E2;border-color:#4A90E2;color:#fff}
.ychk.y24.on{background:#C9A84C;border-color:#C9A84C;color:#fff}
.ychk.y25.on{background:#5BC4A0;border-color:#5BC4A0;color:#fff}
.ydot{width:8px;height:8px;border-radius:50%}
.y23 .ydot{background:#4A90E2}.y24 .ydot{background:#C9A84C}.y25 .ydot{background:#5BC4A0}
/* Legend note */
.legend-note{font-size:10px;color:var(--t2);text-align:right;padding:2px 8px 0;
  display:flex;align-items:center;justify-content:flex-end;gap:10px;transition:opacity .2s}
.legend-note span.pos{color:var(--pos)}.legend-note span.neg{color:var(--neg)}
.legend-note.hide{opacity:0;height:0;overflow:hidden;padding:0;margin:0}
/* Export button */
.cw{position:relative}
.exp-btn{position:absolute;bottom:6px;right:8px;background:var(--input);border:1px solid var(--bd);
  color:var(--t2);font-size:10px;padding:2px 8px;border-radius:4px;cursor:pointer;
  opacity:.5;transition:opacity .2s;z-index:10;font-family:inherit}
.exp-btn:hover{opacity:1;color:var(--gold);border-color:var(--gold)}
/* 图表内嵌筛选项缩小 */
.cw .ychk{font-size:12px;padding:3px 10px}
/* Data source bar */
.hd-src{font-size:13px;color:var(--gold);text-align:center;padding:4px 0;letter-spacing:.3px;font-weight:600}
/* Filter vertical layout */
.fbar-v{background:var(--hover);border:1px solid var(--bd);border-radius:8px;
  padding:10px 14px;margin-bottom:14px;display:flex;flex-direction:column;gap:10px}
.fbar-v .frow{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
/* Show more button */
.show-more-btn{background:var(--input);border:1px solid var(--bd);color:var(--t2);
  font-size:12px;padding:6px 18px;border-radius:6px;cursor:pointer;margin:10px auto 0;
  display:block;transition:all .2s;font-family:inherit}
.show-more-btn:hover{color:var(--gold);border-color:var(--gold)}
/* Modal overlay */
.modal-mask{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.7);
  z-index:200;display:flex;align-items:center;justify-content:center}
.modal-box{background:var(--card);border:1px solid var(--bd);border-radius:12px;
  width:90%;max-width:1200px;max-height:90vh;overflow-y:auto;padding:20px 24px;position:relative}
.modal-close{position:absolute;top:10px;right:14px;background:none;border:none;color:var(--t2);
  font-size:22px;cursor:pointer;z-index:10}
.modal-close:hover{color:var(--gold)}
</style>
</head>
<body>
<!-- 登录/进入看板公告弹窗：每会话弹一次 -->
<div id="loginNoticeMask" style="display:none;position:fixed;inset:0;background:rgba(5,14,30,0.78);z-index:99999;align-items:center;justify-content:center;backdrop-filter:blur(2px)">
  <div style="width:min(520px,92vw);background:#0F2044;border:1px solid #C9A84C;border-radius:12px;padding:24px 28px;box-shadow:0 12px 48px rgba(0,0,0,0.55);font-family:'PingFang SC','Microsoft YaHei',sans-serif;color:#E8EEF8">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">
      <div style="width:6px;height:22px;background:#C9A84C;border-radius:2px"></div>
      <div style="font-size:17px;font-weight:700;letter-spacing:1px">数据使用提示</div>
    </div>
    <div style="font-size:14px;line-height:1.75;color:#C9D6EC">
      Marklines 数据源的 <b style="color:#E8C870">2026 年 4-5 月</b> 仍有部分国家数据缺失，
      <b style="color:#E8C870">销量及增速仅供参考</b>，我们在持续更新中。
    </div>
    <div style="font-size:12px;line-height:1.7;color:#8A9DC0;margin-top:12px;padding-top:10px;border-top:1px solid #1e3a6e">
      如对数据有疑问或合作需求，欢迎联系 · <span style="color:#C9A84C;font-weight:600">罗英 15766781877</span>
    </div>
    <div style="display:flex;justify-content:flex-end;margin-top:18px">
      <button id="loginNoticeOk" style="padding:8px 22px;background:linear-gradient(135deg,#C9A84C,#E8C870);color:#0A1628;border:none;border-radius:6px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit;letter-spacing:1px">我 知 道 了</button>
    </div>
  </div>
</div>
<script>
(function(){
  try{
    if(sessionStorage.getItem('mklsNoticeShown2026M45')) return;
    const mask = document.getElementById('loginNoticeMask');
    if(!mask) return;
    mask.style.display = 'flex';
    function closeNotice(){
      mask.style.display = 'none';
      try{ sessionStorage.setItem('mklsNoticeShown2026M45', '1'); }catch(e){}
    }
    document.getElementById('loginNoticeOk').addEventListener('click', closeNotice);
    mask.addEventListener('click', function(e){ if(e.target===mask) closeNotice(); });
    document.addEventListener('keydown', function onEsc(e){
      if(e.key==='Escape' && mask.style.display!=='none'){ closeNotice(); document.removeEventListener('keydown', onEsc); }
    });
  }catch(e){}
})();
</script>

<!-- 数据源切换Tab -->
<div class="src-tabs" id="srcTabs">
  <div class="src-tab" data-src="zhongqi">中汽协出口（批发）</div>
  <div class="src-tab" data-src="haiguan">海关总署出口（批发）</div>
  <div class="src-tab" data-src="mkls">Marklines销量（终端）</div>
  <div class="src-tab" data-src="inventory">中国品牌海外库存</div>
  <div class="src-tab" data-src="localProd">本地生产</div>
  <div class="src-tab" data-src="priceTrack">价格跟踪</div>
</div>
<div id="srcInventory" style="display:none">
  <div style="display:flex;align-items:center;justify-content:center;min-height:60vh;color:var(--t2)">
    <div style="text-align:center">
      <div style="font-size:48px;margin-bottom:20px;opacity:0.3">📦</div>
      <div style="font-size:20px;font-weight:600;color:var(--t1);margin-bottom:10px">中国品牌海外库存数据</div>
      <div style="font-size:14px;color:var(--t2);margin-bottom:8px">数据正在陆续接入中，敬请期待</div>
      <div style="font-size:14px;color:var(--t2);margin-bottom:8px">已有部分车企数据，欢迎联系交流</div>
      <div style="font-size:15px;color:var(--gold);font-weight:600">罗英 15766781877</div>
    </div>
  </div>
</div>
<div id="srcLocalProd" style="display:none">
  <div style="display:flex;align-items:center;justify-content:center;min-height:60vh;color:var(--t2)">
    <div style="text-align:center">
      <div style="font-size:48px;margin-bottom:20px;opacity:0.3">🏭</div>
      <div style="font-size:20px;font-weight:600;color:var(--t1);margin-bottom:10px">本地生产数据</div>
      <div style="font-size:14px;color:var(--t2)">数据正在陆续接入中，敬请期待</div>
    </div>
  </div>
</div>
<div id="srcPriceTrack" style="display:none">
  <div style="display:flex;align-items:center;justify-content:center;min-height:60vh;color:var(--t2)">
    <div style="text-align:center">
      <div style="font-size:48px;margin-bottom:20px;opacity:0.3">💰</div>
      <div style="font-size:20px;font-weight:600;color:var(--t1);margin-bottom:10px">价格跟踪数据</div>
      <div style="font-size:14px;color:var(--t2)">数据正在陆续接入中，敬请期待</div>
    </div>
  </div>
</div>
<div id="srcHaiguan" style="display:none">
<div class="hd">
  <div class="hd-title">中国<span class="acc">乘用车</span>海外数据看板 · 海关总署</div>
  <div style="text-align:right">
    <div class="hd-meta" id="haiguanMeta">数据加载中</div>
    <div class="hd-src">数据来源：海关总署 | 联系人：罗英 15766781877</div>
  </div>
</div>
<div class="data-footnote" style="margin:0 28px 14px;padding:10px 20px">
  <div class="fn-item" style="margin:0"><span class="fn-label">口径：</span>中国海关统计的整车新车出口数量（不含海外生产） <span class="fn-label">统计维度：</span>分区域、分国家、分动力类型、分车企、分车型 <span class="fn-label">更新频率：</span>双月频 <span class="fn-label">差异说明：</span>统计数据与中汽协的差异来源于：时间差、海外生产等</div>
</div>
<div class="main">

<!-- ═══ 板块0 出口总览 ═══ -->
<div class="sec" id="sec0">
  <div class="sec-hd">
    <div style="display:flex;align-items:center">
      <div class="sec-title"><span class="sec-bar"></span>出口总览</div>
      <div class="tab-bar" id="s0tabs">
        <button class="tab-btn on" data-mode="annual">年度</button>
        <button class="tab-btn" data-mode="monthly">月度</button>
      </div>
    </div>
    <div class="sec-sub"><span id="s0sub">各年度总量与同比增速</span></div>
  </div>
  <div class="sec-body">
    <!-- ▸ 年度模式 -->
    <div id="s0annual">
      <div class="g2">
        <div class="cw">
          <div class="cw-lbl">乘用车出口总量</div>
          <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="c0tot" style="height:280px"></div>
        </div>
        <div class="cw">
          <div class="cw-lbl">分系别出口量</div>
          <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="c0ser" style="height:280px"></div>
          <div class="cw-note">按品牌系别分类；中系=自主品牌</div>
        </div>
      </div>
      <div class="g2 mt14">
        <div class="cw">
          <div class="cw-lbl">分动力类型出口量</div>
          <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="c0eng" style="height:300px"></div>
          <div class="cw-note">新能源 = 纯电动 + 插电混动 + 增程式</div>
        </div>
        <div class="cw">
          <div class="cw-lbl">分区域出口量</div>
          <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="c0reg" style="height:300px"></div>
          <div class="cw-note">欧洲拆分为俄罗斯和非俄两部分</div>
        </div>
      </div>
      <div class="cw mt14">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
          <div class="cw-lbl" style="margin-bottom:0">重点车企出口量（按2025年降序）</div>
          <div id="s0oemExtra" style="display:flex;gap:4px;flex-wrap:wrap;align-items:center"></div>
        </div>
        <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="c0oem" style="height:420px"></div>
        <div class="cw-note">默认展示A/H/US上市车企TOP15；更多车企可筛选追加；支持放大图</div>
      </div>
    </div>
    <!-- ▸ 月度模式（JS动态生成） -->
    <div id="s0monthly" style="display:none"></div>
  </div>
</div>

<!-- ═══ 板块一 分动力类型视角 ═══ -->
<div class="sec" id="sec1">
  <div class="sec-hd">
    <div class="sec-title"><span class="sec-bar"></span>分系别·分动力类型视角</div>
    <div class="sec-sub">按系别、动力类型、年度/月度多维筛选</div>
  </div>
  <div class="sec-body">
    <!-- 筛选栏（竖排） -->
    <div class="fbar-v" id="s1filter">
      <div class="frow">
        <span class="flbl">时间维度</span>
        <div class="tab-bar" id="s1timeTabs">
          <button class="tab-btn on" data-mode="annual">年度</button>
          <button class="tab-btn" data-mode="monthly">月度</button>
        </div>
      </div>
      <div class="frow" id="s1yearBox">
        <span class="flbl">年份</span>
        <!-- 年份 checkbox 由 JS 根据 DATA.meta.years 动态生成 -->
      </div>
      <div class="frow" id="s1seriesBox">
        <span class="flbl">系别</span>
        <label class="ychk on" id="s1s_all"><input type="radio" name="s1ser" value="all" checked><span class="ydot" style="background:#C9A84C"></span>全部</label>
        <label class="ychk" id="s1s_cn"><input type="radio" name="s1ser" value="中系"><span class="ydot" style="background:#4A90E2"></span>中系</label>
        <label class="ychk" id="s1s_jp"><input type="radio" name="s1ser" value="日系"><span class="ydot" style="background:#5BC4A0"></span>日系</label>
        <label class="ychk" id="s1s_eu"><input type="radio" name="s1ser" value="欧系"><span class="ydot" style="background:#E87B5A"></span>欧系</label>
        <label class="ychk" id="s1s_us"><input type="radio" name="s1ser" value="美系"><span class="ydot" style="background:#9B7FD4"></span>美系</label>
        <label class="ychk" id="s1s_kr"><input type="radio" name="s1ser" value="韩系"><span class="ydot" style="background:#5BC4D4"></span>韩系</label>
      </div>
      <div class="frow">
        <span class="flbl">动力类型</span>
        <label class="ychk on" id="s1e_all"><input type="checkbox" value="all" checked><span class="ydot" style="background:#C9A84C"></span>全部</label>
        <label class="ychk" id="s1e_nev"><input type="checkbox" value="新能源"><span class="ydot" style="background:#5BC4A0"></span>新能源</label>
        <label class="ychk" id="s1e_bev"><input type="checkbox" value="纯电动"><span class="ydot" style="background:#4A90E2"></span>纯电动</label>
        <label class="ychk" id="s1e_phev"><input type="checkbox" value="插电混动"><span class="ydot" style="background:#5BC4D4"></span>插电混动</label>
        <label class="ychk" id="s1e_erev"><input type="checkbox" value="增程式"><span class="ydot" style="background:#E87B5A"></span>增程式</label>
        <label class="ychk" id="s1e_ice"><input type="checkbox" value="燃油车"><span class="ydot" style="background:#9B7FD4"></span>燃油车</label>
      </div>
    </div>
    <!-- 图表区：区域 + 国家明细 同一行 -->
    <div class="g2">
      <div class="cw" id="s1regWrap">
        <div class="cw-lbl" id="s1regLbl">按区域</div>
        <div class="legend-note" id="s1regLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="c1reg" style="height:340px"></div>
        <div class="cw-note">联动系别+动力类型筛选；默认前5区域</div>
      </div>
      <div class="cw" id="s1countryWrap">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
          <div class="cw-lbl" style="margin-bottom:0" id="s1drillLbl">国家明细 · 亚洲</div>
          <div id="s1regFilter" style="display:flex;gap:4px;flex-wrap:wrap"></div>
        </div>
        <div class="legend-note" id="s1countryLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="c1country" style="height:340px"></div>
        <div class="cw-note">联动系别+动力类型筛选；点击区域按钮切换国家明细；默认前5国家</div>
      </div>
    </div>
    <!-- 分车企排名（年度全展示 / 月度筛选） -->
    <div class="cw mt14" id="s1oemBox">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap">
        <div class="cw-lbl" style="margin-bottom:0" id="s1oemLbl">重点车企出口量（按2025年降序）</div>
        <div style="display:flex;gap:6px;align-items:center">
          <div class="fbar" id="s1oemFilter" style="margin:0;padding:6px 10px;background:transparent;border:none;flex-wrap:wrap;gap:6px;display:none"></div>
          <button class="show-more-btn" id="s1oemMoreBtn" style="margin:0;padding:3px 12px;font-size:11px;display:none">查看更多</button>
          <button class="show-more-btn" id="s1oemZoomBtn" style="margin:0;padding:3px 12px;font-size:11px;display:none">放大图</button>
        </div>
      </div>
      <div class="legend-note" id="s1oemLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
      <div id="c1oem" style="height:420px"></div>
      <div class="cw-note">联动系别+动力类型筛选；年度默认展示上市车企TOP10；月度默认展示2025年销量TOP1车企；更多车企可筛选追加</div>
    </div>
  </div>
</div>

<!-- ═══ 板块二 分车企视角 ═══ -->
<div class="sec" id="sec2">
  <div class="sec-hd">
    <div class="sec-title"><span class="sec-bar"></span>分车企视角</div>
    <div class="sec-sub">按车企查看动力结构、区域分布、国家及车型明细</div>
  </div>
  <div class="sec-body">
    <!-- 筛选栏（竖排） -->
    <div class="fbar-v" id="s2filter">
      <div class="frow">
        <span class="flbl">时间维度</span>
        <div class="tab-bar" id="s2timeTabs">
          <button class="tab-btn on" data-mode="annual">年度</button>
          <button class="tab-btn" data-mode="monthly">月度</button>
        </div>
      </div>
      <div class="frow" id="s2yearBox">
        <span class="flbl">年份</span>
        <!-- 年份 radio 由 JS 根据 DATA.meta.years 动态生成（含"全部"项） -->
      </div>
      <div class="frow" id="s2oemRow">
        <span class="flbl">选择车企</span>
        <div class="cw-note" id="s2oemHint" style="margin:0;font-style:normal;color:var(--t2)">筛选规则：选"全部"年份时为单选；选具体年份后支持车企复选对比</div>
      </div>
      <div class="frow">
        <div id="s2oemBtns" style="display:flex;gap:4px;flex-wrap:wrap"></div>
        <select id="s2oemSelect" style="background:var(--input);color:var(--t1);border:1px solid var(--bd);border-radius:4px;padding:4px 8px;font-size:12px;font-family:inherit;cursor:pointer">
          <option value="">更多车企 ▾</option>
        </select>
      </div>
      <div class="frow" id="s2brandBox" style="display:none">
        <span class="flbl">品牌筛选</span>
        <div id="s2brandBtns" style="display:flex;gap:4px;flex-wrap:wrap"></div>
      </div>
    </div>
    <!-- 图表区 -->
    <div class="g2">
      <div class="cw">
        <div class="cw-lbl" id="s2engLbl">分动力类型</div>
        <div class="legend-note" id="s2engLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="c2eng" style="height:300px"></div>
        <div class="cw-note">注：新能源 = 纯电动 + 插电混动 + 增程式</div>
      </div>
      <div class="cw">
        <div class="cw-lbl" id="s2shareLbl">动力类型份额占比</div>
        <div id="c2share" style="height:300px"></div>
        <div class="cw-note">受车企筛选联动</div>
      </div>
    </div>
    <div class="g2 mt14">
      <div class="cw">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
          <div class="cw-lbl" style="margin-bottom:0" id="s2regLbl">分区域</div>
          <div id="s2regEnergyFilter" style="display:flex;gap:3px;flex-wrap:wrap"></div>
        </div>
        <div class="legend-note" id="s2regLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="c2reg" style="height:340px"></div>
        <div class="cw-note">受车企和动力类型筛选联动；点击区域按钮联动国家明细</div>
      </div>
      <div class="cw">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
          <div class="cw-lbl" style="margin-bottom:0" id="s2drillLbl">国家明细</div>
          <div id="s2regFilter" style="display:flex;gap:4px;flex-wrap:wrap"></div>
        </div>
        <div class="legend-note" id="s2countryLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="c2countryContainer"><div id="c2country" style="height:340px"></div></div>
        <div class="cw-note">多车企时各自展示TOP5+其他；选"全部区域"查看全球排名</div>
      </div>
    </div>
    <!-- ▸ 分车型出口量 -->
    <div class="cw mt14">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap;gap:8px">
        <div class="cw-lbl" style="margin-bottom:0" id="s2modelLbl">分车型出口量</div>
        <div style="display:flex;align-items:center;gap:18px;flex-wrap:wrap">
          <div style="display:flex;align-items:center;gap:6px">
            <span style="font-size:13px;color:var(--t1);font-weight:600">区域：</span>
            <div id="s2modelRegionFilter" style="display:flex;gap:4px;flex-wrap:wrap"></div>
          </div>
          <div style="display:flex;align-items:center;gap:6px">
            <span style="font-size:13px;color:var(--t1);font-weight:600">动力：</span>
            <div id="s2modelEnergyFilter" style="display:flex;gap:4px;flex-wrap:wrap"></div>
          </div>
          <div style="display:flex;align-items:center;gap:6px">
            <span style="font-size:13px;color:var(--t1);font-weight:600">国家：</span>
            <select id="s2countrySelect" style="background:var(--input);color:var(--t1);border:1px solid var(--bd);border-radius:4px;padding:5px 12px;font-size:13px;font-family:inherit;cursor:pointer;min-width:140px">
              <option value="">全部国家</option>
            </select>
          </div>
        </div>
      </div>
      <div class="legend-note" id="s2modelLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
      <div id="c2modelContainer"></div>
      <div class="cw-note">区域→国家→动力类型三级联动筛选；单车企TOP15，多车企各TOP10；按2025年降序</div>
    </div>
  </div>
</div>

<div class="data-footnote" style="margin:14px 28px;padding:12px 20px;font-size:11px;line-height:1.8;color:var(--t2)">
<div style="font-weight:600;color:var(--t1);margin-bottom:4px">注：各车企包含品牌明细</div>
奇瑞汽车 H9973：奇瑞、星途、捷途、捷途山海、iCAR汽车、奇瑞新能源、奇瑞风云<br>
比亚迪 A002594/H1211：比亚迪、腾势、仰望、方程豹<br>
上汽集团 A600104：荣威、名爵、上汽大通MAXUS、智己汽车、五菱汽车、宝骏、大众、别克、凯迪拉克、雪佛兰<br>
长城汽车 A601633/H2333：哈弗、坦克、欧拉、魏牌、长城<br>
吉利汽车 H0175：吉利汽车、吉利几何、吉利银河、极氪、领克<br>
长安汽车 A000625：长安、长安凯程、长安启源、长安欧尚、深蓝汽车<br>
特斯拉 US.TSLA：特斯拉<br>
江淮汽车 A600418：江淮汽车、江淮瑞风、思皓、江淮钇为<br>
江铃汽车 A000550：江铃集团新能源、江铃晶马汽车、福特<br>
广汽集团 A601238/H2238：广汽传祺、埃安、广汽昊铂<br>
零跑汽车 H9863：零跑汽车<br>
北京汽车 H1958：北京汽车、北京越野、北京汽车制造厂、ROX极石、奔驰<br>
小鹏汽车 H9868/US.XPEV：小鹏<br>
赛力斯 A601127/H9927：鸿蒙智行、东风小康、东风风光、瑞驰汽车<br>
岚图汽车 H7489：岚图汽车<br>
海马汽车 A000572：海马<br>
北汽蓝谷 A600733：ARCFOX极狐<br>
吉利-其他：沃尔沃、Polestar极星、smart、LEVC、睿蓝汽车、远程<br>
长安汽车-其他：福特、林肯、阿维塔、马自达<br>
奇瑞-其他：凯翼、开瑞<br>
广汽集团-其他：丰田、本田<br>
北京汽车-其他：现代<br>
赛力斯-其他：蓝电<br>
宝马（华晨宝马合资）：宝马<br>
一汽集团：红旗、奔腾、一汽、大众、奥迪<br>
东风集团：东风、风神、东风奕派、东风风行、启辰、本田、标致、雪铁龙、起亚、达契亚、猛士
</div>

</div><!-- /main -->
</div><!-- /srcHaiguan -->

<div id="srcZhongqi" style="display:none">
  <div class="hd" style="position:sticky;top:38px;z-index:100">
    <div class="hd-title">中国<span class="acc">乘用车</span>海外数据看板 · 中汽协</div>
    <div style="text-align:right">
      <div class="hd-meta" id="caamMeta">数据加载中</div>
      <div class="hd-src">数据来源：中国汽车工业协会 | 联系人：罗英 15766781877</div>
    </div>
  </div>
  <div class="data-footnote" style="margin:0 28px 14px;padding:10px 20px">
    <div class="fn-item" style="margin:0"><span class="fn-label">口径：</span>中汽协统计的各家车企出海销量（含海外生产） <span class="fn-label">统计维度：</span>分车企、分车型 <span class="fn-label">更新频率：</span>月频</div>
  </div>
  <div class="main" id="caamMain">

  <!-- ═══ 中汽协 板块0 出口总览 ═══ -->
  <div class="sec" id="caamSec0">
    <div class="sec-hd">
      <div style="display:flex;align-items:center">
        <div class="sec-title"><span class="sec-bar"></span>出口总览</div>
        <div class="tab-bar" id="caamS0tabs">
          <button class="tab-btn on" data-mode="annual">年度</button>
          <button class="tab-btn" data-mode="monthly">月度</button>
        </div>
      </div>
      <div class="sec-sub"><span id="caamS0sub">各年度总量与同比增速（2020-2026）</span></div>
    </div>
    <div class="sec-body">
      <div id="caamS0annual">
        <div class="g2">
          <div class="cw">
            <div class="cw-lbl">乘用车出口总量</div>
            <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
            <div id="caamC0tot" style="height:280px"></div>
          </div>
          <div class="cw">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
              <div class="cw-lbl" style="margin-bottom:0">分系别出口量</div>
              <div class="fbar" id="caamASerFilter" style="margin:0;padding:6px 10px;background:transparent;border:none"></div>
            </div>
            <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
            <div id="caamC0ser" style="height:280px"></div>
            <div class="cw-note">按品牌系别分类；可多选筛选</div>
          </div>
        </div>
        <div class="cw mt14">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
            <div class="cw-lbl" style="margin-bottom:0">重点车企出口量（按26年累计降序）</div>
            <div id="caamS0oemExtra" style="display:flex;gap:4px;flex-wrap:wrap;align-items:center"></div>
          </div>
          <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="caamC0oem" style="height:420px"></div>
          <div class="cw-note">默认展示前5上市车企；查看更多展示前15+其他</div>
        </div>
      </div>
      <div id="caamS0monthly" style="display:none"></div>
    </div>
  </div>

  <!-- ═══ 中汽协 板块二 分车企视角 ═══ -->
  <div class="sec" id="caamSec2">
    <div class="sec-hd">
      <div class="sec-title"><span class="sec-bar"></span>分车企-分车型视角</div>
      <div class="sec-sub">按车企查看车型出口结构</div>
    </div>
    <div class="sec-body">
      <div class="fbar-v" id="caamS2filter">
        <div class="frow">
          <span class="flbl">时间维度</span>
          <div class="tab-bar" id="caamS2timeTabs">
            <button class="tab-btn on" data-mode="annual">年度</button>
            <button class="tab-btn" data-mode="monthly">月度</button>
          </div>
        </div>
        <div class="frow" id="caamS2oemRow">
          <span class="flbl">选择车企</span>
          <div id="caamS2oemBtns" style="display:flex;gap:4px;flex-wrap:wrap"></div>
          <select id="caamS2oemSelect" style="background:var(--input);color:var(--t1);border:1px solid var(--bd);border-radius:4px;padding:4px 8px;font-size:12px;font-family:inherit;cursor:pointer">
            <option value="">更多车企 ▾</option>
          </select>
        </div>
        <div class="frow" id="caamS2brandBox" style="display:none">
          <span class="flbl">品牌筛选</span>
          <div id="caamS2brandBtns" style="display:flex;gap:4px;flex-wrap:wrap"></div>
        </div>
      </div>
      <div class="cw mt14">
        <div class="cw-lbl" id="caamS2modelLbl">分车型出口量</div>
        <div class="legend-note" id="caamS2modelLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="caamC2modelContainer"></div>
        <div class="cw-note">默认前5车型；查看更多展示前15+其他；按26年累计降序</div>
      </div>
    </div>
  </div>

  <div class="data-footnote" style="margin:14px 28px;padding:12px 20px;font-size:11px;line-height:1.8;color:var(--t2)">
  <div style="font-weight:600;color:var(--t1);margin-bottom:4px">注：各车企包含品牌明细</div>
  奇瑞汽车 H9973：奇瑞、捷途、星途汽车、iCAR<br>
  比亚迪 A002594/H1211：比亚迪汽车、腾势、仰望、方程豹<br>
  上汽集团 A600104：MG、荣威、大通、智己汽車、五菱、宝骏、大众、别克、凯迪拉克、雪佛兰<br>
  长城汽车 A601633/H2333：哈弗、坦克、欧拉、WEY、长城汽车<br>
  吉利汽车 H0175：吉利、几何汽车、吉利银河、极氪、领克汽车<br>
  长安汽车 A000625：长安汽车、长安启源、深蓝汽车<br>
  特斯拉 US.TSLA：特斯拉<br>
  江淮汽车 A600418：安徽江淮汽车、思皓、江淮瑞风、钇为<br>
  江铃汽车 A000550：江铃汽车、江铃新能源、福特<br>
  广汽集团 A601238/H2238：传祺、广汽埃安<br>
  零跑汽车 H9863：零跑汽车<br>
  北京汽车 H1958：北京、北京汽车制造厂、极石汽车<br>
  小鹏汽车 H9868/US.XPEV：小鹏汽车<br>
  赛力斯 A601127/H9927：赛力斯、东风小康汽车、重庆瑞驰<br>
  岚图汽车 H7489：岚图汽车<br>
  海马汽车 A000572：海马汽车<br>
  北汽蓝谷 A600733：ARCFOX<br>
  吉利-其他：沃尔沃汽车、极星汽车、smart、LEVC、睿蓝汽车<br>
  长安汽车-其他：福特、林肯、阿维塔科技、马自达<br>
  奇瑞-其他：凯翼、开瑞<br>
  广汽集团-其他：丰田、本田<br>
  北京汽车-其他：现代<br>
  赛力斯-其他：蓝电<br>
  宝马（华晨宝马合资）：宝马<br>
  一汽集团：红旗、奔腾、一汽吉林汽车、大众<br>
  东风集团：东风汽车、风神、东风纳米、Forthing、eπ、启辰、本田、标致、雪铁龙、起亚、达契亚、猛士汽车科技公司
  </div>

  </div><!-- /caamMain -->
</div>

<div id="srcMkls" style="display:none">
  <div class="hd" style="position:sticky;top:38px;z-index:100">
    <div class="hd-title">中国<span class="acc">乘用车</span>海外数据看板 · Marklines</div>
    <div style="text-align:right">
      <div class="hd-meta" id="mklsMeta">数据加载中</div>
      <div class="hd-src">数据来源：Marklines | 联系人：罗英 15766781877</div>
    </div>
  </div>
  <div class="data-footnote" style="margin:0 28px 14px;padding:10px 20px">
    <div class="fn-item" style="margin:0"><span class="fn-label">口径：</span>Marklines统计的可跟踪国家各家车企、车型终端实销 <span style="color:#E85A5A;font-weight:600">（部分国家缺失，导致车企实销数据偏小，仅可作趋势研究用）</span> <span class="fn-label">统计维度：</span>分区域、分国家、分动力类型、分车企、分车型 <span class="fn-label">更新频率：</span>月频（部分国家原始数据更新频率较低，这些国家或2-6月更新一次） <span style="color:var(--pos);font-weight:600">注意：<span id="mklsPartialRange">2026年4-5月</span>部分国家数据仍有缺失，销量及增速仅供参考，详情请联系罗英</span></div>
  </div>
  <div class="main" id="mklsMain">

  <!-- ═══ Marklines 终端实销总览 ═══ -->
  <div class="sec" id="mklsSec0">
    <div class="sec-hd">
      <div style="display:flex;align-items:center">
        <div class="sec-title"><span class="sec-bar"></span>终端实销总览</div>
        <div class="tab-bar" id="mklsS0tabs">
          <button class="tab-btn on" data-mode="annual">年度</button>
          <button class="tab-btn" data-mode="monthly">月度</button>
        </div>
      </div>
      <div class="sec-sub"><span id="mklsS0sub">海外轻型车终端实销（不含中国）</span></div>
    </div>
    <div class="sec-body">
      <div id="mklsS0annual">
        <div class="g2">
          <div class="cw">
            <div class="cw-lbl">海外轻型车终端实销总量</div>
            <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
            <div id="mklsC0tot" style="height:280px"></div>
          </div>
          <div class="cw">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap">
              <div class="cw-lbl" style="margin-bottom:0">分系别实销量</div>
              <div style="display:flex;gap:6px;align-items:center">
                <div id="mklsA0serFilter" style="display:flex;gap:3px;flex-wrap:wrap"></div>
                <button class="show-more-btn" id="mklsC0serZoomBtn" style="margin:0;padding:3px 12px;font-size:11px">放大图</button>
              </div>
            </div>
            <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
            <div id="mklsC0ser" style="height:280px"></div>
            <div class="cw-note">按品牌系别分类，除中、日、韩、欧、美系以外列为其他</div>
          </div>
        </div>
        <div class="g2 mt14">
          <div class="cw">
            <div style="margin-bottom:4px">
              <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
                <div class="cw-lbl" style="margin-bottom:0" id="mklsA0engLbl">分动力类型实销量</div>
                <button class="show-more-btn" id="mklsC0engZoomBtn" style="margin:0;padding:3px 12px;font-size:11px">放大图</button>
              </div>
              <div style="display:flex;gap:6px;align-items:flex-start;flex-wrap:wrap">
                <div id="mklsA0engContFilter" style="display:flex;gap:3px;flex-wrap:wrap"></div>
                <select id="mklsA0engCountrySel" style="font-size:11px;padding:3px 8px;background:#1a3060;color:#E8EEF8;border:1px solid #2a5090;border-radius:4px;font-family:inherit;cursor:pointer;min-width:90px;max-width:140px;height:22px;line-height:22px">
                  <option value="">全部国家</option>
                </select>
              </div>
            </div>
            <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
            <div id="mklsC0eng" style="height:300px"></div>
            <div class="cw-note">新能源 = 纯电动 + 插电混动；可按大洲/国家筛选</div>
          </div>
          <div class="cw">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap">
              <div class="cw-lbl" style="margin-bottom:0">分大洲实销量</div>
              <div style="display:flex;gap:6px;align-items:center">
                <div id="mklsA0contFilter" style="display:flex;gap:3px;flex-wrap:wrap"></div>
                <button class="show-more-btn" id="mklsC0contZoomBtn" style="margin:0;padding:3px 12px;font-size:11px">放大图</button>
              </div>
            </div>
            <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
            <div id="mklsC0cont" style="height:300px"></div>
            <div class="cw-note">欧洲拆分为欧洲（非俄）和俄罗斯；可复选大洲</div>
          </div>
        </div>
        <div class="cw mt14">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
            <div class="cw-lbl" style="margin-bottom:0">国家明细</div>
            <div id="mklsS0regFilter" style="display:flex;gap:4px;flex-wrap:wrap"></div>
          </div>
          <div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="mklsC0country" style="height:340px"></div>
          <div class="cw-note">点击大洲按钮切换国家明细；默认前5国家；查看更多展示前15+其他</div>
        </div>
      </div>
      <div id="mklsS0monthly" style="display:none"></div>
    </div>
  </div>

  <!-- ═══ Marklines 分系别·分动力类型视角 ═══ -->
  <div class="sec" id="mklsSec1">
    <div class="sec-hd">
      <div class="sec-title"><span class="sec-bar"></span>分系别·分动力类型视角</div>
      <div class="sec-sub">按系别、动力类型、年度/月度多维筛选</div>
    </div>
    <div class="sec-body">
      <div class="fbar-v" id="mklsS1filter">
        <div class="frow">
          <span class="flbl">时间维度</span>
          <div class="tab-bar" id="mklsS1timeTabs">
            <button class="tab-btn on" data-mode="annual">年度</button>
            <button class="tab-btn" data-mode="monthly">月度</button>
          </div>
        </div>
        <div class="frow" id="mklsS1yearBox"></div>
        <div class="frow" id="mklsS1seriesBox"></div>
        <div class="frow" id="mklsS1energyBox"></div>
      </div>
      <div class="g2">
        <div class="cw" id="mklsS1regWrap">
          <div style="margin-bottom:4px">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
              <div class="cw-lbl" style="margin-bottom:0" id="mklsS1regLbl">分区域销量</div>
              <button class="show-more-btn" id="mklsS1regZoomBtn" style="margin:0;padding:3px 12px;font-size:11px">放大图</button>
            </div>
            <div id="mklsS1regSelFilter" style="display:none;gap:3px;flex-wrap:wrap"></div>
          </div>
          <div class="legend-note" id="mklsS1regLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="mklsC1reg" style="height:340px"></div>
          <div class="cw-note">联动系别+动力类型筛选</div>
        </div>
        <div class="cw" id="mklsS1countryWrap">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
            <div class="cw-lbl" style="margin-bottom:0" id="mklsS1drillLbl">国家明细</div>
            <div style="display:flex;gap:4px;align-items:center">
              <div id="mklsS1regFilter" style="display:flex;gap:3px;flex-wrap:wrap"></div>
              <button class="show-more-btn" id="mklsS1countryZoomBtn" style="margin:0;padding:3px 12px;font-size:11px">放大图</button>
            </div>
          </div>
          <!-- 月度模式：国家多选器（按销量排序，默认首选 TOP1） -->
          <div id="mklsS1countryPickerBox" style="display:none;align-items:flex-start;flex-wrap:wrap;gap:3px;margin-bottom:6px;padding:6px 8px;background:#0A1628;border:1px solid #1e3a6e;border-radius:4px;max-height:86px;overflow-y:auto">
            <span style="font-size:11px;color:#8A9DC0;padding:2px 4px 0 0;white-space:nowrap">国家（可多选）：</span>
            <div id="mklsS1countryPicker" style="display:flex;flex-wrap:wrap;gap:3px;flex:1"></div>
          </div>
          <div class="legend-note" id="mklsS1countryLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="mklsC1country" style="height:340px"></div>
          <div class="cw-note">联动系别+动力类型筛选；点击大洲按钮切换国家明细；年度默认前5国家，月度可多选（默认TOP1）</div>
        </div>
      </div>
      <div class="cw mt14" id="mklsS1oemBox">
        <div style="margin-bottom:4px">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
            <div class="cw-lbl" style="margin-bottom:0" id="mklsS1oemLbl">重点车企终端实销量</div>
            <div style="display:flex;gap:6px;align-items:center">
              <button class="show-more-btn" id="mklsS1oemMoreBtn" style="margin:0;padding:3px 12px;font-size:11px">更多车企</button>
              <button class="show-more-btn" id="mklsS1oemZoomBtn" style="margin:0;padding:3px 12px;font-size:11px">放大图</button>
            </div>
          </div>
          <div style="display:flex;gap:6px;align-items:flex-start;flex-wrap:wrap">
            <div id="mklsS1oemContFilter" style="display:flex;gap:3px;flex-wrap:wrap"></div>
            <select id="mklsS1oemCountrySel" style="font-size:11px;padding:3px 8px;background:#1a3060;color:#E8EEF8;border:1px solid #2a5090;border-radius:4px;font-family:inherit;cursor:pointer;min-width:90px;max-width:140px;height:22px;line-height:22px">
              <option value="">全部国家</option>
            </select>
          </div>
        </div>
        <div class="legend-note" id="mklsS1oemLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="mklsC1oem" style="height:420px"></div>
        <div class="cw-note">联动系别+动力类型+区域/国家筛选；默认展示前10车企；更多车企可筛选追加</div>
      </div>
    </div>
  </div>

  <!-- ═══ Marklines 分车企视角 ═══ -->
  <div class="sec" id="mklsSec2">
    <div class="sec-hd">
      <div class="sec-title"><span class="sec-bar"></span>分车企视角</div>
      <div class="sec-sub">按车企查看动力结构、区域分布、国家及车型明细</div>
    </div>
    <div class="sec-body">
      <div class="fbar-v" id="mklsS2filter">
        <div class="frow">
          <span class="flbl">时间维度</span>
          <div class="tab-bar" id="mklsS2timeTabs">
            <button class="tab-btn on" data-mode="annual">年度</button>
            <button class="tab-btn" data-mode="monthly">月度</button>
          </div>
        </div>
        <div class="frow" id="mklsS2yearBox">
          <span class="flbl">年份</span>
          <label class="ychk on"><input type="radio" name="mklsS2yr" value="all" checked><span class="ydot" style="background:#C9A84C"></span>全部</label>
          <label class="ychk"><input type="radio" name="mklsS2yr" value="2023"><span class="ydot" style="background:#4A90E2"></span>2023</label>
          <label class="ychk"><input type="radio" name="mklsS2yr" value="2024"><span class="ydot" style="background:#C9A84C"></span>2024</label>
          <label class="ychk"><input type="radio" name="mklsS2yr" value="2025"><span class="ydot" style="background:#5BC4A0"></span>2025</label>
          <label class="ychk"><input type="radio" name="mklsS2yr" value="2026"><span class="ydot" style="background:#E85A5A"></span>2026</label>
        </div>
        <div class="frow" id="mklsS2oemRow">
          <span class="flbl">选择车企</span>
          <div class="cw-note" id="mklsS2oemHint" style="margin:0;font-style:normal;color:var(--t2)">选"全部"年份时为单选；选具体年份后支持车企复选对比</div>
        </div>
        <div class="frow">
          <div id="mklsS2oemBtns" style="display:flex;gap:4px;flex-wrap:wrap"></div>
          <select id="mklsS2oemSelect" style="background:var(--input);color:var(--t1);border:1px solid var(--bd);border-radius:4px;padding:4px 8px;font-size:12px;font-family:inherit;cursor:pointer">
            <option value="">更多车企 ▾</option>
          </select>
        </div>
        <div id="mklsS2oemExtraTags" style="display:flex;gap:4px;flex-wrap:wrap"></div>
        <div class="frow" id="mklsS2brandBox" style="display:none">
          <span class="flbl">品牌筛选</span>
          <div id="mklsS2brandBtns" style="display:flex;gap:4px;flex-wrap:wrap"></div>
        </div>
      </div>
      <div class="g2">
        <div class="cw">
          <div class="cw-lbl" id="mklsS2engLbl">分动力类型</div>
          <div class="legend-note" id="mklsS2engLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="mklsC2eng" style="height:300px"></div>
        </div>
        <div class="cw">
          <div class="cw-lbl" id="mklsS2regLbl">分大洲</div>
          <div class="legend-note" id="mklsS2regLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
          <div id="mklsC2reg" style="height:300px"></div>
        </div>
      </div>
      <div class="cw mt14" id="mklsC2countryContainer">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
          <div class="cw-lbl" style="margin-bottom:0" id="mklsS2drillLbl">国家明细</div>
          <div style="display:flex;gap:4px;align-items:center">
            <div id="mklsS2regFilter" style="display:flex;gap:3px;flex-wrap:wrap"></div>
          </div>
        </div>
        <div class="legend-note" id="mklsS2countryLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="mklsC2countryInner" style="height:340px"></div>
        <div class="cw-note">多车企时各自展示TOP5+其他；选"全部区域"查看全球排名</div>
      </div>
      <div class="cw mt14" id="mklsC2modelContainer">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap;gap:8px">
          <div class="cw-lbl" style="margin-bottom:0" id="mklsS2modelLbl">分车型实销量</div>
          <div style="display:flex;align-items:center;gap:18px;flex-wrap:wrap">
            <div style="display:flex;align-items:center;gap:6px">
              <span style="font-size:13px;color:var(--t1);font-weight:600">区域：</span>
              <div id="mklsS2modelRegFilter" style="display:flex;gap:4px;flex-wrap:wrap"></div>
            </div>
            <div style="display:flex;align-items:center;gap:6px">
              <span style="font-size:13px;color:var(--t1);font-weight:600">动力：</span>
              <div id="mklsS2modelEngFilter" style="display:flex;gap:4px;flex-wrap:wrap"></div>
            </div>
            <div style="display:flex;align-items:center;gap:6px">
              <span style="font-size:13px;color:var(--t1);font-weight:600">国家：</span>
              <select id="mklsS2modelCountrySel" style="background:var(--input);color:var(--t1);border:1px solid var(--bd);border-radius:4px;padding:5px 12px;font-size:13px;font-family:inherit;cursor:pointer;min-width:140px">
                <option value="">全部国家</option>
              </select>
            </div>
          </div>
        </div>
        <div class="legend-note" id="mklsS2modelLegend"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>
        <div id="mklsC2modelInner"></div>
        <div class="cw-note">区域→国家→动力类型三级联动筛选；单车企TOP15，多车企各TOP10；按2025年降序<br>注：不同区域的车型名称差异较大，未标准化处理</div>
      </div>
    </div>
  </div>

  </div><!-- /mklsMain -->
</div>

<script>
const DATA = __DATA__;
const CAAM = __CAAM_DATA__;
const MKLS = __MKLS_DATA__;

/* ── Constants ── */
const YC={'2023':'#4A90E2','2024':'#C9A84C','2025':'#5BC4A0'};
const CS=['#4A90E2','#C9A84C','#5BC4A0','#E87B5A','#9B7FD4','#5BC4D4','#E8C870'];
const TT={backgroundColor:'#0D1E3D',borderColor:'#1e3a6e',
  textStyle:{color:'#FFFFFF',fontSize:13},
  extraCssText:'box-shadow:0 4px 20px rgba(0,0,0,.5);border-radius:8px;padding:12px 14px'};
const MONTHS=['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];

function fmt(n){
  if(!n) return '0';
  if(n>=10000) return (n/10000).toFixed(1)+'万';
  if(n>=1000) return (n/1000).toFixed(1)+'K';
  return n.toLocaleString();
}

const _c={};
function gi(id){
  if(!_c[id]){
    const dom=document.getElementById(id);
    if(!dom) return null;
    _c[id]=echarts.init(dom);
  }
  return _c[id];
}
window.addEventListener('resize',()=>Object.values(_c).forEach(c=>c.resize()));

/* ── Excel Export ── */
function colLetter(i){
  let s=''; i++;
  while(i>0){s=String.fromCharCode(((i-1)%26)+65)+s;i=Math.floor((i-1)/26);}
  return s;
}

async function exportChart(chartId, title){
  const chart = _c[chartId];
  if(!chart) return;
  const opt = chart.getOption();
  if(!opt || !opt.series) return;

  const cats = (opt.xAxis && opt.xAxis[0] && opt.xAxis[0].data) || [];
  const sd = opt.series.filter(s=>s.type==='bar'||s.type==='line');
  // 表头：分类 + 各 series 数值列 + 各 series 同比列（数值列在前保持图表 series 引用稳定）
  const valHeaders = sd.map(s=>s.name||'');
  const yoyHeaders = sd.map(s=>(s.name||'')+' 同比%');
  const header = ['分类', ...valHeaders, ...yoyHeaders];
  const rows = [header];
  cats.forEach((cat,ci)=>{
    const valRow=[]; const yoyRow=[];
    sd.forEach(s=>{
      const d=s.data[ci];
      let val=0, yoy=null;
      if(typeof d==='object' && d!==null){ val = d.value||0; yoy = (d.yoy!==undefined && d.yoy!==null) ? d.yoy : null; }
      else { val = d||0; }
      valRow.push(val);
      yoyRow.push(yoy);
    });
    rows.push([cat, ...valRow, ...yoyRow]);
  });

  // Create workbook with SheetJS
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = header.map((_,i)=>({wch:i===0?18:14}));
  // 同比列加百分号格式
  const nS = sd.length;
  for(let si=0; si<nS; si++){
    const col = colLetter(1 + nS + si);  // 同比列字母（A=分类，B..=数值，再后面=同比）
    for(let ri=2; ri<=cats.length+1; ri++){
      const cellRef = col + ri;
      const cell = ws[cellRef];
      if(cell && typeof cell.v === 'number') cell.z = '0.0"%"';
    }
  }
  XLSX.utils.book_append_sheet(wb, ws, '数据');
  const wbBuf = XLSX.write(wb, {bookType:'xlsx',type:'array'});

  // Inject chart via JSZip
  const zip = await JSZip.loadAsync(wbBuf);
  const nR = cats.length;
  const sheet = '数据';

  // Colors from ECharts series
  const getColor = (s,i) => {
    const c = s.itemStyle && s.itemStyle.color;
    if(c && typeof c==='string' && c.startsWith('#')) return c.substring(1).toUpperCase();
    return ['4A90E2','C9A84C','5BC4A0','E87B5A','9B7FD4','5BC4D4','E8C870','D4697F','6BAEE8','8FBC5A','B87FD4','F0A050'][i%12];
  };

  // Build series XML for chart
  let serXml = '';
  sd.forEach((s,si)=>{
    const col = colLetter(si+1);
    const clr = getColor(s,si);
    serXml += `<c:ser>
      <c:idx val="${si}"/><c:order val="${si}"/>
      <c:tx><c:strRef><c:f>${sheet}!$${col}$1</c:f></c:strRef></c:tx>
      <c:spPr><a:solidFill><a:srgbClr val="${clr}"/></a:solidFill></c:spPr>
      <c:cat><c:strRef><c:f>${sheet}!$A$2:$A$${nR+1}</c:f></c:strRef></c:cat>
      <c:val><c:numRef><c:f>${sheet}!$${col}$2:$${col}$${nR+1}</c:f></c:numRef></c:val>
    </c:ser>`;
  });

  const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<c:chart>
  <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="1200"/><a:t>${title||chartId}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
  <c:autoTitleDeleted val="0"/>
  <c:plotArea>
    <c:layout/>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:varyColors val="0"/>
      ${serXml}
      <c:axId val="111"/><c:axId val="222"/>
    </c:barChart>
    <c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:catAx>
    <c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111"/>
      <c:numFmt formatCode="General" sourceLinked="1"/></c:valAx>
  </c:plotArea>
  <c:legend><c:legendPos val="b"/></c:legend>
  <c:plotVisOnly val="1"/>
</c:chart>
</c:chartSpace>`;

  zip.file('xl/charts/chart1.xml', chartXml);

  // Drawing XML (embed chart in sheet)
  const drawXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:twoCellAnchor>
  <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${nR+3}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>${Math.max(nS+1,8)}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${nR+23}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro="">
    <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
    <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
      <c:chart r:id="rId1"/>
    </a:graphicData></a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:twoCellAnchor>
</xdr:wsDr>`;

  zip.file('xl/drawings/drawing1.xml', drawXml);

  // Drawing rels
  zip.file('xl/drawings/_rels/drawing1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`);

  // Update sheet1 rels to include drawing
  const s1RelsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  let s1Rels = '';
  if(zip.file(s1RelsPath)){
    s1Rels = await zip.file(s1RelsPath).async('string');
    s1Rels = s1Rels.replace('</Relationships>',
      `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`);
  } else {
    s1Rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
  }
  zip.file(s1RelsPath, s1Rels);

  // Add drawing ref to sheet XML
  const sheetPath = 'xl/worksheets/sheet1.xml';
  let sheetXml = await zip.file(sheetPath).async('string');
  if(!sheetXml.includes('drawing')){
    sheetXml = sheetXml.replace('</worksheet>',
      '<drawing r:id="rId10"/></worksheet>');
  }
  zip.file(sheetPath, sheetXml);

  // Update [Content_Types].xml
  let ct = await zip.file('[Content_Types].xml').async('string');
  if(!ct.includes('chart+xml')){
    ct = ct.replace('</Types>',
      `<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>`);
  }
  zip.file('[Content_Types].xml', ct);

  // Generate and download
  const blob = await zip.generateAsync({type:'blob'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (title||chartId) + '.xlsx';
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ── Add export buttons to all .cw elements ── */
function addExportButtons(){
  document.querySelectorAll('.cw').forEach(cw=>{
    if(cw.querySelector('.exp-btn')) return;
    // Find chart div (div with id that has an echarts instance)
    const chartDivs = cw.querySelectorAll('div[id]');
    let chartId = null;
    for(const d of chartDivs){
      if(_c[d.id]){ chartId = d.id; break; }
    }
    if(!chartId){
      // Fallback: find any div with style height that looks like a chart container
      for(const d of chartDivs){
        if(d.style.height){ chartId = d.id; break; }
      }
    }
    if(!chartId) return;
    const btn = document.createElement('button');
    btn.className = 'exp-btn';
    btn.textContent = '导出';
    btn.title = '导出Excel数据+图表';
    btn.addEventListener('click', async(e)=>{
      e.stopPropagation();
      // Get current title at click time (may have changed dynamically)
      const lbl = cw.querySelector('.cw-lbl');
      const title = lbl ? lbl.textContent.trim() : chartId;
      btn.textContent = '导出中...';
      btn.disabled = true;
      try{ await exportChart(chartId, title); }catch(ex){ console.error(ex); }
      btn.textContent = '导出';
      btn.disabled = false;
    });
    cw.appendChild(btn);
  });
}
/* ── 弹窗图表导出按钮 ── */
function addModalExport(modalBox, chartId, title){
  const btn = document.createElement('button');
  btn.className = 'exp-btn';
  btn.style.cssText = 'position:absolute;bottom:10px;right:14px';
  btn.textContent = '导出';
  btn.title = '导出Excel数据+图表';
  btn.addEventListener('click', async(e)=>{
    e.stopPropagation();
    btn.textContent = '导出中...'; btn.disabled = true;
    try{
      const chart = echarts.getInstanceByDom(document.getElementById(chartId));
      if(chart){ _c[chartId] = chart; await exportChart(chartId, title); }
    }catch(ex){ console.error(ex); }
    btn.textContent = '导出'; btn.disabled = false;
  });
  modalBox.style.position = 'relative';
  modalBox.appendChild(btn);
}

// Delay to ensure all charts are rendered, re-run on mutations
setTimeout(addExportButtons, 2000);
const _expObs = new MutationObserver(()=>setTimeout(addExportButtons, 500));
_expObs.observe(document.querySelector('.main'), {childList:true, subtree:true});

/* ─────────────────────────────────────────
   buildBar: grouped bar (年度模式 or 月度模式)
   categories: string[]
   dataMap: { cat: {yr:N, yr_yoy:X} } (年度) or { cat: {'1':N,...} } (月度)
   keys: ['2023','2024','2025'] or ['1','2',...,'12']
   keyLabels: optional display labels
   keyColors: optional { key: color }
   showYoy: boolean
───────────────────────────────────────── */
function buildBar(id, categories, dataMap, keys, opts){
  const chart = gi(id);
  if(!chart) return;
  opts = opts || {};
  const keyLabels = opts.keyLabels || {};
  const keyColors = opts.keyColors || YC;
  const showYoy = opts.showYoy !== false;
  const hideLegend = opts.hideLegend || false;
  const showMonthLabel = opts.showMonthLabel || false;
  const isMonthly = opts.isMonthly || false;
  const n = categories.length * keys.length;
  const vFs = n > 24 ? 8 : n > 16 ? 9 : n > 10 ? 10 : 11;
  const yFs = Math.max(7, vFs - 1);
  const maxCatLen = Math.max(...categories.map(c=>c.length));
  const rotate = categories.length > 10 ? -35 : categories.length > 5 && maxCatLen > 6 ? -25 : 0;

  const defaultColors = ['#4A90E2','#C9A84C','#5BC4A0','#E87B5A','#9B7FD4','#5BC4D4',
    '#E8C870','#5BC4A0','#4A90E2','#C9A84C','#E87B5A','#9B7FD4'];

  const mLblFs = n > 60 ? 0 : n > 36 ? 7 : 8;

  const mLblInterval = showMonthLabel ? 5 : 0; // 每隔5个显示（即1月和7月）
  const series = keys.map((k, ki) => ({
    name: keyLabels[k] || (/^\d{4}$/.test(k) ? k+'年' : k),
    type:'bar', barGap: isMonthly ? '8%' : '4%',
    barCategoryGap: categories.length > 7 ? '25%' : isMonthly ? '25%' : '30%',
    itemStyle:{ color: keyColors[k] || defaultColors[ki % defaultColors.length], borderRadius:[3,3,0,0] },
    data: categories.map(cat=>{
      const d = dataMap[cat]||{};
      return { value: d[k]||0, yoy: showYoy ? (d[k+'_yoy']??null) : null, monthKey: k };
    }),
    label:{
      // 强制 show=true，让 formatter 自行决定每个柱子是否输出文字（保证 yoy 总能渲染）
      // position 始终 top（即使 showMonthLabel 模式下，月份标签也放顶部，便于和 yoy 一起堆叠）
      show: true,
      position: 'top',
      formatter(p){
        const v = fmt(p.value);
        const yoy = p.data && p.data.yoy;
        const monthKey = p.data && p.data.monthKey;
        // 月度模式
        if(isMonthly){
          // 抽月份序号（'YYYY-M' → M）；旧整数键也兼容
          let mi = NaN;
          if(monthKey != null){
            const s = String(monthKey);
            mi = s.includes('-') ? parseInt(s.split('-')[1]) : parseInt(s);
          }
          const parts = [];
          // showMonthLabel 模式：在 1月/7月 位置标月份
          if(showMonthLabel && (mi===1 || mi===7)){
            const mLbl = keyLabels[monthKey] || monthKey;
            parts.push(`{m|${mLbl}}`);
          }
          // 有 yoy 一定显示（含值和同比）
          if(yoy !== null && yoy !== undefined){
            const pos = yoy >= 0;
            if(!showMonthLabel) parts.push(`{v|${v}}`);
            parts.push(`{${pos?'r':'g'}|${pos?'▲':'▼'}${Math.abs(yoy).toFixed(1)}%}`);
          } else if(!showMonthLabel && mLblFs > 0 && p.value){
            // 无 yoy + 非 showMonthLabel + 有空间：显示数值
            parts.push(`{v|${v}}`);
          }
          return parts.join('\n');
        }
        // 年度模式
        if(yoy !== null && yoy !== undefined){
          const pos = yoy >= 0;
          return `{v|${v}}\n{${pos?'r':'g'}|${pos?'▲':'▼'}${Math.abs(yoy).toFixed(1)}%}`;
        }
        return `{v|${v}}`;
      },
      rich:{
        v:{fontSize: isMonthly ? Math.max(mLblFs, 8) : vFs, color:'#FFFFFF', lineHeight: isMonthly ? 12 : 18},
        m:{fontSize:9, color:'#C9D6EC', lineHeight:12, padding:[2,0,0,0]},
        r:{fontSize:yFs,color:'#E85A5A',lineHeight:13},
        g:{fontSize:yFs,color:'#4CAF82',lineHeight:13}
      }
    }
  }));

  chart.setOption({
    backgroundColor:'transparent',
    textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#8A9DC0'},
    tooltip:{
      ...TT, trigger:'axis', axisPointer:{type:'shadow'},
      formatter(params){
        let s=`<div style="font-weight:600;margin-bottom:6px;color:#C9A84C">${params[0].axisValue}</div>`;
        params.forEach(p=>{
          const yoy=p.data&&p.data.yoy;
          const yoyS=yoy!=null
            ? `<span style="color:${yoy>=0?'#E85A5A':'#4CAF82'};margin-left:6px">${yoy>=0?'▲':'▼'}${Math.abs(yoy).toFixed(1)}%</span>` : '';
          const c=keyColors[keys[p.seriesIndex]]||defaultColors[p.seriesIndex%defaultColors.length];
          s+=`<div style="display:flex;align-items:center;gap:6px;margin:3px 0">
            <span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${c}"></span>
            ${p.seriesName}：<b>${p.value.toLocaleString()}</b>${yoyS}</div>`;
        });
        return s;
      }
    },
    legend: (hideLegend || showMonthLabel || keys.length > 10) ? {show:false} : {
      data:keys.map(k=>keyLabels[k]||(/^\d{4}$/.test(k)?k+'年':k)), top:6,
      textStyle:{color:'#E8EEF8',fontSize:11}, itemWidth:12, itemHeight:8,
      itemGap:14,
      type: keys.length > 5 ? 'scroll' : 'plain'
    },
    grid:{top: (hideLegend || keys.length > 10) ? 30 : (keys.length > 5 ? 58 : 52), right:12,
      bottom: Math.abs(rotate) > 25 ? 80 : (Math.abs(rotate) > 0 ? 65 : 50),
      left:12, containLabel:true},
    xAxis:{type:'category', data:categories,
      axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
      axisLabel:{color:'#E8EEF8',fontSize: categories.length > 12 ? 10 : 12,interval:0,rotate},splitLine:{show:false}},
    yAxis:{type:'value',
      axisLine:{show:false},axisTick:{show:false},
      axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
      splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
    series
  }, true);
}

/* ── buildEnergyLayered: 燃油车 + 新能源(大类) + 新能源细分 ── */
function buildEnergyLayered(id, bigData, subData, keys, opts){
  opts = opts || {};
  const categories = ['燃油车','新能源','其中：纯电动','其中：插电混动','其中：增程式'];
  const merged = {};
  merged['燃油车'] = bigData['燃油车'] || {};
  merged['新能源'] = bigData['新能源'] || {};
  merged['其中：纯电动'] = subData['纯电动'] || {};
  merged['其中：插电混动'] = subData['插电混动'] || {};
  merged['其中：增程式'] = subData['增程式'] || {};
  buildBar(id, categories, merged, keys, opts);
}

/* ─────────────────────────────────────────
   buildMonthBar: 月度模式柱状图
   X轴 = 01-25 ~ 12-25, 系列 = 分类维度
   monthlyMap: { catName: {'1':N,'2':N,...} }
   catOrder: 排序后的分类列表
───────────────────────────────────────── */
function buildMonthBar(id, monthlyMap, catOrder){
  const chart = gi(id);
  if(!chart) return;
  // x 轴：'YYYY-M' → 'YY-MM'
  const xLabels = MK.map(k=>{ const [y,m] = k.split('-'); return y.substring(2)+'-'+m.padStart(2,'0'); });
  const cats = catOrder || Object.keys(monthlyMap);
  const singleCat = cats.length <= 1;
  const totalBars = cats.length * MK.length;
  const vFs = totalBars > 60 ? 0 : totalBars > 36 ? 8 : totalBars > 20 ? 9 : 10;

  // 月度同比 label：仅在 yoy 数据存在的月份显示（25年月度无可比基数，26M1/M2 vs 25M1/M2 有数据）
  // 数值 label 在多 cat 拥挤时（vFs=0）隐藏，但 yoy label 始终显示
  const yoyRich = {
    v:{fontSize: singleCat ? 10 : Math.max(vFs, 8), color:'#FFFFFF', lineHeight: singleCat ? 14 : 12},
    r:{fontSize: singleCat ? 9 : 8, color:'#E85A5A', lineHeight:12},
    g:{fontSize: singleCat ? 9 : 8, color:'#4CAF82', lineHeight:12}
  };
  function fmtVY(p){
    const yoy = p.data && p.data.yoy;
    const hasYoy = yoy !== null && yoy !== undefined;
    if(!p.value && !hasYoy) return '';
    const v = fmt(p.value);
    const parts = [];
    // 数值：单 cat 总是显示；多 cat 仅在有空间时显示
    if(singleCat || vFs > 0) parts.push(`{v|${v}}`);
    // yoy：永远显示（如果有）
    if(hasYoy){
      const pos = yoy >= 0;
      parts.push(`{${pos?'r':'g'}|${pos?'▲':'▼'}${Math.abs(yoy).toFixed(1)}%}`);
    }
    return parts.join('\n');
  }

  let series;
  if(singleCat){
    const catName = cats[0] || '';
    series = [{
      name: catName, type:'bar', barMaxWidth:48, barCategoryGap:'30%',
      data: MK.map((m,i)=>({
        value: monthlyMap[catName]?.[m]||0,
        yoy: monthlyMap[catName]?.[m+'_yoy']??null,
        itemStyle:{color:mColors[i], borderRadius:[3,3,0,0]}
      })),
      label:{show:true, position:'top', formatter:fmtVY, rich:yoyRich}
    }];
  } else {
    series = cats.map((cat,ci)=>({
      name:cat, type:'bar', barGap:'4%', barCategoryGap:'25%',
      itemStyle:{color:CS[ci%CS.length], borderRadius:[3,3,0,0]},
      data: MK.map(m=>({
        value: monthlyMap[cat]?.[m]||0,
        yoy: monthlyMap[cat]?.[m+'_yoy']??null
      })),
      // show=true 强制显示，formatter 自行决定每个柱子的内容
      label:{show:true, position:'top', formatter:fmtVY, rich:yoyRich}
    }));
  }

  chart.setOption({
    backgroundColor:'transparent',
    textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#8A9DC0'},
    tooltip:{
      ...TT, trigger:'axis', axisPointer:{type:'shadow'},
      formatter(params){
        let s=`<div style="font-weight:600;margin-bottom:6px;color:#C9A84C">${params[0].axisValue}</div>`;
        params.forEach(p=>{
          if(!p.value) return;
          const c = singleCat ? mColors[p.dataIndex] : CS[p.seriesIndex%CS.length];
          const yoy = p.data && p.data.yoy;
          const yoyS = yoy!=null
            ? `<span style="color:${yoy>=0?'#E85A5A':'#4CAF82'};margin-left:6px">${yoy>=0?'▲':'▼'}${Math.abs(yoy).toFixed(1)}%</span>` : '';
          s+=`<div style="display:flex;align-items:center;gap:6px;margin:3px 0">
            <span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${c}"></span>
            ${singleCat ? p.axisValue : p.seriesName}：<b>${p.value.toLocaleString()}</b>${yoyS}</div>`;
        });
        return s;
      }
    },
    legend: singleCat ? {show:false} : {
      data:cats, top:4,
      textStyle:{color:'#E8EEF8',fontSize:12}, itemWidth:12, itemHeight:8
    },
    grid:{top: singleCat?26:42, right:12, bottom:36, left:12, containLabel:true},
    xAxis:{type:'category', data:xLabels,
      axisLine:{lineStyle:{color:'#1e3a6e'}}, axisTick:{show:false},
      axisLabel:{color:'#E8EEF8',fontSize:11,interval:0}, splitLine:{show:false}},
    yAxis:{type:'value',
      axisLine:{show:false}, axisTick:{show:false},
      axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
      splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
    series
  }, true);
}

/* 月度版分动力类型：合并大类+细分 */
function buildMonthEnergyLayered(id, bigMonthly, subMonthly){
  const merged = {};
  merged['燃油车'] = bigMonthly['燃油车'] || {};
  merged['新能源'] = bigMonthly['新能源'] || {};
  merged['其中：纯电动'] = subMonthly['纯电动'] || {};
  merged['其中：插电混动'] = subMonthly['插电混动'] || {};
  merged['其中：增程式'] = subMonthly['增程式'] || {};
  const cats = ['燃油车','新能源','其中：纯电动','其中：插电混动','其中：增程式'];
  buildMonthBar(id, merged, cats);
}

/* ── buildLine: monthly trend ── */
function buildLine(id, seriesMap){
  const chart=gi(id);
  if(!chart) return;
  const names=Object.keys(seriesMap);
  chart.setOption({
    backgroundColor:'transparent',
    textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#8A9DC0'},
    tooltip:{...TT,trigger:'axis',
      formatter(params){
        let s=`<div style="font-weight:600;margin-bottom:6px;color:#C9A84C">${params[0].axisValue}</div>`;
        params.forEach(p=>{
          if(p.value==null) return;
          s+=`<div style="display:flex;align-items:center;gap:6px;margin:3px 0">
            <span style="display:inline-block;width:16px;height:3px;background:${CS[p.seriesIndex%CS.length]};border-radius:1px"></span>
            ${p.seriesName}：<b>${p.value.toLocaleString()}</b></div>`;
        });
        return s;
      }
    },
    legend:{data:names,top:4,textStyle:{color:'#E8EEF8',fontSize:12},itemWidth:16,itemHeight:3},
    grid:{top:46,right:12,bottom:36,left:12,containLabel:true},
    xAxis:{type:'category',data:MONTHS,
      axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
      axisLabel:{color:'#E8EEF8',fontSize:12},splitLine:{show:false}},
    yAxis:{type:'value',
      axisLine:{show:false},axisTick:{show:false},
      axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
      splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
    series:names.map((name,i)=>({
      name,type:'line',smooth:true,symbol:'circle',symbolSize:5,
      lineStyle:{color:CS[i%CS.length],width:2},
      itemStyle:{color:CS[i%CS.length]},
      data:Array.from({length:12},(_,m)=>seriesMap[name]?.[String(m+1)]??null),
      connectNulls:false
    }))
  }, true);
}

/* ══════════════════════════════════════════
   板块0 · 出口总览 (年度 / 月度 切换)
══════════════════════════════════════════ */
// 海关 Tab 元数据驱动的全局常量（从 DATA.meta 取，自动适配年份/月份扩展）
const AY = DATA.meta.years;                                    // ['2023','2024','2025','2026']
const S_ORD = ['中系','日系','欧系','美系','韩系'];
const R_ORD = DATA.meta.regionOrder || ['亚洲','欧洲（非俄）','俄罗斯','北美洲','南美洲','非洲','大洋洲'];
const MK = DATA.meta.monthKeys;                                // ['2025-1', ..., '2026-2']
// 月度短标签：'YYYY-M' → 'YY-MM'
const ML = {};
MK.forEach(k=>{ const [y,m] = k.split('-'); ML[k] = y.substring(2)+'-'+m.padStart(2,'0'); });
// 年度显示标签：完整年份用 '2025'，不完整年份用 '2026M1-2'
const HG_PARTIAL = DATA.meta.partialYear;
const HG_PARTIAL_M = DATA.meta.partialMonths || [];
const YR_LABELS = {};
AY.forEach(y=>{ YR_LABELS[y] = (y===HG_PARTIAL && HG_PARTIAL_M.length<12) ? y+'M1-'+HG_PARTIAL_M[HG_PARTIAL_M.length-1] : y; });
const MC = {};
const _mColorsBase = ['#4A90E2','#C9A84C','#5BC4A0','#E87B5A','#9B7FD4','#5BC4D4',
  '#E8C870','#D4697F','#6BAEE8','#8FBC5A','#B87FD4','#F0A050'];
const mColors = MK.map((_,i)=>_mColorsBase[i%_mColorsBase.length]);
MK.forEach((k,i)=> MC[k]=mColors[i]);

// 数据更新时间显示
(function updateHaiguanMeta(){
  const el = document.getElementById('haiguanMeta');
  if(!el) return;
  const fd = DATA.meta.dataFileDate;
  el.textContent = (fd ? '数据更新：'+fd+' ｜ ' : '') + '截至 '+DATA.meta.latestYear+'年'+DATA.meta.latestMonth+'月';
})();

// 板块1/2 年份筛选框：根据 AY/YR_LABELS 动态生成（不完整年份显示为 2026M1-2）
(function initYearFilters(){
  const palette = {'2023':'#4A90E2','2024':'#C9A84C','2025':'#5BC4A0','2026':'#E85A5A','2027':'#9B7FD4','2028':'#5BC4D4'};
  const colorOf = y => palette[y] || '#6B7DA0';
  // 注入按年份的 ychk CSS（兼容 AY 中任意年份）
  let css = '';
  AY.forEach(y => {
    const c = colorOf(y); const sfx = y.slice(2);
    css += `.ychk.y${sfx}{border-color:${c}66}`;
    css += `.ychk.y${sfx}.on{background:${c};border-color:${c};color:#fff}`;
    css += `.y${sfx} .ydot{background:${c}}`;
  });
  const styleEl = document.createElement('style');
  styleEl.textContent = css;
  document.head.appendChild(styleEl);

  // 板块1：年份复选框（默认全部勾选）
  const s1box = document.getElementById('s1yearBox');
  if(s1box){
    AY.forEach(y => {
      const lbl = document.createElement('label');
      lbl.className = 'ychk y'+y.slice(2)+' on';
      lbl.innerHTML = `<input type="checkbox" value="${y}" checked><span class="ydot"></span>${YR_LABELS[y]||y}`;
      s1box.appendChild(lbl);
    });
  }

  // 板块2：年份单选（"全部"+各年份）
  const s2box = document.getElementById('s2yearBox');
  if(s2box){
    const allLbl = document.createElement('label');
    allLbl.className = 'ychk on'; allLbl.id = 's2y_all';
    allLbl.innerHTML = `<input type="radio" name="s2yr" value="all" checked><span class="ydot" style="background:#C9A84C"></span>全部`;
    s2box.appendChild(allLbl);
    AY.forEach(y => {
      const lbl = document.createElement('label');
      lbl.className = 'ychk y'+y.slice(2);
      lbl.innerHTML = `<input type="radio" name="s2yr" value="${y}"><span class="ydot"></span>${YR_LABELS[y]||y}`;
      s2box.appendChild(lbl);
    });
  }
})();

let s0mode = 'annual';

function renderM0Ser(){
  const checked = [...document.querySelectorAll('#mSerFilter input:checked')].map(c=>c.value);
  if(!checked.length) return;
  const filtered = {};
  checked.forEach(s=>{ if(DATA.monthly.bySeries[s]) filtered[s] = DATA.monthly.bySeries[s]; });
  buildMonthBar('m0ser', filtered, checked.filter(s=>DATA.monthly.bySeries[s]));
}

function renderM0Oem(){
  const codedTop15 = DATA.codedTop15;
  const checkedIdx = [...document.querySelectorAll('#mOemFilter input:checked')].map(c=>parseInt(c.value));
  const selected = checkedIdx.map(i=>codedTop15[i]).filter(Boolean);
  // 追加通过模态框选中的额外车企
  s0extraOems.forEach(oem=>{ if(!selected.includes(oem)) selected.push(oem); });
  if(!selected.length) return;
  const mData = DATA.monthly.oemByEnergy['all'].data;
  const merged = {};
  selected.forEach(oem=>{ merged[oem] = mData[oem] || {}; });
  buildMonthBar('m0oem', merged, selected);
}

// 构建板块0额外车企筛选框
// 板块0车企筛选状态：默认codedTop15
let s0extraOems = [];

(function buildS0OemExtra(){
  const container = document.getElementById('s0oemExtra');
  // "更多车企"按钮
  const moreBtn = document.createElement('button');
  moreBtn.className = 'show-more-btn';
  moreBtn.style.cssText = 'margin:0;padding:3px 12px;font-size:11px';
  moreBtn.textContent = '更多车企';
  moreBtn.addEventListener('click', openS0OemSelector);
  container.appendChild(moreBtn);
  // "放大图"按钮
  const zoomBtn = document.createElement('button');
  zoomBtn.className = 'show-more-btn';
  zoomBtn.style.cssText = 'margin:0;padding:3px 12px;font-size:11px';
  zoomBtn.textContent = '放大图';
  zoomBtn.addEventListener('click', openS0OemZoom);
  container.appendChild(zoomBtn);
})();

function openS0OemSelector(){
  const allOems = DATA.oemByEnergy['all'].order;
  const codedSet = new Set(DATA.codedAll);
  const coded = allOems.filter(o => codedSet.has(o));
  const uncoded = allOems.filter(o => !codedSet.has(o));
  const mask = document.createElement('div'); mask.className='modal-mask';
  let html = '<div class="modal-box" style="max-width:900px"><button class="modal-close">&times;</button>';
  html += '<div class="cw-lbl" style="margin-bottom:12px;font-size:17px">选择展示车企</div>';
  html += '<div style="margin-bottom:8px;font-size:12px;color:var(--t2)">勾选后点击"确认"更新图表；带股票代码车企默认展示TOP15</div>';
  // 带代码车企
  html += '<div style="margin-bottom:6px;font-size:13px;font-weight:700;color:var(--gold)">重点车企（带股票代码）</div>';
  html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
  coded.forEach(oem => {
    const isTop15 = DATA.codedTop15.includes(oem);
    const checked = isTop15 || s0extraOems.includes(oem) ? 'checked' : '';
    const disabled = isTop15 ? 'disabled' : '';
    html += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px">';
    html += '<input type="checkbox" value="'+oem+'" '+checked+' '+disabled+' data-group="coded"> '+oem+'</label>';
  });
  html += '</div>';
  // 其他车企
  html += '<div style="margin-bottom:6px;font-size:13px;font-weight:700;color:var(--t2)">其他车企</div>';
  html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
  uncoded.forEach(oem => {
    const checked = s0extraOems.includes(oem) ? 'checked' : '';
    html += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px">';
    html += '<input type="checkbox" value="'+oem+'" '+checked+' data-group="uncoded"> '+oem+'</label>';
  });
  html += '</div>';
  html += '<div style="text-align:center"><button id="s0oemConfirm" style="background:var(--blue);color:#fff;border:none;padding:8px 32px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">确认</button></div>';
  html += '</div>';
  mask.innerHTML = html;
  document.body.appendChild(mask);
  mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
  mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
  mask.querySelector('#s0oemConfirm').addEventListener('click',()=>{
    const checks = [...mask.querySelectorAll('input[type=checkbox]:not(:disabled)')];
    s0extraOems = checks.filter(c=>c.checked).map(c=>c.value);
    mask.remove();
    renderS0();
  });
}

function openS0OemZoom(){
  const mask = document.createElement('div'); mask.className='modal-mask';
  mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button>'+
    '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">重点车企出口销量（放大视图）</div>'+
    '<div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>'+
    '<div id="c0oemZoom" style="height:600px"></div></div>';
  document.body.appendChild(mask);
  mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
  mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
  setTimeout(()=>{
    if(_c['c0oemZoom']) delete _c['c0oemZoom'];
    if(s0mode === 'annual'){
      const showOems = [...DATA.codedTop15, ...s0extraOems];
      buildBar('c0oemZoom', showOems, DATA.oemByEnergy['all'].data, AY, {keyLabels: YR_LABELS});
    } else {
      // 月度：读取当前checkbox选中的车企
      const checkedIdx = [...document.querySelectorAll('#mOemFilter input:checked')].map(c=>parseInt(c.value));
      const selected = checkedIdx.map(i=>DATA.codedTop15[i]).filter(Boolean);
      s0extraOems.forEach(oem=>{ if(!selected.includes(oem)) selected.push(oem); });
      const mData = DATA.monthly.oemByEnergy['all'].data;
      const merged = {};
      selected.forEach(oem=>{ merged[oem] = mData[oem] || {}; });
      buildMonthBar('c0oemZoom', merged, selected);
    }
    addModalExport(mask.querySelector('.modal-box'), 'c0oemZoom', '重点车企出口销量');
  },100);
}

function renderS0(){
  const sub = document.getElementById('s0sub');
  const annualDiv = document.getElementById('s0annual');
  const monthlyDiv = document.getElementById('s0monthly');

  if(s0mode === 'annual'){
    sub.textContent = '各年度总量与同比增速';
    annualDiv.style.display = '';
    monthlyDiv.style.display = 'none';
    buildBar('c0tot',['中国乘用车总出口'], DATA.total, AY, {keyLabels: YR_LABELS});
    buildBar('c0ser', S_ORD.filter(s=>DATA.bySeries[s]), DATA.bySeries, AY, {keyLabels: YR_LABELS});
    buildEnergyLayered('c0eng', DATA.byEnergyBig, DATA.byEnergySub, AY, {keyLabels: YR_LABELS});
    buildBar('c0reg', R_ORD.filter(r=>DATA.byRegion[r]), DATA.byRegion, AY, {keyLabels: YR_LABELS});
    // 重点车企：默认显示codedTop15，额外车企通过筛选框可加
    const s0ShowOems = [...DATA.codedTop15, ...s0extraOems];
    buildBar('c0oem', s0ShowOems, DATA.oemByEnergy['all'].data, AY, {keyLabels: YR_LABELS});
  } else {
    sub.textContent = MK[0].split('-')[0]+'年'+MK[0].split('-')[1]+'月起至'+DATA.meta.latestYear+'年'+DATA.meta.latestMonth+'月各月出口量';
    annualDiv.style.display = 'none';
    monthlyDiv.style.display = '';
    // 动态生成月度面板（每张图独占一行）
    if(!monthlyDiv._built){
      let html = '';
      // Row1: 总体 + 分系别 同一排（等宽）
      html += `<div class="g2 mt14">`;
      // 左：总体
      html += `<div class="cw"><div class="cw-lbl">乘用车出口总量</div><div id="m0tot" style="height:300px"></div></div>`;
      // 右：分系别（含筛选项）
      html += `<div class="cw">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px">
          <div class="cw-lbl" style="margin-bottom:0">分系别</div>
          <div class="fbar" id="mSerFilter" style="margin:0;padding:6px 10px;background:transparent;border:none">`;
      S_ORD.forEach(s=>{
        const checked = s==='中系' ? 'checked' : '';
        const on = s==='中系' ? ' on' : '';
        html += `<label class="ychk${on}">
          <input type="checkbox" value="${s}" ${checked}><span class="ydot" style="background:${CS[S_ORD.indexOf(s)%CS.length]}"></span>${s}</label>`;
      });
      html += `</div></div><div class="cw-note">可多选系别筛选，默认中系</div><div id="m0ser" style="height:300px"></div></div>`;
      html += `</div>`;
      // Row2: 动力类型 + 分区域 同一排
      html += `<div class="g2 mt14">`;
      // 左：动力类型（只显示燃油车+新能源，细分弹窗展示）
      html += `<div class="cw" id="m0engWrap"><div class="cw-lbl">分动力类型</div><div id="m0eng" style="height:300px"></div>
        <div class="cw-note">注：新能源 = 纯电动 + 插电混动 + 增程式</div></div>`;
      // 右：分区域（只显示亚洲+欧洲，其余弹窗展示）
      html += `<div class="cw" id="m0regWrap"><div class="cw-lbl">分区域</div><div id="m0reg" style="height:300px"></div><div class="cw-note">注：欧洲拆分为"俄罗斯"和"欧洲（非俄）"单独统计</div></div>`;
      html += `</div>`;
      // 重点车企: 默认展示codedTop15中第一个，通过更多车企按钮选择
      const allMOems = DATA.oemByEnergy['all'].order;
      const defaultMOem = DATA.codedTop15[0] || allMOems[0] || '';
      html += `<div class="cw mt14">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap">
          <div class="cw-lbl" style="margin-bottom:0">重点车企出口销量（按2025年出口量排序）</div>
          <div style="display:flex;gap:6px;align-items:center">
            <div class="fbar" id="mOemFilter" style="margin:0;padding:6px 10px;background:transparent;border:none;flex-wrap:wrap;gap:6px">`;
      // 默认只显示codedTop15作为快捷按钮
      DATA.codedTop15.forEach((oem,oi)=>{
        const checked = oem===defaultMOem ? 'checked' : '';
        const on = oem===defaultMOem ? ' on' : '';
        html += `<label class="ychk${on}">
          <input type="checkbox" value="${oi}" ${checked}><span class="ydot" style="background:${CS[oi%CS.length]}"></span>${oem}</label>`;
      });
      html += `</div>
            <button class="show-more-btn" id="mOemMoreBtn" style="margin:0;padding:3px 12px;font-size:11px">更多车企</button>
            <button class="show-more-btn" id="mOemZoomBtn" style="margin:0;padding:3px 12px;font-size:11px">放大图</button>
          </div>
        </div><div class="cw-note">注：默认展示带股票代码的TOP15车企；可通过"更多车企"添加其他车企展示；支持放大图查看</div><div id="m0oem" style="height:320px"></div></div>`;

      monthlyDiv.innerHTML = html;
      // 系别筛选事件
      document.querySelectorAll('#mSerFilter .ychk').forEach(lbl=>{
        lbl.addEventListener('click', function(){
          const cb = this.querySelector('input');
          const next = !cb.checked; cb.checked = next;
          this.classList.toggle('on', next);
          renderM0Ser();
        });
      });
      // 车企筛选事件
      document.querySelectorAll('#mOemFilter .ychk').forEach(lbl=>{
        lbl.addEventListener('click', function(){
          const cb = this.querySelector('input');
          const next = !cb.checked; cb.checked = next;
          this.classList.toggle('on', next);
          renderM0Oem();
        });
      });
      // 更多车企按钮
      const mMoreBtn = document.getElementById('mOemMoreBtn');
      if(mMoreBtn) mMoreBtn.addEventListener('click', openS0OemSelector);
      // 放大图按钮
      const mZoomBtn = document.getElementById('mOemZoomBtn');
      if(mZoomBtn) mZoomBtn.addEventListener('click', openS0OemZoom);
      monthlyDiv._built = true;
    }
    // 渲染月度图表
    buildMonthBar('m0tot', DATA.monthly.total, ['中国乘用车总出口']);
    renderM0Ser();
    const MO = {hideLegend:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};
    const MOL = {hideLegend:true, showMonthLabel:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};

    // 分动力类型：只显示燃油车+新能源大类
    const engBigOnly = {};
    engBigOnly['燃油车'] = DATA.monthly.byEnergyBig['燃油车'] || {};
    engBigOnly['新能源'] = DATA.monthly.byEnergyBig['新能源'] || {};
    buildBar('m0eng', ['燃油车','新能源'], engBigOnly, MK, MOL);
    // 查看更多：弹窗展示细分
    (function(){
      const wrap = document.getElementById('m0engWrap');
      let btn = wrap.querySelector('.show-more-btn');
      if(!btn){ btn=document.createElement('button'); btn.className='show-more-btn'; wrap.appendChild(btn); }
      btn.textContent = '查看细分动力类型';
      btn.onclick = function(){
        const mask = document.createElement('div'); mask.className='modal-mask';
        mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button>'+
          '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">分动力类型（含细分）</div>'+
          '<div id="m0engModal" style="height:500px"></div></div>';
        document.body.appendChild(mask);
        mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
        mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
        setTimeout(()=>{
          const mChart = echarts.init(document.getElementById('m0engModal'));
          const allCats = ['燃油车','新能源','其中：纯电动','其中：插电混动','其中：增程式'];
          const merged = {};
          merged['燃油车'] = DATA.monthly.byEnergyBig['燃油车']||{};
          merged['新能源'] = DATA.monthly.byEnergyBig['新能源']||{};
          merged['其中：纯电动'] = DATA.monthly.byEnergySub['纯电动']||{};
          merged['其中：插电混动'] = DATA.monthly.byEnergySub['插电混动']||{};
          merged['其中：增程式'] = DATA.monthly.byEnergySub['增程式']||{};
          const n2 = allCats.length * MK.length;
          const mLblFs2 = n2>60?0:n2>36?7:8;
          const series2 = MK.map((k,ki)=>({
            name:ML[k]||k, type:'bar', barGap:'8%', barCategoryGap:'25%',
            itemStyle:{color:MC[k], borderRadius:[3,3,0,0]},
            data:allCats.map(cat=>({
              value:(merged[cat]||{})[k]||0,
              yoy:(merged[cat]||{})[k+'_yoy']??null
            })),
            // 始终 show=true：数值在拥挤时省略，yoy（仅 26M1/M2 等有可比基数）总是显示
            label:{show:true, position:'top',
              formatter(p){
                const yoy = p.data && p.data.yoy;
                const hasYoy = yoy !== null && yoy !== undefined;
                if(!p.value && !hasYoy) return '';
                const parts = [];
                if(mLblFs2 > 0 && p.value) parts.push('{v|'+fmt(p.value)+'}');
                if(hasYoy){
                  const pos = yoy >= 0;
                  parts.push('{'+(pos?'r':'g')+'|'+(pos?'▲':'▼')+Math.abs(yoy).toFixed(1)+'%}');
                }
                return parts.join('\n');
              },
              rich:{
                v:{fontSize:Math.max(mLblFs2,8),color:'#FFFFFF',lineHeight:12},
                r:{fontSize:8,color:'#E85A5A',lineHeight:12},
                g:{fontSize:8,color:'#4CAF82',lineHeight:12}
              }}
          }));
          mChart.setOption({
            backgroundColor:'transparent',
            textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#E8EEF8'},
            tooltip:{...TT,trigger:'axis',axisPointer:{type:'shadow'},
              formatter(params){
                let s='<div style="font-weight:600;margin-bottom:6px;color:#C9A84C">'+params[0].axisValue+'</div>';
                params.forEach(p=>{
                  if(!p.value) return;
                  const yoy = p.data && p.data.yoy;
                  const yoyS = yoy!=null
                    ? '<span style="color:'+(yoy>=0?'#E85A5A':'#4CAF82')+';margin-left:6px">'+(yoy>=0?'▲':'▼')+Math.abs(yoy).toFixed(1)+'%</span>' : '';
                  const c = MC[MK[p.seriesIndex]] || '#4A90E2';
                  s+='<div style="display:flex;align-items:center;gap:6px;margin:3px 0">'+
                    '<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:'+c+'"></span>'+
                    p.seriesName+'：<b>'+p.value.toLocaleString()+'</b>'+yoyS+'</div>';
                });
                return s;
              }
            },
            legend:{show:false},
            grid:{top:20,right:12,bottom:50,left:12,containLabel:true},
            xAxis:{type:'category',data:allCats,axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,interval:0},splitLine:{show:false}},
            yAxis:{type:'value',axisLine:{show:false},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
              splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
            series:series2
          },true);
          addModalExport(mask.querySelector('.modal-box'), 'm0engModal', '分动力类型（含细分）');
        },100);
      };
    })();

    // 分区域：只显示亚洲+欧洲（非俄）
    const regShow = ['亚洲','欧洲（非俄）'].filter(r=>DATA.monthly.byRegion[r]);
    buildBar('m0reg', regShow, DATA.monthly.byRegion, MK, MO);
    // 查看更多：弹窗展示全部区域
    (function(){
      const wrap = document.getElementById('m0regWrap');
      let btn = wrap.querySelector('.show-more-btn');
      if(!btn){ btn=document.createElement('button'); btn.className='show-more-btn'; wrap.appendChild(btn); }
      btn.textContent = '查看更多区域';
      btn.onclick = function(){
        const allRegs = R_ORD.filter(r=>DATA.monthly.byRegion[r]);
        const mask = document.createElement('div'); mask.className='modal-mask';
        mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button>'+
          '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">分区域（全部）</div>'+
          '<div id="m0regModal" style="height:500px"></div></div>';
        document.body.appendChild(mask);
        mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
        mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
        setTimeout(()=>{
          const mChart = echarts.init(document.getElementById('m0regModal'));
          const n2 = allRegs.length * MK.length;
          const mLblFs2 = n2>60?0:n2>36?7:8;
          const series2 = MK.map((k,ki)=>({
            name:ML[k]||k, type:'bar', barGap:'8%', barCategoryGap:'25%',
            itemStyle:{color:MC[k], borderRadius:[3,3,0,0]},
            data:allRegs.map(r=>({value:(DATA.monthly.byRegion[r]||{})[k]||0})),
            label:{show:mLblFs2>0, position:'top',
              formatter:p=>p.value?'{v|'+fmt(p.value)+'}':'',
              rich:{v:{fontSize:mLblFs2,color:'#FFFFFF',lineHeight:14}}}
          }));
          mChart.setOption({
            backgroundColor:'transparent',
            textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#E8EEF8'},
            tooltip:{...TT,trigger:'axis',axisPointer:{type:'shadow'}},
            legend:{show:false},
            grid:{top:20,right:12,bottom:50,left:12,containLabel:true},
            xAxis:{type:'category',data:allRegs,axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,interval:0,rotate:allRegs.length>5?-20:0},splitLine:{show:false}},
            yAxis:{type:'value',axisLine:{show:false},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
              splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
            series:series2
          },true);
          addModalExport(mask.querySelector('.modal-box'), 'm0regModal', '分区域（全部）');
        },100);
      };
    })();

    renderM0Oem();
  }
}

// Tab toggle
document.querySelectorAll('#s0tabs .tab-btn').forEach(btn=>{
  btn.addEventListener('click', function(){
    document.querySelectorAll('#s0tabs .tab-btn').forEach(b=>b.classList.remove('on'));
    this.classList.add('on');
    s0mode = this.dataset.mode;
    renderS0();
  });
});

renderS0();

/* ══════════════════════════════════════════
   板块一 · 分动力类型视角
   筛选：年度/月度, 年份复选, 系别复选, 动力类型, 区域下钻
══════════════════════════════════════════ */
let s1mode = 'annual';
let s1years = AY.slice();
let s1seriesVal = 'all'; // 'all' or specific series name
let s1energy = 'all';
let s1region = R_ORD[0];

const ALL_SERIES = ['中系','日系','欧系','美系','韩系'];
// Helper: get selected series array from single-select value
function getS1Series(){ return s1seriesVal === 'all' ? ALL_SERIES : [s1seriesVal]; }


/* ── 工具：合并多系别数据 ──
 *  单系别直接返回源（保留预计算的 _yoy，含部分年同期口径）；
 *  多系别才走合并并重算 yoy（月度按同月同比，年度按全年朴素比）。 */
function sumSeriesData(perSeriesSrc, selectedSeries, keyList){
  if(selectedSeries.length === 1){
    return perSeriesSrc[selectedSeries[0]] || {};
  }
  const merged = {};
  selectedSeries.forEach(s => {
    const sd = perSeriesSrc[s] || {};
    Object.keys(sd).forEach(cat => {
      if(!merged[cat]) merged[cat] = {};
      keyList.forEach(k => {
        merged[cat][k] = (merged[cat][k]||0) + (sd[cat][k]||0);
      });
    });
  });
  recomputeYoyAfterMerge(merged, keyList);
  return merged;
}

function sumSeriesCountryData(perSeriesSrc, selectedSeries, region, keyList){
  if(selectedSeries.length === 1){
    return (perSeriesSrc[selectedSeries[0]]||{})[region] || {};
  }
  const merged = {};
  selectedSeries.forEach(s => {
    const rd = (perSeriesSrc[s]||{})[region] || {};
    Object.keys(rd).forEach(c => {
      if(!merged[c]) merged[c] = {};
      keyList.forEach(k => {
        merged[c][k] = (merged[c][k]||0) + (rd[c][k]||0);
      });
    });
  });
  recomputeYoyAfterMerge(merged, keyList);
  return merged;
}

/* ── 多系别合并后重算 yoy ──
 *  keyList 为年度 ['YYYY',...] 时：cur/prev-1（注意：partial 年会偏，多系别极少用）
 *  keyList 为月度 ['YYYY-M',...] 时：按同月同比 y/(y-1) */
function recomputeYoyAfterMerge(merged, keyList){
  if(!keyList.length) return;
  const isMonthly = /^\d{4}-\d+$/.test(String(keyList[0]));
  if(isMonthly){
    Object.keys(merged).forEach(c => {
      const d = merged[c];
      keyList.forEach(k => {
        const [y, m] = String(k).split('-');
        const prevKey = (parseInt(y)-1) + '-' + m;
        const prev = d[prevKey], cur = d[k];
        if(prev && prev > 0 && cur !== undefined){
          d[k+'_yoy'] = Math.round((cur/prev-1)*1000)/10;
        }
      });
    });
  } else {
    Object.keys(merged).forEach(c => {
      const d = merged[c];
      for(let i=1;i<keyList.length;i++){
        const prev=d[keyList[i-1]], cur=d[keyList[i]];
        if(prev>0) d[keyList[i]+'_yoy']=Math.round((cur/prev-1)*1000)/10;
      }
    });
  }
}

/* ── 获取区域数据（支持系别+动力筛选）── */
function getS1RegionData(isMonthly){
  const s1series = getS1Series();
  const allSelected = s1seriesVal === 'all';
  const keys = isMonthly ? MK : s1years;

  if(allSelected){
    // 全选系别：用全局数据
    if(s1energy === 'all') return isMonthly ? DATA.monthly.byRegion : DATA.byRegion;
    const src = isMonthly ? DATA.monthly.regionByEnergy : DATA.regionByEnergy;
    return src[s1energy] || {};
  }

  // 部分系别：合并 per-series 数据
  if(s1energy === 'all'){
    const src = isMonthly ? DATA.monthly.s1BySeriesRegion : DATA.s1BySeriesRegion;
    return sumSeriesData(src, s1series, keys);
  }
  const src = isMonthly ? DATA.monthly.s1BySeriesEnergyRegion : DATA.s1BySeriesEnergyRegion;
  const perSeries = {};
  s1series.forEach(s => { perSeries[s] = (src[s]||{})[s1energy] || {}; });
  return sumSeriesData(perSeries, s1series, keys);
}

/* ── 获取国家下钻数据 ── */
function getS1CountryData(region, isMonthly){
  const s1series = getS1Series();
  const allSelected = s1seriesVal === 'all';
  const keys = isMonthly ? MK : s1years;

  if(allSelected){
    if(s1energy === 'all'){
      const src = isMonthly ? DATA.monthly.regionCountry : DATA.regionCountry;
      return src[region] || {};
    }
    const src = isMonthly ? DATA.monthly.countryByEnergyRegion : DATA.countryByEnergyRegion;
    return (src[region]&&src[region][s1energy]) || {};
  }

  if(s1energy === 'all'){
    const src = isMonthly ? DATA.monthly.s1BySeriesRegionCountry : DATA.s1BySeriesRegionCountry;
    return sumSeriesCountryData(src, s1series, region, keys);
  }
  const src = isMonthly ? DATA.monthly.s1BySeriesEnergyRegionCountry : DATA.s1BySeriesEnergyRegionCountry;
  const perSeries = {};
  s1series.forEach(s => { perSeries[s] = {}; perSeries[s][region] = ((src[s]||{})[s1energy]||{})[region] || {}; });
  return sumSeriesCountryData(perSeries, s1series, region, keys);
}

/* ── 获取车企排名数据（按系别+动力筛选）── */
function getS1OemData(isMonthly){
  const s1series = getS1Series();
  const key = s1energy === 'all' ? 'all' : s1energy;
  const src = isMonthly ? DATA.monthly.oemByEnergy : DATA.oemByEnergy;
  const info = src[key] || {order:[], data:{}};
  // Filter by selected series
  const oemSer = DATA.oemSeries || {};
  const skip = ['其他','na','NA','Na'];
  const filtered = info.order.filter(oem => {
    if(skip.includes(oem) || skip.includes(oem.trim().toLowerCase())) return false;
    const s = oemSer[oem];
    return s && s1series.includes(s);
  });
  return {order: filtered, data: info.data};
}

/* ── 构建区域筛选按钮 ── */
function buildRegionFilter(){
  const container = document.getElementById('s1regFilter');
  container.innerHTML = '';
  const allRegions = ['全部区域', ...R_ORD];
  allRegions.forEach((r,i)=>{
    const btn = document.createElement('span');
    btn.className = 'ychk' + (s1region===r ? ' on' : '');
    btn.style.cssText = 'cursor:pointer';
    const dotColor = r === '全部区域' ? '#C9A84C' : CS[(i-1)%CS.length];
    btn.innerHTML = `<span class="ydot" style="width:6px;height:6px;background:${dotColor}"></span>${r}`;
    btn.addEventListener('click',()=>{
      s1region = (s1region===r) ? null : r;
      buildRegionFilter();
      renderS1Country();
    });
    container.appendChild(btn);
  });
}

/* ── 渲染区域图（默认前5，可展开） ── */
let s1regExpanded = false;
function renderS1Region(){
  const isMonthly = s1mode === 'monthly';
  const regData = getS1RegionData(isMonthly);
  const allRegions = R_ORD.filter(r=>regData[r]);
  const regions = s1regExpanded ? allRegions : allRegions.slice(0,5);

  if(isMonthly){
    const MOL = {hideLegend:true, showMonthLabel:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};
    buildBar('c1reg', regions, regData, MK, MOL);
  } else {
    buildBar('c1reg', regions, regData, s1years, {keyLabels: YR_LABELS});
  }

  const lbl = document.getElementById('s1regLbl');
  const suffix = s1energy==='all' ? '' : '（'+s1energy+'）';
  lbl.textContent = '按区域' + suffix;

  // 查看更多按钮
  const wrap = document.getElementById('s1regWrap');
  let btn = wrap.querySelector('.show-more-btn');
  if(allRegions.length > 5){
    if(!btn){ btn=document.createElement('button'); btn.className='show-more-btn'; wrap.appendChild(btn);
      btn.addEventListener('click',()=>{ s1regExpanded=!s1regExpanded; renderS1Region(); }); }
    btn.textContent = s1regExpanded ? '收起' : '查看更多（共'+allRegions.length+'项）';
  } else if(btn){ btn.remove(); }
}

/* ── 渲染国家下钻图（默认前5，弹窗查看前15+其他） ── */
function mergeOthers(sorted, countryData, keys, maxShow){
  // 前maxShow个 + 其余合并为"其他"
  if(sorted.length <= maxShow) return {cats: sorted, data: countryData};
  const topN = sorted.slice(0, maxShow);
  const rest = sorted.slice(maxShow);
  const merged = {};
  // copy top countries
  topN.forEach(c=>{ merged[c] = countryData[c]; });
  // merge rest into "其他"
  const other = {};
  keys.forEach(k=>{ other[k] = 0; });
  rest.forEach(c=>{ keys.forEach(k=>{ other[k] += (countryData[c]||{})[k]||0; }); });
  merged['其他'] = other;
  return {cats: [...topN, '其他'], data: merged};
}

function renderS1Country(){
  const drillLbl = document.getElementById('s1drillLbl');
  if(!s1region){
    drillLbl.textContent = '点击右侧区域筛选查看国家详情';
    const chart = gi('c1country'); if(chart) chart.clear();
    const wrap = document.getElementById('s1countryWrap');
    const eb = wrap.querySelector('.show-more-btn'); if(eb) eb.remove();
    return;
  }
  const isMonthly = s1mode === 'monthly';
  const countryData = getS1CountryData(s1region, isMonthly);
  const energySuffix = s1energy==='all' ? '' : '（'+s1energy+'）';
  drillLbl.textContent = '国家明细 · ' + s1region + energySuffix;

  const countries = Object.keys(countryData);
  let sorted;
  if(isMonthly){
    sorted = countries.sort((a,b)=>{
      const sumA = MK.reduce((s,m)=>(s+(countryData[a][m]||0)),0);
      const sumB = MK.reduce((s,m)=>(s+(countryData[b][m]||0)),0);
      return sumB - sumA;
    });
  } else {
    const latestYr = s1years[s1years.length-1];
    sorted = countries.sort((a,b)=>(countryData[b][latestYr]||0)-(countryData[a][latestYr]||0));
  }
  const top = sorted.slice(0,5);

  if(isMonthly){
    const MOL = {hideLegend:true, showMonthLabel:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};
    buildBar('c1country', top, countryData, MK, MOL);
  } else {
    buildBar('c1country', top, countryData, s1years, {keyLabels: YR_LABELS});
  }

  // 查看更多按钮 → 弹窗展示前15+其他
  const wrap = document.getElementById('s1countryWrap');
  let btn = wrap.querySelector('.show-more-btn');
  if(sorted.length > 5){
    if(!btn){ btn=document.createElement('button'); btn.className='show-more-btn'; wrap.appendChild(btn); }
    btn.textContent = '查看更多（共'+sorted.length+'国）';
    btn.onclick = function(){
      const _isMonthly = s1mode === 'monthly';
      const keys = _isMonthly ? MK : s1years;
      const {cats, data} = mergeOthers(sorted, countryData, keys, 15);
      const mask = document.createElement('div');
      mask.className = 'modal-mask';
      mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button>'+
        '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">国家明细 · '+s1region+energySuffix+'（前15+其他）</div>'+
        '<div id="c1countryModal" style="height:500px"></div></div>';
      document.body.appendChild(mask);
      mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
      mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
      setTimeout(()=>{
        const mChart = echarts.init(document.getElementById('c1countryModal'));
        if(_isMonthly){
          const n2 = cats.length * MK.length;
          const mLblFs2 = n2>60?0:n2>36?7:8;
          const series2 = MK.map((k,ki)=>({
            name:ML[k]||k, type:'bar', barGap:'8%', barCategoryGap:'25%',
            itemStyle:{color:MC[k], borderRadius:[3,3,0,0]},
            data:cats.map(cat=>({value:(data[cat]||{})[k]||0})),
            label:{show:mLblFs2>0, position:'top',
              formatter:p=>p.value?'{v|'+fmt(p.value)+'}':'',
              rich:{v:{fontSize:mLblFs2,color:'#FFFFFF',lineHeight:14}}}
          }));
          mChart.setOption({
            backgroundColor:'transparent',
            textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#E8EEF8'},
            tooltip:{...TT,trigger:'axis',axisPointer:{type:'shadow'}},
            legend:{show:false},
            grid:{top:20,right:12,bottom:50,left:12,containLabel:true},
            xAxis:{type:'category',data:cats,axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,interval:0,rotate:cats.length>10?-25:0},splitLine:{show:false}},
            yAxis:{type:'value',axisLine:{show:false},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
              splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
            series:series2
          },true);
        } else {
          const yrs = s1years;
          const series2 = yrs.map((yr,yi)=>({
            name:yr+'年', type:'bar', barGap:'4%', barCategoryGap:'30%',
            itemStyle:{color:YC[yr]||CS[yi%CS.length], borderRadius:[3,3,0,0]},
            data:cats.map(cat=>{
              const d=data[cat]||{};
              return {value:d[yr]||0, yoy:d[yr+'_yoy']??null};
            }),
            label:{show:true, position:'top',
              formatter(p){
                const v=fmt(p.value);
                const yoy=p.data&&p.data.yoy;
                if(yoy!==null&&yoy!==undefined){
                  const pos=yoy>=0;
                  return '{v|'+v+'}\n{'+(pos?'r':'g')+'|'+(pos?'▲':'▼')+Math.abs(yoy).toFixed(1)+'%}';
                }
                return '{v|'+v+'}';
              },
              rich:{v:{fontSize:9,color:'#FFFFFF',lineHeight:16},r:{fontSize:8,color:'#E85A5A',lineHeight:14},g:{fontSize:8,color:'#4CAF82',lineHeight:14}}}
          }));
          mChart.setOption({
            backgroundColor:'transparent',
            textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#E8EEF8'},
            tooltip:{...TT,trigger:'axis',axisPointer:{type:'shadow'}},
            legend:{data:yrs.map(y=>y+'年'),top:4,textStyle:{color:'#E8EEF8',fontSize:12},itemWidth:12,itemHeight:8},
            grid:{top:42,right:12,bottom:50,left:12,containLabel:true},
            xAxis:{type:'category',data:cats,axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,interval:0,rotate:cats.length>10?-25:0},splitLine:{show:false}},
            yAxis:{type:'value',axisLine:{show:false},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
              splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
            series:series2
          },true);
        }
        addModalExport(mask.querySelector('.modal-box'), 'c1countryModal', '国家明细 · '+s1region);
      },100);
    };
  } else if(btn){ btn.remove(); }
}

/* ── 渲染分车企排名 ── */
let s1oemFilterBuilt = false;

function buildS1OemFilter(order){
  const container = document.getElementById('s1oemFilter');
  container.innerHTML = '';
  const defaultOem = order[0] || '';
  order.forEach((oem, i) => {
    const on = oem === defaultOem ? ' on' : '';
    const checked = oem === defaultOem ? 'checked' : '';
    const lbl = document.createElement('label');
    lbl.className = 'ychk' + on;
    lbl.style.cssText = '';
    lbl.innerHTML = `<input type="checkbox" value="${i}" ${checked}><span class="ydot" style="width:6px;height:6px;background:${CS[i%CS.length]}"></span>${oem}`;
    lbl.addEventListener('click', function(){
      const cb = this.querySelector('input');
      const next = !cb.checked; cb.checked = next;
      this.classList.toggle('on', next);
      renderS1OemChart();
    });
    container.appendChild(lbl);
  });
}

function renderS1OemChart(){
  const oemInfo = getS1OemData(true);
  const order = oemInfo.order;
  const data = oemInfo.data;
  const filterBox = document.getElementById('s1oemFilter');
  const selected = [...filterBox.querySelectorAll('input:checked')].map(cb=>{
    return order[parseInt(cb.value)];
  }).filter(Boolean);
  if(!selected.length) return;
  const merged = {};
  selected.forEach(oem=>{ merged[oem] = data[oem] || {}; });
  buildMonthBar('c1oem', merged, selected);
}

let s1extraOems = []; // 板块一年度额外勾选的车企
let s1monthlyOems = []; // 板块一月度选中的车企（独立于年度）
let s1monthlyInited = false;

function renderS1Oem(){
  const isMonthly = s1mode === 'monthly';
  const oemInfo = getS1OemData(isMonthly);
  const order = oemInfo.order;
  const data = oemInfo.data;
  const energySuffix = s1energy==='all' ? '' : '（'+s1energy+'）';
  const serSuffix = s1seriesVal==='all' ? '' : '（'+s1seriesVal+'）';

  const lbl = document.getElementById('s1oemLbl');
  if(isMonthly){
    // 标题中显示所选车企名
    const oemNames = (s1monthlyOems.length > 0 ? s1monthlyOems : [order[0]||'']).join(' vs ');
    lbl.textContent = '重点车企月度出口量 · ' + oemNames + serSuffix + energySuffix;
  } else {
    lbl.textContent = '重点车企出口量' + serSuffix + energySuffix + '（按2025年降序）';
  }

  const filterBox = document.getElementById('s1oemFilter');
  filterBox.style.display = 'none'; // 不再用旧的筛选框

  const moreBtn = document.getElementById('s1oemMoreBtn');
  const zoomBtn = document.getElementById('s1oemZoomBtn');

  if(isMonthly){
    // 月度独立状态：默认展示第1家车企
    if(!s1monthlyInited || !s1monthlyOems.length){
      s1monthlyOems = order.length > 0 ? [order[0]] : [];
      s1monthlyInited = true;
    }
    // 过滤掉不在当前order中的
    s1monthlyOems = s1monthlyOems.filter(o => order.includes(o));
    if(!s1monthlyOems.length && order.length) s1monthlyOems = [order[0]];
    const showOems = [...s1monthlyOems];

    const keys = MK;
    const catLabels = keys.map(k => ML[k] || k);
    const dataMap = {};
    catLabels.forEach((cl,i)=>{
      const mk = keys[i];
      dataMap[cl] = {};
      showOems.forEach(oem=>{
        const od = data[oem]||{};
        dataMap[cl][oem] = od[mk]||0;
        // 同步月度 yoy（仅 26M1/M2 等有可比基数的月份会有值）
        if(od[mk+'_yoy'] !== undefined) dataMap[cl][oem+'_yoy'] = od[mk+'_yoy'];
      });
    });
    if(showOems.length === 1){
      buildBar('c1oem', catLabels, dataMap, [showOems[0]]);
    } else {
      const oemColors = {};
      showOems.forEach((o,i)=>{ oemColors[o]=CS[i%CS.length]; });
      buildBar('c1oem', catLabels, dataMap, showOems, {keyColors:oemColors});
    }
  } else {
    // 年度模式：中系默认展示上市车企TOP10，其他系别直接取排名前10
    const codedSet = new Set(DATA.codedAll || []);
    const codedInOrder = order.filter(o => codedSet.has(o));
    const top10 = codedInOrder.length >= 10 ? codedInOrder.slice(0,10) : order.slice(0,10);
    const showOems = [...top10];
    s1extraOems.forEach(o=>{ if(order.includes(o) && !showOems.includes(o)) showOems.push(o); });
    buildBar('c1oem', showOems, data, s1years, {keyLabels: YR_LABELS});
  }

  // 查看更多按钮（全部车企分组筛选框，年度/月度独立）
  moreBtn.style.display = '';
  moreBtn.textContent = '更多车企';
  moreBtn.onclick = function(){
    const _order = order;
    const _isMonthly = isMonthly;
    const codedSet = new Set(DATA.codedAll || []);
    const codedInOrder = _order.filter(o => codedSet.has(o));
    const uncodedInOrder = _order.filter(o => !codedSet.has(o));
    // 月度：全部可选，当前选中的checked；年度：前10 disabled
    const currentSelected = _isMonthly ? s1monthlyOems : [..._order.slice(0,10), ...s1extraOems];
    const _codedSet = new Set(DATA.codedAll || []);
    const _codedInOrder = _order.filter(o => _codedSet.has(o));
    const top10Names = _isMonthly ? [] : _codedInOrder.slice(0,10); // 月度无锁定，年度锁定带代码TOP10

    const mask = document.createElement('div'); mask.className='modal-mask';
    let h = '<div class="modal-box" style="max-width:1000px"><button class="modal-close">&times;</button>';
    h += '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">'+lbl.textContent+'</div>';
    h += '<div style="margin-bottom:10px;font-size:12px;color:var(--t2)">'+(_isMonthly ? '勾选车企后点击"确认"展示（可多选）' : '前10车企默认展示；勾选其他车企后点击"确认"追加到图表')+'</div>';
    h += '<div style="margin-bottom:6px;font-size:13px;font-weight:700;color:var(--gold)">重点车企（带股票代码）</div>';
    h += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
    codedInOrder.forEach(oem=>{
      const isTop = top10Names.includes(oem);
      const chk = currentSelected.includes(oem) ? 'checked' : '';
      const dis = isTop ? 'disabled' : '';
      h += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px"><input type="checkbox" value="'+oem+'" '+chk+' '+dis+'> '+oem+'</label>';
    });
    h += '</div>';
    if(uncodedInOrder.length > 0){
      h += '<div style="margin-bottom:6px;font-size:13px;font-weight:700;color:var(--t2)">其他车企</div>';
      h += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
      uncodedInOrder.forEach(oem=>{
        const isTop = top10Names.includes(oem);
        const chk = currentSelected.includes(oem) ? 'checked' : '';
        const dis = isTop ? 'disabled' : '';
        h += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px"><input type="checkbox" value="'+oem+'" '+chk+' '+dis+'> '+oem+'</label>';
      });
      h += '</div>';
    }
    h += '<div style="text-align:center;margin-bottom:14px"><button id="s1oemConfirm" style="background:var(--blue);color:#fff;border:none;padding:8px 32px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">确认</button></div>';
    h += '</div>';
    mask.innerHTML = h;
    document.body.appendChild(mask);
    mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
    mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
    mask.querySelector('#s1oemConfirm').addEventListener('click',()=>{
      const checks = [...mask.querySelectorAll('input[type=checkbox]')];
      const selected = checks.filter(c=>c.checked).map(c=>c.value);
      if(_isMonthly){
        s1monthlyOems = selected;
      } else {
        const top10N = _order.slice(0,10);
        s1extraOems = selected.filter(o => !top10N.includes(o));
      }
      mask.remove();
      renderS1Oem();
    });
  };

  // 放大图按钮
  zoomBtn.style.display = '';
  zoomBtn.onclick = function(){
    // 计算当前展示的车企列表
    const _isM = s1mode === 'monthly';
    let zoomOems;
    if(_isM){
      zoomOems = [...s1monthlyOems];
    } else {
      const _oemInfo = getS1OemData(false);
      const __codedSet = new Set(DATA.codedAll || []);
      const _top10 = _oemInfo.order.filter(o => __codedSet.has(o)).slice(0,10);
      zoomOems = [..._top10];
      s1extraOems.forEach(o=>{ if(_oemInfo.order.includes(o) && !zoomOems.includes(o)) zoomOems.push(o); });
    }
    const _data = getS1OemData(_isM).data;
    const mask = document.createElement('div'); mask.className='modal-mask';
    mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">'+lbl.textContent+'（放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="c1oemZoom" style="height:600px"></div></div>';
    document.body.appendChild(mask);
    mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
    mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
    setTimeout(()=>{
      if(_c['c1oemZoom']) delete _c['c1oemZoom'];
      if(_isM){
        const _catLabels = MK.map(k => ML[k] || k);
        const dm = {};
        _catLabels.forEach((cl,i)=>{
          const mk=MK[i]; dm[cl] = {};
          zoomOems.forEach(o=>{
            const od = _data[o]||{};
            dm[cl][o] = od[mk]||0;
            if(od[mk+'_yoy'] !== undefined) dm[cl][o+'_yoy'] = od[mk+'_yoy'];
          });
        });
        const oc = {}; zoomOems.forEach((o,i)=>{ oc[o]=CS[i%CS.length]; });
        buildBar('c1oemZoom', _catLabels, dm, zoomOems, {keyColors:zoomOems.length>1?oc:MC, hideLegend:zoomOems.length===1});
      } else {
        buildBar('c1oemZoom', zoomOems, _data, s1years, {keyLabels: YR_LABELS});
      }
      addModalExport(mask.querySelector('.modal-box'), 'c1oemZoom', lbl.textContent);
    },100);
  };
}

/* ── 主渲染 ── */
function renderS1(){
  const isMonthly = s1mode === 'monthly';
  document.getElementById('s1yearBox').style.display = isMonthly ? 'none' : '';
  document.querySelectorAll('#sec1 .legend-note').forEach(n=>{
    n.classList.toggle('hide', isMonthly);
  });

  renderS1Region();
  buildRegionFilter();
  renderS1Country();
  renderS1Oem();
}

/* ── 事件绑定 ── */
document.querySelectorAll('#s1timeTabs .tab-btn').forEach(btn=>{
  btn.addEventListener('click', function(){
    document.querySelectorAll('#s1timeTabs .tab-btn').forEach(b=>b.classList.remove('on'));
    this.classList.add('on');
    s1mode = this.dataset.mode;
    renderS1();
  });
});

document.querySelectorAll('#s1yearBox .ychk').forEach(lbl=>{
  lbl.addEventListener('click', function(){
    const cb = this.querySelector('input');
    const next = !cb.checked; cb.checked = next;
    this.classList.toggle('on', next);
    s1years = [...document.querySelectorAll('#s1yearBox input:checked')].map(c=>c.value);
    if(s1years.length) renderS1();
  });
});

// 系别单选
document.querySelectorAll('#s1seriesBox .ychk').forEach(lbl=>{
  lbl.addEventListener('click', function(){
    const rb = this.querySelector('input');
    rb.checked = true;
    document.querySelectorAll('#s1seriesBox .ychk').forEach(l=>l.classList.remove('on'));
    this.classList.add('on');
    s1seriesVal = rb.value;
    renderS1();
  });
});

// 动力类型筛选（单选互斥）
document.querySelectorAll('#s1filter [id^="s1e_"]').forEach(lbl=>{
  lbl.addEventListener('click', function(){
    const cb = this.querySelector('input');
    const val = cb.value;
    document.querySelectorAll('#s1filter [id^="s1e_"]').forEach(l=>{
      l.classList.remove('on');
      l.querySelector('input').checked = false;
    });
    cb.checked = true;
    this.classList.add('on');
    s1energy = val === 'all' ? 'all' : val;
    renderS1();
  });
});

renderS1();

/* ══════════════════════════════════════════
   板块二 · 分车企视角
══════════════════════════════════════════ */
let s2mode = 'annual';
let s2yearSel = 'all'; // 'all' 或 AY 中的某个年份
let s2oems = [DATA.sec2.top15.find(o => o !== '其他') || DATA.sec2.top15[0]];
let s2region = null;
let s2country = null;
let s2modelRegion = '全部区域'; // 车型部分独立的区域筛选
let s2modelEnergy = 'all'; // 车型部分动力类型筛选
let s2regEnergy = 'all'; // 分区域动力类型筛选
let s2brandSel = []; // 品牌筛选（空=全部）

// Derived: years array for chart keys
function getS2Years(){ return s2yearSel==='all' ? AY.slice() : [s2yearSel]; }

// Helper: get display label
function getS2OemLabel(){ return s2oems.length===1 ? s2oems[0] : s2oems.join(' vs '); }
function isS2Multi(){ return s2mode==='annual' && s2yearSel!=='all' && s2oems.length>1; }

// Get single-OEM data (for single-select or monthly)
function getS2SingleData(){
  const isMonthly = s2mode === 'monthly';
  // 品牌筛选：合并选中品牌数据
  if(s2brandSel.length > 0 && !isS2Multi()){
    const bSrc = isMonthly ? DATA.sec2.brandMonthly : DATA.sec2.brandAnnual;
    if(s2brandSel.length === 1){
      return bSrc[s2brandSel[0]] || null;
    }
    // 多品牌合并
    const merged = {byEnergyBig:{}, byEnergySub:{}, byRegion:{}, byRegionEnergy:{}, regionCountry:{}, modelByRegionCountry:{}};
    s2brandSel.forEach(b=>{
      const bd = bSrc[b]; if(!bd) return;
      Object.entries(bd.byEnergyBig||{}).forEach(([e,d])=>{
        if(!merged.byEnergyBig[e]) merged.byEnergyBig[e]={};
        Object.keys(d).forEach(k=>{ merged.byEnergyBig[e][k]=(merged.byEnergyBig[e][k]||0)+(d[k]||0); });
      });
      Object.entries(bd.byEnergySub||{}).forEach(([e,d])=>{
        if(!merged.byEnergySub[e]) merged.byEnergySub[e]={};
        Object.keys(d).forEach(k=>{ merged.byEnergySub[e][k]=(merged.byEnergySub[e][k]||0)+(d[k]||0); });
      });
      Object.entries(bd.byRegion||{}).forEach(([r,d])=>{
        if(!merged.byRegion[r]) merged.byRegion[r]={};
        Object.keys(d).forEach(k=>{ merged.byRegion[r][k]=(merged.byRegion[r][k]||0)+(d[k]||0); });
      });
      Object.entries(bd.byRegionEnergy||{}).forEach(([eng,rd])=>{
        if(!merged.byRegionEnergy[eng]) merged.byRegionEnergy[eng]={};
        Object.entries(rd).forEach(([r,d])=>{
          if(!merged.byRegionEnergy[eng][r]) merged.byRegionEnergy[eng][r]={};
          Object.keys(d).forEach(k=>{ merged.byRegionEnergy[eng][r][k]=(merged.byRegionEnergy[eng][r][k]||0)+(d[k]||0); });
        });
      });
      Object.entries(bd.regionCountry||{}).forEach(([r,rd])=>{
        if(!merged.regionCountry[r]) merged.regionCountry[r]={};
        Object.entries(rd).forEach(([c,d])=>{
          if(!merged.regionCountry[r][c]) merged.regionCountry[r][c]={};
          Object.keys(d).forEach(k=>{ merged.regionCountry[r][c][k]=(merged.regionCountry[r][c][k]||0)+(d[k]||0); });
        });
      });
      Object.entries(bd.modelByRegionCountry||{}).forEach(([r,rd])=>{
        if(!merged.modelByRegionCountry[r]) merged.modelByRegionCountry[r]={};
        Object.entries(rd).forEach(([c,cd])=>{
          if(!merged.modelByRegionCountry[r][c]) merged.modelByRegionCountry[r][c]={};
          Object.entries(cd).forEach(([m,d])=>{
            if(!merged.modelByRegionCountry[r][c][m]) merged.modelByRegionCountry[r][c][m]={};
            Object.keys(d).forEach(k=>{ merged.modelByRegionCountry[r][c][m][k]=(merged.modelByRegionCountry[r][c][m][k]||0)+(d[k]||0); });
          });
        });
      });
    });
    return merged;
  }
  const src = isMonthly ? DATA.sec2.monthly : DATA.sec2.annual;
  return src[s2oems[0]] || null;
}

// Build multi-OEM comparison dataMap: { category: { oem1: val, oem2: val } }
function buildMultiOemMap(getter){
  // getter(oemAnnualData, category) => value
  const map = {};
  return { map, put(cat, oem, val){ if(!map[cat]) map[cat]={}; map[cat][oem]=val; } };
}

/* ── 初始化车企筛选按钮（带代码车企TOP15优先） ── */
(function initS2OemFilter(){
  const btnBox = document.getElementById('s2oemBtns');
  const sel = document.getElementById('s2oemSelect');
  const codedOems = DATA.sec2.codedOems || [];
  const uncodedOems = DATA.sec2.uncodedOems || [];
  const top15 = codedOems.slice(0, 15);

  // TOP15 带代码车企 buttons
  top15.forEach((oem, btnIdx) => {
    const btn = document.createElement('span');
    btn.className = 'ychk' + (s2oems.includes(oem) ? ' on' : '');
    btn.style.cssText = 'cursor:pointer';
    btn.dataset.oem = oem;
    btn.innerHTML = '<span class="ydot" style="width:6px;height:6px;background:' + CS[btnIdx % CS.length] + '"></span>' + oem;
    btn.addEventListener('click', function(){
      if(s2mode === 'annual' && s2yearSel !== 'all'){
        const idx = s2oems.indexOf(oem);
        if(idx>=0){ if(s2oems.length>1) s2oems.splice(idx,1); }
        else s2oems.push(oem);
      } else {
        s2oems = [oem];
      }
      s2brandSel = [];
      sel.value = '';
      refreshS2OemHighlight();
      renderS2();
    });
    btnBox.appendChild(btn);
  });

  // Dropdown: 15名之后的带代码车企 + 所有不带代码车企
  const remaining = [...codedOems.slice(15), ...uncodedOems];
  remaining.forEach(oem => {
    const opt = document.createElement('option');
    opt.value = oem;
    opt.textContent = oem;
    sel.appendChild(opt);
  });
  sel.addEventListener('change', function(){
    if(!this.value) return;
    if(s2mode === 'annual' && s2yearSel !== 'all'){
      if(!s2oems.includes(this.value)) s2oems.push(this.value);
    } else {
      s2oems = [this.value];
    }
    s2brandSel = [];
    refreshS2OemHighlight();
    renderS2();
  });
})();

function refreshS2OemHighlight(){
  document.querySelectorAll('#s2oemBtns .ychk').forEach(btn => {
    btn.classList.toggle('on', s2oems.includes(btn.dataset.oem));
  });
  const sel = document.getElementById('s2oemSelect');
  const codedTop15 = (DATA.sec2.codedOems || []).slice(0, 15);
  if(s2oems.every(o => codedTop15.includes(o))){
    sel.value = '';
  }
  // 显示通过下拉框选择的额外车企（可点击取消）
  let extraBox = document.getElementById('s2oemExtraTags');
  if(!extraBox){
    extraBox = document.createElement('div');
    extraBox.id = 's2oemExtraTags';
    extraBox.style.cssText = 'display:flex;gap:4px;flex-wrap:wrap;margin-top:4px';
    document.getElementById('s2oemBtns').parentElement.appendChild(extraBox);
  }
  extraBox.innerHTML = '';
  const extraOems = s2oems.filter(o => !codedTop15.includes(o));
  extraOems.forEach(oem => {
    const tag = document.createElement('span');
    tag.className = 'ychk on';
    tag.style.cssText = 'cursor:pointer;font-size:11px;padding:2px 8px;background:var(--blue);border-color:var(--blue);color:#fff';
    tag.textContent = oem + ' ✕';
    tag.addEventListener('click', ()=>{
      const idx = s2oems.indexOf(oem);
      if(idx >= 0 && s2oems.length > 1) s2oems.splice(idx, 1);
      sel.value = '';
      refreshS2OemHighlight();
      renderS2();
    });
    extraBox.appendChild(tag);
  });
}

/* ── 构建板块二区域筛选按钮 ── */
function buildS2RegionFilter(){
  const container = document.getElementById('s2regFilter');
  container.innerHTML = '';
  const isMonthly = s2mode === 'monthly';
  const src = isMonthly ? DATA.sec2.monthly : DATA.sec2.annual;
  // Collect all regions with data from any selected OEM
  const regSet = new Set();
  s2oems.forEach(oem=>{ const od=src[oem]; if(od) R_ORD.forEach(r=>{ if((od.byRegion||{})[r]) regSet.add(r); }); });
  const availRegions = R_ORD.filter(r => regSet.has(r));

  // 增加「全部区域」选项
  const allRegions = ['全部区域', ...availRegions];

  if(!s2region || !allRegions.includes(s2region)){
    s2region = allRegions[0] || null;
  }

  allRegions.forEach((r, i) => {
    const btn = document.createElement('span');
    btn.className = 'ychk' + (s2region === r ? ' on' : '');
    btn.style.cssText = 'cursor:pointer';
    const dotColor = r === '全部区域' ? '#C9A84C' : CS[(i-1) % CS.length];
    btn.innerHTML = '<span class="ydot" style="width:6px;height:6px;background:' + dotColor + '"></span>' + r;
    btn.addEventListener('click', () => {
      s2region = (s2region === r) ? null : r;
      buildS2RegionFilter();
      renderS2Country();
    });
    container.appendChild(btn);
  });
}

/* ── 渲染分动力类型 ── */
function renderS2Energy(){
  const isMonthly = s2mode === 'monthly';
  document.getElementById('s2engLbl').textContent = '分动力类型 · ' + getS2OemLabel();

  if(isS2Multi()){
    // 多车企对比：X轴=动力类型，bars=各车企（取最新年份）
    const yr = getS2Years()[getS2Years().length-1];
    const catMap = {'燃油车':'燃油车','新能源':'新能源','其中：纯电动':'纯电动','其中：插电混动':'插电混动','其中：增程式':'增程式'};
    const cats = Object.keys(catMap);
    const dataMap = {};
    cats.forEach(cat=>{
      const rawKey = catMap[cat];
      dataMap[cat] = {};
      s2oems.forEach(oem=>{
        const od = DATA.sec2.annual[oem];
        if(!od){ dataMap[cat][oem]=0; return; }
        const src = (rawKey==='燃油车'||rawKey==='新能源') ? (od.byEnergyBig||{})[rawKey]||{} : (od.byEnergySub||{})[rawKey]||{};
        dataMap[cat][oem] = src[yr]||0;
        dataMap[cat][oem+'_yoy'] = src[yr+'_yoy']??null;
      });
    });
    const oemColors = {};
    s2oems.forEach((o,i)=>{ oemColors[o] = CS[i%CS.length]; });
    buildBar('c2eng', cats, dataMap, s2oems, {keyColors:oemColors});
  } else {
    const oemData = getS2SingleData();
    if(!oemData){ const c=gi('c2eng'); if(c) c.clear(); return; }
    if(isMonthly){
      // 月度：只显示燃油车+新能源大类
      const bigOnly = {};
      bigOnly['燃油车'] = oemData.byEnergyBig['燃油车'] || {};
      bigOnly['新能源'] = oemData.byEnergyBig['新能源'] || {};
      const MOL = {hideLegend:true, showMonthLabel:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};
      buildBar('c2eng', ['燃油车','新能源'], bigOnly, MK, MOL);
      // 查看更多：弹窗展示细分
      const engWrap = document.getElementById('c2eng').parentElement;
      let ebtn = engWrap.querySelector('.show-more-btn');
      if(!ebtn){ ebtn=document.createElement('button'); ebtn.className='show-more-btn'; engWrap.appendChild(ebtn); }
      ebtn.textContent = '查看细分动力类型';
      ebtn.onclick = function(){
        const _od = getS2SingleData(); if(!_od) return;
        const mask = document.createElement('div'); mask.className='modal-mask';
        mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button>'+
          '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">分动力类型（含细分） · '+getS2OemLabel()+'</div>'+
          '<div id="c2engModal" style="height:500px"></div></div>';
        document.body.appendChild(mask);
        mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
        mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
        setTimeout(()=>{
          const mChart = echarts.init(document.getElementById('c2engModal'));
          const allCats = ['燃油车','新能源','其中：纯电动','其中：插电混动','其中：增程式'];
          const merged = {};
          merged['燃油车'] = _od.byEnergyBig['燃油车']||{};
          merged['新能源'] = _od.byEnergyBig['新能源']||{};
          merged['其中：纯电动'] = _od.byEnergySub['纯电动']||{};
          merged['其中：插电混动'] = _od.byEnergySub['插电混动']||{};
          merged['其中：增程式'] = _od.byEnergySub['增程式']||{};
          const n2 = allCats.length * MK.length;
          const mLblFs2 = n2>60?0:n2>36?7:8;
          const series2 = MK.map((k,ki)=>({
            name:ML[k]||k, type:'bar', barGap:'8%', barCategoryGap:'25%',
            itemStyle:{color:MC[k], borderRadius:[3,3,0,0]},
            data:allCats.map(cat=>({
              value:(merged[cat]||{})[k]||0,
              yoy:(merged[cat]||{})[k+'_yoy']??null
            })),
            // 始终 show=true：数值在拥挤时省略，yoy（仅 26M1/M2 等有可比基数）总是显示
            label:{show:true, position:'top',
              formatter(p){
                const yoy = p.data && p.data.yoy;
                const hasYoy = yoy !== null && yoy !== undefined;
                if(!p.value && !hasYoy) return '';
                const parts = [];
                if(mLblFs2 > 0 && p.value) parts.push('{v|'+fmt(p.value)+'}');
                if(hasYoy){
                  const pos = yoy >= 0;
                  parts.push('{'+(pos?'r':'g')+'|'+(pos?'▲':'▼')+Math.abs(yoy).toFixed(1)+'%}');
                }
                return parts.join('\n');
              },
              rich:{
                v:{fontSize:Math.max(mLblFs2,8),color:'#FFFFFF',lineHeight:12},
                r:{fontSize:8,color:'#E85A5A',lineHeight:12},
                g:{fontSize:8,color:'#4CAF82',lineHeight:12}
              }}
          }));
          mChart.setOption({
            backgroundColor:'transparent',
            textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#E8EEF8'},
            tooltip:{...TT,trigger:'axis',axisPointer:{type:'shadow'},
              formatter(params){
                let s='<div style="font-weight:600;margin-bottom:6px;color:#C9A84C">'+params[0].axisValue+'</div>';
                params.forEach(p=>{
                  if(!p.value) return;
                  const yoy = p.data && p.data.yoy;
                  const yoyS = yoy!=null
                    ? '<span style="color:'+(yoy>=0?'#E85A5A':'#4CAF82')+';margin-left:6px">'+(yoy>=0?'▲':'▼')+Math.abs(yoy).toFixed(1)+'%</span>' : '';
                  const c = MC[MK[p.seriesIndex]] || '#4A90E2';
                  s+='<div style="display:flex;align-items:center;gap:6px;margin:3px 0">'+
                    '<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:'+c+'"></span>'+
                    p.seriesName+'：<b>'+p.value.toLocaleString()+'</b>'+yoyS+'</div>';
                });
                return s;
              }
            },
            legend:{show:false},
            grid:{top:20,right:12,bottom:50,left:12,containLabel:true},
            xAxis:{type:'category',data:allCats,axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,interval:0},splitLine:{show:false}},
            yAxis:{type:'value',axisLine:{show:false},axisTick:{show:false},
              axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
              splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
            series:series2
          },true);
          addModalExport(mask.querySelector('.modal-box'), 'c2engModal', '分动力类型（含细分） · '+getS2OemLabel());
        },100);
      };
    } else {
      buildEnergyLayered('c2eng', oemData.byEnergyBig || {}, oemData.byEnergySub || {}, getS2Years());
      // 年度模式移除弹窗按钮
      const engWrap = document.getElementById('c2eng').parentElement;
      const eb = engWrap.querySelector('.show-more-btn'); if(eb) eb.remove();
    }
  }
}

/* ── 渲染动力类型份额堆叠百分比图 ── */
function renderS2Share(){
  const chart = gi('c2share');
  if(!chart) return;
  const isMonthly = s2mode === 'monthly';

  document.getElementById('s2shareLbl').textContent = '动力类型份额占比 · ' + getS2OemLabel();

  if(isS2Multi()){
    // 多车企对比：每个车企一组堆叠柱，X轴=车企名
    const yr = getS2Years()[getS2Years().length-1];
    const etypes = ['纯电动','插电混动','增程式','燃油车'];
    const etColors = {'纯电动':'#4A90E2','插电混动':'#5BC4A0','增程式':'#E87B5A','燃油车':'#9B7FD4'};
    const xLabels = s2oems;
    const totals = {};
    s2oems.forEach(oem=>{
      const od = DATA.sec2.annual[oem]; if(!od){ totals[oem]=0; return; }
      let sum=0;
      etypes.forEach(e=>{
        const src = (e==='燃油车') ? (od.byEnergyBig||{}) : (od.byEnergySub||{});
        const key = (e==='燃油车') ? '燃油车' : e;
        sum += (src[key]||{})[yr]||0;
      });
      totals[oem]=sum;
    });
    const series = etypes.map(e=>({
      name:e, type:'bar', stack:'share', barMaxWidth:48,
      itemStyle:{color:etColors[e]},
      data:s2oems.map(oem=>{
        const od = DATA.sec2.annual[oem]; if(!od) return 0;
        const src = (e==='燃油车') ? (od.byEnergyBig||{}) : (od.byEnergySub||{});
        const key = (e==='燃油车') ? '燃油车' : e;
        const val = (src[key]||{})[yr]||0;
        return totals[oem]>0 ? Math.round(val/totals[oem]*1000)/10 : 0;
      }),
      label:{show:true,position:'inside',fontSize:10,color:'#fff',formatter:p=>p.value>5?p.value.toFixed(1)+'%':''}
    }));
    chart.setOption({
      backgroundColor:'transparent',
      textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#8A9DC0'},
      tooltip:{...TT,trigger:'axis',axisPointer:{type:'shadow'}},
      legend:{data:etypes,top:4,textStyle:{color:'#E8EEF8',fontSize:12},itemWidth:12,itemHeight:8},
      grid:{top:42,right:12,bottom:36,left:12,containLabel:true},
      xAxis:{type:'category',data:xLabels,axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
        axisLabel:{color:'#E8EEF8',fontSize:11,interval:0,rotate:xLabels.length>6?30:0},splitLine:{show:false}},
      yAxis:{type:'value',max:100,axisLine:{show:false},axisTick:{show:false},
        axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>v+'%'},
        splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
      series
    },true);
    return;
  }

  const oemData = getS2SingleData();
  if(!oemData){ chart.clear(); return; }

  const etypes = ['纯电动','插电混动','增程式','燃油车'];
  const etColors = {'纯电动':'#4A90E2','插电混动':'#5BC4A0','增程式':'#E87B5A','燃油车':'#9B7FD4'};
  const keys = isMonthly ? MK : getS2Years();
  const xLabels = isMonthly ? MK.map(k=>{ const [y,m]=k.split('-'); return m.padStart(2,'0')+'-'+y.substring(2); }) : keys.map(k=>(YR_LABELS[k]||k)+'年');

  // Compute totals per key
  const totals = {};
  keys.forEach(k => {
    let sum = 0;
    etypes.forEach(e => {
      const src = (e === '燃油车') ? (oemData.byEnergyBig||{}) : (oemData.byEnergySub||{});
      const key = (e === '燃油车') ? '燃油车' : e;
      sum += (src[key]||{})[k] || 0;
    });
    totals[k] = sum;
  });

  const series = etypes.map(e => ({
    name: e, type: 'bar', stack: 'share', barMaxWidth: 48,
    itemStyle: { color: etColors[e] },
    data: keys.map(k => {
      const src = (e === '燃油车') ? (oemData.byEnergyBig||{}) : (oemData.byEnergySub||{});
      const key = (e === '燃油车') ? '燃油车' : e;
      const val = (src[key]||{})[k] || 0;
      const pct = totals[k] > 0 ? Math.round(val / totals[k] * 1000) / 10 : 0;
      return pct;
    }),
    label: {
      show: true, position: 'inside', fontSize: 10, color: '#fff',
      formatter: p => p.value > 5 ? p.value.toFixed(1)+'%' : ''
    }
  }));

  chart.setOption({
    backgroundColor: 'transparent',
    textStyle: {fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif', color:'#8A9DC0'},
    tooltip: {
      ...TT, trigger: 'axis', axisPointer:{type:'shadow'},
      formatter(params){
        let s = '<div style="font-weight:600;margin-bottom:6px;color:#C9A84C">' + params[0].axisValue + '</div>';
        params.forEach(p => {
          if(!p.value) return;
          s += '<div style="display:flex;align-items:center;gap:6px;margin:3px 0">' +
            '<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:'+etColors[p.seriesName]+'"></span>' +
            p.seriesName + '：<b>' + p.value.toFixed(1) + '%</b></div>';
        });
        return s;
      }
    },
    legend: {data:etypes, top:4, textStyle:{color:'#E8EEF8',fontSize:12}, itemWidth:12, itemHeight:8},
    grid: {top:42, right:12, bottom:36, left:12, containLabel:true},
    xAxis: {type:'category', data:xLabels,
      axisLine:{lineStyle:{color:'#1e3a6e'}}, axisTick:{show:false},
      axisLabel:{color:'#E8EEF8',fontSize:12,interval:0}, splitLine:{show:false}},
    yAxis: {type:'value', max:100,
      axisLine:{show:false}, axisTick:{show:false},
      axisLabel:{color:'#E8EEF8',fontSize:11, formatter:v=>v+'%'},
      splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
    series
  }, true);
}

/* ── 渲染分区域 ── */
/* ── 分区域动力类型筛选器 ── */
function buildS2RegEnergyFilter(){
  const container = document.getElementById('s2regEnergyFilter');
  if(!container) return;
  container.innerHTML = '';
  const opts = [
    {key:'all', label:'全部'},
    {key:'新能源', label:'新能源'},
    {key:'纯电动', label:'纯电动'},
    {key:'插电混动', label:'插混'},
    {key:'增程式', label:'增程式'},
    {key:'燃油车', label:'燃油车'}
  ];
  opts.forEach(({key, label}) => {
    const btn = document.createElement('span');
    btn.className = 'ychk' + (s2regEnergy === key ? ' on' : '');
    btn.style.cssText = 'font-size:12px;padding:4px 10px;cursor:pointer';
    btn.textContent = label;
    btn.addEventListener('click', () => {
      s2regEnergy = key;
      buildS2RegEnergyFilter();
      renderS2Region();
    });
    container.appendChild(btn);
  });
}

function renderS2Region(){
  const isMonthly = s2mode === 'monthly';
  const engLabel = s2regEnergy === 'all' ? '' : ' · ' + s2regEnergy;
  document.getElementById('s2regLbl').textContent = '分区域 · ' + getS2OemLabel() + engLabel;

  // 获取区域数据（根据动力类型筛选）
  function getRegData(oemData){
    if(s2regEnergy === 'all') return oemData.byRegion || {};
    return (oemData.byRegionEnergy || {})[s2regEnergy] || {};
  }

  if(isS2Multi()){
    const yr = getS2Years()[getS2Years().length-1];
    const regSet = new Set();
    s2oems.forEach(oem=>{
      const od = DATA.sec2.annual[oem];
      if(od) R_ORD.forEach(r=>{ if((getRegData(od))[r]) regSet.add(r); });
    });
    const regions = R_ORD.filter(r=>regSet.has(r));
    const dataMap = {};
    regions.forEach(r=>{
      dataMap[r] = {};
      s2oems.forEach(oem=>{
        const od = DATA.sec2.annual[oem];
        const regObj = od ? (getRegData(od)[r]||{}) : {};
        dataMap[r][oem] = regObj[yr]||0;
        dataMap[r][oem+'_yoy'] = regObj[yr+'_yoy']??null;
      });
    });
    const oemColors = {};
    s2oems.forEach((o,i)=>{ oemColors[o]=CS[i%CS.length]; });
    buildBar('c2reg', regions, dataMap, s2oems, {keyColors:oemColors});
  } else {
    const oemData = getS2SingleData();
    if(!oemData){ const c=gi('c2reg'); if(c) c.clear(); return; }
    const regData = getRegData(oemData);
    const regions = R_ORD.filter(r => regData[r]);
    if(isMonthly){
      const MOL = {hideLegend:true, showMonthLabel:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};
      buildBar('c2reg', regions, regData, MK, MOL);
    } else {
      buildBar('c2reg', regions, regData, getS2Years(), {keyLabels: YR_LABELS});
    }
  }
}

/* ── 渲染国家明细（默认前5+弹窗前15+其他）── */
function renderS2Country(){
  const drillLbl = document.getElementById('s2drillLbl');
  if(!s2region){
    drillLbl.textContent = '国家明细 · 选择区域查看';
    const container = document.getElementById('c2countryContainer');
    container.querySelectorAll('[id^="c2ctry_"]').forEach(el=>{
      const inst = echarts.getInstanceByDom(el); if(inst) inst.dispose();
      if(el.id && _c[el.id]) delete _c[el.id];
    });
    container.innerHTML = '<div id="c2country" style="height:340px"></div>';
    if(_c['c2country']) delete _c['c2country'];
    return;
  }
  const isMonthly = s2mode === 'monthly';
  drillLbl.textContent = '国家明细 · ' + getS2OemLabel() + ' · ' + s2region;

  if(isS2Multi()){
    // ── 多车企：分图展示，每个OEM各自TOP5国家+其他 ──
    const yr = getS2Years()[getS2Years().length-1];
    const container = document.getElementById('c2countryContainer');
    // 清理旧图表
    container.querySelectorAll('[id^="c2ctry_"]').forEach(el=>{
      const inst = echarts.getInstanceByDom(el);
      if(inst) inst.dispose();
      if(el.id && _c[el.id]) delete _c[el.id];
    });
    container.innerHTML = '';

    s2oems.forEach((oem, oi) => {
      const od = DATA.sec2.annual[oem]; if(!od) return;
      const rc = (od.regionCountry||{})[s2region]||{};
      const countries = Object.keys(rc);
      const sorted = countries.sort((a,b) => ((rc[b]||{})[yr]||0) - ((rc[a]||{})[yr]||0));
      if(!sorted.length) return;

      const top5 = sorted.slice(0, 5);
      const rest = sorted.slice(5);
      const chartData = {};
      top5.forEach(c=>{ chartData[c] = rc[c]; });
      if(rest.length > 0){
        const otherD = {};
        AY.forEach(y=>{ otherD[y] = 0; });
        rest.forEach(c=>{ AY.forEach(y=>{ otherD[y] += ((rc[c]||{})[y]||0); }); });
        chartData['其他'] = otherD;
      }
      const cats = [...top5, ...(chartData['其他'] ? ['其他'] : [])];

      const subWrap = document.createElement('div');
      subWrap.style.cssText = 'background:var(--bg);border:1px solid var(--bd);border-radius:8px;padding:10px;margin-bottom:10px';
      const subTitle = document.createElement('div');
      subTitle.className = 'cw-lbl';
      subTitle.style.cssText = 'font-size:14px;margin-bottom:4px';
      subTitle.textContent = oem + ' · TOP5国家+其他';
      subWrap.appendChild(subTitle);
      const chartDiv = document.createElement('div');
      chartDiv.id = 'c2ctry_' + oi;
      chartDiv.style.height = '260px';
      subWrap.appendChild(chartDiv);
      container.appendChild(subWrap);

      setTimeout(()=>{
        buildBar('c2ctry_'+oi, cats, chartData, [yr]);
      }, 50);
    });
  } else {
    // 单车企：恢复c2country容器
    const container = document.getElementById('c2countryContainer');
    container.querySelectorAll('[id^="c2ctry_"]').forEach(el=>{
      const inst = echarts.getInstanceByDom(el);
      if(inst) inst.dispose();
      if(el.id && _c[el.id]) delete _c[el.id];
    });
    if(!document.getElementById('c2country')){
      container.innerHTML = '<div id="c2country" style="height:340px"></div>';
      if(_c['c2country']) delete _c['c2country'];
    }
    const oemData = getS2SingleData();
    if(!oemData){ const c=gi('c2country'); if(c) c.clear(); return; }
    const countryData = (oemData.regionCountry || {})[s2region] || {};
    const countries = Object.keys(countryData);
    let sorted;
    if(isMonthly){
      sorted = countries.sort((a,b) => {
        const sumA = MK.reduce((s,m) => s + (countryData[a][m]||0), 0);
        const sumB = MK.reduce((s,m) => s + (countryData[b][m]||0), 0);
        return sumB - sumA;
      });
    } else {
      const latestYr = getS2Years()[getS2Years().length - 1];
      sorted = countries.sort((a,b) => (countryData[b][latestYr]||0) - (countryData[a][latestYr]||0));
    }

    const top = sorted.slice(0, 5);
    if(isMonthly){
      const MOL = {hideLegend:true, showMonthLabel:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};
      buildBar('c2country', top, countryData, MK, MOL);
    } else {
      buildBar('c2country', top, countryData, getS2Years(), {keyLabels: YR_LABELS});
    }

    // 查看更多按钮 → 弹窗前15+其他
    const wrap = document.getElementById('c2country').parentElement;
    let btn = wrap.querySelector('.show-more-btn');
    if(sorted.length > 5){
      if(!btn){ btn=document.createElement('button'); btn.className='show-more-btn'; wrap.appendChild(btn); }
      btn.textContent = '查看更多（共'+sorted.length+'国）';
      btn.onclick = function(){
        const _isMonthly = s2mode === 'monthly';
        const keys = _isMonthly ? MK : getS2Years();
        const {cats, data} = mergeOthers(sorted, countryData, keys, 15);
        const mask = document.createElement('div');
        mask.className = 'modal-mask';
        mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button>'+
          '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">国家明细 · '+getS2OemLabel()+' · '+s2region+'（前15+其他）</div>'+
          '<div id="c2countryModal" style="height:500px"></div></div>';
        document.body.appendChild(mask);
        mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
        mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
        setTimeout(()=>{
          const mChart = echarts.init(document.getElementById('c2countryModal'));
          if(_isMonthly){
            const n2 = cats.length * MK.length;
            const mLblFs2 = n2>60?0:n2>36?7:8;
            const series2 = MK.map((k,ki)=>({
              name:ML[k]||k, type:'bar', barGap:'8%', barCategoryGap:'25%',
              itemStyle:{color:MC[k], borderRadius:[3,3,0,0]},
              data:cats.map(cat=>({value:(data[cat]||{})[k]||0})),
              label:{show:mLblFs2>0, position:'top',
                formatter:p=>p.value?'{v|'+fmt(p.value)+'}':'',
                rich:{v:{fontSize:mLblFs2,color:'#FFFFFF',lineHeight:14}}}
            }));
            mChart.setOption({
              backgroundColor:'transparent',
              textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#E8EEF8'},
              tooltip:{...TT,trigger:'axis',axisPointer:{type:'shadow'}},
              legend:{show:false},
              grid:{top:20,right:12,bottom:50,left:12,containLabel:true},
              xAxis:{type:'category',data:cats,axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
                axisLabel:{color:'#E8EEF8',fontSize:11,interval:0,rotate:cats.length>10?-25:0},splitLine:{show:false}},
              yAxis:{type:'value',axisLine:{show:false},axisTick:{show:false},
                axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
                splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
              series:series2
            },true);
          } else {
            const yrs = getS2Years();
            const series2 = yrs.map((yr,yi)=>({
              name:yr+'年', type:'bar', barGap:'4%', barCategoryGap:'30%',
              itemStyle:{color:YC[yr]||CS[yi%CS.length], borderRadius:[3,3,0,0]},
              data:cats.map(cat=>{
                const d=data[cat]||{};
                return {value:d[yr]||0, yoy:d[yr+'_yoy']??null};
              }),
              label:{show:true, position:'top',
                formatter(p){
                  const v=fmt(p.value);
                  const yoy=p.data&&p.data.yoy;
                  if(yoy!==null&&yoy!==undefined){
                    const pos=yoy>=0;
                    return '{v|'+v+'}\n{'+(pos?'r':'g')+'|'+(pos?'▲':'▼')+Math.abs(yoy).toFixed(1)+'%}';
                  }
                  return '{v|'+v+'}';
                },
                rich:{v:{fontSize:9,color:'#FFFFFF',lineHeight:16},r:{fontSize:8,color:'#E85A5A',lineHeight:14},g:{fontSize:8,color:'#4CAF82',lineHeight:14}}}
            }));
            mChart.setOption({
              backgroundColor:'transparent',
              textStyle:{fontFamily:'"PingFang SC","Microsoft YaHei",sans-serif',color:'#E8EEF8'},
              tooltip:{...TT,trigger:'axis',axisPointer:{type:'shadow'}},
              legend:{data:yrs.map(y=>y+'年'),top:4,textStyle:{color:'#E8EEF8',fontSize:12},itemWidth:12,itemHeight:8},
              grid:{top:42,right:12,bottom:50,left:12,containLabel:true},
              xAxis:{type:'category',data:cats,axisLine:{lineStyle:{color:'#1e3a6e'}},axisTick:{show:false},
                axisLabel:{color:'#E8EEF8',fontSize:11,interval:0,rotate:cats.length>10?-25:0},splitLine:{show:false}},
              yAxis:{type:'value',axisLine:{show:false},axisTick:{show:false},
                axisLabel:{color:'#E8EEF8',fontSize:11,formatter:v=>fmt(v)},
                splitLine:{lineStyle:{color:'#1e3a6e',type:'dashed'}}},
              series:series2
            },true);
          }
          addModalExport(mask.querySelector('.modal-box'), 'c2countryModal', '国家明细 · '+getS2OemLabel()+' · '+s2region);
        },100);
      };
    } else if(btn){ btn.remove(); }
  }
}

/* ── 车型区域筛选器（独立于国家明细的区域筛选）── */
function buildS2ModelRegionFilter(){
  const container = document.getElementById('s2modelRegionFilter');
  if(!container) return;
  container.innerHTML = '';

  const src = s2mode === 'monthly' ? DATA.sec2.monthly : DATA.sec2.annual;
  const regSet = new Set();
  s2oems.forEach(oem=>{
    const od = src[oem]; if(!od) return;
    R_ORD.forEach(r=>{ if((od.byRegion||{})[r]) regSet.add(r); });
  });
  const availRegions = ['全部区域', ...R_ORD.filter(r => regSet.has(r))];

  if(!availRegions.includes(s2modelRegion)){
    s2modelRegion = '全部区域';
  }

  availRegions.forEach((r, i) => {
    const btn = document.createElement('span');
    btn.className = 'ychk' + (s2modelRegion === r ? ' on' : '');
    btn.style.cssText = 'font-size:12px;padding:4px 10px;cursor:pointer';
    const dotColor = r === '全部区域' ? '#C9A84C' : CS[(i-1) % CS.length];
    btn.innerHTML = '<span class="ydot" style="width:5px;height:5px;background:' + dotColor + '"></span>' + r;
    btn.addEventListener('click', () => {
      s2modelRegion = r;
      s2country = null; // 切换区域时重置国家
      buildS2ModelRegionFilter();
      buildS2CountryFilter();
      renderS2Model();
    });
    container.appendChild(btn);
  });
}

/* ── 国家筛选器（下拉框，与区域联动，按2025总销量排序）── */
/* ── 车型动力类型筛选器 ── */
function buildS2ModelEnergyFilter(){
  const container = document.getElementById('s2modelEnergyFilter');
  if(!container) return;
  container.innerHTML = '';
  const energyOpts = [
    {key:'all', label:'全部'},
    {key:'新能源', label:'新能源'},
    {key:'纯电动', label:'纯电动'},
    {key:'插电混动', label:'插混'},
    {key:'增程式', label:'增程式'},
    {key:'燃油车', label:'燃油车'}
  ];
  energyOpts.forEach(({key, label}) => {
    const btn = document.createElement('span');
    btn.className = 'ychk' + (s2modelEnergy === key ? ' on' : '');
    btn.style.cssText = 'font-size:12px;padding:4px 10px;cursor:pointer';
    btn.textContent = label;
    btn.addEventListener('click', () => {
      s2modelEnergy = key;
      buildS2ModelEnergyFilter();
      renderS2Model();
    });
    container.appendChild(btn);
  });
}

function buildS2CountryFilter(){
  const sel = document.getElementById('s2countrySelect');
  if(!sel) return;
  while(sel.options.length > 1) sel.remove(1);

  const src = s2mode === 'monthly' ? DATA.sec2.monthly : DATA.sec2.annual;
  // 从选中OEM + 车型区域筛选 获取国家列表
  const countrySet = new Map();
  s2oems.forEach(oem=>{
    const od = src[oem]; if(!od) return;
    const rc = (od.regionCountry||{})[s2modelRegion]||{};
    Object.keys(rc).forEach(c=>{
      const sum = (rc[c]||{})['2025']||0;
      countrySet.set(c, (countrySet.get(c)||0) + sum);
    });
  });
  const sortedCountries = [...countrySet.entries()].sort((a,b)=>b[1]-a[1]).map(e=>e[0]);

  // 第一个选项：全部国家（默认）
  sel.options[0].textContent = '全部国家';

  sortedCountries.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c;
    opt.textContent = c;
    sel.appendChild(opt);
  });

  // 不选择时默认为空（=全部国家）
  if(s2country && !sortedCountries.includes(s2country)){
    s2country = null;
  }
  sel.value = s2country || '';

  if(!sel._bound){
    sel.addEventListener('change', function(){
      s2country = this.value || null;
      renderS2Model();
    });
    sel._bound = true;
  }
}

/* ── 获取车型数据（支持区域+国家筛选）── */
function filterModelsByEnergy(modelData, energyFilter){
  if(!energyFilter || energyFilter === 'all') return modelData;
  const ME = DATA.modelEnergy || {};
  const filtered = {};
  Object.keys(modelData).forEach(model => {
    const me = ME[model] || '';
    let match = false;
    if(energyFilter === '新能源'){
      match = (me === '纯电动' || me === '插电混动' || me === '增程式');
    } else {
      match = (me === energyFilter);
    }
    if(match) filtered[model] = modelData[model];
  });
  return filtered;
}

function getModelData(oemData, region, country, energyFilter){
  let raw;
  if(country){
    raw = ((oemData.modelByRegionCountry||{})[region]||{})[country]||{};
  } else {
    // 未指定国家 → 汇总该区域下所有国家的车型数据
    const allCountryModels = (oemData.modelByRegionCountry||{})[region]||{};
    raw = {};
    Object.values(allCountryModels).forEach(countryModels => {
      Object.keys(countryModels).forEach(model => {
        if(!raw[model]) raw[model] = {};
        const md = countryModels[model];
        Object.keys(md).forEach(k => {
          raw[model][k] = (raw[model][k]||0) + (md[k]||0);
        });
      });
    });
  }
  return filterModelsByEnergy(raw, energyFilter);
}

/* ── 渲染 分车型出口量 ── */
function renderS2Model(){
  const modelLbl = document.getElementById('s2modelLbl');
  const modelLegend = document.getElementById('s2modelLegend');
  const container = document.getElementById('c2modelContainer');
  const isMonthly = s2mode === 'monthly';
  const region = s2modelRegion;
  const country = s2country; // null = 全部国家

  // 清理旧图表实例及缓存
  container.querySelectorAll('[id^="c2model_"]').forEach(el=>{
    const inst = echarts.getInstanceByDom(el);
    if(inst) inst.dispose();
    if(el.id && _c[el.id]) delete _c[el.id];
  });
  container.innerHTML = '';

  const regionLabel = region || '全部区域';
  const countryLabel = country || '全部国家';

  modelLegend.classList.toggle('hide', isMonthly);

  if(isS2Multi()){
    // ── 多车企对比：每个OEM单独一个子图，各展示TOP10 ──
    const energyLabel = s2modelEnergy === 'all' ? '全部动力类型' : s2modelEnergy;
    modelLbl.textContent = '分车型出口量 · ' + regionLabel + ' · ' + countryLabel + ' · ' + energyLabel + '（多车企对比）';
    const yr = getS2Years()[getS2Years().length-1];

    s2oems.forEach((oem, oi) => {
      const od = DATA.sec2.annual[oem]; if(!od) return;
      const modelData = getModelData(od, region, country, s2modelEnergy);
      const models = Object.keys(modelData);
      const sorted = models.sort((a,b) => ((modelData[b]||{})[yr]||0) - ((modelData[a]||{})[yr]||0));
      const top10 = sorted.slice(0, 10);
      if(!top10.length) return;

      const subWrap = document.createElement('div');
      subWrap.style.cssText = 'background:var(--bg);border:1px solid var(--bd);border-radius:8px;padding:10px;margin-bottom:10px';
      const subTitle = document.createElement('div');
      subTitle.className = 'cw-lbl';
      subTitle.style.cssText = 'font-size:14px;margin-bottom:4px';
      subTitle.textContent = oem + ' · TOP10车型';
      subWrap.appendChild(subTitle);
      const chartDiv = document.createElement('div');
      chartDiv.id = 'c2model_' + oi;
      chartDiv.style.height = '280px';
      subWrap.appendChild(chartDiv);
      container.appendChild(subWrap);

      setTimeout(()=>{
        const topData = {};
        top10.forEach(m=>{ topData[m] = modelData[m]; });
        buildBar('c2model_'+oi, top10, topData, [yr]);
      }, 50);
    });
  } else {
    // ── 单车企：展示TOP15 ──
    const energyLabel2 = s2modelEnergy === 'all' ? '全部动力类型' : s2modelEnergy;
    modelLbl.textContent = '分车型出口量 · ' + getS2OemLabel() + ' · ' + regionLabel + ' · ' + countryLabel + ' · ' + energyLabel2;
    const oemData = getS2SingleData();
    if(!oemData) return;
    const modelData = getModelData(oemData, region, country, s2modelEnergy);
    const models = Object.keys(modelData);
    let sorted;
    if(isMonthly){
      sorted = models.sort((a,b) => {
        const sumA = MK.reduce((s,m) => s + ((modelData[a]||{})[m]||0), 0);
        const sumB = MK.reduce((s,m) => s + ((modelData[b]||{})[m]||0), 0);
        return sumB - sumA;
      });
    } else {
      const latestYr = getS2Years()[getS2Years().length - 1];
      sorted = models.sort((a,b) => ((modelData[b]||{})[latestYr]||0) - ((modelData[a]||{})[latestYr]||0));
    }

    const top = sorted.slice(0, 15);
    const chartDiv = document.createElement('div');
    chartDiv.id = 'c2model_0';
    chartDiv.style.height = '420px';
    container.appendChild(chartDiv);

    setTimeout(()=>{
      if(isMonthly){
        const MOL = {hideLegend:true, showMonthLabel:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};
        buildBar('c2model_0', top, modelData, MK, MOL);
      } else {
        buildBar('c2model_0', top, modelData, getS2Years(), {keyLabels: YR_LABELS});
      }
    }, 50);

    if(sorted.length > 15){
      const moreBtn = document.createElement('button');
      moreBtn.className = 'show-more-btn';
      moreBtn.textContent = '查看更多（共'+sorted.length+'款车型）';
      moreBtn.onclick = function(){
        const _isMonthly = s2mode === 'monthly';
        const mask = document.createElement('div'); mask.className='modal-mask';
        mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button>'+
          '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">分车型出口量 · '+getS2OemLabel()+' · '+countryLabel+'（全部车型）</div>'+
          '<div id="c2modelModal" style="height:500px"></div></div>';
        document.body.appendChild(mask);
        mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
        mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
        setTimeout(()=>{
          if(_isMonthly){
            const MOL2 = {hideLegend:true, showMonthLabel:true, isMonthly:true, keyLabels:ML, keyColors:MC, showYoy:true};
            buildBar('c2modelModal', sorted, modelData, MK, MOL2);
          } else {
            buildBar('c2modelModal', sorted, modelData, getS2Years(), {keyLabels: YR_LABELS});
          }
          addModalExport(mask.querySelector('.modal-box'), 'c2modelModal', '分车型出口量 · '+getS2OemLabel()+' · '+countryLabel);
        },100);
      };
      container.appendChild(moreBtn);
    }
  }
}

// 品牌筛选构建
function buildS2BrandFilter(){
  const box = document.getElementById('s2brandBox');
  const btnsDiv = document.getElementById('s2brandBtns');
  if(!box || !btnsDiv) return;
  // 多选车企或无品牌数据时隐藏
  if(isS2Multi() || s2oems.length !== 1){
    box.style.display = 'none';
    return;
  }
  const brands = (DATA.oemBrands||{})[s2oems[0]];
  if(!brands || brands.length <= 1){
    box.style.display = 'none';
    return;
  }
  box.style.display = '';
  btnsDiv.innerHTML = '';
  // "全部"按钮
  const allOn = s2brandSel.length === 0;
  const allBtn = document.createElement('span');
  allBtn.className = 'ychk'+(allOn?' on':'');
  allBtn.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
  allBtn.innerHTML = '<span class="ydot" style="background:#C9A84C"></span>全部';
  allBtn.addEventListener('click',()=>{
    s2brandSel = [];
    buildS2BrandFilter();
    renderS2();
  });
  btnsDiv.appendChild(allBtn);
  // 各品牌
  brands.forEach((b,bi)=>{
    const on = s2brandSel.includes(b.name);
    const btn = document.createElement('span');
    btn.className = 'ychk'+(on?' on':'');
    btn.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
    btn.innerHTML = '<span class="ydot" style="background:'+CS[bi%CS.length]+'"></span>'+b.name;
    btn.addEventListener('click',()=>{
      const idx = s2brandSel.indexOf(b.name);
      if(idx>=0){ s2brandSel.splice(idx,1); }
      else { s2brandSel.push(b.name); }
      buildS2BrandFilter();
      renderS2();
    });
    btnsDiv.appendChild(btn);
  });
}

/* ── 主渲染 ── */
function renderS2(){
  const isMonthly = s2mode === 'monthly';
  document.getElementById('s2yearBox').style.display = isMonthly ? 'none' : '';
  document.querySelectorAll('#sec2 .legend-note').forEach(n => {
    n.classList.toggle('hide', isMonthly);
  });

  buildS2BrandFilter();
  renderS2Energy();
  renderS2Share();
  buildS2RegEnergyFilter();
  renderS2Region();
  buildS2RegionFilter();
  renderS2Country();
  buildS2ModelRegionFilter();
  buildS2ModelEnergyFilter();
  buildS2CountryFilter();
  renderS2Model();
}

/* ── 事件绑定 ── */
document.querySelectorAll('#s2timeTabs .tab-btn').forEach(btn => {
  btn.addEventListener('click', function(){
    document.querySelectorAll('#s2timeTabs .tab-btn').forEach(b => b.classList.remove('on'));
    this.classList.add('on');
    s2mode = this.dataset.mode;
    // 切换到月度时，若多选则只保留第一个
    if(s2mode === 'monthly' && s2oems.length > 1){
      s2oems = [s2oems[0]];
      refreshS2OemHighlight();
    }
    renderS2();
  });
});

// 年份单选
document.querySelectorAll('#s2yearBox .ychk').forEach(lbl => {
  lbl.addEventListener('click', function(){
    const rb = this.querySelector('input');
    rb.checked = true;
    document.querySelectorAll('#s2yearBox .ychk').forEach(l=>l.classList.remove('on'));
    this.classList.add('on');
    const prev = s2yearSel;
    s2yearSel = rb.value;
    // 从"全部"切到具体年份，或从具体年份切到"全部"时，重置为单选
    if((prev==='all' && s2yearSel!=='all') || (prev!=='all' && s2yearSel==='all')){
      s2oems = [s2oems[0]];
      refreshS2OemHighlight();
    }
    renderS2();
  });
});

renderS2();

/* ══════════════════════════════════════════
   中汽协看板
══════════════════════════════════════════ */
if(CAAM){
(function initCAAM(){
  const D = CAAM;
  const YRS = D.meta.years;
  const MK = D.meta.monthKeys;
  const ML_CAAM = D.meta.monthLabels;
  const CAAM_S_ORD = ['中系','日系','欧系','美系','韩系'];

  // 年份颜色（2020-2026扩展）
  const YC_CAAM = {'2020':'#6B7DA0','2021':'#8B6BAE','2022':'#E87B5A','2023':'#4A90E2','2024':'#C9A84C','2025':'#5BC4A0','2026':'#E85A5A'};

  // 月度颜色
  const mColors_caam = ['#4A90E2','#5BC4A0','#C9A84C','#E87B5A','#9B7FD4','#5BC4D4',
    '#E8C870','#6BC4A0','#D49B7F','#7FA0D4','#C4A05B','#E85A7B','#A0D47F','#7FD4C4'];
  const MC_CAAM = {};
  MK.forEach((k,i)=> MC_CAAM[k]=mColors_caam[i%mColors_caam.length]);

  document.getElementById('caamMeta').textContent = (D.meta.dataFileDate ? '数据更新：'+D.meta.dataFileDate+' ｜ ' : '') + '截至 '+D.meta.latestYear+'年'+D.meta.latestMonth+'月';

  // 2026年标注为M1-2（在年份显示中）
  const PARTIAL = D.meta.partialYear;
  const PARTIAL_M = D.meta.partialMonths;
  // 年份显示标签：2026 → 2026M1-2
  const YR_LABELS = {};
  YRS.forEach(y=>{ YR_LABELS[y] = (y===PARTIAL && PARTIAL_M && PARTIAL_M.length<12) ? y+'M1-'+PARTIAL_M[PARTIAL_M.length-1] : y; });

  let caamS0mode = 'annual';
  let caamS2mode = 'annual';
  let caamS2oem = D.codedTop15[0] || D.oemAnnual.order[0];
  let caamS2brandSel = []; // 品牌筛选（空=全部）
  let caamExtraOems = [];        // 年度额外车企
  let caamMonthlyExtraOems = []; // 月度额外车企（独立于年度）

  // 构建年度分系别筛选框
  (function(){
    const container = document.getElementById('caamASerFilter');
    CAAM_S_ORD.forEach((s,si)=>{
      const checked = s==='中系' ? 'checked' : '';
      const on = s==='中系' ? ' on' : '';
      const lbl = document.createElement('label');
      lbl.className = 'ychk'+on;
      lbl.innerHTML = '<input type="checkbox" value="'+s+'" '+checked+'><span class="ydot" style="background:'+CS[si%CS.length]+'"></span>'+s;
      lbl.addEventListener('click', function(){
        const cb = this.querySelector('input');
        const next = !cb.checked; cb.checked = next;
        this.classList.toggle('on', next);
        renderCaamA0Ser();
      });
      container.appendChild(lbl);
    });
  })();

  function renderCaamA0Ser(){
    const checked = [...document.querySelectorAll('#caamASerFilter input:checked')].map(c=>c.value);
    if(!checked.length) return;
    const filtered = {};
    checked.forEach(s=>{ if(D.bySeries[s]) filtered[s] = D.bySeries[s]; });
    buildBar('caamC0ser', checked.filter(s=>D.bySeries[s]), filtered, YRS, {keyColors:YC_CAAM, keyLabels:YR_LABELS});
  }

  /* ── 板块0 渲染 ── */
  function renderCaamS0(){
    const sub = document.getElementById('caamS0sub');
    const annualDiv = document.getElementById('caamS0annual');
    const monthlyDiv = document.getElementById('caamS0monthly');

    if(caamS0mode === 'annual'){
      sub.textContent = '各年度总量与同比增速（2020-2026，其中'+YR_LABELS[PARTIAL]+'为同期可比口径）';
      annualDiv.style.display = '';
      monthlyDiv.style.display = 'none';
      setTimeout(()=>{
        buildBar('caamC0tot',['中汽协乘用车总出口'], D.total, YRS, {keyColors:YC_CAAM, keyLabels:YR_LABELS});
        renderCaamA0Ser();
        // 重点车企：默认展示前5
        const showOems = [...D.codedTop15.slice(0,5), ...caamExtraOems];
        buildBar('caamC0oem', showOems, D.oemAnnual.data, YRS, {keyColors:YC_CAAM, keyLabels:YR_LABELS});
        setTimeout(addExportButtons, 300);
      }, 50);
    } else {
      sub.textContent = '2025年1月以来月度出口量及同比增速';
      annualDiv.style.display = 'none';
      monthlyDiv.style.display = '';
      if(!monthlyDiv._built){
        let html = '';
        html += '<div class="g2 mt14">';
        // 总量
        html += '<div class="cw"><div class="cw-lbl">乘用车出口总量</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="caamM0tot" style="height:300px"></div></div>';
        // 分系别（含筛选框）
        html += '<div class="cw"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><div class="cw-lbl" style="margin-bottom:0">分系别</div>';
        html += '<div class="fbar" id="caamMSerFilter" style="margin:0;padding:6px 10px;background:transparent;border:none">';
        CAAM_S_ORD.forEach((s,si)=>{
          const checked = s==='中系' ? 'checked' : '';
          const on = s==='中系' ? ' on' : '';
          html += '<label class="ychk'+on+'"><input type="checkbox" value="'+s+'" '+checked+'><span class="ydot" style="background:'+CS[si%CS.length]+'"></span>'+s+'</label>';
        });
        html += '</div></div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="caamM0ser" style="height:300px"></div><div class="cw-note">可多选系别筛选，默认中系</div></div>';
        html += '</div>';
        // 重点车企月度
        html += '<div class="cw mt14"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap">';
        html += '<div class="cw-lbl" style="margin-bottom:0">重点车企月度出口量（按26年累计降序）</div>';
        html += '<div style="display:flex;gap:6px;align-items:center">';
        html += '<div class="fbar" id="caamMOemFilter" style="margin:0;padding:6px 10px;background:transparent;border:none;flex-wrap:wrap;gap:6px">';
        D.codedTop15.forEach((oem,oi)=>{
          const checked = oi===0 ? 'checked' : '';
          const on = oi===0 ? ' on' : '';
          html += '<label class="ychk'+on+'"><input type="checkbox" value="'+oi+'" '+checked+'><span class="ydot" style="background:'+CS[oi%CS.length]+'"></span>'+oem+'</label>';
        });
        html += '</div>';
        html += '<button class="show-more-btn" id="caamMOemMoreBtn" style="margin:0;padding:3px 12px;font-size:11px">更多车企</button>';
        html += '</div></div>';
        html += '<div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>';
        html += '<div id="caamM0oem" style="height:320px"></div>';
        html += '<div class="cw-note">可多选车企；筛选独立于年度</div></div>';
        monthlyDiv.innerHTML = html;
        // 系别筛选事件
        document.querySelectorAll('#caamMSerFilter .ychk').forEach(lbl=>{
          lbl.addEventListener('click', function(){
            const cb = this.querySelector('input');
            const next = !cb.checked; cb.checked = next;
            this.classList.toggle('on', next);
            renderCaamM0Ser();
          });
        });
        // 车企筛选事件
        document.querySelectorAll('#caamMOemFilter .ychk').forEach(lbl=>{
          lbl.addEventListener('click', function(){
            const cb = this.querySelector('input');
            const next = !cb.checked; cb.checked = next;
            this.classList.toggle('on', next);
            renderCaamM0Oem();
          });
        });
        // 月度"更多车企"按钮
        document.getElementById('caamMOemMoreBtn').addEventListener('click', function(){
          const allOems = D.oemAnnual.order;
          const codedSet = new Set(D.codedAll);
          const coded = allOems.filter(o=>codedSet.has(o));
          const uncoded = allOems.filter(o=>!codedSet.has(o));
          const mask = document.createElement('div'); mask.className='modal-mask';
          let h = '<div class="modal-box" style="max-width:900px"><button class="modal-close">&times;</button>';
          h += '<div class="cw-lbl" style="margin-bottom:12px;font-size:17px">选择展示车企（中汽协·月度）</div>';
          h += '<div style="margin-bottom:6px;font-size:13px;font-weight:700;color:var(--gold)">重点车企（带股票代码）</div>';
          h += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
          coded.forEach(oem=>{
            const chk = caamMonthlyExtraOems.includes(oem)?'checked':'';
            h += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px"><input type="checkbox" value="'+oem+'" '+chk+'> '+oem+'</label>';
          });
          h += '</div><div style="margin-bottom:6px;font-size:13px;font-weight:700;color:var(--t2)">其他车企</div><div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
          uncoded.forEach(oem=>{
            const chk = caamMonthlyExtraOems.includes(oem)?'checked':'';
            h += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px"><input type="checkbox" value="'+oem+'" '+chk+'> '+oem+'</label>';
          });
          h += '</div><div style="text-align:center"><button id="caamMOemConfirm" style="background:var(--blue);color:#fff;border:none;padding:8px 32px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">确认</button></div></div>';
          mask.innerHTML = h;
          document.body.appendChild(mask);
          mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
          mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
          mask.querySelector('#caamMOemConfirm').addEventListener('click',()=>{
            caamMonthlyExtraOems = [...mask.querySelectorAll('input[type=checkbox]')].filter(c=>c.checked).map(c=>c.value);
            mask.remove();
            renderCaamM0Oem();
          });
        });
        monthlyDiv._built = true;
      }
      // 渲染月度
      const MOL_CAAM = {keyLabels:ML_CAAM, keyColors:MC_CAAM, showYoy:true};
      buildBar('caamM0tot',['中汽协乘用车总出口'], D.monthly_total, MK, MOL_CAAM);
      renderCaamM0Ser();
      renderCaamM0Oem();
      // 延迟添加导出按钮（月度面板动态生成后）
      setTimeout(addExportButtons, 300);
    }
  }

  function renderCaamM0Ser(){
    const checked = [...document.querySelectorAll('#caamMSerFilter input:checked')].map(c=>c.value);
    if(!checked.length) return;
    const filtered = {};
    checked.forEach(s=>{ if(D.monthly_bySeries[s]) filtered[s] = D.monthly_bySeries[s]; });
    const cats = checked.filter(s=>D.monthly_bySeries[s]);
    const MOL = {keyLabels:ML_CAAM, keyColors:MC_CAAM, showYoy:true};
    buildBar('caamM0ser', cats, filtered, MK, MOL);
  }

  function renderCaamM0Oem(){
    const checkedIdx = [...document.querySelectorAll('#caamMOemFilter input:checked')].map(c=>parseInt(c.value));
    const selected = checkedIdx.map(i=>D.codedTop15[i]).filter(Boolean);
    // 月度独立的额外车企
    caamMonthlyExtraOems.forEach(oem=>{ if(!selected.includes(oem)) selected.push(oem); });
    if(!selected.length) return;
    if(selected.length === 1){
      const oem = selected[0];
      const od = D.oemMonthly.data[oem]||{};
      const dataMap = {};
      MK.forEach(mk=>{ dataMap[mk] = {[oem]: od[mk]||0, [oem+'_yoy']: od[mk+'_yoy']??null}; });
      buildBar('caamM0oem', MK, dataMap, [oem], {keyLabels:ML_CAAM, keyColors:MC_CAAM, showYoy:true, hideLegend:true});
    } else {
      const dataMap = {};
      MK.forEach(mk=>{
        dataMap[mk] = {};
        selected.forEach(oem=>{
          const od = D.oemMonthly.data[oem]||{};
          dataMap[mk][oem] = od[mk]||0;
          dataMap[mk][oem+'_yoy'] = od[mk+'_yoy']??null;
        });
      });
      const oemColors = {};
      selected.forEach((o,i)=>{ oemColors[o]=CS[i%CS.length]; });
      buildBar('caamM0oem', MK, dataMap, selected, {keyLabels:ML_CAAM, keyColors:oemColors, showYoy:true});
    }
  }

  /* ── 板块0 更多车企+放大图 ── */
  (function(){
    const container = document.getElementById('caamS0oemExtra');
    const moreBtn = document.createElement('button');
    moreBtn.className = 'show-more-btn';
    moreBtn.style.cssText = 'margin:0;padding:3px 12px;font-size:11px';
    moreBtn.textContent = '更多车企';
    moreBtn.addEventListener('click', function(){
      const allOems = D.oemAnnual.order;
      const codedSet = new Set(D.codedAll);
      const coded = allOems.filter(o=>codedSet.has(o));
      const uncoded = allOems.filter(o=>!codedSet.has(o));
      const mask = document.createElement('div'); mask.className='modal-mask';
      let h = '<div class="modal-box" style="max-width:900px"><button class="modal-close">&times;</button>';
      h += '<div class="cw-lbl" style="margin-bottom:12px;font-size:17px">选择展示车企（中汽协）</div>';
      h += '<div style="margin-bottom:6px;font-size:13px;font-weight:700;color:var(--gold)">重点车企（带股票代码）</div>';
      h += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
      const defaultTop5 = D.codedTop15.slice(0,5);
      coded.forEach(oem=>{
        const isDefault = defaultTop5.includes(oem);
        const chk = isDefault||caamExtraOems.includes(oem)?'checked':'';
        const dis = isDefault?'disabled':'';
        h += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px"><input type="checkbox" value="'+oem+'" '+chk+' '+dis+'> '+oem+'</label>';
      });
      h += '</div><div style="margin-bottom:6px;font-size:13px;font-weight:700;color:var(--t2)">其他车企</div><div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
      uncoded.forEach(oem=>{
        const chk = caamExtraOems.includes(oem)?'checked':'';
        h += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px"><input type="checkbox" value="'+oem+'" '+chk+'> '+oem+'</label>';
      });
      h += '</div><div style="text-align:center"><button id="caamOemConfirm" style="background:var(--blue);color:#fff;border:none;padding:8px 32px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">确认</button></div></div>';
      mask.innerHTML = h;
      document.body.appendChild(mask);
      mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
      mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
      mask.querySelector('#caamOemConfirm').addEventListener('click',()=>{
        caamExtraOems = [...mask.querySelectorAll('input[type=checkbox]:not(:disabled)')].filter(c=>c.checked).map(c=>c.value);
        mask.remove();
        renderCaamS0();
      });
    });
    container.appendChild(moreBtn);
    const zoomBtn = document.createElement('button');
    zoomBtn.className = 'show-more-btn';
    zoomBtn.style.cssText = 'margin:0;padding:3px 12px;font-size:11px';
    zoomBtn.textContent = '查看更多 / 放大图';
    zoomBtn.addEventListener('click', function(){
      const mask = document.createElement('div'); mask.className='modal-mask';
      mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">重点车企出口销量（中汽协·前15+其他）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="caamC0oemZoom" style="height:600px"></div></div>';
      document.body.appendChild(mask);
      mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
      mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
      setTimeout(()=>{
        // 清除旧缓存
        if(_c['caamC0oemZoom']) delete _c['caamC0oemZoom'];
        // 前15带代码车企 + 其余合并为"其他"
        const top15 = D.codedTop15.slice(0,15);
        const rest = [...D.codedAll.slice(15), ...D.uncodedOems, ...caamExtraOems.filter(o=>!top15.includes(o))];
        const zoomData = {};
        top15.forEach(o=>{ zoomData[o] = D.oemAnnual.data[o]||{}; });
        // 合并其他
        const otherD = {};
        YRS.forEach(y=>{ otherD[y] = 0; });
        rest.forEach(o=>{ YRS.forEach(y=>{ otherD[y] += ((D.oemAnnual.data[o]||{})[y]||0); }); });
        if(YRS.some(y=>otherD[y]>0)) zoomData['其他'] = otherD;
        const cats = [...top15, ...(zoomData['其他'] ? ['其他'] : [])];
        buildBar('caamC0oemZoom', cats, zoomData, YRS, {keyColors:YC_CAAM, keyLabels:YR_LABELS});
        addModalExport(mask.querySelector('.modal-box'), 'caamC0oemZoom', '重点车企出口销量（中汽协）');
      },100);
    });
    container.appendChild(zoomBtn);
  })();

  /* ── 板块0 Tab切换 ── */
  document.querySelectorAll('#caamS0tabs .tab-btn').forEach(btn=>{
    btn.addEventListener('click', function(){
      document.querySelectorAll('#caamS0tabs .tab-btn').forEach(b=>b.classList.remove('on'));
      this.classList.add('on');
      caamS0mode = this.dataset.mode;
      renderCaamS0();
    });
  });

  /* ── 板块二 分车企-分车型（前5 + 查看更多前15+其他）── */
  // 品牌筛选构建
  function buildCaamS2BrandFilter(){
    const box = document.getElementById('caamS2brandBox');
    const btnsDiv = document.getElementById('caamS2brandBtns');
    if(!box || !btnsDiv) return;
    const brands = (D.oemBrands||{})[caamS2oem];
    if(!brands || brands.length <= 1){
      box.style.display = 'none';
      return;
    }
    box.style.display = '';
    btnsDiv.innerHTML = '';
    const allOn = caamS2brandSel.length === 0;
    const allBtn = document.createElement('span');
    allBtn.className = 'ychk'+(allOn?' on':'');
    allBtn.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
    allBtn.innerHTML = '<span class="ydot" style="background:#C9A84C"></span>全部';
    allBtn.addEventListener('click',()=>{
      caamS2brandSel = [];
      buildCaamS2BrandFilter();
      renderCaamS2();
    });
    btnsDiv.appendChild(allBtn);
    brands.forEach((b,bi)=>{
      const on = caamS2brandSel.includes(b.name);
      const btn = document.createElement('span');
      btn.className = 'ychk'+(on?' on':'');
      btn.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
      btn.innerHTML = '<span class="ydot" style="background:'+CS[bi%CS.length]+'"></span>'+b.name;
      btn.addEventListener('click',()=>{
        const idx = caamS2brandSel.indexOf(b.name);
        if(idx>=0){ caamS2brandSel.splice(idx,1); }
        else { caamS2brandSel.push(b.name); }
        buildCaamS2BrandFilter();
        renderCaamS2();
      });
      btnsDiv.appendChild(btn);
    });
  }

  function getCaamS2ModelData(){
    const isMonthly = caamS2mode === 'monthly';
    if(caamS2brandSel.length > 0){
      const bSrc = isMonthly ? D.brandModelMonthly : D.brandModelAnnual;
      if(caamS2brandSel.length === 1){
        return bSrc[caamS2brandSel[0]] || {};
      }
      // 多品牌合并
      const merged = {};
      caamS2brandSel.forEach(b=>{
        const bd = bSrc[b]; if(!bd) return;
        Object.entries(bd).forEach(([m,d])=>{
          if(!merged[m]) merged[m]={};
          Object.keys(d).forEach(k=>{ merged[m][k]=(merged[m][k]||0)+(d[k]||0); });
        });
      });
      return merged;
    }
    const src = isMonthly ? D.oemModelMonthly : D.oemModelAnnual;
    if(caamS2oem === '全部车企'){
      // 全市场：跨所有车企按车型名累加，重算同比（月度同月、年度 partial 年走 M1-N 同期）
      const merged = {};
      const yearKeys = isMonthly ? MK : YRS;
      Object.values(src).forEach(oemData=>{
        Object.entries(oemData).forEach(([m,d])=>{
          if(!merged[m]) merged[m] = {};
          yearKeys.forEach(k=>{ merged[m][k] = (merged[m][k]||0) + (d[k]||0); });
        });
      });
      if(isMonthly){
        Object.keys(merged).forEach(m=>{
          const d = merged[m];
          MK.forEach(k=>{
            const [y, mo] = k.split('-');
            if(y === PARTIAL){
              const prevKey = (parseInt(y)-1) + '-' + mo;
              const prev = d[prevKey], cur = d[k];
              if(prev > 0) d[k+'_yoy'] = Math.round((cur/prev-1)*1000)/10;
            }
          });
        });
      } else {
        // 年度：partial 年用 monthly 数据按 M1-N 同期口径；其他年直接 cur/prev
        const monthlyMerged = {};
        const mSrc = D.oemModelMonthly || {};
        Object.values(mSrc).forEach(oemData=>{
          Object.entries(oemData).forEach(([m,d])=>{
            if(!monthlyMerged[m]) monthlyMerged[m] = {};
            Object.keys(d).forEach(k=>{
              if(/^\d{4}-\d+$/.test(k)) monthlyMerged[m][k] = (monthlyMerged[m][k]||0) + (d[k]||0);
            });
          });
        });
        Object.keys(merged).forEach(m=>{
          const d = merged[m];
          for(let i=1;i<YRS.length;i++){
            const yr = YRS[i], prevYr = YRS[i-1];
            if(yr === PARTIAL){
              const md = monthlyMerged[m] || {};
              let curSum = 0, prevSum = 0;
              (PARTIAL_M||[]).forEach(mo=>{
                curSum += (md[yr+'-'+mo]||0);
                prevSum += (md[prevYr+'-'+mo]||0);
              });
              if(prevSum > 0) d[yr+'_yoy'] = Math.round((curSum/prevSum-1)*1000)/10;
            } else {
              const prev=d[prevYr], cur=d[yr];
              if(prev>0) d[yr+'_yoy'] = Math.round((cur/prev-1)*1000)/10;
            }
          }
        });
      }
      return merged;
    }
    return src[caamS2oem] || {};
  }

  function renderCaamS2(){
    const lbl = document.getElementById('caamS2modelLbl');
    const legend = document.getElementById('caamS2modelLegend');
    const container = document.getElementById('caamC2modelContainer');
    const isMonthly = caamS2mode === 'monthly';

    container.querySelectorAll('[id^="caamC2m_"]').forEach(el=>{
      const inst = echarts.getInstanceByDom(el);
      if(inst) inst.dispose();
      if(el.id && _c[el.id]) delete _c[el.id];
    });
    container.innerHTML = '';

    legend.classList.toggle('hide', isMonthly);
    buildCaamS2BrandFilter();

    const modelData = getCaamS2ModelData();
    // 过滤掉"其他"，排序后再决定是否合并
    const realModels = Object.keys(modelData).filter(m => m !== '其他');
    const brandSuffix = caamS2brandSel.length > 0 ? ' · ' + caamS2brandSel.join('+') : '';
    if(!realModels.length && !modelData['其他']){
      lbl.textContent = '分车型出口量 · ' + caamS2oem + brandSuffix + '（无数据）';
      return;
    }

    const keys = isMonthly ? MK : YRS;
    const keyOpts = isMonthly ? {keyLabels:ML_CAAM, keyColors:MC_CAAM, showYoy:true} : {keyColors:YC_CAAM, keyLabels:YR_LABELS};

    // 排序：统一按2026年以来累计销量降序
    const partialYr = PARTIAL;
    let sorted;
    if(isMonthly){
      // 月度模式也按2026年各月累计
      const mk26 = MK.filter(mk => mk.startsWith(partialYr+'-'));
      sorted = realModels.sort((a,b)=>{
        const sumA = mk26.reduce((s,m)=>s+((modelData[a]||{})[m]||0),0);
        const sumB = mk26.reduce((s,m)=>s+((modelData[b]||{})[m]||0),0);
        return sumB - sumA;
      });
      lbl.textContent = '分车型出口量 · ' + caamS2oem + brandSuffix + '（2025年1月以来·按'+YR_LABELS[partialYr]+'累计排序）';
    } else {
      sorted = realModels.sort((a,b)=>((modelData[b]||{})[partialYr]||0)-((modelData[a]||{})[partialYr]||0));
      lbl.textContent = '分车型出口量 · ' + caamS2oem + brandSuffix + '（年度·按'+YR_LABELS[partialYr]+'累计排序）';
    }

    // 默认展示前5
    const top5 = sorted.slice(0, 5);
    const chartDiv = document.createElement('div');
    chartDiv.id = 'caamC2m_0';
    chartDiv.style.height = '320px';
    container.appendChild(chartDiv);

    setTimeout(()=>{
      buildBar('caamC2m_0', top5, modelData, keys, keyOpts);
    }, 50);

    // 查看更多 → 弹窗展示前15 + 其他
    if(sorted.length > 5){
      const moreBtn = document.createElement('button');
      moreBtn.className = 'show-more-btn';
      moreBtn.textContent = '查看更多（共'+sorted.length+'款车型）';
      moreBtn.onclick = function(){
        const top15 = sorted.slice(0, 15);
        const rest = sorted.slice(15);
        // 将rest和原始"其他"合并为一个"其他"
        const mergedData = {};
        top15.forEach(m=>{ mergedData[m] = modelData[m]; });
        // 合并其他
        const otherD = {};
        keys.forEach(k=>{ otherD[k] = 0; });
        rest.forEach(m=>{ keys.forEach(k=>{ otherD[k] += ((modelData[m]||{})[k]||0); }); });
        if(modelData['其他']){ keys.forEach(k=>{ otherD[k] += ((modelData['其他']||{})[k]||0); }); }
        if(keys.some(k=>otherD[k]>0)){
          mergedData['其他'] = otherD;
        }
        const cats = [...top15, ...(mergedData['其他'] ? ['其他'] : [])];

        const mask = document.createElement('div'); mask.className='modal-mask';
        mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">分车型出口量 · '+caamS2oem+'（前15+其他）</div><div id="caamC2mModal" style="height:500px"></div></div>';
        document.body.appendChild(mask);
        mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
        mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
        setTimeout(()=>{
          if(_c['caamC2mModal']) delete _c['caamC2mModal'];
          buildBar('caamC2mModal', cats, mergedData, keys, keyOpts);
          addModalExport(mask.querySelector('.modal-box'), 'caamC2mModal', '分车企-分车型 · '+caamS2oem);
        },100);
      };
      container.appendChild(moreBtn);
    }
  }

  /* ── 板块二 车企选择器 ── */
  (function(){
    const btnBox = document.getElementById('caamS2oemBtns');
    const sel = document.getElementById('caamS2oemSelect');
    const coded = D.codedAll.slice(0,15);
    const remaining = [...D.codedAll.slice(15), ...D.uncodedOems];

    // 全部车企按钮（全市场视角）：跨所有车企按车型排序
    const allBtn = document.createElement('span');
    allBtn.className = 'ychk' + (caamS2oem==='全部车企'?' on':'');
    allBtn.style.cssText = 'cursor:pointer';
    allBtn.dataset.oem = '全部车企';
    allBtn.innerHTML = '<span class="ydot" style="width:6px;height:6px;background:#C9A84C"></span>全部车企';
    allBtn.addEventListener('click', function(){
      caamS2oem = '全部车企';
      caamS2brandSel = [];
      sel.value = '';
      document.querySelectorAll('#caamS2oemBtns .ychk').forEach(b=>b.classList.toggle('on',b.dataset.oem==='全部车企'));
      renderCaamS2();
    });
    btnBox.appendChild(allBtn);

    coded.forEach((oem,i)=>{
      const btn = document.createElement('span');
      btn.className = 'ychk' + (oem===caamS2oem?' on':'');
      btn.style.cssText = 'cursor:pointer';
      btn.dataset.oem = oem;
      btn.innerHTML = '<span class="ydot" style="width:6px;height:6px;background:'+CS[i%CS.length]+'"></span>'+oem;
      btn.addEventListener('click', function(){
        caamS2oem = oem;
        caamS2brandSel = [];
        sel.value = '';
        document.querySelectorAll('#caamS2oemBtns .ychk').forEach(b=>b.classList.toggle('on',b.dataset.oem===oem));
        renderCaamS2();
      });
      btnBox.appendChild(btn);
    });

    remaining.forEach(oem=>{
      const opt = document.createElement('option');
      opt.value = oem; opt.textContent = oem;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', function(){
      if(!this.value) return;
      caamS2oem = this.value;
      caamS2brandSel = [];
      document.querySelectorAll('#caamS2oemBtns .ychk').forEach(b=>b.classList.remove('on'));
      renderCaamS2();
    });
  })();

  /* ── 板块二 Tab切换 ── */
  document.querySelectorAll('#caamS2timeTabs .tab-btn').forEach(btn=>{
    btn.addEventListener('click', function(){
      document.querySelectorAll('#caamS2timeTabs .tab-btn').forEach(b=>b.classList.remove('on'));
      this.classList.add('on');
      caamS2mode = this.dataset.mode;
      renderCaamS2();
    });
  });

  /* ── 首次渲染 + Tab切换时重渲染 ── */
  window._caamRerender = function(){
    // 清除所有caam图表缓存，避免display:none时创建的小尺寸实例
    Object.keys(_c).forEach(id=>{ if(id.startsWith('caam')){ try{_c[id].dispose();}catch(e){} delete _c[id]; } });
    renderCaamS0();
    renderCaamS2();
    // 延迟resize确保DOM完全可见
    setTimeout(()=>{
      Object.entries(_c).forEach(([id,chart])=>{ if(id.startsWith('caam')) try{chart.resize();}catch(e){} });
    }, 200);
  };
  renderCaamS0();
  renderCaamS2();
})();
}

/* ══════════════════════════════════════════
   Marklines看板
══════════════════════════════════════════ */
if(MKLS){
(function initMKLS(){
  const D = MKLS;
  const YRS_M = D.meta.years;
  const MK_M = D.meta.monthKeys;
  const ML_MKLS = D.meta.monthLabels;
  const PARTIAL_M = D.meta.partialYear;
  const YC_MKLS = {'2020':'#6B7DA0','2021':'#8B6BAE','2022':'#E87B5A','2023':'#4A90E2','2024':'#C9A84C','2025':'#5BC4A0','2026':'#E85A5A'};
  const YR_LBL_M = {};
  YRS_M.forEach(y=>{ YR_LBL_M[y] = (y===PARTIAL_M) ? y+'M1-'+D.meta.partialMonths[D.meta.partialMonths.length-1] : y; });
  const mColors = ['#4A90E2','#5BC4A0','#C9A84C','#E87B5A','#9B7FD4','#5BC4D4','#E8C870','#6BC4A0','#D49B7F','#7FA0D4','#C4A05B','#E85A7B','#A0D47F','#7FD4C4'];
  const MC_MKLS = {};
  MK_M.forEach((k,i)=> MC_MKLS[k]=mColors[i%mColors.length]);

  document.getElementById('mklsMeta').textContent = (D.meta.dataFileDate ? '数据更新：'+D.meta.dataFileDate+' ｜ ' : '') + '截至 '+D.meta.latestYear+'年'+D.meta.latestMonth+'月（部分国家数据可能滞后）';

  // 动态月份范围提示（随 partialMonths 自动更新）
  (function(){
    const pm = D.meta.partialMonths || [];
    const py = D.meta.partialYear;
    const el = document.getElementById('mklsPartialRange');
    if(el && py && pm.length){
      const latestM = pm[pm.length-1];
      const startM = Math.max(1, latestM-1);
      el.textContent = py+'年'+startM+'-'+latestM+'月';
    }
  })();

  let mklsS0mode = 'annual';
  let mklsA0serSel = ['中系'];  // 修改2: 年度分系别筛选，默认中系
  let mklsA0engCont = '全部区域';  // 修改3: 年度分动力区域筛选
  let mklsA0engCountry = '';  // 修改3: 年度分动力国家筛选
  let mklsA0contSel = ['亚洲'];  // 修改4: 年度分大洲复选，默认亚洲
  let mklsS2mode = 'annual';
  let mklsS2yearSel = 'all';
  let mklsS2oems = [D.sec2.codedOems[0] || ''];
  let mklsS2region = null;
  let mklsS2modelRegion = '全部区域';
  let mklsS2modelEnergy = 'all';
  let mklsS2modelCountry = null;
  let mklsS2brandSel = [];  // 品牌筛选（空=全部）
  let mklsRegion = '全部区域';

  const MKLS_CONT = D.meta.continentOrder;
  const MKLS_SER = D.meta.seriesList;
  const MKLS_ENG_DISPLAY = ['新能源','其中：纯电动','其中：插电混动','常规混动','燃油车'];

  /* ── 板块0 年度: 分系别筛选渲染 ── */
  function buildMklsA0SerFilter(){
    const container = document.getElementById('mklsA0serFilter');
    if(!container) return;
    container.innerHTML = '';
    MKLS_SER.forEach((s,si)=>{
      const on = mklsA0serSel.includes(s);
      const lbl = document.createElement('label');
      lbl.className = 'ychk'+(on?' on':'');
      lbl.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
      lbl.innerHTML = '<input type="checkbox" value="'+s+'" '+(on?'checked':'')+' style="display:none"><span class="ydot" style="background:'+CS[si%CS.length]+'"></span>'+s;
      lbl.addEventListener('click',function(e){
        e.preventDefault();
        const idx = mklsA0serSel.indexOf(s);
        if(idx>=0){ if(mklsA0serSel.length>1) mklsA0serSel.splice(idx,1); }
        else { mklsA0serSel.push(s); }
        buildMklsA0SerFilter();
        renderMklsA0Ser();
      });
      container.appendChild(lbl);
    });
  }
  function renderMklsA0Ser(){
    buildMklsA0SerFilter();
    const selSeries = mklsA0serSel.length>0 ? mklsA0serSel : MKLS_SER;
    const selList = selSeries.filter(s=>D.bySeries[s]);
    buildBar('mklsC0ser', selList, D.bySeries, YRS_M, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
  }

  /* ── 板块0 年度: 分动力类型筛选渲染（区域+国家） ── */
  function buildMklsA0EngFilter(){
    const container = document.getElementById('mklsA0engContFilter');
    if(!container) return;
    container.innerHTML = '';
    ['全部区域',...MKLS_CONT].forEach((r,i)=>{
      const btn = document.createElement('span');
      btn.className = 'ychk'+(mklsA0engCont===r?' on':'');
      btn.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
      const dotColor = r==='全部区域'?'#C9A84C':CS[(i-1)%CS.length];
      btn.innerHTML = '<span class="ydot" style="width:5px;height:5px;background:'+dotColor+'"></span>'+r;
      btn.addEventListener('click',()=>{
        mklsA0engCont = r;
        mklsA0engCountry = '';
        buildMklsA0EngFilter();
        buildMklsA0CountrySelect();
        renderMklsA0Eng();
      });
      container.appendChild(btn);
    });
  }
  function buildMklsA0CountrySelect(){
    const sel = document.getElementById('mklsA0engCountrySel');
    if(!sel) return;
    sel.innerHTML = '<option value="">全部国家</option>';
    const countries = (D.countryListByContinent||{})[mklsA0engCont] || [];
    countries.forEach(c=>{
      const opt = document.createElement('option');
      opt.value = c; opt.textContent = c;
      if(c===mklsA0engCountry) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.onchange = function(){
      mklsA0engCountry = this.value;
      renderMklsA0Eng();
    };
  }
  function renderMklsA0Eng(){
    buildMklsA0EngFilter();
    buildMklsA0CountrySelect();
    const lbl = document.getElementById('mklsA0engLbl');
    let title = '分动力类型实销量';
    if(mklsA0engCont!=='全部区域') title += ' · '+mklsA0engCont;
    if(mklsA0engCountry) title += ' · '+mklsA0engCountry;
    if(lbl) lbl.textContent = title;
    let src;
    if(mklsA0engCountry){
      src = ((D.energyByCountry||{})[mklsA0engCont]||{})[mklsA0engCountry] || {};
    } else {
      src = (D.energyByContinent||{})[mklsA0engCont] || {};
    }
    const engData = {};
    MKLS_ENG_DISPLAY.forEach(label=>{
      const key = label.replace('其中：','');
      engData[label] = src[key] || {};
    });
    buildBar('mklsC0eng', MKLS_ENG_DISPLAY.filter(l=>Object.keys(engData[l]).length>0), engData, YRS_M, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
  }

  /* ── 板块0 年度: 分大洲复选筛选渲染 ── */
  function buildMklsA0ContFilter(){
    const container = document.getElementById('mklsA0contFilter');
    if(!container) return;
    container.innerHTML = '';
    MKLS_CONT.forEach((c,ci)=>{
      const on = mklsA0contSel.includes(c);
      const lbl = document.createElement('label');
      lbl.className = 'ychk'+(on?' on':'');
      lbl.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
      lbl.innerHTML = '<input type="checkbox" value="'+c+'" '+(on?'checked':'')+' style="display:none"><span class="ydot" style="background:'+CS[ci%CS.length]+'"></span>'+c;
      lbl.addEventListener('click',function(e){
        e.preventDefault();
        const idx = mklsA0contSel.indexOf(c);
        if(idx>=0){ if(mklsA0contSel.length>1) mklsA0contSel.splice(idx,1); }
        else { mklsA0contSel.push(c); }
        buildMklsA0ContFilter();
        renderMklsA0Cont();
      });
      container.appendChild(lbl);
    });
  }
  function renderMklsA0Cont(){
    buildMklsA0ContFilter();
    const selConts = mklsA0contSel.length>0 ? mklsA0contSel : MKLS_CONT;
    const contList = selConts.filter(c=>D.byContinent[c]);
    buildBar('mklsC0cont', contList, D.byContinent, YRS_M, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
  }

  /* ── 板块0 渲染 ── */
  function renderMklsS0(){
    const sub = document.getElementById('mklsS0sub');
    const annualDiv = document.getElementById('mklsS0annual');
    const monthlyDiv = document.getElementById('mklsS0monthly');

    if(mklsS0mode === 'annual'){
      sub.textContent = '海外轻型车终端实销（不含中国）· 年度（'+YR_LBL_M[PARTIAL_M]+'为同期可比，部分国家数据缺失）';
      annualDiv.style.display = '';
      monthlyDiv.style.display = 'none';
      setTimeout(()=>{
        buildBar('mklsC0tot',['海外轻型车终端实销'], D.total, YRS_M, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
        renderMklsA0Ser();
        renderMklsA0Eng();
        renderMklsA0Cont();
        renderMklsCountry();
        setTimeout(addExportButtons, 300);
      }, 50);
    } else {
      sub.textContent = '海外轻型车终端实销（不含中国）· 月度（2025年1月以来，部分国家数据缺失）';
      annualDiv.style.display = 'none';
      monthlyDiv.style.display = '';
      if(!monthlyDiv._built){
        let html = '<div class="g2 mt14">';
        html += '<div class="cw"><div class="cw-lbl">海外轻型车终端实销总量</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsM0tot" style="height:300px"></div></div>';
        // 分系别（含筛选+放大）
        html += '<div class="cw"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><div class="cw-lbl" style="margin-bottom:0">分系别实销量</div><div style="display:flex;gap:6px;align-items:center"><div class="fbar" id="mklsM0serFilter" style="margin:0;padding:4px 8px;background:transparent;border:none;gap:4px">';
        MKLS_SER.forEach((s,si)=>{
          const chk = s==='中系'?'checked':'';
          const on = s==='中系'?' on':'';
          html += '<label class="ychk'+on+'" style="font-size:11px;padding:2px 8px"><input type="checkbox" value="'+s+'" '+chk+'><span class="ydot" style="background:'+CS[si%CS.length]+'"></span>'+s+'</label>';
        });
        html += '</div><button class="show-more-btn" id="mklsM0serZoomBtn" style="margin:0;padding:3px 10px;font-size:11px">放大图</button></div></div>';
        html += '<div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsM0ser" style="height:300px"></div><div class="cw-note">可多选系别筛选，默认中系</div></div>';
        html += '</div>';
        html += '<div class="g2 mt14">';
        // 分动力（含筛选+放大）
        html += '<div class="cw"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><div class="cw-lbl" style="margin-bottom:0">分动力类型实销量</div><div style="display:flex;gap:6px;align-items:center"><div class="fbar" id="mklsM0engFilter" style="margin:0;padding:4px 8px;background:transparent;border:none;gap:4px">';
        ['新能源','纯电动','插电混动','常规混动','燃油车'].forEach((e,ei)=>{
          const chk = e==='新能源'?'checked':'';
          const on = e==='新能源'?' on':'';
          html += '<label class="ychk'+on+'" style="font-size:11px;padding:2px 8px"><input type="checkbox" value="'+e+'" '+chk+'>'+e+'</label>';
        });
        html += '</div><button class="show-more-btn" id="mklsM0engZoomBtn" style="margin:0;padding:3px 10px;font-size:11px">放大图</button></div></div>';
        html += '<div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsM0eng" style="height:300px"></div><div class="cw-note">新能源 = 纯电动 + 插电混动；可多选</div></div>';
        // 分大洲（含筛选+放大）
        html += '<div class="cw"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px"><div class="cw-lbl" style="margin-bottom:0">分大洲实销量</div><div style="display:flex;gap:6px;align-items:center"><div class="fbar" id="mklsM0contFilter" style="margin:0;padding:4px 8px;background:transparent;border:none;gap:4px">';
        MKLS_CONT.forEach((c,ci)=>{
          const chk = c==='亚洲'?'checked':'';
          const on = c==='亚洲'?' on':'';
          html += '<label class="ychk'+on+'" style="font-size:11px;padding:2px 6px"><input type="checkbox" value="'+c+'" '+chk+'>'+c+'</label>';
        });
        html += '</div><button class="show-more-btn" id="mklsM0contZoomBtn" style="margin:0;padding:3px 10px;font-size:11px">放大图</button></div></div>';
        html += '<div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsM0cont" style="height:300px"></div><div class="cw-note">欧洲拆分为欧洲（非俄）和俄罗斯；可多选</div></div>';
        html += '</div>';
        // 区域-国家明细（月度）
        html += '<div class="cw mt14"><div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;flex-wrap:wrap;gap:8px">';
        html += '<div class="cw-lbl" style="margin-bottom:0" id="mklsM0countryLbl">国家明细</div>';
        html += '<div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap">';
        html += '<div style="display:flex;align-items:center;gap:4px"><span style="font-size:13px;color:var(--t1);font-weight:600">大洲：</span><div id="mklsM0regFilter" style="display:flex;gap:3px;flex-wrap:wrap">';
        ['全部区域',...MKLS_CONT].forEach((r,i)=>{
          const on = r==='全部区域'?' on':'';
          const dotColor = r==='全部区域'?'#C9A84C':CS[(i-1)%CS.length];
          html += '<span class="ychk'+on+'" data-region="'+r+'" style="font-size:12px;padding:4px 10px;cursor:pointer"><span class="ydot" style="width:5px;height:5px;background:'+dotColor+'"></span>'+r+'</span>';
        });
        html += '</div></div>';
        html += '<div style="display:flex;align-items:center;gap:4px"><span style="font-size:13px;color:var(--t1);font-weight:600">国家：</span>';
        html += '<span id="mklsM0countryLabel" style="font-size:12px;color:var(--t2);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"></span>';
        html += '<button class="show-more-btn" id="mklsM0countrySelectBtn" style="margin:0;padding:3px 12px;font-size:11px">选择国家</button></div>';
        html += '<button class="show-more-btn" id="mklsM0countryZoomBtn" style="margin:0;padding:3px 12px;font-size:11px">放大图</button>';
        html += '</div></div>';
        html += '<div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div>';
        html += '<div id="mklsM0country" style="height:320px"></div>';
        html += '<div class="cw-note">选择大洲联动国家列表；默认展示TOP1国家月度数据</div></div>';

        monthlyDiv.innerHTML = html;
        // 事件绑定：系别筛选
        document.querySelectorAll('#mklsM0serFilter .ychk').forEach(lbl=>{
          lbl.addEventListener('click', function(){ const cb=this.querySelector('input'); const n=!cb.checked; cb.checked=n; this.classList.toggle('on',n); renderMklsM0Ser(); });
        });
        // 动力筛选
        document.querySelectorAll('#mklsM0engFilter .ychk').forEach(lbl=>{
          lbl.addEventListener('click', function(){ const cb=this.querySelector('input'); const n=!cb.checked; cb.checked=n; this.classList.toggle('on',n); renderMklsM0Eng(); });
        });
        // 大洲筛选
        document.querySelectorAll('#mklsM0contFilter .ychk').forEach(lbl=>{
          lbl.addEventListener('click', function(){ const cb=this.querySelector('input'); const n=!cb.checked; cb.checked=n; this.classList.toggle('on',n); renderMklsM0Cont(); });
        });
        // 放大图按钮
        const zoomHandler = (chartId, filterSel, dataGetter, title) => {
          return function(){
            const mask=document.createElement('div'); mask.className='modal-mask';
            mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">'+title+'（放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="'+chartId+'Zoom" style="height:600px"></div></div>';
            document.body.appendChild(mask);
            mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
            mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
            setTimeout(()=>{
              if(_c[chartId+'Zoom']) delete _c[chartId+'Zoom'];
              const {cats, data} = dataGetter();
              buildBar(chartId+'Zoom', cats, data, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:true});
              addModalExport(mask.querySelector('.modal-box'), chartId+'Zoom', title);
            },100);
          };
        };
        document.getElementById('mklsM0serZoomBtn').addEventListener('click', zoomHandler('mklsM0ser','#mklsM0serFilter',()=>{
          const checked=[...document.querySelectorAll('#mklsM0serFilter input:checked')].map(c=>c.value);
          const d={}; checked.forEach(s=>{if(D.monthly_bySeries[s]) d[s]=D.monthly_bySeries[s];}); return {cats:checked.filter(s=>d[s]),data:d};
        },'分系别实销量'));
        document.getElementById('mklsM0engZoomBtn').addEventListener('click', zoomHandler('mklsM0eng','#mklsM0engFilter',()=>{
          const checked=[...document.querySelectorAll('#mklsM0engFilter input:checked')].map(c=>c.value);
          const d={}; checked.forEach(e=>{if(D.monthly_byEnergy[e]) d[e]=D.monthly_byEnergy[e];}); return {cats:checked.filter(e=>d[e]),data:d};
        },'分动力类型实销量'));
        document.getElementById('mklsM0contZoomBtn').addEventListener('click', zoomHandler('mklsM0cont','#mklsM0contFilter',()=>{
          const checked=[...document.querySelectorAll('#mklsM0contFilter input:checked')].map(c=>c.value);
          const d={}; checked.forEach(c=>{if(D.monthly_byContinent[c]) d[c]=D.monthly_byContinent[c];}); return {cats:checked.filter(c=>d[c]),data:d};
        },'分大洲实销量'));
        // 月度区域-国家筛选
        let mklsM0region = '全部区域';
        let mklsM0country = null;
        document.querySelectorAll('#mklsM0regFilter .ychk').forEach(btn=>{
          btn.addEventListener('click', function(){
            mklsM0region = this.dataset.region;
            mklsM0countries = []; // 切换区域重置国家
            document.querySelectorAll('#mklsM0regFilter .ychk').forEach(b=>b.classList.remove('on'));
            this.classList.add('on');
            ensureMklsM0DefaultCountry();
            renderMklsM0Country();
          });
        });
        let mklsM0countries = [];
        window._mklsM0region = ()=> mklsM0region;
        window._mklsM0countries = ()=> mklsM0countries;

        function updateMklsM0CountryLabel(){
          const lbl = document.getElementById('mklsM0countryLabel');
          if(lbl) lbl.textContent = mklsM0countries.length ? mklsM0countries.join(', ') : '未选择';
        }
        function ensureMklsM0DefaultCountry(){
          const src = (D.monthly_countryByRegion||{})[mklsM0region]||{};
          const countries = Object.keys(src);
          const sorted = countries.sort((a,b)=>{
            const sA = MK_M.reduce((s,m)=>s+((src[a]||{})[m]||0),0);
            const sB = MK_M.reduce((s,m)=>s+((src[b]||{})[m]||0),0);
            return sB-sA;
          });
          if(!mklsM0countries.length || !mklsM0countries.some(c=>sorted.includes(c))){
            mklsM0countries = sorted.length > 0 ? [sorted[0]] : [];
          }
          updateMklsM0CountryLabel();
        }
        // "选择国家"弹窗
        document.getElementById('mklsM0countrySelectBtn').addEventListener('click', function(){
          const src = (D.monthly_countryByRegion||{})[mklsM0region]||{};
          const countries = Object.keys(src);
          const sorted = countries.sort((a,b)=>{
            const sA = MK_M.reduce((s,m)=>s+((src[a]||{})[m]||0),0);
            const sB = MK_M.reduce((s,m)=>s+((src[b]||{})[m]||0),0);
            return sB-sA;
          });
          const mask = document.createElement('div'); mask.className='modal-mask';
          let h = '<div class="modal-box" style="max-width:800px"><button class="modal-close">&times;</button>';
          h += '<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">选择国家 · '+mklsM0region+'</div>';
          h += '<div style="margin-bottom:10px;font-size:12px;color:var(--t2)">可多选国家进行对比（按销量降序排列）</div>';
          h += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px;max-height:400px;overflow-y:auto">';
          sorted.forEach(c=>{
            const chk = mklsM0countries.includes(c)?'checked':'';
            h += '<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px"><input type="checkbox" value="'+c+'" '+chk+'> '+c+'</label>';
          });
          h += '</div><div style="text-align:center"><button id="mklsM0countryConfirm" style="background:var(--blue);color:#fff;border:none;padding:8px 32px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">确认</button></div></div>';
          mask.innerHTML = h;
          document.body.appendChild(mask);
          mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
          mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
          mask.querySelector('#mklsM0countryConfirm').addEventListener('click',()=>{
            mklsM0countries = [...mask.querySelectorAll('input[type=checkbox]:checked')].map(c=>c.value);
            if(!mklsM0countries.length && sorted.length) mklsM0countries = [sorted[0]];
            mask.remove();
            updateMklsM0CountryLabel();
            renderMklsM0Country();
          });
        });
        // 放大图
        document.getElementById('mklsM0countryZoomBtn').addEventListener('click', function(){
          const mask=document.createElement('div'); mask.className='modal-mask';
          mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">国家明细（放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsM0countryZoom" style="height:600px"></div></div>';
          document.body.appendChild(mask);
          mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
          mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
          setTimeout(()=>{
            if(_c['mklsM0countryZoom']) delete _c['mklsM0countryZoom'];
            const region = mklsM0region;
            const src = (D.monthly_countryByRegion||{})[region]||{};
            const sel = mklsM0countries;
            const dm = {};
            MK_M.forEach(mk=>{ dm[mk]={}; sel.forEach(c=>{ dm[mk][c]=(src[c]||{})[mk]||0; dm[mk][c+'_yoy']=(src[c]||{})[mk+'_yoy']??null; }); });
            const cc={}; sel.forEach((c,i)=>{cc[c]=CS[i%CS.length];});
            buildBar('mklsM0countryZoom', MK_M, dm, sel, {keyLabels:ML_MKLS, keyColors:cc, showYoy:true});
            addModalExport(mask.querySelector('.modal-box'), 'mklsM0countryZoom', '国家明细');
          },100);
        });
        // init default
        ensureMklsM0DefaultCountry();
        monthlyDiv._built = true;
      }
      const MOL = {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:true};
      buildBar('mklsM0tot',['海外轻型车终端实销'], D.monthly_total, MK_M, MOL);
      renderMklsM0Ser();
      renderMklsM0Eng();
      renderMklsM0Cont();
      renderMklsM0Country();
      setTimeout(addExportButtons, 300);
    }
  }

  function renderMklsM0Ser(){
    const checked=[...document.querySelectorAll('#mklsM0serFilter input:checked')].map(c=>c.value);
    if(!checked.length) return;
    const d={}; checked.forEach(s=>{if(D.monthly_bySeries[s]) d[s]=D.monthly_bySeries[s];});
    buildBar('mklsM0ser', checked.filter(s=>d[s]), d, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:true});
  }
  function renderMklsM0Eng(){
    const checked=[...document.querySelectorAll('#mklsM0engFilter input:checked')].map(c=>c.value);
    if(!checked.length) return;
    const d={}; checked.forEach(e=>{if(D.monthly_byEnergy[e]) d[e]=D.monthly_byEnergy[e];});
    buildBar('mklsM0eng', checked.filter(e=>d[e]), d, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:true});
  }
  function renderMklsM0Cont(){
    const checked=[...document.querySelectorAll('#mklsM0contFilter input:checked')].map(c=>c.value);
    if(!checked.length) return;
    const d={}; checked.forEach(c=>{if(D.monthly_byContinent[c]) d[c]=D.monthly_byContinent[c];});
    buildBar('mklsM0cont', checked.filter(c=>d[c]), d, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:true});
  }

  function renderMklsM0Country(){
    const region = window._mklsM0region ? window._mklsM0region() : '全部区域';
    const lbl = document.getElementById('mklsM0countryLbl');
    if(!lbl) return;
    // 首次调用确保有默认国家
    if(typeof ensureMklsM0DefaultCountry === 'function') ensureMklsM0DefaultCountry();
    const selCountries = window._mklsM0countries ? window._mklsM0countries() : [];
    if(!selCountries.length){ lbl.textContent='国家明细 · '+region+'（无数据）'; return; }
    lbl.textContent = '国家明细 · '+region+' · '+selCountries.join(' vs ');
    const src = (D.monthly_countryByRegion||{})[region]||{};
    // X轴用月份标签显示
    const catLabels = MK_M.map(k=>ML_MKLS[k]||k);
    const dm = {};
    catLabels.forEach((cl,i)=>{
      const mk = MK_M[i];
      dm[cl] = {};
      selCountries.forEach(c=>{
        dm[cl][c] = (src[c]||{})[mk]||0;
        dm[cl][c+'_yoy'] = (src[c]||{})[mk+'_yoy']??null;
      });
    });
    const countryColors = {};
    selCountries.forEach((c,i)=>{ countryColors[c] = CS[i%CS.length]; });
    buildBar('mklsM0country', catLabels, dm, selCountries, {keyColors:countryColors, showYoy:true});
  }

  /* ── 国家明细 ── */
  function renderMklsCountry(){
    const src = D.countryByRegion[mklsRegion]||{};
    const countries = Object.keys(src);
    const sorted = countries.sort((a,b)=>(src[b]['2025']||0)-(src[a]['2025']||0));
    const top5 = sorted.slice(0,5);
    buildBar('mklsC0country', top5, src, YRS_M, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
    // 查看更多
    const wrap = document.getElementById('mklsC0country').parentElement;
    let mb = wrap.querySelector('.show-more-btn');
    if(sorted.length > 5){
      if(!mb){ mb=document.createElement('button'); mb.className='show-more-btn'; wrap.appendChild(mb); }
      mb.textContent='查看更多';
      mb.onclick=function(){
        const mask=document.createElement('div'); mask.className='modal-mask';
        mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">国家明细 · '+mklsRegion+'（前15+其他）</div><div id="mklsC0countryModal" style="height:500px"></div></div>';
        document.body.appendChild(mask);
        mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
        mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
        setTimeout(()=>{
          if(_c['mklsC0countryModal']) delete _c['mklsC0countryModal'];
          const top15=sorted.slice(0,15); const rest=sorted.slice(15);
          const md={}; top15.forEach(c=>{md[c]=src[c];});
          if(rest.length>0){ const od={}; YRS_M.forEach(y=>{od[y]=0;}); rest.forEach(c=>{YRS_M.forEach(y=>{od[y]+=((src[c]||{})[y]||0);});}); md['其他']=od; }
          buildBar('mklsC0countryModal',[...top15,...(md['其他']?['其他']:[])],md,YRS_M,{keyColors:YC_MKLS,keyLabels:YR_LBL_M});
          addModalExport(mask.querySelector('.modal-box'),'mklsC0countryModal','国家明细 · '+mklsRegion);
        },100);
      };
    } else if(mb){ mb.remove(); }
  }

  // 区域筛选按钮（只保留全部区域+大洲）
  (function(){
    const container = document.getElementById('mklsS0regFilter');
    const allRegs = ['全部区域', ...MKLS_CONT.filter(c=>D.countryByRegion[c])];
    allRegs.forEach((r,i)=>{
      const btn = document.createElement('span');
      btn.className = 'ychk' + (mklsRegion===r?' on':'');
      btn.style.cssText = 'font-size:12px;padding:4px 10px;cursor:pointer';
      const dotColor = r==='全部区域'?'#C9A84C':CS[(i-1)%CS.length];
      btn.innerHTML = '<span class="ydot" style="width:5px;height:5px;background:'+dotColor+'"></span>'+r;
      btn.addEventListener('click',()=>{
        mklsRegion = r;
        container.querySelectorAll('.ychk').forEach(b=>b.classList.remove('on'));
        btn.classList.add('on');
        renderMklsCountry();
      });
      container.appendChild(btn);
    });
  })();

  // 分系别年度放大图（联动筛选）
  document.getElementById('mklsC0serZoomBtn').addEventListener('click', function(){
    const selSeries = mklsA0serSel.length>0 ? mklsA0serSel : MKLS_SER;
    const selList = selSeries.filter(s=>D.bySeries[s]);
    const titleSer = selSeries.join('+');
    const mask=document.createElement('div'); mask.className='modal-mask';
    mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">分系别实销量（'+titleSer+'·放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsC0serZoom" style="height:600px"></div></div>';
    document.body.appendChild(mask);
    mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
    mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
    setTimeout(()=>{
      if(_c['mklsC0serZoom']) delete _c['mklsC0serZoom'];
      buildBar('mklsC0serZoom', selList, D.bySeries, YRS_M, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
      addModalExport(mask.querySelector('.modal-box'), 'mklsC0serZoom', '分系别实销量（'+titleSer+'）');
    },100);
  });

  // 分动力年度放大图（联动区域+国家筛选）
  document.getElementById('mklsC0engZoomBtn').addEventListener('click', function(){
    let title = '分动力类型实销量';
    if(mklsA0engCont!=='全部区域') title += ' · '+mklsA0engCont;
    if(mklsA0engCountry) title += ' · '+mklsA0engCountry;
    const mask=document.createElement('div'); mask.className='modal-mask';
    mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">'+title+'（放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsC0engZoom" style="height:600px"></div></div>';
    document.body.appendChild(mask);
    mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
    mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
    setTimeout(()=>{
      if(_c['mklsC0engZoom']) delete _c['mklsC0engZoom'];
      let src;
      if(mklsA0engCountry){ src=((D.energyByCountry||{})[mklsA0engCont]||{})[mklsA0engCountry]||{}; }
      else { src=(D.energyByContinent||{})[mklsA0engCont]||{}; }
      const engData = {};
      MKLS_ENG_DISPLAY.forEach(label=>{ const key=label.replace('其中：',''); engData[label]=src[key]||{}; });
      buildBar('mklsC0engZoom', MKLS_ENG_DISPLAY.filter(l=>Object.keys(engData[l]).length>0), engData, YRS_M, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
      addModalExport(mask.querySelector('.modal-box'), 'mklsC0engZoom', title);
    },100);
  });

  // 分大洲年度放大图（联动复选筛选）
  document.getElementById('mklsC0contZoomBtn').addEventListener('click', function(){
    const selConts = mklsA0contSel.length>0 ? mklsA0contSel : MKLS_CONT;
    const contList = selConts.filter(c=>D.byContinent[c]);
    const mask=document.createElement('div'); mask.className='modal-mask';
    mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">分大洲实销量（放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsC0contZoom" style="height:600px"></div></div>';
    document.body.appendChild(mask);
    mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
    mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
    setTimeout(()=>{
      if(_c['mklsC0contZoom']) delete _c['mklsC0contZoom'];
      buildBar('mklsC0contZoom', contList, D.byContinent, YRS_M, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
      addModalExport(mask.querySelector('.modal-box'), 'mklsC0contZoom', '分大洲实销量');
    },100);
  });

  // Tab切换
  document.querySelectorAll('#mklsS0tabs .tab-btn').forEach(btn=>{
    btn.addEventListener('click', function(){
      document.querySelectorAll('#mklsS0tabs .tab-btn').forEach(b=>b.classList.remove('on'));
      this.classList.add('on');
      mklsS0mode = this.dataset.mode;
      renderMklsS0();
    });
  });

  /* ══════════════════════════════════════════
     Marklines 板块1 · 分系别·分动力类型视角
  ══════════════════════════════════════════ */
  let mklsS1mode = 'annual';
  let mklsS1years = YRS_M.filter(y=>['2023','2024','2025','2026'].includes(y));
  let mklsS1series = 'all';
  let mklsS1energy = 'all';
  let mklsS1region = '全部区域';
  let mklsS1extraOems = [];
  let mklsS1oemCont = '全部区域';
  let mklsS1oemCountry = '';
  let mklsS1regSel = ['亚洲'];  // 月度区域筛选默认亚洲
  let mklsS1countrySel = null;  // 月度国家明细多选；null=走默认 TOP1
  const MKLS_S_ORD = D.meta.seriesList.slice(0,5);
  const MKLS_ENERGIES = ['all','新能源','纯电动','插电混动','常规混动','燃油车'];

  // 构建筛选栏
  (function buildMklsS1Filters(){
    // 年份
    const yBox = document.getElementById('mklsS1yearBox');
    yBox.innerHTML = '<span class="flbl">年份</span>';
    ['2023','2024','2025','2026'].forEach(y=>{
      const lbl = document.createElement('label');
      lbl.className = 'ychk on';
      lbl.innerHTML = '<input type="checkbox" value="'+y+'" checked><span class="ydot" style="background:'+YC_MKLS[y]+'"></span>'+YR_LBL_M[y];
      lbl.addEventListener('click', function(){
        const cb = this.querySelector('input');
        const next = !cb.checked; cb.checked = next;
        this.classList.toggle('on', next);
        mklsS1years = [...document.querySelectorAll('#mklsS1yearBox input:checked')].map(c=>c.value);
        if(!mklsS1years.length){ mklsS1years=['2025']; }
        renderMklsS1();
      });
      yBox.appendChild(lbl);
    });
    // 系别
    const sBox = document.getElementById('mklsS1seriesBox');
    sBox.innerHTML = '<span class="flbl">系别</span>';
    ['all',...MKLS_S_ORD].forEach(s=>{
      const lbl = document.createElement('label');
      lbl.className = 'ychk'+(mklsS1series===s?' on':'');
      lbl.innerHTML = '<input type="radio" name="mklsS1ser" value="'+s+'"><span class="ydot" style="background:'+(s==='all'?'#C9A84C':CS[MKLS_S_ORD.indexOf(s)%CS.length])+'"></span>'+(s==='all'?'全部':s);
      lbl.addEventListener('click', function(){
        this.querySelector('input').checked = true;
        sBox.querySelectorAll('.ychk').forEach(l=>l.classList.remove('on'));
        this.classList.add('on');
        mklsS1series = this.querySelector('input').value;
        mklsS1countrySel = null;  // 切系别重置国家多选
        renderMklsS1();
      });
      sBox.appendChild(lbl);
    });
    // 动力类型
    const eBox = document.getElementById('mklsS1energyBox');
    eBox.innerHTML = '<span class="flbl">动力类型</span>';
    MKLS_ENERGIES.forEach(e=>{
      const lbl = document.createElement('label');
      lbl.className = 'ychk'+(mklsS1energy===e?' on':'');
      const label = e==='all'?'全部':e;
      lbl.innerHTML = '<input type="radio" name="mklsS1eng" value="'+e+'">'+label;
      lbl.addEventListener('click', function(){
        this.querySelector('input').checked = true;
        eBox.querySelectorAll('.ychk').forEach(l=>l.classList.remove('on'));
        this.classList.add('on');
        mklsS1energy = this.querySelector('input').value;
        mklsS1countrySel = null;  // 切动力重置国家多选
        renderMklsS1();
      });
      eBox.appendChild(lbl);
    });
  })();

  function getMklsS1Series(){ return mklsS1series==='all' ? MKLS_S_ORD : [mklsS1series]; }

  function getMklsS1RegionData(isMonthly){
    const allSel = mklsS1series==='all';
    if(allSel){
      if(mklsS1energy==='all'){
        return isMonthly ? D.monthly_byContinent : D.byContinent;
      }
      return (isMonthly ? D.monthly_regionByEnergy : D.regionByEnergy)[mklsS1energy] || {};
    }
    // 按系别筛选
    const series = getMklsS1Series();
    if(mklsS1energy==='all'){
      const src = isMonthly ? D.monthly_s1BySeriesRegion : D.s1BySeriesRegion;
      const merged = {};
      series.forEach(s=>{ Object.entries(src[s]||{}).forEach(([r,d])=>{ if(!merged[r]) merged[r]={}; Object.keys(d).forEach(k=>{ merged[r][k]=(merged[r][k]||0)+(d[k]||0); }); }); });
      return merged;
    }
    const src = isMonthly ? D.monthly_s1BySeriesEnergyRegion : D.s1BySeriesEnergyRegion;
    const merged = {};
    series.forEach(s=>{ const ed=(src[s]||{})[mklsS1energy]||{}; Object.entries(ed).forEach(([r,d])=>{ if(!merged[r]) merged[r]={}; Object.keys(d).forEach(k=>{ merged[r][k]=(merged[r][k]||0)+(d[k]||0); }); }); });
    return merged;
  }

  function getMklsS1CountryData(region, isMonthly){
    const allSel = mklsS1series==='all';
    if(allSel){
      if(mklsS1energy==='all'){
        return (isMonthly ? D.monthly_regionCountry : D.regionCountry)[region] || {};
      }
      const src = isMonthly ? D.monthly_countryByEnergyRegion : D.countryByEnergyRegion;
      return ((src[region]||{})[mklsS1energy]) || {};
    }
    const series = getMklsS1Series();
    const src = mklsS1energy==='all'
      ? (isMonthly ? D.monthly_s1BySeriesRegionCountry : D.s1BySeriesRegionCountry)
      : (isMonthly ? D.monthly_s1BySeriesEnergyRegionCountry : D.s1BySeriesEnergyRegionCountry);
    // 单系别 fast-path：直接返源，保留预计算的 _yoy（部分年同期口径）
    if(series.length === 1){
      if(mklsS1energy==='all') return (src[series[0]]||{})[region] || {};
      return ((src[series[0]]||{})[mklsS1energy]||{})[region] || {};
    }
    // 多系别合并：遍历源的所有键（含 _yoy）做累加
    const merged = {};
    series.forEach(s=>{
      const rd = mklsS1energy==='all' ? ((src[s]||{})[region]||{})
                                      : (((src[s]||{})[mklsS1energy]||{})[region]||{});
      Object.entries(rd).forEach(([c,d])=>{
        if(!merged[c]) merged[c]={};
        Object.keys(d).forEach(k=>{ merged[c][k]=(merged[c][k]||0)+(d[k]||0); });
      });
    });
    return merged;
  }

  function getMklsS1OemData(isMonthly){
    let info;
    const key = mklsS1energy==='all' ? 'all' : mklsS1energy;
    if(mklsS1oemCountry){
      const cSrc = isMonthly ? D.monthly_oemByEnergyCountry : D.oemByEnergyCountry;
      info = ((cSrc||{})[key]||{})[mklsS1oemCont]||{};
      info = info[mklsS1oemCountry] || {order:[], data:{}};
    } else if(mklsS1oemCont !== '全部区域'){
      const cSrc = isMonthly ? D.monthly_oemByEnergyCont : D.oemByEnergyCont;
      info = ((cSrc||{})[key]||{})[mklsS1oemCont] || {order:[], data:{}};
    } else {
      const src = isMonthly ? D.monthly_oemByEnergy : D.oemByEnergy;
      info = src[key] || {order:[], data:{}};
    }
    const oemSer = D.oemSeries || {};
    const series = getMklsS1Series();
    const filtered = info.order.filter(oem=>{
      const s = oemSer[oem];
      return s && series.includes(s);
    });
    return {order: filtered, data: info.data};
  }

  // OEM区域+国家筛选
  function buildMklsS1OemContFilter(){
    const container = document.getElementById('mklsS1oemContFilter');
    if(!container) return;
    container.innerHTML = '';
    ['全部区域',...MKLS_CONT].forEach((r,i)=>{
      const btn = document.createElement('span');
      btn.className = 'ychk'+(mklsS1oemCont===r?' on':'');
      btn.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
      const dotColor = r==='全部区域'?'#C9A84C':CS[(i-1)%CS.length];
      btn.innerHTML = '<span class="ydot" style="width:5px;height:5px;background:'+dotColor+'"></span>'+r;
      btn.addEventListener('click',()=>{
        mklsS1oemCont = r;
        mklsS1oemCountry = '';
        mklsS1extraOems = [];
        buildMklsS1OemContFilter();
        buildMklsS1OemCountrySelect();
        renderMklsS1Oem();
      });
      container.appendChild(btn);
    });
  }
  function buildMklsS1OemCountrySelect(){
    const sel = document.getElementById('mklsS1oemCountrySel');
    if(!sel) return;
    sel.innerHTML = '<option value="">全部国家</option>';
    const countries = (D.countryListByContinent||{})[mklsS1oemCont] || [];
    countries.forEach(c=>{
      const opt = document.createElement('option');
      opt.value = c; opt.textContent = c;
      if(c===mklsS1oemCountry) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.onchange = function(){
      mklsS1oemCountry = this.value;
      mklsS1extraOems = [];
      renderMklsS1Oem();
    };
  }

  // 区域筛选按钮
  function buildMklsS1RegFilter(){
    const container = document.getElementById('mklsS1regFilter');
    container.innerHTML = '';
    ['全部区域',...D.meta.continentOrder].forEach((r,i)=>{
      const btn = document.createElement('span');
      btn.className = 'ychk'+(mklsS1region===r?' on':'');
      btn.style.cssText = 'font-size:12px;padding:4px 10px;cursor:pointer';
      const dotColor = r==='全部区域'?'#C9A84C':CS[(i-1)%CS.length];
      btn.innerHTML = '<span class="ydot" style="width:5px;height:5px;background:'+dotColor+'"></span>'+r;
      btn.addEventListener('click',()=>{
        mklsS1region = (mklsS1region===r)?null:r;
        if(!mklsS1region) mklsS1region='全部区域';
        mklsS1countrySel = null;  // 切换大洲，重置国家多选
        buildMklsS1RegFilter();
        renderMklsS1Country();
      });
      container.appendChild(btn);
    });
  }

  function buildMklsS1RegSelFilter(){
    // 月度时在放大图按钮旁显示区域复选
    const container = document.getElementById('mklsS1regSelFilter');
    if(!container) return;
    const isMonthly = mklsS1mode==='monthly';
    container.style.display = isMonthly ? '' : 'none';
    if(!isMonthly) return;
    container.innerHTML = '';
    MKLS_CONT.forEach((c,ci)=>{
      const on = mklsS1regSel.includes(c);
      const lbl = document.createElement('label');
      lbl.className = 'ychk'+(on?' on':'');
      lbl.style.cssText = 'font-size:11px;padding:2px 6px;cursor:pointer';
      lbl.innerHTML = '<input type="checkbox" value="'+c+'" '+(on?'checked':'')+' style="display:none"><span class="ydot" style="background:'+CS[ci%CS.length]+'"></span>'+c;
      lbl.addEventListener('click',function(e){
        e.preventDefault();
        const idx = mklsS1regSel.indexOf(c);
        if(idx>=0){ if(mklsS1regSel.length>1) mklsS1regSel.splice(idx,1); }
        else { mklsS1regSel.push(c); }
        buildMklsS1RegSelFilter();
        renderMklsS1Region();
      });
      container.appendChild(lbl);
    });
  }
  function renderMklsS1Region(){
    const isMonthly = mklsS1mode==='monthly';
    const energySuffix = mklsS1energy==='all'?'':'（'+mklsS1energy+'）';
    document.getElementById('mklsS1regLbl').textContent = '分区域销量'+energySuffix;
    buildMklsS1RegSelFilter();
    const regData = getMklsS1RegionData(isMonthly);
    const contList = D.meta.continentOrder.filter(c=>regData[c]);
    if(isMonthly){
      const selConts = mklsS1regSel.length>0 ? mklsS1regSel.filter(c=>regData[c]) : contList;
      buildBar('mklsC1reg', selConts, regData, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:true});
    } else {
      buildBar('mklsC1reg', contList, regData, mklsS1years, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
    }
  }

  /* 月度国家多选 picker */
  function buildMklsS1CountryPicker(sorted){
    const box = document.getElementById('mklsS1countryPickerBox');
    const pick = document.getElementById('mklsS1countryPicker');
    if(!box || !pick) return;
    if(mklsS1mode !== 'monthly'){ box.style.display = 'none'; return; }
    box.style.display = 'flex';
    pick.innerHTML = '';
    sorted.forEach((c, i)=>{
      const on = mklsS1countrySel && mklsS1countrySel.indexOf(c) >= 0;
      const lbl = document.createElement('span');
      lbl.className = 'ychk'+(on?' on':'');
      lbl.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer;line-height:1.5';
      lbl.innerHTML = '<span class="ydot" style="width:5px;height:5px;background:'+CS[i%CS.length]+'"></span>'+c;
      lbl.addEventListener('click', function(){
        if(!Array.isArray(mklsS1countrySel)) mklsS1countrySel = [];
        const idx = mklsS1countrySel.indexOf(c);
        if(idx >= 0){
          if(mklsS1countrySel.length > 1) mklsS1countrySel.splice(idx, 1);  // 至少保留一个
        } else {
          mklsS1countrySel.push(c);
        }
        renderMklsS1Country();
      });
      pick.appendChild(lbl);
    });
  }

  function renderMklsS1Country(){
    const drillLbl = document.getElementById('mklsS1drillLbl');
    if(!mklsS1region){ drillLbl.textContent='选择大洲查看'; return; }
    const isMonthly = mklsS1mode==='monthly';
    const energySuffix = mklsS1energy==='all'?'':'（'+mklsS1energy+'）';
    drillLbl.textContent = '国家明细 · '+mklsS1region+energySuffix;
    // 放大图按钮：年月度都可用
    const zoomBtn = document.getElementById('mklsS1countryZoomBtn');
    if(zoomBtn) zoomBtn.style.display = '';
    const countryData = getMklsS1CountryData(mklsS1region, isMonthly);
    const countries = Object.keys(countryData);
    let sorted;
    if(isMonthly){
      sorted = countries.sort((a,b)=>{ const sA=MK_M.reduce((s,m)=>s+((countryData[a]||{})[m]||0),0); const sB=MK_M.reduce((s,m)=>s+((countryData[b]||{})[m]||0),0); return sB-sA; });
    } else {
      const ly = mklsS1years[mklsS1years.length-1];
      sorted = countries.sort((a,b)=>((countryData[b]||{})[ly]||0)-((countryData[a]||{})[ly]||0));
    }
    // 确定主图展示国家
    let mainCountries;
    if(isMonthly){
      // 月度：优先用 mklsS1countrySel（并过滤掉不在当前区域的），否则默认 TOP1
      if(Array.isArray(mklsS1countrySel)){
        mklsS1countrySel = mklsS1countrySel.filter(c => sorted.indexOf(c) >= 0);
        if(!mklsS1countrySel.length) mklsS1countrySel = sorted.slice(0, 1);
      } else {
        mklsS1countrySel = sorted.slice(0, 1);
      }
      mainCountries = mklsS1countrySel.slice();
      buildMklsS1CountryPicker(sorted);
    } else {
      mainCountries = sorted.slice(0, 5);
      const box = document.getElementById('mklsS1countryPickerBox');
      if(box) box.style.display = 'none';
    }
    if(isMonthly){
      buildBar('mklsC1country', mainCountries, countryData, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:true});
    } else {
      buildBar('mklsC1country', mainCountries, countryData, mklsS1years, {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
    }
    // 查看更多（用专属 ID 避免与放大图按钮的 .show-more-btn 冲突）
    const wrap = document.getElementById('mklsS1countryWrap');
    let mb = document.getElementById('mklsS1countryMoreBtn');
    const showMoreThreshold = isMonthly ? 1 : 5;
    if(sorted.length > showMoreThreshold){
      if(!mb){ mb=document.createElement('button'); mb.id='mklsS1countryMoreBtn'; mb.className='show-more-btn'; wrap.appendChild(mb); }
      mb.textContent='查看更多';
      mb.onclick=function(){
        const topN2 = isMonthly ? 10 : 15;
        const noteText = isMonthly ? '前10国家，其余合并为其他' : '前15国家，其余合并为其他';
        const mask=document.createElement('div'); mask.className='modal-mask';
        mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">国家明细 · '+mklsS1region+energySuffix+'</div><div style="font-size:12px;color:var(--t2);margin-bottom:8px">'+noteText+'</div><div id="mklsC1countryModal" style="height:500px"></div></div>';
        document.body.appendChild(mask);
        mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
        mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
        setTimeout(()=>{
          if(_c['mklsC1countryModal']) delete _c['mklsC1countryModal'];
          const topSlice=sorted.slice(0,topN2);
          const rest=sorted.slice(topN2);
          const md={}; topSlice.forEach(c=>{md[c]=countryData[c];});
          if(rest.length>0){ const keys2=isMonthly?MK_M:mklsS1years; const od={}; keys2.forEach(k=>{od[k]=0;}); rest.forEach(c=>{keys2.forEach(k=>{od[k]+=((countryData[c]||{})[k]||0);});}); md['其他']=od; }
          const cats=[...topSlice,...(md['其他']?['其他']:[])];
          if(isMonthly){ buildBar('mklsC1countryModal',cats,md,MK_M,{keyLabels:ML_MKLS,keyColors:MC_MKLS,showYoy:true}); }
          else { buildBar('mklsC1countryModal',cats,md,mklsS1years,{keyColors:YC_MKLS,keyLabels:YR_LBL_M}); }
          addModalExport(mask.querySelector('.modal-box'),'mklsC1countryModal','国家明细 · '+mklsS1region);
        },100);
      };
    } else if(mb){ mb.remove(); }
  }

  function renderMklsS1Oem(){
    buildMklsS1OemContFilter();
    buildMklsS1OemCountrySelect();
    const isMonthly = mklsS1mode==='monthly';
    const oemInfo = getMklsS1OemData(isMonthly);
    const order = oemInfo.order;
    const data = oemInfo.data;
    const energySuffix = mklsS1energy==='all'?'':'（'+mklsS1energy+'）';
    const serSuffix = mklsS1series==='all'?'':'（'+mklsS1series+'）';
    let regSuffix = '';
    if(mklsS1oemCont!=='全部区域') regSuffix += ' · '+mklsS1oemCont;
    if(mklsS1oemCountry) regSuffix += ' · '+mklsS1oemCountry;
    const lbl = document.getElementById('mklsS1oemLbl');
    lbl.textContent = '重点车企终端实销量'+serSuffix+energySuffix+regSuffix+'（按2025年降序）';

    const defaultN = isMonthly ? 5 : 10;
    const topN = order.slice(0,defaultN);
    const showOems = [...topN];
    mklsS1extraOems.forEach(o=>{ if(order.includes(o) && !showOems.includes(o)) showOems.push(o); });

    if(isMonthly){
      // 月度：X轴=车企，月份作为图例
      buildBar('mklsC1oem',showOems,data,MK_M,{keyLabels:ML_MKLS,keyColors:MC_MKLS,showYoy:false});
    } else {
      buildBar('mklsC1oem',showOems,data,mklsS1years,{keyColors:YC_MKLS,keyLabels:YR_LBL_M});
    }

    // 更多车企
    document.getElementById('mklsS1oemMoreBtn').onclick = function(){
      const mask=document.createElement('div'); mask.className='modal-mask';
      let h='<div class="modal-box" style="max-width:1000px"><button class="modal-close">&times;</button>';
      h+='<div class="cw-lbl" style="margin-bottom:8px;font-size:17px">'+lbl.textContent+'</div>';
      h+='<div style="margin-bottom:10px;font-size:12px;color:var(--t2)">前10车企默认展示；勾选其他车企后点击"确认"追加</div>';
      h+='<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px">';
      order.forEach((oem,oi)=>{
        const isTop=oi<10;
        const chk=(isTop||mklsS1extraOems.includes(oem))?'checked':'';
        const dis=isTop?'disabled':'';
        h+='<label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#C9D6EC;cursor:pointer;padding:3px 8px;border:1px solid var(--bd);border-radius:4px"><input type="checkbox" value="'+oem+'" '+chk+' '+dis+'> '+oem+'</label>';
      });
      h+='</div><div style="text-align:center"><button id="mklsS1oemConfirm" style="background:var(--blue);color:#fff;border:none;padding:8px 32px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">确认</button></div></div>';
      mask.innerHTML=h;
      document.body.appendChild(mask);
      mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
      mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
      mask.querySelector('#mklsS1oemConfirm').addEventListener('click',()=>{
        mklsS1extraOems=[...mask.querySelectorAll('input[type=checkbox]:not(:disabled)')].filter(c=>c.checked).map(c=>c.value);
        mask.remove();
        renderMklsS1Oem();
      });
    };
    // 放大图
    document.getElementById('mklsS1oemZoomBtn').onclick = function(){
      const mask=document.createElement('div'); mask.className='modal-mask';
      mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">'+lbl.textContent+'（放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsC1oemZoom" style="height:600px"></div></div>';
      document.body.appendChild(mask);
      mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
      mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
      setTimeout(()=>{
        if(_c['mklsC1oemZoom']) delete _c['mklsC1oemZoom'];
        buildBar('mklsC1oemZoom',showOems,data,mklsS1years,{keyColors:YC_MKLS,keyLabels:YR_LBL_M});
        addModalExport(mask.querySelector('.modal-box'),'mklsC1oemZoom',lbl.textContent);
      },100);
    };
  }

  function renderMklsS1(){
    const isMonthly = mklsS1mode==='monthly';
    document.getElementById('mklsS1yearBox').style.display = isMonthly?'none':'';
    document.querySelectorAll('#mklsSec1 .legend-note').forEach(n=>n.classList.toggle('hide',isMonthly));
    renderMklsS1Region();
    buildMklsS1RegFilter();
    renderMklsS1Country();
    renderMklsS1Oem();
    setTimeout(addExportButtons, 300);
  }

  // 板块1 分大洲放大图
  document.getElementById('mklsS1regZoomBtn').addEventListener('click', function(){
    const mask=document.createElement('div'); mask.className='modal-mask';
    mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">'+document.getElementById('mklsS1regLbl').textContent+'（放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsC1regZoom" style="height:600px"></div></div>';
    document.body.appendChild(mask);
    mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
    mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
    setTimeout(()=>{
      if(_c['mklsC1regZoom']) delete _c['mklsC1regZoom'];
      const isM = mklsS1mode==='monthly';
      const regData = getMklsS1RegionData(isM);
      const contList = MKLS_CONT.filter(c=>regData[c]);
      if(isM){ buildBar('mklsC1regZoom',contList,regData,MK_M,{keyLabels:ML_MKLS,keyColors:MC_MKLS,showYoy:true}); }
      else { buildBar('mklsC1regZoom',contList,regData,mklsS1years,{keyColors:YC_MKLS,keyLabels:YR_LBL_M}); }
      addModalExport(mask.querySelector('.modal-box'),'mklsC1regZoom',document.getElementById('mklsS1regLbl').textContent);
    },100);
  });

  // 板块1 国家明细放大图：主图的放大版本（显示当前选中国家，月度用多选，年度用 Top5）
  document.getElementById('mklsS1countryZoomBtn').addEventListener('click', function(){
    const mask=document.createElement('div'); mask.className='modal-mask';
    mask.innerHTML='<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">'+document.getElementById('mklsS1drillLbl').textContent+'（放大视图）</div><div class="legend-note"><span class="pos">红色▲ 同比增长</span><span class="neg">绿色▼ 同比下降</span></div><div id="mklsC1countryZoom" style="height:600px"></div></div>';
    document.body.appendChild(mask);
    mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
    mask.addEventListener('click',e=>{if(e.target===mask) mask.remove();});
    setTimeout(()=>{
      if(_c['mklsC1countryZoom']) delete _c['mklsC1countryZoom'];
      const isM = mklsS1mode==='monthly';
      const countryData = getMklsS1CountryData(mklsS1region, isM);
      const countries = Object.keys(countryData);
      let sorted;
      if(isM){ sorted=countries.sort((a,b)=>{const sA=MK_M.reduce((s,m)=>s+((countryData[a]||{})[m]||0),0);const sB=MK_M.reduce((s,m)=>s+((countryData[b]||{})[m]||0),0);return sB-sA;}); }
      else { const ly=mklsS1years[mklsS1years.length-1]; sorted=countries.sort((a,b)=>((countryData[b]||{})[ly]||0)-((countryData[a]||{})[ly]||0)); }
      // 与主图保持一致：月度跟随 mklsS1countrySel，年度默认前5
      let disp;
      if(isM){
        disp = Array.isArray(mklsS1countrySel) ? mklsS1countrySel.filter(c=>sorted.indexOf(c)>=0) : [];
        if(!disp.length) disp = sorted.slice(0, 1);
      } else {
        disp = sorted.slice(0, 5);
      }
      if(isM){ buildBar('mklsC1countryZoom',disp,countryData,MK_M,{keyLabels:ML_MKLS,keyColors:MC_MKLS,showYoy:true}); }
      else { buildBar('mklsC1countryZoom',disp,countryData,mklsS1years,{keyColors:YC_MKLS,keyLabels:YR_LBL_M}); }
      addModalExport(mask.querySelector('.modal-box'),'mklsC1countryZoom',document.getElementById('mklsS1drillLbl').textContent);
    },100);
  });

  // Time tab
  document.querySelectorAll('#mklsS1timeTabs .tab-btn').forEach(btn=>{
    btn.addEventListener('click', function(){
      document.querySelectorAll('#mklsS1timeTabs .tab-btn').forEach(b=>b.classList.remove('on'));
      this.classList.add('on');
      mklsS1mode = this.dataset.mode;
      mklsS1countrySel = null;  // 切时间维度重置国家多选
      renderMklsS1();
    });
  });

  renderMklsS1();

  /* ── 板块二 分车企 ── */
  function getMklsS2OemLabel(){ return mklsS2oems.length===1 ? mklsS2oems[0] : mklsS2oems.join(' vs '); }
  function isMklsS2Multi(){ return mklsS2mode==='annual' && mklsS2yearSel!=='all' && mklsS2oems.length>1; }
  function getMklsS2Years(){
    if(mklsS2yearSel==='all') return YRS_M.filter(y=>['2020','2021','2022','2023','2024','2025','2026'].includes(y));
    return [mklsS2yearSel];
  }
  function getMklsS2SingleData(){
    const isMonthly = mklsS2mode==='monthly';
    // 品牌筛选：合并选中品牌数据
    if(mklsS2brandSel.length > 0 && !isMklsS2Multi()){
      const bSrc = isMonthly ? D.sec2.brandMonthly : D.sec2.brandAnnual;
      if(mklsS2brandSel.length === 1){
        return bSrc[mklsS2brandSel[0]] || null;
      }
      // 多品牌合并
      const merged = {byEnergy:{}, byContinent:{}, countryByRegion:{}, models:{}};
      mklsS2brandSel.forEach(b=>{
        const bd = bSrc[b]; if(!bd) return;
        // 合并 byEnergy
        Object.entries(bd.byEnergy||{}).forEach(([e,d])=>{
          if(!merged.byEnergy[e]) merged.byEnergy[e]={};
          Object.keys(d).forEach(k=>{ merged.byEnergy[e][k]=(merged.byEnergy[e][k]||0)+(d[k]||0); });
        });
        // 合并 byContinent
        Object.entries(bd.byContinent||{}).forEach(([c,d])=>{
          if(!merged.byContinent[c]) merged.byContinent[c]={};
          Object.keys(d).forEach(k=>{ merged.byContinent[c][k]=(merged.byContinent[c][k]||0)+(d[k]||0); });
        });
        // 合并 countryByRegion
        Object.entries(bd.countryByRegion||{}).forEach(([r,rd])=>{
          if(!merged.countryByRegion[r]) merged.countryByRegion[r]={};
          Object.entries(rd).forEach(([c,d])=>{
            if(!merged.countryByRegion[r][c]) merged.countryByRegion[r][c]={};
            Object.keys(d).forEach(k=>{ merged.countryByRegion[r][c][k]=(merged.countryByRegion[r][c][k]||0)+(d[k]||0); });
          });
        });
        // 合并 models
        Object.entries(bd.models||{}).forEach(([m,d])=>{
          if(!merged.models[m]) merged.models[m]={};
          Object.keys(d).forEach(k=>{ merged.models[m][k]=(merged.models[m][k]||0)+(d[k]||0); });
        });
      });
      return merged;
    }
    const src = isMonthly ? D.sec2.monthly : D.sec2.annual;
    return src[mklsS2oems[0]] || null;
  }

  // 品牌筛选构建
  function buildMklsS2BrandFilter(){
    const box = document.getElementById('mklsS2brandBox');
    const btnsDiv = document.getElementById('mklsS2brandBtns');
    if(!box || !btnsDiv) return;
    // 多选车企或无品牌数据时隐藏
    if(isMklsS2Multi() || mklsS2oems.length !== 1){
      box.style.display = 'none';
      return;
    }
    const brands = (D.oemBrands||{})[mklsS2oems[0]];
    if(!brands || brands.length <= 1){
      box.style.display = 'none';
      return;
    }
    box.style.display = '';
    btnsDiv.innerHTML = '';
    // "全部"按钮
    const allOn = mklsS2brandSel.length === 0;
    const allBtn = document.createElement('span');
    allBtn.className = 'ychk'+(allOn?' on':'');
    allBtn.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
    allBtn.innerHTML = '<span class="ydot" style="background:#C9A84C"></span>全部';
    allBtn.addEventListener('click',()=>{
      mklsS2brandSel = [];
      buildMklsS2BrandFilter();
      renderMklsS2();
    });
    btnsDiv.appendChild(allBtn);
    // 各品牌
    brands.forEach((b,bi)=>{
      const on = mklsS2brandSel.includes(b.name);
      const btn = document.createElement('span');
      btn.className = 'ychk'+(on?' on':'');
      btn.style.cssText = 'font-size:11px;padding:2px 8px;cursor:pointer';
      btn.innerHTML = '<span class="ydot" style="background:'+CS[bi%CS.length]+'"></span>'+b.name;
      btn.addEventListener('click',()=>{
        const idx = mklsS2brandSel.indexOf(b.name);
        if(idx>=0){ mklsS2brandSel.splice(idx,1); }
        else { mklsS2brandSel.push(b.name); }
        buildMklsS2BrandFilter();
        renderMklsS2();
      });
      btnsDiv.appendChild(btn);
    });
  }

  /* ── 初始化车企筛选按钮 ── */
  (function initMklsS2OemFilter(){
    const btnBox = document.getElementById('mklsS2oemBtns');
    const sel = document.getElementById('mklsS2oemSelect');
    const codedOems = D.sec2.codedOems || [];
    const uncodedOems = D.sec2.uncodedOems || [];
    const top15 = codedOems.slice(0, 15);

    top15.forEach((oem, btnIdx) => {
      const btn = document.createElement('span');
      btn.className = 'ychk' + (mklsS2oems.includes(oem) ? ' on' : '');
      btn.style.cssText = 'cursor:pointer';
      btn.dataset.oem = oem;
      btn.innerHTML = '<span class="ydot" style="width:6px;height:6px;background:' + CS[btnIdx % CS.length] + '"></span>' + oem;
      btn.addEventListener('click', function(){
        if(mklsS2mode === 'annual' && mklsS2yearSel !== 'all'){
          const idx = mklsS2oems.indexOf(oem);
          if(idx>=0){ if(mklsS2oems.length>1) mklsS2oems.splice(idx,1); }
          else mklsS2oems.push(oem);
        } else {
          mklsS2oems = [oem];
        }
        mklsS2brandSel = [];
        sel.value = '';
        refreshMklsS2OemHighlight();
        renderMklsS2();
      });
      btnBox.appendChild(btn);
    });

    const remaining = [...codedOems.slice(15), ...uncodedOems];
    remaining.forEach(oem => {
      const opt = document.createElement('option');
      opt.value = oem;
      opt.textContent = oem;
      sel.appendChild(opt);
    });
    sel.addEventListener('change', function(){
      if(!this.value) return;
      if(mklsS2mode === 'annual' && mklsS2yearSel !== 'all'){
        if(!mklsS2oems.includes(this.value)) mklsS2oems.push(this.value);
      } else {
        mklsS2oems = [this.value];
      }
      mklsS2brandSel = [];
      refreshMklsS2OemHighlight();
      renderMklsS2();
    });
  })();

  function refreshMklsS2OemHighlight(){
    document.querySelectorAll('#mklsS2oemBtns .ychk').forEach(btn => {
      btn.classList.toggle('on', mklsS2oems.includes(btn.dataset.oem));
    });
    const sel = document.getElementById('mklsS2oemSelect');
    const codedTop15 = (D.sec2.codedOems || []).slice(0, 15);
    if(mklsS2oems.every(o => codedTop15.includes(o))){
      sel.value = '';
    }
    const extraBox = document.getElementById('mklsS2oemExtraTags');
    extraBox.innerHTML = '';
    const extraOems = mklsS2oems.filter(o => !codedTop15.includes(o));
    extraOems.forEach(oem => {
      const tag = document.createElement('span');
      tag.className = 'ychk on';
      tag.style.cssText = 'cursor:pointer;font-size:11px;padding:2px 8px;background:var(--blue);border-color:var(--blue);color:#fff';
      tag.textContent = oem + ' \u2715';
      tag.addEventListener('click', ()=>{
        const idx = mklsS2oems.indexOf(oem);
        if(idx >= 0 && mklsS2oems.length > 1) mklsS2oems.splice(idx, 1);
        sel.value = '';
        refreshMklsS2OemHighlight();
        renderMklsS2();
      });
      extraBox.appendChild(tag);
    });
  }

  /* ── 构建板块二区域筛选按钮（国家明细） ── */
  function buildMklsS2RegFilter(){
    const container = document.getElementById('mklsS2regFilter');
    container.innerHTML = '';
    const isMonthly = mklsS2mode === 'monthly';
    const src = isMonthly ? D.sec2.monthly : D.sec2.annual;
    const contSet = new Set();
    mklsS2oems.forEach(oem=>{ const od=src[oem]; if(od) MKLS_CONT.forEach(c=>{ if((od.byContinent||{})[c]) contSet.add(c); }); });
    const availConts = MKLS_CONT.filter(c => contSet.has(c));
    const allRegs = ['全部区域', ...availConts];
    if(!mklsS2region || !allRegs.includes(mklsS2region)){
      mklsS2region = allRegs[0] || null;
    }
    allRegs.forEach((r, i) => {
      const btn = document.createElement('span');
      btn.className = 'ychk' + (mklsS2region === r ? ' on' : '');
      btn.style.cssText = 'cursor:pointer;font-size:12px;padding:3px 8px';
      const dotColor = r === '全部区域' ? '#C9A84C' : CS[(i-1) % CS.length];
      btn.innerHTML = '<span class="ydot" style="width:5px;height:5px;background:' + dotColor + '"></span>' + r;
      btn.addEventListener('click', () => {
        mklsS2region = r;
        buildMklsS2RegFilter();
        renderMklsS2Country();
      });
      container.appendChild(btn);
    });
  }

  /* ── 渲染分动力类型 ── */
  function renderMklsS2Energy(){
    const isMonthly = mklsS2mode === 'monthly';
    document.getElementById('mklsS2engLbl').textContent = '分动力类型 · ' + getMklsS2OemLabel();
    if(isMklsS2Multi()){
      const yr = getMklsS2Years()[0];
      const engList = D.meta.energyList;
      const dataMap = {};
      engList.forEach(e=>{
        dataMap[e] = {};
        mklsS2oems.forEach(oem=>{
          const od = D.sec2.annual[oem];
          const src = od ? (od.byEnergy||{}) : {};
          dataMap[e][oem] = (src[e]||{})[yr]||0;
          dataMap[e][oem+'_yoy'] = (src[e]||{})[yr+'_yoy']??null;
        });
      });
      const oemColors = {};
      mklsS2oems.forEach((o,i)=>{ oemColors[o] = CS[i%CS.length]; });
      buildBar('mklsC2eng', engList.filter(e=>mklsS2oems.some(o=>dataMap[e][o]>0)), dataMap, mklsS2oems, {keyColors:oemColors});
    } else {
      const oemData = getMklsS2SingleData();
      if(!oemData){ const c=gi('mklsC2eng'); if(c) c.clear(); return; }
      if(isMonthly){
        const engList = D.meta.energyList.filter(e=>oemData.byEnergy[e]);
        buildBar('mklsC2eng', engList, oemData.byEnergy, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:false});
      } else {
        const engList = D.meta.energyList.filter(e=>(oemData.byEnergy||{})[e]);
        buildBar('mklsC2eng', engList, oemData.byEnergy||{}, getMklsS2Years(), {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
      }
    }
  }

  /* ── 渲染分大洲 ── */
  function renderMklsS2Region(){
    const isMonthly = mklsS2mode === 'monthly';
    document.getElementById('mklsS2regLbl').textContent = '分大洲 · ' + getMklsS2OemLabel();
    if(isMklsS2Multi()){
      const yr = getMklsS2Years()[0];
      const contSet = new Set();
      mklsS2oems.forEach(oem=>{
        const od = D.sec2.annual[oem];
        if(od) MKLS_CONT.forEach(c=>{ if((od.byContinent||{})[c]) contSet.add(c); });
      });
      const conts = MKLS_CONT.filter(c=>contSet.has(c));
      const dataMap = {};
      conts.forEach(c=>{
        dataMap[c] = {};
        mklsS2oems.forEach(oem=>{
          const od = D.sec2.annual[oem];
          const regObj = od ? ((od.byContinent||{})[c]||{}) : {};
          dataMap[c][oem] = regObj[yr]||0;
          dataMap[c][oem+'_yoy'] = regObj[yr+'_yoy']??null;
        });
      });
      const oemColors = {};
      mklsS2oems.forEach((o,i)=>{ oemColors[o]=CS[i%CS.length]; });
      buildBar('mklsC2reg', conts, dataMap, mklsS2oems, {keyColors:oemColors});
    } else {
      const oemData = getMklsS2SingleData();
      if(!oemData){ const c=gi('mklsC2reg'); if(c) c.clear(); return; }
      if(isMonthly){
        const contList = MKLS_CONT.filter(c=>oemData.byContinent[c]);
        buildBar('mklsC2reg', contList, oemData.byContinent, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:false});
      } else {
        const contList = MKLS_CONT.filter(c=>(oemData.byContinent||{})[c]);
        buildBar('mklsC2reg', contList, oemData.byContinent||{}, getMklsS2Years(), {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
      }
    }
  }

  /* ── 渲染国家明细 ── */
  function renderMklsS2Country(){
    const drillLbl = document.getElementById('mklsS2drillLbl');
    const region = mklsS2region;
    if(!region){
      drillLbl.textContent = '国家明细 · 选择区域查看';
      const inner = document.getElementById('mklsC2countryInner');
      inner.querySelectorAll('[id^="mklsC2ctry_"]').forEach(el=>{
        const inst = echarts.getInstanceByDom(el); if(inst) inst.dispose();
        if(el.id && _c[el.id]) delete _c[el.id];
      });
      inner.innerHTML = '';
      inner.style.height = '340px';
      return;
    }
    const isMonthly = mklsS2mode === 'monthly';
    drillLbl.textContent = '国家明细 · ' + getMklsS2OemLabel() + ' · ' + region;

    if(isMklsS2Multi()){
      const yr = getMklsS2Years()[0];
      const inner = document.getElementById('mklsC2countryInner');
      inner.querySelectorAll('[id^="mklsC2ctry_"]').forEach(el=>{
        const inst = echarts.getInstanceByDom(el); if(inst) inst.dispose();
        if(el.id && _c[el.id]) delete _c[el.id];
      });
      inner.innerHTML = '';
      inner.style.height = 'auto';

      mklsS2oems.forEach((oem, oi) => {
        const od = D.sec2.annual[oem]; if(!od) return;
        const rc = (od.countryByRegion||{})[region]||{};
        const countries = Object.keys(rc);
        const sorted = countries.sort((a,b) => ((rc[b]||{})[yr]||0) - ((rc[a]||{})[yr]||0));
        if(!sorted.length) return;
        const top5 = sorted.slice(0, 5);
        const rest = sorted.slice(5);
        const chartData = {};
        top5.forEach(c=>{ chartData[c] = rc[c]; });
        if(rest.length > 0){
          const otherD = {};
          [yr].forEach(y=>{ otherD[y] = 0; });
          rest.forEach(c=>{ [yr].forEach(y=>{ otherD[y] += ((rc[c]||{})[y]||0); }); });
          chartData['\u5176\u4ed6'] = otherD;
        }
        const cats = [...top5, ...(chartData['\u5176\u4ed6'] ? ['\u5176\u4ed6'] : [])];
        const subWrap = document.createElement('div');
        subWrap.style.cssText = 'background:var(--bg);border:1px solid var(--bd);border-radius:8px;padding:10px;margin-bottom:10px';
        const subTitle = document.createElement('div');
        subTitle.className = 'cw-lbl';
        subTitle.style.cssText = 'font-size:14px;margin-bottom:4px';
        subTitle.textContent = oem + ' \u00b7 TOP5\u56fd\u5bb6+\u5176\u4ed6';
        subWrap.appendChild(subTitle);
        const chartDiv = document.createElement('div');
        chartDiv.id = 'mklsC2ctry_' + oi;
        chartDiv.style.height = '260px';
        subWrap.appendChild(chartDiv);
        inner.appendChild(subWrap);
        setTimeout(()=>{
          buildBar('mklsC2ctry_'+oi, cats, chartData, [yr], {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
        }, 50);
      });
    } else {
      const inner = document.getElementById('mklsC2countryInner');
      inner.querySelectorAll('[id^="mklsC2ctry_"]').forEach(el=>{
        const inst = echarts.getInstanceByDom(el); if(inst) inst.dispose();
        if(el.id && _c[el.id]) delete _c[el.id];
      });
      if(_c['mklsC2countryInner']) { try{_c['mklsC2countryInner'].dispose();}catch(e){} delete _c['mklsC2countryInner']; }
      inner.innerHTML = '';
      inner.style.height = '340px';
      inner.id = 'mklsC2countryInner';

      const oemData = getMklsS2SingleData();
      if(!oemData){ return; }
      const countryData = (oemData.countryByRegion || {})[region] || {};
      const countries = Object.keys(countryData);
      let sorted;
      if(isMonthly){
        sorted = countries.sort((a,b)=>{
          const sumA = MK_M.reduce((s,m)=>s+((countryData[a]||{})[m]||0),0);
          const sumB = MK_M.reduce((s,m)=>s+((countryData[b]||{})[m]||0),0);
          return sumB - sumA;
        });
      } else {
        const latestYr = getMklsS2Years()[getMklsS2Years().length-1];
        sorted = countries.sort((a,b) => ((countryData[b]||{})[latestYr]||0) - ((countryData[a]||{})[latestYr]||0));
      }
      const top5 = sorted.slice(0,5);
      if(isMonthly){
        buildBar('mklsC2countryInner', top5, countryData, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:false});
      } else {
        buildBar('mklsC2countryInner', top5, countryData, getMklsS2Years(), {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
      }

      // 查看更多
      const wrap = document.getElementById('mklsC2countryContainer');
      let btn = wrap.querySelector('.show-more-btn');
      if(sorted.length > 5){
        if(!btn){ btn=document.createElement('button'); btn.className='show-more-btn'; wrap.appendChild(btn); }
        btn.textContent = '\u67e5\u770b\u66f4\u591a\uff08\u5171'+sorted.length+'\u56fd\uff09';
        btn.onclick = function(){
          const _isMonthly = mklsS2mode === 'monthly';
          const keys = _isMonthly ? MK_M : getMklsS2Years();
          const top15 = sorted.slice(0,15);
          const rest2 = sorted.slice(15);
          const md = {};
          top15.forEach(c=>{ md[c]=countryData[c]; });
          if(rest2.length>0){
            const od={}; keys.forEach(k=>{od[k]=0;}); rest2.forEach(c=>{keys.forEach(k=>{od[k]+=((countryData[c]||{})[k]||0);});}); md['\u5176\u4ed6']=od;
          }
          const cats=[...top15,...(md['\u5176\u4ed6']?['\u5176\u4ed6']:[])];
          const mask = document.createElement('div'); mask.className='modal-mask';
          mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">\u56fd\u5bb6\u660e\u7ec6 \u00b7 '+getMklsS2OemLabel()+' \u00b7 '+region+'\uff08\u524d15+\u5176\u4ed6\uff09</div><div id="mklsC2countryModal" style="height:500px"></div></div>';
          document.body.appendChild(mask);
          mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
          mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
          setTimeout(()=>{
            if(_c['mklsC2countryModal']) delete _c['mklsC2countryModal'];
            if(_isMonthly){
              buildBar('mklsC2countryModal', cats, md, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:false});
            } else {
              buildBar('mklsC2countryModal', cats, md, getMklsS2Years(), {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
            }
            addModalExport(mask.querySelector('.modal-box'), 'mklsC2countryModal', '\u56fd\u5bb6\u660e\u7ec6 \u00b7 '+getMklsS2OemLabel()+' \u00b7 '+region);
          },100);
        };
      } else if(btn){ btn.remove(); }
    }
  }

  /* ── 车型区域筛选器 ── */
  function buildMklsS2ModelRegFilter(){
    const container = document.getElementById('mklsS2modelRegFilter');
    if(!container) return;
    container.innerHTML = '';
    const src = mklsS2mode === 'monthly' ? D.sec2.monthly : D.sec2.annual;
    const contSet = new Set();
    mklsS2oems.forEach(oem=>{
      const od = src[oem]; if(!od) return;
      MKLS_CONT.forEach(c=>{ if((od.byContinent||{})[c]) contSet.add(c); });
    });
    const availRegs = ['\u5168\u90e8\u533a\u57df', ...MKLS_CONT.filter(c=>contSet.has(c))];
    if(!availRegs.includes(mklsS2modelRegion)){
      mklsS2modelRegion = '\u5168\u90e8\u533a\u57df';
    }
    availRegs.forEach((r, i) => {
      const btn = document.createElement('span');
      btn.className = 'ychk' + (mklsS2modelRegion === r ? ' on' : '');
      btn.style.cssText = 'font-size:12px;padding:4px 10px;cursor:pointer';
      const dotColor = r === '\u5168\u90e8\u533a\u57df' ? '#C9A84C' : CS[(i-1) % CS.length];
      btn.innerHTML = '<span class="ydot" style="width:5px;height:5px;background:' + dotColor + '"></span>' + r;
      btn.addEventListener('click', () => {
        mklsS2modelRegion = r;
        mklsS2modelCountry = null;
        buildMklsS2ModelRegFilter();
        buildMklsS2ModelCountryFilter();
        renderMklsS2Model();
      });
      container.appendChild(btn);
    });
  }

  /* ── 车型动力类型筛选器 ── */
  function buildMklsS2ModelEngFilter(){
    const container = document.getElementById('mklsS2modelEngFilter');
    if(!container) return;
    container.innerHTML = '';
    const energyOpts = [
      {key:'all', label:'\u5168\u90e8'},
      {key:'\u65b0\u80fd\u6e90', label:'\u65b0\u80fd\u6e90'},
      {key:'\u7eaf\u7535\u52a8', label:'\u7eaf\u7535\u52a8'},
      {key:'\u63d2\u7535\u6df7\u52a8', label:'\u63d2\u6df7'},
      {key:'\u5e38\u89c4\u6df7\u52a8', label:'\u5e38\u89c4\u6df7\u52a8'},
      {key:'\u71c3\u6cb9\u8f66', label:'\u71c3\u6cb9\u8f66'}
    ];
    energyOpts.forEach(({key, label}) => {
      const btn = document.createElement('span');
      btn.className = 'ychk' + (mklsS2modelEnergy === key ? ' on' : '');
      btn.style.cssText = 'font-size:12px;padding:4px 10px;cursor:pointer';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        mklsS2modelEnergy = key;
        buildMklsS2ModelEngFilter();
        renderMklsS2Model();
      });
      container.appendChild(btn);
    });
  }

  /* ── 车型国家下拉 ── */
  function buildMklsS2ModelCountryFilter(){
    const sel = document.getElementById('mklsS2modelCountrySel');
    if(!sel) return;
    while(sel.options.length > 1) sel.remove(1);
    const src = mklsS2mode === 'monthly' ? D.sec2.monthly : D.sec2.annual;
    const countrySet = new Map();
    mklsS2oems.forEach(oem=>{
      const od = src[oem]; if(!od) return;
      const rc = (od.countryByRegion||{})[mklsS2modelRegion]||{};
      Object.keys(rc).forEach(c=>{
        const sum = (rc[c]||{})['2025']||0;
        countrySet.set(c, (countrySet.get(c)||0) + sum);
      });
    });
    const sortedCountries = [...countrySet.entries()].sort((a,b)=>b[1]-a[1]).map(e=>e[0]);
    sortedCountries.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c;
      opt.textContent = c;
      sel.appendChild(opt);
    });
    if(mklsS2modelCountry && !sortedCountries.includes(mklsS2modelCountry)){
      mklsS2modelCountry = null;
    }
    sel.value = mklsS2modelCountry || '';
    if(!sel._bound){
      sel.addEventListener('change', function(){
        mklsS2modelCountry = this.value || null;
        renderMklsS2Model();
      });
      sel._bound = true;
    }
  }

  /* ── 获取车型数据（支持区域+国家+动力筛选）── */
  function filterMklsModelsByEnergy(modelData, energyFilter){
    if(!energyFilter || energyFilter === 'all') return modelData;
    const ME = D.modelEnergy || {};
    const filtered = {};
    Object.keys(modelData).forEach(model => {
      const me = ME[model] || '';
      let match = false;
      if(energyFilter === '\u65b0\u80fd\u6e90'){
        match = (me === '\u7eaf\u7535\u52a8' || me === '\u63d2\u7535\u6df7\u52a8');
      } else {
        match = (me === energyFilter);
      }
      if(match) filtered[model] = modelData[model];
    });
    return filtered;
  }

  function getMklsModelData(oemData, region, country, energyFilter){
    let raw;
    if(country){
      raw = ((oemData.modelByRegionCountry||{})[region]||{})[country]||{};
    } else {
      const allCountryModels = (oemData.modelByRegionCountry||{})[region]||{};
      raw = {};
      Object.values(allCountryModels).forEach(countryModels => {
        Object.keys(countryModels).forEach(model => {
          if(!raw[model]) raw[model] = {};
          const md = countryModels[model];
          Object.keys(md).forEach(k => {
            raw[model][k] = (raw[model][k]||0) + (md[k]||0);
          });
        });
      });
    }
    return filterMklsModelsByEnergy(raw, energyFilter);
  }

  /* ── 渲染分车型 ── */
  function renderMklsS2Model(){
    const modelLbl = document.getElementById('mklsS2modelLbl');
    const modelLegend = document.getElementById('mklsS2modelLegend');
    const container = document.getElementById('mklsC2modelInner');
    const isMonthly = mklsS2mode === 'monthly';
    const region = mklsS2modelRegion;
    const country = mklsS2modelCountry;

    container.querySelectorAll('[id^="mklsC2mdl_"]').forEach(el=>{
      const inst = echarts.getInstanceByDom(el);
      if(inst) inst.dispose();
      if(el.id && _c[el.id]) delete _c[el.id];
    });
    container.innerHTML = '';

    const regionLabel = region || '\u5168\u90e8\u533a\u57df';
    const countryLabel = country || '\u5168\u90e8\u56fd\u5bb6';
    const energyLabel = mklsS2modelEnergy === 'all' ? '\u5168\u90e8\u52a8\u529b\u7c7b\u578b' : mklsS2modelEnergy;
    modelLegend.classList.toggle('hide', isMonthly);

    if(isMklsS2Multi()){
      modelLbl.textContent = '\u5206\u8f66\u578b\u5b9e\u9500\u91cf \u00b7 ' + regionLabel + ' \u00b7 ' + countryLabel + ' \u00b7 ' + energyLabel + '\uff08\u591a\u8f66\u4f01\u5bf9\u6bd4\uff09';
      const yr = getMklsS2Years()[0];
      mklsS2oems.forEach((oem, oi) => {
        const od = D.sec2.annual[oem]; if(!od) return;
        const modelData = getMklsModelData(od, region, country, mklsS2modelEnergy);
        const models = Object.keys(modelData);
        const sorted = models.sort((a,b) => ((modelData[b]||{})[yr]||0) - ((modelData[a]||{})[yr]||0));
        const top10 = sorted.slice(0, 10);
        if(!top10.length) return;
        const subWrap = document.createElement('div');
        subWrap.style.cssText = 'background:var(--bg);border:1px solid var(--bd);border-radius:8px;padding:10px;margin-bottom:10px';
        const subTitle = document.createElement('div');
        subTitle.className = 'cw-lbl';
        subTitle.style.cssText = 'font-size:14px;margin-bottom:4px';
        subTitle.textContent = oem + ' \u00b7 TOP10\u8f66\u578b';
        subWrap.appendChild(subTitle);
        const chartDiv = document.createElement('div');
        chartDiv.id = 'mklsC2mdl_' + oi;
        chartDiv.style.height = '280px';
        subWrap.appendChild(chartDiv);
        container.appendChild(subWrap);
        setTimeout(()=>{
          const topData = {};
          top10.forEach(m=>{ topData[m] = modelData[m]; });
          buildBar('mklsC2mdl_'+oi, top10, topData, [yr], {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
        }, 50);
      });
    } else {
      modelLbl.textContent = '\u5206\u8f66\u578b\u5b9e\u9500\u91cf \u00b7 ' + getMklsS2OemLabel() + ' \u00b7 ' + regionLabel + ' \u00b7 ' + countryLabel + ' \u00b7 ' + energyLabel;
      const oemData = getMklsS2SingleData();
      if(!oemData) return;

      let modelData, models, sorted;
      if(isMonthly){
        // 月度用简单models数据（无区域/国家/动力筛选）
        modelData = oemData.models || {};
        models = Object.keys(modelData);
        sorted = models.sort((a,b)=>{
          const sumA = MK_M.reduce((s,m)=>s+((modelData[a]||{})[m]||0),0);
          const sumB = MK_M.reduce((s,m)=>s+((modelData[b]||{})[m]||0),0);
          return sumB - sumA;
        });
      } else {
        modelData = getMklsModelData(oemData, region, country, mklsS2modelEnergy);
        models = Object.keys(modelData);
        const latestYr = getMklsS2Years()[getMklsS2Years().length-1];
        sorted = models.sort((a,b) => ((modelData[b]||{})[latestYr]||0) - ((modelData[a]||{})[latestYr]||0));
      }

      const top15 = sorted.slice(0, 15);
      const chartDiv = document.createElement('div');
      chartDiv.id = 'mklsC2mdl_0';
      chartDiv.style.height = '420px';
      container.appendChild(chartDiv);
      setTimeout(()=>{
        if(isMonthly){
          buildBar('mklsC2mdl_0', top15, modelData, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:false});
        } else {
          buildBar('mklsC2mdl_0', top15, modelData, getMklsS2Years(), {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
        }
      }, 50);

      if(sorted.length > 15){
        const moreBtn = document.createElement('button');
        moreBtn.className = 'show-more-btn';
        moreBtn.textContent = '\u67e5\u770b\u66f4\u591a\uff08\u5171'+sorted.length+'\u6b3e\u8f66\u578b\uff09';
        moreBtn.onclick = function(){
          const _isMonthly = mklsS2mode === 'monthly';
          const mask = document.createElement('div'); mask.className='modal-mask';
          mask.innerHTML = '<div class="modal-box"><button class="modal-close">&times;</button><div class="cw-lbl" style="margin-bottom:8px;font-size:17px">\u5206\u8f66\u578b\u5b9e\u9500\u91cf \u00b7 '+getMklsS2OemLabel()+' \u00b7 '+countryLabel+'\uff08\u5168\u90e8\u8f66\u578b\uff09</div><div id="mklsC2mdlModal" style="height:500px"></div></div>';
          document.body.appendChild(mask);
          mask.querySelector('.modal-close').addEventListener('click',()=>mask.remove());
          mask.addEventListener('click',e=>{ if(e.target===mask) mask.remove(); });
          setTimeout(()=>{
            if(_c['mklsC2mdlModal']) delete _c['mklsC2mdlModal'];
            if(_isMonthly){
              buildBar('mklsC2mdlModal', sorted, modelData, MK_M, {keyLabels:ML_MKLS, keyColors:MC_MKLS, showYoy:false});
            } else {
              buildBar('mklsC2mdlModal', sorted, modelData, getMklsS2Years(), {keyColors:YC_MKLS, keyLabels:YR_LBL_M});
            }
            addModalExport(mask.querySelector('.modal-box'), 'mklsC2mdlModal', '\u5206\u8f66\u578b \u00b7 '+getMklsS2OemLabel()+' \u00b7 '+countryLabel);
          },100);
        };
        container.appendChild(moreBtn);
      }
    }
  }

  /* ── 主渲染 ── */
  function renderMklsS2(){
    const isMonthly = mklsS2mode === 'monthly';
    document.getElementById('mklsS2yearBox').style.display = isMonthly ? 'none' : '';
    document.querySelectorAll('#mklsSec2 .legend-note').forEach(n => {
      n.classList.toggle('hide', isMonthly);
    });
    buildMklsS2BrandFilter();
    renderMklsS2Energy();
    renderMklsS2Region();
    buildMklsS2RegFilter();
    renderMklsS2Country();
    buildMklsS2ModelRegFilter();
    buildMklsS2ModelEngFilter();
    buildMklsS2ModelCountryFilter();
    renderMklsS2Model();
    setTimeout(addExportButtons, 300);
  }

  /* ── 事件绑定 ── */
  // Tab切换
  document.querySelectorAll('#mklsS2timeTabs .tab-btn').forEach(btn=>{
    btn.addEventListener('click', function(){
      document.querySelectorAll('#mklsS2timeTabs .tab-btn').forEach(b=>b.classList.remove('on'));
      this.classList.add('on');
      mklsS2mode = this.dataset.mode;
      if(mklsS2mode === 'monthly' && mklsS2oems.length > 1){
        mklsS2oems = [mklsS2oems[0]];
        refreshMklsS2OemHighlight();
      }
      renderMklsS2();
    });
  });

  // 年份单选
  document.querySelectorAll('#mklsS2yearBox .ychk').forEach(lbl => {
    lbl.addEventListener('click', function(){
      const rb = this.querySelector('input');
      rb.checked = true;
      document.querySelectorAll('#mklsS2yearBox .ychk').forEach(l=>l.classList.remove('on'));
      this.classList.add('on');
      const prev = mklsS2yearSel;
      mklsS2yearSel = rb.value;
      if((prev==='all' && mklsS2yearSel!=='all') || (prev!=='all' && mklsS2yearSel==='all')){
        mklsS2oems = [mklsS2oems[0]];
        refreshMklsS2OemHighlight();
      }
      renderMklsS2();
    });
  });

  /* ── Tab切换时重渲染 ── */
  window._mklsRerender = function(){
    Object.keys(_c).forEach(id=>{ if(id.startsWith('mkls')){ try{_c[id].dispose();}catch(e){} delete _c[id]; } });
    renderMklsS0();
    renderMklsS1();
    renderMklsS2();
    setTimeout(()=>{
      Object.entries(_c).forEach(([id,chart])=>{ if(id.startsWith('mkls')) try{chart.resize();}catch(e){} });
    }, 200);
  };

  renderMklsS0();
  renderMklsS2();
})();
}

/* ══════════════════════════════════════════
   数据源Tab切换
══════════════════════════════════════════ */
function switchTab(src){
  document.querySelectorAll('#srcTabs .src-tab').forEach(t => t.classList.remove('on'));
  document.querySelectorAll('#srcTabs .src-tab').forEach(t => { if(t.dataset.src===src) t.classList.add('on'); });
  ['srcHaiguan','srcZhongqi','srcMkls','srcInventory','srcLocalProd','srcPriceTrack'].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.style.display = 'none';
  });
  const tabMap = {haiguan:'srcHaiguan',zhongqi:'srcZhongqi',mkls:'srcMkls',inventory:'srcInventory',localProd:'srcLocalProd',priceTrack:'srcPriceTrack'};
  const target = document.getElementById(tabMap[src]);
  if(target) target.style.display = '';
  location.hash = src;
  window.scrollTo(0, 0);
  // 切换后重新渲染图表（display:none→block需要重绘）
  if(src === 'haiguan'){
    setTimeout(()=>{
      Object.entries(_c).forEach(([id,chart])=>{ if(!id.startsWith('caam') && !id.startsWith('mkls')) try{chart.resize();}catch(e){} });
    }, 150);
  }
  if(src === 'zhongqi' && window._caamRerender){
    setTimeout(window._caamRerender, 100);
  }
  if(src === 'mkls' && window._mklsRerender){
    setTimeout(window._mklsRerender, 100);
  }
}
document.querySelectorAll('#srcTabs .src-tab').forEach(tab => {
  tab.addEventListener('click', function(){ switchTab(this.dataset.src); });
});
// 刷新时恢复上次Tab（所有Tab初始隐藏，由JS决定显示）
(function(){
  const h = location.hash.replace('#','');
  const target = ['haiguan','zhongqi','mkls','inventory','localProd','priceTrack'].includes(h) ? h : 'zhongqi';
  switchTab(target);
})();
</script>
</body>
</html>'''

final = HTML.replace('__DATA__', DATA_JSON).replace('__CAAM_DATA__', CAAM_JSON).replace('__MKLS_DATA__', MKLS_JSON)
out = os.path.join(HERE, 'index.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(final)

print(f'Done -> {out} ({len(final):,} bytes)')
