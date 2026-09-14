# -*- coding: utf-8 -*-
"""港股（HKFRS/IFRS）年报附注五科目分析表生成器（Mac 版，openpyxl）。

港股 PP&E / 无形资产附注按「账面净值滚动」披露（年初净值 → 添置 → 轉讓 → 出售(净值) → 折舊 → 年末净值，
并另给年初/年末的成本与累计折旧），在建工程是 PP&E 里的一列而非独立附注；研发资本化多数为 0；
政府补助只有「其他收入·政府補貼」「遞延收入」「现金流量表·收取的政府補貼」三个口径。
本生成器按 references/*.md 的区块结构落表，把上述口径差异内建成公式，产出：
    固定资产 / 在建工程 / 无形资产 / 资本化研发 / 政府补助分析  五张 sheet。

用法:  python3 build_hk_notes_model.py <spec.json>
spec.json 字段见 README 段（文件末尾 SPEC_DOC）。金额单位跟年报（千元）。
所有历史值硬编码（深蓝），预测假设深红，公式黑；校验行绿底，合计行黄底。
"""
import sys, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

FONT = "Microsoft YaHei"
NAVY = "FF1F4E78"; RED = "FFC00000"; BLACK = "FF000000"; WHITE = "FFFFFFFF"
F_SEC = PatternFill("solid", fgColor="FFD9E2F3")
F_TOT = PatternFill("solid", fgColor="FFFFE699")
F_CHK = PatternFill("solid", fgColor="FFE2EFDA")
F_HDR = PatternFill("solid", fgColor=NAVY)
NUM = r'#,##0;\(#,##0\);\-'
PCT1 = "0.0%"; PCT2 = "0.00%"


class Sheet:
    """按 conventions.md 的布局写表：B=科目, C=来源, D.. 年份, 最后一列=假设理由。"""

    def __init__(self, wb, name, hist, fc, unit):
        self.ws = wb.create_sheet(name)
        self.name = name
        self.hist = list(hist); self.fc = list(fc); self.years = self.hist + self.fc
        self.col = {y: 4 + i for i, y in enumerate(self.years)}
        self.reason_col = 4 + len(self.years)
        self.unit = unit
        self.r = 6
        self.rows = {}
        ws = self.ws
        ws.column_dimensions["A"].width = 2
        ws.column_dimensions["B"].width = 44
        ws.column_dimensions["C"].width = 18
        for y in self.years:
            ws.column_dimensions[CL(self.col[y])].width = 13
        ws.column_dimensions[CL(self.reason_col)].width = 60
        ws.sheet_view.showGridLines = False

    # ---------- low-level ----------
    def L(self, y):
        return CL(self.col[y])

    def cell(self, r, c, v=None, color=BLACK, bold=False, fill=None, fmt=None, size=10, align=None):
        cell = self.ws.cell(r, c)
        if v is not None:
            cell.value = v
        cell.font = Font(name=FONT, size=size, bold=bold, color=color)
        if fill is not None:
            cell.fill = fill
        if fmt:
            cell.number_format = fmt
        if align:
            cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=(align == "left"))
        return cell

    def header(self, title, unit_line, check_line):
        self.cell(1, 2, title, NAVY, True, size=14)
        self.cell(2, 2, unit_line, BLACK)
        self.cell(3, 2, check_line, BLACK)
        self.cell(5, 2, "科目", WHITE, True, F_HDR)
        self.cell(5, 3, "附注 / 来源", WHITE, True, F_HDR)
        for y in self.years:
            self.cell(5, self.col[y], f"{y}{'A' if y in self.hist else 'E'}", WHITE, True, F_HDR, align="center")
        if self.fc:
            self.cell(5, self.reason_col, "假设理由 / 说明", WHITE, True, F_HDR)
        self.ws.freeze_panes = "D6"

    def section(self, title):
        r = self.r
        for c in range(2, self.reason_col + 1):
            self.cell(r, c, None, BLACK, True, F_SEC)
        self.cell(r, 2, title, BLACK, True, F_SEC)
        self.r += 1
        return r

    def blank(self):
        self.r += 1

    def row(self, label, source=None, hist=None, fc=None, key=None, fmt=NUM, total=False, check=False,
            hist_color=NAVY, fc_color=BLACK, reason=None, indent=0):
        """hist: dict{year: value|formula}; fc: dict{year: value|formula} 或 callable(y)->str。
        值为 str 且以 '=' 开头 → 公式(黑)；数值 → 硬编码(hist 深蓝 / fc 深红)。"""
        r = self.r
        fill = F_TOT if total else (F_CHK if check else None)
        self.cell(r, 2, ("　" * indent) + label, BLACK, total or check, fill)
        if isinstance(source, str) and source.startswith("="):
            source = "＝" + source[1:]   # 来源列是文字标签，不能以 = 开头（conventions.md）
        self.cell(r, 3, source, BLACK, False, fill)
        for y in self.years:
            v = None
            if y in self.hist and hist is not None:
                v = hist(y) if callable(hist) else hist.get(y)
                color = hist_color
            elif y in self.fc and fc is not None:
                v = fc(y) if callable(fc) else fc.get(y)
                color = fc_color
            else:
                color = BLACK
            if isinstance(v, str) and v.startswith("="):
                color = BLACK
            self.cell(r, self.col[y], v, color, total or check, fill, fmt, align="right")
        if reason:
            self.cell(r, self.reason_col, reason, BLACK, False, fill, align="left")
        if key:
            self.rows[key] = r
        self.r += 1
        return r

    def note_block(self, title, lines):
        self.section(title)
        for ln in lines:
            self.cell(self.r, 2, ln, BLACK)
            self.ws.merge_cells(start_row=self.r, start_column=2, end_row=self.r, end_column=self.reason_col)
            self.ws.cell(self.r, 2).alignment = Alignment(wrap_text=True, vertical="top")
            self.ws.row_dimensions[self.r].height = max(15, 15 * (1 + len(ln) // 110))
            self.r += 1
        self.blank()

    def ref(self, key, y, sheet=None):
        """未定义的 key 先给占位符 @@key|col@@，finalize() 时回填行号（支持前向引用）。"""
        sh = f"{sheet}!" if sheet else ""
        if key in self.rows:
            return f"{sh}{self.L(y)}{self.rows[key]}"
        return f"{sh}@@{key}|{self.L(y)}@@"

    def finalize(self):
        import re
        pat = re.compile(r"@@([A-Za-z0-9_\-]+)\|([A-Z]+)@@")
        for row in self.ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "@@" in c.value:
                    def sub(m):
                        k, col = m.group(1), m.group(2)
                        if k not in self.rows:
                            raise KeyError(f"unresolved row key {k!r} in {self.name}!{c.coordinate}")
                        return f"{col}{self.rows[k]}"
                    c.value = pat.sub(sub, c.value)

    def prev(self, y):
        return self.years[self.years.index(y) - 1]


def V(d, y, default=0):
    v = d.get(str(y), d.get(y)) if isinstance(d, dict) else None
    return default if v is None else v


# ======================================================================
def build_cip(wb, S, hist, fc, unit):
    """在建工程（港股 = PP&E 附注在建工程列，单项目）。"""
    P = S["ppe"]; C = P["cip"]; classes = P["classes"]
    sh = Sheet(wb, "在建工程", hist, fc, unit)
    note = P.get("note", "PP&E附注")
    sh.header(f"{note} 在建工程 — PP&E附注在建工程列滚动 + 转固分配 ({hist[0]}-{fc[-1] if fc else hist[-1]})",
              f"单位：{unit} ｜ 颜色：深蓝=年报实际 ｜ 深红=可调假设 ｜ 黑=公式 ｜ 黄底=合计 ｜ 绿底=校验",
              "勾稽：期末 = 期初 + 本年增加 − 转入固定资产 ｜ 港股无独立在建工程附注，数据来自 PP&E 附注「在建工程」列（按账面净值滚动，不计折旧）")
    h0 = hist[0]
    # 一
    sh.section("一、CIP总额勾稽（与 PP&E 附注在建工程列对账）")
    sh.row("在建工程（账面价值）", note, {y: V(C["nbv"], y) for y in hist}, (lambda y: f"={sh.ref('f_end', y)}") if fc else None, key="cip_nbv")
    sh.row("工程物资", "港股无此科目", {y: 0 for y in hist}, {y: 0 for y in fc}, key="cip_mat", hist_color=BLACK, fc_color=BLACK)
    sh.row("合计", None, lambda y: f"={sh.ref('cip_nbv', y)}+{sh.ref('cip_mat', y)}", lambda y: f"={sh.ref('cip_nbv', y)}+{sh.ref('cip_mat', y)}", key="cip_tot", total=True)
    sh.row("PP&E附注披露 在建工程期末净值", note, {y: V(C["nbv"], y) for y in hist}, key="cip_disc")
    sh.row("校验：合计 − 披露（应=0）", None, lambda y: f"={sh.ref('cip_tot', y)}-{sh.ref('cip_disc', y)}", key="cip_chk", check=True)
    sh.blank()
    # 二
    sh.section("二、CIP期末账面价值 — 按项目")
    sh.row("在建工程（年报未按项目披露，单列）", note, {y: V(C["nbv"], y) for y in hist}, key="proj1")
    sh.row("合计", None, lambda y: f"={sh.ref('proj1', y)}", key="proj_tot", total=True)
    sh.row("校验：vs 一.账面价值（应=0）", None, lambda y: f"={sh.ref('proj_tot', y)}-{sh.ref('cip_nbv', y)}", check=True)
    sh.blank()
    # 三 滚动（历史）
    sh.section("三、在建工程变动滚动表（历史，来自 PP&E 附注在建工程列）")
    yrs = hist
    sh.row("期初余额", note, {**{h0: V(C["open_first"], h0, V(C["nbv"], h0) - V(C["additions"], h0) + V(C["transfer_out"], h0))},
                                **{y: f"={sh.ref('r_end', sh.prev(y))}" if False else None for y in hist[1:]}}, key="r_open")
    # 期初 = 上年期末(公式)：需先知道 r_end 行号 → 先占位，后面回填
    sh.row("本年增加（添置）", note, {y: V(C["additions"], y) for y in hist}, key="r_add")
    sh.row("重分类（如有）", note, {y: V(C.get("reclass", {}), y) for y in hist}, key="r_recl")
    sh.row("企业合并（如有）", note, {y: 0 for y in hist}, key="r_comb")
    sh.row("本年转入固定资产（−）", note, {y: -V(C["transfer_out"], y) for y in hist}, key="r_tr")
    sh.row("其他减少（减值/处置，如有）", note, {y: -V(C.get("other_dec", {}), y) for y in hist}, key="r_oth")
    sh.row("期末（公式）", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['r_open']}:{sh.L(y)}{sh.rows['r_oth']})", key="r_end", total=True)
    sh.row("校验：期末公式 − 披露期末（应≈0，差异=减值）", None, lambda y: f"={sh.ref('r_end', y)}-{sh.ref('cip_disc', y)}", key="r_chk", check=True)
    sh.row("转固率 = 转入 ÷ (期初 + 本年增加)", None,
           lambda y: f'=IFERROR(-{sh.ref("r_tr", y)}/({sh.ref("r_open", y)}+{sh.ref("r_add", y)}+{sh.ref("r_recl", y)}+{sh.ref("r_comb", y)}),"")',
           key="r_rate", fmt=PCT1)
    # 回填期初 = 上年期末
    for y in hist[1:]:
        sh.cell(sh.rows["r_open"], sh.col[y], f"={sh.ref('r_end', sh.prev(y))}", BLACK, fmt=NUM, align="right")
    sh.blank()
    # 四 预算：省略
    sh.section("四、预算追踪 — 年报未披露在建项目预算，本节省略")
    sh.blank()
    # 五 转固分配
    sh.section("五、CIP→PP&E 转固分配 — 按资产类别（PP&E 附注「轉讓」行）")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(P["transfer_in"][cls], y) for y in hist}, key=f"alloc_{i}")
    sh.row("转固合计", None, lambda y: "=" + "+".join(sh.ref(f"alloc_{i}", y) for i in range(len(classes))), key="alloc_tot", total=True)
    sh.row("校验：vs 三.转入固定资产（应=0）", None, lambda y: f"={sh.ref('alloc_tot', y)}+{sh.ref('r_tr', y)}", check=True)
    for i, cls in enumerate(classes):
        sh.row(f"{cls} 占比", None, lambda y, i=i: f'=IFERROR({sh.ref(f"alloc_{i}", y)}/{sh.ref("alloc_tot", y)},"")', key=f"share_{i}", fmt=PCT1)
    sh.blank()
    if fc:
        n3 = hist[-3:]
        # 六 假设
        sh.section("六、预测假设输入（深红色为可调）")
        sh.row("6.1 本年新增CIP投入", "近3年平均", None,
               {y: f"=AVERAGE({sh.L(n3[0])}{sh.rows['r_add']}:{sh.L(n3[-1])}{sh.rows['r_add']})" for y in fc},
               key="a_add", reason=f"默认 = {n3[0]}-{n3[-1]} 年平均添置；可按资本开支指引改深红硬编码")
        sh.row("6.2 整体转固率", "近3年平均", None,
               {y: f"=AVERAGE({sh.L(n3[0])}{sh.rows['r_rate']}:{sh.L(n3[-1])}{sh.rows['r_rate']})" for y in fc},
               key="a_rate", fmt=PCT1, reason="默认 = 近3年转固率平均；产能集中投放年可上调")
        for i, cls in enumerate(classes):
            sh.row(f"6.3 转固分配比例 — {cls}", "近3年平均", None,
                   {y: f"=AVERAGE({sh.L(n3[0])}{sh.rows[f'share_{i}']}:{sh.L(n3[-1])}{sh.rows[f'share_{i}']})" for y in fc},
                   key=f"a_share_{i}", fmt=PCT1)
        sh.blank()
        # 七 滚动预测
        sh.section("七、CIP滚动预测 + 转固预测")
        sh.row("7.1 期初", None, None, lambda y: f"={sh.ref('r_end', y - 1) if (y - 1) in hist else sh.ref('f_end', y - 1)}", key="f_open")
        sh.row("　本年增加（=6.1）", None, None, lambda y: f"={sh.ref('a_add', y)}", key="f_add")
        sh.row("　转固（= −(期初+增加)×6.2）", None, None, lambda y: f"=-({sh.ref('f_open', y)}+{sh.ref('f_add', y)})*{sh.ref('a_rate', y)}", key="f_tr")
        sh.row("　期末", None, None, lambda y: f"={sh.ref('f_open', y)}+{sh.ref('f_add', y)}+{sh.ref('f_tr', y)}", key="f_end", total=True)
        sh.row("7.2 转固按资产类别拆分（固定资产表 8.4 引用源）", None, None, None)
        for i, cls in enumerate(classes):
            sh.row(f"　{cls}", None, None, lambda y, i=i: f"=-{sh.ref('f_tr', y)}*{sh.ref(f'a_share_{i}', y)}", key=f"f_alloc_{i}")
        sh.row("　合计", None, None, lambda y: "=" + "+".join(sh.ref(f"f_alloc_{i}", y) for i in range(len(classes))), total=True)
        sh.blank()
    rem = S.get("remarks", {}).get("在建工程", [])
    sh.note_block("八、说明与备注", [
        "【数据范围】港股无独立在建工程附注，本表全部取自 PP&E 附注「在建工程」列；港股按账面净值滚动，在建工程不计折旧，故净值=成本。",
        "【勾稽关系】期末 = 期初 + 本年添置 + 重分类 − 转入固定资产 − 其他减少；「轉讓」行各资产类别之和 = 在建工程列转出额（附注内恒成立）。",
        "【局限】未按项目披露，无法做项目级转固预测；预算追踪不可得。预测为参数化估计（6.1-6.3 深红可调）。",
    ] + rem)
    sh.finalize()
    return sh


def build_ppe(wb, S, hist, fc, unit, cip):
    P = S["ppe"]; classes = P["classes"]; N = len(classes); note = P.get("note", "PP&E附注")
    sh = Sheet(wb, "固定资产", hist, fc, unit)
    sh.header(f"{note} 固定资产（物業、廠房及設備）— {N} 类汇总 + 预测勾稽 ({hist[0]}-{fc[-1] if fc else hist[-1]})",
              f"单位：{unit} ｜ 颜色：深蓝=年报实际 ｜ 深红=可调假设 ｜ 黑=公式 ｜ 黄底=合计 ｜ 绿底=校验",
              "勾稽：NBV = 原价 − 累计折旧 − 减值 ｜ 本期计提 = 平均原价 × 折旧率 ｜ 期末原价 = 期初 + 购置 + CIP转入 − 处置原价（处置原价由港股净值口径倒算，见 8.6）")
    h0 = hist[0]
    cost = P["cost"]; acc = P["accdep"]; imp = P.get("impair", {}); dep = P["depreciation"]
    add = P["additions"]; tr = P["transfer_in"]; disp = P["disposal_nbv"]

    def nbv_h(cls, y):
        return V(cost[cls], y) - V(acc[cls], y) - V(imp.get(cls, {}), y)

    # 一 NBV
    sh.section(f"一、期末账面净值（NBV）— {N} 类汇总（不含在建工程）")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: nbv_h(cls, y) for y in hist},
               lambda y, i=i: f"={sh.ref(f'cost_{i}', y)}-{sh.ref(f'acc_{i}', y)}-{sh.ref(f'imp_{i}', y)}", key=f"nbv_{i}")
    sh.row("合计 NBV（不含在建工程）", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['nbv_0']}:{sh.L(y)}{sh.rows[f'nbv_{N-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['nbv_0']}:{sh.L(y)}{sh.rows[f'nbv_{N-1}']})", key="nbv_tot", total=True)
    sh.row("在建工程 NBV（=在建工程表）", "在建工程!一", lambda y: f"=在建工程!{cip.L(y)}{cip.rows['cip_nbv']}",
           lambda y: f"=在建工程!{cip.L(y)}{cip.rows['cip_nbv']}", key="nbv_cip")
    sh.row("合计 PP&E NBV（含在建工程）", None, lambda y: f"={sh.ref('nbv_tot', y)}+{sh.ref('nbv_cip', y)}",
           lambda y: f"={sh.ref('nbv_tot', y)}+{sh.ref('nbv_cip', y)}", key="nbv_all", total=True)
    sh.row("BS 披露 物業、廠房及設備", "合併資產負債表", {y: V(P["bs_nbv"], y) for y in hist}, key="bs")
    sh.row("校验：模型 − BS（应=0）", None, lambda y: f"={sh.ref('nbv_all', y)}-{sh.ref('bs', y)}", check=True)
    sh.blank()
    # 二 原价
    sh.section(f"二、期末账面原价（成本）— {N} 类汇总")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(cost[cls], y) for y in hist},
               lambda y, i=i: f"={sh.ref(f'cost_{i}', sh.prev(y))}+{sh.ref(f'buy_{i}', y)}+{sh.ref(f'cipin_{i}', y)}-{sh.ref(f'dispc_{i}', y)}", key=f"cost_{i}")
    sh.row("合计原价", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['cost_0']}:{sh.L(y)}{sh.rows[f'cost_{N-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['cost_0']}:{sh.L(y)}{sh.rows[f'cost_{N-1}']})", key="cost_tot", total=True)
    sh.blank()
    # 三 累计折旧
    sh.section(f"三、累计折旧 期末 — {N} 类汇总")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(acc[cls], y) for y in hist},
               lambda y, i=i: f"={sh.ref(f'acc_{i}', sh.prev(y))}*(1-{sh.ref(f'a_disp_{i}', y)})+{sh.ref(f'dep_{i}', y)}", key=f"acc_{i}")
    sh.row("合计累计折旧", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['acc_0']}:{sh.L(y)}{sh.rows[f'acc_{N-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['acc_0']}:{sh.L(y)}{sh.rows[f'acc_{N-1}']})", key="acc_tot", total=True)
    sh.blank()
    # 四 减值
    sh.section(f"四、减值准备 期末 — {N} 类汇总")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(imp.get(cls, {}), y) for y in hist}, {y: V(imp.get(cls, {}), hist[-1]) for y in fc}, key=f"imp_{i}",
               reason="预测默认延续末年期末值（年报未计提减值则为0）" if i == 0 else None)
    sh.row("合计减值", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['imp_0']}:{sh.L(y)}{sh.rows[f'imp_{N-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['imp_0']}:{sh.L(y)}{sh.rows[f'imp_{N-1}']})", total=True)
    sh.blank()
    # 五 计提
    sh.section(f"五、本期计提折旧 — {N} 类汇总（附注「折舊費用」行）")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(dep[cls], y) for y in hist},
               lambda y, i=i: f"=({sh.ref(f'cost_{i}', y)}+{sh.ref(f'cost_{i}', sh.prev(y))})/2*{sh.ref(f'a_rate_{i}', y)}", key=f"dep_{i}")
    sh.row("合计计提折旧", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['dep_0']}:{sh.L(y)}{sh.rows[f'dep_{N-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['dep_0']}:{sh.L(y)}{sh.rows[f'dep_{N-1}']})", key="dep_tot", total=True)
    dbf = P.get("dep_by_function", {})
    if dbf:
        sh.row("附：折旧按功能分布（附注披露，仅历史）", None, None, None)
        for lab, d in dbf.items():
            sh.row(f"　{lab}", note, {y: V(d, y) for y in hist}, key=f"dbf_{lab}")
        sh.row("　合计（应=五.合计）", None, lambda y: "=" + "+".join(sh.ref(f"dbf_{lab}", y) for lab in dbf), key="dbf_tot", total=True)
        sh.row("　校验：功能分布合计 − 计提合计（应=0）", None, lambda y: f"={sh.ref('dbf_tot', y)}-{sh.ref('dep_tot', y)}", check=True)
    sh.blank()
    # 六 折旧率
    sh.section("六、隐含折旧率反推（本期计提 ÷ 平均原价）")
    for i, cls in enumerate(classes):
        sh.row(f"{cls} 折旧率", None,
               lambda y, i=i: (f'=IFERROR({sh.ref(f"dep_{i}", y)}/(({sh.ref(f"cost_{i}", y)}+{sh.ref(f"cost_{i}", sh.prev(y))})/2),"")' if y != h0 else None),
               lambda y, i=i: f"={sh.ref(f'a_rate_{i}', y)}", key=f"rate_{i}", fmt=PCT2)
    sh.blank()
    # 七 新旧度
    sh.section("七、资产新旧度（累计折旧 ÷ 期末原价）")
    sh.row("整体新旧度", "老化指标", lambda y: f'=IFERROR({sh.ref("acc_tot", y)}/{sh.ref("cost_tot", y)},"")',
           lambda y: f'=IFERROR({sh.ref("acc_tot", y)}/{sh.ref("cost_tot", y)},"")', fmt=PCT1)
    sh.blank()
    # 八 假设
    sh.section("八、预测假设输入（深红色为可调）")
    n2 = hist[-2:]; n3 = hist[-3:]
    sh.row("8.1 购置增长率（历史=购置YoY反推）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", None,
               lambda y, i=i: (f'=IFERROR({sh.ref(f"buy_{i}", y)}/{sh.ref(f"buy_{i}", sh.prev(y))}-1,"")' if y != h0 else None),
               {y: 0 for y in fc}, key=f"a_grow_{i}", fmt=PCT1,
               reason="默认 0%（购置维持末年水平）；按资本开支指引调整" if i == 0 else None)
    sh.row("8.2 处置率（处置原价 ÷ 期初原价）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", None,
               lambda y, i=i: (f'=IFERROR({sh.ref(f"dispc_{i}", y)}/{sh.ref(f"cost_{i}", sh.prev(y))},"")' if y != h0 else None),
               {y: f"=IFERROR(AVERAGE({sh.L(n3[0])}{{r}}:{sh.L(n3[-1])}{{r}}),0)" for y in fc}, key=f"a_disp_{i}", fmt=PCT2,
               reason=f"默认 = {n3[0]}-{n3[-1]} 年平均处置率" if i == 0 else None)
    sh.row(f"8.3 折旧率（锁定 {n2[0]}-{n2[-1]} 平均）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", None, None,
               {y: f"=AVERAGE({sh.L(n2[0])}{{rate}}:{sh.L(n2[-1])}{{rate}})" for y in fc}, key=f"a_rate_{i}", fmt=PCT2,
               reason="默认 = 近2年隐含折旧率平均；类别结构变化大时可深红手调" if i == 0 else None)
    sh.row("8.4 在建工程转入（历史=附注「轉讓」；预测=在建工程表 7.2）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", note, {y: V(tr[cls], y) for y in hist},
               lambda y, i=i: f"=在建工程!{cip.L(y)}{cip.rows[f'f_alloc_{i}']}" if cip.rows.get(f"f_alloc_{i}") else 0, key=f"cipin_{i}")
    sh.row("　合计", None, lambda y: "=" + "+".join(sh.ref(f"cipin_{i}", y) for i in range(N)),
           lambda y: "=" + "+".join(sh.ref(f"cipin_{i}", y) for i in range(N)), total=True)
    sh.row("8.5 购置（历史=附注「添置」；预测=上年×(1+8.1)）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", note, {y: V(add[cls], y) for y in hist},
               lambda y, i=i: f"={sh.ref(f'buy_{i}', sh.prev(y))}*(1+{sh.ref(f'a_grow_{i}', y)})", key=f"buy_{i}")
    sh.row("　合计", None, lambda y: "=" + "+".join(sh.ref(f"buy_{i}", y) for i in range(N)),
           lambda y: "=" + "+".join(sh.ref(f"buy_{i}", y) for i in range(N)), total=True)
    sh.row("8.6 处置原价（历史=期初+购置+转入−期末 倒算；预测=期初原价×8.2）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", "倒算",
               lambda y, i=i: (f"={sh.ref(f'cost_{i}', sh.prev(y))}+{sh.ref(f'buy_{i}', y)}+{sh.ref(f'cipin_{i}', y)}-{sh.ref(f'cost_{i}', y)}" if y != h0 else 0),
               lambda y, i=i: f"={sh.ref(f'cost_{i}', sh.prev(y))}*{sh.ref(f'a_disp_{i}', y)}", key=f"dispc_{i}")
    sh.row("　合计", None, lambda y: "=" + "+".join(sh.ref(f"dispc_{i}", y) for i in range(N)),
           lambda y: "=" + "+".join(sh.ref(f"dispc_{i}", y) for i in range(N)), total=True)
    sh.row("8.7 附：处置账面净值（附注「出售」行，仅历史；处置原价−此值=处置转出的累计折旧）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", note, {y: V(disp[cls], y) for y in hist}, key=f"dispn_{i}")
    # 回填 8.2 / 8.3 预测公式里的行号占位
    for i in range(N):
        for y in fc:
            c = sh.ws.cell(sh.rows[f"a_disp_{i}"], sh.col[y])
            c.value = c.value.replace("{r}", str(sh.rows[f"a_disp_{i}"]))
            c2 = sh.ws.cell(sh.rows[f"a_rate_{i}"], sh.col[y])
            c2.value = c2.value.replace("{rate}", str(sh.rows[f"rate_{i}"]))
    sh.blank()
    life = P.get("useful_life", {})
    rem = S.get("remarks", {}).get("固定资产", [])
    sh.note_block("备注与说明", [
        "【勾稽关系】NBV = 原价 − 累计折旧 − 减值；合计NBV + 在建工程 = 资产负债表「物業、廠房及設備」；期末原价 = 期初 + 添置 + 在建工程转入 − 处置原价；本期计提 = 平均原价 × 折旧率。",
        "【口径说明】港股 PP&E 附注按账面净值滚动（年初净值→添置→轉讓→出售→折舊→年末净值），另给年初/年末成本与累计折旧。「出售」是处置资产的净值，处置原价不直接披露，本表由成本滚动倒算（8.6），处置转出的累计折旧 = 8.6 − 8.7。",
        "【可使用年期（会计政策）】" + "；".join(f"{k}：{v}" for k, v in life.items()) if life else "【可使用年期】见年报会计政策。",
        f"【预测假设说明】8.1 购置增长率默认 0%；8.2 处置率默认近3年平均；8.3 折旧率锁定近2年隐含值平均；8.4 在建工程转入联动在建工程表 7.2；减值默认延续。首列 {h0}A 为基准年（取自 {h0+1} 年报比较列），预测自 {fc[0] if fc else '-'}E 起。",
    ] + rem)
    sh.finalize()
    return sh


def build_rd(wb, S, hist, fc, unit):
    R = S["rd"]; note = R.get("note", "")
    sh = Sheet(wb, "资本化研发", hist, fc, unit)
    sh.header(f"{note} 研发支出 — 费用化 / 资本化 链路 ({hist[0]}-{fc[-1] if fc else hist[-1]})",
              f"链路：研发投入 → 费用化 → 研發開支 / 资本化 → 開發成本 → 转入无形资产 ｜ 单位：{unit}",
              "港股无「研发投入情况表」；费用化=利润表研發開支，资本化=无形资产附注「開發成本」添置（本公司为0时全部费用化）")
    h0 = hist[0]
    sh.section("一、研发投入拆分（港股按功能列报，仅可得研發開支总额及其中的折旧摊销）")
    sh.row("研發開支（利润表，全部费用化）", "合併全面收益表", {y: V(R["rd_expense"], y) for y in hist},
           lambda y: f"={sh.ref('rev', y)}*{sh.ref('a_ratio', y)}", key="rd")
    sh.row("　其中：物業、廠房及設備折旧（研发口径）", "PP&E附注折舊分析", {y: V(R.get("rd_dep", {}), y) for y in hist}, key="rd_dep")
    sh.row("　其中：无形资产摊销（研发口径）", "無形資產附注攤銷分析", {y: V(R.get("rd_amort", {}), y) for y in hist}, key="rd_am")
    if R.get("design_dev"):
        sh.row("附：設計及開發開支（集团按性质，含研发外委）", "附注 按性質劃分的開支", {y: V(R["design_dev"], y) for y in hist}, key="dd")
    sh.blank()
    sh.section("二、费用化 vs 资本化")
    sh.row("费用化研发支出", "=一.研發開支", lambda y: f"={sh.ref('rd', y)}", lambda y: f"={sh.ref('total', y)}-{sh.ref('cap', y)}", key="exp")
    sh.row("资本化研发支出（開發成本添置）", note or "無形資產附注", {y: V(R.get("capitalized", {}), y) for y in hist},
           lambda y: f"={sh.ref('total', y)}*{sh.ref('a_caprate', y)}", key="cap")
    sh.row("合计核对（=费用化+资本化）", None, lambda y: f"={sh.ref('exp', y)}+{sh.ref('cap', y)}", lambda y: f"={sh.ref('rd', y)}", key="total", total=True)
    sh.row("资本化率", None, lambda y: f'=IFERROR({sh.ref("cap", y)}/{sh.ref("total", y)},"")', {y: 0 for y in fc}, key="a_caprate", fmt=PCT1,
           reason="深红可调；公司政策为开发支出全部费用化时默认 0%")
    sh.blank()
    sh.section("三、開發成本（资产负债表科目）— 按项目")
    sh.row("開發成本期末余额（年报无此科目）", note or "無形資產附注", {y: V(R.get("dev_cost_balance", {}), y) for y in hist}, key="dc")
    sh.blank()
    sh.section("三-2、開發成本滚动表（历史+预测）")
    sh.row("期初余额", None, {h0: V(R.get("dev_cost_open", {}), h0)}, None, key="dc_open")
    sh.row("本年新增（内部开发资本化）", None, lambda y: f"={sh.ref('cap', y)}", lambda y: f"={sh.ref('cap', y)}", key="dc_add")
    sh.row("转入无形资产（−）", note or "無形資產附注", {y: -V(R.get("dev_to_intangible", {}), y) for y in hist},
           lambda y: f"=-({sh.ref('dc_open', y)}+0.5*{sh.ref('dc_add', y)})*{sh.ref('a_trrate', y)}", key="dc_tr")
    sh.row("计入当期损益（资本化失败转出）（−）", None, {y: 0 for y in hist}, {y: 0 for y in fc}, key="dc_pl", fc_color=RED)
    sh.row("期末余额", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['dc_open']}:{sh.L(y)}{sh.rows['dc_pl']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['dc_open']}:{sh.L(y)}{sh.rows['dc_pl']})", key="dc_end", total=True)
    sh.row("校验：期末公式 − 披露（应=0）", None, lambda y: f"={sh.ref('dc_end', y)}-{sh.ref('dc', y)}", check=True)
    sh.row("历史转入率 / 预测转入率假设", None, lambda y: f'=IFERROR(-{sh.ref("dc_tr", y)}/{sh.ref("dc_open", y)},"")', {y: 0 for y in fc}, key="a_trrate", fmt=PCT1)
    for y in hist[1:] + fc:
        sh.cell(sh.rows["dc_open"], sh.col[y], f"={sh.ref('dc_end', sh.prev(y))}", BLACK, fmt=NUM, align="right")
    sh.blank()
    sh.section("四、研发占收入比")
    sh.row("收益（营业收入）", "合併全面收益表", {y: V(R["revenue"], y) for y in hist},
           {y: f"={{prev}}*(1+{{g}})" for y in fc}, key="rev")
    sh.row("收入增速假设", None, lambda y: (f'=IFERROR({sh.ref("rev", y)}/{sh.ref("rev", sh.prev(y))}-1,"")' if y != h0 else None),
           {y: V(R.get("rev_growth_fc", {}), y, 0.2) for y in fc}, key="a_g", fmt=PCT1, reason="预测收入增速，深红可调（默认20%）")
    sh.row("研发费用率假设（研發開支 ÷ 收益）", None, lambda y: f'=IFERROR({sh.ref("rd", y)}/{sh.ref("rev", y)},"")',
           {y: f"=AVERAGE({sh.L(hist[-2])}{{r}}:{sh.L(hist[-1])}{{r}})" for y in fc}, key="a_ratio", fmt=PCT1, reason="预测默认锁定近2年研发费用率平均")
    for y in fc:
        c = sh.ws.cell(sh.rows["rev"], sh.col[y])
        c.value = c.value.replace("{prev}", sh.ref("rev", sh.prev(y))).replace("{g}", sh.ref("a_g", y))
        c2 = sh.ws.cell(sh.rows["a_ratio"], sh.col[y])
        c2.value = c2.value.replace("{r}", str(sh.rows["a_ratio"]))
    sh.blank()
    rem = S.get("remarks", {}).get("资本化研发", [])
    sh.note_block("五、备注", [
        "【两步流程】研发投入先按会计政策判断费用化/资本化；资本化部分进「開發成本」，达到可使用状态转入无形资产并开始摊销。",
        "【预测逻辑】研發開支 = 预测收益 × 研发费用率（四）；资本化 = 合计 × 资本化率（二，深红）；转入无形资产 = (期初+0.5×新增)×转入率。",
        "【联动】三-2「转入无形资产」行是无形资产表 9.3 的引用源。",
    ] + rem)
    sh.finalize()
    return sh


def build_intangibles(wb, S, hist, fc, unit, rd):
    I = S["intangibles"]; classes = I["classes"]; M = len(classes); note = I.get("note", "無形資產附注")
    indefinite = set(I.get("indefinite", []))
    sh = Sheet(wb, "无形资产", hist, fc, unit)
    sh.header(f"{note} 无形资产 — {M} 类完整勾稽 ({hist[0]}-{fc[-1] if fc else hist[-1]})",
              f"单位：{unit} ｜ 颜色：深蓝=年报实际 ｜ 深红=可调假设 ｜ 黑=公式 ｜ 黄底=合计 ｜ 绿底=校验",
              f"{M} 类资产：" + " / ".join(classes) + (f" ｜ 无限年期不摊销：{'、'.join(indefinite)}" if indefinite else ""))
    h0 = hist[0]
    cost = I["cost"]; acc = I["accamort"]; imp = I.get("impair", {}); am = I["amort"]; add = I["additions"]; disp = I["disposal_nbv"]
    n2 = hist[-2:]; n3 = hist[-3:]
    sh.section(f"一、期末账面净值（NBV）— {M} 类汇总")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(cost[cls], y) - V(acc[cls], y) - V(imp.get(cls, {}), y) for y in hist},
               lambda y, i=i: f"={sh.ref(f'cost_{i}', y)}-{sh.ref(f'acc_{i}', y)}-{sh.ref(f'imp_{i}', y)}", key=f"nbv_{i}")
    sh.row("合计 NBV", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['nbv_0']}:{sh.L(y)}{sh.rows[f'nbv_{M-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['nbv_0']}:{sh.L(y)}{sh.rows[f'nbv_{M-1}']})", key="nbv_tot", total=True)
    sh.row("BS 披露 無形資產", "合併資產負債表", {y: V(I["bs_nbv"], y) for y in hist}, key="bs")
    sh.row("校验：模型 − BS（应=0）", None, lambda y: f"={sh.ref('nbv_tot', y)}-{sh.ref('bs', y)}", check=True)
    sh.blank()
    sh.section(f"二、期末账面原值 — {M} 类汇总（预测引用十.原价滚动表）")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(cost[cls], y) for y in hist}, lambda y, i=i: f"={sh.ref(f'rc_end_{i}', y)}", key=f"cost_{i}")
    sh.row("合计原价", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['cost_0']}:{sh.L(y)}{sh.rows[f'cost_{M-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['cost_0']}:{sh.L(y)}{sh.rows[f'cost_{M-1}']})", key="cost_tot", total=True)
    sh.blank()
    sh.section(f"三、期末累计摊销 — {M} 类汇总（预测引用十一.滚动表）")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(acc[cls], y) for y in hist}, lambda y, i=i: f"={sh.ref(f'ra_end_{i}', y)}", key=f"acc_{i}")
    sh.row("合计累计摊销", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['acc_0']}:{sh.L(y)}{sh.rows[f'acc_{M-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['acc_0']}:{sh.L(y)}{sh.rows[f'acc_{M-1}']})", key="acc_tot", total=True)
    sh.blank()
    sh.section(f"四、期末减值准备 — {M} 类汇总")
    for i, cls in enumerate(classes):
        sh.row(cls, note, {y: V(imp.get(cls, {}), y) for y in hist}, lambda y, i=i: f"={sh.ref(f'imp_{i}', sh.prev(y))}", key=f"imp_{i}")
    sh.row("合计减值", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['imp_0']}:{sh.L(y)}{sh.rows[f'imp_{M-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['imp_0']}:{sh.L(y)}{sh.rows[f'imp_{M-1}']})", total=True)
    sh.blank()
    sh.section(f"五、本期计提摊销 — {M} 类汇总（附注「攤銷費用」行）")
    for i, cls in enumerate(classes):
        if cls in indefinite:
            f = lambda y, i=i: 0
        else:
            f = lambda y, i=i: (f"=MIN(({sh.ref(f'cost_{i}', y)}+{sh.ref(f'cost_{i}', sh.prev(y))})/2*{sh.ref(f'a_rate_{i}', y)},"
                                f"MAX(0,{sh.ref(f'nbv_{i}', sh.prev(y))}+({sh.ref(f'cost_{i}', y)}-{sh.ref(f'cost_{i}', sh.prev(y))})"
                                f"-({sh.ref(f'imp_{i}', y)}-{sh.ref(f'imp_{i}', sh.prev(y))})))")
        sh.row(cls, note, {y: V(am[cls], y) for y in hist}, f, key=f"am_{i}")
    sh.row("合计计提摊销", None, lambda y: f"=SUM({sh.L(y)}{sh.rows['am_0']}:{sh.L(y)}{sh.rows[f'am_{M-1}']})",
           lambda y: f"=SUM({sh.L(y)}{sh.rows['am_0']}:{sh.L(y)}{sh.rows[f'am_{M-1}']})", key="am_tot", total=True)
    abf = I.get("amort_by_function", {})
    if abf:
        sh.row("附：摊销按功能分布（附注披露，仅历史）", None, None, None)
        for lab, d in abf.items():
            sh.row(f"　{lab}", note, {y: V(d, y) for y in hist}, key=f"abf_{lab}")
        sh.row("　合计（应=五.合计）", None, lambda y: "=" + "+".join(sh.ref(f"abf_{lab}", y) for lab in abf), key="abf_tot", total=True)
        sh.row("　校验（应=0）", None, lambda y: f"={sh.ref('abf_tot', y)}-{sh.ref('am_tot', y)}", check=True)
    sh.blank()
    sh.section("六、隐含摊销率反推（本期计提 ÷ 平均原价）")
    for i, cls in enumerate(classes):
        sh.row(f"{cls} 摊销率", None,
               lambda y, i=i: (f'=IFERROR({sh.ref(f"am_{i}", y)}/(({sh.ref(f"cost_{i}", y)}+{sh.ref(f"cost_{i}", sh.prev(y))})/2),"")' if y != h0 else None),
               lambda y, i=i: f"={sh.ref(f'a_rate_{i}', y)}", key=f"rate_{i}", fmt=PCT2)
    sh.blank()
    sh.section("七、资产新旧度（累计摊销 ÷ 原价合计）")
    sh.row("整体新旧度", "老化指标", lambda y: f'=IFERROR({sh.ref("acc_tot", y)}/{sh.ref("cost_tot", y)},"")',
           lambda y: f'=IFERROR({sh.ref("acc_tot", y)}/{sh.ref("cost_tot", y)},"")', fmt=PCT1)
    sh.blank()
    sh.section("八、内部研发形成的无形资产占比 — 年报未披露（公司无資本化開發成本），本节省略")
    sh.blank()
    sh.section("九、预测假设输入（深红色为可调）")
    sh.row("9.1 购置增长率（历史=购置YoY反推）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", None,
               lambda y, i=i: (f'=IFERROR({sh.ref(f"buy_{i}", y)}/{sh.ref(f"buy_{i}", sh.prev(y))}-1,"")' if y != h0 else None),
               {y: 0 for y in fc}, key=f"a_grow_{i}", fmt=PCT1, reason="默认 0%" if i == 0 else None)
    sh.row(f"9.2 摊销率（默认锁定 {n2[0]}-{n2[-1]} 平均；无限年期类别=0）", None, None, None)
    for i, cls in enumerate(classes):
        if cls in indefinite:
            sh.row(f"　{cls}", None, None, {y: 0 for y in fc}, key=f"a_rate_{i}", fmt=PCT2, reason="无限可使用年期，不摊销，每年做减值测试")
        else:
            sh.row(f"　{cls}", None, None, {y: f"=AVERAGE({sh.L(n2[0])}{{rate}}:{sh.L(n2[-1])}{{rate}})" for y in fc}, key=f"a_rate_{i}", fmt=PCT2)
    sh.row("9.3 内部研发资本化转入（联动资本化研发表三-2「转入无形资产」）", None, None, None)
    rd_cls = I.get("rd_class")
    for i, cls in enumerate(classes):
        if cls == rd_cls:
            sh.row(f"　{cls}", "资本化研发!三-2", {y: V(I.get("rd_capitalized", {}).get(cls, {}), y) for y in hist},
                   lambda y: f"=-资本化研发!{rd.L(y)}{rd.rows['dc_tr']}", key=f"rdin_{i}")
        else:
            sh.row(f"　{cls}", None, {y: 0 for y in hist}, {y: 0 for y in fc}, key=f"rdin_{i}", hist_color=BLACK, fc_color=BLACK)
    sh.row("　合计内部研发资本化", None, lambda y: "=" + "+".join(sh.ref(f"rdin_{i}", y) for i in range(M)),
           lambda y: "=" + "+".join(sh.ref(f"rdin_{i}", y) for i in range(M)), total=True)
    sh.row("9.4 购置（历史=附注「添置」；预测=上年×(1+9.1)）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", note, {y: V(add[cls], y) for y in hist},
               lambda y, i=i: f"={sh.ref(f'buy_{i}', sh.prev(y))}*(1+{sh.ref(f'a_grow_{i}', y)})", key=f"buy_{i}")
    sh.row("9.5 处置原价（历史=倒算；预测=近3年平均）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", "倒算",
               lambda y, i=i: (f"={sh.ref(f'cost_{i}', sh.prev(y))}+{sh.ref(f'buy_{i}', y)}+{sh.ref(f'rdin_{i}', y)}+{sh.ref(f'comb_{i}', y)}-{sh.ref(f'cost_{i}', y)}" if y != h0 else 0),
               {y: f"=AVERAGE({sh.L(n3[0])}{{r}}:{sh.L(n3[-1])}{{r}})" for y in fc}, key=f"dispc_{i}")
    sh.row("9.6 附：处置账面净值（附注「出售」行，仅历史）", None, None, None)
    for i, cls in enumerate(classes):
        sh.row(f"　{cls}", note, {y: V(disp.get(cls, {}), y) for y in hist}, key=f"dispn_{i}")
    sh.blank()
    sh.section(f"十、原价滚动表（{M} 类，公式驱动）")
    for i, cls in enumerate(classes):
        sh.row(f"【{cls}】期初原价", None, {h0: V(I.get("cost_open_first", {}).get(cls, {}), h0, None)}, None, key=f"rc_open_{i}")
        sh.row("　购置（=9.4）", None, lambda y, i=i: f"={sh.ref(f'buy_{i}', y)}", lambda y, i=i: f"={sh.ref(f'buy_{i}', y)}")
        sh.row("　内部研发资本化（=9.3）", None, lambda y, i=i: f"={sh.ref(f'rdin_{i}', y)}", lambda y, i=i: f"={sh.ref(f'rdin_{i}', y)}")
        sh.row("　企业合并增加", note, {y: V(I.get("combination", {}).get(cls, {}), y) for y in hist}, {y: 0 for y in fc}, key=f"comb_{i}")
        sh.row("　处置（−）（=9.5）", None, lambda y, i=i: f"=-{sh.ref(f'dispc_{i}', y)}", lambda y, i=i: f"=-{sh.ref(f'dispc_{i}', y)}")
        sh.row("　外币折算差额", note, {y: 0 for y in hist}, {y: 0 for y in fc}, hist_color=BLACK, fc_color=BLACK)
        sh.row("　期末原价", None, lambda y, i=i: f"=SUM({sh.L(y)}{sh.rows[f'rc_open_{i}']}:{sh.L(y)}{sh.rows[f'rc_open_{i}'] + 5})",
               lambda y, i=i: f"=SUM({sh.L(y)}{sh.rows[f'rc_open_{i}']}:{sh.L(y)}{sh.rows[f'rc_open_{i}'] + 5})", key=f"rc_end_{i}", total=True)
        sh.row("　校验：vs 二.期末原价（历史应=0）", None, lambda y, i=i: f"={sh.ref(f'rc_end_{i}', y)}-{sh.ref(f'cost_{i}', y)}", check=True)
        for y in hist[1:] + fc:
            sh.cell(sh.rows[f"rc_open_{i}"], sh.col[y], f"={sh.ref(f'rc_end_{i}', sh.prev(y))}", BLACK, fmt=NUM, align="right")
        if sh.ws.cell(sh.rows[f"rc_open_{i}"], sh.col[h0]).value is None:
            # 首年期初 = 首年期末 − 首年购置 − 研发转入 − 合并 + 处置（由披露倒推）
            sh.cell(sh.rows[f"rc_open_{i}"], sh.col[h0],
                    f"={sh.ref(f'cost_{i}', h0)}-{sh.ref(f'buy_{i}', h0)}-{sh.ref(f'rdin_{i}', h0)}-{sh.ref(f'comb_{i}', h0)}+{sh.ref(f'dispc_{i}', h0)}",
                    BLACK, fmt=NUM, align="right")
    sh.blank()
    sh.section(f"十一、累计摊销滚动表（{M} 类）")
    for i, cls in enumerate(classes):
        sh.row(f"【{cls}】期初累计摊销", None, {h0: V(I.get("acc_open_first", {}).get(cls, {}), h0, None)}, None, key=f"ra_open_{i}")
        sh.row("　本期计提（=五）", None, lambda y, i=i: f"={sh.ref(f'am_{i}', y)}", lambda y, i=i: f"={sh.ref(f'am_{i}', y)}")
        sh.row("　处置/外币减少（−）（历史倒算；预测=−(处置原价−处置净值率)≈−处置原价）", "倒算",
               lambda y, i=i: (f"={sh.ref(f'acc_{i}', y)}-{sh.ref(f'ra_open_{i}', y)}-{sh.ref(f'am_{i}', y)}" if y != h0 else 0),
               lambda y, i=i: f"=-{sh.ref(f'dispc_{i}', y)}", key=f"ra_dec_{i}")
        sh.row("　期末累计摊销", None, lambda y, i=i: f"=SUM({sh.L(y)}{sh.rows[f'ra_open_{i}']}:{sh.L(y)}{sh.rows[f'ra_dec_{i}']})",
               lambda y, i=i: f"=SUM({sh.L(y)}{sh.rows[f'ra_open_{i}']}:{sh.L(y)}{sh.rows[f'ra_dec_{i}']})", key=f"ra_end_{i}", total=True)
        sh.row("　校验：vs 三.期末累计摊销（历史应=0）", None, lambda y, i=i: f"={sh.ref(f'ra_end_{i}', y)}-{sh.ref(f'acc_{i}', y)}", check=True)
        for y in hist[1:] + fc:
            sh.cell(sh.rows[f"ra_open_{i}"], sh.col[y], f"={sh.ref(f'ra_end_{i}', sh.prev(y))}", BLACK, fmt=NUM, align="right")
        if sh.ws.cell(sh.rows[f"ra_open_{i}"], sh.col[h0]).value is None:
            sh.cell(sh.rows[f"ra_open_{i}"], sh.col[h0], f"={sh.ref(f'acc_{i}', h0)}-{sh.ref(f'am_{i}', h0)}", BLACK, fmt=NUM, align="right")
    # 回填占位
    for i in range(M):
        for y in fc:
            c = sh.ws.cell(sh.rows[f"a_rate_{i}"], sh.col[y])
            if isinstance(c.value, str):
                c.value = c.value.replace("{rate}", str(sh.rows[f"rate_{i}"]))
            c2 = sh.ws.cell(sh.rows[f"dispc_{i}"], sh.col[y])
            if isinstance(c2.value, str):
                c2.value = c2.value.replace("{r}", str(sh.rows[f"dispc_{i}"]))
    sh.blank()
    rem = S.get("remarks", {}).get("无形资产", [])
    sh.note_block("十二、备注与说明", [
        "【勾稽关系】NBV = 原价 − 累计摊销 − 减值 = 资产负债表「無形資產」；期末原价 = 期初 + 购置 + 内部研发资本化 + 企业合并 − 处置原价；期末累计摊销 = 期初 + 本期计提 − 处置转出。",
        "【口径说明】港股附注按净值滚动，「出售」为处置净值；处置原价与处置转出的累计摊销由成本/累计摊销滚动倒算（9.5、十一）。",
        "【预测假设说明】9.1 购置增长率默认 0%；9.2 摊销率锁定近2年平均、无限年期类别为 0；9.3 内部研发资本化联动资本化研发表；9.5 处置默认近3年平均；减值默认延续。摊销封顶：预测计提不超过上年NBV+原价增量，防止摊成负NBV。",
    ] + rem)
    sh.finalize()
    return sh


def build_grants(wb, S, hist, unit, company):
    G = S["grants"]
    sh = Sheet(wb, "政府补助分析", hist, [], unit)
    sh.header(f"{company} 政府补助全景分析 ({hist[0]}-{hist[-1]})",
              f"单位：{unit}（汇总区为亿元） ｜ 数据来源：各年度年报附注（其他收入 / 遞延收入 / 合併現金流量表） ｜ 制表：广发汽车",
              "港股口径：计入损益的政府补贴全部在「其他收入」；与资产相关部分先入「遞延收入」再按资产年期摊销转入其他收入；无营业外收入/非经常性损益口径")
    h0 = hist[0]
    oi = G["other_income_grant"]; da = G["deferred_amort"]; de = G["deferred_end"]; rec = G["grants_received"]; cf = G.get("cf_grants_received", rec)
    # 核心指标汇总（亿元）：公式引用明细区，先占位后回填
    sh.section("核心指标汇总（亿元）")
    sum_rows = {}
    for key, lab in [("s_oi", "其他收入·政府补贴合计"), ("s_pl", "计入当期损益的政府补助（=其他收入政府补贴）"),
                     ("s_asset", "　其中：与资产相关（递延收入摊销）"), ("s_inc", "　其中：与收益相关（当期直接计入）"),
                     ("s_de", "递延收入期末余额"), ("s_new", "　当年新增（收到，与资产相关）"), ("s_out", "　当年转出（摊销）"),
                     ("s_cf", "政府补助现金流入（投资活动·收購非流動資產收取的政府補貼）")]:
        sum_rows[key] = sh.row(lab, "=明细区/100000", None, None, key=key, fmt="0.00")
    sh.blank()
    sh.section(f"一、其他收入拆分（{unit}）")
    sh.row("与日常活动相关的政府补贴（其他收入）", "附注 其他收入", {y: V(oi, y) for y in hist}, key="oi")
    sh.row("其他（港股其他收入仅政府补贴一项）", "附注 其他收入", {y: 0 for y in hist}, key="oi_oth", hist_color=BLACK)
    sh.row("合计", None, lambda y: f"={sh.ref('oi', y)}+{sh.ref('oi_oth', y)}", key="oi_tot", total=True)
    sh.row("政府补贴 YoY", None, lambda y: (f'=IF({sh.ref("oi", sh.prev(y))}=0,"",{sh.ref("oi", y)}/{sh.ref("oi", sh.prev(y))}-1)' if y != h0 else None), fmt=PCT1)
    sh.row("政府补贴占其他收入比重", None, lambda y: f'=IF({sh.ref("oi_tot", y)}=0,"",{sh.ref("oi", y)}/{sh.ref("oi_tot", y)})', fmt=PCT1)
    sh.blank()
    sh.section(f"二、计入当期损益的政府补助分类（{unit}）")
    sh.row("与资产相关 — 递延收入摊销计入其他收入", "附注 遞延收入 / 现金流量资料「政府補貼攤銷」", {y: V(da, y) for y in hist}, key="pl_asset")
    sh.row("与收益相关 — 当期直接计入其他收入（倒挤）", "=一.合计 − 资产相关", lambda y: f"={sh.ref('oi', y)}-{sh.ref('pl_asset', y)}", key="pl_inc")
    sh.row("与收益相关 — 冲减财务费用 / 营业外收入", "港股无此口径", {y: 0 for y in hist}, key="pl_fin", hist_color=BLACK)
    sh.row("合计", None, lambda y: f"={sh.ref('pl_asset', y)}+{sh.ref('pl_inc', y)}+{sh.ref('pl_fin', y)}", key="pl_tot", total=True)
    sh.row("与资产相关占比", None, lambda y: f'=IF({sh.ref("pl_tot", y)}=0,"",{sh.ref("pl_asset", y)}/{sh.ref("pl_tot", y)})', fmt=PCT1)
    sh.row("与收益相关占比", None, lambda y: f'=IF({sh.ref("pl_tot", y)}=0,"",{sh.ref("pl_inc", y)}/{sh.ref("pl_tot", y)})', fmt=PCT1)
    sh.blank()
    sh.section(f"三、递延收入（政府補貼）变动（{unit}）")
    sh.row("期初余额", "上年期末", {h0: V(G.get("deferred_open_first", {}), h0, None)}, None, key="de_open")
    sh.row("本年新增（收到与资产相关补贴）", "附注 遞延收入", {y: V(rec, y) for y in hist}, key="de_new")
    sh.row("本年退回", "附注", {y: 0 for y in hist}, key="de_ret", hist_color=BLACK)
    sh.row("本年转出（摊销计入其他收入）", "附注 遞延收入 / 现金流量资料", {y: V(da, y) for y in hist}, key="de_out")
    sh.row("其他变动（倒挤：披露期末 − 滚动期末，年报未解释）", "倒挤", lambda y: f"={sh.ref('de_disc', y)}-({sh.ref('de_open', y)}+{sh.ref('de_new', y)}-{sh.ref('de_ret', y)}-{sh.ref('de_out', y)})", key="de_oth")
    sh.row("期末余额（滚动公式，不含其他变动）", None, lambda y: f"={sh.ref('de_open', y)}+{sh.ref('de_new', y)}-{sh.ref('de_ret', y)}-{sh.ref('de_out', y)}", key="de_end", total=True)
    sh.row("期末余额（年报披露值）", "附注 遞延收入 / 資產負債表", {y: V(de, y) for y in hist}, key="de_disc")
    sh.row("差异校验：滚动 − 披露（应=0；非0即为「其他变动」）", None, lambda y: f"={sh.ref('de_end', y)}-{sh.ref('de_disc', y)}", key="de_chk", check=True)
    sh.row("期末 YoY", None, lambda y: (f'=IF({sh.ref("de_disc", sh.prev(y))}=0,"",{sh.ref("de_disc", y)}/{sh.ref("de_disc", sh.prev(y))}-1)' if y != h0 else None), fmt=PCT1)
    sh.row("当年转出率（转出 ÷ 期初）", None, lambda y: f'=IF({sh.ref("de_open", y)}=0,"",{sh.ref("de_out", y)}/{sh.ref("de_open", y)})', fmt=PCT1)
    for y in hist[1:]:
        sh.cell(sh.rows["de_open"], sh.col[y], f"={sh.ref('de_disc', sh.prev(y))}", BLACK, fmt=NUM, align="right")
    if sh.ws.cell(sh.rows["de_open"], sh.col[h0]).value is None:
        sh.cell(sh.rows["de_open"], sh.col[h0], f"={sh.ref('de_disc', h0)}-{sh.ref('de_new', h0)}+{sh.ref('de_out', h0)}", BLACK, fmt=NUM, align="right")
        sh.cell(sh.rows["de_open"], 3, "倒算（首年）")
    sh.blank()
    sh.section("四、非经常性损益中的政府补助 — 港股无此口径，本节省略")
    sh.blank()
    sh.section("五、增值税加计抵减等 — 港股其他收入未含此项，本节省略")
    sh.blank()
    sh.section(f"六、政府补助现金流口径（{unit}）")
    sh.row("收到的政府补贴现金（投资活动：就收購非流動資產收取的政府補貼）", "合併現金流量表", {y: V(cf, y) for y in hist}, key="cf")
    sh.row("递延收入新增（参考，=三.新增）", None, lambda y: f"={sh.ref('de_new', y)}")
    sh.row("计入损益（参考，=二.合计）", None, lambda y: f"={sh.ref('pl_tot', y)}")
    sh.row("现金收到 ÷ 计入损益 倍数", None, lambda y: f'=IF({sh.ref("pl_tot", y)}=0,"",{sh.ref("cf", y)}/{sh.ref("pl_tot", y)})', fmt="0.00")
    sh.blank()
    # 回填汇总区
    link = {"s_oi": "oi_tot", "s_pl": "pl_tot", "s_asset": "pl_asset", "s_inc": "pl_inc", "s_de": "de_disc", "s_new": "de_new", "s_out": "de_out", "s_cf": "cf"}
    for k, src in link.items():
        for y in hist:
            sh.cell(sum_rows[k], sh.col[y], f"={sh.ref(src, y)}/100000", BLACK, fmt="0.00", align="right")
    rem = S.get("remarks", {}).get("政府补助分析", [])
    sh.note_block("备注与说明", [
        "【口径】港股「其他收入·政府補貼」= 当期计入损益的全部政府补助，其中与资产相关部分来自递延收入按资产年期直线摊销（现金流量资料「政府補貼攤銷」行），其余为与收益相关的当期补贴。",
        "【递延收入】期末 = 期初 + 本年收到（与资产相关）− 本年摊销；滚动与披露的差额记为「其他变动」并在校验行显示，年报未解释的年份在备注点名。",
        "【现金流】投资活动「就收購非流動資產收取的政府補貼」= 与资产相关补贴的现金流入，应与递延收入新增基本一致（个别年份差异为与成本相关的小额补贴走经营活动）。",
    ] + rem)
    sh.finalize()
    return sh


def build(spec_path):
    S = json.load(open(spec_path, encoding="utf-8"))
    hist = [int(y) for y in S["hist_years"]]; fc = [int(y) for y in S.get("fc_years", [])]
    unit = S.get("unit", "千元"); company = S["company"]
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    cip = build_cip(wb, S, hist, fc, unit)
    ppe = build_ppe(wb, S, hist, fc, unit, cip)
    rd = build_rd(wb, S, hist, fc, unit)
    intan = build_intangibles(wb, S, hist, fc, unit, rd)
    grants = build_grants(wb, S, hist, unit, company)
    # 排序：固定资产 / 在建工程 / 无形资产 / 资本化研发 / 政府补助分析
    order = ["固定资产", "在建工程", "无形资产", "资本化研发", "政府补助分析"]
    wb._sheets = [wb[n] for n in order]
    out = S["output"]
    wb.save(out)
    print(json.dumps({"status": "ok", "output": out, "sheets": order}, ensure_ascii=False))
    return out


SPEC_DOC = """
spec.json（金额单位与年报一致，默认千元；年份键可为字符串）：
{
 "company": "零跑汽车", "unit": "千元", "output": ".../XX_附注科目分析_2022-2025.xlsx",
 "hist_years": [2021,...,2025], "fc_years": [2026,2027,2028],
 "ppe": {"note": "附注14", "classes": [...], "cost": {cls:{y:v}}, "accdep": {...}, "impair": {...可选},
         "additions": {cls:{y}}, "transfer_in": {cls:{y}}, "disposal_nbv": {cls:{y}}, "depreciation": {cls:{y}},
         "bs_nbv": {y: v 含在建工程}, "dep_by_function": {label:{y}}, "useful_life": {cls: "20年"},
         "cip": {"nbv": {y}, "additions": {y}, "transfer_out": {y 正数}, "open_first": {首年: v}}},
 "intangibles": {"note": "附注16", "classes": [...], "cost","accamort","impair","additions","disposal_nbv","amort",
                 "bs_nbv", "indefinite": [不摊销类别], "amort_by_function": {...}, "combination": {cls:{y}}, "rd_class": 承接开发成本的类别},
 "rd": {"rd_expense": {y}, "revenue": {y}, "rd_dep": {y}, "rd_amort": {y}, "design_dev": {y}, "capitalized": {y}, "dev_cost_balance": {y},
        "dev_to_intangible": {y}, "rev_growth_fc": {y: 0.2}},
 "grants": {"other_income_grant": {y}, "deferred_amort": {y}, "deferred_end": {y}, "grants_received": {y}, "cf_grants_received": {y},
            "deferred_open_first": {首年: v 可选}},
 "remarks": {"固定资产": [...], "在建工程": [...], "无形资产": [...], "资本化研发": [...], "政府补助分析": [...]}
}
"""

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__ + SPEC_DOC); sys.exit(1)
    build(sys.argv[1])
