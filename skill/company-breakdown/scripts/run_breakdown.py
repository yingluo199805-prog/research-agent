#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公司财报拆分总控驱动：按一家公司的 company.json 依次跑
    ① 收入分配拆分（A股 run_v2_batch / 港股 run_h_share_batch，JSON 配置）
    ② 年报附注截图底稿（build_digao，可选）
    ③ 附注五表并合并进收入分配工作簿（build_hk_notes_model，merge_into → 目录页）
    ④ formulas 重算自检（recalc_check）
并打印验收清单。数字的抽取与 spec 的填写仍由分析师完成（见 SKILL.md 第 2-4 步），本脚本只负责确定性的生成与校验。

用法:
    python3 run_breakdown.py <company.json> [--skip-split] [--skip-digao] [--skip-notes] [--only-check]

company.json（路径相对该文件所在目录）:
{
  "company": "小米集团", "short": "小米", "market": "H",              # H = 港股(HKFRS/IFRS/US GAAP) ; A = A股
  "revenue_split": {"engine": "h_share", "config": "revenue_split.json"},   # engine: h_share | a_share
  "notes_spec": "notes_spec.json",                                    # build_hk_notes_model 的 spec（港股）；A 股按 references 手工建表时可省略
  "digao_spec": "研究底稿/digao_spec.json",                            # 可选
  "combined_output": "小米_财报拆分_2018-2025.xlsx"                     # 合并版产物名（写进 notes_spec.output / merge_into）
}
"""
import sys, os, json, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
ENG = os.path.join(HERE, "..", "engines")
RS = os.path.join(ENG, "revenue-split", "scripts")
NA = os.path.join(ENG, "notes-analysis", "scripts")
sys.path.insert(0, RS)
import common  # noqa: E402  (REVENUE_SPLIT_BASE / --base)


def run(cmd, cwd):
    print("\n$", " ".join(os.path.basename(c) if i == 1 else c for i, c in enumerate(cmd)))
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    out = "\n".join(l for l in (r.stdout + r.stderr).splitlines() if "Warning" not in l and "warn(" not in l and "it/s" not in l)
    print(out)
    if r.returncode != 0:
        raise SystemExit(f"step failed: {cmd[1]}")
    return out


def main():
    args = sys.argv[1:]
    if not args:
        print(__doc__); sys.exit(1)
    if "--base" in args:
        i = args.index("--base"); common.set_base(args[i + 1]); del args[i:i + 2]
    spec_path = os.path.abspath(args[0]); cdir = os.path.dirname(spec_path)
    S = json.load(open(spec_path, encoding="utf-8"))
    P = lambda rel: rel if os.path.isabs(rel) else os.path.join(cdir, rel)
    skip = {a for a in args[1:] if a.startswith("--")}
    short = S["short"]; combined = P(S["combined_output"])
    print(f"===== {S['company']} ({S['market']}股) 财报拆分 =====\n根目录 {common.BASE}\n公司目录 {cdir}")

    # ① 收入分配拆分
    rs = S["revenue_split"]; split_xlsx = None
    if rs["engine"] == "h_share":
        cfg = json.load(open(P(rs["config"]), encoding="utf-8"))
        split_xlsx = cfg["outfile"] if os.path.isabs(cfg["outfile"]) else os.path.join(common.BASE, cfg["outfile"])
        if "--skip-split" not in skip and "--only-check" not in skip:
            run([sys.executable, os.path.join(RS, "run_h_share_batch.py"), "--config", P(rs["config"])], RS)
    elif rs["engine"] == "a_share":
        cfg = json.load(open(P(rs["config"]), encoding="utf-8"))
        split_xlsx = os.path.join(common.BASE, cfg["outfile"])
        if "--skip-split" not in skip and "--only-check" not in skip:
            run([sys.executable, os.path.join(RS, "run_v2_batch.py"), "--config", P(rs["config"])], RS)
    else:
        raise SystemExit("revenue_split.engine 必须是 h_share 或 a_share")
    if not os.path.exists(split_xlsx):
        raise SystemExit(f"收入分配工作簿不存在：{split_xlsx}")

    # ② 截图底稿
    if S.get("digao_spec") and "--skip-digao" not in skip and "--only-check" not in skip:
        run([sys.executable, os.path.join(NA, "build_digao.py"), P(S["digao_spec"])], NA)

    # ③ 附注五表 → 合并
    if S.get("notes_spec"):
        ns_path = P(S["notes_spec"]); ns = json.load(open(ns_path, encoding="utf-8"))
        changed = False
        if ns.get("merge_into") != split_xlsx:
            ns["merge_into"] = split_xlsx; changed = True
        if ns.get("output") != combined:
            ns["output"] = combined; changed = True
        if changed:
            json.dump(ns, open(ns_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            print(f"notes_spec 已写入 merge_into={split_xlsx}\n            output={combined}")
        if "--skip-notes" not in skip and "--only-check" not in skip:
            run([sys.executable, os.path.join(NA, "build_hk_notes_model.py"), ns_path], NA)
        # ④ 自检
        out = run([sys.executable, os.path.join(NA, "recalc_check.py"), combined, "--sheets", "固定资产,在建工程,无形资产,资本化研发,政府补助分析"], NA)
        errs = [l for l in out.splitlines() if l.startswith("公式错误") and not l.startswith("公式错误: 0")]
        blanks = [l for l in out.splitlines() if l.startswith("预测列有公式但算不出值") and not l.startswith("预测列有公式但算不出值: 0")]
        nonzero = [l for l in out.splitlines() if "<-- 非零" in l]
    else:
        errs = blanks = nonzero = []
        print("（未提供 notes_spec：A 股附注五表按 engines/notes-analysis/references 手工建表后，用 recalc_check.py 单独自检）")

    # 验收清单
    print("\n===== 验收清单 =====")
    print(f"[{'x' if os.path.exists(split_xlsx) else ' '}] 收入分配拆分工作簿：{split_xlsx}")
    print(f"[{'x' if S.get('notes_spec') and os.path.exists(combined) else ' '}] 合并版工作簿（目录+三表+底稿+收入分配+附注五表）：{combined}")
    print(f"[{'x' if not errs else ' '}] 公式错误 = 0" + ("" if not errs else "  <-- " + "; ".join(errs)))
    print(f"[{'x' if not blanks else ' '}] 预测列无空值" + ("" if not blanks else "  <-- " + "; ".join(blanks)))
    print(f"[{'x' if not nonzero else '!'}] 校验行归零" + ("" if not nonzero else f"  <-- {len(nonzero)} 行非零，须在该表备注里逐条点名原因：\n      " + "\n      ".join(l.strip()[:120] for l in nonzero)))
    print("[ ] 人工：建模值 vs 底稿截图抽样比对（每表 3 个数）；比率区间（折旧率/摊销率/转固率/供应商占比）；备注区写了特殊事项")
    print("[ ] 人工：收入分配表 供应商占比合理（整车 62%-85%，硬件/互联网 75%-88%）、员工/折摊趋势能讲通")


if __name__ == "__main__":
    main()
