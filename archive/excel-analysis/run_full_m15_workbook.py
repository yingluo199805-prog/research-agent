# -*- coding: utf-8
"""Orchestrate full workbook M1-5 fill + user layout + formatting."""

from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

import pythoncom
import win32com.client as win32

from relayout_m15_sheet import ALL_SHEETS, relayout_sheet
from update_workbook_m15_fixed import update_sheet


def replace_sheet_from_template(output_path: Path, template_path: Path, sheet_name: str) -> None:
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        out_wb = excel.Workbooks.Open(str(output_path.resolve()))
        tpl_wb = excel.Workbooks.Open(str(template_path.resolve()), ReadOnly=True)
        tpl_ws = tpl_wb.Worksheets(sheet_name)
        out_ws = out_wb.Worksheets(sheet_name)
        tpl_ws.Cells.Copy()
        out_ws.Cells.Clear()
        out_ws.Range("A1").PasteSpecial(Paste=-4104)
        excel.CutCopyMode = False
        out_wb.Save()
        out_wb.Close(SaveChanges=True)
        tpl_wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cache", required=True, type=Path)
    parser.add_argument("--base-workbook", required=True, type=Path)
    parser.add_argument("--template-workbook", required=True, type=Path)
    parser.add_argument("--output-workbook", required=True, type=Path)
    parser.add_argument(
        "--europe-workbook",
        type=Path,
        default=None,
        help="Workbook with finished 欧洲-核心数据; used as output seed when provided",
    )
    args = parser.parse_args()

    cache = json.loads(args.cache.read_text(encoding="utf-8"))
    out = args.output_workbook.resolve()
    base = args.base_workbook.resolve()
    template = args.template_workbook.resolve()
    europe_done = args.europe_workbook.resolve() if args.europe_workbook else None

    if out.exists():
        out.unlink()
    shutil.copy2(europe_done or base, out)

    for sheet in ALL_SHEETS:
        print(f"processing {sheet}...", flush=True)
        if sheet == "欧洲-核心数据":
            if europe_done:
                print("skip 欧洲-核心数据 (already relayouted in seed workbook)", flush=True)
                continue
            replace_sheet_from_template(out, template, sheet)
            relayout_sheet(out, sheet, template, preserve_summary=True)
            continue

        update_sheet(cache, base, out, sheet, in_place=True, skip_if_m15=False)
        relayout_sheet(out, sheet, template, preserve_summary=False)

    print(f"done -> {out}")


if __name__ == "__main__":
    main()
