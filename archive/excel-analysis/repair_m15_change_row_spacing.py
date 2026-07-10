# -*- coding: utf-8 -*-
"""Align M1-5 share-change rows by widening them to the same part width."""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32


TARGET_HEADER = "2026M1-5相较于2025M1-5"
XL_SHIFT_TO_RIGHT = -4161


def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def col_letter(col: int) -> str:
    out = ""
    while col:
        col, rem = divmod(col - 1, 26)
        out = chr(65 + rem) + out
    return out


def find_block_start(ws: Any, row: int, col: int) -> int:
    cur = col
    while cur > 1:
        if text(ws.Cells(row, cur - 1).Value) == "":
            return cur
        cur -= 1
    return 1


def find_end_row(ws: Any, header_row: int, start_col: int, max_row: int) -> int:
    row = header_row + 1
    while row <= max_row:
        value = text(ws.Cells(row, start_col).Value)
        if not value or value.startswith(("一、", "二、", "三、", "四、", "五、", "六、", "七、")):
            break
        row += 1
    return row - 1


def repair(input_path: Path, output_path: Path) -> None:
    input_path = input_path.resolve()
    output_path = output_path.resolve()
    if output_path.exists():
        output_path.unlink()
    shutil.copy2(input_path, output_path)

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    try:
        wb = excel.Workbooks.Open(str(output_path))
        for ws in wb.Worksheets:
            used = ws.UsedRange
            last_row = int(used.Row + used.Rows.Count - 1)
            hits: list[tuple[int, int, int]] = []
            found = used.Find(What=TARGET_HEADER, LookAt=1, SearchOrder=1)
            if found is not None:
                first_address = found.Address
                while True:
                    row = int(found.Row)
                    col = int(found.Column)
                    start_col = find_block_start(ws, row, col)
                    end_row = find_end_row(ws, row, start_col, last_row)
                    hits.append((row, col, end_row))
                    found = used.FindNext(found)
                    if found is None or found.Address == first_address:
                        break
            for row, col, end_row in sorted(hits, key=lambda item: (item[1], item[0]), reverse=True):
                address = f"{col_letter(col + 1)}{row}:{col_letter(col + 3)}{end_row}"
                ws.Range(address).Insert(Shift=XL_SHIFT_TO_RIGHT)
                ws.Range(address).ColumnWidth = 3
        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
        excel.Quit()
        pythoncom.CoUninitialize()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-workbook", required=True, type=Path)
    parser.add_argument("--output-workbook", required=True, type=Path)
    args = parser.parse_args()
    repair(args.input_workbook, args.output_workbook)


if __name__ == "__main__":
    main()
