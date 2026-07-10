import argparse
import json
import os
import shutil
import time
import zipfile
from datetime import datetime

import openpyxl


HEADER11 = ["年", "月", "国家", "区域", "细分区域", "品牌", "车型", "能源类型（分为）", "数量", "系别", "车企划分"]
COL_LETTERS = [chr(ord("A") + i) for i in range(11)]


def xml_escape(value):
    return str(value).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def encode_row(row_num, values):
    parts = [f'<row r="{row_num}">']
    for col_idx, value in enumerate(values[:11]):
        if value is None:
            continue
        ref = COL_LETTERS[col_idx] + str(row_num)
        if isinstance(value, bool):
            parts.append(f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>')
        elif isinstance(value, (int, float)):
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            parts.append(f'<c r="{ref}"><v>{value}</v></c>')
        else:
            parts.append(f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{xml_escape(value)}</t></is></c>')
    parts.append("</row>")
    return "".join(parts).encode("utf-8")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--main", required=True)
    ap.add_argument("--payload", required=True)
    args = ap.parse_args()

    with open(args.payload, "r", encoding="utf-8") as f:
        payload = json.load(f)
    new_rows = payload["cleaned"]
    if not new_rows:
        raise SystemExit("No new rows")

    wb = openpyxl.load_workbook(args.main, read_only=True, data_only=True)
    ws = wb["海关CV"]
    ws.reset_dimensions()
    old_rows = []
    existing_target = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        vals = list(row[:11])
        if vals[0] == 2026 and vals[1] == 4:
            existing_target += 1
        old_rows.append(vals)
    wb.close()
    if existing_target:
        raise SystemExit(f"2026-04 already exists in 海关CV: {existing_target} rows")

    backup = args.main.replace(".xlsx", f"_preHgZipAppend_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    shutil.copy2(args.main, backup)

    total_rows = 1 + len(old_rows) + len(new_rows)
    sheet_path = "xl/worksheets/sheet2.xml"
    tmp = args.main + ".hg.tmp"

    sheet_open = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
        'mc:Ignorable="x14ac">'
        f'<dimension ref="A1:K{total_rows}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>'
        "<sheetData>"
    ).encode("utf-8")
    sheet_close = b"</sheetData></worksheet>"

    t0 = time.time()
    with zipfile.ZipFile(args.main, "r") as zin, zipfile.ZipFile(
        tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6, allowZip64=True
    ) as zout:
        for item in zin.infolist():
            if item.filename == sheet_path:
                with zout.open(item, "w", force_zip64=True) as fp:
                    fp.write(sheet_open)
                    fp.write(encode_row(1, HEADER11))
                    rnum = 2
                    for row in old_rows:
                        fp.write(encode_row(rnum, row))
                        rnum += 1
                    for row in new_rows:
                        fp.write(encode_row(rnum, row))
                        rnum += 1
                    fp.write(sheet_close)
            else:
                zout.writestr(item, zin.read(item.filename))
    os.replace(tmp, args.main)

    print(json.dumps({
        "status": "appended",
        "backup": backup,
        "old_rows": len(old_rows),
        "new_rows": len(new_rows),
        "total_rows_with_header": total_rows,
        "elapsed_sec": round(time.time() - t0, 1),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
