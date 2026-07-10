# -*- coding: utf-8 -*-
"""Column layout templates matching user rearranged 欧洲-核心数据 summary block."""

from __future__ import annotations

from typing import Any

M15_WIDTH = 4
BLOCK_WIDTH = 14
BLOCK_STRIDE = 16

SALES_HEADERS: list[str | None] = [
    "2023",
    "2024",
    "2025",
    "2025Q1",
    "2026Q1",
    "2025M1-5",
    "2026M1-5",
    "25同比绝对量",
    "24-25 YoY",
    "26Q1绝对增量",
    "26Q1 YoY",
    "26M1-5增量",
    "26M1-5 YoY",
]

SHARE_HEADERS: list[str | None] = [
    "2023",
    "2024",
    "2025",
    "2025Q1",
    "2026Q1",
    "2025M1-5",
    "2026M1-5",
    "24年25变化",
    None,
    "25Q1与26Q1变化",
    None,
    "25M1-5→26M1-5变化",
    None,
]

HEADER_ALIASES: dict[str, str] = {
    "25同比绝对量": "25同比绝对量",
    "25同比": "25同比绝对量",
    "26Q1绝对增量": "26Q1绝对增量",
    "26Q1绝对量增量": "26Q1绝对增量",
    "26Q1增量": "26Q1绝对增量",
    "26M1-5增量": "26M1-5增量",
    "25M1-5→26M1-5变化": "25M1-5→26M1-5变化",
    "25M1-5→26M1-5变化": "25M1-5→26M1-5变化",
}


def normalize_header(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if isinstance(value, int) and 1900 <= value <= 2100:
        return str(value)
    text = str(value).strip()
    if not text:
        return None
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return HEADER_ALIASES.get(text, text)


def is_share_header_row(headers: dict[int, str | None]) -> bool:
    labels = {v for v in headers.values() if v}
    if "市占率" in labels or "占比" in labels:
        return True
    if "24年25变化" in labels or "25Q1与26Q1变化" in labels:
        return True
    return False


def target_headers(share: bool) -> list[str | None]:
    return SHARE_HEADERS if share else SALES_HEADERS


def q1_col_after_inserts(orig_qcol: int, orig_qcols: list[int]) -> int:
    return orig_qcol + M15_WIDTH * sum(1 for q in orig_qcols if q < orig_qcol)


def m15_col_after_inserts(orig_qcol: int, orig_qcols: list[int]) -> int:
    return (orig_qcol + 1) + M15_WIDTH * sum(1 for q in orig_qcols if q < orig_qcol)


def insert_col_for_q(orig_qcol: int) -> int:
    return orig_qcol + 1
