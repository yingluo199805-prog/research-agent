# -*- coding: utf-8 -*-
"""build-revenue-split 公共配置：工作根目录 / 年报文本缓存目录。

根目录优先级：环境变量 REVENUE_SPLIT_BASE > 默认 <仓库>/outputs/收入分配。
年报转出的文本缓存在 <BASE>/_年报文本/<prefix>_<FY>.txt（不再写 /tmp）。
"""
import os

_HERE = os.path.dirname(os.path.abspath(__file__))
_DEFAULT_BASE = os.path.abspath(os.path.join(_HERE, "..", "..", "..", "outputs", "收入分配"))

BASE = os.environ.get("REVENUE_SPLIT_BASE") or _DEFAULT_BASE
TXT_DIR = os.path.join(BASE, "_年报文本")


def set_base(path):
    """命令行 --base 覆盖（须在其他模块读 BASE 之前调用）。"""
    global BASE, TXT_DIR
    BASE = os.path.abspath(path)
    TXT_DIR = os.path.join(BASE, "_年报文本")
    os.environ["REVENUE_SPLIT_BASE"] = BASE


def txt_path(prefix, fy):
    return os.path.join(TXT_DIR, f"{prefix}_{fy}.txt")


def base_path(*parts):
    return os.path.join(BASE, *parts)
