#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Verify that a .docx conforms to the Jinan Cigarette Factory format standard.

Usage:
    python verify_gongwen.py output.docx

Zero external dependencies. Prints a checklist and exits 0 only when all checks pass.
"""

from __future__ import annotations

import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


EXPECTED = {
    "GongwenTitle":    ("方正小标宋简体", "44", "center", False),
    "GongwenSubtitle": ("楷体_GB2312",   "32", "center", False),
    "GongwenBlank":    ("仿宋_GB2312",   "32", "left",   False),
    "GongwenH1":       ("黑体",          "32", "both",   True),
    "GongwenH2":       ("楷体_GB2312",   "32", "both",   True),
    "GongwenH3":       ("仿宋_GB2312",   "32", "both",   True),
    "GongwenH4":       ("仿宋_GB2312",   "32", "both",   True),
    "GongwenBody":     ("仿宋_GB2312",   "32", "both",   True),
    "GongwenSign":     ("仿宋_GB2312",   "32", "left",   False),
}

EXPECTED_MARGINS = {"top": "2098", "bottom": "1984", "left": "1587", "right": "1474"}


def check(results: list[tuple[bool, str]], ok: bool, label: str, detail: str = "") -> None:
    results.append((bool(ok), f"{'PASS' if ok else 'FAIL'}  {label}" + (f" — {detail}" if detail else "")))


def verify(path: Path) -> int:
    results: list[tuple[bool, str]] = []
    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
            check(results, bad is None, "ZIP 包完整", f"testzip={bad}")
            doc = ET.fromstring(zf.read("word/document.xml"))
            ET.fromstring(zf.read("word/styles.xml"))
    except Exception as exc:  # noqa: BLE001
        print(f"FAIL  无法读取 docx: {exc}")
        return 1

    sect = doc.find(f".//{q('sectPr')}")
    if sect is None:
        check(results, False, "存在节设置 sectPr")
    else:
        pg_sz = sect.find(q("pgSz"))
        check(results, pg_sz is not None and pg_sz.get(q("w")) == "11906"
              and pg_sz.get(q("h")) == "16838", "纸张 A4",
              f"{pg_sz.attrib if pg_sz is not None else 'missing'}")
        pg_mar = sect.find(q("pgMar"))
        if pg_mar is None:
            check(results, False, "存在页边距设置")
        else:
            for key, expected in EXPECTED_MARGINS.items():
                actual = pg_mar.get(q(key))
                check(results, actual == expected, f"页边距 {key}",
                      f"期望 {expected}，实际 {actual}")

    paragraphs = doc.findall(f".//{q('body')}/{q('p')}")
    check(results, len(paragraphs) >= 4, "段落数量充足", f"共 {len(paragraphs)} 段")

    texts: list[str] = []
    present_styles: set[str] = set()
    filled = 0
    for idx, para in enumerate(paragraphs, 1):
        ppr = para.find(q("pPr"))
        if ppr is None:
            check(results, False, f"第 {idx} 段缺少 pPr")
            continue
        style_el = ppr.find(q("pStyle"))
        style = style_el.get(q("val")) if style_el is not None else None
        text = "".join(t.text or "" for t in para.findall(f".//{q('t')}"))
        texts.append(text)
        if text:
            filled += 1

        if style not in EXPECTED:
            check(results, False, f"第 {idx} 段样式已注册", f"style={style}")
            continue
        present_styles.add(style)

        font_exp, sz_exp, align_exp, indent_exp = EXPECTED[style]
        run = para.find(q("r"))
        rpr = run.find(q("rPr")) if run is not None else None
        fonts = rpr.find(q("rFonts")) if rpr is not None else None
        sz = rpr.find(q("sz")) if rpr is not None else None
        actual_font = fonts.get(q("eastAsia")) if fonts is not None else None
        actual_sz = sz.get(q("val")) if sz is not None else None

        spacing = ppr.find(q("spacing"))
        line = spacing.get(q("line")) if spacing is not None else None
        rule = spacing.get(q("lineRule")) if spacing is not None else None
        jc = ppr.find(q("jc"))
        align = jc.get(q("val")) if jc is not None else None
        ind = ppr.find(q("ind"))
        has_indent = ind is not None and ind.get(q("firstLineChars")) == "200"

        check(results, actual_font == font_exp, f"第 {idx} 段字体({style})",
              f"期望 {font_exp}，实际 {actual_font}")
        check(results, actual_sz == sz_exp, f"第 {idx} 段字号({style})",
              f"期望 {sz_exp} 半磅，实际 {actual_sz}")
        check(results, align == align_exp, f"第 {idx} 段对齐({style})",
              f"期望 {align_exp}，实际 {align}")
        check(results, line == "580" and rule == "exact", f"第 {idx} 段固定行距({style})",
              f"期望 line=580 exact，实际 line={line} rule={rule}")
        before_auto = spacing.get(q("beforeAutospacing")) if spacing is not None else None
        after_auto = spacing.get(q("afterAutospacing")) if spacing is not None else None
        check(results, before_auto == "0" and after_auto == "0",
              f"第 {idx} 段段前后非自动({style})",
              f"beforeAuto={before_auto} afterAuto={after_auto}")
        if style != "GongwenBlank":
            check(results, has_indent == indent_exp, f"第 {idx} 段首行缩进({style})",
                  f"期望 {'2字符' if indent_exp else '无'}，实际 {'2字符' if has_indent else '无'}")

    check(results, filled >= 4, "正文内容非空", f"非空段落 {filled} 段")

    # 标题块（主标题 + 可选副标题）后恰好一个空行
    blank_idx = 1 if len(texts) > 1 and texts[1] == "" else 2
    if len(texts) > blank_idx:
        check(results, texts[blank_idx] == "", "标题块后存在空行",
              f"第 {blank_idx + 1} 段")
        if len(texts) > blank_idx + 1:
            check(results, bool(texts[blank_idx + 1]), "空行后紧跟正文",
                  f"第 {blank_idx + 2} 段")

    # 四级编号样式
    joined = "\n".join(texts)
    for label, pattern, used_style in (
        ("一级编号 一、", r"^一、", "GongwenH1"),
        ("二级编号 （一）", r"^（一）", "GongwenH2"),
        ("三级编号 1．", r"^1．", "GongwenH3"),
        ("四级编号 （1）", r"^（1）", "GongwenH4"),
    ):
        if used_style in present_styles:
            check(results, bool(re.search(pattern, joined, re.M)), label)
        else:
            check(results, True, f"{label}（文档中未使用该层级，跳过）")

    for ok, line_text in results:
        print(line_text)
    passed = sum(1 for ok, _ in results if ok)
    print(f"\n通过 {passed}/{len(results)} 项检查")
    return 0 if all(ok for ok, _ in results) else 1


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python verify_gongwen.py output.docx")
        return 2
    return verify(Path(sys.argv[1]).resolve())


if __name__ == "__main__":
    raise SystemExit(main())
