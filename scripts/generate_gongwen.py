#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate a Jinan Cigarette Factory (济南卷烟厂) official-format .docx.

Zero external dependencies: only Python's standard library is required.
The output is a valid OOXML (Word .docx) package built directly with zipfile,
so it works in WorkBuddy, Doubao, CI runners, or any plain Python 3 environment.

Usage:
    python generate_gongwen.py content.json -o output.docx
    python generate_gongwen.py --self-test -o sample.docx

Format enforced (per 济南卷烟厂公文格式):
    Page      A4; margins top 3.7cm, bottom 3.5cm, left 2.8cm, right 2.6cm
    Title     方正小标宋简体, 二号 (22pt), centered, exact 29pt line spacing
    Subtitle  楷体_GB2312, 三号 (16pt), centered, exact 29pt line spacing
    H1        黑体, 三号 (16pt), marker 一、
    H2        楷体_GB2312, 三号 (16pt), marker （一）
    H3        仿宋_GB2312, 三号 (16pt), marker 1．
    H4        仿宋_GB2312, 三号 (16pt), marker （1）
    Body      仿宋_GB2312, 三号 (16pt), justified, exact 29pt line spacing,
              first-line indent 2 characters, one blank line below the title block

Note on font names: the factory spec writes 仿宋GB2312 / 楷体GB2312; the names
installed on Chinese Windows are 仿宋_GB2312 / 楷体_GB2312 (with underscore).
The script uses the installed names by default. Override via FONTS below if the
target machine uses different names.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

# --------------------------------------------------------------------------
# Format constants — edit here if the factory standard changes
# --------------------------------------------------------------------------

FONTS = {
    "title": "方正小标宋简体",
    "subtitle": "楷体_GB2312",
    "h1": "黑体",
    "h2": "楷体_GB2312",
    "h3": "仿宋_GB2312",
    "h4": "仿宋_GB2312",
    "body": "仿宋_GB2312",
    "sign": "仿宋_GB2312",
    "blank": "仿宋_GB2312",
}

SIZE_TITLE_PT = 22  # 二号
SIZE_BODY_PT = 16   # 三号
LINE_PT = 29        # 固定值 29 磅

MARGINS_CM = {"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6}
PAGE_CM = {"width": 21.0, "height": 29.7}  # A4

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Paragraph style definitions: (styleId, font key, size pt, alignment, indent chars)
STYLE_DEFS = {
    "title":    ("GongwenTitle",    "title",    SIZE_TITLE_PT, "center",  0),
    "subtitle": ("GongwenSubtitle", "subtitle", SIZE_BODY_PT,  "center",  0),
    "blank":    ("GongwenBlank",    "blank",    SIZE_BODY_PT,  "left",    0),
    "h1":       ("GongwenH1",       "h1",       SIZE_BODY_PT,  "both",    2),
    "h2":       ("GongwenH2",       "h2",       SIZE_BODY_PT,  "both",    2),
    "h3":       ("GongwenH3",       "h3",       SIZE_BODY_PT,  "both",    2),
    "h4":       ("GongwenH4",       "h4",       SIZE_BODY_PT,  "both",    2),
    "body":     ("GongwenBody",     "body",     SIZE_BODY_PT,  "both",    2),
    "sign":     ("GongwenSign",     "sign",     SIZE_BODY_PT,  "left",    0),
}

HEADING_ORDER = ("h1", "h2", "h3", "h4")

CN_DIGITS = "零一二三四五六七八九"


# --------------------------------------------------------------------------
# Unit conversion helpers
# --------------------------------------------------------------------------

def cm_to_twips(cm: float) -> int:
    return int(round(cm * 1440.0 / 2.54))


def pt_to_half_points(pt: float) -> int:
    return int(round(pt * 2))


def pt_to_twips(pt: float) -> int:
    return int(round(pt * 20))


def cn_number(n: int) -> str:
    """Convert an integer to Chinese numerals for 一、/（一）markers."""
    if n <= 0:
        return str(n)
    if n < 10:
        return CN_DIGITS[n]
    if n < 20:
        return "十" + (CN_DIGITS[n % 10] if n % 10 else "")
    if n < 100:
        tens, ones = divmod(n, 10)
        return CN_DIGITS[tens] + "十" + (CN_DIGITS[ones] if ones else "")
    if n < 1000:
        hundreds, rest = divmod(n, 100)
        tail = cn_number(rest) if rest else ""
        if rest and rest < 10:
            tail = "零" + tail
        return CN_DIGITS[hundreds] + "百" + tail
    return str(n)


def marker(kind: str, n: int) -> str:
    if kind == "h1":
        return cn_number(n) + "、"
    if kind == "h2":
        return "（" + cn_number(n) + "）"
    if kind == "h3":
        return f"{n}．"
    if kind == "h4":
        return f"（{n}）"
    return ""


# Already-numbered headings are left untouched; otherwise a marker is added.
MARKER_RE = {
    "h1": re.compile(r"^[一二三四五六七八九十百零]+、"),
    "h2": re.compile(r"^（[一二三四五六七八九十百零]+）"),
    "h3": re.compile(r"^\d+\s*[．.]"),
    "h4": re.compile(r"^（\d+）"),
}


# --------------------------------------------------------------------------
# OOXML building blocks
# --------------------------------------------------------------------------

def rpr_xml(font_key: str, size_pt: float, bold: bool = False) -> str:
    font = FONTS[font_key]
    half = pt_to_half_points(size_pt)
    bold_xml = "<w:b/><w:bCs/>" if bold else ""
    return (
        "<w:rPr>"
        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'
        f"{bold_xml}"
        f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>'
        '<w:lang w:val="zh-CN" w:eastAsia="zh-CN"/>'
        "</w:rPr>"
    )


def ppr_xml(kind: str, indent_chars: int | None = None, align: str | None = None) -> str:
    style_id, _font, _size, default_align, default_indent = STYLE_DEFS[kind]
    align = align or default_align
    indent_chars = default_indent if indent_chars is None else indent_chars
    line = pt_to_twips(LINE_PT)

    parts = [
        f'<w:pStyle w:val="{style_id}"/>',
        '<w:keepNext/>' if kind in ("title", "subtitle") else "",
        '<w:snapToGrid w:val="0"/>',
        f'<w:spacing w:before="0" w:after="0" '
        f'w:beforeAutospacing="0" w:afterAutospacing="0" '
        f'w:line="{line}" w:lineRule="exact"/>',
    ]
    if indent_chars:
        # firstLineChars is in 1/100 character units; firstLine is a twip fallback.
        fallback = pt_to_twips(SIZE_BODY_PT * 2) * indent_chars // 2
        parts.append(f'<w:ind w:firstLineChars="{indent_chars * 100}" w:firstLine="{fallback}"/>')
    parts.append(f'<w:jc w:val="{align}"/>')
    return "<w:pPr>" + "".join(parts) + "</w:pPr>"


def paragraph_xml(text: str, kind: str, indent_chars=None, align=None, bold=False) -> str:
    _style_id, font_key, size_pt, _a, _i = STYLE_DEFS[kind]
    safe = escape(text).replace("\n", "<w:br/>")
    return (
        "<w:p>"
        + ppr_xml(kind, indent_chars=indent_chars, align=align)
        + "<w:r>"
        + rpr_xml(font_key, size_pt, bold=bold)
        + f'<w:t xml:space="preserve">{safe}</w:t>'
        + "</w:r>"
        + "</w:p>"
    )


def styles_xml() -> str:
    blocks = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<w:styles xmlns:w="{W_NS}">',
        "<w:docDefaults><w:rPrDefault><w:rPr>"
        f'<w:rFonts w:ascii="{FONTS["body"]}" w:hAnsi="{FONTS["body"]}" '
        f'w:eastAsia="{FONTS["body"]}" w:cs="{FONTS["body"]}"/>'
        f'<w:sz w:val="{pt_to_half_points(SIZE_BODY_PT)}"/>'
        "</w:rPr></w:rPrDefault></w:docDefaults>",
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/>'
        + ppr_xml("body")
        + rpr_xml("body", SIZE_BODY_PT)
        + "</w:style>",
    ]
    for kind, (style_id, _font, _size, _align, _indent) in STYLE_DEFS.items():
        if style_id == "GongwenBody":
            continue
        blocks.append(
            f'<w:style w:type="paragraph" w:styleId="{style_id}">'
            f'<w:name w:val="{style_id}"/>'
            '<w:basedOn w:val="Normal"/>'
            + ppr_xml(kind)
            + rpr_xml(kind, SIZE_TITLE_PT if kind == "title" else SIZE_BODY_PT)
            + "</w:style>"
        )
    blocks.append("</w:styles>")
    return "".join(blocks)


def document_xml(paragraphs: list[str]) -> str:
    page_w, page_h = cm_to_twips(PAGE_CM["width"]), cm_to_twips(PAGE_CM["height"])
    sect = (
        "<w:sectPr>"
        f'<w:pgSz w:w="{page_w}" w:h="{page_h}"/>'
        f'<w:pgMar w:top="{cm_to_twips(MARGINS_CM["top"])}" '
        f'w:right="{cm_to_twips(MARGINS_CM["right"])}" '
        f'w:bottom="{cm_to_twips(MARGINS_CM["bottom"])}" '
        f'w:left="{cm_to_twips(MARGINS_CM["left"])}" '
        'w:header="851" w:footer="992" w:gutter="0"/>'
        '<w:cols w:space="425"/>'
        "</w:sectPr>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NS}"><w:body>'
        + "".join(paragraphs)
        + sect
        + "</w:body></w:document>"
    )


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Office Word</Application></Properties>"""


def core_xml(title: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f"<dc:title>{escape(title)}</dc:title>"
        "<dc:creator>济南卷烟厂</dc:creator>"
        "<cp:lastModifiedBy>济南卷烟厂</cp:lastModifiedBy>"
        "</cp:coreProperties>"
    )


# --------------------------------------------------------------------------
# Content model
# --------------------------------------------------------------------------

def normalize_content(data: dict) -> tuple[str, str, list[dict]]:
    """Accept both the full schema and the simple {title, sign, body:[...]} form."""
    title = data.get("main_title") or data.get("title") or ""
    if not title:
        raise ValueError("缺少主标题：请提供 main_title（或 title）")
    subtitle = data.get("subtitle", "") or ""

    items: list[dict] = []
    if "paragraphs" in data:
        for raw in data["paragraphs"] or []:
            if isinstance(raw, str):
                items.append({"type": "body", "text": raw})
            elif isinstance(raw, dict):
                items.append(dict(raw))
            else:
                raise ValueError("paragraphs 元素必须是字符串或对象")
    elif "body" in data:
        items = [{"type": "body", "text": t} for t in (data.get("body") or [])]
    else:
        raise ValueError("缺少正文：请提供 paragraphs 或 body")

    sign = data.get("sign")
    if sign:
        items.append({"type": "sign", "text": sign})
    return title, subtitle, items


def build_paragraphs(title: str, subtitle: str, items: list[dict]) -> list[str]:
    out = [paragraph_xml(title, "title")]
    if subtitle:
        out.append(paragraph_xml(subtitle, "subtitle"))

    # 正文在标题下空一行（若调用方已显式给出空行则不重复插入）
    if not items or items[0].get("type") != "blank":
        out.append(paragraph_xml("", "blank"))

    counters = dict.fromkeys(HEADING_ORDER, 0)
    for idx, item in enumerate(items):
        kind = (item.get("type") or "body").strip().lower()
        if kind not in STYLE_DEFS:
            raise ValueError(f"第 {idx + 1} 段类型无效: {kind}")

        text = str(item.get("text", "") or "")
        if kind in HEADING_ORDER:
            level = HEADING_ORDER.index(kind)
            counters[kind] += 1
            for lower in HEADING_ORDER[level + 1:]:
                counters[lower] = 0
            if not MARKER_RE[kind].match(text) and item.get("auto_number", True):
                text = marker(kind, counters[kind]) + text

        out.append(paragraph_xml(
            text,
            kind,
            indent_chars=item.get("indent_chars"),
            align=item.get("align"),
            bold=bool(item.get("bold", False)),
        ))
    return out


def self_test_data() -> dict:
    return {
        "main_title": "关于开展公文格式规范化培训的通知",
        "subtitle": "——以标准铸根基 以规范促提升",
        "paragraphs": [
            {"type": "body", "text": "为统一济南卷烟厂公文排版标准，提升文稿规范水平，现就公文格式有关要求通知如下。"},
            {"type": "h1", "text": "总体要求"},
            {"type": "body", "text": "各部门要严格贯彻执行公文格式标准，做到要素齐全、层级清晰、版面规范。"},
            {"type": "h2", "text": "统一思想认识"},
            {"type": "body", "text": "公文格式是企业管理基础工作的重要组成部分，体现工作作风与专业水准。"},
            {"type": "h3", "text": "明确责任分工"},
            {"type": "body", "text": "综合管理部门负责标准宣贯，各业务部门负责本部门文稿的自查与落实。"},
            {"type": "h4", "text": "细化时间节点"},
            {"type": "body", "text": "自本通知印发之日起，新起草文稿一律按本标准排版。"},
            {"type": "h2", "text": "强化检查考核"},
            {"type": "body", "text": "定期对公文格式执行情况开展检查，检查结果纳入部门基础管理评价。"},
        ],
        "sign": "（济南卷烟厂）",
    }


# --------------------------------------------------------------------------
# Package writer
# --------------------------------------------------------------------------

def write_docx(path: Path, paras: list[str], title: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": ROOT_RELS,
        "word/document.xml": document_xml(paras),
        "word/styles.xml": styles_xml(),
        "word/_rels/document.xml.rels": DOC_RELS,
        "docProps/core.xml": core_xml(title),
        "docProps/app.xml": APP_XML,
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in parts.items():
            zf.writestr(name, content.encode("utf-8"))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成符合济南卷烟厂公文格式的 docx")
    parser.add_argument("input", nargs="?", help="内容 JSON 文件；与 --self-test 二选一")
    parser.add_argument("-o", "--output", required=True, help="输出 .docx 路径")
    parser.add_argument("--self-test", action="store_true", help="生成覆盖全部标题层级的示例文档")
    args = parser.parse_args(argv)

    if args.self_test == bool(args.input):
        parser.error("请且仅请指定 content.json 或 --self-test 其中之一")

    try:
        if args.self_test:
            data = self_test_data()
        else:
            with open(args.input, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            if not isinstance(data, dict):
                raise ValueError("内容 JSON 根节点必须是对象")

        title, subtitle, items = normalize_content(data)
        paras = build_paragraphs(title, subtitle, items)
        out = Path(args.output).resolve()
        write_docx(out, paras, title)

        with zipfile.ZipFile(out) as zf:
            bad = zf.testzip()
        if bad is not None:
            raise RuntimeError(f"生成的 docx 校验失败: {bad}")

        print(f"已生成: {out}")
        print(f"段落数: {len(paras)}")
        print(f"页边距(cm): 上{MARGINS_CM['top']} 下{MARGINS_CM['bottom']} "
              f"左{MARGINS_CM['left']} 右{MARGINS_CM['right']}")
        print(f"正文字号: 三号({SIZE_BODY_PT}pt) / 固定行距: {LINE_PT}磅")
        return 0
    except Exception as exc:  # noqa: BLE001 - CLI should report any failure clearly
        print(f"错误: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
