#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数学讲义批量格式化工具。

扫描目录下所有 .docx 文件，统一排版为标准中文数学讲义格式；
原文件保持不动，结果写入输出目录（默认 <输入目录>/formatted/）。

默认格式规范（可在下方「格式配置」处调整）：
    页面      A4 纵向；上下边距 1.6cm，左右 1.7cm
    页眉页脚   华文行楷 10 号；页脚居中「第 N 页」；页眉默认沿用原内容并统一字体，
              可用 --header 参数或启动时输入统一设置页眉（居中、带下框线）
    讲义标题   黑体 小三(15pt)，居中                        —— 文档首个非空段落
    一级标题   黑体 四号(14pt)，居中                        —— “一、…”“第2讲 …”
    二级标题   黑体 小四(12pt)                              —— “（一）…”“1.1 …”“练习”“作业”等
    例题      “例1”等前缀黑体加粗，段落顶格                  —— “例1…”“例题 2 …”
    题目标签   【】指示标签用华文行楷加粗，段落顶格             —— “【典例1】…”“【变式1】…”
    正文      中文宋体、西文 Times New Roman，五号(10.5pt)，
              1.5 倍行距，首行缩进 2 字符，两端对齐；
              “解：/证明：/分析：”等前导词自动加粗
    其他      删除多余空行；表格、图片、公式原样保留；
              正文中原有的加粗/斜体/颜色等强调保留

用法：
    python format_handout.py                          # 交互模式：启动后按提示输入路径
    python format_handout.py <目录或.docx文件> [-o 输出目录]

示例：
    python format_handout.py                          # 启动后输入 D:/讲义 等路径
    python format_handout.py D:/讲义                  # 输出到 D:/讲义/formatted
    python format_handout.py 七年级下 -o 已排版       # 输出到 七年级下/已排版
    python format_handout.py 第3讲.docx               # 单文件，输出到 formatted/
    python format_handout.py D:/讲义 --header "八年级数学讲义"   # 统一设置页眉
"""

from __future__ import annotations

import argparse
import copy
import re
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.run import Run

__version__ = "1.0.0"

# ================================ 格式配置（按需调整） ================================

LATIN_FONT = "Times New Roman"    # 西文字体
BODY_FONT = "宋体"                # 正文中文字体
HEADING_FONT = "黑体"             # 标题中文字体
HEADER_FOOTER_FONT = "华文行楷"    # 页眉/页脚字体
HEADER_FOOTER_SIZE = 10           # 页眉/页脚字号（10 号 = 10pt）

LABEL_FONT = "华文行楷"            # 【典例1】【变式1】等题目标签字体
LABEL_BOLD = True                 # 标签是否加粗

HEADER_TEXT = ""                  # 预设页眉文字（留空 = 不修改原文档页眉内容）

PAGE_CM = dict(top=1.6, bottom=1.6, left=1.7, right=1.7)   # A4 页边距

# 各类段落样式：size=字号(pt)  align=对齐  line=行距(倍)
#               before/after=段前/段后间距(pt)  indent=首行缩进(字符)
#               outline=大纲级别(None 表示正文)  force_bold=是否强制统一加粗状态
STYLE_TITLE = dict(font=HEADING_FONT, size=15, bold=False,
                   align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0,
                   before=0, after=12, indent=0, outline=0, force_bold=True)
STYLE_H1 = dict(font=HEADING_FONT, size=14, bold=False,
                align=WD_ALIGN_PARAGRAPH.CENTER, line=1.0,
                before=12, after=12, indent=0, outline=0, force_bold=True)
STYLE_H2 = dict(font=HEADING_FONT, size=12, bold=False,
                align=WD_ALIGN_PARAGRAPH.LEFT, line=1.0,
                before=8, after=8, indent=0, outline=1, force_bold=True)
STYLE_BODY = dict(font=BODY_FONT, size=10.5, bold=False,
                  align=WD_ALIGN_PARAGRAPH.JUSTIFY, line=1.5,
                  before=0, after=0, indent=2, outline=None, force_bold=False)
STYLE_EXAMPLE = dict(STYLE_BODY, indent=0)   # 例题段：顶格，“例N”前缀另加粗

# ================================ 结构识别 ================================

RE_EXAMPLE = re.compile(r"^(?:例题?\s*\d{1,3}|例\s*[一二三四五六七八九十]{1,3})\s*[、．.:：)）]?\s*")
RE_BRACKET_LABEL = re.compile(r"^【[^】]{1,12}】\s*[、．.:：]?\s*")   # 【典例1】【变式1】【知识点】…
RE_UNIT_TITLE = re.compile(r"^第\s*[一二三四五六七八九十百\d]+\s*[讲章节单元卷]")
RE_CN_ENUM = re.compile(r"^[一二三四五六七八九十百]+\s*[、．.]")
RE_PAREN_ENUM = re.compile(r"^[（(][一二三四五六七八九十百]+\s*[)）]")
RE_NUM_MULTI = re.compile(r"^\d+(?:\.\d+)+\s*[、．.\s]?")

SECTION_HEADS = (
    "知识梳理", "知识点", "知识要点", "知识清单", "知识归纳", "学习目标", "教学目标",
    "重难点", "重点难点", "考情分析", "知识回顾", "温故知新", "课前预习", "自主预习",
    "新课导入", "例题精讲", "典例剖析", "典例分析", "精讲精练", "随堂练习", "课堂练习",
    "巩固练习", "课堂检测", "当堂检测", "达标检测", "拓展提高", "拓展提升", "思维拓展",
    "方法总结", "归纳总结", "本课小结", "课堂小结", "参考答案", "温馨提示",
    "练习", "习题", "例题", "作业",
)
# 整行仅为板块名（可带编号/冒号）时视为二级标题，如“练习”“作业：”“巩固练习（一）”
RE_SECTION_HEAD = re.compile(
    r"^(?:%s)\s*(?:[（(]\s*[一二三四五六七八九十\d]{1,3}\s*[)）]|[一二三四五六七八九十\d]{1,3})?\s*[：:]?\s*$"
    % "|".join(sorted(SECTION_HEADS, key=len, reverse=True))
)

# 段首前导词（后接内容时仅加粗前导词）
RE_LABEL = re.compile(
    r"^(?:注意|证明|分析|思路|提示|评注|反思|小结|知识点|知识梳理|作业|练习|习题|解|证|注)\s*[:：]"
)


def _is_h1(text: str) -> bool:
    return bool(RE_UNIT_TITLE.match(text) and len(text) <= 30) \
        or bool(RE_CN_ENUM.match(text) and len(text) <= 40)


def _is_h2(text: str) -> bool:
    return bool(RE_PAREN_ENUM.match(text) and len(text) <= 40) \
        or bool(RE_NUM_MULTI.match(text) and len(text) <= 40)


def _classify(text: str, style_name: str) -> tuple[str, int]:
    """返回 (段落类型, 前缀标签长度)。"""
    m = RE_BRACKET_LABEL.match(text)
    if m:
        return "labeled", m.end()
    m = RE_EXAMPLE.match(text)
    if m:
        return "example", m.end()
    if _is_h1(text):
        return "h1", 0
    if _is_h2(text) or RE_SECTION_HEAD.match(text):
        return "h2", 0
    m = RE_LABEL.match(text)
    if m:
        return "body", m.end()
    # 尊重原文档中的 Word 标题样式
    name = style_name or ""
    if "标题" in name or "heading" in name.lower():
        digits = re.search(r"\d", name)
        return ("h2" if digits and int(digits.group()) >= 2 else "h1"), 0
    return "body", 0


# ================================ 底层格式工具 ================================

def _iter_runs(paragraph):
    """段落中所有 run（含超链接内部的 run；公式 m:r 不受影响）。"""
    return [Run(r, paragraph) for r in paragraph._p.findall(".//" + qn("w:r"))]


def _set_run_font(run: Run, east: str, size: float, bold=None, latin: str = LATIN_FONT) -> None:
    """统一 run 字体：中文 east、西文 latin、字号 size；bold=None 表示不改加粗。"""
    run.font.name = latin
    run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rPr.get_or_add_rFonts().set(qn("w:eastAsia"), east)


def _set_run_text(r_el, text: str) -> None:
    """重设 run 的文本（保留 rPr 等格式属性）。"""
    for tag in ("w:t", "w:tab", "w:br", "w:cr"):
        for el in r_el.findall(qn(tag)):
            r_el.remove(el)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r_el.append(t)


def _bold_prefix(paragraph, n_chars: int, font: str, bold: bool = True) -> None:
    """把段落前 n_chars 个字符设为指定字体（必要时拆分 run），其余内容保持原样。"""
    left = n_chars
    for run in _iter_runs(paragraph):
        if left <= 0:
            break
        text = run.text
        if not text:
            continue
        if len(text) > left:  # 拆分为「标签前半 + 原样后半」
            tail = copy.deepcopy(run._element)
            run._element.addnext(tail)
            _set_run_text(run._element, text[:left])
            _set_run_text(tail, text[left:])
        _set_run_font(run, font, STYLE_BODY["size"], bold=bold)
        left -= min(left, len(text))


def _has_content(paragraph) -> bool:
    """段落是否含公式、图片等非文本内容（删除空段时必须保留）。"""
    p = paragraph._p
    for tag in ("w:drawing", "w:pict", "w:object", "w:fldSimple",
                "m:oMath", "m:oMathPara", "w:hyperlink"):
        if p.findall(".//" + qn(tag)):
            return True
    return False


def _apply_style(paragraph, spec: dict) -> None:
    pf = paragraph.paragraph_format
    pf.alignment = spec["align"]
    pf.line_spacing = spec["line"]
    pf.space_before = Pt(spec["before"])
    pf.space_after = Pt(spec["after"])
    pf.keep_with_next = spec["outline"] is not None   # 标题与后文同页

    pPr = paragraph._p.get_or_add_pPr()

    # 首行缩进（按字符，Word 精确模式），并清除原有各类缩进
    ind = pPr.get_or_add_ind()
    for attr in ("w:firstLine", "w:firstLineChars", "w:hanging", "w:hangingChars",
                 "w:left", "w:leftChars", "w:right", "w:rightChars"):
        ind.attrib.pop(qn(attr), None)
    if spec["indent"]:
        ind.set(qn("w:firstLineChars"), str(spec["indent"] * 100))
        ind.set(qn("w:firstLine"), str(int(spec["size"] * 20 * spec["indent"])))

    # 关闭“对齐到网格”，保证行距设置生效
    snap = pPr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        pPr.append(snap)
    snap.set(qn("w:val"), "0")

    # 大纲级别（Word 导航窗格/目录可用）
    if spec["outline"] is not None:
        lvl = pPr.find(qn("w:outlineLvl"))
        if lvl is None:
            lvl = OxmlElement("w:outlineLvl")
            pPr.append(lvl)
        lvl.set(qn("w:val"), str(spec["outline"]))

    bold = spec["bold"] if spec["force_bold"] else None
    for run in _iter_runs(paragraph):
        _set_run_font(run, spec["font"], spec["size"], bold=bold)


# ================================ 页面与页脚 ================================

def _footer_field_run():
    """构造页码域的内嵌 run（与页脚同字体字号）。"""
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), HEADER_FOOTER_FONT)
    rFonts.set(qn("w:hAnsi"), HEADER_FOOTER_FONT)
    rFonts.set(qn("w:eastAsia"), HEADER_FOOTER_FONT)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(HEADER_FOOTER_SIZE * 2))
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(HEADER_FOOTER_SIZE * 2))
    rPr.append(rFonts)
    rPr.append(sz)
    rPr.append(szCs)
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(rPr)
    r.append(t)
    return r


def _add_page_number(section) -> None:
    footer = section.footer
    footer.is_linked_to_previous = False
    for extra in footer.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for el in list(para._p):   # 清空原有页脚内容
        if el.tag != qn("w:pPr"):
            para._p.remove(el)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run("第 ")
    _set_run_font(run, HEADER_FOOTER_FONT, HEADER_FOOTER_SIZE, bold=False,
                  latin=HEADER_FOOTER_FONT)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE   \\* MERGEFORMAT ")
    fld.append(_footer_field_run())
    para._p.append(fld)
    run = para.add_run(" 页")
    _set_run_font(run, HEADER_FOOTER_FONT, HEADER_FOOTER_SIZE, bold=False,
                  latin=HEADER_FOOTER_FONT)


def _format_header(section) -> None:
    """统一页眉格式（不添加、不删除内容；无自定义页眉则跳过）。"""
    header = section.header
    if header.is_linked_to_previous:
        return
    for para in header.paragraphs:
        pf = para.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        for run in _iter_runs(para):
            _set_run_font(run, HEADER_FOOTER_FONT, HEADER_FOOTER_SIZE, bold=False,
                          latin=HEADER_FOOTER_FONT)


def _set_header_text(section, text: str) -> None:
    """用指定文字替换页眉内容：居中、带下框线、华文行楷 10 号。"""
    header = section.header
    header.is_linked_to_previous = False
    for extra in header.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    for el in list(para._p):   # 清空原有页眉内容
        if el.tag != qn("w:pPr"):
            para._p.remove(el)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    # 页眉下框线
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "auto")
    pBdr.append(bottom)
    pPr.insert_element_before(pBdr, "w:shd", "w:tabs", "w:spacing",
                              "w:ind", "w:jc", "w:rPr")
    run = para.add_run(text)
    _set_run_font(run, HEADER_FOOTER_FONT, HEADER_FOOTER_SIZE, bold=False,
                  latin=HEADER_FOOTER_FONT)


def _setup_page(doc: Document, header_text: str | None = None) -> None:
    for section in doc.sections:
        section.orientation = WD_ORIENT.PORTRAIT
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(PAGE_CM["top"])
        section.bottom_margin = Cm(PAGE_CM["bottom"])
        section.left_margin = Cm(PAGE_CM["left"])
        section.right_margin = Cm(PAGE_CM["right"])
        _add_page_number(section)
        if header_text:
            _set_header_text(section, header_text)
        else:
            _format_header(section)


def _setup_normal_style(doc: Document) -> None:
    """统一 Normal 样式，使表格等未直接设置格式的部分也继承宋体小四。"""
    style = doc.styles["Normal"]
    style.font.name = LATIN_FONT
    style.font.size = Pt(STYLE_BODY["size"])
    rPr = style.element.get_or_add_rPr()
    rPr.get_or_add_rFonts().set(qn("w:eastAsia"), BODY_FONT)


# ================================ 文档处理 ================================

def process_file(src: Path, dst: Path, header_text: str | None = None) -> Counter:
    doc = Document(str(src))
    _setup_page(doc, header_text)
    _setup_normal_style(doc)

    stats = Counter()
    paras = doc.paragraphs
    title_idx = next((i for i, p in enumerate(paras) if p.text.strip()), None)

    to_delete = []
    for i, para in enumerate(paras):
        text = para.text.strip()
        if not text:
            if not _has_content(para):
                to_delete.append(para)
            continue

        is_title = i == title_idx and len(text) <= 40 and (
            RE_UNIT_TITLE.match(text)
            or not (_is_h1(text) or _is_h2(text) or RE_EXAMPLE.match(text)
                    or RE_SECTION_HEAD.match(text) or RE_BRACKET_LABEL.match(text))
        )
        if is_title:
            kind, label = "title", 0
        else:
            kind, label = _classify(text, para.style.name if para.style is not None else "")

        if kind == "title":
            _apply_style(para, STYLE_TITLE)
            stats["讲义标题"] += 1
        elif kind == "h1":
            _apply_style(para, STYLE_H1)
            stats["一级标题"] += 1
        elif kind == "h2":
            _apply_style(para, STYLE_H2)
            stats["二级标题"] += 1
        elif kind == "example":
            _apply_style(para, STYLE_EXAMPLE)
            _bold_prefix(para, label, HEADING_FONT)
            stats["例题"] += 1
        elif kind == "labeled":
            _apply_style(para, STYLE_EXAMPLE)
            _bold_prefix(para, label, LABEL_FONT, bold=LABEL_BOLD)
            stats["题目标签"] += 1
        else:
            _apply_style(para, STYLE_BODY)
            if label:
                _bold_prefix(para, label, BODY_FONT)
            stats["正文"] += 1

    for para in to_delete:
        para._p.getparent().remove(para._p)
    stats["删除空段"] = len(to_delete)

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst))
    return stats


# ================================ 命令行入口 ================================

def _parse_args(argv=None):
    ap = argparse.ArgumentParser(
        prog="format_handout.py",
        description="批量将 Word 文件统一为标准数学讲义格式（原文件保持不动）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  python format_handout.py                 # 交互模式：启动后按提示输入路径\n"
               "  python format_handout.py D:/讲义\n"
               "  python format_handout.py 七年级下 -o 已排版\n"
               "  python format_handout.py 第3讲.docx\n",
    )
    ap.add_argument("input", nargs="?", help="包含 .docx 的目录，或单个 .docx 文件；省略则进入交互模式")
    ap.add_argument("-o", "--output", help="输出目录（默认：<输入目录>/formatted）")
    ap.add_argument("--header", help="统一设置页眉文字（替换原文档页眉内容；默认不修改）")
    ap.add_argument("-V", "--version", action="version", version=f"%(prog)s {__version__}")
    return ap.parse_args(argv)


def _prompt_path() -> Path | None:
    """交互式询问路径；直接回车返回 None。粘贴的路径带引号会自动去掉。"""
    raw = input("\n请输入讲义文件夹路径（或单个 .docx 文件；直接回车退出）：")
    raw = raw.strip().strip('"').strip("'").strip()
    if not raw:
        return None
    return Path(raw).expanduser()


def _process_batch(src: Path, output_arg: str | None, header_text: str | None = None) -> tuple[int, int]:
    """处理一批文件，返回 (成功数, 失败数)。"""
    out = Path(output_arg) if output_arg else (
        src.parent / "formatted" if src.is_file() else src / "formatted")

    if src.is_file():
        files = [src]
    else:
        out_resolved = out.resolve()
        files = sorted(p for p in src.rglob("*.docx")
                       if not p.name.startswith("~$")
                       and out_resolved not in p.resolve().parents)
    if not files:
        print(f"[错误] 未在 {src} 下找到 .docx 文件（旧版 .doc 请先在 Word 中另存为 .docx）。")
        return 0, 0

    ok = fail = 0
    for f in files:
        dst = (out / f.relative_to(src)).resolve() if src.is_dir() else out / f.name
        if dst == f.resolve():
            dst = dst.with_name(dst.stem + "_格式化" + dst.suffix)
        try:
            stats = process_file(f, dst, header_text)
            ok += 1
            detail = " · ".join(f"{k} {v}" for k, v in stats.items())
            print(f"[完成] {f.name}  ->  {dst}   ({detail})")
        except PermissionError:
            fail += 1
            print(f"[失败] {f.name}: 文件被占用，请先关闭 Word 后重试")
        except Exception as exc:
            fail += 1
            print(f"[失败] {f.name}: {exc}")

    print(f"\n处理完毕：成功 {ok} 个，失败 {fail} 个。输出目录：{out.resolve()}")
    return ok, fail


def main(argv=None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    args = _parse_args(argv)

    if args.input is not None:                      # 命令行模式：路径作为参数
        src = Path(args.input)
        if not src.exists():
            print(f"[错误] 路径不存在：{src}")
            return 1
        ok, fail = _process_batch(src, args.output, args.header)
        return 0 if fail == 0 else 2

    # 交互模式：先询问页眉，再输入路径；处理完可继续输入下一个文件夹，直接回车退出
    if args.header is None:
        hint = f"（直接回车使用「{HEADER_TEXT}」）" if HEADER_TEXT else "（直接回车 = 沿用原页眉）"
        try:
            raw = input(f"\n请输入页眉文字{hint}：").strip()
        except EOFError:
            raw = ""
        if raw:
            args.header = raw
        elif HEADER_TEXT:
            args.header = HEADER_TEXT

    total_ok = total_fail = 0
    while True:
        try:
            src = _prompt_path()
        except (EOFError, KeyboardInterrupt):
            print()
            break
        if src is None:
            break
        if not src.exists():
            print(f"[提示] 路径不存在：{src}，请重新输入。")
            continue
        ok, fail = _process_batch(src, args.output, args.header)
        total_ok += ok
        total_fail += fail

    if total_ok or total_fail:
        print(f"\n本次运行共处理：成功 {total_ok} 个，失败 {total_fail} 个。")
    try:
        input("\n按回车键退出...")
    except EOFError:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
