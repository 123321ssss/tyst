#!/usr/bin/env python3
"""
Stage 2: Deterministically build a fully compliant Taiyuan University of Science and
Technology (TYUST) doctoral/master's thesis .docx from thesis_ir.json + metadata.json
on top of the cleaned official TYUST skeleton templates.

Key guarantees:
  - Preserves official Cover (with TYUST logo), Chinese/English Title Pages, and
    Originality/Authorization Declaration verbatim
  - True Word 'oddPage' section breaks for Abstract, ABSTRACT, TOC, Symbols, every Chapter,
    References, Appendices, Achievements, and Acknowledgements
  - Per-section odd/even headers (宋体/TNR 五号 10.5pt, single bottom border)
  - Spec-compliant centered bottom PAGE footers (Times New Roman 五号 10.5pt, no dashes):
    upperRoman for frontmatter (Abstract..TOC/Symbols), decimal starting at 1 from Chapter 1
  - Three-line tables (top/bottom 1.5pt, header bottom 0.75pt, repeat header, cantSplit)
  - Centered equations with right-aligned numbering (ch-seq) via exact tab stops (8cm / 16cm)
  - Lossless OMML equation XML & image embedding with aspect-ratio fitting + keep_with_next
  - Real Word TOC field (TOC \\o "1-3" \\h \\u) matched with cleaned toc 1/2/3 styles
  - Optional standalone spine .docx generation (书脊)
"""

import argparse
import json
import os
import re
import sys
import zipfile
import shutil
from copy import deepcopy
from lxml import etree
from PIL import Image
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
O_NS = "urn:schemas-microsoft-com:office:office"
V_NS = "urn:schemas-microsoft-com:vml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
from math_engine import latex_to_omml, normalize_omml_font, split_inline_math_runs

TEMPLATES_DIR = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "resources", "templates"))

CN_DIGITS = ["", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
             "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十"]

RE_CITATION = re.compile(r"(\[\d+(?:[-–—,，]\d+)*\])")


def to_cn_num(n: int) -> str:
    if 1 <= n < len(CN_DIGITS):
        return CN_DIGITS[n]
    return str(n)


def ensure_child(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def set_rpr_font(rPr, east_asia="宋体", latin="Times New Roman", size_pt=12.0,
                 bold=False, italic=False, sup=False, sub=False):
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), east_asia)
    rFonts.set(qn("w:cs"), latin)

    half_pt = str(int(round(size_pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        el = ensure_child(rPr, tag)
        el.set(qn("w:val"), half_pt)

    for tag in ("w:b", "w:bCs"):
        el = ensure_child(rPr, tag)
        el.set(qn("w:val"), "1" if bold else "0")

    if italic:
        for tag in ("w:i", "w:iCs"):
            el = ensure_child(rPr, tag)
            el.set(qn("w:val"), "1")

    if sup or sub:
        va = ensure_child(rPr, "w:vertAlign")
        va.set(qn("w:val"), "superscript" if sup else "subscript")


def add_run(p, text: str, east_asia="宋体", latin="Times New Roman", size_pt=12.0,
            bold=False, italic=False, sup=False, sub=False):
    r = p.add_run(text)
    rPr = r._r.get_or_add_rPr()
    set_rpr_font(rPr, east_asia=east_asia, latin=latin, size_pt=size_pt,
                 bold=bold, italic=italic, sup=sup, sub=sub)
    return r


def set_para_format(p, align="both", before_pt=0.0, after_pt=0.0,
                    line_pt=None, line_rule="auto",
                    first_line_chars=0, left_chars=0, hanging_chars=0,
                    outline_lvl=None, keep_with_next=False,
                    before_lines=None, after_lines=None):
    pPr = p._p.get_or_add_pPr()

    # Alignment
    jc = ensure_child(pPr, "w:jc")
    jc.set(qn("w:val"), align)

    # Spacing
    sp = ensure_child(pPr, "w:spacing")
    sp.attrib.clear()
    if before_lines is not None:
        sp.set(qn("w:beforeLines"), str(before_lines))
        sp.set(qn("w:before"), str(int(round(before_lines * 3.12))))
    else:
        sp.set(qn("w:before"), str(int(round(before_pt * 20))))
    if after_lines is not None:
        sp.set(qn("w:afterLines"), str(after_lines))
        sp.set(qn("w:after"), str(int(round(after_lines * 3.12))))
    else:
        sp.set(qn("w:after"), str(int(round(after_pt * 20))))

    if line_pt is not None:
        sp.set(qn("w:line"), str(int(round(line_pt * 20))))
        sp.set(qn("w:lineRule"), line_rule)
    else:
        sp.set(qn("w:line"), "240")
        sp.set(qn("w:lineRule"), "auto")

    # Indentation
    ind = ensure_child(pPr, "w:ind")
    ind.attrib.clear()
    # 1 char ≈ 240 dxa at 12pt (or 210 dxa at 10.5pt)
    if left_chars:
        ind.set(qn("w:leftChars"), str(left_chars))
        ind.set(qn("w:left"), str(int(round(left_chars * 2.1 if hanging_chars else left_chars * 2.4))))
    if hanging_chars:
        ind.set(qn("w:hangingChars"), str(hanging_chars))
        ind.set(qn("w:hanging"), str(int(round(hanging_chars * 2.1))))
    elif first_line_chars:
        ind.set(qn("w:firstLineChars"), str(first_line_chars))
        ind.set(qn("w:firstLine"), str(int(round(first_line_chars * 2.4))))
    else:
        ind.set(qn("w:firstLine"), "0")

    # Outline level
    old_ol = pPr.find(qn("w:outlineLvl"))
    if old_ol is not None:
        pPr.remove(old_ol)
    if outline_lvl is not None:
        ol = OxmlElement("w:outlineLvl")
        ol.set(qn("w:val"), str(outline_lvl))
        pPr.append(ol)

    if keep_with_next:
        ensure_child(pPr, "w:keepNext")

    # Disable widow control / snapToGrid mismatches where needed
    snap = pPr.find(qn("w:snapToGrid"))
    if snap is not None:
        pPr.remove(snap)


def add_text_with_math_and_cites(p, text: str, east_asia="宋体", latin="Times New Roman",
                                 size_pt=12.0, bold=False, italic=False, sup=False, sub=False,
                                 auto_sup_citations=True, format_inline_math=True):
    """Render plain text with optional superscript [n] citations and inline math italic/sup/sub."""
    if not text:
        return
    if sup or sub:
        add_run(p, text, east_asia=east_asia, latin=latin, size_pt=size_pt,
                bold=bold, italic=italic, sup=sup, sub=sub)
        return

    if auto_sup_citations and RE_CITATION.search(text):
        parts = RE_CITATION.split(text)
        for idx, part in enumerate(parts):
            if not part:
                continue
            is_cite = (idx % 2 == 1)
            if is_cite:
                add_run(p, part, east_asia=east_asia, latin=latin, size_pt=size_pt,
                        bold=bold, italic=italic, sup=True, sub=False)
            elif format_inline_math:
                for mr in split_inline_math_runs(part):
                    add_run(p, mr["text"], east_asia=east_asia, latin=latin, size_pt=size_pt,
                            bold=bold, italic=(italic or mr["italic"]), sup=mr["sup"], sub=mr["sub"])
            else:
                add_run(p, part, east_asia=east_asia, latin=latin, size_pt=size_pt,
                        bold=bold, italic=italic, sup=False, sub=False)
    elif format_inline_math:
        for mr in split_inline_math_runs(text):
            add_run(p, mr["text"], east_asia=east_asia, latin=latin, size_pt=size_pt,
                    bold=bold, italic=(italic or mr["italic"]), sup=mr["sup"], sub=mr["sub"])
    else:
        add_run(p, text, east_asia=east_asia, latin=latin, size_pt=size_pt,
                bold=bold, italic=italic, sup=False, sub=False)


def add_segments_to_para(p, segments, east_asia="宋体", latin="Times New Roman",
                         size_pt=12.0, auto_sup_citations=True, footnote_collector=None,
                         ole_collector=None, ir_dir=None, format_inline_math=True):
    """Render lossless segments (text runs, OMML math XML, MathType OLE, footnote_refs) into paragraph p."""
    for seg in segments:
        stype = seg.get("type")
        if stype == "omml":
            try:
                omml_el = etree.fromstring(seg["xml"].encode("utf-8"))
                if omml_el.tag == f"{{{M_NS}}}oMathPara":
                    for inner in omml_el.findall(f"{{{M_NS}}}oMath"):
                        normalize_omml_font(inner, font_pt=size_pt)
                        p._p.append(inner)
                else:
                    normalize_omml_font(omml_el, font_pt=size_pt)
                    p._p.append(omml_el)
            except Exception:
                pass
        elif stype == "ole_object":
            if ole_collector is not None and ir_dir:
                r_el = ole_collector.build_run_element(seg, ir_dir, size_pt=size_pt)
                if r_el is not None:
                    p._p.append(r_el)
        elif stype == "footnote_ref":
            fn_id = seg.get("id")
            if footnote_collector is not None and fn_id is not None:
                new_id = footnote_collector.register(fn_id)
                if new_id is not None:
                    r = OxmlElement("w:r")
                    rPr = OxmlElement("w:rPr")
                    set_rpr_font(rPr, east_asia="宋体", latin="Times New Roman", size_pt=size_pt, sup=True)
                    r.append(rPr)
                    fn_ref = OxmlElement("w:footnoteReference")
                    fn_ref.set(qn("w:id"), str(new_id))
                    r.append(fn_ref)
                    p._p.append(r)
        elif stype == "text":
            txt = seg.get("text", "")
            if not txt:
                continue
            add_text_with_math_and_cites(
                p, txt, east_asia=east_asia, latin=latin, size_pt=size_pt,
                bold=seg.get("bold", False), italic=seg.get("italic", False),
                sup=seg.get("sup", False), sub=seg.get("sub", False),
                auto_sup_citations=auto_sup_citations, format_inline_math=format_inline_math,
            )


def replace_para_text_keep_style(p, new_text: str):
    """Replace all text in a cover/title-page paragraph while preserving its exact first-run rPr."""
    first_rPr = None
    for r in p.runs:
        rPr = r._r.find(qn("w:rPr"))
        if rPr is not None:
            first_rPr = deepcopy(rPr)
            break
    # Remove existing runs
    for child in list(p._p):
        if child.tag == qn("w:r"):
            p._p.remove(child)
    # Handle multiline cover titles cleanly with <w:br/>
    lines = new_text.split("\n")
    for i, ln in enumerate(lines):
        r = OxmlElement("w:r")
        if first_rPr is not None:
            r.append(deepcopy(first_rPr))
        t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = ln
        r.append(t)
        if i < len(lines) - 1:
            r.append(OxmlElement("w:br"))
        p._p.append(r)


def populate_frontmatter(doc: Document, meta: dict):
    """Fill Cover (Sec 0) and CN/EN Title Pages (Sec 1) in-place on the official skeleton."""
    title_zh = meta.get("title_zh", "中文题名")
    title_en = meta.get("title_en", "English Title")
    author_zh = meta.get("author_zh", "张三")
    author_en = meta.get("author_en", "Zhang San")
    sup_zh = f"{meta.get('supervisor_zh', '李四')}  {meta.get('supervisor_title_zh', '教授')}".strip()
    sup_en_title = meta.get("supervisor_title_en", "Professor")
    sup_en = f"{sup_en_title}: {meta.get('supervisor_en', 'Li Si')}"
    major_zh = meta.get("major_zh", "XXXXXXXX")
    major_en = meta.get("major_en", "Mechanical Engineering")
    college_zh = meta.get("college_zh", "XXXXXXXX学院")
    college_en = meta.get("college_en", "School of Mechanical Engineering")
    clc = meta.get("clc_code", "TH16")
    udc = meta.get("udc_code", "621")
    school_code = meta.get("school_code", "10109")
    secret = meta.get("secret_level", "公开")
    cover_date = meta.get("cover_date_zh", "2026年5月")
    submit_date = meta.get("submit_date_zh", cover_date)
    defense_date = meta.get("defense_date_zh", "2026年6月")
    chair_zh = f"{meta.get('chair_zh', '王五')}  {meta.get('chair_title_zh', '教授')}".strip()
    submit_date_en = meta.get("submit_date_en", "May, 2026")

    # Replace paragraphs in Sec 0 and Sec 1
    for p in doc.paragraphs:
        t = p.text.strip()
        if t == "中文题名":
            replace_para_text_keep_style(p, title_zh)
            # Enforce spec single line spacing on cover/title page Chinese title
            set_para_format(p, align="center", before_pt=0.0, after_pt=0.0)
        elif t == "English Title":
            replace_para_text_keep_style(p, title_en)
            set_para_format(p, align="center", before_pt=0.0, after_pt=0.0)
        elif re.fullmatch(r"20\d{2}年\d{1,2}月", t):
            replace_para_text_keep_style(p, cover_date)
        elif t == "作者英文姓名":
            replace_para_text_keep_style(p, author_en)
        elif t == "英文专业名称":
            replace_para_text_keep_style(p, major_en)
        elif t.startswith("英文学院名称"):
            replace_para_text_keep_style(p, college_en)
        elif t.startswith("Professor:") or "导师英文姓名" in t:
            replace_para_text_keep_style(p, sup_en)
        elif "英文提交论文月份" in t:
            replace_para_text_keep_style(p, submit_date_en)

    # Table 0: Cover 3x2 table
    if len(doc.tables) >= 1:
        t0 = doc.tables[0]
        if len(t0.rows) >= 3:
            replace_para_text_keep_style(t0.rows[0].cells[1].paragraphs[0], author_zh)
            replace_para_text_keep_style(t0.rows[1].cells[1].paragraphs[0], sup_zh)
            replace_para_text_keep_style(t0.rows[2].cells[1].paragraphs[0], major_zh)

    # Table 1: Chinese Title Page top 2x5 classification table
    if len(doc.tables) >= 2:
        t1 = doc.tables[1]
        if len(t1.rows) >= 2 and len(t1.columns) >= 5:
            replace_para_text_keep_style(t1.rows[0].cells[1].paragraphs[0], clc)
            replace_para_text_keep_style(t1.rows[0].cells[4].paragraphs[0], school_code)
            replace_para_text_keep_style(t1.rows[1].cells[1].paragraphs[0], udc)
            replace_para_text_keep_style(t1.rows[1].cells[4].paragraphs[0], secret)

    # Table 2: Chinese Title Page bottom 7x2 info table
    if len(doc.tables) >= 3:
        t2 = doc.tables[2]
        vals = [author_zh, sup_zh, college_zh, major_zh, submit_date, defense_date, chair_zh]
        for idx, val in enumerate(vals):
            if idx < len(t2.rows):
                replace_para_text_keep_style(t2.rows[idx].cells[1].paragraphs[0], val)

    # Compact excess blank spacer paragraphs on Cover (Sec 0) & CN Title Page (Sec 1)
    # when CN or EN title wraps to multiple lines, so each strictly stays on 1 page!
    zh_lines = max(1, (len(title_zh) + 19) // 20)
    en_lines = max(1, (len(title_en) + 41) // 42)
    extra_lines = (zh_lines - 1) + (en_lines - 1)
    if extra_lines > 0:
        compact_frontmatter_spacers(doc, extra_lines)

    # Guarantee Sec 0, 1, 2 have explicit empty, borderless headers & footers (never inherit Header border)
    for i in range(min(3, len(doc.sections))):
        s = doc.sections[i]
        for part in (s.header, s.even_page_header, s.first_page_header, s.footer, s.even_page_footer, s.first_page_footer):
            part.is_linked_to_previous = False
            clear_hf(part)
            p_empty = OxmlElement("w:p")
            pPr = OxmlElement("w:pPr")
            pBdr = OxmlElement("w:pBdr")
            btm = OxmlElement("w:bottom")
            btm.set(qn("w:val"), "none")
            pBdr.append(btm)
            pPr.append(pBdr)
            p_empty.append(pPr)
            part._element.append(p_empty)


def compact_frontmatter_spacers(doc: Document, extra_lines: int):
    """
    Reclaim excess blank spacer paragraphs in Cover (before Sec 0 break) and Chinese Title Page
    (between Sec 0 break and Page Break before English Title Page) while preserving the single
    blank line between Chinese Title and English Title.
    """
    body = doc.element.body
    children = list(body)

    def is_pure_blank_p(el):
        if el.tag != qn("w:p"):
            return False
        if el.find(".//" + qn("w:drawing")) is not None or el.find(".//" + qn("w:pict")) is not None:
            return False
        if el.find(".//" + qn("w:br")) is not None:
            return False
        pPr = el.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            return False
        txt = "".join(t.text for t in el.findall(".//" + qn("w:t")) if t.text).strip()
        return txt == ""

    def shrink_p(el, pt_val=6.0):
        pPr = ensure_child(el, "w:pPr")
        sp = ensure_child(pPr, "w:spacing")
        sp.attrib.clear()
        sp.set(qn("w:before"), "0")
        sp.set(qn("w:after"), "0")
        sp.set(qn("w:line"), str(int(pt_val * 20)))
        sp.set(qn("w:lineRule"), "exact")

    # Find Sec 0 boundary and English Title Page break
    sec0_idx = None
    en_br_idx = None
    for idx, ch in enumerate(children):
        if ch.tag == qn("w:p"):
            pPr = ch.find(qn("w:pPr"))
            if sec0_idx is None and pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                sec0_idx = idx
            elif sec0_idx is not None and en_br_idx is None:
                if any(b.get(qn("w:type")) == "page" for b in ch.findall(".//" + qn("w:br"))):
                    en_br_idx = idx
                    break

    # 1. Compact Cover (0 .. sec0_idx)
    if sec0_idx is not None:
        cover_blanks = [ch for ch in children[:sec0_idx] if is_pure_blank_p(ch)]
        to_remove = []
        if len(cover_blanks) >= 2:
            to_remove.append(cover_blanks[0])  # top blank 1
        if extra_lines >= 2 and len(cover_blanks) >= 3:
            to_remove.append(cover_blanks[1])  # top blank 2
        if extra_lines >= 2 and len(cover_blanks) >= 8:
            to_remove.append(cover_blanks[-3])  # second blank above cover table
        if extra_lines >= 3 and len(cover_blanks) >= 9:
            to_remove.append(cover_blanks[-2])  # first blank above cover date
        for el in to_remove:
            if el.getparent() is body:
                body.remove(el)

    # 2. Compact Chinese Title Page (sec0_idx+1 .. en_br_idx)
    if sec0_idx is not None and en_br_idx is not None:
        tp_els = children[sec0_idx + 1:en_br_idx]
        tp_blanks = [ch for ch in tp_els if is_pure_blank_p(ch)]
        if len(tp_blanks) >= 6:
            shrink_p(tp_blanks[0], 6.0)   # small gap between T1 and logo
            shrink_p(tp_blanks[3], 10.5)  # 五号 gap between CN & EN title
            remove_candidates = [tp_blanks[1], tp_blanks[4]]
            if extra_lines >= 2:
                remove_candidates.append(tp_blanks[2])
            if extra_lines >= 3 and len(tp_blanks) >= 6:
                remove_candidates.append(tp_blanks[5])
            if len(tp_blanks) >= 7:
                remove_candidates.append(tp_blanks[6])
            for el in remove_candidates:
                if el.getparent() is body:
                    body.remove(el)
        if extra_lines >= 2 and len(doc.tables) >= 3:
            t2 = doc.tables[2]
            for row in t2.rows:
                trPr = row._tr.get_or_add_trPr()
                trH = trPr.find(qn("w:trHeight"))
                if trH is not None:
                    old_v = int(trH.get(qn("w:val")) or 520)
                    trH.set(qn("w:val"), str(min(old_v, 460)))
                for cell in row.cells:
                    for cp in cell.paragraphs:
                        pPr = cp._p.get_or_add_pPr()
                        sp = ensure_child(pPr, "w:spacing")
                        sp.set(qn("w:before"), "0")
                        sp.set(qn("w:after"), "0")
                        sp.set(qn("w:line"), "260")
                        sp.set(qn("w:lineRule"), "auto")


def clear_hf(hf):
    el = hf._element
    for ch in list(el):
        el.remove(ch)


def set_header_content(hdr, text: str):
    """Set centered 宋体 / Times New Roman 五号 (10.5pt) header with single bottom border."""
    clear_hf(hdr)
    p_el = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    pPr.append(jc)
    pBdr = OxmlElement("w:pBdr")
    btm = OxmlElement("w:bottom")
    btm.set(qn("w:val"), "single")
    btm.set(qn("w:color"), "auto")
    btm.set(qn("w:sz"), "6")
    btm.set(qn("w:space"), "1")
    pBdr.append(btm)
    pPr.append(pBdr)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "240")
    sp.set(qn("w:lineRule"), "auto")
    pPr.append(sp)
    p_el.append(pPr)

    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    set_rpr_font(rPr, east_asia="宋体", latin="Times New Roman", size_pt=10.5, bold=False)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    r.append(t)
    p_el.append(r)
    hdr._element.append(p_el)


def set_footer_page_number(ftr):
    """Set clean bottom-centered PAGE field in Times New Roman 五号 (10.5pt), no dashes, no textbox."""
    clear_hf(ftr)
    p_el = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    pPr.append(jc)
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"), "0")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "240")
    sp.set(qn("w:lineRule"), "auto")
    pPr.append(sp)
    p_el.append(pPr)

    def make_field_run(kind, text=None):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        set_rpr_font(rPr, east_asia="Times New Roman", latin="Times New Roman", size_pt=10.5, bold=False)
        r.append(rPr)
        if kind == "fldChar":
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), text)
            r.append(fc)
        elif kind == "instrText":
            it = OxmlElement("w:instrText")
            it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            it.text = text
            r.append(it)
        elif kind == "t":
            t = OxmlElement("w:t")
            t.text = text
            r.append(t)
        return r

    p_el.append(make_field_run("fldChar", "begin"))
    p_el.append(make_field_run("instrText", " PAGE "))
    p_el.append(make_field_run("fldChar", "separate"))
    p_el.append(make_field_run("t", "1"))
    p_el.append(make_field_run("fldChar", "end"))
    ftr._element.append(p_el)


def configure_section(sec, odd_hdr_text: str, even_hdr_text: str,
                      pg_fmt: str = "decimal", pg_start: int = None, is_odd_page: bool = True):
    """Apply TYUST A4 margins, oddPage break, odd/even headers, and centered 10.5pt TNR PAGE footer."""
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.54)
    sec.bottom_margin = Cm(2.54)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(2.0)
    sec.header_distance = Cm(1.9)
    sec.footer_distance = Cm(2.2)
    sec.different_first_page_header_footer = False

    if is_odd_page:
        tp = ensure_child(sec._sectPr, "w:type")
        tp.set(qn("w:val"), "oddPage")

    sec.header.is_linked_to_previous = False
    sec.even_page_header.is_linked_to_previous = False
    sec.footer.is_linked_to_previous = False
    sec.even_page_footer.is_linked_to_previous = False

    set_header_content(sec.header, odd_hdr_text)
    set_header_content(sec.even_page_header, even_hdr_text)
    set_footer_page_number(sec.footer)
    set_footer_page_number(sec.even_page_footer)

    pg = sec._sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        sec._sectPr.append(pg)
    pg.attrib.clear()
    pg.set(qn("w:fmt"), pg_fmt)
    if pg_start is not None:
        pg.set(qn("w:start"), str(pg_start))


def insert_toc_field(doc: Document):
    """Insert a real Word TOC field (TOC \\o '1-3' \\h \\u) updated later by Word/WPS COM."""
    p = doc.add_paragraph()
    set_para_format(p, align="both", before_pt=6.0, after_pt=0.0)
    p_el = p._p

    def fld_run(kind, val):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        set_rpr_font(rPr, east_asia="宋体", latin="Times New Roman", size_pt=12.0, bold=False)
        r.append(rPr)
        if kind == "fldChar":
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), val)
            if val == "begin":
                fc.set(qn("w:dirty"), "true")
            r.append(fc)
        elif kind == "instrText":
            it = OxmlElement("w:instrText")
            it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            it.text = val
            r.append(it)
        elif kind == "t":
            t = OxmlElement("w:t")
            t.text = val
            r.append(t)
        return r

    p_el.append(fld_run("fldChar", "begin"))
    p_el.append(fld_run("instrText", ' TOC \\o "1-3" \\h \\z \\u '))
    p_el.append(fld_run("fldChar", "separate"))
    p_el.append(fld_run("t", "请右键选择“更新域”或运行 finalize_word.py 自动刷新目录"))
    p_el.append(fld_run("fldChar", "end"))


def _prevent_adjacent_table_merge(doc: Document):
    """
    Word automatically merges two consecutive <w:tbl> elements unless separated by a <w:p>.
    Insert a 1pt-tall invisible separator paragraph when the previous body block is a table.
    """
    body = doc.element.body
    for ch in reversed(list(body)):
        if ch.tag == qn("w:sectPr"):
            continue
        if ch.tag == qn("w:tbl"):
            p_sep = doc.add_paragraph()
            pPr = p_sep._p.get_or_add_pPr()
            sp = ensure_child(pPr, "w:spacing")
            sp.attrib.clear()
            sp.set(qn("w:before"), "0")
            sp.set(qn("w:after"), "0")
            sp.set(qn("w:line"), "20")  # 1pt exact
            sp.set(qn("w:lineRule"), "exact")
            rPr = ensure_child(pPr, "w:rPr")
            sz = ensure_child(rPr, "w:sz")
            sz.set(qn("w:val"), "2")
        break


def add_three_line_table(doc: Document, rows: list, caption_zh: str, caption_en: str,
                         tbl_num: str, notes: list, bilingual: bool = False, format_inline_math: bool = True):
    """Render a TYUST-compliant three-line table (top/bottom 1.5pt, header bottom 0.75pt, no vertical borders)."""
    # 1. Caption above table
    p_cap = doc.add_paragraph()
    set_para_format(p_cap, align="center", before_pt=6.0, after_pt=6.0, keep_with_next=True)
    cap_txt = f"表{tbl_num} {caption_zh}".strip() if caption_zh else f"表{tbl_num}"
    add_text_with_math_and_cites(p_cap, cap_txt, east_asia="楷体", latin="Times New Roman",
                                 size_pt=10.5, auto_sup_citations=False, format_inline_math=format_inline_math)

    if bilingual and caption_en:
        p_cap_en = doc.add_paragraph()
        set_para_format(p_cap_en, align="center", before_pt=0.0, after_pt=6.0, keep_with_next=True)
        add_run(p_cap_en, f"Table {tbl_num} {caption_en}", east_asia="楷体", latin="Times New Roman", size_pt=10.5)

    if not rows:
        return

    _prevent_adjacent_table_merge(doc)
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    tbl = doc.add_table(rows=n_rows, cols=n_cols)

    # Table-level properties: centered, fixed full width 16.0 cm (9072 dxa), three-line borders
    tblPr = tbl._tbl.tblPr
    jc = ensure_child(tblPr, "w:jc")
    jc.set(qn("w:val"), "center")

    tblW = ensure_child(tblPr, "w:tblW")
    tblW.set(qn("w:w"), "9072")
    tblW.set(qn("w:type"), "dxa")

    tblLayout = ensure_child(tblPr, "w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")

    # Cell margins (top/bottom 60 dxa ≈ 3pt, left/right 80 dxa ≈ 4pt)
    tblCellMar = ensure_child(tblPr, "w:tblCellMar")
    tblCellMar.clear()
    for side, val in (("top", "60"), ("bottom", "60"), ("left", "80"), ("right", "80")):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), val)
        m.set(qn("w:type"), "dxa")
        tblCellMar.append(m)

    # Table borders: top=12 (1.5pt), bottom=12 (1.5pt), others none
    tblBorders = ensure_child(tblPr, "w:tblBorders")
    tblBorders.clear()
    for side, val, sz in (
        ("top", "single", "12"),
        ("bottom", "single", "12"),
        ("left", "none", "0"),
        ("right", "none", "0"),
        ("insideH", "none", "0"),
        ("insideV", "none", "0"),
    ):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), val)
        b.set(qn("w:sz"), sz)
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        tblBorders.append(b)

    # Compute proportional column widths based on average/max text length (visual width weighted)
    col_weights = []
    for c in range(n_cols):
        lens = []
        for r in rows:
            txt = r[c] if c < len(r) else ""
            # Approximate visual width: CJK=2, ASCII=1
            vw = sum(2 if ord(ch) > 127 else 1 for ch in txt.replace("\n", " "))
            lens.append(vw)
        # Use damped max so one huge cell doesn't crush narrow columns
        w = max(8, min(60, int(round(max(lens) * 0.65 + (sum(lens) / max(1, len(lens))) * 0.35))))
        col_weights.append(w)
    total_w = sum(col_weights)
    col_widths_dxa = [max(900, int(round(9072 * w / total_w))) for w in col_weights]
    # Re-normalize sum to 9072
    diff = 9072 - sum(col_widths_dxa)
    col_widths_dxa[-1] += diff

    # Update tblGrid
    tblGrid = tbl._tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        tblGrid.clear()
        for w_dxa in col_widths_dxa:
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(w_dxa))
            tblGrid.append(gc)

    # Font size: 10.5pt (五号) for <=4 cols, 9pt (小五) for >=5 cols so wide tables never overflow
    cell_font_pt = 9.0 if n_cols >= 5 else 10.5

    for r_idx, row_vals in enumerate(rows):
        row = tbl.rows[r_idx]
        trPr = row._tr.get_or_add_trPr()
        ensure_child(trPr, "w:cantSplit")
        if r_idx == 0:
            ensure_child(trPr, "w:tblHeader")

        for c_idx in range(n_cols):
            cell = row.cells[c_idx]
            val = row_vals[c_idx] if c_idx < len(row_vals) else ""
            tcPr = cell._tc.get_or_add_tcPr()

            tcW = ensure_child(tcPr, "w:tcW")
            tcW.set(qn("w:w"), str(col_widths_dxa[c_idx]))
            tcW.set(qn("w:type"), "dxa")

            vAlign = ensure_child(tcPr, "w:vAlign")
            vAlign.set(qn("w:val"), "center")

            # Explicit cell borders for clean rendering in both Word and WPS
            tcBdr = ensure_child(tcPr, "w:tcBorders")
            tcBdr.clear()
            top_b = ("single", "12") if r_idx == 0 else ("nil", "0")
            if r_idx == 0:
                btm_b = ("single", "6")  # 0.75pt header rule
            elif r_idx == n_rows - 1:
                btm_b = ("single", "12")  # 1.5pt bottom rule
            else:
                btm_b = ("nil", "0")
            for side, (bval, bsz) in (
                ("top", top_b),
                ("bottom", btm_b),
                ("left", ("nil", "0")),
                ("right", ("nil", "0")),
            ):
                b = OxmlElement(f"w:{side}")
                b.set(qn("w:val"), bval)
                if bval != "nil":
                    b.set(qn("w:sz"), bsz)
                    b.set(qn("w:space"), "0")
                    b.set(qn("w:color"), "auto")
                tcBdr.append(b)

            # Populate cell paragraphs
            lines = val.split("\n") if val else [""]
            p0 = cell.paragraphs[0]
            for li, ln in enumerate(lines):
                cp = p0 if li == 0 else cell.add_paragraph()
                # Header row centered; long descriptive cells left-aligned, short cells centered
                c_align = "center" if (r_idx == 0 or len(val) <= 18) else "left"
                set_para_format(cp, align=c_align, before_pt=1.0, after_pt=1.0)
                add_text_with_math_and_cites(cp, ln, east_asia="宋体", latin="Times New Roman",
                                             size_pt=cell_font_pt, bold=(r_idx == 0),
                                             auto_sup_citations=False, format_inline_math=format_inline_math)

    # Notes below table (小五 9pt 宋体, 两端对齐, 首行缩进 2 字符, 单倍行距)
    for note in notes:
        p_n = doc.add_paragraph()
        set_para_format(p_n, align="both", before_pt=2.0, after_pt=2.0, first_line_chars=200)
        add_text_with_math_and_cites(p_n, note, east_asia="宋体", latin="Times New Roman",
                                     size_pt=9.0, auto_sup_citations=False, format_inline_math=format_inline_math)


def add_figure_block(doc: Document, fig_blk: dict, ir_dir: str, fig_num: str,
                     bilingual: bool = False, format_inline_math: bool = True):
    """Insert figure image(s) fitted within 14.8cm x 18.0cm + centered 楷体五号 caption below."""
    imgs = fig_blk.get("images", [])
    for img_rel in imgs:
        img_abs = os.path.join(ir_dir, img_rel) if not os.path.isabs(img_rel) else img_rel
        if not os.path.exists(img_abs):
            continue
        p_img = doc.add_paragraph()
        set_para_format(p_img, align="center", before_pt=6.0, after_pt=3.0, keep_with_next=True)
        # Convert WMF/EMF figures to high-DPI PNG because python-docx add_picture does not recognize WMF/EMF
        embed_path = img_abs
        if img_abs.lower().endswith((".emf", ".wmf")):
            try:
                png_cand = os.path.splitext(img_abs)[0] + "_conv.png"
                with Image.open(img_abs) as im_vec:
                    im_vec.load(dpi=300)
                    im_vec.save(png_cand, "PNG")
                embed_path = png_cand
            except Exception:
                continue
        # Calculate safe dimensions preserving aspect ratio
        max_w_cm = 14.8
        max_h_cm = 17.5
        try:
            with Image.open(embed_path) as im:
                w_px, h_px = im.size
                dpi_x, dpi_y = im.info.get("dpi", (96, 96))
                if not dpi_x or dpi_x < 50:
                    dpi_x = 96
                if not dpi_y or dpi_y < 50:
                    dpi_y = 96
                w_cm = w_px / dpi_x * 2.54
                h_cm = h_px / dpi_y * 2.54
                # If image is reasonably wide, scale nicely to max_w_cm (or preserve if small icon)
                target_w = min(max_w_cm, max(w_cm, 11.5)) if w_px >= 600 else min(max_w_cm, w_cm)
                scale = target_w / max(w_cm, 0.01)
                target_h = h_cm * scale
                if target_h > max_h_cm:
                    scale = max_h_cm / max(h_cm, 0.01)
                    target_w = w_cm * scale
                p_img.add_run().add_picture(embed_path, width=Cm(round(target_w, 2)))
        except Exception:
            p_img.add_run().add_picture(embed_path, width=Cm(14.0))

    # Figure caption below image
    cap_zh = fig_blk.get("caption_zh", "").strip()
    p_cap = doc.add_paragraph()
    set_para_format(p_cap, align="center", before_pt=3.0, after_pt=6.0)
    cap_txt = f"图{fig_num} {cap_zh}".strip() if cap_zh else f"图{fig_num}"
    add_text_with_math_and_cites(p_cap, cap_txt, east_asia="楷体", latin="Times New Roman",
                                 size_pt=10.5, auto_sup_citations=True, format_inline_math=format_inline_math)

    if bilingual and fig_blk.get("caption_en"):
        p_cap_en = doc.add_paragraph()
        set_para_format(p_cap_en, align="center", before_pt=0.0, after_pt=6.0)
        add_run(p_cap_en, f"Fig. {fig_num} {fig_blk['caption_en']}", east_asia="楷体", latin="Times New Roman", size_pt=10.5)

    if fig_blk.get("subcaption_zh"):
        p_sub = doc.add_paragraph()
        set_para_format(p_sub, align="center", before_pt=0.0, after_pt=4.0)
        add_text_with_math_and_cites(p_sub, fig_blk["subcaption_zh"], east_asia="楷体", latin="Times New Roman",
                                     size_pt=10.5, auto_sup_citations=False, format_inline_math=format_inline_math)
    if bilingual and fig_blk.get("subcaption_en"):
        p_sub_en = doc.add_paragraph()
        set_para_format(p_sub_en, align="center", before_pt=0.0, after_pt=4.0)
        add_run(p_sub_en, fig_blk["subcaption_en"], east_asia="楷体", latin="Times New Roman", size_pt=10.5)

    for note in fig_blk.get("notes", []):
        p_n = doc.add_paragraph()
        set_para_format(p_n, align="center", before_pt=0.0, after_pt=2.0)
        add_text_with_math_and_cites(p_n, note, east_asia="宋体", latin="Times New Roman",
                                     size_pt=9.0, auto_sup_citations=False, format_inline_math=format_inline_math)


class OleObjectCollector:
    """Collect MathType / OLE equation objects and rewrite them cleanly into the output docx."""
    def __init__(self):
        self.items = []  # dicts: img_rid, img_zip, img_abs, ole_rid, ole_zip, ole_abs
        self.seq = 0

    def build_run_element(self, seg: dict, ir_dir: str, size_pt: float = 12.0):
        ole_rel = seg.get("ole_path")
        if not ole_rel:
            return None
        ole_abs = os.path.join(ir_dir, ole_rel) if not os.path.isabs(ole_rel) else ole_rel
        if not os.path.isfile(ole_abs):
            return None

        img_rel = seg.get("image_path")
        img_abs = None
        if img_rel:
            cand = os.path.join(ir_dir, img_rel) if not os.path.isabs(img_rel) else img_rel
            if os.path.isfile(cand):
                img_abs = cand

        self.seq += 1
        ole_ext = os.path.splitext(ole_abs)[1] or ".bin"
        ole_rid = f"rId98{self.seq * 2 - 1:03d}"
        ole_zip = f"embeddings/tyust_ole_{self.seq:03d}{ole_ext}"

        img_rid = None
        img_zip = None
        if img_abs:
            img_ext = os.path.splitext(img_abs)[1] or ".wmf"
            img_rid = f"rId98{self.seq * 2:03d}"
            img_zip = f"media/tyust_ole_{self.seq:03d}{img_ext}"

        self.items.append({
            "ole_rid": ole_rid,
            "ole_zip": ole_zip,
            "ole_abs": ole_abs,
            "img_rid": img_rid,
            "img_zip": img_zip,
            "img_abs": img_abs,
        })

        obj_el = etree.fromstring(seg["xml"].encode("utf-8"))
        # Rewrite r:id and unique ShapeID / ObjectID
        shape_id = f"_x0000_i{9000 + self.seq}"
        for shp in obj_el.findall(f".//{{{V_NS}}}shape"):
            shp.set("id", shape_id)
        for imd in obj_el.findall(f".//{{{V_NS}}}imagedata"):
            if img_rid:
                imd.set(f"{{{R_NS}}}id", img_rid)
        for ole in obj_el.findall(f".//{{{O_NS}}}OLEObject"):
            ole.set(f"{{{R_NS}}}id", ole_rid)
            ole.set("ShapeID", shape_id)
            ole.set("ObjectID", f"_{1900000000 + self.seq}")

        r_el = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        if seg.get("position"):
            pos = OxmlElement("w:position")
            pos.set(qn("w:val"), str(seg["position"]))
            rPr.append(pos)
        set_rpr_font(rPr, east_asia="宋体", latin="Times New Roman", size_pt=size_pt)
        r_el.append(rPr)
        r_el.append(obj_el)
        return r_el


def inject_ole_objects_into_docx(docx_path: str, collector: OleObjectCollector):
    """Write embedded OLE binaries, preview images, rels, and Content_Types into the .docx archive."""
    if not collector.items:
        return
    tmp_path = docx_path + ".oletmp"
    rel_img_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    rel_ole_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"
    ext_mime = {
        "bin": "application/vnd.openxmlformats-officedocument.oleObject",
        "wmf": "image/x-wmf",
        "emf": "image/x-emf",
        "png": "image/png",
    }
    needed_exts = set()
    for it in collector.items:
        needed_exts.add(os.path.splitext(it["ole_zip"])[1].lstrip(".").lower())
        if it["img_zip"]:
            needed_exts.add(os.path.splitext(it["img_zip"])[1].lstrip(".").lower())

    with zipfile.ZipFile(docx_path, "r") as zin:
        rels_root = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
        existing_rids = {el.get("Id") for el in rels_root}
        for it in collector.items:
            if it["ole_rid"] not in existing_rids:
                rel = etree.SubElement(rels_root, "Relationship")
                rel.set("Id", it["ole_rid"])
                rel.set("Type", rel_ole_type)
                rel.set("Target", it["ole_zip"])
            if it["img_rid"] and it["img_rid"] not in existing_rids:
                rel = etree.SubElement(rels_root, "Relationship")
                rel.set("Id", it["img_rid"])
                rel.set("Type", rel_img_type)
                rel.set("Target", it["img_zip"])
        rels_bytes = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        ct_root = etree.fromstring(zin.read("[Content_Types].xml"))
        existing_exts = {
            el.get("Extension", "").lower()
            for el in ct_root
            if el.tag.endswith("Default")
        }
        for ext in sorted(needed_exts):
            if ext and ext not in existing_exts and ext in ext_mime:
                d_el = etree.SubElement(ct_root, "Default")
                d_el.set("Extension", ext)
                d_el.set("ContentType", ext_mime[ext])
        ct_bytes = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/_rels/document.xml.rels":
                    zout.writestr(item, rels_bytes)
                elif item.filename == "[Content_Types].xml":
                    zout.writestr(item, ct_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))
            for it in collector.items:
                with open(it["ole_abs"], "rb") as f:
                    zout.writestr("word/" + it["ole_zip"], f.read())
                if it["img_abs"] and it["img_zip"]:
                    with open(it["img_abs"], "rb") as f:
                        zout.writestr("word/" + it["img_zip"], f.read())

    shutil.move(tmp_path, docx_path)


def _split_top_level_latex_lines(raw_tex: str) -> list:
    """Split a LaTeX string on top-level `\\\\` (brace depth == 0), preserving `\\substack{... \\\\ ...}`."""
    s = (raw_tex or "").strip()
    if not s:
        return []
    lines = []
    buf = []
    depth = 0
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if ch == "\\" and i + 1 < n:
            nxt = s[i + 1]
            if nxt == "\\" and depth == 0:
                part = "".join(buf).strip()
                if part:
                    lines.append(part)
                buf = []
                i += 2
                continue
            buf.append(ch)
            buf.append(nxt)
            i += 2
            continue
        if ch == "{":
            depth += 1
        elif ch == "}" and depth > 0:
            depth -= 1
        buf.append(ch)
        i += 1
    tail = "".join(buf).strip()
    if tail:
        lines.append(tail)
    return lines


def add_equation_block(doc: Document, eq_blk: dict, eq_num_str: str,
                       ole_collector: OleObjectCollector = None, ir_dir: str = None,
                       format_inline_math: bool = True):
    """
    Render a TYUST equation using a borderless 3-column layout table (total width = 16.0 cm = 9072 dxa):
      - Col 0 (1.85 cm = 1050 dxa): symmetric blank spacer
      - Col 1 (12.30 cm = 6972 dxa): centered formula (LaTeX->OMML, native OMML, MathType OLE, or formatted fallback)
      - Col 2 (1.85 cm = 1050 dxa): right-aligned equation number (e.g. （2-1）), vertically centered
    """
    _prevent_adjacent_table_merge(doc)
    tbl = doc.add_table(rows=1, cols=3)

    tblPr = tbl._tbl.tblPr
    jc = ensure_child(tblPr, "w:jc")
    jc.set(qn("w:val"), "center")

    tblW = ensure_child(tblPr, "w:tblW")
    tblW.set(qn("w:w"), "9072")
    tblW.set(qn("w:type"), "dxa")

    tblLayout = ensure_child(tblPr, "w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")

    tblCellMar = ensure_child(tblPr, "w:tblCellMar")
    tblCellMar.clear()
    for side, val in (("top", "40"), ("bottom", "40"), ("left", "0"), ("right", "0")):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), val)
        m.set(qn("w:type"), "dxa")
        tblCellMar.append(m)

    tblBorders = ensure_child(tblPr, "w:tblBorders")
    tblBorders.clear()
    for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        tblBorders.append(b)

    col_widths_dxa = [1050, 6972, 1050]
    tblGrid = tbl._tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        tblGrid.clear()
        for w_dxa in col_widths_dxa:
            gc = OxmlElement("w:gridCol")
            gc.set(qn("w:w"), str(w_dxa))
            tblGrid.append(gc)

    row = tbl.rows[0]
    trPr = row._tr.get_or_add_trPr()
    ensure_child(trPr, "w:cantSplit")

    for c_idx, w_dxa in enumerate(col_widths_dxa):
        cell = row.cells[c_idx]
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = ensure_child(tcPr, "w:tcW")
        tcW.set(qn("w:w"), str(w_dxa))
        tcW.set(qn("w:type"), "dxa")
        vAlign = ensure_child(tcPr, "w:vAlign")
        vAlign.set(qn("w:val"), "center")
        tcBdr = ensure_child(tcPr, "w:tcBorders")
        tcBdr.clear()
        for side in ("top", "bottom", "left", "right"):
            b = OxmlElement(f"w:{side}")
            b.set(qn("w:val"), "nil")
            tcBdr.append(b)

    # Col 0: empty spacer paragraph
    p0 = row.cells[0].paragraphs[0]
    set_para_format(p0, align="center", before_pt=0.0, after_pt=0.0)
    add_run(p0, "", size_pt=12.0)

    # Col 2: right-aligned equation number
    p2 = row.cells[2].paragraphs[0]
    set_para_format(p2, align="right", before_pt=0.0, after_pt=0.0)
    add_run(p2, eq_num_str, east_asia="宋体", latin="Times New Roman", size_pt=12.0)

    # Col 1: centered equation content
    col1 = row.cells[1]
    latex_lines = eq_blk.get("latex_lines")
    if not latex_lines and eq_blk.get("latex"):
        latex_lines = _split_top_level_latex_lines(eq_blk["latex"])

    if latex_lines:
        is_multi = len(latex_lines) > 1
        for li, tex in enumerate(latex_lines):
            cp = col1.paragraphs[0] if li == 0 else col1.add_paragraph()
            set_para_format(
                cp,
                align="center",
                before_pt=(2.0 if is_multi else 4.0),
                after_pt=(2.0 if is_multi else 4.0),
                line_pt=18.0,
                line_rule="atLeast",
            )
            omml_el = latex_to_omml(tex, font_pt=12.0)
            cp._p.append(omml_el)
    else:
        segs = eq_blk.get("segments", [])
        has_rich = any(s.get("type") in ("omml", "ole_object") for s in segs)
        cp = col1.paragraphs[0]
        set_para_format(cp, align="center", before_pt=4.0, after_pt=4.0, line_pt=18.0, line_rule="atLeast")
        if has_rich:
            add_segments_to_para(
                cp, segs, east_asia="宋体", latin="Times New Roman", size_pt=12.0,
                auto_sup_citations=False, ole_collector=ole_collector, ir_dir=ir_dir,
                format_inline_math=False,
            )
        else:
            eq_txt = eq_blk.get("text", "")
            add_text_with_math_and_cites(
                cp, eq_txt, east_asia="宋体", latin="Times New Roman", size_pt=12.0,
                auto_sup_citations=False, format_inline_math=format_inline_math,
            )


class FootnoteCollector:
    """Collect footnotes referenced in body and write them into word/footnotes.xml post-save."""
    def __init__(self, source_footnotes: dict):
        self.source = source_footnotes or {}
        self.items = []  # [(new_id, text)]
        self.next_id = 1

    def register(self, orig_id: str):
        txt = self.source.get(str(orig_id))
        if not txt:
            return None
        nid = self.next_id
        self.next_id += 1
        self.items.append((nid, txt))
        return nid


def inject_footnotes_into_docx(docx_path: str, collector: FootnoteCollector):
    """Write collected footnotes cleanly into word/footnotes.xml with 小五 (9pt) 宋体/TNR left-aligned."""
    if not collector.items:
        return
    tmp_path = docx_path + ".fntmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        has_fn = "word/footnotes.xml" in zin.namelist()
        if has_fn:
            fn_root = etree.fromstring(zin.read("word/footnotes.xml"))
            # Remove any existing non-separator footnotes
            for fn in list(fn_root.findall(qn("w:footnote"))):
                if fn.get(qn("w:type")) not in ("separator", "continuationSeparator"):
                    fn_root.remove(fn)
        else:
            fn_root = etree.Element(qn("w:footnotes"), nsmap={"w": W_NS})
            for fid, ftype, tag in (("-1", "separator", "w:separator"), ("0", "continuationSeparator", "w:continuationSeparator")):
                f_el = etree.SubElement(fn_root, qn("w:footnote"))
                f_el.set(qn("w:type"), ftype)
                f_el.set(qn("w:id"), fid)
                p = etree.SubElement(f_el, qn("w:p"))
                r = etree.SubElement(p, qn("w:r"))
                etree.SubElement(r, qn(tag))

        for nid, txt in collector.items:
            f_el = etree.SubElement(fn_root, qn("w:footnote"))
            f_el.set(qn("w:id"), str(nid))
            p = etree.SubElement(f_el, qn("w:p"))
            pPr = etree.SubElement(p, qn("w:pPr"))
            jc = etree.SubElement(pPr, qn("w:jc"))
            jc.set(qn("w:val"), "left")
            sp = etree.SubElement(pPr, qn("w:spacing"))
            sp.set(qn("w:before"), "0")
            sp.set(qn("w:after"), "0")
            sp.set(qn("w:line"), "240")
            sp.set(qn("w:lineRule"), "auto")

            # Footnote ref mark
            r_ref = etree.SubElement(p, qn("w:r"))
            rPr1 = etree.SubElement(r_ref, qn("w:rPr"))
            set_rpr_font(rPr1, east_asia="宋体", latin="Times New Roman", size_pt=9.0, sup=True)
            etree.SubElement(r_ref, qn("w:footnoteRef"))

            # Space + footnote body text
            r_txt = etree.SubElement(p, qn("w:r"))
            rPr2 = etree.SubElement(r_txt, qn("w:rPr"))
            set_rpr_font(rPr2, east_asia="宋体", latin="Times New Roman", size_pt=9.0)
            t = etree.SubElement(r_txt, qn("w:t"))
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = " " + txt

        fn_bytes = etree.tostring(fn_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/footnotes.xml":
                    zout.writestr(item, fn_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))
            if not has_fn:
                zout.writestr("word/footnotes.xml", fn_bytes)

    shutil.move(tmp_path, docx_path)


def select_skeleton_template(meta: dict) -> str:
    lvl = meta.get("degree_level", "master")
    dtype = meta.get("degree_type", "academic")
    disc = meta.get("discipline", "science")

    lvl_code = "doc" if lvl == "doctor" else "mas"
    type_code = "pro" if dtype == "professional" else "aca"
    # Note: official templates provide humanities for master; doctoral humanities reuses doctoral skeleton
    disc_code = "hum" if (disc == "humanities" and lvl == "master") else "sci"
    fname = f"{lvl_code}_{disc_code}_{type_code}.docx"
    path = os.path.join(TEMPLATES_DIR, fname)
    if not os.path.exists(path):
        raise FileNotFoundError(f"Skeleton template not found: {path}. Run prepare_templates.py first.")
    return path


def format_chapter_heading(ch_num: int, title: str, discipline: str, language: str) -> str:
    if language == "en":
        return f"Chapter {ch_num} {title}".strip()
    if discipline == "humanities":
        return f"{to_cn_num(ch_num)}、{title}".strip()
    return f"{ch_num}  {title}".strip()


def format_sub_heading(level: int, ch_num: int, sec_idx: int, subsec_idx: int,
                       title: str, discipline: str, language: str) -> str:
    if language == "en":
        num = f"{ch_num}.{sec_idx}" if level == 2 else f"{ch_num}.{sec_idx}.{subsec_idx}"
        return f"{num} {title}".strip()
    if discipline == "humanities":
        if level == 2:
            return f"（{to_cn_num(sec_idx)}）{title}".strip()
        return f"{subsec_idx}.{title}".strip()
    # Science: 2 half-width spaces between number and title
    num = f"{ch_num}.{sec_idx}" if level == 2 else f"{ch_num}.{sec_idx}.{subsec_idx}"
    return f"{num}  {title}".strip()


def build_spine(meta: dict, spine_out: str):
    tpl = os.path.join(TEMPLATES_DIR, "spine_template.docx")
    if not os.path.exists(tpl):
        return
    doc = Document(tpl)
    title_zh = meta.get("title_zh", "学位论文题名")
    author_zh = meta.get("author_zh", "姓名")
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.text = r.text.replace("{{SPINE_TITLE}}", title_zh).replace("{{SPINE_AUTHOR}}", author_zh)
    os.makedirs(os.path.dirname(os.path.abspath(spine_out)), exist_ok=True)
    doc.save(spine_out)
    print(f"[OK] Spine saved -> {spine_out}")


def build_thesis(ir_path: str, meta_path: str, output_path: str,
                 h1_align_override=None, bilingual_override=None, spine_output=None,
                 format_inline_math: bool = True):
    with open(ir_path, "r", encoding="utf-8") as f:
        ir = json.load(f)
    with open(meta_path, "r", encoding="utf-8") as f:
        meta = json.load(f)

    ir_dir = os.path.dirname(os.path.abspath(ir_path))
    discipline = meta.get("discipline", "science")
    language = meta.get("language", "zh")
    degree_level = meta.get("degree_level", "master")
    h1_align = h1_align_override or meta.get("h1_align", "center")
    bilingual = bilingual_override if bilingual_override is not None else bool(meta.get("bilingual_captions", False))

    if language == "en":
        even_hdr = (
            "Doctoral Dissertation of Taiyuan University of Science and Technology"
            if degree_level == "doctor"
            else "Master's Thesis of Taiyuan University of Science and Technology"
        )
    else:
        even_hdr = "太原科技大学博士学位论文" if degree_level == "doctor" else "太原科技大学硕士学位论文"

    tpl_path = select_skeleton_template(meta)
    doc = Document(tpl_path)
    populate_frontmatter(doc, meta)

    fn_collector = FootnoteCollector(ir.get("footnotes", {}))
    ole_collector = OleObjectCollector()

    # Track whether we have used the skeleton's pre-existing trailing section (Sec 3)
    used_initial_sec3 = False

    def begin_odd_section(odd_hdr_text: str, pg_fmt: str = "decimal", pg_start: int = None):
        nonlocal used_initial_sec3
        if not used_initial_sec3:
            sec = doc.sections[-1]
            used_initial_sec3 = True
        else:
            sec = doc.add_section(WD_SECTION_START.ODD_PAGE)
        configure_section(sec, odd_hdr_text=odd_hdr_text, even_hdr_text=even_hdr,
                          pg_fmt=pg_fmt, pg_start=pg_start, is_odd_page=True)
        return sec

    def add_major_title(text: str, is_english_bold: bool = False, in_toc: bool = True, align: str = "center"):
        p = doc.add_paragraph()
        set_para_format(p, align=align, before_pt=24.0, after_pt=18.0,
                        outline_lvl=(0 if in_toc else None), keep_with_next=True)
        add_run(p, text, east_asia="黑体", latin="Times New Roman",
                size_pt=16.0, bold=is_english_bold)
        return p

    def add_body_para(text: str = "", segments=None):
        p = doc.add_paragraph()
        set_para_format(p, align="both", before_pt=0.0, after_pt=0.0,
                        line_pt=20.0, line_rule="exact", first_line_chars=200)
        if segments:
            add_segments_to_para(
                p, segments, east_asia="宋体", latin="Times New Roman",
                size_pt=12.0, auto_sup_citations=True, footnote_collector=fn_collector,
                ole_collector=ole_collector, ir_dir=ir_dir, format_inline_math=format_inline_math,
            )
        else:
            add_text_with_math_and_cites(
                p, text, east_asia="宋体", latin="Times New Roman",
                size_pt=12.0, auto_sup_citations=True, format_inline_math=format_inline_math,
            )
        return p

    def render_abstract_zh(is_first_front: bool):
        begin_odd_section("摘    要", pg_fmt="upperRoman", pg_start=(1 if is_first_front else None))
        add_major_title("摘    要", is_english_bold=False, in_toc=True)
        abs_zh = ir.get("abstract_zh", {})
        paras = abs_zh.get("paragraphs", [])
        if paras:
            for pa in paras:
                add_body_para(pa.get("text", ""), pa.get("segments"))
        else:
            add_body_para("（请补充中文摘要内容，硕士约 500 字，博士约 1000 字。）")
        # Blank line before keywords
        p_blank = doc.add_paragraph()
        set_para_format(p_blank, align="both", line_pt=20.0, line_rule="exact")
        add_run(p_blank, "", size_pt=12.0)
        # Keywords
        p_kw = doc.add_paragraph()
        set_para_format(p_kw, align="both", line_pt=20.0, line_rule="exact", first_line_chars=0)
        add_run(p_kw, "关键词：", east_asia="宋体", latin="Times New Roman", size_pt=12.0, bold=True)
        kws = abs_zh.get("keywords", [])
        kw_str = "；".join(kws) if kws else "关键词1；关键词2；关键词3"
        add_run(p_kw, kw_str, east_asia="宋体", latin="Times New Roman", size_pt=12.0, bold=False)

    def render_abstract_en(is_first_front: bool):
        begin_odd_section("ABSTRACT", pg_fmt="upperRoman", pg_start=(1 if is_first_front else None))
        add_major_title("ABSTRACT", is_english_bold=True, in_toc=True)
        abs_en = ir.get("abstract_en", {})
        paras = abs_en.get("paragraphs", [])
        if paras:
            for pa in paras:
                p = doc.add_paragraph()
                set_para_format(p, align="both", line_pt=20.0, line_rule="exact", first_line_chars=200)
                if pa.get("segments"):
                    add_segments_to_para(p, pa["segments"], east_asia="宋体", latin="Times New Roman",
                                         size_pt=12.0, auto_sup_citations=False, ole_collector=ole_collector,
                                         ir_dir=ir_dir, format_inline_math=format_inline_math)
                else:
                    add_text_with_math_and_cites(p, pa.get("text", ""), east_asia="宋体", latin="Times New Roman",
                                                 size_pt=12.0, auto_sup_citations=False,
                                                 format_inline_math=format_inline_math)
        else:
            p = doc.add_paragraph()
            set_para_format(p, align="both", line_pt=20.0, line_rule="exact", first_line_chars=200)
            add_run(p, "[English Abstract to be completed.]", latin="Times New Roman", size_pt=12.0)
        # Blank line before Keywords
        p_blank = doc.add_paragraph()
        set_para_format(p_blank, align="both", line_pt=20.0, line_rule="exact")
        add_run(p_blank, "", size_pt=12.0)
        # Keywords
        p_kw = doc.add_paragraph()
        set_para_format(p_kw, align="both", line_pt=20.0, line_rule="exact", first_line_chars=0)
        add_run(p_kw, "Keywords: ", east_asia="宋体", latin="Times New Roman", size_pt=12.0, bold=True)
        kws = abs_en.get("keywords", [])
        kw_str = "; ".join(kws) if kws else "Keyword 1; Keyword 2; Keyword 3"
        add_run(p_kw, kw_str, east_asia="宋体", latin="Times New Roman", size_pt=12.0, bold=False)

    # ---- FRONTMATTER SECTIONS (Roman page numbers starting at I) ----
    if language == "en":
        render_abstract_en(is_first_front=True)
        render_abstract_zh(is_first_front=False)
    else:
        render_abstract_zh(is_first_front=True)
        render_abstract_en(is_first_front=False)

    # TOC Section
    toc_title = "Contents" if language == "en" else "目    录"
    begin_odd_section(toc_title, pg_fmt="upperRoman", pg_start=None)
    add_major_title(toc_title, is_english_bold=(language == "en"), in_toc=False)
    insert_toc_field(doc)

    # Optional Symbols Section
    if ir.get("symbols"):
        sym_title = "Nomenclature" if language == "en" else "符号说明"
        begin_odd_section(sym_title, pg_fmt="upperRoman", pg_start=None)
        add_major_title(sym_title, is_english_bold=(language == "en"), in_toc=True)
        for blk in ir["symbols"]:
            if blk["type"] == "table":
                add_three_line_table(doc, blk.get("rows", []), blk.get("caption_zh", ""), "", "0-1",
                                     blk.get("notes", []), bilingual=False, format_inline_math=format_inline_math)
            elif blk["type"] == "paragraph":
                add_body_para(blk.get("text", ""), blk.get("segments"))

    # ---- BODY CHAPTERS (Arabic page numbers starting at 1 on Chapter 1) ----
    for ch_idx, ch in enumerate(ir.get("chapters", [])):
        ch_num = ch.get("number", ch_idx + 1)
        ch_title_clean = ch.get("title", "").strip()
        full_ch_heading = format_chapter_heading(ch_num, ch_title_clean, discipline, language)

        begin_odd_section(full_ch_heading, pg_fmt="decimal", pg_start=(1 if ch_idx == 0 else None))
        add_major_title(full_ch_heading, is_english_bold=(language == "en"), in_toc=True,
                        align=("center" if language == "en" else h1_align))

        sec_idx = 0
        subsec_idx = 0
        fig_seq = 0
        tbl_seq = 0
        eq_seq = 0

        for blk in ch.get("blocks", []):
            btype = blk.get("type")
            if btype == "heading":
                lvl = blk.get("level", 2)
                htitle = blk.get("title", "").strip()
                if lvl == 2:
                    sec_idx += 1
                    subsec_idx = 0
                    htxt = format_sub_heading(2, ch_num, sec_idx, 0, htitle, discipline, language)
                    p = doc.add_paragraph()
                    set_para_format(p, align="both", before_pt=24.0, after_pt=6.0,
                                    outline_lvl=1, keep_with_next=True)
                    add_text_with_math_and_cites(p, htxt, east_asia="黑体", latin="Times New Roman",
                                                 size_pt=14.0, bold=(language == "en"),
                                                 auto_sup_citations=False, format_inline_math=format_inline_math)
                else:
                    if sec_idx == 0:
                        sec_idx = 1
                    subsec_idx += 1
                    htxt = format_sub_heading(3, ch_num, sec_idx, subsec_idx, htitle, discipline, language)
                    p = doc.add_paragraph()
                    set_para_format(p, align="both", before_pt=12.0, after_pt=6.0,
                                    outline_lvl=2, keep_with_next=True)
                    add_text_with_math_and_cites(p, htxt, east_asia="黑体", latin="Times New Roman",
                                                 size_pt=12.0, bold=(language == "en"),
                                                 auto_sup_citations=False, format_inline_math=format_inline_math)
            elif btype == "paragraph":
                add_body_para(blk.get("text", ""), blk.get("segments"))
            elif btype == "quote":
                # Independent quote (humanities): 楷体小四, 段前段后各1行, 左缩进2字符+首行缩进2字符, 固定值20磅
                p = doc.add_paragraph()
                set_para_format(p, align="both", before_lines=100, after_lines=100,
                                line_pt=20.0, line_rule="exact", left_chars=200, first_line_chars=200)
                if blk.get("segments"):
                    add_segments_to_para(p, blk["segments"], east_asia="楷体", latin="Times New Roman",
                                         size_pt=12.0, auto_sup_citations=False, footnote_collector=fn_collector,
                                         ole_collector=ole_collector, ir_dir=ir_dir, format_inline_math=format_inline_math)
                else:
                    add_text_with_math_and_cites(p, blk.get("text", ""), east_asia="楷体", latin="Times New Roman",
                                                 size_pt=12.0, auto_sup_citations=False,
                                                 format_inline_math=format_inline_math)
            elif btype == "figure":
                fig_seq += 1
                add_figure_block(doc, blk, ir_dir, f"{ch_num}-{fig_seq}", bilingual=bilingual,
                                 format_inline_math=format_inline_math)
            elif btype == "table":
                tbl_seq += 1
                add_three_line_table(doc, blk.get("rows", []), blk.get("caption_zh", ""),
                                     blk.get("caption_en", ""), f"{ch_num}-{tbl_seq}",
                                     blk.get("notes", []), bilingual=bilingual,
                                     format_inline_math=format_inline_math)
            elif btype == "equation":
                eq_seq += 1
                eq_num_str = blk.get("number") or f"（{ch_num}-{eq_seq}）"
                add_equation_block(doc, blk, eq_num_str, ole_collector=ole_collector, ir_dir=ir_dir,
                                   format_inline_math=format_inline_math)

    # ---- REFERENCES SECTION (oddPage, continuous Arabic page numbers) ----
    ref_title = "References" if language == "en" else "参考文献"
    begin_odd_section(ref_title, pg_fmt="decimal", pg_start=None)
    add_major_title(ref_title, is_english_bold=(language == "en"), in_toc=True)
    for idx, ref_txt in enumerate(ir.get("references", []), 1):
        clean_ref = re.sub(r"^\[(\d+)\]\s*", r"[\1] ", ref_txt.strip())
        if not re.match(r"^\[\d+\]\s", clean_ref):
            clean_ref = f"[{idx}] {clean_ref}"
        p = doc.add_paragraph()
        set_para_format(p, align="both", before_pt=0.0, after_pt=0.0,
                        line_pt=16.0, line_rule="exact", left_chars=150, hanging_chars=150)
        add_run(p, clean_ref, east_asia="宋体", latin="Times New Roman", size_pt=10.5)

    # ---- OPTIONAL APPENDICES (each starts on oddPage) ----
    for app in ir.get("appendices", []):
        letter = app.get("letter", "A")
        atitle = app.get("title", "").strip()
        full_app = f"Appendix {letter} {atitle}".strip() if language == "en" else f"附录{letter}  {atitle}".strip()
        begin_odd_section(full_app, pg_fmt="decimal", pg_start=None)
        add_major_title(full_app, is_english_bold=(language == "en"), in_toc=True)
        fig_seq = tbl_seq = eq_seq = 0
        for blk in app.get("blocks", []):
            bt = blk.get("type")
            if bt == "paragraph":
                add_body_para(blk.get("text", ""), blk.get("segments"))
            elif bt == "figure":
                fig_seq += 1
                add_figure_block(doc, blk, ir_dir, f"{letter}-{fig_seq}", bilingual=bilingual,
                                 format_inline_math=format_inline_math)
            elif bt == "table":
                tbl_seq += 1
                add_three_line_table(doc, blk.get("rows", []), blk.get("caption_zh", ""),
                                     blk.get("caption_en", ""), f"{letter}-{tbl_seq}",
                                     blk.get("notes", []), bilingual=bilingual,
                                     format_inline_math=format_inline_math)
            elif bt == "equation":
                eq_seq += 1
                add_equation_block(doc, blk, blk.get("number") or f"（{letter}-{eq_seq}）",
                                   ole_collector=ole_collector, ir_dir=ir_dir,
                                   format_inline_math=format_inline_math)

    # ---- ACADEMIC ACHIEVEMENTS SECTION (oddPage) ----
    ach_title = "Academic Achievements During Degree Studies" if language == "en" else "攻读学位期间取得的学术成果"
    begin_odd_section(ach_title, pg_fmt="decimal", pg_start=None)
    add_major_title(ach_title, is_english_bold=(language == "en"), in_toc=True)
    if ir.get("achievements"):
        for blk in ir["achievements"]:
            txt = blk.get("text", "").strip()
            is_sub = bool(re.match(r"^[一二三四五]、", txt))
            p = doc.add_paragraph()
            set_para_format(p, align="both", line_pt=20.0, line_rule="exact",
                            first_line_chars=(0 if is_sub else 0),
                            left_chars=(0 if is_sub else 150),
                            hanging_chars=(0 if is_sub else 150))
            add_run(p, txt, east_asia="宋体", latin="Times New Roman", size_pt=12.0, bold=is_sub)
    else:
        for cat, placeholder in (
            ("一、学术论文", "[1] （待填写：列出所有作者全名，格式同参考文献，末尾标注收录情况与对应章节，如：SCI 一区，对应论文第 3 章）"),
            ("二、国家发明专利", "[1] （待填写：发明人. 专利名称. 专利类型. 授权/公布号, 日期.）"),
            ("三、科研项目", "[1] （待填写：项目类型, 项目名称, 项目编号, 资助单位, 起止时间, 主持/参与.）"),
        ):
            p_cat = doc.add_paragraph()
            set_para_format(p_cat, align="both", before_pt=6.0, after_pt=0.0, line_pt=20.0, line_rule="exact")
            add_run(p_cat, cat, east_asia="宋体", latin="Times New Roman", size_pt=12.0, bold=True)
            p_item = doc.add_paragraph()
            set_para_format(p_item, align="both", line_pt=20.0, line_rule="exact", left_chars=150, hanging_chars=150)
            add_run(p_item, placeholder, east_asia="宋体", latin="Times New Roman", size_pt=12.0)

    # ---- ACKNOWLEDGEMENTS SECTION (oddPage) ----
    ack_title = "Acknowledgments" if language == "en" else "致    谢"
    begin_odd_section(ack_title, pg_fmt="decimal", pg_start=None)
    add_major_title(ack_title, is_english_bold=(language == "en"), in_toc=True)
    if ir.get("acknowledgements"):
        for blk in ir["acknowledgements"]:
            if blk["type"] == "paragraph":
                add_body_para(blk.get("text", ""), blk.get("segments"))
    else:
        add_body_para("（待填写：在此对指导教师、实验室同门、资助项目以及家人朋友致以诚挚感谢。）")

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    inject_footnotes_into_docx(output_path, fn_collector)
    inject_ole_objects_into_docx(output_path, ole_collector)
    print(f"[OK] Built thesis -> {output_path} (sections={len(doc.sections)}, template={os.path.basename(tpl_path)})")

    if spine_output:
        build_spine(meta, spine_output)


def main():
    parser = argparse.ArgumentParser(description="Build TYUST-compliant thesis .docx from IR and metadata.")
    parser.add_argument("--ir", required=True, help="Path to thesis_ir.json")
    parser.add_argument("--meta", required=True, help="Path to metadata.json")
    parser.add_argument("--output", required=True, help="Output formatted thesis .docx")
    parser.add_argument("--spine-output", default=None, help="Optional output path for spine .docx")
    parser.add_argument("--h1-align", choices=["center", "both"], default=None, help="Override H1 alignment")
    parser.add_argument("--bilingual-captions", action="store_true", help="Include English figure/table captions")
    parser.add_argument("--no-inline-math", action="store_true", help="Disable italic/sub/sup formatting on inline math symbols")
    args = parser.parse_args()
    build_thesis(
        args.ir,
        args.meta,
        args.output,
        h1_align_override=args.h1_align,
        bilingual_override=(True if args.bilingual_captions else None),
        spine_output=args.spine_output,
        format_inline_math=not args.no_inline_math,
    )


if __name__ == "__main__":
    main()
