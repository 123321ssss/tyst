#!/usr/bin/env python3
"""
Stage 1: Lossless extraction of an existing Word (.docx) thesis draft into a structured
Intermediate Representation (thesis_ir.json) + extracted media assets + review_hints report.

Preserves:
  - Plain text & inline rich runs (bold, italic, superscript, subscript)
  - Native Word OMML equations (<m:oMath>, <m:oMathPara>) as raw XML
  - Embedded images (extracted as original binary files into media/)
  - Tables (auto-distinguishing normal data tables vs 2-column equation tables)
  - Footnotes (from word/footnotes.xml)
  - Auto-detected headings (from Word styles/outlineLvl OR regex patterns), captions,
    abstracts, keywords, chapters, references, appendices, achievements, acknowledgements.
"""

import argparse
import json
import os
import re
import zipfile
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph

O_NS = "urn:schemas-microsoft-com:office:office"
V_NS = "urn:schemas-microsoft-com:vml"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CN_NUM_MAP = {
    "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
    "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
    "十一": 11, "十二": 12, "十三": 13, "十四": 14, "十五": 15,
}

RE_EQ_NUM = re.compile(r"^\s*[（(]\s*([A-Za-z]?\d+[-–—\.]\d+)\s*[）)]\s*$")
RE_CHAPTER_CN = re.compile(r"^第\s*(\d+|[一二三四五六七八九十]+)\s*章\s*(.*)$")
RE_CHAPTER_EN = re.compile(r"^Chapter\s+(\d+)\s*[:：]?\s*(.+)$", re.IGNORECASE)
RE_CHAPTER_NUM = re.compile(r"^(\d{1,2})\s{1,6}([^\d\s].{1,45})$")
RE_CHAPTER_HUM = re.compile(r"^([一二三四五六七八九十]+)、\s*(.{1,40})$")
RE_SEC_NUM = re.compile(r"^(\d+\.\d+)\s+(.+)$")
RE_SUBSEC_NUM = re.compile(r"^(\d+\.\d+\.\d+)\s+(.+)$")
RE_SEC_HUM = re.compile(r"^（([一二三四五六七八九十]+)）\s*(.{1,45})$")
RE_SUBSEC_HUM = re.compile(r"^(\d+)\.\s*([^\d\s].{1,45})$")
RE_FIG_CAP = re.compile(r"^图\s*([A-Za-z]?\d+[-–—\.]\d+)\s*(.*)$")
RE_FIG_CAP_EN = re.compile(r"^Fig\.?\s*([A-Za-z]?\d+[-–—\.]\d+)\s*(.*)$", re.IGNORECASE)
RE_TBL_CAP = re.compile(r"^表\s*([A-Za-z]?\d+[-–—\.]\d+)\s*(.*)$")
RE_TBL_CAP_EN = re.compile(r"^Table\s*([A-Za-z]?\d+[-–—\.]\d+)\s*(.*)$", re.IGNORECASE)
RE_SUBFIG_CAP = re.compile(r"^[（(]a[）)].+")
RE_NOTE = re.compile(r"^(资料来源|图片来源|数据来源|注)\s*[：:]")
RE_ABSTRACT_ZH = re.compile(r"^摘\s*要$")
RE_ABSTRACT_EN = re.compile(r"^ABSTRACT$", re.IGNORECASE)
RE_KEYWORDS_ZH = re.compile(r"^关键词\s*[：:]\s*(.+)$")
RE_KEYWORDS_EN = re.compile(r"^Key\s*words\s*[：:]\s*(.+)$", re.IGNORECASE)
RE_TOC_TITLE = re.compile(r"^(目\s*录|Contents)$", re.IGNORECASE)
RE_SYMBOLS_TITLE = re.compile(r"^(符号说明|主要符号表|缩略词及符号说明)$")
RE_REFS_TITLE = re.compile(r"^(参\s*考\s*文\s*献|References)$", re.IGNORECASE)
RE_APPENDIX_TITLE = re.compile(r"^(附录\s*([A-Z])|Appendix\s*([A-Z]))\s*(.*)$", re.IGNORECASE)
RE_ACHIEVE_TITLE = re.compile(r"^(攻读.*学位期间.*成果|攻读.*期间.*学术成果|Academic Achievements.*)$", re.IGNORECASE)
RE_ACK_TITLE = re.compile(r"^(致\s*谢|Acknowledgments?)$", re.IGNORECASE)
RE_REF_ITEM = re.compile(r"^\[\d+\]")


def parse_cn_or_int(token: str) -> int:
    token = token.strip()
    if token.isdigit():
        return int(token)
    return CN_NUM_MAP.get(token, 1)


def normalize_heading_title(raw_title: str) -> str:
    """Clean excessive internal spaces in 2-char Chinese headings like '绪  论' -> '绪论'."""
    t = raw_title.strip()
    # Collapse spaces between two single Chinese chars (e.g. '绪  论' -> '绪论')
    if re.fullmatch(r"[\u4e00-\u9fff]\s+[\u4e00-\u9fff]", t):
        t = re.sub(r"\s+", "", t)
    return t


def extract_footnotes_map(docx_path: str) -> dict:
    """Extract {footnote_id: plain_text} from word/footnotes.xml if present."""
    fn_map = {}
    with zipfile.ZipFile(docx_path, "r") as z:
        if "word/footnotes.xml" not in z.namelist():
            return fn_map
        root = etree.fromstring(z.read("word/footnotes.xml"))
        for fn in root.findall(qn("w:footnote")):
            fn_type = fn.get(qn("w:type"))
            if fn_type in ("separator", "continuationSeparator"):
                continue
            fn_id = fn.get(qn("w:id"))
            texts = [t.text for t in fn.findall(".//" + qn("w:t")) if t.text]
            txt = "".join(texts).strip()
            if fn_id is not None and txt:
                fn_map[fn_id] = txt
    return fn_map


class MediaExtractor:
    def __init__(self, doc: Document, out_dir: str):
        self.doc = doc
        self.media_dir = os.path.join(out_dir, "media")
        self.emb_dir = os.path.join(out_dir, "embeddings")
        os.makedirs(self.media_dir, exist_ok=True)
        os.makedirs(self.emb_dir, exist_ok=True)
        self.rid_to_relpath = {}
        self.ole_rid_to_relpath = {}
        self.counter = 0
        self.ole_counter = 0

    def extract_from_element(self, el) -> list:
        """Find all blip/imagedata embeds inside an XML element, extract binary to media/, return metadata list."""
        images = []
        # DrawingML blips
        for blip in el.findall(f".//{{{A_NS}}}blip"):
            rid = blip.get(f"{{{R_NS}}}embed")
            if not rid:
                continue
            rel_path = self._save_rid(rid)
            if rel_path:
                images.append({"path": rel_path})
        # VML imagedata fallback
        for imd in el.findall(f".//{{{V_NS}}}imagedata"):
            rid = imd.get(f"{{{R_NS}}}id")
            if not rid:
                continue
            rel_path = self._save_rid(rid)
            if rel_path:
                images.append({"path": rel_path})
        return images

    def extract_ole_object(self, obj_el, position: str = None):
        """Extract a MathType/Word <w:object> containing <o:OLEObject> + preview <v:imagedata>."""
        ole_el = obj_el.find(f"{{{O_NS}}}OLEObject")
        if ole_el is None:
            return None
        ole_rid = ole_el.get(f"{{{R_NS}}}id")
        if not ole_rid:
            return None
        ole_rel_path = self._save_ole_rid(ole_rid)
        if not ole_rel_path:
            return None

        imd = obj_el.find(f".//{{{V_NS}}}imagedata")
        img_old_rid = imd.get(f"{{{R_NS}}}id") if imd is not None else None
        img_rel_path = self._save_rid(img_old_rid, prefix="ole_img_") if img_old_rid else None

        return {
            "type": "ole_object",
            "xml": etree.tostring(obj_el, encoding="unicode"),
            "ole_path": ole_rel_path,
            "ole_old_rid": ole_rid,
            "image_path": img_rel_path,
            "image_old_rid": img_old_rid,
            "prog_id": ole_el.get("ProgID", ""),
            "position": position,
        }

    def _save_ole_rid(self, rid: str):
        if rid in self.ole_rid_to_relpath:
            return self.ole_rid_to_relpath[rid]
        if rid not in self.doc.part.rels:
            return None
        rel = self.doc.part.rels[rid]
        try:
            blob = rel.target_part.blob
            ext = os.path.splitext(rel.target_ref)[1] or ".bin"
        except Exception:
            return None
        self.ole_counter += 1
        fname = f"ole_{self.ole_counter:03d}{ext}"
        full_path = os.path.join(self.emb_dir, fname)
        with open(full_path, "wb") as f:
            f.write(blob)
        rel_path = f"embeddings/{fname}"
        self.ole_rid_to_relpath[rid] = rel_path
        return rel_path

    def _save_rid(self, rid: str, prefix: str = "img_"):
        if rid in self.rid_to_relpath:
            return self.rid_to_relpath[rid]
        if rid not in self.doc.part.rels:
            return None
        rel = self.doc.part.rels[rid]
        try:
            blob = rel.target_part.blob
            ext = os.path.splitext(rel.target_ref)[1] or ".png"
        except Exception:
            return None
        self.counter += 1
        fname = f"{prefix}{self.counter:03d}{ext}"
        full_path = os.path.join(self.media_dir, fname)
        with open(full_path, "wb") as f:
            f.write(blob)
        rel_path = f"media/{fname}"
        self.rid_to_relpath[rid] = rel_path
        return rel_path


def extract_paragraph_segments(p_el, media_ext: MediaExtractor) -> tuple:
    """
    Serialize paragraph children into lossless segments:
      - text runs (with bold/italic/sup/sub)
      - omml (<m:oMath> / <m:oMathPara>) raw XML
      - footnote_ref (id)
      - image (extracted media path)
    Returns (plain_text, segments, images).
    """
    segments = []
    images = []
    plain_parts = []

    for child in p_el:
        tag = child.tag
        if tag == qn("w:pPr"):
            continue
        elif tag in (f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara"):
            xml_str = etree.tostring(child, encoding="unicode")
            segments.append({"type": "omml", "xml": xml_str})
            math_txt = "".join(t.text for t in child.findall(f".//{{{M_NS}}}t") if t.text)
            if math_txt:
                plain_parts.append(math_txt)
        elif tag == qn("w:hyperlink"):
            # Recurse into hyperlink runs
            sub_txt, sub_segs, sub_imgs = extract_paragraph_segments(child, media_ext)
            segments.extend(sub_segs)
            images.extend(sub_imgs)
            plain_parts.append(sub_txt)
        elif tag == qn("w:r"):
            rPr = child.find(qn("w:rPr"))
            pos_el = rPr.find(qn("w:position")) if rPr is not None else None
            r_position = pos_el.get(qn("w:val")) if pos_el is not None else None

            # Prioritize embedded OLE object (e.g. MathType Equation.3) before drawing/pict
            obj_el = child.find(qn("w:object"))
            handled_ole = False
            if obj_el is not None:
                ole_seg = media_ext.extract_ole_object(obj_el, position=r_position)
                if ole_seg is not None:
                    segments.append(ole_seg)
                    handled_ole = True

            # Check embedded drawing/pict (if not already handled as an OLE object)
            if not handled_ole and (child.find(qn("w:drawing")) is not None or child.find(qn("w:pict")) is not None):
                found_imgs = media_ext.extract_from_element(child)
                for im in found_imgs:
                    images.append(im)
                    segments.append({"type": "image", "path": im["path"]})
            # Check footnote reference
            fn_ref = child.find(qn("w:footnoteReference"))
            if fn_ref is not None:
                fn_id = fn_ref.get(qn("w:id"))
                if fn_id is not None:
                    segments.append({"type": "footnote_ref", "id": fn_id})
            # Run formatting
            bold = False
            italic = False
            sup = False
            sub = False
            if rPr is not None:
                b_el = rPr.find(qn("w:b"))
                if b_el is not None and b_el.get(qn("w:val")) not in ("0", "false"):
                    bold = True
                i_el = rPr.find(qn("w:i"))
                if i_el is not None and i_el.get(qn("w:val")) not in ("0", "false"):
                    italic = True
                va = rPr.find(qn("w:vertAlign"))
                if va is not None:
                    v = va.get(qn("w:val"))
                    sup = (v == "superscript")
                    sub = (v == "subscript")
            # Extract run text + breaks + tabs
            r_text_parts = []
            for rch in child:
                if rch.tag == qn("w:t") and rch.text:
                    r_text_parts.append(rch.text)
                elif rch.tag == qn("w:tab"):
                    r_text_parts.append("\t")
                elif rch.tag == qn("w:br"):
                    if rch.get(qn("w:type")) != "page":
                        r_text_parts.append("\n")
            r_text = "".join(r_text_parts)
            if r_text:
                plain_parts.append(r_text)
                # Merge adjacent text segments with identical formatting
                if (
                    segments
                    and segments[-1]["type"] == "text"
                    and segments[-1]["bold"] == bold
                    and segments[-1]["italic"] == italic
                    and segments[-1]["sup"] == sup
                    and segments[-1]["sub"] == sub
                ):
                    segments[-1]["text"] += r_text
                else:
                    segments.append({
                        "type": "text",
                        "text": r_text,
                        "bold": bold,
                        "italic": italic,
                        "sup": sup,
                        "sub": sub,
                    })

    return "".join(plain_parts), segments, images


def get_outline_level(p_el, style_name: str) -> int:
    """Return 0/1/2 for Heading 1/2/3 or explicit outlineLvl, else -1."""
    sn = (style_name or "").lower()
    if sn in ("heading 1", "标题 1", "heading1"):
        return 0
    if sn in ("heading 2", "标题 2", "heading2"):
        return 1
    if sn in ("heading 3", "标题 3", "heading3"):
        return 2
    pPr = p_el.find(qn("w:pPr"))
    if pPr is not None:
        ol = pPr.find(qn("w:outlineLvl"))
        if ol is not None and ol.get(qn("w:val")) is not None:
            v = int(ol.get(qn("w:val")))
            if v in (0, 1, 2):
                return v
    return -1


def is_independent_quote(p_el) -> bool:
    """Detect humanities-style independent quotation (楷体 + left indent >= 200 chars or >= 350 dxa)."""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return False
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return False
    left = int(ind.get(qn("w:left") or "0") or 0)
    left_chars = int(ind.get(qn("w:leftChars") or "0") or 0)
    if left < 300 and left_chars < 150:
        return False
    # Check if run uses 楷体
    for rFonts in p_el.findall(".//" + qn("w:rFonts")):
        ea = rFonts.get(qn("w:eastAsia")) or ""
        if "楷" in ea or "Kai" in ea:
            return True
    return False


def check_equation_table(tbl: Table, media_ext: MediaExtractor):
    """
    Detect if a 2-column (formula | (2-1)) or 3-column (empty | formula | (2-1)) table
    is an equation layout table. If so, return list of equation block dicts; else return None.
    """
    ncols = len(tbl.columns)
    if ncols not in (2, 3) or len(tbl.rows) == 0:
        return None
    eqs = []
    for row in tbl.rows:
        if ncols == 3 and row.cells[0].text.strip():
            return None
        right_txt = row.cells[ncols - 1].text.strip()
        m = RE_EQ_NUM.match(right_txt)
        if not m:
            return None
        eq_num = m.group(1).replace("–", "-").replace("—", "-")
        # Extract rich segments from formula cell paragraphs
        left_cell = row.cells[ncols - 2]
        cell_segs = []
        cell_texts = []
        for pi, p in enumerate(left_cell.paragraphs):
            ptxt, psegs, _ = extract_paragraph_segments(p._p, media_ext)
            if pi > 0 and ptxt.strip():
                cell_segs.append({"type": "text", "text": "  ", "bold": False, "italic": False, "sup": False, "sub": False})
                cell_texts.append("  ")
            cell_segs.extend(psegs)
            cell_texts.append(ptxt.strip())
        eqs.append({
            "type": "equation",
            "number": f"（{eq_num}）",
            "text": "".join(cell_texts).strip(),
            "segments": cell_segs,
        })
    return eqs


def extract_table_data(tbl: Table, media_ext: MediaExtractor) -> dict:
    """Extract a normal Word table into rows of cell text & rich segments."""
    rows_out = []
    for row in tbl.rows:
        r_cells = []
        for cell in row.cells:
            # Join multiple paragraphs inside a table cell with '\n'
            p_texts = []
            for p in cell.paragraphs:
                t = p.text.strip()
                if t:
                    p_texts.append(t)
            r_cells.append("\n".join(p_texts))
        rows_out.append(r_cells)
    return {
        "type": "table",
        "caption_zh": "",
        "caption_en": "",
        "notes": [],
        "rows": rows_out,
    }


def extract_thesis(input_path: str, out_dir: str, discipline: str = "science", split_colon_headings: bool = True):
    os.makedirs(out_dir, exist_ok=True)
    doc = Document(input_path)
    media_ext = MediaExtractor(doc, out_dir)
    footnotes_map = extract_footnotes_map(input_path)

    ir = {
        "source_file": os.path.abspath(input_path),
        "detected_metadata": {},
        "abstract_zh": {"paragraphs": [], "keywords": []},
        "abstract_en": {"paragraphs": [], "keywords": []},
        "symbols": [],
        "chapters": [],
        "references": [],
        "appendices": [],
        "achievements": [],
        "acknowledgements": [],
        "footnotes": footnotes_map,
    }
    review_hints = []

    state = "frontmatter"
    current_chapter = None
    current_appendix = None
    pending_table_caption = None  # {"zh": ..., "en": ...}
    last_block_ref = None  # reference to last figure/table block for attaching captions/notes
    frontmatter_texts = []
    chapter_counter = 0

    def append_block(blk):
        nonlocal last_block_ref
        if state == "chapters":
            if current_chapter is None:
                # Create implicit Chapter 1 if body text appears before any chapter heading
                nonlocal chapter_counter
                chapter_counter += 1
                ch = {"number": chapter_counter, "title": "绪论", "blocks": []}
                ir["chapters"].append(ch)
                current_chapter_set(ch)
            current_chapter["blocks"].append(blk)
            last_block_ref = blk
        elif state == "appendices" and current_appendix is not None:
            current_appendix["blocks"].append(blk)
            last_block_ref = blk
        elif state == "symbols":
            ir["symbols"].append(blk)
            last_block_ref = blk
        elif state == "achievements":
            ir["achievements"].append(blk)
            last_block_ref = blk
        elif state == "acknowledgements":
            ir["acknowledgements"].append(blk)
            last_block_ref = blk

    def current_chapter_set(ch):
        nonlocal current_chapter
        current_chapter = ch

    def start_chapter(num: int, title: str):
        nonlocal state, current_chapter, chapter_counter, pending_table_caption, last_block_ref
        state = "chapters"
        chapter_counter = num if num > 0 else (chapter_counter + 1)
        ch = {
            "number": chapter_counter,
            "title": normalize_heading_title(title),
            "blocks": [],
        }
        ir["chapters"].append(ch)
        current_chapter = ch
        pending_table_caption = None
        last_block_ref = None

    body_children = list(doc.element.body.iterchildren())

    for idx, child in enumerate(body_children):
        # -------- TABLE ELEMENT --------
        if isinstance(child, CT_Tbl):
            tbl = Table(child, doc)
            if state == "frontmatter":
                # Parse key-value metadata table from draft cover
                for row in tbl.rows:
                    if len(row.cells) >= 2:
                        k = row.cells[0].text.strip().rstrip("：:")
                        v = row.cells[1].text.strip()
                        if k and v:
                            ir["detected_metadata"][k] = v
                continue

            if state == "toc":
                continue

            # Check if this table is actually an equation layout table
            eq_blocks = check_equation_table(tbl, media_ext)
            if eq_blocks is not None:
                for eq in eq_blocks:
                    append_block(eq)
                continue

            tbl_blk = extract_table_data(tbl, media_ext)
            if pending_table_caption is not None:
                tbl_blk["caption_zh"] = pending_table_caption.get("zh", "")
                tbl_blk["caption_en"] = pending_table_caption.get("en", "")
                pending_table_caption = None
            append_block(tbl_blk)
            continue

        # -------- PARAGRAPH ELEMENT --------
        if not isinstance(child, CT_P):
            continue

        p = Paragraph(child, doc)
        st_name = p.style.name if p.style else ""
        outline = get_outline_level(child, st_name)
        raw_txt, segments, images = extract_paragraph_segments(child, media_ext)
        txt = raw_txt.strip()

        has_ole = any(s.get("type") == "ole_object" for s in segments)
        if not txt and not images and not has_ole:
            continue

        # 0. Guard inside TOC state so TOC entries like '参考文献\t15' / '附录A  标题\t17' never hijack state
        if state == "toc":
            is_toc_para = (
                st_name.lower().startswith("toc")
                or bool(re.search(r"\t\s*([IVXLCDM]+|\d+)\s*$", txt))
                or child.find(".//" + qn("w:fldChar")) is not None
                or child.find(".//" + qn("w:instrText")) is not None
            )
            if is_toc_para:
                continue

        # 1. Detect top-level structural markers (work from any non-TOC state)
        if RE_ABSTRACT_ZH.match(txt):
            state = "abstract_zh"
            last_block_ref = None
            continue
        if RE_ABSTRACT_EN.match(txt):
            state = "abstract_en"
            last_block_ref = None
            continue
        if RE_TOC_TITLE.match(txt):
            state = "toc"
            last_block_ref = None
            continue
        if RE_SYMBOLS_TITLE.match(txt) and len(txt) <= 20:
            state = "symbols"
            last_block_ref = None
            continue
        if RE_REFS_TITLE.match(txt):
            state = "references"
            last_block_ref = None
            continue
        if RE_ACHIEVE_TITLE.match(txt) and len(txt) <= 45:
            state = "achievements"
            last_block_ref = None
            continue
        if RE_ACK_TITLE.match(txt):
            state = "acknowledgements"
            last_block_ref = None
            continue
        m_app = RE_APPENDIX_TITLE.match(txt)
        if m_app and len(txt) <= 50 and "：" not in txt:
            state = "appendices"
            app_letter = (m_app.group(2) or m_app.group(3) or "A").upper()
            app_title = (m_app.group(4) or "").strip()
            current_appendix = {"letter": app_letter, "title": app_title, "blocks": []}
            ir["appendices"].append(current_appendix)
            last_block_ref = None
            continue

        # Skip TOC entries if old thesis already had a TOC
        if state == "toc":
            # Check if a chapter heading starts here to exit TOC
            m_ch_cn = RE_CHAPTER_CN.match(txt)
            m_ch_en = RE_CHAPTER_EN.match(txt)
            if not (m_ch_cn or m_ch_en or outline == 0):
                continue

        # 2. Frontmatter handling
        if state == "frontmatter":
            frontmatter_texts.append(txt)
            continue

        # 3. Abstract ZH / EN handling
        if state == "abstract_zh":
            m_kw = RE_KEYWORDS_ZH.match(txt)
            if m_kw:
                kws = [k.strip() for k in re.split(r"[；;，,、]+", m_kw.group(1)) if k.strip()]
                ir["abstract_zh"]["keywords"] = kws
                continue
            # Check if Chapter 1 starts immediately after Chinese abstract (no EN abstract in draft!)
            m_ch_cn = RE_CHAPTER_CN.match(txt)
            m_ch_en = RE_CHAPTER_EN.match(txt)
            if (m_ch_cn and "：" not in txt and len(txt) <= 50) or (m_ch_en and len(txt) <= 70) or outline == 0:
                pass  # fall through to chapter detection below
            else:
                ir["abstract_zh"]["paragraphs"].append({"text": txt, "segments": segments})
                continue

        if state == "abstract_en":
            m_kw = RE_KEYWORDS_EN.match(txt)
            if m_kw:
                kws = [k.strip() for k in re.split(r"[;；]+", m_kw.group(1)) if k.strip()]
                ir["abstract_en"]["keywords"] = kws
                continue
            m_ch_cn = RE_CHAPTER_CN.match(txt)
            m_ch_en = RE_CHAPTER_EN.match(txt)
            if (m_ch_cn and "：" not in txt and len(txt) <= 50) or (m_ch_en and len(txt) <= 70) or outline == 0:
                pass  # fall through to chapter detection below
            else:
                ir["abstract_en"]["paragraphs"].append({"text": txt, "segments": segments})
                continue

        # 4. References handling
        if state == "references":
            ref_txt = txt
            # Normalize leading [n] spacing
            ref_txt = re.sub(r"^\[(\d+)\]\s*", r"[\1] ", ref_txt)
            ir["references"].append(ref_txt)
            continue

        # 5. Chapter heading detection (from abstract/toc/chapters)
        m_ch_cn = RE_CHAPTER_CN.match(txt)
        if m_ch_cn and len(txt) <= 60 and not ("：" in txt and len(txt) > 35):
            ch_num = parse_cn_or_int(m_ch_cn.group(1))
            ch_title = m_ch_cn.group(2).strip()
            start_chapter(ch_num, ch_title)
            continue

        m_ch_en = RE_CHAPTER_EN.match(txt)
        if m_ch_en and len(txt) <= 80:
            start_chapter(int(m_ch_en.group(1)), m_ch_en.group(2).strip())
            continue

        if outline == 0 and state in ("abstract_zh", "abstract_en", "toc", "chapters"):
            # Strip existing leading chapter number if present
            m_n = RE_CHAPTER_NUM.match(txt)
            m_h = RE_CHAPTER_HUM.match(txt)
            if m_n:
                start_chapter(int(m_n.group(1)), m_n.group(2))
            elif m_h:
                start_chapter(parse_cn_or_int(m_h.group(1)), m_h.group(2))
            else:
                start_chapter(chapter_counter + 1, txt)
            continue

        # Humanities explicit chapter heading without outlineLvl (e.g. '一、绪论')
        if discipline == "humanities" and state in ("abstract_zh", "abstract_en", "toc", "chapters"):
            m_h = RE_CHAPTER_HUM.match(txt)
            if m_h and "。" not in txt and "：" not in txt and len(txt) <= 35:
                expected = chapter_counter + 1
                num = parse_cn_or_int(m_h.group(1))
                if num == expected:
                    start_chapter(num, m_h.group(2))
                    continue

        # Science explicit numeric chapter heading without outlineLvl (e.g. '1  绪论')
        if discipline == "science" and state in ("abstract_zh", "abstract_en", "toc", "chapters"):
            m_n = RE_CHAPTER_NUM.match(txt)
            if m_n and "。" not in txt and "：" not in txt and len(txt) <= 35:
                num = int(m_n.group(1))
                if num == chapter_counter + 1:
                    start_chapter(num, m_n.group(2))
                    continue

        # 6. Inside chapters / appendices / achievements / acknowledgements / symbols
        # 6a. Standalone image paragraph
        if images and not txt:
            fig_blk = {
                "type": "figure",
                "images": [im["path"] for im in images],
                "caption_zh": "",
                "caption_en": "",
                "subcaption_zh": "",
                "subcaption_en": "",
                "notes": [],
            }
            append_block(fig_blk)
            continue

        # 6b. Figure caption immediately following a figure
        m_fig = RE_FIG_CAP.match(txt)
        if m_fig and len(txt) <= 90 and "。" not in txt:
            cap_text = m_fig.group(2).strip()
            if last_block_ref and last_block_ref.get("type") == "figure" and not last_block_ref.get("caption_zh"):
                last_block_ref["caption_zh"] = cap_text
                last_block_ref["orig_num"] = m_fig.group(1)
            else:
                append_block({
                    "type": "figure",
                    "images": [im["path"] for im in images],
                    "caption_zh": cap_text,
                    "caption_en": "",
                    "subcaption_zh": "",
                    "subcaption_en": "",
                    "notes": [],
                    "orig_num": m_fig.group(1),
                })
            continue

        m_fig_en = RE_FIG_CAP_EN.match(txt)
        if m_fig_en and last_block_ref and last_block_ref.get("type") == "figure":
            last_block_ref["caption_en"] = m_fig_en.group(2).strip()
            continue

        if RE_SUBFIG_CAP.match(txt) and last_block_ref and last_block_ref.get("type") == "figure" and len(txt) <= 100:
            if re.search(r"[\u4e00-\u9fff]", txt):
                last_block_ref["subcaption_zh"] = txt
            else:
                last_block_ref["subcaption_en"] = txt
            continue

        # 6c. Table caption (precedes the table)
        m_tbl = RE_TBL_CAP.match(txt)
        if m_tbl and len(txt) <= 90 and "。" not in txt:
            pending_table_caption = {"zh": m_tbl.group(2).strip(), "en": "", "orig_num": m_tbl.group(1)}
            continue

        m_tbl_en = RE_TBL_CAP_EN.match(txt)
        if m_tbl_en and pending_table_caption is not None:
            pending_table_caption["en"] = m_tbl_en.group(2).strip()
            continue

        # 6d. Table/Figure source or note
        if RE_NOTE.match(txt) and last_block_ref and last_block_ref.get("type") in ("table", "figure"):
            last_block_ref["notes"].append(txt)
            continue

        # 6e. Section (H2) and Subsection (H3) detection
        if state in ("chapters", "appendices"):
            # H3 first (e.g. 1.2.1  标题)
            m_h3 = RE_SUBSEC_NUM.match(txt)
            if (outline == 2 or (m_h3 and "。" not in txt and len(txt) <= 65)):
                if m_h3:
                    append_block({"type": "heading", "level": 3, "orig_num": m_h3.group(1), "title": m_h3.group(2).strip()})
                else:
                    m_h3_hum = RE_SUBSEC_HUM.match(txt)
                    t = m_h3_hum.group(2).strip() if m_h3_hum else txt
                    append_block({"type": "heading", "level": 3, "orig_num": m_h3_hum.group(1) if m_h3_hum else "", "title": t})
                continue

            # H2 (e.g. 1.1  标题 OR outline-style '3.1 标题：正文...')
            m_h2 = RE_SEC_NUM.match(txt)
            if outline == 1:
                t = m_h2.group(2).strip() if m_h2 else txt
                append_block({"type": "heading", "level": 2, "orig_num": m_h2.group(1) if m_h2 else "", "title": t})
                continue

            if m_h2:
                sec_num_str = m_h2.group(1)
                rest = m_h2.group(2).strip()
                ch_prefix = int(sec_num_str.split(".")[0])
                if current_chapter and ch_prefix == current_chapter["number"]:
                    if "：" not in rest and "。" not in rest and len(rest) <= 60:
                        append_block({"type": "heading", "level": 2, "orig_num": sec_num_str, "title": rest})
                        continue
                    elif split_colon_headings and "：" in rest:
                        head_part, body_part = rest.split("：", 1)
                        if 2 <= len(head_part) <= 45 and len(body_part.strip()) >= 10:
                            append_block({"type": "heading", "level": 2, "orig_num": sec_num_str, "title": head_part.strip()})
                            append_block({
                                "type": "paragraph",
                                "text": body_part.strip(),
                                "segments": [{"type": "text", "text": body_part.strip(), "bold": False, "italic": False, "sup": False, "sub": False}],
                            })
                            review_hints.append(
                                f"已将行内冒号节标题拆分为二级标题+正文: {sec_num_str} {head_part.strip()}"
                            )
                            continue

            # Humanities H2/H3 without outlineLvl
            if discipline == "humanities":
                m_h2_hum = RE_SEC_HUM.match(txt)
                if m_h2_hum and "。" not in txt and "：" not in txt and len(txt) <= 40:
                    append_block({"type": "heading", "level": 2, "orig_num": m_h2_hum.group(1), "title": m_h2_hum.group(2).strip()})
                    continue
                m_h3_hum = RE_SUBSEC_HUM.match(txt)
                if m_h3_hum and "。" not in txt and "：" not in txt and len(txt) <= 40:
                    append_block({"type": "heading", "level": 3, "orig_num": m_h3_hum.group(1), "title": m_h3_hum.group(2).strip()})
                    continue

        # 6f. Standalone OMML or MathType OLE equation paragraph
        has_omml = any(s["type"] == "omml" for s in segments)
        has_eq_ole = any(
            s.get("type") == "ole_object" and ("Equation" in (s.get("prog_id") or "") or not s.get("prog_id"))
            for s in segments
        )
        if (has_omml or has_eq_ole) and state in ("chapters", "appendices"):
            # Check if trailing equation number exists in text
            m_tail = re.search(r"[（(]\s*([A-Za-z]?\d+[-–—\.]\d+)\s*[）)]\s*$", txt)
            non_math_len = sum(len(s["text"].strip()) for s in segments if s["type"] == "text")
            if m_tail or non_math_len <= 15:
                eq_num = f"（{m_tail.group(1)}）" if m_tail else ""
                # Strip trailing eq_num and trailing whitespace runs so formula centers cleanly
                cleaned_segs = [dict(s) for s in segments]
                if m_tail:
                    # Remove eq_num from the last non-empty text segment
                    for si in range(len(cleaned_segs) - 1, -1, -1):
                        if cleaned_segs[si].get("type") == "text":
                            new_t = re.sub(r"\s*[（(]\s*[A-Za-z]?\d+[-–—\.]\d+\s*[）)]\s*$", "", cleaned_segs[si]["text"])
                            cleaned_segs[si]["text"] = new_t
                            if new_t.strip():
                                break
                cleaned_segs = [
                    s for s in cleaned_segs
                    if not (s.get("type") == "text" and not s.get("text", "").strip())
                ]
                append_block({
                    "type": "equation",
                    "number": eq_num,
                    "text": txt,
                    "segments": cleaned_segs,
                })
                continue

        # 6g. Independent quotation (humanities)
        if state in ("chapters", "appendices") and is_independent_quote(child):
            append_block({"type": "quote", "text": txt, "segments": segments})
            continue

        # 6h. Normal paragraph
        # Flag divider notes like '后续章节研究内容规划与详细大纲（第 3–6 章）'
        if re.match(r"^后续章节.*大纲", txt):
            review_hints.append(f"检测到草稿过渡提示行（已保留在正文，正式定稿可删）: {txt}")
        append_block({"type": "paragraph", "text": txt, "segments": segments})

    # Post-process detected frontmatter lines for title_zh / title_en
    if frontmatter_texts:
        ir["detected_metadata"]["frontmatter_lines"] = frontmatter_texts
        for line in frontmatter_texts:
            clean_line = " ".join(line.split())
            if re.search(r"学位论文", clean_line) and len(clean_line) <= 30:
                continue
            if not ir["detected_metadata"].get("title_zh") and re.search(r"[\u4e00-\u9fff]", clean_line):
                ir["detected_metadata"]["title_zh"] = re.sub(r"\s+", "", clean_line)
            elif not ir["detected_metadata"].get("title_en") and re.search(r"[A-Za-z]{4,}", clean_line):
                # Fix accidental missing space like Laser3D -> Laser 3D
                ir["detected_metadata"]["title_en"] = clean_line

    # Compute summary stats & review hints
    total_figs = sum(1 for ch in ir["chapters"] for b in ch["blocks"] if b["type"] == "figure")
    total_tbls = sum(1 for ch in ir["chapters"] for b in ch["blocks"] if b["type"] == "table")
    total_eqs = sum(1 for ch in ir["chapters"] for b in ch["blocks"] if b["type"] == "equation")
    body_chars = sum(
        len(b.get("text", ""))
        for ch in ir["chapters"]
        for b in ch["blocks"]
        if b["type"] in ("paragraph", "quote")
    )
    abs_zh_chars = sum(len(p["text"]) for p in ir["abstract_zh"]["paragraphs"])

    if not ir["abstract_en"]["paragraphs"]:
        review_hints.append("旧论文缺少英文摘要（ABSTRACT）与英文关键词（Keywords），建议补全。")
    if len(ir["references"]) < 30:
        review_hints.append(f"当前参考文献仅 {len(ir['references'])} 篇（硕士规范建议 ≥30 篇，博士 ≥80 篇）。")
    if not ir["achievements"]:
        review_hints.append("旧论文缺少『攻读学位期间取得的学术成果』章节（构建时将自动生成标准占位节）。")
    if not ir["acknowledgements"]:
        review_hints.append("旧论文缺少『致谢』章节（构建时将自动生成标准占位节）。")
    for ch in ir["chapters"]:
        if "（大纲）" in ch["title"] or "（撰写要点）" in ch["title"]:
            review_hints.append(f"第 {ch['number']} 章标题含草稿后缀: '{ch['title']}'")
    for ref in ir["references"]:
        if "[注:" in ref or "[待补充]" in ref:
            review_hints.append(f"参考文献含待核对/待补充标注: {ref[:65]}")

    plain_eqs = 0
    for ch in ir["chapters"] + ir["appendices"]:
        for b in ch["blocks"]:
            if b.get("type") == "equation":
                segs = b.get("segments", [])
                has_rich = any(s.get("type") in ("omml", "ole_object") for s in segs) or bool(b.get("latex")) or bool(b.get("latex_lines"))
                if not has_rich:
                    plain_eqs += 1
    if plain_eqs > 0:
        review_hints.append(
            f"发现 {plain_eqs} 个纯文本公式，建议在 Step 2 补充 latex / latex_lines 字段以生成 Word 原生 OMML 公式"
        )

    ir_path = os.path.join(out_dir, "thesis_ir.json")
    with open(ir_path, "w", encoding="utf-8") as f:
        json.dump(ir, f, ensure_ascii=False, indent=2)

    report = {
        "source_file": os.path.abspath(input_path),
        "ir_path": ir_path,
        "stats": {
            "chapters": len(ir["chapters"]),
            "chapter_titles": [f"{c['number']} {c['title']}" for c in ir["chapters"]],
            "abstract_zh_chars": abs_zh_chars,
            "abstract_en_paras": len(ir["abstract_en"]["paragraphs"]),
            "body_chars": body_chars,
            "figures": total_figs,
            "tables": total_tbls,
            "equations": total_eqs,
            "references": len(ir["references"]),
            "appendices": len(ir["appendices"]),
            "footnotes": len(footnotes_map),
            "extracted_images": media_ext.counter,
        },
        "detected_metadata": ir["detected_metadata"],
        "review_hints": review_hints,
    }
    report_path = os.path.join(out_dir, "extract_report.json")
    with open(report_path, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print(f"[OK] Extracted -> {ir_path}")
    print(f"     Chapters={report['stats']['chapters']}, Figs={total_figs}, Tbls={total_tbls}, Eqs={total_eqs}, Refs={len(ir['references'])}, BodyChars={body_chars}")
    if review_hints:
        print(f"     Review hints ({len(review_hints)}):")
        for h in review_hints[:10]:
            print(f"       - {h}")


def main():
    parser = argparse.ArgumentParser(description="Losslessly extract an existing .docx thesis into thesis_ir.json.")
    parser.add_argument("--input", required=True, help="Path to source thesis .docx")
    parser.add_argument("--out-dir", required=True, help="Output directory for thesis_ir.json and media/")
    parser.add_argument("--discipline", choices=["science", "humanities"], default="science")
    parser.add_argument("--no-split-colon-headings", action="store_true", help="Do not split '3.1 Title：Body' into H2 + paragraph")
    args = parser.parse_args()
    extract_thesis(args.input, args.out_dir, args.discipline, split_colon_headings=not args.no_split_colon_headings)


if __name__ == "__main__":
    main()
