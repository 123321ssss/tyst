#!/usr/bin/env python3
"""
Generate cleaned official TYUST thesis skeleton templates from the 6 official .docx templates,
plus a standalone spine_template.docx (书脊模板).

Each cleaned skeleton preserves 100% of the official Cover (Sec 0), Chinese & English Title Pages (Sec 1),
TYUST logo image, and Originality/Authorization Declaration (Sec 2), while:
  - Stripping placeholder body chapters (Sec 3+) so build_thesis.py injects content cleanly
  - Fixing known official-template defects ('Bold' font name, oversized/bold toc 1/2 styles,
    leftover MDPI styles, inconsistent footer/header styles, sample footnotes)
  - Configuring odd/even headers, updateFields=true, and circled per-page footnotes (①②③)
"""

import argparse
import os
import shutil
import zipfile
from lxml import etree
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

TEMPLATE_MAP = {
    "doc_sci_pro": (
        "太原科技大学博士、硕士学位论文模板-理工类",
        "太原科技大学博士学位论文模板-理工类（专业学位）.docx",
    ),
    "doc_sci_aca": (
        "太原科技大学博士、硕士学位论文模板-理工类",
        "太原科技大学博士学位论文模板-理工类（学术学位）.docx",
    ),
    "mas_sci_pro": (
        "太原科技大学博士、硕士学位论文模板-理工类",
        "太原科技大学硕士学位论文模板-理工类（专业学位）.docx",
    ),
    "mas_sci_aca": (
        "太原科技大学博士、硕士学位论文模板-理工类",
        "太原科技大学硕士学位论文模板-理工类（学术学位）.docx",
    ),
    "mas_hum_pro": (
        "太原科技大学硕士学位论文模板-人文社科类",
        "太原科技大学硕士学位论文模板-人文社科类（专业学位）.docx",
    ),
    "mas_hum_aca": (
        "太原科技大学硕士学位论文模板-人文社科类",
        "太原科技大学硕士学位论文模板-人文社科类（学术学位）.docx",
    ),
}


def set_rfonts(rPr, east_asia="宋体", latin="Times New Roman"):
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), latin)
    rFonts.set(qn("w:hAnsi"), latin)
    rFonts.set(qn("w:eastAsia"), east_asia)
    rFonts.set(qn("w:cs"), latin)


def set_sz(rPr, half_pts: int):
    for tag in ("w:sz", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), str(half_pts))


def set_bold(rPr, val: bool):
    for tag in ("w:b", "w:bCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), "1" if val else "0")


def ensure_child(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def configure_toc_style(doc, style_name: str, left_chars: int, left_dxa: int):
    try:
        st = doc.styles[style_name]
    except KeyError:
        return
    el = st.element
    pPr = ensure_child(el, "w:pPr")
    rPr = ensure_child(el, "w:rPr")

    # Spacing: before 6pt (120 dxa), after 0, single line spacing
    sp = ensure_child(pPr, "w:spacing")
    sp.attrib.clear()
    sp.set(qn("w:before"), "120")
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "240")
    sp.set(qn("w:lineRule"), "auto")

    # Indent: no firstLine, left 0/2/4 chars
    ind = ensure_child(pPr, "w:ind")
    ind.attrib.clear()
    ind.set(qn("w:left"), str(left_dxa))
    ind.set(qn("w:leftChars"), str(left_chars))
    ind.set(qn("w:right"), "0")
    ind.set(qn("w:firstLine"), "0")

    # Alignment: both (justified)
    jc = ensure_child(pPr, "w:jc")
    jc.set(qn("w:val"), "both")

    # Right tab stop at 16.0 cm (9072 dxa) with dot leader
    old_tabs = pPr.find(qn("w:tabs"))
    if old_tabs is not None:
        pPr.remove(old_tabs)
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:leader"), "dot")
    tab.set(qn("w:pos"), "9072")
    tabs.append(tab)
    pPr.append(tabs)

    # Font: 宋体 / Times New Roman, 小四号 (12pt = 24 half-pt), not bold
    set_rfonts(rPr, east_asia="宋体", latin="Times New Roman")
    set_sz(rPr, 24)
    set_bold(rPr, False)


def configure_styles(doc):
    # 1. Normal style: 宋体 / Times New Roman, 小四号 (12pt), both
    if "Normal" in doc.styles:
        el = doc.styles["Normal"].element
        pPr = ensure_child(el, "w:pPr")
        jc = ensure_child(pPr, "w:jc")
        jc.set(qn("w:val"), "both")
        rPr = ensure_child(el, "w:rPr")
        set_rfonts(rPr, east_asia="宋体", latin="Times New Roman")
        set_sz(rPr, 24)

    # 2. Header style: 宋体 / Times New Roman, 五号 (10.5pt = 21), centered, NO default bottom border
    #    (build_thesis.py draws the single bottom border explicitly on sections 3+ so frontmatter NEVER gets a stray rule)
    for name in ("Header", "header"):
        if name in doc.styles:
            el = doc.styles[name].element
            pPr = ensure_child(el, "w:pPr")
            jc = ensure_child(pPr, "w:jc")
            jc.set(qn("w:val"), "center")
            old_bdr = pPr.find(qn("w:pBdr"))
            if old_bdr is not None:
                pPr.remove(old_bdr)
            rPr = ensure_child(el, "w:rPr")
            set_rfonts(rPr, east_asia="宋体", latin="Times New Roman")
            set_sz(rPr, 21)
            set_bold(rPr, False)
            break

    # 3. Footnote Text style: 宋体 / Times New Roman, 小五号 (9pt = 18), left, single spacing
    for name in ("Footnote Text", "footnote text"):
        if name in doc.styles:
            el = doc.styles[name].element
            pPr = ensure_child(el, "w:pPr")
            jc = ensure_child(pPr, "w:jc")
            jc.set(qn("w:val"), "left")
            sp = ensure_child(pPr, "w:spacing")
            sp.attrib.clear()
            sp.set(qn("w:before"), "0")
            sp.set(qn("w:after"), "0")
            sp.set(qn("w:line"), "240")
            sp.set(qn("w:lineRule"), "auto")
            rPr = ensure_child(el, "w:rPr")
            set_rfonts(rPr, east_asia="宋体", latin="Times New Roman")
            set_sz(rPr, 18)
            break

    # 4. TOC 1 / 2 / 3 per TYUST 2026 spec
    configure_toc_style(doc, "toc 1", left_chars=0, left_dxa=0)
    configure_toc_style(doc, "toc 2", left_chars=200, left_dxa=480)
    configure_toc_style(doc, "toc 3", left_chars=400, left_dxa=960)

    # 5. Remove leftover MDPI style if present
    styles_el = doc.styles.element
    for st_el in list(styles_el.findall(qn("w:style"))):
        name_el = st_el.find(qn("w:name"))
        if name_el is not None and "MDPI" in (name_el.get(qn("w:val")) or ""):
            styles_el.remove(st_el)


def clear_hf_part(hf):
    """Remove all paragraphs/content from a header or footer part, leaving one borderless empty paragraph."""
    el = hf._element
    for child in list(el):
        el.remove(child)
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pBdr = OxmlElement("w:pBdr")
    btm = OxmlElement("w:bottom")
    btm.set(qn("w:val"), "none")
    pBdr.append(btm)
    pPr.append(pBdr)
    p.append(pPr)
    el.append(p)


def configure_settings(doc):
    doc.settings.odd_and_even_pages_header_footer = True
    settings_el = doc.settings.element

    # Ensure updateFields = true so Word updates TOC on open if needed
    uf = settings_el.find(qn("w:updateFields"))
    if uf is None:
        uf = OxmlElement("w:updateFields")
        settings_el.append(uf)
    uf.set(qn("w:val"), "true")

    # Configure footnotePr: ①②③ (decimalEnclosedCircle) restarted eachPage
    fn_pr = settings_el.find(qn("w:footnotePr"))
    if fn_pr is None:
        fn_pr = OxmlElement("w:footnotePr")
        settings_el.insert(0, fn_pr)
    for tag, val in (("w:numFmt", "decimalEnclosedCircle"), ("w:numRestart", "eachPage")):
        el = fn_pr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            fn_pr.insert(0, el)
        el.set(qn("w:val"), val)


def fix_frontmatter_font_bugs(doc):
    """Fix the broken font name 'Bold' on English title page (Ph.D.Dissertation / Master's Thesis)."""
    for p in doc.paragraphs:
        for r in p.runs:
            rPr = r._r.find(qn("w:rPr"))
            if rPr is None:
                continue
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is not None:
                for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
                    if rFonts.get(qn(f"w:{attr}")) == "Bold":
                        rFonts.set(qn(f"w:{attr}"), "Times New Roman")


def strip_comments_in_doc(doc):
    """Remove all Word review comment markers from document.xml body."""
    body = doc.element.body
    for tag in ("w:commentRangeStart", "w:commentRangeEnd", "w:commentReference"):
        for el in list(body.findall(".//" + qn(tag))):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)


def strip_after_declaration(doc):
    """
    Keep Sec 0 (Cover), Sec 1 (CN/EN Title Pages), and Sec 2 (Declaration) intact.
    Delete all body elements after Sec 2's terminating sectPr paragraph (except body's final sectPr).
    """
    body = doc.element.body
    children = list(body)
    sect_paras = []
    for idx, ch in enumerate(children):
        if ch.tag == qn("w:p"):
            pPr = ch.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                sect_paras.append(idx)

    if len(sect_paras) < 3:
        raise RuntimeError(f"Expected >=3 section-break paragraphs before Abstract, found {len(sect_paras)}")

    sec2_end_idx = sect_paras[2]  # End of Declaration section (Sec 2)

    # Remove every element after sec2_end_idx except the final <w:sectPr>
    for ch in children[sec2_end_idx + 1:]:
        if ch.tag != qn("w:sectPr"):
            body.remove(ch)

    # Configure all remaining sections' page geometry cleanly
    for i, sec in enumerate(doc.sections):
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.top_margin = Cm(2.54)
        sec.bottom_margin = Cm(2.54)
        sec.left_margin = Cm(3.0)
        sec.right_margin = Cm(2.0)
        sec.header_distance = Cm(1.9)
        sec.footer_distance = Cm(2.2)
        sec.different_first_page_header_footer = False

        # Frontmatter sections (0, 1, 2) must have explicit empty, borderless headers & footers
        if i <= 2:
            sec.header.is_linked_to_previous = False
            sec.even_page_header.is_linked_to_previous = False
            sec.first_page_header.is_linked_to_previous = False
            sec.footer.is_linked_to_previous = False
            sec.even_page_footer.is_linked_to_previous = False
            sec.first_page_footer.is_linked_to_previous = False
            clear_hf_part(sec.header)
            clear_hf_part(sec.even_page_header)
            clear_hf_part(sec.first_page_header)
            clear_hf_part(sec.footer)
            clear_hf_part(sec.even_page_footer)
            clear_hf_part(sec.first_page_footer)
            pg = sec._sectPr.find(qn("w:pgNumType"))
            if pg is not None:
                sec._sectPr.remove(pg)
        else:
            # Sec 3 (placeholder trailing section, becomes Abstract section on oddPage)
            sec.header.is_linked_to_previous = False
            sec.even_page_header.is_linked_to_previous = False
            sec.footer.is_linked_to_previous = False
            sec.even_page_footer.is_linked_to_previous = False
            clear_hf_part(sec.header)
            clear_hf_part(sec.even_page_header)
            clear_hf_part(sec.footer)
            clear_hf_part(sec.even_page_footer)
            tp = ensure_child(sec._sectPr, "w:type")
            tp.set(qn("w:val"), "oddPage")


def clean_zip_residuals(docx_path: str):
    """
    Strip residual sample footnotes (keeping separator & continuationSeparator) and
    unused sample embedded OLE/images (keeping image1.png which is the TYUST logo!).
    """
    tmp_path = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8", errors="ignore")
        # Find which media targets are still referenced in document.xml.rels
        keep_media = set()
        if "word/_rels/document.xml.rels" in zin.namelist():
            rels_root = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
            for rel in list(rels_root):
                target = rel.get("Target", "")
                if target.startswith("media/"):
                    rid = rel.get("Id", "")
                    if f'"{rid}"' in doc_xml:
                        keep_media.add("word/" + target)

        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                data = zin.read(name)

                # Skip unused sample media, embeddings, and official template review comments
                if name.startswith("word/media/") and name not in keep_media:
                    continue
                if name.startswith("word/embeddings/"):
                    continue
                if "comments" in name.lower():
                    continue

                # Clean sample footnotes from word/footnotes.xml
                if name == "word/footnotes.xml":
                    root = etree.fromstring(data)
                    for fn in list(root.findall(qn("w:footnote"))):
                        fn_type = fn.get(qn("w:type"))
                        if fn_type not in ("separator", "continuationSeparator"):
                            root.remove(fn)
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")

                # Remove comments overrides from [Content_Types].xml
                if name == "[Content_Types].xml":
                    ct_root = etree.fromstring(data)
                    for ov in list(ct_root):
                        part = ov.get("PartName", "")
                        if "comments" in part.lower():
                            ct_root.remove(ov)
                    data = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

                # Prune dangling media/oleObject/comments relationships from document.xml.rels
                if name == "word/_rels/document.xml.rels":
                    rels_root = etree.fromstring(data)
                    for rel in list(rels_root):
                        target = rel.get("Target", "")
                        rid = rel.get("Id", "")
                        if "comments" in target.lower():
                            rels_root.remove(rel)
                        elif (target.startswith("media/") or target.startswith("embeddings/")) and (f'"{rid}"' not in doc_xml):
                            rels_root.remove(rel)
                    data = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

                zout.writestr(item, data)

    shutil.move(tmp_path, docx_path)


def create_spine_template(out_path: str):
    """
    Create spine_template.docx per TYUST spec Section 五:
    书脊：上方为论文题目，中间为作者姓名，下方为“太原科技大学”。
    距上下边界均为5 cm左右。字体为四号仿宋，行距16磅，段前40磅，段后0磅。
    """
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(5.0)
    sec.bottom_margin = Cm(5.0)
    sec.left_margin = Cm(9.5)
    sec.right_margin = Cm(9.5)

    # Single-column vertical table to represent the centered spine cleanly
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = 1  # center
    cell = tbl.cell(0, 0)
    cell.width = Cm(2.0)

    # Set vertical text direction (tbRl) on cell
    tcPr = cell._tc.get_or_add_tcPr()
    textDir = OxmlElement("w:textDirection")
    textDir.set(qn("w:val"), "tbRl")
    tcPr.append(textDir)
    vAlign = OxmlElement("w:vAlign")
    vAlign.set(qn("w:val"), "center")
    tcPr.append(vAlign)

    # Set cell height to 19.7 cm (29.7 - 5 - 5)
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), str(int(19.7 * 567)))
    trH.set(qn("w:hRule"), "exact")
    trPr.append(trH)

    p = cell.paragraphs[0]
    p.alignment = 1  # center
    pPr = p._p.get_or_add_pPr()
    sp = ensure_child(pPr, "w:spacing")
    sp.set(qn("w:before"), "800")  # 40pt
    sp.set(qn("w:after"), "0")
    sp.set(qn("w:line"), "320")  # 16pt exact
    sp.set(qn("w:lineRule"), "exact")

    r = p.add_run("{{SPINE_TITLE}}　　　{{SPINE_AUTHOR}}　　　太原科技大学")
    rPr = r._r.get_or_add_rPr()
    set_rfonts(rPr, east_asia="仿宋", latin="Times New Roman")
    set_sz(rPr, 28)  # 四号 = 14pt

    doc.save(out_path)


def main():
    parser = argparse.ArgumentParser(description="Prepare cleaned TYUST skeleton templates.")
    parser.add_argument(
        "--src-dir",
        required=True,
        help="Directory containing the official TYUST templates",
    )
    parser.add_argument(
        "--out-dir",
        default=os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "resources", "templates")),
        help="Output directory for cleaned skeleton templates",
    )
    args = parser.parse_args()
    os.makedirs(args.out_dir, exist_ok=True)

    for key, (sub, fname) in TEMPLATE_MAP.items():
        src = os.path.join(args.src_dir, sub, fname)
        dst = os.path.join(args.out_dir, f"{key}.docx")
        doc = Document(src)
        fix_frontmatter_font_bugs(doc)
        configure_styles(doc)
        configure_settings(doc)
        strip_after_declaration(doc)
        strip_comments_in_doc(doc)
        doc.save(dst)
        clean_zip_residuals(dst)
        # Quick verify
        vdoc = Document(dst)
        print(f"[OK] {key}.docx -> sections={len(vdoc.sections)}, paras={len(vdoc.paragraphs)}, tables={len(vdoc.tables)}, size={os.path.getsize(dst)}B")

    spine_path = os.path.join(args.out_dir, "spine_template.docx")
    create_spine_template(spine_path)
    print(f"[OK] spine_template.docx -> size={os.path.getsize(spine_path)}B")


if __name__ == "__main__":
    main()
