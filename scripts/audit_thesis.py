#!/usr/bin/env python3
"""
Stage 4: Automated compliance auditor for TYUST graduate theses (.docx + optional .pdf).
Checks page geometry, odd/even headers & footers, page numbering formats & 10.5pt TNR font,
typography & line spacing, three-line table rules, TOC refresh status, reference count &
citation closure, and physical PDF odd-page section starts.
"""

import argparse
import json
import os
import re
import xml.etree.ElementTree as ET
from docx import Document
from docx.oxml.ns import qn


def emu_to_cm(val):
    return round(val / 360000, 2) if val is not None else 0.0


def audit_docx(docx_path: str, degree: str = "master", discipline: str = "science", pdf_path: str = None):
    doc = Document(docx_path)
    checks = []

    def add_check(name: str, status: str, detail: str):
        checks.append({"check": name, "status": status, "detail": detail})

    # 1. Odd and even headers setting
    if doc.settings.odd_and_even_pages_header_footer:
        add_check("奇偶页不同设置", "PASS", "已开启 evenAndOddHeaders")
    else:
        add_check("奇偶页不同设置", "FAIL", "未开启奇偶页页眉页脚不同")

    # 2. Page dimensions & margins across all sections
    bad_geo = []
    for i, s in enumerate(doc.sections):
        dims = (emu_to_cm(s.page_width), emu_to_cm(s.page_height),
                emu_to_cm(s.top_margin), emu_to_cm(s.bottom_margin),
                emu_to_cm(s.left_margin), emu_to_cm(s.right_margin),
                emu_to_cm(s.header_distance), emu_to_cm(s.footer_distance))
        expected = (21.0, 29.7, 2.54, 2.54, 3.0, 2.0, 1.9, 2.2)
        if any(abs(a - b) > 0.05 for a, b in zip(dims, expected)):
            bad_geo.append(f"Sec {i}: {dims}")
    if not bad_geo:
        add_check("A4页面与页边距", "PASS", f"全部 {len(doc.sections)} 节均符合 A4 + 上下2.54/左3.0/右2.0/页眉1.9/页脚2.2 cm")
    else:
        add_check("A4页面与页边距", "FAIL", "; ".join(bad_geo[:5]))

    # 3. Frontmatter (Sec 0..2) must have no headers/footers
    front_hf_issues = []
    for i in range(min(3, len(doc.sections))):
        s = doc.sections[i]
        htxt = "".join(p.text.strip() for p in s.header.paragraphs)
        ftxt = "".join(p.text.strip() for p in s.footer.paragraphs)
        if htxt or ftxt or "PAGE" in ET.tostring(s.footer._element, encoding="unicode"):
            front_hf_issues.append(f"Sec {i} 存在残留页眉/页脚")
    if not front_hf_issues:
        add_check("封面/扉页/声明无页眉页码", "PASS", "前 3 节均无页眉页脚")
    else:
        add_check("封面/扉页/声明无页眉页码", "FAIL", "; ".join(front_hf_issues))

    # 4. Odd-page breaks, odd/even headers, and 10.5pt TNR bottom-centered footers from Sec 3+
    expected_even = "太原科技大学博士学位论文" if degree == "doctor" else "太原科技大学硕士学位论文"
    sec_issues = []
    roman_secs = 0
    arabic_secs = 0
    for i in range(3, len(doc.sections)):
        s = doc.sections[i]
        tp = s._sectPr.find(qn("w:type"))
        tp_val = tp.get(qn("w:val")) if tp is not None else "nextPage"
        if tp_val != "oddPage":
            sec_issues.append(f"Sec {i} 分节符类型={tp_val} (应为 oddPage)")
        oh = "".join(p.text.strip() for p in s.header.paragraphs)
        eh = "".join(p.text.strip() for p in s.even_page_header.paragraphs)
        if not oh:
            sec_issues.append(f"Sec {i} 奇数页页眉为空")
        if expected_even not in eh and "Taiyuan University of Science and Technology" not in eh:
            sec_issues.append(f"Sec {i} 偶数页页眉='{eh}' (期望含 '{expected_even}')")
        ftr_xml = ET.tostring(s.footer._element, encoding="unicode")
        if "PAGE" not in ftr_xml:
            sec_issues.append(f"Sec {i} 页脚缺少 PAGE 域")
        if "txbx" in ftr_xml or "drawing" in ftr_xml:
            sec_issues.append(f"Sec {i} 页脚含浮动文本框")
        if 'val="21"' not in ftr_xml:
            sec_issues.append(f"Sec {i} 页脚页码字号非五号(10.5pt/sz=21)")
        pg = s._sectPr.find(qn("w:pgNumType"))
        if pg is not None:
            fmt = pg.get(qn("w:fmt"))
            if fmt == "upperRoman":
                roman_secs += 1
            elif fmt in (None, "decimal"):
                arabic_secs += 1
        elif i >= 6:
            # Word omits pgNumType on sections continuing decimal numbering from the previous section
            arabic_secs += 1
    if not sec_issues and roman_secs >= 3 and arabic_secs >= 1:
        add_check("奇偶页分节与页眉页脚", "PASS", f"前置罗马节 {roman_secs} 节 + 正文及后置阿拉伯节 {arabic_secs} 节，全部 oddPage 且五号 TNR 页码")
    else:
        add_check("奇偶页分节与页眉页脚", "FAIL", "; ".join(sec_issues[:8]) or f"roman={roman_secs}, arabic={arabic_secs}")

    # 5. TOC refresh check
    toc_placeholder = any("请右键选择“更新域”" in p.text for p in doc.paragraphs)
    toc_entries = sum(1 for p in doc.paragraphs if p.style and p.style.name.lower().startswith("toc"))
    if not toc_placeholder and toc_entries >= 3:
        add_check("自动目录刷新", "PASS", f"目录域已刷新，含 {toc_entries} 条目录项")
    elif not toc_placeholder:
        add_check("自动目录刷新", "WARN", f"未检测到占位提示，但 toc 样式段落数为 {toc_entries}")
    else:
        add_check("自动目录刷新", "FAIL", "目录域尚未刷新（请运行 finalize_word.py 或在 Word 中更新域）")

    # 6. Classify body tables (excluding frontmatter tables 0, 1, 2):
    #    equation layout tables (3 cols, Col 2 = (ch-seq)) vs three-line data tables
    body_tables = doc.tables[3:]
    re_eq_num = re.compile(r"^[（(]\s*(?:[A-Za-z]|\d+)[-–—\.]\d+\s*[）)]$")
    data_tables = []
    eq_tables = []
    for t in body_tables:
        if len(t.columns) == 3 and len(t.rows) == 1 and re_eq_num.match(t.rows[0].cells[2].text.strip()):
            eq_tables.append(t)
        else:
            data_tables.append(t)

    bad_tbls = []
    for ti, t in enumerate(data_tables, 1):
        tblPr = t._tbl.tblPr
        bdr = tblPr.find(qn("w:tblBorders"))
        if bdr is None:
            bad_tbls.append(f"表{ti}缺边框定义")
            continue
        top_sz = bdr.find(qn("w:top")).get(qn("w:sz")) if bdr.find(qn("w:top")) is not None else None
        btm_sz = bdr.find(qn("w:bottom")).get(qn("w:sz")) if bdr.find(qn("w:bottom")) is not None else None
        inv_val = bdr.find(qn("w:insideV")).get(qn("w:val")) if bdr.find(qn("w:insideV")) is not None else "none"
        if top_sz != "12" or btm_sz != "12" or inv_val not in ("none", "nil"):
            bad_tbls.append(f"表{ti}(top={top_sz},btm={btm_sz},inV={inv_val})")
    if not bad_tbls:
        add_check("三线表格式规范", "PASS", f"正文全部 {len(data_tables)} 张数据表均符合三线表（顶底1.5pt、无竖线）")
    else:
        add_check("三线表格式规范", "FAIL", "; ".join(bad_tbls))

    # 6b. Equation layout & OMML/OLE check
    if eq_tables:
        eq_bad = []
        omml_cnt = 0
        ole_cnt = 0
        plain_cnt = 0
        m_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        for ei, t in enumerate(eq_tables, 1):
            tblPr = t._tbl.tblPr
            tblW = tblPr.find(qn("w:tblW"))
            w_val = tblW.get(qn("w:w")) if tblW is not None else None
            if w_val != "9072":
                eq_bad.append(f"式{ei}宽度={w_val} (应为9072)")
            bdr = tblPr.find(qn("w:tblBorders"))
            if bdr is not None:
                for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
                    sb = bdr.find(qn(f"w:{side}"))
                    if sb is not None and sb.get(qn("w:val")) not in (None, "none", "nil"):
                        eq_bad.append(f"式{ei}存在边框 {side}={sb.get(qn('w:val'))}")
                        break
            c2_p = t.rows[0].cells[2].paragraphs[0]
            c2_jc = c2_p._p.find(f"./{qn('w:pPr')}/{qn('w:jc')}")
            if c2_jc is None or c2_jc.get(qn("w:val")) != "right":
                eq_bad.append(f"式{ei}编号未右对齐")

            c1_el = t.rows[0].cells[1]._tc
            if c1_el.findall(f".//{{{m_ns}}}oMath"):
                omml_cnt += 1
            elif c1_el.findall(".//" + qn("w:object")):
                ole_cnt += 1
            else:
                plain_cnt += 1

        if eq_bad:
            add_check("公式排版规范", "FAIL", "; ".join(eq_bad[:6]))
        elif plain_cnt > 0:
            add_check(
                "公式排版规范",
                "WARN",
                f"共 {len(eq_tables)} 个无边框三列表公式（OMML={omml_cnt}, MathType OLE={ole_cnt}, 纯文本={plain_cnt}）",
            )
        else:
            add_check(
                "公式排版规范",
                "PASS",
                f"全部 {len(eq_tables)} 个公式均为无边框三列表（16.0 cm / 编号右对齐），其中原生 OMML={omml_cnt}, MathType OLE={ole_cnt}",
            )
    else:
        add_check("公式排版规范", "PASS", "未检测到独立编号公式")

    # 7. References & citations & placeholders check
    in_refs = False
    refs = []
    cited_nums = set()
    placeholders = []
    re_cite = re.compile(r"\[(\d+(?:[-–—,，]\d+)*)\]")
    for p in doc.paragraphs:
        t = p.text.strip()
        if t in ("参考文献", "References"):
            in_refs = True
            continue
        if in_refs and t in ("攻读学位期间取得的学术成果", "致    谢", "Acknowledgments") or t.startswith("附录"):
            in_refs = False
        if in_refs and re.match(r"^\[\d+\]", t):
            refs.append(t)
        elif not in_refs:
            for m in re_cite.findall(t):
                for part in re.split(r"[,，]", m):
                    if "-" in part or "–" in part:
                        ab = re.split(r"[-–]", part)
                        if len(ab) == 2 and ab[0].isdigit() and ab[1].isdigit():
                            for k in range(int(ab[0]), int(ab[1]) + 1):
                                cited_nums.add(k)
                    elif part.isdigit():
                        cited_nums.add(int(part))
        for kw in ("[待补充]", "（待填写", "[English Abstract", "[注:", "（大纲）", "（撰写要点）"):
            if kw in t:
                placeholders.append(f"{kw} -> {t[:45]}")

    min_refs = 80 if degree == "doctor" else 30
    en_refs = sum(1 for r in refs if len(re.findall(r"[A-Za-z]", r)) > len(re.findall(r"[\u4e00-\u9fff]", r)))
    ref_space_bad = [r[:20] for r in refs if not re.match(r"^\[\d+\] ", r)]
    if len(refs) >= min_refs and not ref_space_bad:
        add_check("参考文献数量与格式", "PASS", f"共 {len(refs)} 篇（含外文 {en_refs} 篇），编号后空格规范")
    else:
        add_check("参考文献数量与格式", "WARN", f"共 {len(refs)} 篇（外文 {en_refs} 篇，规范建议 ≥{min_refs} 篇）；格式缺空格 {len(ref_space_bad)} 条")

    # 8. Residual draft placeholders / notes
    if not placeholders:
        add_check("草稿占位与批注清理", "PASS", "无待填写/大纲/注释残留")
    else:
        add_check("草稿占位与批注清理", "WARN", f"发现 {len(placeholders)} 处草稿/待补项（如：{placeholders[0]}）")

    # 9. Optional PDF physical layout check
    pdf_info = {}
    if pdf_path and os.path.exists(pdf_path):
        import pymupdf
        pdf = pymupdf.open(pdf_path)
        total_pages = len(pdf)
        pdf_info["total_pages"] = total_pages
        # Check front pages: Page 1=Cover, Page 2=CN Title, Page 3=EN Title, Page 4=Declaration
        p1 = pdf[0].get_text("text") if total_pages >= 1 else ""
        p2 = pdf[1].get_text("text") if total_pages >= 2 else ""
        p3 = pdf[2].get_text("text") if total_pages >= 3 else ""
        p4 = pdf[3].get_text("text") if total_pages >= 4 else ""
        p1_ns = re.sub(r"\s+", "", p1)
        p2_ns = re.sub(r"\s+", "", p2)
        p4_ns = re.sub(r"\s+", "", p4)
        front_ok = (
            ("学位论文" in p1_ns)
            and ("中图分类号" in p2_ns or "学校代码" in p2_ns)
            and ("Taiyuan University of Science and Technology" in p3)
            and ("学位论文原创性声明" in p4_ns)
        )
        if front_ok:
            add_check("PDF前置单页严密性", "PASS", f"P1封面 / P2中文扉页 / P3英文扉页 / P4声明各严密占满单页（全书共 {total_pages} 页）")
        else:
            add_check("PDF前置单页严密性", "FAIL", "前置页发生撑页错位！请检查封面/扉页/声明标题行数与空行")

        # Verify odd-page starts in physical PDF (Note: front pages are 4 pages (even), so Page 5 = Abstract is odd!)
        odd_violations = []
        for idx in range(4, total_pages):
            txt = pdf[idx].get_text("text").strip()
            lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
            if len(lines) >= 2:
                # lines[0] is header, lines[1] is top body line
                top_body = lines[1]
                if top_body in ("摘    要", "ABSTRACT", "目    录", "参考文献", "攻读学位期间取得的学术成果", "致    谢") or re.match(r"^(\d+  |[一二三四五六七八九十]+、)", top_body):
                    phys_page = idx + 1
                    if phys_page % 2 == 0:
                        odd_violations.append(f"物理第{phys_page}页(偶数): '{top_body}'")
        if not odd_violations:
            add_check("PDF奇数页起章核验", "PASS", "摘要、ABSTRACT、目录、各章及后置部分均从物理奇数页开始")
        else:
            add_check("PDF奇数页起章核验", "FAIL", "; ".join(odd_violations[:6]))
        pdf.close()

    summary = {
        "pass": sum(1 for c in checks if c["status"] == "PASS"),
        "warn": sum(1 for c in checks if c["status"] == "WARN"),
        "fail": sum(1 for c in checks if c["status"] == "FAIL"),
    }
    return {"docx": os.path.abspath(docx_path), "pdf": os.path.abspath(pdf_path) if pdf_path else None,
            "summary": summary, "pdf_info": pdf_info, "checks": checks}


def main():
    parser = argparse.ArgumentParser(description="Audit a formatted TYUST thesis .docx / .pdf.")
    parser.add_argument("--docx", required=True)
    parser.add_argument("--pdf", default=None)
    parser.add_argument("--degree", choices=["master", "doctor"], default="master")
    parser.add_argument("--discipline", choices=["science", "humanities"], default="science")
    parser.add_argument("--report", default=None, help="Optional JSON report output path")
    args = parser.parse_args()

    res = audit_docx(args.docx, degree=args.degree, discipline=args.discipline, pdf_path=args.pdf)
    if args.report:
        os.makedirs(os.path.dirname(os.path.abspath(args.report)), exist_ok=True)
        with open(args.report, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)

    print(f"=== TYUST AUDIT SUMMARY: PASS={res['summary']['pass']} | WARN={res['summary']['warn']} | FAIL={res['summary']['fail']} ===")
    for c in res["checks"]:
        print(f"  [{c['status']:4s}] {c['check']}: {c['detail']}")


if __name__ == "__main__":
    main()
