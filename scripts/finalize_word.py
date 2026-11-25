#!/usr/bin/env python3
"""
Stage 3: Refresh Word TOC & all fields via Windows COM (Microsoft Word preferred,
fallback to WPS Office Kwps.Application), save updated .docx, export PDF, and render
preview PNGs with PyMuPDF for visual QA.
"""

import argparse
import os
import sys
import time


def update_and_export_via_com(docx_path: str, pdf_path: str = None):
    import pythoncom
    import win32com.client

    abs_docx = os.path.abspath(docx_path)
    abs_pdf = os.path.abspath(pdf_path) if pdf_path else None
    if abs_pdf:
        os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)

    pythoncom.CoInitialize()
    last_err = None
    for prog_id in ("Word.Application", "Kwps.Application"):
        app = None
        doc = None
        try:
            app = win32com.client.DispatchEx(prog_id)
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass

            doc = app.Documents.Open(abs_docx, ConfirmConversions=False, ReadOnly=False)

            # Two-pass repaginate & TOC update (pass 1 populates TOC, pass 2 fixes any shifted page numbers)
            for _ in range(2):
                try:
                    doc.Repaginate()
                except Exception:
                    pass
                try:
                    doc.Fields.Update()
                except Exception:
                    pass
                try:
                    for i in range(1, doc.TablesOfContents.Count + 1):
                        doc.TablesOfContents(i).Update()
                except Exception:
                    pass
                # Also update fields in headers/footers (skip Sec 1..3 frontmatter which must stay empty)
                try:
                    for s_idx in range(4, doc.Sections.Count + 1):
                        sec = doc.Sections(s_idx)
                        for hf_idx in (1, 2, 3):
                            try:
                                sec.Footers(hf_idx).Range.Fields.Update()
                            except Exception:
                                pass
                            try:
                                sec.Headers(hf_idx).Range.Fields.Update()
                            except Exception:
                                pass
                except Exception:
                    pass

            # Normalize TOC entry run formatting (clear inherited bold/Heiti from headings like 摘要/ABSTRACT)
            try:
                for i in range(1, doc.TablesOfContents.Count + 1):
                    toc_rng = doc.TablesOfContents(i).Range
                    toc_rng.Font.Bold = False
                    toc_rng.Font.Size = 12
                    toc_rng.Font.NameFarEast = "宋体"
                    toc_rng.Font.NameAscii = "Times New Roman"
                    toc_rng.Font.NameOther = "Times New Roman"
            except Exception:
                pass

            doc.Save()

            if abs_pdf:
                # 17 = wdExportFormatPDF / wdFormatPDF
                try:
                    doc.ExportAsFixedFormat(OutputFileName=abs_pdf, ExportFormat=17)
                except Exception:
                    doc.SaveAs2(abs_pdf, FileFormat=17)

            print(f"[OK] Finalized via {prog_id}: TOC refreshed & saved -> {abs_docx}")
            if abs_pdf:
                print(f"[OK] Exported PDF -> {abs_pdf}")
            return prog_id
        except Exception as e:
            last_err = e
            print(f"[WARN] COM engine {prog_id} failed: {e}")
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            time.sleep(0.3)

    pythoncom.CoUninitialize()
    raise RuntimeError(f"All Office COM engines failed. Last error: {last_err}")


def render_pdf_previews(pdf_path: str, preview_dir: str, dpi: int = 135, all_pages: bool = False):
    import pymupdf

    os.makedirs(preview_dir, exist_ok=True)
    pdf = pymupdf.open(pdf_path)
    total = len(pdf)
    rendered = []

    # Decide which pages to render
    if all_pages or total <= 28:
        indices = list(range(total))
    else:
        # Always render frontmatter (first 12 pages), last 4 pages, plus pages with images/tables/headings
        picks = set(range(min(12, total)))
        picks.update(range(max(0, total - 4), total))
        for idx in range(total):
            page = pdf[idx]
            txt = page.get_text("text")
            if page.get_images() or "表1-" in txt or "表2-" in txt or "表5-" in txt or "参考文献" in txt or "致    谢" in txt:
                picks.add(idx)
        indices = sorted(picks)

    mat = pymupdf.Matrix(dpi / 72.0, dpi / 72.0)
    for idx in indices:
        page = pdf[idx]
        pix = page.get_pixmap(matrix=mat, alpha=False)
        out_png = os.path.join(preview_dir, f"page_{idx + 1:03d}.png")
        pix.save(out_png)
        rendered.append(out_png)

    # Also write a compact page text map for fast diagnostics
    map_lines = []
    for idx in range(total):
        txt = pdf[idx].get_text("text").strip().splitlines()
        first_line = txt[0].strip()[:60] if txt else "(BLANK PAGE)"
        last_line = txt[-1].strip()[:30] if txt else ""
        map_lines.append(f"Page {idx + 1:03d} | top='{first_line}' | bottom='{last_line}'")
    map_path = os.path.join(preview_dir, "page_map.txt")
    with open(map_path, "w", encoding="utf-8") as f:
        f.write("\n".join(map_lines))

    pdf.close()
    print(f"[OK] Rendered {len(rendered)}/{total} PDF pages -> {preview_dir} (see page_map.txt)")
    return total, rendered


def main():
    parser = argparse.ArgumentParser(description="Refresh Word TOC via COM, export PDF, and render preview PNGs.")
    parser.add_argument("--docx", required=True, help="Formatted thesis .docx path")
    parser.add_argument("--pdf", default=None, help="Optional output PDF path")
    parser.add_argument("--preview-dir", default=None, help="Optional directory for PNG page previews")
    parser.add_argument("--dpi", type=int, default=135, help="Preview PNG DPI")
    parser.add_argument("--all-pages", action="store_true", help="Render every PDF page to PNG")
    args = parser.parse_args()

    pdf_path = args.pdf
    if args.preview_dir and not pdf_path:
        pdf_path = os.path.splitext(args.docx)[0] + ".pdf"

    update_and_export_via_com(args.docx, pdf_path=pdf_path)
    if pdf_path and args.preview_dir:
        render_pdf_previews(pdf_path, args.preview_dir, dpi=args.dpi, all_pages=args.all_pages)


if __name__ == "__main__":
    main()
