#!/usr/bin/env python3
"""
Math typesetting engine for TYUST thesis formatting:
  1. LaTeX -> MathML (latex2mathml) -> Word native editable OMML (<m:oMath>) via
     Microsoft Office's official MML2OMML.XSL stylesheet.
  2. Uniform OMML font & size normalization (Cambria Math, 12pt / sz=24).
  3. Inline mathematical symbol parser converting plain-text tokens like H_L, θ_L*,
     U_Z*, Δy_blur,max, cosθ_L, 10^5, )^T into upright/italic runs + true Word
     superscripts/subscripts while protecting URLs, file paths, and normal prose.
"""

import os
import re
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_XSLT_CACHE = None

# Standard MML2OMML.XSL search paths on Windows
_XSL_CANDIDATES = [
    os.environ.get("TYUST_MML2OMML_XSL", ""),
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office15\MML2OMML.XSL",
]


def find_mml2omml_xsl():
    """Locate Microsoft Office's MML2OMML.XSL stylesheet."""
    for p in _XSL_CANDIDATES:
        if p and os.path.isfile(p):
            return p
    # Fallback: check Windows registry for Office ClickToRun / InstallRoot
    try:
        import winreg
        for root_key in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for sub in (
                r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                r"SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Word\InstallRoot",
            ):
                try:
                    with winreg.OpenKey(root_key, sub) as k:
                        for val_name in ("InstallationPath", "Path"):
                            try:
                                base, _ = winreg.QueryValueEx(k, val_name)
                                for rel in (r"root\Office16\MML2OMML.XSL", r"Office16\MML2OMML.XSL", "MML2OMML.XSL"):
                                    cand = os.path.join(base, rel)
                                    if os.path.isfile(cand):
                                        return cand
                            except OSError:
                                pass
                except OSError:
                    pass
    except Exception:
        pass
    return None


def _get_xslt():
    global _XSLT_CACHE
    if _XSLT_CACHE is not None:
        return _XSLT_CACHE
    xsl_path = find_mml2omml_xsl()
    if not xsl_path:
        raise FileNotFoundError(
            "Cannot locate MML2OMML.XSL. Set TYUST_MML2OMML_XSL or ensure Microsoft Office is installed."
        )
    _XSLT_CACHE = etree.XSLT(etree.parse(xsl_path))
    return _XSLT_CACHE


def preprocess_latex(latex_str: str) -> str:
    """
    Normalize LaTeX macros for clean MathML -> OMML output:
      - Strip enclosing $, $$, \\[ \\]
      - Convert operators that latex2mathml renders italic or raw (\\arg, \\det, \\min, \\max,
        \\sup, \\inf, \\lim) into upright \\mathrm{...}
      - Convert single-letter \\mathrm{X} -> \\text{X} so Word keeps upright <m:nor/> (e.g. ^T, m)
      - Convert \\# to \\text{#}
      - Map horizontal spacing commands (\\qquad, \\quad, \\;\\;, \\;, \\,) into preserved
        \\text{...} spaces so MML2OMML.XSL does not drop <mspace/>
      - Map \\boldsymbol{X} -> \\mathbf{X}
    """
    s = (latex_str or "").strip()
    if s.startswith("$$") and s.endswith("$$") and len(s) > 4:
        s = s[2:-2].strip()
    elif s.startswith("$") and s.endswith("$") and len(s) > 2:
        s = s[1:-1].strip()
    elif s.startswith(r"\[") and s.endswith(r"\]") and len(s) > 4:
        s = s[2:-2].strip()

    # Upright operator normalization
    s = re.sub(r"\\(arg|det|min|max|sup|inf|lim)(?![A-Za-z])", r"\\mathrm{\1}", s)
    # Single-letter \mathrm{X} -> \text{X} so Cambria Math renders upright (<m:nor/>)
    s = re.sub(r"\\mathrm\{([A-Za-z])\}", r"\\text{\1}", s)
    # Handle \# cleanly as upright #
    s = re.sub(r"\\#", r"\\text{#}", s)
    # Handle \boldsymbol
    s = re.sub(r"\\boldsymbol\b", r"\\mathbf", s)
    # Convert LaTeX spacing macros into explicit \text{...} spaces (longest first)
    s = s.replace(r"\qquad", r"\text{      }")
    s = s.replace(r"\quad", r"\text{    }")
    s = s.replace(r"\;\;", r"\text{  }")
    s = s.replace(r"\;", r"\text{ }")
    s = s.replace(r"\,", "\\text{\u2009}")
    return s


def _make_wrpr(font_pt: float = 12.0, math_font: str = "Cambria Math") -> etree._Element:
    rPr = etree.Element(f"{{{W_NS}}}rPr")
    rFonts = etree.SubElement(rPr, f"{{{W_NS}}}rFonts")
    rFonts.set(f"{{{W_NS}}}ascii", math_font)
    rFonts.set(f"{{{W_NS}}}hAnsi", math_font)
    rFonts.set(f"{{{W_NS}}}eastAsia", "宋体")
    rFonts.set(f"{{{W_NS}}}cs", math_font)
    half_pt = str(int(round(font_pt * 2)))
    for tag in ("sz", "szCs"):
        el = etree.SubElement(rPr, f"{{{W_NS}}}{tag}")
        el.set(f"{{{W_NS}}}val", half_pt)
    return rPr


def _apply_wrpr_to_element(parent_el, font_pt: float = 12.0, math_font: str = "Cambria Math", before_tag: str = None):
    """Ensure parent_el has a normalized <w:rPr> with Cambria Math & font_pt."""
    wrpr = parent_el.find(f"{{{W_NS}}}rPr")
    if wrpr is None:
        wrpr = _make_wrpr(font_pt=font_pt, math_font=math_font)
        if before_tag:
            target = parent_el.find(before_tag)
            if target is not None:
                idx = list(parent_el).index(target)
                parent_el.insert(idx, wrpr)
                return
        # Insert after <m:rPr> if present, else at index 0
        mrpr = parent_el.find(f"{{{M_NS}}}rPr")
        if mrpr is not None:
            idx = list(parent_el).index(mrpr) + 1
            parent_el.insert(idx, wrpr)
        else:
            parent_el.insert(0, wrpr)
    else:
        rFonts = wrpr.find(f"{{{W_NS}}}rFonts")
        if rFonts is None:
            rFonts = etree.Element(f"{{{W_NS}}}rFonts")
            wrpr.insert(0, rFonts)
        rFonts.set(f"{{{W_NS}}}ascii", math_font)
        rFonts.set(f"{{{W_NS}}}hAnsi", math_font)
        rFonts.set(f"{{{W_NS}}}eastAsia", "宋体")
        rFonts.set(f"{{{W_NS}}}cs", math_font)
        half_pt = str(int(round(font_pt * 2)))
        for tag in ("sz", "szCs"):
            el = wrpr.find(f"{{{W_NS}}}{tag}")
            if el is None:
                el = etree.SubElement(wrpr, f"{{{W_NS}}}{tag}")
            el.set(f"{{{W_NS}}}val", half_pt)


def normalize_omml_font(omml_el: etree._Element, font_pt: float = 12.0, math_font: str = "Cambria Math"):
    """Normalize every <m:r> and <m:ctrlPr> in an <m:oMath> element to Cambria Math at font_pt, preserving spaces."""
    for mr in omml_el.findall(f".//{{{M_NS}}}r"):
        _apply_wrpr_to_element(mr, font_pt=font_pt, math_font=math_font, before_tag=f"{{{M_NS}}}t")
    for ctrl in omml_el.findall(f".//{{{M_NS}}}ctrlPr"):
        _apply_wrpr_to_element(ctrl, font_pt=font_pt, math_font=math_font)
    for mt in omml_el.findall(f".//{{{M_NS}}}t"):
        if mt.text and (mt.text != mt.text.strip() or not mt.text.strip() or "\u2009" in mt.text):
            mt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return omml_el


def latex_to_omml(latex_str: str, font_pt: float = 12.0) -> etree._Element:
    """
    Convert a LaTeX math expression into a Word-native editable <m:oMath> element
    styled with Cambria Math at `font_pt` (default 12.0pt / 小四号).
    """
    import latex2mathml.converter

    clean_tex = preprocess_latex(latex_str)
    mml_str = latex2mathml.converter.convert(clean_tex)
    mml_root = etree.fromstring(mml_str.encode("utf-8"))
    xslt = _get_xslt()
    result = xslt(mml_root).getroot()
    if result is None:
        raise ValueError(f"MML2OMML.XSL returned empty result for LaTeX: {latex_str}")

    if result.tag == f"{{{M_NS}}}oMathPara":
        inner = result.find(f"{{{M_NS}}}oMath")
        if inner is None:
            raise ValueError(f"MML2OMML.XSL oMathPara had no oMath for LaTeX: {latex_str}")
        omml_el = inner
    elif result.tag == f"{{{M_NS}}}oMath":
        omml_el = result
    else:
        inner = result.find(f".//{{{M_NS}}}oMath")
        if inner is None:
            raise ValueError(f"Unexpected OMML root tag {result.tag} for LaTeX: {latex_str}")
        omml_el = inner

    normalize_omml_font(omml_el, font_pt=font_pt)
    return omml_el


# ==================== INLINE MATH FORMATTER ====================

_RE_PROTECT = re.compile(
    r"(https?://[^\s\u4e00-\u9fff）)]+|"
    r"www\.[^\s\u4e00-\u9fff）)]+|"
    r"(?:DOI|CSTR):[^\s\u4e00-\u9fff]+|"
    r"[A-Za-z]:\\[^\s\u4e00-\u9fff]+)"
)

_GREEK_SET = set("ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψωθφσλτρηκμñû∂")
_UPRIGHT_BASES = {"Δ", "res", "det", "min", "max"}

# Order matters: longer / more specific patterns first
_RE_INLINE_TOKEN = re.compile(
    r"(?:"
    # 1. trig/math function + Greek symbol + optional subscript + optional *
    r"(?<![A-Za-z0-9_])(?P<fn>sin|cos|tan|cot|sec|csc|arg|exp|ln|log)(?P<fn_sym>[θφβαψωγδλστμνρηξζχκ])"
    r"(?:_(?P<fn_sub>\{[^}]+\}|[A-Za-z0-9]+(?:,[A-Za-z0-9]+)*))?(?P<fn_star>\*)?(?![A-Za-z0-9_])"
    r"|"
    # 2. Scientific power 10^5 or 10^{-5}
    r"(?<![A-Za-z0-9_.])(?P<ten>10)\^(?P<ten_exp>\{[^}]+\}|-?\d+)(?![A-Za-z0-9_])"
    r"|"
    # 3. Closing paren/bracket transpose/power: )^T, )^2, )^{-1/2}
    r"(?<=[)\]])\^(?P<paren_exp>\{[^}]+\}|[A-Za-z0-9\-\+\*/]+)"
    r"|"
    # 4. Variable with subscript (+ optional * or ^exp)
    r"(?<![A-Za-z0-9_])(?P<v_base>Δ?[A-Za-zΑ-Ωα-ωθφσλτρηκμñû∂]|res|det|min|max)"
    r"_(?P<v_sub>\{[^}]+\}|[A-Za-z0-9]+(?:,[A-Za-z0-9]+)*)"
    r"(?P<v_sup>\*|\^(?:\{[^}]+\}|[A-Za-z0-9\-\+\*]+))?(?![A-Za-z0-9_])"
    r"|"
    # 5. Greek symbol with * (no underscore), e.g., α*, ψ*
    r"(?<![A-Za-z0-9_])(?P<g_star_base>[αβγδεζηθικλμνξοπρστυφχψω])\*(?![A-Za-z0-9_])"
    r"|"
    # 6. Single-letter variable with ^exp (no underscore), e.g., x^2, s_p already handled
    r"(?<![A-Za-z0-9_])(?P<s_base>[A-Za-zα-ωΑ-Ωθφσλ])\^(?P<s_exp>\{[^}]+\}|[A-Za-z0-9\-\+\*]+)(?![A-Za-z0-9_])"
    r")"
)


def _strip_braces(s: str) -> str:
    if s and len(s) >= 2 and s.startswith("{") and s.endswith("}"):
        return s[1:-1]
    return s


def _emit_base_runs(out: list, base: str):
    if base.startswith("Δ") and len(base) > 1:
        out.append({"text": "Δ", "italic": False, "sup": False, "sub": False})
        rest = base[1:]
        out.append({"text": rest, "italic": rest not in _UPRIGHT_BASES, "sup": False, "sub": False})
    else:
        out.append({"text": base, "italic": base not in _UPRIGHT_BASES, "sup": False, "sub": False})


def _merge_runs(runs: list) -> list:
    merged = []
    for r in runs:
        if not r["text"]:
            continue
        if (
            merged
            and merged[-1]["italic"] == r["italic"]
            and merged[-1]["sup"] == r["sup"]
            and merged[-1]["sub"] == r["sub"]
        ):
            merged[-1]["text"] += r["text"]
        else:
            merged.append(dict(r))
    return merged


def _split_chunk(chunk: str) -> list:
    runs = []
    pos = 0
    for m in _RE_INLINE_TOKEN.finditer(chunk):
        if m.start() > pos:
            runs.append({"text": chunk[pos:m.start()], "italic": False, "sup": False, "sub": False})

        if m.group("fn") is not None:
            runs.append({"text": m.group("fn"), "italic": False, "sup": False, "sub": False})
            runs.append({"text": m.group("fn_sym"), "italic": True, "sup": False, "sub": False})
            if m.group("fn_sub"):
                runs.append({"text": _strip_braces(m.group("fn_sub")), "italic": False, "sup": False, "sub": True})
            if m.group("fn_star"):
                runs.append({"text": "*", "italic": False, "sup": True, "sub": False})

        elif m.group("ten") is not None:
            runs.append({"text": "10", "italic": False, "sup": False, "sub": False})
            runs.append({"text": _strip_braces(m.group("ten_exp")), "italic": False, "sup": True, "sub": False})

        elif m.group("paren_exp") is not None:
            runs.append({"text": _strip_braces(m.group("paren_exp")), "italic": False, "sup": True, "sub": False})

        elif m.group("v_base") is not None:
            _emit_base_runs(runs, m.group("v_base"))
            runs.append({"text": _strip_braces(m.group("v_sub")), "italic": False, "sup": False, "sub": True})
            v_sup = m.group("v_sup")
            if v_sup:
                sup_txt = "*" if v_sup == "*" else _strip_braces(v_sup[1:])
                runs.append({"text": sup_txt, "italic": False, "sup": True, "sub": False})

        elif m.group("g_star_base") is not None:
            runs.append({"text": m.group("g_star_base"), "italic": True, "sup": False, "sub": False})
            runs.append({"text": "*", "italic": False, "sup": True, "sub": False})

        elif m.group("s_base") is not None:
            runs.append({"text": m.group("s_base"), "italic": True, "sup": False, "sub": False})
            runs.append({"text": _strip_braces(m.group("s_exp")), "italic": False, "sup": True, "sub": False})

        pos = m.end()

    if pos < len(chunk):
        runs.append({"text": chunk[pos:], "italic": False, "sup": False, "sub": False})
    return runs


def split_inline_math_runs(text: str) -> list:
    """
    Split plain text into runs with proper italics and true superscripts/subscripts
    for mathematical variables, while leaving URLs, DOIs, and file paths untouched.
    """
    if not text or ("_" not in text and "^" not in text and "*" not in text):
        return [{"text": text, "italic": False, "sup": False, "sub": False}] if text else []

    out = []
    pos = 0
    for m in _RE_PROTECT.finditer(text):
        if m.start() > pos:
            out.extend(_split_chunk(text[pos:m.start()]))
        out.append({"text": m.group(0), "italic": False, "sup": False, "sub": False})
        pos = m.end()
    if pos < len(text):
        out.extend(_split_chunk(text[pos:]))
    return _merge_runs(out)
