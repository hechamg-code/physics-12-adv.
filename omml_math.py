import re
from lxml import etree
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

A14_NS = "http://schemas.microsoft.com/office/drawing/2010/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

SUP_MAP = {
    "⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4",
    "⁵": "5", "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9",
    "⁺": "+", "⁻": "−", "ⁿ": "n", "ⁱ": "i"
}
SUB_MAP = {
    "₀": "0", "₁": "1", "₂": "2", "₃": "3", "₄": "4",
    "₅": "5", "₆": "6", "₇": "7", "₈": "8", "₉": "9",
    "₊": "+", "₋": "−", "ₑ": "e", "ᵢ": "i", "ₒ": "o", "ₓ": "x"
}

UNIT_SLASHES = {
    "N/C", "V/m", "C/m", "C/m²", "C/m³", "μC/m", "μC/m²", "pC/m²",
    "nC/m", "nC/m²", "N·m²/C²", "C²/(N·m²)", "J/C", "kg/mol"
}

def _m_r(text, size_pt=13, bold=False, color="0F172A", italic=None):
    r = etree.Element(f"{{{M_NS}}}r")
    if italic is False:
        rPr_m = etree.SubElement(r, f"{{{M_NS}}}rPr")
        etree.SubElement(rPr_m, f"{{{M_NS}}}sty", val="p")
    rPr = etree.SubElement(r, f"{{{A_NS}}}rPr", lang="en-US", sz=str(int(size_pt * 100)), b="1" if bold else "0")
    sf = etree.SubElement(rPr, f"{{{A_NS}}}solidFill")
    etree.SubElement(sf, f"{{{A_NS}}}srgbClr", val=color)
    t = etree.SubElement(r, f"{{{M_NS}}}t")
    t.text = text
    return r

def _m_frac(num_elems, den_elems):
    f = etree.Element(f"{{{M_NS}}}f")
    num = etree.SubElement(f, f"{{{M_NS}}}num")
    for el in num_elems:
        num.append(el)
    den = etree.SubElement(f, f"{{{M_NS}}}den")
    for el in den_elems:
        den.append(el)
    return f

def _m_sup(base_elems, sup_elems):
    s = etree.Element(f"{{{M_NS}}}sSup")
    e = etree.SubElement(s, f"{{{M_NS}}}e")
    for el in base_elems:
        e.append(el)
    sup = etree.SubElement(s, f"{{{M_NS}}}sup")
    for el in sup_elems:
        sup.append(el)
    return s

def _m_sub(base_elems, sub_elems):
    s = etree.Element(f"{{{M_NS}}}sSub")
    e = etree.SubElement(s, f"{{{M_NS}}}e")
    for el in base_elems:
        e.append(el)
    sub = etree.SubElement(s, f"{{{M_NS}}}sub")
    for el in sub_elems:
        sub.append(el)
    return s

def _m_rad(rad_elems):
    rad = etree.Element(f"{{{M_NS}}}rad")
    radPr = etree.SubElement(rad, f"{{{M_NS}}}radPr")
    etree.SubElement(radPr, f"{{{M_NS}}}degHide", val="1")
    etree.SubElement(rad, f"{{{M_NS}}}deg")
    e = etree.SubElement(rad, f"{{{M_NS}}}e")
    for el in rad_elems:
        e.append(el)
    return rad

def _normalize_unicode_scripts(s):
    """Convert Unicode superscripts/subscripts into ^(...) and _(...) for uniform parsing."""
    # Protect unit slashes temporarily
    for i, u in enumerate(sorted(UNIT_SLASHES, key=len, reverse=True)):
        s = s.replace(u, f"__UNIT{i}__")

    # Convert unicode superscript runs
    out = []
    i = 0
    while i < len(s):
        if s[i] in SUP_MAP:
            buf = []
            while i < len(s) and s[i] in SUP_MAP:
                buf.append(SUP_MAP[s[i]])
                i += 1
            out.append(f"^{{{''.join(buf)}}}")
        elif s[i] in SUB_MAP:
            buf = []
            while i < len(s) and s[i] in SUB_MAP:
                buf.append(SUB_MAP[s[i]])
                i += 1
            out.append(f"_{{{''.join(buf)}}}")
        else:
            out.append(s[i])
            i += 1
    s = "".join(out)

    # Restore units
    for i, u in enumerate(sorted(UNIT_SLASHES, key=len, reverse=True)):
        s = s.replace(f"__UNIT{i}__", u)
    return s

def _find_matching_paren(s, start, open_ch="(", close_ch=")"):
    depth = 0
    for idx in range(start, len(s)):
        if s[idx] == open_ch:
            depth += 1
        elif s[idx] == close_ch:
            depth -= 1
            if depth == 0:
                return idx
    return -1

def _split_top_level_ops(expr):
    """
    Splits expr by top-level additive/relational operators: =, +, -, −, ⇒, ,, ;
    keeping operators as separate tokens, while respecting (), [], {}.
    """
    tokens = []
    buf = []
    depth = 0
    i = 0
    ops = set("=+−-⇒,;≈<>≤≥")
    while i < len(expr):
        ch = expr[i]
        if ch in "([{":
            depth += 1
            buf.append(ch)
        elif ch in ")]}":
            depth = max(0, depth - 1)
            buf.append(ch)
        elif depth == 0 and ch in ops:
            # Handle unary minus at start or after another op
            if ch in "-−" and (not "".join(buf).strip()) and (not tokens or tokens[-1] in ops):
                buf.append(ch)
            else:
                part = "".join(buf).strip()
                if part:
                    tokens.append(part)
                tokens.append(ch)
                buf = []
        else:
            buf.append(ch)
        i += 1
    part = "".join(buf).strip()
    if part:
        tokens.append(part)
    return tokens

def _split_top_level_slash(term):
    """Splits a multiplicative term by top-level '/' (ignoring '/' inside unit tokens)."""
    # Do not split if the whole term ends with a known unit containing '/'
    for u in UNIT_SLASHES:
        if term == u:
            return [term]
    parts = []
    buf = []
    depth = 0
    for i, ch in enumerate(term):
        if ch in "([{":
            depth += 1
            buf.append(ch)
        elif ch in ")]}":
            depth = max(0, depth - 1)
            buf.append(ch)
        elif depth == 0 and ch == "/":
            # Check if this slash is part of a unit like N/C or C/m²
            left_tail = "".join(buf).split()[-1] if "".join(buf).split() else ""
            right_head = term[i+1:].split()[0] if term[i+1:].split() else ""
            cand = f"{left_tail}/{right_head}".rstrip(".,;:)")
            if cand in UNIT_SLASHES:
                buf.append(ch)
            else:
                parts.append("".join(buf).strip())
                buf = []
        else:
            buf.append(ch)
    parts.append("".join(buf).strip())
    return parts

def _parse_atom_list(s, size_pt=13, bold=False, color="0F172A"):
    """Parses a string without top-level '=' or '/' into OMML elements (handling √, ^, _, ())."""
    elems = []
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        # Square root: √(...) or √digit
        if ch == "√":
            if i + 1 < n and s[i + 1] == "(":
                end = _find_matching_paren(s, i + 1, "(", ")")
                if end != -1:
                    inner = s[i + 2:end]
                    elems.append(_expr_to_elems(inner, size_pt, bold, color))
                    rad_el = _m_rad(elems.pop())
                    elems.append(rad_el)
                    i = end + 1
                    continue
            elif i + 1 < n and (s[i + 1].isalnum()):
                j = i + 1
                while j < n and (s[j].isalnum() or s[j] == "."):
                    j += 1
                if j < n and s[j] == "_":
                    sub_elems, next_j = _consume_script(s, j + 1, size_pt, bold, color)
                    rad_el = _m_rad([_m_sub([_m_r(s[i + 1:j], size_pt, bold, color)], sub_elems)])
                    elems.append(rad_el)
                    i = next_j
                else:
                    rad_el = _m_rad([_m_r(s[i + 1:j], size_pt, bold, color)])
                    elems.append(rad_el)
                    i = j
                continue

        # Parenthesized or bracketed group that might have a trailing ^ or _ (e.g. (1.60×10^-19)^2 or [ q1/r1 + q2/r2 ])
        if ch in "([":
            close_ch = ")" if ch == "(" else "]"
            end = _find_matching_paren(s, i, ch, close_ch)
            if end != -1:
                inner = s[i + 1:end]
                group_elems = [_m_r(ch, size_pt, bold, color)] + _expr_to_elems(inner, size_pt, bold, color) + [_m_r(close_ch, size_pt, bold, color)]
                i = end + 1
                # Check for ^ or _ right after ) or ]
                if i < n and s[i] in "^_":
                    op = s[i]
                    script_elems, next_i = _consume_script(s, i + 1, size_pt, bold, color)
                    if op == "^":
                        elems.append(_m_sup(group_elems, script_elems))
                    else:
                        elems.append(_m_sub(group_elems, script_elems))
                    i = next_i
                else:
                    elems.extend(group_elems)
                continue

        # Plain run until special char
        j = i
        while j < n and s[j] not in "√([^_[":
            j += 1
        if j == i:
            # Unmatched special char (e.g. stray '(' or '[' or '√') -> emit literally and advance
            elems.append(_m_r(s[i], size_pt, bold, color))
            i += 1
            continue
        chunk = s[i:j]
        if j < n and s[j] in "^_":
            # Attach ^ or _ to the last token of chunk (e.g. "10", "e", "r", "E", "q_inner")
            op = s[j]
            script_elems, next_i = _consume_script(s, j + 1, size_pt, bold, color)
            # Split chunk into prefix + base_token
            m = re.match(r"^(.*?)([A-Za-z0-9α-ωΑ-ΩεθφλσρμΔΦ]+)$", chunk)
            if m:
                prefix, base = m.group(1), m.group(2)
                if prefix:
                    elems.append(_m_r(prefix, size_pt, bold, color))
                base_el = [_m_r(base, size_pt, bold, color)]
                if op == "^":
                    elems.append(_m_sup(base_el, script_elems))
                else:
                    elems.append(_m_sub(base_el, script_elems))
                i = next_i
                continue
            else:
                if chunk:
                    elems.append(_m_r(chunk, size_pt, bold, color))
                if elems:
                    last = elems.pop()
                    elems.append(_m_sup([last], script_elems) if op == "^" else _m_sub([last], script_elems))
                i = next_i
                continue
        else:
            if chunk:
                elems.append(_m_r(chunk, size_pt, bold, color))
            i = j

    return elems if elems else [_m_r(" ", size_pt, bold, color)]

def _consume_script(s, start, size_pt, bold, color):
    if start < len(s) and s[start] == "{":
        end = _find_matching_paren(s, start, "{", "}")
        if end != -1:
            inner = s[start + 1:end]
            return [_m_r(inner, size_pt, bold, color)], end + 1
    # Single word/number (e.g. _inner, _enc, ^2, ^-1)
    m = re.match(r"^([−-]?\d+|[A-Za-z0-9]+)", s[start:])
    if m:
        txt = m.group(1)
        return [_m_r(txt, size_pt, bold, color)], start + len(txt)
    return [_m_r(s[start:start+1], size_pt, bold, color)], start + 1

def _expr_to_elems(expr, size_pt=13, bold=False, color="0F172A"):
    tokens = _split_top_level_ops(expr.strip())
    elems = []
    for idx, tok in enumerate(tokens):
        if tok in set("=+−-⇒,;≈<>≤≥"):
            if tok == ",":
                elems.append(_m_r(",  ", size_pt, bold, color))
            elif tok == ";":
                elems.append(_m_r(";  ", size_pt, bold, color))
            else:
                elems.append(_m_r(f" {tok} ", size_pt, bold, color))
        else:
            slash_parts = _split_top_level_slash(tok)
            if len(slash_parts) == 1:
                elems.extend(_parse_atom_list(slash_parts[0], size_pt, bold, color))
            else:
                # If wrapped in outer |...| (e.g. |ΔV / Δs|), keep the absolute-value bars outside the fraction
                wrap_abs = (
                    len(slash_parts) == 2
                    and slash_parts[0].startswith("|")
                    and slash_parts[-1].endswith("|")
                    and slash_parts[0].count("|") == 1
                    and slash_parts[-1].count("|") == 1
                )
                num_str = slash_parts[0][1:].strip() if wrap_abs else slash_parts[0]
                den_str = slash_parts[-1][:-1].strip() if wrap_abs else slash_parts[-1]
                cur = _parse_atom_list(num_str, size_pt, bold, color)
                if wrap_abs:
                    den = _parse_atom_list(den_str, size_pt, bold, color)
                    cur = [_m_r("|", size_pt, bold, color), _m_frac(cur, den), _m_r("|", size_pt, bold, color)]
                else:
                    for d_str in slash_parts[1:]:
                        den = _parse_atom_list(d_str, size_pt, bold, color)
                        cur = [_m_frac(cur, den)]
                elems.extend(cur)
    return elems

def add_omml_paragraph(tf, expr_str, size_pt=13, bold=False, color="0F172A", align=PP_ALIGN.CENTER, is_first=False):
    """Adds a native PowerPoint OMML equation paragraph to a text frame."""
    p = tf.paragraphs[0] if is_first else tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(2)
    p.space_after = Pt(3)
    a14m = etree.SubElement(p._p, f"{{{A14_NS}}}m")
    omath_para = etree.SubElement(a14m, f"{{{M_NS}}}oMathPara")
    omath = etree.SubElement(omath_para, f"{{{M_NS}}}oMath")

    norm = _normalize_unicode_scripts(expr_str.strip())
    for el in _expr_to_elems(norm, size_pt=size_pt, bold=bold, color=color):
        omath.append(el)
    return p


_AR_RE = re.compile(r"[\u0600-\u06FF]")
_MATH_CHARS = (
    r"A-Za-z0-9α-ωΑ-ΩεθφλσρμΔΦ"
    r"₀-₉⁰-⁹¹²³⁺⁻₊₋ₑᵢₒₓⁿⁱ"
    r"_·×÷/\^(){}\[\]|.,'+\-−=<>≤≥≈⇒√∮∂→°′∞\u20d7\u0302"
)
# Matches an equation/inequality span (including superscripts ¹²³⁰-⁹, subscripts, ∂, roots, fractions)
_EQ_SPAN_RE = re.compile(
    r"(?<![A-Za-z0-9_])"
    r"(?:[+\-−]?[A-Za-z0-9α-ωΑ-ΩεθφλσρμΔΦ∮√|(∂][" + _MATH_CHARS + r"\s]*)"
    r"[=<>≤≥≈](?:[" + _MATH_CHARS + r"\s]*)"
)

def _has_arabic(s):
    return bool(_AR_RE.search(s))

def _needs_omml_math(eq):
    """Returns True if the equation has fractions, roots, or scripts that benefit from OMML."""
    # Ignore pure unit slashes when checking '/'
    tmp = eq
    for u in sorted(UNIT_SLASHES, key=len, reverse=True):
        tmp = tmp.replace(u, "")
    return ("/" in tmp) or ("√" in tmp) or ("^" in tmp) or any(c in tmp for c in SUP_MAP) or any(c in tmp for c in SUB_MAP) or ("_" in tmp) or ("∮" in tmp) or ("∂" in tmp)

def _balance_eq_span(eq_str):
    """Trim trailing punctuation and any unmatched opening/closing delimiters from an extracted equation."""
    eq = eq_str.strip()
    trailing = ""
    while eq and eq[-1] in ".,;:":
        trailing = eq[-1] + trailing
        eq = eq[:-1].rstrip()
    # Trim trailing unmatched '(' or '[' (e.g. before an Arabic parenthetical comment)
    while eq and eq[-1] in "([":
        trailing = eq[-1] + trailing
        eq = eq[:-1].rstrip()
    # Trim trailing unmatched ')' or ']' if there are more closing than opening
    for op, cl in (("(", ")"), ("[", "]")):
        while eq.endswith(cl) and eq.count(cl) > eq.count(op):
            trailing = cl + trailing
            eq = eq[:-1].rstrip()
    while eq and eq[-1] in ".,;:":
        trailing = eq[-1] + trailing
        eq = eq[:-1].rstrip()
    return eq, trailing

def _add_rtl_para(tf, text, font_size, color_rgb, bold, line_spacing, align, font_name, is_first):
    p = tf.paragraphs[0] if is_first else tf.add_paragraph()
    p.alignment = align
    p.line_spacing = line_spacing
    pPr = p._p.get_or_add_pPr()
    pPr.set("rtl", "1")
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color_rgb is not None:
        run.font.color.rgb = color_rgb
    return p

def _clean_after_text(after):
    after_clean = after.strip()
    if after_clean.startswith("(") and after_clean.endswith(")."):
        after_clean = after_clean[1:-2].strip()
    elif after_clean.startswith("(") and after_clean.endswith(")"):
        after_clean = after_clean[1:-1].strip()
    return after_clean

def add_rich_arabic_math_text(tf, text, font_size=14, color_rgb=None, math_color_hex="0F172A",
                              bold=False, line_spacing=1.2, align=PP_ALIGN.RIGHT,
                              math_align=PP_ALIGN.CENTER, font_name="Arial", math_size_pt=None):
    """
    Renders multi-line Arabic + Math text into a text frame:
    - Arabic prose lines stay pure RTL (no bidi scrambling).
    - Mathematical equations with fractions (/), roots (√), or exponents/subscripts
      are rendered on clean native PowerPoint OMML (<a14:m>) lines with real stacked fractions!
    """
    if math_size_pt is None:
        math_size_pt = max(11, font_size - 1)

    lines = text.split("\n")
    first_para = True

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        # Strip optional step prefix like "1) " or "a) " so it doesn't get swallowed into math
        step_prefix = ""
        m_step = re.match(r"^(\d+\)|[a-zA-Z]\))\s+", line)
        if m_step:
            step_prefix = m_step.group(1)
            line = line[m_step.end():].strip()

        # Case 1: Pure math line (no Arabic characters)
        if not _has_arabic(line):
            parts = [p.strip() for p in re.split(r"\s{2,}\|\s{2,}", line) if p.strip()]
            for idx, part in enumerate(parts):
                if idx == 0 and step_prefix:
                    if _needs_omml_math(part):
                        _add_rtl_para(tf, step_prefix, font_size, color_rgb, bold, line_spacing, align, font_name, first_para)
                        first_para = False
                        add_omml_paragraph(tf, part, size_pt=math_size_pt, bold=bold,
                                           color=math_color_hex, align=math_align, is_first=first_para)
                    else:
                        _add_rtl_para(tf, f"{step_prefix} {part}", font_size, color_rgb, bold, line_spacing, math_align, font_name, first_para)
                else:
                    if _needs_omml_math(part):
                        add_omml_paragraph(tf, part, size_pt=math_size_pt, bold=bold,
                                           color=math_color_hex, align=math_align, is_first=first_para)
                    else:
                        _add_rtl_para(tf, part, font_size, color_rgb, bold, line_spacing, math_align, font_name, first_para)
                first_para = False
            continue

        # Case 2: Mixed Arabic + one or more equations
        segments = [s.strip() for s in re.split(r"\s{2,}\|\s{2,}", line) if s.strip()]
        for seg_idx, seg in enumerate(segments):
            cur_prefix = step_prefix if seg_idx == 0 else ""
            omml_matches = [m for m in _EQ_SPAN_RE.finditer(seg) if _needs_omml_math(m.group(0))]

            if not omml_matches:
                full_seg = f"{cur_prefix} {seg}".strip() if cur_prefix else seg
                _add_rtl_para(tf, full_seg, font_size, color_rgb, bold, line_spacing, align, font_name, first_para)
                first_para = False
                continue

            if len(omml_matches) == 1:
                m = omml_matches[0]
                eq_str, trailing = _balance_eq_span(m.group(0))
                before = seg[:m.start()].strip()
                after = (trailing + seg[m.end():]).strip()

                ar_parts = []
                if cur_prefix:
                    ar_parts.append(cur_prefix)
                if before:
                    ar_parts.append(before.rstrip(":"))
                after_clean = _clean_after_text(after) if after else ""
                if after_clean and len(after_clean) <= 45:
                    ar_parts.append(after_clean)
                    after_clean = ""

                ar_line = " ".join(ar_parts).strip()
                if ar_line:
                    if not ar_line.endswith(":"):
                        ar_line = ar_line.rstrip("=.") + ":"
                    _add_rtl_para(tf, ar_line, font_size, color_rgb, bold, line_spacing, align, font_name, first_para)
                    first_para = False

                add_omml_paragraph(tf, eq_str, size_pt=math_size_pt, bold=False,
                                   color=math_color_hex, align=math_align, is_first=first_para)
                first_para = False

                if after_clean:
                    _add_rtl_para(tf, after_clean, font_size, color_rgb, bold, line_spacing, align, font_name, first_para)
                    first_para = False
            else:
                # Multiple OMML equations on the same segment: interleave Arabic prose and OMML equations
                pos = 0
                pending_prefix = cur_prefix
                for idx, m in enumerate(omml_matches):
                    eq_str, trailing = _balance_eq_span(m.group(0))
                    before = seg[pos:m.start()].strip()
                    before = before.lstrip("،.؛,: ").strip() if pos > 0 else before
                    ar_parts = []
                    if pending_prefix:
                        ar_parts.append(pending_prefix)
                        pending_prefix = ""
                    if before:
                        ar_parts.append(before.rstrip(":"))
                    ar_line = " ".join(ar_parts).strip()
                    if ar_line:
                        if not ar_line.endswith(":"):
                            ar_line = ar_line.rstrip("=.") + ":"
                        _add_rtl_para(tf, ar_line, font_size, color_rgb, bold, line_spacing, align, font_name, first_para)
                        first_para = False

                    add_omml_paragraph(tf, eq_str, size_pt=math_size_pt, bold=False,
                                       color=math_color_hex, align=math_align, is_first=first_para)
                    first_para = False
                    pos = m.end()
                    if trailing:
                        # Re-include trailing punctuation/delimiters for the next prose slice
                        pos -= len(trailing)

                tail = seg[pos:].strip().lstrip("،.؛,: ").strip()
                if tail and tail not in (")", ").", "]", "]."):
                    _add_rtl_para(tf, tail, font_size, color_rgb, bold, line_spacing, align, font_name, first_para)
                    first_para = False

