# -*- coding: utf-8 -*-
"""DOCX → Markdown reverse helpers: footnotes/endnotes and OMML → LaTeX.

Used by ``office_to_md._docx_to_markdown_structured`` so a pandoc-produced
DOCX (or any Word file with real footnotes / OMML equations) round-trips to
pandoc-native ``[^id]`` / ``$...$`` / ``$$...$$`` Markdown.
"""
from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = "{%s}" % W_NS
M = "{%s}" % M_NS

# Local-name helpers (tolerate prefixed or Clark notation)
def _ln(tag: str) -> str:
    if tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    if ":" in tag:
        return tag.split(":", 1)[-1]
    return tag


def _w_val(el: ET.Element, attr: str = "id") -> Optional[str]:
    return el.get(f"{{{W_NS}}}{attr}") or el.get(attr)


# ── Footnotes / endnotes ─────────────────────────────────────

def _note_text_from_element(note_el: ET.Element) -> str:
    """Flatten a w:footnote / w:endnote element to inline Markdown."""
    parts: List[str] = []
    for p in note_el.iter():
        if _ln(p.tag) != "p":
            continue
        # skip the footnoteRef run marker paragraph lead
        inline = _paragraph_inline_md(p, footnotes={}, endnotes={}, note_mode=True)
        if inline.strip():
            parts.append(inline.strip())
    return " ".join(parts).strip()


def load_notes_from_docx(doc) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Return (footnotes, endnotes) maps of id-str → markdown body.

    Reads ``word/footnotes.xml`` / ``word/endnotes.xml`` via the document
    package relationships. Ids match ``w:footnoteReference/@w:id``.
    """
    footnotes: Dict[str, str] = {}
    endnotes: Dict[str, str] = {}

    def _load(reltype_substr: str, target: Dict[str, str]) -> None:
        try:
            rels = doc.part.rels
        except Exception:
            return
        for rel in rels.values():
            reltype = getattr(rel, "reltype", "") or ""
            if reltype_substr not in reltype:
                continue
            try:
                blob = rel.target_part.blob
            except Exception:
                continue
            try:
                root = ET.fromstring(blob)
            except ET.ParseError:
                continue
            for child in root:
                if _ln(child.tag) not in ("footnote", "endnote"):
                    continue
                note_type = _w_val(child, "type")
                if note_type in ("separator", "continuationSeparator"):
                    continue
                nid = _w_val(child, "id")
                if nid is None:
                    continue
                body = _note_text_from_element(child)
                if body:
                    # Strip a leading footnoteRef marker digit/space if present
                    body = re.sub(r"^[\s\u200b]*", "", body)
                    target[str(nid)] = body

    _load("footnotes", footnotes)
    _load("endnotes", endnotes)
    return footnotes, endnotes


def remap_note_ids(notes: Dict[str, str], prefix: str) -> Tuple[Dict[str, str], Dict[str, str]]:
    """Map Word ids → stable pandoc ids like fn1, en1.

    Returns (word_id → pandoc_id, pandoc_id → body).
    """
    word_to_pandoc: Dict[str, str] = {}
    pandoc_to_body: Dict[str, str] = {}
    # Sort numerically when possible for stable ordering
    def sort_key(k: str):
        try:
            return (0, int(k))
        except ValueError:
            return (1, k)

    n = 1
    for wid in sorted(notes.keys(), key=sort_key):
        pid = f"{prefix}{n}"
        word_to_pandoc[wid] = pid
        pandoc_to_body[pid] = notes[wid]
        n += 1
    return word_to_pandoc, pandoc_to_body


# ── OMML → LaTeX ─────────────────────────────────────────────

_ACCENT = {
    "dot": r"\dot",
    "ddot": r"\ddot",
    "hat": r"\hat",
    "check": r"\check",
    "tilde": r"\tilde",
    "acute": r"\acute",
    "grave": r"\grave",
    "bar": r"\bar",
    "vec": r"\vec",
}

_NARY = {
    "∑": r"\sum",
    "∫": r"\int",
    "∬": r"\iint",
    "∭": r"\iiint",
    "∮": r"\oint",
    "∏": r"\prod",
    "⋃": r"\bigcup",
    "⋂": r"\bigcap",
    "⋁": r"\bigvee",
    "⋀": r"\bigwedge",
}

_FUNC = {
    "sin": r"\sin",
    "cos": r"\cos",
    "tan": r"\tan",
    "log": r"\log",
    "ln": r"\ln",
    "lim": r"\lim",
    "max": r"\max",
    "min": r"\min",
    "sup": r"\sup",
    "inf": r"\inf",
    "exp": r"\exp",
    "det": r"\det",
}


def _text_of(el: ET.Element) -> str:
    bits = []
    for t in el.iter():
        if _ln(t.tag) == "t" and t.text:
            bits.append(t.text)
    return "".join(bits)


def _children_elems(el: ET.Element) -> List[ET.Element]:
    return [c for c in list(el) if isinstance(c.tag, str)]


def _first_child(el: ET.Element, name: str) -> Optional[ET.Element]:
    for c in el:
        if _ln(c.tag) == name:
            return c
    return None


def _latex_escape_text(s: str) -> str:
    # Keep math operators; escape only TeX specials that aren't operators
    repl = {
        "\\": r"\backslash{}",
        "{": r"\{",
        "}": r"\}",
        "%": r"\%",
        "&": r"\&",
        "#": r"\#",
        "_": r"\_",
    }
    out = []
    for ch in s:
        out.append(repl.get(ch, ch))
    return "".join(out)


def omml_to_latex(el: ET.Element) -> str:
    """Convert an ``m:oMath`` / ``m:oMathPara`` / OMML subtree to LaTeX."""
    name = _ln(el.tag)

    if name == "oMathPara":
        inner = _first_child(el, "oMath")
        return omml_to_latex(inner) if inner is not None else "".join(
            omml_to_latex(c) for c in _children_elems(el) if _ln(c.tag) != "oMathParaPr"
        )

    if name == "oMath":
        return "".join(
            omml_to_latex(c) for c in _children_elems(el) if _ln(c.tag) != "oMathParaPr"
        )

    if name == "r":
        # math run: concatenate m:t, ignore rPr
        return _latex_escape_text(_text_of(el))

    if name == "t":
        return _latex_escape_text(el.text or "")

    if name == "f":  # fraction
        num = _first_child(el, "num")
        den = _first_child(el, "den")
        num_l = omml_to_latex(num) if num is not None else ""
        den_l = omml_to_latex(den) if den is not None else ""
        return r"\frac{" + num_l + "}{" + den_l + "}"

    if name == "sSup":
        base = _first_child(el, "e")
        sup = _first_child(el, "sup")
        return (
            "{"
            + (omml_to_latex(base) if base is not None else "")
            + "}^{"
            + (omml_to_latex(sup) if sup is not None else "")
            + "}"
        )

    if name == "sSub":
        base = _first_child(el, "e")
        sub = _first_child(el, "sub")
        return (
            "{"
            + (omml_to_latex(base) if base is not None else "")
            + "}_{"
            + (omml_to_latex(sub) if sub is not None else "")
            + "}"
        )

    if name == "sSubSup":
        base = _first_child(el, "e")
        sub = _first_child(el, "sub")
        sup = _first_child(el, "sup")
        return (
            "{"
            + (omml_to_latex(base) if base is not None else "")
            + "}_{"
            + (omml_to_latex(sub) if sub is not None else "")
            + "}^{"
            + (omml_to_latex(sup) if sup is not None else "")
            + "}"
        )

    if name == "sPre":
        # left scripts
        base = _first_child(el, "e")
        sub = _first_child(el, "sub")
        sup = _first_child(el, "sup")
        return (
            "{}_{"
            + (omml_to_latex(sub) if sub is not None else "")
            + "}^{"
            + (omml_to_latex(sup) if sup is not None else "")
            + "}{"
            + (omml_to_latex(base) if base is not None else "")
            + "}"
        )

    if name == "rad":
        deg = _first_child(el, "deg")
        e = _first_child(el, "e")
        body = omml_to_latex(e) if e is not None else ""
        if deg is not None and _text_of(deg).strip():
            return r"\sqrt[" + omml_to_latex(deg) + "]{" + body + "}"
        return r"\sqrt{" + body + "}"

    if name == "nary":
        e = _first_child(el, "e")
        sub = _first_child(el, "sub")
        sup = _first_child(el, "sup")
        narypr = _first_child(el, "naryPr")
        op = r"\int"
        if narypr is not None:
            ch = _first_child(narypr, "chr")
            if ch is not None:
                val = ch.get(f"{{{M_NS}}}val") or ch.get("val") or ""
                op = _NARY.get(val, val or op)
        out = op
        if sub is not None:
            out += "_{" + omml_to_latex(sub) + "}"
        if sup is not None:
            out += "^{" + omml_to_latex(sup) + "}"
        out += "{" + (omml_to_latex(e) if e is not None else "") + "}"
        return out

    if name == "d":  # delimiter
        dpr = _first_child(el, "dPr")
        beg, end = "(", ")"
        if dpr is not None:
            b = _first_child(dpr, "begChr")
            en = _first_child(dpr, "endChr")
            # Empty string means invisible fence (\left. / \right.)
            if b is not None:
                beg = b.get(f"{{{M_NS}}}val")
                if beg is None:
                    beg = b.get("val")
                if beg is None:
                    beg = "("
                if beg == "":
                    beg = "."
            if en is not None:
                end = en.get(f"{{{M_NS}}}val")
                if end is None:
                    end = en.get("val")
                if end is None:
                    end = ")"
                if end == "":
                    end = "."
        bodies = [omml_to_latex(c) for c in _children_elems(el) if _ln(c.tag) == "e"]
        body = ",".join(bodies)
        # Escape fence chars that need \left/\right forms
        def _fence(ch: str) -> str:
            if ch in ".()[]|/":
                return ch
            if ch in "{}":
                return "\\" + ch
            return ch
        return r"\left" + _fence(beg) + body + r"\right" + _fence(end)

    if name == "acc":
        e = _first_child(el, "e")
        apr = _first_child(el, "accPr")
        cmd = r"\hat"
        if apr is not None:
            ch = _first_child(apr, "chr")
            if ch is not None:
                val = ch.get(f"{{{M_NS}}}val") or ch.get("val") or ""
                # map common combining chars roughly
                if val in ("\u0302", "^"):
                    cmd = r"\hat"
                elif val in ("\u0307",):
                    cmd = r"\dot"
                elif val in ("\u0308",):
                    cmd = r"\ddot"
                elif val in ("\u0304", "\u00af"):
                    cmd = r"\bar"
                elif val in ("\u20d7",):
                    cmd = r"\vec"
                elif val in ("\u0303", "~"):
                    cmd = r"\tilde"
        return cmd + "{" + (omml_to_latex(e) if e is not None else "") + "}"

    if name == "func":
        fname = _first_child(el, "fName")
        e = _first_child(el, "e")
        name_txt = _text_of(fname).strip() if fname is not None else ""
        cmd = _FUNC.get(name_txt, name_txt)
        if cmd and not cmd.startswith("\\") and name_txt.isalpha():
            cmd = "\\" + name_txt
        return cmd + "{" + (omml_to_latex(e) if e is not None else "") + "}"

    if name == "bar":
        e = _first_child(el, "e")
        return r"\overline{" + (omml_to_latex(e) if e is not None else "") + "}"

    if name == "groupChr":
        e = _first_child(el, "e")
        body = omml_to_latex(e) if e is not None else ""
        gpr = _first_child(el, "groupChrPr")
        ch = ""
        pos = ""
        if gpr is not None:
            chr_el = _first_child(gpr, "chr")
            pos_el = _first_child(gpr, "pos")
            if chr_el is not None:
                ch = chr_el.get(f"{{{M_NS}}}val") or chr_el.get("val") or ""
            if pos_el is not None:
                pos = pos_el.get(f"{{{M_NS}}}val") or pos_el.get("val") or ""
        # underbrace U+23DF, overbrace U+23DE
        if ch in (chr(0x23DF), "_") or pos == "bot":
            return r"\underbrace{" + body + "}"
        if ch in (chr(0x23DE), "^") or pos == "top":
            return r"\overbrace{" + body + "}"
        return body

    if name == "borderBox":
        e = _first_child(el, "e")
        return r"\boxed{" + (omml_to_latex(e) if e is not None else "") + "}"

    if name == "box":
        e = _first_child(el, "e")
        return omml_to_latex(e) if e is not None else ""

    if name == "limLow":
        e = _first_child(el, "e")
        lim = _first_child(el, "lim")
        return (
            "{"
            + (omml_to_latex(e) if e is not None else "")
            + "}_{"
            + (omml_to_latex(lim) if lim is not None else "")
            + "}"
        )

    if name == "limUpp":
        e = _first_child(el, "e")
        lim = _first_child(el, "lim")
        return (
            "{"
            + (omml_to_latex(e) if e is not None else "")
            + "}^{"
            + (omml_to_latex(lim) if lim is not None else "")
            + "}"
        )

    if name == "phant":
        e = _first_child(el, "e")
        return r"\phantom{" + (omml_to_latex(e) if e is not None else "") + "}"

    if name == "eqArr":
        rows = [_first_child(c, "e") if _ln(c.tag) == "e" else c for c in _children_elems(el) if _ln(c.tag) == "e"]
        parts = [omml_to_latex(r) for r in rows if r is not None]
        return r"\begin{align}" + r" \\ ".join(parts) + r"\end{align}"

    if name == "m":  # matrix
        rows = []
        for mr in _children_elems(el):
            if _ln(mr.tag) != "mr":
                continue
            cells = [omml_to_latex(e) for e in _children_elems(mr) if _ln(e.tag) == "e"]
            rows.append(" & ".join(cells))
        return r"\begin{matrix}" + r" \\ ".join(rows) + r"\end{matrix}"

    if name in ("e", "num", "den", "sub", "sup", "deg", "fName", "lim"):
        return "".join(omml_to_latex(c) for c in _children_elems(el))

    # Properties / unknown: skip props, recurse into content-bearing children
    if name.endswith("Pr") or name in ("ctrlPr", "argPr", "rPr"):
        return ""

    # Fallback: concatenate descendant m:t (lossy but better than dropping)
    texts = _text_of(el)
    if texts:
        return _latex_escape_text(texts)
    return "".join(omml_to_latex(c) for c in _children_elems(el))


def latex_is_display(para_el: ET.Element) -> bool:
    """True if paragraph is essentially a display equation (oMathPara)."""
    has_omath_para = False
    has_other_text = False
    for child in para_el:
        ln = _ln(child.tag)
        if ln == "oMathPara":
            has_omath_para = True
        elif ln == "pPr":
            continue
        elif ln == "r":
            if _text_of(child).strip():
                has_other_text = True
        elif ln == "oMath":
            # bare oMath alone in paragraph → treat as display if no other text
            pass
    if has_omath_para and not has_other_text:
        return True
    # sole oMath child (besides pPr)
    content = [c for c in para_el if _ln(c.tag) not in ("pPr",)]
    if len(content) == 1 and _ln(content[0].tag) == "oMath":
        return True
    return False


# ── Paragraph inline walk ────────────────────────────────────

def _run_to_md(run_el: ET.Element, note_mode: bool = False) -> str:
    """Convert a w:r element to markdown, handling footnote refs inside."""
    # Footnote / endnote reference
    for child in run_el:
        ln = _ln(child.tag)
        if ln == "footnoteReference":
            return ""  # caller handles via marker — see paragraph walker
        if ln == "endnoteReference":
            return ""
        if ln == "footnoteRef" or ln == "endnoteRef":
            return ""  # marker in the note body itself

    text = ""
    for child in run_el:
        if _ln(child.tag) == "t" and child.text:
            text += child.text
        elif _ln(child.tag) == "tab":
            text += "\t"
        elif _ln(child.tag) == "br":
            text += "\n"

    if not text:
        return ""

    # Style from rPr
    bold = italic = code = False
    rPr = _first_child(run_el, "rPr")
    if rPr is not None:
        if _first_child(rPr, "b") is not None or _first_child(rPr, "bCs") is not None:
            bold = True
        if _first_child(rPr, "i") is not None or _first_child(rPr, "iCs") is not None:
            italic = True
        rfonts = _first_child(rPr, "rFonts")
        if rfonts is not None:
            ascii_f = rfonts.get(f"{{{W_NS}}}ascii") or rfonts.get("ascii") or ""
            if any(x in ascii_f for x in ("Mono", "Courier", "Consolas")):
                code = True

    if code:
        text = f"`{text}`"
    if bold and italic:
        text = f"***{text}***"
    elif bold:
        text = f"**{text}**"
    elif italic:
        text = f"*{text}*"
    return text


def _paragraph_inline_md(
    para_el: ET.Element,
    footnotes: Dict[str, str],
    endnotes: Dict[str, str],
    note_mode: bool = False,
    fn_id_map: Optional[Dict[str, str]] = None,
    en_id_map: Optional[Dict[str, str]] = None,
) -> str:
    """Walk a paragraph's children emitting inline MD + math + footnote refs.

    ``footnotes`` / ``endnotes`` here are word_id → pandoc_id maps when
    ``fn_id_map`` is None; if maps provided, use those for refs.
    """
    fn_map = fn_id_map if fn_id_map is not None else footnotes
    en_map = en_id_map if en_id_map is not None else endnotes

    parts: List[str] = []
    for child in para_el:
        ln = _ln(child.tag)
        if ln == "pPr":
            continue
        if ln == "r":
            # Check for footnote/endnote reference inside run
            fn_ref = None
            en_ref = None
            for sub in child:
                if _ln(sub.tag) == "footnoteReference":
                    fn_ref = _w_val(sub, "id")
                elif _ln(sub.tag) == "endnoteReference":
                    en_ref = _w_val(sub, "id")
            if fn_ref is not None and not note_mode:
                pid = fn_map.get(str(fn_ref), f"fn{fn_ref}")
                parts.append(f"[^{pid}]")
                continue
            if en_ref is not None and not note_mode:
                pid = en_map.get(str(en_ref), f"en{en_ref}")
                parts.append(f"[^{pid}]")
                continue
            parts.append(_run_to_md(child, note_mode=note_mode))
        elif ln == "oMath":
            latex = omml_to_latex(child).strip()
            if latex:
                parts.append(f"${latex}$")
        elif ln == "oMathPara":
            latex = omml_to_latex(child).strip()
            if latex:
                # Inline context fallback (display handled at para level)
                parts.append(f"${latex}$")
        elif ln == "hyperlink":
            # Flatten hyperlink text
            for sub in child:
                if _ln(sub.tag) == "r":
                    parts.append(_run_to_md(sub, note_mode=note_mode))
    return "".join(parts)


def paragraph_to_markdown(
    para,
    fn_id_map: Dict[str, str],
    en_id_map: Dict[str, str],
    heading_styles: Optional[Dict[str, int]] = None,
) -> Optional[str]:
    """Convert a python-docx Paragraph to a Markdown block/line.

    Returns None for empty paragraphs. Display math becomes a ``$$...$$`` block.
    """
    heading_styles = heading_styles or {}
    para_el = para._p
    style_name = para.style.name if para.style else ""

    # Display math paragraph
    if latex_is_display(para_el):
        # Find oMathPara or oMath
        latex = ""
        for child in para_el:
            ln = _ln(child.tag)
            if ln in ("oMathPara", "oMath"):
                latex = omml_to_latex(child).strip()
                break
        if latex:
            return f"$$\n{latex}\n$$"
        return None

    inline = _paragraph_inline_md(
        para_el,
        footnotes={},
        endnotes={},
        fn_id_map=fn_id_map,
        en_id_map=en_id_map,
    )
    text = inline.strip()
    if not text:
        return None

    if style_name in heading_styles:
        level = heading_styles[style_name]
        # Strip any accidental math-only edge cases already handled
        return f"{'#' * level} {text}"

    if style_name.startswith("List Bullet"):
        return f"- {text}"
    if style_name.startswith("List Number"):
        return f"1. {text}"
    if style_name in ("Quote", "Intense Quote"):
        return f"> {text}"

    return text


def format_note_definitions(pandoc_to_body: Dict[str, str]) -> str:
    """Emit ``[^id]: body`` definitions block."""
    if not pandoc_to_body:
        return ""
    lines = []
    for pid, body in pandoc_to_body.items():
        body = re.sub(r"\s+", " ", body).strip()
        lines.append(f"[^{pid}]: {body}")
    return "\n".join(lines)
