# -*- coding: utf-8 -*-
"""DOCX→MD footnotes + OMML math reverse conversion."""
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.docx_reverse import (  # noqa: E402
    omml_to_latex,
    load_notes_from_docx,
    remap_note_ids,
    format_note_definitions,
    _ln,
)
from backend.converters.office_to_md import OfficeToMdConverter  # noqa: E402
from xml.etree import ElementTree as ET
import zipfile


HAS_PANDOC = shutil.which("pandoc") is not None


def _write_sample_md(path: Path) -> None:
    path.write_text(
        "# Roundtrip\n\n"
        "Inline math $E=mc^2$ and display:\n\n"
        "$$\n\\frac{a}{b+1}\n$$\n\n"
        "A note[^fn1] here.\n\n"
        "[^fn1]: Footnote body with *emphasis*.\n",
        encoding="utf-8",
    )


@pytest.mark.skipif(not HAS_PANDOC, reason="pandoc required")
def test_omml_to_latex_from_pandoc_docx(tmp_path):
    md = tmp_path / "s.md"
    docx = tmp_path / "s.docx"
    _write_sample_md(md)
    subprocess.check_call(["pandoc", str(md), "-o", str(docx)])
    z = zipfile.ZipFile(docx)
    root = ET.fromstring(z.read("word/document.xml"))
    maths = [el for el in root.iter() if _ln(el.tag) == "oMath"]
    assert len(maths) >= 2
    latex0 = omml_to_latex(maths[0])
    latex1 = omml_to_latex(maths[1])
    assert "E" in latex0 and "c" in latex0 and "2" in latex0
    assert r"\frac" in latex1
    assert "a" in latex1 and "b" in latex1


@pytest.mark.skipif(not HAS_PANDOC, reason="pandoc required")
def test_load_footnotes_from_pandoc_docx(tmp_path):
    from docx import Document

    md = tmp_path / "s.md"
    docx = tmp_path / "s.docx"
    _write_sample_md(md)
    subprocess.check_call(["pandoc", str(md), "-o", str(docx)])
    doc = Document(str(docx))
    fn, en = load_notes_from_docx(doc)
    assert fn, "expected at least one footnote"
    assert any("Footnote body" in v for v in fn.values())
    assert any("emphasis" in v for v in fn.values())
    fmap, fbody = remap_note_ids(fn, "fn")
    block = format_note_definitions(fbody)
    assert block.startswith("[^fn")
    assert "Footnote body" in block


@pytest.mark.skipif(not HAS_PANDOC, reason="pandoc required")
def test_docx_to_md_footnotes_and_math(tmp_path):
    md = tmp_path / "s.md"
    docx = tmp_path / "s.docx"
    _write_sample_md(md)
    subprocess.check_call(["pandoc", str(md), "-o", str(docx)])

    conv = OfficeToMdConverter(output_dir=str(tmp_path))
    out = conv._docx_to_markdown_structured(docx)
    assert "[^fn" in out
    assert "[^fn" in out and "]:" in out
    assert "Footnote body" in out
    assert "$" in out
    assert r"\frac" in out
    assert "$$" in out


@pytest.mark.skipif(not HAS_PANDOC, reason="pandoc required")
def test_roundtrip_md_docx_md_footnote_math(tmp_path):
    """MD → DOCX (pandoc) → MD (our extractor) keeps footnotes + math."""
    md = tmp_path / "s.md"
    docx = tmp_path / "s.docx"
    _write_sample_md(md)
    subprocess.check_call(["pandoc", str(md), "-o", str(docx)])
    conv = OfficeToMdConverter(output_dir=str(tmp_path))
    out = conv._docx_to_markdown_structured(docx)
    # Survive another MD→DOCX→MD hop via pandoc for the extracted md
    md2 = tmp_path / "round.md"
    md2.write_text(out, encoding="utf-8")
    docx2 = tmp_path / "round.docx"
    subprocess.check_call(["pandoc", str(md2), "-o", str(docx2)])
    out2 = conv._docx_to_markdown_structured(docx2)
    assert "Footnote body" in out2
    assert r"\frac" in out2 or "frac" in out2
    assert "$" in out2
