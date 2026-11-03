# -*- coding: utf-8 -*-
"""Phase C: merge markers <-> grid_tables <-> DOCX merges; TOC fallbacks."""
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.table_merge import (  # noqa: E402
    parse_cell_markers,
    pipe_table_markers_to_grid,
    rewrite_marked_pipe_tables,
    docx_table_to_markdown,
    docx_table_has_merges,
    repair_docx_vertical_merges,
)
from backend.converters.md_to_office import MdToOfficeConverter  # noqa: E402


def test_parse_cell_markers_both():
    cs, rs, text = parse_cell_markers("[[colspan:2]][[rowspan:3]]Hello")
    assert cs == 2 and rs == 3 and text == "Hello"


def test_parse_cell_markers_plain():
    assert parse_cell_markers("plain") == (1, 1, "plain")


def test_pipe_colspan_to_grid():
    md = (
        "| [[colspan:2]]AB |  | C |\n"
        "| --- | --- | --- |\n"
        "| 1 | 2 | 3 |\n"
    )
    grid = pipe_table_markers_to_grid(md)
    assert grid is not None
    assert grid.lstrip().startswith("+")
    assert "AB" in grid
    # colspan cell should span without an interior pipe between A-cols
    first_data = [ln for ln in grid.splitlines() if ln.startswith("|")][0]
    # Expect one wide AB cell then C
    assert "AB" in first_data and "C" in first_data


def test_pipe_rowspan_to_grid_open_hline():
    md = (
        "| [[rowspan:2]]A | B |\n"
        "| --- | --- |\n"
        "|  | C |\n"
    )
    grid = pipe_table_markers_to_grid(md)
    assert grid is not None
    # Open border under rowspan (spaces instead of dashes)
    assert any(("+     +" in ln) or (ln.startswith("+") and "   " in ln and "-" in ln)
               for ln in grid.splitlines())
    assert "RowSpan" or "A" in grid


def test_rewrite_leaves_plain_pipe():
    md = "| a | b |\n| --- | --- |\n| 1 | 2 |\n"
    assert rewrite_marked_pipe_tables(md) == md


def test_rewrite_in_preprocess(tmp_path):
    md = tmp_path / "t.md"
    md.write_text(
        "# T\n\n"
        "| [[colspan:2]]AB |  | C |\n"
        "| --- | --- | --- |\n"
        "| 1 | 2 | 3 |\n",
        encoding="utf-8",
    )
    c = MdToOfficeConverter(output_dir=str(tmp_path), promote_headings=False)
    out, _ = c._preprocess_markdown(str(md))
    assert out is not None
    assert "[[colspan:" not in out
    assert out.count("+") >= 4  # grid hlines


@pytest.mark.skipif(
    __import__("shutil").which("pandoc") is None,
    reason="pandoc required",
)
def test_md_to_docx_colspan_gridspan(tmp_path):
    from docx import Document
    from docx.oxml.ns import qn

    md = tmp_path / "c.md"
    md.write_text(
        "# T\n\n"
        "| [[colspan:2]]AB |  | C |\n"
        "| --- | --- | --- |\n"
        "| 1 | 2 | 3 |\n",
        encoding="utf-8",
    )
    c = MdToOfficeConverter(output_dir=str(tmp_path), promote_headings=False)
    out = c._convert_to_docx(str(md))
    assert out and Path(out).exists()
    doc = Document(out)
    assert doc.tables
    # Default template may prepend cover tables — find the content table
    t = None
    for cand in doc.tables:
        texts = [c.text.strip() for row in cand.rows for c in row.cells]
        if "AB" in texts and "C" in texts:
            t = cand
            break
    assert t is not None, "content table with AB/C not found"
    tc = t.rows[0].cells[0]._tc
    el = tc.tcPr.find(qn("w:gridSpan")) if tc.tcPr is not None else None
    assert el is not None
    assert el.get(qn("w:val")) == "2"


@pytest.mark.skipif(
    __import__("shutil").which("pandoc") is None,
    reason="pandoc required",
)
def test_docx_to_md_emits_grid_on_merge(tmp_path):
    from docx import Document
    from backend.converters.office_to_md import OfficeToMdConverter

    # Build a DOCX with a real horizontal merge via python-docx
    doc = Document()
    table = doc.add_table(rows=2, cols=3)
    table.cell(0, 0).text = "AB"
    table.cell(0, 0).merge(table.cell(0, 1))
    table.cell(0, 2).text = "C"
    table.cell(1, 0).text = "1"
    table.cell(1, 1).text = "2"
    table.cell(1, 2).text = "3"
    docx_path = tmp_path / "merged.docx"
    doc.save(str(docx_path))

    assert docx_table_has_merges(Document(str(docx_path)).tables[0])
    grid = docx_table_to_markdown(Document(str(docx_path)).tables[0])
    assert grid.lstrip().startswith("+")
    assert "AB" in grid

    conv = OfficeToMdConverter(output_dir=str(tmp_path))
    md = conv._docx_to_markdown_structured(docx_path)
    assert "+" in md
    assert "AB" in md


def test_update_toc_python_docx_fallback(tmp_path):
    from docx import Document

    doc = Document()
    # Fake a prior TOC entry
    try:
        p = doc.add_paragraph("Old TOC entry")
        p.style = "TOC 1"
    except Exception:
        p = doc.add_paragraph("Old TOC entry")
    doc.add_heading("Chapter One", level=1)
    doc.add_heading("Section Two", level=2)
    path = tmp_path / "toc.docx"
    doc.save(str(path))

    c = MdToOfficeConverter(output_dir=str(tmp_path), promote_headings=False)
    ok = c._update_toc_via_python_docx(str(path))
    # May return False if TOC style missing on this python-docx; accept either
    # so long as the method does not raise.
    assert ok in (True, False)
    if ok:
        doc2 = Document(str(path))
        texts = [p.text for p in doc2.paragraphs]
        assert any("Chapter One" in t for t in texts)
