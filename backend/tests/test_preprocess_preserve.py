# -*- coding: utf-8 -*-
"""Preprocess must preserve heading numbers, fig: captions, footnotes, deflists."""
import logging
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.md_to_office import MdToOfficeConverter  # noqa: E402
from backend.converters.book_assemble import (  # noqa: E402
    assemble_book_markdown,
    looks_like_book_manifest,
)


SAMPLE = """# 1 Introduction

See ![fig:overview](a.png){#fig:overview}.

Term One
: A pandoc definition.

Here is a note[^n1].

[^n1]: Footnote body stays attached.

## 1.1 Details
"""


def test_preprocess_preserves_heading_numbers_and_fig(tmp_path):
    md = tmp_path / "doc.md"
    md.write_text(SAMPLE, encoding="utf-8")
    c = MdToOfficeConverter(output_dir=str(tmp_path), promote_headings=False)
    out, _temps = c._preprocess_markdown(str(md))
    assert out is not None
    assert "# 1 Introduction" in out
    assert "## 1.1 Details" in out
    assert "![fig:overview]" in out
    assert "fig:overview" in out


def test_preprocess_leaves_pandoc_footnotes(tmp_path):
    md = tmp_path / "doc.md"
    md.write_text(SAMPLE, encoding="utf-8")
    c = MdToOfficeConverter(output_dir=str(tmp_path), promote_headings=False)
    out, _ = c._preprocess_markdown(str(md))
    assert "## 脚注" not in out
    assert "[^n1]: Footnote body stays attached." in out
    assert "[^n1]" in out


def test_preprocess_leaves_definition_lists(tmp_path):
    md = tmp_path / "doc.md"
    md.write_text(SAMPLE, encoding="utf-8")
    c = MdToOfficeConverter(output_dir=str(tmp_path), promote_headings=False)
    out, _ = c._preprocess_markdown(str(md))
    assert "- **Term One**:" not in out
    assert "Term One" in out
    assert ": A pandoc definition." in out


def test_scholarly_args_degrade_without_crossref(tmp_path, caplog):
    c = MdToOfficeConverter(
        output_dir=str(tmp_path),
        enable_crossref=True,
        enable_citeproc=True,
    )
    with caplog.at_level(logging.WARNING):
        args = c._scholarly_pandoc_args()
    assert "--citeproc" in args
    # crossref may or may not be installed; if missing, warning is logged
    if "--filter" not in args:
        assert any("pandoc-crossref" in r.message for r in caplog.records)


def test_scholarly_bibliography_path(tmp_path):
    bib = tmp_path / "refs.bib"
    bib.write_text("@article{x, title={T}, author={A}, year={2020}}", encoding="utf-8")
    c = MdToOfficeConverter(
        output_dir=str(tmp_path),
        bibliography=str(bib),
    )
    args = c._scholarly_pandoc_args()
    assert "--citeproc" in args
    assert "--bibliography" in args
    assert str(bib.resolve()) in args


def test_book_assemble_chapters_list(tmp_path):
    ch1 = tmp_path / "01.md"
    ch2 = tmp_path / "02.md"
    ch1.write_text("# Chapter One\n\nHello.\n", encoding="utf-8")
    ch2.write_text("# Chapter Two\n\nWorld.\n", encoding="utf-8")
    manifest = tmp_path / "book.chapters"
    manifest.write_text("title: Demo Book\n01.md\n02.md\n", encoding="utf-8")
    assert looks_like_book_manifest(str(manifest))
    assembled, meta = assemble_book_markdown(str(manifest), work_dir=str(tmp_path))
    body = Path(assembled).read_text(encoding="utf-8")
    assert "Chapter One" in body and "Chapter Two" in body
    assert meta["title"] == "Demo Book"
    assert len(meta["chapters"]) == 2


def test_book_assemble_yaml(tmp_path):
    (tmp_path / "a.md").write_text("# A\n", encoding="utf-8")
    (tmp_path / "b.md").write_text("# B\n", encoding="utf-8")
    manifest = tmp_path / "demo.book.yml"
    manifest.write_text(
        "title: YAML Book\nchapters:\n  - a.md\n  - b.md\n",
        encoding="utf-8",
    )
    assert looks_like_book_manifest(str(manifest))
    assembled, meta = assemble_book_markdown(str(manifest), work_dir=str(tmp_path))
    body = Path(assembled).read_text(encoding="utf-8")
    assert body.index("# A") < body.index("# B")
    assert meta["title"] == "YAML Book"
