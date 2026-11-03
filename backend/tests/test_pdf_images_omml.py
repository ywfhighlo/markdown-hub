# -*- coding: utf-8 -*-
"""PDF→MD image reading-order / OCR merge / snapshot filter + OMML edge cases."""
import struct
import sys
import zlib
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.office_to_md import OfficeToMdConverter  # noqa: E402
from backend.converters.docx_reverse import omml_to_latex  # noqa: E402
from xml.etree import ElementTree as ET

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _m(tag, text=None, children=None, **attrs):
    el = ET.Element("{%s}%s" % (M_NS, tag))
    for k, v in attrs.items():
        el.set("{%s}%s" % (M_NS, k), v)
    if text is not None:
        el.text = text
    for c in children or []:
        el.append(c)
    return el


def _t(s):
    return _m("r", children=[_m("t", s)])


def _png_bytes(w=200, h=150):
    """Uncompressed-ish RGB PNG large enough to pass the 5KB filter."""
    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    # Vary pixels so zlib doesn't crush below 5KB
    rows = []
    for y in range(h):
        row = bytearray()
        row.append(0)
        for x in range(w):
            row.extend(((x * 3 + y) & 255, (y * 5) & 255, (x * y) & 255))
        rows.append(bytes(row))
    raw = b"".join(rows)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw, 1))
        + chunk(b"IEND", b"")
    )


@pytest.fixture
def conv(tmp_path):
    return OfficeToMdConverter(output_dir=str(tmp_path))


def test_analyze_pdf_image_accepts_int_size_without_typeerror(conv):
    info = conv._analyze_pdf_image(100, 400, 400)
    assert info["size"] == 100
    assert info["skip_reason"]

    big_int = conv._analyze_pdf_image(50 * 1024, 400, 400)
    assert big_int["size"] == 50 * 1024
    assert big_int.get("skip_reason") is None

    raw = conv._analyze_pdf_image(b"x" * 6000, 400, 400)
    assert raw["size"] == 6000
    assert raw.get("skip_reason") is None


def test_merge_ocr_keeps_extracted_image_refs(conv):
    ocr = "Scanned line one.\nScanned line two."
    imgs = "![图片 1-1](doc_assets/page_1_img_1.png)\n\n![图片 1-2](doc_assets/page_1_img_2.jpg)"
    out = conv._merge_ocr_with_extracted_images(ocr, imgs)
    assert "Scanned line one" in out
    assert "doc_assets/page_1_img_1.png" in out
    assert "提取的图片" in out
    assert conv._merge_ocr_with_extracted_images(ocr, "") == ocr
    assert "page_1_img_1" in conv._merge_ocr_with_extracted_images("", imgs)


def test_merge_ocr_skips_duplicate_when_ocr_already_has_assets(conv):
    ocr = "text\n\n![x](foo_assets/a.png)\n"
    imgs = "![y](foo_assets/b.png)"
    out = conv._merge_ocr_with_extracted_images(ocr, imgs)
    assert out == ocr.rstrip()
    assert "b.png" not in out


def _has_fitz():
    import importlib.util
    return (
        importlib.util.find_spec("fitz") is not None
        or importlib.util.find_spec("pymupdf") is not None
    )


@pytest.mark.skipif(not _has_fitz(), reason="PyMuPDF required")
def test_pdf_extract_inserts_image_in_reading_order(tmp_path, conv, monkeypatch):
    import fitz

    monkeypatch.setattr(conv, "_ocr_pdf", lambda p: "")

    png = tmp_path / "big.png"
    data = _png_bytes(220, 160)
    assert len(data) >= 5 * 1024
    png.write_bytes(data)

    pdf = tmp_path / "flow.pdf"
    doc = fitz.open()
    page = doc.new_page(width=400, height=500)
    page.insert_text((50, 40), "Above image", fontsize=12)
    page.insert_image(fitz.Rect(50, 80, 260, 200), filename=str(png))
    page.insert_text((50, 280), "Below image", fontsize=12)
    doc.save(str(pdf))
    doc.close()

    md = conv._extract_pdf_with_fitz(pdf)
    assert "Above image" in md
    assert "Below image" in md
    assert "![" in md and "_assets/" in md
    # Prefer embedded img over full-page snapshot for this case
    assert "page_1_img_" in md
    i_above = md.index("Above image")
    i_img = md.index("page_1_img_")
    i_below = md.index("Below image")
    assert i_above < i_img < i_below


@pytest.mark.skipif(not _has_fitz(), reason="PyMuPDF required")
def test_pdf_snapshot_filter_uses_png_bytes_not_filesize_int(tmp_path, conv, monkeypatch):
    import fitz

    monkeypatch.setattr(conv, "_ocr_pdf", lambda p: "")

    pdf = tmp_path / "snap.pdf"
    doc = fitz.open()
    page = doc.new_page(width=400, height=400)
    page.insert_text((40, 40), "Figure 1: demo chart", fontsize=14)
    page.draw_rect(fitz.Rect(40, 80, 280, 260), color=(0, 0, 1), width=2)
    doc.save(str(pdf))
    doc.close()

    seen = {"called_with_bytes": False, "called_with_int": False}
    real = conv._analyze_pdf_image

    def wrap(image_data, width, height):
        if isinstance(image_data, (bytes, bytearray, memoryview)):
            seen["called_with_bytes"] = True
        if isinstance(image_data, int):
            seen["called_with_int"] = True
        return real(image_data, width, height)

    monkeypatch.setattr(conv, "_analyze_pdf_image", wrap)
    conv._extract_pdf_with_fitz(pdf)
    assert seen["called_with_bytes"], "snapshot filter must pass PNG bytes"
    assert not seen["called_with_int"], "must not pass st_size int into _analyze_pdf_image"


def test_omml_limlow_limupp_boxed_phantom_delim_nested():
    lim = _m(
        "limLow",
        children=[_m("e", children=[_t("lim")]), _m("lim", children=[_t("n")])],
    )
    assert omml_to_latex(lim) == "{lim}_{n}"

    upp = _m(
        "limUpp",
        children=[_m("e", children=[_t("x")]), _m("lim", children=[_t("*")])],
    )
    assert "}^{" in omml_to_latex(upp)

    boxed = _m("borderBox", children=[_m("e", children=[_t("42")])])
    assert omml_to_latex(boxed) == r"\boxed{42}"

    phant = _m("phant", children=[_m("e", children=[_t("x")])])
    assert omml_to_latex(phant) == r"\phantom{x}"

    d = _m(
        "d",
        children=[
            _m("e", children=[_t("a")]),
            _m("e", children=[_t("b")]),
        ],
    )
    assert omml_to_latex(d) == r"\left(a,b\right)"

    nested = _m(
        "f",
        children=[
            _m(
                "num",
                children=[
                    _m(
                        "f",
                        children=[
                            _m("num", children=[_t("1")]),
                            _m("den", children=[_t("2")]),
                        ],
                    )
                ],
            ),
            _m("den", children=[_t("3")]),
        ],
    )
    assert omml_to_latex(nested) == r"\frac{\frac{1}{2}}{3}"


def test_omml_groupchr_underbrace():
    g = _m(
        "groupChr",
        children=[
            _m("groupChrPr", children=[_m("pos", val="bot")]),
            _m("e", children=[_t("abc")]),
        ],
    )
    assert omml_to_latex(g) == r"\underbrace{abc}"
