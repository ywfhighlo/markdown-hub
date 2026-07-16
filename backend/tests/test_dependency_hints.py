"""Tests for dependency_hints: platform install hints + error classification.

Covers ``poppler_install_hint()`` (per-platform install guidance shown to
users when Poppler can't be located) and ``is_poppler_missing_error()``
(the classifier that routes pdf2image failures to that hint instead of a
bare exception message).
"""
import sys
from pathlib import Path

import pytest

# Make backend importable when run from project root
ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

import backend.dependency_hints as dh  # noqa: E402


# ─────────────────────────────────────────────
# poppler_install_hint — per-platform content
# ─────────────────────────────────────────────

def test_hint_windows(monkeypatch):
    monkeypatch.setattr(dh.platform, "system", lambda: "Windows")
    hint = dh.poppler_install_hint()
    assert "Poppler" in hint
    assert "poppler-windows" in hint          # manual download URL
    assert "MARKDOWN_HUB_NO_AUTO_DOWNLOAD" in hint  # auto-download option
    assert "--poppler-path" in hint           # explicit override option


def test_hint_macos(monkeypatch):
    monkeypatch.setattr(dh.platform, "system", lambda: "Darwin")
    hint = dh.poppler_install_hint()
    assert "brew install poppler" in hint
    assert "--poppler-path" in hint


def test_hint_linux(monkeypatch):
    monkeypatch.setattr(dh.platform, "system", lambda: "Linux")
    hint = dh.poppler_install_hint()
    # All three distro families listed so users recognise their own
    assert "apt install poppler-utils" in hint
    assert "dnf install poppler-utils" in hint
    assert "pacman -S poppler" in hint
    assert "--poppler-path" in hint


def test_hint_folds_in_context(monkeypatch):
    monkeypatch.setattr(dh.platform, "system", lambda: "Linux")
    hint = dh.poppler_install_hint(context="auto-download failed")
    assert "auto-download failed" in hint


def test_hint_never_empty(monkeypatch):
    for plat in ("Windows", "Darwin", "Linux", "FreeBSD"):
        monkeypatch.setattr(dh.platform, "system", lambda p=plat: p)
        assert len(dh.poppler_install_hint()) > 20


# ─────────────────────────────────────────────
# is_poppler_missing_error — classification
# ─────────────────────────────────────────────

def test_classifies_string_poppler_error():
    assert dh.is_poppler_missing_error(
        Exception("Unable to get page count. Is poppler installed?"))
    assert dh.is_poppler_missing_error(Exception("pdfinfo not found in PATH"))
    assert dh.is_poppler_missing_error(Exception("pdftoppm: command not found"))


def test_does_not_classify_unrelated_error():
    assert not dh.is_poppler_missing_error(Exception("File not found"))
    assert not dh.is_poppler_missing_error(ValueError("bad value"))
    assert not dh.is_poppler_missing_error(Exception("network unreachable"))


def test_classifies_pdf2image_exception_type():
    # If pdf2image is installed, its dedicated exception types are recognised
    # by isinstance, not just string matching.
    try:
        from pdf2image.exceptions import PDFInfoNotInstalledError
    except ImportError:
        pytest.skip("pdf2image not installed in this env")
    assert dh.is_poppler_missing_error(PDFInfoNotInstalledError("missing"))
