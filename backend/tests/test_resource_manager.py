"""Tests for resource_manager: Poppler bin directory discovery.

Covers ``get_poppler_bin_path()`` — the helper that locates the
``pdftoppm`` executable inside the extracted Poppler cache and is used
by both ``office_to_md._ocr_pdf`` and ``batch_pdf_to_png.batch_convert``.

Auto-download is disabled via ``MARKDOWN_HUB_NO_AUTO_DOWNLOAD=1`` so the
tests never touch the network; they only exercise cache-hit / cache-miss
/ layout-edge-case branches.
"""
import sys
from pathlib import Path

import pytest

# Make backend importable when run from project root
ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

import backend.resource_manager as rm  # noqa: E402


# ─────────────────────────────────────────────
# Fixtures
# ─────────────────────────────────────────────

@pytest.fixture
def fake_cache(tmp_path, monkeypatch):
    """Redirect CACHE_ROOT to a temp dir and disable auto-download."""
    monkeypatch.setattr(rm, "CACHE_ROOT", tmp_path)
    monkeypatch.setenv("MARKDOWN_HUB_NO_AUTO_DOWNLOAD", "1")
    return tmp_path


def _seed_poppler(cache_root: Path, version: str = "24.02.0") -> Path:
    """Create a fake extracted poppler cache matching the oschwartz10612 layout:
    ``poppler/poppler-<ver>/Library/bin/pdftoppm.exe`` plus a matching
    ``.version`` marker so ``is_cached()`` returns True.
    """
    poppler_dir = cache_root / "poppler"
    bin_dir = poppler_dir / f"poppler-{version}" / "Library" / "bin"
    bin_dir.mkdir(parents=True)
    # Use the Windows exe name; tests run on win32. The rglob lookup in
    # get_poppler_bin_path picks the name via platform.system(), so this
    # matches the production code path on this host.
    (bin_dir / "pdftoppm.exe").write_bytes(b"FAKE")
    (poppler_dir / ".version").write_text(version, encoding="utf-8")
    return bin_dir


# ─────────────────────────────────────────────
# get_poppler_bin_path
# ─────────────────────────────────────────────

def test_finds_bin_when_cached(fake_cache):
    """Cached poppler with correct layout -> returns the bin directory."""
    expected_bin = _seed_poppler(fake_cache)
    result = rm.get_poppler_bin_path()
    assert result is not None
    assert result == expected_bin
    assert (result / "pdftoppm.exe").is_file()


def test_returns_none_when_cache_empty(fake_cache):
    """No cache, auto-download disabled -> None (no network call)."""
    result = rm.get_poppler_bin_path()
    assert result is None


def test_returns_none_when_exe_missing(fake_cache):
    """Cache marked present (version ok) but pdftoppm.exe absent -> None."""
    poppler_dir = fake_cache / "poppler"
    poppler_dir.mkdir()
    (poppler_dir / ".version").write_text("24.02.0", encoding="utf-8")
    result = rm.get_poppler_bin_path()
    assert result is None


def test_returns_none_on_version_mismatch(fake_cache):
    """Version marker mismatch -> is_cached False; with download disabled -> None.

    Guards against silently reusing a stale poppler from a previous version
    when the spec version has moved on.
    """
    _seed_poppler(fake_cache, version="23.08.0")  # spec expects 24.02.0
    result = rm.get_poppler_bin_path()
    assert result is None
