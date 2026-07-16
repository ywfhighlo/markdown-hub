"""Tests for resource_manager: Poppler bin directory discovery.

Covers ``get_poppler_bin_path()`` — the helper that locates the
``pdftoppm`` executable inside the extracted Poppler cache and is used
by both ``office_to_md._ocr_pdf`` and ``batch_pdf_to_png.batch_convert``.

Auto-download is disabled via ``MARKDOWN_HUB_NO_AUTO_DOWNLOAD=1`` so the
tests never touch the network; they only exercise cache-hit / cache-miss
/ layout-edge-case branches.
"""
import logging
import sys
from pathlib import Path
from urllib.error import URLError

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


# ─────────────────────────────────────────────
# _download — progress logging + retries
# ─────────────────────────────────────────────

class _FakeResponse:
    """Minimal stand-in for an HTTPResponse: context manager + chunked read."""

    def __init__(self, data: bytes, content_length=None):
        self._data = data
        self._pos = 0
        self._cl = content_length

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def getheader(self, name):
        if name.lower() == "content-length" and self._cl is not None:
            return str(self._cl)
        return None

    def read(self, size=-1):
        if self._pos >= len(self._data):
            return b""
        chunk = self._data[self._pos:self._pos + size]
        self._pos += len(chunk)
        return chunk


def _no_sleep(monkeypatch):
    """Avoid real time.sleep during retry backoff in tests."""
    monkeypatch.setattr(rm.time, "sleep", lambda s: None)


def test_download_success_with_progress(tmp_path, monkeypatch, caplog):
    """A normal download writes the full payload and logs percentage progress."""
    data = b"x" * (12 * 1024 * 1024)  # 12 MB -> logs at 10% increments
    monkeypatch.setattr(rm, "urlopen",
                        lambda req, timeout=None: _FakeResponse(data, len(data)))
    _no_sleep(monkeypatch)
    dest = tmp_path / "out.bin"
    caplog.set_level(logging.INFO, logger=rm.logger.name)

    ok = rm._download("http://x/file", dest)

    assert ok
    assert dest.read_bytes() == data
    msgs = [r.message for r in caplog.records]
    assert any("Download progress" in m and "%" in m for m in msgs)
    assert any("Download complete" in m for m in msgs)


def test_download_progress_without_content_length(tmp_path, monkeypatch, caplog):
    """When Content-Length is missing, progress logs bytes with 'size unknown'."""
    data = b"x" * (6 * 1024 * 1024)  # 6 MB -> one 5 MB progress line
    monkeypatch.setattr(rm, "urlopen",
                        lambda req, timeout=None: _FakeResponse(data, None))
    _no_sleep(monkeypatch)
    dest = tmp_path / "out.bin"
    caplog.set_level(logging.INFO, logger=rm.logger.name)

    ok = rm._download("http://x/file", dest)

    assert ok
    msgs = [r.message for r in caplog.records]
    assert any("size unknown" in m for m in msgs)


def test_download_retries_then_succeeds(tmp_path, monkeypatch):
    """First attempt fails (URLError), second succeeds -> True after one retry."""
    data = b"hello"
    attempts = {"n": 0}

    def fake_urlopen(req, timeout=None):
        attempts["n"] += 1
        if attempts["n"] == 1:
            raise URLError("boom")
        return _FakeResponse(data, len(data))

    monkeypatch.setattr(rm, "urlopen", fake_urlopen)
    _no_sleep(monkeypatch)
    dest = tmp_path / "out.bin"

    ok = rm._download("http://x/file", dest)

    assert ok
    assert attempts["n"] == 2
    assert dest.read_bytes() == data


def test_download_all_retries_fail_cleans_partial(tmp_path, monkeypatch):
    """Every attempt fails -> False, and any pre-existing partial is removed."""

    def _always_fail(req, timeout=None):
        raise URLError("boom")

    monkeypatch.setattr(rm, "urlopen", _always_fail)
    _no_sleep(monkeypatch)
    dest = tmp_path / "out.bin"
    dest.write_bytes(b"partial-junk")  # simulate a leftover partial

    ok = rm._download("http://x/file", dest)

    assert not ok
    assert not dest.exists()  # partial cleaned up between/after retries


def test_download_short_read_raises_and_retries(tmp_path, monkeypatch):
    """Content-Length promises 100 but only 50 arrive -> OSError -> retry -> fail."""

    def fake_urlopen(req, timeout=None):
        return _FakeResponse(b"x" * 50, content_length=100)

    monkeypatch.setattr(rm, "urlopen", fake_urlopen)
    _no_sleep(monkeypatch)
    dest = tmp_path / "out.bin"

    ok = rm._download("http://x/file", dest, max_retries=2)

    assert not ok  # short read -> OSError -> retried -> still short -> False
    assert not dest.exists()
