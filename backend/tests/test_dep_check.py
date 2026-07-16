"""Tests for dep_check: platform-specific install hints.

Covers ``install_hint_for()`` — the single source of truth for how to
install each native dependency on Windows/macOS/Linux. Converters call
this when a dependency is missing so users see an actionable command
instead of a bare "未安装".
"""
import sys
from pathlib import Path

import pytest

# Make backend importable when run from project root
ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

import backend.converters.dep_check as dc  # noqa: E402


# ─────────────────────────────────────────────
# install_hint_for — every known command has a hint
# ─────────────────────────────────────────────

@pytest.mark.parametrize("cmd", [
    "pandoc", "tesseract", "java", "mmdc", "drawio", "graphviz", "soffice",
])
def test_known_cmd_returns_nonempty_hint(cmd):
    hint = dc.install_hint_for(cmd)
    assert isinstance(hint, str) and len(hint) > 3


# ─────────────────────────────────────────────
# graphviz — newly added, three platforms
# ─────────────────────────────────────────────

def test_graphviz_hint_windows(monkeypatch):
    monkeypatch.setattr(dc.sys, "platform", "win32")
    assert "graphviz.org" in dc.install_hint_for("graphviz")


def test_graphviz_hint_macos(monkeypatch):
    monkeypatch.setattr(dc.sys, "platform", "darwin")
    assert "brew install graphviz" in dc.install_hint_for("graphviz")


def test_graphviz_hint_linux(monkeypatch):
    monkeypatch.setattr(dc.sys, "platform", "linux")
    assert "apt install graphviz" in dc.install_hint_for("graphviz")


# ─────────────────────────────────────────────
# soffice — newly added, three platforms
# ─────────────────────────────────────────────

def test_soffice_hint_windows(monkeypatch):
    monkeypatch.setattr(dc.sys, "platform", "win32")
    assert "libreoffice.org" in dc.install_hint_for("soffice")


def test_soffice_hint_macos(monkeypatch):
    monkeypatch.setattr(dc.sys, "platform", "darwin")
    assert "brew install --cask libreoffice" in dc.install_hint_for("soffice")


def test_soffice_hint_linux(monkeypatch):
    monkeypatch.setattr(dc.sys, "platform", "linux")
    assert "apt install libreoffice" in dc.install_hint_for("soffice")


# ─────────────────────────────────────────────
# existing commands still resolve (regression guard)
# ─────────────────────────────────────────────

def test_java_hint_linux(monkeypatch):
    monkeypatch.setattr(dc.sys, "platform", "linux")
    assert "apt" in dc.install_hint_for("java") or "openjdk" in dc.install_hint_for("java")


def test_tesseract_hint_macos(monkeypatch):
    monkeypatch.setattr(dc.sys, "platform", "darwin")
    assert "brew install tesseract" in dc.install_hint_for("tesseract")


def test_unknown_cmd_falls_back():
    hint = dc.install_hint_for("nonexistent-tool-xyz")
    assert isinstance(hint, str) and len(hint) > 0  # "请参考官方文档" fallback
