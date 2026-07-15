"""Tests for the pandoc code block highlight argument builder."""
import logging
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.md_to_office import MdToOfficeConverter  # noqa: E402


def _make_converter(theme: str, tmp_path) -> MdToOfficeConverter:
    """Build a converter with the given highlight theme, without running real conversion."""
    c = MdToOfficeConverter(output_dir=str(tmp_path), code_highlight_theme=theme)
    return c


@pytest.mark.parametrize("theme,expected_arg", [
    ("pygments",   "--highlight-style=pygments"),
    ("tango",      "--highlight-style=tango"),
    ("espresso",   "--highlight-style=espresso"),
    ("zenburn",    "--highlight-style=zenburn"),
    ("kate",       "--highlight-style=kate"),
    ("monochrome", "--highlight-style=monochrome"),
    ("breezedark", "--highlight-style=breezedark"),
    ("haddock",    "--highlight-style=haddock"),
])
def test_builtin_themes(tmp_path, theme, expected_arg):
    c = _make_converter(theme, tmp_path)
    args = c._highlight_style_args()
    assert args == [expected_arg], f"Theme '{theme}' should produce {expected_arg}, got {args}"


@pytest.mark.parametrize("off_value", ["off", "none", "disable", "OFF", "Off"])
def test_off_values(tmp_path, off_value):
    c = _make_converter(off_value, tmp_path)
    args = c._highlight_style_args()
    assert args == ["--no-highlight"], f"'{off_value}' should disable highlight, got {args}"


def test_empty_string_falls_back_to_default(tmp_path):
    """Empty string theme should use the default (pygments), not disable."""
    c = _make_converter("", tmp_path)
    args = c._highlight_style_args()
    assert args == ["--highlight-style=pygments"]


def test_unknown_theme_falls_back(tmp_path, caplog):
    c = _make_converter("does-not-exist", tmp_path)
    with caplog.at_level(logging.WARNING):
        args = c._highlight_style_args()
    assert args == ["--highlight-style=pygments"], "Unknown theme should fall back to pygments"
    assert any("pygments" in r.message for r in caplog.records), "Should log a warning about fallback"


def test_default_theme_is_pygments(tmp_path):
    """When no theme is specified, default should be pygments."""
    c = MdToOfficeConverter(output_dir=str(tmp_path))
    assert c._highlight_style_args() == ["--highlight-style=pygments"]
