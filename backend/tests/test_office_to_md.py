"""Tests for office_to_md.py — table merge, heading detection, text cleaning, safe import."""
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.office_to_md import (
    _safe_import,
    OfficeToMdConverter,
    _split_table_row,
    _render_table_cell,
)


# ─────────────────────────────────────────
# _safe_import
# ─────────────────────────────────────────

class TestSafeImport(unittest.TestCase):
    def test_import_existing_module(self):
        result = _safe_import("os")
        self.assertIs(result, sys.modules["os"])

    def test_import_nonexistent_returns_none(self):
        result = _safe_import("nonexistent_module_xyz_123")
        self.assertIsNone(result)

    def test_import_error_returns_none(self):
        with patch("importlib.import_module", side_effect=ImportError("boom")):
            result = _safe_import("anything")
        self.assertIsNone(result)


# ─────────────────────────────────────────
# _split_table_row
# ─────────────────────────────────────────

class TestTableRowSplit(unittest.TestCase):
    def test_basic(self):
        self.assertEqual(_split_table_row("| a | b | c |"), ["a", "b", "c"])

    def test_escaped_pipe(self):
        self.assertEqual(_split_table_row(r"| a \| b | c |"), ["a | b", "c"])

    def test_double_backslash(self):
        # r"| a \\ | b |" has 2 backslash chars; bs=2 (even) → pipe is separator.
        # bs//2=1 backslash goes into the first cell → 'a \' (a, space, backslash).
        self.assertEqual(_split_table_row(r"| a \\ | b |"), ["a \\\\", "b"])

    def test_triple_backslash(self):
        self.assertEqual(_split_table_row(r"| a \\\| b |"), ["a \| b"])

    def test_backticks_with_escaped_pipe(self):
        self.assertEqual(_split_table_row(r"| `5 \| 3` | long |"), ["`5 | 3`", "long"])

    def test_alignment_markers(self):
        self.assertEqual(_split_table_row("| :---: | ---: |"), [":---:", "---:"])

    def test_single_column(self):
        self.assertEqual(_split_table_row("| a |"), ["a"])

    def test_empty(self):
        self.assertEqual(_split_table_row("|  "), [])

    def test_render_then_split_idempotent(self):
        """Render a cell with pipe, then split — should round-trip cleanly."""
        original = r"| `cmd \| opt` | note |"
        cells = _split_table_row(original)
        rendered = "| " + " | ".join(_render_table_cell(c) for c in cells) + " |"
        self.assertEqual(_split_table_row(rendered), cells)


# ─────────────────────────────────────────
# _render_table_cell
# ─────────────────────────────────────────

class TestRenderTableCell(unittest.TestCase):
    def test_escapes_pipe(self):
        self.assertEqual(_render_table_cell("a | b"), r"a \| b")

    def test_no_pipe_unchanged(self):
        self.assertEqual(_render_table_cell("plain text"), "plain text")

    def test_empty_string(self):
        self.assertEqual(_render_table_cell(""), "")

    def test_multiple_pipes(self):
        self.assertEqual(_render_table_cell("a | b | c"), r"a \| b \| c")


# ─────────────────────────────────────────
# _detect_table_regions
# ─────────────────────────────────────────

class TestDetectTableRegions(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_no_tables(self):
        lines = ["# Heading", "Some paragraph text", ""]
        self.assertEqual(self.conv._detect_table_regions(lines), [])

    def test_two_rows(self):
        lines = ["| a | b |", "| c | d |"]
        self.assertEqual(self.conv._detect_table_regions(lines), [(0, 2)])

    def test_three_rows(self):
        lines = ["| a | b |", "| c | d |", "| e | f |"]
        self.assertEqual(self.conv._detect_table_regions(lines), [(0, 3)])

    def test_single_row_ignored(self):
        lines = ["| only | one |"]
        self.assertEqual(self.conv._detect_table_regions(lines), [])

    def test_separator_breaks_region(self):
        # The separator line is not a definitive row, so it breaks the scan.
        lines = ["| a | b |", "|---|----|", "| c | d |"]
        self.assertEqual(self.conv._detect_table_regions(lines), [])

    def test_two_tables_with_blank_line(self):
        lines = ["| a | b |", "| c | d |", "", "| x | y |", "| z | w |"]
        self.assertEqual(self.conv._detect_table_regions(lines), [(0, 2), (3, 5)])


# ─────────────────────────────────────────
# _is_definitive_table_row
# ─────────────────────────────────────────

class TestIsDefinitiveTableRow(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_valid(self):
        self.assertTrue(self.conv._is_definitive_table_row("| a | b |"))

    def test_valid_with_content(self):
        self.assertTrue(self.conv._is_definitive_table_row("| hello world | 123 |"))

    def test_not_pipe_start(self):
        self.assertFalse(self.conv._is_definitive_table_row("a | b"))

    def test_single_column(self):
        self.assertFalse(self.conv._is_definitive_table_row("| only |"))

    def test_dashes_only(self):
        self.assertFalse(self.conv._is_definitive_table_row("| --- | --- |"))

    def test_escaped_pipe_valid(self):
        self.assertTrue(self.conv._is_definitive_table_row(r"| a \| b | c |"))

    def test_too_long_cells(self):
        long = "| " + "x" * 200 + " | " + "y" * 200 + " |"
        self.assertFalse(self.conv._is_definitive_table_row(long))


# ─────────────────────────────────────────
# Table row merging
# ─────────────────────────────────────────

class TestTableRowMerging(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_should_merge_similar_rows(self):
        self.assertTrue(self.conv._should_merge_table_rows("| a | b |", "| c | d |"))

    def test_different_column_counts(self):
        self.assertFalse(self.conv._should_merge_table_rows("| a | b |", "| c | d | e |"))

    def test_merge_combines_content(self):
        merged = self.conv._merge_table_rows("| cmd | |", "| --help | show help |")
        cells = _split_table_row(merged)
        self.assertEqual(len(cells), 2)
        self.assertIn("cmd", cells[0])

    def test_merge_escapes_pipes(self):
        merged = self.conv._merge_table_rows(r"| a \| b | c |", r"| d | e |")
        cells = _split_table_row(merged)
        self.assertEqual(len(cells), 2)


# ─────────────────────────────────────────
# Heading thresholds
# ─────────────────────────────────────────

class TestHeadingThreshold(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_empty(self):
        self.assertEqual(self.conv._compute_heading_thresholds([]), {})

    def test_body_only(self):
        self.assertEqual(self.conv._compute_heading_thresholds([12.0] * 10), {})

    def test_headings_detected(self):
        sizes = [12.0] * 10 + [16.0] * 3 + [20.0] * 2 + [24.0] * 1
        self.assertEqual(self.conv._compute_heading_thresholds(sizes), {1: 24.0, 2: 20.0, 3: 16.0})

    def test_max_five_levels(self):
        sizes = [12.0] + [30.0, 28.0, 26.0, 24.0, 22.0, 20.0, 18.0, 16.0]
        self.assertLessEqual(len(self.conv._compute_heading_thresholds(sizes)), 5)


# ─────────────────────────────────────────
# Heading level detection
# ─────────────────────────────────────────

class TestDetectHeadingLevel(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def make_span(self, text, size, bold=False, flags=0):
        font = "Arial Bold" if bold else "Arial"
        return {"text": text, "size": size, "font": font, "flags": flags,
                "bbox": (0, 0, 100, 20), "color": 0}

    def test_empty_spans(self):
        self.assertIsNone(self.conv._detect_heading_level([], {1: 20.0}))

    def test_empty_thresholds(self):
        span = self.make_span("Hello", 20.0, bold=True)
        self.assertIsNone(self.conv._detect_heading_level([span], {}))

    def test_h1(self):
        span = self.make_span("Chapter One", 24.0, bold=True, flags=16)
        self.assertEqual(self.conv._detect_heading_level([span], {1: 24.0, 2: 20.0}), 1)

    def test_h2(self):
        span = self.make_span("Section", 20.0, bold=True, flags=16)
        self.assertEqual(self.conv._detect_heading_level([span], {1: 24.0, 2: 20.0}), 2)

    def test_too_long_not_heading(self):
        span = self.make_span(" ".join(["word"] * 50), 24.0, bold=True, flags=16)
        self.assertIsNone(self.conv._detect_heading_level([span], {1: 24.0}))

    def test_single_word_no_bold_not_heading(self):
        span = self.make_span("Hello", 20.0, bold=False, flags=0)
        self.assertIsNone(self.conv._detect_heading_level([span], {1: 18.0}))


# ─────────────────────────────────────────
# Bold span detection
# ─────────────────────────────────────────

class TestIsBoldSpan(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_flag_16(self):
        self.assertTrue(self.conv._is_bold_span({"font": "Arial", "flags": 16}))

    def test_bold_in_font_name(self):
        self.assertTrue(self.conv._is_bold_span({"font": "Arial Bold", "flags": 0}))

    def test_black_in_font_name(self):
        self.assertTrue(self.conv._is_bold_span({"font": "Arial Black", "flags": 0}))

    def test_regular_not_bold(self):
        self.assertFalse(self.conv._is_bold_span({"font": "Arial", "flags": 0}))

    def test_chinese_bold(self):
        self.assertTrue(self.conv._is_bold_span({"font": "粗体", "flags": 0}))


# ─────────────────────────────────────────
# Text cleaning
# ─────────────────────────────────────────

class TestCleanText(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_strips_whitespace(self):
        self.assertEqual(self.conv._clean_text("  hello  "), "hello")

    def test_removes_control_chars(self):
        self.assertEqual(self.conv._clean_text("hello\x00world"), "helloworld")

    def test_preserves_chinese(self):
        self.assertEqual(self.conv._clean_text("你好世界"), "你好世界")


# ─────────────────────────────────────────
# Anchor generation
# ─────────────────────────────────────────

class TestGenerateAnchor(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_simple(self):
        anchor = self.conv._generate_anchor("Hello World")
        self.assertIn("hello", anchor)
        self.assertNotIn(" ", anchor)

    def test_chinese(self):
        anchor = self.conv._generate_anchor("第一章 入门")
        self.assertTrue(len(anchor) > 0)
        self.assertEqual(anchor, anchor.lower().replace(" ", "-"))

    def test_special_chars_removed(self):
        anchor = self.conv._generate_anchor("Hello! World?")
        self.assertNotIn("!", anchor)
        self.assertNotIn("?", anchor)


# ─────────────────────────────────────────
# Markdown title cleaning
# ─────────────────────────────────────────

class TestCleanMarkdownTitle(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_strips_hash(self):
        self.assertEqual(self.conv._clean_markdown_title("# Hello"), "Hello")

    def test_strips_multiple_hashes(self):
        self.assertEqual(self.conv._clean_markdown_title("## Section"), "Section")

    def test_strips_trailing_anchor(self):
        self.assertEqual(self.conv._clean_markdown_title("Title {#anchor}"), "Title")

    def test_preserves_chinese(self):
        self.assertEqual(self.conv._clean_markdown_title("# 你好"), "你好")


if __name__ == "__main__":
    unittest.main(verbosity=2)
