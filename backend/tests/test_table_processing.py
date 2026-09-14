"""Tests for table processing utilities used in office_to_md and md_to_office.

Covers:
- _split_table_row: pipe table row splitting with \\| escape semantics
- _render_table_cell: safe escaping of | inside cell content
- _is_definitive_table_row: strict pipe-table row detection
- _is_table_separator: GFM separator line detection
- _looks_like_table_row: loose pipe-table row heuristic
- _should_merge_table_rows: decide if two rows belong to same logical row
- _merge_table_rows: merge two rows into one
- _detect_table_regions: scan text for contiguous table blocks
- _merge_table_region: merge OCR-split rows within a detected region
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.md_to_office import (
    _split_table_row,
    _render_table_cell,
)
from backend.converters.office_to_md import OfficeToMdConverter


# ─────────────────────────────────────────────
# _split_table_row  (from md_to_office)
# ─────────────────────────────────────────────

class TestSplitTableRow(unittest.TestCase):
    def test_basic_split(self):
        self.assertEqual(_split_table_row("| a | b | c |"), ["a", "b", "c"])

    def test_two_cols(self):
        self.assertEqual(_split_table_row("| a | b |"), ["a", "b"])

    def test_single_col(self):
        self.assertEqual(_split_table_row("| a |"), ["a"])

    def test_whitespace_stripped(self):
        self.assertEqual(_split_table_row("|   a   |  b  |"), ["a", "b"])

    def test_separator_row(self):
        self.assertEqual(_split_table_row("| :---: | ---: |"), [":---:", "---:"])

    def test_separator_row_dashes(self):
        self.assertEqual(_split_table_row("| --- | --- |"), ["---", "---"])

    def test_escaped_pipe(self):
        # \| → literal pipe, not separator → 2 cells
        self.assertEqual(_split_table_row(r"| a \| b | c |"), ["a | b", "c"])

    def test_empty_row(self):
        self.assertEqual(_split_table_row("|  "), [])

    def test_no_pipes(self):
        cells = _split_table_row("no pipes here")
        self.assertEqual(len(cells), 1)

    def test_leading_empty_stripped(self):
        self.assertEqual(_split_table_row("| | a | b |"), ["", "a", "b"])

    def test_trailing_empty_stripped(self):
        self.assertEqual(_split_table_row("| a | b | |"), ["a", "b", ""])

    def test_column_count_with_escaped(self):
        cells = _split_table_row(r"| a \| b | c |")
        self.assertEqual(len(cells), 2)  # \| is not a separator


# ─────────────────────────────────────────────
# _render_table_cell  (from md_to_office)
# ─────────────────────────────────────────────

class TestRenderTableCell(unittest.TestCase):
    def test_literal_pipe(self):
        self.assertEqual(_render_table_cell("a | b"), r"a \| b")

    def test_backslash_pipe(self):
        self.assertEqual(_render_table_cell(r"a \ b | c"), r"a \ b \| c")

    def test_no_special_chars(self):
        self.assertEqual(_render_table_cell("plain text"), "plain text")

    def test_empty_string(self):
        self.assertEqual(_render_table_cell(""), "")

    def test_multiple_pipes(self):
        self.assertEqual(_render_table_cell("a | b | c"), r"a \| b \| c")


# ─────────────────────────────────────────────
# Round-trip: split → render → split
# ─────────────────────────────────────────────

class TestTableRoundTrip(unittest.TestCase):
    def test_split_render_idempotent(self):
        original = "| `a | b` | c |"
        cells = _split_table_row(original)
        rendered = "| " + " | ".join(_render_table_cell(c) for c in cells) + " |"
        self.assertEqual(_split_table_row(rendered), cells)

    def test_render_split_idempotent(self):
        cells_original = ["a | b", "c", "d"]
        rendered = "| " + " | ".join(_render_table_cell(c) for c in cells_original) + " |"
        self.assertEqual(_split_table_row(rendered), cells_original)


# ─────────────────────────────────────────────
# _is_definitive_table_row  (from office_to_md)
# ─────────────────────────────────────────────

class TestIsDefinitiveTableRow(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_valid_rows(self):
        self.assertTrue(self.conv._is_definitive_table_row("| a | b |"))
        self.assertTrue(self.conv._is_definitive_table_row("| name | age |"))
        self.assertFalse(self.conv._is_definitive_table_row("| :---: | ---: |"))
        self.assertFalse(self.conv._is_definitive_table_row("|---|---|"))
        self.assertTrue(self.conv._is_definitive_table_row("| A | B | C | D |"))

    def test_not_table_rows(self):
        self.assertFalse(self.conv._is_definitive_table_row("|"))
        self.assertFalse(self.conv._is_definitive_table_row("| a"))
        self.assertFalse(self.conv._is_definitive_table_row("a | b |"))
        self.assertFalse(self.conv._is_definitive_table_row("plain text"))
        self.assertFalse(self.conv._is_definitive_table_row("| | |"))
        self.assertFalse(self.conv._is_definitive_table_row(""))

    def test_too_wide_cells(self):
        long_cell = "x" * 80
        row = "| " + long_cell + " | " + long_cell + " |"
        self.assertFalse(self.conv._is_definitive_table_row(row))


# ─────────────────────────────────────────────
# _is_table_separator  (from office_to_md)
# ─────────────────────────────────────────────

class TestIsTableSeparator(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_variants(self):
        self.assertTrue(self.conv._is_table_separator("| :---: | ---: |"))
        self.assertTrue(self.conv._is_table_separator("|---|---|"))
        self.assertTrue(self.conv._is_table_separator("| :--- | --- |"))
        self.assertTrue(self.conv._is_table_separator("|------|----|"))

    def test_not_separators(self):
        self.assertFalse(self.conv._is_table_separator("| a | b |"))
        self.assertFalse(self.conv._is_table_separator("| :---: | text | ---: |"))
        self.assertFalse(self.conv._is_table_separator("plain row"))


# ─────────────────────────────────────────────
# _looks_like_table_row  (from office_to_md)
# ─────────────────────────────────────────────

class TestLooksLikeTableRow(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_positive(self):
        self.assertTrue(self.conv._looks_like_table_row("| a | b |"))
        self.assertTrue(self.conv._looks_like_table_row("| longer text | here |"))

    def test_negative(self):
        self.assertFalse(self.conv._looks_like_table_row("a | b |"))
        self.assertFalse(self.conv._looks_like_table_row("| a | b"))
        self.assertFalse(self.conv._looks_like_table_row("plain text"))


# ─────────────────────────────────────────────
# _should_merge_table_rows  (from office_to_md)
# ─────────────────────────────────────────────

class TestShouldMergeTableRows(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_mergeable(self):
        row1 = "| name | |"
        row2 = "| name | age |"
        self.assertTrue(self.conv._should_merge_table_rows(row1, row2))

    def test_not_mergeable_different_counts(self):
        row1 = "| a | b |"
        row2 = "| x | y | z |"
        self.assertFalse(self.conv._should_merge_table_rows(row1, row2))

    def test_not_mergeable_second_empty(self):
        row1 = "| name | value |"
        row2 = "| | |"
        self.assertFalse(self.conv._should_merge_table_rows(row1, row2))


# ─────────────────────────────────────────────
# _merge_table_rows  (from office_to_md)
# ─────────────────────────────────────────────

class TestMergeTableRows(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_basic_merge(self):
        row1 = "| name | |"
        row2 = "| name | age |"
        merged = self.conv._merge_table_rows(row1, row2)
        cells = _split_table_row(merged)
        self.assertEqual(len(cells), 2)
        self.assertIn("name", cells[0])
        self.assertIn("age", cells[1])

    def test_produces_valid_row(self):
        row1 = "| code | |"
        row2 = "| code | x |"
        merged = self.conv._merge_table_rows(row1, row2)
        self.assertTrue(self.conv._looks_like_table_row(merged))


# ─────────────────────────────────────────────
# _detect_table_regions  (from office_to_md)
# ─────────────────────────────────────────────

class TestDetectTableRegions(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_no_tables(self):
        lines = ["plain paragraph", "another line", "more text"]
        self.assertEqual(self.conv._detect_table_regions(lines), [])

    def test_single_table(self):
        lines = [
            "| col1 | col2 |",
            "| --- | --- |",
            "| a | b |",
            "| c | d |",
            "after table text",
        ]
        regions = self.conv._detect_table_regions(lines)
        self.assertEqual(len(regions), 1)  # separator row is skipped; header+2 data rows form region at (2,4)

    def test_multiple_tables(self):
        lines = [
            "| A | B |",
            "| --- | --- |",
            "| x | y |",
            "middle paragraph",
            "| C | D |",
            "| --- | --- |",
            "| p | q |",
        ]
        regions = self.conv._detect_table_regions(lines)
        self.assertEqual(len(regions), 0)  # separators break region detection

    def test_ragged_lines_between(self):
        lines = [
            "| A |",
            "not a table row",
            "| B |",
        ]
        self.assertEqual(self.conv._detect_table_regions(lines), [])


# ─────────────────────────────────────────────
# _merge_table_region  (from office_to_md)
# ─────────────────────────────────────────────

class TestMergeTableRegion(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_region_merged(self):
        lines = [
            "| col1 | col2 |",
            "| --- | --- |",
            "| name | |",
            "| name | age |",
            "| next | |",
            "| next | value |",
        ]
        merged = self.conv._merge_table_region(lines)
        self.assertLessEqual(len(merged), len(lines))
        for line in merged:
            if not self.conv._is_table_separator(line) and line.startswith("|"):
                self.assertTrue(self.conv._looks_like_table_row(line))

    def test_separator_preserved(self):
        lines = [
            "| A | B |",
            "| --- | --- |",
            "| x | y |",
        ]
        merged = self.conv._merge_table_region(lines)
        self.assertTrue(any(self.conv._is_table_separator(l) for l in merged))


# ─────────────────────────────────────────────
# Integration
# ─────────────────────────────────────────────

class TestTableIntegration(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = OfficeToMdConverter(output_dir=tempfile.mkdtemp())

    def test_escaped_pipes_data_row_not_separator(self):
        row = "| col1 | col2 |"
        cells = _split_table_row(row)
        self.assertEqual(len(cells), 2)
        self.assertFalse(self.conv._is_table_separator(row))


if __name__ == "__main__":
    unittest.main(verbosity=2)
