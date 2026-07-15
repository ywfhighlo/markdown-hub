"""Tests for pipe table parsing and \\| escape handling.

Covers the core helper `_split_table_row` and the integration via
`MdToOfficeConverter._optimize_table_column_widths`.
"""
import os
import sys
from pathlib import Path

import pytest

# Make backend importable when run from project root
ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.md_to_office import (  # noqa: E402
    _split_table_row,
    _render_table_cell,
    MdToOfficeConverter,
)


# ─────────────────────────────────────────────
# _split_table_row: escape semantics
# ─────────────────────────────────────────────

@pytest.mark.parametrize("row,expected", [
    ("| a | b | c |",               ["a", "b", "c"]),
    ("| a | b |",                   ["a", "b"]),
    (r"| a \| b | c |",             ["a | b", "c"]),
    (r"| a \\| b | c |",            ["a \\", "b", "c"]),
    (r"| a \\\| b | c |",           ["a \\| b", "c"]),
    (r"| `5 \| 3` | long |",        ["`5 | 3`", "long"]),
    ("| :---: | ---: |",            [":---:", "---:"]),
    (r"| --- \| --- | --- |",       ["--- | ---", "---"]),
])
def test_split_table_row(row, expected):
    assert _split_table_row(row) == expected


@pytest.mark.parametrize("row,expected_cols", [
    # For complex backslash sequences, we verify column count (the separator
    # detection correctness), since exact backslash rendering is pandoc's job.
    (r"| a \\\\ | b | c |", 3),       # 4 backslashes + space: still 3 columns
    (r"| a \\\\\\| b |",    2),       # 3 backslashes + |: escaped, so 2 columns
])
def test_split_table_row_column_count(row, expected_cols):
    """Verify that backslash sequences don't corrupt column counting."""
    cells = _split_table_row(row)
    assert len(cells) == expected_cols, f"Expected {expected_cols} cols, got {len(cells)}: {cells}"


def test_split_table_row_single_column():
    assert _split_table_row("| a |") == ["a"]


def test_split_table_row_empty():
    assert _split_table_row("|  ") == []


# ─────────────────────────────────────────────
# _render_table_cell: escape back
# ─────────────────────────────────────────────

def test_render_cell_literal_pipe():
    assert _render_table_cell("a | b") == r"a \| b"


def test_render_cell_no_pipe():
    assert _render_table_cell("plain") == "plain"


def test_round_trip_split_then_render():
    """Splitting then rendering should be idempotent for the pipe character."""
    original = r"| `a \| b` | c |"
    cells = _split_table_row(original)
    rendered = "| " + " | ".join(_render_table_cell(c) for c in cells) + " |"
    # The rendered output should still parse to the same cells
    assert _split_table_row(rendered) == cells


# ─────────────────────────────────────────────
# _optimize_table_column_widths: integration
# ─────────────────────────────────────────────

@pytest.fixture(scope="module")
def converter(tmp_path_factory):
    out_dir = tmp_path_factory.mktemp("out")
    return MdToOfficeConverter(output_dir=str(out_dir))


def test_optimize_preserves_escape(converter):
    """A table with escaped pipe should be optimized and stay 2 columns."""
    md = (
        "| cmd (with \\|) | note |\n"
        "| :---: | ---: |\n"
        "| `a \\| b` | long text |\n"
        "| `x` | y |\n"
    )
    out = converter._optimize_table_column_widths(md)
    # Every row should split into exactly 2 cells
    table_lines = [l for l in out.split('\n') if l.startswith('|') and '---' not in l.replace(':', '')]
    for line in table_lines:
        cells = _split_table_row(line)
        assert len(cells) == 2, f"Expected 2 columns, got {len(cells)}: {line!r}"


def test_optimize_skips_invalid_separator(converter):
    """Separator row containing \\| should skip optimization."""
    md = (
        "| a | c |\n"
        "| --- \\| --- | --- |\n"
        "| x | y |\n"
    )
    out = converter._optimize_table_column_widths(md)
    assert out.strip() == md.strip(), "Table with invalid separator should be returned as-is"


def test_optimize_skips_ragged_rows(converter):
    """Rows with inconsistent column count should skip optimization."""
    md = (
        "| a | b |\n"
        "| --- | --- |\n"
        "| x | y | z |\n"
    )
    out = converter._optimize_table_column_widths(md)
    assert out.strip() == md.strip(), "Ragged table should be returned as-is"


def test_optimize_preserves_alignment_colons(converter):
    """Alignment colons in separator should be preserved."""
    md = (
        "| name | value |\n"
        "| :--- | ---: |\n"
        "| abc | 1 |\n"
        "| def | 22 |\n"
    )
    out = converter._optimize_table_column_widths(md)
    # The separator line should still contain the colons
    sep_line = [l for l in out.split('\n') if '---' in l and ':' in l]
    assert len(sep_line) == 1
    assert ':---' in sep_line[0] and '---:' in sep_line[0]
