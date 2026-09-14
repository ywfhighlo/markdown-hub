"""Tests for MdToOfficeConverter utility functions.

Covers: _highlight_style_args, _get_title_from_md, _parse_align,
        _bbox_inside_rect, _strip_html_to_md, _inline_to_runs
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

from backend.converters.md_to_office import (
    _split_table_row,
    _render_table_cell,
    MdToOfficeConverter,
)


# ─────────────────────────────────────────
# _highlight_style_args
# ─────────────────────────────────────────

class TestHighlightStyleArgs(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = MdToOfficeConverter(output_dir=tempfile.mkdtemp())

    def test_off_variants(self):
        for val in ('', 'off', 'none', 'disable'):
            self.conv.code_highlight_theme = val
            self.assertEqual(self.conv._highlight_style_args(), ['--no-highlight'],
                           f"Failed for {val!r}")

    def test_builtin_themes(self):
        for theme in ('pygments', 'tango', 'espresso', 'zenburn',
                      'kate', 'monochrome', 'breezedark', 'haddock'):
            self.conv.code_highlight_theme = theme
            self.assertEqual(self.conv._highlight_style_args(),
                           [f'--highlight-style={theme}'],
                           f"Failed for {theme!r}")

    def test_unknown_theme_fallback(self):
        self.conv.code_highlight_theme = 'unknown-nonsense-theme'
        self.assertEqual(self.conv._highlight_style_args(), ['--highlight-style=pygments'])

    def test_none_theme(self):
        self.conv.code_highlight_theme = None
        self.assertEqual(self.conv._highlight_style_args(), ['--no-highlight'])


# ─────────────────────────────────────────
# _get_title_from_md
# ─────────────────────────────────────────

class TestGetTitleFromMd(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = MdToOfficeConverter(output_dir=tempfile.mkdtemp())

    def test_yaml_title(self):
        content = "---\ntitle: My Document Title\n---\n# Intro"
        title = self.conv._get_title_from_md(content, Path("/tmp/test.md"))
        self.assertEqual(title, "My Document Title")

    def test_h1_heading(self):
        content = "# This Is H1\n\nSome text"
        title = self.conv._get_title_from_md(content, Path("/tmp/test.md"))
        self.assertEqual(title, "This Is H1")

    def test_yaml_title_over_h1(self):
        content = "---\ntitle: YAML Title\n---\n# H1 Title"
        title = self.conv._get_title_from_md(content, Path("/tmp/test.md"))
        self.assertEqual(title, "YAML Title")

    def test_no_title_fallback(self):
        content = "Just plain text\n\nNo headings here"
        title = self.conv._get_title_from_md(content, Path("/tmp/my_file.md"))
        self.assertEqual(title, "my_file")

    def test_deep_h1_not_picked(self):
        content = "## This is H2\n# But this is H1"
        title = self.conv._get_title_from_md(content, Path("/tmp/test.md"))
        self.assertEqual(title, "But this is H1")


# ─────────────────────────────────────────
# _parse_align
# ─────────────────────────────────────────

class TestParseAlign(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = MdToOfficeConverter(output_dir=tempfile.mkdtemp())

    def test_center(self):
        self.assertEqual(self.conv._parse_align("text-align:center"), "center")
        self.assertEqual(self.conv._parse_align("text-align: center"), "center")
        self.assertEqual(self.conv._parse_align("align:center;"), "center")

    def test_right(self):
        self.assertEqual(self.conv._parse_align("text-align:right"), "right")
        self.assertEqual(self.conv._parse_align("align=right"), "right")

    def test_left_default(self):
        self.assertEqual(self.conv._parse_align(""), "left")
        self.assertEqual(self.conv._parse_align("text-align:left"), "left")
        self.assertEqual(self.conv._parse_align("normal"), "left")
        self.assertEqual(self.conv._parse_align("inherit"), "left")


# ─────────────────────────────────────────
# _bbox_inside_rect
# ─────────────────────────────────────────

class TestBboxInsideRect(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = MdToOfficeConverter(output_dir=tempfile.mkdtemp())

    def test_fully_inside(self):
        bbox = (10, 10, 50, 50)
        rect = (0, 0, 100, 100)
        self.assertTrue(self.conv._bbox_inside_rect(bbox, rect))

    def test_on_edge(self):
        bbox = (0, 0, 100, 100)
        rect = (0, 0, 100, 100)
        self.assertTrue(self.conv._bbox_inside_rect(bbox, rect))

    def test_partial_overlap(self):
        bbox = (90, 90, 110, 110)
        rect = (0, 0, 100, 100)
        self.assertFalse(self.conv._bbox_inside_rect(bbox, rect))

    def test_outside(self):
        bbox = (200, 200, 300, 300)
        rect = (0, 0, 100, 100)
        self.assertFalse(self.conv._bbox_inside_rect(bbox, rect))

    def test_with_margin(self):
        bbox = (99, 99, 101, 101)
        rect = (0, 0, 100, 100)
        self.assertTrue(self.conv._bbox_inside_rect(bbox, rect, margin=2.0))


# ─────────────────────────────────────────
# _strip_html_to_md
# ─────────────────────────────────────────

class TestStripHtmlToMd(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = MdToOfficeConverter(output_dir=tempfile.mkdtemp())

    def test_simple_tags(self):
        html = "<p>Hello</p><p>World</p>"
        md = self.conv._strip_html_to_md(html)
        self.assertIn("Hello", md)
        self.assertIn("World", md)

    def test_strong_bold(self):
        html = "<strong>bold text</strong>"
        md = self.conv._strip_html_to_md(html)
        self.assertIn("**bold text**", md)

    def test_em_italic(self):
        html = "<em>italic text</em>"
        md = self.conv._strip_html_to_md(html)
        self.assertIn("*italic text*", md)

    def test_code(self):
        html = "<code>inline code</code>"
        md = self.conv._strip_html_to_md(html)
        self.assertIn("`inline code`", md)

    def test_link(self):
        html = '<a href="https://example.com">Example</a>'
        md = self.conv._strip_html_to_md(html)
        self.assertIn("[Example](https://example.com)", md)

    def test_image(self):
        html = '<img src="img.png" alt="alt text" />'
        md = self.conv._strip_html_to_md(html)
        self.assertIn("![alt text]", md)
        self.assertIn("img.png", md)

    def test_br_converted(self):
        html = "Line1<br>Line2"
        md = self.conv._strip_html_to_md(html)
        self.assertIn("\n", md)

    def test_html_entities(self):
        html = "&lt;script&gt; &amp; &quot;test&quot;"
        md = self.conv._strip_html_to_md(html)
        self.assertIn("<script>", md)
        self.assertIn("&", md)


# ─────────────────────────────────────────
# _inline_to_runs
# ─────────────────────────────────────────

class TestInlineToRuns(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = MdToOfficeConverter(output_dir=tempfile.mkdtemp())

    def test_single_run(self):
        runs = self.conv._inline_to_runs("plain text")
        self.assertEqual(len(runs), 1)
        self.assertEqual(runs[0]['text'], "plain text")
        self.assertFalse(runs[0]['bold'])
        self.assertFalse(runs[0]['italic'])
        self.assertFalse(runs[0]['code'])
        self.assertIsNone(runs[0]['link'])

    def test_unicode_text(self):
        runs = self.conv._inline_to_runs("Chinese test")
        self.assertEqual(len(runs), 1)
        self.assertEqual(runs[0]['text'], "Chinese test")


# ─────────────────────────────────────────
# _strip_empty_runs
# ─────────────────────────────────────────

class TestStripEmptyRuns(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = MdToOfficeConverter(output_dir=tempfile.mkdtemp())

    def test_removes_empty(self):
        runs = [
            {'text': 'hello', 'bold': False, 'italic': False, 'code': False, 'link': None},
            {'text': '', 'bold': False, 'italic': False, 'code': False, 'link': None},
            {'text': 'world', 'bold': False, 'italic': False, 'code': False, 'link': None},
        ]
        cleaned = self.conv._strip_empty_runs(runs)
        self.assertEqual(len(cleaned), 2)
        self.assertEqual(cleaned[0]['text'], 'hello')
        self.assertEqual(cleaned[1]['text'], 'world')

    def test_all_empty(self):
        runs = [{'text': '', 'bold': False, 'italic': False, 'code': False, 'link': None}]
        cleaned = self.conv._strip_empty_runs(runs)
        self.assertEqual(cleaned, [])


# ─────────────────────────────────────────
# _annotate_runs
# ─────────────────────────────────────────

class TestAnnotateRuns(unittest.TestCase):
    def setUp(self):
        import tempfile
        self.conv = MdToOfficeConverter(output_dir=tempfile.mkdtemp())

    def test_adds_bold(self):
        runs = [{'text': 'hello', 'bold': False, 'italic': False, 'code': False, 'link': None}]
        annotated = self.conv._annotate_runs(runs, bold=True)
        self.assertTrue(annotated[0]['bold'])

    def test_preserves_other_flags(self):
        runs = [{'text': 'hello', 'bold': True, 'italic': True, 'code': False, 'link': None}]
        annotated = self.conv._annotate_runs(runs, code=True)
        self.assertTrue(annotated[0]['bold'])
        self.assertTrue(annotated[0]['italic'])
        self.assertTrue(annotated[0]['code'])


# ─────────────────────────────────────────
# _render_table_cell round-trip
# ─────────────────────────────────────────

class TestTableCellEscape(unittest.TestCase):
    def test_pipe_escaped(self):
        self.assertEqual(_render_table_cell("a | b"), r"a \| b")

    def test_backslash_preserved(self):
        self.assertEqual(_render_table_cell("a \ b"), r"a \ b")

    def test_empty_cell(self):
        self.assertEqual(_render_table_cell(""), "")


# ─────────────────────────────────────────
# _split_table_row completeness
# ─────────────────────────────────────────

class TestSplitTableRowCompleteness(unittest.TestCase):
    def test_basic_cases(self):
        cases = [
            ("| A | B |", ["A", "B"]),
            ("| A | B | C |", ["A", "B", "C"]),
            ("| :---: | ---: |", [":---:", "---:"]),
            ("| --- | --- |", ["---", "---"]),
        ]
        for row, expected in cases:
            result = _split_table_row(row)
            self.assertEqual(result, expected, f"Failed for {row!r}")


if __name__ == "__main__":
    unittest.main(verbosity=2)
