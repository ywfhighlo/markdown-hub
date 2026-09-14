"""Tests for dep_check module utility functions.

Covers: lib_available, lib_error, command_available, command_info,
        resolve_command, feature_snapshot, reset_cache, parser_ready, generator_ready
"""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(ROOT))

import backend.converters.dep_check as dc


class TestLibAvailable(unittest.TestCase):
    def setUp(self):
        dc._LIB_CACHE.clear()
        dc._LIB_ERROR_CACHE.clear()

    def test_unknown_lib_not_available(self):
        result = dc.lib_available("i-do-not-exist-abc123")
        self.assertFalse(result)

    def test_lib_error_includes_reason(self):
        result = dc.lib_error("i-do-not-exist-abc123")
        self.assertIsInstance(result, str)
        self.assertTrue(len(result) > 0)  # non-empty = not available


class TestCommandAvailable(unittest.TestCase):
    def setUp(self):
        dc._CMD_PATH_CACHE.clear()
        dc._CMD_INFO_CACHE.clear()

    def tearDown(self):
        dc._CMD_PATH_CACHE.clear()
        dc._CMD_INFO_CACHE.clear()

    def test_python3_available(self):
        self.assertTrue(dc.command_available("python3"))

    def test_nonexistent_command(self):
        result = dc.command_available("this-command-does-not-exist-xyz123")
        self.assertFalse(result)

    def test_command_info_for_missing(self):
        info = dc.command_info("nonexistent-cmd-abc")
        self.assertIsInstance(info, str)
        self.assertTrue(len(info) > 0)


class TestResolveCommand(unittest.TestCase):
    def setUp(self):
        dc._CMD_PATH_CACHE.clear()
        dc._CMD_INFO_CACHE.clear()

    def tearDown(self):
        dc._CMD_PATH_CACHE.clear()
        dc._CMD_INFO_CACHE.clear()

    def test_resolve_python3(self):
        result = dc.resolve_command("python3")
        self.assertIsInstance(result, str)
        self.assertTrue(len(result) > 0)

    def test_resolve_nonexistent_returns_empty(self):
        result = dc.resolve_command("i-am-not-a-command-xyz123")
        self.assertEqual(result, "")


class TestFeatureSnapshot(unittest.TestCase):
    def test_returns_dict(self):
        snap = dc.feature_snapshot()
        self.assertIsInstance(snap, dict)

    def test_values_are_dicts(self):
        snap = dc.feature_snapshot()
        for category, features in snap.items():
            self.assertIsInstance(features, dict,
                               f"Category {category!r} should be a dict")


class TestResetCache(unittest.TestCase):
    def setUp(self):
        dc.reset_cache()

    def test_clears_lib_caches(self):
        dc.lib_available("some-fake-lib-xyz")
        self.assertIn("some-fake-lib-xyz", dc._LIB_CACHE)
        dc.reset_cache()
        self.assertNotIn("some-fake-lib-xyz", dc._LIB_CACHE)

    def test_clears_command_caches(self):
        dc.command_available("some-fake-cmd-xyz")
        self.assertIn("some-fake-cmd-xyz", dc._CMD_PATH_CACHE)
        dc.reset_cache()
        self.assertNotIn("some-fake-cmd-xyz", dc._CMD_PATH_CACHE)


class TestParserReady(unittest.TestCase):
    def test_returns_tuple(self):
        result = dc.parser_ready("pdf")
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 2)

    def test_pdf_parser_known(self):
        ok, deps = dc.parser_ready("pdf")
        self.assertIsInstance(ok, bool)
        self.assertIsInstance(deps, list)

    def test_unknown_format(self):
        ok, deps = dc.parser_ready("nonexistent-format-xyz")
        self.assertFalse(ok)
        self.assertIsInstance(deps, list)


class TestGeneratorReady(unittest.TestCase):
    def test_returns_tuple(self):
        result = dc.generator_ready("docx")
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)  # (ok, inputs, outputs)

    def test_docx_generator_known(self):
        ok, inputs, outputs = dc.generator_ready("docx")
        self.assertIsInstance(ok, bool)
        self.assertIsInstance(inputs, list)
        self.assertIsInstance(outputs, list)

    def test_unknown_format(self):
        ok, inputs, outputs = dc.generator_ready("nonexistent-format-xyz")
        self.assertFalse(ok)


class TestCheckLibInternal(unittest.TestCase):
    def setUp(self):
        dc._LIB_CACHE.clear()
        dc._LIB_ERROR_CACHE.clear()

    def test_check_lib_returns_tuple(self):
        ok, err = dc._check_lib("nonexistent-lib-xyz-abc")
        self.assertIsInstance(ok, bool)
        self.assertIsInstance(err, str)

    def test_check_lib_populates_cache(self):
        dc._check_lib("nonexistent-lib-xyz")
        self.assertIn("nonexistent-lib-xyz", dc._LIB_CACHE)
        self.assertIn("nonexistent-lib-xyz", dc._LIB_ERROR_CACHE)

    def test_cache_hit(self):
        first = dc._check_lib("nonexistent-lib-xyz")
        second = dc._check_lib("nonexistent-lib-xyz")
        self.assertEqual(first, second)
        self.assertIn("nonexistent-lib-xyz", dc._LIB_CACHE)


class TestCheckCommandInternal(unittest.TestCase):
    def setUp(self):
        dc._CMD_PATH_CACHE.clear()
        dc._CMD_INFO_CACHE.clear()

    def tearDown(self):
        dc._CMD_PATH_CACHE.clear()
        dc._CMD_INFO_CACHE.clear()

    def test_check_command_returns_tuple(self):
        ok, info, path = dc._check_command("python3")
        self.assertIsInstance(ok, bool)
        self.assertIsInstance(info, str)
        self.assertIsInstance(path, str)

    def test_check_command_populates_caches(self):
        dc._check_command("python3")
        self.assertIn("python3", dc._CMD_PATH_CACHE)
        self.assertIn("python3", dc._CMD_INFO_CACHE)


if __name__ == "__main__":
    unittest.main(verbosity=2)
