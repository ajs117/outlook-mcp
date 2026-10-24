"""Unit tests for the protected-sender keep-list:

    .venv/Scripts/python.exe -m unittest tests.test_protected

Config I/O is redirected at OUTLOOK_MCP_CONFIG so no real config is touched.
"""
import json
import os
import tempfile
import unittest
from pathlib import Path

from outlook_mcp import protected


class Match(unittest.TestCase):
    def test_glob_matches_domain(self):
        self.assertEqual(
            protected.match("hello@edfenergy.com", ["*@edfenergy.com"]), "*@edfenergy.com"
        )

    def test_glob_does_not_match_subdomain(self):
        self.assertIsNone(protected.match("hello@home.edfenergy.com", ["*@edfenergy.com"]))

    def test_bare_pattern_is_substring(self):
        self.assertEqual(protected.match("no-reply@dwp.gov.uk", ["dwp"]), "dwp")

    def test_case_insensitive(self):
        self.assertEqual(protected.match("Alerts@NHS.net", ["*@nhs.net"]), "*@nhs.net")

    def test_no_patterns_no_match(self):
        self.assertIsNone(protected.match("a@b.com", []))

    def test_blank_address(self):
        self.assertIsNone(protected.match("", ["dwp"]))

    def test_first_matching_pattern_returned(self):
        self.assertEqual(protected.match("a@dwp.gov.uk", ["gov.uk", "dwp"]), "gov.uk")


class Partition(unittest.TestCase):
    def setUp(self):
        self.msgs = [{"s": "a@dwp.gov.uk"}, {"s": "deals@shop.com"}]
        self.sender_of = lambda m: m["s"]

    def test_splits(self):
        keep, held = protected.partition(self.msgs, ["dwp"], self.sender_of)
        self.assertEqual(keep, [{"s": "deals@shop.com"}])
        self.assertEqual(held, [{"s": "a@dwp.gov.uk"}])

    def test_empty_patterns_holds_nothing(self):
        keep, held = protected.partition(self.msgs, [], self.sender_of)
        self.assertEqual(len(keep), 2)
        self.assertEqual(held, [])


class ConfigIO(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "cfg.json"
        os.environ["OUTLOOK_MCP_CONFIG"] = str(self.path)

    def tearDown(self):
        os.environ.pop("OUTLOOK_MCP_CONFIG", None)
        self.tmp.cleanup()

    def test_missing_file_is_empty_list(self):
        self.assertEqual(protected.get_patterns(), [])

    def test_roundtrip_dedupes_and_sorts(self):
        protected.set_patterns(["ZZZ", "aaa", "aaa", "  "])
        self.assertEqual(protected.get_patterns(), ["aaa", "zzz"])

    def test_preserves_other_config_keys(self):
        self.path.write_text(json.dumps({"other": 1}), encoding="utf-8")
        protected.set_patterns(["dwp"])
        self.assertEqual(json.loads(self.path.read_text(encoding="utf-8"))["other"], 1)

    def test_corrupt_file_does_not_raise(self):
        self.path.write_text("{not json", encoding="utf-8")
        self.assertEqual(protected.get_patterns(), [])


if __name__ == "__main__":
    unittest.main()
