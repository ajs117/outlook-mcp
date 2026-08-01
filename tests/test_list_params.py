"""Unit tests for _build_list_params — the OData query-shaping rules for
message listings. Pure (no network/auth), so runnable with plain stdlib:

    .venv/Scripts/python.exe -m unittest tests.test_list_params

Regression focus: Graph rejects `$search` combined with `$skip`/`$orderby`/
`$filter` (HTTP 400). These tests lock that in.
"""
import unittest

from outlook_mcp.graph_client import _build_list_params, MESSAGE_FIELDS


def build(**over):
    base = dict(
        top=25,
        skip=0,
        filter_query=None,
        order_by="receivedDateTime desc",
        search=None,
        fetch_all=False,
        max_results=500,
        include_preview=False,
    )
    base.update(over)
    return _build_list_params(**base)


class BuildListParams(unittest.TestCase):
    def test_search_omits_skip_orderby_filter(self):
        p = build(search="delivered", filter_query="isRead eq false")
        self.assertEqual(p["$search"], '"delivered"')
        self.assertNotIn("$skip", p)
        self.assertNotIn("$orderby", p)
        self.assertNotIn("$filter", p)

    def test_search_with_nonzero_skip_still_omits_skip(self):
        # Regression: previously $skip was added for every non-fetch_all call,
        # including searches, which made Graph return 400.
        p = build(search="foo", skip=40)
        self.assertNotIn("$skip", p)

    def test_paged_non_search_has_skip_and_orderby(self):
        p = build(skip=10)
        self.assertEqual(p["$skip"], 10)
        self.assertEqual(p["$orderby"], "receivedDateTime desc")

    def test_non_search_filter_passthrough(self):
        p = build(filter_query="isRead eq false")
        self.assertEqual(p["$filter"], "isRead eq false")

    def test_fetch_all_never_sends_skip(self):
        p = build(fetch_all=True, skip=10, max_results=300)
        self.assertNotIn("$skip", p)
        self.assertEqual(p["$top"], 300)

    def test_top_capped_at_999(self):
        self.assertEqual(build(top=5000)["$top"], 999)
        self.assertEqual(build(fetch_all=True, max_results=5000)["$top"], 999)

    def test_include_preview_toggles_bodypreview(self):
        self.assertTrue(build(include_preview=True)["$select"].endswith(",bodyPreview"))
        self.assertEqual(build(include_preview=False)["$select"], MESSAGE_FIELDS)


if __name__ == "__main__":
    unittest.main()
