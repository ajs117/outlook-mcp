"""Unit tests for sender grouping and brand labelling. Pure, no network:

    .venv/Scripts/python.exe -m unittest tests.test_senders

Regression focus: two subdomains of one company are two senders, and the
grouping must say so rather than quietly merging billing into marketing.
"""
import unittest

from outlook_mcp.senders import brand_for, domain_of, group_senders


def msg(ref, frm, unread=False, date="2026-09-01T00:00:00Z"):
    return {"ref": ref, "from": frm, "unread": unread, "date": date}


class BrandFor(unittest.TestCase):
    def test_strips_esp_subdomain(self):
        self.assertEqual(brand_for("noreply@emails.thewhitecompany.com"), "Thewhitecompany")

    def test_strips_glued_email_prefix(self):
        self.assertEqual(brand_for("news@emaildebenhams.com"), "Debenhams")

    def test_plain_domain(self):
        self.assertEqual(brand_for("jobalerts-noreply@linkedin.com"), "Linkedin")

    def test_two_letter_esp_label_dropped(self):
        self.assertEqual(brand_for("linkedin@em.linkedin.com"), "Linkedin")

    def test_co_uk_tail_stripped(self):
        self.assertEqual(brand_for("noreply@marks.sainsburys.co.uk"), "Sainsburys")

    def test_leading_letter_not_eaten(self):
        self.assertEqual(brand_for("hello@edfenergy.com"), "Edfenergy")

    def test_mail_suffix_not_stripped_from_real_name(self):
        self.assertEqual(brand_for("no-reply@royalmail.com"), "Royalmail")

    def test_no_at_sign_still_labelled(self):
        self.assertEqual(brand_for("bounces.ebay.com"), "Ebay")

    def test_empty(self):
        self.assertEqual(brand_for(""), "unknown")


class DomainOf(unittest.TestCase):
    def test_keeps_subdomain(self):
        self.assertEqual(domain_of("hello@home.edfenergy.com"), "home.edfenergy.com")


class GroupSenders(unittest.TestCase):
    def test_counts_and_read_pct(self):
        rows = [
            msg("m1", "a@x.com", unread=True),
            msg("m2", "a@x.com", unread=False),
            msg("m3", "a@x.com", unread=False),
        ]
        (g,) = group_senders(rows)
        self.assertEqual((g["count"], g["read"], g["unread"], g["read_pct"]), (3, 2, 1, 67))

    def test_subdomains_stay_separate_and_are_flagged(self):
        rows = [
            msg("m1", "hello@edfenergy.com"),
            msg("m2", "hello@home.edfenergy.com"),
        ]
        groups = group_senders(rows)
        self.assertEqual(len(groups), 2)
        self.assertEqual({g["brand"] for g in groups}, {"Edfenergy"})
        self.assertTrue(all(g["brand_shared"] for g in groups))

    def test_unique_brand_not_flagged(self):
        (g,) = group_senders([msg("m1", "a@x.com")])
        self.assertNotIn("brand_shared", g)

    def test_case_insensitive_sender_merge(self):
        groups = group_senders([msg("m1", "A@X.com"), msg("m2", "a@x.com")])
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]["count"], 2)

    def test_sorted_by_count_desc(self):
        rows = [msg("m1", "a@x.com"), msg("m2", "b@y.com"), msg("m3", "b@y.com")]
        self.assertEqual([g["sender"] for g in group_senders(rows)], ["b@y.com", "a@x.com"])

    def test_date_range_tracked(self):
        rows = [
            msg("m1", "a@x.com", date="2026-01-05T00:00:00Z"),
            msg("m2", "a@x.com", date="2026-09-05T00:00:00Z"),
        ]
        (g,) = group_senders(rows)
        self.assertEqual(g["oldest"], "2026-01-05T00:00:00Z")
        self.assertEqual(g["newest"], "2026-09-05T00:00:00Z")

    def test_blank_sender_skipped(self):
        self.assertEqual(group_senders([msg("m1", "")]), [])


if __name__ == "__main__":
    unittest.main()
