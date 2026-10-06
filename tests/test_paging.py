"""Unit tests for paging (no Outlook needed): DASL filter construction, timestamp handling and tie-boundary paging.

Run from the project folder:  .venv/Scripts/python.exe -m unittest tests.test_paging -v
"""
import random
import sys
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import outlook_client as oc  # noqa: E402


class FakeItem:
    Class = 43

    def __init__(self, entry_id, when):
        self.EntryID = entry_id
        self.ReceivedTime = when  # local wall-clock, like the COM value


class FakeItems:
    """Emulates Restrict/Sort/GetFirst/GetNext. Restrict returns everything (a valid superset: the Python-side
    cursor does the exact filtering) and ties are shuffled on every call, like Outlook's arbitrary order."""

    def __init__(self, items, log):
        self.items, self.log = items, log

    def Restrict(self, f):
        self.log.append(f)
        return self

    def Sort(self, *_):
        pass

    def _ordered(self):
        items = sorted(self.items, key=lambda i: i.ReceivedTime, reverse=True)
        out, k = [], 0
        while k < len(items):  # shuffle inside each equal-timestamp group
            g = [i for i in items if i.ReceivedTime == items[k].ReceivedTime]
            random.shuffle(g)
            out += g
            k += len(g)
        return out

    def GetFirst(self):
        self.it = iter(self._ordered())
        return next(self.it, None)

    def GetNext(self):
        return next(self.it, None)


class FakeFolder:
    def __init__(self, items):
        self.log = []
        self.Items = FakeItems(items, self.log)


def local(y, mo, d, h=0, mi=0, s=0):
    return datetime(y, mo, d, h, mi, s)


class FakeFolderRef:
    FolderPath = r"\\me@example.com\Inbox"
    StoreID = "STORE"


class MeetingRequest(FakeItem):
    Class = 53
    MessageClass = "IPM.Schedule.Meeting.Request"
    Subject = "Invitation: sync"
    SenderName = "Someone"
    UnRead = True
    Parent = FakeFolderRef()


class Base(unittest.TestCase):
    def setUp(self):
        self._summary = oc._summary
        oc._summary = lambda item: {"entry_id": item.EntryID, "received": oc._iso(item.ReceivedTime)}

    def tearDown(self):
        oc._summary = self._summary


class TimestampTests(unittest.TestCase):
    def test_aware_is_converted_to_utc(self):
        self.assertEqual(oc._parse_ts("2026-09-30T10:00:00+02:00", "x"), datetime(2026, 9, 30, 8, tzinfo=timezone.utc))
        self.assertEqual(oc._parse_ts("2026-09-30T10:00:00Z", "x"), datetime(2026, 9, 30, 10, tzinfo=timezone.utc))

    def test_naive_is_local_time(self):
        got = oc._parse_ts("2026-09-30T10:00:00", "x")
        self.assertEqual(got, datetime(2026, 9, 30, 10).astimezone(timezone.utc))

    def test_date_only(self):
        self.assertEqual(oc._parse_ts("2026-09-30", "x"), datetime(2026, 9, 30).astimezone(timezone.utc))

    def test_invalid(self):
        with self.assertRaises(oc.OutlookError):
            oc._parse_ts("yesterday", "before")

    def test_dasl_literal_is_utc(self):
        self.assertEqual(oc._dasl_ts(oc._parse_ts("2026-09-30T10:00:00+02:00", "x")), "2026-09-30 08:00:00")

    def test_iso_output_roundtrips_into_before(self):
        """next_before (what we output) must parse back to the same instant."""
        t = local(2026, 10, 1, 13, 37, 36)
        out = oc._iso(t)
        self.assertEqual(oc._parse_ts(out, "before"), t.astimezone(timezone.utc))
        self.assertTrue(out.endswith("+00:00"))


class FilterTests(unittest.TestCase):
    def test_no_params(self):
        self.assertEqual(oc._range_conditions(), ([], None))

    def test_all_params_combine(self):
        conds, cursor = oc._range_conditions(True, "2026-09-01T00:00:00Z", "2026-09-30T12:00:00Z")
        self.assertEqual(conds, ['"urn:schemas:httpmail:read" = 0',
                                 "\"urn:schemas:httpmail:datereceived\" >= '2026-09-01 00:00:00'",
                                 "\"urn:schemas:httpmail:datereceived\" < '2026-09-30 12:00:00'"])
        self.assertEqual(cursor, (datetime(2026, 9, 30, 12, tzinfo=timezone.utc), ""))

    def test_entry_id_widens_the_filter_by_one_second(self):
        conds, cursor = oc._range_conditions(False, None, "2026-09-30T12:00:00Z", "ABC")
        self.assertEqual(conds, ["\"urn:schemas:httpmail:datereceived\" < '2026-09-30 12:00:01'"])
        self.assertEqual(cursor[1], "ABC")

    def test_entry_id_without_before_is_rejected(self):
        with self.assertRaises(oc.OutlookError):
            oc._range_conditions(before_entry_id="ABC")

    def test_text_filter_still_built_independently(self):
        self.assertIn("ci_startswith", oc._text_condition("foo", False)[0])


class PagingTests(Base):
    def make_folder(self, n=57, group=7, seed=1):
        random.seed(seed)
        base = local(2026, 9, 1, 8)
        items = [FakeItem(f"E{k:03d}", base + timedelta(minutes=k // group)) for k in range(n)]  # ties of `group`
        random.shuffle(items)
        return FakeFolder(items), n

    def walk(self, folder, limit):
        ids, before, before_id, pages = [], None, None, 0
        while True:
            conds, cursor = oc._range_conditions(False, None, before, before_id)
            page = oc._page(oc._ranked(folder, conds, limit, cursor), limit)
            ids += [m["entry_id"] for m in page["messages"]]
            pages += 1
            self.assertLess(pages, 200)
            if not page["has_more"]:
                return ids, pages
            before, before_id = page["next_before"], page["next_before_entry_id"]

    def test_pages_cover_everything_exactly_once_with_ties_on_boundaries(self):
        for limit in (1, 3, 4, 7, 10, 14, 57, 100):
            folder, n = self.make_folder()
            ids, _ = self.walk(folder, limit)
            self.assertEqual(len(ids), n, f"limit={limit}")
            self.assertEqual(len(set(ids)), n, f"duplicates at limit={limit}")

    def test_order_is_newest_first_then_entry_id_desc(self):
        folder, _ = self.make_folder()
        ids, _ = self.walk(folder, 5)
        self.assertEqual(ids[0], "E056")
        self.assertEqual(ids, sorted(ids, reverse=True))

    def test_has_more_only_when_more_exist(self):
        folder, n = self.make_folder(n=10, group=2)
        conds, cursor = oc._range_conditions()
        self.assertFalse(oc._page(oc._ranked(folder, conds, 10, cursor), 10)["has_more"])
        self.assertTrue(oc._page(oc._ranked(folder, conds, 9, cursor), 9)["has_more"])
        self.assertFalse(oc._page(oc._ranked(FakeFolder([]), conds, 5, cursor), 5)["has_more"])

    def test_non_mail_items_are_paged_too(self):
        base = local(2026, 9, 1, 8)
        folder = FakeFolder([FakeItem(f"M{k}", base + timedelta(minutes=k)) for k in range(3)] +
                            [MeetingRequest("MR", base + timedelta(minutes=1))])
        conds, cursor = oc._range_conditions()
        page = oc._page(oc._ranked(folder, conds, 100, cursor), 100)
        self.assertEqual(sorted(m["entry_id"] for m in page["messages"]), ["M0", "M1", "M2", "MR"])

    def test_non_mail_summary_is_minimal_and_typed(self):
        got = self._summary(MeetingRequest("MR", local(2026, 9, 1, 8)))
        self.assertEqual(got["type"], "IPM.Schedule.Meeting.Request")
        self.assertEqual(sorted(got), ["entry_id", "folder", "received", "sender_name", "store_id", "subject",
                                       "type", "unread"])
        self.assertTrue(got["unread"])

    def test_empty_page_has_no_cursor(self):
        page = oc._page([], 5)
        self.assertEqual((page["messages"], page["has_more"], page["next_before"]), ([], False, None))

    def test_without_entry_id_tied_messages_at_the_boundary_are_dropped(self):
        """Documents why before_entry_id exists: timestamp-only paging is strictly-before."""
        folder, n = self.make_folder(n=14, group=7)
        conds, cursor = oc._range_conditions()
        p1 = oc._page(oc._ranked(folder, conds, 3, cursor), 3)
        conds, cursor = oc._range_conditions(before=p1["next_before"])
        p2 = oc._page(oc._ranked(folder, conds, 100, cursor), 100)
        self.assertLess(len(p1["messages"]) + len(p2["messages"]), n)

    def test_before_filter_reaches_outlook_as_utc_dasl(self):
        folder, _ = self.make_folder(n=3)
        conds, cursor = oc._range_conditions(True, "2026-09-01T00:00:00+00:00", "2026-09-02T00:00:00+00:00")
        oc._ranked(folder, conds, 5, cursor)
        self.assertEqual(folder.log, ['@SQL="urn:schemas:httpmail:read" = 0 AND '
                                      "\"urn:schemas:httpmail:datereceived\" >= '2026-09-01 00:00:00' AND "
                                      "\"urn:schemas:httpmail:datereceived\" < '2026-09-02 00:00:00'"])

    def test_multi_folder_merge_pages_correctly(self):
        random.seed(3)
        base = local(2026, 9, 1, 8)
        a = FakeFolder([FakeItem(f"A{k:02d}", base + timedelta(minutes=k // 3)) for k in range(20)])
        b = FakeFolder([FakeItem(f"B{k:02d}", base + timedelta(minutes=k // 3)) for k in range(20)])
        seen, before, before_id = [], None, None
        for _ in range(50):
            conds, cursor = oc._range_conditions(False, None, before, before_id)
            page = oc._page(oc._ranked(a, conds, 6, cursor) + oc._ranked(b, conds, 6, cursor), 6)
            seen += [m["entry_id"] for m in page["messages"]]
            if not page["has_more"]:
                break
            before, before_id = page["next_before"], page["next_before_entry_id"]
        self.assertEqual(len(seen), 40)
        self.assertEqual(len(set(seen)), 40)


if __name__ == "__main__":
    unittest.main()
