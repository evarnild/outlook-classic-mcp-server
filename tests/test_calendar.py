"""Unit tests for calendar_client (no Outlook needed): range/DASL construction, paging with ties and recurring
occurrences, query filtering, and create_event behaviour against a fake appointment (never sends).

Run from the project folder:  .venv/Scripts/python.exe -m unittest tests.test_calendar -v
"""
import random
import sys
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import calendar_client as cal  # noqa: E402
import outlook_client as oc  # noqa: E402


class FakeEvent:
    Class = 26

    def __init__(self, entry_id, start, subject="Meeting", location="", organizer="boss@example.com",
                 minutes=30):
        self.EntryID, self.Start, self.Subject = entry_id, start, subject
        self.End = start + timedelta(minutes=minutes)
        self.Location, self.Organizer = location, organizer


class FakeItems:
    """Sort/IncludeRecurrences/Restrict/GetFirst/GetNext. Restrict returns everything (the Python cursor and
    filters do the exact work); ties come back in random order, like Outlook."""

    def __init__(self, items, log):
        self.items, self.log, self.IncludeRecurrences = items, log, False

    def Sort(self, *a):
        self.log.append(("sort", a))

    def Restrict(self, f):
        self.log.append(("restrict", f, self.IncludeRecurrences))
        return self

    def GetFirst(self):
        items = sorted(self.items, key=lambda i: i.Start)
        out, k = [], 0
        while k < len(items):
            g = [i for i in items if i.Start == items[k].Start]
            random.shuffle(g)
            out += g
            k += len(g)
        self.it = iter(out)
        return next(self.it, None)

    def GetNext(self):
        return next(self.it, None)


class FakeFolder:
    def __init__(self, items):
        self.log = []
        self.Items = FakeItems(items, self.log)


def t(h, m=0, day=1):
    return datetime(2026, 10, day, h, m)


class SummaryTests(unittest.TestCase):
    def test_occurrence_without_folder_parent_still_summarises(self):
        class Occurrence(FakeEvent):
            @property
            def Parent(self):
                raise AttributeError("occurrence parent is the series, not a folder")
        ev = Occurrence("SERIES", t(9), subject="Weekly")
        ev.AllDayEvent, ev.IsRecurring, ev.MeetingStatus, ev.BusyStatus, ev.ResponseStatus = False, True, 0, 2, 0
        out = cal._summary(ev, "STORE")
        self.assertEqual((out["store_id"], out["is_recurring"], out["subject"]), ("STORE", True, "Weekly"))
        self.assertEqual(out["start_local"], "2026-10-01T09:00:00")


class OccurrenceTests(unittest.TestCase):
    def test_finds_the_occurrence_including_a_moved_one(self):
        series = [FakeEvent("SERIES", t(14, day=d)) for d in (1, 8)] + [FakeEvent("SERIES", t(14, 30, day=15))]  # moved
        folder = FakeFolder(series + [FakeEvent("OTHER", t(14, 30, day=15))])
        found = cal._find_occurrence(folder, "SERIES", t(14, 30, day=15).astimezone(timezone.utc))
        self.assertIsNotNone(found)
        self.assertEqual((found.EntryID, found.Start), ("SERIES", t(14, 30, day=15)))

    def test_unknown_start_returns_none(self):
        folder = FakeFolder([FakeEvent("SERIES", t(14, day=1))])
        self.assertIsNone(cal._find_occurrence(folder, "SERIES", t(14, day=2).astimezone(timezone.utc)))


class RangeTests(unittest.TestCase):
    def test_default_window_is_today_for_seven_days(self):
        s, e = cal._window(None, None)
        self.assertEqual(e - s, timedelta(days=7))
        self.assertEqual(s.astimezone().replace(tzinfo=None).time(), datetime.min.time())

    def test_explicit_range_converted_to_utc(self):
        s, e = cal._window("2026-10-07T09:00:00+02:00", "2026-10-07T18:00:00+02:00")
        self.assertEqual((s, e), (datetime(2026, 10, 7, 7, tzinfo=timezone.utc), datetime(2026, 10, 7, 16, tzinfo=timezone.utc)))

    def test_invalid_ranges_rejected(self):
        with self.assertRaises(oc.OutlookError):
            cal._window("2026-10-07", "2026-10-06")
        with self.assertRaises(oc.OutlookError):
            cal._window("2026-01-01", "2028-01-01")

    def test_dasl_overlap_filter_in_utc(self):
        s, e = cal._window("2026-10-07T00:00:00Z", "2026-10-08T00:00:00Z")
        self.assertEqual(cal._range_conditions(s, e), [
            "\"urn:schemas:calendar:dtstart\" < '2026-10-08 00:00:00'",
            "\"urn:schemas:calendar:dtend\" > '2026-10-07 00:00:00'"])

    def test_cursor(self):
        self.assertIsNone(cal._cursor(None, None))
        c = cal._cursor("2026-10-07T09:00:00Z", None)
        self.assertEqual(c[0], datetime(2026, 10, 7, 9, tzinfo=timezone.utc))
        self.assertGreater(c[1], "FFFFFFFF")  # excludes everything at that second
        with self.assertRaises(oc.OutlookError):
            cal._cursor(None, "ABC")


class PagingTests(unittest.TestCase):
    def setUp(self):
        self._summary = cal._summary
        cal._summary = lambda item, store_id: {"entry_id": item.EntryID, "start": oc._iso(item.Start), "subject": item.Subject}

    def tearDown(self):
        cal._summary = self._summary

    def folder(self, n=40, group=6):
        random.seed(2)
        items = [FakeEvent(f"E{k:03d}", t(8) + timedelta(minutes=30 * (k // group))) for k in range(n)]
        random.shuffle(items)
        return FakeFolder(items), n

    def walk(self, folder, limit, words=()):
        ids, after, after_id, pages = [], None, None, 0
        while True:
            ranked = cal._ranked(folder, ["x"], limit, cal._cursor(after, after_id), words)
            page = cal._page(ranked, limit, "S")
            ids += [e["entry_id"] for e in page["events"]]
            pages += 1
            self.assertLess(pages, 200)
            if not page["has_more"]:
                return ids
            after, after_id = page["next_after"], page["next_after_entry_id"]

    def test_pages_cover_everything_once_with_ties_on_boundaries(self):
        for limit in (1, 2, 5, 6, 7, 13, 40, 100):
            folder, n = self.folder()
            ids = self.walk(folder, limit)
            self.assertEqual(len(ids), n, f"limit={limit}")
            self.assertEqual(len(set(ids)), n, f"duplicates at limit={limit}")

    def test_earliest_first_then_entry_id(self):
        folder, _ = self.folder()
        ids = self.walk(folder, 7)
        self.assertEqual(ids, sorted(ids))

    def test_recurring_occurrences_sharing_an_entry_id_are_all_returned(self):
        series = [FakeEvent("SERIES", t(9, day=d), subject="Weekly sync") for d in range(1, 8)]
        folder = FakeFolder(series + [FakeEvent("OTHER", t(9, day=3))])
        ids = self.walk(folder, 2)
        self.assertEqual(ids.count("SERIES"), 7)
        self.assertEqual(len(ids), 8)

    def test_query_filters_title_location_organizer_with_all_words(self):
        items = [FakeEvent("A", t(9), subject="Dentist Camille", location="Paris"),
                 FakeEvent("B", t(10), subject="Team sync", location="Room 4", organizer="anna@example.com"),
                 FakeEvent("C", t(11), subject="Dentist", location="Berlin")]
        folder = FakeFolder(items)
        self.assertEqual(self.walk(folder, 10, ["dentist"]), ["A", "C"])
        self.assertEqual(self.walk(folder, 10, ["dentist", "paris"]), ["A"])
        self.assertEqual(self.walk(folder, 10, ["anna"]), ["B"])
        self.assertEqual(self.walk(folder, 10, ["nothing"]), [])

    def test_recurrences_are_expanded_after_sort_and_before_restrict(self):
        folder, _ = self.folder(n=3)
        cal._ranked(folder, ["x"], 5)
        self.assertEqual([e[0] for e in folder.log], ["sort", "restrict"])
        self.assertTrue(folder.log[1][2])  # IncludeRecurrences was already true when Restrict ran

    def test_empty_calendar(self):
        page = cal._page(cal._ranked(FakeFolder([]), ["x"], 5), 5, "S")
        self.assertEqual((page["events"], page["has_more"], page["next_after"]), ([], False, None))


class FakeRecipient:
    def __init__(self, addr):
        self.Address, self.Type = addr, 0


class FakeRecipients(list):
    def Add(self, addr):
        r = FakeRecipient(addr)
        self.append(r)
        return r

    def ResolveAll(self):
        self.resolved = True


class FakeAppointment:
    EntryID = "NEW"
    MeetingStatus = 0
    ReminderSet = False
    Class = 26

    class Parent:
        StoreID = "STORE"
        FolderPath = "\\\\me@example.com\\Calendar"

    def __init__(self):
        self.Recipients = FakeRecipients()
        self.saved = False

    @property
    def Start(self):  # what Outlook would report back: local wall clock of StartUTC
        return self.StartUTC.astimezone().replace(tzinfo=None)

    @property
    def End(self):
        return self.EndUTC.astimezone().replace(tzinfo=None)

    def __setattr__(self, name, value):
        if name in ("Start", "End"):
            raise AssertionError("create_event must set StartUTC/EndUTC, not Start/End (DST bug in Outlook's OM)")
        if name in ("StartUTC", "EndUTC"):
            assert value.tzinfo is not None, "StartUTC/EndUTC must be timezone-aware"
        object.__setattr__(self, name, value)

    def Save(self):
        self.saved = True

    def Send(self):
        raise AssertionError("create_event must never send")

    def Delete(self):
        raise AssertionError("create_event must never delete")


class FakeCalFolder:
    StoreID = "STORE"

    def __init__(self):
        self.created = []

        class Items:
            def Add(inner, kind):
                assert kind == 1  # olAppointmentItem
                item = FakeAppointment()
                self.created.append(item)
                return item
        self.Items = Items()


class CreateEventTests(unittest.TestCase):
    def setUp(self):
        self.folder = FakeCalFolder()
        self._cal, self._sum = cal._calendar, cal._summary
        cal._calendar = lambda account: (object(), self.folder)
        cal._summary = lambda item, store_id: {"entry_id": item.EntryID, "start_local": item.Start.isoformat(),
                                               "end_local": item.End.isoformat()}
        self._fp = oc._folder_display_path
        oc._folder_display_path = lambda f: "me@example.com/Calendar"

    def tearDown(self):
        cal._calendar, cal._summary, oc._folder_display_path = self._cal, self._sum, self._fp

    def test_local_naive_start_and_default_duration(self):
        out = cal.create_event("Lunch", "2026-10-07T12:00")
        item = self.folder.created[0]
        self.assertEqual((item.Start, item.End), (datetime(2026, 10, 7, 12), datetime(2026, 10, 7, 13)))
        self.assertEqual(item.StartUTC, datetime(2026, 10, 7, 12).astimezone(timezone.utc))
        self.assertTrue(item.saved)
        self.assertEqual(out["message"], cal.EVENT_NOTE_PLAIN)
        self.assertEqual(item.MeetingStatus, 0)

    def test_offset_start_is_converted_to_local(self):
        cal.create_event("Call", "2026-10-07T10:00:00+00:00", duration_minutes=45)
        item = self.folder.created[0]
        expected = datetime(2026, 10, 7, 10, tzinfo=timezone.utc).astimezone().replace(tzinfo=None)
        self.assertEqual((item.Start, item.End), (expected, expected + timedelta(minutes=45)))
        self.assertEqual(item.StartUTC, datetime(2026, 10, 7, 10, tzinfo=timezone.utc))

    def test_explicit_end_wins_over_duration(self):
        cal.create_event("Workshop", "2026-10-07T09:00", end="2026-10-07T17:00", duration_minutes=5)
        item = self.folder.created[0]
        self.assertEqual(item.End, datetime(2026, 10, 7, 17))

    def test_all_day_end_is_exclusive_midnight(self):
        cal.create_event("Holiday", "2026-10-07", end="2026-10-09", all_day=True)
        item = self.folder.created[0]
        self.assertTrue(item.AllDayEvent)
        self.assertEqual((item.Start, item.End), (datetime(2026, 10, 7), datetime(2026, 10, 10)))
        self.assertEqual(item.StartUTC, datetime(2026, 10, 7).astimezone(timezone.utc))

    def test_all_day_uses_each_dates_own_utc_offset(self):
        """Midnight local in winter and in summer map to different UTC offsets; each must use its own date's."""
        cal.create_event("Winter holiday", "2027-01-06", end="2027-01-07", all_day=True)
        item = self.folder.created[0]
        self.assertEqual(item.StartUTC, datetime(2027, 1, 6).astimezone(timezone.utc))
        self.assertEqual(item.EndUTC, datetime(2027, 1, 8).astimezone(timezone.utc))
        self.assertEqual((item.Start, item.End), (datetime(2027, 1, 6), datetime(2027, 1, 8)))

    def test_timed_event_in_another_dst_period_keeps_its_wall_clock(self):
        cal.create_event("Winter call", "2027-01-04T03:00")
        self.assertEqual(self.folder.created[0].Start, datetime(2027, 1, 4, 3))
        cal.create_event("Summer call", "2026-07-04T03:00")
        self.assertEqual(self.folder.created[1].Start, datetime(2026, 7, 4, 3))

    def test_single_all_day_event(self):
        cal.create_event("Birthday", "2026-10-07", all_day=True)
        item = self.folder.created[0]
        self.assertEqual((item.Start, item.End), (datetime(2026, 10, 7), datetime(2026, 10, 8)))

    def test_attendees_make_an_unsent_meeting(self):
        out = cal.create_event("Sync", "2026-10-07T09:00", attendees=["a@example.com", " b@example.com "])
        item = self.folder.created[0]
        self.assertEqual(item.MeetingStatus, 1)
        self.assertEqual([(r.Address, r.Type) for r in item.Recipients],
                         [("a@example.com", 1), ("b@example.com", 1)])
        self.assertTrue(item.Recipients.resolved)
        self.assertEqual(out["message"], cal.EVENT_NOTE_MEETING)  # Send() on the fake would raise

    def test_optional_fields(self):
        cal.create_event("Dentist", "2026-10-07T09:00", location="Paris", body="Bring card", reminder_minutes=30,
                         busy_status="oof")
        item = self.folder.created[0]
        self.assertEqual((item.Location, item.Body, item.ReminderSet, item.ReminderMinutesBeforeStart,
                          item.BusyStatus), ("Paris", "Bring card", True, 30, 3))

    def test_validation_errors_create_nothing(self):
        bad = [dict(subject="  ", start="2026-10-07T09:00"),
               dict(subject="x", start="2026-10-07T09:00", end="2026-10-07T08:00"),
               dict(subject="x", start="2026-10-07T09:00", duration_minutes=0),
               dict(subject="x", start="2026-10-09", end="2026-10-07", all_day=True),
               dict(subject="x", start="2026-10-07T09:00", attendees=["not an address"]),
               dict(subject="x", start="2026-10-07T09:00", busy_status="away"),
               dict(subject="x", start="soon")]
        for kw in bad:
            with self.assertRaises(oc.OutlookError, msg=str(kw)):
                cal.create_event(**kw)
        self.assertEqual(self.folder.created, [])


if __name__ == "__main__":
    unittest.main()
