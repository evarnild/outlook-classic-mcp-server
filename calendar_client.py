"""Calendar logic (main calendar of an account) on top of outlook_client. Runs on the COM worker thread.

Reads events (recurring series are expanded into occurrences) and creates new events. It never sends invitations,
and never edits, deletes or moves existing events. Events with attendees are saved as meetings whose invitations
have NOT been sent; the user sends them from Outlook.
"""
import re
from datetime import date, datetime, time, timedelta, timezone

import pywintypes

import outlook_client as oc
from constants import OL_APPOINTMENT_CLASS, OL_APPOINTMENT_ITEM, OL_FOLDER_CALENDAR

MAX_WINDOW_DAYS = 366
DEFAULT_WINDOW_DAYS = 7
MAX_PER_INSTANT = 200  # events that can overlap a single second
BUSY = {"free": 0, "tentative": 1, "busy": 2, "oof": 3, "working_elsewhere": 4}
BUSY_NAMES = {v: k for k, v in BUSY.items()}
RESPONSE_NAMES = {0: "none", 1: "organizer", 2: "tentative", 3: "accepted", 4: "declined", 5: "not_responded"}
RECURRENCE_NAMES = {0: "daily", 1: "weekly", 2: "monthly", 3: "monthly_nth", 5: "yearly", 6: "yearly_nth"}
SENSITIVITY_NAMES = {0: "normal", 1: "personal", 2: "private", 3: "confidential"}
ATTENDEE_TYPES = {1: "required", 2: "optional", 3: "resource"}
_MAX_ID = "\U0010ffff"  # sorts after every entry_id


# ------------------------------------------------------------------- helpers

def _calendar(account):
    acct = oc._find_account(account)
    try:
        return acct, acct.DeliveryStore.GetDefaultFolder(OL_FOLDER_CALENDAR)
    except pywintypes.com_error:
        raise oc.OutlookError(f"Account '{acct.SmtpAddress}' has no calendar. Accounts: {oc._account_names()}")


def _local_fields(dt):
    """COM datetime -> naive local wall-clock datetime (pywin32 mislabels the tz, so use the fields)."""
    return datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


def _local_iso(dt):
    return _local_fields(dt).isoformat()


def _local_midnight_utc(day):
    """Midnight local time on `day` as an aware UTC datetime, using the offset in force on that date."""
    return datetime.combine(day, time.min).astimezone(timezone.utc)


def _local_date(dt_utc):
    return dt_utc.astimezone().date()


def _window(start, end):
    """(start, end) strings or None -> aware UTC datetimes. Defaults: today (local) for 7 days."""
    if start:
        s = oc._parse_ts(start, "start")
    else:
        s = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0).astimezone(timezone.utc)
    e = oc._parse_ts(end, "end") if end else s + timedelta(days=DEFAULT_WINDOW_DAYS)
    if e <= s:
        raise oc.OutlookError("'end' must be after 'start'.")
    if e - s > timedelta(days=MAX_WINDOW_DAYS):
        raise oc.OutlookError(f"Date range too large; use at most {MAX_WINDOW_DAYS} days per call.")
    return s, e


def _range_conditions(s, e):
    """DASL overlap filter (UTC literals): events starting before e and ending after s."""
    return [f"\"urn:schemas:calendar:dtstart\" < '{oc._dasl_ts(e)}'",
            f"\"urn:schemas:calendar:dtend\" > '{oc._dasl_ts(s)}'"]


def _cursor(after, after_entry_id):
    """Items must rank strictly after the cursor. Without an entry_id, everything at that second is excluded."""
    if after_entry_id and not after:
        raise oc.OutlookError("'after_entry_id' only makes sense together with 'after'.")
    if not after:
        return None
    t = oc._parse_ts(after, "after").replace(microsecond=0)
    return (t, after_entry_id or _MAX_ID)


def _rank(item):
    """Sort key, oldest start first: (start to the second, entry_id). Occurrences of a series share an entry_id
    but never a start time, so this is a total order."""
    return (oc._utc(item.Start).replace(microsecond=0), item.EntryID)


def _matches(item, words):
    if not words:
        return True
    hay = " ".join(str(getattr(item, a, "") or "") for a in ("Subject", "Location", "Organizer")).lower()
    return all(w in hay for w in words)


def _ranked(folder, conditions, limit, cursor=None, words=()):
    """Up to limit+1 (rank, item) pairs of events, earliest first, ranking strictly after cursor.

    Reads the whole tie group at the cut so consecutive pages neither skip nor repeat events.
    """
    items = folder.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True  # must come after Sort and before Restrict
    items = items.Restrict("@SQL=" + " AND ".join(conditions))
    got = []
    item = items.GetFirst()
    while item is not None:
        if getattr(item, "Class", None) == OL_APPOINTMENT_CLASS:
            rank = _rank(item)
            if len(got) > limit and rank[0] > got[-1][0][0]:
                break
            if (cursor is None or rank > cursor) and _matches(item, words):
                got.append((rank, item))
        item = items.GetNext()
    got.sort(key=lambda p: p[0])
    return got[:limit + 1]


def _summary(item, store_id):
    """store_id comes from the calendar folder: an expanded occurrence's Parent is its series, not a folder."""
    def attr(name, default=None):
        try:
            v = getattr(item, name)
            return default if v is None else v
        except Exception:
            return default
    return {
        "entry_id": item.EntryID,
        "store_id": store_id,
        "subject": attr("Subject"),
        "start": oc._iso(item.Start),
        "end": oc._iso(item.End),
        "start_local": _local_iso(item.Start),
        "end_local": _local_iso(item.End),
        "all_day": bool(attr("AllDayEvent", False)),
        "location": attr("Location") or None,
        "organizer": attr("Organizer"),
        "is_recurring": bool(attr("IsRecurring", False)),
        "is_meeting": attr("MeetingStatus", 0) != 0,
        "busy_status": BUSY_NAMES.get(attr("BusyStatus", 2), "busy"),
        "response": RESPONSE_NAMES.get(attr("ResponseStatus", 0), "none"),
    }


def _page(ranked, limit, store_id):
    msgs = [_summary(item, store_id) for _, item in ranked[:limit]]
    return {"events": msgs, "has_more": len(ranked) > limit,
            "next_after": msgs[-1]["start"] if msgs else None,
            "next_after_entry_id": msgs[-1]["entry_id"] if msgs else None}


def _get_event_item(entry_id, store_id):
    oc._check_store_allowed(store_id)
    try:
        item = oc._namespace().GetItemFromID(entry_id, store_id)
    except pywintypes.com_error:
        raise oc.OutlookError("Event not found. Use the entry_id and store_id exactly as returned by list_events.")
    if getattr(item, "Class", None) != OL_APPOINTMENT_CLASS:
        raise oc.OutlookError("That item is not a calendar event.")
    return item


def _find_occurrence(folder, entry_id, start_utc):
    """The occurrence of series entry_id starting at start_utc (to the second), or None.

    Looks it up in the expanded calendar view, the same one list_events reads, because
    RecurrencePattern.GetOccurrence fails for occurrences that were moved from the series' schedule.
    """
    s = start_utc.replace(microsecond=0)
    for rank, item in _ranked(folder, _range_conditions(s, s + timedelta(seconds=1)), MAX_PER_INSTANT):
        if rank == (s, entry_id):
            return item
    return None


# --------------------------------------------------------------------- tools

def list_calendars():
    out = []
    for a in oc._accounts():
        try:
            cal = a.DeliveryStore.GetDefaultFolder(OL_FOLDER_CALENDAR)
            out.append({"account": a.SmtpAddress, "calendar": oc._folder_display_path(cal), "items": cal.Items.Count,
                        "default": oc._norm(a.SmtpAddress) == oc._norm(oc.CONFIG["default_account"])})
        except pywintypes.com_error:
            out.append({"account": a.SmtpAddress, "calendar": None, "items": 0, "default": False})
    return {"calendars": out}


def list_events(start=None, end=None, query=None, account=None, limit=50, after=None, after_entry_id=None):
    limit = oc._clamp_limit(limit)
    s, e = _window(start, end)
    acct, cal = _calendar(account)
    words = [w.lower() for w in (query or "").split()]
    ranked = _ranked(cal, _range_conditions(s, e), limit, _cursor(after, after_entry_id), words)
    return {"calendar": oc._folder_display_path(cal), "range": {"start": s.isoformat(), "end": e.isoformat()},
            **_page(ranked, limit, cal.StoreID)}


def get_event(entry_id, store_id, occurrence_start=None, max_chars=oc.DEFAULT_BODY_CHARS):
    item = _get_event_item(entry_id, store_id)
    recurrence = None
    if item.IsRecurring:
        rp = item.GetRecurrencePattern()
        recurrence = {"type": RECURRENCE_NAMES.get(rp.RecurrenceType, str(rp.RecurrenceType)),
                      "interval": rp.Interval, "no_end_date": bool(rp.NoEndDate),
                      "pattern_start": _local_iso(rp.PatternStartDate),
                      "pattern_end": None if rp.NoEndDate else _local_iso(rp.PatternEndDate)}
        if occurrence_start:
            want = oc._parse_ts(occurrence_start, "occurrence_start")
            folder = oc._namespace().GetStoreFromID(store_id).GetDefaultFolder(OL_FOLDER_CALENDAR)
            item = _find_occurrence(folder, entry_id, want)
            if item is None:
                raise oc.OutlookError("This series has no occurrence starting at "
                                      f"{want.isoformat()}. Use a `start` value returned by list_events.")
    elif occurrence_start:
        raise oc.OutlookError("occurrence_start was given but this event is not recurring.")
    max_chars = max(1, int(max_chars))
    body = item.Body or ""
    attendees = []
    for r in item.Recipients:
        addr = r.Address
        try:
            if r.AddressEntry.Type == "EX":
                addr = r.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        except Exception:
            pass
        attendees.append({"name": r.Name, "address": addr, "type": ATTENDEE_TYPES.get(r.Type, str(r.Type)),
                          "response": RESPONSE_NAMES.get(r.MeetingResponseStatus, "none")})
    out = _summary(item, store_id)
    out.update({
        "attendees": attendees,
        "body": body[:max_chars],
        "truncated": len(body) > max_chars,
        "categories": item.Categories or None,
        "sensitivity": SENSITIVITY_NAMES.get(item.Sensitivity, "normal"),
        "reminder_minutes": item.ReminderMinutesBeforeStart if item.ReminderSet else None,
        "recurrence": recurrence,
    })
    return out


EVENT_NOTE_PLAIN = "Event saved to the calendar."
EVENT_NOTE_MEETING = ("Saved as a meeting WITHOUT sending invitations; the user must open it in Outlook and "
                      "click Send to invite the attendees.")


def create_event(subject, start, end=None, duration_minutes=60, all_day=False, location=None, body=None,
                 attendees=None, reminder_minutes=None, busy_status=None, account=None):
    if not subject or not subject.strip():
        raise oc.OutlookError("A subject is required.")
    acct, cal = _calendar(account)
    s_utc = oc._parse_ts(start, "start")
    if all_day:
        first = _local_date(s_utc)
        last = _local_date(oc._parse_ts(end, "end")) if end else first
        if last < first:
            raise oc.OutlookError("'end' must not be before 'start'.")
        s_utc, e_utc = _local_midnight_utc(first), _local_midnight_utc(last + timedelta(days=1))  # end is exclusive
    elif end:
        e_utc = oc._parse_ts(end, "end")
    else:
        if int(duration_minutes) <= 0:
            raise oc.OutlookError("'duration_minutes' must be positive.")
        e_utc = s_utc + timedelta(minutes=int(duration_minutes))
    if e_utc <= s_utc:
        raise oc.OutlookError("'end' must be after 'start'.")
    attendee_list = [a.strip() for a in (attendees or []) if a and a.strip()]
    oc._addr_list(attendee_list, "attendees")  # validates
    if busy_status is not None and busy_status not in BUSY:
        raise oc.OutlookError(f"Invalid busy_status '{busy_status}'. Valid: {sorted(BUSY)}")

    item = cal.Items.Add(OL_APPOINTMENT_ITEM)  # created directly in the account's main calendar
    item.Subject = subject.strip()
    item.AllDayEvent = bool(all_day)
    # Set the times through the UTC properties with aware datetimes. Start/End are converted with today's DST
    # offset by Outlook's object model, which puts events in the other half of the year an hour off.
    item.StartUTC = s_utc
    item.EndUTC = e_utc
    if location:
        item.Location = location
    if body:
        item.Body = body
    if busy_status is not None:
        item.BusyStatus = BUSY[busy_status]
    if reminder_minutes is not None:
        item.ReminderSet = True
        item.ReminderMinutesBeforeStart = int(reminder_minutes)
    if attendee_list:
        item.MeetingStatus = 1  # a meeting, but unsent until the user clicks Send in Outlook
        try:
            item.SendUsingAccount = acct
        except pywintypes.com_error:
            pass
        for a in attendee_list:
            item.Recipients.Add(a).Type = 1
        item.Recipients.ResolveAll()
    item.Save()
    out = _summary(item, cal.StoreID)
    out["calendar"] = oc._folder_display_path(cal)
    out["message"] = EVENT_NOTE_MEETING if attendee_list else EVENT_NOTE_PLAIN
    return out
