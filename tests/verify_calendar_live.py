"""Live calendar check against your real Outlook.

Read-only part: list_calendars, list_events over a wide window (cross-checked against a brute-force scan of the
calendar folder), paging, query filter, get_event incl. a recurring occurrence, local/UTC consistency.
Write part (skip with --no-create): creates THREE events titled "[MCP TEST] ..." months away (a timed one, an all-day one, and one with
an attendee = yourself, to check that no invitation is sent), in the opposite daylight-saving period and verifies them. Delete them manually afterwards.

Run from the project folder:  .venv/Scripts/python.exe tests/verify_calendar_live.py [--no-create]
"""
import asyncio
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
FAILS = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}{(' - ' + str(detail)) if detail else ''}")
    if not ok:
        FAILS.append(name)


async def brute_force(window_start, window_end):
    """Non-recurring events overlapping the window, read straight from the calendar folder (ground truth)."""
    import outlook_client as oc

    def go():
        cal = oc._find_account(None).DeliveryStore.GetDefaultFolder(9)
        out = set()
        for item in cal.Items:
            if getattr(item, "Class", None) != 26 or item.IsRecurring:
                continue
            s, e = oc._utc(item.Start), oc._utc(item.End)
            if s < window_end and e > window_start:
                out.add((item.EntryID, s.replace(microsecond=0).isoformat()))
        return out
    return await oc.run(go)


async def main(create):
    params = StdioServerParameters(command=sys.executable, args=[str(ROOT / "server.py")], cwd=str(ROOT))
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()

            async def call(name, **a):
                return json.loads((await s.call_tool(name, a)).content[0].text)

            async def walk(**kw):
                events, pages, extra = [], 0, {}
                while True:
                    res = await call("list_events", **kw, **extra)
                    assert "error" not in res, res
                    events += res["events"]
                    pages += 1
                    if not res["has_more"]:
                        return events, pages
                    extra = {"after": res["next_after"], "after_entry_id": res["next_after_entry_id"]}
                    assert pages < 500

            print("calendars")
            cals = (await call("list_calendars"))["calendars"]
            default = next(c for c in cals if c["default"] or c["calendar"])
            check("main calendar found", bool(default["calendar"]), f"{default['calendar']} ({default['items']} items)")

            now = datetime.now(timezone.utc)
            ws, we = now - timedelta(days=180), now + timedelta(days=180)
            print("list_events (+-180 days)")
            evs, pages = await walk(start=ws.isoformat(), end=we.isoformat(), limit=100)
            keys = [(e["entry_id"], e["start"]) for e in evs]
            print(f"  {len(evs)} events in {pages} page(s), {sum(e['is_recurring'] for e in evs)} are recurring occurrences")
            check("no duplicate occurrences", len(keys) == len(set(keys)))
            check("sorted by start", [e["start"] for e in evs] == sorted(e["start"] for e in evs))
            check("all overlap the window", all(e["start"] < we.isoformat() and e["end"] > ws.isoformat() for e in evs))
            truth = await brute_force(ws, we)
            got_plain = {(e["entry_id"], e["start"]) for e in evs if not e["is_recurring"]}
            check("non-recurring events equal a brute-force scan of the folder", got_plain == truth,
                  f"{len(got_plain)} vs {len(truth)}; missing {len(truth - got_plain)}, extra {len(got_plain - truth)}")
            small, spages = await walk(start=ws.isoformat(), end=we.isoformat(), limit=4)
            check("limit=4 paging returns the same events", [(e['entry_id'], e['start']) for e in small] == keys, f"{spages} pages")
            check("local time matches UTC conversion", all(
                datetime.fromisoformat(e["start"]).astimezone().replace(tzinfo=None).isoformat() == e["start_local"]
                for e in evs))

            print("query / defaults / errors")
            if evs:
                word = next((w for e in evs for w in (e["subject"] or "").split() if len(w) >= 4), None)
                if word:
                    q = await call("list_events", start=ws.isoformat(), end=we.isoformat(), query=word.upper(), limit=100)
                    check("query filter (case-insensitive)", len(q["events"]) >= 1 and all(
                        word.lower() in " ".join(str(e.get(k) or "") for k in ("subject", "location", "organizer")).lower()
                        for e in q["events"]), f"'{word}' -> {len(q['events'])}")
            d = await call("list_events")
            check("default window = today + 7 days", "error" not in d and (
                datetime.fromisoformat(d["range"]["end"]) - datetime.fromisoformat(d["range"]["start"])) == timedelta(days=7))
            check("too-large range gives a readable error", "error" in await call("list_events", start="2026-01-01", end="2028-01-01"))
            check("bad date gives a readable error", "Invalid" in (await call("list_events", start="tomorrowish")).get("error", ""))

            print("get_event")
            plain = next((e for e in evs if not e["is_recurring"]), None)
            if plain:
                g = await call("get_event", entry_id=plain["entry_id"], store_id=plain["store_id"], max_chars=100)
                check("get_event fields", {"attendees", "body", "recurrence", "start_local"} <= set(g) and len(g["body"]) <= 100)
                check("non-recurring has no recurrence", g["recurrence"] is None)
            rec = [e for e in evs if e["is_recurring"]]
            if len(rec) >= 2:
                same = next(((a, b) for a in rec for b in rec if a is not b and a["entry_id"] == b["entry_id"]), None)
                if same:
                    a, b = same
                    ga = await call("get_event", entry_id=a["entry_id"], store_id=a["store_id"], occurrence_start=a["start"])
                    gb = await call("get_event", entry_id=b["entry_id"], store_id=b["store_id"], occurrence_start=b["start"])
                    check("occurrence_start returns that occurrence", ga["start"] == a["start"] and gb["start"] == b["start"] and ga["start"] != gb["start"])
                    check("recurrence pattern described", ga["recurrence"] and ga["recurrence"]["type"])

            if create:
                print("create_event (adds 3 test events to your calendar)")
                me = (await call("list_accounts"))["accounts"][0]["smtp_address"]
                folders = (await call("list_folders", depth=2))["folders"]
                tracked = {f["path"]: f["items"] for f in folders if f["path"].rsplit("/", 1)[-1] in ("Outbox", "Sent", "Sent Mail")}
                # a date in the opposite DST period from today, so the daylight-saving handling is exercised
                today = datetime.now()
                month = 1 if 4 <= today.month <= 10 else 7
                day = datetime(today.year, month, 15, 3, 0)
                if day < today + timedelta(days=30):
                    day = day.replace(year=day.year + 1)
                print(f"  (test date {day:%Y-%m-%d}, UTC offset there {day.astimezone().utcoffset()}, now {today.astimezone().utcoffset()})")
                c1 = await call("create_event", subject="[MCP TEST] calendar event", start=day.isoformat(),
                                duration_minutes=15, location="Nowhere", body="Created by verify_calendar_live.py. Safe to delete.")
                check("create_event (local naive start)", "error" not in c1 and c1["start_local"] == day.isoformat(), c1.get("start_local") or c1)
                check("end = start + 15 min", c1.get("end_local") == (day + timedelta(minutes=15)).isoformat())
                c2 = await call("create_event", subject="[MCP TEST] calendar meeting (invitation NOT sent)",
                                start=(day + timedelta(hours=1)).isoformat(), attendees=[me])
                check("meeting with attendee is saved as unsent meeting", "error" not in c2 and c2["is_meeting"] and "WITHOUT sending" in c2["message"], c2.get("message"))
                c3 = await call("create_event", subject="[MCP TEST] calendar all-day event", start=day.date().isoformat(),
                                end=(day + timedelta(days=1)).date().isoformat(), all_day=True)
                check("all-day event spans midnight to midnight (end exclusive)", "error" not in c3 and c3["all_day"] and
                      c3["start_local"] == day.replace(hour=0).isoformat() and
                      c3["end_local"] == (day.replace(hour=0) + timedelta(days=2)).isoformat(), (c3.get("start_local"), c3.get("end_local")))
                check("meeting start is exactly as requested", c2.get("start_local") == (day + timedelta(hours=1)).isoformat(), c2.get("start_local"))
                found = await call("list_events", start=(day - timedelta(days=1)).astimezone(timezone.utc).isoformat(),
                                   end=(day + timedelta(days=3)).astimezone(timezone.utc).isoformat(), query="MCP TEST")
                check("created events show up in list_events", {c1["entry_id"], c2["entry_id"], c3["entry_id"]} <= {e["entry_id"] for e in found["events"]})
                g1 = await call("get_event", entry_id=c1["entry_id"], store_id=c1["store_id"])
                check("get_event reads back location and body", g1["location"] == "Nowhere" and "Safe to delete" in g1["body"])
                g2 = await call("get_event", entry_id=c2["entry_id"], store_id=c2["store_id"])
                check("attendee recorded", any(a["address"].lower() == me.lower() for a in g2["attendees"]), [a["address"] for a in g2["attendees"]])
                await asyncio.sleep(3)
                after = {f["path"]: f["items"] for f in (await call("list_folders", depth=2))["folders"] if f["path"] in tracked}
                check("nothing sent: Outbox / Sent counts unchanged", after == tracked, {k: (tracked[k], after[k]) for k in tracked if tracked[k] != after[k]})

    print(f"\n{len(FAILS)} failure(s)" + (": " + ", ".join(FAILS) if FAILS else ""))
    return 1 if FAILS else 0


sys.exit(asyncio.run(main("--no-create" not in sys.argv)))
