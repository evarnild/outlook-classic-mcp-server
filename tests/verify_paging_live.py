"""Live paging check against the real mailboxes (read-only, metadata only).

Pages through all unread messages of each Inbox with limit=100 and compares the total with Outlook's own
unread counter and with count_only; also checks search_messages paging and before/since/unread_only combos.

Run from the project folder:  .venv/Scripts/python.exe tests/verify_paging_live.py
"""
import asyncio
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parent.parent
FAILS = []


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}{(' - ' + str(detail)) if detail else ''}")
    if not ok:
        FAILS.append(name)


async def main():
    params = StdioServerParameters(command=sys.executable, args=[str(ROOT / "server.py")], cwd=str(ROOT))
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()

            async def call(name, **a):
                return json.loads((await s.call_tool(name, a)).content[0].text)

            async def walk(tool, **kw):
                """Follow next_before / next_before_entry_id until has_more is false."""
                msgs, pages, extra = [], 0, {}
                while True:
                    res = await call(tool, **kw, **extra)
                    assert "error" not in res, res
                    msgs += res["messages"]
                    pages += 1
                    if not res["has_more"]:
                        return msgs, pages
                    extra = {"before": res["next_before"], "before_entry_id": res["next_before_entry_id"]}
                    assert pages < 500

            folders = {f["path"]: f for f in (await call("list_folders", depth=1))["folders"]}
            now = datetime.now(timezone.utc)

            accounts = [a["smtp_address"] for a in (await call("list_accounts"))["accounts"]]
            for acct in accounts:
                path = f"{acct}/Inbox"
                print(f"\n{path}")
                unread_before = folders[path]["unread"]
                msgs, pages = await walk("list_messages", folder=path, unread_only=True, limit=100)
                ids = [m["entry_id"] for m in msgs]
                cnt = (await call("list_messages", folder=path, unread_only=True, count_only=True))["count"]
                unread_after = {f["path"]: f for f in (await call("list_folders", depth=1))["folders"]}[path]["unread"]
                print(f"  paged unread: {len(ids)} in {pages} page(s) | folder unread counter: {unread_before}->{unread_after} | count_only: {cnt}")
                check("no duplicates", len(ids) == len(set(ids)))
                check("total equals folder unread counter", len(ids) in (unread_before, unread_after), f"{len(ids)} vs {unread_before}/{unread_after}")
                check("count_only matches paging", cnt == len(ids))
                recv = [m["received"] for m in msgs]
                check("strictly newest first", recv == sorted(recv, reverse=True))
                check("received timestamps are real UTC (none in the future)", all(datetime.fromisoformat(x) <= now for x in recv))

                # small pages must give the same set (stresses tie-breaking on the boundaries)
                small, spages = await walk("list_messages", folder=path, unread_only=True, limit=7)
                check("limit=7 paging returns the same messages", [m["entry_id"] for m in small] == ids, f"{spages} pages")

                # combining before + since + unread_only
                if len(recv) >= 30:
                    hi, lo = recv[9], recv[min(len(recv) - 1, 29)]
                    sel = await call("list_messages", folder=path, unread_only=True, limit=100, before=hi, since=lo)
                    expect = [m for m in msgs if lo <= m["received"] < hi]
                    check("before + since + unread_only combine", [m["entry_id"] for m in sel["messages"]] == [m["entry_id"] for m in expect], f"{len(expect)} expected")
                    c2 = await call("list_messages", folder=path, unread_only=True, before=hi, since=lo, count_only=True)
                    check("count_only with before + since", c2["count"] == len(expect), c2["count"])
                    check("before is strict (boundary message excluded)", all(m["received"] < hi for m in sel["messages"]))

            print("\nsearch paging")
            first = (await call("list_messages", limit=1))["messages"][0]
            sender = first.get("sender_address") or first["sender_name"]
            full, pages = await walk("search_messages", query="", sender=sender, since="2026-01-01", limit=100)
            fid = [m["entry_id"] for m in full]
            print(f"  newest message's sender since 2026-01-01: {len(fid)} messages in {pages} page(s)")
            check("search: no duplicates", len(fid) == len(set(fid)))
            check("search: newest first", [m["received"] for m in full] == sorted((m["received"] for m in full), reverse=True))
            small, sp = await walk("search_messages", query="", sender=sender, since="2026-01-01", limit=9)
            check("search: limit=9 paging returns the same messages", [m["entry_id"] for m in small] == fid, f"{sp} pages")
            s1 = await call("search_messages", query="a", sender=sender, since="2026-01-01", limit=3)
            check("search: has_more / next_before present", {"has_more", "next_before", "next_before_entry_id"} <= set(s1))

            print("\nbackwards compatibility (new params omitted)")
            old = await call("list_messages", limit=20)
            check("list_messages default still returns 20 newest", len(old["messages"]) == 20 and old["folder"].endswith("/Inbox"))
            check("search default still covers Inbox + Sent", any("Sent" in p for p in (await call("search_messages", query="a", limit=3))["searched"]))
            bad = await call("list_messages", before="not a date")
            check("invalid before gives a readable error", "Invalid 'before'" in bad.get("error", ""))

    print(f"\n{len(FAILS)} failure(s)" + (": " + ", ".join(FAILS) if FAILS else ""))
    return 1 if FAILS else 0


sys.exit(asyncio.run(main()))
