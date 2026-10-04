"""Smoke test: drives server.py over stdio, calls every read tool, checks search behaviour and timing,
and (unless --no-draft) creates ONE draft to yourself, subject prefixed [MCP TEST]. Prints metadata only,
never message bodies. Accounts and search words are discovered from your own mailbox at runtime.

Run from the project folder:  .venv/Scripts/python.exe tests/smoke_test.py [--no-draft]
"""
import asyncio
import json
import re
import sys
import time
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parent.parent
FAILS = []
WORD = r"[^\W\d_]{6,}"  # a 6+ letter word


def check(name, ok, detail=""):
    print(f"  [{'PASS' if ok else 'FAIL'}] {name}{(' - ' + str(detail)) if detail else ''}")
    if not ok:
        FAILS.append(name)


async def main(make_draft):
    params = StdioServerParameters(command=sys.executable, args=[str(ROOT / "server.py")], cwd=str(ROOT))
    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()

            async def call(name, **a):
                t = time.perf_counter()
                res = json.loads((await s.call_tool(name, a)).content[0].text)
                return res, time.perf_counter() - t

            print("accounts / folders")
            acc, _ = await call("list_accounts")
            check("list_accounts", len(acc["accounts"]) >= 1)
            me = next((a for a in acc["accounts"] if a["default"]), acc["accounts"][0])["smtp_address"]
            stores = {a["store"] for a in acc["accounts"]}
            fl, _ = await call("list_folders", depth=2)
            check("list_folders", any(f["path"] == f"{me}/Inbox" for f in fl["folders"]))
            check("archive hidden by default", all(f["path"].split("/")[0] in stores for f in fl["folders"]))
            fla, _ = await call("list_folders", depth=1, include_archive=True)
            arch = next((f["path"].split("/")[0] for f in fla["folders"] if f["path"].split("/")[0] not in stores), None)
            print(f"  (archive store configured: {bool(arch)})")
            sent = next((f["path"] for f in fl["folders"] if f["path"].startswith(me + "/")
                         and f["path"].rsplit("/", 1)[-1] in ("Sent", "Sent Items", "Sent Mail")), None)

            print("list_messages")
            lm, t = await call("list_messages", limit=20)
            msgs = lm["messages"]
            check("default inbox, 20 newest", len(msgs) == 20 and lm["folder"].endswith("/Inbox"), f"{t:.1f}s")
            check("newest first", [m["received"] for m in msgs] == sorted((m["received"] for m in msgs), reverse=True))
            check("limit capped at 100", len((await call("list_messages", limit=500))[0]["messages"]) <= 100)
            un, _ = await call("list_messages", limit=5, unread_only=True)
            check("unread_only", all(m["unread"] for m in un["messages"]))
            sn, _ = await call("list_messages", limit=50, since=msgs[9]["received"])
            check("since filter", all(m["received"] >= msgs[9]["received"] for m in sn["messages"]))

            print("get_message / get_thread")
            m = next(x for x in msgs if x.get("subject") and "type" not in x)
            full, _ = await call("get_message", entry_id=m["entry_id"], store_id=m["store_id"], max_chars=200)
            check("get_message truncation", len(full["body"]) <= 200 and "truncated" in full)
            th, _ = await call("get_thread", entry_id=m["entry_id"], store_id=m["store_id"])
            check("get_thread contains the message", any(x["entry_id"] == m["entry_id"] for x in th["messages"]))

            print("search (main use case)")
            subj = m["subject"]
            kw = next(w for x in msgs for w in re.findall(WORD, x.get("subject") or ""))
            r1, t1 = await call("search_messages", query=subj, limit=5)
            check("find message by its exact subject", any(x["entry_id"] == m["entry_id"] for x in r1["messages"]), f"{t1:.1f}s")
            r2, t2 = await call("search_messages", query="", sender=m["sender_address"], limit=5)
            check("filter by sender address only (no text)", len(r2["messages"]) >= 1, f"{t2:.1f}s")
            body_words = re.findall(WORD, full["body"])
            if body_words:
                r3, t3 = await call("search_messages", query=body_words[0], limit=20)
                check("body word search finds source message", any(x["entry_id"] == m["entry_id"] for x in r3["messages"]), f"{t3:.1f}s")
            words = subj.split()
            t4 = 0
            if len(words) >= 3:
                r4, t4 = await call("search_messages", query=f"{words[-1]} {words[0]}", limit=20)
                check("multi-word query (words not adjacent)", any(x["entry_id"] == m["entry_id"] for x in r4["messages"]), f"{t4:.1f}s")
            r5, _ = await call("search_messages", query=kw, limit=20)
            r5u, _ = await call("search_messages", query=kw.upper(), limit=20)
            check("case-insensitive", len(r5["messages"]) == len(r5u["messages"]) > 0)
            r6, t6 = await call("search_messages", query="ü", limit=5)
            check("non-ASCII query", "error" not in r6, f"{t6:.1f}s")
            t7 = 0
            if sent:
                r7, t7 = await call("search_messages", query="a", folder=sent, limit=5)
                check("search Sent folder", "error" not in r7 and len(r7["messages"]) > 0, f"{t7:.1f}s")
            r8, t8 = await call("search_messages", query=kw, limit=100, since="2020-01-01")
            check("wide search, limit 100", "error" not in r8 and len(r8["messages"]) <= 100, f"{len(r8['messages'])} hits, {t8:.1f}s")
            rp, tp = await call("search_messages", query='"' + " ".join(subj.split()[:2]) + '"', limit=20)
            check("quoted phrase", "error" not in rp and len(rp["messages"]) > 0, f"{tp:.1f}s")
            check("fast searches all under 20s (first one is cold)", max(t1, t2, t4, t6, t7, t8) < 20, f"max {max(t1, t2, t4, t6, t7, t8):.1f}s")
            rx, tx = await call("search_messages", query=kw, exhaustive=True, limit=20)
            check("exhaustive finds at least as much as fast", len(rx["messages"]) >= len(r5["messages"]), f"{len(r5['messages'])} vs {len(rx['messages'])}, exhaustive {tx:.0f}s")
            rq, _ = await call("search_messages", query="it's \"quoted\" [x] 100%_", limit=3)
            check("special characters don't break the query", "error" not in rq)
            rd, _ = await call("search_messages", query="a", limit=100)
            check("default search covers Inbox and Sent", any(p.endswith("/Inbox") for p in rd["searched"]) and (not sent or sent in rd["searched"]), rd["searched"])
            rn, _ = await call("search_messages", query=kw, limit=10)
            check("archive not searched by default", all(p.split("/")[0] in stores for p in rn["searched"]))
            if arch:
                ra, ta = await call("search_messages", query=kw, include_archive=True, limit=10)
                check("archive search only with flag", any(p.startswith(arch) for p in ra["searched"]) and "error" not in ra, f"{ta:.1f}s")
                check("archive fast search under 20s", ta < 20, f"{ta:.1f}s")

            if make_draft:
                print("draft")
                d, _ = await call("create_draft", to=[me], subject="[MCP TEST] smoke test draft", body="Created by smoke_test.py. Safe to delete.")
                check("create_draft in account Drafts", "error" not in d and d["folder"].startswith(f"{me}/Drafts"), d.get("folder"))
                g, _ = await call("get_message", entry_id=d["entry_id"], store_id=d["store_id"])
                check("draft is not sent", g["subject"].startswith("[MCP TEST]") and "Drafts" in g["folder"])

    print(f"\n{len(FAILS)} failure(s)" + (": " + ", ".join(FAILS) if FAILS else ""))
    return 1 if FAILS else 0


sys.exit(asyncio.run(main("--no-draft" not in sys.argv)))
