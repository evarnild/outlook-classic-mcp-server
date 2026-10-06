# Outlook Classic MCP server (read mail and calendar, create drafts and events)

A small local MCP server that lets Claude read and search your mail and calendar in **Outlook classic** (via COM)
and create **drafts** and **calendar events**. It can never send mail or invitations, and never deletes, moves or
edits existing items: there is no code path for it. You review and send
drafts yourself in Outlook. Unofficial; not affiliated with Microsoft or Anthropic.

- Windows 11, Outlook classic (the "new Outlook" has no COM support), Python 3.11+ (tested on 3.14).
- No credentials anywhere: it only talks to your local Outlook, which handles IMAP/SMTP auth itself.
- stdio transport. Logs go to stderr and `outlook_mcp.log` (rotating). Email bodies are never logged.

## Setup

```powershell
git clone https://github.com/evarnild/outlook-classic-mcp-server
cd outlook-classic-mcp-server
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

Outlook classic should be installed and your accounts configured. The server starts Outlook if it isn't running.
Optionally copy `config.example.toml` to `config.toml` and edit it (see Configuration).

In the snippets below, replace `C:\path\to\outlook-classic-mcp-server` with the folder you cloned into.

## Claude Desktop

Settings > Developer > Edit Config, add to `claude_desktop_config.json`, then fully restart Claude Desktop:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\path\\to\\outlook-classic-mcp-server\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\outlook-classic-mcp-server\\server.py"]
    }
  }
}
```

## Claude Code

```powershell
claude mcp add outlook --scope user -- "C:\path\to\outlook-classic-mcp-server\.venv\Scripts\python.exe" "C:\path\to\outlook-classic-mcp-server\server.py"
```

(`--scope user` makes it available in every project; drop it for the current project only.)
Check with `claude mcp list`.

## Tools

| Tool | What it does |
|---|---|
| `list_accounts` | Accounts with SMTP address, store name, default flag |
| `list_folders` | Mail folders with item / unread counts |
| `list_messages` | Newest-first list with 200-char previews (default: default account's Inbox) |
| `search_messages` | Text / sender / date search; default scope is Inbox **and Sent** of every account |
| `get_message` | Full headers, plain-text body (truncated to `max_chars`, default 8000), attachment names/sizes |
| `get_thread` | Messages in the same conversation (best effort, by conversation topic, Inbox + Sent) |
| `create_draft` | New draft in the chosen account's Drafts folder |
| `create_reply_draft` | Reply / reply-all draft with your text above the quoted original |
| `update_draft` | Edit an existing draft; refuses anything not in a Drafts folder |
| `list_calendars` | Each account's main calendar (path, event count) |
| `list_events` | Events overlapping a date range, earliest first; recurring series are expanded into occurrences |
| `get_event` | Full event: attendees and responses, description, reminder, recurrence; one occurrence of a series |
| `create_event` | New event in the main calendar (timed or all-day); with attendees it is saved as an **unsent** meeting |

Every message has an `entry_id` and `store_id`; pass both back to `get_message`, `get_thread`, and the
reply/update draft tools. Dates are ISO 8601. `limit` defaults to 20 (max 100).

### Paging

`list_messages` and `search_messages` return at most 100 messages per call, newest first, plus `has_more`,
`next_before` and `next_before_entry_id`. To get older messages, repeat the call with `before=<next_before>` and
`before_entry_id=<next_before_entry_id>` while `has_more` is true. Passing both makes messages that share the same
timestamp at a page boundary neither skipped nor repeated. `before` is strict (received strictly before it) and
combines with `since`, `unread_only`, `sender`, `folder` and `limit`.

`list_messages(count_only=true, ...)` returns just `{"count": N}` for the matching messages, e.g. unread before a date.

Timestamps: results are UTC (`+00:00`). In `since` / `before`, a value without an offset is read as the machine's
local time. Non-mail items in a folder (meeting requests, reports) are listed too, as minimal entries with a `type`
field (their message class), so counts match Outlook's own unread counter.

### Calendar

Only each account's **main (default) calendar** is used; other calendars and sub-calendars are ignored.

- `list_events(start, end, query, account, limit)` returns events that overlap `[start, end)`, earliest first
  (default: today for 7 days, at most 366 days). Each event has UTC `start`/`end`, the local wall-clock
  `start_local`/`end_local` as Outlook shows them, `all_day` (the end of an all-day event is exclusive), location,
  organizer, `is_recurring`, `is_meeting`, busy status and your response. `query` filters on title, location and
  organizer (all words must match). Paging works like mail: repeat with `after=<next_after>` and
  `after_entry_id=<next_after_entry_id>` while `has_more` is true.
- All occurrences of a recurring series share one `entry_id`; pass an occurrence's `start` as `occurrence_start` to
  `get_event` to get that occurrence (also works for occurrences that were moved).
- `create_event(subject, start, end | duration_minutes, all_day, location, body, attendees, reminder_minutes,
  busy_status, account)` saves a new event immediately. A `start` without an offset is local time. **It never sends
  invitations:** with `attendees` the event is saved as a meeting whose invitations are *not* sent; open it in
  Outlook and click Send if you want to invite them. There is no tool to edit or delete events.
- Outlook's object model converts `Start`/`End` with today's daylight-saving offset, which puts events in the other
  half of the year an hour off. Events are therefore created through the `StartUTC`/`EndUTC` properties, and event
  times are filtered with UTC DASL literals (the locale-dependent Jet date syntax is not used).

### Search

- Every word must match, in any order, in subject or body. Use `"double quotes"` for an exact phrase.
- **Fast mode (default)** uses Outlook's content index: the subject matches anywhere in a word, the body matches
  at the *start* of a word. Typically 1-5 seconds, including the archive.
- **`exhaustive=true`** scans bodies for substrings anywhere in a word (finds "Quarterreport" when searching
  "report"), but takes 20+ seconds per large folder.
- Without `folder`, it searches Inbox + Sent for each account. Pass `folder="me@example.com/Projects"`
  or similar for other folders (see `list_folders`).

### Archive

The local archive store (set as `archive_store` in `config.toml`) is never touched unless a tool is called with
`include_archive=true`. The tool descriptions tell Claude to do so only when you explicitly ask for the archive,
e.g. "search the archive for ...".

## Configuration (`config.toml`, optional, git-ignored)

```toml
allowed_accounts = []                       # empty = all accounts; or ["me@example.com"] to restrict
default_account = "me@example.com"
archive_store = "my-archive"                # display name of a local archive store; optional
```

See `config.example.toml`.

`allowed_accounts` filters every tool, including lookups by `entry_id`.

## Safety

- No send, delete, move, mark-as-read, rules or attachment download code exists. The only writes are
  `Items.Add` + `.Save()` for new drafts and new calendar events, and `.Save()` on existing drafts. Events are
  created but never edited, deleted or sent; invitations to attendees are never sent.
- Email content is untrusted: it can contain text that tries to instruct the AI. Tool descriptions say to treat it
  as data. Having no send tool is the main guardrail; still read drafts before sending.
- Drafts are created directly in the account's own Drafts folder (not moved), with `SendUsingAccount` set.

## Troubleshooting

- **"A program is trying to access email addresses" prompt from Outlook.** This appears when Outlook doesn't
  recognise your antivirus / Windows Security as up to date. Fix that (Windows Security status, or your
  antivirus); this server does not try to bypass the prompt.
- **Outlook COM errors** are returned as readable messages including the HRESULT. If Outlook can't be started, close
  it, reopen it once manually and retry.
- **"Folder not found"** errors list the valid options. Folder paths look like `account/Inbox`,
  `me@gmail.com/[Gmail]/Sent Mail`.
- **MCP Inspector** in PowerShell: `npx.cmd @modelcontextprotocol/inspector .venv\Scripts\python.exe server.py`
  (`npx.cmd` avoids the blocked `npx.ps1` script; run it from this folder).
- `mcp` is pinned `<2` because SDK 2.x renamed `FastMCP`.

## Tests

```powershell
.venv\Scripts\python.exe -m unittest discover -s tests -t . -p "test_*.py"   # offline unit tests, no Outlook needed
.venv\Scripts\python.exe tests\smoke_test.py                # needs Outlook; add --no-draft to skip the draft
.venv\Scripts\python.exe tests\verify_paging_live.py        # read-only; pages through your unread mail
.venv\Scripts\python.exe tests\verify_calendar_live.py       # calendar checks; add --no-create to skip creating events
```

`smoke_test.py` calls every read tool, checks search behaviour and timings, and (unless `--no-draft`) creates **one**
draft to yourself (`[MCP TEST] smoke test draft`) in your default account's Drafts folder; delete test drafts manually.
`verify_paging_live.py` pages through all unread mail of each account and compares the total with Outlook's own
unread counters. `verify_calendar_live.py` cross-checks `list_events` against a brute-force scan of your calendar and
(unless `--no-create`) creates three `[MCP TEST]` events months away, one with yourself as attendee, and verifies
nothing was sent; delete those events manually. The tests discover your accounts at runtime and print metadata only,
never message bodies.

## License

[MIT](LICENSE)
