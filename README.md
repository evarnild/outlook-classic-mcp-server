# Outlook Classic MCP server (read + drafts only)

A small local MCP server that lets Claude read and search your mail in **Outlook classic** (via COM) and
create **drafts**. It can never send, delete or move mail: there is no code path for it. You review and send
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

- No send, delete, move, mark-as-read, rules, attachment download or calendar code exists. The only writes are
  `Items.Add` + `.Save()` for new drafts and `.Save()` on existing drafts.
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
.venv\Scripts\python.exe -m unittest tests.test_paging       # offline unit tests, no Outlook needed
.venv\Scripts\python.exe tests\smoke_test.py                # needs Outlook; add --no-draft to skip the draft
.venv\Scripts\python.exe tests\verify_paging_live.py        # read-only; pages through your unread mail
```

`smoke_test.py` calls every read tool, checks search behaviour and timings, and (unless `--no-draft`) creates **one**
draft to yourself (`[MCP TEST] smoke test draft`) in your default account's Drafts folder; delete test drafts manually.
`verify_paging_live.py` pages through all unread mail of each account and compares the total with Outlook's own
unread counters. The tests discover your accounts at runtime and print metadata only, never message bodies.

## License

[MIT](LICENSE)
