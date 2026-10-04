"""All Outlook COM logic. Every COM call runs on a single dedicated worker thread.

The only writes are drafts: create_draft / create_reply_draft save new drafts, update_draft edits
an existing draft. Nothing in this module sends, deletes or moves mail.
"""
import html as _html
import logging
import re
import tomllib
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

from constants import OL_FOLDER_DRAFTS, OL_FOLDER_INBOX, OL_FOLDER_SENT_MAIL, OL_MAIL_CLASS, OL_MAIL_ITEM

log = logging.getLogger("outlook_mcp")

MAX_LIMIT = 100
DEFAULT_BODY_CHARS = 8000
ARCHIVE_HINT = "Only set include_archive=true if the user explicitly asked to search the archive."


class OutlookError(Exception):
    """Error with a message that is safe and useful to show to the model."""


def load_config():
    path = Path(__file__).with_name("config.toml")
    cfg = {"allowed_accounts": [], "default_account": None, "archive_store": None}
    try:
        with open(path, "rb") as f:
            cfg.update(tomllib.load(f))
    except FileNotFoundError:
        pass
    return cfg


CONFIG = load_config()

# ---------------------------------------------------------------- COM worker

_executor = ThreadPoolExecutor(max_workers=1, initializer=pythoncom.CoInitialize)
_ns = None  # only ever touched on the worker thread
_app = None


def _namespace():
    global _ns, _app
    if _ns is None:
        try:
            _app = win32com.client.Dispatch("Outlook.Application")
            _ns = _app.GetNamespace("MAPI")
        except pywintypes.com_error as e:
            raise OutlookError(f"Could not start or connect to Outlook classic (HRESULT {_hr(e)}). "
                               "Make sure Outlook classic (not new Outlook) is installed.")
    return _ns


def _hr(e):
    try:
        return hex(e.hresult & 0xFFFFFFFF)
    except Exception:
        return "?"


def _wrap(fn, args, kwargs):
    try:
        return fn(*args, **kwargs)
    except OutlookError as e:
        return {"error": str(e)}
    except pywintypes.com_error as e:
        log.warning("COM error in %s: %s", fn.__name__, _hr(e))
        detail = e.excepinfo[2] if getattr(e, "excepinfo", None) and len(e.excepinfo) > 2 else e.strerror
        return {"error": f"Outlook COM error (HRESULT {_hr(e)}): {detail}"}


async def run(fn, *args, **kwargs):
    import asyncio
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(_executor, lambda: _wrap(fn, args, kwargs))


# ------------------------------------------------------------------- helpers

def _norm(s):
    return (s or "").strip().lower()


def _accounts():
    """Allowed accounts as list of COM Account objects."""
    allowed = {_norm(a) for a in CONFIG["allowed_accounts"]}
    out = []
    for acct in _namespace().Accounts:
        if not allowed or _norm(acct.SmtpAddress) in allowed or _norm(acct.DisplayName) in allowed:
            out.append(acct)
    return out


def _account_names():
    return [a.SmtpAddress for a in _accounts()]


def _find_account(name):
    name = name or CONFIG["default_account"]
    accts = _accounts()
    if not name:
        if len(accts) == 1:
            return accts[0]
        raise OutlookError(f"No account given. Valid accounts: {_account_names()}")
    for a in accts:
        if _norm(name) in (_norm(a.SmtpAddress), _norm(a.DisplayName)):
            return a
    raise OutlookError(f"Account '{name}' not found. Valid accounts: {_account_names()}")


def _store_names(include_archive):
    names = [a.DeliveryStore.DisplayName for a in _accounts()]
    if include_archive and CONFIG["archive_store"]:
        names.append(CONFIG["archive_store"])
    return names


def _find_store(name, include_archive):
    for store in _namespace().Stores:
        if _norm(store.DisplayName) == _norm(name) and _norm(store.DisplayName) in map(_norm, _store_names(include_archive)):
            return store
    if CONFIG["archive_store"] and _norm(name) == _norm(CONFIG["archive_store"]):
        raise OutlookError("That folder is in the archive store. " + ARCHIVE_HINT)
    return None


def _is_archive(store_name):
    return bool(CONFIG["archive_store"]) and _norm(store_name) == _norm(CONFIG["archive_store"])


def _resolve_folder(path, include_archive=False):
    """'me@example.com/Inbox' or 'Inbox' (default account) -> Folder."""
    parts = [p for p in re.split(r"[\\/]+", path or "") if p]
    if not parts:
        raise OutlookError("Empty folder path.")
    store = _find_store(parts[0], include_archive)
    if store is not None:
        parts = parts[1:]
    else:
        store = _find_account(None).DeliveryStore
    if not parts:
        raise OutlookError(f"Give a folder inside the store, e.g. '{store.DisplayName}/Inbox'.")
    if _norm(parts[0]) == "inbox" and not _is_archive(store.DisplayName):
        folder = _default_folder(store, OL_FOLDER_INBOX)
        parts = parts[1:]
    else:
        folder = store.GetRootFolder()
    for part in parts:
        for sub in folder.Folders:
            if _norm(sub.Name) == _norm(part):
                folder = sub
                break
        else:
            raise OutlookError(f"Folder '{part}' not found in '{folder.FolderPath}'. "
                               f"Valid subfolders: {[f.Name for f in folder.Folders]}")
    return folder


def _default_folder(store, kind):
    return store.GetDefaultFolder(kind)


def _folder_display_path(folder):
    return folder.FolderPath.lstrip("\\").replace("\\", "/")


def _check_store_allowed(store_id):
    """Item lookups by id are restricted to allowed accounts (+ archive store)."""
    store = _namespace().GetStoreFromID(store_id)
    name = store.DisplayName
    if _norm(name) in map(_norm, _store_names(False)) or _is_archive(name):
        return
    raise OutlookError("That store is not exposed by this server.")


def _get_item(entry_id, store_id):
    _check_store_allowed(store_id)
    try:
        item = _namespace().GetItemFromID(entry_id, store_id)
    except pywintypes.com_error:
        raise OutlookError("Message not found. Use the entry_id and store_id exactly as returned by a list/search tool.")
    if getattr(item, "Class", None) != OL_MAIL_CLASS:
        raise OutlookError("That item is not an email message.")
    return item


def _utc(dt):
    """COM datetime -> aware UTC datetime.

    pywin32 returns the machine-local wall-clock time but labels it with a bogus tzinfo (it said "UTC+00:00"
    for a 13:37 CEST message, which is really 11:37 UTC). So use the fields only and convert from local time.
    """
    return datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second).astimezone(timezone.utc)


def _iso(dt):
    try:
        return _utc(dt).isoformat()
    except Exception:
        return None


def _when(item):
    try:
        return item.ReceivedTime
    except Exception:
        return item.SentOn


def _sender_address(item):
    try:
        if item.SenderEmailType == "EX":
            return item.Sender.GetExchangeUser().PrimarySmtpAddress
    except Exception:
        pass
    return item.SenderEmailAddress


def _recipients(item):
    to, cc = [], []
    for r in item.Recipients:
        addr = r.Address
        try:
            if r.AddressEntry.Type == "EX":
                addr = r.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        except Exception:
            pass
        entry = f"{r.Name} <{addr}>" if addr and r.Name and r.Name != addr else (addr or r.Name)
        if r.Type == 1:
            to.append(entry)
        elif r.Type == 2:
            cc.append(entry)
    return to, cc


def _preview(item, n=200):
    try:
        return re.sub(r"\s+", " ", item.Body[:n * 3]).strip()[:n]
    except Exception:
        return ""


def _minimal_summary(item):
    """Non-mail item (meeting request, report, ...): identity and basics only, flagged with its message class."""
    def attr(name):
        try:
            return getattr(item, name)
        except Exception:
            return None
    return {
        "entry_id": item.EntryID,
        "store_id": item.Parent.StoreID,
        "folder": _folder_display_path(item.Parent),
        "type": attr("MessageClass") or "unknown",
        "subject": attr("Subject"),
        "sender_name": attr("SenderName"),
        "received": _iso(_when(item)),
        "unread": bool(attr("UnRead")),
    }


def _summary(item):
    if getattr(item, "Class", None) != OL_MAIL_CLASS:
        return _minimal_summary(item)
    return {
        "entry_id": item.EntryID,
        "store_id": item.Parent.StoreID,
        "folder": _folder_display_path(item.Parent),
        "subject": item.Subject,
        "sender_name": item.SenderName,
        "sender_address": _sender_address(item),
        "received": _iso(_when(item)),
        "unread": bool(item.UnRead),
        "has_attachments": item.Attachments.Count > 0,
        "preview": _preview(item),
    }


def _q(s):
    """Escape a value for a single-quoted DASL string."""
    return str(s).replace("'", "''")


def _parse_ts(value, name):
    """ISO 8601 string -> aware UTC datetime. Timestamps without an offset are machine-local time."""
    try:
        dt = datetime.fromisoformat(str(value).strip())
    except ValueError:
        raise OutlookError(f"Invalid '{name}' value '{value}'; use ISO 8601, e.g. 2026-09-30 or "
                           "2026-09-30T08:00:00+00:00.")
    return dt.astimezone(timezone.utc)  # naive datetimes are taken as local time


def _dasl_ts(dt):
    """DASL date literals are UTC."""
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


DATE = '"urn:schemas:httpmail:datereceived"'


def _range_conditions(unread_only=False, since=None, before=None, before_entry_id=None):
    """DASL conditions shared by list/search. Returns (conditions, cursor).

    cursor is the (received, entry_id) rank of the page boundary (or None); items must rank strictly below it.
    Without before_entry_id the cursor is (before, ""), i.e. strictly older than `before`. With it, the whole
    second of `before` is fetched so the tie-break on entry_id can be applied exactly.
    """
    conds = []
    if unread_only:
        conds.append('"urn:schemas:httpmail:read" = 0')
    if since:
        conds.append(f"{DATE} >= '{_dasl_ts(_parse_ts(since, 'since'))}'")
    cursor = None
    if before_entry_id and not before:
        raise OutlookError("'before_entry_id' only makes sense together with 'before'.")
    if before:
        b = _parse_ts(before, "before").replace(microsecond=0)
        upper = b + timedelta(seconds=1) if before_entry_id else b
        conds.append(f"{DATE} < '{_dasl_ts(upper)}'")
        cursor = (b, before_entry_id or "")
    return conds, cursor


def _clamp_limit(limit):
    return max(1, min(int(limit), MAX_LIMIT))


def _query_folder(folder, conditions, limit):
    """Newest-first mail summaries from one folder matching DASL conditions."""
    items = folder.Items
    if conditions:
        items = items.Restrict("@SQL=" + " AND ".join(conditions))
    items.Sort("[ReceivedTime]", True)
    out = []
    item = items.GetFirst()
    while item is not None and len(out) < limit:
        if getattr(item, "Class", None) == OL_MAIL_CLASS:
            out.append(_summary(item))
        item = items.GetNext()
    return out


def _rank(item):
    """Sort key, newest first: (received to the second, entry_id). Gives a total order even for equal times."""
    return (_utc(_when(item)).replace(microsecond=0), item.EntryID)


def _ranked(folder, conditions, limit, cursor=None):
    """Up to limit+1 (rank, item) pairs of items in folder, newest first, ranking strictly below cursor.

    Non-mail items (meeting requests etc.) are included so counts match Outlook's own; _summary flags them.

    The extra item tells the caller whether more exist. Because Outlook orders items with the same
    timestamp arbitrarily, the whole tie group at the cut is read and ordered by entry_id, so consecutive
    pages neither skip nor repeat messages.
    """
    items = folder.Items
    if conditions:
        items = items.Restrict("@SQL=" + " AND ".join(conditions))
    items.Sort("[ReceivedTime]", True)
    got = []
    item = items.GetFirst()
    while item is not None:
        rank = _rank(item)
        if len(got) > limit and rank[0] < got[-1][0][0]:
            break  # past the tie group at the cut
        if cursor is None or rank < cursor:
            got.append((rank, item))
        item = items.GetNext()
    got.sort(key=lambda p: p[0], reverse=True)
    return got[:limit + 1]


def _page(ranked, limit):
    """Merge ranked pairs (possibly from several folders) into one page with paging info."""
    ranked = sorted(ranked, key=lambda p: p[0], reverse=True)
    msgs = [_summary(item) for _, item in ranked[:limit]]
    return {"messages": msgs, "has_more": len(ranked) > limit,
            "next_before": msgs[-1]["received"] if msgs else None,
            "next_before_entry_id": msgs[-1]["entry_id"] if msgs else None}


def _mail_folders(root):
    """Recursively yield mail folders under root."""
    for f in root.Folders:
        try:
            if f.DefaultItemType == 0:
                yield f
            yield from _mail_folders(f)
        except pywintypes.com_error:
            continue


# --------------------------------------------------------------------- tools

def list_accounts():
    out = []
    for a in _accounts():
        out.append({"display_name": a.DisplayName, "smtp_address": a.SmtpAddress,
                    "store": a.DeliveryStore.DisplayName,
                    "default": _norm(a.SmtpAddress) == _norm(CONFIG["default_account"])})
    return {"accounts": out}


def list_folders(account=None, depth=2, include_archive=False):
    stores = [_find_account(account).DeliveryStore] if account else [a.DeliveryStore for a in _accounts()]
    if include_archive and not account and CONFIG["archive_store"]:
        arch = _find_store(CONFIG["archive_store"], True)
        if arch is not None:
            stores.append(arch)
    out = []

    def walk(folder, level):
        for f in folder.Folders:
            try:
                if f.DefaultItemType != 0:
                    continue
                out.append({"path": _folder_display_path(f), "items": f.Items.Count, "unread": f.UnReadItemCount})
                if level < depth:
                    walk(f, level + 1)
            except pywintypes.com_error:
                continue

    for store in stores:
        walk(store.GetRootFolder(), 1)
    return {"folders": out}


def list_messages(folder=None, limit=20, unread_only=False, since=None, include_archive=False,
                  before=None, before_entry_id=None, count_only=False):
    limit = _clamp_limit(limit)
    if folder is None:
        folder = f"{_find_account(None).DeliveryStore.DisplayName}/Inbox"
    f = _resolve_folder(folder, include_archive)
    conds, cursor = _range_conditions(unread_only, since, before, before_entry_id)
    if count_only:
        if cursor and cursor[1]:
            raise OutlookError("'before_entry_id' is not supported with count_only; use 'before' alone.")
        count = f.Items.Restrict("@SQL=" + " AND ".join(conds)).Count if conds else f.Items.Count
        return {"folder": _folder_display_path(f), "count": count}
    out = _page(_ranked(f, conds, limit, cursor), limit)
    return {"folder": _folder_display_path(f), **out}


def _text_condition(query, exhaustive):
    """DASL condition for free text. Each word (or "quoted phrase") must match, in subject or body.

    Fast mode uses Outlook's content index: subject substring match, body word-prefix / phrase match.
    Exhaustive mode scans bodies for substrings (finds mid-word matches, but takes tens of seconds).
    """
    terms = [a or b for a, b in re.findall(r'"([^"]+)"|(\S+)', query)]
    conds = []
    for t in terms[:8]:
        q = _q(t)
        phrase = " " in t
        if exhaustive:
            body = f"\"urn:schemas:httpmail:textdescription\" LIKE '%{q}%'"
        elif phrase:
            body = f"\"urn:schemas:httpmail:textdescription\" ci_phrasematch '{q}'"
        else:
            body = f"\"urn:schemas:httpmail:textdescription\" ci_startswith '{q}'"
        conds.append(f"(\"urn:schemas:httpmail:subject\" LIKE '%{q}%' OR {body})")
    return conds


def search_messages(query, folder=None, sender=None, since=None, limit=20, include_archive=False,
                    exhaustive=False, before=None, before_entry_id=None):
    limit = _clamp_limit(limit)
    has_text = bool(query and query.strip())
    conds = []  # non-text filters
    if sender:
        s = _q(sender)
        conds.append(f"(\"urn:schemas:httpmail:fromemail\" LIKE '%{s}%' OR "
                     f"\"urn:schemas:httpmail:fromname\" LIKE '%{s}%')")
    range_conds, cursor = _range_conditions(False, since, before, before_entry_id)
    conds += range_conds

    if folder:
        folders = [_resolve_folder(folder, include_archive)]
    else:
        folders = []
        for a in _accounts():
            for kind in (OL_FOLDER_INBOX, OL_FOLDER_SENT_MAIL):
                try:
                    folders.append(_default_folder(a.DeliveryStore, kind))
                except pywintypes.com_error:
                    log.warning("no default folder %s for one account", kind)
        if include_archive and CONFIG["archive_store"]:
            arch = _find_store(CONFIG["archive_store"], True)
            if arch is not None:
                folders += [f for f in _mail_folders(arch.GetRootFolder())
                            if _norm(f.Name) != "deleted items"]
    results = []
    for f in folders:
        try:
            text = _text_condition(query, exhaustive) if has_text else []
            results += _ranked(f, text + conds, limit, cursor)
        except pywintypes.com_error as e:
            if exhaustive or not has_text:
                raise
            log.warning("indexed search failed on a folder (%s); retrying exhaustively", _hr(e))
            results += _ranked(f, _text_condition(query, True) + conds, limit, cursor)
    return {"searched": [_folder_display_path(f) for f in folders], **_page(results, limit)}


def get_message(entry_id, store_id, max_chars=DEFAULT_BODY_CHARS):
    item = _get_item(entry_id, store_id)
    max_chars = max(1, int(max_chars))
    body = item.Body or ""
    to, cc = _recipients(item)
    return {
        "entry_id": item.EntryID,
        "store_id": store_id,
        "folder": _folder_display_path(item.Parent),
        "subject": item.Subject,
        "from": {"name": item.SenderName, "address": _sender_address(item)},
        "to": to,
        "cc": cc,
        "date": _iso(_when(item)),
        "unread": bool(item.UnRead),
        "body": body[:max_chars],
        "truncated": len(body) > max_chars,
        "attachments": [{"name": a.FileName, "size": a.Size} for a in item.Attachments],
    }


def _thread_topic_items(item, limit):
    topic = item.ConversationTopic
    if not topic:
        topic = re.sub(r"^\s*((re|fw|fwd|aw|wg)\s*:\s*)+", "", item.Subject or "", flags=re.I)
    store = item.Parent.Store
    found = []
    folders = [item.Parent]
    try:
        sent = store.GetDefaultFolder(OL_FOLDER_SENT_MAIL)
        if sent.EntryID != item.Parent.EntryID:
            folders.append(sent)
    except pywintypes.com_error:
        pass
    for f in folders:
        found += _query_folder(f, [f"\"urn:schemas:httpmail:thread-topic\" = '{_q(topic)}'"], limit)
    return found


def get_thread(entry_id, store_id, limit=10):
    item = _get_item(entry_id, store_id)
    limit = _clamp_limit(limit)
    msgs = _thread_topic_items(item, limit)
    msgs.sort(key=lambda m: m["received"] or "")
    return {"subject": item.ConversationTopic or item.Subject, "messages": msgs[-limit:]}


# ---------------------------------------------------------------- draft tools

DRAFT_NOTE = "Saved as a draft only; nothing was sent. The user reviews and sends it in Outlook."


def _addr_list(addrs, field):
    cleaned = [a.strip() for a in (addrs or []) if a and a.strip()]
    for a in cleaned:
        if not re.fullmatch(r"[^@\s;,<>]+@[^@\s;,<>]+", a):
            raise OutlookError(f"'{a}' in '{field}' is not a plain email address.")
    return "; ".join(cleaned)


def _text_to_html(text):
    return "<div>" + _html.escape(text).replace("\r\n", "\n").replace("\n", "<br>\n") + "</div>"


def _is_drafts_folder(folder):
    try:
        return folder.EntryID == folder.Store.GetDefaultFolder(OL_FOLDER_DRAFTS).EntryID
    except pywintypes.com_error:
        return False


def _draft_result(item):
    return {"entry_id": item.EntryID, "store_id": item.Parent.StoreID,
            "folder": _folder_display_path(item.Parent), "subject": item.Subject, "message": DRAFT_NOTE}


def create_draft(to, subject, body, cc=None, bcc=None, account=None, html=False):
    acct = _find_account(account)
    to_s = _addr_list(to, "to")
    if not to_s:
        raise OutlookError("At least one recipient is required in 'to'.")
    drafts = acct.DeliveryStore.GetDefaultFolder(OL_FOLDER_DRAFTS)
    item = drafts.Items.Add(OL_MAIL_ITEM)  # created directly in the account's Drafts folder
    item.SendUsingAccount = acct
    item.To = to_s
    item.CC = _addr_list(cc, "cc")
    item.BCC = _addr_list(bcc, "bcc")
    item.Subject = subject
    if html:
        item.HTMLBody = body
    else:
        item.Body = body
    item.Save()
    return _draft_result(item)


def create_reply_draft(entry_id, store_id, body, reply_all=False):
    orig = _get_item(entry_id, store_id)
    reply = orig.ReplyAll() if reply_all else orig.Reply()
    html_body = reply.HTMLBody or ""
    m = re.search(r"<body[^>]*>", html_body, re.I)
    if m:
        reply.HTMLBody = html_body[:m.end()] + _text_to_html(body) + "<br>" + html_body[m.end():]
    else:
        reply.Body = body + "\r\n\r\n" + (reply.Body or "")
    reply.Save()
    return _draft_result(reply)


def update_draft(entry_id, store_id, body=None, subject=None, to=None):
    item = _get_item(entry_id, store_id)
    if getattr(item, "Sent", False) or not _is_drafts_folder(item.Parent):
        raise OutlookError("Refusing to edit: this message is not in a Drafts folder. Only drafts can be updated.")
    if body is None and subject is None and to is None:
        raise OutlookError("Nothing to update: give body, subject and/or to.")
    if to is not None:
        to_s = _addr_list(to, "to")
        if not to_s:
            raise OutlookError("'to' must contain at least one address.")
        item.To = to_s
    if subject is not None:
        item.Subject = subject
    if body is not None:
        if item.BodyFormat == 2:  # HTML draft: keep it HTML
            item.HTMLBody = _text_to_html(body)
        else:
            item.Body = body
    item.Save()
    return _draft_result(item)
