"""Inspection helpers for debugging Mail.app drafts via AppleScript.

Each helper makes a small targeted AppleScript call. Multi-line content
(message body) is written to a temp file by AppleScript and read back from
Python — sidesteps any newline-parsing problems.

Read-only. Never mutates Mail state.
"""
from __future__ import annotations

import os
import sys
import tempfile
from typing import Optional

# Make the package importable when running directly from repo root.
_HERE = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.dirname(_HERE)
_SRC = os.path.join(_REPO, "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

from apple_mail_mcp.applescript import MailBridge, _as_escape  # noqa: E402


def find_drafts_with_subject(bridge: MailBridge, subject: str) -> list[int]:
    """Return all draft IDs (across all accounts' Drafts mailboxes) whose
    subject equals `subject`. Empty list if none."""
    safe = _as_escape(subject)
    script = f'''tell application "Mail"
    set out to ""
    repeat with acc in (every account)
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    try
                        if name of mb contains "raft" then
                            repeat with d in (every message of mb whose subject is "{safe}")
                                try
                                    set out to out & ((id of d) as string) & linefeed
                                end try
                            end repeat
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat
    return out
end tell'''
    raw = bridge._run_applescript(script, timeout=20) or ""
    ids = []
    for line in raw.splitlines():
        line = line.strip()
        if line and line.isdigit():
            ids.append(int(line))
    return ids


def get_subject(bridge: MailBridge, msg_id: int) -> Optional[str]:
    return _get_str_field(bridge, msg_id, "subject of foundMsg")


def get_message_id_header(bridge: MailBridge, msg_id: int) -> Optional[str]:
    return _get_str_field(bridge, msg_id, "message id of foundMsg")


def get_account_and_mailbox(bridge: MailBridge, msg_id: int) -> Optional[tuple[str, str]]:
    """Return (account_name, mailbox_name) for the message, or None."""
    script = _wrap_find_msg(msg_id, '''
    set out to ""
    try
        set out to (name of (mailbox of foundMsg)) as string
    end try
    set out2 to ""
    try
        set out2 to (name of (account of (mailbox of foundMsg))) as string
    end try
    return out2 & "<<<>>>" & out''')
    raw = bridge._run_applescript(script, timeout=20)
    if not raw or "<<<>>>" not in raw:
        return None
    acct, mbox = raw.split("<<<>>>", 1)
    return (acct.strip(), mbox.strip())


def get_recipients(bridge: MailBridge, msg_id: int, kind: str) -> list[str]:
    """kind ∈ {'to', 'cc', 'bcc'}. Returns ['Name <addr>', ...]."""
    assert kind in ("to", "cc", "bcc")
    script = _wrap_find_msg(msg_id, f'''
    set out to ""
    try
        repeat with r in (every {kind} recipient of foundMsg)
            try
                set nm to (name of r) as string
            on error
                set nm to ""
            end try
            try
                set ad to (address of r) as string
            on error
                set ad to ""
            end try
            if nm is not "" then
                set out to out & nm & " <" & ad & ">" & linefeed
            else
                set out to out & ad & linefeed
            end if
        end repeat
    end try
    return out''')
    raw = bridge._run_applescript(script, timeout=20) or ""
    return [ln for ln in (l.strip() for l in raw.splitlines()) if ln]


def get_headers(bridge: MailBridge, msg_id: int) -> list[tuple[str, str]]:
    """Return list of (header_name, header_value) for the message.
    Uses '<<<:>>>' as the separator (very unlikely in real header content)."""
    script = _wrap_find_msg(msg_id, '''
    set out to ""
    try
        repeat with h in (every header of foundMsg)
            try
                set nm to (name of h) as string
            on error
                set nm to ""
            end try
            try
                set ct to (content of h) as string
            on error
                set ct to ""
            end try
            set out to out & nm & "<<<:>>>" & ct & linefeed
        end repeat
    end try
    return out''')
    raw = bridge._run_applescript(script, timeout=20) or ""
    headers = []
    for line in raw.splitlines():
        if "<<<:>>>" in line:
            name, val = line.split("<<<:>>>", 1)
            headers.append((name.strip(), val.strip()))
    return headers


def get_body(bridge: MailBridge, msg_id: int) -> Optional[str]:
    """Return the full body content of the message. Uses a temp file to avoid
    string-marshalling issues with multi-line content."""
    fd, tmp_path = tempfile.mkstemp(prefix="mail_body_", suffix=".txt")
    os.close(fd)
    safe_path = tmp_path.replace('"', '\\"')
    script = _wrap_find_msg(msg_id, f'''
    try
        set bodyText to (content of foundMsg) as string
    on error
        set bodyText to ""
    end try
    set fileRef to open for access POSIX file "{safe_path}" with write permission
    set eof of fileRef to 0
    try
        write bodyText as «class utf8» to fileRef
    on error
        try
            write bodyText to fileRef
        end try
    end try
    close access fileRef
    return "OK"''')
    result = bridge._run_applescript(script, timeout=30)
    body = None
    if result == "OK":
        try:
            with open(tmp_path, "rb") as f:
                raw = f.read()
            body = raw.decode("utf-8", errors="replace")
        except OSError:
            body = None
    try:
        os.unlink(tmp_path)
    except OSError:
        pass
    return body


def msg_exists(bridge: MailBridge, msg_id: int) -> bool:
    """Quick check: does a message with this id exist anywhere?"""
    script = f'''tell application "Mail"
    set found to false
    repeat with acc in (every account)
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    try
                        if (count of (every message of mb whose id is {msg_id})) > 0 then
                            set found to true
                            exit repeat
                        end if
                    end try
                end repeat
            end if
        end try
        if found then exit repeat
    end repeat
    if found then
        return "yes"
    else
        return "no"
    end if
end tell'''
    return (bridge._run_applescript(script, timeout=15) or "").strip() == "yes"


# --- internal helpers ---

def _wrap_find_msg(msg_id: int, body: str) -> str:
    """Wrap a snippet that uses `foundMsg` as the message reference."""
    return f'''tell application "Mail"
    set foundMsg to missing value
    repeat with acc in (every account)
        if foundMsg is not missing value then exit repeat
        try
            if enabled of acc is true then
                repeat with mb in (every mailbox of acc)
                    if foundMsg is not missing value then exit repeat
                    try
                        set tgts to (every message of mb whose id is {msg_id})
                        if (count of tgts) > 0 then
                            set foundMsg to item 1 of tgts
                        end if
                    end try
                end repeat
            end if
        end try
    end repeat
    if foundMsg is missing value then return ""
{body}
end tell'''


def _get_str_field(bridge: MailBridge, msg_id: int, prop_expr: str) -> Optional[str]:
    """Helper: get a single string-valued property as `prop_expr` (e.g.
    'subject of foundMsg'). Returns None if not found / property missing."""
    script = _wrap_find_msg(msg_id, f'''
    set out to ""
    try
        set out to ({prop_expr}) as string
    end try
    return out''')
    raw = bridge._run_applescript(script, timeout=15)
    if raw is None:
        return None
    raw = raw.strip()
    return raw or None


def dump(bridge: MailBridge, msg_id: int) -> dict:
    """Return a dict with all interesting fields of the draft."""
    if not msg_exists(bridge, msg_id):
        return {"id": msg_id, "exists": False}
    acct_mbox = get_account_and_mailbox(bridge, msg_id)
    headers = get_headers(bridge, msg_id)
    return {
        "id": msg_id,
        "exists": True,
        "subject": get_subject(bridge, msg_id),
        "message_id_header": get_message_id_header(bridge, msg_id),
        "account": acct_mbox[0] if acct_mbox else None,
        "mailbox": acct_mbox[1] if acct_mbox else None,
        "to": get_recipients(bridge, msg_id, "to"),
        "cc": get_recipients(bridge, msg_id, "cc"),
        "bcc": get_recipients(bridge, msg_id, "bcc"),
        "headers": headers,
        "in_reply_to": next((v for n, v in headers if n.lower() == "in-reply-to"), None),
        "references": next((v for n, v in headers if n.lower() == "references"), None),
        "body": get_body(bridge, msg_id),
    }
