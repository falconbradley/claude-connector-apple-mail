"""
Bridge to Apple Mail via JXA (JavaScript for Automation).

Replaces direct SQLite / .emlx filesystem access with a clean scripting
interface through Mail.app.  Every method builds a JXA script string, executes
it via ``osascript -l JavaScript``, and parses the JSON that comes back.

Requirements
------------
- Mail.app must be running.
- The calling process (Claude Desktop, Terminal, etc.) must have Automation
  permission for Mail.app in System Settings -> Privacy & Security -> Automation.
"""

from __future__ import annotations

import json
import logging
import re
import subprocess
import tempfile
import time
from datetime import datetime
from email.utils import parseaddr as _parseaddr
from pathlib import Path
from typing import Any, Optional

logger = logging.getLogger("apple_mail_mcp.applescript")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_RE_PREFIX = re.compile(r"^(Re|Fwd|Fw)\s*:\s*", re.IGNORECASE)

# Flag color name → flagIndex integer (0 = no flag, 1–7 = red … gray).
# Confirmed via JXA probe: msg.flagIndex() returns 1 for a red-flagged message.
_FLAG_COLOR_MAP: dict[str, int] = {
    "red": 1, "orange": 2, "yellow": 3,
    "green": 4, "blue": 5, "purple": 6, "gray": 7,
}
_FLAG_COLOR_ORDER = ["red", "orange", "yellow", "green", "blue", "purple", "gray"]


def _strip_subject_prefixes(subj: str) -> str:
    """Remove leading Re:/Fwd:/Fw: prefixes to get the base subject."""
    prev = None
    while prev != subj:
        prev = subj
        subj = _RE_PREFIX.sub("", subj).strip()
    return subj


def _parse_address(addr: str) -> tuple[str, str]:
    """Split 'Name <email>' into (name, email). Falls back to ('', addr)."""
    name, email = _parseaddr(addr)
    return name or "", email or addr


def _js_escape(value: str) -> str:
    """Escape a Python string for safe embedding inside a JS string literal."""
    return (
        value
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("'", "\\'")
        .replace("\n", "\\n")
        .replace("\r", "\\r")
        .replace("\t", "\\t")
    )


# ---------------------------------------------------------------------------
# MailBridge
# ---------------------------------------------------------------------------

class MailBridge:
    """Bridge to Mail.app via JXA (JavaScript for Automation)."""

    _message_cache: dict[int, tuple[str, str, Optional[int]]]  # msg_id -> (account, mailbox, index)
    _nonempty_mailboxes: set[tuple[str, str]]  # (account_name, mailbox_name)

    def __init__(self) -> None:
        """Verify Mail.app is running. Raise RuntimeError with helpful message if not."""
        self._message_cache = {}
        self._nonempty_mailboxes = set()

        # Check if Mail.app is running and pre-scan mailbox counts in one call.
        # This avoids per-mailbox IMAP queries during search_messages.
        init_script = """
        (function() {
            var se = Application("System Events");
            var procs = se.processes.whose({name: "Mail"});
            if (procs.length === 0) return JSON.stringify({"running": false});

            var mail = Application("Mail");
            var accounts = mail.accounts();
            var mboxes = [];
            for (var i = 0; i < accounts.length; i++) {
                var acctName = accounts[i].name();
                var mbs = accounts[i].mailboxes();
                for (var j = 0; j < mbs.length; j++) {
                    var mb = mbs[j];
                    var mc = mb.messages.length;
                    if (mc > 0) {
                        mboxes.push({"account": acctName, "mailbox": mb.name(), "count": mc});
                    }
                }
            }
            return JSON.stringify({"running": true, "nonempty": mboxes});
        })();
        """
        try:
            result = self._run_jxa(init_script, timeout=120)
        except RuntimeError:
            raise RuntimeError(
                "Mail.app is not running. Please open Mail.app and try again."
            )

        if not result or not result.get("running", False):
            raise RuntimeError(
                "Mail.app is not running. Please open Mail.app and try again."
            )

        # Cache non-empty mailboxes to skip slow IMAP queries during search
        for mb in result.get("nonempty", []):
            self._nonempty_mailboxes.add((mb["account"], mb["mailbox"]))

        logger.info(
            "MailBridge initialised — %d non-empty mailboxes.",
            len(self._nonempty_mailboxes),
        )

    # ------------------------------------------------------------------
    # JXA execution
    # ------------------------------------------------------------------

    def _run_jxa(self, script: str, timeout: int = 30) -> Any:
        """Execute JXA script via osascript, return parsed JSON.

        The script MUST produce a JSON string as its final expression
        (typically via ``JSON.stringify(...)``).

        Raises:
            RuntimeError: on timeout, permission error, or non-zero exit.
        """
        truncated = script[:200].replace("\n", " ")
        logger.debug("Running JXA: %s ...", truncated)

        t0 = time.monotonic()
        try:
            # Write script to temp file — avoids arg-length limits and is
            # measurably faster than passing large scripts via -e.
            with tempfile.NamedTemporaryFile(
                mode="w", suffix=".js", delete=False, prefix="jxa_"
            ) as f:
                f.write(script)
                script_path = f.name
            try:
                proc = subprocess.run(
                    ["osascript", "-l", "JavaScript", script_path],
                    capture_output=True,
                    text=True,
                    timeout=timeout,
                )
            finally:
                Path(script_path).unlink(missing_ok=True)
        except subprocess.TimeoutExpired:
            elapsed = time.monotonic() - t0
            logger.warning("JXA script timed out after %.1fs", elapsed)
            raise RuntimeError(
                f"Mail.app is not responding (timed out after {timeout}s). "
                "It may be busy or frozen — try again in a moment."
            )

        elapsed = time.monotonic() - t0

        if proc.returncode != 0:
            stderr = proc.stderr.strip()
            logger.warning(
                "JXA failed (rc=%d, %.1fs): %s", proc.returncode, elapsed, stderr
            )
            if "not running" in stderr.lower():
                raise RuntimeError(
                    "Mail.app is not running. Please open Mail.app and try again."
                )
            if "not allowed" in stderr.lower() or "permission" in stderr.lower():
                raise RuntimeError(
                    "Automation permission denied. Grant permission in "
                    "System Settings -> Privacy & Security -> Automation, "
                    "then try again."
                )
            raise RuntimeError(f"JXA script failed: {stderr}")

        stdout = proc.stdout.strip()
        if not stdout:
            logger.debug("JXA returned empty output (%.1fs)", elapsed)
            return None

        try:
            parsed = json.loads(stdout)
        except json.JSONDecodeError as exc:
            logger.warning("JXA output is not valid JSON: %s", exc)
            raise RuntimeError(f"Unexpected output from Mail.app: {stdout[:300]}")

        logger.debug("JXA completed in %.1fs", elapsed)
        return parsed

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_stats(self) -> dict:
        """Return overall statistics.

        Returns:
            {"total_messages": int, "unread_messages": int,
             "mailbox_count": int, "account_count": int}
        """
        script = """
        (function() {
            var mail = Application("Mail");
            var accounts = mail.accounts();
            var totalMessages = 0;
            var unreadMessages = 0;
            var mailboxCount = 0;

            for (var i = 0; i < accounts.length; i++) {
                var mboxes = accounts[i].mailboxes();
                mailboxCount += mboxes.length;
                for (var j = 0; j < mboxes.length; j++) {
                    var mc = mboxes[j].messages.length;
                    totalMessages += mc;
                    unreadMessages += mboxes[j].unreadCount();
                }
            }
            return JSON.stringify({
                "total_messages": totalMessages,
                "unread_messages": unreadMessages,
                "mailbox_count": mailboxCount,
                "account_count": accounts.length
            });
        })();
        """
        result = self._run_jxa(script, timeout=60)
        if result is None:
            return {
                "total_messages": 0,
                "unread_messages": 0,
                "mailbox_count": 0,
                "account_count": 0,
            }
        return result

    def list_mailboxes(self) -> list[dict]:
        """List every mailbox across all accounts.

        Returns:
            [{"name": str, "account_name": str, "unread_count": int,
              "message_count": int}, ...]
        """
        script = """
        (function() {
            var mail = Application("Mail");
            var accounts = mail.accounts();
            var result = [];
            for (var i = 0; i < accounts.length; i++) {
                var acctName = accounts[i].name();
                var mboxes = accounts[i].mailboxes();
                for (var j = 0; j < mboxes.length; j++) {
                    var mb = mboxes[j];
                    result.push({
                        "name": mb.name(),
                        "account_name": acctName,
                        "unread_count": mb.unreadCount(),
                        "message_count": mb.messages.length
                    });
                }
            }
            return JSON.stringify(result);
        })();
        """
        result = self._run_jxa(script, timeout=60)
        return result if isinstance(result, list) else []

    def search_messages(
        self,
        *,
        mailbox_name: Optional[str] = None,
        account_name: Optional[str] = None,
        subject_contains: Optional[str] = None,
        sender_contains: Optional[str] = None,
        to_address_contains: Optional[str] = None,
        since: Optional[datetime] = None,
        before: Optional[datetime] = None,
        is_unread: Optional[bool] = None,
        is_flagged: Optional[bool] = None,
        has_attachments: Optional[bool] = None,
        limit: int = 50,
        offset: int = 0,
    ) -> tuple[int, list[dict]]:
        """Search messages across mailboxes.

        Returns:
            (total_matching, page_of_results)

        Two-round bulk-fetch approach:
          Round 1: bulk fetch ids + dates (+ filter properties) for ALL
                   non-empty mailboxes, filter in JS, sort, paginate.
          Round 2: bulk fetch remaining output properties for ONLY the
                   mailboxes that contain page results. Extract specific
                   message indices to complete the result set.

        In "query" mode (subject_contains == sender_contains), matches
        subject OR sender.  When set separately, both must match (AND).
        """
        # Account / mailbox JS filters
        acct_filter = ""
        if account_name:
            safe_acct = _js_escape(account_name)
            acct_filter = (
                f'if (acctName.toLowerCase().indexOf("{safe_acct}".toLowerCase()) === -1) continue;'
            )
        mbox_filter = ""
        if mailbox_name:
            safe_mbox = _js_escape(mailbox_name)
            mbox_filter = (
                f'if (mbName.toLowerCase().indexOf("{safe_mbox}".toLowerCase()) === -1) continue;'
            )

        # Date JS vars
        since_js = f'var sinceDate = new Date("{since.isoformat()}");' if since else ""
        before_js = f'var beforeDate = new Date("{before.isoformat()}");' if before else ""

        # Phase 1: determine which bulk properties to fetch and which
        # JS post-filters to apply.  Always fetch ids + dates.  Conditionally
        # fetch subjects/senders/flags only when needed for filtering.
        phase1_filters: list[str] = []
        bulk_fetches = ["var ids = msgs.id();", "var dates = msgs.dateReceived();"]
        loop_vars: list[str] = []

        # Date filters
        if since:
            phase1_filters.append("if (!dr || dr < sinceDate) continue;")
        if before:
            phase1_filters.append("if (!dr || dr > beforeDate) continue;")

        # Text filters — bulk fetch subjects/senders only when needed
        query_mode = (
            subject_contains
            and sender_contains
            and subject_contains == sender_contains
        )
        if query_mode:
            safe_q = _js_escape(subject_contains)
            phase1_filters.append(
                f'if (subj.toLowerCase().indexOf("{safe_q}".toLowerCase()) === -1 '
                f'&& sndr.toLowerCase().indexOf("{safe_q}".toLowerCase()) === -1) continue;'
            )
            bulk_fetches.append("var subjects = msgs.subject();")
            bulk_fetches.append("var senders = msgs.sender();")
            loop_vars.append('var subj = subjects[k] || "";')
            loop_vars.append('var sndr = senders[k] || "";')
        else:
            if subject_contains:
                safe_subj = _js_escape(subject_contains)
                phase1_filters.append(
                    f'if (subj.toLowerCase().indexOf("{safe_subj}".toLowerCase()) === -1) continue;'
                )
                bulk_fetches.append("var subjects = msgs.subject();")
                loop_vars.append('var subj = subjects[k] || "";')
            if sender_contains:
                safe_sndr = _js_escape(sender_contains)
                phase1_filters.append(
                    f'if (sndr.toLowerCase().indexOf("{safe_sndr}".toLowerCase()) === -1) continue;'
                )
                bulk_fetches.append("var senders = msgs.sender();")
                loop_vars.append('var sndr = senders[k] || "";')

        # Recipient (To/CC) filter
        if to_address_contains:
            safe_recip = _js_escape(to_address_contains.lower())
            bulk_fetches.append("var toAddrs = msgs.toRecipients.address();")
            bulk_fetches.append("var ccAddrs = msgs.ccRecipients.address();")
            loop_vars.append("var _ta = toAddrs[k] || [];")
            loop_vars.append("var _ca = ccAddrs[k] || [];")
            phase1_filters.append(
                f'var _recipMatch = false;'
                f' for (var _t = 0; _t < _ta.length; _t++) {{'
                f' if (_ta[_t] && _ta[_t].toLowerCase().indexOf("{safe_recip}") !== -1) {{ _recipMatch = true; break; }}'
                f' }}'
                f' if (!_recipMatch) {{ for (var _c = 0; _c < _ca.length; _c++) {{'
                f' if (_ca[_c] && _ca[_c].toLowerCase().indexOf("{safe_recip}") !== -1) {{ _recipMatch = true; break; }}'
                f' }} }}'
                f' if (!_recipMatch) continue;'
            )

        # Flag filters
        if is_unread is True:
            phase1_filters.append("if (readFlags[k]) continue;")
            bulk_fetches.append("var readFlags = msgs.readStatus();")
        elif is_unread is False:
            phase1_filters.append("if (!readFlags[k]) continue;")
            bulk_fetches.append("var readFlags = msgs.readStatus();")

        if is_flagged is True:
            phase1_filters.append("if (!flagFlags[k]) continue;")
            bulk_fetches.append("var flagFlags = msgs.flaggedStatus();")
        elif is_flagged is False:
            phase1_filters.append("if (flagFlags[k]) continue;")
            bulk_fetches.append("var flagFlags = msgs.flaggedStatus();")

        bulk_fetch_js = "\n                        ".join(bulk_fetches)
        loop_vars_js = "\n                            ".join(loop_vars)
        filter_js = "\n                            ".join(phase1_filters)

        # Build a JS set of non-empty mailbox keys to skip slow IMAP queries
        nonempty_keys = [
            f'"{_js_escape(a)}|{_js_escape(m)}"'
            for a, m in self._nonempty_mailboxes
        ]
        nonempty_set_js = "var _ne = {" + ",".join(
            f"{k}: 1" for k in nonempty_keys
        ) + "};"

        # ---------------------------------------------------------------
        # Round 1: filter + paginate using minimal bulk fetches
        # ---------------------------------------------------------------
        script_r1 = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts();
            {since_js}
            {before_js}
            {nonempty_set_js}

            var candidates = [];
            var _timing = [];
            var scanStart = Date.now();

            for (var i = 0; i < accounts.length; i++) {{
                var acctName = accounts[i].name();
                {acct_filter}
                var mboxes = accounts[i].mailboxes();

                for (var j = 0; j < mboxes.length; j++) {{
                    var mbName = mboxes[j].name();
                    {mbox_filter}
                    if (!_ne[acctName + "|" + mbName]) continue;

                    try {{
                        var t0 = Date.now();
                        var msgs = mboxes[j].messages;
                        {bulk_fetch_js}

                        for (var k = 0; k < ids.length; k++) {{
                            var dr = dates[k];
                            {loop_vars_js}
                            {filter_js}
                            var _c = {{
                                id: ids[k], date: dr, msgIdx: k,
                                acctName: acctName, mbName: mbName
                            }};
                            if (typeof subj !== "undefined") _c.subj = subj;
                            if (typeof sndr !== "undefined") _c.sndr = sndr;
                            if (typeof readFlags !== "undefined") _c.read = readFlags[k] ? true : false;
                            if (typeof flagFlags !== "undefined") _c.flag = flagFlags[k] ? true : false;
                            candidates.push(_c);
                        }}
                        _timing.push({{mbox: acctName + "/" + mbName, msgs: ids.length, ms: Date.now() - t0}});
                    }} catch(e) {{}}
                }}
            }}

            candidates.sort(function(a, b) {{
                if (!a.date && !b.date) return 0;
                if (!a.date) return 1;
                if (!b.date) return -1;
                return b.date - a.date;
            }});

            var total = candidates.length;
            var page = candidates.slice({offset}, {offset + limit});

            return JSON.stringify({{
                "total": total,
                "page": page,
                "_timing": _timing,
                "_scan_ms": Date.now() - scanStart
            }});
        }})();
        """

        r1 = self._run_jxa(script_r1, timeout=300)
        if r1 is None:
            return 0, []

        total: int = r1.get("total", 0)
        page: list[dict] = r1.get("page", [])

        # Log Round 1 timing
        timing = r1.get("_timing", [])
        scan_ms = r1.get("_scan_ms")
        if timing:
            parts = [f"{t['mbox']}({t['msgs']}msgs, {t['ms']}ms)" for t in timing]
            logger.info("Search R1: %s", ", ".join(parts))
        if scan_ms is not None:
            logger.info("Search R1 total: %d candidates in %dms", total, scan_ms)

        if not page:
            return total, []

        # ---------------------------------------------------------------
        # Round 2: bulk-fetch display properties for page mailboxes only.
        # Fetches subjects, senders, readStatus, flaggedStatus from only
        # the mailboxes that contain page results.  Skips nice-to-have
        # properties (dateSent, messageSize, messageId) to keep fast —
        # those are available via get_email when needed.
        # ---------------------------------------------------------------
        # Group page items by mailbox
        mbox_groups: dict[str, list[dict]] = {}
        for item in page:
            key = f"{item['acctName']}|{item['mbName']}"
            mbox_groups.setdefault(key, []).append(item)

        # Build targets for Round 2
        targets_js_parts: list[str] = []
        for key, items in mbox_groups.items():
            acct_name_r2, mbox_name_r2 = key.split("|", 1)
            indices = [str(item["msgIdx"]) for item in items]
            targets_js_parts.append(
                f'{{"acct": "{_js_escape(acct_name_r2)}", '
                f'"mbox": "{_js_escape(mbox_name_r2)}", '
                f'"indices": [{",".join(indices)}]}}'
            )
        targets_js = "[" + ",".join(targets_js_parts) + "]"

        script_r2 = f"""
        (function() {{
            var mail = Application("Mail");
            var targets = {targets_js};
            var results = {{}};

            for (var t = 0; t < targets.length; t++) {{
                var tgt = targets[t];
                try {{
                    var accts = mail.accounts.whose({{name: tgt.acct}});
                    if (accts.length === 0) continue;
                    var mboxes = accts[0].mailboxes.whose({{name: tgt.mbox}});
                    if (mboxes.length === 0) continue;
                    var msgs = mboxes[0].messages;
                    var _subj = msgs.subject();
                    var _sndr = msgs.sender();
                    var _read = msgs.readStatus();
                    var _flag = msgs.flaggedStatus();
                    for (var q = 0; q < tgt.indices.length; q++) {{
                        var idx = tgt.indices[q];
                        var key = tgt.acct + "|" + tgt.mbox + "|" + idx;
                        results[key] = {{
                            "subj": _subj[idx] || "",
                            "sndr": _sndr[idx] || "",
                            "read": _read[idx] ? true : false,
                            "flag": _flag[idx] ? true : false
                        }};
                    }}
                }} catch(e) {{}}
            }}
            return JSON.stringify(results);
        }})();
        """

        r2 = self._run_jxa(script_r2, timeout=300)
        if r2 is None:
            r2 = {}

        # ---------------------------------------------------------------
        # Merge Round 1 + Round 2 into final results
        # ---------------------------------------------------------------
        results: list[dict] = []
        for item in page:
            r2_key = f"{item['acctName']}|{item['mbName']}|{item['msgIdx']}"
            extra = r2.get(r2_key, {})
            # date is already ISO string from JSON.stringify of JS Date
            dr_raw = item.get("date")
            dr_str = dr_raw if isinstance(dr_raw, str) else None
            # Prefer Round 1 data when available, fall back to Round 2
            results.append({
                "id": item["id"],
                "msg_idx": item["msgIdx"],
                "subject": item.get("subj") or extra.get("subj", ""),
                "sender": item.get("sndr") or extra.get("sndr", ""),
                "date_received": dr_str,
                "date_sent": dr_str,  # approximate; use get_email for exact
                "is_read": item.get("read") if "read" in item else extra.get("read", True),
                "is_flagged": item.get("flag") if "flag" in item else extra.get("flag", False),
                "has_attachments": False,
                "mailbox_name": item["mbName"],
                "account_name": item["acctName"],
                "message_id": None,
                "in_reply_to": None,
                "size": 0,
            })

        # Update message cache
        for msg in results:
            msg_id = msg.get("id")
            if msg_id is not None:
                self._message_cache[msg_id] = (
                    msg.get("account_name", ""),
                    msg.get("mailbox_name", ""),
                    msg.get("msg_idx"),
                )

        return total, results

    def get_message_id_header(self, message_id: int) -> Optional[str]:
        """Get just the RFC 2822 Message-ID header for a message.

        Returns the bare Message-ID string (no angle brackets) or None.
        """
        location = self._find_message(message_id)
        if location is None:
            return None

        acct_name, mbox_name, msg_idx = location
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)

        if msg_idx is not None:
            msg_lookup_js = f"""
            var msg = mb.messages[{msg_idx}];
            if (msg.id() !== {message_id}) {{
                var msgs = mb.messages.whose({{id: {message_id}}});
                if (msgs.length === 0) return JSON.stringify(null);
                msg = msgs[0];
            }}"""
        else:
            msg_lookup_js = f"""
            var msgs = mb.messages.whose({{id: {message_id}}});
            if (msgs.length === 0) return JSON.stringify(null);
            var msg = msgs[0];"""

        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
            if (accounts.length === 0) return JSON.stringify(null);
            var acct = accounts[0];
            var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
            if (mboxes.length === 0) return JSON.stringify(null);
            var mb = mboxes[0];
            {msg_lookup_js}
            return JSON.stringify(msg.messageId());
        }})();
        """
        result = self._run_jxa(script, timeout=15)
        return result if isinstance(result, str) else None

    def get_message(self, message_id: int) -> Optional[dict]:
        """Get a single message by Mail.app id.

        Returns search-format dict plus ``body_text``, ``to_recipients``,
        and ``cc_recipients``, or None if not found.
        """
        location = self._find_message(message_id)
        if location is None:
            logger.warning("Message %d not found in any mailbox.", message_id)
            return None

        acct_name, mbox_name, msg_idx = location
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)

        # Use direct index access when available (from search cache),
        # fall back to whose({id:}) for messages found via _find_message.
        if msg_idx is not None:
            msg_lookup_js = f"""
            var msg = mb.messages[{msg_idx}];
            if (msg.id() !== {message_id}) {{
                var msgs = mb.messages.whose({{id: {message_id}}});
                if (msgs.length === 0) return JSON.stringify(null);
                msg = msgs[0];
            }}"""
        else:
            msg_lookup_js = f"""
            var msgs = mb.messages.whose({{id: {message_id}}});
            if (msgs.length === 0) return JSON.stringify(null);
            var msg = msgs[0];"""

        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
            if (accounts.length === 0) return JSON.stringify(null);
            var acct = accounts[0];
            var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
            if (mboxes.length === 0) return JSON.stringify(null);
            var mb = mboxes[0];
            {msg_lookup_js}

            var dr = msg.dateReceived();
            var ds = msg.dateSent();

            // Recipients
            var toRecips = [];
            try {{
                var toR = msg.toRecipients();
                for (var i = 0; i < toR.length; i++) {{
                    var addr = toR[i].address();
                    var nm = toR[i].name();
                    toRecips.push(nm ? (nm + " <" + addr + ">") : addr);
                }}
            }} catch(e) {{}}

            var ccRecips = [];
            try {{
                var ccR = msg.ccRecipients();
                for (var i = 0; i < ccR.length; i++) {{
                    var addr = ccR[i].address();
                    var nm = ccR[i].name();
                    ccRecips.push(nm ? (nm + " <" + addr + ">") : addr);
                }}
            }} catch(e) {{}}

            var bodyText = "";
            try {{
                bodyText = msg.content() || "";
            }} catch(e) {{}}

            var msgIdHeader = null;
            try {{
                msgIdHeader = msg.messageId();
            }} catch(e) {{}}

            var attachCount = 0;
            try {{
                attachCount = msg.mailAttachments.length;
            }} catch(e) {{}}

            return JSON.stringify({{
                "id": msg.id(),
                "subject": msg.subject() || "",
                "sender": msg.sender() || "",
                "date_received": dr ? dr.toISOString() : null,
                "date_sent": ds ? ds.toISOString() : null,
                "is_read": msg.readStatus() ? true : false,
                "is_flagged": msg.flaggedStatus() ? true : false,
                "has_attachments": attachCount > 0,
                "mailbox_name": "{safe_mbox}",
                "account_name": "{safe_acct}",
                "message_id": msgIdHeader,
                "in_reply_to": null,
                "size": msg.messageSize() || 0,
                "body_text": bodyText,
                "to_recipients": toRecips,
                "cc_recipients": ccRecips
            }});
        }})();
        """
        return self._run_jxa(script, timeout=30)

    def get_message_source(self, message_id: int) -> Optional[str]:
        """Get raw RFC 2822 source of a message (for HTML extraction)."""
        location = self._find_message(message_id)
        if location is None:
            logger.warning("Message %d not found for source retrieval.", message_id)
            return None

        acct_name, mbox_name, _ = location
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)

        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
            if (accounts.length === 0) return JSON.stringify(null);
            var acct = accounts[0];
            var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
            if (mboxes.length === 0) return JSON.stringify(null);
            var mb = mboxes[0];

            var msgs = mb.messages.whose({{id: {message_id}}});
            if (msgs.length === 0) return JSON.stringify(null);

            var src = msgs[0].source();
            return JSON.stringify({{"source": src}});
        }})();
        """
        result = self._run_jxa(script, timeout=30)
        if result is None:
            return None
        return result.get("source")

    def get_thread_messages(self, message_id: int) -> list[dict]:
        """Get all messages in same conversation thread.

        Strategy: find the target message's subject, strip Re:/Fwd: prefixes,
        then search the same mailbox for messages with the same base subject.
        Results are sorted chronologically (oldest first).
        """
        # Get the target message first
        msg = self.get_message(message_id)
        if msg is None:
            return []

        base_subject = _strip_subject_prefixes(msg.get("subject", ""))
        if not base_subject:
            return [msg]

        acct_name = msg.get("account_name", "")
        mbox_name = msg.get("mailbox_name", "")
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)
        safe_subj = _js_escape(base_subject)

        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
            if (accounts.length === 0) return JSON.stringify([]);
            var acct = accounts[0];
            var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
            if (mboxes.length === 0) return JSON.stringify([]);
            var mb = mboxes[0];

            var matching = mb.messages.whose({{subject: {{_contains: "{safe_subj}"}}}});
            var count = matching.length;
            if (count === 0) return JSON.stringify([]);

            // Cap at 200 to avoid slowness
            if (count > 200) count = 200;

            var ids = matching.id();
            var subjects = matching.subject();
            var senders = matching.sender();
            var datesRecv = matching.dateReceived();
            var datesSent = matching.dateSent();
            var readFlags = matching.readStatus();
            var flagFlags = matching.flaggedStatus();
            var sizes = matching.messageSize();
            var msgIds;
            try {{
                msgIds = matching.messageId();
            }} catch(e) {{
                msgIds = [];
            }}

            var results = [];
            for (var k = 0; k < count; k++) {{
                var subj = subjects[k] || "";
                // Strip Re:/Fwd: and compare base
                var base = subj.replace(/^(Re|Fwd|Fw)\\s*:\\s*/gi, "").trim();
                // Repeat stripping
                var prev = "";
                while (prev !== base) {{
                    prev = base;
                    base = base.replace(/^(Re|Fwd|Fw)\\s*:\\s*/gi, "").trim();
                }}
                if (base !== "{safe_subj}") continue;

                var dr = datesRecv[k];
                var ds = datesSent[k];
                results.push({{
                    "id": ids[k],
                    "subject": subj,
                    "sender": senders[k] || "",
                    "date_received": dr ? dr.toISOString() : null,
                    "date_sent": ds ? ds.toISOString() : null,
                    "is_read": readFlags[k] ? true : false,
                    "is_flagged": flagFlags[k] ? true : false,
                    "has_attachments": false,
                    "mailbox_name": "{safe_mbox}",
                    "account_name": "{safe_acct}",
                    "message_id": (msgIds && msgIds.length > k) ? (msgIds[k] || null) : null,
                    "in_reply_to": null,
                    "size": sizes[k] || 0
                }});
            }}

            // Sort chronologically (oldest first)
            results.sort(function(a, b) {{
                if (!a.date_received && !b.date_received) return 0;
                if (!a.date_received) return -1;
                if (!b.date_received) return 1;
                return new Date(a.date_received) - new Date(b.date_received);
            }});

            return JSON.stringify(results);
        }})();
        """
        result = self._run_jxa(script, timeout=60)
        if not isinstance(result, list) or len(result) == 0:
            return [msg]
        return result

    def list_attachments(self, message_id: int) -> list[dict]:
        """List attachments for a message.

        Returns:
            [{"index": int, "name": str, "mime_type": str, "file_size": int}]
        """
        location = self._find_message(message_id)
        if location is None:
            logger.warning("Message %d not found for attachment listing.", message_id)
            return []

        acct_name, mbox_name, _ = location
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)

        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
            if (accounts.length === 0) return JSON.stringify([]);
            var acct = accounts[0];
            var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
            if (mboxes.length === 0) return JSON.stringify([]);
            var mb = mboxes[0];

            var msgs = mb.messages.whose({{id: {message_id}}});
            if (msgs.length === 0) return JSON.stringify([]);
            var msg = msgs[0];

            var atts = msg.mailAttachments();
            var result = [];
            for (var i = 0; i < atts.length; i++) {{
                var att = atts[i];
                result.push({{
                    "index": i,
                    "name": att.name() || ("attachment_" + i),
                    "mime_type": att.mimeType() || "application/octet-stream",
                    "file_size": att.fileSize() || 0
                }});
            }}
            return JSON.stringify(result);
        }})();
        """
        result = self._run_jxa(script, timeout=30)
        return result if isinstance(result, list) else []

    def get_attachment(
        self, message_id: int, attachment_index: int
    ) -> Optional[tuple[str, str, bytes]]:
        """Save attachment to temp file, read it, return (filename, mime_type, raw_bytes).

        Returns None if the message or attachment is not found.
        """
        location = self._find_message(message_id)
        if location is None:
            logger.warning("Message %d not found for attachment download.", message_id)
            return None

        acct_name, mbox_name, _ = location
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)

        with tempfile.TemporaryDirectory(prefix="apple_mail_att_") as tmpdir:
            safe_tmpdir = _js_escape(tmpdir)

            # First get the attachment metadata and save it
            script = f"""
            (function() {{
                var mail = Application("Mail");
                var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
                if (accounts.length === 0) return JSON.stringify(null);
                var acct = accounts[0];
                var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
                if (mboxes.length === 0) return JSON.stringify(null);
                var mb = mboxes[0];

                var msgs = mb.messages.whose({{id: {message_id}}});
                if (msgs.length === 0) return JSON.stringify(null);
                var msg = msgs[0];

                var atts = msg.mailAttachments();
                if ({attachment_index} >= atts.length) return JSON.stringify(null);

                var att = atts[{attachment_index}];
                var fileName = att.name() || "attachment_{attachment_index}";
                var mimeType = att.mimeType() || "application/octet-stream";

                var savePath = "{safe_tmpdir}/" + fileName;
                mail.save(att, {{in: Path(savePath)}});

                return JSON.stringify({{
                    "filename": fileName,
                    "mime_type": mimeType,
                    "saved_path": savePath
                }});
            }})();
            """
            result = self._run_jxa(script, timeout=60)
            if result is None:
                return None

            filename: str = result.get("filename", f"attachment_{attachment_index}")
            mime_type: str = result.get("mime_type", "application/octet-stream")
            saved_path: str = result.get("saved_path", "")

            if not saved_path:
                logger.warning("No saved_path returned for attachment.")
                return None

            path = Path(saved_path)
            if not path.exists():
                # Try to find the file in the temp dir (name might differ)
                files = list(Path(tmpdir).iterdir())
                if files:
                    path = files[0]
                else:
                    logger.warning("Attachment file not found at %s", saved_path)
                    return None

            try:
                raw_bytes = path.read_bytes()
            except OSError as exc:
                logger.warning("Failed to read saved attachment: %s", exc)
                return None

            logger.info(
                "Retrieved attachment %r (%s, %d bytes) from message %d",
                filename,
                mime_type,
                len(raw_bytes),
                message_id,
            )
            return filename, mime_type, raw_bytes

    def create_draft(
        self,
        *,
        to_addresses: list[str],
        subject: str,
        body: str,
        cc_addresses: list[str] | None = None,
        bcc_addresses: list[str] | None = None,
    ) -> dict:
        """Create a draft email in Mail.app.

        Returns:
            {"success": bool, "message_id": str | None}
            message_id is the RFC 2822 Message-ID for constructing a message:// link,
            or None if Mail.app did not expose one on the saved draft.
        """
        safe_subject = _js_escape(subject)
        safe_body = _js_escape(body)

        def recip_js(kind: str, addrs: list[str]) -> str:
            cls_map = {"to": "ToRecipient", "cc": "CcRecipient", "bcc": "BccRecipient"}
            field_map = {"to": "toRecipients", "cc": "ccRecipients", "bcc": "bccRecipients"}
            cls = cls_map[kind]
            field = field_map[kind]
            lines = []
            for addr in addrs:
                name, email = _parse_address(addr)
                lines.append(
                    f'draft.{field}.push('
                    f'mail.{cls}({{address: "{_js_escape(email)}", name: "{_js_escape(name)}"}}'
                    f'));'
                )
            return "\n        ".join(lines)

        to_js = recip_js("to", to_addresses)
        cc_js = recip_js("cc", cc_addresses or [])
        bcc_js = recip_js("bcc", bcc_addresses or [])

        script = f"""
        (function() {{
            var mail = Application("Mail");

            // Record time before saving so the fallback search can exclude
            // pre-existing drafts with the same subject.
            var createdAfter = new Date();

            // Use JXA constructor style — mail.make() is not supported for
            // outgoing messages in Mail.app's JXA dictionary.
            var draft = mail.OutgoingMessage({{
                subject: "{safe_subject}",
                content: "{safe_body}",
                visible: false
            }});
            mail.outgoingMessages.push(draft);

            {to_js}
            {cc_js}
            {bcc_js}

            draft.save();

            // Draft Message-IDs are typically not assigned until send; this is
            // expected to return null in most cases, so the fallback below is
            // the real code path.
            var msgId = null;
            try {{ msgId = draft.messageId(); }} catch(e) {{}}

            if (!msgId) {{
                var accounts = mail.accounts();
                outer: for (var i = 0; i < accounts.length; i++) {{
                    var mboxes = accounts[i].mailboxes();
                    for (var j = 0; j < mboxes.length; j++) {{
                        if (mboxes[j].name().toLowerCase().indexOf("draft") === -1) continue;
                        try {{
                            var candidates = mboxes[j].messages.whose({{subject: "{safe_subject}"}});
                            var latest = -1, latestDate = null;
                            for (var k = 0; k < candidates.length; k++) {{
                                // Per-message try/catch: a bad message must not abort the loop.
                                var ds = null;
                                try {{ ds = candidates[k].dateSent(); }} catch(e) {{}}
                                if (!ds) {{ try {{ ds = candidates[k].dateReceived(); }} catch(e) {{}} }}
                                if (!ds || ds < createdAfter) continue;
                                if (!latestDate || ds > latestDate) {{ latestDate = ds; latest = k; }}
                            }}
                            if (latest >= 0) {{ try {{ msgId = candidates[latest].messageId(); }} catch(e2) {{}} }}
                        }} catch(e) {{}}
                        if (msgId) break outer;
                    }}
                }}
            }}

            return JSON.stringify({{"success": true, "message_id": msgId}});
        }})();
        """
        result = self._run_jxa(script, timeout=30)
        if result is None:
            return {"success": False, "message_id": None}
        return result

    def create_reply_draft(
        self,
        message_id: int,
        body: str,
        *,
        reply_all: bool = False,
        cc_addresses: list[str] | None = None,
        bcc_addresses: list[str] | None = None,
        include_quoted: bool = True,
    ) -> dict:
        """Create a reply draft to an existing message via Mail.app's native reply command.

        Mail.app's `reply` command sets the In-Reply-To and References headers
        and pre-populates the recipient list and "Re: ..." subject — none of
        which we can do reliably by constructing an OutgoingMessage from
        scratch. The user's body is prepended to Mail's auto-generated quoted
        block when include_quoted=True; otherwise it replaces the content.

        Returns:
            {"success": bool, "message_id": str | None, "subject": str | None,
             "to_addresses": list[str], "cc_addresses": list[str]}
            message_id is the RFC 2822 Message-ID for the saved draft (typically
            None until the draft is sent — same as create_draft).
        """
        location = self._find_message(message_id)
        if location is None:
            raise ValueError(f"Message {message_id} not found.")

        acct_name, mbox_name, _ = location
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)
        safe_body = _js_escape(body)
        reply_all_js = "true" if reply_all else "false"
        include_quoted_js = "true" if include_quoted else "false"

        def recip_js(kind: str, addrs: list[str]) -> str:
            cls_map = {"cc": "CcRecipient", "bcc": "BccRecipient"}
            field_map = {"cc": "ccRecipients", "bcc": "bccRecipients"}
            cls = cls_map[kind]
            field = field_map[kind]
            lines = []
            for addr in addrs:
                name, email = _parse_address(addr)
                lines.append(
                    f'draft.{field}.push('
                    f'mail.{cls}({{address: "{_js_escape(email)}", name: "{_js_escape(name)}"}}'
                    f'));'
                )
            return "\n        ".join(lines)

        extra_cc_js = recip_js("cc", cc_addresses or [])
        extra_bcc_js = recip_js("bcc", bcc_addresses or [])

        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
            if (accounts.length === 0) return JSON.stringify({{"success": false, "error": "account not found"}});
            var acct = accounts[0];
            var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
            if (mboxes.length === 0) return JSON.stringify({{"success": false, "error": "mailbox not found"}});
            var mb = mboxes[0];
            var msgs = mb.messages.whose({{id: {message_id}}});
            if (msgs.length === 0) return JSON.stringify({{"success": false, "error": "message not found"}});
            var src = msgs[0];

            var createdAfter = new Date();

            // Mail's native reply: sets In-Reply-To/References, To/Cc, and "Re:" subject.
            mail.reply(src, {{openingWindow: false, replyToAll: {reply_all_js}}});

            // Mail.app populates the auto-quoted body asynchronously after reply
            // returns. ~1 s is sufficient on modern hardware.
            delay(1.0);

            // Find the freshly-created outgoing message: scan from the end and
            // pick the first "Re:..." we haven't seen before by matching subject
            // against the source message.
            var srcSubject = "";
            try {{ srcSubject = src.subject() || ""; }} catch(e) {{}}
            var expectedSubject = "Re: " + srcSubject.replace(/^Re:\\s*/i, "");
            var draft = null;
            var outs = mail.outgoingMessages();
            for (var k = outs.length - 1; k >= 0; k--) {{
                try {{
                    var s = outs[k].subject() || "";
                    if (s === expectedSubject || s === "Re: " + srcSubject) {{
                        draft = outs[k];
                        break;
                    }}
                }} catch(e) {{}}
            }}
            if (!draft) return JSON.stringify({{"success": false, "error": "reply outgoing message not found"}});

            // Capture the auto-populated recipients and the auto-quoted body.
            var subj = "";
            try {{ subj = draft.subject() || ""; }} catch(e) {{}}
            var toAddrs = [];
            try {{
                var to = draft.toRecipients();
                for (var m = 0; m < to.length; m++) {{
                    var addr = to[m].address();
                    var nm = to[m].name();
                    toAddrs.push(nm ? (nm + " <" + addr + ">") : addr);
                }}
            }} catch(e) {{}}
            var ccAddrs = [];
            try {{
                var cc = draft.ccRecipients();
                for (var m = 0; m < cc.length; m++) {{
                    var addr = cc[m].address();
                    var nm = cc[m].name();
                    ccAddrs.push(nm ? (nm + " <" + addr + ">") : addr);
                }}
            }} catch(e) {{}}

            var quoted = "";
            try {{ quoted = draft.content() || ""; }} catch(e) {{}}

            // Prepend user body. include_quoted=false replaces the auto-quoted block.
            var newBody;
            if ({include_quoted_js}) {{
                newBody = "{safe_body}\\n\\n" + quoted;
            }} else {{
                newBody = "{safe_body}";
            }}
            try {{ draft.content = newBody; }} catch(e) {{}}

            // Push extra cc/bcc on top of what reply pre-populated.
            {extra_cc_js}
            {extra_bcc_js}

            draft.save();

            // Message-ID is typically null on save (only assigned at send time),
            // so the fallback below is the real code path.
            var msgId = null;
            try {{ msgId = draft.messageId(); }} catch(e) {{}}

            if (!msgId) {{
                // Allow Mail.app a moment to commit the draft to the Drafts mailbox.
                delay(2.0);
                var safeSubj = subj;
                var accts = mail.accounts();
                outer: for (var i = 0; i < accts.length; i++) {{
                    var dmboxes = accts[i].mailboxes();
                    for (var j = 0; j < dmboxes.length; j++) {{
                        if (dmboxes[j].name().toLowerCase().indexOf("draft") === -1) continue;
                        try {{
                            var candidates = dmboxes[j].messages.whose({{subject: safeSubj}});
                            var latest = -1, latestDate = null;
                            for (var c = 0; c < candidates.length; c++) {{
                                var ds = null;
                                try {{ ds = candidates[c].dateSent(); }} catch(e) {{}}
                                if (!ds) {{ try {{ ds = candidates[c].dateReceived(); }} catch(e) {{}} }}
                                if (!ds || ds < createdAfter) continue;
                                if (!latestDate || ds > latestDate) {{ latestDate = ds; latest = c; }}
                            }}
                            if (latest >= 0) {{ try {{ msgId = candidates[latest].messageId(); }} catch(e2) {{}} }}
                        }} catch(e) {{}}
                        if (msgId) break outer;
                    }}
                }}
            }}

            return JSON.stringify({{
                "success": true,
                "message_id": msgId,
                "subject": subj,
                "to_addresses": toAddrs,
                "cc_addresses": ccAddrs
            }});
        }})();
        """
        result = self._run_jxa(script, timeout=90)
        if result is None:
            return {
                "success": False,
                "message_id": None,
                "subject": None,
                "to_addresses": [],
                "cc_addresses": [],
            }
        return result

    def get_flag(self, message_id: int) -> dict:
        """Return flag status and color index for a message.

        Returns:
            {"is_flagged": bool, "color_index": int, "flag_color": str | None}
            flag_color is a color name like "red" or None when unflagged.
        """
        location = self._find_message(message_id)
        if location is None:
            raise ValueError(f"Message {message_id} not found.")

        acct_name, mbox_name, _ = location
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)

        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
            if (accounts.length === 0) return JSON.stringify(null);
            var acct = accounts[0];
            var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
            if (mboxes.length === 0) return JSON.stringify(null);
            var mb = mboxes[0];

            var msgs = mb.messages.whose({{id: {message_id}}});
            if (msgs.length === 0) return JSON.stringify(null);
            var msg = msgs[0];

            var isFlagged = msg.flaggedStatus() ? true : false;
            // flagIndex() returns -1 when unflagged, 1-7 for colors.
            // Fall back to isFlagged ? 1 : -1 if the property is unavailable.
            var colorIdx = isFlagged ? 1 : -1;
            try {{
                colorIdx = msg.flagIndex();
            }} catch(e) {{}}
            return JSON.stringify({{is_flagged: isFlagged, color_index: colorIdx}});
        }})();
        """
        result = self._run_jxa(script, timeout=30)
        if result is None:
            raise ValueError(f"Message {message_id} not found.")

        color_index: int = result.get("color_index", -1)
        flag_color: Optional[str] = (
            _FLAG_COLOR_ORDER[color_index - 1] if 1 <= color_index <= 7 else None
        )
        return {
            "is_flagged": bool(result.get("is_flagged", False)),
            "color_index": color_index,
            "flag_color": flag_color,
        }

    def set_flag(self, message_id: int, flag: Optional[str] = None) -> dict:
        """Set or remove the flag on a message.

        Args:
            flag: A color string ("red", "orange", etc.) or None to remove the flag.

        Returns:
            {"success": bool, "is_flagged": bool}
        """
        location = self._find_message(message_id)
        if location is None:
            raise ValueError(f"Message {message_id} not found.")

        acct_name, mbox_name, _ = location
        safe_acct = _js_escape(acct_name)
        safe_mbox = _js_escape(mbox_name)
        # color_index: 1-7 = set color, -1/None = unflag via flaggedStatus = false
        color_index = _FLAG_COLOR_MAP[flag] if flag is not None else -1

        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts.whose({{name: "{safe_acct}"}});
            if (accounts.length === 0) return JSON.stringify({{success: false, error: "account not found"}});
            var acct = accounts[0];
            var mboxes = acct.mailboxes.whose({{name: "{safe_mbox}"}});
            if (mboxes.length === 0) return JSON.stringify({{success: false, error: "mailbox not found"}});
            var mb = mboxes[0];

            var msgs = mb.messages.whose({{id: {message_id}}});
            if (msgs.length === 0) return JSON.stringify({{success: false, error: "message not found"}});
            var msg = msgs[0];

            var colorIdx = {color_index};
            if (colorIdx < 1) {{
                // Unflag: flaggedStatus = false sets flagIndex to -1
                msg.flaggedStatus = false;
            }} else {{
                // Set specific color via flagIndex (confirmed in JXA probe)
                try {{
                    msg.flagIndex = colorIdx;
                }} catch(e) {{
                    // Fallback: boolean flag (always red)
                    msg.flaggedStatus = true;
                }}
            }}

            var isFlagged = msg.flaggedStatus() ? true : false;
            // Read back the actual flag index so the caller knows the true resulting color.
            var actualIdx = isFlagged ? 1 : -1;
            try {{ actualIdx = msg.flagIndex(); }} catch(e) {{}}
            return JSON.stringify({{success: true, is_flagged: isFlagged, color_index: actualIdx}});
        }})();
        """
        result = self._run_jxa(script, timeout=30)
        if result is None:
            return {"success": False, "is_flagged": False}
        return result

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _find_message(self, message_id: int) -> Optional[tuple[str, str, Optional[int]]]:
        """Find which account/mailbox contains this message.

        Uses cache first, then searches non-empty mailboxes.

        Returns:
            (account_name, mailbox_name, msg_index_or_None) or None.
        """
        # Check cache
        cached = self._message_cache.get(message_id)
        if cached is not None:
            return cached

        # Search only non-empty mailboxes (cached from init)
        logger.debug("Cache miss for message %d, searching non-empty mailboxes...", message_id)
        nonempty_keys = [
            f'"{_js_escape(a)}|{_js_escape(m)}"'
            for a, m in self._nonempty_mailboxes
        ]
        nonempty_set_js = "var _ne = {" + ",".join(
            f"{k}: 1" for k in nonempty_keys
        ) + "};"
        script = f"""
        (function() {{
            var mail = Application("Mail");
            var accounts = mail.accounts();
            {nonempty_set_js}
            for (var i = 0; i < accounts.length; i++) {{
                var acctName = accounts[i].name();
                var mboxes = accounts[i].mailboxes();
                for (var j = 0; j < mboxes.length; j++) {{
                    var mb = mboxes[j];
                    var mbName = mb.name();
                    if (!_ne[acctName + "|" + mbName]) continue;
                    try {{
                        var msgs = mb.messages.whose({{id: {message_id}}});
                        if (msgs.length > 0) {{
                            return JSON.stringify({{
                                "account_name": acctName,
                                "mailbox_name": mbName
                            }});
                        }}
                    }} catch(e) {{
                        // Skip inaccessible mailboxes
                    }}
                }}
            }}
            return JSON.stringify(null);
        }})();
        """
        try:
            result = self._run_jxa(script, timeout=60)
        except RuntimeError:
            logger.warning("Failed to locate message %d.", message_id)
            return None

        if result is None:
            return None

        acct = result.get("account_name", "")
        mbox = result.get("mailbox_name", "")
        location = (acct, mbox, None)  # no index from whose-based lookup

        # Cache for future lookups
        self._message_cache[message_id] = location
        logger.debug("Found message %d in %s / %s", message_id, acct, mbox)
        return location
