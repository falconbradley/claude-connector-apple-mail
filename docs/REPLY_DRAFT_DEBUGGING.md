# Reply draft debugging log

Persistent log of attempts to fix reply-draft quoting issues in
`create_reply_draft`. The point of this file is durability: anyone
(future Brad, future Claude) picking this up should be able to read this
top-to-bottom and know what's been tried, what worked, what didn't, and
why — without spelunking through git history.

**How to update this file:**
- Append new attempts under "Attempts" in chronological order. Don't
  rewrite history — strike through with `~~text~~` if something
  superseded a prior finding, but keep the original entry.
- Update "Current state" at the top whenever a fix lands or a new
  symptom is confirmed.
- Record _both_ failures and successes. A confirmed dead end is as
  valuable as a confirmed fix — it saves the next person from re-running
  the same experiment.

Related code:
- `src/apple_mail_mcp/applescript.py::MailBridge.create_reply_draft`
- `src/apple_mail_mcp/applescript.py::_build_quoted_reply_body` (used by
  `create_draft`, _not_ the reply path)
- `tools/dev_harness.py` — exercises the real code path
- `tools/probe_strategies.py` — A/B/C strategy comparison
- `tools/draft_inspect.py` — read-only draft dump helpers

---

## Current state

**As of 2026-05-24 (v0.4.17)** — reply drafts work correctly. Mail's
native gray vertical-bar blockquote is preserved, the user's reply
text is placed above the quote, and the compose window is left open
and foregrounded for review.

The strategy is `reply with opening window` + System Events Cmd+V
paste, landed in #9 (commit `93cfc46`). The earlier "missing
blockquote bar" report was against an older _installed_ version of
the extension, not the current code.

**Caveats users should still know about:**
- Requires Accessibility permission for Claude Desktop (responsible
  app for the System Events keystroke). If denied, the saved draft
  will have Mail's native quote but no user body. STATS line in the
  AppleScript output emits `sysevents_paste=yes/no` for diagnostics.
- Replies against a Drafts-mailbox source hang Mail.app; the tool
  refuses with a clear error.
- Saved-draft Message-ID is empty until send, so `draft_link` is
  sometimes `null`.

---

## Strategy reference

Three reply-draft strategies have been explored in the codebase. The
current production path is **B'** (a variant of B).

| ID | Strategy | Outcome |
|----|----------|---------|
| A | `reply without opening window` + `set content of foundMsg to "..."` + `save` | Lost native blockquote — `set content` only accepts plain text and discards Mail's styled rich-text body. (#5, then superseded.) |
| B | `reply with opening window` + `set visible false` + `set content` + `save` | Same blockquote loss as A; the hidden window doesn't change anything. |
| B' | `reply with opening window` + System Events Cmd+V paste + `save` (window left open) | **Current.** Preserves blockquote in principle. Requires Accessibility permission for the host app (Claude Desktop). #9 / commit `93cfc46`. |
| C | `make new outgoing message with properties {subject, content}` + `save` | No threading headers (no `In-Reply-To` / `References`). Not viable for replies. |

---

## Constraints we've already discovered

These are load-bearing facts. Don't re-derive them from scratch.

1. **`set content of outgoing-message` accepts plain text only.** It
   silently discards any HTML structure Mail's `reply` populated in the
   compose window — including the native blockquote bar.
2. **`reply` against a Drafts-mailbox message hangs Mail.app
   indefinitely.** Reply source must come from Inbox or another
   received-mail mailbox. Current code rejects with a clear error.
3. **System Events keystrokes require Accessibility permission.**
   The _responsible app_ is whichever process invoked the MCP server
   (typically Claude Desktop). Without it, keystrokes are silently
   dropped by TCC — saved draft contains only Mail's native quote, no
   user body. The diagnostic STATS line emits `sysevents_paste=yes/no`.
4. **OutgoingMessage subject-match isn't unique enough to identify the
   newly-created draft.** Use a pre/post snapshot of
   `count of outgoing messages` and take the last item, or snapshot
   draft IDs before the reply call and diff afterwards.
5. **Draft Message-IDs are not assigned until send.** The `message id`
   property of a saved draft is typically empty. The
   `message://` link returned to the user may be `None`.

---

## Attempts

Format per attempt:

```
### YYYY-MM-DD — short title
**Hypothesis:** what we thought was happening
**Change:** what we did
**Result:** what happened (with concrete evidence — STATS line,
inspected draft body, screenshot path, etc.)
**Verdict:** fixed / regressed / no change / inconclusive
**Notes:** any context for the next attempt
```

---

### 2026-05-05 — Initial fix: clipboard + System Events paste (commit `93cfc46`, PR #9)

**Hypothesis:** `set content` discards Mail's native blockquote because
it accepts plain text only. Pasting the user body into the live compose
window via Cmd+V should leave the surrounding rich-text intact.

**Change:** Switched from `reply without opening window` + `set content`
to `reply with opening window`, put user body on clipboard, send
`keystroke "v" using command down` via System Events, then `save`.
Compose window left open and foregrounded for user review.

**Result:** Worked in dev (with Terminal/iTerm holding Accessibility
permission). Production path (Claude Desktop as the responsible app)
not yet broadly validated.

**Verdict:** Improvement, but the blockquote-missing reports suggest
this path doesn't fully solve the problem in production.

**Notes:** First place to look on a new report is the STATS line in
the AppleScript output — `sysevents_paste=no` means TCC denied the
keystroke and the bug is permissions, not styling.

---

### 2026-05-24 — Confirmed v0.4.17 baseline: behaves as designed

**Hypothesis:** Earlier "missing blockquote bar" reports might be
against a stale installed version, not the current code.

**Change:** Brad upgraded his installed extension to v0.4.17 (built
from commit `4d72d37` plus the new `get_selected_emails` tool) and
exercised the reply-draft flow end-to-end via Claude Desktop.

**Result:** Reply draft was created with a proper Mail blockquote on
the quoted original, the user reply text placed above the quote, and
the compose window left open and foregrounded for review. Behaves
exactly as the design intends.

**Verdict:** No bug present in current code. Closing this thread.

**Notes:** If future reports surface a missing blockquote bar, first
check (a) which version is actually installed (Finder → Get Info on
the `.mcpb` shows version in Comments), and (b) Accessibility
permission for Claude Desktop. Re-open this debugging log with a new
attempt entry if real evidence of regression appears.

---

<!-- New attempts go below this line. Copy the template above. -->
