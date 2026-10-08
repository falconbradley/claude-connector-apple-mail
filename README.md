# Apple Mail MCP

A Claude Desktop extension that gives **Claude fast access to Apple Mail** on macOS. Reads come straight from Mail.app's local message store — searching a 250k-message mailbox takes **milliseconds**, not tens of seconds — while Mail.app remains the sync and auth engine (iCloud, Gmail, anything Mail supports), so no IMAP credentials or passwords are ever handled. Writes (drafts, flags) go through Mail.app's native scripting interface.

Packaged as an [MCPB desktop extension](https://support.claude.com/en/articles/12922929-building-desktop-extensions-with-mcpb) with the Apple Mail icon and one-click install.

---

## What it does

| Tool | Description |
|------|-------------|
| `get_stats` | Total messages, unread count, mailbox and account counts |
| `list_mailboxes` | Every account/folder with message counts |
| `search_emails` | Rich search: free text, sender, recipient (To/CC), subject, date range, read/flagged status, attachments. Every result includes clickable open-in-Mail links |
| `get_email` | Full email with decoded plain-text body, recipients, flag color, and metadata |
| `preview_email` | Show an email as an inline preview card in the chat — sender, recipients, date, flags, attachments, rendered body, and an Open-in-Mail button. Same JSON as `get_email` for the model |
| `get_email_link` | Links that open an email in Mail.app — `scope="message"` for the single email, `scope="thread"` for its conversation |
| `open_email_in_mail` | Open an email directly in Mail.app (for chat UIs that block `message://` links) |
| `get_selected_emails` | The message(s) currently selected in Mail.app's viewer — id, subject, sender, mailbox, and open-in-Mail links |
| `get_email_html` | HTML body of a message |
| `get_thread` | All messages in a conversation thread (summaries only, no bodies) |
| `preview_thread` | Show a whole conversation as one inline card — every message in the thread, collapsed to a line each with the newest open, each expanding to its full body |
| `list_email_attachments` | Enumerate attachments for any email |
| `get_email_attachment` | Retrieve attachment content (base64) |
| `create_email_draft` | Create a draft email saved to Mail.app's Drafts mailbox, returns a `message://` link to open it |
| `create_email_reply_draft` | Reply to an existing message — preserves `In-Reply-To`/`References` headers so the reply threads correctly in the recipient's client |
| `get_email_flag` | Get the flag status and color (e.g. `"orange"`) for an email |
| `set_email_flag` | Set or remove a color flag on an email (red/orange/yellow/green/blue/purple/gray, or null to remove) |

### Timestamps are UTC

Every timestamp the server returns (`date_sent`, `date_received`) is UTC, and the
`since` / `before` filters read a value with no offset as UTC to match — so a
`date_sent` copied out of a result round-trips exactly. To filter on a local wall
clock (e.g. "today" as it reads in Mail.app), pass the offset explicitly:

```
since = "2026-08-24T00:00:00-07:00"   # local midnight
since = "2026-08-24T00:00:00"         # UTC midnight — 17:00 the previous day in PDT
```

## How it works

Mail.app remains the **sync and auth engine** — it holds your Apple ID / iCloud credentials natively and continuously mirrors every account to disk. This server has two engines on top of that:

### Fast read path (default, needs Full Disk Access)

Reads are served directly from Mail.app's local message store:

- **`~/Library/Mail/V*/MailData/Envelope Index`** — Mail's SQLite index of every message (subjects, senders, recipients, dates, read/flag state). Searches complete in **milliseconds** instead of tens of seconds.
- **`.emlx` files** — raw RFC 2822 messages on disk, parsed for bodies, headers, HTML, and attachments.

No credentials are ever handled: the server is a read-only consumer of data Mail.app has already synced. The store is opened read-only (`PRAGMA query_only`) and never mutated. The only extra requirement is **Full Disk Access**, granted once in System Settings — see [Permissions](#permissions) for the `uv` gotcha.

The schema of the Envelope Index varies across macOS releases, so the server introspects it at runtime and adapts (falling back to the documented `flags` bitfield when dedicated columns are absent). Account UUIDs in mailbox URLs are resolved to display names ("iCloud", "Work Gmail") via the system accounts store. Run `uv run python -m apple_mail_mcp.selftest` from a terminal with Full Disk Access to verify the fast path on your machine.

### Inline email previews (MCP Apps)

`preview_email` renders a message as a card directly in the chat transcript, the
way the Gmail and Superhuman connectors do. It uses the
[MCP Apps extension](https://github.com/modelcontextprotocol/ext-apps): the
server ships a `ui://apple-mail/email-preview` HTML resource, the tool points at
it via `_meta.ui.resourceUri`, and hosts that support the extension (Claude
Desktop, claude.ai) load it in a sandboxed iframe beside the tool call and feed
it the tool's `structuredContent`.

- The card shows sender (with avatar), recipients, date in the host's locale and
  time zone, unread/flag state, attachment chips, and the body. HTML mail is
  sanitized client-side (allowlisted tags and attributes, `javascript:` and
  `data:` links stripped, inline styles scrubbed of `url()`/`expression()`); plain
  text is linkified with quoted lines styled. A toggle switches between the two.
- **Open in Mail** goes through the host's `ui/open-link` to the connector's
  localhost redirector, which pops the message open in Mail.app.
- Remote images are never loaded. The card declares no CSP allowances, so the
  host runs it under the strictest default: no network, no nested frames.
  Tracking pixels are dropped; other remote images become labelled placeholders.
- The model only sees `get_email`-shaped JSON in the tool's text content. The
  HTML body rides in `structuredContent`, which goes to the iframe, not the
  model's context.
- Hosts without MCP Apps ignore the UI hint and get a normal text result.

The card lives in `src/apple_mail_mcp/ui/email_preview.html` as one
self-contained document — no SDK, no external assets.

### Clickable open-in-Mail links

Every search/thread/email result carries two links:

- **`mail_link`** — the raw `message://<Message-ID>` URL. Works in Terminal (`open '<url>'`), Notes, Reminders, task managers, and Safari — but most chat UIs (including Claude Desktop and Claude Code) block custom URL schemes in rendered links.
- **`open_link`** — `http://127.0.0.1:<port>/open/<id>?t=<token>`. Chat UIs open http links fine: the click routes through your browser to a tiny localhost-only server inside the extension, which tells macOS to front Mail.app on that message. Requests require a per-install random token (persisted, so links in old conversations keep working); the endpoint's only capability is focusing Mail — it never serves message content. **The tab closes itself**: Mail.app coming to the front is the real confirmation, so the success page closes the tab the browser opened for the click instead of leaving it to pile up. Browsers only permit this for a tab with no history to go back to, which is exactly this case; where it is refused the page stays and explains itself. Error pages never auto-close.

There is also an `open_email_in_mail` tool so Claude can jump to a message directly without any clicking.

### Message links vs. thread links

`get_email_link` takes a `scope`, because "link me to that email" and "link
me to that thread" are different requests:

- **`scope="message"`** (default) — that one email.
- **`scope="thread"`** — its conversation. Returns a `thread_link`
  (`http://127.0.0.1:<port>/thread/<id>?t=<token>`) alongside the thread's
  size and newest message. The conversation is resolved when the link is
  *clicked*, not when it is generated, so a thread link in an old chat
  transcript follows the thread as new replies arrive.

**Mail.app cannot open a conversation.** Its only registered URL schemes are
`mailto:`, `message:` and `mail-pref-pane:`, and `message://` always opens a
single message in its own window — even when that message is already visible
in the front viewer with Organize by Conversation on. (Mail's scripting
interface can't get there either: `message viewer`'s `selected messages` is
nominally settable but selects the wrong message on current macOS, and
`visible messages` errors outright.) So a `thread_link` fronts the
conversation's **newest** message and says so on the result page; Mail's own
conversation grouping shows the rest.

To actually read a back-and-forth, use **`preview_thread`** — it renders
every message of the conversation inline in the chat, so nothing depends on
what Mail.app is willing to open.

### JXA fallback + writes

When Full Disk Access is missing, reads transparently fall back to the original **JXA (JavaScript for Automation)** bridge (the two-round bulk-fetch search, ~37s across 60k+ messages). Search results include an `engine` field (`"sqlite"` or `"applescript"`) so you can tell which path served them. Set `APPLE_MAIL_MCP_DISABLE_FAST=1` to force the JXA path.

The fallback scales poorly, and past a couple hundred thousand messages some searches simply cannot finish in time. MCP clients abandon a tool call after 60s, so the fallback holds itself to a **55s budget** shared across both rounds:

- **Round 1 overruns the budget** → the search fails with an error naming the cause and suggesting how to narrow it. Previously the internal limit was 300s, so the client timed out at 60s while `osascript` kept running for another four minutes, holding Mail.app busy and slowing every later call.
- **Only Round 2 (display properties) is left short** → results are returned with just the Round 1 fields rather than failing. When you filtered on subject or sender those are already populated, so the degraded result usually looks identical to the full one.
- **`has_attachments` is rejected outright.** Mail.app exposes no bulk attachment property and reading it per message costs minutes, so this filter needs the fast engine. On the fallback it raises instead of silently returning unfiltered results, and search results report `has_attachments: null` ("not determined") rather than a misleading `false` — use `get_email` for an authoritative answer on one message.

Writes — drafts, reply drafts, flag changes — always go through Mail.app scripting (Automation permission), so Mail.app owns every mutation and syncs it back to the server (e.g. iCloud) itself.

---

## Requirements

- macOS 13 Ventura or later
- Apple Mail running with at least one configured account
- Python 3.11+
- Claude Desktop (with extension support)

---

## Installation

### Option 1: Desktop Extension (recommended)

Download the latest `.mcpb` from [Releases](../../releases), then **double-click** to install.

Or build from source:

```bash
git clone https://github.com/falconbradley/claude-connector-apple-mail.git
cd claude-connector-apple-mail
./build.sh
```

Then double-click `dist/apple-mail-<version>.mcpb` (or drag it into Claude Desktop).

The extension appears in **Settings > Extensions** with the Apple Mail icon.

### Option 2: Manual MCP config

Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "apple-mail": {
      "command": "uv",
      "args": ["run", "--project", "/path/to/apple-mail-mcp", "apple-mail-mcp"]
    }
  }
}
```

### Permissions

Two macOS permissions matter:

1. **Full Disk Access** (for the fast read path): System Settings > Privacy & Security > Full Disk Access > enable **uv**. Claude Desktop launches extension servers through a helper that makes the spawned process itself responsible for permissions, so macOS attributes FDA to the `uv` launcher binary — enabling Claude Desktop alone is *not* sufficient. If `uv` isn't in the list, add it with **+** (press ⌘⇧G to paste a path): the installed extension uses `~/Library/Application Support/Claude/uv-runtime/<version>/uv` — the fallback's error message prints the exact path — while a manual `claude_desktop_config.json` setup uses whichever `uv` is on your `PATH` (e.g. `~/.local/bin/uv`). Then quit Claude (⌘Q) and reopen it; macOS reads this permission only at launch. Without FDA, reads still work via the slow AppleScript fallback. (Enable your terminal app too if you want to run the selftest.)
2. **Automation** (for writes and the fallback): Mail.app must be running; macOS prompts automatically on first use — click **OK**. If the prompt doesn't appear, check System Settings > Privacy & Security > Automation.

### Using several Apple connectors

This connector is one of a family of Claude Desktop extensions for Apple apps — **Mail**, [Messages](https://github.com/falconbradley/claude-connector-apple-messages), [Contacts](https://github.com/falconbradley/claude-connector-apple-contacts), [Calendar](https://github.com/falconbradley/claude-connector-apple-calendar), [Reminders](https://github.com/falconbradley/claude-connector-apple-reminders), [Notes](https://github.com/falconbradley/claude-connector-apple-notes) — and they share the same setup quirks:

- **Install them one at a time.** Opening several `.mcpb` files at once can leave Claude Desktop showing only the last install dialog, so the others silently never install. Approve each dialog before opening the next, then check **Settings → Extensions**.
- **Permissions belong to `uv`, not Claude.** Claude Desktop launches every extension through the same bundled `uv` and macOS attributes their privacy grants to it. Full Disk Access granted once to `~/Library/Application Support/Claude/uv-runtime/<version>/uv` covers Mail, Notes, Messages, and Reminders together; the Contacts, Calendars, and Reminders panes list the connectors as **uv**. Permission errors print the exact path in use, ready to paste.
- **Re-grant after Claude Desktop updates `uv`.** The `<version>` folder changes and macOS treats the new binary as a new app. Symptoms: Mail and Notes searches report `"engine": "applescript"` and get slow, Messages reads and Reminders tags fail with a Full Disk Access error.
- **Restart after granting.** Quit Claude (⌘Q) and reopen it — macOS reads Full Disk Access only at launch.
- **Verify.** Ask Claude for each connector's stats (`get_stats`). For Mail and Notes, a search result's `engine` should be `"sqlite"`.

---

## Usage examples

Once installed, just ask Claude naturally:

- *"Show me my unread emails from this week"*
- *"Search for emails from alice@example.com about the Q4 budget"*
- *"Show me emails sent to bill@example.com in the last month"*
- *"What attachments are in the last email from my accountant?"*
- *"Summarise the email thread about the contract renewal"*
- *"Show me the whole back-and-forth with Alice"* (renders the conversation as a card)
- *"Link me to that thread"* (vs. *"link me to that email"*)
- *"Find flagged emails with PDF attachments"*
- *"Draft a reply to John's email about the project update"*
- *"Reply-all to that thread saying I'll review by EOD"*
- *"Create a draft email to the team announcing Friday's meeting"*
- *"Flag this email as orange"*
- *"What color is the flag on that email from Sarah?"*

---

## Building from source

```bash
# Install mcpb CLI (one time)
npm install -g @anthropic-ai/mcpb

# Build the extension
./build.sh

# Or manually:
mcpb validate manifest.json
mcpb pack . dist/apple-mail-<version>.mcpb
```

### Project layout

```
apple-mail-mcp/
├── manifest.json              # MCPB desktop extension manifest
├── icon.png                   # Apple Mail icon (512x512)
├── icons/                     # Multi-size icons
│   ├── icon-128.png
│   ├── icon-256.png
│   └── icon-512.png
├── pyproject.toml             # Python package + dependencies
├── build.sh                   # Test, check, and pack build script
└── src/
    └── apple_mail_mcp/
        ├── __init__.py
        ├── server.py          # MCP tools (FastMCP)
        ├── hybrid.py          # Fast SQLite reads with JXA fallback
        ├── envelope.py        # Envelope Index (SQLite) read engine
        ├── applescript.py     # JXA bridge to Mail.app
        ├── emlx.py            # MIME body extraction utilities
        ├── weblink.py         # Localhost redirector for open-in-Mail links
        ├── preview.py         # Inline preview card payload + ui:// resource
        ├── ui/
        │   └── email_preview.html  # The MCP Apps card (self-contained)
        └── models.py          # Pydantic data models
```

---

## Performance notes

Search performance depends on mailbox size and which filters are active. Bulk property fetches are conditional — only the properties needed for active filters are fetched.

| Scenario | Approx. time |
|----------|-------------|
| Init (one-time mailbox prescan) | ~12s |
| Search with date filter only | ~14s |
| Search with text (subject/sender) | ~24s |
| Search with recipient (To/CC) filter | ~68s |
| Full search (all filters) | ~91s |

Times measured against ~61K messages across 7 mailboxes. Searches without optional filters add zero overhead for those properties.

---

## Roadmap

**Phase 1 — Read**
- [x] List mailboxes and accounts
- [x] Search emails (subject, sender, recipient, date, flags, attachments)
- [x] Read full message body (plain text + HTML)
- [x] Thread view
- [x] List and retrieve attachments
- [x] `message://` links to open emails in Mail.app
- [x] Inline email preview cards in the chat (MCP Apps)

**Phase 2 — Write (in progress)**
- [x] Create draft emails (saved to Drafts with a `message://` link to open)
- [x] Reply to a thread (preserves `In-Reply-To`/`References` headers)
- [x] Set, change, or remove color flags on emails
- [ ] Mark as read / unread
- [ ] Move to folder
- [ ] Delete (move to Trash)

---

## Security & privacy

- Read operations never modify your mail: the Envelope Index is opened with `PRAGMA query_only` and `.emlx` files are only ever read. Write operations go through Mail.app scripting and are limited to: creating drafts (saved locally, never sent automatically) and setting/removing flags on messages.
- No data leaves your machine — this is a local MCP server. Mail.app keeps sole custody of account credentials (iCloud sign-in, OAuth, etc.).
- The open-in-Mail link redirector binds to 127.0.0.1 only, requires a per-install random token on every request, and can only focus Mail.app on a message — it never serves message content.
- The inline preview card runs in the host's sandboxed iframe with no network access; email HTML is sanitized before rendering and remote images are never fetched, so a message can't phone home when previewed.
- Full Disk Access (read-only usage) powers the fast search path; without it the extension degrades to Automation-only scripting.
- macOS-only (`"platforms": ["darwin"]` in manifest).
- Attachment data is returned as base64 only when explicitly requested.

---

## Troubleshooting

**"Mail.app is not running"**
Open Mail.app before using the extension. It must be running for JXA scripting to work.

**"Automation permission denied"**
Go to **System Settings > Privacy & Security > Automation** and ensure Claude Desktop (or Terminal) is allowed to control Mail.app. Then restart Claude Desktop.

**Search is slow or times out**
Large mailboxes (50k+ messages) take longer. Use date filters (`since`) to narrow the search window. The first search after startup includes a one-time ~12s mailbox prescan.

**Extension doesn't appear after install**
Make sure you're running a recent version of Claude Desktop that supports MCPB extensions. Restart Claude Desktop after installing.

---

## Releasing

Every connector in the family releases the same way:

1. Bump the version in `pyproject.toml`, `manifest.json`, and `src/apple_mail_mcp/__init__.py`, then run `uv lock` so `uv.lock` matches. CI fails if the three disagree.
2. Add a section for the version to [CHANGELOG.md](CHANGELOG.md).
3. Commit, tag `vX.Y.Z`, and push the tag: `git push origin main vX.Y.Z`.

The [release workflow](.github/workflows/release.yml) then checks the tag matches all three version files, runs `./build.sh` (tests, manifest and tool checks, pack), and publishes `apple-mail-X.Y.Z.mcpb` to a GitHub release whose notes are that version's CHANGELOG section.

## License

[MIT](LICENSE)
