# Apple Mail MCP

A Claude Desktop extension that gives **Claude access to Apple Mail** on macOS via Mail.app's native scripting interface. No IMAP credentials, no database access, no Full Disk Access needed — just Automation permission, which macOS prompts for automatically.

Packaged as an [MCPB desktop extension](https://support.claude.com/en/articles/12922929-building-desktop-extensions-with-mcpb) with the Apple Mail icon and one-click install.

---

## What it does

| Tool | Description |
|------|-------------|
| `get_stats` | Total messages, unread count, mailbox and account counts |
| `list_mailboxes` | Every account/folder with message counts |
| `search_emails` | Rich search: free text, sender, recipient (To/CC), subject, date range, read/flagged status, attachments |
| `get_email` | Full email with decoded plain-text body, recipients, flag color, and metadata |
| `get_email_link` | Get a `message://` URL that opens the email directly in Mail.app |
| `get_email_html` | HTML body of a message |
| `get_thread` | All messages in a conversation thread |
| `list_email_attachments` | Enumerate attachments for any email |
| `get_email_attachment` | Retrieve attachment content (base64) |
| `create_email_draft` | Create a draft email saved to Mail.app's Drafts mailbox, returns a `message://` link to open it |
| `create_email_reply_draft` | Reply to an existing message — preserves `In-Reply-To`/`References` headers so the reply threads correctly in the recipient's client |
| `get_email_flag` | Get the flag status and color (e.g. `"orange"`) for an email |
| `set_email_flag` | Set or remove a color flag on an email (red/orange/yellow/green/blue/purple/gray, or null to remove) |

## How it works

Mail.app remains the **sync and auth engine** — it holds your Apple ID / iCloud credentials natively and continuously mirrors every account to disk. This server has two engines on top of that:

### Fast read path (default, needs Full Disk Access)

Reads are served directly from Mail.app's local message store:

- **`~/Library/Mail/V*/MailData/Envelope Index`** — Mail's SQLite index of every message (subjects, senders, recipients, dates, read/flag state). Searches complete in **milliseconds** instead of tens of seconds.
- **`.emlx` files** — raw RFC 2822 messages on disk, parsed for bodies, headers, HTML, and attachments.

No credentials are ever handled: the server is a read-only consumer of data Mail.app has already synced. The store is opened read-only (`PRAGMA query_only`) and never mutated. The only extra requirement is **Full Disk Access** for the host process (Claude Desktop), granted once in System Settings.

The schema of the Envelope Index varies across macOS releases, so the server introspects it at runtime and adapts (falling back to the documented `flags` bitfield when dedicated columns are absent). Run `uv run python -m apple_mail_mcp.selftest` from a terminal with Full Disk Access to verify the fast path on your machine.

### JXA fallback + writes

When Full Disk Access is missing, reads transparently fall back to the original **JXA (JavaScript for Automation)** bridge (the two-round bulk-fetch search, ~37s across 60k+ messages). Search results include an `engine` field (`"sqlite"` or `"applescript"`) so you can tell which path served them. Set `APPLE_MAIL_MCP_DISABLE_FAST=1` to force the JXA path.

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

Then double-click `dist/apple-mail.mcpb` (or drag it into Claude Desktop).

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

1. **Full Disk Access** (for the fast read path): System Settings > Privacy & Security > Full Disk Access > enable **Claude Desktop** (and your terminal app if you want to run the selftest). Without it, reads still work via the slow AppleScript fallback.
2. **Automation** (for writes and the fallback): Mail.app must be running; macOS prompts automatically on first use — click **OK**. If the prompt doesn't appear, check System Settings > Privacy & Security > Automation.

---

## Usage examples

Once installed, just ask Claude naturally:

- *"Show me my unread emails from this week"*
- *"Search for emails from alice@example.com about the Q4 budget"*
- *"Show me emails sent to bill@example.com in the last month"*
- *"What attachments are in the last email from my accountant?"*
- *"Summarise the email thread about the contract renewal"*
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
mcpb pack . dist/apple-mail.mcpb
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
├── build.sh                   # Validate + pack build script
└── src/
    └── apple_mail_mcp/
        ├── __init__.py
        ├── server.py          # MCP tools (FastMCP)
        ├── applescript.py     # JXA bridge to Mail.app
        ├── emlx.py            # MIME body extraction utilities
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

**Phase 2 — Write (in progress)**
- [x] Create draft emails (saved to Drafts with a `message://` link to open)
- [x] Reply to a thread (preserves `In-Reply-To`/`References` headers)
- [x] Set, change, or remove color flags on emails
- [ ] Mark as read / unread
- [ ] Move to folder
- [ ] Delete (move to Trash)

---

## Security & privacy

- Read operations never modify your mail. Write operations are limited to: creating drafts (saved locally, never sent automatically) and setting/removing flags on messages.
- No data leaves your machine — this is a local MCP server.
- Only requires Automation permission, not Full Disk Access.
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

## License

[MIT](LICENSE)
