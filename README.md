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
| `get_email` | Full email with decoded plain-text body, recipients, and metadata |
| `get_email_link` | Get a `message://` URL that opens the email directly in Mail.app |
| `get_email_html` | HTML body of a message |
| `get_thread` | All messages in a conversation thread |
| `list_email_attachments` | Enumerate attachments for any email |
| `get_email_attachment` | Retrieve attachment content (base64) |
| `create_email_draft` | Create a draft email saved to Mail.app's Drafts mailbox, returns a `message://` link to open it |

## How it works

All communication with Mail.app is done through **JXA (JavaScript for Automation)** via `osascript -l JavaScript`. This means:

- No direct database or filesystem access required
- No Full Disk Access needed — only Automation permission (macOS prompts automatically)
- Mail.app handles all IMAP/account authentication natively

The server uses a **two-round bulk-fetch search architecture** optimised for large mailboxes:

1. **Round 1** — For each non-empty mailbox, bulk-fetch message IDs + dates + conditional filter properties (subjects, senders, recipients, flags). Apply all filters as JavaScript post-processing. Sort by date, paginate.
2. **Round 2** — For only the mailboxes containing page results, bulk-fetch display properties (subject, sender, read/flagged status) to complete the result set.

This approach avoids Mail.app's extremely slow `whose` queries and per-message property access, achieving ~37s search times across 60k+ messages.

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

Mail.app must be running. On first use, macOS will prompt you to grant Automation permission — just click **OK**. No Full Disk Access is needed.

If the prompt doesn't appear, check **System Settings > Privacy & Security > Automation** and ensure your host process (Claude Desktop or Terminal) is allowed to control Mail.app.

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
- *"Create a draft email to the team announcing Friday's meeting"*

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
- [ ] Mark as read / unread
- [ ] Flag / unflag
- [ ] Move to folder
- [ ] Delete (move to Trash)

---

## Security & privacy

- Read operations never modify your mail. The only write operation is `create_email_draft`, which saves a draft locally — it does not send anything.
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
