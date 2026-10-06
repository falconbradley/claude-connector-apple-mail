"""Name the binary macOS holds responsible for this server's privacy grants.

Claude Desktop starts extension servers through a helper that disclaims
responsibility for its children, so macOS attributes every privacy grant —
Full Disk Access, Contacts, Calendars, Reminders, Automation — to the `uv`
that `uv run` launched, not to Claude. Enabling Claude in System Settings
therefore does nothing for an extension.

That `uv` lives under a versioned folder
(`~/Library/Application Support/Claude/uv-runtime/uv-<version>-<arch>/uv`)
which changes whenever Claude Desktop updates its bundled copy, so the
error messages here print the exact path in use rather than a pattern for
the user to resolve.

This module is shared verbatim by every claude-connector-apple-* repo.
"""

from __future__ import annotations

import os
import subprocess
from pathlib import Path
from typing import Optional

_GENERIC = "~/Library/Application Support/Claude/uv-runtime/<version>/uv"


def responsible_binary() -> Optional[str]:
    """Absolute path of the `uv` that launched this process, if found."""
    # `uv run` exports its own path to the child as UV.
    uv = os.environ.get("UV")
    if uv and Path(uv).name == "uv" and Path(uv).is_file():
        return uv
    # Otherwise walk up the process tree looking for it.
    pid = os.getppid()
    for _ in range(4):
        if pid <= 1:
            return None
        try:
            out = subprocess.run(
                ["ps", "-o", "ppid=,comm=", "-p", str(pid)],
                capture_output=True, text=True, timeout=2,
            ).stdout.strip()
        except (OSError, subprocess.SubprocessError):
            return None
        parts = out.split(None, 1)
        if len(parts) != 2 or not parts[0].isdigit():
            return None
        ppid, comm = int(parts[0]), parts[1]
        if Path(comm).name == "uv" and Path(comm).is_absolute():
            return comm
        pid = ppid
    return None


def _display(path: Optional[str]) -> str:
    if not path:
        return _GENERIC
    home = str(Path.home())
    return "~" + path[len(home):] if path.startswith(home + "/") else path


def full_disk_access_steps() -> str:
    """How to grant Full Disk Access to the binary that actually needs it."""
    return (
        "To fix: System Settings → Privacy & Security → Full Disk Access → "
        "click +, press ⌘⇧G, and paste:\n\n"
        f"    {_display(responsible_binary())}\n\n"
        "Then quit Claude (⌘Q) and reopen it — macOS reads this permission "
        "only at launch. Claude Desktop runs extensions through this uv, so "
        "enabling Claude itself is not enough. One grant covers every Apple "
        "connector. The path changes when Claude Desktop updates its bundled "
        "uv; if this stops working after an update, grant the new path."
    )


def privacy_pane_steps(pane: str, access: str = "") -> str:
    """How to enable a prompted permission (Contacts, Calendars, Reminders)."""
    suffix = f" with {access}" if access else ""
    return (
        f"To fix: System Settings → Privacy & Security → {pane} → enable "
        f"'uv'{suffix}. It is listed as uv, not Claude, because Claude "
        "Desktop runs extensions through this binary:\n\n"
        f"    {_display(responsible_binary())}\n\n"
        "Then quit Claude (⌘Q) and reopen it. The path changes when Claude "
        "Desktop updates its bundled uv, which can mean answering the macOS "
        "prompt again."
    )
