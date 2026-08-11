"""
Hybrid bridge: fast SQLite reads, JXA writes.

Routes every read operation to EnvelopeIndexBridge (Mail.app's local
SQLite store — milliseconds) and falls back to the AppleScript/JXA
MailBridge when the fast path is unavailable (typically: the host process
has no Full Disk Access) or fails mid-call. Write operations (drafts,
flags) always go through JXA, since the local store must never be
mutated behind Mail.app's back.

Message ids are shared between both engines: the Envelope Index ROWID is
the same integer id Mail.app exposes via its scripting interface.

Set APPLE_MAIL_MCP_DISABLE_FAST=1 to force the JXA path for everything.
"""

from __future__ import annotations

import logging
import os
import time
from typing import Any, Optional

from .applescript import MailBridge
from .envelope import EnvelopeIndexBridge, EnvelopeUnavailable

logger = logging.getLogger("apple_mail_mcp.hybrid")

# How long to wait before re-probing the fast path after a failed init
# (e.g. the user grants Full Disk Access mid-session).
_ENVELOPE_RETRY_INTERVAL = 60.0

_READ_METHODS = frozenset(
    {
        "get_stats",
        "list_mailboxes",
        "search_messages",
        "get_message",
        "get_message_id_header",
        "get_message_source",
        "get_thread_messages",
        "list_attachments",
        "get_flag",
    }
)


class HybridBridge:
    """Drop-in replacement for MailBridge with a fast read path."""

    def __init__(self) -> None:
        self._mail: Optional[MailBridge] = None
        self._env: Optional[EnvelopeIndexBridge] = None
        self._env_error: Optional[str] = None
        self._env_last_attempt = 0.0
        # Which engine served the most recent read ("sqlite" | "applescript")
        self.last_engine: Optional[str] = None

    # ------------------------------------------------------------------
    # Engine acquisition
    # ------------------------------------------------------------------

    def _envelope(self) -> Optional[EnvelopeIndexBridge]:
        if os.environ.get("APPLE_MAIL_MCP_DISABLE_FAST"):
            return None
        if self._env is not None:
            return self._env
        now = time.monotonic()
        if now - self._env_last_attempt < _ENVELOPE_RETRY_INTERVAL:
            return None
        self._env_last_attempt = now
        try:
            self._env = EnvelopeIndexBridge()
            self._env_error = None
            logger.info("Fast read path active (Envelope Index).")
        except EnvelopeUnavailable as exc:
            self._env_error = str(exc)
            logger.warning(
                "Fast read path unavailable, using AppleScript bridge: %s", exc
            )
        return self._env

    def _jxa(self) -> MailBridge:
        if self._mail is None:
            try:
                self._mail = MailBridge()
                logger.info("AppleScript bridge initialised.")
            except (RuntimeError, OSError) as exc:
                hint = ""
                if self._env is None and self._env_error:
                    hint = (
                        "\n\nNote: the fast local-store read path is also "
                        f"unavailable: {self._env_error}"
                    )
                raise RuntimeError(
                    "Could not connect to Mail.app. Make sure Mail.app is open "
                    "and that this process has Automation permission in System "
                    "Settings -> Privacy & Security -> Automation."
                    f"\n\nUnderlying error: {exc}{hint}"
                )
        return self._mail

    def fast_path_status(self) -> str:
        if os.environ.get("APPLE_MAIL_MCP_DISABLE_FAST"):
            return "disabled via APPLE_MAIL_MCP_DISABLE_FAST"
        if self._env is not None:
            return "active"
        return self._env_error or "not yet probed"

    # ------------------------------------------------------------------
    # Read dispatch
    # ------------------------------------------------------------------

    def _read(self, method: str, *args: Any, **kwargs: Any) -> Any:
        assert method in _READ_METHODS
        env = self._envelope()
        if env is not None:
            try:
                result = getattr(env, method)(*args, **kwargs)
                self.last_engine = "sqlite"
                return result
            except ValueError:
                # Semantic "not found" — trust it, don't burn 30s of JXA
                # re-checking.
                self.last_engine = "sqlite"
                raise
            except Exception:
                logger.exception(
                    "Fast path failed for %s; falling back to AppleScript.",
                    method,
                )
        result = getattr(self._jxa(), method)(*args, **kwargs)
        self.last_engine = "applescript"
        return result

    def get_stats(self) -> dict:
        return self._read("get_stats")

    def list_mailboxes(self) -> list[dict]:
        return self._read("list_mailboxes")

    def search_messages(self, **kwargs: Any) -> tuple[int, list[dict]]:
        return self._read("search_messages", **kwargs)

    def get_message(self, message_id: int) -> Optional[dict]:
        return self._read("get_message", message_id)

    def get_message_id_header(self, message_id: int) -> Optional[str]:
        return self._read("get_message_id_header", message_id)

    def get_message_source(self, message_id: int) -> Optional[str]:
        return self._read("get_message_source", message_id)

    def get_thread_messages(self, message_id: int) -> list[dict]:
        return self._read("get_thread_messages", message_id)

    def list_attachments(self, message_id: int) -> list[dict]:
        return self._read("list_attachments", message_id)

    def get_attachment(
        self, message_id: int, attachment_index: int
    ) -> Optional[tuple[str, str, bytes]]:
        env = self._envelope()
        if env is not None:
            try:
                result = env.get_attachment(message_id, attachment_index)
                if result is not None:
                    self.last_engine = "sqlite"
                    return result
                # None = body not on disk (stub / not downloaded): let
                # Mail.app fetch it from the server via JXA.
            except Exception:
                logger.exception("Fast attachment fetch failed; trying JXA.")
        result = self._jxa().get_attachment(message_id, attachment_index)
        self.last_engine = "applescript"
        return result

    def get_flag(self, message_id: int) -> dict:
        """Flag status. The index answers is_flagged instantly; the exact
        color comes from JXA only when the message is actually flagged and
        the index couldn't provide a color."""
        env = self._envelope()
        if env is not None:
            try:
                result = env.get_flag(message_id)
                self.last_engine = "sqlite"
                if not result["is_flagged"]:
                    return {
                        "is_flagged": False,
                        "color_index": -1,
                        "flag_color": None,
                    }
                color_index = result.get("color_index", -1)
                if 0 <= color_index <= 6:
                    from .applescript import _FLAG_COLOR_ORDER

                    return {
                        "is_flagged": True,
                        "color_index": color_index,
                        "flag_color": _FLAG_COLOR_ORDER[color_index],
                    }
                # Flagged but color unknown -> ask Mail.app, tolerate failure
                try:
                    return self._jxa().get_flag(message_id)
                except Exception:
                    logger.warning(
                        "Could not resolve flag color for %d via JXA.", message_id
                    )
                    return {
                        "is_flagged": True,
                        "color_index": -1,
                        "flag_color": None,
                    }
            except ValueError:
                raise
            except Exception:
                logger.exception("Fast get_flag failed; falling back to JXA.")
        result = self._jxa().get_flag(message_id)
        self.last_engine = "applescript"
        return result

    def get_selected_messages(self) -> list[dict]:
        """UI-state read — only Mail.app knows the current selection, so
        this is always JXA regardless of the fast path."""
        result = self._jxa().get_selected_messages()
        self.last_engine = "applescript"
        return result

    # ------------------------------------------------------------------
    # Writes: always JXA (Mail.app owns the store)
    # ------------------------------------------------------------------

    def set_flag(self, message_id: int, flag: Optional[str] = None) -> dict:
        return self._jxa().set_flag(message_id, flag)

    def create_draft(self, **kwargs: Any) -> dict:
        return self._jxa().create_draft(**kwargs)

    def create_reply_draft(self, *args: Any, **kwargs: Any) -> dict:
        return self._jxa().create_reply_draft(*args, **kwargs)
