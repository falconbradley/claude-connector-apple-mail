"""
Localhost HTTP redirector so chat UIs can open emails with one click.

Chat clients (Claude Desktop / Claude Code) block the message:// scheme
in rendered links but happily open http(s) URLs in the default browser.
This module runs a tiny localhost-only HTTP server; search results carry
an open_link like

    http://127.0.0.1:46325/open/268481?t=<token>

Clicking it opens the browser, which hits this server, which resolves
the message's RFC Message-ID and hands the message:// URL to macOS
`open` — Mail.app fronts with the message selected.

A second route, /thread/<message_id>, is the thread-scoped equivalent.
Mail.app has no conversation URL — its only registered schemes are
mailto:, message: and mail-pref-pane:, and message:// always opens one
message in its own window — so this route opens the *newest* message in
the conversation and says so on the result page. Reading the actual
back-and-forth is what the preview_thread card is for.

Security posture:
  - Bound to 127.0.0.1 only.
  - Every request must carry a per-install random token (persisted to
    disk so links in old chat transcripts keep working across restarts).
  - The only action is focusing Mail.app on a message; no message
    content is ever served over HTTP. The thread route serves a message
    *count* and nothing else — no subjects, senders or bodies.

Multiple server instances (Claude Desktop + a Claude Code session) share
the persisted port: the first instance binds it, later instances detect
the sibling via /ping and emit links pointing at the same port.
"""

from __future__ import annotations

import hmac
import json
import logging
import re
import secrets
import subprocess
import threading
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Callable, NamedTuple, Optional
from urllib.parse import parse_qs, quote, urlparse

logger = logging.getLogger("apple_mail_mcp.weblink")

_DEFAULT_PORT = 46325
_DEFAULT_STATE = (
    Path.home() / "Library" / "Application Support" / "apple-mail-mcp"
    / "weblink.json"
)
_PING_BODY = b"apple-mail-mcp-weblink"
_OPEN_PATH_RE = re.compile(r"^/open/(\d{1,12})$")
_THREAD_PATH_RE = re.compile(r"^/thread/(\d{1,12})$")

_PAGE = """<!doctype html><meta charset="utf-8">
<title>{title}</title>
<body style="font-family: -apple-system, sans-serif; margin: 3em; color: #333">
<h3>{title}</h3><p>{detail}</p>{script}</body>
"""

# Clicking a link hands the URL to the browser, which opens a tab for it --
# so every click used to leave a dead "Opened in Mail" tab behind, and they
# pile up fast. Mail.app coming to the front is the real confirmation, so on
# success the tab closes itself.
#
# A browser only lets a page close its own tab when that tab has no session
# history to go back to, which is exactly the case here: the tab was created
# for this URL. Where it is refused anyway (Safari, or a tab the user
# navigated by hand) the close silently fails and the page stays readable --
# which is the behaviour this replaces. Error pages never carry this: the
# whole point of those is to be read.
#
# A 204 No Content is NOT a substitute. It stops a *same-tab* navigation,
# but a tab the browser opened for the click still exists -- just blank and
# untitled, which is strictly worse than a page that explains itself.
_CLOSE_SCRIPT = (
    "<script>setTimeout(function(){try{window.close();}catch(e){}},150);</script>"
)


class ThreadAnchor(NamedTuple):
    """Where a /thread/<id> link should point Mail.app.

    Mail cannot be told to open a conversation, so a thread link opens one
    member of it: ``rfc_id`` is the newest message's Message-ID, and
    ``total`` is how many messages the conversation holds (shown on the
    result page so the user knows the rest exist).
    """

    message_id: int
    rfc_id: str
    total: int


class WebLinkServer:
    """Serves /open/<id> and /thread/<id> links that focus Mail.app."""

    def __init__(
        self,
        resolve_rfc_id: Callable[[int], Optional[str]],
        state_path: Optional[Path] = None,
        opener: Optional[Callable[[str], bool]] = None,
        preferred_port: int = _DEFAULT_PORT,
        resolve_thread: Optional[Callable[[int], Optional[ThreadAnchor]]] = None,
    ) -> None:
        self._resolve_rfc_id = resolve_rfc_id
        self._resolve_thread = resolve_thread
        self._opener = opener or self._open_with_macos
        self._state_path = state_path or _DEFAULT_STATE
        self._preferred_port = preferred_port
        self._lock = threading.Lock()
        self._httpd: Optional[ThreadingHTTPServer] = None
        self._started = False
        self.port: Optional[int] = None
        self.token: Optional[str] = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def open_link(self, message_id: int) -> Optional[str]:
        """Return the clickable http:// link for a message, or None if the
        redirector could not be started."""
        if not self.ensure_started():
            return None
        return f"http://127.0.0.1:{self.port}/open/{message_id}?t={self.token}"

    def thread_link(self, message_id: int) -> Optional[str]:
        """Return the clickable http:// link for the message's *conversation*.

        None when no thread resolver was supplied or the redirector could
        not start. The id in the path is the message the caller asked
        about; which member of the thread actually opens is decided at
        click time, so a link in an old transcript follows the thread as
        it grows.
        """
        if self._resolve_thread is None:
            return None
        if not self.ensure_started():
            return None
        return f"http://127.0.0.1:{self.port}/thread/{message_id}?t={self.token}"

    def ensure_started(self) -> bool:
        with self._lock:
            if self._started:
                return self.port is not None
            self._started = True
            try:
                self._start()
            except Exception:
                logger.exception("Web link redirector failed to start.")
                self.port = None
            return self.port is not None

    def shutdown(self) -> None:
        with self._lock:
            if self._httpd is not None:
                self._httpd.shutdown()
                self._httpd = None
            self._started = False

    # ------------------------------------------------------------------
    # Startup / state
    # ------------------------------------------------------------------

    def _start(self) -> None:
        state = self._load_state()
        self.token = state["token"]
        wanted_port = state.get("port", self._preferred_port)

        try:
            self._bind_and_serve(wanted_port)
        except OSError:
            if self._sibling_alive(wanted_port):
                # Another instance of this server owns the port; reuse it
                # in generated links, nothing to serve from here.
                logger.info("Web links served by sibling on port %d.", wanted_port)
                self.port = wanted_port
                return
            # Port taken by an unrelated process — fall back to ephemeral.
            self._bind_and_serve(0)

        state["port"] = self.port
        self._save_state(state)

    def _bind_and_serve(self, port: int) -> None:
        handler = self._make_handler()
        self._httpd = ThreadingHTTPServer(("127.0.0.1", port), handler)
        self._httpd.daemon_threads = True
        self.port = self._httpd.server_address[1]
        thread = threading.Thread(
            target=self._httpd.serve_forever, name="weblink", daemon=True
        )
        thread.start()
        logger.info("Web link redirector listening on 127.0.0.1:%d", self.port)

    def _sibling_alive(self, port: int) -> bool:
        try:
            with urllib.request.urlopen(
                f"http://127.0.0.1:{port}/ping?t={self.token}", timeout=1.0
            ) as resp:
                return resp.read(64).strip() == _PING_BODY
        except OSError:
            return False

    def _load_state(self) -> dict:
        try:
            state = json.loads(self._state_path.read_text())
            if isinstance(state.get("token"), str) and state["token"]:
                return state
        except (OSError, ValueError):
            pass
        return {"token": secrets.token_urlsafe(16), "port": self._preferred_port}

    def _save_state(self, state: dict) -> None:
        try:
            self._state_path.parent.mkdir(parents=True, exist_ok=True)
            self._state_path.write_text(json.dumps(state))
        except OSError as exc:
            logger.warning("Could not persist weblink state: %s", exc)

    # ------------------------------------------------------------------
    # Request handling
    # ------------------------------------------------------------------

    @staticmethod
    def _open_with_macos(mail_link: str) -> bool:
        proc = subprocess.run(
            ["open", mail_link], capture_output=True, timeout=15
        )
        return proc.returncode == 0

    def _make_handler(self) -> type:
        server = self

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, fmt: str, *args) -> None:
                logger.debug("weblink: " + fmt, *args)

            def _reply(
                self, status: int, title: str, detail: str = "",
                self_close: bool = False,
            ) -> None:
                body = _PAGE.format(
                    title=title,
                    detail=detail,
                    script=_CLOSE_SCRIPT if self_close else "",
                ).encode()
                self.send_response(status)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def do_GET(self) -> None:  # noqa: N802 (http.server API)
                parsed = urlparse(self.path)
                token = (parse_qs(parsed.query).get("t") or [""])[0]
                if not (server.token and hmac.compare_digest(token, server.token)):
                    self._reply(403, "Forbidden")
                    return

                if parsed.path == "/ping":
                    self.send_response(200)
                    self.send_header("Content-Length", str(len(_PING_BODY)))
                    self.end_headers()
                    self.wfile.write(_PING_BODY)
                    return

                match = _OPEN_PATH_RE.fullmatch(parsed.path)
                if match:
                    self._serve_message(int(match.group(1)))
                    return

                match = _THREAD_PATH_RE.fullmatch(parsed.path)
                if match:
                    self._serve_thread(int(match.group(1)))
                    return

                self._reply(404, "Not found")

            # -- routes ------------------------------------------------

            def _serve_message(self, message_id: int) -> None:
                try:
                    rfc_id = server._resolve_rfc_id(message_id)
                except Exception:
                    logger.exception("weblink: resolve failed for %d", message_id)
                    rfc_id = None
                if not rfc_id:
                    self._not_found(message_id)
                    return
                self._hand_to_mail(
                    rfc_id,
                    "Opened in Mail",
                    "The message should now be front-most in Mail.app. "
                    "You can close this tab.",
                )

            def _serve_thread(self, message_id: int) -> None:
                if server._resolve_thread is None:  # pragma: no cover - wired at startup
                    self._reply(404, "Not found")
                    return
                try:
                    anchor = server._resolve_thread(message_id)
                except Exception:
                    logger.exception("weblink: thread resolve failed for %d", message_id)
                    anchor = None
                if anchor is None or not anchor.rfc_id:
                    self._not_found(message_id)
                    return
                # Mail.app has no conversation URL, so the best it can do is
                # front the newest message. Say so rather than implying the
                # whole thread opened.
                if anchor.total > 1:
                    detail = (
                        f"Mail is showing the most recent of {anchor.total} messages "
                        "in this conversation. Mail.app can only open one message at "
                        "a time — with View \u25b8 Organize by Conversation turned on, "
                        "its message list groups the rest of the thread under the "
                        "same row. You can close this tab."
                    )
                else:
                    detail = (
                        "This conversation has just the one message, now front-most "
                        "in Mail.app. You can close this tab."
                    )
                self._hand_to_mail(anchor.rfc_id, "Opened thread in Mail", detail)

            # -- shared ------------------------------------------------

            def _not_found(self, message_id: int) -> None:
                self._reply(
                    404,
                    "Message not found",
                    f"No email with id {message_id} — it may have been "
                    "deleted, or Mail's index may have changed.",
                )

            def _hand_to_mail(self, rfc_id: str, title: str, detail: str) -> None:
                mail_link = f"message://{quote(f'<{rfc_id}>', safe='')}"
                if server._opener(mail_link):
                    self._reply(200, title, detail, self_close=True)
                else:
                    self._reply(
                        500,
                        "Could not open Mail",
                        f'Try this link directly: <a href="{mail_link}">'
                        f"{mail_link}</a>",
                    )

        return Handler
