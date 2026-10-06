"""Tests for launcher: naming the uv that macOS holds responsible.

Runs under pytest or standalone (`uv run python tests/test_launcher.py`).
"""

from __future__ import annotations

import os
import stat
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from apple_mail_mcp import launcher  # noqa: E402


def _fake_uv(tmp: str) -> str:
    path = Path(tmp) / "uv-runtime" / "uv-9.9.9-darwin-arm64" / "uv"
    path.parent.mkdir(parents=True)
    path.write_text("#!/bin/sh\n")
    path.chmod(path.stat().st_mode | stat.S_IXUSR)
    return str(path)


def _with_env(value):
    old = os.environ.get("UV")
    if value is None:
        os.environ.pop("UV", None)
    else:
        os.environ["UV"] = value
    return old


def test_prefers_uv_env_var():
    with tempfile.TemporaryDirectory() as tmp:
        fake = _fake_uv(tmp)
        old = _with_env(fake)
        try:
            assert launcher.responsible_binary() == fake
            assert fake in launcher.full_disk_access_steps()
        finally:
            _with_env(old)


def test_ignores_env_var_that_is_not_uv():
    old = _with_env(sys.executable)
    try:
        found = launcher.responsible_binary()
        assert found is None or Path(found).name == "uv"
    finally:
        _with_env(old)


def test_falls_back_to_generic_path():
    old = _with_env(None)
    try:
        launcher.responsible_binary = lambda: None  # type: ignore[assignment]
        text = launcher.full_disk_access_steps()
        assert "uv-runtime/<version>/uv" in text
        assert "Full Disk Access" in text
    finally:
        _with_env(old)
        import importlib
        importlib.reload(launcher)


def test_home_is_abbreviated():
    home = str(Path.home())
    assert launcher._display(home + "/x/uv") == "~/x/uv"
    assert launcher._display("/opt/uv") == "/opt/uv"


def test_pane_steps_name_uv_not_claude():
    text = launcher.privacy_pane_steps("Calendars", "Full Access")
    assert "enable 'uv' with Full Access" in text
    assert "Calendars" in text


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for t in tests:
        try:
            t()
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"FAIL {t.__name__}: {exc!r}")
    print(f"{len(tests) - failed}/{len(tests)} checks passed")
    sys.exit(1 if failed else 0)
