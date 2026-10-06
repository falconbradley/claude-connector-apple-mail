"""Fail unless the running server's tools match manifest.json's tool list.

Importing the server also catches SDK drift (e.g. mcp 2.0 removing
FastMCP) that static tests pass happily against. Run inside the project
environment: `uv run python .github/check_tools.py`.
Shared verbatim by every claude-connector-apple-* repo.
"""

import asyncio
import importlib
import json
import sys

manifest = json.load(open("manifest.json"))
pkg = manifest["name"].replace("-", "_") + "_mcp"
server = importlib.import_module(f"{pkg}.server")
code = {t.name for t in asyncio.run(server.mcp.list_tools())}
listed = {t["name"] for t in manifest["tools"]}
if code != listed:
    print("::error::manifest/server tool mismatch")
    print(f"  only in manifest.json: {sorted(listed - code)}")
    print(f"  only in the server:    {sorted(code - listed)}")
    sys.exit(1)
print(f"✓ {server.mcp.name} serves {len(code)} tools, matching manifest.json.")
