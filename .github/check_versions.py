"""Fail unless manifest.json, pyproject.toml, __init__.py (and a tag) agree.

The packer reads only manifest.json and the server reports __init__.py's
version over the wire, so any drift ships a mislabelled bundle.
Shared verbatim by every claude-connector-apple-* repo.

usage: python3 .github/check_versions.py [vX.Y.Z]
"""

import json
import re
import sys
from pathlib import Path

manifest = json.loads(Path("manifest.json").read_text())
pkg = manifest["name"].replace("-", "_") + "_mcp"
found = {
    "manifest.json": manifest["version"],
    "pyproject.toml": re.search(
        r'^version = "(.*)"', Path("pyproject.toml").read_text(), re.M
    ).group(1),
    f"src/{pkg}/__init__.py": re.search(
        r'^__version__ = "(.*)"',
        Path(f"src/{pkg}/__init__.py").read_text(), re.M,
    ).group(1),
}
lock = re.search(
    rf'name = "{manifest["name"]}-mcp"\nversion = "(.*)"',
    Path("uv.lock").read_text(),
)
if lock:
    found["uv.lock"] = lock.group(1)
if len(sys.argv) > 1:
    found["tag"] = sys.argv[1].removeprefix("v")

for where, version in found.items():
    print(f"{where}: {version}")
if len(set(found.values())) != 1:
    print("::error::version mismatch — set every file above to the same version")
    sys.exit(1)
