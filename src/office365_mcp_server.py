#!/usr/bin/env python3
"""
Backwards-compatible entrypoint shim for m365-mac-mcp.
Deprecated: update your MCP client Arguments to `src/m365_mac_mcp_server.py`.
"""

import sys
from pathlib import Path

_SRC_DIR = str(Path(__file__).resolve().parent)
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)

import m365_mac_mcp_server as _server  # noqa: E402
from m365_mac_mcp_server import *  # noqa: E402, F401, F403

_server.logger.warning(
    "src/office365_mcp_server.py is deprecated; update your MCP Arguments to src/m365_mac_mcp_server.py."
)

for _attr in (
    "_ACTIVE_PROFILE_LABEL",
    "_PROFILE_APPS",
    "_PROFILE_COMPACT",
    "_REGISTERED_TOOL_NAMES",
    "_parse_mcp_profile",
):
    if hasattr(_server, _attr):
        globals()[_attr] = getattr(_server, _attr)

if __name__ == "__main__":
    _server.main()
