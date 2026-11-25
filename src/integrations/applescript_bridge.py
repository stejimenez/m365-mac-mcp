"""
AppleScript Bridge for Microsoft 365 for Mac MCP (m365-mac-mcp — macOS 27.x / Office 16.x).

Facade composing modular domain mixins (`integrations.applescript.*`) for hardened,
injection-safe, timeout-protected live automation of Microsoft Word, Excel, and PowerPoint.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from integrations.applescript import (
    AppleScriptBaseMixin,
    AppleScriptExcelCoreMixin,
    AppleScriptExcelIOMixin,
    AppleScriptPowerPointMixin,
    AppleScriptTimeoutError,
    AppleScriptWordMixin,
    _FS,
    _GS,
    _OFFICE_APPS,
    _RS,
    _US,
)


class AppleScriptBridge(
    AppleScriptWordMixin,
    AppleScriptExcelCoreMixin,
    AppleScriptExcelIOMixin,
    AppleScriptPowerPointMixin,
    AppleScriptBaseMixin,
):
    """Hardened bridge for communicating with Microsoft Office apps via AppleScript."""


__all__ = [
    "AppleScriptBridge",
    "AppleScriptTimeoutError",
    "_FS",
    "_GS",
    "_OFFICE_APPS",
    "_RS",
    "_US",
]
