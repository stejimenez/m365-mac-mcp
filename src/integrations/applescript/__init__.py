"""Modular AppleScript bridge domain mixins."""

from integrations.applescript.base import (
    AppleScriptBaseMixin,
    AppleScriptTimeoutError,
    _FS,
    _GS,
    _OFFICE_APPS,
    _RS,
    _US,
)
from integrations.applescript.excel_core import AppleScriptExcelCoreMixin
from integrations.applescript.excel_io import AppleScriptExcelIOMixin
from integrations.applescript.ppt_bridge import AppleScriptPowerPointMixin
from integrations.applescript.word_bridge import AppleScriptWordMixin

__all__ = [
    "AppleScriptBaseMixin",
    "AppleScriptExcelCoreMixin",
    "AppleScriptExcelIOMixin",
    "AppleScriptPowerPointMixin",
    "AppleScriptTimeoutError",
    "AppleScriptWordMixin",
    "_FS",
    "_GS",
    "_OFFICE_APPS",
    "_RS",
    "_US",
]
