"""Shared constants and named-style presets for ExcelController."""

import re
from typing import Any, Dict

# Built-in Excel Named Cell Style / Financial Model Presets
# Registered on demand via wb.add_named_style() so they work even on brand-new openpyxl workbooks.
_NAMED_CELL_STYLE_PRESETS: Dict[str, Dict[str, Any]] = {
    "Normal": {
        "font": {"name": "Calibri", "size": 11, "color": "000000"},
    },
    "Good": {
        "font": {"name": "Calibri", "size": 11, "color": "006100"},
        "fill": "C6EFCE",
    },
    "Bad": {
        "font": {"name": "Calibri", "size": 11, "color": "9C0006"},
        "fill": "FFC7CE",
    },
    "Neutral": {
        "font": {"name": "Calibri", "size": 11, "color": "9C6500"},
        "fill": "FFEB9C",
    },
    "Input": {
        "font": {"name": "Calibri", "size": 11, "color": "3F3F76"},
        "fill": "FFCC99",
        "border": {"all": "thin", "color": "7F7F7F"},
    },
    "Calculation": {
        "font": {"name": "Calibri", "size": 11, "bold": True, "color": "FA7D00"},
        "fill": "F2F2F2",
        "border": {"all": "thin", "color": "7F7F7F"},
    },
    "Check Cell": {
        "font": {"name": "Calibri", "size": 11, "bold": True, "color": "FFFFFF"},
        "fill": "A5A5A5",
        "border": {"all": "double", "color": "3F3F3F"},
    },
    "Linked Cell": {
        "font": {"name": "Calibri", "size": 11, "color": "FA7D00"},
        "border": {"bottom": "double", "color": "FF8001"},
    },
    "Note": {
        "font": {"name": "Calibri", "size": 11, "color": "000000"},
        "fill": "FFFFCC",
        "border": {"all": "thin", "color": "B2B2B2"},
    },
    "Output": {
        "font": {"name": "Calibri", "size": 11, "bold": True, "color": "3F3F3F"},
        "fill": "F2F2F2",
        "border": {"all": "thin", "color": "3F3F3F"},
    },
    "Warning Text": {
        "font": {"name": "Calibri", "size": 11, "color": "FF0000"},
    },
    "Title": {
        "font": {"name": "Calibri", "size": 18, "bold": True, "color": "1F4E79"},
    },
    "Heading 1": {
        "font": {"name": "Calibri", "size": 15, "bold": True, "color": "1F4E79"},
        "border": {"bottom": "thick", "color": "4472C4"},
    },
    "Heading 2": {
        "font": {"name": "Calibri", "size": 13, "bold": True, "color": "1F4E79"},
        "border": {"bottom": "thick", "color": "8EA9DB"},
    },
    "Heading 3": {
        "font": {"name": "Calibri", "size": 11, "bold": True, "color": "1F4E79"},
        "border": {"bottom": "medium", "color": "8EA9DB"},
    },
    "Heading 4": {
        "font": {"name": "Calibri", "size": 11, "bold": True, "color": "1F4E79"},
    },
    "Total": {
        "font": {"name": "Calibri", "size": 11, "bold": True, "color": "000000"},
        "border": {"top": "thin", "bottom": "double", "color": "4472C4"},
    },
}

_TABLE_STYLE_RE = re.compile(
    r"^(TableStyleLight([1-9]|1[0-9]|2[01])|TableStyleMedium([1-9]|1[0-9]|2[0-8])|TableStyleDark([1-9]|1[01]))$"
)

