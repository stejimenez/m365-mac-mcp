"""
Input validation, path resolution, and JSON serialization utilities
for Microsoft 365 for Mac MCP (m365-mac-mcp).
"""

import datetime
import math
import re
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, List, Optional

_CELL_REF_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$")
_RANGE_REF_RE = re.compile(
    r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*:\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$"
)


def resolve_file_path(
    file_path: str,
    must_exist: bool = False,
    extensions: Optional[List[str]] = None,
) -> Path:
    """Expand ``~`` and resolve a POSIX file path safely.

    Handles paths with spaces, parentheses, apostrophes, and Unicode cleanly.

    Args:
        file_path: Input path string.
        must_exist: Raise ValueError if the path does not exist.
        extensions: Optional list of allowed lowercase extensions (e.g. ['.docx']).

    Returns:
        Resolved pathlib.Path.
    """
    if not file_path or not isinstance(file_path, str) or not file_path.strip():
        raise ValueError("File path cannot be empty")

    if "\x00" in file_path:
        raise ValueError("File path contains invalid null byte")

    path = Path(file_path).expanduser().resolve()

    if must_exist and not path.exists():
        raise FileNotFoundError(f"File does not exist: {path}")

    if extensions:
        allowed = {
            ext.lower() if ext.startswith(".") else f".{ext.lower()}"
            for ext in extensions
        }
        if path.suffix.lower() not in allowed:
            raise ValueError(
                f"File '{path.name}' must have one of these extensions: {sorted(allowed)}"
            )

    return path


def validate_file_path(
    file_path: str,
    must_exist: bool = False,
    extensions: Optional[List[str]] = None,
) -> str:
    """Validate and return a resolved file path as a string."""
    return str(resolve_file_path(file_path, must_exist=must_exist, extensions=extensions))


def validate_cell_ref(cell: str) -> str:
    """Validate an A1-style Excel cell reference (e.g. 'A1', 'BC42')."""
    if not cell or not isinstance(cell, str):
        raise ValueError("Cell reference must be a non-empty string")
    cleaned = cell.strip().upper().replace("$", "")
    if not _CELL_REF_RE.match(cleaned):
        raise ValueError(f"Invalid Excel cell reference: '{cell}' (expected e.g. 'A1')")
    return cleaned


def validate_range_ref(range_ref: str) -> str:
    """Validate an A1-style Excel cell or range reference (e.g. 'A1' or 'A1:D10')."""
    if not range_ref or not isinstance(range_ref, str):
        raise ValueError("Range reference must be a non-empty string")
    cleaned = range_ref.strip().upper().replace("$", "")
    if _CELL_REF_RE.match(cleaned) or _RANGE_REF_RE.match(cleaned):
        return cleaned
    raise ValueError(
        f"Invalid Excel range reference: '{range_ref}' (expected e.g. 'A1' or 'A1:D10')"
    )


def sanitize_excel_write_value(value: Any) -> Any:
    """Sanitize a value before writing to an Excel cell so there are no silent coercion bugs.

    - Preserves ``bool`` (checked before ``int``), ``int``, ``float``, ``str``, ``None``,
      ``datetime``, ``date``, ``time``.
    - Converts ``NaN`` / ``Inf`` / ``-Inf`` floats to ``None`` (Excel cannot store IEEE NaN/Inf).
    - Converts ``Decimal`` to ``float``.
    - Converts unsupported complex objects to ``str``.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        return value
    if isinstance(value, Decimal):
        f = float(value)
        return None if (math.isnan(f) or math.isinf(f)) else f
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value
    if isinstance(value, str):
        return value
    return str(value)


def sanitize_for_json(value: Any) -> Any:
    """Recursively convert values into strictly JSON-serializable primitives.

    Ensures MCP responses never fail serialization in Antigravity or Bionic 1.2.1+10.
    """
    if value is None or isinstance(value, (bool, int, str)):
        return value
    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return None
        return value
    if isinstance(value, Decimal):
        f = float(value)
        return None if (math.isnan(f) or math.isinf(f)) else f
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    if isinstance(value, dict):
        return {str(k): sanitize_for_json(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [sanitize_for_json(v) for v in value]
    return str(value)


def validate_input(data: Dict[str, Any], schema: Dict[str, Any]) -> Dict[str, Any]:
    """Validate input data against a simple field schema."""
    validated: Dict[str, Any] = {}

    for field, rules in schema.items():
        value = data.get(field)

        if rules.get("required", False) and value is None:
            raise ValueError(f"Required field '{field}' is missing")

        if value is None:
            continue

        expected_type = rules.get("type")
        if expected_type and not isinstance(value, expected_type):
            raise ValueError(f"Field '{field}' must be of type {expected_type}")

        if isinstance(value, str):
            min_length = rules.get("min_length")
            max_length = rules.get("max_length")
            if min_length and len(value) < min_length:
                raise ValueError(f"Field '{field}' must be at least {min_length} characters")
            if max_length and len(value) > max_length:
                raise ValueError(f"Field '{field}' must be at most {max_length} characters")
            pattern = rules.get("pattern")
            if pattern and not re.match(pattern, value):
                raise ValueError(f"Field '{field}' does not match required pattern")

        if isinstance(value, (int, float)) and not isinstance(value, bool):
            min_value = rules.get("min_value")
            max_value = rules.get("max_value")
            if min_value is not None and value < min_value:
                raise ValueError(f"Field '{field}' must be at least {min_value}")
            if max_value is not None and value > max_value:
                raise ValueError(f"Field '{field}' must be at most {max_value}")

        choices = rules.get("choices")
        if choices and value not in choices:
            raise ValueError(f"Field '{field}' must be one of {choices}")

        validated[field] = value

    return validated
