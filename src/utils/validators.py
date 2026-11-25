"""
Input validation, path resolution, and JSON serialization utilities
for Microsoft 365 for Mac MCP (m365-mac-mcp).
"""

import contextvars
import datetime
import math
import os
import re
import tempfile
from decimal import Decimal
from pathlib import Path
from typing import Any, Dict, List, Optional

_CELL_REF_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$")
_RANGE_REF_RE = re.compile(
    r"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*:\$?[A-Za-z]{1,3}\$?[1-9][0-9]*$"
)
_MCP_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_ACTIVE_CLIENT_PROJECT_DIR: Optional[Path] = None
_SESSION_ID_CTX: contextvars.ContextVar[Optional[str]] = contextvars.ContextVar(
    "m365_mcp_session_id", default=None
)
_SESSION_PROJECT_DIRS: Dict[str, Optional[Path]] = {}
_MAX_TRACKED_SESSIONS = 128


def _is_path_inside(child: Path, parent: Path) -> bool:
    try:
        return child == parent or child.is_relative_to(parent)
    except AttributeError:
        return child == parent or str(child).startswith(str(parent) + "/")


def _is_inside_project_root(resolved: Path) -> bool:
    return _is_path_inside(resolved, _MCP_PROJECT_ROOT)


def set_current_session_id(session_id: Optional[str]) -> contextvars.Token:
    """Bind the active MCP session identifier for async-safe working_dir isolation."""
    return _SESSION_ID_CTX.set(session_id)


def reset_current_session_id(token: contextvars.Token) -> None:
    """Restore the previous MCP session context token."""
    _SESSION_ID_CTX.reset(token)


def get_current_session_id() -> Optional[str]:
    """Return the active MCP session identifier, or None outside an MCP request context."""
    return _SESSION_ID_CTX.get()


def clear_session_project_dir(session_id: Optional[str] = None) -> None:
    """Remove a tracked session's project directory mapping."""
    target = session_id if session_id is not None else _SESSION_ID_CTX.get()
    if target is not None:
        _SESSION_PROJECT_DIRS.pop(target, None)


def set_client_project_dir(project_dir: Optional[str | Path]) -> Optional[Path]:
    """Register the active client/Bionic Project Working Folder (Priority #2; never inside MCP project root).

    Isolated per MCP session when an async session context is active; falls back to
    process-wide storage for single-client / direct Python callers.
    """
    global _ACTIVE_CLIENT_PROJECT_DIR
    sess_id = _SESSION_ID_CTX.get()

    if project_dir is None or not str(project_dir).strip():
        if sess_id is not None:
            _SESSION_PROJECT_DIRS[sess_id] = None
        else:
            _ACTIVE_CLIENT_PROJECT_DIR = None
        return None

    raw = str(project_dir).strip()
    if raw.startswith("file://"):
        from urllib.parse import unquote, urlparse
        raw = unquote(urlparse(raw).path)
    candidate = Path(raw).expanduser().resolve()
    if _is_inside_project_root(candidate):
        return None

    if sess_id is not None:
        if len(_SESSION_PROJECT_DIRS) >= _MAX_TRACKED_SESSIONS and sess_id not in _SESSION_PROJECT_DIRS:
            oldest_key = next(iter(_SESSION_PROJECT_DIRS))
            _SESSION_PROJECT_DIRS.pop(oldest_key, None)
        _SESSION_PROJECT_DIRS[sess_id] = candidate
    else:
        _ACTIVE_CLIENT_PROJECT_DIR = candidate
    return candidate


def get_client_project_dir() -> Optional[Path]:
    """Return the active client/Bionic Project Working Folder if safely configured outside MCP project root."""
    sess_id = _SESSION_ID_CTX.get()
    if sess_id is not None:
        candidate = _SESSION_PROJECT_DIRS.get(sess_id)
        if candidate is not None and not _is_inside_project_root(candidate):
            return candidate
        return None
    if _ACTIVE_CLIENT_PROJECT_DIR is not None and not _is_inside_project_root(_ACTIVE_CLIENT_PROJECT_DIR):
        return _ACTIVE_CLIENT_PROJECT_DIR
    return None


def get_blocked_write_dirs() -> List[Path]:
    """Return normalized directories blocked for writes via M365_MCP_BLOCKED_DIRS (colon- or semicolon-separated)."""
    raw = (os.getenv("M365_MCP_BLOCKED_DIRS") or "").strip()
    if not raw:
        return []
    parts = [p.strip() for p in re.split(r"[:;]", raw) if p.strip()]
    resolved_dirs: List[Path] = []
    for part in parts:
        try:
            resolved_dirs.append(Path(part).expanduser().resolve())
        except Exception:
            continue
    return resolved_dirs


def is_write_restricted_to_project_dir() -> bool:
    """Return True when M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR enforces strict project-folder-only writes."""
    raw = (os.getenv("M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR") or "").strip().lower()
    return raw in {"1", "true", "yes", "on"}


def is_within_trusted_project_dir(path: str | Path) -> bool:
    """Return True ONLY when `path` resides strictly inside the active client/Bionic Project Working Folder."""
    proj = get_client_project_dir()
    if proj is None:
        return False
    resolved = Path(path).expanduser().resolve()
    if _is_inside_project_root(resolved):
        return False
    return _is_path_inside(resolved, proj)


def get_grant_access_timeout(default: float = 45.0) -> float:
    """Return extended timeout (30-50s, safely under MCP client 60-90s cap) for external Grant Access prompts."""
    raw = (os.getenv("M365_MCP_GRANT_ACCESS_TIMEOUT") or "").strip()
    if raw:
        try:
            val = float(raw)
            return max(15.0, min(val, 55.0))
        except ValueError:
            pass
    return default


def get_background_mode() -> str:
    """Return active background/focus-protection mode from M365_MCP_BACKGROUND_MODE.

    Modes:
      - 'no_focus' (default): Never steal OS focus (`open -g -j`, ignore caller `activate=True`),
        minimize newly opened MCP Office windows, and restore any previously active user
        workbook/document inside the same Office app so stray keystrokes cannot hit MCP files.
      - 'strict': Same focus suppression, but hides newly opened MCP windows (`visible = false`)
        instead of minimizing them for maximum same-app isolation.
      - 'off': Legacy behavior — honors caller's `activate` flag without window isolation.
    """
    raw = (
        os.getenv("M365_MCP_BACKGROUND_MODE")
        or os.getenv("OFFICE_MCP_BACKGROUND_MODE")
        or "no_focus"
    ).strip().lower()
    if raw in {"", "no_focus", "nofocus", "no-focus", "on", "true", "1", "yes", "default", "background"}:
        return "no_focus"
    if raw in {"strict", "hidden", "hide"}:
        return "strict"
    if raw in {"off", "false", "0", "no", "none", "legacy"}:
        return "off"
    return "no_focus"


def get_default_output_dir() -> Optional[Path]:
    """Return the optional safe default output folder from M365_MCP_OUTPUT_DIR (Priority #3; never inside MCP project root)."""
    raw = (os.getenv("M365_MCP_OUTPUT_DIR") or os.getenv("OFFICE365_MCP_OUTPUT_DIR") or "").strip()
    if not raw:
        return None
    candidate = Path(raw).expanduser().resolve()
    if _is_inside_project_root(candidate):
        return None
    return candidate


def assert_safe_write_path(path: Path, raw_input: Optional[str] = None) -> Path:
    """Block writing user Office deliverables inside the MCP project folder or outside configured write fences."""
    resolved = path.expanduser().resolve()
    if _is_inside_project_root(resolved):
        rel_note = ""
        if raw_input and not Path(raw_input).expanduser().is_absolute():
            rel_note = f" (bare relative path '{raw_input}' resolved against server CWD)"
        raise PermissionError(
            f"Refusing to write output file inside the MCP server project folder ('{_MCP_PROJECT_ROOT}'){rel_note}. "
            "Provide an explicit absolute or ~/ user path (such as your Project Working Folder, '~/Desktop/...', or '~/Documents/...'), "
            "or set optional M365_MCP_OUTPUT_DIR."
        )

    for blocked in get_blocked_write_dirs():
        if _is_path_inside(resolved, blocked):
            raise PermissionError(
                f"Refusing to write output file '{resolved}' inside blocked directory '{blocked}' "
                "(configured via M365_MCP_BLOCKED_DIRS). Choose a destination outside blocked directories."
            )

    if is_write_restricted_to_project_dir():
        temp_staging_root = (Path(tempfile.gettempdir()) / "m365-mac-mcp").resolve()
        if not _is_path_inside(resolved, temp_staging_root):
            proj = get_client_project_dir()
            allowed_root = proj if proj is not None else get_default_output_dir()
            if allowed_root is None:
                raise PermissionError(
                    "Refusing to write output file: M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR is enabled, "
                    "but no active Project Working Folder or M365_MCP_OUTPUT_DIR is configured. "
                    "Call check_office_status(working_dir='...') first or configure M365_MCP_OUTPUT_DIR."
                )
            if not _is_path_inside(resolved, allowed_root):
                raise PermissionError(
                    f"Refusing to write output file '{resolved}' outside the active Project Working Folder "
                    f"'{allowed_root}' (M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR is enabled). "
                    f"Save inside '{allowed_root}' or disable M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR."
                )

    return resolved


def resolve_file_path(
    file_path: str,
    must_exist: bool = False,
    extensions: Optional[List[str]] = None,
    for_write: bool = False,
) -> Path:
    """Expand ``~`` and resolve a POSIX file path safely.

    Strict resolution precedence:
      1. Explicit absolute / ``~`` user path (always wins).
      2. Active client / Bionic Project Working Folder (MCP ``roots/list`` or ``check_office_status(working_dir=...)``).
      3. Optional ``M365_MCP_OUTPUT_DIR`` environment variable.
      4. Process CWD (writes inside the MCP server's own project folder are ALWAYS blocked).
    """
    if not file_path or not isinstance(file_path, str) or not file_path.strip():
        raise ValueError("File path cannot be empty")

    if "\x00" in file_path:
        raise ValueError("File path contains invalid null byte")

    raw_expanded = Path(file_path).expanduser()
    if raw_expanded.is_absolute():
        # Priority 1: explicit user path
        path = raw_expanded.resolve()
    else:
        # Priority 2: client/Bionic Project Working Folder; Priority 3: M365_MCP_OUTPUT_DIR
        project_dir = get_client_project_dir()
        default_dir = get_default_output_dir()
        base_dirs = [d for d in (project_dir, default_dir) if d is not None]

        if for_write and base_dirs:
            path = (base_dirs[0] / raw_expanded).resolve()
        elif not for_write and base_dirs:
            chosen: Optional[Path] = None
            for base in base_dirs:
                cand = (base / raw_expanded).resolve()
                if cand.exists():
                    chosen = cand
                    break
            if chosen is not None:
                path = chosen
            else:
                cwd_cand = raw_expanded.resolve()
                if cwd_cand.exists() and not _is_inside_project_root(cwd_cand):
                    path = cwd_cand
                else:
                    path = (base_dirs[0] / raw_expanded).resolve()
        else:
            path = raw_expanded.resolve()

    if for_write:
        assert_safe_write_path(path, raw_input=file_path)

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
        # Decode double-escaped Unicode literals (e.g. 'WACC \\u2193' -> 'WACC ↓', '-\\u0394NWC' -> '-ΔNWC')
        # inside display text (never touch Excel formulas starting with '=').
        if not value.startswith("=") and "\\u" in value:
            value = re.sub(
                r"\\u([0-9a-fA-F]{4})",
                lambda m: chr(int(m.group(1), 16)),
                value,
            )
        return value
    return str(value)


_EXCEL_FORMULA_QUOTE_TOKEN_RE = re.compile(
    r'("(?:""|[^"])*"|\'(?:\'\'|[^\'])*\'!|\'(?:\'\'|[^\'])*\')'
)


def sanitize_excel_formula_quotes(formula: str) -> str:
    """Repair common LLM/JSON-wrapper quote bugs inside Excel formulas.

    1. Unescapes backslash-escaped double quotes (``\\"`` -> ``"``), which are never valid in
       Excel formula syntax (Excel escapes quotes via ``""``) and are commonly leaked when
       double-encoded JSON wrappers (such as Bionic's ``params_json``) triple-escape quotes.
    2. Converts standalone single-quoted string literals (e.g., ``=IF(B2>0,'Go','No-Go')``)
       into valid Excel double-quoted string literals (``=IF(B2>0,"Go","No-Go")``), while
       strictly preserving quoted sheet references (``'Project Cash Flows'!``) and apostrophes
       inside double-quoted strings (``"Bob's Store"``).
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula

    fixed = formula
    if '\\"' in fixed:
        fixed = fixed.replace('\\"', '"')

    if "'" in fixed:
        def _repl(match: re.Match) -> str:
            tok = match.group(1)
            # Keep valid double-quoted strings and valid 'Sheet Name'! references untouched
            if tok.startswith('"') or tok.endswith("'!"):
                return tok
            # Standalone '...' is ALWAYS invalid Excel syntax -> convert to "..."
            inner = tok[1:-1].replace("''", "'").replace('"', '""')
            return f'"{inner}"'

        fixed = _EXCEL_FORMULA_QUOTE_TOKEN_RE.sub(_repl, fixed)

    return fixed


def has_invalid_excel_formula_quotes(formula: str) -> bool:
    """Return True if an Excel formula contains illegal ``\\"`` or standalone ``'text'`` literals."""
    if not isinstance(formula, str) or not formula.startswith("="):
        return False
    return sanitize_excel_formula_quotes(formula) != formula


def atomic_save_via_temp(save_fn, target_path: Path) -> None:
    """Write via a sibling temp file and atomically swap into target_path via os.replace().

    Prevents truncating or corrupting an existing Office file if serialization fails
    mid-write or a cloud-sync client (Dropbox/OneDrive/iCloud) transiently locks the file.
    """
    import uuid as _uuid

    target = Path(target_path)
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = target.parent / f".{target.stem}.tmp.{_uuid.uuid4().hex[:8]}{target.suffix}"
    try:
        save_fn(str(tmp_path))
        os.replace(str(tmp_path), str(target))
    except Exception:
        try:
            if tmp_path.exists():
                tmp_path.unlink()
        except Exception:
            pass
        raise


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
