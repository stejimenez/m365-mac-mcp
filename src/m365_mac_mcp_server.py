#!/usr/bin/env python3
"""
Microsoft 365 for Mac MCP (m365-mac-mcp) — Hardened for macOS 27.x & Office 16.x,
Bionic 1.2.1+10, and Antigravity MCP clients).

Supports both:
- File-level read-write workflows (python-docx, openpyxl, csv, python-pptx)
- Live application control via hardened AppleScript (Microsoft Word & Excel)
"""

import functools
import inspect
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

# Ensure src/ is always importable regardless of launch working directory
SRC_DIR = str(Path(__file__).resolve().parent)
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

# Compatible with MCP Python SDK v2 (MCPServer) and v1 (FastMCP)
try:
    from mcp.server import MCPServer as _ServerClass
except ImportError:
    from mcp.server.fastmcp import FastMCP as _ServerClass  # type: ignore

from controllers.excel_controller import ExcelController
from controllers.powerpoint_controller import PowerPointController
from controllers.word_controller import WordController
from integrations.applescript_bridge import AppleScriptBridge
from utils.config import Config
from utils.logger import setup_logger
from utils.validators import (
    get_background_mode,
    get_blocked_write_dirs,
    get_client_project_dir,
    get_default_output_dir,
    is_write_restricted_to_project_dir,
    reset_current_session_id,
    sanitize_for_json,
    set_client_project_dir,
    set_current_session_id,
)

logger = setup_logger(__name__)
config = Config()

_VALID_PROFILE_APPS = {"excel", "word", "powerpoint"}
_AVAILABLE_PROFILES = [
    "full",
    "compact",
    "excel",
    "word",
    "powerpoint",
    "excel,word",
    "word,powerpoint",
    "excel,compact",
]


def _parse_mcp_profile(raw: str):
    """Parse M365_MCP_PROFILE (or legacy OFFICE_MCP_PROFILE) into (label, allowed_apps_or_None, compact_only)."""
    cleaned = (raw or "full").strip().lower()
    if not cleaned or cleaned == "full":
        return "full", None, False
    if cleaned == "compact":
        return "compact", None, True

    tokens = [t.strip() for t in cleaned.replace(";", ",").split(",") if t.strip()]
    apps: set[str] = set()
    compact_only = False
    unknown: List[str] = []

    alias_map = {
        "xls": "excel",
        "xlsx": "excel",
        "doc": "word",
        "docx": "word",
        "ppt": "powerpoint",
        "pptx": "powerpoint",
    }

    for tok in tokens:
        mapped = alias_map.get(tok, tok)
        if mapped in _VALID_PROFILE_APPS:
            apps.add(mapped)
        elif mapped == "compact":
            compact_only = True
        elif mapped == "full":
            return "full", None, False
        else:
            unknown.append(tok)

    if unknown or (not apps and not compact_only):
        logger.warning(
            f"Unknown M365_MCP_PROFILE='{raw}' (invalid tokens: {unknown}). "
            f"Falling back to 'full'. Valid values: {', '.join(_AVAILABLE_PROFILES)}."
        )
        return "full", None, False

    ordered = [a for a in ("excel", "word", "powerpoint") if a in apps]
    if compact_only:
        ordered.append("compact")
    return ",".join(ordered), (apps if apps else None), compact_only


def _resolve_profile_env() -> str:
    raw = os.environ.get("M365_MCP_PROFILE")
    if raw is not None and raw.strip():
        return raw
    legacy = os.environ.get("OFFICE_MCP_PROFILE")
    if legacy is not None and legacy.strip():
        logger.warning(
            "OFFICE_MCP_PROFILE is deprecated; rename it to M365_MCP_PROFILE in your MCP config."
        )
        return legacy
    return "full"


_ACTIVE_PROFILE_LABEL, _PROFILE_APPS, _PROFILE_COMPACT = _parse_mcp_profile(
    _resolve_profile_env()
)
_REGISTERED_TOOL_NAMES: List[str] = []

async def _sync_client_roots_from_context(ctx: Any) -> None:
    """Best-effort sync of MCP client workspace root (e.g. Bionic/VS Code/Cline Project Folder) via roots/list."""
    try:
        session = getattr(ctx, "session", None)
        if session is None:
            return
        caps = getattr(session, "client_capabilities", None)
        if caps is None or getattr(caps, "roots", None) is None:
            return
        import asyncio
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", DeprecationWarning)
            roots_result = await asyncio.wait_for(session.list_roots(), timeout=1.5)
        roots = getattr(roots_result, "roots", None) or []
        for root in roots:
            uri = str(getattr(root, "uri", "") or "")
            if uri.startswith("file://"):
                if set_client_project_dir(uri) is not None:
                    return
    except Exception:
        pass


class _M365MCPServer(_ServerClass):  # type: ignore[misc,valid-type]
    """MCPServer subclass so MCP SDK v2 binds our session-isolation & roots/list hook during __init__."""

    @staticmethod
    def _extract_session_key(ctx: Any) -> Optional[str]:
        """Return a stable per-connection session key across MCP SDK v1 and v2.

        In MCP SDK v2, ``ctx.session`` is a per-request ``ServerSession`` wrapper around
        a persistent ``_connection`` object; keying on ``id(_connection)`` preserves
        ``working_dir`` / ``roots/list`` across tool calls on the same connection while
        still isolating concurrent client connections.
        """
        session = getattr(ctx, "session", None)
        if session is None:
            return None
        conn = getattr(session, "_connection", None)
        target = conn if conn is not None else session
        return f"mcp_session_{id(target)}"

    async def _handle_call_tool(self, ctx: Any, params: Any) -> Any:
        sess_key = self._extract_session_key(ctx)
        tok = set_current_session_id(sess_key) if sess_key is not None else None
        try:
            tool_name = getattr(params, "name", "")
            tool_args = getattr(params, "arguments", None) or {}
            if (tool_name == "check_office_status" and not tool_args.get("working_dir")) or get_client_project_dir() is None:
                await _sync_client_roots_from_context(ctx)
            return await super()._handle_call_tool(ctx, params)
        finally:
            if tok is not None:
                reset_current_session_id(tok)


mcp = _M365MCPServer(
    "m365-mac-mcp",
    instructions=(
        "m365-mac-mcp (Excel, Word, PowerPoint on macOS). "
        "Prefer batch file tools for edits/formatting and *_live_* tools for live Excel recalc."
    ),
)

applescript = AppleScriptBridge(default_timeout=config.get("applescript_timeout", 30.0))
word = WordController(max_documents=config.get("max_documents", 20))
excel = ExcelController(max_workbooks=config.get("max_workbooks", 20))
powerpoint = PowerPointController()


def _compact_json_schema(node: Any, is_properties_map: bool = False) -> None:
    """Strip Pydantic schema noise ('title', verbose nullable 'anyOf', 'default': null) in place.

    Reduces per-turn MCP tool schema tokens by ~30%+ without changing parameter names,
    types, non-null defaults, or Pydantic's runtime validation of Optional[T] arguments.
    """
    if isinstance(node, dict):
        if not is_properties_map:
            node.pop("title", None)
            if (
                "anyOf" in node
                and isinstance(node["anyOf"], list)
                and len(node["anyOf"]) == 2
            ):
                non_null = [item for item in node["anyOf"] if item != {"type": "null"}]
                if len(non_null) == 1 and isinstance(non_null[0], dict):
                    merged = dict(non_null[0])
                    for k, v in node.items():
                        if k != "anyOf":
                            merged[k] = v
                    node.clear()
                    node.update(merged)
            if "default" in node and node["default"] is None:
                node.pop("default", None)
            if node.get("additionalProperties") is True:
                node.pop("additionalProperties", None)
        for k, v in node.items():
            _compact_json_schema(v, is_properties_map=(k == "properties"))
    elif isinstance(node, list):
        for item in node:
            _compact_json_schema(item, is_properties_map=False)


def _append_debug_log(tool_name: str, args: tuple, kwargs: Dict[str, Any], result: Any, elapsed_ms: float) -> None:
    """Opt-in JSONL wire logger (active only when M365_MCP_DEBUG_LOG is set)."""
    log_path_raw = os.environ.get("M365_MCP_DEBUG_LOG", "").strip()
    if not log_path_raw:
        return
    try:
        log_path = Path(os.path.expanduser(log_path_raw))
        log_path.parent.mkdir(parents=True, exist_ok=True)
        arg_summary: Dict[str, Any] = {}
        if args:
            arg_summary["_args"] = sanitize_for_json(list(args))
        if kwargs:
            arg_summary.update(sanitize_for_json(kwargs))
        args_str = json.dumps(arg_summary, separators=(",", ":"), ensure_ascii=False, default=str)
        if len(args_str) > 1500:
            args_str = args_str[:1500] + "...<truncated>"

        record: Dict[str, Any] = {
            "ts": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ"),
            "tool": tool_name,
            "ms": round(elapsed_ms, 1),
            "args": args_str,
        }
        if isinstance(result, dict):
            if "status" in result:
                record["status"] = result["status"]
            if "error" in result:
                record["error"] = str(result["error"])[:500]
            if "warning" in result:
                record["warning"] = str(result["warning"])[:500]
            if "warnings" in result:
                record["warnings"] = result["warnings"]
            if "blocked_count" in result:
                record["blocked_count"] = result["blocked_count"]
        res_str = (
            result
            if isinstance(result, str)
            else json.dumps(sanitize_for_json(result), separators=(",", ":"), ensure_ascii=False, default=str)
        )
        record["response_preview"] = res_str[:1200] + ("...<truncated>" if len(res_str) > 1200 else "")
        with log_path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(record, separators=(",", ":"), ensure_ascii=False) + "\n")
    except Exception:
        pass


def _wrap_compact_json(fn):
    """Wrap an async tool function so MCP returns compact JSON (no indent=2 whitespace)."""
    @functools.wraps(fn)
    async def wrapped(*args, **kwargs):
        t0 = time.perf_counter()
        res = await fn(*args, **kwargs)
        _append_debug_log(fn.__name__, args, kwargs, res, (time.perf_counter() - t0) * 1000.0)
        if isinstance(res, str):
            return res
        return json.dumps(
            sanitize_for_json(res),
            separators=(",", ":"),
            ensure_ascii=False,
            default=str,
        )

    wrapped.__signature__ = inspect.signature(fn).replace(return_annotation=str)
    wrapped.__annotations__ = {**getattr(fn, "__annotations__", {}), "return": str}
    return wrapped


def _tool(app: str = "core", compact: bool = True):
    """Register an MCP tool respecting M365_MCP_PROFILE (full, compact, excel, word, powerpoint, combos)."""
    def decorator(fn):
        app_allowed = (_PROFILE_APPS is None) or (app == "core") or (app in _PROFILE_APPS)
        compact_allowed = (not _PROFILE_COMPACT) or compact
        if app_allowed and compact_allowed:
            _REGISTERED_TOOL_NAMES.append(fn.__name__)
            wrapped = _wrap_compact_json(fn)
            try:
                mcp.tool(structured_output=False)(wrapped)
            except TypeError:
                # Fallback for MCP SDK v1 (FastMCP) which does not accept structured_output
                mcp.tool()(wrapped)
            tool_mgr = getattr(mcp, "_tool_manager", None)
            tools_map = getattr(tool_mgr, "_tools", None)
            if isinstance(tools_map, dict) and fn.__name__ in tools_map:
                _compact_json_schema(tools_map[fn.__name__].parameters)
        # Return original fn so direct Python callers/tests continue receiving structured dicts
        return fn
    return decorator


def _error_response(action: str, exc: Exception) -> Dict[str, Any]:
    """Return a concise, structured, actionable error payload so local models recover cleanly."""
    logger.error(f"{action} failed: {exc}")
    msg = str(exc)
    hint: Optional[str] = None

    awaiting_user_action = False
    registered = set(_REGISTERED_TOOL_NAMES)
    if isinstance(exc, PermissionError) and "mcp server project folder" in msg.lower():
        hint = "Provide an explicit user output path outside the MCP server project folder (e.g., '~/Desktop/Model.xlsx' or '~/Documents/Report.docx')."
    elif isinstance(exc, PermissionError) and (
        "m365_mcp_blocked_dirs" in msg.lower()
        or "m365_mcp_restrict_write_to_project_dir" in msg.lower()
    ):
        awaiting_user_action = True
        hint = (
            "HARD STOP IMMEDIATELY: write destination is blocked by server write-fence configuration "
            "(M365_MCP_BLOCKED_DIRS or M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR). "
            "Do NOT retry or silently save elsewhere; ask the user for an allowed output folder."
        )
    elif "-1743" in msg or "not authorized to send apple events" in msg.lower():
        awaiting_user_action = True
        hint = (
            "HARD STOP IMMEDIATELY: macOS denied Automation permission (-1743). "
            "Do NOT fall back to uncached file-level reads for formulas or run downstream steps. "
            "Ask the user to enable permission in System Settings > Privacy & Security > Automation for Microsoft Office and reply 'continue'."
        )
    elif isinstance(exc, PermissionError):
        awaiting_user_action = True
        hint = (
            "HARD STOP IMMEDIATELY: macOS denied filesystem access to this file or folder (e.g., Desktop/Documents "
            "permission revoked in System Settings > Privacy & Security > Files & Folders, or file is read-only). "
            "Do NOT retry or save alternate copies silently; tell the user which path was blocked and ask for an allowed folder."
        )
    elif isinstance(exc, TimeoutError):
        awaiting_user_action = True
        hint = (
            "Microsoft Office is likely waiting on a macOS 'Grant Access', Password, Enable Macros, or Repair modal dialog. "
            "HARD STOP IMMEDIATELY: do NOT assume the operation succeeded, do NOT retry, do NOT save alternate copies, "
            "do NOT fall back to file-level read_range for uncached formulas, do NOT calculate mentally, and do NOT run downstream steps. "
            "Tell the user which file/action is waiting on Office and ask them to approve/dismiss the dialog and reply 'continue' when back."
        )
    elif "unsaved manual changes" in msg.lower():
        awaiting_user_action = True
        hint = (
            "HARD STOP IMMEDIATELY: the user has unsaved manual edits open in Microsoft Excel. "
            "Do NOT call excel_live_save, do NOT call excel_live_close(save_changes='no'), and do NOT save alternate copies elsewhere. "
            "Ask the user to save (Cmd+S) or discard their changes in Excel and reply 'continue'."
        )
    elif action == "excel_live_open" and ("-50" in msg or "parameter error" in msg.lower()):
        hint = (
            "Microsoft Excel rejected the saved .xlsx file (Parameter error -50), typically due to invalid formula syntax "
            "(e.g., 1-argument NPV()), a circular self-reference, or malformed OOXML. Call audit_workbook_formulas(workbook_id=...), "
            "fix any flagged formulas/inconsistencies, call save_workbook, then retry excel_live_open."
        )
    elif isinstance(exc, FileNotFoundError):
        hint = "Verify the file path exists (supports ~ expansion and spaces) before retrying."
    elif isinstance(exc, ValueError):
        low = msg.lower()
        if "workbook" in low and "not found" in low:
            if "list_active_workbooks" in registered:
                hint = "Call list_active_workbooks or create_workbook/open_workbook to obtain a valid workbook_id."
            else:
                hint = "Call create_workbook or open_workbook to obtain a valid workbook_id (or omit workbook_id when 1 workbook is active)."
        elif "document" in low and "not found" in low:
            if "list_active_documents" in registered:
                hint = "Call list_active_documents or create_document/open_document to obtain a valid document_id (or omit document_id when 1 document is active)."
            else:
                hint = "Call create_document or open_document to obtain a valid document_id (or omit document_id when 1 document is active)."
        elif "presentation" in low and "not found" in low:
            if "list_active_presentations" in registered:
                hint = "Call list_active_presentations or create_presentation to obtain a valid presentation_id (or omit presentation_id when 1 presentation is active)."
            else:
                hint = "Call create_presentation to obtain a valid presentation_id (or omit presentation_id when 1 presentation is active)."
        elif "worksheet" in low and "not found" in low:
            hint = "Call list_worksheets to verify exact sheet names (including spaces/capitalization)."
        elif "range" in low or "cell reference" in low:
            hint = "Use A1-style references like 'A1' or 'A1:G10'."
    elif isinstance(exc, IndexError):
        if "read_document" in registered:
            hint = "Use read_document or read_tables to check valid 0-based indices before updating."
        else:
            hint = "Verify valid 0-based index bounds before retrying."

    payload: Dict[str, Any] = {
        "status": "error",
        "action": action,
        "error_type": type(exc).__name__,
        "error": msg,
    }
    if awaiting_user_action:
        payload["awaiting_user_action"] = True
    if hint:
        payload["hint"] = hint
    return sanitize_for_json(payload)


# ======================================================================
# WORD FILE-LEVEL READ-WRITE TOOLS (python-docx)
# ======================================================================

@_tool(app="word")
async def create_document(
    title: str = "New Document",
    template_path: Optional[str] = None,
    add_title_heading: bool = False,
    open_in_word: bool = False,
) -> Dict[str, Any]:
    """Create a new Word document (.docx) for file-level editing. Parent directories are auto-created on save_document."""
    try:
        return await word.create_document(
            title=title,
            template_path=template_path,
            add_title_heading=add_title_heading,
            open_in_word=open_in_word,
        )
    except Exception as exc:
        return _error_response("create_document", exc)


@_tool(app="word")
async def open_document(file_path: str) -> Dict[str, Any]:
    """Open an existing Word document (.docx, .dotx, or .doc) for reading and editing."""
    try:
        return await word.open_document(file_path=file_path)
    except Exception as exc:
        return _error_response("open_document", exc)


@_tool(app="word")
async def read_document(
    document_id: Optional[str] = None,
    include_tables: bool = True,
    include_styles: bool = False,
    include_metadata: bool = True,
    include_full_text: bool = False,
    max_paragraphs: int = 200,
    start_paragraph: int = 0,
) -> Dict[str, Any]:
    """Read indexed paragraphs (with non-Normal styles), tables, and metadata from a Word document (paginated via max_paragraphs/start_paragraph). document_id auto-resolves if 1 document is active."""
    try:
        return await word.read_document(
            document_id=document_id,
            include_tables=include_tables,
            include_styles=include_styles,
            include_metadata=include_metadata,
            include_full_text=include_full_text,
            max_paragraphs=max_paragraphs,
            start_paragraph=start_paragraph,
        )
    except Exception as exc:
        return _error_response("read_document", exc)


@_tool(app="word")
async def add_heading(
    document_id: Optional[str] = None,
    text: str = "",
    level: int = 1,
    style: Optional[str] = None,
    formatting: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Append a heading (level 0-9) to a Word document with optional explicit formatting. document_id auto-resolves if 1 document is active."""
    try:
        return await word.add_heading(
            document_id=document_id,
            text=text,
            level=level,
            style=style,
            formatting=formatting,
        )
    except Exception as exc:
        return _error_response("add_heading", exc)


@_tool(app="word", compact=False)
async def format_paragraph(
    document_id: Optional[str] = None,
    index: int = 0,
    formatting: Optional[Dict[str, Any]] = None,
    style: Optional[str] = None,
) -> Dict[str, Any]:
    """Apply explicit style and formatting to an existing Word paragraph by 0-based index."""
    try:
        return await word.format_paragraph(
            document_id=document_id, index=index, formatting=formatting or {}, style=style
        )
    except Exception as exc:
        return _error_response("format_paragraph", exc)


@_tool(app="word")
async def add_paragraph(
    document_id: Optional[str] = None,
    text: str = "",
    style: Optional[str] = None,
    formatting: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Append a paragraph to a Word document with optional style and formatting. document_id auto-resolves if 1 document is active."""
    try:
        return await word.add_paragraph(
            document_id=document_id,
            text=text,
            style=style,
            formatting=formatting,
        )
    except Exception as exc:
        return _error_response("add_paragraph", exc)


@_tool(app="word", compact=False)
async def insert_paragraph(
    document_id: Optional[str] = None,
    index: int = 0,
    text: str = "",
    style: Optional[str] = None,
    formatting: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Insert a paragraph before the specified 0-based paragraph index."""
    try:
        return await word.insert_paragraph(
            document_id=document_id,
            index=index,
            text=text,
            style=style,
            formatting=formatting,
        )
    except Exception as exc:
        return _error_response("insert_paragraph", exc)


@_tool(app="word", compact=False)
async def replace_document_text(
    document_id: Optional[str] = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = True,
    include_tables: bool = True,
) -> Dict[str, Any]:
    """Find and replace text across paragraphs and tables in a Word document."""
    try:
        return await word.replace_text(
            document_id=document_id,
            find_text=find_text,
            replace_text=replace_text,
            match_case=match_case,
            include_tables=include_tables,
        )
    except Exception as exc:
        return _error_response("replace_document_text", exc)


@_tool(app="word", compact=False)
async def delete_paragraph(document_id: Optional[str] = None, index: int = 0) -> Dict[str, Any]:
    """Delete a paragraph by its 0-based index from a Word document."""
    try:
        return await word.delete_paragraph(document_id=document_id, index=index)
    except Exception as exc:
        return _error_response("delete_paragraph", exc)


@_tool(app="word", compact=False)
async def add_list(
    document_id: Optional[str] = None,
    items: Optional[List[str]] = None,
    list_type: str = "bullet",
    style: Optional[str] = None,
    formatting: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Append a bulleted or numbered list to a Word document with optional formatting."""
    try:
        return await word.add_list(
            document_id=document_id,
            items=items or [],
            list_type=list_type,
            style=style,
            formatting=formatting,
        )
    except Exception as exc:
        return _error_response("add_list", exc)


@_tool(app="word")
async def add_table(
    document_id: Optional[str] = None,
    rows: int = 0,
    columns: int = 0,
    data: Optional[List[List[Any]]] = None,
    style: Optional[str] = None,
    column_widths: Optional[List[float]] = None,
    header_formatting: Optional[Dict[str, Any]] = None,
    body_formatting: Optional[Dict[str, Any]] = None,
    alternating_row_color: Optional[str] = None,
    borders: Optional[Dict[str, Any]] = None,
    cell_padding: Optional[Dict[str, float]] = None,
    repeat_header: bool = False,
    table_alignment: Optional[str] = None,
    column_alignments: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """WORD DOCUMENTS ONLY (not PowerPoint — use add_text_to_slide or batch_add_slides for slides). Add a styled table to a Word document and populate cells in one call. document_id auto-resolves if 1 Word document is active."""
    try:
        if document_id and (
            document_id in powerpoint.active_presentations
            or any(document_id in p.get("slides", {}) for p in powerpoint.active_presentations.values())
        ):
            raise ValueError(
                "add_table is for Microsoft Word documents only — not PowerPoint presentations or slides. "
                "Use add_text_to_slide or batch_add_slides for PowerPoint content."
            )
        return await word.add_table(
            document_id=document_id,
            rows=rows,
            columns=columns,
            data=data,
            style=style,
            column_widths=column_widths,
            header_formatting=header_formatting,
            body_formatting=body_formatting,
            alternating_row_color=alternating_row_color,
            borders=borders,
            cell_padding=cell_padding,
            repeat_header=repeat_header,
            table_alignment=table_alignment,
            column_alignments=column_alignments,
        )
    except Exception as exc:
        return _error_response("add_table", exc)


@_tool(app="word", compact=False)
async def format_table(
    document_id: Optional[str] = None,
    table_index: int = 0,
    style: Optional[str] = None,
    column_widths: Optional[List[float]] = None,
    header_formatting: Optional[Dict[str, Any]] = None,
    body_formatting: Optional[Dict[str, Any]] = None,
    alternating_row_color: Optional[str] = None,
    borders: Optional[Dict[str, Any]] = None,
    cell_padding: Optional[Dict[str, float]] = None,
    repeat_header: bool = False,
    table_alignment: Optional[str] = None,
    column_alignments: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """Style an existing Word table (header shading, borders, padding, column widths, fonts)."""
    try:
        return await word.format_table(
            document_id=document_id,
            table_index=table_index,
            style=style,
            column_widths=column_widths,
            header_formatting=header_formatting,
            body_formatting=body_formatting,
            alternating_row_color=alternating_row_color,
            borders=borders,
            cell_padding=cell_padding,
            repeat_header=repeat_header,
            table_alignment=table_alignment,
            column_alignments=column_alignments,
        )
    except Exception as exc:
        return _error_response("format_table", exc)


@_tool(app="word", compact=False)
async def update_table_cell(
    document_id: Optional[str] = None,
    table_index: int = 0,
    row: int = 0,
    column: int = 0,
    text: str = "",
    formatting: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Update a single cell (and optional formatting) in an existing Word table (0-based indices)."""
    try:
        return await word.update_table_cell(
            document_id=document_id,
            table_index=table_index,
            row=row,
            column=column,
            text=text,
            formatting=formatting,
        )
    except Exception as exc:
        return _error_response("update_table_cell", exc)


@_tool(app="word", compact=False)
async def set_document_properties(
    document_id: Optional[str] = None,
    title: Optional[str] = None,
    author: Optional[str] = None,
    subject: Optional[str] = None,
    comments: Optional[str] = None,
) -> Dict[str, Any]:
    """Set core document metadata properties (title, author, subject, comments)."""
    try:
        return await word.set_document_properties(
            document_id=document_id,
            title=title,
            author=author,
            subject=subject,
            comments=comments,
        )
    except Exception as exc:
        return _error_response("set_document_properties", exc)


@_tool(app="word")
async def save_document(
    document_id: Optional[str] = None,
    file_path: Optional[str] = None,
    format: str = "docx",
) -> Dict[str, Any]:
    """Save a Word document in-place or to a new path (.docx, .pdf, .rtf, .txt, .doc). document_id auto-resolves if 1 document is active."""
    try:
        return await word.save_document(
            document_id=document_id, file_path=file_path, format=format
        )
    except Exception as exc:
        return _error_response("save_document", exc)


@_tool(app="word")
async def close_document(
    document_id: Optional[str] = None,
    save_before_close: bool = False,
) -> Dict[str, Any]:
    """Close an in-memory Word document and clean up temporary files. document_id auto-resolves if 1 document is active."""
    try:
        return await word.close_document(
            document_id=document_id, save_before_close=save_before_close
        )
    except Exception as exc:
        return _error_response("close_document", exc)


# ======================================================================
# WORD LIVE-APP TOOLS (Microsoft Word 16.115.1 via AppleScript)
# ======================================================================

@_tool(app="word")
async def word_live_list_documents() -> Dict[str, Any]:
    """List all documents currently open in the live Microsoft Word app."""
    try:
        docs = await applescript.list_open_word_documents()
        return {"status": "success", "count": len(docs), "documents": docs}
    except Exception as exc:
        return _error_response("word_live_list_documents", exc)


def _reject_internal_temp_live_open(file_path: str, app_name: str, save_tool: str) -> None:
    """Block live Office open on internal /tmp or macOS /var/folders temp UUID files that trigger Grant Access dialogs."""
    try:
        target = Path(os.path.expanduser(file_path)).resolve()
        temp_roots = [
            excel.temp_dir.resolve(),
            word.temp_dir.resolve(),
            Path(config.get("temp_dir", "/tmp/m365-mac-mcp")).resolve(),
            Path("/tmp/m365-mac-mcp").resolve(),
            Path("/tmp/office365_mcp").resolve(),
        ]
        for root in temp_roots:
            if target == root or root in target.parents:
                raise ValueError(
                    f"Cannot open internal temporary file '{file_path}' in live {app_name} "
                    f"(macOS sandbox triggers a Grant Access dialog). Call {save_tool}(file_path=...) "
                    f"to save to a normal folder first, then open the saved file."
                )
    except ValueError:
        raise
    except Exception:
        pass


@_tool(app="word")
async def word_live_open(
    file_path: str,
    read_only: bool = False,
    activate: bool = False,
) -> Dict[str, Any]:
    """Open a Word file directly in the live Microsoft Word application."""
    try:
        _reject_internal_temp_live_open(file_path, "Microsoft Word", "save_document")
        return await applescript.open_word_file(
            file_path=file_path, read_only=read_only, activate=activate
        )
    except Exception as exc:
        return _error_response("word_live_open", exc)


@_tool(app="word", compact=False)
async def word_live_create(
    title: Optional[str] = None,
    activate: bool = False,
) -> Dict[str, Any]:
    """Create a new document directly in the live Microsoft Word application."""
    try:
        return await applescript.create_word_document(title=title, activate=activate)
    except Exception as exc:
        return _error_response("word_live_create", exc)


@_tool(app="word")
async def word_live_read(
    document_name: Optional[str] = None,
    include_tables: bool = True,
    include_full_text: bool = False,
    max_paragraphs: int = 200,
    start_paragraph: int = 0,
) -> Dict[str, Any]:
    """Read paragraphs, selection/cursor state, and tables from a live Word document (paginated via max_paragraphs/start_paragraph; pass include_full_text=True only if raw joined text is needed)."""
    try:
        return await applescript.read_live_word_document(
            document_name=document_name,
            include_tables=include_tables,
            include_full_text=include_full_text,
            max_paragraphs=max_paragraphs,
            start_paragraph=start_paragraph,
        )
    except Exception as exc:
        return _error_response("word_live_read", exc)


@_tool(app="word", compact=False)
async def word_live_write(
    text: str,
    document_name: Optional[str] = None,
    mode: str = "append",
    add_newline: bool = True,
) -> Dict[str, Any]:
    """Write text into a live Microsoft Word document."""
    try:
        return await applescript.write_live_word_text(
            text=text,
            document_name=document_name,
            mode=mode,
            add_newline=add_newline,
        )
    except Exception as exc:
        return _error_response("word_live_write", exc)


@_tool(app="word", compact=False)
async def word_live_replace_text(
    find_text: str,
    replace_text: str,
    document_name: Optional[str] = None,
    match_case: bool = False,
) -> Dict[str, Any]:
    """Find and replace all occurrences in a live Microsoft Word document."""
    try:
        return await applescript.replace_live_word_text(
            find_text=find_text,
            replace_text=replace_text,
            document_name=document_name,
            match_case=match_case,
        )
    except Exception as exc:
        return _error_response("word_live_replace_text", exc)


def _find_dirty_in_memory_document(
    document_name: Optional[str] = None,
    file_path: Optional[str] = None,
) -> Optional[str]:
    """Return the document_id of a matching active in-memory Word document with unsaved edits."""
    targets: set[str] = set()
    if document_name and str(document_name).strip():
        targets.add(Path(str(document_name).strip()).name.lower())
    if file_path and str(file_path).strip():
        targets.add(Path(str(file_path).strip()).expanduser().name.lower())

    dirty_ids: List[str] = []
    for doc_id, entry in word.active_documents.items():
        if not entry.get("dirty"):
            continue
        dirty_ids.append(doc_id)
        if targets:
            meta = entry.get("metadata") or {}
            candidates = {
                Path(str(entry.get("working_path") or "")).name.lower(),
                Path(str(meta.get("file_path") or "")).name.lower(),
                Path(str(meta.get("source_path") or "")).name.lower(),
            }
            if targets & candidates:
                return doc_id
    if not targets and len(dirty_ids) == 1:
        return dirty_ids[0]
    return None


@_tool(app="word")
async def word_live_save(
    document_name: Optional[str] = None,
    file_path: Optional[str] = None,
    file_format: str = "docx",
) -> Dict[str, Any]:
    """Save a live Word document in-place or Save As (.docx, .pdf, .rtf, .txt, .doc)."""
    try:
        dirty_doc_id = _find_dirty_in_memory_document(document_name=document_name, file_path=file_path)
        if dirty_doc_id:
            return {
                "status": "error",
                "action": "word_live_save",
                "error_type": "UnsavedInMemoryEditsError",
                "error": (
                    f"In-memory Word document '{dirty_doc_id}' has unsaved file-level edits "
                    "(from add_heading/add_paragraph/add_table/replace_document_text). "
                    "Calling word_live_save would save Word's stale copy over disk and lose those edits."
                ),
                "hint": f"Call save_document(document_id='{dirty_doc_id}', file_path=...) instead to persist your file-level edits first.",
            }
        return await applescript.save_word_document(
            document_name=document_name,
            file_path=file_path,
            file_format=file_format,
        )
    except Exception as exc:
        return _error_response("word_live_save", exc)


def _normalize_save_changes(save_changes: Union[bool, str]) -> str:
    """Accept bool (True/False) or str ('yes'/'no'/'ask') for live close tools."""
    if isinstance(save_changes, bool):
        return "yes" if save_changes else "no"
    return str(save_changes or "no").strip().lower()


@_tool(app="word")
async def word_live_close(
    document_name: Optional[str] = None,
    save_changes: Union[bool, str] = "no",
) -> Dict[str, Any]:
    """Close a live Word document cleanly without triggering unsaved-changes dialogs."""
    try:
        return await applescript.close_word_document(
            document_name=document_name,
            save_changes=_normalize_save_changes(save_changes),
        )
    except Exception as exc:
        return _error_response("word_live_close", exc)


# ======================================================================
# EXCEL FILE-LEVEL READ-WRITE TOOLS (openpyxl + csv)
# ======================================================================

@_tool(app="excel")
async def create_workbook(
    title: str = "New Workbook",
    template_path: Optional[str] = None,
    sheet_name: str = "Sheet1",
    sheet_names: Optional[List[str]] = None,
    include_title_header: bool = False,
    open_in_excel: bool = False,
    header_metadata: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Create a new Excel workbook (.xlsx). Pass sheet_names=['IS','BS','CF'] to create multiple tabs in one call. Optional header_metadata={'entity':'Company/Project','period':'Years 0-6','currency':'USD','magnitude':'auto'} auto-writes a 4-line institutional header block (Rows 1-4; tables start at Row 6) with Smart Magnitude scaling ($/$1,000's/$M/$B)."""
    try:
        return await excel.create_workbook(
            title=title,
            template_path=template_path,
            sheet_name=sheet_name,
            sheet_names=sheet_names,
            include_title_header=include_title_header,
            open_in_excel=open_in_excel,
            header_metadata=header_metadata,
        )
    except Exception as exc:
        return _error_response("create_workbook", exc)


@_tool(app="excel")
async def open_workbook(file_path: str) -> Dict[str, Any]:
    """Open an existing Excel (.xlsx, .xlsm with macro preservation) or CSV (.csv) file."""
    try:
        return await excel.open_workbook(file_path=file_path)
    except Exception as exc:
        return _error_response("open_workbook", exc)


@_tool(app="excel")
async def list_worksheets(workbook_id: Optional[str] = None) -> Dict[str, Any]:
    """List all worksheet names in an open workbook. workbook_id auto-resolves if 1 workbook is active."""
    try:
        resolved_id = excel._resolve_workbook_id(workbook_id)
        sheets = await excel.list_worksheets(workbook_id=resolved_id)
        return {"status": "success", "workbook_id": resolved_id, "worksheets": sheets}
    except Exception as exc:
        return _error_response("list_worksheets", exc)


@_tool(app="excel")
async def add_worksheet(
    workbook_id: Optional[str] = None,
    sheet_name: str = "",
    position: Optional[int] = None,
) -> Dict[str, Any]:
    """Add a new worksheet to an open workbook. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.add_worksheet(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            position=position,
        )
    except Exception as exc:
        return _error_response("add_worksheet", exc)


@_tool(app="excel", compact=False)
async def rename_worksheet(
    workbook_id: Optional[str] = None,
    old_name: str = "",
    new_name: str = "",
) -> Dict[str, Any]:
    """Rename an existing worksheet in an open workbook. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.rename_worksheet(
            workbook_id=excel._resolve_workbook_id(workbook_id), old_name=old_name, new_name=new_name
        )
    except Exception as exc:
        return _error_response("rename_worksheet", exc)


@_tool(app="excel")
async def delete_worksheet(
    workbook_id: Optional[str] = None,
    sheet_name: str = "",
) -> Dict[str, Any]:
    """Delete a worksheet from an open workbook. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.delete_worksheet(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
        )
    except Exception as exc:
        return _error_response("delete_worksheet", exc)


@_tool(app="excel", compact=False)
async def set_active_worksheet(
    workbook_id: Optional[str] = None,
    sheet_name: str = "",
) -> Dict[str, Any]:
    """Set the active worksheet of an open workbook. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.set_active_worksheet(
            workbook_id=excel._resolve_workbook_id(workbook_id), sheet_name=sheet_name
        )
    except Exception as exc:
        return _error_response("set_active_worksheet", exc)


@_tool(app="excel", compact=False)
async def read_cell(
    workbook_id: Optional[str] = None,
    cell: str = "",
    sheet_name: Optional[str] = None,
    include_evaluated: bool = True,
) -> Dict[str, Any]:
    """Read a single cell's value, formula, cached evaluated value, and number format. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.read_cell(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            cell=cell,
            include_evaluated=include_evaluated,
        )
    except Exception as exc:
        return _error_response("read_cell", exc)


@_tool(app="excel")
async def read_range(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    cell_range: Optional[str] = None,
    range_ref: Optional[str] = None,
    range: Optional[str] = None,
    evaluate_formulas: bool = False,
    include_formulas: Optional[bool] = None,
    max_rows: int = 150,
    max_cols: int = 30,
) -> Dict[str, Any]:
    """Read a cell range (or used range) as a JSON-safe 2D matrix (default cap 150x30 to protect context). Accepts cell_range, range_ref, or range; workbook_id auto-resolves if 1 workbook is active."""
    try:
        effective_range = (
            cell_range
            if cell_range is not None
            else (range_ref if range_ref is not None else range)
        )
        effective_eval = (not include_formulas) if include_formulas is not None else evaluate_formulas
        return await excel.read_range(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            cell_range=effective_range,
            evaluate_formulas=effective_eval,
            max_rows=max_rows,
            max_cols=max_cols,
        )
    except Exception as exc:
        return _error_response("read_range", exc)


@_tool(app="excel", compact=False)
async def read_tables(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    max_rows_per_table: int = 100,
) -> Dict[str, Any]:
    """Read structured Excel Tables (names, ranges, headers, up to max_rows_per_table rows) from a sheet or workbook. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.read_tables(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            max_rows_per_table=max_rows_per_table,
        )
    except Exception as exc:
        return _error_response("read_tables", exc)


@_tool(app="excel")
async def inspect_workbook(
    workbook_id: Optional[str] = None,
    file_path: Optional[str] = None,
    sheet_names: Optional[List[str]] = None,
    max_rows_per_sheet: int = 50,
    max_cols_per_sheet: int = 20,
    include_formulas: bool = True,
    include_evaluated: bool = True,
) -> Dict[str, Any]:
    """Inspect all (or selected via sheet_names) sheets (headers, sparse formulas, cached values, tables) in ONE call."""
    try:
        return await excel.inspect_workbook(
            workbook_id=workbook_id,
            file_path=file_path,
            sheet_names=sheet_names,
            max_rows_per_sheet=max_rows_per_sheet,
            max_cols_per_sheet=max_cols_per_sheet,
            include_formulas=include_formulas,
            include_evaluated=include_evaluated,
        )
    except Exception as exc:
        return _error_response("inspect_workbook", exc)


@_tool(app="excel")
async def audit_workbook_formulas(
    workbook_id: Optional[str] = None,
    file_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Audit all sheets (read-only) for Excel errors (#REF!/#NAME?/#VALUE!), broken cross-sheet links, external links, and suspicious literals. Pass workbook_id or file_path."""
    try:
        return await excel.audit_workbook_formulas(
            workbook_id=workbook_id,
            file_path=file_path,
        )
    except Exception as exc:
        return _error_response("audit_workbook_formulas", exc)


@_tool(app="excel")
async def write_cell(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    cell: str = "",
    value: Any = None,
    formatting: Optional[Dict[str, Any]] = None,
    number_format: Optional[str] = None,
    allow_formula_overwrite: bool = False,
) -> Dict[str, Any]:
    """Write a properly typed value (bool, int, float, str, None) to a cell. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.write_cell(
            workbook_id=workbook_id,
            sheet_name=sheet_name,
            cell=cell,
            value=value,
            formatting=formatting,
            number_format=number_format,
            allow_formula_overwrite=allow_formula_overwrite,
        )
    except Exception as exc:
        return _error_response("write_cell", exc)


@_tool(app="excel", compact=False)
async def write_range(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    start_cell: str = "A1",
    data: Optional[List[List[Any]]] = None,
    formatting: Optional[Dict[str, Any]] = None,
    header_formatting: Optional[Dict[str, Any]] = None,
    number_format: Optional[str] = None,
    allow_formula_overwrite: bool = False,
) -> Dict[str, Any]:
    """Write a 2D array/matrix of typed values starting at start_cell. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.write_range(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            start_cell=start_cell,
            data=data or [],
            formatting=formatting,
            header_formatting=header_formatting,
            number_format=number_format,
            allow_formula_overwrite=allow_formula_overwrite,
        )
    except Exception as exc:
        return _error_response("write_range", exc)


@_tool(app="excel", compact=False)
async def add_formula(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    cell: str = "",
    formula: str = "",
    number_format: Optional[str] = None,
    formatting: Optional[Dict[str, Any]] = None,
    allow_formula_overwrite: bool = False,
) -> Dict[str, Any]:
    """Write an Excel formula to a cell (e.g. '=SUM(B2:E2)' or "='Balance Sheet'!B2"). workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.add_formula(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            cell=cell,
            formula=formula,
            number_format=number_format,
            formatting=formatting,
            allow_formula_overwrite=allow_formula_overwrite,
        )
    except Exception as exc:
        return _error_response("add_formula", exc)


@_tool(app="excel", compact=False)
async def set_column_widths(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    widths: Optional[Dict[str, float]] = None,
    auto_fit: bool = False,
    min_width: float = 10.0,
    max_width: float = 60.0,
    padding: float = 4.0,
) -> Dict[str, Any]:
    """Set explicit Excel column widths (e.g. {'A': 42, 'B:G': 15}) and/or auto-fit columns. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.set_column_widths(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            widths=widths,
            auto_fit=auto_fit,
            min_width=min_width,
            max_width=max_width,
            padding=padding,
        )
    except Exception as exc:
        return _error_response("set_column_widths", exc)


@_tool(app="excel", compact=False)
async def format_range(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    cell_range: Optional[str] = None,
    range_ref: Optional[str] = None,
    formatting: Optional[Dict[str, Any]] = None,
    number_format: Optional[str] = None,
) -> Dict[str, Any]:
    """Style an Excel cell/range ('A1:G1', '1:1', 'B:G'). workbook_id auto-resolves if 1 workbook is active."""
    try:
        effective_range = cell_range if cell_range is not None else range_ref
        if not effective_range:
            raise ValueError("Provide cell_range or range_ref (e.g. 'A1:G1')")
        return await excel.format_range(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            cell_range=effective_range,
            formatting=formatting,
            number_format=number_format,
        )
    except Exception as exc:
        return _error_response("format_range", exc)


@_tool(app="excel", compact=False)
async def format_worksheet(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    column_widths: Optional[Dict[str, float]] = None,
    row_heights: Optional[Dict[str, float]] = None,
    freeze_panes: Optional[str] = None,
    auto_fit_columns: bool = False,
    min_col_width: float = 10.0,
    max_col_width: float = 60.0,
    col_padding: float = 4.0,
    styles: Optional[List[Dict[str, Any]]] = None,
    merge_ranges: Optional[List[str]] = None,
    show_gridlines: Optional[bool] = None,
    auto_color_financial_model: bool = False,
) -> Dict[str, Any]:
    """Batch-format one worksheet in one call. workbook_id auto-resolves if 1 workbook is active. Supports comma-separated ranges and auto_color_financial_model=True."""
    try:
        return await excel.format_worksheet(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            column_widths=column_widths,
            row_heights=row_heights,
            freeze_panes=freeze_panes,
            auto_fit_columns=auto_fit_columns,
            min_col_width=min_col_width,
            max_col_width=max_col_width,
            col_padding=col_padding,
            styles=styles,
            merge_ranges=merge_ranges,
            show_gridlines=show_gridlines,
            auto_color_financial_model=auto_color_financial_model,
        )
    except Exception as exc:
        return _error_response("format_worksheet", exc)


@_tool(app="excel")
async def batch_write_ranges(
    workbook_id: Optional[str] = None,
    operations: Optional[List[Dict[str, Any]]] = None,
    allow_formula_overwrite: bool = False,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Write multiple 2D ranges or drag-fill formulas across sheets in one call. workbook_id auto-resolves if 1 workbook is active; optional top-level sheet_name sets the default sheet for ops omitting sheet_name (and 'Sheet!A1:C3' prefixes in range are also parsed). Supports 2D data [[r1c1,r1c2]] or drag/fill via {'sheet_name':'...','range':'C4:H4','fill_formula':'=C2*C3'}."""
    try:
        return await excel.batch_write_ranges(
            workbook_id=workbook_id,
            operations=operations,
            allow_formula_overwrite=allow_formula_overwrite,
            sheet_name=sheet_name,
        )
    except Exception as exc:
        return _error_response("batch_write_ranges", exc)


@_tool(app="excel")
async def batch_format_workbook(
    workbook_id: Optional[str] = None,
    operations: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    """Batch-format worksheets in one call. workbook_id auto-resolves if 1 workbook is active. One-shot workbook polish: operations=[{'auto_color_financial_model':True,'auto_fit_columns':True,'freeze_panes':'auto'}] (omitting sheet_name applies workbook-wide flags to ALL tabs, auto-detects multi-table headers/totals on ANY row/col, applies Smart Magnitude, and freezes below each table's header). For range rules, ALWAYS include explicit 'sheet_name'."""
    try:
        return await excel.batch_format_workbook(
            workbook_id=workbook_id, operations=operations
        )
    except Exception as exc:
        return _error_response("batch_format_workbook", exc)


@_tool(app="excel", compact=False)
async def list_cell_styles(workbook_id: Optional[str] = None) -> Dict[str, Any]:
    """List available Excel named cell styles and built-in financial model presets. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.list_cell_styles(workbook_id=excel._resolve_workbook_id(workbook_id))
    except Exception as exc:
        return _error_response("list_cell_styles", exc)


@_tool(app="excel", compact=False)
async def apply_cell_style(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    cell_range: str = "",
    style_name: str = "",
    overrides: Optional[Dict[str, Any]] = None,
    number_format: Optional[str] = None,
) -> Dict[str, Any]:
    """Apply a named Excel cell style (e.g. 'Good', 'Bad', 'Input', 'Calculation', 'Total', 'Heading 1') to a range. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.apply_cell_style(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            cell_range=cell_range,
            style_name=style_name,
            overrides=overrides,
            number_format=number_format,
        )
    except Exception as exc:
        return _error_response("apply_cell_style", exc)


@_tool(app="excel")
async def add_data_validation(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    cell_range: str = "",
    validation_type: str = "",
    formula1: Any = None,
    formula2: Any = None,
    operator: Optional[str] = None,
    allow_blank: bool = True,
    prompt_title: Optional[str] = None,
    prompt: Optional[str] = None,
    error_title: Optional[str] = None,
    error: Optional[str] = None,
    error_style: str = "stop",
    show_dropdown: bool = True,
) -> Dict[str, Any]:
    """Add an Excel data validation rule (list, whole, decimal, date, time, textLength, custom). workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.add_data_validation(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            cell_range=cell_range,
            validation_type=validation_type,
            formula1=formula1,
            formula2=formula2,
            operator=operator,
            allow_blank=allow_blank,
            prompt_title=prompt_title,
            prompt=prompt,
            error_title=error_title,
            error=error,
            error_style=error_style,
            show_dropdown=show_dropdown,
        )
    except Exception as exc:
        return _error_response("add_data_validation", exc)


@_tool(app="excel", compact=False)
async def list_data_validations(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """List all data validation rules defined on an Excel worksheet. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.list_data_validations(
            workbook_id=excel._resolve_workbook_id(workbook_id), sheet_name=sheet_name
        )
    except Exception as exc:
        return _error_response("list_data_validations", exc)


@_tool(app="excel", compact=False)
async def clear_data_validations(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    cell_range: Optional[str] = None,
) -> Dict[str, Any]:
    """Clear all data validation rules on a sheet, or only rules matching cell_range. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.clear_data_validations(
            workbook_id=excel._resolve_workbook_id(workbook_id), sheet_name=sheet_name, cell_range=cell_range
        )
    except Exception as exc:
        return _error_response("clear_data_validations", exc)


@_tool(app="excel")
async def add_excel_table(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    table_name: str = "",
    data_range: str = "",
    range: Optional[str] = None,
    style_name: str = "TableStyleMedium9",
    show_first_column: bool = False,
    show_last_column: bool = False,
    show_row_stripes: bool = True,
    show_column_stripes: bool = False,
) -> Dict[str, Any]:
    """Create a structured Excel Table (TableStyleLight*/Medium*/Dark*) with banded row/column toggles. workbook_id auto-resolves if 1 workbook is active."""
    try:
        effective_range = (data_range or range or "").strip()
        if not effective_range:
            raise ValueError(
                "add_excel_table requires data_range (e.g. data_range='A1:F5'), "
                "sheet_name (e.g. 'Quarterly Sales'), and table_name (e.g. 'SalesTable')."
            )
        effective_name = (table_name or "").strip() or "Table1"
        return await excel.add_table(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            table_name=effective_name,
            data_range=effective_range,
            style_name=style_name,
            show_first_column=show_first_column,
            show_last_column=show_last_column,
            show_row_stripes=show_row_stripes,
            show_column_stripes=show_column_stripes,
        )
    except Exception as exc:
        return _error_response("add_excel_table", exc)


@_tool(app="excel", compact=False)
async def update_excel_table_style(
    workbook_id: Optional[str] = None,
    table_name: str = "",
    sheet_name: Optional[str] = None,
    style_name: Optional[str] = None,
    show_first_column: Optional[bool] = None,
    show_last_column: Optional[bool] = None,
    show_row_stripes: Optional[bool] = None,
    show_column_stripes: Optional[bool] = None,
) -> Dict[str, Any]:
    """Update the table style (TableStyleLight*/Medium*/Dark*) and banded toggles of an existing Excel Table. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.update_table_style(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            table_name=table_name,
            sheet_name=sheet_name,
            style_name=style_name,
            show_first_column=show_first_column,
            show_last_column=show_last_column,
            show_row_stripes=show_row_stripes,
            show_column_stripes=show_column_stripes,
        )
    except Exception as exc:
        return _error_response("update_excel_table_style", exc)


@_tool(app="excel")
async def create_chart(
    workbook_id: Optional[str] = None,
    sheet_name: Optional[str] = None,
    chart_type: str = "",
    data_range: str = "",
    chart_title: str = "",
    position: str = "E5",
) -> Dict[str, Any]:
    """Create a bar, line, or pie chart from a worksheet data range. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.create_chart(
            workbook_id=excel._resolve_workbook_id(workbook_id),
            sheet_name=sheet_name,
            chart_type=chart_type,
            data_range=data_range,
            chart_title=chart_title,
            position=position,
        )
    except Exception as exc:
        return _error_response("create_chart", exc)


@_tool(app="excel")
async def save_workbook(
    workbook_id: Optional[str] = None,
    file_path: Optional[str] = None,
    format: str = "xlsx",
    sheet_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Save a workbook in-place or to a new path (.xlsx, .xlsm, .csv, .pdf). workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.save_workbook(
            workbook_id=workbook_id,
            file_path=file_path,
            format=format,
            sheet_name=sheet_name,
        )
    except Exception as exc:
        return _error_response("save_workbook", exc)


@_tool(app="excel")
async def close_workbook(
    workbook_id: Optional[str] = None,
    save_before_close: bool = False,
) -> Dict[str, Any]:
    """Close an in-memory workbook, release memory, and clean up temporary files. workbook_id auto-resolves if 1 workbook is active."""
    try:
        return await excel.close_workbook(
            workbook_id=workbook_id, save_before_close=save_before_close
        )
    except Exception as exc:
        return _error_response("close_workbook", exc)


# ======================================================================
# EXCEL LIVE-APP TOOLS (Microsoft Excel 16.115 via AppleScript)
# ======================================================================

@_tool(app="excel")
async def excel_live_list_workbooks() -> Dict[str, Any]:
    """List all workbooks currently open in the live Microsoft Excel app (with sheets)."""
    try:
        wbs = await applescript.list_open_excel_workbooks()
        return {"status": "success", "count": len(wbs), "workbooks": wbs}
    except Exception as exc:
        return _error_response("excel_live_list_workbooks", exc)


@_tool(app="excel")
async def excel_live_open(
    file_path: str,
    read_only: bool = False,
    activate: bool = False,
) -> Dict[str, Any]:
    """Open an .xlsx, .xlsm, or .csv file directly in live Microsoft Excel."""
    try:
        _reject_internal_temp_live_open(file_path, "Microsoft Excel", "save_workbook")
        return await applescript.open_excel_file(
            file_path=file_path, read_only=read_only, activate=activate
        )
    except Exception as exc:
        return _error_response("excel_live_open", exc)


@_tool(app="excel", compact=False)
async def excel_live_create(activate: bool = False) -> Dict[str, Any]:
    """Create a new workbook directly in live Microsoft Excel."""
    try:
        return await applescript.create_excel_workbook(activate=activate)
    except Exception as exc:
        return _error_response("excel_live_create", exc)


@_tool(app="excel", compact=False)
async def excel_live_list_worksheets(
    workbook_name: Optional[str] = None,
) -> Dict[str, Any]:
    """List worksheets in a specific (or active) live Excel workbook."""
    try:
        return await applescript.list_live_excel_worksheets(workbook_name=workbook_name)
    except Exception as exc:
        return _error_response("excel_live_list_worksheets", exc)


@_tool(app="excel", compact=False)
async def excel_live_add_worksheet(
    sheet_name: str,
    workbook_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Add a new worksheet to a specific (or active) live Excel workbook."""
    try:
        return await applescript.add_live_excel_worksheet(
            sheet_name=sheet_name, workbook_name=workbook_name
        )
    except Exception as exc:
        return _error_response("excel_live_add_worksheet", exc)


@_tool(app="excel", compact=False)
async def excel_live_rename_worksheet(
    old_name: str,
    new_name: str,
    workbook_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Rename a worksheet in a specific (or active) live Excel workbook."""
    try:
        return await applescript.rename_live_excel_worksheet(
            old_name=old_name, new_name=new_name, workbook_name=workbook_name
        )
    except Exception as exc:
        return _error_response("excel_live_rename_worksheet", exc)


@_tool(app="excel", compact=False)
async def excel_live_activate_worksheet(
    sheet_name: str,
    workbook_name: Optional[str] = None,
) -> Dict[str, Any]:
    """Switch the active worksheet in a specific (or active) live Excel workbook."""
    try:
        return await applescript.activate_live_excel_worksheet(
            sheet_name=sheet_name, workbook_name=workbook_name
        )
    except Exception as exc:
        return _error_response("excel_live_activate_worksheet", exc)


def _attach_dirty_in_memory_warning(result: Dict[str, Any]) -> Dict[str, Any]:
    """Warn if an in-memory workbook matching the live Excel workbook has unsaved file-level edits."""
    if not isinstance(result, dict) or result.get("status") != "success":
        return result
    live_name = str(result.get("workbook_name") or "").strip().lower()
    if not live_name:
        return result
    for wb_id, entry in excel.active_workbooks.items():
        if not entry.get("dirty"):
            continue
        meta = entry.get("metadata") or {}
        candidates = {
            Path(str(entry.get("working_path") or "")).name.lower(),
            Path(str(meta.get("file_path") or "")).name.lower(),
            Path(str(meta.get("original_path") or "")).name.lower(),
        }
        if live_name in candidates:
            result["warning"] = (
                f"In-memory workbook '{wb_id}' has unsaved file-level edits not yet visible in live Excel. "
                "Call save_workbook() first so live Excel reloads them, or use excel_live_batch_write_cells for direct live edits."
            )
            break
    return result


@_tool(app="excel", compact=False)
async def excel_live_read_range(
    range_ref: Optional[str] = None,
    cell_range: Optional[str] = None,
    sheet_name: Optional[str] = None,
    workbook_name: Optional[str] = None,
    include_formulas: bool = False,
    include_display_values: Optional[bool] = None,
) -> Dict[str, Any]:
    """Read live evaluated values (plus formatted display_values on compact ranges & '###' clipping warnings) from Excel."""
    try:
        effective_ref = range_ref if range_ref is not None else cell_range
        if not effective_ref:
            raise ValueError("Provide range_ref or cell_range (e.g. 'B2:F15')")
        kwargs: Dict[str, Any] = {
            "range_ref": effective_ref,
            "sheet_name": sheet_name,
            "workbook_name": workbook_name,
            "include_formulas": include_formulas,
        }
        if include_display_values is not None:
            kwargs["include_display_values"] = include_display_values
        res = await applescript.read_live_excel_range(**kwargs)
        return _attach_dirty_in_memory_warning(res)
    except Exception as exc:
        return _error_response("excel_live_read_range", exc)


@_tool(app="excel")
async def excel_live_write_cell(
    cell_ref: str,
    value: Any,
    sheet_name: Optional[str] = None,
    workbook_name: Optional[str] = None,
    is_formula: bool = False,
    number_format: Optional[str] = None,
    recalculate: bool = True,
    allow_formula_overwrite: bool = False,
) -> Dict[str, Any]:
    """Write a typed value or formula to a cell in a targeted live Excel workbook & sheet."""
    try:
        return await applescript.write_live_excel_cell(
            cell_ref=cell_ref,
            value=value,
            sheet_name=sheet_name,
            workbook_name=workbook_name,
            is_formula=is_formula,
            number_format=number_format,
            recalculate=recalculate,
            allow_formula_overwrite=allow_formula_overwrite,
        )
    except Exception as exc:
        return _error_response("excel_live_write_cell", exc)


@_tool(app="excel", compact=False)
async def excel_live_write_range(
    start_cell: str,
    data: List[List[Any]],
    sheet_name: Optional[str] = None,
    workbook_name: Optional[str] = None,
    number_format: Optional[str] = None,
    recalculate: bool = True,
    allow_formula_overwrite: bool = False,
) -> Dict[str, Any]:
    """Write a 2D array of values/formulas into a targeted live Excel workbook & sheet."""
    try:
        return await applescript.write_live_excel_range(
            start_cell=start_cell,
            data=data,
            sheet_name=sheet_name,
            workbook_name=workbook_name,
            number_format=number_format,
            recalculate=recalculate,
            allow_formula_overwrite=allow_formula_overwrite,
        )
    except Exception as exc:
        return _error_response("excel_live_write_range", exc)


@_tool(app="excel")
async def excel_live_batch_write_cells(
    cells: Any,
    sheet_name: Optional[str] = None,
    workbook_name: Optional[str] = None,
    recalculate: bool = True,
    allow_formula_overwrite: bool = False,
) -> Dict[str, Any]:
    """Write multiple non-contiguous cells/formulas in ONE live Excel call and recalculate once.

    Pass `cells` as {'G2': '=...', 'G11': 'N/A'} or [{'cell': 'G2', 'value': '=...', 'sheet_name': '...', 'number_format': '0.0%'}].
    """
    try:
        return await applescript.batch_write_live_excel_cells(
            cells=cells,
            sheet_name=sheet_name,
            workbook_name=workbook_name,
            recalculate=recalculate,
            allow_formula_overwrite=allow_formula_overwrite,
        )
    except Exception as exc:
        return _error_response("excel_live_batch_write_cells", exc)


@_tool(app="excel")
async def excel_live_batch_read_ranges(
    ranges: List[Dict[str, Any]],
    workbook_name: Optional[str] = None,
    include_formulas: bool = False,
    include_display_values: Optional[bool] = None,
) -> Dict[str, Any]:
    """Read multiple ranges across one or more live Excel sheets in ONE call (auto-includes formatted display_values on compact tables & flags '###' clipped cells).

    Pass `ranges` as [{'sheet_name': 'Income Statement', 'range_ref': 'B2:F16'}, {'sheet_name': 'Balance Sheet', 'range_ref': 'F23:F27'}].
    """
    try:
        kwargs: Dict[str, Any] = {
            "ranges": ranges,
            "workbook_name": workbook_name,
            "include_formulas": include_formulas,
        }
        if include_display_values is not None:
            kwargs["include_display_values"] = include_display_values
        res = await applescript.batch_read_live_excel_ranges(**kwargs)
        return _attach_dirty_in_memory_warning(res)
    except Exception as exc:
        return _error_response("excel_live_batch_read_ranges", exc)


def _find_dirty_in_memory_workbook(
    workbook_name: Optional[str] = None,
    file_path: Optional[str] = None,
) -> Optional[str]:
    """Return the workbook_id of a matching active in-memory workbook with unsaved file-level edits."""
    targets: set[str] = set()
    if workbook_name and str(workbook_name).strip():
        targets.add(Path(str(workbook_name).strip()).name.lower())
    if file_path and str(file_path).strip():
        targets.add(Path(str(file_path).strip()).expanduser().name.lower())

    dirty_ids: List[str] = []
    for wb_id, entry in excel.active_workbooks.items():
        if not entry.get("dirty"):
            continue
        dirty_ids.append(wb_id)
        if targets:
            meta = entry.get("metadata") or {}
            candidates = {
                Path(str(entry.get("working_path") or "")).name.lower(),
                Path(str(meta.get("file_path") or "")).name.lower(),
                Path(str(meta.get("original_path") or "")).name.lower(),
            }
            if targets & candidates:
                return wb_id
    if not targets and len(dirty_ids) == 1:
        return dirty_ids[0]
    return None


@_tool(app="excel")
async def excel_live_save(
    workbook_name: Optional[str] = None,
    file_path: Optional[str] = None,
    file_format: str = "xlsx",
) -> Dict[str, Any]:
    """Save a live Excel workbook in-place or Save As (.xlsx, .xlsm, .csv, .pdf)."""
    try:
        dirty_wb_id = _find_dirty_in_memory_workbook(workbook_name=workbook_name, file_path=file_path)
        if dirty_wb_id:
            return {
                "status": "error",
                "action": "excel_live_save",
                "error_type": "UnsavedInMemoryEditsError",
                "error": (
                    f"In-memory workbook '{dirty_wb_id}' has unsaved file-level edits "
                    "(from batch_write_ranges/write_cell/batch_format_workbook). "
                    "Calling excel_live_save would save Excel's stale copy over disk and lose those edits."
                ),
                "hint": (
                    f"Call save_workbook(workbook_id='{dirty_wb_id}', file_path=...) instead — "
                    "M365_MCP_LIVE_SYNC=auto will write your edits and automatically reload Excel."
                ),
            }
        return await applescript.save_excel_workbook(
            workbook_name=workbook_name,
            file_path=file_path,
            file_format=file_format,
        )
    except Exception as exc:
        return _error_response("excel_live_save", exc)


@_tool(app="excel")
async def excel_live_close(
    workbook_name: Optional[str] = None,
    save_changes: Union[bool, str] = "no",
) -> Dict[str, Any]:
    """Close a live Excel workbook cleanly without triggering unsaved-changes dialogs."""
    try:
        return await applescript.close_excel_workbook(
            workbook_name=workbook_name,
            save_changes=_normalize_save_changes(save_changes),
        )
    except Exception as exc:
        return _error_response("excel_live_close", exc)


@_tool(app="excel")
async def excel_live_circle_invalid(
    sheet_name: Optional[str] = None,
    workbook_name: Optional[str] = None,
    clear: bool = False,
) -> Dict[str, Any]:
    """Circle invalid data-validation entries (`clear=False`) or clear validation circles (`clear=True`) in live Excel."""
    try:
        return await applescript.excel_live_circle_invalid(
            sheet_name=sheet_name, workbook_name=workbook_name, clear=clear
        )
    except Exception as exc:
        return _error_response("excel_live_circle_invalid", exc)


@_tool(app="excel", compact=False)
async def excel_live_custom_views(
    action: str = "list",
    view_name: Optional[str] = None,
    workbook_name: Optional[str] = None,
) -> Dict[str, Any]:
    """List (`action='list'`) or display (`action='show'`) Custom Views in live Microsoft Excel."""
    try:
        return await applescript.excel_live_custom_views(
            action=action, view_name=view_name, workbook_name=workbook_name
        )
    except Exception as exc:
        return _error_response("excel_live_custom_views", exc)


@_tool(app="excel", compact=False)
async def excel_live_update_table_style(
    table_name: str,
    sheet_name: Optional[str] = None,
    workbook_name: Optional[str] = None,
    style_name: Optional[str] = None,
    show_row_stripes: Optional[bool] = None,
    show_column_stripes: Optional[bool] = None,
    show_first_column: Optional[bool] = None,
    show_last_column: Optional[bool] = None,
) -> Dict[str, Any]:
    """Update a live Excel Table (list object) style and banded row/column toggles via AppleScript."""
    try:
        return await applescript.excel_live_update_table_style(
            table_name=table_name,
            sheet_name=sheet_name,
            workbook_name=workbook_name,
            style_name=style_name,
            show_row_stripes=show_row_stripes,
            show_column_stripes=show_column_stripes,
            show_first_column=show_first_column,
            show_last_column=show_last_column,
        )
    except Exception as exc:
        return _error_response("excel_live_update_table_style", exc)


@_tool(app="excel", compact=False)
async def excel_view_capabilities() -> Dict[str, Any]:
    """Report supported Excel 16.115 / openpyxl features vs. manual UI-only View ribbon controls (e.g. Focus Cell)."""
    return sanitize_for_json(applescript.get_excel_view_capabilities())


# ======================================================================
# POWERPOINT TOOLS (preserved)
# ======================================================================

@_tool(app="powerpoint", compact=False)
async def create_presentation(
    title: str,
    theme: str = "default",
    template_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Create a new PowerPoint presentation (.pptx). Parent directories are auto-created on save_presentation."""
    try:
        return sanitize_for_json(
            await powerpoint.create_presentation(
                title=title, theme=theme, template_path=template_path
            )
        )
    except Exception as exc:
        return _error_response("create_presentation", exc)


@_tool(app="powerpoint", compact=False)
async def add_slide(
    presentation_id: Optional[str] = None,
    layout: str = "Title and Content",
    position: Optional[int] = None,
) -> Dict[str, Any]:
    """Add a new slide to a PowerPoint presentation. presentation_id auto-resolves if 1 presentation is active."""
    try:
        return sanitize_for_json(
            await powerpoint.add_slide(
                presentation_id=presentation_id, layout=layout, position=position
            )
        )
    except Exception as exc:
        return _error_response("add_slide", exc)


@_tool(app="powerpoint")
async def batch_add_slides(
    slides: List[Dict[str, Any]],
    presentation_id: Optional[str] = None,
) -> Dict[str, Any]:
    """Add and populate multiple slides (title, subtitle, bullets/content, notes) in ONE call. presentation_id auto-resolves if 1 presentation is active."""
    try:
        return sanitize_for_json(
            await powerpoint.batch_add_slides(
                presentation_id=presentation_id, slides=slides
            )
        )
    except Exception as exc:
        return _error_response("batch_add_slides", exc)


@_tool(app="powerpoint", compact=False)
async def add_text_to_slide(
    slide_id: Optional[str] = None,
    text: str = "",
    placeholder: str = "content",
    formatting: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Add text content to a PowerPoint slide."""
    try:
        return sanitize_for_json(
            await powerpoint.add_text(
                slide_id=slide_id,
                text=text,
                placeholder=placeholder,
                formatting=formatting or {},
            )
        )
    except Exception as exc:
        return _error_response("add_text_to_slide", exc)


@_tool(app="powerpoint", compact=False)
async def add_image_to_slide(
    slide_id: Optional[str] = None,
    image_source: str = "",
    position: Optional[Dict[str, float]] = None,
    size: Optional[Dict[str, float]] = None,
) -> Dict[str, Any]:
    """Add an image to a PowerPoint slide."""
    try:
        return sanitize_for_json(
            await powerpoint.add_image(
                slide_id=slide_id,
                image_source=image_source,
                position=position or {"x": 1.0, "y": 1.0},
                size=size or {"width": 4.0, "height": 3.0},
            )
        )
    except Exception as exc:
        return _error_response("add_image_to_slide", exc)


@_tool(app="powerpoint", compact=False)
async def add_speaker_notes(slide_id: Optional[str] = None, notes: str = "") -> Dict[str, Any]:
    """Add speaker notes to a PowerPoint slide."""
    try:
        return sanitize_for_json(
            await powerpoint.add_speaker_notes(slide_id=slide_id, notes=notes)
        )
    except Exception as exc:
        return _error_response("add_speaker_notes", exc)


@_tool(app="powerpoint", compact=False)
async def save_presentation(
    presentation_id: Optional[str] = None,
    file_path: str = "",
    format: str = "pptx",
) -> Dict[str, Any]:
    """Save a PowerPoint presentation to file. presentation_id auto-resolves if 1 presentation is active."""
    try:
        return sanitize_for_json(
            await powerpoint.save_presentation(
                presentation_id=presentation_id, file_path=file_path, format=format
            )
        )
    except Exception as exc:
        return _error_response("save_presentation", exc)


# ======================================================================
# UTILITY & STATUS TOOLS
# ======================================================================

@_tool(app="word", compact=False)
async def list_active_documents() -> List[Dict[str, Any]]:
    """List all in-memory Word documents currently managed by the server."""
    return await word.list_documents()


@_tool(app="excel", compact=False)
async def list_active_workbooks() -> List[Dict[str, Any]]:
    """List all in-memory Excel workbooks currently managed by the server."""
    return await excel.list_workbooks()


@_tool(app="powerpoint", compact=False)
async def list_active_presentations() -> List[Dict[str, Any]]:
    """List all in-memory PowerPoint presentations currently managed by the server."""
    return sanitize_for_json(await powerpoint.list_presentations())


_SKILLS_DIR = Path(__file__).resolve().parent.parent / "skills"


def _discover_finance_skills() -> List[Dict[str, str]]:
    """Scan skills/*/SKILL.md live from disk so models always see current versions and slugs."""
    skills: List[Dict[str, str]] = []
    if not _SKILLS_DIR.is_dir():
        return skills
    for skill_md in sorted(_SKILLS_DIR.glob("*/SKILL.md")):
        try:
            txt = skill_md.read_text(encoding="utf-8")
            fm = txt.split("---", 2)[1] if txt.startswith("---") and txt.count("---") >= 2 else ""
            name_m = re.search(r"^name:\s*(.+?)\s*$", fm, re.MULTILINE)
            ver_m = re.search(r"^version:\s*(.+?)\s*$", fm, re.MULTILINE)
            desc_m = re.search(r"^description:\s*(.+?)\s*$", fm, re.MULTILINE)
            slug = (name_m.group(1).strip().strip('"').strip("'") if name_m else skill_md.parent.name)
            ver = (ver_m.group(1).strip().strip('"').strip("'") if ver_m else "")
            desc = (desc_m.group(1).strip().strip('"').strip("'") if desc_m else "")
            desc = re.sub(r"^\[v[^\]]+\]\s*", "", desc)
            skills.append({"name": slug, "version": ver, "summary": desc[:110], "_path": str(skill_md)})
        except Exception:
            continue
    return skills


@_tool(app="core")
async def check_office_status(
    skill_name: Optional[str] = None,
    working_dir: Optional[str] = None,
) -> Dict[str, Any]:
    """Check installed Office versions, active profile/tools, and live Finance skills. Pass skill_name (e.g. 'capex-project-analyzer') to load the full live SKILL.md playbook, and optional working_dir to register your Project Working Folder for bare relative paths."""
    try:
        if working_dir and str(working_dir).strip():
            set_client_project_dir(str(working_dir).strip())
        versions = await applescript.get_office_version_info()
        discovered = _discover_finance_skills()
        active_proj_dir = get_client_project_dir()
        default_out = get_default_output_dir()
        blocked_dirs = get_blocked_write_dirs()
        restrict_to_proj = is_write_restricted_to_project_dir()
        profile_info: Dict[str, Any] = {
            "active_profile": _ACTIVE_PROFILE_LABEL,
            "registered_tool_count": len(_REGISTERED_TOOL_NAMES),
            "registered_tools": list(_REGISTERED_TOOL_NAMES),
            "available_profiles": list(_AVAILABLE_PROFILES),
            "live_sync_mode": excel.get_live_sync_mode(),
            "available_live_sync_modes": ["auto", "off"],
            "formula_guard_mode": excel.get_formula_guard_mode(),
            "available_formula_guard_modes": ["literals_only", "strict", "off"],
            "background_mode": get_background_mode(),
            "available_background_modes": ["no_focus", "strict", "off"],
        }
        if active_proj_dir is not None:
            profile_info["active_project_dir"] = str(active_proj_dir)
        if default_out is not None:
            profile_info["default_output_dir"] = str(default_out)
        if blocked_dirs:
            profile_info["blocked_write_dirs"] = [str(d) for d in blocked_dirs]
        if restrict_to_proj:
            profile_info["restrict_write_to_project_dir"] = True
        payload: Dict[str, Any] = {
            "status": "running",
            "server_version": "2.4.0",
            "word_available": versions["word"]["installed"],
            "excel_available": versions["excel"]["installed"],
            "powerpoint_available": versions["powerpoint"]["installed"],
            "profile": profile_info,
            "applications": versions,
            "active_in_memory": {
                "documents": len(word.active_documents),
                "workbooks": len(excel.active_workbooks),
                "presentations": len(powerpoint.active_presentations),
            },
        }
        if skill_name and str(skill_name).strip():
            query = str(skill_name).strip().lower()
            matched = next((s for s in discovered if s["name"].lower() == query), None)
            if matched is None:
                matched = next((s for s in discovered if query in s["name"].lower()), None)
            if matched is not None:
                raw_skill_txt = Path(matched["_path"]).read_text(encoding="utf-8")
                payload["loaded_skill"] = {
                    "name": matched["name"],
                    "version": matched["version"],
                    "precedence_rule": (
                        "MANDATORY PLAYBOOK OVERRIDE: Follow this <mandatory_playbook>'s exact sheet_names, "
                        "row coordinates, formulas, and step order verbatim; they strictly override generic modeling priors."
                    ),
                    "content": (
                        f'<mandatory_playbook name="{matched["name"]}" version="{matched["version"]}">\n'
                        f"{raw_skill_txt}\n"
                        f"</mandatory_playbook>"
                    ),
                }
            else:
                payload["skill_error"] = f"Skill '{skill_name}' not found."
                payload["finance_skills"] = [
                    {"name": s["name"], "version": s["version"], "summary": s["summary"]}
                    for s in discovered
                ]
        elif discovered:
            payload["finance_skills"] = {
                "instruction": (
                    "For finance builds, audits, formatting, what-if, or boardroom briefs, "
                    "call check_office_status(skill_name='<name>') to load the exact live playbook before editing."
                ),
                "available": [
                    {"name": s["name"], "version": s["version"], "summary": s["summary"]}
                    for s in discovered
                ],
            }
        return sanitize_for_json(payload)
    except Exception as exc:
        return _error_response("check_office_status", exc)


# ======================================================================
# RESOURCES
# ======================================================================

@mcp.resource("m365-mac-mcp://status")
async def get_server_status() -> str:
    """Get server status summary."""
    status = await check_office_status()
    return json.dumps(status, separators=(",", ":"), ensure_ascii=False, default=str)


def main() -> None:
    """Start the m365-mac-mcp stdio server."""
    logger.info(
        f"Starting Microsoft 365 for Mac MCP (m365-mac-mcp, stdio) — profile='{_ACTIVE_PROFILE_LABEL}', "
        f"tools={len(_REGISTERED_TOOL_NAMES)}, live_sync='{excel.get_live_sync_mode()}', "
        f"formula_guard='{excel.get_formula_guard_mode()}', background_mode='{get_background_mode()}'. "
        f"Valid M365_MCP_PROFILE: {', '.join(_AVAILABLE_PROFILES)}; "
        f"M365_MCP_LIVE_SYNC: auto | off; M365_MCP_FORMULA_GUARD: literals_only | strict | off; "
        f"M365_MCP_BACKGROUND_MODE: no_focus | strict | off."
    )
    mcp.run()


if __name__ == "__main__":
    main()
