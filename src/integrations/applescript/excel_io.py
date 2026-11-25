"""Live Microsoft Excel range/cell read, batch read/write, display-value extraction, and formula guard."""

import os
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

from utils.logger import setup_logger
from utils.validators import (
    sanitize_excel_formula_quotes,
    sanitize_for_json,
    validate_cell_ref,
    validate_range_ref,
)

logger = setup_logger(__name__)
from integrations.applescript.base import (
    _FS,
    _GS,
    _RS,
    _US,
)

class AppleScriptExcelIOMixin:
    """Live Microsoft Excel range/cell read, batch read/write, display-value extraction, and formula guard."""

    @staticmethod
    def _process_live_display_strings(
        values: List[List[Any]],
        str_block: str,
        range_str: str,
        include_display_values: Any = "auto",
    ) -> Tuple[Optional[List[List[str]]], List[str]]:
        """Backfill Excel error tokens from `string value`, detect '###' clipped columns, and return formatted display strings.

        Token-efficient: in 'auto' mode, returns `display_values` only for compact ranges (<= 50 cells)
        where at least one cell's formatted display string differs meaningfully from its raw value.
        """
        if not str_block or not values:
            return None, []
        str_rows = [row.split(_US) for row in str_block.split(_RS)]

        # Parse top-left (col, row) for exact clipped cell coordinates
        start_col_idx, start_row_idx = 1, 1
        try:
            first_part = (range_str or "A1").split(":", 1)[0].strip().replace("$", "")
            m = re.match(r"^([A-Za-z]+)(\d+)$", first_part)
            if m:
                col_s, row_s = m.group(1).upper(), int(m.group(2))
                c_num = 0
                for ch in col_s:
                    c_num = c_num * 26 + (ord(ch) - ord("A") + 1)
                start_col_idx, start_row_idx = c_num, row_s
        except Exception:
            pass

        def _col_to_letters(n: int) -> str:
            out = ""
            while n > 0:
                n, rem = divmod(n - 1, 26)
                out = chr(65 + rem) + out
            return out

        clipped_cells: List[str] = []
        has_meaningful_format_diff = False
        total_cells = sum(len(r) for r in values)

        for r_i, row in enumerate(values):
            if r_i >= len(str_rows):
                break
            srow = str_rows[r_i]
            for c_i, cell_val in enumerate(row):
                if c_i >= len(srow):
                    break
                sv = srow[c_i].strip()
                # Backfill Excel error tokens (#DIV/0!, #NUM!, #N/A, #VALUE!, #REF!, #NAME?)
                if (cell_val == "" or cell_val is None) and sv.startswith("#") and not set(sv) == {"#"}:
                    row[c_i] = sv
                # Detect Excel column-width overflow ('###' / '########')
                if len(sv) >= 3 and set(sv) == {"#"}:
                    if len(clipped_cells) < 10:
                        clipped_cells.append(f"{_col_to_letters(start_col_idx + c_i)}{start_row_idx + r_i}")
                # Check if formatted display string adds real formatting/scaling info over raw value
                raw_s = "" if cell_val is None else str(cell_val).strip()
                if sv and raw_s and sv != raw_s:
                    if any(ch in sv for ch in ("$", "€", "£", "¥", "%", ",", "(", "x")) or "E+" in raw_s or "E-" in raw_s:
                        has_meaningful_format_diff = True

        mode = include_display_values
        if isinstance(mode, bool):
            want_display = mode
        elif isinstance(mode, str) and mode.strip().lower() in ("true", "1", "yes", "always"):
            want_display = True
        elif isinstance(mode, str) and mode.strip().lower() in ("false", "0", "no", "never", "off"):
            want_display = False
        else:
            # 'auto' (default): attach formatted display strings on compact KPI/summary tables (<= 50 cells)
            want_display = total_cells <= 50 and (has_meaningful_format_diff or bool(clipped_cells))

        return (str_rows if want_display else None), clipped_cells

    @staticmethod
    def _trim_and_cap_live_grid(
        values: List[List[Any]],
        display_values: Optional[List[List[str]]],
        max_cells: int = 3000,
    ) -> Tuple[List[List[Any]], Optional[List[List[str]]], bool, int]:
        """Trim trailing all-empty rows/cols from live Excel reads and cap giant ranges to protect context."""
        if not values:
            return values, display_values, False, 0

        def _is_empty(v: Any) -> bool:
            return v is None or v == ""

        end_row = len(values)
        while end_row > 0 and all(_is_empty(v) for v in values[end_row - 1]):
            end_row -= 1
        trimmed = values[:end_row]
        if not trimmed:
            return [], (None if display_values is None else []), False, 0

        max_used_col = 0
        for row in trimmed:
            for idx in range(len(row) - 1, -1, -1):
                if not _is_empty(row[idx]):
                    if idx + 1 > max_used_col:
                        max_used_col = idx + 1
                    break
        if max_used_col == 0:
            return [], (None if display_values is None else []), False, 0

        trimmed = [row[:max_used_col] for row in trimmed]
        trimmed_disp = (
            [row[:max_used_col] for row in display_values[:end_row]]
            if display_values is not None
            else None
        )
        orig_rows = len(trimmed)
        truncated = False
        if orig_rows * max_used_col > max_cells:
            keep_rows = max(1, max_cells // max_used_col)
            trimmed = trimmed[:keep_rows]
            if trimmed_disp is not None:
                trimmed_disp = trimmed_disp[:keep_rows]
            truncated = True
        return trimmed, trimmed_disp, truncated, orig_rows

    async def read_live_excel_range(
        self,
        range_ref: str,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        include_formulas: bool = False,
        include_display_values: Any = "auto",
    ) -> Dict[str, Any]:
        """Read cell or range values (and optionally formulas & formatted display strings) from a live Excel workbook."""
        clean_ref = validate_range_ref(range_ref)
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set refStr to item 3 of argv
            set incFormulas to (item 4 of argv) is "true"
            set us to character id 31
            set rs to character id 30
            set gs to character id 29

            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                if shTarget is "" then
                    set sh to active sheet of wb
                else
                    set sh to worksheet shTarget of wb
                end if
                set wbName to name of wb
                set shName to name of sh
                set rng to range refStr of sh
                set rawVals to value of rng
                set rawStrs to missing value
                try
                    set rawStrs to string value of rng
                end try
                set rawFmls to missing value
                if incFormulas then
                    set rawFmls to formula of rng
                end if
            end tell

            -- Normalize scalar / 1D list / 2D list outside Excel's tell block (fast & avoids sub-range indexing quirks)
            set valBlock to my matrix_to_text(rawVals, us, rs)
            set fmlBlock to ""
            if incFormulas then
                set fmlBlock to my matrix_to_text(rawFmls, us, rs)
            end if
            set strBlock to my matrix_to_text(rawStrs, us, rs)
            return wbName & us & shName & gs & valBlock & gs & fmlBlock & gs & strBlock
        end run

        on matrix_to_text(rawData, us, rs)
            if rawData is missing value then
                return ""
            end if
            if class of rawData is not list then
                return (rawData as text)
            end if
            if (count of rawData) is 0 then
                return ""
            end if
            if class of (item 1 of rawData) is not list then
                set matrixRows to {rawData}
            else
                set matrixRows to rawData
            end if
            set outRows to {}
            repeat with rowItem in matrixRows
                set rowList to contents of rowItem
                if class of rowList is not list then
                    set rowList to {rowList}
                end if
                set cellTexts to {}
                repeat with cellItem in rowList
                    set v to contents of cellItem
                    if v is missing value then
                        set end of cellTexts to ""
                    else
                        set end of cellTexts to (v as text)
                    end if
                end repeat
                set prevDelims to AppleScript's text item delimiters
                set AppleScript's text item delimiters to us
                set end of outRows to (cellTexts as text)
                set AppleScript's text item delimiters to prevDelims
            end repeat
            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set joined to outRows as text
            set AppleScript's text item delimiters to prevDelims
            return joined
        end matrix_to_text
        """
        raw = await self.execute_applescript(
            script,
            args=[
                workbook_name or "",
                sheet_name or "",
                clean_ref,
                "true" if include_formulas else "false",
            ],
            timeout=30.0,
        )
        sections = raw.split(_GS)
        header = sections[0].split(_US) if sections else ["", ""]
        val_block = sections[1] if len(sections) > 1 else ""
        fml_block = sections[2] if len(sections) > 2 else ""
        str_block = sections[3] if len(sections) > 3 else ""

        values = [row.split(_US) for row in val_block.split(_RS)] if val_block else []
        display_values, clipped_cells = self._process_live_display_strings(
            values, str_block, clean_ref, include_display_values=include_display_values
        )
        values, display_values, grid_truncated, orig_rows = self._trim_and_cap_live_grid(
            values, display_values, max_cells=3000
        )
        result: Dict[str, Any] = {
            "status": "success",
            "workbook_name": header[0],
            "sheet_name": header[1] if len(header) > 1 else "",
            "range": clean_ref,
            "rows": len(values),
            "columns": len(values[0]) if values else 0,
            "values": values,
        }
        if grid_truncated:
            result["truncated"] = True
            result["total_populated_rows"] = orig_rows
            result["pagination_hint"] = (
                f"Live range capped to {len(values)} of {orig_rows} populated rows to protect context. "
                "Request smaller targeted ranges (e.g. KPI summary or specific row slices)."
            )
        if display_values is not None:
            result["display_values"] = display_values
        if clipped_cells:
            result["clipped_cells"] = clipped_cells
            result["warning"] = (
                f"Column width clipping ('###') detected in live Excel at {', '.join(clipped_cells)}; "
                "widen columns via batch_format_workbook(auto_fit_columns=true) and re-save."
            )
        if include_formulas:
            result["formulas"] = self._extract_sparse_formulas(fml_block, clean_ref)
        return result

    @staticmethod
    def _extract_sparse_formulas(fml_block: str, range_ref: str) -> Dict[str, str]:
        """Extract only formula cells ('=...') from a serialized formula matrix into a sparse A1->formula map."""
        if not fml_block:
            return {}
        start_part = (range_ref or "A1").split(":")[0].replace("$", "").strip().upper()
        m = re.match(r"^([A-Z]+)([0-9]+)$", start_part)
        if m:
            col_letters, row_str = m.group(1), m.group(2)
            start_col = 0
            for ch in col_letters:
                start_col = start_col * 26 + (ord(ch) - ord("A") + 1)
            start_row = int(row_str)
        else:
            start_col, start_row = 1, 1

        def _idx_to_col(idx: int) -> str:
            out = ""
            while idx > 0:
                idx, rem = divmod(idx - 1, 26)
                out = chr(65 + rem) + out
            return out

        sparse: Dict[str, str] = {}
        for r_off, row_str_val in enumerate(fml_block.split(_RS)):
            for c_off, cell_val in enumerate(row_str_val.split(_US)):
                if cell_val.startswith("="):
                    coord = f"{_idx_to_col(start_col + c_off)}{start_row + r_off}"
                    sparse[coord] = cell_val
        return sparse

    @staticmethod
    def _sanitize_live_excel_formula(formula: str, sheet_names: Sequence[str]) -> str:
        """Auto-quote same-workbook sheet names with spaces and block non-existent sheet refs.

        Prevents live Microsoft Excel from popping up a modal 'Cannot find <Sheet>. Copy from:'
        file picker when a formula has an unquoted sheet name with spaces (e.g. `=Cash Flow!F15`)
        or references a sheet that does not exist in the open workbook.
        """
        if not formula or not formula.startswith("="):
            return formula

        formula = sanitize_excel_formula_quotes(formula)
        if "!" not in formula:
            return formula

        existing = list(sheet_names)
        existing_set = set(existing)
        sheets_needing_quotes = sorted(
            [
                name
                for name in existing
                if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name)
                or re.match(r"^[A-Za-z]{1,3}[0-9]+$", name)
            ],
            key=len,
            reverse=True,
        )

        # Auto-quote unquoted same-workbook sheet names outside string literals and already-quoted refs
        if sheets_needing_quotes:
            token_re = re.compile(r'("(?:""|[^"])*"|\'(?:\'\'|[^\'])*\'!)')
            parts = token_re.split(formula)
            for idx in range(0, len(parts), 2):
                seg = parts[idx]
                if not seg or "!" not in seg:
                    continue
                for sname in sheets_needing_quotes:
                    escaped = sname.replace("'", "''")
                    pat = re.compile(r"(^|[=(,+\-*/:; &<>^])" + re.escape(sname) + r"!")
                    seg = pat.sub(rf"\1'{escaped}'!", seg)
                parts[idx] = seg
            formula = "".join(parts)

        # Guard against non-existent sheet references that would trigger Excel's modal file picker
        if existing:
            stripped = re.sub(r'"(?:""|[^"])*"', '""', formula)
            ref_re = re.compile(r"(?:'((?:''|[^'])+)'|([A-Za-z0-9_.\[\]]+))!")
            for match in ref_re.finditer(stripped):
                q, u = match.group(1), match.group(2)
                ref = q.replace("''", "'") if q is not None else u
                if not ref:
                    continue
                if "[" in ref or "]" in ref or ref not in existing_set:
                    raise ValueError(
                        f"Formula '{formula}' references unknown/external sheet '{ref}' "
                        f"(would trigger Excel's modal file picker). Available sheets: {existing}. "
                        "Use single quotes for sheet names with spaces, e.g. ='Cash Flow'!F15."
                    )

        return formula

    _warned_legacy_formula_guard: bool = False

    @classmethod
    def get_formula_guard_mode(cls, raw: Optional[str] = None) -> str:
        """Parse M365_MCP_FORMULA_GUARD (or legacy OFFICE_MCP_FORMULA_GUARD): 'literals_only' (default) | 'strict' | 'off'."""
        if raw is None:
            env_val = os.environ.get("M365_MCP_FORMULA_GUARD")
            if env_val is None or not env_val.strip():
                legacy = os.environ.get("OFFICE_MCP_FORMULA_GUARD")
                if legacy is not None and legacy.strip():
                    if not cls._warned_legacy_formula_guard:
                        logger.warning(
                            "OFFICE_MCP_FORMULA_GUARD is deprecated; rename it to M365_MCP_FORMULA_GUARD."
                        )
                        cls._warned_legacy_formula_guard = True
                    env_val = legacy
            val = (env_val or "literals_only").strip().lower()
        else:
            val = raw.strip().lower()
        if not val or val in ("literals_only", "literals", "literal", "default"):
            return "literals_only"
        if val in ("strict", "all", "on"):
            return "strict"
        if val in ("off", "false", "0", "no", "none", "disabled"):
            return "off"
        return "literals_only"

    @staticmethod
    def _formula_guard_response(mode: str, blocked: List[Dict[str, Any]]) -> Dict[str, Any]:
        return sanitize_for_json({
            "status": "confirmation_required",
            "awaiting_user_action": True,
            "guard_mode": mode,
            "blocked_count": len(blocked),
            "blocked_cells": blocked[:10],
            "instruction": (
                "HARD STOP and ask the user for approval before retrying with allow_formula_overwrite=true. "
                "Briefly explain what you want to change and why."
            ),
        })

    async def write_live_excel_cell(
        self,
        cell_ref: str,
        value: Any,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        is_formula: bool = False,
        number_format: Optional[str] = None,
        recalculate: bool = True,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write a typed value or formula into a specific cell of a targeted live Excel workbook/sheet."""
        clean_cell = validate_cell_ref(cell_ref)
        guard_mode = "off" if allow_formula_overwrite else self.get_formula_guard_mode()

        # Determine explicit type tag so AppleScript preserves booleans, numbers, nulls, formulas, strings
        if value is None:
            val_type = "null"
            val_str = ""
        elif is_formula or (isinstance(value, str) and value.startswith("=")):
            val_type = "formula"
            val_str = value if str(value).startswith("=") else f"={value}"
            if "!" in val_str:
                try:
                    ws_info = await self.list_live_excel_worksheets(workbook_name=workbook_name)
                    val_str = self._sanitize_live_excel_formula(
                        val_str, ws_info.get("worksheets", [])
                    )
                except ValueError:
                    raise
                except Exception:
                    pass
        elif isinstance(value, bool):
            val_type = "bool"
            val_str = "true" if value else "false"
        elif isinstance(value, (int, float)):
            val_type = "number"
            val_str = str(value)
        else:
            val_type = "string"
            val_str = str(value)

        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set cellRef to item 3 of argv
            set valType to item 4 of argv
            set valStr to item 5 of argv
            set numFmt to item 6 of argv
            set doCalc to (item 7 of argv) is "true"
            set guardMode to item 8 of argv
            set us to character id 31
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                if shTarget is "" then
                    set sh to active sheet of wb
                else
                    set sh to worksheet shTarget of wb
                end if
                set prevAlerts to display alerts
                set display alerts to false
                set prevScreen to true
                set prevEvents to true
                try
                    set prevScreen to screen updating
                end try
                try
                    set prevEvents to enable events
                end try
                try
                    try
                        set screen updating to false
                    end try
                    try
                        set enable events to false
                    end try
                    set targetCell to range cellRef of sh
                    set fmlUpdated to 0
                    if guardMode is not "off" then
                        set existingFml to ""
                        try
                            set rawExisting to formula of targetCell
                            if rawExisting is not missing value then
                                set existingFml to rawExisting as text
                            end if
                        end try
                        if (length of existingFml) > 0 and (character 1 of existingFml is "=") then
                            if guardMode is "strict" or valType is not "formula" then
                                try
                                    set enable events to prevEvents
                                end try
                                try
                                    set screen updating to prevScreen
                                end try
                                set display alerts to prevAlerts
                                return "BLOCKED" & us & (name of wb) & us & (name of sh) & us & existingFml
                            else
                                set fmlUpdated to 1
                            end if
                        end if
                    end if
                    if valType is "null" then
                        clear contents targetCell
                    else if valType is "formula" then
                        set formula of targetCell to valStr
                    else if valType is "bool" then
                        if valStr is "true" then
                            set value of targetCell to true
                        else
                            set value of targetCell to false
                        end if
                    else if valType is "number" then
                        set value of targetCell to (valStr as number)
                    else
                        set value of targetCell to valStr
                    end if
                    if numFmt is not "" then
                        set number format of targetCell to numFmt
                    end if
                    if doCalc then
                        calculate
                    end if
                    set evalVal to value of targetCell
                    if evalVal is missing value then
                        set evalText to ""
                    else
                        set evalText to evalVal as text
                    end if
                    try
                        set enable events to prevEvents
                    end try
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                    return (name of wb) & us & (name of sh) & us & evalText & us & (fmlUpdated as text)
                on error errMsg number errNum
                    try
                        set enable events to prevEvents
                    end try
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[
                workbook_name or "",
                sheet_name or "",
                clean_cell,
                val_type,
                val_str,
                number_format or "",
                "true" if recalculate else "false",
                guard_mode,
            ],
            timeout=20.0,
        )
        parts = raw.split(_US)
        if parts and parts[0] == "BLOCKED":
            return self._formula_guard_response(
                guard_mode,
                [{
                    "sheet": parts[2] if len(parts) > 2 else (sheet_name or ""),
                    "cell": clean_cell,
                    "existing_formula": parts[3] if len(parts) > 3 else "",
                    "proposed_value": value,
                }],
            )
        out: Dict[str, Any] = {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "sheet_name": parts[1] if len(parts) > 1 else "",
            "cell": clean_cell,
            "written_value": value,
            "evaluated_value": parts[2] if len(parts) > 2 else "",
            "number_format": number_format,
        }
        if len(parts) > 3 and parts[3] == "1":
            out["formulas_updated"] = 1
        return sanitize_for_json(out)

    async def write_live_excel_range(
        self,
        start_cell: str,
        data: List[List[Any]],
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        number_format: Optional[str] = None,
        recalculate: bool = True,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write a 2D array into a targeted live Excel workbook & worksheet."""
        if not data:
            raise ValueError("data must be a non-empty 2D list")
        clean_start = validate_cell_ref(start_cell)
        guard_mode = "off" if allow_formula_overwrite else self.get_formula_guard_mode()

        # If any cell contains a cross-sheet formula, fetch sheet names once to auto-quote & guard
        has_cross_sheet_formula = any(
            isinstance(v, str) and v.startswith("=") and "!" in v
            for row in data
            for v in row
        )
        live_sheets: List[str] = []
        if has_cross_sheet_formula:
            try:
                ws_info = await self.list_live_excel_worksheets(workbook_name=workbook_name)
                live_sheets = list(ws_info.get("worksheets", []))
            except Exception:
                live_sheets = []

        # Encode rows with type tags: <type>:<value> separated by _US / _RS
        encoded_rows: List[str] = []
        max_cols = 0
        for row in data:
            max_cols = max(max_cols, len(row))
            encoded_cells: List[str] = []
            for cell_val in row:
                if cell_val is None:
                    encoded_cells.append("N:")
                elif isinstance(cell_val, bool):
                    encoded_cells.append(f"B:{'true' if cell_val else 'false'}")
                elif isinstance(cell_val, (int, float)):
                    encoded_cells.append(f"D:{cell_val}")
                elif isinstance(cell_val, str) and cell_val.startswith("="):
                    safe_fml = (
                        self._sanitize_live_excel_formula(cell_val, live_sheets)
                        if "!" in cell_val and live_sheets
                        else cell_val
                    )
                    encoded_cells.append(f"F:{safe_fml}")
                else:
                    encoded_cells.append(f"S:{cell_val}")
            encoded_rows.append(_US.join(encoded_cells))
        payload = _RS.join(encoded_rows)

        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set startRef to item 3 of argv
            set payload to item 4 of argv
            set numFmt to item 5 of argv
            set doCalc to (item 6 of argv) is "true"
            set guardMode to item 7 of argv
            set us to character id 31
            set rs to character id 30
            set gs to character id 29

            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set rowList to text items of payload
            set AppleScript's text item delimiters to prevDelims

            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                if shTarget is "" then
                    set sh to active sheet of wb
                else
                    set sh to worksheet shTarget of wb
                end if

                set anchorCell to range startRef of sh
                set startR to first row index of anchorCell
                set startC to first column index of anchorCell

                set prevAlerts to display alerts
                set display alerts to false
                set prevScreen to true
                set prevEvents to true
                try
                    set prevScreen to screen updating
                end try
                try
                    set prevEvents to enable events
                end try
                try
                    try
                        set screen updating to false
                    end try
                    try
                        set enable events to false
                    end try

                    set fmlUpdated to 0
                    if guardMode is not "off" then
                        set blockedRows to {}
                        set rIdx to 0
                        repeat with rStr in rowList
                            set AppleScript's text item delimiters to us
                            set cellList to text items of (rStr as text)
                            set AppleScript's text item delimiters to prevDelims
                            set cIdx to 0
                            repeat with entry in cellList
                                set entryStr to entry as text
                                set tagChar to ""
                                set rawVal to ""
                                if (length of entryStr) >= 2 then
                                    set tagChar to character 1 of entryStr
                                    if (length of entryStr) > 2 then
                                        set rawVal to text 3 thru -1 of entryStr
                                    end if
                                end if
                                set targetCell to cell (startC + cIdx) of row (startR + rIdx) of sh
                                set existingFml to ""
                                try
                                    set rawEx to formula of targetCell
                                    if rawEx is not missing value then
                                        set existingFml to rawEx as text
                                    end if
                                end try
                                if (length of existingFml) > 0 and (character 1 of existingFml is "=") then
                                    if guardMode is "strict" or tagChar is not "F" then
                                        set cellAddr to get address of targetCell without row absolute and column absolute
                                        set end of blockedRows to ((name of sh) & us & cellAddr & us & existingFml & us & rawVal)
                                    else
                                        set fmlUpdated to fmlUpdated + 1
                                    end if
                                end if
                                set cIdx to cIdx + 1
                            end repeat
                            set rIdx to rIdx + 1
                        end repeat
                        if (count of blockedRows) > 0 then
                            try
                                set enable events to prevEvents
                            end try
                            try
                                set screen updating to prevScreen
                            end try
                            set display alerts to prevAlerts
                            set AppleScript's text item delimiters to rs
                            set joinedBlocked to blockedRows as text
                            set AppleScript's text item delimiters to prevDelims
                            return "BLOCKED" & gs & joinedBlocked
                        end if
                    end if

                    set rIdx to 0
                    repeat with rStr in rowList
                        set AppleScript's text item delimiters to us
                        set cellList to text items of (rStr as text)
                        set AppleScript's text item delimiters to prevDelims
                        set cIdx to 0
                        repeat with entry in cellList
                            set entryStr to entry as text
                            set targetCell to cell (startC + cIdx) of row (startR + rIdx) of sh
                            if (length of entryStr) >= 2 then
                                set tagChar to character 1 of entryStr
                                if (length of entryStr) > 2 then
                                    set rawVal to text 3 thru -1 of entryStr
                                else
                                    set rawVal to ""
                                end if
                                if tagChar is "N" then
                                    clear contents targetCell
                                else if tagChar is "B" then
                                    if rawVal is "true" then
                                        set value of targetCell to true
                                    else
                                        set value of targetCell to false
                                    end if
                                else if tagChar is "D" then
                                    set value of targetCell to (rawVal as number)
                                else if tagChar is "F" then
                                    set formula of targetCell to rawVal
                                else
                                    set value of targetCell to rawVal
                                end if
                                if numFmt is not "" then
                                    set number format of targetCell to numFmt
                                end if
                            end if
                            set cIdx to cIdx + 1
                        end repeat
                        set rIdx to rIdx + 1
                    end repeat

                    if doCalc then
                        calculate
                    end if
                    try
                        set enable events to prevEvents
                    end try
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                    return (name of wb) & us & (name of sh) & us & (fmlUpdated as text)
                on error errMsg number errNum
                    try
                        set enable events to prevEvents
                    end try
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[
                workbook_name or "",
                sheet_name or "",
                clean_start,
                payload,
                number_format or "",
                "true" if recalculate else "false",
                guard_mode,
            ],
            timeout=45.0,
        )
        if raw.startswith(f"BLOCKED{_GS}"):
            blocked_raw = raw.split(_GS, 1)[1]
            blocked_list: List[Dict[str, Any]] = []
            for row_s in blocked_raw.split(_RS):
                if not row_s:
                    continue
                p = row_s.split(_US)
                if len(p) >= 3:
                    blocked_list.append({
                        "sheet": p[0],
                        "cell": p[1],
                        "existing_formula": p[2],
                        "proposed_value": p[3] if len(p) > 3 else "",
                    })
            return self._formula_guard_response(guard_mode, blocked_list)
        parts = raw.split(_US)
        out: Dict[str, Any] = {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "sheet_name": parts[1] if len(parts) > 1 else "",
            "start_cell": clean_start,
            "rows_written": len(data),
            "columns_written": max_cols,
        }
        if len(parts) > 2 and parts[2].isdigit() and int(parts[2]) > 0:
            out["formulas_updated"] = int(parts[2])
        return out

    async def batch_write_live_excel_cells(
        self,
        cells: Any,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        recalculate: bool = True,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write multiple non-contiguous cells/formulas in ONE osascript call and recalculate once.

        `cells` may be either:
          - A dict mapping cell ref to value: `{'G2': '=...', 'G11': 'N/A'}`
          - A list of dicts: `[{'cell': 'G2', 'value': '=...', 'sheet_name': 'Income Statement', 'number_format': '0.0%'}]`
        """
        normalized_items: List[Dict[str, Any]] = []
        if isinstance(cells, dict):
            for cell_ref, val in cells.items():
                normalized_items.append({
                    "sheet_name": sheet_name or "",
                    "cell": validate_cell_ref(str(cell_ref)),
                    "value": val,
                    "number_format": "",
                    "is_formula": False,
                })
        elif isinstance(cells, list):
            for item in cells:
                if not isinstance(item, dict):
                    raise ValueError("Each item in cells list must be a dict with 'cell' and 'value'")
                raw_cell = item.get("cell") or item.get("cell_ref")
                if not raw_cell:
                    raise ValueError("Each cell update must include 'cell' (e.g. 'G2')")
                normalized_items.append({
                    "sheet_name": item.get("sheet_name") or sheet_name or "",
                    "cell": validate_cell_ref(str(raw_cell)),
                    "value": item.get("value"),
                    "number_format": str(item.get("number_format") or ""),
                    "is_formula": bool(item.get("is_formula", False)),
                })
        else:
            raise ValueError("cells must be a dict ({'G2': ...}) or non-empty list of dicts")

        if not normalized_items:
            raise ValueError("cells cannot be empty")

        # Sanitize any cross-sheet formulas once against the live workbook's sheet list
        has_cross_sheet = any(
            (it["is_formula"] or (isinstance(it["value"], str) and it["value"].startswith("=")))
            and "!" in str(it["value"] or "")
            for it in normalized_items
        )
        live_sheets: List[str] = []
        if has_cross_sheet:
            try:
                ws_info = await self.list_live_excel_worksheets(workbook_name=workbook_name)
                live_sheets = list(ws_info.get("worksheets", []))
            except Exception:
                live_sheets = []

        encoded_rows: List[str] = []
        for it in normalized_items:
            val = it["value"]
            if val is None:
                vtype, vstr = "N", ""
            elif it["is_formula"] or (isinstance(val, str) and val.startswith("=")):
                vtype = "F"
                vstr = val if str(val).startswith("=") else f"={val}"
                if "!" in vstr and live_sheets:
                    vstr = self._sanitize_live_excel_formula(vstr, live_sheets)
            elif isinstance(val, bool):
                vtype, vstr = "B", ("true" if val else "false")
            elif isinstance(val, (int, float)):
                vtype, vstr = "D", str(val)
            else:
                vtype, vstr = "S", str(val)
            it["_encoded_value"] = vstr
            encoded_rows.append(_US.join([it["sheet_name"], it["cell"], vtype, vstr, it["number_format"]]))

        guard_mode = "off" if allow_formula_overwrite else self.get_formula_guard_mode()
        payload = _RS.join(encoded_rows)
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set payload to item 2 of argv
            set doCalc to (item 3 of argv) is "true"
            set guardMode to item 4 of argv
            set us to character id 31
            set rs to character id 30
            set gs to character id 29

            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set rowList to text items of payload
            set AppleScript's text item delimiters to prevDelims

            set targetRefs to {}
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set wbName to name of wb
                set prevAlerts to display alerts
                set display alerts to false
                set prevScreen to true
                set prevEvents to true
                try
                    set prevScreen to screen updating
                end try
                try
                    set prevEvents to enable events
                end try
                try
                    try
                        set screen updating to false
                    end try
                    try
                        set enable events to false
                    end try
                    set fmlUpdated to 0
                    if guardMode is not "off" then
                        set blockedRows to {}
                        repeat with rStr in rowList
                            set AppleScript's text item delimiters to us
                            set parts to text items of (rStr as text)
                            set AppleScript's text item delimiters to prevDelims
                            if (count of parts) >= 5 then
                                set shName to item 1 of parts
                                set cellRef to item 2 of parts
                                set tagChar to item 3 of parts
                                set rawVal to item 4 of parts
                                if shName is "" then
                                    set sh to active sheet of wb
                                else
                                    set sh to worksheet shName of wb
                                end if
                                set targetCell to range cellRef of sh
                                set existingFml to ""
                                try
                                    set rawEx to formula of targetCell
                                    if rawEx is not missing value then
                                        set existingFml to rawEx as text
                                    end if
                                end try
                                if (length of existingFml) > 0 and (character 1 of existingFml is "=") then
                                    if guardMode is "strict" or tagChar is not "F" then
                                        set end of blockedRows to ((name of sh) & us & cellRef & us & existingFml & us & rawVal)
                                    else
                                        set fmlUpdated to fmlUpdated + 1
                                    end if
                                end if
                            end if
                        end repeat
                        if (count of blockedRows) > 0 then
                            try
                                set enable events to prevEvents
                            end try
                            try
                                set screen updating to prevScreen
                            end try
                            set display alerts to prevAlerts
                            set AppleScript's text item delimiters to rs
                            set joinedBlocked to blockedRows as text
                            set AppleScript's text item delimiters to prevDelims
                            return "BLOCKED" & gs & joinedBlocked
                        end if
                    end if

                    repeat with rStr in rowList
                        set AppleScript's text item delimiters to us
                        set parts to text items of (rStr as text)
                        set AppleScript's text item delimiters to prevDelims
                        if (count of parts) >= 5 then
                            set shName to item 1 of parts
                            set cellRef to item 2 of parts
                            set tagChar to item 3 of parts
                            set rawVal to item 4 of parts
                            set numFmt to item 5 of parts
                            if shName is "" then
                                set sh to active sheet of wb
                            else
                                set sh to worksheet shName of wb
                            end if
                            set actualSh to name of sh
                            set targetCell to range cellRef of sh
                            if tagChar is "N" then
                                clear contents targetCell
                            else if tagChar is "F" then
                                set formula of targetCell to rawVal
                            else if tagChar is "B" then
                                if rawVal is "true" then
                                    set value of targetCell to true
                                else
                                    set value of targetCell to false
                                end if
                            else if tagChar is "D" then
                                set value of targetCell to (rawVal as number)
                            else
                                set value of targetCell to rawVal
                            end if
                            if numFmt is not "" then
                                set number format of targetCell to numFmt
                            end if
                            set end of targetRefs to {actualSh, cellRef, targetCell}
                        end if
                    end repeat

                    if doCalc then
                        calculate
                    end if

                    set outRows to {}
                    repeat with entry in targetRefs
                        set actualSh to item 1 of entry
                        set cellRef to item 2 of entry
                        set cObj to item 3 of entry
                        set evalVal to value of cObj
                        if evalVal is missing value then
                            set evalText to ""
                        else
                            set evalText to evalVal as text
                        end if
                        set end of outRows to (actualSh & us & cellRef & us & evalText)
                    end repeat

                    try
                        set enable events to prevEvents
                    end try
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                on error errMsg number errNum
                    try
                        set enable events to prevEvents
                    end try
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell

            set AppleScript's text item delimiters to rs
            set joinedRows to outRows as text
            set AppleScript's text item delimiters to prevDelims
            return wbName & gs & joinedRows & gs & (fmlUpdated as text)
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", payload, "true" if recalculate else "false", guard_mode],
            timeout=45.0,
        )
        if raw.startswith(f"BLOCKED{_GS}"):
            blocked_raw = raw.split(_GS, 1)[1]
            blocked_list: List[Dict[str, Any]] = []
            for row_s in blocked_raw.split(_RS):
                if not row_s:
                    continue
                p = row_s.split(_US)
                if len(p) >= 3:
                    blocked_list.append({
                        "sheet": p[0],
                        "cell": p[1],
                        "existing_formula": p[2],
                        "proposed_value": p[3] if len(p) > 3 else "",
                    })
            return self._formula_guard_response(guard_mode, blocked_list)
        sections = raw.split(_GS)
        wb_name = sections[0] if sections else (workbook_name or "")
        rows_raw = sections[1] if len(sections) > 1 else ""
        fml_updated_str = sections[2] if len(sections) > 2 else "0"

        results: List[Dict[str, Any]] = []
        if rows_raw:
            for row in rows_raw.split(_RS):
                if not row:
                    continue
                parts = row.split(_US)
                if len(parts) >= 3:
                    results.append({
                        "sheet_name": parts[0],
                        "cell": parts[1],
                        "evaluated_value": parts[2],
                    })

        out: Dict[str, Any] = {
            "status": "success",
            "workbook_name": wb_name,
            "cells_written": len(results),
            "recalculated": recalculate,
            "results": results,
        }
        if fml_updated_str.isdigit() and int(fml_updated_str) > 0:
            out["formulas_updated"] = int(fml_updated_str)
        return sanitize_for_json(out)

    async def batch_read_live_excel_ranges(
        self,
        ranges: List[Dict[str, Any]],
        workbook_name: Optional[str] = None,
        include_formulas: bool = False,
        include_display_values: Any = "auto",
    ) -> Dict[str, Any]:
        """Read multiple ranges across one or more live Excel sheets in ONE osascript call.

        Each item in `ranges` should be `{'sheet_name': 'Income Statement', 'range_ref': 'B2:F16'}`
        (`range` and `cell_range` are also accepted as aliases for `range_ref`).
        """
        if not ranges or not isinstance(ranges, list):
            raise ValueError("ranges must be a non-empty list of {'sheet_name', 'range_ref'} dicts")

        req_rows: List[str] = []
        cleaned_specs: List[Dict[str, str]] = []
        for item in ranges:
            if not isinstance(item, dict):
                raise ValueError("Each range spec must be a dict with 'range_ref'")
            raw_ref = (
                item.get("range_ref")
                or item.get("range")
                or item.get("cell_range")
                or item.get("cell_ref")
                or item.get("cell")
            )
            if not raw_ref:
                raise ValueError("Each range spec must include 'range_ref' (e.g. 'B2:F16')")
            clean_ref = validate_range_ref(str(raw_ref))
            sh_name = str(
                item.get("sheet_name")
                or item.get("sheet")
                or item.get("worksheet")
                or item.get("worksheet_name")
                or item.get("tab")
                or ""
            )
            cleaned_specs.append({"sheet_name": sh_name, "range_ref": clean_ref})
            req_rows.append(_US.join([sh_name, clean_ref]))

        payload = _RS.join(req_rows)
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set payload to item 2 of argv
            set incFormulas to (item 3 of argv) is "true"
            set us to character id 31
            set rs to character id 30
            set gs to character id 29
            set fs to character id 28

            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set reqList to text items of payload
            set AppleScript's text item delimiters to prevDelims

            set rawResults to {}
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set wbName to name of wb
                try
                    calculate
                end try
                repeat with reqStr in reqList
                    set AppleScript's text item delimiters to us
                    set parts to text items of (reqStr as text)
                    set AppleScript's text item delimiters to prevDelims
                    if (count of parts) >= 2 then
                        set shName to item 1 of parts
                        set refStr to item 2 of parts
                        if shName is "" then
                            set sh to active sheet of wb
                        else
                            set sh to worksheet shName of wb
                        end if
                        set actualSh to name of sh
                        set rng to range refStr of sh
                        set rVals to value of rng
                        set rStrs to missing value
                        try
                            set rStrs to string value of rng
                        end try
                        set rFmls to missing value
                        if incFormulas then
                            set rFmls to formula of rng
                        end if
                        set end of rawResults to {actualSh, refStr, rVals, rFmls, rStrs}
                    end if
                end repeat
            end tell

            set outBlocks to {}
            repeat with itemRec in rawResults
                set actualSh to item 1 of itemRec
                set refStr to item 2 of itemRec
                set rVals to item 3 of itemRec
                set rFmls to item 4 of itemRec
                set rStrs to item 5 of itemRec
                set valBlock to my matrix_to_text(rVals, us, rs)
                set fmlBlock to ""
                if incFormulas then
                    set fmlBlock to my matrix_to_text(rFmls, us, rs)
                end if
                set strBlock to my matrix_to_text(rStrs, us, rs)
                set end of outBlocks to (actualSh & us & refStr & gs & valBlock & gs & fmlBlock & gs & strBlock)
            end repeat

            set AppleScript's text item delimiters to fs
            set joinedBlocks to outBlocks as text
            set AppleScript's text item delimiters to prevDelims
            return wbName & fs & joinedBlocks
        end run

        on matrix_to_text(rawData, us, rs)
            if rawData is missing value then
                return ""
            end if
            if class of rawData is not list then
                return (rawData as text)
            end if
            if (count of rawData) is 0 then
                return ""
            end if
            if class of (item 1 of rawData) is not list then
                set matrixRows to {rawData}
            else
                set matrixRows to rawData
            end if
            set outRows to {}
            repeat with rowItem in matrixRows
                set rowList to contents of rowItem
                if class of rowList is not list then
                    set rowList to {rowList}
                end if
                set cellTexts to {}
                repeat with cellItem in rowList
                    set v to contents of cellItem
                    if v is missing value then
                        set end of cellTexts to ""
                    else
                        set end of cellTexts to (v as text)
                    end if
                end repeat
                set prevDelims to AppleScript's text item delimiters
                set AppleScript's text item delimiters to us
                set end of outRows to (cellTexts as text)
                set AppleScript's text item delimiters to prevDelims
            end repeat
            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set joined to outRows as text
            set AppleScript's text item delimiters to prevDelims
            return joined
        end matrix_to_text
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", payload, "true" if include_formulas else "false"],
            timeout=45.0,
        )
        blocks = raw.split(_FS)
        wb_name = blocks[0] if blocks else (workbook_name or "")
        results: List[Dict[str, Any]] = []
        all_clipped: List[str] = []
        for blk in blocks[1:]:
            if not blk:
                continue
            secs = blk.split(_GS)
            hdr = secs[0].split(_US) if secs else ["", ""]
            val_block = secs[1] if len(secs) > 1 else ""
            fml_block = secs[2] if len(secs) > 2 else ""
            str_block = secs[3] if len(secs) > 3 else ""
            values = [row.split(_US) for row in val_block.split(_RS)] if val_block else []
            range_str = hdr[1] if len(hdr) > 1 else ""
            display_values, clipped_cells = self._process_live_display_strings(
                values, str_block, range_str, include_display_values=include_display_values
            )
            values, display_values, grid_truncated, orig_rows = self._trim_and_cap_live_grid(
                values, display_values, max_cells=3000
            )
            entry: Dict[str, Any] = {
                "sheet_name": hdr[0],
                "range": range_str,
                "rows": len(values),
                "columns": len(values[0]) if values else 0,
                "values": values,
            }
            if grid_truncated:
                entry["truncated"] = True
                entry["total_populated_rows"] = orig_rows
            if display_values is not None:
                entry["display_values"] = display_values
            if clipped_cells:
                entry["clipped_cells"] = clipped_cells
                all_clipped.extend(f"{hdr[0]}!{c}" for c in clipped_cells)
            if include_formulas:
                entry["formulas"] = self._extract_sparse_formulas(fml_block, range_str)
            results.append(entry)

        out: Dict[str, Any] = {
            "status": "success",
            "workbook_name": wb_name,
            "ranges_read": len(results),
            "results": results,
        }
        if all_clipped:
            out["clipped_cells"] = all_clipped[:15]
            out["warning"] = (
                f"Column width clipping ('###') detected in live Excel at {', '.join(all_clipped[:10])}; "
                "widen columns via batch_format_workbook(auto_fit_columns=true) and re-save."
            )
        return sanitize_for_json(out)
