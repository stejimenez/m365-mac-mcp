"""
Excel Controller for Microsoft 365 for Mac MCP (m365-mac-mcp).

Provides full, non-blocking file-level read-write operations for Microsoft Excel
(.xlsx, .xlsm with VBA preservation, and .csv) using openpyxl + csv, plus optional
live Excel integration via AppleScriptBridge.
"""

import asyncio
import atexit
import csv
import os
import re
import sys
import tempfile
import time
import uuid
from copy import copy
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from integrations.applescript_bridge import AppleScriptBridge
from utils.logger import setup_logger
from utils.validators import (
    resolve_file_path,
    sanitize_excel_write_value,
    sanitize_for_json,
    validate_cell_ref,
    validate_range_ref,
)

logger = setup_logger(__name__)

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


class ExcelController:
    """Controller for Excel file-level read-write operations (.xlsx, .xlsm, .csv)."""

    def __init__(self, max_workbooks: int = 20):
        self.applescript = AppleScriptBridge()
        self.active_workbooks: Dict[str, Dict[str, Any]] = {}
        self.max_workbooks = max_workbooks
        self.temp_dir = Path(tempfile.gettempdir()) / "office365_mcp" / "excel"
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self._owned_temp_files: set[Path] = set()
        atexit.register(self._cleanup_temp_files)

    def _cleanup_temp_files(self) -> None:
        """Close open workbooks and remove controller-owned temporary files on shutdown."""
        for entry in list(self.active_workbooks.values()):
            try:
                entry["workbook_object"].close()
            except Exception:
                pass
        for path in list(self._owned_temp_files):
            try:
                if path.exists():
                    path.unlink()
            except Exception:
                pass
            self._owned_temp_files.discard(path)

    def _evict_if_needed(self) -> None:
        """Keep active workbook count bounded to prevent unbounded memory growth."""
        while len(self.active_workbooks) >= self.max_workbooks:
            oldest_id = next(iter(self.active_workbooks))
            entry = self.active_workbooks.pop(oldest_id)
            try:
                entry["workbook_object"].close()
            except Exception:
                pass
            temp_file = entry.get("temp_file")
            if temp_file and temp_file in self._owned_temp_files:
                try:
                    if temp_file.exists():
                        temp_file.unlink()
                except Exception:
                    pass
                self._owned_temp_files.discard(temp_file)
            logger.info(
                f"Evicted oldest in-memory workbook {oldest_id} to stay within limit ({self.max_workbooks})"
            )

    def _get_entry(self, workbook_id: str) -> Dict[str, Any]:
        if workbook_id not in self.active_workbooks:
            raise ValueError(f"Workbook '{workbook_id}' not found. Open or create a workbook first.")
        return self.active_workbooks[workbook_id]

    def _get_sheet(self, entry: Dict[str, Any], sheet_name: Optional[str]):
        wb = entry["workbook_object"]
        if not sheet_name:
            return wb.active
        if sheet_name not in wb.sheetnames:
            raise ValueError(
                f"Worksheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}"
            )
        return wb[sheet_name]

    async def _save_working_copy(self, entry: Dict[str, Any]) -> None:
        """Persist working copy asynchronously without blocking the event loop."""
        wb = entry["workbook_object"]
        target = entry["working_path"]
        await asyncio.to_thread(wb.save, str(target))

    def _sync_metadata(self, entry: Dict[str, Any]) -> None:
        wb = entry["workbook_object"]
        entry["metadata"]["worksheet_count"] = len(wb.worksheets)
        entry["metadata"]["worksheets"] = list(wb.sheetnames)
        entry["metadata"]["active_sheet"] = wb.active.title if wb.active else None

    # ------------------------------------------------------------------
    # Create / Open / Close
    # ------------------------------------------------------------------

    async def create_workbook(
        self,
        title: str = "New Workbook",
        template_path: Optional[str] = None,
        sheet_name: str = "Sheet1",
        include_title_header: bool = False,
        open_in_excel: bool = False,
    ) -> Dict[str, Any]:
        """Create a new Excel workbook (.xlsx)."""
        self._evict_if_needed()
        workbook_id = str(uuid.uuid4())
        keep_vba = False

        if template_path:
            tpl = resolve_file_path(
                template_path, must_exist=True, extensions=[".xlsx", ".xlsm", ".xltx"]
            )
            keep_vba = tpl.suffix.lower() == ".xlsm"
            wb = await asyncio.to_thread(load_workbook, str(tpl), keep_vba=keep_vba)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = sheet_name or "Sheet1"
            if include_title_header and title:
                ws["A1"] = title
                ws["A1"].font = Font(bold=True, size=14)

        ext = ".xlsm" if keep_vba else ".xlsx"
        temp_file = self.temp_dir / f"{workbook_id}{ext}"
        await asyncio.to_thread(wb.save, str(temp_file))
        self._owned_temp_files.add(temp_file)

        opened_in_app = False
        open_note: Optional[str] = None
        if open_in_excel:
            # Opening unnamed /tmp UUID files in sandboxed Microsoft Excel triggers macOS
            # "Grant Access" dialogs. Keep the temp file as working fallback and instruct the
            # model to save into the chat's configured workspace folder before opening live.
            open_note = (
                "Ignored open_in_excel=True on unnamed temporary file to avoid macOS Excel "
                "sandbox 'Grant Access' prompts. Save the workbook first with save_workbook into "
                "the chat's configured workspace folder (or ~/Desktop / ~/Documents), then call "
                "excel_live_open."
            )
            logger.info(open_note)

        metadata = {
            "workbook_id": workbook_id,
            "title": title,
            "file_path": str(temp_file),
            "source_path": None,
            "is_temporary": True,
            "keep_vba": keep_vba,
            "worksheet_count": len(wb.worksheets),
            "worksheets": list(wb.sheetnames),
            "active_sheet": wb.active.title,
            "opened_in_excel": opened_in_app,
            "created_at": time.time(),
        }
        if open_note:
            metadata["note"] = open_note

        self.active_workbooks[workbook_id] = {
            "metadata": metadata,
            "workbook_object": wb,
            "working_path": temp_file,
            "temp_file": temp_file,
        }
        logger.info(f"Created workbook '{title}' ({workbook_id})")
        return sanitize_for_json(metadata)

    async def open_workbook(self, file_path: str) -> Dict[str, Any]:
        """Open an existing Excel (.xlsx, .xlsm) or CSV (.csv) file."""
        self._evict_if_needed()
        resolved = resolve_file_path(
            file_path, must_exist=True, extensions=[".xlsx", ".xlsm", ".csv"]
        )
        workbook_id = str(uuid.uuid4())
        suffix = resolved.suffix.lower()
        temp_file: Optional[Path] = None
        keep_vba = suffix == ".xlsm"

        if suffix == ".csv":
            wb = Workbook()
            ws = wb.active
            ws.title = resolved.stem[:31] or "Sheet1"

            def _load_csv() -> None:
                with open(resolved, "r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.reader(f)
                    for row in reader:
                        parsed_row: List[Any] = []
                        for val in row:
                            parsed_row.append(self._parse_csv_scalar(val))
                        ws.append(parsed_row)

            await asyncio.to_thread(_load_csv)
            temp_file = self.temp_dir / f"{workbook_id}.xlsx"
            await asyncio.to_thread(wb.save, str(temp_file))
            self._owned_temp_files.add(temp_file)
            working_path = temp_file
        else:
            wb = await asyncio.to_thread(load_workbook, str(resolved), keep_vba=keep_vba)
            working_path = resolved

        metadata = {
            "workbook_id": workbook_id,
            "title": resolved.stem,
            "file_path": str(resolved),
            "source_path": str(resolved),
            "source_format": suffix.lstrip("."),
            "is_temporary": suffix == ".csv",
            "keep_vba": keep_vba,
            "worksheet_count": len(wb.worksheets),
            "worksheets": list(wb.sheetnames),
            "active_sheet": wb.active.title if wb.active else None,
            "opened_at": time.time(),
        }

        self.active_workbooks[workbook_id] = {
            "metadata": metadata,
            "workbook_object": wb,
            "working_path": working_path,
            "temp_file": temp_file,
        }
        logger.info(f"Opened workbook '{resolved}' ({workbook_id})")
        return sanitize_for_json(metadata)

    @staticmethod
    def _parse_csv_scalar(val: str) -> Any:
        """Convert plain CSV strings to int/float/bool cleanly without corrupting leading zeros."""
        if val == "":
            return None
        low = val.lower()
        if low == "true":
            return True
        if low == "false":
            return False
        # Preserve strings with leading zeros (e.g., ZIP codes "01234")
        if len(val) > 1 and val[0] == "0" and val[1].isdigit():
            return val
        try:
            return int(val)
        except ValueError:
            pass
        try:
            f = float(val)
            return sanitize_excel_write_value(f)
        except ValueError:
            return val

    async def close_workbook(
        self,
        workbook_id: str,
        save_before_close: bool = False,
    ) -> Dict[str, Any]:
        """Close an in-memory workbook, release resources, and remove temp files."""
        entry = self._get_entry(workbook_id)
        saved_path = None
        if save_before_close and not entry["metadata"].get("is_temporary", True):
            await self._save_working_copy(entry)
            saved_path = str(entry["working_path"])

        self.active_workbooks.pop(workbook_id, None)
        try:
            entry["workbook_object"].close()
        except Exception:
            pass

        temp_file = entry.get("temp_file")
        if temp_file and temp_file in self._owned_temp_files:
            try:
                if temp_file.exists():
                    temp_file.unlink()
            except Exception:
                pass
            self._owned_temp_files.discard(temp_file)

        return {
            "status": "success",
            "workbook_id": workbook_id,
            "saved_before_close": bool(saved_path),
            "saved_path": saved_path,
        }

    # ------------------------------------------------------------------
    # Worksheet Management
    # ------------------------------------------------------------------

    async def list_worksheets(self, workbook_id: str) -> List[str]:
        entry = self._get_entry(workbook_id)
        return list(entry["workbook_object"].sheetnames)

    async def add_worksheet(
        self,
        workbook_id: str,
        sheet_name: str,
        position: Optional[int] = None,
    ) -> Dict[str, Any]:
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        if sheet_name in wb.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' already exists")

        ws = wb.create_sheet(sheet_name, position) if position is not None else wb.create_sheet(sheet_name)
        self._sync_metadata(entry)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "sheet_index": wb.worksheets.index(ws),
            "worksheets": list(wb.sheetnames),
        }

    async def rename_worksheet(
        self,
        workbook_id: str,
        old_name: str,
        new_name: str,
    ) -> Dict[str, Any]:
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        if old_name not in wb.sheetnames:
            raise ValueError(f"Worksheet '{old_name}' not found")
        if new_name in wb.sheetnames:
            raise ValueError(f"Worksheet '{new_name}' already exists")

        ws = wb[old_name]
        ws.title = new_name
        self._sync_metadata(entry)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "workbook_id": workbook_id,
            "old_name": old_name,
            "new_name": ws.title,
            "worksheets": list(wb.sheetnames),
        }

    async def delete_worksheet(self, workbook_id: str, sheet_name: str) -> Dict[str, Any]:
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' not found")
        if len(wb.worksheets) <= 1:
            raise ValueError("Cannot delete the only remaining worksheet in a workbook")

        del wb[sheet_name]
        self._sync_metadata(entry)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "workbook_id": workbook_id,
            "deleted_sheet": sheet_name,
            "worksheets": list(wb.sheetnames),
            "active_sheet": wb.active.title,
        }

    async def set_active_worksheet(self, workbook_id: str, sheet_name: str) -> Dict[str, Any]:
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' not found")
        wb.active = wb.sheetnames.index(sheet_name)
        self._sync_metadata(entry)
        await self._save_working_copy(entry)
        return {
            "status": "success",
            "workbook_id": workbook_id,
            "active_sheet": wb.active.title,
        }

    # ------------------------------------------------------------------
    # Reading (Cells, Ranges, Formulas vs Evaluated Values, Tables)
    # ------------------------------------------------------------------

    def _load_cached_evaluated_sheet(self, entry: Dict[str, Any], sheet_title: str):
        """Load cached formula values from disk via openpyxl(data_only=True) if available."""
        working_path: Path = entry["working_path"]
        if not working_path.exists() or working_path.suffix.lower() not in (".xlsx", ".xlsm"):
            return None, None
        wb_eval = load_workbook(str(working_path), data_only=True, read_only=True)
        if sheet_title not in wb_eval.sheetnames:
            wb_eval.close()
            return None, None
        return wb_eval, wb_eval[sheet_title]

    async def read_cell(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        cell: str,
        include_evaluated: bool = True,
    ) -> Dict[str, Any]:
        """Read a single cell's value, formula, cached evaluated value, and number format."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)
        clean_cell = validate_cell_ref(cell)

        cell_obj = ws[clean_cell]
        raw_val = cell_obj.value
        is_formula = isinstance(raw_val, str) and raw_val.startswith("=")

        evaluated_val = raw_val
        if include_evaluated and is_formula:
            def _fetch_eval():
                wb_e, ws_e = self._load_cached_evaluated_sheet(entry, ws.title)
                if ws_e is None:
                    return None
                try:
                    return ws_e[clean_cell].value
                finally:
                    wb_e.close()
            evaluated_val = await asyncio.to_thread(_fetch_eval)

        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "cell": clean_cell,
            "value": raw_val,
            "is_formula": is_formula,
            "formula": raw_val if is_formula else None,
            "evaluated_value": evaluated_val,
            "number_format": cell_obj.number_format,
            "data_type": cell_obj.data_type,
        })

    async def read_range(
        self,
        workbook_id: str,
        sheet_name: Optional[str] = None,
        cell_range: Optional[str] = None,
        evaluate_formulas: bool = False,
        max_rows: int = 500,
    ) -> Dict[str, Any]:
        """Read a cell range (or used range) as a JSON-safe 2D matrix."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)

        if cell_range:
            clean_ref = validate_range_ref(cell_range)
            if ":" in clean_ref:
                min_col, min_row, max_col, max_row = range_boundaries(clean_ref)
            else:
                col_l, row_i = coordinate_from_string(clean_ref)
                min_col = max_col = column_index_from_string(col_l)
                min_row = max_row = row_i
        else:
            min_col = ws.min_column or 1
            min_row = ws.min_row or 1
            max_col = ws.max_column or 1
            max_row = ws.max_row or 1

        truncated = False
        if (max_row - min_row + 1) > max_rows:
            max_row = min_row + max_rows - 1
            truncated = True

        def _extract() -> List[List[Any]]:
            if evaluate_formulas:
                wb_e, ws_e = self._load_cached_evaluated_sheet(entry, ws.title)
                if ws_e is not None:
                    try:
                        return [
                            list(row)
                            for row in ws_e.iter_rows(
                                min_row=min_row,
                                max_row=max_row,
                                min_col=min_col,
                                max_col=max_col,
                                values_only=True,
                            )
                        ]
                    finally:
                        wb_e.close()
            return [
                list(row)
                for row in ws.iter_rows(
                    min_row=min_row,
                    max_row=max_row,
                    min_col=min_col,
                    max_col=max_col,
                    values_only=True,
                )
            ]

        rows_data = await asyncio.to_thread(_extract)
        if not cell_range and rows_data:
            rows_data = self._trim_2d_grid(rows_data)
            if rows_data:
                max_row = min_row + len(rows_data) - 1
                max_col = min_col + len(rows_data[0]) - 1
        actual_range = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{max_row}"
        )

        payload: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "range": actual_range,
            "rows": len(rows_data),
            "columns": len(rows_data[0]) if rows_data else 0,
            "evaluate_formulas": evaluate_formulas,
            "data": rows_data,
        }
        if truncated:
            payload["truncated"] = True
        return sanitize_for_json(payload)

    @staticmethod
    def _trim_2d_grid(rows: List[List[Any]]) -> List[List[Any]]:
        """Trim trailing all-empty (None / '') rows and columns from a 2D cell matrix."""
        def _is_empty(v: Any) -> bool:
            return v is None or v == ""

        end_row = len(rows)
        while end_row > 0 and all(_is_empty(v) for v in rows[end_row - 1]):
            end_row -= 1
        trimmed = rows[:end_row]
        if not trimmed:
            return []

        max_used_col = 0
        for row in trimmed:
            for idx in range(len(row) - 1, -1, -1):
                if not _is_empty(row[idx]):
                    if idx + 1 > max_used_col:
                        max_used_col = idx + 1
                    break
        if max_used_col == 0:
            return []
        return [row[:max_used_col] for row in trimmed]

    async def read_tables(
        self,
        workbook_id: str,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """List and read all structured Excel Tables in a sheet or entire workbook."""
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        sheets = [self._get_sheet(entry, sheet_name)] if sheet_name else list(wb.worksheets)

        tables_out: List[Dict[str, Any]] = []
        for ws in sheets:
            for tbl in ws.tables.values():
                min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
                rows_data = [
                    list(row)
                    for row in ws.iter_rows(
                        min_row=min_row,
                        max_row=max_row,
                        min_col=min_col,
                        max_col=max_col,
                        values_only=True,
                    )
                ]
                headers = rows_data[0] if rows_data else []
                body = rows_data[1:] if len(rows_data) > 1 else []
                tables_out.append({
                    "sheet_name": ws.title,
                    "table_name": tbl.name,
                    "ref": tbl.ref,
                    "headers": headers,
                    "row_count": len(body),
                    "rows": body,
                })

        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "table_count": len(tables_out),
            "tables": tables_out,
        })

    async def inspect_workbook(
        self,
        workbook_id: Optional[str] = None,
        file_path: Optional[str] = None,
        max_rows_per_sheet: int = 50,
        max_cols_per_sheet: int = 20,
        include_formulas: bool = True,
        include_evaluated: bool = True,
    ) -> Dict[str, Any]:
        """Inspect all worksheets (structure, headers, formulas, cached values, tables) in ONE read-only call."""
        if not workbook_id and not file_path:
            raise ValueError("Provide either workbook_id or file_path to inspect_workbook")

        owned_wb = None
        if workbook_id:
            entry = self._get_entry(workbook_id)
            wb = entry["workbook_object"]
            disk_path: Optional[Path] = entry["working_path"]
        else:
            disk_path = resolve_file_path(file_path or "", must_exist=True, extensions=[".xlsx", ".xlsm"])
            wb = await asyncio.to_thread(load_workbook, str(disk_path), data_only=False)
            owned_wb = wb

        def _scan() -> Dict[str, Any]:
            wb_eval = None
            if include_evaluated and disk_path and disk_path.exists() and disk_path.suffix.lower() in (".xlsx", ".xlsm"):
                try:
                    wb_eval = load_workbook(str(disk_path), data_only=True, read_only=True)
                except Exception:
                    wb_eval = None
            try:
                sheets_out: List[Dict[str, Any]] = []
                for ws in wb.worksheets:
                    max_r = min(ws.max_row or 0, max_rows_per_sheet)
                    max_c = min(ws.max_column or 0, max_cols_per_sheet)
                    truncated_rows = (ws.max_row or 0) > max_rows_per_sheet
                    truncated_cols = (ws.max_column or 0) > max_cols_per_sheet

                    raw_grid: List[List[Any]] = []
                    if max_r > 0 and max_c > 0:
                        raw_grid = [
                            list(row)
                            for row in ws.iter_rows(
                                min_row=1, max_row=max_r, min_col=1, max_col=max_c, values_only=True
                            )
                        ]

                    evaluated_grid: Optional[List[List[Any]]] = None
                    if wb_eval is not None and ws.title in wb_eval.sheetnames and max_r > 0 and max_c > 0:
                        ws_e = wb_eval[ws.title]
                        evaluated_grid = [
                            list(row)
                            for row in ws_e.iter_rows(
                                min_row=1, max_row=max_r, min_col=1, max_col=max_c, values_only=True
                            )
                        ]

                    # Build ONE unified values grid + sparse A1->formula map (eliminates duplicating 2D matrices)
                    sparse_formulas: Dict[str, str] = {}
                    unified_values: List[List[Any]] = []
                    for r_idx, raw_row in enumerate(raw_grid, start=1):
                        eval_row = (
                            evaluated_grid[r_idx - 1]
                            if evaluated_grid is not None and r_idx - 1 < len(evaluated_grid)
                            else None
                        )
                        out_row: List[Any] = []
                        for c_idx, raw_val in enumerate(raw_row, start=1):
                            is_formula = isinstance(raw_val, str) and raw_val.startswith("=")
                            if is_formula:
                                if include_formulas:
                                    sparse_formulas[f"{get_column_letter(c_idx)}{r_idx}"] = raw_val
                                eval_val = (
                                    eval_row[c_idx - 1]
                                    if eval_row is not None and c_idx - 1 < len(eval_row)
                                    else None
                                )
                                out_row.append(eval_val if include_evaluated else None)
                            else:
                                out_row.append(raw_val)
                        unified_values.append(out_row)

                    unified_values = self._trim_2d_grid(unified_values)
                    actual_rows = len(unified_values)
                    actual_cols = len(unified_values[0]) if unified_values else 0

                    sheet_info: Dict[str, Any] = {
                        "sheet_name": ws.title,
                        "total_rows": ws.max_row or 0,
                        "total_columns": ws.max_column or 0,
                        "used_range": f"A1:{get_column_letter(max(actual_cols or ws.max_column or 1, 1))}{max(actual_rows or ws.max_row or 1, 1)}",
                    }
                    if ws.freeze_panes:
                        sheet_info["freeze_panes"] = ws.freeze_panes
                    tables_list = [{"name": t.name, "ref": t.ref} for t in ws.tables.values()]
                    if tables_list:
                        sheet_info["tables"] = tables_list
                    dv_count = len(ws.data_validations.dataValidation)
                    if dv_count:
                        sheet_info["data_validation_count"] = dv_count
                    if truncated_rows or truncated_cols:
                        sheet_info["truncated"] = True
                    sheet_info["values"] = unified_values
                    if include_formulas and sparse_formulas:
                        sheet_info["formulas"] = sparse_formulas
                    sheets_out.append(sheet_info)

                out: Dict[str, Any] = {
                    "status": "success",
                    "sheet_count": len(wb.worksheets),
                    "sheet_names": list(wb.sheetnames),
                    "sheets": sheets_out,
                }
                if workbook_id:
                    out["workbook_id"] = workbook_id
                if disk_path:
                    out["file_path"] = str(disk_path)
                ext_links = len(getattr(wb, "_external_links", []) or [])
                if ext_links:
                    out["external_links_count"] = ext_links
                return out
            finally:
                if wb_eval is not None:
                    wb_eval.close()

        try:
            return sanitize_for_json(await asyncio.to_thread(_scan))
        finally:
            if owned_wb is not None:
                owned_wb.close()

    async def audit_workbook_formulas(
        self,
        workbook_id: Optional[str] = None,
        file_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Audit every sheet in a workbook (read-only) for Excel errors, broken links, and suspicious literals."""
        if not workbook_id and not file_path:
            raise ValueError("Provide either workbook_id or file_path to audit_workbook_formulas")

        excel_error_tokens = {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!"}
        owned_wb = None
        if workbook_id:
            entry = self._get_entry(workbook_id)
            wb = entry["workbook_object"]
            disk_path: Optional[Path] = entry["working_path"]
        else:
            disk_path = resolve_file_path(file_path or "", must_exist=True, extensions=[".xlsx", ".xlsm"])
            wb = await asyncio.to_thread(load_workbook, str(disk_path), data_only=False)
            owned_wb = wb

        def _audit() -> Dict[str, Any]:
            wb_eval = None
            if disk_path and disk_path.exists() and disk_path.suffix.lower() in (".xlsx", ".xlsm"):
                try:
                    wb_eval = load_workbook(str(disk_path), data_only=True, read_only=True)
                except Exception:
                    wb_eval = None
            try:
                error_cells: List[Dict[str, Any]] = []
                sheet_ref_warnings: List[str] = []
                suspicious_literals: List[Dict[str, Any]] = []
                total_formulas = 0

                for ws in wb.worksheets:
                    ws_e = wb_eval[ws.title] if (wb_eval is not None and ws.title in wb_eval.sheetnames) else None
                    eval_rows = (
                        [list(r) for r in ws_e.iter_rows(values_only=True)]
                        if ws_e is not None
                        else []
                    )

                    max_r = ws.max_row or 0
                    max_c = ws.max_column or 0
                    col_formula_counts: Dict[int, int] = {}
                    col_nonempty_cells: Dict[int, List[Tuple[int, Any]]] = {}

                    for r in range(1, max_r + 1):
                        row_label = ws.cell(row=r, column=1).value
                        for c in range(1, max_c + 1):
                            val = ws.cell(row=r, column=c).value
                            coord = f"{get_column_letter(c)}{r}"
                            cached_val = (
                                eval_rows[r - 1][c - 1]
                                if (r - 1 < len(eval_rows) and c - 1 < len(eval_rows[r - 1]))
                                else None
                            )

                            is_fml = isinstance(val, str) and val.startswith("=")
                            if is_fml:
                                total_formulas += 1
                                col_formula_counts[c] = col_formula_counts.get(c, 0) + 1
                                sheet_ref_warnings.extend(
                                    self._inspect_formula_sheet_refs(
                                        wb, val, cell_ref=f"{ws.title}!{coord}"
                                    )
                                )
                            elif val is not None and r > 1:
                                col_nonempty_cells.setdefault(c, []).append((r, val))

                            # Detect Excel error tokens in either formula text or cached value
                            err_hit = None
                            if isinstance(val, str) and val.strip() in excel_error_tokens:
                                err_hit = val.strip()
                            elif isinstance(cached_val, str) and cached_val.strip() in excel_error_tokens:
                                err_hit = cached_val.strip()
                            if err_hit:
                                error_cells.append({
                                    "sheet_name": ws.title,
                                    "cell": coord,
                                    "row_label": row_label,
                                    "error": err_hit,
                                    "formula": val if is_fml else None,
                                    "cached_value": cached_val,
                                })

                            # Flag literal empty quotes ('""') anywhere
                            if isinstance(val, str) and val.strip() == '""':
                                suspicious_literals.append({
                                    "sheet_name": ws.title,
                                    "cell": coord,
                                    "row_label": row_label,
                                    "value": val,
                                    "reason": "Literal quote characters ('\"\"') instead of empty cell or formula",
                                })

                    # Flag hardcoded literals sitting inside predominantly-formula columns (>= 3 formulas)
                    for c, f_count in col_formula_counts.items():
                        if f_count >= 3:
                            for r, lit_val in col_nonempty_cells.get(c, []):
                                if isinstance(lit_val, str) and lit_val.strip() == '""':
                                    continue  # already flagged above
                                coord = f"{get_column_letter(c)}{r}"
                                suspicious_literals.append({
                                    "sheet_name": ws.title,
                                    "cell": coord,
                                    "row_label": ws.cell(row=r, column=1).value,
                                    "value": lit_val,
                                    "reason": f"Non-formula value in formula column {get_column_letter(c)} ({f_count} formulas in column)",
                                })

                ext_count = len(getattr(wb, "_external_links", []) or [])
                audit_out: Dict[str, Any] = {
                    "status": "success",
                    "sheet_count": len(wb.worksheets),
                    "sheet_names": list(wb.sheetnames),
                    "total_formulas": total_formulas,
                    "external_links_count": ext_count,
                    "error_count": len(error_cells),
                    "sheet_ref_warning_count": len(sheet_ref_warnings),
                    "suspicious_literal_count": len(suspicious_literals),
                    "clean": (
                        len(error_cells) == 0
                        and len(sheet_ref_warnings) == 0
                        and ext_count == 0
                    ),
                }
                if workbook_id:
                    audit_out["workbook_id"] = workbook_id
                if disk_path:
                    audit_out["file_path"] = str(disk_path)
                if error_cells:
                    audit_out["error_cells"] = error_cells
                if sheet_ref_warnings:
                    audit_out["sheet_ref_warnings"] = sheet_ref_warnings
                if suspicious_literals:
                    audit_out["suspicious_literals"] = suspicious_literals[:50]
                return audit_out
            finally:
                if wb_eval is not None:
                    wb_eval.close()

        try:
            return sanitize_for_json(await asyncio.to_thread(_audit))
        finally:
            if owned_wb is not None:
                owned_wb.close()

    # ------------------------------------------------------------------
    # Writing (Cells, Ranges, Formulas, Tables, Charts)
    # ------------------------------------------------------------------

    @staticmethod
    def get_formula_guard_mode(raw: Optional[str] = None) -> str:
        """Parse OFFICE_MCP_FORMULA_GUARD ('literals_only' default | 'strict' | 'off'), case-insensitive."""
        val = (
            os.environ.get("OFFICE_MCP_FORMULA_GUARD", "literals_only")
            if raw is None
            else raw
        ).strip().lower()
        if not val or val in ("literals_only", "literals", "literal", "default"):
            return "literals_only"
        if val in ("strict", "all", "on"):
            return "strict"
        if val in ("off", "false", "0", "no", "none", "disabled"):
            return "off"
        logger.warning(
            f"Unknown OFFICE_MCP_FORMULA_GUARD='{val}'. Falling back to 'literals_only'. "
            "Valid values: literals_only, strict, off."
        )
        return "literals_only"

    @staticmethod
    def _formula_guard_response(mode: str, blocked: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Return a compact, structured confirmation_required payload when formula writes are blocked."""
        return sanitize_for_json({
            "status": "confirmation_required",
            "guard_mode": mode,
            "blocked_count": len(blocked),
            "blocked_cells": blocked[:10],
            "instruction": (
                "STOP and ask the user for approval before retrying with allow_formula_overwrite=true. "
                "Briefly explain what you want to change and why."
            ),
        })

    async def write_cell(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        cell: str,
        value: Any,
        formatting: Optional[Dict[str, Any]] = None,
        number_format: Optional[str] = None,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write a properly typed value to a specific cell."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)
        clean_cell = validate_cell_ref(cell)

        wb = entry["workbook_object"]
        clean_value = sanitize_excel_write_value(value)
        is_new_formula = isinstance(clean_value, str) and clean_value.startswith("=")

        cell_obj = ws[clean_cell]
        existing_val = cell_obj.value
        existing_is_formula = isinstance(existing_val, str) and existing_val.startswith("=")
        guard_mode = self.get_formula_guard_mode()
        if existing_is_formula and not allow_formula_overwrite and guard_mode != "off":
            if guard_mode == "strict" or not is_new_formula:
                return self._formula_guard_response(
                    guard_mode,
                    [{
                        "sheet": ws.title,
                        "cell": clean_cell,
                        "existing_formula": existing_val,
                        "proposed_value": clean_value,
                    }],
                )

        formula_warnings: List[str] = []
        if is_new_formula:
            clean_value = self._normalize_formula_sheet_refs(wb, clean_value)
            formula_warnings = self._inspect_formula_sheet_refs(wb, clean_value, cell_ref=clean_cell)
        cell_obj.value = clean_value

        if number_format:
            cell_obj.number_format = number_format
        if formatting:
            self._apply_cell_formatting(cell_obj, formatting)

        await self._save_working_copy(entry)

        result: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "cell": clean_cell,
            "value": clean_value,
            "value_type": type(clean_value).__name__,
            "number_format": cell_obj.number_format,
        }
        if existing_is_formula:
            result["formulas_updated"] = 1
        if formula_warnings:
            result["formula_warnings"] = formula_warnings
        return sanitize_for_json(result)

    async def write_range(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        start_cell: str,
        data: List[List[Any]],
        formatting: Optional[Dict[str, Any]] = None,
        header_formatting: Optional[Dict[str, Any]] = None,
        number_format: Optional[str] = None,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write a 2D list of typed values to a worksheet range."""
        if not data:
            raise ValueError("data must be a non-empty 2D list")

        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)
        clean_start = validate_cell_ref(start_cell)

        col_letter, start_row = coordinate_from_string(clean_start)
        start_col = column_index_from_string(col_letter)
        max_cols = 0

        wb = entry["workbook_object"]

        # Atomic pre-write formula guard scan
        guard_mode = self.get_formula_guard_mode()
        if not allow_formula_overwrite and guard_mode != "off":
            blocked: List[Dict[str, Any]] = []
            for row_idx, row_data in enumerate(data):
                for col_idx, raw_val in enumerate(row_data):
                    existing = ws.cell(row=start_row + row_idx, column=start_col + col_idx).value
                    if isinstance(existing, str) and existing.startswith("="):
                        clean_val = sanitize_excel_write_value(raw_val)
                        new_is_fml = isinstance(clean_val, str) and clean_val.startswith("=")
                        if guard_mode == "strict" or not new_is_fml:
                            blocked.append({
                                "sheet": ws.title,
                                "cell": f"{get_column_letter(start_col + col_idx)}{start_row + row_idx}",
                                "existing_formula": existing,
                                "proposed_value": clean_val,
                            })
            if blocked:
                return self._formula_guard_response(guard_mode, blocked)

        formula_warnings: List[str] = []
        formulas_updated = 0
        for row_idx, row_data in enumerate(data):
            max_cols = max(max_cols, len(row_data))
            for col_idx, raw_val in enumerate(row_data):
                cell_obj = ws.cell(row=start_row + row_idx, column=start_col + col_idx)
                if isinstance(cell_obj.value, str) and cell_obj.value.startswith("="):
                    formulas_updated += 1
                clean_val = sanitize_excel_write_value(raw_val)
                if isinstance(clean_val, str) and clean_val.startswith("="):
                    clean_val = self._normalize_formula_sheet_refs(wb, clean_val)
                    formula_warnings.extend(
                        self._inspect_formula_sheet_refs(wb, clean_val, cell_ref=cell_obj.coordinate)
                    )
                cell_obj.value = clean_val
                if number_format and (row_idx > 0 or not header_formatting):
                    cell_obj.number_format = number_format
                if row_idx == 0 and header_formatting:
                    self._apply_cell_formatting(cell_obj, header_formatting)
                elif formatting:
                    self._apply_cell_formatting(cell_obj, formatting)

        await self._save_working_copy(entry)

        end_row = start_row + len(data) - 1
        end_col = start_col + max(max_cols, 1) - 1
        end_cell = f"{get_column_letter(end_col)}{end_row}"

        result: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "range": f"{clean_start}:{end_cell}",
            "rows_written": len(data),
            "columns_written": max_cols,
        }
        if formulas_updated:
            result["formulas_updated"] = formulas_updated
        if formula_warnings:
            result["formula_warnings"] = formula_warnings
        return sanitize_for_json(result)

    async def batch_write_ranges(
        self,
        workbook_id: str,
        operations: List[Dict[str, Any]],
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write multiple 2D ranges across one or more worksheets in a single call and save once."""
        if not operations:
            raise ValueError("operations must be a non-empty list of write specifications")

        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]

        # Atomic pre-write formula guard scan across ALL operations before mutating any cell
        guard_mode = self.get_formula_guard_mode()
        if not allow_formula_overwrite and guard_mode != "off":
            blocked: List[Dict[str, Any]] = []
            for op in operations:
                sheet_name = op.get("sheet_name")
                start_cell = op.get("start_cell", "A1")
                data = op.get("data")
                if not data:
                    raise ValueError(f"Each operation must include non-empty 'data' (sheet={sheet_name})")
                ws = self._get_sheet(entry, sheet_name)
                clean_start = validate_cell_ref(start_cell)
                col_letter, start_row = coordinate_from_string(clean_start)
                start_col = column_index_from_string(col_letter)
                for row_idx, row_data in enumerate(data):
                    for col_idx, raw_val in enumerate(row_data):
                        existing = ws.cell(row=start_row + row_idx, column=start_col + col_idx).value
                        if isinstance(existing, str) and existing.startswith("="):
                            clean_val = sanitize_excel_write_value(raw_val)
                            new_is_fml = isinstance(clean_val, str) and clean_val.startswith("=")
                            if guard_mode == "strict" or not new_is_fml:
                                blocked.append({
                                    "sheet": ws.title,
                                    "cell": f"{get_column_letter(start_col + col_idx)}{start_row + row_idx}",
                                    "existing_formula": existing,
                                    "proposed_value": clean_val,
                                })
            if blocked:
                return self._formula_guard_response(guard_mode, blocked)

        results: List[Dict[str, Any]] = []
        all_warnings: List[str] = []
        formulas_updated = 0

        for op in operations:
            sheet_name = op.get("sheet_name")
            start_cell = op.get("start_cell", "A1")
            data = op.get("data")
            if not data:
                raise ValueError(f"Each operation must include non-empty 'data' (sheet={sheet_name})")
            formatting = op.get("formatting")
            header_formatting = op.get("header_formatting")
            number_format = op.get("number_format")

            ws = self._get_sheet(entry, sheet_name)
            clean_start = validate_cell_ref(start_cell)
            col_letter, start_row = coordinate_from_string(clean_start)
            start_col = column_index_from_string(col_letter)
            max_cols = 0

            for row_idx, row_data in enumerate(data):
                max_cols = max(max_cols, len(row_data))
                for col_idx, raw_val in enumerate(row_data):
                    cell_obj = ws.cell(row=start_row + row_idx, column=start_col + col_idx)
                    if isinstance(cell_obj.value, str) and cell_obj.value.startswith("="):
                        formulas_updated += 1
                    clean_val = sanitize_excel_write_value(raw_val)
                    if isinstance(clean_val, str) and clean_val.startswith("="):
                        clean_val = self._normalize_formula_sheet_refs(wb, clean_val)
                        all_warnings.extend(
                            self._inspect_formula_sheet_refs(
                                wb, clean_val, cell_ref=f"{ws.title}!{cell_obj.coordinate}"
                            )
                        )
                    cell_obj.value = clean_val
                    if number_format and (row_idx > 0 or not header_formatting):
                        cell_obj.number_format = number_format
                    if row_idx == 0 and header_formatting:
                        self._apply_cell_formatting(cell_obj, header_formatting)
                    elif formatting:
                        self._apply_cell_formatting(cell_obj, formatting)

            end_row = start_row + len(data) - 1
            end_col = start_col + max(max_cols, 1) - 1
            results.append({
                "sheet_name": ws.title,
                "range": f"{clean_start}:{get_column_letter(end_col)}{end_row}",
                "rows_written": len(data),
                "columns_written": max_cols,
            })

        await self._save_working_copy(entry)
        out: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "operations_completed": len(results),
            "results": results,
        }
        if formulas_updated:
            out["formulas_updated"] = formulas_updated
        if all_warnings:
            out["formula_warnings"] = all_warnings
        return sanitize_for_json(out)

    async def add_formula(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        cell: str,
        formula: str,
        number_format: Optional[str] = None,
        formatting: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Write an Excel formula to a cell with optional number format."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)
        clean_cell = validate_cell_ref(cell)

        wb = entry["workbook_object"]
        if not formula or not isinstance(formula, str):
            raise ValueError("formula must be a non-empty string")
        normalized = formula if formula.startswith("=") else f"={formula}"
        normalized = self._normalize_formula_sheet_refs(wb, normalized)
        formula_warnings = self._inspect_formula_sheet_refs(wb, normalized, cell_ref=clean_cell)

        cell_obj = ws[clean_cell]
        cell_obj.value = normalized
        if number_format:
            cell_obj.number_format = number_format
        if formatting:
            self._apply_cell_formatting(cell_obj, formatting)

        await self._save_working_copy(entry)

        result: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "cell": clean_cell,
            "formula": normalized,
            "number_format": cell_obj.number_format,
        }
        if formula_warnings:
            result["formula_warnings"] = formula_warnings
        return sanitize_for_json(result)

    async def add_table(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        table_name: str,
        data_range: str,
        style_name: str = "TableStyleMedium9",
        show_first_column: bool = False,
        show_last_column: bool = False,
        show_row_stripes: bool = True,
        show_column_stripes: bool = False,
    ) -> Dict[str, Any]:
        """Create a structured Excel Table with explicit TableStyleLight/Medium/Dark and banding."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)
        clean_ref = validate_range_ref(data_range)
        if ":" not in clean_ref:
            raise ValueError("data_range must be a multi-cell range like 'A1:D10'")

        safe_name = "".join(c if (c.isalnum() or c == "_") else "_" for c in table_name)
        if not safe_name or not (safe_name[0].isalpha() or safe_name[0] == "_"):
            safe_name = f"Table_{safe_name}"

        tbl = Table(displayName=safe_name, ref=clean_ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name=style_name,
            showFirstColumn=bool(show_first_column),
            showLastColumn=bool(show_last_column),
            showRowStripes=bool(show_row_stripes),
            showColumnStripes=bool(show_column_stripes),
        )
        ws.add_table(tbl)
        await self._save_working_copy(entry)

        result: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "table_name": safe_name,
            "range": clean_ref,
            "style": style_name,
            "show_first_column": bool(show_first_column),
            "show_last_column": bool(show_last_column),
            "show_row_stripes": bool(show_row_stripes),
            "show_column_stripes": bool(show_column_stripes),
        }
        if style_name and not _TABLE_STYLE_RE.match(style_name):
            result["style_note"] = (
                f"'{style_name}' is not a standard TableStyleLight1-21 / TableStyleMedium1-28 / "
                "TableStyleDark1-11 name; Excel will use it if defined in the workbook or fall back to default."
            )
        return sanitize_for_json(result)

    async def update_table_style(
        self,
        workbook_id: str,
        table_name: str,
        sheet_name: Optional[str] = None,
        style_name: Optional[str] = None,
        show_first_column: Optional[bool] = None,
        show_last_column: Optional[bool] = None,
        show_row_stripes: Optional[bool] = None,
        show_column_stripes: Optional[bool] = None,
    ) -> Dict[str, Any]:
        """Update the table style and banded row/column toggles of an existing Excel Table."""
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        sheets = [self._get_sheet(entry, sheet_name)] if sheet_name else list(wb.worksheets)

        target_ws = None
        target_tbl = None
        for ws in sheets:
            if table_name in ws.tables:
                target_ws = ws
                target_tbl = ws.tables[table_name]
                break
        if target_tbl is None or target_ws is None:
            available = [
                f"{ws.title}:{name}" for ws in wb.worksheets for name in ws.tables.keys()
            ]
            raise ValueError(
                f"Excel table '{table_name}' not found. Available tables: {available}"
            )

        info = target_tbl.tableStyleInfo or TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        if style_name is not None:
            info.name = style_name
        if show_first_column is not None:
            info.showFirstColumn = bool(show_first_column)
        if show_last_column is not None:
            info.showLastColumn = bool(show_last_column)
        if show_row_stripes is not None:
            info.showRowStripes = bool(show_row_stripes)
        if show_column_stripes is not None:
            info.showColumnStripes = bool(show_column_stripes)
        target_tbl.tableStyleInfo = info

        await self._save_working_copy(entry)
        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": target_ws.title,
            "table_name": target_tbl.name,
            "style": info.name,
            "show_first_column": bool(info.showFirstColumn),
            "show_last_column": bool(info.showLastColumn),
            "show_row_stripes": bool(info.showRowStripes),
            "show_column_stripes": bool(info.showColumnStripes),
        })

    async def create_chart(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        chart_type: str,
        data_range: str,
        chart_title: str = "",
        position: str = "E5",
    ) -> Dict[str, Any]:
        """Create a bar, line, or pie chart from a worksheet range."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)
        clean_ref = validate_range_ref(data_range)
        clean_pos = validate_cell_ref(position)

        min_col, min_row, max_col, max_row = range_boundaries(clean_ref)

        chart_key = chart_type.lower()
        if chart_key == "bar":
            chart = BarChart()
        elif chart_key == "line":
            chart = LineChart()
        elif chart_key == "pie":
            chart = PieChart()
        else:
            raise ValueError(f"Unsupported chart type: '{chart_type}'. Use bar, line, or pie.")

        if chart_title:
            chart.title = chart_title

        data = Reference(
            ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row
        )
        chart.add_data(data, titles_from_data=True)
        ws.add_chart(chart, clean_pos)

        await self._save_working_copy(entry)

        return {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "chart_type": chart_key,
            "position": clean_pos,
            "data_range": clean_ref,
        }

    # ------------------------------------------------------------------
    # Save / Export
    # ------------------------------------------------------------------

    @staticmethod
    def get_live_sync_mode(raw: Optional[str] = None) -> str:
        """Parse OFFICE_MCP_LIVE_SYNC ('auto' default | 'off'), case-insensitive."""
        val = (os.environ.get("OFFICE_MCP_LIVE_SYNC", "auto") if raw is None else raw).strip().lower()
        if not val or val in ("auto", "on", "true", "1", "yes"):
            return "auto"
        if val in ("off", "false", "0", "no", "none"):
            return "off"
        logger.warning(
            f"Unknown OFFICE_MCP_LIVE_SYNC='{val}'. Falling back to 'auto'. Valid values: auto, off."
        )
        return "auto"

    async def save_workbook(
        self,
        workbook_id: str,
        file_path: Optional[str] = None,
        format: str = "xlsx",
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Save workbook in-place or to a new path (.xlsx, .xlsm, .csv, .pdf)."""
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        fmt = format.lower().lstrip(".")

        if file_path:
            save_path = resolve_file_path(file_path, must_exist=False)
        else:
            if entry["metadata"].get("is_temporary", True):
                raise ValueError(
                    "Workbook has no saved path yet; provide file_path to save."
                )
            save_path = Path(entry["working_path"])

        if not save_path.suffix:
            save_path = save_path.with_suffix(f".{fmt}")
        save_path.parent.mkdir(parents=True, exist_ok=True)

        if fmt in ("xlsx", "xlsm"):
            live_sync_mode = self.get_live_sync_mode()
            reloaded_in_excel = False
            view_restored = False
            live_sync_warning: Optional[str] = None
            open_live_wb: Optional[Dict[str, Any]] = None
            captured_view: Dict[str, Any] = {}

            if live_sync_mode == "auto":
                try:
                    open_live_wb = await self.applescript.find_open_excel_workbook_by_path(
                        str(save_path)
                    )
                except Exception as exc:
                    live_sync_warning = f"Could not check live Excel open state: {exc}"
                    logger.warning(live_sync_warning)

                if open_live_wb is not None:
                    wb_live_name = open_live_wb.get("name") or save_path.name
                    if not open_live_wb.get("saved", False):
                        raise RuntimeError(
                            f"File '{wb_live_name}' is currently open in Microsoft Excel with unsaved "
                            "manual changes. Save in Excel first (Cmd+S or excel_live_save) before "
                            "running save_workbook so your unsaved edits are not lost."
                        )
                    try:
                        captured_view = await self.applescript.capture_excel_workbook_view_state(
                            wb_live_name
                        )
                        await self.applescript.close_excel_workbook(
                            workbook_name=wb_live_name, save_changes="no"
                        )
                    except Exception as exc:
                        live_sync_warning = (
                            f"Excel has '{wb_live_name}' open but could not auto-close it "
                            f"(close & reopen Excel to see updates): {exc}"
                        )
                        logger.warning(live_sync_warning)
                        open_live_wb = None

            await asyncio.to_thread(wb.save, str(save_path))
            entry["working_path"] = save_path
            entry["metadata"]["file_path"] = str(save_path)
            entry["metadata"]["is_temporary"] = False
            logger.info(f"Saved workbook to {save_path}")

            if open_live_wb is not None:
                try:
                    reopen_res = await self.applescript.open_excel_file(
                        str(save_path), read_only=False, activate=False
                    )
                    reloaded_in_excel = True
                    reopened_name = reopen_res.get("workbook_name") or save_path.name
                    if captured_view:
                        view_restored = await self.applescript.restore_excel_workbook_view_state(
                            reopened_name, captured_view
                        )
                except Exception as exc:
                    live_sync_warning = (
                        f"Saved to disk cleanly, but could not auto-reopen in Excel: {exc}"
                    )
                    logger.warning(live_sync_warning)

            out: Dict[str, Any] = {
                "status": "success",
                "workbook_id": workbook_id,
                "file_path": str(save_path),
                "format": fmt,
            }
            if reloaded_in_excel:
                out["reloaded_in_excel"] = True
                out["view_restored"] = view_restored
            if live_sync_warning:
                out["live_sync_warning"] = live_sync_warning
            return sanitize_for_json(out)

        if fmt == "csv":
            ws = self._get_sheet(entry, sheet_name)

            def _write_csv() -> int:
                count = 0
                with open(save_path, "w", encoding="utf-8", newline="") as f:
                    writer = csv.writer(f)
                    for row in ws.iter_rows(values_only=True):
                        writer.writerow(["" if v is None else v for v in row])
                        count += 1
                return count

            rows_exported = await asyncio.to_thread(_write_csv)
            return {
                "status": "success",
                "workbook_id": workbook_id,
                "sheet_name": ws.title,
                "file_path": str(save_path),
                "format": "csv",
                "rows_exported": rows_exported,
            }

        if fmt == "pdf":
            await self._save_working_copy(entry)
            open_res = await self.applescript.open_excel_file(
                str(entry["working_path"]), read_only=True, activate=False
            )
            live_name = open_res["workbook_name"]
            try:
                if sheet_name:
                    await self.applescript.activate_live_excel_worksheet(
                        sheet_name=sheet_name, workbook_name=live_name
                    )
                await self.applescript.save_excel_workbook(
                    workbook_name=live_name,
                    file_path=str(save_path),
                    file_format="pdf",
                )
            finally:
                try:
                    await self.applescript.close_excel_workbook(
                        workbook_name=live_name, save_changes="no"
                    )
                except Exception:
                    pass
            return {
                "status": "success",
                "workbook_id": workbook_id,
                "file_path": str(save_path),
                "format": "pdf",
                "method": "applescript_export",
            }

        raise ValueError(f"Unsupported format: '{format}'. Supported: xlsx, xlsm, csv, pdf.")

    # ------------------------------------------------------------------
    # Explicit Column Widths, Range & Batch Worksheet Formatting
    # ------------------------------------------------------------------

    async def set_column_widths(
        self,
        workbook_id: str,
        sheet_name: Optional[str] = None,
        widths: Optional[Dict[str, float]] = None,
        auto_fit: bool = False,
        min_width: float = 10.0,
        max_width: float = 60.0,
        padding: float = 4.0,
    ) -> Dict[str, Any]:
        """Set explicit column widths (e.g. {'A': 42, 'B:G': 15}) and/or auto-fit columns."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)

        applied: Dict[str, float] = {}
        if auto_fit:
            fitted = self._compute_auto_fit_widths(
                ws, min_width=min_width, max_width=max_width, padding=padding
            )
            for col_letter, width in fitted.items():
                ws.column_dimensions[col_letter].width = width
                applied[col_letter] = width

        if widths:
            for col_spec, raw_width in widths.items():
                width_val = float(raw_width)
                if width_val <= 0:
                    raise ValueError(f"Column width for '{col_spec}' must be positive")
                for col_letter in self._expand_column_spec(col_spec):
                    ws.column_dimensions[col_letter].width = width_val
                    applied[col_letter] = width_val

        await self._save_working_copy(entry)
        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "auto_fit": auto_fit,
            "column_widths": applied,
        })

    async def format_range(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        cell_range: str,
        formatting: Optional[Dict[str, Any]] = None,
        number_format: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Apply font, fill, border, alignment, and/or number_format to an existing range."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)
        bounds_list = self._resolve_multi_format_bounds(ws, cell_range)

        cells_formatted = 0
        norm_refs: List[str] = []
        for min_col, min_row, max_col, max_row, norm_ref in bounds_list:
            norm_refs.append(norm_ref)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    cell_obj = ws.cell(row=r, column=c)
                    if number_format:
                        cell_obj.number_format = str(number_format)
                    if formatting:
                        self._apply_cell_formatting(cell_obj, formatting)
                    cells_formatted += 1

        await self._save_working_copy(entry)
        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "range": ", ".join(norm_refs),
            "cells_formatted": cells_formatted,
        })

    async def format_worksheet(
        self,
        workbook_id: str,
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
        """Batch-format a worksheet in a single call (styles, widths, heights, freeze panes, merges)."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)

        styled_ranges = 0
        cells_formatted = 0
        if auto_color_financial_model:
            cells_formatted += self._auto_color_financial_sheet(ws)
        if styles:
            for rule in styles:
                ref, style_name, num_fmt, fmt = self._extract_style_rule(rule)
                for min_col, min_row, max_col, max_row, _ in self._resolve_multi_format_bounds(ws, ref):
                    for r in range(min_row, max_row + 1):
                        for c in range(min_col, max_col + 1):
                            cell_obj = ws.cell(row=r, column=c)
                            if style_name:
                                self._apply_named_style_to_cell(entry["workbook_object"], cell_obj, style_name)
                            if num_fmt:
                                cell_obj.number_format = str(num_fmt)
                            if fmt:
                                self._apply_cell_formatting(cell_obj, fmt)
                            cells_formatted += 1
                styled_ranges += 1

        applied_widths: Dict[str, float] = {}
        if auto_fit_columns:
            fitted = self._compute_auto_fit_widths(
                ws, min_width=min_col_width, max_width=max_col_width, padding=col_padding
            )
            for col_letter, width in fitted.items():
                ws.column_dimensions[col_letter].width = width
                applied_widths[col_letter] = width

        if column_widths:
            for col_spec, raw_width in column_widths.items():
                width_val = float(raw_width)
                if width_val <= 0:
                    raise ValueError(f"Column width for '{col_spec}' must be positive")
                for col_letter in self._expand_column_spec(col_spec):
                    ws.column_dimensions[col_letter].width = width_val
                    applied_widths[col_letter] = width_val

        applied_heights: Dict[str, float] = {}
        if row_heights:
            for row_spec, raw_height in row_heights.items():
                height_val = float(raw_height)
                if height_val <= 0:
                    raise ValueError(f"Row height for '{row_spec}' must be positive")
                for row_idx in self._expand_row_spec(row_spec):
                    ws.row_dimensions[row_idx].height = height_val
                    applied_heights[str(row_idx)] = height_val

        if freeze_panes is not None:
            if freeze_panes == "":
                ws.freeze_panes = None
            else:
                ws.freeze_panes = validate_cell_ref(freeze_panes)

        merged: List[str] = []
        if merge_ranges:
            for mref in merge_ranges:
                clean_mref = validate_range_ref(mref)
                if ":" not in clean_mref:
                    raise ValueError(f"Merge range '{mref}' must be a multi-cell range like 'A1:G1'")
                ws.merge_cells(clean_mref)
                merged.append(clean_mref)

        if show_gridlines is not None:
            ws.views.sheetView[0].showGridLines = bool(show_gridlines)

        await self._save_working_copy(entry)
        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "styled_ranges": styled_ranges,
            "cells_formatted": cells_formatted,
            "column_widths": applied_widths,
            "row_heights": applied_heights,
            "freeze_panes": ws.freeze_panes,
            "merged_ranges": merged,
            "show_gridlines": ws.views.sheetView[0].showGridLines,
        })

    async def batch_format_workbook(
        self,
        workbook_id: str,
        operations: List[Dict[str, Any]],
    ) -> Dict[str, Any]:
        """Batch-format multiple worksheets in a single call and save once.

        Each item in operations may contain:
          sheet_name, column_widths, row_heights, freeze_panes, auto_fit_columns,
          min_col_width, max_col_width, col_padding, styles, merge_ranges, show_gridlines.
        """
        if not operations:
            raise ValueError("operations must be a non-empty list of worksheet format specs")

        entry = self._get_entry(workbook_id)
        sheet_summaries: List[Dict[str, Any]] = []
        op_warnings: List[str] = []
        total_cells = 0

        known_op_keys = {
            "sheet_name", "column_widths", "row_heights", "freeze_panes",
            "auto_fit_columns", "min_col_width", "max_col_width", "col_padding",
            "auto_color_financial_model", "merge_ranges", "show_gridlines",
            "styles", "rules", "cell_styles", "formats", "range_styles",
            "style_rules", "cell_formats", "formatting_rules", "number_formats",
        }
        style_list_keys = (
            "styles", "rules", "cell_styles", "formats",
            "range_styles", "style_rules", "cell_formats", "formatting_rules",
        )

        for op in operations:
            ws = self._get_sheet(entry, op.get("sheet_name"))
            styled_ranges = 0
            cells_formatted = 0

            is_flat_range_op = any(
                k in op for k in ("range", "cells", "cell_range", "range_ref")
            )
            if not is_flat_range_op:
                unknown_keys = sorted(set(op.keys()) - known_op_keys)
                if unknown_keys:
                    op_warnings.append(
                        f"Sheet '{ws.title}' ignored unrecognized keys {unknown_keys}; "
                        f"use 'styles': [{{'range': 'A1:G1', 'formatting': {{...}}}}]."
                    )

            if op.get("auto_color_financial_model"):
                cells_formatted += self._auto_color_financial_sheet(ws)

            combined_rules: List[Dict[str, Any]] = []
            if is_flat_range_op:
                # Fallback: accept flat per-range op like {'sheet_name': 'IS', 'range': 'A1:G1', 'bold': True}
                flat_rule = {k: v for k, v in op.items() if k not in known_op_keys}
                combined_rules.append(flat_rule)
            for skey in style_list_keys:
                val = op.get(skey)
                if isinstance(val, list):
                    combined_rules.extend(val)
            num_fmts_map = op.get("number_formats")
            if isinstance(num_fmts_map, dict):
                for nf_range, nf_code in num_fmts_map.items():
                    combined_rules.append({"range": str(nf_range), "number_format": str(nf_code)})

            for rule in combined_rules:
                ref, style_name, num_fmt, fmt = self._extract_style_rule(rule, sheet_title=ws.title)
                for min_col, min_row, max_col, max_row, _ in self._resolve_multi_format_bounds(ws, ref):
                    for r in range(min_row, max_row + 1):
                        for c in range(min_col, max_col + 1):
                            cell_obj = ws.cell(row=r, column=c)
                            if style_name:
                                self._apply_named_style_to_cell(entry["workbook_object"], cell_obj, style_name)
                            if num_fmt:
                                cell_obj.number_format = str(num_fmt)
                            if fmt:
                                self._apply_cell_formatting(cell_obj, fmt)
                            cells_formatted += 1
                styled_ranges += 1

            applied_widths: Dict[str, float] = {}
            if op.get("auto_fit_columns"):
                fitted = self._compute_auto_fit_widths(
                    ws,
                    min_width=float(op.get("min_col_width", 10.0)),
                    max_width=float(op.get("max_col_width", 60.0)),
                    padding=float(op.get("col_padding", 4.0)),
                )
                for col_letter, width in fitted.items():
                    ws.column_dimensions[col_letter].width = width
                    applied_widths[col_letter] = width

            for col_spec, raw_width in (op.get("column_widths") or {}).items():
                width_val = float(raw_width)
                if width_val <= 0:
                    raise ValueError(f"Column width for '{col_spec}' must be positive")
                for col_letter in self._expand_column_spec(col_spec):
                    ws.column_dimensions[col_letter].width = width_val
                    applied_widths[col_letter] = width_val

            for row_spec, raw_height in (op.get("row_heights") or {}).items():
                height_val = float(raw_height)
                if height_val <= 0:
                    raise ValueError(f"Row height for '{row_spec}' must be positive")
                for row_idx in self._expand_row_spec(row_spec):
                    ws.row_dimensions[row_idx].height = height_val

            if "freeze_panes" in op and op["freeze_panes"] is not None:
                fp = op["freeze_panes"]
                ws.freeze_panes = None if fp == "" else validate_cell_ref(fp)

            for mref in op.get("merge_ranges") or []:
                clean_mref = validate_range_ref(mref)
                if ":" not in clean_mref:
                    raise ValueError(f"Merge range '{mref}' must be a multi-cell range")
                ws.merge_cells(clean_mref)

            if "show_gridlines" in op and op["show_gridlines"] is not None:
                ws.views.sheetView[0].showGridLines = bool(op["show_gridlines"])

            total_cells += cells_formatted
            sheet_summaries.append({
                "sheet_name": ws.title,
                "styled_ranges": styled_ranges,
                "cells_formatted": cells_formatted,
                "column_widths_set": len(applied_widths),
                "freeze_panes": ws.freeze_panes,
            })

        await self._save_working_copy(entry)
        out: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "sheets_formatted": len(sheet_summaries),
            "total_cells_formatted": total_cells,
            "sheets": sheet_summaries,
        }
        if op_warnings:
            out["warnings"] = op_warnings
        return sanitize_for_json(out)

    # ------------------------------------------------------------------
    # Named Cell Styles & Financial Model Presets
    # ------------------------------------------------------------------

    async def list_cell_styles(self, workbook_id: str) -> Dict[str, Any]:
        """List workbook named styles and built-in Excel financial-model style presets."""
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "workbook_named_styles": list(wb.named_styles),
            "built_in_presets": list(_NAMED_CELL_STYLE_PRESETS.keys()),
        })

    async def apply_cell_style(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        cell_range: str,
        style_name: str,
        overrides: Optional[Dict[str, Any]] = None,
        number_format: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Apply a named Excel cell style (e.g. 'Good', 'Input', 'Total', 'Heading 1') to a range."""
        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        ws = self._get_sheet(entry, sheet_name)

        resolved_style = self._ensure_named_style(wb, style_name)
        min_col, min_row, max_col, max_row, normalized_ref = self._resolve_format_bounds(
            ws, cell_range
        )

        cells_styled = 0
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell_obj = ws.cell(row=r, column=c)
                cell_obj.style = resolved_style
                if number_format:
                    cell_obj.number_format = str(number_format)
                if overrides:
                    self._apply_cell_formatting(cell_obj, overrides)
                cells_styled += 1

        await self._save_working_copy(entry)
        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "range": normalized_ref,
            "style_name": resolved_style,
            "cells_styled": cells_styled,
        })

    # ------------------------------------------------------------------
    # Data Validation Rules (openpyxl DataValidation)
    # ------------------------------------------------------------------

    async def add_data_validation(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        cell_range: str,
        validation_type: str,
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
        """Add an Excel data validation rule (list, whole, decimal, date, time, textLength, custom)."""
        valid_types = {"list", "whole", "decimal", "date", "time", "textLength", "custom"}
        vtype = str(validation_type).strip()
        if vtype not in valid_types:
            raise ValueError(
                f"Unsupported validation_type '{validation_type}'. Use one of {sorted(valid_types)}."
            )

        valid_operators = {
            "between", "notBetween", "equal", "notEqual",
            "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual",
        }
        if operator is not None and operator not in valid_operators:
            raise ValueError(
                f"Unsupported operator '{operator}'. Use one of {sorted(valid_operators)}."
            )

        entry = self._get_entry(workbook_id)
        wb = entry["workbook_object"]
        ws = self._get_sheet(entry, sheet_name)
        _, _, _, _, normalized_ref = self._resolve_format_bounds(ws, cell_range)

        # Format formula1 cleanly for 'list' validations when passed a Python list of choices
        f1_str: Optional[str] = None
        if formula1 is not None:
            if vtype == "list" and isinstance(formula1, (list, tuple)):
                joined = ",".join(str(item).replace('"', "") for item in formula1)
                f1_str = f'"{joined}"'
            else:
                f1_str = str(formula1)
                if f1_str.startswith("="):
                    f1_str = self._normalize_formula_sheet_refs(wb, f1_str)

        f2_str: Optional[str] = str(formula2) if formula2 is not None else None
        if f2_str and f2_str.startswith("="):
            f2_str = self._normalize_formula_sheet_refs(wb, f2_str)

        # Note: openpyxl's showDropDown is inverted in OOXML (False shows the dropdown arrow!)
        dv = DataValidation(
            type=vtype,
            operator=operator,
            formula1=f1_str,
            formula2=f2_str,
            allow_blank=bool(allow_blank),
            showDropDown=(not bool(show_dropdown)) if vtype == "list" else None,
            errorStyle=error_style if error_style in ("stop", "warning", "information") else "stop",
        )
        if prompt_title or prompt:
            dv.showInputMessage = True
            if prompt_title:
                dv.promptTitle = str(prompt_title)
            if prompt:
                dv.prompt = str(prompt)
        if error_title or error:
            dv.showErrorMessage = True
            if error_title:
                dv.errorTitle = str(error_title)
            if error:
                dv.error = str(error)

        ws.add_data_validation(dv)
        dv.add(normalized_ref)
        await self._save_working_copy(entry)

        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "range": normalized_ref,
            "validation_type": vtype,
            "operator": operator,
            "formula1": f1_str,
            "formula2": f2_str,
            "allow_blank": bool(allow_blank),
            "total_validations_on_sheet": len(ws.data_validations.dataValidation),
        })

    async def list_data_validations(
        self,
        workbook_id: str,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """List all data validation rules defined on a worksheet."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)

        rules = []
        for idx, dv in enumerate(ws.data_validations.dataValidation):
            rules.append({
                "index": idx,
                "type": dv.type,
                "operator": dv.operator,
                "formula1": dv.formula1,
                "formula2": dv.formula2,
                "ranges": str(dv.sqref),
                "allow_blank": bool(dv.allow_blank),
                "prompt_title": dv.promptTitle,
                "prompt": dv.prompt,
                "error_title": dv.errorTitle,
                "error": dv.error,
                "error_style": dv.errorStyle,
            })

        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "count": len(rules),
            "validations": rules,
        })

    async def clear_data_validations(
        self,
        workbook_id: str,
        sheet_name: Optional[str] = None,
        cell_range: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Clear all data validations on a sheet, or only rules matching a specific range."""
        entry = self._get_entry(workbook_id)
        ws = self._get_sheet(entry, sheet_name)
        existing = list(ws.data_validations.dataValidation)

        if not cell_range:
            removed = len(existing)
            ws.data_validations.dataValidation = []
        else:
            _, _, _, _, normalized_ref = self._resolve_format_bounds(ws, cell_range)
            kept = []
            removed = 0
            for dv in existing:
                ranges = str(dv.sqref).split()
                if normalized_ref in ranges:
                    remaining = [r for r in ranges if r != normalized_ref]
                    if remaining:
                        dv.sqref = " ".join(remaining)
                        kept.append(dv)
                    removed += 1
                else:
                    kept.append(dv)
            ws.data_validations.dataValidation = kept

        await self._save_working_copy(entry)
        return sanitize_for_json({
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "removed_count": removed,
            "remaining_count": len(ws.data_validations.dataValidation),
        })

    # ------------------------------------------------------------------
    # Formatting, Formula & Dimension Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _normalize_hex_color(color: Any) -> Optional[str]:
        """Normalize a hex color string ('#1F4E79', '1f4e79', 'FF1F4E79') for openpyxl."""
        if color is None:
            return None
        raw = str(color).strip().lstrip("#").upper()
        if len(raw) == 3 and all(ch in "0123456789ABCDEF" for ch in raw):
            return "".join(ch * 2 for ch in raw)
        if len(raw) in (6, 8) and all(ch in "0123456789ABCDEF" for ch in raw):
            return raw
        logger.warning(f"Ignoring invalid hex color: '{color}'")
        return None

    @classmethod
    def _make_side(cls, spec: Any, default_color: str = "000000") -> Side:
        """Build an openpyxl Side from a bool, style string, or {'style', 'color'} dict."""
        valid_styles = {
            "thin", "medium", "thick", "double", "dashed", "dotted",
            "hair", "mediumDashed", "dashDot", "mediumDashDot",
            "dashDotDot", "mediumDashDotDot", "slantDashDot",
        }
        if spec is None or spec is False:
            return Side(style=None)
        if spec is True:
            return Side(style="thin", color=default_color)
        if isinstance(spec, str):
            s = spec.strip()
            if not s or s.lower() == "none":
                return Side(style=None)
            style_val = s if s in valid_styles else "thin"
            return Side(style=style_val, color=default_color)
        if isinstance(spec, dict):
            raw_style = str(spec.get("style", "thin")).strip()
            if not raw_style or raw_style.lower() == "none":
                return Side(style=None)
            style_val = raw_style if raw_style in valid_styles else "thin"
            col = cls._normalize_hex_color(spec.get("color")) or default_color
            return Side(style=style_val, color=col)
        return Side(style="thin", color=default_color)

    @staticmethod
    def _expand_column_spec(col_spec: Union[str, int]) -> List[str]:
        """Expand 'A', 'B:G', or 1-based column index into uppercase column letters."""
        if isinstance(col_spec, int):
            if col_spec < 1:
                raise ValueError(f"Column index must be >= 1, got {col_spec}")
            return [get_column_letter(col_spec)]
        spec = str(col_spec).strip().upper().replace("$", "")
        if not spec:
            raise ValueError("Column specification cannot be empty")
        if spec.isdigit():
            idx = int(spec)
            if idx < 1:
                raise ValueError(f"Column index must be >= 1, got {idx}")
            return [get_column_letter(idx)]
        if ":" in spec:
            parts = spec.split(":", 1)
            start_s, end_s = parts[0].strip(), parts[1].strip()
            if not (start_s.isalpha() and end_s.isalpha()):
                raise ValueError(f"Invalid column range: '{col_spec}' (expected e.g. 'B:G')")
            start_idx = column_index_from_string(start_s)
            end_idx = column_index_from_string(end_s)
            if end_idx < start_idx:
                start_idx, end_idx = end_idx, start_idx
            return [get_column_letter(i) for i in range(start_idx, end_idx + 1)]
        if spec.isalpha():
            # Validate via column_index_from_string
            column_index_from_string(spec)
            return [spec]
        raise ValueError(f"Invalid column specification: '{col_spec}' (expected e.g. 'A' or 'B:G')")

    @staticmethod
    def _expand_row_spec(row_spec: Union[str, int]) -> List[int]:
        """Expand 1, '1', or '2:10' into 1-based row indices."""
        if isinstance(row_spec, int):
            if row_spec < 1:
                raise ValueError(f"Row index must be >= 1, got {row_spec}")
            return [row_spec]
        spec = str(row_spec).strip()
        if ":" in spec:
            start_s, end_s = spec.split(":", 1)
            start_r, end_r = int(start_s.strip()), int(end_s.strip())
            if start_r < 1 or end_r < 1:
                raise ValueError(f"Row indices must be >= 1: '{row_spec}'")
            if end_r < start_r:
                start_r, end_r = end_r, start_r
            return list(range(start_r, end_r + 1))
        idx = int(spec)
        if idx < 1:
            raise ValueError(f"Row index must be >= 1, got {idx}")
        return [idx]

    @classmethod
    def _resolve_format_bounds(cls, ws, cell_range: str) -> Tuple[int, int, int, int, str]:
        """Resolve 'A1', 'A1:G10', row range '1:1', or column range 'B:G' to 1-based bounds."""
        raw = str(cell_range).strip().upper().replace("$", "")
        if re.match(r"^[1-9][0-9]*:[1-9][0-9]*$", raw):
            r1, r2 = (int(x) for x in raw.split(":", 1))
            if r2 < r1:
                r1, r2 = r2, r1
            max_col = max(ws.max_column or 1, 1)
            return 1, r1, max_col, r2, f"A{r1}:{get_column_letter(max_col)}{r2}"
        if re.match(r"^[A-Z]{1,3}:[A-Z]{1,3}$", raw):
            c1_s, c2_s = raw.split(":", 1)
            c1 = column_index_from_string(c1_s)
            c2 = column_index_from_string(c2_s)
            if c2 < c1:
                c1, c2 = c2, c1
            max_row = max(ws.max_row or 1, 1)
            return c1, 1, c2, max_row, f"{get_column_letter(c1)}1:{get_column_letter(c2)}{max_row}"

        clean_ref = validate_range_ref(raw)
        if ":" in clean_ref:
            min_col, min_row, max_col, max_row = range_boundaries(clean_ref)
            return min_col, min_row, max_col, max_row, clean_ref
        col_l, row_i = coordinate_from_string(clean_ref)
        col_i = column_index_from_string(col_l)
        return col_i, row_i, col_i, row_i, clean_ref

    @staticmethod
    def _compute_auto_fit_widths(
        ws,
        min_width: float = 10.0,
        max_width: float = 60.0,
        padding: float = 4.0,
    ) -> Dict[str, float]:
        """Estimate display width per populated column to prevent clipped text and '###'."""
        widths: Dict[str, float] = {}
        max_col = ws.max_column or 0
        max_row = ws.max_row or 0
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            best = 0.0
            for row_idx in range(1, max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val = cell.value
                if val is None:
                    continue
                if isinstance(val, str) and val.startswith("="):
                    # Formulas evaluate to numbers/text; reserve comfortable width for formatted values
                    num_fmt = cell.number_format or ""
                    est = 14.0 if any(ch in num_fmt for ch in ("$", "#", "0", "%")) else 12.0
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    num_fmt = cell.number_format or ""
                    if "%" in num_fmt:
                        est = float(len(f"{val * 100:.1f}%")) + 1.0
                    elif any(ch in num_fmt for ch in ("$", ",")):
                        est = float(len(f"$({abs(val):,.2f})"))
                    else:
                        est = float(len(str(val)))
                else:
                    lines = str(val).splitlines() or [""]
                    est = float(max(len(line) for line in lines))
                if cell.font and cell.font.bold:
                    est *= 1.1
                if est > best:
                    best = est
            if best > 0:
                widths[col_letter] = round(max(min_width, min(max_width, best + padding)), 2)
        return widths

    @staticmethod
    def _normalize_formula_sheet_refs(wb, formula: str) -> str:
        """Ensure same-workbook sheet references with spaces/special chars are properly single-quoted.

        Prevents Excel from flagging unquoted same-workbook references (e.g. `=Balance Sheet!B2`)
        as broken or external workbook links, without altering already-quoted references or
        string literals.
        """
        if not formula or not formula.startswith("="):
            return formula

        sheets_needing_quotes = [
            name
            for name in wb.sheetnames
            if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name)
            or re.match(r"^[A-Za-z]{1,3}[0-9]+$", name)
        ]
        if not sheets_needing_quotes:
            return formula

        # Longest first so 'Balance Sheet 2026' matches before 'Balance Sheet'
        sheets_needing_quotes.sort(key=len, reverse=True)

        # Tokenize formula into double-quoted strings, single-quoted sheet refs, and other segments
        token_re = re.compile(r'("(?:""|[^"])*"|\'(?:\'\'|[^\'])*\'!)')
        parts = token_re.split(formula)
        for idx in range(0, len(parts), 2):
            seg = parts[idx]
            if not seg or "!" not in seg:
                continue
            for sheet_name in sheets_needing_quotes:
                escaped_sheet = sheet_name.replace("'", "''")
                # Match unquoted sheet_name immediately followed by '!' and preceded by start or delimiter
                pattern = re.compile(
                    r"(^|[=(,+\-*/:; &<>^])" + re.escape(sheet_name) + r"!"
                )
                seg = pattern.sub(rf"\1'{escaped_sheet}'!", seg)
            parts[idx] = seg
        return "".join(parts)

    @staticmethod
    def _inspect_formula_sheet_refs(
        wb,
        formula: str,
        cell_ref: Optional[str] = None,
    ) -> List[str]:
        """Validate cross-sheet references in a formula and flag non-existent or external sheets.

        Prevents unintended Excel external-link prompts when a formula references a misspelled
        sheet name (e.g. `Balance_Sheet` when the workbook has `Balance Sheet`) or an accidental
        `[Workbook.xlsx]` external prefix.
        """
        if not formula or not formula.startswith("=") or "!" not in formula:
            return []

        warnings: List[str] = []
        existing = list(wb.sheetnames)
        existing_set = set(existing)
        norm_lookup = {
            re.sub(r"[\s_\-]+", "", name).lower(): name for name in existing
        }

        # Strip double-quoted string literals so "Sheet!A1" inside text is ignored
        stripped = re.sub(r'"(?:""|[^"])*"', '""', formula)
        loc = f" in {cell_ref}" if cell_ref else ""

        # Find quoted ('Sheet Name'!) and unquoted (Sheet_Name!) sheet references before '!'
        ref_re = re.compile(
            r"(?:'((?:''|[^'])+)'|([A-Za-z0-9_.\[\]]+))!"
        )
        seen: set[str] = set()
        for match in ref_re.finditer(stripped):
            quoted_name, unquoted_name = match.group(1), match.group(2)
            raw_ref = quoted_name.replace("''", "'") if quoted_name is not None else unquoted_name
            if not raw_ref or raw_ref in seen:
                continue
            seen.add(raw_ref)

            if "[" in raw_ref or "]" in raw_ref:
                warnings.append(
                    f"Formula{loc} contains external workbook reference '{raw_ref}!'. "
                    f"Remove '[...]' to reference a sheet in this workbook ({existing})."
                )
                continue

            if raw_ref not in existing_set:
                norm_key = re.sub(r"[\s_\-]+", "", raw_ref).lower()
                suggestion = norm_lookup.get(norm_key)
                if suggestion:
                    quoted_sug = f"'{suggestion}'" if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", suggestion) else suggestion
                    warnings.append(
                        f"Formula{loc} references non-existent sheet '{raw_ref}' (may trigger an external link prompt). "
                        f"Did you mean {quoted_sug}!?"
                    )
                else:
                    warnings.append(
                        f"Formula{loc} references non-existent sheet '{raw_ref}' (may trigger an external link prompt). "
                        f"Available sheets: {existing}."
                    )

        return warnings

    @classmethod
    def _ensure_named_style(cls, wb, style_name: str) -> str:
        """Ensure a named cell style exists on the workbook, registering built-in presets on demand."""
        if not style_name or not isinstance(style_name, str):
            raise ValueError("style_name must be a non-empty string")
        cleaned = style_name.strip()

        # Exact match in workbook already
        if cleaned in wb.named_styles:
            return cleaned

        # Case-insensitive match against existing workbook styles
        for existing in wb.named_styles:
            if existing.lower() == cleaned.lower():
                return existing

        # Match against built-in presets (case-insensitive)
        preset_key = None
        for key in _NAMED_CELL_STYLE_PRESETS:
            if key.lower() == cleaned.lower():
                preset_key = key
                break

        if preset_key is None:
            raise ValueError(
                f"Named style '{style_name}' not found. "
                f"Available presets: {list(_NAMED_CELL_STYLE_PRESETS.keys())}, "
                f"workbook styles: {list(wb.named_styles)}"
            )

        spec = _NAMED_CELL_STYLE_PRESETS[preset_key]
        ns = NamedStyle(name=preset_key)
        if "font" in spec:
            ns.font = Font(**spec["font"])
        if "fill" in spec:
            hex_bg = cls._normalize_hex_color(spec["fill"])
            if hex_bg:
                ns.fill = PatternFill(start_color=hex_bg, end_color=hex_bg, fill_type="solid")
        if "border" in spec:
            bspec = spec["border"]
            bcol = cls._normalize_hex_color(bspec.get("color")) or "000000"
            if "all" in bspec:
                s = cls._make_side(bspec["all"], bcol)
                ns.border = Border(left=s, right=s, top=s, bottom=s)
            else:
                ns.border = Border(
                    left=cls._make_side(bspec.get("left"), bcol),
                    right=cls._make_side(bspec.get("right"), bcol),
                    top=cls._make_side(bspec.get("top"), bcol),
                    bottom=cls._make_side(bspec.get("bottom"), bcol),
                )
        wb.add_named_style(ns)
        return preset_key

    @classmethod
    def _resolve_multi_format_bounds(
        cls, ws, cell_range: Union[str, List[str]]
    ) -> List[Tuple[int, int, int, int, str]]:
        """Resolve a single range, comma-separated multi-range ('G2:G4, G6:G9, G15'), or list of ranges."""
        parts = (
            [str(p).strip() for p in cell_range if str(p).strip()]
            if isinstance(cell_range, (list, tuple))
            else [p.strip() for p in str(cell_range).split(",") if p.strip()]
        )
        if not parts:
            raise ValueError("Range specification cannot be empty")
        return [cls._resolve_format_bounds(ws, part) for part in parts]

    def _auto_color_financial_sheet(self, ws) -> int:
        """Auto-color data cells by financial-model convention: blue inputs, black formulas, green cross-sheet links."""
        colored = 0
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0
        for r in range(2, max_r + 1):
            for c in range(2, max_c + 1):
                cell_obj = ws.cell(row=r, column=c)
                val = cell_obj.value
                if val is None:
                    continue
                if isinstance(val, str) and val.startswith("="):
                    color = "008000" if "!" in val else "000000"
                    self._apply_cell_formatting(cell_obj, {"font_color": color})
                    colored += 1
                elif isinstance(val, (int, float)) and not isinstance(val, bool):
                    self._apply_cell_formatting(cell_obj, {"font_color": "0000FF"})
                    colored += 1
        return colored

    @staticmethod
    def _extract_style_rule(
        rule: Dict[str, Any],
        sheet_title: Optional[str] = None,
    ) -> Tuple[str, Optional[str], Optional[str], Optional[Dict[str, Any]]]:
        """Accept both nested {'range', 'formatting': {...}} and flat {'range', 'bold': True, ...} style rules."""
        ref = (
            rule.get("range")
            or rule.get("cells")
            or rule.get("cell_range")
            or rule.get("range_ref")
        )
        if not ref:
            loc = f" on sheet '{sheet_title}'" if sheet_title else ""
            raise ValueError(f"Each style entry{loc} must include a 'range' (e.g. 'A1:G1')")
        raw_style_val = rule.get("style")
        style_name = (
            rule.get("cell_style")
            or rule.get("style_name")
            or (raw_style_val if isinstance(raw_style_val, str) else None)
        )
        num_fmt = (
            rule.get("number_format")
            or rule.get("num_format")
            or rule.get("num_fmt")
            or rule.get("format_code")
        )
        reserved = {
            "range", "cells", "cell_range", "range_ref",
            "formatting", "style", "format",
            "number_format", "num_format", "num_fmt", "format_code",
            "cell_style", "style_name",
        }
        merged_fmt: Dict[str, Any] = {}
        for k, v in rule.items():
            if k not in reserved:
                merged_fmt[k] = v
        for nested_key in ("style", "format", "formatting"):
            nested_val = rule.get(nested_key)
            if isinstance(nested_val, dict):
                merged_fmt.update(nested_val)
        return str(ref), style_name, num_fmt, (merged_fmt or None)

    @classmethod
    def _apply_named_style_to_cell(cls, wb, cell, style_name: str) -> None:
        resolved = cls._ensure_named_style(wb, style_name)
        cell.style = resolved

    def _apply_cell_formatting(self, cell, formatting: Dict[str, Any]) -> None:
        """Apply optional named cell_style plus font, fill, alignment, per-side borders, and number_format."""
        # 0. Optional named cell style applied first so explicit overrides merge cleanly on top
        raw_style = formatting.get("cell_style", formatting.get("style_name"))
        if raw_style:
            wb = getattr(getattr(cell, "parent", None), "parent", None)
            if wb is not None:
                self._apply_named_style_to_cell(wb, cell, str(raw_style))

        # 1. Font (merge with existing cell.font; supports both flat keys and nested 'font': {...})
        nested_font = formatting.get("font") if isinstance(formatting.get("font"), dict) else {}
        font_src: Dict[str, Any] = {**nested_font, **formatting}
        font_keys = {
            "bold", "italic", "underline", "strike", "strikethrough",
            "font_size", "size", "font_color", "color", "font_name", "font_family", "name",
        }
        if nested_font or any(k in formatting for k in (font_keys - {"size", "name"})):
            base_font = copy(cell.font) if cell.font else Font()
            font_kwargs: Dict[str, Any] = {
                "name": base_font.name,
                "size": base_font.size,
                "bold": base_font.bold,
                "italic": base_font.italic,
                "underline": base_font.underline,
                "strike": base_font.strike,
                "color": base_font.color,
            }
            if "bold" in font_src:
                font_kwargs["bold"] = bool(font_src["bold"])
            if "italic" in font_src:
                font_kwargs["italic"] = bool(font_src["italic"])
            if "underline" in font_src:
                u = font_src["underline"]
                font_kwargs["underline"] = "single" if u is True else (None if not u else str(u))
            if "strike" in font_src or "strikethrough" in font_src:
                font_kwargs["strike"] = bool(font_src.get("strike", font_src.get("strikethrough")))
            raw_sz = formatting.get("font_size", nested_font.get("size", nested_font.get("font_size")))
            if raw_sz is not None:
                font_kwargs["size"] = float(raw_sz)
            raw_fn = formatting.get("font_name") or formatting.get("font_family") or nested_font.get("name") or nested_font.get("font_name")
            if raw_fn:
                font_kwargs["name"] = str(raw_fn)
            raw_fc = formatting.get("font_color", formatting.get("color", nested_font.get("color", nested_font.get("font_color"))))
            if raw_fc is not None:
                norm_fc = self._normalize_hex_color(raw_fc)
                if norm_fc:
                    font_kwargs["color"] = norm_fc
            cell.font = Font(**font_kwargs)

        # 2. Fill (bg_color / background_color / fill_color / fill)
        raw_fill = formatting.get(
            "bg_color",
            formatting.get(
                "background_color",
                formatting.get("fill_color", formatting.get("fill")),
            ),
        )
        if raw_fill is not None:
            fill_type = "solid"
            color_val = raw_fill
            if isinstance(raw_fill, dict):
                fill_type = str(raw_fill.get("fill_type", raw_fill.get("patternType", "solid")))
                color_val = (
                    raw_fill.get("color")
                    or raw_fill.get("start_color")
                    or raw_fill.get("fgColor")
                    or raw_fill.get("bgColor")
                )
            if color_val is None or str(color_val).lower() == "none":
                cell.fill = PatternFill(fill_type=None)
            else:
                norm_bg = self._normalize_hex_color(color_val)
                if norm_bg:
                    cell.fill = PatternFill(start_color=norm_bg, end_color=norm_bg, fill_type=fill_type)

        # 3. Alignment (merge with existing cell.alignment; supports flat keys or nested 'alignment': {...})
        nested_align = formatting.get("alignment") if isinstance(formatting.get("alignment"), dict) else {}
        align_src: Dict[str, Any] = {**nested_align, **formatting}
        align_keys = {"horizontal", "vertical", "wrap_text", "indent", "text_rotation"}
        if nested_align or any(k in formatting for k in align_keys):
            base_align = copy(cell.alignment) if cell.alignment else Alignment()
            align_kwargs: Dict[str, Any] = {
                "horizontal": base_align.horizontal,
                "vertical": base_align.vertical,
                "wrap_text": base_align.wrap_text,
                "indent": base_align.indent,
                "text_rotation": base_align.text_rotation,
            }
            if "horizontal" in align_src:
                align_kwargs["horizontal"] = str(align_src["horizontal"])
            if "vertical" in align_src:
                align_kwargs["vertical"] = str(align_src["vertical"])
            if "wrap_text" in align_src:
                align_kwargs["wrap_text"] = bool(align_src["wrap_text"])
            if "indent" in align_src:
                align_kwargs["indent"] = int(align_src["indent"])
            if "text_rotation" in align_src:
                align_kwargs["text_rotation"] = int(align_src["text_rotation"])
            cell.alignment = Alignment(**align_kwargs)

        # 4. Border (supports bool, style string, dict with top/bottom/left/right/all/color, or border_* keys)
        border_spec = formatting.get("border", formatting.get("borders"))
        has_side_keys = any(
            k in formatting for k in ("border_top", "border_bottom", "border_left", "border_right")
        )
        if border_spec is not None or has_side_keys:
            base_border = copy(cell.border) if cell.border else Border()
            left_s = base_border.left
            right_s = base_border.right
            top_s = base_border.top
            bottom_s = base_border.bottom
            default_bcol = self._normalize_hex_color(formatting.get("border_color")) or "000000"

            if isinstance(border_spec, dict):
                default_bcol = self._normalize_hex_color(border_spec.get("color")) or default_bcol
                if "all" in border_spec:
                    all_side = self._make_side(border_spec["all"], default_bcol)
                    left_s = right_s = top_s = bottom_s = all_side
                if "left" in border_spec:
                    left_s = self._make_side(border_spec["left"], default_bcol)
                if "right" in border_spec:
                    right_s = self._make_side(border_spec["right"], default_bcol)
                if "top" in border_spec:
                    top_s = self._make_side(border_spec["top"], default_bcol)
                if "bottom" in border_spec:
                    bottom_s = self._make_side(border_spec["bottom"], default_bcol)
            elif border_spec is not None:
                all_side = self._make_side(border_spec, default_bcol)
                left_s = right_s = top_s = bottom_s = all_side

            if "border_left" in formatting:
                left_s = self._make_side(formatting["border_left"], default_bcol)
            if "border_right" in formatting:
                right_s = self._make_side(formatting["border_right"], default_bcol)
            if "border_top" in formatting:
                top_s = self._make_side(formatting["border_top"], default_bcol)
            if "border_bottom" in formatting:
                bottom_s = self._make_side(formatting["border_bottom"], default_bcol)

            cell.border = Border(left=left_s, right=right_s, top=top_s, bottom=bottom_s)

        # 5. Number format
        raw_nf = formatting.get("number_format") or formatting.get("num_format") or formatting.get("num_fmt") or formatting.get("format_code")
        if raw_nf:
            cell.number_format = str(raw_nf)

    async def get_workbook_info(self, workbook_id: str) -> Dict[str, Any]:
        entry = self._get_entry(workbook_id)
        self._sync_metadata(entry)
        return sanitize_for_json(entry["metadata"])

    async def list_workbooks(self) -> List[Dict[str, Any]]:
        return [sanitize_for_json(data["metadata"]) for data in self.active_workbooks.values()]
