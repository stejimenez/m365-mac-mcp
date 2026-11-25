"""
Excel Controller for Microsoft 365 for Mac MCP (m365-mac-mcp).

Facade composing modular domain mixins (`controllers.excel.*`) for full,
non-blocking file-level read-write operations (.xlsx, .xlsm, .csv) plus
optional live Excel integration via AppleScriptBridge.
"""

import asyncio
import atexit
import os
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from controllers.excel import (
    ExcelChartsTablesMixin,
    ExcelFormatterMixin,
    ExcelIOMixin,
    ExcelReaderMixin,
    ExcelWriterMixin,
    _NAMED_CELL_STYLE_PRESETS,
    _TABLE_STYLE_RE,
)
from integrations.applescript_bridge import AppleScriptBridge
from utils.logger import setup_logger
from utils.validators import atomic_save_via_temp

logger = setup_logger(__name__)


class ExcelController(
    ExcelIOMixin,
    ExcelReaderMixin,
    ExcelWriterMixin,
    ExcelChartsTablesMixin,
    ExcelFormatterMixin,
):
    """Controller for Excel file-level read-write operations (.xlsx, .xlsm, .csv)."""

    def __init__(self, max_workbooks: int = 20):
        self.applescript = AppleScriptBridge()
        self.active_workbooks: Dict[str, Dict[str, Any]] = {}
        self.max_workbooks = max_workbooks
        self.temp_dir = Path(tempfile.gettempdir()) / "m365-mac-mcp" / "excel"
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self._owned_temp_files: set[Path] = set()
        self._wb_locks: Dict[Tuple[int, str], asyncio.Lock] = {}
        atexit.register(self._cleanup_temp_files)

    def _get_wb_lock(self, workbook_id: str) -> asyncio.Lock:
        """Return an event-loop-safe asyncio.Lock serializing concurrent mutations/saves on one workbook."""
        try:
            loop_id = id(asyncio.get_running_loop())
        except RuntimeError:
            loop_id = 0
        key = (loop_id, workbook_id)
        lock = self._wb_locks.get(key)
        if lock is None:
            lock = asyncio.Lock()
            self._wb_locks[key] = lock
        return lock

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
        """Keep active workbook count bounded; evict clean (saved) workbooks before dirty ones."""
        while len(self.active_workbooks) >= self.max_workbooks:
            clean_ids = [wid for wid, e in self.active_workbooks.items() if not e.get("dirty")]
            evict_id = clean_ids[0] if clean_ids else next(iter(self.active_workbooks))
            entry = self.active_workbooks.pop(evict_id)
            was_dirty = bool(entry.get("dirty"))
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
            if was_dirty:
                logger.warning(
                    f"Evicted DIRTY in-memory workbook {evict_id} with unsaved edits to stay within limit ({self.max_workbooks})"
                )
            else:
                logger.info(
                    f"Evicted oldest clean in-memory workbook {evict_id} to stay within limit ({self.max_workbooks})"
                )

    def _resolve_workbook_id(self, workbook_id: Optional[str]) -> str:
        if workbook_id:
            if workbook_id not in self.active_workbooks:
                raise ValueError(f"Workbook '{workbook_id}' not found. Open or create a workbook first.")
            return workbook_id
        if len(self.active_workbooks) == 1:
            return next(iter(self.active_workbooks))
        if not self.active_workbooks:
            raise ValueError("No active workbooks in memory. Open or create a workbook first.")
        raise ValueError(
            f"workbook_id is required when multiple workbooks are active: {list(self.active_workbooks.keys())}"
        )

    def _get_entry(self, workbook_id: Optional[str]) -> Dict[str, Any]:
        resolved_id = self._resolve_workbook_id(workbook_id)
        return self.active_workbooks[resolved_id]

    def _get_sheet(self, entry: Dict[str, Any], sheet_name: Optional[str]):
        wb = entry["workbook_object"]
        if not sheet_name:
            return wb.active
        if sheet_name not in wb.sheetnames:
            raise ValueError(
                f"Worksheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}"
            )
        return wb[sheet_name]

    async def _save_working_copy(self, entry: Dict[str, Any], force: bool = False) -> None:
        """Mark workbook dirty in memory; defer disk serialization until save_workbook unless force=True.

        Deferring disk writes avoids redundant full-.xlsx zip writes on multi-step edits,
        preserves original cached formula values on disk until explicit save, and ensures
        save_workbook's M365_MCP_LIVE_SYNC unsaved-edits guard is never bypassed.
        """
        entry["dirty"] = True
        if force:
            wb = entry["workbook_object"]
            target = entry["working_path"]
            await asyncio.to_thread(wb.save, str(target))
            entry["dirty"] = False

    def _sync_metadata(self, entry: Dict[str, Any]) -> None:
        wb = entry["workbook_object"]
        entry["metadata"]["worksheet_count"] = len(wb.worksheets)
        entry["metadata"]["worksheets"] = list(wb.sheetnames)
        entry["metadata"]["active_sheet"] = wb.active.title if wb.active else None

