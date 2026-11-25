"""Cell/range/table readers, workbook inspection, and formula auditing."""

import asyncio
import re
from pathlib import Path
from typing import (
    Any,
    Dict,
    List,
    Optional,
    Tuple,
)

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.utils.cell import coordinate_from_string

from utils.logger import setup_logger
from utils.sheet_layout import analyze_worksheet_layout
from utils.validators import (
    has_invalid_excel_formula_quotes,
    resolve_file_path,
    sanitize_for_json,
    validate_cell_ref,
    validate_range_ref,
)

logger = setup_logger(__name__)


class ExcelReaderMixin:
    """Cell/range/table readers, workbook inspection, and formula auditing."""

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
        workbook_id = entry["metadata"]["workbook_id"]
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
        max_rows: int = 150,
        max_cols: int = 30,
    ) -> Dict[str, Any]:
        """Read a cell range (or used range) as a JSON-safe 2D matrix."""
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
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

        orig_rows = max_row - min_row + 1
        orig_cols = max_col - min_col + 1
        truncated = False
        if orig_rows > max_rows:
            max_row = min_row + max_rows - 1
            truncated = True
        if orig_cols > max_cols:
            max_col = min_col + max_cols - 1
            truncated = True

        def _extract() -> tuple[List[List[Any]], int]:
            uncached_formulas = 0
            if evaluate_formulas:
                wb_e, ws_e = self._load_cached_evaluated_sheet(entry, ws.title)
                if ws_e is not None:
                    try:
                        eval_rows = [
                            list(row)
                            for row in ws_e.iter_rows(
                                min_row=min_row,
                                max_row=max_row,
                                min_col=min_col,
                                max_col=max_col,
                                values_only=True,
                            )
                        ]
                        raw_rows = [
                            list(row)
                            for row in ws.iter_rows(
                                min_row=min_row,
                                max_row=max_row,
                                min_col=min_col,
                                max_col=max_col,
                                values_only=True,
                            )
                        ]
                        for e_row, r_row in zip(eval_rows, raw_rows):
                            for e_val, r_val in zip(e_row, r_row):
                                if e_val is None and isinstance(r_val, str) and r_val.startswith("="):
                                    uncached_formulas += 1
                        return eval_rows, uncached_formulas
                    finally:
                        wb_e.close()
            raw_rows = [
                list(row)
                for row in ws.iter_rows(
                    min_row=min_row,
                    max_row=max_row,
                    min_col=min_col,
                    max_col=max_col,
                    values_only=True,
                )
            ]
            return raw_rows, 0

        rows_data, uncached_formula_count = await asyncio.to_thread(_extract)
        if rows_data:
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
            payload["total_rows"] = orig_rows
            payload["total_columns"] = orig_cols
            next_start = f"{get_column_letter(min_col)}{max_row + 1}"
            payload["pagination_hint"] = (
                f"Range truncated to {actual_range} of {orig_rows}x{orig_cols}. "
                f"Pass cell_range='{next_start}:...' (or higher max_rows/max_cols) to read further slices."
            )
        if uncached_formula_count > 0:
            payload["uncached_formula_count"] = uncached_formula_count
            payload["warning"] = (
                f"{uncached_formula_count} formula cell(s) in this range have no cached value (null). "
                "File-level reads cannot calculate formulas—do NOT retry read_range with other ranges "
                "or estimate values mentally. Save the workbook and read live via excel_live_open + excel_live_batch_read_ranges."
            )
        return sanitize_for_json(payload)

    @staticmethod
    def _ensure_valid_excel_table_headers(wb: Any) -> List[str]:
        """Ensure every Excel Table header cell is non-empty, non-formula, and unique.

        Microsoft Excel refuses to open .xlsx files (Parameter error -50 / repair prompt)
        when an OOXML Table has a blank, formula, or duplicate column header.
        """
        repaired: List[str] = []
        for ws in wb.worksheets:
            for tbl in ws.tables.values():
                try:
                    min_col, min_row, max_col, _ = range_boundaries(tbl.ref)
                except Exception:
                    continue
                seen: set[str] = set()
                for offset, col_idx in enumerate(range(min_col, max_col + 1), start=1):
                    cell = ws.cell(row=min_row, column=col_idx)
                    raw = cell.value
                    header = str(raw).strip() if raw is not None else ""
                    if not header or header.startswith("="):
                        header = f"Column{offset}"
                        cell.value = header
                        repaired.append(f"{ws.title}!{cell.coordinate} -> '{header}'")
                    base = header
                    suffix = 2
                    while header.lower() in seen:
                        header = f"{base}_{suffix}"
                        suffix += 1
                    if header != (str(raw).strip() if raw is not None else ""):
                        cell.value = header
                    seen.add(header.lower())
        return repaired

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
        max_rows_per_table: int = 100,
    ) -> Dict[str, Any]:
        """List and read structured Excel Tables in a sheet or entire workbook (capped per table to protect context)."""
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        wb = entry["workbook_object"]
        sheets = [self._get_sheet(entry, sheet_name)] if sheet_name else list(wb.worksheets)
        row_cap = max(1, int(max_rows_per_table))

        tables_out: List[Dict[str, Any]] = []
        for ws in sheets:
            for tbl in ws.tables.values():
                min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
                total_body_rows = max(0, max_row - min_row)
                eff_max_row = min(max_row, min_row + row_cap)
                rows_data = [
                    list(row)
                    for row in ws.iter_rows(
                        min_row=min_row,
                        max_row=eff_max_row,
                        min_col=min_col,
                        max_col=max_col,
                        values_only=True,
                    )
                ]
                headers = rows_data[0] if rows_data else []
                body = rows_data[1:] if len(rows_data) > 1 else []
                tbl_info: Dict[str, Any] = {
                    "sheet_name": ws.title,
                    "table_name": tbl.name,
                    "ref": tbl.ref,
                    "headers": headers,
                    "row_count": total_body_rows,
                    "rows": body,
                }
                if total_body_rows > row_cap:
                    tbl_info["truncated"] = True
                    tbl_info["returned_rows"] = len(body)
                tables_out.append(tbl_info)

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
        sheet_names: Optional[List[str]] = None,
        max_rows_per_sheet: int = 50,
        max_cols_per_sheet: int = 20,
        include_formulas: bool = True,
        include_evaluated: bool = True,
    ) -> Dict[str, Any]:
        """Inspect all (or selected) worksheets (structure, headers, sparse formulas, cached values, tables) in ONE read-only call."""
        if not workbook_id and not file_path:
            raise ValueError("Provide either workbook_id or file_path to inspect_workbook")

        owned_wb = None
        if workbook_id:
            entry = self._get_entry(workbook_id)
            workbook_id = entry["metadata"]["workbook_id"]
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
                target_sheets = list(wb.worksheets)
                missing_sheets: List[str] = []
                if sheet_names:
                    requested = [str(s).strip() for s in sheet_names if str(s).strip()]
                    if requested:
                        exact_map = {ws.title: ws for ws in wb.worksheets}
                        lower_map = {ws.title.strip().lower(): ws for ws in wb.worksheets}
                        matched = []
                        seen_titles: set[str] = set()
                        for req in requested:
                            ws_hit = exact_map.get(req) or lower_map.get(req.lower())
                            if ws_hit is None:
                                missing_sheets.append(req)
                            elif ws_hit.title not in seen_titles:
                                matched.append(ws_hit)
                                seen_titles.add(ws_hit.title)
                        if not matched:
                            raise ValueError(
                                f"None of the requested sheet_names {requested} matched. "
                                f"Available sheets: {list(wb.sheetnames)}"
                            )
                        target_sheets = matched

                # Auto-scale preview rows on large (>6 tab) unfiltered workbooks to prevent 65K context overflow
                auto_scaled_preview = (
                    not sheet_names
                    and len(target_sheets) > 6
                    and max_rows_per_sheet == 50
                )
                eff_max_rows_per_sheet = 15 if auto_scaled_preview else max_rows_per_sheet

                sheets_out: List[Dict[str, Any]] = []
                for ws in target_sheets:
                    max_r = min(ws.max_row or 0, eff_max_rows_per_sheet)
                    max_c = min(ws.max_column or 0, max_cols_per_sheet)
                    truncated_rows = (ws.max_row or 0) > eff_max_rows_per_sheet
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

                    sheet_layout = analyze_worksheet_layout(ws)
                    detected_tables = [b.to_summary_dict() for b in sheet_layout.data_blocks()]
                    sheet_info: Dict[str, Any] = {
                        "sheet_name": ws.title,
                        "total_rows": ws.max_row or 0,
                        "total_columns": ws.max_column or 0,
                        "used_range": f"A1:{get_column_letter(max(actual_cols or ws.max_column or 1, 1))}{max(actual_rows or ws.max_row or 1, 1)}",
                    }
                    if detected_tables:
                        sheet_info["detected_tables"] = detected_tables
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
                if sheet_names:
                    out["inspected_sheet_count"] = len(sheets_out)
                if missing_sheets:
                    out["missing_sheets"] = missing_sheets
                if auto_scaled_preview:
                    out["preview_note"] = (
                        f"Large workbook ({len(target_sheets)} sheets): auto-capped preview to top {eff_max_rows_per_sheet} rows/sheet "
                        "to protect context window. Pass sheet_names=[...] for deeper inspection of specific tabs."
                    )
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

    _R1C1_TOKEN_RE = re.compile(
        r'"(?:""|[^"])*"|'
        r"(?P<sheet>(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
        r"(?P<col_abs>\$?)(?P<col>[A-Z]{1,3})(?P<row_abs>\$?)(?P<row>[1-9]\d*)(?!\s*\()"
    )

    @classmethod
    def _formula_to_r1c1_signature(cls, formula: str, row: int, col: int) -> str:
        """Convert an A1 Excel formula at (row, col) into a normalized R1C1 relative signature."""
        def _repl(match: re.Match) -> str:
            txt = match.group(0)
            if txt.startswith('"'):
                return txt  # preserve string literals untouched
            col_letters = match.group("col")
            try:
                c_num = column_index_from_string(col_letters)
            except Exception:
                return txt
            r_num = int(match.group("row"))
            # Guard against huge invalid column letters (Excel max XFD = 16384)
            if c_num > 16384 or r_num > 1048576:
                return txt
            sheet_raw = match.group("sheet") or ""
            if sheet_raw:
                s_name = sheet_raw[:-1]
                if s_name.startswith("'") and s_name.endswith("'") and len(s_name) >= 2:
                    s_name = s_name[1:-1].replace("''", "'")
                sheet_norm = f"{s_name.upper()}!"
            else:
                sheet_norm = ""
            r_part = f"R{r_num}" if match.group("row_abs") == "$" else (
                "R" if r_num == row else f"R[{r_num - row}]"
            )
            c_part = f"C{c_num}" if match.group("col_abs") == "$" else (
                "C" if c_num == col else f"C[{c_num - col}]"
            )
            return f"{sheet_norm}{r_part}{c_part}"

        # Normalize whitespace outside quoted strings and uppercase non-string tokens
        parts = re.split(r'("(?:""|[^"])*")', formula.strip())
        rebuilt: List[str] = []
        for idx, part in enumerate(parts):
            if idx % 2 == 1:
                rebuilt.append(part)
            else:
                compact = re.sub(r"\s+", "", part).upper()
                rebuilt.append(cls._R1C1_TOKEN_RE.sub(_repl, compact))
        sig = "".join(rebuilt)
        # Canonicalize pure top-level commutative additions (e.g. =RC[-1]+R[-1]C == =R[-1]C+RC[-1])
        # so cumulative rows written in either operand order don't trigger false-positive R1C1 flags.
        if sig.startswith("=") and "+" in sig and not any(ch in sig[1:] for ch in '-*/()^,"&<>='):
            terms = sig[1:].split("+")
            if all(terms):
                sig = "=" + "+".join(sorted(terms))
        return sig

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
            workbook_id = entry["metadata"]["workbook_id"]
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

                calc_row_keywords = (
                    "total",
                    "subtotal",
                    "sub-total",
                    "sous-total",
                    "margin",
                    "marge",
                    "margen",
                    "%",
                    "ebitda",
                    "ebit",
                    "net income",
                    "résultat net",
                    "resultat net",
                    "utilidad neta",
                    "gross profit",
                    "marge brute",
                    "utilidad bruta",
                    "operating income",
                    "check",
                    "cagr",
                    "growth",
                    "croissance",
                    "crecimiento",
                    "sum",
                    "fcf",
                    "free cash flow",
                    "npv",
                    "van",
                    "irr",
                    "tir",
                    "mirr",
                    "wacc",
                    "eps",
                    "ratio",
                    "multiple",
                )

                row_inconsistencies: List[Dict[str, Any]] = []
                total_col_keywords = ("total", "sum", "avg", "average", "cagr", "ytd", "full year", "check")

                for ws in wb.worksheets:
                    max_r = ws.max_row or 0
                    max_c = ws.max_column or 0
                    sheet_layout = analyze_worksheet_layout(ws)
                    block_index_map = {id(b): idx for idx, b in enumerate(sheet_layout.blocks)}
                    col_formula_counts: Dict[Tuple[int, int], int] = {}
                    row_formula_counts: Dict[Tuple[int, int], int] = {}
                    row_formula_entries: Dict[Tuple[int, int], List[Tuple[int, str, str, Any, str]]] = {}
                    numeric_literals: List[Tuple[int, int, int, Any, Any]] = []

                    # Sparse cell iteration via ws._cells prevents hanging/OOM on ghost 1,048,576-row sheets
                    cells_dict = getattr(ws, "_cells", None)
                    if isinstance(cells_dict, dict):
                        iter_coords = sorted(
                            (rc for rc, cobj in list(cells_dict.items()) if cobj is not None and cobj.value is not None)
                        )
                    else:
                        iter_coords = [
                            (r, c)
                            for r in range(1, max_r + 1)
                            for c in range(1, max_c + 1)
                            if ws.cell(row=r, column=c).value is not None
                        ]

                    # Bound wb_eval read strictly to populated bounds so ghost 1,048,576-row sheets never OOM
                    ws_e = wb_eval[ws.title] if (wb_eval is not None and ws.title in wb_eval.sheetnames) else None
                    if ws_e is not None and iter_coords:
                        max_pop_r = max(r for r, _ in iter_coords)
                        max_pop_c = max(c for _, c in iter_coords)
                        eval_rows = [
                            list(r)
                            for r in ws_e.iter_rows(
                                min_row=1,
                                max_row=max_pop_r,
                                min_col=1,
                                max_col=max_pop_c,
                                values_only=True,
                            )
                        ]
                    else:
                        eval_rows = []

                    for r, c in iter_coords:
                            val = ws._cells[(r, c)].value if isinstance(cells_dict, dict) else ws.cell(row=r, column=c).value
                            if val is None:
                                continue
                            blk = sheet_layout.block_at(r, c)
                            blk_id = block_index_map.get(id(blk), -1) if blk is not None else -1
                            row_label = (
                                ws.cell(row=r, column=blk.label_col).value
                                if blk is not None
                                else ws.cell(row=r, column=1).value
                            )
                            coord = f"{get_column_letter(c)}{r}"
                            cached_val = (
                                eval_rows[r - 1][c - 1]
                                if (r - 1 < len(eval_rows) and c - 1 < len(eval_rows[r - 1]))
                                else None
                            )

                            is_fml = isinstance(val, str) and val.startswith("=")
                            if is_fml:
                                total_formulas += 1
                                col_formula_counts[(blk_id, c)] = col_formula_counts.get((blk_id, c), 0) + 1
                                row_formula_counts[(blk_id, r)] = row_formula_counts.get((blk_id, r), 0) + 1
                                r1c1_sig = self._formula_to_r1c1_signature(val, r, c)
                                row_formula_entries.setdefault((blk_id, r), []).append((c, coord, val, cached_val, r1c1_sig))
                                sheet_ref_warnings.extend(
                                    self._inspect_formula_sheet_refs(
                                        wb, val, cell_ref=f"{ws.title}!{coord}"
                                    )
                                )
                                # Catch circular self-references & invalid 1-argument NPV() directly as error_cells
                                s_no_str = re.sub(r'"[^"]*"', '""', val.strip())
                                col_let = get_column_letter(c)
                                self_pat = rf"(?<![A-Za-z0-9_!'$])\$?{col_let}\$?{r}(?![A-Za-z0-9_])"
                                if re.search(self_pat, s_no_str, re.IGNORECASE):
                                    error_cells.append({
                                        "sheet_name": ws.title,
                                        "cell": coord,
                                        "row_label": row_label,
                                        "error": "#CIRCULAR!",
                                        "formula": val,
                                        "cached_value": cached_val,
                                        "reason": f"Self-referencing circular formula ({coord} references itself)",
                                    })
                                if re.search(r"(?<![A-Za-z0-9_])NPV\s*\(\s*[^,()]+\s*\)", s_no_str, re.IGNORECASE):
                                    error_cells.append({
                                        "sheet_name": ws.title,
                                        "cell": coord,
                                        "row_label": row_label,
                                        "error": "#SYNTAX!",
                                        "formula": val,
                                        "cached_value": cached_val,
                                        "reason": "Invalid 1-argument NPV() — NPV requires at least 2 arguments: NPV(rate, values)",
                                    })
                                if has_invalid_excel_formula_quotes(val):
                                    error_cells.append({
                                        "sheet_name": ws.title,
                                        "cell": coord,
                                        "row_label": row_label,
                                        "error": "#SYNTAX!",
                                        "formula": val,
                                        "cached_value": cached_val,
                                        "reason": "Invalid quote syntax in Excel formula (backslash-escaped \\\" or standalone single-quoted string literal); Excel requires double quotes \"...\"",
                                    })
                            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                                # Layout-agnostic literal filtering: skip TITLE_BLOCK, SECTION_BANNER,
                                # header rows, code/label columns, and MATRIX_TABLE axis headers on ANY row/col!
                                skip_literal = False
                                if blk is not None:
                                    if blk.kind in ("TITLE_BLOCK", "SECTION_BANNER"):
                                        skip_literal = True
                                    elif r in blk.header_rows or r in blk.subheader_rows:
                                        skip_literal = True
                                    elif c <= blk.label_col:
                                        skip_literal = True
                                    elif blk.kind == "MATRIX_TABLE" and (r == blk.matrix_header_row or c == blk.matrix_axis_col):
                                        skip_literal = True
                                elif r <= 1 or c <= 1:
                                    skip_literal = True
                                if not skip_literal:
                                    numeric_literals.append((blk_id, r, c, val, row_label))

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

                    # High-precision SOX 404 literal check (layout-agnostic across arbitrary tables/rows/cols):
                    yr0_headers = {"0", "0.0", "year 0", "yr 0", "y0", "t0", "t=0", "period 0"}
                    for blk_id, r, c, lit_val, row_label in numeric_literals:
                        blk = sheet_layout.blocks[blk_id] if 0 <= blk_id < len(sheet_layout.blocks) else None
                        # Never flag table/matrix axis header rows (e.g. WACC 0.085..0.165 or years 0..6)
                        if blk is not None and (r in blk.header_rows or r == blk.min_row):
                            continue
                        first_val_col = blk.value_start_col if blk is not None else 2
                        # Standard Period-0 / Year-0 zero (Revenue/Tax=0) or unity discount factor (1.0) is not a SOX override
                        if lit_val in (0, 1) and c == first_val_col:
                            hdr_rows_check = (blk.header_rows if (blk is not None and blk.header_rows) else [1, 2])
                            hdr_texts = {
                                str(ws.cell(row=hr, column=c).value).strip().lower()
                                for hr in hdr_rows_check
                                if hr < r and ws.cell(row=hr, column=c).value is not None
                            }
                            if hdr_texts & yr0_headers:
                                continue
                        row_f = row_formula_counts.get((blk_id, r), 0)
                        col_f = col_formula_counts.get((blk_id, c), 0)
                        label_low = str(row_label or "").strip().lower()
                        is_calc_label = bool(label_low) and any(kw in label_low for kw in calc_row_keywords)
                        if row_f >= 2:
                            coord = f"{get_column_letter(c)}{r}"
                            suspicious_literals.append({
                                "sheet_name": ws.title,
                                "cell": coord,
                                "row_label": row_label,
                                "value": lit_val,
                                "reason": f"Hardcoded number in formula row {r} ({row_f} sibling formulas in row)",
                            })
                        elif col_f >= 3 and is_calc_label:
                            coord = f"{get_column_letter(c)}{r}"
                            suspicious_literals.append({
                                "sheet_name": ws.title,
                                "cell": coord,
                                "row_label": row_label,
                                "value": lit_val,
                                "reason": f"Hardcoded number on calculated row '{row_label}' in formula column {get_column_letter(c)}",
                            })

                    # Deterministic R1C1 financial-logic check — strictly scoped per (TableBlock, row)
                    # so side-by-side tables on the same row never collide and headers come from the block's own header row!
                    for (blk_id, r), entries in row_formula_entries.items():
                        if len(entries) < 3:
                            continue
                        blk = sheet_layout.blocks[blk_id] if 0 <= blk_id < len(sheet_layout.blocks) else None
                        sig_counts: Dict[str, int] = {}
                        sig_example: Dict[str, Tuple[str, str]] = {}
                        for _, coord, fml, _, sig in entries:
                            sig_counts[sig] = sig_counts.get(sig, 0) + 1
                            sig_example.setdefault(sig, (coord, fml))
                        dominant_sig, dominant_cnt = max(sig_counts.items(), key=lambda kv: kv[1])
                        if dominant_cnt < 2 or (dominant_cnt / len(entries)) < 0.6:
                            continue
                        row_label = (
                            ws.cell(row=r, column=blk.label_col).value
                            if blk is not None
                            else ws.cell(row=r, column=1).value
                        )
                        min_c_row = entries[0][0]
                        max_c_row = entries[-1][0]
                        dom_coord, dom_fml = sig_example[dominant_sig]
                        dom_is_agg = dominant_sig.startswith(("=SUM(", "=AVERAGE(", "=NPV(", "=IRR("))
                        blk_hdr_rows = (
                            blk.header_rows
                            if (blk is not None and blk.header_rows)
                            else ([blk.matrix_header_row] if (blk is not None and blk.matrix_header_row) else [1, 2])
                        )

                        for c, coord, fml, cached_val, sig in entries:
                            if sig == dominant_sig:
                                continue
                            # Skip legitimate Period-1 seed formula on the leftmost column of the table
                            if c == min_c_row:
                                continue
                            # Skip legitimate rightmost Total/Summary column (e.g., =SUM(B2:M2))
                            col_hdr = " ".join(
                                str(ws.cell(row=hr, column=c).value or "").strip().lower()
                                for hr in blk_hdr_rows
                                if hr < r
                            )
                            if c == max_c_row and (
                                any(kw in col_hdr for kw in total_col_keywords)
                                or (not dom_is_agg and sig.startswith(("=SUM(", "=AVERAGE(", "=NPV(", "=IRR(")))
                            ):
                                continue
                            expected_fml: Optional[str] = None
                            try:
                                expected_fml = Translator(dom_fml, origin=dom_coord).translate_formula(coord)
                            except Exception:
                                expected_fml = None
                            col_hdr_display = None
                            for hr in reversed(blk_hdr_rows):
                                if hr < r and ws.cell(row=hr, column=c).value is not None:
                                    col_hdr_display = ws.cell(row=hr, column=c).value
                                    break
                            item: Dict[str, Any] = {
                                "sheet_name": ws.title,
                                "cell": coord,
                                "row_label": row_label,
                                "col_header": col_hdr_display,
                                "formula": fml,
                                "r1c1_signature": sig,
                                "dominant_r1c1_pattern": dominant_sig,
                                "sibling_matches": dominant_cnt,
                                "cached_value": cached_val,
                                "reason": (
                                    f"Financial-logic inconsistency on row '{row_label or r}': "
                                    f"R1C1 signature '{sig}' breaks dominant row pattern '{dominant_sig}' "
                                    f"({dominant_cnt}/{len(entries)} sibling cells) — possible wrong-month / wrong-period reference."
                                ),
                            }
                            if expected_fml:
                                item["expected_formula"] = expected_fml
                            row_inconsistencies.append(item)

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
                    "row_inconsistency_count": len(row_inconsistencies),
                    "clean": (
                        len(error_cells) == 0
                        and len(sheet_ref_warnings) == 0
                        and len(row_inconsistencies) == 0
                        and ext_count == 0
                    ),
                }
                if workbook_id:
                    audit_out["workbook_id"] = workbook_id
                if disk_path:
                    audit_out["file_path"] = str(disk_path)
                if error_cells:
                    audit_out["error_cells"] = error_cells[:50]
                    if len(error_cells) > 50:
                        audit_out["error_cells_truncated"] = True
                if sheet_ref_warnings:
                    audit_out["sheet_ref_warnings"] = sheet_ref_warnings[:50]
                    if len(sheet_ref_warnings) > 50:
                        audit_out["sheet_ref_warnings_truncated"] = True
                if suspicious_literals:
                    audit_out["suspicious_literals"] = suspicious_literals[:50]
                if row_inconsistencies:
                    audit_out["row_inconsistencies"] = row_inconsistencies[:50]
                return audit_out
            finally:
                if wb_eval is not None:
                    wb_eval.close()

        try:
            return sanitize_for_json(await asyncio.to_thread(_audit))
        finally:
            if owned_wb is not None:
                owned_wb.close()
