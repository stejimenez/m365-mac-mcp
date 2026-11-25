"""Excel Tables (ListObjects), Charts, and Data Validation management."""

from typing import Any, Dict, Optional

from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from controllers.excel.constants import _TABLE_STYLE_RE
from utils.logger import setup_logger
from utils.validators import sanitize_for_json, validate_cell_ref, validate_range_ref

logger = setup_logger(__name__)


class ExcelChartsTablesMixin:
    """Excel Tables (ListObjects), Charts, and Data Validation management."""

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
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id = entry["metadata"]["workbook_id"]
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
        """Create a bar, line, or pie chart from a worksheet range (supports same-sheet or cross-sheet 'Sheet'!A1:H5 data_range)."""
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        ws = self._get_sheet(entry, sheet_name)
        raw_range = str(data_range or "").strip()
        data_ws = ws
        if "!" in raw_range:
            src_sheet_raw, cell_part = raw_range.rsplit("!", 1)
            src_sheet = src_sheet_raw.strip()
            if src_sheet.startswith("'") and src_sheet.endswith("'") and len(src_sheet) >= 2:
                src_sheet = src_sheet[1:-1].replace("''", "'")
            data_ws = self._get_sheet(entry, src_sheet)
            raw_range = cell_part.strip()
        clean_ref = validate_range_ref(raw_range)
        clean_pos = validate_cell_ref(position)

        min_col, min_row, max_col, max_row = range_boundaries(clean_ref)

        chart_key = chart_type.strip().lower().replace("-", "_").replace(" ", "_")
        if chart_key in {
            "bar", "column", "col", "column_clustered", "clustered_column",
            "bar_clustered", "clustered_bar", "column_stacked", "stacked_column",
            "bar_stacked", "stacked_bar",
        }:
            chart = BarChart()
            if "stacked" in chart_key:
                chart.grouping = "stacked"
        elif chart_key in {"line", "line_markers", "line_stacked"}:
            chart = LineChart()
        elif chart_key in {"pie", "pie_3d", "doughnut"}:
            chart = PieChart()
        else:
            raise ValueError(f"Unsupported chart type: '{chart_type}'. Use bar, line, or pie.")

        chart.width = 20
        chart.height = 12
        chart.style = 10
        if chart_title:
            chart.title = chart_title
            if hasattr(chart.title, "overlay"):
                chart.title.overlay = False
        if hasattr(chart, "legend") and chart.legend is not None:
            chart.legend.position = "b"
            chart.legend.overlay = False
        if hasattr(chart, "x_axis") and chart.x_axis is not None:
            chart.x_axis.delete = False
            chart.x_axis.tickLblPos = "low"
        if hasattr(chart, "y_axis") and chart.y_axis is not None:
            chart.y_axis.delete = False

        def _is_formula(v: Any) -> bool:
            return isinstance(v, str) and v.strip().startswith("=")

        def _is_num(v: Any) -> bool:
            return isinstance(v, (int, float)) and not isinstance(v, bool)

        def _is_text(v: Any) -> bool:
            return isinstance(v, str) and bool(v.strip()) and not v.strip().startswith("=")

        if max_col > min_col and max_row > min_row:
            first_row_vals = [
                data_ws.cell(row=min_row, column=c).value
                for c in range(min_col, max_col + 1)
            ]
            first_row_all_formulas = all(_is_formula(v) for v in first_row_vals if v is not None) and any(
                v is not None for v in first_row_vals
            )

            # Case A: Caller passed data-only matrix (e.g. C4:G8 where Row 4 is all formulas).
            # Auto-detect adjacent header row (Row 3) and category column (Col B) so Excel never
            # eats numeric data row 1 as legend labels like '-6961603.893'.
            if first_row_all_formulas:
                adj_header_ok = (
                    min_row > 1
                    and any(
                        data_ws.cell(row=min_row - 1, column=c).value is not None
                        for c in range(min_col, max_col + 1)
                    )
                    and not all(
                        _is_formula(data_ws.cell(row=min_row - 1, column=c).value)
                        for c in range(min_col, max_col + 1)
                        if data_ws.cell(row=min_row - 1, column=c).value is not None
                    )
                )
                adj_cat_ok = (
                    min_col > 1
                    and all(
                        data_ws.cell(row=r, column=min_col - 1).value is not None
                        and not _is_formula(data_ws.cell(row=r, column=min_col - 1).value)
                        for r in range(min_row, max_row + 1)
                    )
                )
                data_min_row = min_row - 1 if adj_header_ok else min_row
                data = Reference(
                    data_ws, min_col=min_col, min_row=data_min_row, max_col=max_col, max_row=max_row
                )
                chart.add_data(data, titles_from_data=adj_header_ok)
                if adj_cat_ok:
                    cats = Reference(
                        data_ws, min_col=min_col - 1, min_row=min_row, max_row=max_row
                    )
                    chart.set_categories(cats)
            else:
                first_col_vals = [
                    data_ws.cell(row=r, column=min_col).value
                    for r in range(min_row + 1, max_row + 1)
                ]
                body_vals = [
                    data_ws.cell(row=r, column=c).value
                    for r in range(min_row + 1, max_row + 1)
                    for c in range(min_col + 1, max_col + 1)
                ]
                has_text_labels = any(_is_text(v) for v in first_col_vals)
                has_numeric_vals = any(_is_num(v) for v in first_col_vals)
                # Also recognize numeric category column (e.g. CF Scale 0.8..1.2 or Years 2020..2026)
                # when first column is non-formula inputs while the chart body is formulas.
                first_col_is_cat = (has_text_labels and not has_numeric_vals) or (
                    has_numeric_vals
                    and not any(_is_formula(v) for v in first_col_vals)
                    and any(_is_formula(v) for v in body_vals)
                )
                if first_col_is_cat:
                    data = Reference(
                        data_ws, min_col=min_col + 1, min_row=min_row, max_col=max_col, max_row=max_row
                    )
                    cats = Reference(
                        data_ws, min_col=min_col, min_row=min_row + 1, max_row=max_row
                    )
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(cats)
                else:
                    data = Reference(
                        data_ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row
                    )
                    chart.add_data(data, titles_from_data=True)
        else:
            data = Reference(
                data_ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row
            )
            chart.add_data(data, titles_from_data=True)

        # Prevent Excel for Mac from rendering negative BarChart values as hollow white outlines
        if isinstance(chart, BarChart):
            for ser in chart.series:
                ser.invertIfNegative = False

        ws.add_chart(chart, clean_pos)

        await self._save_working_copy(entry)

        return {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "source_sheet": data_ws.title,
            "chart_type": chart_key,
            "position": clean_pos,
            "data_range": clean_ref,
        }

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
        workbook_id = entry["metadata"]["workbook_id"]
        wb = entry["workbook_object"]
        ws = self._get_sheet(entry, sheet_name)
        _, _, _, _, normalized_ref = self._resolve_format_bounds(ws, cell_range)

        # Format formula1 cleanly for 'list' validations (both Python lists and comma-separated strings)
        f1_str: Optional[str] = None
        if formula1 is not None:
            if vtype == "list" and isinstance(formula1, (list, tuple)):
                joined = ",".join(str(item).replace('"', "") for item in formula1)
                f1_str = f'"{joined}"'
            else:
                f1_str = str(formula1).strip()
                if f1_str.startswith("="):
                    f1_str = self._normalize_formula_sheet_refs(wb, f1_str)
                elif vtype == "list" and f1_str and not (f1_str.startswith('"') and f1_str.endswith('"')):
                    # Excel OOXML requires literal dropdown lists to be double-quoted ("Base,Upside,Downside");
                    # leave range references (e.g. "Sheet1!A1:A3" or "$A$1:$A$3") unquoted.
                    looks_like_range = ("!" in f1_str) or (":" in f1_str and "," not in f1_str)
                    if not looks_like_range:
                        f1_str = f'"{f1_str.replace(chr(34), "")}"'

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
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id = entry["metadata"]["workbook_id"]
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
