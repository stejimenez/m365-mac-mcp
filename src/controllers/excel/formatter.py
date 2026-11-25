"""Wall Street auto-coloring, cell/worksheet/batch formatting, named styles, and auto-fit widths."""

import re
from copy import copy
from typing import Any, Dict, List, Optional, Tuple, Union

from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.utils.cell import coordinate_from_string

from controllers.excel.constants import _NAMED_CELL_STYLE_PRESETS
from utils.logger import setup_logger
from utils.sheet_layout import (
    FACTOR_KEYWORDS,
    INT_COUNT_KEYWORDS,
    MULT_KEYWORDS,
    PAYBACK_KEYWORDS,
    PCT_KEYWORDS,
    analyze_worksheet_layout,
    build_magnitude_spec,
    is_pct_label,
    is_per_unit_label,
    resolve_workbook_magnitude,
)
from utils.validators import sanitize_for_json, validate_cell_ref, validate_range_ref

logger = setup_logger(__name__)


class ExcelFormatterMixin:
    """Wall Street auto-coloring, cell/worksheet/batch formatting, named styles, and auto-fit widths."""

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
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id: Optional[str] = None,
        operations: Optional[List[Dict[str, Any]]] = None,
    ) -> Dict[str, Any]:
        """Batch-format multiple worksheets in a single call and save once.

        Each item in operations may contain:
          sheet_name, column_widths, row_heights, freeze_panes, auto_fit_columns,
          min_col_width, max_col_width, col_padding, styles, merge_ranges, show_gridlines.
        """
        if not operations:
            raise ValueError("operations must be a non-empty list of worksheet format specs")

        workbook_id = self._resolve_workbook_id(workbook_id)
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        sheet_summaries: List[Dict[str, Any]] = []
        op_warnings: List[str] = []
        total_cells = 0

        known_op_keys = {
            "sheet_name", "sheet", "worksheet", "worksheet_name", "tab",
            "column_widths", "column_width", "row_heights", "freeze_panes",
            "auto_fit_columns", "min_col_width", "max_col_width", "col_padding",
            "auto_color_financial_model", "merge_ranges", "show_gridlines",
            "styles", "rules", "cell_styles", "formats", "range_styles",
            "style_rules", "cell_formats", "formatting_rules", "number_formats",
        }
        style_list_keys = (
            "styles", "rules", "cell_styles", "formats",
            "range_styles", "style_rules", "cell_formats", "formatting_rules",
        )
        wb_obj = entry["workbook_object"]
        all_sheet_names = list(wb_obj.sheetnames)
        sheet_level_flags = (
            "auto_color_financial_model", "auto_fit_columns", "freeze_panes", "show_gridlines",
        )
        range_specific_keys = (
            "range", "cells", "cell_range", "range_ref",
            "column_widths", "column_width", "row_heights", "merge_ranges", "number_formats",
        ) + style_list_keys

        # Pre-normalize universal aliases (e.g. {"operation": "auto_color_financial_model", "sheet_name": None})
        # and expand workbook-wide operations (sheet_name=None/"*"/"all") across all worksheets.
        expanded_ops: List[Dict[str, Any]] = []
        for raw_op in operations:
            if not isinstance(raw_op, dict):
                raise ValueError("Each item in operations must be a JSON object/dict")
            op = dict(raw_op)

            # Universal alias: {"operation": "auto_color_financial_model" | "auto_fit_columns" | "freeze_panes", ...}
            op_alias = op.pop("operation", None) or op.pop("action", None)
            if isinstance(op_alias, str) and op_alias.strip():
                alias_key = op_alias.strip().lower().replace("-", "_").replace(" ", "_")
                if alias_key in ("auto_color_financial_model", "auto_color", "wall_street_format"):
                    op["auto_color_financial_model"] = True
                elif alias_key in ("auto_fit_columns", "autofit", "auto_fit"):
                    op["auto_fit_columns"] = True
                elif alias_key in ("freeze_panes", "freeze"):
                    raw_sn = op.get("sheet_name")
                    if (
                        isinstance(raw_sn, str)
                        and raw_sn not in all_sheet_names
                        and re.match(r"^[A-Za-z]{1,3}\d+$", raw_sn.strip())
                    ):
                        op["freeze_panes"] = raw_sn.strip()
                        op["sheet_name"] = None
                    else:
                        op["freeze_panes"] = str(op.pop("cell", None) or op.pop("value", None) or "auto")
                else:
                    op["operation"] = op_alias  # keep unknown alias so warning surfaces cleanly

            req_sheet = self._extract_batch_sheet_name(op)
            # Extract sheet prefix from embedded 'Sheet!A1:B2' range if sheet_name was omitted
            if not req_sheet:
                for rk in ("range", "cells", "cell_range", "range_ref"):
                    rv = op.get(rk)
                    if isinstance(rv, str) and "!" in rv and "," not in rv:
                        sp_raw, cell_part = rv.rsplit("!", 1)
                        sp = sp_raw.strip()
                        if sp.startswith("'") and sp.endswith("'") and len(sp) >= 2:
                            sp = sp[1:-1].replace("''", "'")
                        if sp in all_sheet_names:
                            req_sheet = sp
                            op["sheet_name"] = sp
                            op[rk] = cell_part.strip()
                            break

            has_range_specific = any(op.get(k) is not None for k in range_specific_keys)
            has_sheet_flags = any(op.get(k) is not None for k in sheet_level_flags)

            if req_sheet in (None, "", "*", "all", "ALL"):
                if has_range_specific and len(all_sheet_names) > 1:
                    raise ValueError(
                        f"Multi-sheet workbook requires explicit 'sheet_name' on range formatting ops "
                        f"(available sheets: {all_sheet_names}). Omit 'sheet_name' ONLY for workbook-wide "
                        f"flags like {{'auto_color_financial_model': true, 'auto_fit_columns': true}}."
                    )
                if has_sheet_flags and not has_range_specific:
                    for sname in all_sheet_names:
                        sub = dict(op)
                        sub["sheet_name"] = sname
                        expanded_ops.append(sub)
                    continue
            expanded_ops.append(op)

        for op in expanded_ops:
            ws = self._get_sheet(entry, self._extract_batch_sheet_name(op))
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
                if any(k not in ("range", "cells", "cell_range", "range_ref") for k in flat_rule):
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
            raw_col_widths = op.get("column_widths")
            if raw_col_widths is None and op.get("column_width") is not None:
                # Universal alias: singular {"column_width": 15, "range": "B:F"} or {"column_width": 15, "range": "B1:F5"}
                single_w = float(op["column_width"])
                if single_w <= 0:
                    raise ValueError("column_width must be positive")
                range_hint = (
                    op.get("range")
                    or op.get("cells")
                    or op.get("cell_range")
                    or op.get("range_ref")
                )
                if isinstance(range_hint, str) and range_hint.strip():
                    col_spec_clean = re.sub(r"\d+", "", range_hint.split(",")[0].strip())
                    if col_spec_clean:
                        raw_col_widths = {col_spec_clean: single_w}

            # Default auto_fit_columns=True when auto_color_financial_model=True and no explicit column_widths were given
            explicit_auto_fit = op.get("auto_fit_columns")
            should_auto_fit = (
                bool(explicit_auto_fit)
                if explicit_auto_fit is not None
                else bool(op.get("auto_color_financial_model") and raw_col_widths is None)
            )
            if should_auto_fit:
                fitted = self._compute_auto_fit_widths(
                    ws,
                    min_width=float(op.get("min_col_width", 12.0)),
                    max_width=float(op.get("max_col_width", 60.0)),
                    padding=float(op.get("col_padding", 5.0)),
                )
                for col_letter, width in fitted.items():
                    ws.column_dimensions[col_letter].width = width
                    applied_widths[col_letter] = width

            if isinstance(raw_col_widths, list):
                start_col_idx = 1
                range_hint = (
                    op.get("range")
                    or op.get("cells")
                    or op.get("cell_range")
                    or op.get("range_ref")
                )
                if isinstance(range_hint, str) and range_hint.strip():
                    try:
                        first_ref = range_hint.split(",")[0].strip().split(":")[0].strip()
                        col_l, _ = coordinate_from_string(validate_cell_ref(first_ref))
                        start_col_idx = column_index_from_string(col_l)
                    except Exception:
                        start_col_idx = 1
                raw_col_widths = {
                    get_column_letter(start_col_idx + idx): width
                    for idx, width in enumerate(raw_col_widths)
                    if width is not None
                }
            elif raw_col_widths is not None and not isinstance(raw_col_widths, dict):
                raise ValueError(
                    "column_widths must be a mapping like {'A': 22, 'B:F': 14} "
                    "or a list of positive numbers like [22, 14, 14]."
                )

            for col_spec, raw_width in (raw_col_widths or {}).items():
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
                fp = str(op["freeze_panes"]).strip()
                if fp == "":
                    ws.freeze_panes = None
                elif fp.lower() == "auto":
                    lay_fp = analyze_worksheet_layout(ws)
                    ws.freeze_panes = lay_fp.recommended_freeze_cell or "B2"
                else:
                    ws.freeze_panes = validate_cell_ref(fp)

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

    async def list_cell_styles(self, workbook_id: str) -> Dict[str, Any]:
        """List workbook named styles and built-in Excel financial-model style presets."""
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id = entry["metadata"]["workbook_id"]
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
            # Use sparse populated max_row to avoid formatting 1,048,576 empty ghost rows
            cells_dict = getattr(ws, "_cells", None)
            if isinstance(cells_dict, dict) and cells_dict:
                pop_rows = [
                    r for (r, _c), cobj in cells_dict.items()
                    if cobj is not None and cobj.value is not None
                ]
                max_row = max(pop_rows) if pop_rows else min(max(ws.max_row or 1, 1), 100)
            else:
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
        """Estimate display width per populated column via sparse ws._cells (O(populated), ghost-row safe)."""
        widths: Dict[str, float] = {}
        lay = analyze_worksheet_layout(ws)
        overflow_cells = lay.overflow_banner_cells
        best_by_col: Dict[int, float] = {}

        cells_dict = getattr(ws, "_cells", None)
        if isinstance(cells_dict, dict):
            cell_iter = (
                (r, c, cobj)
                for (r, c), cobj in list(cells_dict.items())
                if cobj is not None and cobj.value is not None
            )
        else:
            max_col = ws.max_column or 0
            max_row = ws.max_row or 0
            cell_iter = (
                (r, c, ws.cell(row=r, column=c))
                for c in range(1, max_col + 1)
                for r in range(1, max_row + 1)
            )

        for row_idx, col_idx, cell in cell_iter:
            if (row_idx, col_idx) in overflow_cells:
                continue
            val = cell.value
            if val is None:
                continue
            if isinstance(val, str) and val.startswith("="):
                # Formulas evaluate to numbers/text; reserve comfortable width for formatted values
                num_fmt = cell.number_format or ""
                est = 15.0 if any(ch in num_fmt for ch in ("$", "€", "£", "#", "0", "%")) else 12.0
            elif isinstance(val, (int, float)) and not isinstance(val, bool):
                num_fmt = cell.number_format or ""
                if "%" in num_fmt:
                    est = float(len(f"{val * 100:.1f}%")) + 1.0
                elif any(ch in num_fmt for ch in ("$", "€", "£", ",")):
                    # Account for Option-A trailing-comma scaling (e.g. $#,##0, divides display by 10^3n)
                    first_sec = num_fmt.split(";")[0]
                    trailing_commas = len(first_sec) - len(first_sec.rstrip(","))
                    disp_val = abs(val) / (1000 ** trailing_commas) if trailing_commas > 0 else abs(val)
                    est = float(len(f"$({disp_val:,.0f})")) if ".00" not in first_sec else float(len(f"$({disp_val:,.2f})"))
                else:
                    est = float(len(str(val)))
            else:
                lines = str(val).splitlines() or [""]
                est = float(max(len(line) for line in lines))
            if cell.font and cell.font.bold:
                est *= 1.1
            if est > best_by_col.get(col_idx, 0.0):
                best_by_col[col_idx] = est

        for col_idx in sorted(best_by_col):
            best = best_by_col[col_idx]
            if best > 0:
                col_letter = get_column_letter(col_idx)
                widths[col_letter] = round(max(min_width, min(max_width, best + padding)), 2)
        return widths

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
        """Apply layout-agnostic institutional Wall Street styling across arbitrary multi-table sheets.

        Powered by IMP-003 (`analyze_worksheet_layout`) and IMP-001 (Option-A Smart Magnitude):
          - NEVER assumes headers are on Row 1 or row labels are in Column A.
          - Supports stacked & side-by-side multi-table layouts, indented Wall Street sheets,
            4-line institutional title blocks, section banners, mid-table sub-headers, 2-way
            sensitivity matrices, and multilingual formula-AST + keyword subtotal/grand-total tiers.
          - Applies Option-A trailing-comma currency scaling ($#,##0, / $#,##0,,) ONLY when an
            institutional magnitude header is present (syncing Row 4 automatically so cell display
            and header declaration always match 100%; Assumptions tabs remain in exact base currency
            with transparent Row 4 labeling; per-share/per-unit rows stay in exact $#,##0.00).
        Runs BEFORE explicit caller `styles`, so explicit overrides always win.
        """
        colored = 0
        max_r = ws.max_row or 0
        max_c = ws.max_column or 0
        if max_r == 0 or max_c == 0:
            return 0

        wb = getattr(ws, "parent", None)
        header_meta = getattr(wb, "_m365_header_metadata", None) if wb is not None else None
        all_layouts = (
            {s.title: analyze_worksheet_layout(s) for s in wb.worksheets}
            if wb is not None
            else {ws.title: analyze_worksheet_layout(ws)}
        )
        layout = all_layouts[ws.title]
        mag_spec = (
            resolve_workbook_magnitude(wb, layouts=all_layouts, header_metadata=header_meta)
            if wb is not None
            else build_magnitude_spec(currency="USD", scale_name="exact", display_scaled=False)
        )

        # Golden Safety Rule: Option-A trailing-comma scaling ($#,##0, / $#,##0,,) is applied ONLY
        # when the workbook/sheet visibly declares that magnitude in its institutional header block
        # (never silently divide numbers on an unheaded legacy 'Line Item ($)' sheet!).
        has_institutional_header = bool(
            header_meta is not None
            or layout.declared_magnitude is not None
            or any(l.declared_magnitude is not None for l in all_layouts.values())
        )
        non_monetary_sheets = set((header_meta or {}).get("non_monetary_sheets") or [])
        is_non_monetary_sheet = (ws.title or "").strip().lower() in non_monetary_sheets
        title_low = (ws.title or "").lower()
        is_assumptions_sheet = any(kw in title_low for kw in ("assumption", "input", "driver"))

        exact_currency_fmt = f'{mag_spec.currency_symbol}#,##0;({mag_spec.currency_symbol}#,##0);"-"'
        stmt_currency_fmt = (
            mag_spec.currency_format
            if (has_institutional_header and not is_non_monetary_sheet)
            else (
                mag_spec.currency_format
                if not mag_spec.display_scaled
                else exact_currency_fmt
            )
        )

        # 1. Style 4-line institutional TITLE_BLOCK (and synchronize Row 4 magnitude label when auto-resolved)
        if layout.title_block is not None:
            tb = layout.title_block
            tcol = tb.min_col
            tb_span_max_col = max(
                [b.max_col for b in layout.data_blocks()] or [max_c]
            )
            tb_rows = list(range(tb.min_row, tb.max_row + 1))

            # Sync Row 4 (magnitude line) automatically when created via header_metadata(magnitude='auto')
            if header_meta is not None and not is_non_monetary_sheet and len(tb_rows) >= 4:
                mag_row = tb.min_row + 3
                if is_assumptions_sheet:
                    ws.cell(row=mag_row, column=tcol).value = (
                        f"Input Basis: Exact {mag_spec.currency_code} {mag_spec.currency_symbol} "
                        f"(Statements Presented {mag_spec.header_label})"
                    )
                else:
                    ws.cell(row=mag_row, column=tcol).value = mag_spec.header_label

            for idx, r in enumerate(tb_rows):
                c_obj = ws.cell(row=r, column=tcol)
                if idx == 0:
                    self._apply_cell_formatting(c_obj, {"bold": True, "font_size": 14, "font_color": "#1F4E79", "horizontal": "left"})
                elif idx == 1:
                    self._apply_cell_formatting(c_obj, {"bold": True, "font_size": 11, "font_color": "#1F2937", "horizontal": "left"})
                elif idx == 2:
                    self._apply_cell_formatting(c_obj, {"italic": True, "font_size": 10, "font_color": "#4B5563", "horizontal": "left"})
                else:
                    self._apply_cell_formatting(c_obj, {"bold": True, "italic": True, "font_size": 10, "font_color": "#1F4E79", "horizontal": "left"})
                colored += 1
            # Clean thin bottom rule beneath the last line of the institutional title block
            last_tb_r = tb_rows[-1]
            for c in range(tcol, tb_span_max_col + 1):
                self._apply_cell_formatting(ws.cell(row=last_tb_r, column=c), {"border": {"bottom": "thin", "color": "#1F4E79"}})
            if last_tb_r + 1 <= max_r and all(ws.cell(row=last_tb_r + 1, column=c).value is None for c in range(1, max_c + 1)):
                ws.row_dimensions[last_tb_r + 1].height = 8

        # 2. Style SECTION_BANNER blocks (e.g. 'Table 2: WACC vs. Exit Multiple')
        for blk in layout.blocks:
            if blk.kind != "SECTION_BANNER":
                continue
            for r in range(blk.min_row, blk.max_row + 1):
                c_obj = ws.cell(row=r, column=blk.min_col)
                self._apply_cell_formatting(c_obj, {"bold": True, "font_size": 11, "font_color": "#1F4E79", "horizontal": "left"})
                colored += 1

        # 3. Style every MATRIX_TABLE (2-way sensitivity tables at ANY row/col on ANY sheet)
        for blk in layout.blocks:
            if blk.kind != "MATRIX_TABLE":
                continue
            hr = blk.matrix_header_row or blk.min_row
            ac = blk.matrix_axis_col or blk.min_col
            corner = ws.cell(row=hr, column=ac)
            if corner.value is not None:
                self._apply_cell_formatting(corner, {"bold": True, "font_color": "#1F4E79", "bg_color": "#F2F4F8"})
                colored += 1
            for c in range(ac + 1, blk.max_col + 1):
                hc = ws.cell(row=hr, column=c)
                if hc.value is None:
                    continue
                self._apply_cell_formatting(hc, {"bold": True, "font_color": "0000FF", "bg_color": "#F2F4F8", "horizontal": "center"})
                cur_nf = (hc.number_format or "").strip()
                if (not cur_nf or cur_nf == "General") and isinstance(hc.value, (int, float)) and not isinstance(hc.value, bool):
                    hc.number_format = "0.0%" if 0 < abs(float(hc.value)) <= 2.5 else '0.0"x"'
                colored += 1
            for r in range(hr + 1, blk.max_row + 1):
                lc = ws.cell(row=r, column=ac)
                if lc.value is not None:
                    self._apply_cell_formatting(lc, {"bold": True, "font_color": "0000FF", "bg_color": "#F2F4F8", "horizontal": "right"})
                    cur_nf = (lc.number_format or "").strip()
                    if (not cur_nf or cur_nf == "General") and isinstance(lc.value, (int, float)) and not isinstance(lc.value, bool):
                        lc.number_format = "0.0%" if 0 < abs(float(lc.value)) <= 2.5 else '0.0"x"'
                    colored += 1
                for c in range(ac + 1, blk.max_col + 1):
                    cell_obj = ws.cell(row=r, column=c)
                    val = cell_obj.value
                    if val is None:
                        continue
                    is_fml = isinstance(val, str) and val.startswith("=")
                    is_num = isinstance(val, (int, float)) and not isinstance(val, bool)
                    if is_fml:
                        self._apply_cell_formatting(cell_obj, {"font_color": "008000" if "!" in val else "000000"})
                        colored += 1
                    elif is_num:
                        self._apply_cell_formatting(cell_obj, {"font_color": "0000FF"})
                        colored += 1
                    cur_nf = (cell_obj.number_format or "").strip()
                    if (is_fml or is_num) and (not cur_nf or cur_nf == "General"):
                        cell_obj.number_format = stmt_currency_fmt

        # 4. Style every DATA_TABLE independently at its own (min_row..max_row, min_col..max_col)
        for blk in layout.blocks:
            if blk.kind != "DATA_TABLE":
                continue
            hdr_set = set(blk.header_rows)
            blk_currency_fmt = exact_currency_fmt if (blk.is_assumptions_block or is_assumptions_sheet) else stmt_currency_fmt

            # 4a. Style Table Header Row(s) — navy white bold, integer '0' on numeric year/period headers
            for hr in blk.header_rows:
                row_pop = sum(1 for c in range(blk.min_col, blk.max_col + 1) if ws.cell(row=hr, column=c).value is not None)
                for c in range(blk.min_col, blk.max_col + 1):
                    hcell = ws.cell(row=hr, column=c)
                    if row_pop >= 2 and hcell.value is None:
                        continue
                    if not hcell.fill or hcell.fill.fill_type is None:
                        self._apply_cell_formatting(
                            hcell,
                            {
                                "bold": True,
                                "font_color": "#FFFFFF",
                                "bg_color": "#1F4E79",
                                "horizontal": "center" if (row_pop >= 2 and c >= blk.value_start_col) else "left",
                            },
                        )
                        if isinstance(hcell.value, (int, float)) and not isinstance(hcell.value, bool):
                            cur_nf = (hcell.number_format or "").strip()
                            if not cur_nf or cur_nf == "General":
                                hcell.number_format = "0"
                        colored += 1

            # 4b. Style Body Rows (SUBHEADER, DETAIL, SUBTOTAL, GRAND TOTAL)
            for r in range(blk.min_row, blk.max_row + 1):
                if r in hdr_set:
                    continue

                if r in blk.subheader_rows:
                    lbl_cell = ws.cell(row=r, column=blk.label_col)
                    if not lbl_cell.fill or lbl_cell.fill.fill_type is None:
                        self._apply_cell_formatting(
                            lbl_cell,
                            {"bold": True, "font_color": "#1F4E79", "bg_color": "#F2F4F8"},
                        )
                        colored += 1
                    continue

                row_label_raw = str(ws.cell(row=r, column=blk.label_col).value or "").strip()
                row_label = row_label_raw.lower()
                # Strip parenthetical method qualifiers (e.g. 'Enterprise Value (Gordon Growth)', 'Equity Value (Exit Multiple)')
                # so '(Gordon Growth)' / '(Exit Multiple)' never override currency or per-share formatting.
                row_label_core = re.sub(r"\([^)]*\)", "", row_label).strip() or row_label
                is_pct_row = is_pct_label(row_label_core)
                is_grand_total = r in blk.total_rows
                is_subtotal = r in blk.subtotal_rows

                # Determine row-level smart number format (with per-share / per-unit exact-currency protection)
                if is_per_unit_label(row_label_core):
                    row_num_fmt: Optional[str] = mag_spec.per_unit_format
                elif is_pct_row:
                    row_num_fmt = "0.0%"
                elif any(kw in row_label_core for kw in MULT_KEYWORDS):
                    row_num_fmt = '0.00"x"'
                elif any(kw in row_label_core for kw in FACTOR_KEYWORDS):
                    row_num_fmt = "0.000"
                elif any(kw in row_label_core for kw in PAYBACK_KEYWORDS):
                    row_num_fmt = "0.00"
                elif any(kw in row_label_core for kw in INT_COUNT_KEYWORDS):
                    row_num_fmt = "#,##0"
                else:
                    row_num_fmt = blk_currency_fmt

                tier_style: Optional[Dict[str, Any]] = None
                if is_grand_total:
                    tier_style = {
                        "bold": True,
                        "bg_color": "#D9E1F2",
                        "border": {"top": "thin", "bottom": "double"},
                    }
                elif is_subtotal:
                    tier_style = {
                        "bold": True,
                        "bg_color": "#EDF2F8",
                        "border": {"top": "thin", "bottom": "thin"},
                    }

                if tier_style is not None:
                    label_tier_style = {**tier_style, "font_color": "#000000"}
                    for lc in range(blk.min_col, blk.label_col + 1):
                        lbl_cell = ws.cell(row=r, column=lc)
                        if not lbl_cell.fill or lbl_cell.fill.fill_type is None:
                            self._apply_cell_formatting(lbl_cell, label_tier_style)

                for c in range(blk.value_start_col, blk.max_col + 1):
                    cell_obj = ws.cell(row=r, column=c)
                    val = cell_obj.value
                    if val is None:
                        continue
                    is_fml = isinstance(val, str) and val.startswith("=")
                    is_num = isinstance(val, (int, float)) and not isinstance(val, bool)

                    if is_fml:
                        color = "008000" if "!" in val else "000000"
                        fmt_dict: Dict[str, Any] = {"font_color": color}
                        if tier_style is not None and (not cell_obj.fill or cell_obj.fill.fill_type is None):
                            fmt_dict.update(tier_style)
                        self._apply_cell_formatting(cell_obj, fmt_dict)
                        colored += 1
                    elif is_num:
                        fmt_dict = {"font_color": "0000FF"}
                        if tier_style is not None and (not cell_obj.fill or cell_obj.fill.fill_type is None):
                            fmt_dict.update(tier_style)
                        self._apply_cell_formatting(cell_obj, fmt_dict)
                        colored += 1

                    # Apply smart number format ONLY if cell is still unformatted ('General' / empty)
                    cur_nf = (cell_obj.number_format or "").strip()
                    if (is_fml or is_num) and (not cur_nf or cur_nf == "General"):
                        if row_num_fmt:
                            cell_obj.number_format = row_num_fmt

        # Institutional print/PDF setup: fit all columns cleanly to 1 landscape page wide
        try:
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
        except Exception:
            pass

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
