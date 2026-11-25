"""Cell/range/batch writers, formula guard, coercion, and sheet-reference normalization."""

import os
import re
from typing import (
    Any,
    Dict,
    List,
    Optional,
    Tuple,
)

from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

from utils.logger import setup_logger
from utils.validators import (
    sanitize_excel_formula_quotes,
    sanitize_excel_write_value,
    sanitize_for_json,
    validate_cell_ref,
)

logger = setup_logger(__name__)


class ExcelWriterMixin:
    """Cell/range/batch writers, formula guard, coercion, and sheet-reference normalization."""

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
        logger.warning(
            f"Unknown M365_MCP_FORMULA_GUARD='{val}'. Falling back to 'literals_only'. "
            "Valid values: literals_only, strict, off."
        )
        return "literals_only"

    @staticmethod
    def _formula_guard_response(mode: str, blocked: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Return a compact, structured confirmation_required payload when formula writes are blocked."""
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

    _CURRENCY_STR_RE = re.compile(
        r"^(?P<neg1>-)?(?P<lparen>\()?\s*\$\s*(?P<neg2>-)?(?P<num>(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s*(?P<rparen>\))?$"
    )

    _PERCENT_STR_RE = re.compile(
        r"^(?P<num>[+-]?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?)%$"
    )

    _PLAIN_NUM_STR_RE = re.compile(r"^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?$")

    _MISSING_EQ_FORMULA_RE = re.compile(
        r"^(?:IF|IFERROR|SUM|SUMIF|SUMIFS|SUMPRODUCT|NPV|IRR|MIRR|XNPV|XIRR|"
        r"AVERAGE|MIN|MAX|COUNT|COUNTIF|COUNTIFS|INDEX|MATCH|VLOOKUP|XLOOKUP|"
        r"CHOOSE|OFFSET|PV|FV|PMT|RATE)\s*\(.*\)$",
        re.IGNORECASE,
    )
    _CELL_OR_SHEET_REF_RE = re.compile(
        r"(?:'[^']+'!|[A-Za-z_][A-Za-z0-9_]*!|\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+|\$[A-Z]{1,3}\$?\d+)"
    )
    _MISSING_EQ_SHEET_LINK_RE = re.compile(
        r"^[+-]?\s*(?:'[^']+'!|[A-Za-z_][A-Za-z0-9_]*!)\$?[A-Z]{1,3}\$?\d+$"
    )

    @classmethod
    def _coerce_excel_write_value(cls, raw_value: Any) -> Tuple[Any, Optional[str]]:
        """Normalize empty strings to blank cells and parse unambiguous formatted numeric strings.

        Returns ``(coerced_value, inferred_number_format)``.
        - ``""`` -> ``(None, None)`` so blank cells never trigger ``#VALUE!`` in Excel arithmetic.
        - Unambiguous Excel formula strings missing leading ``=`` (e.g. ``IF(MAX('Sheet'!B1:H1)<0,...)``)
          are auto-prefixed with ``=`` and quote-sanitized.
        - ``"$25,000,000"`` / ``"($4,000,000)"`` -> numeric value + currency format.
        - ``"10.5%"`` / ``"12%"`` -> decimal float (``0.105`` / ``0.12``) + percent format.
        - ``"6"`` / ``"0"`` -> numeric int/float (leading-zero codes like ``"00123"`` stay text).
        - Formulas (``=...``), text with units (``"$510M"``), and literal ``'""'`` stay untouched.
        """
        val = sanitize_excel_write_value(raw_value)
        if not isinstance(val, str):
            return val, None
        if val.startswith("="):
            return val, None
        if val == "":
            return None, None

        s = val.strip()
        if not s:
            return None, None

        # Auto-repair unambiguous Excel formula strings where the caller omitted the leading '='
        if (
            cls._MISSING_EQ_FORMULA_RE.match(s)
            and cls._CELL_OR_SHEET_REF_RE.search(s)
            and s.count("(") == s.count(")")
        ) or cls._MISSING_EQ_SHEET_LINK_RE.match(s):
            return sanitize_excel_formula_quotes("=" + s), None

        m_curr = cls._CURRENCY_STR_RE.match(s)
        if m_curr:
            has_lp = bool(m_curr.group("lparen"))
            has_rp = bool(m_curr.group("rparen"))
            if has_lp == has_rp:
                num_txt = m_curr.group("num").replace(",", "")
                is_neg = bool(m_curr.group("neg1") or m_curr.group("neg2") or (has_lp and has_rp))
                has_dec = "." in num_txt
                parsed: Any = float(num_txt) if has_dec else int(num_txt)
                if is_neg:
                    parsed = -parsed
                fmt = '$#,##0.00;($#,##0.00);"-"' if has_dec else '$#,##0;($#,##0);"-"'
                return parsed, fmt

        m_pct = cls._PERCENT_STR_RE.match(s)
        if m_pct:
            num_txt = m_pct.group("num").replace(",", "")
            parsed_pct = round(float(num_txt) / 100.0, 10)
            dec_part = num_txt.split(".", 1)[1] if "." in num_txt else ""
            fmt = "0.00%" if len(dec_part) >= 2 else "0.0%"
            return parsed_pct, fmt

        if cls._PLAIN_NUM_STR_RE.match(s):
            return (float(s) if "." in s else int(s)), None

        return val, None

    @staticmethod
    def _build_compact_cell_map(
        ws,
        start_row: int,
        start_col: int,
        num_rows: int,
        num_cols: int,
    ) -> Optional[Dict[str, str]]:
        """Build a compact ground-truth coordinate map for tables up to 60 rows (prevents off-by-1 row errors)."""
        if num_rows < 2 or num_rows > 60 or num_cols < 2 or num_cols > 8:
            return None
        val_col_letter = get_column_letter(start_col + 1)
        is_kv_table = num_cols <= 3
        cmap: Dict[str, str] = {}
        for r in range(start_row, start_row + num_rows):
            lbl = ws.cell(row=r, column=start_col).value
            v = ws.cell(row=r, column=start_col + 1).value
            if not isinstance(lbl, str) or not lbl.strip():
                continue
            lbl_low = lbl.strip().lower()
            if r == start_row and lbl_low in ("parameter", "item", "assumption", "metric", "kpi", "driver", "field", "name", "input", "line item"):
                continue
            short_lbl = lbl.strip()[:36]
            if is_kv_table:
                if v is None:
                    continue
                if isinstance(v, str) and v.startswith("="):
                    cmap[f"{val_col_letter}{r}"] = f"{short_lbl}: <formula>"
                else:
                    val_preview = str(v)[:32]
                    cmap[f"{val_col_letter}{r}"] = f"{short_lbl} = {val_preview}"
            else:
                # Multi-column financial statement / schedule: return compact row-number anchor
                if any(ws.cell(row=r, column=c).value is not None for c in range(start_col + 1, start_col + num_cols)):
                    cmap[f"R{r}"] = short_lbl
        return cmap if len(cmap) >= 2 else None

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
        workbook_id = entry["metadata"]["workbook_id"]
        ws = self._get_sheet(entry, sheet_name)
        clean_cell = validate_cell_ref(cell)

        wb = entry["workbook_object"]
        clean_value, inferred_fmt = self._coerce_excel_write_value(value)
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
        elif inferred_fmt and (not cell_obj.number_format or cell_obj.number_format == "General"):
            cell_obj.number_format = inferred_fmt
        if formatting:
            self._apply_cell_formatting(cell_obj, formatting)

        await self._save_working_copy(entry)

        result: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "sheet_name": ws.title,
            "cell": clean_cell,
            "value": clean_value,
            "value_type": "formula" if is_new_formula else type(clean_value).__name__,
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
        workbook_id = entry["metadata"]["workbook_id"]
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
                        clean_val, _ = self._coerce_excel_write_value(raw_val)
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
                clean_val, inferred_fmt = self._coerce_excel_write_value(raw_val)
                if isinstance(clean_val, str) and clean_val.startswith("="):
                    clean_val = self._normalize_formula_sheet_refs(wb, clean_val)
                    formula_warnings.extend(
                        self._inspect_formula_sheet_refs(wb, clean_val, cell_ref=cell_obj.coordinate)
                    )
                cell_obj.value = clean_val
                if number_format and (row_idx > 0 or not header_formatting):
                    cell_obj.number_format = number_format
                elif inferred_fmt and (not cell_obj.number_format or cell_obj.number_format == "General"):
                    cell_obj.number_format = inferred_fmt
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

    @staticmethod
    def _extract_batch_sheet_name(
        op: Dict[str, Any], default_sheet_name: Optional[str] = None
    ) -> Optional[str]:
        """Resolve sheet_name from explicit aliases, A1 'Sheet!Cell' prefix, or batch default."""
        explicit = (
            op.get("sheet_name")
            or op.get("sheet")
            or op.get("worksheet")
            or op.get("worksheet_name")
            or op.get("tab")
        )
        if explicit:
            return str(explicit).strip()
        for key in ("start_cell", "range", "cell_range", "cell"):
            raw = op.get(key)
            if raw and "!" in str(raw):
                prefix = str(raw).rsplit("!", 1)[0].strip()
                if prefix.startswith("'") and prefix.endswith("'") and len(prefix) >= 2:
                    prefix = prefix[1:-1].replace("''", "'")
                if prefix:
                    return prefix
        return default_sheet_name

    @staticmethod
    def _extract_batch_start_cell(op: Dict[str, Any]) -> str:
        """Resolve start_cell from LLM-friendly aliases (start_cell, range, cell_range, cell)."""
        raw = (
            op.get("start_cell")
            or op.get("range")
            or op.get("cell_range")
            or op.get("cell")
            or "A1"
        )
        text = str(raw).strip()
        if "!" in text:
            text = text.rsplit("!", 1)[1].strip()
        if ":" in text:
            text = text.split(":", 1)[0].strip()
        return text or "A1"

    @staticmethod
    def _normalize_batch_data(op: Dict[str, Any], raw_data: Any) -> List[List[Any]]:
        """Normalize flat 1D lists and auto-orient 1xN / Nx1 matrices to match an explicit range shape."""
        if not isinstance(raw_data, list) or not raw_data:
            sheet_label = op.get("sheet_name") or op.get("sheet")
            format_only_keys = {
                "number_format", "number_formats", "formatting", "format",
                "cell_formats", "cell_format", "column_widths", "row_heights",
                "styles", "bold", "italic", "fill_color", "font_color", "freeze_panes",
            }
            if any(k in op for k in format_only_keys):
                raise ValueError(
                    f"Formatting-only operation detected (sheet={sheet_label}): "
                    "call batch_format_workbook (not batch_write_ranges) to apply number_format, styles, or column_widths."
                )
            raise ValueError(
                f"Each operation must include non-empty 'data', 'values', 'value', or 'fill_formula' (sheet={sheet_label})"
            )

        # Determine if caller passed an explicit A1:B2 range to infer intended 1xN vs Nx1 orientation
        raw_ref = str(
            op.get("range")
            or op.get("cell_range")
            or op.get("start_cell")
            or op.get("cell")
            or ""
        ).strip()
        if "!" in raw_ref:
            raw_ref = raw_ref.rsplit("!", 1)[1].strip()

        target_is_horizontal = False
        target_is_vertical = False
        if ":" in raw_ref:
            try:
                left_ref, right_ref = [p.strip() for p in raw_ref.split(":", 1)]
                c1_let, r1 = coordinate_from_string(validate_cell_ref(left_ref))
                c2_let, r2 = coordinate_from_string(validate_cell_ref(right_ref))
                c1, c2 = column_index_from_string(c1_let), column_index_from_string(c2_let)
                if r1 == r2 and c2 > c1:
                    target_is_horizontal = True
                elif c1 == c2 and r2 > r1:
                    target_is_vertical = True
            except Exception:
                pass

        # Flat 1D list of scalars -> wrap as horizontal row (or vertical column if target range is vertical)
        if not isinstance(raw_data[0], list):
            if target_is_vertical:
                return [[v] for v in raw_data]
            return [list(raw_data)]

        rows: List[List[Any]] = [list(r) if isinstance(r, list) else [r] for r in raw_data]

        # Auto-transpose Nx1 vertical data written into an explicit 1xN horizontal range (e.g., B3:F3)
        if target_is_horizontal and len(rows) > 1 and all(len(r) == 1 for r in rows):
            return [[r[0] for r in rows]]

        # Auto-transpose 1xN horizontal data written into an explicit Nx1 vertical range (e.g., A4:A10)
        if target_is_vertical and len(rows) == 1 and len(rows[0]) > 1:
            return [[v] for v in rows[0]]

        return rows

    @classmethod
    def _resolve_batch_op_data(cls, op: Dict[str, Any]) -> List[List[Any]]:
        """Resolve a batch_write_ranges operation into a 2D matrix, expanding Excel-style `fill_formula` across ranges."""
        raw_data = op.get("data") if op.get("data") is not None else op.get("values")
        fill_fml = op.get("fill_formula")
        if fill_fml is None and raw_data is None and op.get("formula"):
            fill_fml = op.get("formula")
        if fill_fml is None and raw_data is None and "value" in op:
            single_val = op["value"]
            if isinstance(single_val, list):
                raw_data = single_val
            else:
                return [[single_val]]
        if fill_fml is None and raw_data is not None and not isinstance(raw_data, list):
            return [[raw_data]]

        if fill_fml is not None:
            fml_str = str(fill_fml).strip()
            if not fml_str:
                raise ValueError(
                    "Invalid empty 'fill_formula': pass an Excel formula starting with '=' "
                    "(e.g. {'sheet_name': '...', 'range': 'C8:H8', 'fill_formula': '=C7-C9'})."
                )
            if not fml_str.startswith("=") and re.match(
                r"^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$", fml_str
            ):
                raise ValueError(
                    f"Invalid 'fill_formula' value '{fml_str}': 'fill_formula' must be an Excel formula starting with '=' "
                    f"(e.g. '=C7-C9'), NOT a cell range. Put literal labels/values and 'fill_formula' in SEPARATE operations: "
                    f"e.g. {{'sheet_name': '...', 'range': 'A8:B8', 'data': [['OpEx', 0]]}} and "
                    f"{{'sheet_name': '...', 'range': 'C8:H8', 'fill_formula': '=C7-C9'}}."
                )
            if raw_data is not None:
                raise ValueError(
                    "Do not combine 'data'/'values' and 'fill_formula' in the same operation — put them in SEPARATE operations: "
                    "e.g. {'sheet_name': '...', 'range': 'A8:B8', 'data': [['OpEx', 0]]} and "
                    "{'sheet_name': '...', 'range': 'C8:H8', 'fill_formula': '=C7-C9'}."
                )
            if not fml_str.startswith("="):
                fml_str = "=" + fml_str
            raw_ref = str(
                op.get("range")
                or op.get("cell_range")
                or op.get("start_cell")
                or op.get("cell")
                or ""
            ).strip()
            if "!" in raw_ref:
                raw_ref = raw_ref.rsplit("!", 1)[1].strip()
            if not raw_ref:
                raise ValueError("fill_formula requires 'range' (e.g. 'C4:H4' or 'C3:G9')")

            if ":" in raw_ref:
                left_ref, right_ref = [p.strip() for p in raw_ref.split(":", 1)]
                c1_let, r1 = coordinate_from_string(validate_cell_ref(left_ref))
                c2_let, r2 = coordinate_from_string(validate_cell_ref(right_ref))
                c1, c2 = column_index_from_string(c1_let), column_index_from_string(c2_let)
                if r2 < r1 or c2 < c1:
                    raise ValueError(f"Invalid fill_formula range '{raw_ref}': end cell precedes start cell")
            else:
                c1_let, r1 = coordinate_from_string(validate_cell_ref(raw_ref))
                c1 = c2 = column_index_from_string(c1_let)
                r2 = r1

            origin = f"{get_column_letter(c1)}{r1}"
            translator = Translator(fml_str, origin=origin)
            matrix: List[List[Any]] = []
            for r in range(r1, r2 + 1):
                row_vals: List[Any] = []
                for c in range(c1, c2 + 1):
                    target = f"{get_column_letter(c)}{r}"
                    row_vals.append(fml_str if target == origin else translator.translate_formula(target))
                matrix.append(row_vals)
            return matrix

        return cls._normalize_batch_data(op, raw_data)

    async def batch_write_ranges(
        self,
        workbook_id: Optional[str] = None,
        operations: Optional[List[Dict[str, Any]]] = None,
        allow_formula_overwrite: bool = False,
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Write multiple 2D ranges across one or more worksheets in a single call and save once."""
        if not operations:
            raise ValueError("operations must be a non-empty list of write specifications")

        workbook_id = self._resolve_workbook_id(workbook_id)
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        wb = entry["workbook_object"]
        default_sheet = sheet_name

        # Universal normalization:
        # 1. Expand invented plural 'fill_formulas' lists cleanly.
        # 2. Expand dict-shaped 'fill_formula': {'range': 'D2:H2', 'formula': '=C2*1.12'}
        #    (even when combined with 'value'/'data' on another range in the same op)
        #    so Python dict reprs are never written literally into Excel cells.
        normalized_ops: List[Dict[str, Any]] = []
        for op in operations:
            if isinstance(op, dict) and "fill_formulas" in op:
                ff_list = op.get("fill_formulas")
                if not isinstance(ff_list, list) or not ff_list or not all(isinstance(x, dict) for x in ff_list):
                    raise ValueError(
                        "Invalid 'fill_formulas' key: use singular 'fill_formula' per operation "
                        "(e.g. {'sheet_name': '...', 'range': 'C4:H4', 'fill_formula': '=C2*C3'})."
                    )
                parent_sheet = self._extract_batch_sheet_name(op, default_sheet_name=default_sheet)
                for item in ff_list:
                    expanded_op = dict(item)
                    if parent_sheet and not self._extract_batch_sheet_name(expanded_op):
                        expanded_op["sheet_name"] = parent_sheet
                    if "formula" in expanded_op and "fill_formula" not in expanded_op:
                        expanded_op["fill_formula"] = expanded_op.pop("formula")
                    normalized_ops.append(expanded_op)
            elif isinstance(op, dict) and isinstance(op.get("fill_formula"), dict):
                ff_dict = op["fill_formula"]
                fml_str = ff_dict.get("formula") or ff_dict.get("fill_formula") or ff_dict.get("expression")
                if not fml_str or not isinstance(fml_str, str):
                    raise ValueError(
                        "Dict 'fill_formula' must include a string 'formula' key "
                        "(e.g. {'range': 'C4:H4', 'fill_formula': '=C2*C3'})."
                    )
                parent_sheet = self._extract_batch_sheet_name(op, default_sheet_name=default_sheet)
                has_payload = any(op.get(k) is not None for k in ("data", "values", "value"))
                if has_payload:
                    val_op = {k: v for k, v in op.items() if k != "fill_formula"}
                    normalized_ops.append(val_op)
                ff_op: Dict[str, Any] = {
                    "range": ff_dict.get("range") or ff_dict.get("cell_range") or op.get("range") or op.get("cell_range") or op.get("start_cell") or op.get("cell"),
                    "fill_formula": fml_str,
                }
                if parent_sheet:
                    ff_op["sheet_name"] = parent_sheet
                normalized_ops.append(ff_op)
            else:
                normalized_ops.append(op)
        operations = normalized_ops

        # Atomic pre-write formula guard scan across ALL operations before mutating any cell
        guard_mode = self.get_formula_guard_mode()
        if not allow_formula_overwrite and guard_mode != "off":
            blocked: List[Dict[str, Any]] = []
            for op in operations:
                op_sheet = self._extract_batch_sheet_name(op, default_sheet_name=default_sheet)
                start_cell = self._extract_batch_start_cell(op)
                data = self._resolve_batch_op_data(op)
                ws = self._get_sheet(entry, op_sheet)
                clean_start = validate_cell_ref(start_cell)
                col_letter, start_row = coordinate_from_string(clean_start)
                start_col = column_index_from_string(col_letter)
                for row_idx, row_data in enumerate(data):
                    for col_idx, raw_val in enumerate(row_data):
                        existing = ws.cell(row=start_row + row_idx, column=start_col + col_idx).value
                        if isinstance(existing, str) and existing.startswith("="):
                            clean_val, _ = self._coerce_excel_write_value(raw_val)
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
        range_warnings: List[str] = []
        formulas_updated = 0

        for op in operations:
            op_sheet = self._extract_batch_sheet_name(op, default_sheet_name=default_sheet)
            start_cell = self._extract_batch_start_cell(op)
            data = self._resolve_batch_op_data(op)
            formatting = op.get("formatting")
            header_formatting = op.get("header_formatting")
            number_format = op.get("number_format")

            ws = self._get_sheet(entry, op_sheet)
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
                    clean_val, inferred_fmt = self._coerce_excel_write_value(raw_val)
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
                    elif inferred_fmt and (not cell_obj.number_format or cell_obj.number_format == "General"):
                        cell_obj.number_format = inferred_fmt
                    if row_idx == 0 and header_formatting:
                        self._apply_cell_formatting(cell_obj, header_formatting)
                    elif formatting:
                        self._apply_cell_formatting(cell_obj, formatting)

            end_row = start_row + len(data) - 1
            end_col = start_col + max(max_cols, 1) - 1
            actual_range = f"{clean_start}:{get_column_letter(end_col)}{end_row}"

            # Detect column/row spill against explicit caller-provided A1:B2 range
            raw_ref = str(
                op.get("range") or op.get("cell_range") or op.get("start_cell") or ""
            ).strip()
            if "!" in raw_ref:
                raw_ref = raw_ref.rsplit("!", 1)[1].strip()
            if ":" in raw_ref and not op.get("fill_formula"):
                try:
                    l_ref, r_ref = [p.strip() for p in raw_ref.split(":", 1)]
                    rc1_let, rr1 = coordinate_from_string(validate_cell_ref(l_ref))
                    rc2_let, rr2 = coordinate_from_string(validate_cell_ref(r_ref))
                    exp_cols = column_index_from_string(rc2_let) - column_index_from_string(rc1_let) + 1
                    exp_rows = rr2 - rr1 + 1
                    if max_cols > exp_cols or len(data) > exp_rows:
                        hint_start = " Check start_cell (did you mean 'A1' instead of 'B1'?)." if start_col > 1 and max_cols > exp_cols else ""
                        range_warnings.append(
                            f"Range shape mismatch on '{ws.title}': requested '{raw_ref}' "
                            f"({exp_rows}x{exp_cols}) but data was {len(data)}x{max_cols} -> "
                            f"spilled to '{actual_range}'.{hint_start}"
                        )
                except Exception:
                    pass

            res_item: Dict[str, Any] = {
                "sheet_name": ws.title,
                "range": actual_range,
                "rows_written": len(data),
                "columns_written": max_cols,
            }
            cmap = self._build_compact_cell_map(ws, start_row, start_col, len(data), max_cols)
            if cmap:
                res_item["cell_map"] = cmap
            results.append(res_item)

        await self._save_working_copy(entry)
        out: Dict[str, Any] = {
            "status": "success",
            "workbook_id": workbook_id,
            "operations_completed": len(results),
            "results": results,
        }
        # Compact cross-turn state anchor (~30 tokens) so Preserve Thinking = Off models always see which tabs remain
        if 2 <= len(wb.worksheets) <= 8:
            has_hdr_meta = getattr(wb, "_m365_header_metadata", None) is not None
            min_body_r = 5 if has_hdr_meta else 1
            wb_prog: Dict[str, str] = {}
            for s_ws in wb.worksheets:
                body_cells = 0
                body_fmls = 0
                for (r_k, _c_k), c_cell in s_ws._cells.items():
                    if r_k >= min_body_r and c_cell.value is not None:
                        body_cells += 1
                        if isinstance(c_cell.value, str) and c_cell.value.startswith("="):
                            body_fmls += 1
                if body_cells == 0:
                    wb_prog[s_ws.title] = "empty"
                elif body_fmls > 0:
                    wb_prog[s_ws.title] = f"{body_cells} cells ({body_fmls} fml)"
                else:
                    wb_prog[s_ws.title] = f"{body_cells} cells"
            out["workbook_progress"] = wb_prog
        if formulas_updated:
            out["formulas_updated"] = formulas_updated
        if all_warnings:
            out["formula_warnings"] = all_warnings
        if range_warnings:
            out["warnings"] = range_warnings
        return sanitize_for_json(out)

    async def add_formula(
        self,
        workbook_id: str,
        sheet_name: Optional[str],
        cell: str,
        formula: str,
        number_format: Optional[str] = None,
        formatting: Optional[Dict[str, Any]] = None,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write an Excel formula to a cell with optional number format."""
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        ws = self._get_sheet(entry, sheet_name)
        clean_cell = validate_cell_ref(cell)

        wb = entry["workbook_object"]
        if not formula or not isinstance(formula, str):
            raise ValueError("formula must be a non-empty string")
        normalized = formula if formula.startswith("=") else f"={formula}"
        normalized = self._normalize_formula_sheet_refs(wb, normalized)
        formula_warnings = self._inspect_formula_sheet_refs(wb, normalized, cell_ref=clean_cell)

        cell_obj = ws[clean_cell]
        existing_val = cell_obj.value
        was_formula = isinstance(existing_val, str) and existing_val.startswith("=")
        guard_mode = self.get_formula_guard_mode()
        if was_formula and guard_mode == "strict" and not allow_formula_overwrite:
            return self._formula_guard_response(
                guard_mode,
                [{
                    "sheet_name": ws.title,
                    "cell": clean_cell,
                    "existing_formula": existing_val,
                    "proposed_value": normalized,
                }],
            )

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
        if was_formula:
            result["formulas_updated"] = 1
        if formula_warnings:
            result["formula_warnings"] = formula_warnings
        return sanitize_for_json(result)

    @staticmethod
    def _normalize_formula_sheet_refs(wb, formula: str) -> str:
        """Ensure same-workbook sheet references with spaces/special chars are properly single-quoted.

        Prevents Excel from flagging unquoted same-workbook references (e.g. `=Balance Sheet!B2`)
        as broken or external workbook links, without altering already-quoted references or
        string literals.
        """
        if not formula or not formula.startswith("="):
            return formula

        formula = sanitize_excel_formula_quotes(formula)

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
