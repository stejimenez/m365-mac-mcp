"""Workbook creation, loading, closing, worksheet CRUD, save & live-sync helpers."""

import asyncio
import csv
import os
import re
import time
import uuid
from pathlib import Path
from typing import (
    Any,
    Dict,
    List,
    Optional,
)

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from utils.logger import setup_logger
from utils.sheet_layout import (
    NON_CURRENCY_SCALE_KEYWORDS,
    PCT_KEYWORDS,
    analyze_worksheet_layout,
    build_magnitude_spec,
    is_per_unit_label,
    normalize_currency,
)
from utils.validators import (
    assert_safe_write_path,
    atomic_save_via_temp,
    get_client_project_dir,
    has_invalid_excel_formula_quotes,
    is_within_trusted_project_dir,
    resolve_file_path,
    sanitize_excel_write_value,
    sanitize_for_json,
)

logger = setup_logger(__name__)


class ExcelIOMixin:
    """Workbook creation, loading, closing, worksheet CRUD, save & live-sync helpers."""

    async def create_workbook(
        self,
        title: str = "New Workbook",
        template_path: Optional[str] = None,
        sheet_name: str = "Sheet1",
        sheet_names: Optional[List[str]] = None,
        include_title_header: bool = False,
        open_in_excel: bool = False,
        header_metadata: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Create a new Excel workbook (.xlsx), optionally with multiple worksheets & 4-line institutional headers."""
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
            clean_names = [str(n).strip() for n in (sheet_names or []) if str(n).strip()]
            if clean_names:
                ws.title = clean_names[0][:31]
                for extra_name in clean_names[1:]:
                    if extra_name[:31] not in wb.sheetnames:
                        wb.create_sheet(title=extra_name[:31])
            else:
                ws.title = (sheet_name or "Sheet1")[:31]
            if include_title_header and title and not header_metadata:
                ws["A1"] = title
                ws["A1"].font = Font(bold=True, size=14)

        # IMP-001: Optional 4-line institutional worksheet header block (Rows 1-4, blank spacer Row 5, table starts Row 6)
        resolved_header_meta: Optional[Dict[str, Any]] = None
        if isinstance(header_metadata, dict) and header_metadata:
            entity = str(
                header_metadata.get("entity")
                or header_metadata.get("company")
                or header_metadata.get("project")
                or header_metadata.get("company_name")
                or header_metadata.get("project_name")
                or title
            ).strip()
            period = str(
                header_metadata.get("period")
                or header_metadata.get("years")
                or header_metadata.get("horizon")
                or header_metadata.get("time_horizon")
                or ""
            ).strip()
            currency_code, currency_sym = normalize_currency(header_metadata.get("currency") or "USD")
            raw_mag = str(header_metadata.get("magnitude") or "auto").strip().lower()
            subtitles = header_metadata.get("sheet_subtitles") if isinstance(header_metadata.get("sheet_subtitles"), dict) else {}
            non_monetary = {
                str(s).strip().lower()
                for s in (header_metadata.get("non_monetary_sheets") or [])
                if str(s).strip()
            }
            resolved_header_meta = {
                "entity": entity,
                "period": period,
                "currency": currency_code,
                "currency_symbol": currency_sym,
                "magnitude": raw_mag or "auto",
                "non_monetary_sheets": sorted(non_monetary),
            }
            initial_spec = (
                build_magnitude_spec(currency=currency_code, scale_name=raw_mag, display_scaled=True)
                if raw_mag not in ("", "auto", "smart")
                else build_magnitude_spec(currency=currency_code, scale_name="exact", display_scaled=False)
            )
            for s_ws in wb.worksheets:
                st_low = s_ws.title.lower()
                stmt_name = str(subtitles.get(s_ws.title) or s_ws.title).strip()
                s_ws["A1"] = entity
                s_ws["A2"] = stmt_name
                if period:
                    s_ws["A3"] = period
                if st_low not in non_monetary:
                    if any(kw in st_low for kw in ("assumption", "input", "driver")):
                        s_ws["A4"] = (
                            f"Input Basis: Exact {currency_code} {currency_sym} "
                            f"(Statements Presented {initial_spec.header_label})"
                            if raw_mag not in ("", "auto", "smart")
                            else f"Input Basis: Exact {currency_code} {currency_sym} (Smart-Scaled Statements)"
                        )
                    else:
                        s_ws["A4"] = (
                            initial_spec.header_label
                            if raw_mag not in ("", "auto", "smart")
                            else f"In {currency_code} {currency_sym} (Smart Magnitude — Auto)"
                        )
                s_ws.row_dimensions[5].height = 8
            setattr(wb, "_m365_header_metadata", resolved_header_meta)

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

        metadata: Dict[str, Any] = {
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
        if resolved_header_meta is not None:
            metadata["header_metadata"] = resolved_header_meta
            metadata["table_start_row"] = 6
            metadata["recommended_freeze_panes"] = "B7"
        if open_note:
            metadata["note"] = open_note

        self.active_workbooks[workbook_id] = {
            "metadata": metadata,
            "workbook_object": wb,
            "working_path": temp_file,
            "temp_file": temp_file,
            "dirty": False,
            "header_metadata": resolved_header_meta,
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
            "dirty": False,
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
        workbook_id: Optional[str] = None,
        save_before_close: bool = False,
    ) -> Dict[str, Any]:
        """Close an in-memory workbook, release resources, and remove temp files."""
        workbook_id = self._resolve_workbook_id(workbook_id)
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        saved_path = None
        if save_before_close and not entry["metadata"].get("is_temporary", True):
            save_res = await self.save_workbook(workbook_id)
            saved_path = save_res.get("file_path") or str(entry["working_path"])

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

    async def list_worksheets(self, workbook_id: str) -> List[str]:
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        return list(entry["workbook_object"].sheetnames)

    async def add_worksheet(
        self,
        workbook_id: str,
        sheet_name: str,
        position: Optional[int] = None,
    ) -> Dict[str, Any]:
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id = entry["metadata"]["workbook_id"]
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
        workbook_id = entry["metadata"]["workbook_id"]
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

    _warned_legacy_live_sync: bool = False

    @classmethod
    def get_live_sync_mode(cls, raw: Optional[str] = None) -> str:
        """Parse M365_MCP_LIVE_SYNC (or legacy OFFICE_MCP_LIVE_SYNC): 'auto' (default) | 'off'."""
        if raw is None:
            env_val = os.environ.get("M365_MCP_LIVE_SYNC")
            if env_val is None or not env_val.strip():
                legacy = os.environ.get("OFFICE_MCP_LIVE_SYNC")
                if legacy is not None and legacy.strip():
                    if not cls._warned_legacy_live_sync:
                        logger.warning(
                            "OFFICE_MCP_LIVE_SYNC is deprecated; rename it to M365_MCP_LIVE_SYNC."
                        )
                        cls._warned_legacy_live_sync = True
                    env_val = legacy
            val = (env_val or "auto").strip().lower()
        else:
            val = raw.strip().lower()
        if not val or val in ("auto", "on", "true", "1", "yes"):
            return "auto"
        if val in ("off", "false", "0", "no", "none"):
            return "off"
        logger.warning(
            f"Unknown M365_MCP_LIVE_SYNC='{val}'. Falling back to 'auto'. Valid values: auto, off."
        )
        return "auto"

    def _quick_save_sanity_warnings(self, wb: Workbook) -> List[str]:
        """Fast (~1-2ms) post-build linter run during save_workbook; returns [] when workbook is clean."""
        warnings: List[str] = []
        error_tokens = {"#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#N/A", "#NUM!", "#NULL!"}
        ref_warnings: List[str] = []

        sheet_macro_max: Dict[str, float] = {}
        sheet_macro_min: Dict[str, float] = {}

        for ws in wb.worksheets:
            lay = analyze_worksheet_layout(ws)
            non_empty = 0
            unformatted_nums = 0
            total_nums = 0
            unformatted_formulas = 0
            total_formulas_ws = 0
            sheet_errors: List[str] = []
            circular_refs: List[str] = []
            invalid_formulas: List[str] = []
            sign_warnings: List[str] = []
            max_r = ws.max_row or 0
            max_c = ws.max_column or 0
            # Sparse cell iteration via ws._cells prevents ghost 1,048,576-row hang & empty Cell allocation
            cells_dict = getattr(ws, "_cells", None)
            if isinstance(cells_dict, dict):
                populated_cells = [
                    (r, c, cobj)
                    for (r, c), cobj in sorted(list(cells_dict.items()))
                    if cobj is not None and cobj.value is not None
                ]
            else:
                populated_cells = [
                    (r, c, ws.cell(row=r, column=c))
                    for r in range(1, max_r + 1)
                    for c in range(1, max_c + 1)
                    if ws.cell(row=r, column=c).value is not None
                ]
            for r, c, cell in populated_cells:
                    val = cell.value
                    if val is None:
                        continue
                    non_empty += 1
                    blk = lay.block_at(r, c)
                    if isinstance(val, str):
                        s = val.strip()
                        if s in error_tokens:
                            sheet_errors.append(f"{ws.title}!{cell.coordinate} ({s})")
                        elif s.startswith("="):
                            total_formulas_ws += 1
                            nf = (cell.number_format or "").strip()
                            if not nf or nf == "General":
                                unformatted_formulas += 1
                            ref_warnings.extend(
                                self._inspect_formula_sheet_refs(
                                    wb, s, cell_ref=f"{ws.title}!{cell.coordinate}"
                                )
                            )
                            # Strip quoted string literals before structural formula checks
                            s_no_str = re.sub(r'"[^"]*"', '""', s)
                            # Self-referencing circular formula check (e.g. D17 = "=D17+E16")
                            col_let = get_column_letter(c)
                            self_pat = rf"(?<![A-Za-z0-9_!'$])\$?{col_let}\$?{r}(?![A-Za-z0-9_])"
                            if re.search(self_pat, s_no_str, re.IGNORECASE):
                                circular_refs.append(f"{ws.title}!{cell.coordinate} ({s})")
                            # Invalid 1-argument NPV(...) check (causes Excel Parameter error -50 on open)
                            if re.search(r"(?<![A-Za-z0-9_])NPV\s*\(\s*[^,()]+\s*\)", s_no_str, re.IGNORECASE):
                                invalid_formulas.append(
                                    f"{ws.title}!{cell.coordinate} ({s} — NPV() requires >=2 arguments: NPV(rate, values))"
                                )
                            # Invalid backslash-escaped \" or standalone single-quoted string literals
                            if has_invalid_excel_formula_quotes(s):
                                invalid_formulas.append(
                                    f"{ws.title}!{cell.coordinate} ({s} — invalid quote syntax in formula)"
                                )
                            # Exponentiation against a non-numeric text header cell (e.g. ^C$6 where C6='Year 1' -> #VALUE!)
                            for exp_m in re.finditer(r"\^\s*\$?([A-Za-z]{1,3})\$?([1-9]\d*)(?![A-Za-z0-9_!])", s_no_str):
                                tgt_coord = f"{exp_m.group(1).upper()}{exp_m.group(2)}"
                                tgt_val = ws[tgt_coord].value
                                if isinstance(tgt_val, str) and not tgt_val.startswith("="):
                                    invalid_formulas.append(
                                        f"{ws.title}!{cell.coordinate} (exponent ^{exp_m.group(1)}{exp_m.group(2)} references text '{tgt_val}' -> #VALUE!; use numeric header 0..N or COLUMN()-offset)"
                                    )
                            # Blank first-arg discount rate in NPV(X$Y, ...) (classic sensitivity table column shift)
                            for npv_m in re.finditer(r"(?<![A-Za-z0-9_])NPV\s*\(\s*\$?([A-Za-z]{1,3})\$?([1-9]\d*)\s*,", s_no_str, re.IGNORECASE):
                                rate_coord = f"{npv_m.group(1).upper()}{npv_m.group(2)}"
                                if ws[rate_coord].value is None:
                                    invalid_formulas.append(
                                        f"{ws.title}!{cell.coordinate} (NPV rate {rate_coord} is blank — check column alignment)"
                                    )
                            # Bare $B$N referencing 0/blank alongside Assumptions! when Assumptions!B{N} is a non-zero input
                            if "assumptions!" in s_no_str.lower() and "assumption" not in ws.title.lower():
                                assump_ws = next((sw for sw in wb.worksheets if sw.title.lower() == "assumptions"), None)
                                if assump_ws is not None:
                                    for bare_m in re.finditer(r"(?<![A-Za-z0-9_!'])\$([A-Za-z]{1,3})\$(\d+)(?![A-Za-z0-9_:])", s_no_str):
                                        b_coord = f"{bare_m.group(1).upper()}{bare_m.group(2)}"
                                        same_v = ws[b_coord].value
                                        assump_v = assump_ws[b_coord].value
                                        if (same_v is None or same_v == 0) and isinstance(assump_v, (int, float)) and not isinstance(assump_v, bool) and assump_v != 0:
                                            invalid_formulas.append(
                                                f"{ws.title}!{cell.coordinate} (bare '${bare_m.group(1)}${bare_m.group(2)}' is {same_v} on '{ws.title}' while Assumptions!{b_coord}={assump_v} — missing 'Assumptions!' prefix?)"
                                            )
                            # Layout-agnostic cash-flow outflow sign-inversion warning (works on ANY row/col)
                            lcol = blk.label_col if blk is not None else 1
                            if c > lcol:
                                rlbl = str(ws.cell(row=r, column=lcol).value or "").strip()
                                if rlbl.startswith("-") and re.match(
                                    r"^=\s*(?:'[^']+'!|[A-Za-z_][A-Za-z0-9_]*!)?\$?[A-Za-z]{1,3}\$?\d+$",
                                    s,
                                ):
                                    sign_warnings.append(
                                        f"{ws.title}!{cell.coordinate} ('{rlbl}' has positive direct link '{s}' — cash outflows should normally be negative '=-...')"
                                    )
                    elif isinstance(val, (int, float)) and not isinstance(val, bool):
                        # Skip header rows / account-code columns when checking formatting or magnitude
                        is_hdr_or_code = (
                            blk is not None
                            and (
                                r in blk.header_rows
                                or r in blk.subheader_rows
                                or c <= blk.label_col
                            )
                        )
                        if not is_hdr_or_code:
                            abs_v = abs(val)
                            if abs_v >= 1000 or (0 < abs_v < 1):
                                total_nums += 1
                                nf = (cell.number_format or "").strip()
                                if not nf or nf == "General":
                                    unformatted_nums += 1
                            # Track macro monetary literals per sheet for 1,000x magnitude-mismatch detection
                            lcol = blk.label_col if blk is not None else 1
                            rlbl_low = str(ws.cell(row=r, column=lcol).value or "").strip().lower()
                            if (
                                abs_v >= 5
                                and not is_per_unit_label(rlbl_low)
                                and not any(kw in rlbl_low for kw in NON_CURRENCY_SCALE_KEYWORDS)
                                and not any(kw in rlbl_low for kw in PCT_KEYWORDS)
                                and (blk is None or blk.kind != "MATRIX_TABLE")
                            ):
                                sheet_macro_max[ws.title] = max(sheet_macro_max.get(ws.title, 0.0), float(abs_v))
                                prev_min = sheet_macro_min.get(ws.title)
                                sheet_macro_min[ws.title] = float(abs_v) if prev_min is None else min(prev_min, float(abs_v))

            if non_empty == 0:
                warnings.append(f"Worksheet '{ws.title}' is empty.")
            if sheet_errors:
                warnings.append(f"Excel error tokens found: {', '.join(sheet_errors[:5])}.")
            if circular_refs:
                warnings.append(f"Self-referencing circular formulas found: {', '.join(circular_refs[:5])}.")
            if invalid_formulas:
                warnings.append(f"Invalid formula syntax found: {', '.join(invalid_formulas[:5])}.")
            if sign_warnings:
                warnings.append(f"Possible outflow sign error: {', '.join(sign_warnings[:3])}.")
            if (unformatted_nums >= 4 and total_nums > 0 and (unformatted_nums / total_nums) >= 0.8) or (
                unformatted_formulas >= 6 and total_formulas_ws > 0 and (unformatted_formulas / total_formulas_ws) >= 0.8
            ):
                warnings.append(
                    f"Sheet '{ws.title}' has unformatted numeric/formula cells still in 'General' format; "
                    "call batch_format_workbook(auto_color_financial_model=true, auto_fit_columns=true) with currency/percent formats."
                )
            if any(kw in ws.title.lower() for kw in ("chart", "graph")) and len(getattr(ws, "_charts", [])) == 0:
                warnings.append(
                    f"Worksheet '{ws.title}' has 0 charts; call create_chart(...) before finalizing."
                )
            if (
                non_empty >= 4
                and total_formulas_ws == 0
                and any(kw in ws.title.lower() for kw in ("sensitivity", "dcf", "valuation", "cash flow", "income statement", "balance sheet"))
            ):
                warnings.append(
                    f"Worksheet '{ws.title}' has 0 formulas; populate calculated cells with dynamic Excel formulas before saving."
                )
            # Detect exact-share-count (>=1M) mixed with $M-scaled monetary inputs (<10K) on Assumptions
            if any(kw in ws.title.lower() for kw in ("assumption", "input", "driver")):
                max_shares = 0.0
                max_money_in = 0.0
                for (r_i, c_i), c_obj in ws._cells.items():
                    v_i = c_obj.value
                    if isinstance(v_i, (int, float)) and not isinstance(v_i, bool) and abs(v_i) > 0:
                        lbl_i = str(ws.cell(row=r_i, column=max(1, c_i - 1)).value or "").strip().lower()
                        if "share" in lbl_i:
                            max_shares = max(max_shares, abs(float(v_i)))
                        elif any(k in lbl_i for k in ("revenue", "debt", "cash", "capex", "lease", "minority", "salvage")):
                            if "%" not in lbl_i and "growth" not in lbl_i and "margin" not in lbl_i and "rate" not in lbl_i:
                                max_money_in = max(max_money_in, abs(float(v_i)))
                if max_shares >= 1_000_000 and 0 < max_money_in < 10_000:
                    warnings.append(
                        f"Magnitude mismatch on '{ws.title}': Diluted Shares is stored as exact count ({int(max_shares):,}) "
                        f"while monetary inputs (max={max_money_in:g}) appear to be in $M — store ALL monetary inputs in exact base currency "
                        f"(e.g. 120000000 instead of 120) so Share Price = Equity / Shares is not off by 1,000,000x."
                    )
            # Flag multi-row numeric/formula tables whose label column (Col A) is completely blank
            for blk in lay.data_blocks():
                if blk.kind == "DATA_TABLE" and (blk.max_row - blk.min_row) >= 4 and blk.max_col >= 3:
                    data_rows_blk = [ri for ri in range(blk.min_row + 1, blk.max_row + 1) if ri not in blk.header_rows]
                    if len(data_rows_blk) >= 4:
                        labeled_cnt = sum(
                            1 for ri in data_rows_blk
                            if ws.cell(row=ri, column=blk.label_col).value not in (None, "")
                        )
                        if labeled_cnt == 0:
                            warnings.append(
                                f"Worksheet '{ws.title}' table R{blk.min_row}:R{blk.max_row} has {len(data_rows_blk)} numeric/formula rows "
                                f"with missing row labels in column {get_column_letter(blk.label_col)}."
                            )
            if max_c >= 2 and non_empty >= 4:
                label_cols = {b.label_col for b in lay.data_blocks()} or {1}
                has_long_label = any(
                    len(str(ws.cell(row=ri, column=lc).value or "")) > 11
                    for lc in label_cols
                    for ri in range(1, max_r + 1)
                    if (ri, lc) not in lay.overflow_banner_cells
                )
                if max_c >= 3 or has_long_label:
                    has_any_width = any(
                        (
                            (cd := ws.column_dimensions.get(get_column_letter(ci))) is not None
                            and cd.width is not None
                        )
                        for ci in range(1, max_c + 1)
                    )
                    if not has_any_width:
                        warnings.append(
                            f"Worksheet '{ws.title}' has no column widths set (labels may clip); "
                            "pass auto_fit_columns=true in batch_format_workbook."
                        )

        # IMP-001 Safeguard: Detect cross-sheet 1,000x / 1,000,000x magnitude mismatch
        # (e.g., Assumptions stored as 25,000,000 while a statement tab hardcodes -25 in $M)
        if len(sheet_macro_max) >= 2:
            overall_max = max(sheet_macro_max.values())
            if overall_max >= 1_000_000:
                for sh_name, sh_max in sheet_macro_max.items():
                    if 5 <= sh_max <= 500 and (overall_max / sh_max) >= 10_000:
                        warnings.append(
                            f"Possible magnitude mismatch: '{sh_name}' has macro numeric literals around {sh_max:g} "
                            f"while another sheet uses full units around {overall_max:,.0f} — verify consistent units!"
                        )

        if ref_warnings:
            warnings.extend(ref_warnings[:5])
        return warnings

    async def save_workbook(
        self,
        workbook_id: Optional[str] = None,
        file_path: Optional[str] = None,
        format: str = "xlsx",
        sheet_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Save workbook in-place or to a new path (.xlsx, .xlsm, .csv, .pdf)."""
        workbook_id = self._resolve_workbook_id(workbook_id)
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        wb = entry["workbook_object"]
        fmt = format.lower().lstrip(".")

        if file_path:
            save_path = resolve_file_path(file_path, must_exist=False, for_write=True)
        else:
            if entry["metadata"].get("is_temporary", True):
                raise ValueError(
                    "Workbook has no saved path yet; provide file_path to save."
                )
            save_path = assert_safe_write_path(Path(entry["working_path"]))

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

            async with self._get_wb_lock(workbook_id):
                self._ensure_valid_excel_table_headers(wb)
                sanity_warnings = self._quick_save_sanity_warnings(wb)
                await asyncio.to_thread(atomic_save_via_temp, wb.save, save_path)
                entry["working_path"] = save_path
                entry["metadata"]["file_path"] = str(save_path)
                entry["metadata"]["is_temporary"] = False
                entry["dirty"] = False
            logger.info(f"Saved workbook to {save_path}")

            reopen_timed_out = False
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
                except TimeoutError as exc:
                    reopen_timed_out = True
                    live_sync_warning = (
                        f"Saved to disk cleanly, but Microsoft Excel timed out auto-reopening '{save_path.name}' "
                        f"(likely waiting on a macOS 'Grant Access' or modal dialog): {exc}"
                    )
                    logger.warning(live_sync_warning)
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
            if reopen_timed_out:
                out["awaiting_user_action"] = True
                out["hint"] = (
                    "HARD STOP IMMEDIATELY: the file saved to disk, but Microsoft Excel is likely waiting on a "
                    "macOS 'Grant Access' or modal dialog while reloading. Do NOT run live reads or downstream steps; "
                    "ask the user to click 'Grant Access' / dismiss the Excel dialog and reply 'continue'."
                )
            if sanity_warnings:
                out["sanity_warnings"] = sanity_warnings
                out["warning_count"] = len(sanity_warnings)
                out["attention_required"] = True
                out["next_step"] = "Resolve sanity_warnings (fix formulas, formatting, column widths, or missing charts) BEFORE calling excel_live_open."
            if get_client_project_dir() is not None and not is_within_trusted_project_dir(save_path):
                out["grant_access_notice"] = (
                    "Saved outside the active Project Working Folder; opening this file live in Microsoft Excel "
                    "may prompt the user for macOS Grant Access."
                )
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
            pdf_staging: Optional[Path] = None
            source_for_pdf = Path(entry["working_path"])
            in_trusted_project = is_within_trusted_project_dir(save_path)
            needs_staging = (
                entry.get("dirty")
                or entry["metadata"].get("is_temporary", True)
                or not source_for_pdf.exists()
                or (in_trusted_project and not is_within_trusted_project_dir(source_for_pdf))
            )
            if needs_staging:
                ext = ".xlsm" if entry["metadata"].get("keep_vba") else ".xlsx"
                # Stage inside save_path.parent when exporting within the trusted Project Working Folder
                # so LaunchServices opens it cleanly with zero stray Grant Access prompts; unlink in finally.
                stage_dir = save_path.parent if in_trusted_project else self.temp_dir
                stage_prefix = ".m365_pdf_stage_" if in_trusted_project else "pdf_stage_"
                pdf_staging = stage_dir / f"{stage_prefix}{uuid.uuid4()}{ext}"
                await asyncio.to_thread(wb.save, str(pdf_staging))
                source_for_pdf = pdf_staging
            open_res = await self.applescript.open_excel_file(
                str(source_for_pdf),
                read_only=not in_trusted_project,
                activate=False,
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
                if pdf_staging and pdf_staging.exists():
                    try:
                        pdf_staging.unlink()
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

    async def get_workbook_info(self, workbook_id: str) -> Dict[str, Any]:
        entry = self._get_entry(workbook_id)
        workbook_id = entry["metadata"]["workbook_id"]
        self._sync_metadata(entry)
        return sanitize_for_json(entry["metadata"])

    async def list_workbooks(self) -> List[Dict[str, Any]]:
        return [sanitize_for_json(data["metadata"]) for data in self.active_workbooks.values()]
