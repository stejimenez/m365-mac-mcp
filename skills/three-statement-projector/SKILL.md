---
name: three-statement-projector
description: Extend a historical 3-statement Excel model (Income Statement, Balance Sheet, Cash Flow) with dynamic multi-year projections, Base/Bull/Bear scenarios, and live Balance Check = 0 verification.
---

# 3-Statement Financial Projector

## Goal
Build dynamic, fully linked multi-year financial projections (IS, BS, CF) driven by an `Assumptions` tab with a `Base` / `Bull` / `Bear` selector—verified live in Excel so `Balance Check = 0` in every projected year.

## Mandatory Pre-Flight Check (HALT if missing)
Before writing any projection cells, confirm with the user:
- Projection horizon (e.g., 5 years: `FY2026–FY2030`) and units (`$M`).
- Core operating drivers (or confirm using trailing historical averages as Base Case): Revenue growth %, Gross Margin %, R&D/SG&A % of Revenue, Effective Tax Rate %, D&A % of Revenue, CapEx % of Revenue, Working Capital days (DSO, DIO, DPO), and Dividends/Buybacks or Debt schedule assumptions.

## Exact `m365-mac-mcp` Workflow
1. **Audit Historicals First:** Call `inspect_workbook(file_path=...)` and `audit_workbook_formulas(file_path=...)` to verify the historical base.
2. **Open & Add Assumptions Tab:** Call `open_workbook`, then `add_worksheet(workbook_id, "Assumptions", position=0)`.
3. **Add Scenario Selector:** Call `add_data_validation` (`validation_type="list"`, `formula1=["Base", "Bull", "Bear"]`) on `Assumptions!B2` and populate Base/Bull/Bear driver rows with `=INDEX(...)` or `=IF($B$2="Bull", ..., IF($B$2="Bear", ..., ...))` active-case formulas.
4. **Populate Projections in ONE Call:** Call `batch_write_ranges` across `Assumptions`, `Income Statement`, `Balance Sheet`, and `Cash Flow`:
   - **Income Statement:** Dynamic revenue growth, margins, OpEx, tax, Net Income (zero hardcoded totals/margins).
   - **Balance Sheet:** Working capital linked via DSO/DIO/DPO, PP&E roll-forward (`Prior PP&E + CapEx - D&A`), Retained Earnings roll-forward (`Prior RE + 'Income Statement'!NI - Distributions`), **Cash linked to `'Cash Flow'!Ending Cash`**, and `Balance Check = Total Assets - Total L&E`.
   - **Cash Flow:** Net Income linked to `'Income Statement'`, D&A, Working Capital changes (`-ΔAR`, `-ΔInv`, `+ΔAP`), CapEx, Financing, Net Change in Cash, Ending Cash (`Prior Cash + Net Change`).
5. **Format & Audit:** Call `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, widths, `freeze_panes="B2"`), then `save_workbook` and `audit_workbook_formulas`.
6. **Live Balance Verification:** Call `excel_live_open` + `excel_live_batch_read_ranges` on `Balance Check` across all projected years to confirm `0` in every year.

## Guardrails
- Always single-quote sheet names with spaces (`'Income Statement'!G15`, `'Cash Flow'!G16`).
- Use JSON `null` for blank cells—never the literal string `"None"`.
- Respect `OFFICE_MCP_FORMULA_GUARD`: never overwrite historical formulas with literals.
- Output a compact FY projection KPI summary table (Revenue, EBITDA/EBIT, Net Income, CFO, FCF, Ending Cash, Balance Check).
