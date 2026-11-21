---
name: three-statement-projector
description: Extend a historical 3-statement Excel model (IS, BS, CF) with dynamic multi-year IFRS/FASB-compliant projections, Base/Bull/Bear scenarios, SOX controls, and live Balance Check = 0 verification via m365-mac-mcp.
---

# 3-Statement Financial Projector (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Project FY26–FY30 with Base/Bull/Bear scenarios"*, *"Build a 5-year 3-statement forecast from these historicals."*
- **Deliverable:** A dynamically linked multi-year 3-statement model driven by an `Assumptions` tab with a `Base`/`Bull`/`Bear` selector, verified live in Excel with `Balance Check = 0` in every projected period.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm (or HALT if unspecified) before writing any cells:
- Accounting standard: **IFRS** or **US GAAP (FASB)** (governs Cash Flow interest/dividend placement under `IAS 7` vs. `ASC 230` and lease presentation under `IFRS 16` vs. `ASC 842`).
- Projection horizon (e.g., 5 years) and currency/units (`$M`).
- Core drivers (or explicit approval to use trailing historical averages as Base Case): Revenue growth %, Gross Margin %, OpEx % of Revenue, Effective Tax Rate %, D&A % of Revenue, CapEx % of Revenue, Working Capital days (DSO/DIO/DPO), and Debt/Distribution assumptions.

## 3. Exact `m365-mac-mcp` Tool Sequence
1. **Audit Historicals:** Call `inspect_workbook(file_path=...)` and `audit_workbook_formulas(file_path=...)` to verify the historical base is clean.
2. **Open & Add Assumptions Tab:** Call `open_workbook`, then `add_worksheet(workbook_id, "Assumptions", position=0)`.
3. **Add Scenario Selector:** Call `add_data_validation(validation_type="list", formula1=["Base", "Bull", "Bear"])` on `Assumptions!B2` and wire active drivers via `=IF($B$2="Bull", ..., IF($B$2="Bear", ..., ...))`.
4. **Populate Projections in ONE Call:** Call `batch_write_ranges` across `Assumptions`, `Income Statement`, `Balance Sheet`, and `Cash Flow`:
   - **IS:** Dynamic Revenue, COGS, OpEx, EBIT, Taxes (`IAS 12`/`ASC 740`), Net Income (zero hardcoded totals/margins).
   - **BS:** Working capital from DSO/DIO/DPO, PP&E roll-forward (`Prior + CapEx - D&A` per `IAS 16`/`ASC 360`), Retained Earnings roll-forward (`Prior RE + 'Income Statement'!NI - Distributions`), **Ending Cash linked to `'Cash Flow'!Ending Cash`**, and `Balance Check = Assets - L&E`.
   - **CF:** Net Income linked to `'Income Statement'`, `+D&A`, Working Capital deltas (`-ΔAR`, `-ΔInv`, `+ΔAP`), CapEx, Financing, Net Change in Cash, Ending Cash (`Prior Cash + Net Change`).
5. **Format & Save:** Call `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, widths, `freeze_panes="B2"`), then `save_workbook`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS / FASB Consistency:** Classify Interest Paid as Operating and Dividends Paid as Financing under FASB `ASC 230` (or consistently per `IAS 7` under IFRS with explicit row labels). Never use undocumented Balance Sheet plugs.
- **SOX 404 Controls:** Separate all inputs on `Assumptions` (blue font); zero hardcoded calculated lines; respect `OFFICE_MCP_FORMULA_GUARD` (STOP on `confirmation_required`).
- **Formula Hygiene:** Single-quote spaced sheet names (`'Income Statement'!G15`, `'Cash Flow'!G16`); use JSON `null` (never `"None"`) for blank cells.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` confirming `Balance Check == 0` across all projected years, then output a compact projected KPI table.
