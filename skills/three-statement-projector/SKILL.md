---
name: three-statement-projector
description: Build dynamic IFRS/FASB multi-year 3-statement Excel projections with Base/Bull/Bear scenarios and live Balance Check = 0.
---

# 3-Statement Financial Projector (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Project FY26–FY30 with Base/Bull/Bear scenarios"*, *"Build a 5-year 3-statement forecast from these historicals."*
- **Deliverable:** Dynamically linked multi-year IS, BS, and CF projections driven by an `Assumptions` tab with a `Base`/`Bull`/`Bear` dropdown and live-verified `Balance Check = 0`.

## 2. Mandatory Pre-Flight / HALT Inputs
**HARD STOP (ZERO TOOLS):** If ANY item below is missing from the user's prompt, reply ONLY with 1–3 concise questions and call **ZERO tools** (do NOT call `create_workbook` or invent assumptions):
- Accounting basis: **IFRS** or **US GAAP (FASB)** (`IAS 7` vs. `ASC 230` Cash Flow classification; `IFRS 16` vs. `ASC 842` leases).
- Forecast horizon (e.g., 5 years) and currency/units (`$M`).
- Core drivers (or approval to use trailing historical averages as Base): Revenue growth %, Gross Margin %, OpEx % of Rev, Tax Rate %, D&A %, CapEx %, DSO/DIO/DPO, and Debt/Distributions.

## 3. Exact m365-mac-mcp Tool Sequence
1. **Audit Historicals:** Call `inspect_workbook(file_path=...)` and `audit_workbook_formulas(file_path=...)` to verify the base.
2. **Open & Add Assumptions Tab:** Call `open_workbook`, then `add_worksheet(workbook_id, "Assumptions", position=0)`.
3. **Add Scenario Dropdown:** Call `add_data_validation(validation_type="list", formula1=["Base", "Bull", "Bear"])` on `Assumptions!B2` and wire active drivers via `=IF($B$2="Bull", ..., IF($B$2="Bear", ..., ...))`.
4. **Write Projections (1 `batch_write_ranges` Call/Sheet — Labels/Inputs + `fill_formula` across projected columns, ZERO mental math):** Across `Assumptions`, `Income Statement`, `Balance Sheet`, and `Cash Flow`:
   - **IS:** Dynamic Revenue, COGS, OpEx, EBIT, Taxes (`IAS 12`/`ASC 740`), Net Income.
   - **BS:** Working capital via DSO/DIO/DPO, PP&E roll-forward (`Prior + CapEx - D&A` per `IAS 16`/`ASC 360`), Retained Earnings (`Prior RE + 'Income Statement'!NI - Distributions`), **Cash linked to `'Cash Flow'!Ending Cash`**, and `Balance Check = Assets - L&E`.
   - **CF:** Net Income link, `+D&A`, Working Capital deltas (`-ΔAR`, `-ΔInv`, `+ΔAP`), CapEx, Financing, Net Change in Cash, Ending Cash.
5. **Format, Save & Close:** Call `batch_format_workbook` (`auto_color_financial_model=True`, navy `#1F4E79` headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, widths, `freeze_panes="B2"`), then `save_workbook` and `close_workbook(workbook_id=...)`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS / FASB:** Interest Paid = Operating and Dividends Paid = Financing under FASB `ASC 230` (or consistently labeled per `IAS 7` under IFRS). Zero undocumented Balance Sheet plugs.
- **SOX 404:** Centralize blue inputs on `Assumptions`; zero hardcoded calculated rows; obey `M365_MCP_FORMULA_GUARD`.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` proving `Balance Check == 0` across every projected year, output a compact KPI table in plain finance language (never mention internal tool/skill names), and add a 1-line **Fresh-Chat Handoff Note**: *Saved to `<path>`. If your chat context is >50%, start a **New Chat** with this file path before running a DCF valuation, sensitivity test, or executive brief.*
