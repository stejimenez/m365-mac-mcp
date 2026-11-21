---
name: dcf-investment-template-builder
description: Create a reusable institutional DCF & Multi-Stream Cash Flow Valuation template (CAPM WACC, multi-segment UFCF, dual Terminal Value, EV-to-Equity Bridge, scenarios, and charts) via m365-mac-mcp.
---

# DCF & Multi-Stream Investment Valuation Template Builder

## Goal
Build a clean, reusable, dynamic Excel valuation template for complex companies or multi-stream investment projects—complete with CAPM WACC, multi-segment operating cash flows, Gordon Growth & Exit Multiple Terminal Value, EV-to-Equity Bridge, and Base/Bull/Bear scenarios.

## Mandatory Pre-Flight Check (HALT if missing)
Confirm with the user:
- Company / project name, currency & units (`$M`), and forecast horizon (default 5 years).
- Business segments / cash flow streams to model (e.g., 2–4 segments or single consolidated stream).
- WACC inputs (Risk-Free Rate, Equity Risk Premium, Beta, Pre-Tax Cost of Debt, Tax Rate, Target D/E) or a direct WACC %.
- Terminal Growth Rate % and/or Exit EV/EBITDA multiple, plus Bridge items (Total Debt, Cash, Minority Interest, Preferred Stock, Diluted Shares).

## Exact `m365-mac-mcp` Workflow
1. **Create Workbook:** Call `create_workbook(title=..., sheet_name="Dashboard")` and add worksheets: `WACC & Inputs`, `Segment Cash Flows`, `DCF Valuation`, `Sensitivity`.
2. **Add Controls:** On `Dashboard!B2`, call `add_data_validation` (`validation_type="list"`, `formula1=["Base", "Bull", "Bear"]`) for live scenario switching.
3. **Populate All Tabs in ONE Call:** Call `batch_write_ranges`:
   - **`WACC & Inputs`:** CAPM Cost of Equity (`=Rf + Beta * ERP`), After-Tax Cost of Debt (`=Kd * (1 - TaxRate)`), WACC (`=We*Ke + Wd*Kd_after_tax`), and Base/Bull/Bear driver table.
   - **`Segment Cash Flows`:** Per-segment Revenue, Margin, EBITDA, D&A, CapEx, and ΔNWC rolling cleanly into **Consolidated Operating Totals** (`=SUM(...)`).
   - **`DCF Valuation`:**
     - EBIT, Taxes, NOPAT, `+D&A`, `-CapEx`, `-ΔNWC` → **Unlevered Free Cash Flow (UFCF)**.
     - Discount periods (`1..N`, with optional `-0.5` mid-year convention toggle), Discount Factors, **PV of Explicit UFCF**.
     - **Terminal Value (Gordon Growth):** `=FinalUFCF * (1 + g) / (WACC - g)` and PV of TV.
     - **Terminal Value (Exit Multiple):** `=FinalEBITDA * ExitMultiple` and PV of TV.
     - **Enterprise Value → Equity Value Bridge:** `EV - Total Debt + Cash - Minority Interest - Preferred = Equity Value`; `Implied Share Price = Equity Value / Diluted Shares`.
   - **`Sensitivity`:** 2-way formula table for **WACC vs. Terminal Growth Rate** (Implied Share Price / EV) and **WACC vs. Exit Multiple**.
   - **`Dashboard`:** Executive valuation summary pulling live links from `DCF Valuation` + `add_excel_table` + `create_chart` (UFCF & PV breakdown).
4. **Format, Audit & Live Verify:**
   - Call `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, `0.0"x"`, widths, `freeze_panes="B2"`).
   - Call `save_workbook`, `audit_workbook_formulas` (confirm `clean: true`), and `excel_live_open` + `excel_live_batch_read_ranges` on `Dashboard` to output live recalculated EV, Equity Value, and Implied Share Price.

## Guardrails
- Quote all spaced sheet names (`'WACC & Inputs'!C12`, `'Segment Cash Flows'!G15`).
- Zero hardcoded totals, discount factors, or EV bridge outputs.
- Respect `OFFICE_MCP_FORMULA_GUARD` on any follow-up edits.
