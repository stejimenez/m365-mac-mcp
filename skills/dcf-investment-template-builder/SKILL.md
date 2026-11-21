---
name: dcf-investment-template-builder
description: Create a reusable IFRS/FASB/SOX-compliant DCF & Multi-Stream Cash Flow Valuation template (CAPM WACC, multi-segment UFCF, dual Terminal Value, EV-to-Equity Bridge, scenarios, and charts) via m365-mac-mcp.
---

# DCF & Multi-Stream Valuation Template Builder (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Build a reusable DCF valuation template"*, *"Create a multi-segment cash flow investment model with WACC, Terminal Value, and Equity Bridge."*
- **Deliverable:** A modular 5-tab valuation workbook (`Dashboard`, `WACC & Inputs`, `Segment Cash Flows`, `DCF Valuation`, `Sensitivity`) with dynamic `Base`/`Bull`/`Bear` switching, dual Terminal Value methods, and live-verified Equity Value per share.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm (or HALT if unspecified) before building:
- Reporting framework: **IFRS** or **US GAAP (FASB)** (critical for `IFRS 16` vs. `ASC 842` lease-liability treatment in Debt, EBITDA, and EV/EBITDA exit multiples).
- Company/project name, currency & units (`$M`), forecast horizon (default 5 years), and segment/stream names.
- WACC inputs (`Rf`, `ERP`, `Beta`, `Kd`, `Tax Rate`, target `D/E`) or direct WACC %, Terminal Growth `g` and/or Exit EV/EBITDA multiple, plus EV Bridge items (Total Debt, Lease Liabilities, Cash, Minority Interest, Preferred Stock, Diluted Shares).

## 3. Exact `m365-mac-mcp` Tool Sequence
1. **Create Workbook:** Call `create_workbook(title=..., sheet_name="Dashboard")` and `add_worksheet` for `WACC & Inputs`, `Segment Cash Flows`, `DCF Valuation`, and `Sensitivity`.
2. **Add Scenario Control:** Call `add_data_validation(validation_type="list", formula1=["Base", "Bull", "Bear"])` on `Dashboard!B2`.
3. **Populate All 5 Tabs in ONE Call:** Call `batch_write_ranges`:
   - **`WACC & Inputs`:** CAPM `Ke = Rf + Beta * ERP`, `Kd_after_tax = Kd * (1 - TaxRate)`, `WACC = We*Ke + Wd*Kd_after_tax`, and Base/Bull/Bear driver schedules.
   - **`Segment Cash Flows`:** Independent segment Revenue, EBITDA, D&A, CapEx, and ΔNWC rolling into consolidated totals (`=SUM(...)`).
   - **`DCF Valuation`:** EBIT, Taxes (`IAS 12`/`ASC 740`), NOPAT, `+D&A`, `-CapEx`, `-ΔNWC` → **Unlevered FCF**; Discount Factors (optional `-0.5` mid-year convention toggle); **PV of Explicit UFCF**; **Terminal Value (Gordon Growth `=FinalUFCF*(1+g)/(WACC-g)` + Exit Multiple `=FinalEBITDA*Multiple`)**; **EV → Equity Bridge** (`EV - Debt - Leases + Cash - Minority Interest - Preferred = Equity Value`; `Share Price = Equity Value / Diluted Shares`).
   - **`Sensitivity`:** 2-way formula tables for **WACC vs. Terminal Growth** and **WACC vs. Exit Multiple**.
   - **`Dashboard`:** Executive valuation summary linked to `DCF Valuation` + `add_excel_table` + `create_chart`.
4. **Format & Save:** Call `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, `0.0"x"`, widths, `freeze_panes="B2"`), then `save_workbook`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS (`IFRS 16`) / FASB (`ASC 842`) Consistency:** Never mix pre-IFRS-16 EBITDA with post-IFRS-16 lease debt in the EV Bridge; ensure Terminal Growth `g < WACC` (flag error otherwise).
- **SOX 404 Controls:** Single source of truth for all assumptions on `WACC & Inputs`; zero hardcoded discount factors, Terminal Values, or Equity Bridge outputs; respect `OFFICE_MCP_FORMULA_GUARD`.
- **Formula Hygiene:** Single-quote spaced sheet names (`'WACC & Inputs'!C12`); use JSON `null` (never `"None"`) for blank cells.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` on `Dashboard`, then output ONLY a compact Valuation Summary table (WACC, PV of UFCF, PV of TV, EV, Equity Value, Implied Share Price).
