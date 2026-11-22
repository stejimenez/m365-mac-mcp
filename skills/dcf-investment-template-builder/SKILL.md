---
name: dcf-investment-template-builder
description: Build a reusable IFRS/FASB/SOX multi-segment DCF valuation template (CAPM WACC, dual TV, EV-to-Equity Bridge, charts).
---

# DCF & Multi-Stream Valuation Template Builder (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Build a reusable DCF valuation template"*, *"Create a multi-segment cash flow valuation model with WACC, Terminal Value, and Equity Bridge."*
- **Deliverable:** A 5-tab template (`Dashboard`, `WACC & Inputs`, `Segment Cash Flows`, `DCF Valuation`, `Sensitivity`) with `Base`/`Bull`/`Bear` switching, dual Terminal Value methods, and live-verified Share Price.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm concisely before building:
- Reporting framework: **IFRS** or **US GAAP (FASB)** (governs `IFRS 16` vs. `ASC 842` lease liabilities in Debt, EBITDA, and EV/EBITDA exit multiples).
- Company/project name, currency & units (`$M`), forecast horizon (default 5 years), and segment/stream names.
- WACC inputs (`Rf`, `ERP`, `Beta`, `Kd`, `Tax Rate`, target `D/E`) or direct WACC %, Terminal Growth `g` and/or Exit EV/EBITDA multiple, plus EV Bridge items (Total Debt, Lease Liabilities, Cash, Minority Interest, Preferred Stock, Diluted Shares).

## 3. Exact m365-mac-mcp Tool Sequence
1. **Create All 5 Tabs (1 Call):** Call `create_workbook(title=..., sheet_names=["Dashboard", "WACC & Inputs", "Segment Cash Flows", "DCF Valuation", "Sensitivity"])`.
2. **Add Scenario Control:** Call `add_data_validation(validation_type="list", formula1=["Base", "Bull", "Bear"])` on `Dashboard!B2`.
3. **Write All 5 Tabs (1 Call):** Call `batch_write_ranges`:
   - **`WACC & Inputs`:** CAPM `Ke = Rf + Beta * ERP`, `Kd_after_tax = Kd * (1 - TaxRate)`, `WACC = We*Ke + Wd*Kd_after_tax`, and Base/Bull/Bear drivers.
   - **`Segment Cash Flows`:** Per-segment Revenue, EBITDA, D&A, CapEx, ΔNWC rolling into consolidated totals (`=SUM(...)`).
   - **`DCF Valuation`:** EBIT, Taxes (`IAS 12`/`ASC 740`), NOPAT, `+D&A`, `-CapEx`, `-ΔNWC` → **UFCF**; Discount Factors (optional `-0.5` mid-year toggle); **PV of UFCF**; **Terminal Value (Gordon Growth `=FinalUFCF*(1+g)/(WACC-g)` + Exit Multiple `=FinalEBITDA*Multiple`)**; **EV → Equity Bridge** (`EV - Debt - Leases + Cash - Minority Interest - Preferred = Equity Value`; `Implied Share Price = Equity Value / Diluted Shares`).
   - **`Sensitivity`:** 2-way formula tables for **WACC vs. Terminal Growth** and **WACC vs. Exit Multiple**.
   - **`Dashboard`:** Executive valuation summary linked to `DCF Valuation` + `add_excel_table` + `create_chart`.
4. **Format, Save & Close:** Call `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, `0.0"x"`, widths, `freeze_panes="B2"`), then `save_workbook` and `close_workbook(workbook_id=...)`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS (`IFRS 16`) / FASB (`ASC 842`) Parity:** Never mix pre-IFRS-16 EBITDA with post-IFRS-16 lease debt in the EV Bridge; enforce Terminal Growth `g < WACC`.
- **SOX 404:** Single source of truth for inputs on `WACC & Inputs`; zero hardcoded discount factors, Terminal Values, or EV Bridge outputs.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` on `Dashboard`, output ONLY a compact Valuation Summary table in plain finance language (never mention internal tool/skill names), and add a 1-line **Fresh-Chat Handoff Note**: *Saved to `<path>`. If the bottom-right **Context** indicator is >50%, start a **New Chat** with this file path for sensitivity tests or an executive brief.*
