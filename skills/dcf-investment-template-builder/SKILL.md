---
name: dcf-investment-template-builder
version: 2.3.1A
description: "[v2.3.1A] Build a reusable IFRS/FASB/SOX multi-segment DCF valuation template (CAPM WACC, dual TV, EV-to-Equity Bridge, charts)."
---

# DCF & Multi-Stream Valuation Template Builder (IFRS / FASB / SOX) (v2.3.1A)

## 1. Trigger & Scope
- **Triggers:** *"Build a reusable DCF valuation template"*, *"Create a multi-segment cash flow valuation model with WACC, Terminal Value, and Equity Bridge."*
- **Deliverable:** A 5-tab template (`Dashboard`, `Assumptions`, `Segment Cash Flows`, `DCF Valuation`, `Sensitivity`) with a pure-input `Assumptions` tab, `Base`/`Bull`/`Bear` switching, dual Terminal Value methods, and live-verified Share Price.

## 2. Mandatory Pre-Flight / HALT Inputs
**HARD STOP (ZERO TOOLS):** If ANY item below is missing, reply ONLY with 1–3 concise questions and call **ZERO tools** (NEVER call `create_workbook` or invent assumptions):
- Reporting framework: **IFRS** or **US GAAP (FASB)** (`IFRS 16` vs. `ASC 842` lease liabilities in Debt, EBITDA, and EV/EBITDA multiples).
- Company name, currency/units (`$M`), forecast horizon (default 5 yrs), and segment names.
- WACC inputs (`Rf`, `ERP`, `Beta`, `Kd`, `Tax Rate`, target `D/C`) or WACC %, Terminal `g` and/or Exit Multiple, plus EV Bridge items (Debt, Leases, Cash, Minority Interest, Preferred, Diluted Shares).

## 3. Exact m365-mac-mcp Tool Sequence
1. **Create All 5 Tabs (1 Call):** `create_workbook(title=..., sheet_names=["Dashboard", "Assumptions", "Segment Cash Flows", "DCF Valuation", "Sensitivity"])`.
2. **Add Scenario Control:** `add_data_validation(validation_type="list", formula1=["Base", "Bull", "Bear"])` on `Dashboard!B2`.
3. **Write All 5 Tabs (1 `batch_write_ranges` Call/Sheet — 2D `data` arrays + `fill_formula` with `$`-locked `Assumptions!$B$...` refs, ZERO mental math):**
   - **`Assumptions`:** **RAW BLUE INPUTS ONLY — ZERO calculated results or KPIs.** Store CAPM inputs (`Rf`, `ERP`, `Beta`, `Kd`, `Tax`, `D/C`), Base/Bull/Bear segment drivers, Terminal `g` & Exit Multiple, and EV Bridge raw inputs.
   - **`Segment Cash Flows`:** Per-segment Revenue, EBITDA, D&A, CapEx, ΔNWC referencing `Assumptions!$B$...` & `Dashboard!$B$2`, rolling into consolidated totals (`=SUM(...)`).
   - **`DCF Valuation`:** Computed CAPM WACC (`Ke=Rf+Beta*ERP`, `Kd_at=Kd*(1-Tax)`, `WACC=We*Ke+Wd*Kd_at`), EBIT, Taxes (`IAS 12`/`ASC 740`), NOPAT, `+D&A`, `-CapEx`, `-ΔNWC` → **UFCF**; Discount Factors (optional `-0.5` mid-year); **PV of UFCF**; **Terminal Value (Gordon Growth `=FinalUFCF*(1+g)/(WACC-g)` + Exit Multiple `=FinalEBITDA*Multiple`)**; **EV → Equity Bridge** (`EV - Debt - Leases + Cash - Minority - Preferred = Equity`; `Share Price = Equity / Shares`).
   - **`Sensitivity`:** 2-way formula tables for **WACC vs. Terminal Growth** and **WACC vs. Exit Multiple**.
   - **`Dashboard`:** **OUTPUTS ONLY** — Executive valuation summary linked to `'DCF Valuation'!` + `add_excel_table` + `create_chart`.
4. **Format, Save & Close:** `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, `0.0"x"`, widths, `freeze_panes="B2"`), `save_workbook`, `close_workbook(workbook_id=...)`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS (`IFRS 16`) / FASB (`ASC 842`) Parity:** Never mix pre-IFRS-16 EBITDA with post-IFRS-16 lease debt in the EV Bridge; enforce Terminal Growth `g < WACC`.
- **SOX 404:** Centralize ALL raw inputs exclusively on `Assumptions` (zero calculated results on `Assumptions`); zero hardcoded WACC, discount factors, Terminal Values, or EV Bridge outputs.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` on `Dashboard`, output ONLY a compact Valuation Summary table in plain finance language (never mention internal tool/skill names), and add a 1-line **Fresh-Chat Handoff Note**: *Saved to `<path>`. If context >50%, start a **New Chat** with this path.*
