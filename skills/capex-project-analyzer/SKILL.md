---
name: capex-project-analyzer
description: Build an IFRS/FASB/SOX-compliant Capital Budgeting & CapEx Analysis workbook (NPV, IRR, MIRR, Payback, Discounted Payback, 2-way sensitivity, and charts) via m365-mac-mcp.
---

# CapEx Project Analyzer (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Evaluate a \$20M CapEx project"*, *"Build an NPV/IRR/MIRR capital budgeting model"*, *"Analyze payback and returns for a new facility/equipment investment."*
- **Deliverable:** A 3-tab institutional Excel capital-budgeting model (`Executive Summary`, `Project Cash Flows`, `Sensitivity & Charts`) with live-verified `NPV`, `IRR`, `MIRR`, `PI`, `Payback`, and Go/No-Go decision.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm (or HALT if unspecified) before creating the workbook:
- Accounting standard: **IFRS (`IAS 16`)** or **US GAAP (`ASC 360`)** (and tax depreciation basis: Straight-Line vs. MACRS).
- Initial CapEx outlay (Year 0), project life (`Years 1..N`), and annual Revenue / OpEx (or EBITDA) ramp.
- Corporate Tax Rate %, Discount Rate (WACC %), Working Capital % of Revenue (with Year `N` NWC recovery), and Terminal / Salvage Value.

## 3. Exact `m365-mac-mcp` Tool Sequence
1. **Create Workbook:** Call `create_workbook(title=..., sheet_name="Executive Summary")` and `add_worksheet` for `Project Cash Flows` and `Sensitivity & Charts`.
2. **Write All 3 Tabs in ONE Call:** Call `batch_write_ranges`:
   - **`Executive Summary`:** Input block (blue inputs) + dynamic KPI Decision Box:
     - **NPV:** `=NPV(WACC, 'Project Cash Flows'!C16:I16) + 'Project Cash Flows'!B16` *(Year 0 MUST stay outside `NPV()`)*.
     - **IRR / MIRR:** `=IRR('Project Cash Flows'!B16:I16)`, `=MIRR('Project Cash Flows'!B16:I16, WACC, ReinvestRate)`.
     - **Profitability Index:** `=(NPV - Year0Outlay) / -Year0Outlay`.
     - **Simple & Discounted Payback** (formula-driven from cumulative cash flows) + `=IF(NPV>0, "ACCEPT (NPV > 0)", "REJECT (NPV <= 0)")`.
   - **`Project Cash Flows` (`Year 0..N`):** Revenue, Operating Costs, EBITDA, Depreciation, EBIT, Taxes (`IAS 12`/`ASC 740`), NOPAT, `+Depreciation`, `-ΔNWC` (recovered in Year `N`), `-Initial & Maintenance CapEx`, `+After-Tax Salvage Value` (`=Salvage - (Salvage - EndingBookValue) * TaxRate`), **Unlevered FCF**, **Cumulative UFCF**, **Discount Factor (`=1/(1+WACC)^t`)**, **PV of UFCF**, **Cumulative PV**.
   - **`Sensitivity & Charts`:** 2-way formula table testing **WACC** (±200 bps, 50 bps steps) vs. **Revenue / CapEx** (±20%, 10% steps).
3. **Chart & Format:** Call `create_chart` (Cash Flow & Cumulative NPV chart) and `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, widths, `freeze_panes="B2"`).
4. **Save:** Call `save_workbook`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS (`IAS 16`) / FASB (`ASC 360`) Capitalization:** Separate capitalizable CapEx from expensed start-up/opex costs; compute after-tax salvage using ending book value (`Cost - Accumulated Depreciation`). Exclude sunk costs and financing interest (already captured in WACC).
- **SOX 404 Controls:** Centralize every input on `Executive Summary`; zero hardcoded KPIs, tax shields, or salvage taxes; respect `OFFICE_MCP_FORMULA_GUARD`.
- **Excel Traps:** NEVER put Year 0 inside `=NPV(...)`; single-quote spaced sheet names; use `null` (never `"None"`) for blank cells.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` on `Executive Summary` KPIs, then output ONLY a compact Go/No-Go KPI summary table.
