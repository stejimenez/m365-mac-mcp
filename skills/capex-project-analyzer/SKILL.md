---
name: capex-project-analyzer
description: Build a complete Capital Budgeting / CapEx Investment Analysis workbook (NPV, IRR, MIRR, Payback, Discounted Payback, 2-way sensitivity, and charts) via m365-mac-mcp.
---

# CapEx Project Analyzer

## Goal
Create an institutional capital-budgeting Excel model that evaluates a proposed CapEx investment's after-tax unlevered cash flows, returns (`NPV`, `IRR`, `MIRR`, `Profitability Index`, `Payback`), sensitivity to WACC/Revenue, and Go/No-Go decision.

## Mandatory Pre-Flight Check (HALT if missing)
Confirm with the user before building:
- Initial CapEx outlay (Year 0) and project life (Years `1..N`).
- Expected annual Revenue & OpEx (or EBITDA) ramp.
- Depreciation method (Straight-Line vs. MACRS) and useful life.
- Corporate Tax Rate %, Discount Rate (WACC %), Working Capital % of Revenue (if any), and Terminal / Salvage Value.

## Exact `m365-mac-mcp` Workflow
1. **Create Workbook:** Call `create_workbook(title=..., sheet_name="Executive Summary")` and add worksheets `Project Cash Flows` and `Sensitivity & Charts`.
2. **Write Model in ONE Call:** Call `batch_write_ranges` across all 3 tabs:
   - **`Executive Summary`:** Input assumptions block (blue inputs) + KPI Decision Box linked dynamically to `Project Cash Flows`:
     - **NPV:** `=NPV(WACC, 'Project Cash Flows'!C16:I16) + 'Project Cash Flows'!B16` *(CRITICAL: Year 0 cash flow MUST stay outside Excel's `NPV()` function!)*.
     - **IRR:** `=IRR('Project Cash Flows'!B16:I16)`.
     - **MIRR:** `=MIRR('Project Cash Flows'!B16:I16, WACC, ReinvestRate)`.
     - **Profitability Index:** `=(NPV - Year0Outlay) / -Year0Outlay`.
     - **Simple & Discounted Payback** (formula-driven from cumulative cash flow rows).
     - **Decision Flag:** `=IF(NPV>0, "ACCEPT (NPV > 0)", "REJECT (NPV <= 0)")`.
   - **`Project Cash Flows`:** Columns `Year 0` through `Year N`:
     - Revenue, Operating Costs, EBITDA, Depreciation, EBIT, Tax (`=MAX(0, EBIT*TaxRate)` or full tax shield), NOPAT, `+Depreciation`, `-Change in NWC`, `-Initial & Maintenance CapEx`, `+After-Tax Salvage Value`, **Unlevered Free Cash Flow (UFCF)**, **Cumulative UFCF**, **Discount Factor (`=1/(1+WACC)^t`)**, **PV of UFCF**, **Cumulative PV**.
   - **`Sensitivity & Charts`:** 2-way formula grid testing **WACC** (±200 bps in 50 bps steps) vs. **Revenue / CapEx** (±20% in 10% steps).
3. **Charts & Formatting:**
   - Call `create_chart` (bar/column or line chart of Annual & Cumulative Cash Flows on `Sensitivity & Charts`).
   - Call `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, widths, `freeze_panes="B2"`).
4. **Save, Audit & Live Verify:** Call `save_workbook`, `audit_workbook_formulas`, and `excel_live_open` + `excel_live_batch_read_ranges` on `Executive Summary` KPIs to report live recalculated `NPV`, `IRR`, `MIRR`, `PI`, `Payback`, and Go/No-Go recommendation.

## Guardrails
- NEVER include Year 0 inside Excel's `NPV(rate, values)` range—always add Year 0 outside.
- Zero hardcoded KPIs or totals; use `null` (not `"None"`) for blank cells.
