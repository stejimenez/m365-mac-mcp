---
name: capex-project-analyzer
description: Build an IFRS/FASB/SOX CapEx capital-budgeting workbook (NPV, IRR, MIRR, Payback, sensitivity, chart) via m365-mac-mcp.
---

# CapEx Project Analyzer (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Evaluate a \$20M CapEx project"*, *"Build an NPV/IRR/MIRR capital budgeting model"*, *"Analyze payback for a plant/equipment investment."*
- **Deliverable:** A 3-tab Excel model (`Executive Summary`, `Project Cash Flows`, `Sensitivity & Charts`) with live-verified `NPV`, `IRR`, `MIRR`, `PI`, `Payback`, and Go/No-Go decision.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm concisely before building:
- Accounting basis: **IFRS (`IAS 16`)** or **US GAAP (`ASC 360`)** + depreciation method (Straight-Line vs. MACRS).
- Initial Year-0 CapEx outlay, project life (`Years 1..N`), and annual Revenue / OpEx (or EBITDA) ramp.
- Corporate Tax Rate %, WACC %, Working Capital % of Revenue (recovered in Year `N`), and Terminal / Salvage Value.

## 3. Exact m365-mac-mcp Tool Sequence
1. **Create All 3 Tabs (1 Call):** Call `create_workbook(title=..., sheet_names=["Executive Summary", "Project Cash Flows", "Sensitivity & Charts"])`.
2. **Write Model (1 Call):** Call `batch_write_ranges` across all 3 tabs:
   - **`Executive Summary`:** Blue input block + dynamic KPI Decision Box:
     - **NPV:** `=NPV(WACC, 'Project Cash Flows'!C16:I16) + 'Project Cash Flows'!B16` *(Year 0 MUST stay outside `NPV()`)*.
     - **IRR / MIRR:** `=IRR('Project Cash Flows'!B16:I16)`, `=MIRR('Project Cash Flows'!B16:I16, WACC, ReinvestRate)`.
     - **Profitability Index:** `=(NPV - Year0Outlay) / -Year0Outlay`; formula-driven **Simple & Discounted Payback**; `=IF(NPV>0, "ACCEPT (NPV > 0)", "REJECT (NPV <= 0)")`.
   - **`Project Cash Flows` (`Year 0..N`):** Revenue, OpEx, EBITDA, Depreciation, EBIT, Tax (`IAS 12`/`ASC 740`), NOPAT, `+D&A`, `-ΔNWC` (recovered Year `N`), `-CapEx`, `+After-Tax Salvage` (`=Salvage - (Salvage - EndingBookValue) * TaxRate`), **Unlevered FCF**, **Cumulative UFCF**, **Discount Factor (`=1/(1+WACC)^t`)**, **PV of UFCF**, **Cumulative PV**.
   - **`Sensitivity & Charts`:** 2-way formula grid testing **WACC** (±200 bps, 50 bps steps) vs. **Revenue / CapEx** (±20%, 10% steps).
3. **Chart, Format, Save & Close:** Call `create_chart` (Annual & Cumulative Cash Flow chart), `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, widths, `freeze_panes="B2"`), `save_workbook`, and `close_workbook(workbook_id=...)`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS (`IAS 16`) / FASB (`ASC 360`):** Separate capitalizable CapEx from expensed start-up costs; tax salvage gain/loss against ending book value (`Cost - Accum. Depr.`); exclude sunk costs and financing interest (captured in WACC).
- **SOX 404:** Centralize all inputs on `Executive Summary`; zero hardcoded KPIs, tax shields, or salvage taxes.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` on `Executive Summary`, output ONLY a compact Go/No-Go KPI table in plain finance language (never mention internal tool/skill names), and add a 1-line **Fresh-Chat Handoff Note**: *Saved to `<path>`. If the bottom-right **Context** indicator is >50%, start a **New Chat** with this file path for follow-up analysis.*
