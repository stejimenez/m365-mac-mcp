---
name: capex-project-analyzer
description: Build an IFRS/FASB/SOX CapEx capital-budgeting workbook (NPV, IRR, MIRR, Payback, sensitivity, chart) via m365-mac-mcp.
---

# CapEx Project Analyzer (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Evaluate a \$20M CapEx project"*, *"Build an NPV/IRR/MIRR capital budgeting model"*, *"Analyze payback for a plant/equipment investment."*
- **Deliverable:** A 3-tab Excel model (`Executive Summary`, `Project Cash Flows`, `Sensitivity & Charts`) with live-verified `NPV`, `IRR`, `MIRR`, `PI`, `Payback`, and Go/No-Go decision.

## 2. Mandatory Pre-Flight / HALT Inputs
**HARD STOP (ZERO TOOLS):** If ANY item below is missing from the user's prompt, reply ONLY with 1–3 concise questions and call **ZERO tools** (do NOT call `create_workbook` or invent "standard" assumptions):
- Accounting basis: **IFRS (`IAS 16`)** or **US GAAP (`ASC 360`)** + depreciation method (Straight-Line vs. MACRS).
- Initial Year-0 CapEx outlay, project life (`Years 1..N`), and annual Revenue / OpEx (or EBITDA) ramp.
- Corporate Tax Rate %, WACC %, Working Capital % of Revenue (recovered in Year `N`), and Terminal / Salvage Value.

## 3. Exact m365-mac-mcp Tool Sequence
1. **Create Tabs (1 Call):** `create_workbook(title=..., sheet_names=["Executive Summary", "Project Cash Flows", "Sensitivity & Charts"])`.
2. **Write Model (1 `batch_write_ranges` Call/Sheet — 3 Calls Total; Labels/Inputs + `fill_formula` across ranges, ZERO mental math in `<think>`):**
   - **Call 1 — `Project Cash Flows` (`A1`, Cols `B..H` = `Year 0..N`):** R1 Headers, R2 Revenue (`fill_formula` `D2:H2`), R3 OpEx, R4 EBITDA, R5 Depr., R6 EBIT, R7 Tax (`IAS 12`/`ASC 740`), R8 NOPAT, R9 `+D&A`, R10 `-ΔNWC` (recovered Yr `N`), R11 `-CapEx`, R12 `+After-Tax Salvage` (`=Salvage-(Salvage-EndBV)*Tax`), R13 **UFCF**, R14 Cum. UFCF, R15 Disc. Factor (`=1/(1+WACC)^t`), R16 PV of UFCF, R17 Cum. PV.
   - **Call 2 — `Executive Summary` (`A1`):** Blue inputs + KPI Box linked to `'Project Cash Flows'!B13` & `C13:H13`:
     - **NPV:** `=NPV(WACC,'Project Cash Flows'!C13:H13)+'Project Cash Flows'!B13` *(Year 0 outside `NPV()`)*.
     - **IRR / MIRR / PI / Payback:** `=IRR(...)`, `=MIRR(...,WACC,WACC)`, `=(NPV-Yr0)/-Yr0`, Simple & Discounted Payback, **Go/No-Go:** `=NPV>0`.
   - **Call 3 — `Sensitivity & Charts` (`A1`):** 2-way **WACC** (±200 bps, 50 bps steps) vs. **CapEx/Rev** (±20%, 10% steps) via ONE 2D `fill_formula` op (no Python).
3. **Chart, Format, Save & Close:** `create_chart` (Cash Flow), `batch_format_workbook` (`auto_color_financial_model=True`, navy headers, `#D9E1F2` totals, `$#,##0;($#,##0);"-"`, `0.0%`, widths, `freeze_panes="B2"`), `save_workbook`, `close_workbook`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS (`IAS 16`) / FASB (`ASC 360`):** Separate capitalizable CapEx from expensed start-up costs; tax salvage gain/loss against ending book value (`Cost - Accum. Depr.`); exclude sunk costs and financing interest (captured in WACC).
- **SOX 404:** Centralize all inputs on `Executive Summary`; zero hardcoded KPIs, tax shields, or salvage taxes.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` on `Executive Summary`, output ONLY a compact Go/No-Go KPI table in plain finance language (never mention internal tool/skill names), and add a 1-line **Fresh-Chat Handoff Note**: *Saved to `<path>`. If the bottom-right **Context** indicator is >50%, start a **New Chat** with this file path for follow-up analysis.*
