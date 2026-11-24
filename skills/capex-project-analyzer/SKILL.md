---
name: capex-project-analyzer
version: 2.3.1C
description: "[v2.3.1C] Build an IFRS/FASB/SOX CapEx capital-budgeting workbook (NPV, IRR, MIRR, Payback, sensitivity, chart) via m365-mac-mcp."
---

# CapEx Project Analyzer (IFRS / FASB / SOX) (v2.3.1C)

## 1. Trigger & Scope
- **Triggers:** *"Evaluate a \$20M CapEx project"*, *"Build an NPV/IRR/MIRR capital budgeting model."*
- **Deliverable:** 4-tab model (`Executive Summary`, `Assumptions`, `Project Cash Flows`, `Sensitivity & Charts`), pure-input `Assumptions`, live-verified `NPV`, `IRR`, `MIRR`, `PI`, `Payback`, Go/No-Go.

## 2. Mandatory Pre-Flight / HALT Inputs
**HARD STOP (ZERO TOOLS):** If ANY item below is missing, ask 1–3 concise questions and call **ZERO tools** (NEVER call `create_workbook` or invent assumptions):
- Basis (**IFRS `IAS 16`** / **US GAAP `ASC 360`**) + depr. method, Yr-0 CapEx, project life, Revenue/OpEx (or EBITDA) ramp, Tax %, WACC %, NWC % (recovered Yr `N`), Salvage.

## 3. Exact m365-mac-mcp Tool Sequence
1. **Create Tabs (1 Call):** `create_workbook(title=..., sheet_names=["Executive Summary", "Assumptions", "Project Cash Flows", "Sensitivity & Charts"])`.
2. **Write Model (1 `batch_write_ranges`/Sheet; `fill_formula` is an op key INSIDE `batch_write_ranges`; lock `Assumptions!$B$...` with `$`, ZERO mental math):**
   - **Call 1 — `Assumptions` (`A1`):** **RAW BLUE INPUTS ONLY (ZERO KPIs):** Basis, Depr. Method, Life, Yr-0 CapEx, Sunk Cost (excluded), Yr-1 Rev, Growth %, EBITDA %, Tax %, WACC %, NWC %, Salvage.
   - **Call 2 — `Project Cash Flows` (`A1:H1`=`["Year",0,1,2,3,4,5,6]`; Col `A`=ASCII Labels, `B`=Yr 0, `C..H`=Yrs 1..6; no `""` in numeric cells):** R2 Rev (`B2=0`, `C2=Assumptions!$B$...`, `D2:H2` `fill_formula`), R3 OpEx, R4 EBITDA, R5 Depr., R6 EBIT, R7 Tax (`IAS 12`/`ASC 740`), R8 NOPAT, R9 `+D&A`, R10 `-Change in NWC` (`B10=0`, `C10=-C2*NWC%`, `D10:G10=-(D2-C2)*NWC%`, **`H10=+G2*NWC%` recovery**), **R11 `-CapEx` (`B11=-Assumptions!$B$5` NEGATIVE in Yr 0; `C11:H11=0`)**, R12 `+After-Tax Salvage` (`B12:G12=0`, `H12=Salvage-(Salvage-EndBV)*Tax`), R13 **UFCF** (**`B13=B11`**, `C13:H13=C8+C9+C10+C11+C12`), R14 Cum. UFCF (`B14=B13`, `C14:H14` `fill_formula: "=B14+C13"`), R15 Disc. Factor (`B15=1`, `C15:H15=1/(1+WACC)^C$1`), R16 PV (**`B16=B13`**, `C16:H16=C13*C15`), R17 Cum. PV (`B17=B16`, `C17:H17` `fill_formula: "=B17+C16"`).
   - **Call 3 — `Executive Summary` (`A1`):** **OUTPUTS ONLY:** **NPV** (`B2`: `=NPV(Assumptions!$B$...,'Project Cash Flows'!C13:H13)+'Project Cash Flows'!B13`), **IRR** (`B3`), **MIRR** (`B4`), **PI** (`B5`: `=(B2+Assumptions!$B$5)/Assumptions!$B$5`), Simple & Discounted Payback (prefix ALL refs with `'Project Cash Flows'!`), **Go/No-Go** (`=IF(B2>0,"Go","No-Go")`).
   - **Call 4 — `Sensitivity & Charts` (`A1`):** R1 Title, R2 blank, `C3:G3` WACC (±200 bps, 100 bps steps), `B4:B8` CF Scale (80%–120%), `C4:G8` ONE 2D `fill_formula`: `=NPV(C$3,'Project Cash Flows'!$C$13:$H$13)*$B4+'Project Cash Flows'!$B$13`.
3. **Chart, Format, THEN Save (BEFORE `excel_live_open`):** `create_chart`, `batch_format_workbook` (`auto_fit_columns: true`, `auto_color_financial_model: true`, `freeze_panes: "B2"`, navy `#1F4E79` white bold headers, `#D9E1F2` bold totals, `$#,##0;($#,##0);"-"`, `0.0%`), `save_workbook`, `close_workbook`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS (`IAS 16`) / FASB (`ASC 360`):** Exclude sunk costs & financing interest; tax salvage vs. ending book value (`Cost - Accum. Depr.`).
- **SOX 404:** Raw inputs ONLY on `Assumptions` (zero KPIs on `Assumptions`); zero hardcoded KPIs, CapEx links, or salvage taxes.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` (`clean: true`, `external_links_count: 0`) AND `excel_live_open` + `excel_live_batch_read_ranges` on `Executive Summary`, output ONLY a compact Go/No-Go KPI table in plain finance language (no tool/skill names), plus 1-line **Fresh-Chat Handoff Note**: *Saved to `<path>`. If context >50%, start a **New Chat** with this path.*
