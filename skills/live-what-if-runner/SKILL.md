---
name: live-what-if-runner
description: Run non-destructive, SOX-auditable live Excel sensitivity and scenario stress tests via native Microsoft Excel recalculation and automatically revert all inputs to baseline.
---

# Live Excel What-If & Sensitivity Runner

## 1. Trigger & Scope
- **Triggers:** *"Run a ±10% Revenue stress test"*, *"Test WACC ±100 bps live in Excel without changing the file"*, *"Check live 3-statement sensitivity."*
- **Deliverable:** A clean **Baseline vs. Scenario(s) vs. Reverted** KPI table computed by Microsoft Excel's native calculation engine, with all inputs provably restored to baseline.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm (or HALT if unspecified):
- Target workbook path (or open workbook name in Excel).
- Exact input driver cell(s) and shock values/percentages (e.g., `Income Statement!F2` = `450000`, or `±10%`).
- Output KPI cells to track (e.g., Net Income, CFO, FCF, Enterprise Value, and mandatory `Balance Check`).

## 3. Exact `m365-mac-mcp` Tool Sequence
1. **Open Live:** Call `excel_live_open(file_path=..., activate=True)` (skip if already open in Excel).
2. **Capture Baseline in ONE Call:** Call `excel_live_batch_read_ranges(workbook_name=..., ranges=[...], include_formulas=True)` covering all driver cells (to record exact original values and confirm they are inputs, not formulas) and all output KPI cells.
3. **Execute Each Scenario:**
   - Write shocked input(s) via `excel_live_write_cell(..., recalculate=True)` or `excel_live_batch_write_cells(..., recalculate=True)`.
   - Read all output KPIs in ONE `excel_live_batch_read_ranges` call and compute exact numeric deltas vs. baseline.
4. **Mandatory Revert & Proof:** Restore every driver cell to its exact baseline value with `recalculate=True`, then run ONE final `excel_live_batch_read_ranges` call proving every tracked KPI matches baseline identically.
5. **Leave Open (Default):** Do NOT call `excel_live_save` or `excel_live_close` unless explicitly instructed.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX Change-Control & Integrity:** Never leave test shocks persisted in a production workbook; always log Baseline, Shocked, Delta, and Reverted values as an audit trail.
- **Accounting Identity Check:** Always include `Balance Check (Assets - L&E)` in tracked KPIs—under both IFRS (`IAS 1`) and FASB (`ASC 210`), a valid 3-statement shock MUST keep `Balance Check` unchanged (`0` if balanced at baseline).
- **`OFFICE_MCP_FORMULA_GUARD` Compliance:** If `excel_live_write_cell` returns `status: "confirmation_required"`, **STOP**—the target cell contains a formula, not an input! Ask the user which blue input cell to shock instead.
- **Tool Hygiene:** NEVER use Bionic's built-in viewer or shell `osascript` loops.

## 5. Verification & Output Gate
Confirm the final reverted read matches baseline 100% (`Delta = 0` on revert), then output ONLY a compact Markdown table (`Sheet!Cell | Metric | Baseline | Scenario(s) | Delta | Reverted | Status`).
