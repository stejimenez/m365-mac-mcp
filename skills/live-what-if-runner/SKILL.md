---
name: live-what-if-runner
description: Run non-destructive live Excel sensitivity stress tests with native recalculation and automatic baseline revert.
---

# Live Excel What-If & Sensitivity Runner

## 1. Trigger & Scope
- **Triggers:** *"Run a ±10% Revenue stress test"*, *"Test WACC ±100 bps live in Excel without saving"*, *"Check live 3-statement sensitivity."*
- **Deliverable:** A **Baseline vs. Scenario(s) vs. Reverted** KPI table recalculated natively in Microsoft Excel, with every input restored to baseline.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm (ask concisely only if missing):
- Target workbook — call `excel_live_list_workbooks()` first if the user refers to their open workbook.
- Driver input cell(s) and shock values/percentages (e.g., `Income Statement!F2` ±10%).
- Output KPI cells to track (always include `Balance Check` if a 3-statement model).

## 3. Exact m365-mac-mcp Tool Sequence
1. **Locate / Open Live:** Call `excel_live_list_workbooks()` or `excel_live_open(file_path=..., activate=True)`.
2. **Capture Baseline (1 Call):** Call `excel_live_batch_read_ranges(workbook_name=..., ranges=[...], include_formulas=True)` covering all driver cells (verifying they are inputs, not formulas) and all output KPI cells.
3. **Run Each Scenario:**
   - Apply shock via `excel_live_write_cell(..., recalculate=True)` or `excel_live_batch_write_cells(..., recalculate=True)`.
   - Read all tracked KPIs in ONE `excel_live_batch_read_ranges` call and compute exact deltas vs. baseline.
4. **Mandatory Revert & Proof:** Restore every driver cell to its exact baseline value (`recalculate=True`), then run ONE final `excel_live_batch_read_ranges` confirming all KPIs match baseline identically.
5. **Leave Open:** Do NOT call `excel_live_save` or `excel_live_close` unless explicitly asked.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX 404 Integrity:** Never leave test shocks saved in a production model; report Baseline, Shocked, Delta, and Reverted values as an audit trail.
- **IFRS (`IAS 1`) / FASB (`ASC 210`) Balance Check:** A valid 3-statement shock MUST keep `Balance Check (Assets - L&E)` unchanged (`0` if balanced at baseline).
- **Formula Guard:** If `confirmation_required` is returned, STOP—the targeted cell is a formula, not an input; ask which blue input cell to shock.

## 5. Verification & Output Gate
Verify the reverted read equals baseline (`Delta = 0`), then output ONLY a compact Markdown table (`Sheet!Cell | Metric | Baseline | Scenario(s) | Delta | Reverted | Status`) in plain finance language (never mentioning internal tool/skill names).
