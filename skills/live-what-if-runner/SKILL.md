---
name: live-what-if-runner
description: Run non-destructive live Excel sensitivity and scenario stress tests (recalculating via native Microsoft Excel) and automatically revert all inputs back to baseline.
---

# Live Excel What-If & Sensitivity Runner

## Goal
Test driver changes (e.g., Revenue ±10%, WACC ±100 bps, Margin ±200 bps, CapEx shocks) using Microsoft Excel's real calculation engine, capture exact recalculated deltas across all linked tabs, and leave the original workbook untouched.

## Mandatory Pre-Flight Check
Confirm (or HALT if unspecified):
- Target workbook path (or open workbook name).
- Driver cell(s) and shock values to test (e.g., `Income Statement!F2` = `450000`, or `±10%`).
- Output KPI cells to track (e.g., Net Income, FCF, Enterprise Value, Balance Check).

## Exact `m365-mac-mcp` Workflow
1. **Open Live:** Call `excel_live_open(file_path=..., activate=True)` (skip if already open).
2. **Capture Baseline in ONE Call:** Call `excel_live_batch_read_ranges(workbook_name=..., ranges=[...], include_formulas=True)` covering both the driver cells (to record exact original values) and all output KPI cells.
3. **Run Each Scenario:**
   - Write the shocked driver value(s) using `excel_live_write_cell(..., recalculate=True)` or `excel_live_batch_write_cells(..., recalculate=True)`.
   - Read all output KPI cells in ONE `excel_live_batch_read_ranges` call and record the exact recalculated values and deltas vs. baseline.
4. **Revert Cleanly:** Restore every driver cell to its exact baseline value with `recalculate=True`, and run ONE final `excel_live_batch_read_ranges` check confirming outputs match baseline identically.
5. **Leave Open (Default):** Do NOT call `excel_live_save` or `excel_live_close` unless the user explicitly asked to save/close.

## Guardrails
- **Never overwrite a formula cell as a driver:** If `excel_live_write_cell` returns `status: "confirmation_required"` (`OFFICE_MCP_FORMULA_GUARD`), STOP—the target cell is a formula, not an input! Ask the user which input cell to shock instead.
- Output a compact **Baseline vs. Scenario(s) vs. Reverted** Markdown table with exact numeric deltas.
