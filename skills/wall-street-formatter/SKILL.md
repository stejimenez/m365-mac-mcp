---
name: wall-street-formatter
version: 2.3.1D
description: "[v2.3.1D] Format Excel financial statements Wall Street style (multi-table auto-detection, Smart Magnitude, colors, borders, widths) without altering values."
---

# Wall Street Financial Model Formatter (v2.3.1D)

## 1. Trigger & Scope
- **Triggers:** *"Format this workbook Wall Street style"*, *"Make this model board-ready"*, *"Style these financial statements."*
- **Deliverable:** Institutional `.xlsx`/`.xlsm` styling aligned with IFRS (`IAS 1`) / FASB (`ASC 210`) presentation clarity and SOX 404 zero-mutation controls.

## 2. Mandatory Pre-Flight / HALT Inputs
- Target file path — if the user says *"the workbook I have open"*, call `excel_live_list_workbooks()` to get its path automatically (do not force the user to type a path).
- Currency symbol/scale (default `$#,##0;($#,##0);"-"` if USD; proceed without halting).

## 3. Exact Office MCP Tool Sequence
1. **Inspect Once:** Call `inspect_workbook(file_path=...)` to inspect `detected_tables` (multi-table regions, header rows & label columns on ANY row/col), currency/magnitude cues, percentages, and subtotals/totals.
2. **Open:** Call `open_workbook(file_path=...)` if not already loaded.
3. **Single Batch Format:** Call `batch_format_workbook(workbook_id=..., operations=[{"auto_color_financial_model": true, "auto_fit_columns": true, "freeze_panes": "auto"}])` across all tabs (plus explicit per-sheet `styles` overrides ONLY if custom branding is requested):
   - Auto-detects title blocks, section banners, stacked & side-by-side table headers on **ANY row or column**, 2-way sensitivity matrices, blue `#0000FF` inputs, black `#000000` formulas, green `#008000` cross-sheet links, Smart Magnitude currency formats (with exact `$#,##0.00` per-share/unit protection), `0.0%` margins/rates, `0.00"x"` multiples, two-tier `#EDF2F8` subtotals & `#D9E1F2` double-underlined totals, unclipped auto-fitted widths, and dynamic `freeze_panes`.
4. **Save & Close:** Call `save_workbook(workbook_id=...)` (auto-syncs open Excel window via `M365_MCP_LIVE_SYNC=auto`) and `close_workbook(workbook_id=...)`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX 404 Zero-Mutation:** Formatting-only runs MUST NEVER change any value, formula, or line classification.
- **IFRS (`IAS 1`) / FASB (`ASC 210`):** Preserve negative parentheses `($#,##0)`, Current vs. Non-Current visual hierarchy, and double-underline totals.
- Check `batch_format_workbook` `"warnings"` before saving.

## 5. Verification & Output Gate
Confirm `audit_workbook_formulas(file_path=...)` returns `clean: true` and `external_links_count: 0`, then output ONLY a compact Markdown table (in plain finance language, never mentioning internal tool/skill names) of sheets formatted, ranges styled, and number formats applied.
