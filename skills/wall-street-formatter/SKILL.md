---
name: wall-street-formatter
description: Format Excel financial statements Wall Street style (colors, borders, formats, widths) via m365-mac-mcp without altering values.
---

# Wall Street Financial Model Formatter

## 1. Trigger & Scope
- **Triggers:** *"Format this workbook Wall Street style"*, *"Make this model board-ready"*, *"Style these financial statements."*
- **Deliverable:** Institutional `.xlsx`/`.xlsm` styling aligned with IFRS (`IAS 1`) / FASB (`ASC 210`) presentation clarity and SOX 404 zero-mutation controls.

## 2. Mandatory Pre-Flight / HALT Inputs
- Target file path — if the user says *"the workbook I have open"*, call `excel_live_list_workbooks()` to get its path automatically (do not force the user to type a path).
- Currency symbol/scale (default `$#,##0;($#,##0);"-"` if USD; proceed without halting).

## 3. Exact m365-mac-mcp Tool Sequence
1. **Inspect Once:** Call `inspect_workbook(file_path=...)` to map headers, used range, currency rows, percentage/margin/CAGR ranges, multiples, and subtotals/totals.
2. **Open:** Call `open_workbook(file_path=...)` if not already loaded.
3. **Single Batch Format:** Call `batch_format_workbook(workbook_id=..., operations=[...])` across all tabs:
   - `freeze_panes: "B2"`, `column_widths`: `{"A": 44, "B:G": 15}` (match sheet's last column).
   - `auto_color_financial_model: True` (blue `#0000FF` inputs, black `#000000` formulas, green `#008000` cross-sheet links).
   - **Header Row 1:** `bold: True`, `font_color: "#FFFFFF"`, `bg_color: "#1F4E79"`, `horizontal: "center"`.
   - **Currency rows:** `number_format: '$#,##0;($#,##0);"-"'`.
   - **Percentage / Margin / CAGR:** `number_format: "0.0%"` across the **entire** contiguous range (e.g., `B5:G5`, `G2:G16`—never stop one column short).
   - **Multiples:** `number_format: '0.0"x"'`.
   - **Subtotals / Totals:** `bold: True`, `bg_color: "#D9E1F2"`, `border: {"top": "thin", "bottom": "double"}` (`thin`/`thin` for intermediate subtotals).
4. **Save & Close:** Call `save_workbook(workbook_id=...)` (auto-syncs open Excel window via `M365_MCP_LIVE_SYNC=auto`) and `close_workbook(workbook_id=...)`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX 404 Zero-Mutation:** Formatting-only runs MUST NEVER change any value, formula, or line classification.
- **IFRS (`IAS 1`) / FASB (`ASC 210`):** Preserve negative parentheses `($#,##0)`, Current vs. Non-Current visual hierarchy, and double-underline totals.
- Check `batch_format_workbook` `"warnings"` before saving.

## 5. Verification & Output Gate
Confirm `audit_workbook_formulas(file_path=...)` returns `clean: true` and `external_links_count: 0`, then output ONLY a compact Markdown table (in plain finance language, never mentioning internal tool/skill names) of sheets formatted, ranges styled, and number formats applied.
