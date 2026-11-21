---
name: wall-street-formatter
description: Apply institutional Wall Street financial-model formatting (colors, borders, number formats, widths, freeze panes) to any Excel workbook via m365-mac-mcp while preserving IFRS/FASB presentation standards and SOX spreadsheet controls.
---

# Wall Street Financial Model Formatter

## 1. Trigger & Scope
- **Triggers:** *"Format this workbook Wall Street style"*, *"Make this model board-ready"*, *"Apply institutional formatting to these financial statements."*
- **Deliverable:** A cleanly styled `.xlsx`/`.xlsm` workbook adhering to Wall Street color coding, IFRS (`IAS 1`) / FASB (`ASC 210`) financial-statement presentation clarity, and SOX Section 404 zero-mutation formatting controls.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm (or HALT only if target file is missing):
- Target workbook path (or active `workbook_id`).
- Currency symbol & scale label (`$M`, `€K`, etc. — default `$#,##0;($#,##0);"-"` if already USD).
- Whether to keep the file open in live Excel (`OFFICE_MCP_LIVE_SYNC=auto` reloads it automatically on save).

## 3. Exact `m365-mac-mcp` Tool Sequence
1. **Inspect Once:** Call `inspect_workbook(file_path=...)` to map every tab's headers, used range, dollar rows, percentage/margin/CAGR rows, multiple (`x`) rows, and subtotal/total rows.
2. **Open:** Call `open_workbook(file_path=...)` if not already loaded.
3. **Single Batch Format Call:** Call `batch_format_workbook(workbook_id=..., operations=[...])` across all tabs:
   - `freeze_panes: "B2"` on every statement tab; `column_widths`: `{"A": 44, "B:G": 15}` (adjust final column per sheet).
   - `auto_color_financial_model: True` (blue `#0000FF` hardcoded inputs, black `#000000` same-sheet formulas, green `#008000` cross-sheet links).
   - **Header Row 1:** `{"range": "A1:G1", "formatting": {"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E79", "horizontal": "center"}}`.
   - **Dollar rows:** `number_format: '$#,##0;($#,##0);"-"'` (negative values in parentheses per IFRS/FASB convention).
   - **Percentage / Margin / CAGR ranges:** `number_format: "0.0%"` across the **full** contiguous range (e.g., `B5:G5`, `G2:G16`—never stop one column short).
   - **Multiples / Ratios (if present):** `number_format: '0.0"x"'`.
   - **Subtotal / Total rows:** `{"bold": True, "bg_color": "#D9E1F2", "border": {"top": "thin", "bottom": "double"}}` (`{"top": "thin", "bottom": "thin"}` for intermediate subtotals).
4. **Save & Close:** Call `save_workbook(workbook_id=...)` (auto-reloads open Excel view via `OFFICE_MCP_LIVE_SYNC=auto`) and `close_workbook(workbook_id=...)`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX 404 Non-Mutation Rule:** Formatting-only runs MUST NEVER alter a single cell value, formula, or accounting classification.
- **IFRS (`IAS 1`) / FASB (`ASC 210`) Presentation:** Preserve negative parentheses `($#,##0)`, explicit Current vs. Non-Current visual hierarchy, and double-underline totals.
- **Tool Hygiene:** NEVER use Bionic's built-in viewer or `.internal/skills/.../SKILL.md`. Inspect any `"warnings"` returned by `batch_format_workbook` before saving.

## 5. Verification & Output Gate
Verify `audit_workbook_formulas(file_path=...)` reports `clean: true` and `external_links_count: 0`, then output ONLY a compact Markdown table listing each sheet formatted, ranges styled, and number formats applied.
