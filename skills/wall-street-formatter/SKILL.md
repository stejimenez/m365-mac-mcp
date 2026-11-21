---
name: wall-street-formatter
description: Apply institutional Wall Street financial-model formatting (colors, borders, number formats, column widths, freeze panes) to any Excel workbook via m365-mac-mcp using plain language.
---

# Wall Street Financial Model Formatter

## Goal
Transform any raw or partially styled `.xlsx` / `.xlsm` workbook into an institutional, board-ready financial model using a single `batch_format_workbook` call—without modifying a single value or formula.

## Exact `m365-mac-mcp` Workflow
1. **Inspect Once:** Call `inspect_workbook(file_path=...)` (or by `workbook_id`) to map every worksheet's headers, used range, dollar rows, percentage/margin/CAGR rows, multiple (`x`) rows, and subtotal/total rows.
2. **Open:** Call `open_workbook(file_path=...)` if not already open in memory.
3. **Batch Format in ONE Call:** Call `batch_format_workbook(workbook_id=..., operations=[...])` across all worksheets:
   - `freeze_panes: "B2"` on every financial statement tab.
   - `column_widths`: `{"A": 44, "B:G": 15}` (adjust last column letter per sheet; widen `A` if labels are long).
   - `auto_color_financial_model: True` (automatically colors numeric inputs blue `#0000FF`, same-sheet formulas black `#000000`, cross-sheet links green `#008000`).
   - **Header Row 1:** `{"range": "A1:G1", "formatting": {"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E79", "horizontal": "center"}}`.
   - **Dollar rows:** `number_format: '$#,##0;($#,##0);"-"'` across all currency columns.
   - **Percentage / Margin / CAGR rows & columns:** `number_format: "0.0%"` across the **entire** contiguous range (e.g., `B5:G5`, `G2:G16`—never stop one column short and leave unformatted decimal cells).
   - **Multiples / Ratios (if present):** `number_format: '0.0"x"'`.
   - **Subtotal & Total rows:** `{"bold": True, "bg_color": "#D9E1F2", "border": {"top": "thin", "bottom": "double"}}` (use `{"top": "thin", "bottom": "thin"}` for intermediate subtotals if requested).
4. **Save & Close:** Call `save_workbook(workbook_id=...)` (with `OFFICE_MCP_LIVE_SYNC=auto`, an already-open Excel window automatically reloads and restores your exact view) and `close_workbook(workbook_id=...)`.

## Guardrails
- NEVER use Bionic's built-in file viewer; use `m365-mac-mcp` tools only.
- NEVER modify cell values or formulas in a formatting-only request.
- Check the `batch_format_workbook` response for any `"warnings"` key before saving.
- Output ONLY a concise summary table of sheets formatted, ranges styled, and number formats applied.
