---
name: model-auditor
description: Perform a read-only institutional audit of an Excel financial model for formula errors, broken cross-sheet links, stray external links, hardcoded totals, and invalid data via m365-mac-mcp.
---

# Financial Model Auditor

## Goal
Run a fast, non-destructive due-diligence audit on any `.xlsx` / `.xlsm` workbook before it goes to management, clients, or an investment committee.

## Exact `m365-mac-mcp` Workflow
1. **Formula & Link Audit:** Call `audit_workbook_formulas(file_path=...)` to detect:
   - Formula errors (`#REF!`, `#NAME?`, `#VALUE!`, `#DIV/0!`, `#N/A`, `#NUM!`, `#NULL!`).
   - External workbook links (`[...]`) and non-existent/misspelled sheet references.
   - Suspicious hardcoded literals sitting inside predominantly-formula columns or literal `"None"` / `'""'` strings.
2. **Structure & Formula Map Inspection:** Call `inspect_workbook(file_path=...)` once to review the sparse `formulas` map and `values` grid on every tab. Check for:
   - Hardcoded numbers inside Subtotal / Total / Margin / CAGR / Balance Check rows.
   - Sign-convention inconsistencies (e.g., positive vs. negative CapEx, Debt, or Dividends).
   - Wrong denominators in margin or YoY growth formulas (e.g., dividing by `B2` instead of `C2`).
   - Balance Sheet check rows (`Assets - L&E`) that are non-zero.
3. **Optional Live Check (if workbook is open in Excel or live values are needed):**
   - Call `excel_live_open(file_path=...)` and `excel_live_batch_read_ranges` on key check cells (e.g., Balance Check, Cash reconciliation) to verify live calculated values.
   - Optionally call `excel_live_circle_invalid(clear=False)` to highlight data-validation violations in Excel.

## Guardrails
- **Strictly read-only by default:** Do NOT edit or overwrite any cell unless the user explicitly asks you to fix findings.
- If asked to fix findings and a write returns `status: "confirmation_required"` (`OFFICE_MCP_FORMULA_GUARD`), **STOP** and ask the user for approval before retrying with `allow_formula_overwrite=true`.
- Output a compact **Traffic-Light Audit Table** (`PASS` / `WARN` / `FAIL`) listing exact `Sheet!Cell`, issue, current formula/value, and recommended fix.
