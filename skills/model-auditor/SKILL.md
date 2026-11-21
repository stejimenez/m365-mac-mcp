---
name: model-auditor
description: Perform a SOX-grade, read-only institutional audit of an Excel financial model for formula errors, broken links, hardcoded totals, IFRS/FASB classification issues, and balance-check breaks via m365-mac-mcp.
---

# Financial Model Auditor (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Audit this model"*, *"Check for formula errors and hardcoded totals"*, *"Run a SOX/IFRS/GAAP review before I send this workbook."*
- **Deliverable:** A strictly read-only, SOX-compliant **Traffic-Light Audit Report** (`PASS` / `WARN` / `FAIL`) identifying formula defects, control breaks, and IFRS/FASB accounting inconsistencies—without modifying the workbook.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm (or HALT only if missing):
- Target workbook path (or active `workbook_id`).
- Reporting framework to audit against: **IFRS** or **US GAAP (FASB)** (if unspecified, audit both general accounting identities and flag framework-sensitive items such as `IAS 7` vs. `ASC 230` cash flow lines or `IFRS 16` vs. `ASC 842` leases).

## 3. Exact `m365-mac-mcp` Tool Sequence
1. **Automated Formula & Link Scan:** Call `audit_workbook_formulas(file_path=...)` to flag:
   - Excel error tokens (`#REF!`, `#NAME?`, `#VALUE!`, `#DIV/0!`, `#N/A`, `#NUM!`, `#NULL!`).
   - External workbook links (`[...]`), broken/unquoted sheet references, and suspicious hardcoded literals inside formula columns or literal `"None"` / `'""'` strings.
2. **Structure & Sparse Formula Audit:** Call `inspect_workbook(file_path=...)` once to inspect the `values` grid and sparse `formulas` map on every tab for:
   - Hardcoded numbers inside Subtotal, Total, Margin, CAGR, or Balance Check rows.
   - Wrong denominators in margin or growth formulas (e.g., `=C4/B2` instead of `=C4/C2`).
   - Sign-convention breaks (CapEx, Debt, Dividends/Buybacks) and Balance Sheet imbalances (`Assets - L&E ≠ 0`).
3. **Optional Live Excel Verification:** If cached values are empty or live check is requested, call `excel_live_open(file_path=...)`, `excel_live_batch_read_ranges` on check cells, and optionally `excel_live_circle_invalid(clear=False)` to surface data-validation breaches.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX Section 404 Spreadsheet Controls:** Flag any undocumented plugs, hardcoded overrides inside calculated lines, circular references, or missing reconciliation checks (`Balance Check`, `Cash Flow → Balance Sheet Cash`).
- **IFRS vs. FASB Checks:** Verify Cash Flow classification consistency (`IAS 7` vs. `ASC 230` for interest/dividends), Current vs. Non-Current separation (`IAS 1` / `ASC 210`), Retained Earnings roll-forward integrity, and Lease/Debt consistency (`IFRS 16` / `ASC 842`).
- **Zero-Mutation Default:** Do NOT edit any cell unless explicitly instructed. If later asked to fix issues and a write returns `status: "confirmation_required"` (`OFFICE_MCP_FORMULA_GUARD`), **STOP** and ask the user for approval before retrying with `allow_formula_overwrite=true`.

## 5. Verification & Output Gate
Cross-check `audit_workbook_formulas` counts against `inspect_workbook` (and live check cells if opened), then output ONLY a compact Traffic-Light Audit Table (`Status | Sheet!Cell | Category (Formula / SOX / IFRS-FASB) | Issue | Recommended Fix`).
