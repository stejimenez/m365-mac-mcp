---
name: model-auditor
description: Read-only SOX, IFRS, and FASB audit of Excel models for formula errors, broken links, hardcoded totals, and balance breaks.
---

# Financial Model Auditor (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers:** *"Audit this model"*, *"Check for formula errors or hardcoded totals"*, *"Run a SOX/IFRS/GAAP review."*
- **Deliverable:** Strictly read-only Traffic-Light Audit Report (`PASS` / `WARN` / `FAIL`) covering formula integrity, SOX 404 controls, and IFRS/FASB consistency.

## 2. Mandatory Pre-Flight / HALT Inputs
- Target workbook path — if user says *"my open workbook"*, call `excel_live_list_workbooks()` to resolve the path automatically.
- Reporting framework (**IFRS** vs. **FASB / US GAAP**): if unstated, proceed immediately and flag framework-sensitive items (`IAS 7` vs. `ASC 230`, `IFRS 16` vs. `ASC 842`) without halting.

## 3. Exact m365-mac-mcp Tool Sequence
1. **Automated Formula & Link Audit:** Call `audit_workbook_formulas(file_path=...)` to detect `#REF!`, `#NAME?`, `#VALUE!`, `#DIV/0!`, `#N/A`, `#NUM!`, `#NULL!`, external `[...]` links, broken/unquoted sheet refs, and suspicious literals in formula columns.
2. **Structure & Sparse Formula Review:** Call `inspect_workbook(file_path=...)` once to inspect `values` and sparse `formulas` across every tab for:
   - Hardcoded numbers inside Subtotal, Total, Margin, CAGR, or Balance Check rows.
   - Wrong denominators in margins/growth (e.g., `=C4/B2` instead of `=C4/C2`).
   - Sign-convention breaks (CapEx, Debt, Dividends/Buybacks) and Balance Sheet imbalances (`Assets - L&E ≠ 0`).
3. **Optional Live Check:** If cached values are `null` or requested, call `excel_live_open(file_path=...)`, `excel_live_batch_read_ranges` on check cells, and optionally `excel_live_circle_invalid(clear=False)`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX Section 404 Controls:** Flag undocumented Balance Sheet plugs, hardcoded overrides inside calculated lines, circular references, and missing reconciliations (`Cash Flow → BS Cash`, `Balance Check`).
- **IFRS vs. FASB Checks:** Verify Cash Flow placement (`IAS 7` vs. `ASC 230`), Current/Non-Current separation (`IAS 1` / `ASC 210`), Retained Earnings roll-forward, and Lease/Debt parity (`IFRS 16` / `ASC 842`).
- **Read-Only Default:** Never edit cells unless asked; if asked to fix and `confirmation_required` fires, stop for approval.

## 5. Verification & Output Gate
Reconcile `audit_workbook_formulas` counts with `inspect_workbook` (and live reads if opened), then output ONLY a compact Traffic-Light Audit Table (`Status | Sheet!Cell | Category | Issue | Recommended Fix`) in plain finance language (never mentioning internal tool/skill names).
