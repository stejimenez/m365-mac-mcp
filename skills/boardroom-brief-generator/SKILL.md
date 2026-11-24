---
name: boardroom-brief-generator
version: 2.3.1A
description: "[v2.3.1A] Turn live-recalculated Excel KPIs into a SOX-traceable executive Word/PDF memo or PowerPoint deck via m365-mac-mcp."
---

# Boardroom Brief Generator (Excel → Word / PDF / PowerPoint) (v2.3.1A)

## 1. Trigger & Scope
- **Triggers:** *"Turn this Excel model into a 1-page CFO memo"*, *"Create an executive Word/PDF brief and PowerPoint summary from this workbook."*
- **Deliverable:** A styled Word memo (`.docx` + optional `.pdf`) and/or PowerPoint deck (`.pptx`) whose numbers reconcile 100% to live-recalculated Excel cells.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm concisely (if missing):
- Source Excel workbook (or call `excel_live_list_workbooks()` if open) + desired format(s) (`.docx`, `.pdf`, `.pptx`) and save folder.
- Reporting basis label (**IFRS** or **US GAAP / FASB**) and currency/units (`$M`).
- Ensure the active MCP profile includes Word/PowerPoint (`M365_MCP_PROFILE=EXCEL,WORD` or `FULL`, or `@`-mention your secondary `m365-mac-mcp` Word & PPT server).

## 3. Exact m365-mac-mcp Tool Sequence
1. **Audit & Read Live Numbers:**
   - Call `audit_workbook_formulas(file_path=...)` to confirm the source model is clean.
   - Call `excel_live_open(file_path=...)` + `excel_live_batch_read_ranges` on all KPI/summary ranges so every number comes verbatim from live Microsoft Excel.
2. **Build Styled Word Memo (if requested):**
   - Call `create_document(title=...)`.
   - Add navy `#1F4E79` `add_heading(..., level=1)` sections (*Executive Summary*, *Financial Highlights*, *Risks & Sensitivities*, *Recommendation*).
   - Add crisp bullets (`add_paragraph`, `add_list`) and KPI table via `add_table(document_id=..., rows=..., columns=..., data=..., column_widths=..., header_formatting={"bold": True, "bg_color": "1F4E79", "font_color": "FFFFFF"}, alternating_row_color="F2F6FA", repeat_header=True)`.
   - Call `save_document(..., file_path=".../Executive_Brief.docx")` (and `format="pdf"` if requested), then `close_document`.
3. **Build PowerPoint Deck (if requested):**
   - Call `create_presentation(title=...)` → `add_slide` + `add_text_to_slide` + `add_speaker_notes` (3–5 slides) → `save_presentation`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX Traceability:** Every figure MUST match Step 1's live Excel read verbatim—never extrapolate or round inconsistently. Note the source workbook filename and plain-English formula check status (e.g., *"Formula Audit: Passed — 0 errors, 0 external links"*) in the memo footer; never print internal tool or skill names.
- **IFRS / FASB Disclosure:** Explicitly label accounting basis (IFRS vs. US GAAP), currency/units, and reconcile any non-GAAP metrics (Adjusted EBITDA, FCF) to GAAP/IFRS Net Income or CFO.

## 5. Verification & Output Gate
Confirm all output files saved cleanly and every KPI matches Step 1's live Excel read, then output ONLY a 3-line confirmation with saved file paths and audit status.
