---
name: boardroom-brief-generator
description: Convert a live-verified Excel financial model into a SOX-traceable executive Word memo (.docx + .pdf) and/or PowerPoint deck (.pptx) with zero copy-paste and full IFRS/FASB disclosure alignment via m365-mac-mcp.
---

# Boardroom Brief Generator (Excel → Word / PDF / PowerPoint)

## 1. Trigger & Scope
- **Triggers:** *"Turn this Excel model into a 1-page CFO memo"*, *"Create an executive Word/PDF brief and PowerPoint summary from this workbook."*
- **Deliverable:** A styled executive Word memo (`.docx` and optional `.pdf`) and/or PowerPoint slide deck (`.pptx`) whose figures reconcile 100% to live-recalculated Excel cells.

## 2. Mandatory Pre-Flight / HALT Inputs
Confirm (or HALT if unspecified):
- Source Excel workbook path and target output format(s) (`.docx`, `.pdf`, `.pptx`) + output folder.
- Audience & accounting basis label (**IFRS** or **US GAAP / FASB**, currency & units e.g. `$ in millions`).
- Verify `OFFICE_MCP_PROFILE` includes `word` (and `powerpoint` if slides are requested, e.g., `excel,word` or `full`).

## 3. Exact `m365-mac-mcp` Tool Sequence
1. **Audit & Read Live Numbers First:**
   - Call `audit_workbook_formulas(file_path=...)` to verify the source model is clean before reporting numbers.
   - Call `excel_live_open(file_path=...)` + `excel_live_batch_read_ranges` on all KPI/summary ranges so every metric comes verbatim from Microsoft Excel's native calculation engine.
2. **Generate Styled Word Memo (if requested):**
   - Call `create_document(title=...)`.
   - Add styled navy `#1F4E79` headings (`add_heading`, `level=1`) for *Executive Summary*, *Financial Highlights*, *Key Sensitivities & Risks*, and *Recommendation*.
   - Add concise executive narrative (`add_paragraph`, `add_list`) and KPI summary table via `add_table(document_id=..., rows=..., columns=..., data=..., column_widths=..., header_formatting={"bold": True, "bg_color": "1F4E79", "font_color": "FFFFFF"}, alternating_row_color="F2F6FA", repeat_header=True)`.
   - Call `save_document(..., file_path=".../Executive_Brief.docx")` (and `format="pdf"` if PDF requested), then `close_document`.
3. **Generate PowerPoint Deck (if requested):**
   - Call `create_presentation(title=...)` → `add_slide` + `add_text_to_slide` + `add_speaker_notes` (3–5 crisp slides) → `save_presentation`.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX Traceability & RAG Absolutism:** Every figure in the Word memo or slides MUST match the live Excel read verbatim—never estimate, round inconsistently, or extrapolate unverified numbers. State the source workbook filename and audit status (`clean: true`) in the memo footer/note.
- **IFRS / FASB Disclosure Clarity:** Clearly label reporting basis (IFRS vs. US GAAP), currency/units (`$M`), and any non-GAAP metrics (e.g., Adjusted EBITDA, FCF) reconciled to GAAP/IFRS Net Income or Operating Cash Flow.
- **Tool Hygiene:** NEVER read `.internal/skills/.../SKILL.md` or use Bionic's built-in viewer.

## 5. Verification & Output Gate
Confirm all output files saved cleanly and every KPI reconciles to Step 1's live Excel read, then output ONLY a 3-line summary with saved file paths and source model audit status.
