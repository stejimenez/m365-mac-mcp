---
name: boardroom-brief-generator
description: Convert a live or saved Excel financial model into a styled executive Word memo (.docx + .pdf) and/or PowerPoint slide deck (.pptx) with zero manual copy-paste via m365-mac-mcp.
---

# Boardroom Brief Generator (Excel → Word / PDF / PowerPoint)

## Goal
Extract live recalculated KPIs from an Excel financial model and generate a clean, executive-ready Word memo (`.docx` + `.pdf`) and/or PowerPoint deck (`.pptx`)—guaranteeing 100% numeric consistency with the underlying model.

## Prerequisite
Ensure `OFFICE_MCP_PROFILE` includes `word` (and `powerpoint` if generating slides), e.g., `OFFICE_MCP_PROFILE=full` or `excel,word`.

## Exact `m365-mac-mcp` Workflow
1. **Read Verified Live Numbers from Excel:**
   - Call `excel_live_open(file_path=...)` + `excel_live_batch_read_ranges` on the key KPI/summary ranges so every number comes from Excel's real calculation engine (never from uncached `null` values or mental arithmetic).
   - Optionally call `audit_workbook_formulas(file_path=...)` first and note whether the model passed cleanly.
2. **Build Styled Word Executive Memo (if requested):**
   - Call `create_document(title=...)`.
   - Call `add_heading(..., level=1, formatting={"font_name": "Calibri", "font_color": "1F4E79"})` for Title & section headers (*Executive Summary*, *Financial Highlights*, *Key Risks & Sensitivities*, *Recommendation*).
   - Call `add_paragraph` / `add_list` with concise, data-backed executive bullets citing exact numbers from Step 1.
   - Call `add_table(document_id=..., rows=..., columns=..., data=..., column_widths=..., header_formatting={"bold": True, "bg_color": "1F4E79", "font_color": "FFFFFF"}, alternating_row_color="F2F6FA", repeat_header=True)`.
   - Call `save_document(..., file_path=".../Executive_Brief.docx")` and optionally export PDF via `save_document(..., file_path=".../Executive_Brief.pdf", format="pdf")`, then `close_document`.
3. **Build PowerPoint Executive Deck (if requested):**
   - Call `create_presentation(title=...)` → `add_slide` + `add_text_to_slide` + `add_speaker_notes` (3–5 crisp slides: Title, Financial Performance, Valuation/Returns, Sensitivities & Recommendation) → `save_presentation`.

## Guardrails
- **RAG Absolutism:** Every single number in the Word memo or PowerPoint slides MUST come verbatim from Step 1's Excel read—never estimate or invent missing metrics.
- Keep executive tables scannable (max 5–6 columns) and output the saved file paths in a 3-line confirmation.
