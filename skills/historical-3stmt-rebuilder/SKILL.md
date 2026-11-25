---
name: historical-3stmt-rebuilder
version: 2.3.1D
description: "[v2.3.1D] Rebuild raw company Excel financials (EN/FR/ES) into a new IFRS/FASB/SOX formula-driven 3-statement model with Smart Magnitude & live proof."
---

# Historical 3-Statement Model Rebuilder (IFRS / FASB / SOX) (v2.3.1D)

## 1. Trigger & Scope
- **Triggers (Natural Language — EN / FR / ES):** *"Build the clean formula model from this company file"* / *"Construis le modèle 3 états avec formules"* / *"Reconstruye este Excel en un modelo IFRS/FASB limpio con todas las fórmulas y validaciones."*
- **Deliverable:** A **brand-new** `.xlsx` 3-statement workbook preserving the company's **unique verbatim account names**, with 100% live formulas, cross-statement links, data validations, and live Excel proof that rebuilt totals equal raw reported totals (`Delta = 0`).

## 2. Mandatory Pre-Flight / HALT Inputs
- If an approved Mapping Blueprint from `financial-statement-extractor` is already in this chat, reuse it directly (do NOT re-read the raw workbook).
- Otherwise, locate the source path (via `excel_live_list_workbooks()` if open), call `open_workbook` + `list_worksheets` + `inspect_workbook(workbook_id=..., sheet_names=[...], max_rows_per_sheet=150)` on the 3 statement tabs, then `close_workbook`.
- Confirm output path (default: `<SourceName>_Clean_Model.xlsx` in the same folder—**never** overwrite the raw source file) and IFRS vs. FASB basis if ambiguous.

## 3. Exact Office MCP Tool Sequence
1. **Create New 3-Tab Workbook + 4-Line Headers (1 Call):** Call `create_workbook(title=..., sheet_names=["Income Statement", "Balance Sheet", "Cash Flow"], header_metadata={"entity": "<Company>", "period": "Fiscal Years: ...", "currency": "<ISO>", "magnitude": "auto"})` *(tables start Row 6; keep EN/FR/ES tab names if requested)*.
2. **Populate All 3 Tabs (1 `batch_write_ranges` Call/Sheet — Labels/Inputs + `fill_formula`, ZERO mental math):**
   - Preserve every company-specific account name **verbatim** across all historical years plus trailing CAGR/Avg on IS.
   - Replace every hardcoded subtotal, total, margin (`%`), and CAGR with dynamic formulas; link `Cash Flow!Net Income` to `'Income Statement'!Net Income`.
   - Keep audited CF & Equity inputs faithful to the filing, and add **Validation Check** rows: `Balance Check (=Total Assets - Total L&E)`, `Cash Check (=BS Cash - CF Ending Cash)`, `NI Link Check`.
3. **Controls, Format & Save:** Call `add_data_validation`, `batch_format_workbook(operations=[{"auto_color_financial_model": true, "auto_fit_columns": true, "freeze_panes": "auto"}])`, `save_workbook(file_path="<Company>_Clean_Model.xlsx")` (resolve `sanity_warnings`), and `close_workbook`.
4. **Live Reconciliation Gate:** Call `audit_workbook_formulas(file_path=...)` (`clean: true`, `external_links_count: 0`), then `excel_live_open` + `excel_live_batch_read_ranges` on all rebuilt totals and check rows. Compare live rebuilt totals against the raw source file's reported totals—fix any row-index typo before sign-off.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **IFRS / FASB Fidelity:** Preserve `IAS 1`/`ASC 210` Current vs. Non-Current structure and `IAS 7`/`ASC 230` Cash Flow classification; reconcile any historical FX/M&A or Retained Earnings delta in an explicit labeled line (zero silent plugs per SOX 404).
- **Multilingual Support:** Understand prompts and source labels in EN/FR/ES, preserve exact account names, and respond in the user's language.

## 5. Verification & Output Gate
Output ONLY a compact **Rebuilt vs. Source Reconciliation Table** (`Metric | Years Verified | Max Delta | Balance Check | Audit`) in plain finance language (EN/FR/ES, never exposing internal tool/skill names), followed by this **Fresh-Chat Handoff Note** (localized to EN/FR/ES):
> *Saved to `<path>`. Next, you can ask to format it Wall Street style, project multi-year statements, or build a DCF valuation. If your chat context is above ~50%, start a **New Chat** and share this file path for best speed and accuracy.*
