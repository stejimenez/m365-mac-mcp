---
name: financial-statement-extractor
description: Read-only discovery of raw company Excel files (EN/FR/ES) to map historical years and verbatim accounts into an IFRS/FASB blueprint.
---

# Financial Statement Extractor & Account Mapper (IFRS / FASB / SOX)

## 1. Trigger & Scope
- **Triggers (Natural Language — EN / FR / ES):** *"Extract the historical financials from this file"* / *"Extrais les états financiers historiques de ce fichier"* / *"Extrae los estados financieros históricos de este Excel y mapéalos bajo IFRS/GAAP."*
- **Deliverable:** A strictly **read-only, compact Account Mapping Blueprint** (in the user's language) preserving every unique company account name **verbatim in its original language** across all available historical years—ready for approval before running `historical-3stmt-rebuilder` in the **same chat**.

## 2. Mandatory Pre-Flight / HALT Inputs
- Source Excel file path — if the user refers to their open workbook, call `excel_live_list_workbooks()` first to resolve the path automatically.
- Accounting framework (**IFRS** vs. **US GAAP / FASB**): auto-detect from statement labels across EN/FR/ES (*Revenue / Chiffre d'affaires / Ingresos*, *Current / Courant / Corriente*, *Finance Costs / Charges financières / Gastos Financieros*) and state the detected basis; ask only if genuinely ambiguous.

## 3. Exact m365-mac-mcp Tool Sequence
1. **Low-Token Tab Discovery (Mandatory on Multi-Tab Files):**
   - Call `open_workbook(file_path=...)` and `list_worksheets(workbook_id=...)` (~50 tokens) to locate ONLY the 3 core statements (EN/FR/ES names: *Income Statement / Compte de résultat / Estado de Resultados*, *Balance Sheet / Bilan / Balance General*, *Cash Flow / Flux de trésorerie / Flujo de Efectivo*)—skip Cover, Footnotes, Ratios, Quarterly, and Chart tabs.
2. **Targeted 1-Call Read (Only the 3 Statement Tabs — Both Values & Sparse Formulas):**
   - Call `inspect_workbook(workbook_id=..., sheet_names=[<IS_tab>, <BS_tab>, <CF_tab>], max_rows_per_sheet=150, max_cols_per_sheet=20)` in ONE call (returns both evaluated values and sparse A1 formulas so you never guess Input vs. Formula rows; omit `sheet_names` only if <=3 tabs). If any statement tab reports `truncated: true`, call `read_range` once for the remaining rows.
3. **Close Source Cleanly:** Call `close_workbook(workbook_id=...)` (read-only; never save or modify the raw file).
4. **Build Mapping Blueprint:** Group every line item—preserving the company's **exact verbatim account names**—into standard IFRS (`IAS 1`/`IAS 7`) or FASB (`ASC 210`/`ASC 230`) sections:
   - Tag **Input rows** vs. **Calculated Subtotals/Totals/Margins**.
   - Record available years (3, 5, 10+ yrs), currency/units, sign conventions, and any FX/M&A working-capital or Retained Earnings reconciliation lines.

## 4. Accounting (IFRS / FASB), SOX & Formula Guardrails
- **SOX 404 Non-Mutation & Fidelity:** NEVER modify the source workbook and NEVER rename, translate, merge, or fabricate company accounts—keep every unique reported account name and historical value verbatim.
- **IFRS / FASB Classification:** Explicitly flag framework-sensitive items (`IAS 7` vs. `ASC 230` interest/dividends; `IFRS 16` vs. `ASC 842` leases).

## 5. Verification & Output Gate
Verify every discovered tab's reported headline totals (`Revenue`, `Net Income`, `Total Assets`, `Total L&E`, `CFO`, `Ending Cash`) are captured verbatim, output ONLY the compact Mapping Blueprint in plain finance language (EN/FR/ES, never exposing internal tool/skill names), and end with: *"Stay in this same chat and reply to confirm (e.g., 'Build it' / 'Construis-le' / 'Constrúyelo') to generate the clean formula-driven 3-statement workbook."*
