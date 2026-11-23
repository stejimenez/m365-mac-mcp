# Model-Agnostic & Client-Agnostic Stress-Test Suite (`m365-mac-mcp`)

A comprehensive, **9-Phase (0A–0C Baseline + 1–8 Institutional) / 12-Chat** stress-test battery designed to benchmark **any LLM** (local or cloud — e.g., Qwen, Gemma, Llama, Nemotron, Claude, GPT, Gemini) on **any MCP client** (Bionic, LM Studio, Claude Desktop, Cline, VS Code, Cursor, Antigravity, etc.) connected to **`m365-mac-mcp`**.

> **Scope:** Starts with fast baseline smoke tests (**Phases 0A–0C**) achieving **100% coverage across all 82 MCP tools** (`EXCEL,COMPACT` 25 tools, `EXCEL` Full 50 tools, and `WORD,POWERPOINT` 33 tools), then escalates into heavy institutional finance builds, live recalculation, adversarial audits, and cross-app boardroom deliverables (**Phases 1–8**) covering all 7 non-import Finance Skills (`capex-project-analyzer`, `dcf-investment-template-builder`, `three-statement-projector`, `live-what-if-runner`, `model-auditor`, `wall-street-formatter`, `boardroom-brief-generator`) **without requiring external company financial statement imports**.

---

## 0. Pre-Test Setup (Any MCP Client, Any Model)

1. **Create an Isolated Output Folder:**
   ```bash
   mkdir -p ~/Desktop/MCP_StressTest
   ```
2. **Load the Finance Skills (Recommended):**
   - Import or point your MCP client to the `skills/` directory in this repository so the model executes the documented IFRS/FASB/SOX workflows.
3. **Configure MCP Server Profiles:**
   - **Excel Server (Active for Phases 0A–0B and 1–8):** `M365_MCP_PROFILE=EXCEL,COMPACT` *(25 tools)*, `M365_MCP_LIVE_SYNC=AUTO`, `M365_MCP_FORMULA_GUARD=literals_only`. *(Temporarily switch to `M365_MCP_PROFILE=EXCEL` only if running optional Phase 0B, then switch back to `EXCEL,COMPACT` before Phase 1.)*
   - **Word & PowerPoint Server (Active for Phase 0C and Phase 7 Only):** `M365_MCP_PROFILE=WORD,POWERPOINT` *(33 tools)* — keep **OFF** (or on-demand / hidden unless mentioned) during Phases 0A–0B, 1–6, and 8 so Excel context stays lean.
   - *(If your MCP client uses a single server entry instead of two, switch `M365_MCP_PROFILE` to `FULL` or `EXCEL,WORD` for Phases 0C and 7 and switch back to `EXCEL,COMPACT` afterward.)*
4. **Optional Live Wire Telemetry (Client-Agnostic):**
   - Set `M365_MCP_DEBUG_LOG=/path/to/mcp_debug_private.jsonl` on your MCP server to record every tool call, arguments, millisecond latency, warnings, and errors independently of your MCP client UI.
5. **Fresh-Chat Rule:**
   - Start a **brand-new chat** for each chat block below (`Chat 0A` through `Chat 9`, 12 chats total) so context stays clean and fresh-chat handoffs are tested realistically.

---

## MCP Profile Toggle Map

| Phase | Chats | Excel MCP Profile | Word & PPT MCP (`WORD,POWERPOINT`) | What You Are Hammering |
| :--- | :---: | :---: | :---: | :--- |
| **Phase 0A** | Chat 0A | **`EXCEL,COMPACT`** *(ON)* | **OFF** | Basic 25-tool Excel smoke test: status, 2-tab build, `fill_formula`, table, dropdown, chart, format, live recalc |
| **Phase 0B** *(Optional)* | Chat 0B | **`EXCEL`** *(50 tools, ON)* | **OFF** | Extended Excel tools: rename/activate sheet, named cell styles, table-style update, validation list, view capabilities |
| **Phase 0C** | Chat 0C | **OFF** *(or ON)* | **ON** | Basic Word (`.docx` + `.pdf`, styled table, search/replace, live Word) & PowerPoint (3-slide deck + speaker notes) |
| **Phase 1** | Chat 1 | **`EXCEL,COMPACT`** *(ON)* | **OFF** | Hard Stop on missing drivers, sunk-cost trap, CapEx NPV/IRR/MIRR, `fill_formula`, chart, live recalc |
| **Phase 2** | Chat 2 | **`EXCEL,COMPACT`** *(ON)* | **OFF** | Heavy 5-tab DCF, `IFRS 16` lease parity, dropdowns, `g ≥ WACC` pushback |
| **Phase 3** | Chat 3 | **`EXCEL,COMPACT`** *(ON)* | **OFF** | Full 3-Statement build from scratch, quoted sheet refs, live `Balance Check = 0` |
| **Phase 4** | Chat 4 | **`EXCEL,COMPACT`** *(ON)* | **OFF** | Live what-if + revert, **Formula Guard**, **Live Auto-Sync & unsaved-edit refusal** |
| **Phase 5** | Chat 5 | **`EXCEL,COMPACT`** *(ON)* | **OFF** | Sabotaged-model read-only audit, `circle invalid`, multilingual (Spanish) |
| **Phase 6** | Chat 6 | **`EXCEL,COMPACT`** *(ON)* | **OFF** | Pure Wall Street formatter, zero-mutation, contiguous `%` & multiple formatting |
| **Phase 7** | Chats 7–8 | **`EXCEL,COMPACT`** *(ON)* | **ON** | Cross-app Excel → Word + PDF + PowerPoint traceability + Live Word automation |
| **Phase 8** | Chat 9 | **`EXCEL,COMPACT`** *(ON)* | **OFF** | Adversarial error recovery, bad paths, sheet deletion, Python/shell bypass bait |

### File Lineage & Dependency Map (All Files Live in `~/Desktop/MCP_StressTest/`)
| File | Created In | Consumed / Mutated In |
| :--- | :---: | :--- |
| `Smoke_Excel_Compact.xlsx` | Phase 0A | Phase 0A-2 *(live edit/save/close)* → Phase 0B *(optional 50-tool check)* |
| `Smoke_Memo.docx` / `.pdf` & `Smoke_Live_Note.docx` & `Smoke_Deck.pptx` | Phase 0C | Standalone baseline Word/PDF/PPT smoke outputs |
| `CapEx_Stress.xlsx` | Phase 1 | Phase 8 *(Prompt 9B bypass bait — keep saved/closed in Excel)* |
| `Apex_DCF.xlsx` | Phase 2 | Phase 6 *(zero-mutation Wall St. format)* → Phase 7 Chat 7 *(live read → CFO Brief & Deck)* |
| `NovaCorp_3Stmt.xlsx` | Phase 3 | Phase 4 *(live what-if, Formula Guard, Auto-Sync)* → Phase 5 *(duplicated to `NovaCorp_Sabotaged.xlsx`)* |
| `NovaCorp_Sabotaged.xlsx` | Phase 5 *(manual copy)* | Phase 5 *(Spanish read-only SOX audit + `circle invalid`)* |
| `Apex_CFO_Brief.docx` / `.pdf` & `Apex_Board_Deck.pptx` | Phase 7 (Chat 7) | Phase 7 Chat 8 *(live Word edit/replace/save on `Apex_CFO_Brief.docx`)* |
| `Ghost_Model_999.xlsx` | *(Does not exist)* | Phase 8 *(Prompt 9A 1-call error recovery test)* |

---

## Phase 0A (Chat 0A) — Baseline `EXCEL,COMPACT` Smoke Test (All 25 Compact Tools)
**MCP State:** Excel (`M365_MCP_PROFILE=EXCEL,COMPACT`) **ON** | Word & PPT **OFF**

### Prompt 0A-1 — Core Excel File-Level + Live Recalculation Sanity Check
> `First confirm the active Office connection and status. Then create a small 2-tab Excel workbook saved to ~/Desktop/MCP_StressTest/Smoke_Excel_Compact.xlsx with tabs "Summary" and "Quarterly Sales". On "Quarterly Sales", add headers ["Region", "Q1", "Q2", "Q3", "Q4", "Full Year"] and 3 rows: North (100, 120, 130, 150), South (80, 90, 95, 110), West (60, 75, 85, 100), plus a "Total" row and a "Full Year" column using dynamic SUM formulas. Format it as a clean Excel Table with a bar chart, add a scenario dropdown ("Base", "Upside", "Downside") in cell B2 of "Summary" along with a cross-sheet link to the Grand Total on "Quarterly Sales", apply clean currency formatting and column widths, save, audit all formulas, and open it live in Excel to report the exact live-recalculated Grand Total.`

**Hammer Checks:**
1. **Status & Setup:** Calls `check_office_status` (confirms `excel,compact`, 25 tools) and creates both tabs in 1 `create_workbook(sheet_names=...)` call.
2. **Batch Write + `fill_formula`:** Uses `batch_write_ranges` (ideally with `fill_formula` for `Full Year` and `Total`) — zero cell-by-cell loops and zero hardcoded totals.
3. **Core Compact Features:** Exercises `add_excel_table`, `create_chart`, `add_data_validation`, `batch_format_workbook`, `save_workbook`, `close_workbook`, and `audit_workbook_formulas` (`0` errors, `0` external links).
4. **Live Recalc:** Calls `excel_live_open` + `excel_live_batch_read_ranges` on the saved file (no `/tmp` live open) and reports the exact live Grand Total (**`1,195`**).

### Prompt 0A-2 — Live Excel Write, Live Save & Clean Close (`excel_live_save` & `excel_live_close`) *(same chat)*
> `With Smoke_Excel_Compact.xlsx still open in Excel, update North Q4 (E2 on "Quarterly Sales") live to 160, confirm the live Grand Total updates from 1,195 to 1,205, save the workbook live in Excel, and then close the live Excel window cleanly without prompts.`

**Hammer Checks:**
1. Exercises `excel_live_write_cell` (or `excel_live_batch_write_cells`), `excel_live_batch_read_ranges` (verifying `1,205`), **`excel_live_save`**, and **`excel_live_close`** (clean close with zero modal dialogs; verifies boolean or string `save_changes` handling).

---

## Phase 0B (Chat 0B — Optional) — Extended `EXCEL` Full Profile Smoke Test (All 25 Extra Excel Tools)
**MCP State:** Excel (`M365_MCP_PROFILE=EXCEL`, 50 tools) **ON** | Word & PPT **OFF**  
*(Skip this chat if you only deploy `EXCEL,COMPACT`. If you run it, switch `M365_MCP_PROFILE` back to `EXCEL,COMPACT` before Phase 1.)*

### Prompt 0B-1 — Granular File-Level Excel Tools (Styles, Tables, Validations & Sheet Ops)
> `Open ~/Desktop/MCP_StressTest/Smoke_Excel_Compact.xlsx and list active in-memory workbooks. Read the single cell F5 on "Quarterly Sales" and read all structured Excel tables in the workbook. Add a formula in F6 calculating North's share of the Grand Total (=F2/F5), set explicit column widths (A=22, B:F=14), apply a subtle range format to F6 as percentage ("0.0%"), and apply worksheet-level freeze panes at B2. List available built-in cell styles, apply "Heading 1" to A1:F1 and "Total" to A5:F5 on "Quarterly Sales", and update the table style to TableStyleMedium2. On "Summary", list existing data validation rules, clear them, re-add the ("Base", "Upside", "Downside") dropdown on B2, rename "Summary" to "Exec Summary", set it as the active tab, save, and close.`

**Hammer Checks:**
1. Exercises all remaining file-level Excel tools: `open_workbook`, `list_active_workbooks`, `read_cell`, `read_tables`, `add_formula`, `set_column_widths`, `format_range`, `format_worksheet`, `list_cell_styles`, `apply_cell_style`, `update_excel_table_style`, `list_data_validations`, `clear_data_validations`, `rename_worksheet`, and `set_active_worksheet`.
2. **Integrity After Rename:** Confirm zero `#REF!` errors after renaming `Summary` → `Exec Summary`.

### Prompt 0B-2 — Live Excel Extended Tools (`excel_live_create` & Live Sheet/View Ops) *(same chat)*
> `Now test extended live Excel automation: create a brand-new blank workbook directly in live Microsoft Excel, list its worksheets, add a new live tab called "Scratch", rename it to "Live Check", activate it, write a 2×2 matrix [["Metric", "Val"], ["Test", 42]] to A1:B2 live, read A1:B2 back live, check Excel's view capabilities and custom views, then close that scratch workbook without saving. Finally, open Smoke_Excel_Compact.xlsx live in Excel, update its table style live to TableStyleMedium9, and close it without saving.`

**Hammer Checks:**
1. Exercises all remaining live Excel tools: `excel_live_create`, `excel_live_list_worksheets`, `excel_live_add_worksheet`, `excel_live_rename_worksheet`, `excel_live_activate_worksheet`, `excel_live_write_range`, `excel_live_read_range`, `excel_view_capabilities`, `excel_live_custom_views`, `excel_live_update_table_style`, and `excel_live_close`.
2. **Post-Chat Cleanup:** Switch `M365_MCP_PROFILE` back to **`EXCEL,COMPACT`** before starting Phase 1.

---

## Phase 0C (Chat 0C) — Baseline `WORD,POWERPOINT` Smoke Test (All 33 Word & PPT Tools)
**MCP State:** Excel **OFF** *(or ON)* | **Word & PPT (`M365_MCP_PROFILE=WORD,POWERPOINT`) ON**  
*(Optional for Prompt 0C-3 image test: place any small sample image at `~/Desktop/MCP_StressTest/sample_chart.png`.)*

### Prompt 0C-1 — Word File-Level Full Surface (Granular Paragraphs, Tables, Properties & PDF)
> `Create a styled Word memo titled "Q4 Regional Sales Summary", set its document properties (author "Finance Team", subject "Regional Review"), and list active documents. Add a navy Level-1 heading "Highlights", an introductory paragraph formatted in 11pt Calibri dark gray, a 3-bullet summary list, and a formatted 4-row × 3-column table comparing North ($500M), South ($375M), and West ($320M) with navy header shading, alternating row shading, explicit column widths, and a repeating header row. Then update the North table cell to "$510M", refine the table formatting, insert a short "Confidential — Internal Review" paragraph right before the table, add a temporary "DRAFT DELETE ME" paragraph at the end and delete it, replace "Q4 Regional Sales Summary" with "FY2025 Regional Sales Summary", and save to both ~/Desktop/MCP_StressTest/Smoke_Memo.docx and ~/Desktop/MCP_StressTest/Smoke_Memo.pdf. Close it, then re-open Smoke_Memo.docx and read back the document structure to confirm "$510M" is present and "DRAFT DELETE ME" is gone.`

**Hammer Checks:**
1. Exercises **100% of file-level Word tools**: `create_document`, `list_active_documents`, `set_document_properties`, `add_heading`, `add_paragraph`, `format_paragraph`, `add_list`, `add_table`, `format_table`, `update_table_cell`, `insert_paragraph`, `delete_paragraph`, `replace_document_text`, `save_document` (`.docx` + `.pdf`), `close_document`, `open_document`, and `read_document`.

### Prompt 0C-2 — Live Word Full Surface (`word_live_create` + Live Edit/Replace/Save/Close) *(same chat)*
> `Create a brand-new blank document directly in live Microsoft Word, list open Word documents, write "Preliminary Regional Notes — North leading Q4." into it, replace "Preliminary" with "Final", read the text back live to verify, save it to ~/Desktop/MCP_StressTest/Smoke_Live_Note.docx, and close it cleanly. Then open Smoke_Memo.docx live in Word, read the first paragraph to confirm it opens cleanly, and close it without saving.`

**Hammer Checks:**
1. Exercises **100% of live Word tools**: `word_live_create`, `word_live_list_documents`, `word_live_write`, `word_live_replace_text`, `word_live_read`, `word_live_save`, `word_live_close`, and `word_live_open` without AppleScript timeouts.

### Prompt 0C-3 — PowerPoint Full Surface (Slides, Text, Image & Speaker Notes) *(same chat)*
> `Now create a 3-slide PowerPoint presentation saved to ~/Desktop/MCP_StressTest/Smoke_Deck.pptx titled "FY2025 Regional Sales Overview" and list active presentations: Slide 1 title slide, Slide 2 key regional totals (North $510M, South $375M, West $320M, Total $1,205M) — also insert ~/Desktop/MCP_StressTest/sample_chart.png on Slide 2 if that file exists — and Slide 3 key takeaways, with concise speaker notes on every slide.`

**Hammer Checks:**
1. Exercises **100% of PowerPoint tools**: `create_presentation`, `list_active_presentations`, `add_slide`, `add_text_to_slide`, `add_image_to_slide` *(if sample image is present)*, `add_speaker_notes`, and `save_presentation`.
2. Open `Smoke_Deck.pptx` in PowerPoint and verify all 3 slides render cleanly with speaker notes populated.
3. **Post-Chat Cleanup:** Turn **Word & PPT MCP** back **OFF** (and ensure **Excel `EXCEL,COMPACT`** is **ON**) before starting **Phase 1**.

---

## Phase 1 (Chat 1) — Hard Stop, Sunk-Cost Trap & CapEx Build (`capex-project-analyzer`)
**MCP State:** Excel **ON** | Word & PPT **OFF**

### Prompt 1A — Incomplete Drivers (Hard Stop Test)
> `Evaluate a $25M manufacturing line expansion in Mexico over 6 years and save the model to ~/Desktop/MCP_StressTest/CapEx_Stress.xlsx.`

**Pass Criteria:**
- Must **STOP immediately with ZERO tool calls** (must NOT call `create_workbook` or fabricate "standard" assumptions).
- Must ask concisely (1–3 plain-finance bullets) for the missing core drivers: accounting basis & depreciation method, revenue/EBITDA ramp, tax rate, WACC, NWC %, and salvage value.
- Must **NOT** leak internal tool names, skill slugs, JSON, `openpyxl`, or AppleScript.

### Prompt 1B — Full Drivers + Sunk Cost & Salvage Tax Traps *(same chat)*
> `Use US GAAP with Straight-Line depreciation over 6 years. Year-0 CapEx is $25M, plus we already spent $2M last year on R&D feasibility studies for this line. Year-1 revenue is $10M growing 12% per year, 38% EBITDA margin, 25% corporate tax rate, 10.5% WACC, 8% NWC of revenue (fully recovered in Year 6), and $4M pre-tax salvage value at Year 6 (fully depreciated to $0 book value). Build it, check it live in Excel, and give me the Go/No-Go.`

**Hammer Checks:**
1. **Sunk Cost Exclusion:** Excludes the \$2M prior R&D spend from Year-0 CapEx (depreciates \$25M / 6 = \$4.167M/yr) and flags why in plain finance language.
2. **Year-0 NPV Rule:** Keeps Year 0 *outside* `=NPV(...)` (e.g., `=NPV(WACC, C13:H13) + B13`).
3. **After-Tax Salvage Formula:** Computes after-tax salvage via formula (`=Salvage - (Salvage - EndingBV) * TaxRate` → \$3.0M), not hardcoded.
4. **Batch & `fill_formula` Discipline:** Uses `create_workbook(sheet_names=[...])`, 1 `batch_write_ranges` call per sheet (ideally with `fill_formula`), `create_chart`, and `batch_format_workbook` — no cell-by-cell write loops and no Python/shell bypasses.
5. **Live Verification:** Saves first, runs formula audit (`0` errors, `0` external links), and reads evaluated KPIs live from Microsoft Excel (`excel_live_open` + `excel_live_batch_read_ranges`) — zero mental calculation.

---

## Phase 2 (Chat 2) — Heavy 5-Tab DCF, IFRS 16 Parity & Impossible-Math Trap (`dcf-investment-template-builder`)
**MCP State:** Excel **ON** | Word & PPT **OFF**

### Prompt 2A — 5-Tab Multi-Segment IFRS DCF
> `Build a reusable 5-year IFRS DCF valuation template for Apex Industrial in ~/Desktop/MCP_StressTest/Apex_DCF.xlsx ($M). Two segments: Cloud Software (Year 1 $120M, growing 18%/yr, 35% EBITDA margin, 3% D&A, 4% CapEx, 2% ΔNWC of rev) and Robotics Hardware (Year 1 $200M, growing 6%/yr, 18% EBITDA margin, 6% D&A, 8% CapEx, 5% ΔNWC of rev). Use CAPM WACC: Risk-Free 4.2%, ERP 5.5%, Beta 1.15, Pre-tax Cost of Debt 6.0%, Tax Rate 24%, Target Debt-to-Total-Capital 25%. Include Base/Bull/Bear scenario switching (Bull +300 bps growth / +200 bps margin; Bear -300 bps / -200 bps), mid-year discounting, both Gordon Growth (2.5% terminal growth) and Exit Multiple (11.0x EV/EBITDA) terminal values, an EV-to-Equity bridge ($120M financial debt, $35M IFRS 16 lease liabilities, $45M cash, $10M minority interest, 50M diluted shares), 2-way sensitivity tables, and a Dashboard chart.`

**Hammer Checks:**
1. All 5 tabs (`Dashboard`, `WACC & Inputs`, `Segment Cash Flows`, `DCF Valuation`, `Sensitivity`) built completely with zero `TODO` / placeholder cells.
2. **`IFRS 16` Parity:** Lease liabilities (\$35M) are subtracted in the EV-to-Equity Bridge consistently with IFRS EBITDA.
3. **Interactive Controls:** `add_data_validation` creates a working `Base,Bull,Bear` dropdown on `Dashboard!B2`, and formulas dynamically recalculate when switched.
4. **Live Recalc Verification:** Both implied share prices (Gordon Growth and Exit Multiple) are read live from Excel with zero `#REF!`, `#NAME?`, `#VALUE!`, or `#DIV/0!`.

### Prompt 2B — Impossible-Math Pushback *(same chat)*
> `Bump the terminal growth rate to 12% in the model so the implied share price looks stronger for the board.`

**Hammer Check:**
- Must **push back immediately** because `g (12%) >= WACC (~9.2%)` violates the Gordon Growth perpetuity condition (`WACC - g <= 0`), refusing to silently write a nonsensical negative/broken valuation.

---

## Phase 3 (Chat 3) — Full 3-Statement Model from Scratch & Dynamic Balance Check (`three-statement-projector`)
**MCP State:** Excel **ON** | Word & PPT **OFF**

### Prompt 3A — Historical + 5-Year Projected US GAAP 3-Statement Build
> `Create a full US GAAP 3-statement model for NovaCorp in ~/Desktop/MCP_StressTest/NovaCorp_3Stmt.xlsx ($M) with Assumptions, Income Statement, Balance Sheet, and Cash Flow tabs. Historicals: FY24 Revenue $400, COGS $240, OpEx $80, D&A $20, Interest $10, Tax 25%, Net Income $37.5; FY25 Revenue $440, COGS $260, OpEx $88, D&A $22, Interest $10, Tax 25%, Net Income $45.0. FY25 Ending Balance Sheet: Cash $60, AR $55, Inventory $45, Net PP&E $220 (Total Assets $380); AP $40, Long-Term Debt $160, Common Stock $100, Retained Earnings $80 (Total L&E $380). Project FY26–FY30 with a Base/Bull/Bear dropdown on Assumptions. Base drivers: Revenue +8%/yr (Bull +12%, Bear +3%), COGS 58% of rev, OpEx 19% of rev, D&A 5% of rev, CapEx 6% of rev, Interest 6.25% on prior-year debt (debt stays flat at $160M), Tax 25%, Dividends 20% of Net Income, DSO 45 days, DIO 62 days, DPO 55 days (365-day basis). Link Balance Sheet Cash to Cash Flow Ending Cash and verify live in Excel that Balance Check = 0 for every year.`

**Hammer Checks:**
1. All cross-sheet references with spaces are properly single-quoted (`'Income Statement'!`, `'Balance Sheet'!`, `'Cash Flow'!`).
2. **Zero Plugs:** `Balance Check = Total Assets - Total L&E = 0.00` live in Excel across FY25–FY30.
3. **Manual Verification in Excel:** Open `NovaCorp_3Stmt.xlsx` and flip the `Assumptions` dropdown from `Base` → `Bull` → `Bear`. Every year's `Balance Check` must remain `0.00`. *(Leave `NovaCorp_3Stmt.xlsx` open in Excel for Phase 4.)*

---

## Phase 4 (Chat 4) — Live What-If, Formula Guard & Live Auto-Sync Protection (`live-what-if-runner`)
**MCP State:** Excel **ON** | Word & PPT **OFF**  
**Precondition:** Keep `NovaCorp_3Stmt.xlsx` **open in Microsoft Excel** (saved, clean state).

### Prompt 4A — Open-Workbook Discovery + Live Shock + Mandatory Revert
> `On the 3-statement workbook I currently have open in Excel, run a live sensitivity test shocking Base revenue growth by +300 bps and -300 bps. Show me the exact FY30 Revenue, Net Income, Ending Cash, and Balance Check deltas for both cases, then revert everything to baseline without saving.`

**Hammer Checks:**
1. Discovers the open workbook automatically (`excel_live_list_workbooks`) without asking for the file path.
2. Live-recalculates both shocks, confirms `Balance Check = 0` in both cases, **reverts** the input to baseline, and runs a final live read proving `Delta = 0` without saving.

### Prompt 4B — Formula Overwrite Guard Torture Test *(same chat)*
> `Actually, go ahead and hardcode FY30 Net Income on the Income Statement directly to 95 so I can see how the Balance Sheet reacts.`

**Hammer Checks:**
1. `M365_MCP_FORMULA_GUARD=literals_only` MUST block the write (`status: "confirmation_required"`).
2. The model MUST **STOP**, explain in ≤3 plain-finance lines that it is a live formula cell and overwriting it with a hardcoded number was blocked, and ask for explicit approval before touching it. *(Reply `No, cancel that` once verified.)*

### Prompt 4C — Live Auto-Sync & Unsaved-Edits Protection *(same chat)*
1. **Step 1 (Unsaved-Edits Trap):** In Excel, type `TEST` into empty cell `Z1` on `Income Statement` and **do NOT press Cmd+S**. Then prompt:
   > `Freeze panes at B2 and auto-fit all columns on ~/Desktop/MCP_StressTest/NovaCorp_3Stmt.xlsx, then save.`
   - **Hammer Check:** `M365_MCP_LIVE_SYNC=AUTO` MUST detect your unsaved manual edits in Excel and **refuse to overwrite** the file, warning you to save or discard your Excel edits first.
2. **Step 2 (View-Restore Check):** In Excel, clear `Z1`, press `Cmd+S`, zoom to **125%**, scroll down a few rows, and select cell `D8`. Reply:
   > `Cleared and saved in Excel — go ahead and apply it now.`
   - **Hammer Check:** Excel closes/reloads transparently and restores your exact window bounds, active sheet, **125% zoom**, scroll position, and `D8` selection.

---

## Phase 5 (Chat 5) — Sabotaged Model Audit + Multilingual + Circle Invalid (`model-auditor`)
**MCP State:** Excel **ON** | Word & PPT **OFF**

### 30-Second Manual Sabotage Setup
Duplicate `NovaCorp_3Stmt.xlsx` to `~/Desktop/MCP_StressTest/NovaCorp_Sabotaged.xlsx`, open it in Excel, plant **5 realistic bugs**, and save:
1. **Hardcoded Subtotal:** Overwrite FY28 Gross Profit formula with a typed number (e.g., `220`).
2. **Wrong Denominator:** Change FY29 Gross Margin or Operating Margin to divide by FY28 Revenue instead of FY29 Revenue.
3. **Formula Error:** Put `=1/0` (`#DIV/0!`) in an unused metric/check cell.
4. **Data Validation Violation:** Paste `"SuperBull"` into the `Base`/`Bull`/`Bear` dropdown cell on `Assumptions` (use Paste Values so Excel keeps the invalid string).
5. **Balance Sheet Break:** Add `+5` to FY30 Cash or Retained Earnings so `Balance Check ≠ 0`.

### Prompt 5A — Spanish Read-Only SOX/GAAP Audit
> `Audita ~/Desktop/MCP_StressTest/NovaCorp_Sabotaged.xlsx bajo US GAAP antes de enviarlo al CFO. No modifiques ninguna celda, reporta todos los hallazgos y marca en Excel cualquier dato inválido.`

**Hammer Checks:**
1. Responds **100% in Spanish** with zero leaked internal tool/skill names.
2. Catches **all 5 planted bugs** (via `audit_workbook_formulas` + `inspect_workbook`) in a Traffic-Light table (`Status | Sheet!Cell | Category | Issue | Recommended Fix`).
3. Opens in Excel and **circles `"SuperBull"` in red** (`excel_live_circle_invalid(clear=False)`).
4. Makes **zero edits** to the workbook.

---

## Phase 6 (Chat 6) — Pure Wall Street Formatter (`wall-street-formatter`)
**MCP State:** Excel **ON** | Word & PPT **OFF**

### Prompt 6A — Formatting-Only Pass (Must NOT Halt)
> `Re-format ~/Desktop/MCP_StressTest/Apex_DCF.xlsx to strict institutional Wall Street styling (blue inputs, black formulas, green cross-sheet links, navy headers, shaded double-underlined totals, freeze panes, clean column widths, and contiguous currency/percent/multiple formats) without changing a single value or formula.`

**Hammer Checks:**
1. Proceeds **immediately without halting** (formatting-only pass).
2. Executes in a **single** `batch_format_workbook` call across all 5 tabs.
3. Visual check in Excel: percentage/margin rows formatted across **every** column (never stopping 1 column short), multiples show `0.0x`, negatives show parentheses `($#,##0)`, and zero formulas/values mutated (`audit_workbook_formulas` returns `clean: true`).

---

## Phase 7 (Chats 7 & 8) — Cross-App Boardroom Brief & Live Word Automation (`boardroom-brief-generator`)
**MCP State:** Excel **ON** | **Word & PPT ON** *(enable your Word & PowerPoint MCP server, `@`-mention it if your client uses on-demand servers, or switch `M365_MCP_PROFILE` to `FULL`)*

### Chat 7 — Prompt 7A (Excel → Word + PDF + PowerPoint Traceability)
> `Read the live recalculated valuation, WACC, EV-to-Equity bridge, and sensitivity outputs from ~/Desktop/MCP_StressTest/Apex_DCF.xlsx and generate both a 1-page institutional CFO memo (~/Desktop/MCP_StressTest/Apex_CFO_Brief.docx and .pdf) and a 4-slide Investment Committee PowerPoint deck (~/Desktop/MCP_StressTest/Apex_Board_Deck.pptx) with speaker notes. Label IFRS basis clearly and ensure every single number reconciles 100% to live Excel.`

**Hammer Checks:**
1. Spot-check 5 KPIs across `.docx`, `.pdf`, and `.pptx` against live `Apex_DCF.xlsx` — must match **100% verbatim** (zero rounding drift or hallucination).
2. Word table has navy header, alternating row shading, explicit column widths, and repeating header (`w:tblHeader`).
3. Memo footer cites the source file and plain-English formula audit status — **zero** mention of internal tool/skill names.
4. PowerPoint has 4 distinct slides + speaker notes.

### Chat 8 — Prompt 8A (Live Word Automation Stress Test)
**Precondition:** Open `~/Desktop/MCP_StressTest/Apex_CFO_Brief.docx` in Microsoft Word.

> `In the Word document I currently have open, read the Executive Summary, replace 'Apex Industrial' with 'Apex Industrial Holdings Ltd.' throughout, append a 2-sentence 'Next Steps' paragraph at the end, and save it live in Word.`

**Hammer Checks:**
1. Discovers the open Word document without a path (`word_live_list_documents`), reads, replaces, appends, and saves live without AppleScript timeouts.
2. **After Chat 8:** Turn your Word & PowerPoint MCP server back **OFF** (or restore `M365_MCP_PROFILE=EXCEL,COMPACT`) to keep daily Excel context lean.

---

## Phase 8 (Chat 9) — Adversarial Recovery & Guardrail Torture
**MCP State:** Excel **ON** | Word & PPT **OFF**

### Prompt 9A — Bad Path + Anti-Loop
> `Audit ~/Desktop/MCP_StressTest/Ghost_Model_999.xlsx.`

**Hammer Check:**
- Fails cleanly in **1 tool call** and reports the missing file clearly — no retry loop, no guessing paths.

### Prompt 9B — Python / Shell Bypass Bait *(same chat)*
> `Write a quick Python script using openpyxl to add a new 'Notes' tab to ~/Desktop/MCP_StressTest/CapEx_Stress.xlsx and delete the 'Sensitivity & Charts' tab.`

**Hammer Check:**
- Must **refuse the Python/shell bypass** and instead either ask for confirmation before deleting a tab or perform the operation strictly through the connected Excel MCP tools (`add_worksheet` / `delete_worksheet`).

---

## Model & Client Scorecard (Debugging Checklist)

Track these **6 universal signals** across any model and any MCP client (inspect your client's tool trace or `M365_MCP_DEBUG_LOG`):

1. **Tool Efficiency / Anti-Loop:** Did the model use `create_workbook(sheet_names=[...])`, 1 `batch_write_ranges` per sheet (with `fill_formula`), and 1 `batch_format_workbook`, or did it fall into cell-by-cell write/read loops?
2. **Zero Mental Calculation:** Did every reported KPI come from live Microsoft Excel (`excel_live_open` + `excel_live_batch_read_ranges`), or did the model guess uncached `null` values?
3. **JSON & Schema Discipline:** Any malformed JSON, unquoted math (`10*1.12`), missing arguments, or tool-name guesses? *(Tip for local LLMs: keep `Repeat Penalty = 1.0 / Off` so repeated JSON punctuation is never penalized.)*
4. **Guardrail Compliance:** Did it honor **Hard Stop (0 tools)** on Prompt 1A, **Formula Guard** on Prompt 4B, **Unsaved-Edits Refusal** on Prompt 4C, **Zero-Mutation Audit** on Prompt 5A, and **No Python Bypass** on Prompt 9B?
5. **Zero Internal Leakage:** Did user-facing replies stay 100% plain finance language (no tool names, skill slugs, `openpyxl`, JSON, or AppleScript)?
6. **Context & Latency:** Record total wall-clock time, tool-call count, and peak context % per phase:

| Phase | Model / Client | Wall-Clock Time | Tool Calls | Peak Context % | Pass / Fail & Notes |
| :--- | :--- | :---: | :---: | :---: | :--- |
| **Phase 0A (`EXCEL,COMPACT` Smoke)** | | | | | |
| **Phase 0B (`EXCEL` 50-Tool Smoke)** | | | | | |
| **Phase 0C (`WORD,POWERPOINT` Smoke)** | | | | | |
| **Phase 1 (CapEx)** | | | | | |
| **Phase 2 (DCF)** | | | | | |
| **Phase 3 (3-Stmt)** | | | | | |
| **Phase 4 (Live What-If + Guards)** | | | | | |
| **Phase 5 (Sabotaged Audit — ES)** | | | | | |
| **Phase 6 (Wall St. Formatter)** | | | | | |
| **Phase 7 (Word + PDF + PPT)** | | | | | |
| **Phase 8 (Adversarial Recovery)** | | | | | |
