# Microsoft 365 for Mac MCP

*Dual-mode (file-level + live AppleScript) MCP server for Microsoft Excel, Word, and PowerPoint — verified on **macOS 27.x** & **Microsoft Office 16.x** (`m365-mac-mcp`).*

![Release](https://img.shields.io/badge/Release-v2.1.2-blue)
![macOS](https://img.shields.io/badge/macOS-27.x%20(Apple%20Silicon)-000000?logo=apple)
![Microsoft Office](https://img.shields.io/badge/Microsoft%20Office%20for%20Mac-16.x-D83B01?logo=microsoftoffice)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![MCP](https://img.shields.io/badge/MCP-SDK%20v1%20%7C%20v2-4B32C3)
![License](https://img.shields.io/badge/License-Non--Commercial%20%2F%20Source--Available-orange)

A Model Context Protocol (MCP) server (`m365-mac-mcp`) that enables AI assistants to create, inspect, format, audit, and automate Microsoft Office documents (**Excel**, **Word**, and **PowerPoint**) on macOS.

This server combines **fast file-level manipulation** (`openpyxl`, `python-docx`, `python-pptx`) with **hardened live macOS Office automation** via AppleScript (`osascript`) for real-time Excel formula recalculation, live Word editing, and seamless view-preserving sync with open workbooks.

> **Author:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)  
> **Licensing:** [Custom Non-Commercial / Source-Available License](LICENSE) — **free for personal & academic use**; **commercial license required** for business/commercial use (`stejimenez@gmail.com`). Not an official Microsoft MCP Server.

---

## Highlights & Architecture

- **Dual-Mode Engine (File-Level + Live App):**
  - **File-level tools** (`openpyxl`, `python-docx`, `python-pptx`) handle bulk reads/writes, multi-sheet batch formatting, charts, Excel tables, data validation, named styles, formula auditing, and rich `.docx` table/paragraph styling without blocking the UI.
  - **Live-app tools** (`excel_live_*`, `word_live_*`) drive Microsoft Excel and Microsoft Word directly via serialized, timeout-guarded AppleScript (`osascript -` + `argv`) for real Excel formula evaluation, live edits, custom views, invalid-data circles, and PDF/RTF/TXT export.
- **Token-Efficient by Design (Zero Loss of Exactitude):**
  - **Compact JSON wire responses** (`separators=(",", ":")`, `structured_output=False`) cut 2D spreadsheet response payloads by **46%–74%** with zero data loss.
  - **Sparse A1 formula maps** (`{"D2": "=(C2/B2)-1", ...}`) in `inspect_workbook`, `excel_live_read_range`, and `excel_live_batch_read_ranges` eliminate duplicated 2D matrices while making formula coordinates explicit.
  - **Deduplicated Word reads (`read_document`):** Omits duplicate `full_text` (opt-in via `include_full_text=True`), duplicate `headings`, default `"style": "Normal"`, and empty metadata/tables — cutting Word read tokens by **~40%–50%** with zero data loss.
  - **Compacted MCP tool schemas** strip Pydantic `"title"`, redundant `"default": null`, and verbose nullable `anyOf` wrappers (**-35%** per-turn schema overhead) while preserving full runtime validation.
  - **Trimmed trailing blanks & omitted empty boilerplate** keep responses tight without hiding non-empty data or truncation flags.
- **Live Excel Auto-Sync (`OFFICE_MCP_LIVE_SYNC=auto`):** When `save_workbook` writes to a file already open in Excel with no unsaved edits, it automatically captures the window bounds, active sheet, zoom, scroll position, and selection, reloads the saved file in Excel, and restores your exact view. If you have unsaved manual edits in Excel, it refuses to overwrite so your work is never lost.
- **Composable Tool Profiles (`OFFICE_MCP_PROFILE`):** Trim the exposed toolset (e.g., `excel,compact`, `excel`, `word`, `word,powerpoint`) to cut schema token overhead by up to 60%+ on local LLMs (LM Studio / Bionic).
- **Hardened for Modern macOS & Office for Mac:**
  - Verified on macOS (Apple Silicon) with **Microsoft Excel 16.115** and **Microsoft Word 16.115.1**.
  - Per-event-loop `asyncio.Lock` serialization around `osascript` calls prevents parallel AppleEvent races.
  - Automatic quoting of spaced sheet names in formulas (e.g., `=Cash Flow!F15` → `='Cash Flow'!F15`) and pre-write validation prevents Excel modal dialogs and stray `xl/externalLinks`.
  - Compatible with both **MCP SDK v2** and **MCP SDK v1**, with strict `stderr`-only logging so `stdio` JSON-RPC streams are never polluted.

---

## Features

### Excel (File-Level & Live App)
- **Workbook & Sheet Management:** Create, open, close, save (`.xlsx`, `.xlsm`, `.csv`, `.pdf`), add/rename/delete/activate worksheets, and one-call `inspect_workbook`.
- **Bulk & Batch Operations:** Read/write cells and ranges (with `include_formulas` / `evaluate_formulas` aliases), `batch_write_ranges`, `batch_format_workbook`, `excel_live_batch_read_ranges`, and `excel_live_batch_write_cells`.
- **Financial-Model Styling & Formatting:**
  - `format_range`, `format_worksheet`, `batch_format_workbook`, and `set_column_widths` (explicit widths or auto-fit with padding).
  - Supports comma-separated multi-ranges (`"B2:F3, B6:F7"`), merged font/alignment/border preservation, per-side borders (`thin`, `medium`, `thick`, `double`, `dashed`, `dotted`, `none`), and `auto_color_financial_model=True` (blue inputs, black same-sheet formulas, green cross-sheet links).
  - 16 built-in named cell style presets (`Normal`, `Good`, `Bad`, `Neutral`, `Input`, `Calculation`, `Check Cell`, `Linked Cell`, `Note`, `Output`, `Warning Text`, `Title`, `Heading 1`–`Heading 4`, `Total`) via `list_cell_styles` / `apply_cell_style`.
- **Tables, Charts & Data Validation:**
  - Native Excel Tables (`add_excel_table`, `update_excel_table_style`, `excel_live_update_table_style`) with `TableStyleLight*` / `Medium*` / `Dark*` and banded row/column toggles.
  - Charts (`create_chart`: bar, column, line, pie).
  - Data validation rules (`add_data_validation`, `list_data_validations`, `clear_data_validations`) plus live invalid-data circles (`excel_live_circle_invalid`).
- **Formula Auditing & Live Recalculation:**
  - One-call `audit_workbook_formulas` checks for `#REF!`, `#NAME?`, `#VALUE!`, `#DIV/0!`, `#N/A`, external links, and broken sheet references.
  - Live Excel reads use bulk single-AppleEvent matrix extraction (10–50× faster and reliable on vertical `N×1` ranges) with live recalculation (`recalculate=True`).
  - Live Custom Views (`excel_live_custom_views`) and capability inspection (`excel_view_capabilities`).

### Word (File-Level & Live App)
- **Document Operations:** Create, open (including legacy `.doc` conversion), close, read paragraphs/tables, set document properties, search & replace, and export to `.docx`, `.pdf`, `.rtf`, or `.txt`.
- **Rich `.docx` Styling:**
  - Full formatting on `add_heading`, `add_paragraph`, `insert_paragraph`, `add_list`, `format_paragraph` (fonts, colors, alignment, line spacing, `keep_with_next`, indents, shading/background, bottom borders).
  - Full table styling on `add_table`, `format_table`, and `update_table_cell` (column widths in inches with `autofit=False`, header/body styles, alternating row shading, XML borders, cell padding in pt, `repeat_header` / `w:tblHeader`, table & per-column alignment).
- **Live Word Automation:** List open documents (with multi-tier cascade hardened for Word 16.115.1), open, read, write/append, search & replace, save/export, and close directly in Microsoft Word.

### PowerPoint
- Create presentations with custom titles and themes
- Add slides with standard layouts
- Insert text, images, and speaker notes
- Save presentations in multiple formats

### Finance Skills *(Work in Progress)*
Ready-to-use, pure-Markdown (**3–4 KB** each) **IFRS / FASB / SOX-compliant Finance Skills** live in [`skills/`](skills/) so non-technical finance professionals can run complex workflows in plain English:
- **Wall Street Formatter** (`wall-street-formatter`) · **Model Auditor** (`model-auditor`) · **Live What-If Runner** (`live-what-if-runner`)
- **3-Statement Projector** (`three-statement-projector`) · **CapEx Project Analyzer** (`capex-project-analyzer`) · **DCF & Valuation Template Builder** (`dcf-investment-template-builder`) · **Boardroom Brief Generator** (`boardroom-brief-generator`)

> **Full Installation & Plain-English Usage Guide (WIP):** See [**`skills/README.md`**](skills/README.md).

---

## Prerequisites

### System Requirements
- **macOS** 10.15 or later (recommended: Apple Silicon macOS 27.x)
- **Python 3.10+** (recommended: Python 3.12 in a local `.venv`)
- **Microsoft 365 / Office for Mac** (Excel, Word, and/or PowerPoint)

### Recommended & Tested Reference Setup

This server was developed, benchmarked, and end-to-end verified on the following 100% local Apple Silicon setup:

| Component | Verified Configuration |
| --- | --- |
| **Hardware** | **Mac mini (Apple M4 Pro, `Mac16,11`)** — 14-core CPU (10P + 4E), **64 GB** Unified Memory |
| **Operating System** | **macOS 27.2.1** (Build `26B131`, Apple Silicon) |
| **Microsoft 365 for Mac** | Excel `16.115`, Word `16.115.1`, PowerPoint `16.115.1` |
| **MCP Client & Runtime** | **Bionic `1.2.1+10`** + **LM Studio** (local inference) |
| **Local Inference Model** | **`Qwen3.6 35B A3B — Q8_0`** (Context `65536`, GPU Offload `40`, CPU Threads `6`, Eval Batch `2048`, Phys Batch `512`, Flash Attention `ON`, Temp `0.3`, Reasoning Budget `2048`, Preserve Thinking `OFF`, Unified + GPU KV Cache `ON`, Top-K `20`, Top-P `0.95`, Min-P `0.05`, Repeat Penalty `1.1`, Experts `8`) — see [`Sample Model Config/`](Sample%20Model%20Config/) |
| **Python Environment** | Python `3.12.16` (`.venv`), `mcp` `2.3.0`, `openpyxl` `3.1.5`, `python-docx` `1.2.0` |
| **Recommended MCP Env** | `OFFICE_MCP_PROFILE=EXCEL,COMPACT` (or `EXCEL`), `OFFICE_MCP_LIVE_SYNC=AUTO`, `OFFICE_MCP_FORMULA_GUARD=LITERALS_ONLY`, `OFFICE365_MCP_APPLESCRIPT_TIMEOUT=30` |

### Hardware Tier & MCP Profile Quick Summary

| Apple Silicon Tier | Best Local Model (LM Studio / MLX) | Recommended Profile | Safe Context | Expected Speed & Risk |
| --- | --- | --- | --- | --- |
| **16 GB M4** | `Qwen3.5-9B` *(MLX 4/6-bit)* or `Qwen3-14B` *(4-bit, tight ctx)* | `EXCEL,COMPACT` *(25 tools, ~2.4K tok)* | `12K–16K` *(Q8 KV)* | ~20–32 tok/s; use 2-phase chunking on big builds |
| **32 GB M4 / M4 Pro** | `Qwen3.6 35B A3B` *(MLX/GGUF 4-bit)* or `gpt-oss-20b` | `EXCEL,COMPACT` or `EXCEL` | `32K` *(Q8 KV)* | ~20–30+ tok/s; all 7 skills in 1 shot, low risk |
| **64 GB M4 Pro / Max** | `Qwen3.6 35B A3B — Q8_0` *(verified reference)* | `EXCEL,COMPACT` or `FULL` *(82 tools)* | `65K–128K` | ~18–26 tok/s (Pro) / 35–50+ (Max); very low risk |

> **Full Hardware, Local Model & `OFFICE_MCP_PROFILE` Decision Guide:** See [**`docs/HARDWARE_AND_PROFILE_GUIDE.md`**](docs/HARDWARE_AND_PROFILE_GUIDE.md).

### Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/stejimenez/m365-mac-mcp.git
   cd m365-mac-mcp
   ```

2. **Create a virtual environment and install dependencies**
   ```bash
   python3 -m venv .venv
   .venv/bin/pip install -r requirements.txt
   ```

   `requirements.txt` installs:
   - `mcp>=1.2.0`
   - `python-docx>=1.1.0`
   - `openpyxl>=3.1.2`
   - `python-pptx>=1.0.0`
   - `aiofiles>=23.2.1`, `typing-extensions>=4.12.0`
   - `pytest>=8.0.0`, `pytest-asyncio>=0.23.0`

3. **Grant macOS Automation permissions**

   Live Office tools use AppleScript to control Office apps:
   1. Open **System Settings** → **Privacy & Security** → **Automation**
   2. Locate your host application (Terminal, iTerm2, VS Code, Bionic, Claude Desktop, etc.)
   3. Enable:
      - **Microsoft Excel**
      - **Microsoft Word**
      - **Microsoft PowerPoint**
   4. Click **OK** if macOS prompts on first run.

---

## Configuration

### Environment Variables

> All environment variables are **optional** — if omitted or left blank in Bionic, Claude Desktop, or Cline, the server automatically applies the documented defaults below.

| Variable | Default | Description |
| --- | --- | --- |
| `OFFICE_MCP_PROFILE` | `full` | Comma-separated, case-insensitive tool profile filter (see table below). |
| `OFFICE_MCP_LIVE_SYNC` | `auto` | `auto` (reloads open saved Excel workbooks with view restore on `save_workbook`) or `off`. |
| `OFFICE_MCP_FORMULA_GUARD` | `literals_only` | Protects existing Excel formulas from accidental overwrites: `literals_only` (default), `strict`, or `off`. |
| `OFFICE365_MCP_LOG_LEVEL` | `INFO` | Logging verbosity (`DEBUG`, `INFO`, `WARNING`, `ERROR`). All logs go to `stderr`. |
| `OFFICE365_MCP_APPLESCRIPT_TIMEOUT` | `30.0` | Timeout in seconds for each `osascript` call. |
| `OFFICE365_MCP_ENABLE_APPLESCRIPT` | `true` | Enable or disable AppleScript bridge (`true`/`false`). |
| `OFFICE365_MCP_TEMP_DIR` | `/tmp/office365_mcp` | Temporary directory for intermediate/unnamed files. |

### Tool Profiles (`OFFICE_MCP_PROFILE`)

To keep context usage low and tool selection fast on local models (such as LM Studio / Bionic), set `OFFICE_MCP_PROFILE` to expose only the apps you need:

| `OFFICE_MCP_PROFILE` | Tools | Est. Schema Tokens | Best For |
| --- | --- | --- | --- |
| `excel,compact` | 25 | ~2.4K | **Recommended** for daily Excel modeling, live recalc & all 6 Excel Finance Skills on local models |
| `excel` | 50 | ~4.8K | Full Excel toolset (all single-op + batch tools, views, cell styles) |
| `word` | 26 | ~2.2K | Word-only document creation, styling, and live automation |
| `powerpoint` | 8 | ~0.5K | PowerPoint-only slide decks |
| `word,powerpoint` | 33 | ~2.7K | Secondary MCP entry for Word + PowerPoint |
| `compact` | 38 | ~3.5K | Core Excel + Word tools combined |
| `excel,word` | 75 | ~7.0K | Full Excel + Word without PowerPoint |
| `full` *(default)* | 82 | ~7.4K | All Excel, Word, and PowerPoint tools |

- Values are **case-insensitive** (`EXCEL,COMPACT`, `Excel,Compact`, `excel,compact`, or `XLS,Compact` are equivalent) and accept any comma-separated combination of `excel` (`xls`/`xlsx`), `word` (`doc`/`docx`), `powerpoint` (`ppt`/`pptx`), and `compact`.
- Unknown values log a warning to `stderr` and fall back safely to `full`.
- Call `check_office_status` (or read `office365://status`) at any time to inspect the active profile, registered tool count, `live_sync_mode`, and `formula_guard_mode`.

### Live Excel Auto-Sync (`OFFICE_MCP_LIVE_SYNC`)

Controls how `save_workbook` behaves when the target `.xlsx`/`.xlsm` file is already open in Microsoft Excel (case-insensitive: `auto` or `off`):

- **`auto` *(default)*:** If the file is open in Excel with **no unsaved changes**, captures the window position/size, active sheet, zoom, scroll offset, and selection → closes Excel → writes via `openpyxl` → reopens in Excel and restores the exact view. If the file has **unsaved manual edits** in Excel, refuses to overwrite and returns a clear warning so your edits are never lost.
- **`off`:** Never checks or touches live Excel during `save_workbook`.

> **Note on `excel_live_close`:** `OFFICE_MCP_LIVE_SYNC=auto` applies to file-level `save_workbook`. If you explicitly call `excel_live_close`, Excel will close the workbook and leave it closed—simply omit `excel_live_close` if you want the workbook to stay open in Excel.

### Formula Overwrite Guard (`OFFICE_MCP_FORMULA_GUARD`)

Prevents AI models from silently destroying live Excel formulas across all 6 write tools (`write_cell`, `write_range`, `batch_write_ranges`, `excel_live_write_cell`, `excel_live_write_range`, `excel_live_batch_write_cells`). Values are case-insensitive:

- **`literals_only` *(default)*:** Atomically blocks **Formula → Hardcoded Literal / Blank (`None` / `""`)** overwrites before modifying any cell and returns `status: "confirmation_required"` with the blocked cells so the model stops and asks for your approval (retrying with `allow_formula_overwrite=true` only if approved). Legitimate **Formula → New Formula** updates are allowed and reported via `"formulas_updated": N`.
- **`strict`:** Blocks **any** overwrite of an existing formula cell (including Formula → Formula) unless `allow_formula_overwrite=true` — ideal for finalized valuation/board models where only input cells should be edited.
- **`off`:** Disables the pre-write formula guard.

---

### MCP Client Configuration

#### For Bionic / LM Studio (Recommended Local-Model Setup)

Register two entries pointing to the same server `.venv`:
1. **`Microsoft 365 for Mac MCP` (Excel)** — always enabled:
   - **Command:** `/path/to/m365-mac-mcp/.venv/bin/python`
   - **Args:** `["/path/to/m365-mac-mcp/src/office365_mcp_server.py"]`
   - **Env:** `OFFICE_MCP_PROFILE=excel` (or `excel,compact`), `OFFICE_MCP_LIVE_SYNC=auto`
   - **Description:** `Microsoft 365 for Mac MCP (Excel/Office) for reading, editing, formatting, auditing, and live recalculating .xlsx/.xlsm workbooks.`
2. **`M365 Word & PPT MCP`** — set to **Hide from agent unless mentioned** so it uses zero context until you `@`-mention it:
   - **Env:** `OFFICE_MCP_PROFILE=word,powerpoint`

**Recommended LM Studio / Bionic settings for `Qwen3.6 35B A3B — Q8_0` (full screenshots & system prompt in [`Sample Model Config/`](Sample%20Model%20Config/)):**
- **Context & Performance:** `Context Length = 65536`, `GPU Offload = 40`, `CPU Thread Pool Size = 6`, `Evaluation Batch Size = 2048`, `Physical Batch Size = 512`, `Max Concurrent Predictions = 1`, `Flash Attention = ON`
- **Generation & Reasoning:** `Temperature = 0.3`, `Context Overflow = Stop at Limit`, `Reasoning Budget = 2048` (`"I have to answer now."`), `Preserve Thinking = OFF`
- **Memory:** `Offload KV Cache to GPU = ON`, `Unified KV Cache = ON`, `Context Checkpoints = 32`, `Keep Model in Memory = ON`, `Try mmap() = ON`, `Speculative Decoding = Off`
- **Sampling:** `Top K = 20`, `Top P = 0.95`, `Min P = 0.05`, `Repeat Penalty = 1.1`, `Number of Experts = 8`
- **System Prompt:** Copy [`Sample Model Config/SYSTEM_PROMPT.md`](Sample%20Model%20Config/SYSTEM_PROMPT.md) (`UNIVERSAL STERLING ENGINE V3.4 — FINANCE & M365 SPECIALIST`) into Bionic/LM Studio's System Prompt so the agent stays concise, avoids tool loops, blocks Bionic's built-in Office/shell bypasses, respects `OFFICE_MCP_FORMULA_GUARD`, and uses your custom Finance `skills/` cleanly.
  > **Disclosure:** Applying `Sample Model Config/SYSTEM_PROMPT.md` **highly customizes the model for financial modeling, corporate valuation, Excel/VBA, and Microsoft 365 work only**. If you also use Bionic / LM Studio for general chat or non-finance coding, apply this prompt inside a dedicated Finance Project rather than as your global default.

#### For Claude Desktop
Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "m365-mac-mcp": {
      "command": "/path/to/m365-mac-mcp/.venv/bin/python",
      "args": ["/path/to/m365-mac-mcp/src/office365_mcp_server.py"],
      "env": {
        "OFFICE_MCP_PROFILE": "full",
        "OFFICE_MCP_LIVE_SYNC": "auto"
      }
    }
  }
}
```

#### For Cline (VS Code)
Edit your Cline MCP settings:

```json
{
  "mcpServers": {
    "m365-mac-mcp": {
      "command": "/path/to/m365-mac-mcp/.venv/bin/python",
      "args": ["/path/to/m365-mac-mcp/src/office365_mcp_server.py"],
      "env": {
        "OFFICE_MCP_PROFILE": "excel,word",
        "OFFICE_MCP_LIVE_SYNC": "auto"
      }
    }
  }
}
```

---

## Tool Reference (82 Tools)

### Excel File-Level Tools (30)
| Category | Tools |
| --- | --- |
| **Lifecycle** | `create_workbook`, `open_workbook`, `close_workbook`, `save_workbook`, `list_active_workbooks` |
| **Inspection & Audit** | `inspect_workbook`, `audit_workbook_formulas` *(both accept `workbook_id` or `file_path` directly)* |
| **Worksheets** | `list_worksheets`, `add_worksheet`, `rename_worksheet`, `delete_worksheet`, `set_active_worksheet` |
| **Read & Write** | `read_cell`, `read_range`, `write_cell`, `write_range`, `add_formula`, `batch_write_ranges` |
| **Formatting & Styles** | `format_range`, `format_worksheet`, `batch_format_workbook`, `set_column_widths`, `list_cell_styles`, `apply_cell_style` |
| **Tables & Charts** | `add_excel_table`, `update_excel_table_style`, `create_chart` |
| **Data Validation** | `add_data_validation`, `list_data_validations`, `clear_data_validations` |

### Excel Live-App Tools (18)
| Category | Tools |
| --- | --- |
| **Lifecycle** | `excel_live_list_workbooks`, `excel_live_create`, `excel_live_open`, `excel_live_save`, `excel_live_close` |
| **Worksheets** | `excel_live_list_worksheets`, `excel_live_add_worksheet`, `excel_live_rename_worksheet`, `excel_live_activate_worksheet` |
| **Read & Write** | `excel_live_read_range`, `excel_live_write_cell`, `excel_live_write_range`, `excel_live_batch_read_ranges`, `excel_live_batch_write_cells` |
| **Views & Inspection** | `excel_live_circle_invalid`, `excel_live_custom_views`, `excel_live_update_table_style`, `excel_view_capabilities` |

#### Verified Excel for Mac `16.115` Scripting Capabilities
| Feature | Support | Tool / Notes |
| --- | --- | --- |
| Live recalculation & bulk range reads | Supported | `excel_live_read_range`, `excel_live_batch_read_ranges`, `excel_live_write_cell`, `excel_live_batch_write_cells` |
| Circle / Clear Invalid Data | Supported | `excel_live_circle_invalid` (`circle invalid` / `clear circles`) |
| Custom Views (list & show) | Supported | `excel_live_custom_views` (`list` / `show`) |
| Native Table Styles | Supported | `excel_live_update_table_style` & `update_excel_table_style` |
| Focus Cell / Focus Cell Color / Auto-Highlight | Manual UI only | Not exposed in `Excel.sdef`; reported by `excel_view_capabilities` |

### Word File-Level Tools (18)
| Category | Tools |
| --- | --- |
| **Lifecycle & Metadata** | `create_document`, `open_document`, `close_document`, `save_document`, `list_active_documents`, `set_document_properties` |
| **Reading** | `read_document` *(deduplicated paragraphs; opt-in `include_full_text=True`)*, `read_tables` |
| **Content & Editing** | `add_heading`, `add_paragraph`, `insert_paragraph`, `delete_paragraph`, `replace_document_text`, `add_list` |
| **Styling & Tables** | `format_paragraph`, `add_table`, `format_table`, `update_table_cell` |

### Word Live-App Tools (8)
`word_live_list_documents`, `word_live_create`, `word_live_open`, `word_live_read`, `word_live_write`, `word_live_replace_text`, `word_live_save`, `word_live_close`

### PowerPoint Tools (7)
`create_presentation`, `add_slide`, `add_text_to_slide`, `add_image_to_slide`, `add_speaker_notes`, `save_presentation`, `list_active_presentations`

### Utility & MCP Resources (1 Tool + 1 Resource)
- `check_office_status` — Reports installed Office apps, active profile, registered tool count, and `live_sync_mode`.
- `office365://status` — Read-only MCP resource returning the same server status payload.

---

## Example Workflows

### 1. Audit & Format an Excel Financial Model (File-Level)
```python
# 1. Inspect and audit a workbook directly by file path
summary = await inspect_workbook(file_path="~/Desktop/Model.xlsx")
audit = await audit_workbook_formulas(file_path="~/Desktop/Model.xlsx")

# 2. Open, batch-format across sheets, add validation & table, and save
wb = await open_workbook(file_path="~/Desktop/Model.xlsx")
await batch_format_workbook(
    workbook_id=wb["workbook_id"],
    operations=[
        {
            "sheet_name": "Income Statement",
            "freeze_panes": "B2",
            "auto_fit_columns": True,
            "auto_color_financial_model": True,
            "styles": [
                {"range": "A1:G1", "bold": True, "bg_color": "1F4E78", "font_color": "FFFFFF"},
                {"range": "B2:F15", "number_format": "#,##0"}
            ]
        }
    ]
)
await add_excel_table(
    workbook_id=wb["workbook_id"],
    sheet_name="Income Statement",
    table_range="A1:G16",
    table_name="IS_Table",
    table_style="TableStyleLight9"
)
await add_data_validation(
    workbook_id=wb["workbook_id"],
    sheet_name="Income Statement",
    cell_range="I1",
    validation_type="list",
    formula1=["Base", "Bull", "Bear"]
)
await save_workbook(workbook_id=wb["workbook_id"])
```

### 2. Live Excel Recalculation Check
```python
await excel_live_open(file_path="~/Desktop/Model.xlsx", activate=True)
baseline = await excel_live_batch_read_ranges(
    workbook_name="Model.xlsx",
    ranges=[
        {"sheet_name": "Income Statement", "cell_range": "F2:F15"},
        {"sheet_name": "Balance Sheet", "cell_range": "F12:F27"},
    ]
)
await excel_live_write_cell(
    workbook_name="Model.xlsx",
    sheet_name="Income Statement",
    cell="F2",
    value=450000,
    recalculate=True
)
updated = await excel_live_read_range(
    workbook_name="Model.xlsx",
    sheet_name="Income Statement",
    cell_range="F15"
)
await excel_live_save(workbook_name="Model.xlsx")
```

### 3. Create a Styled Word Report
```python
doc = await create_document(title="Executive Summary")
await add_heading(
    document_id=doc["document_id"],
    text="Q4 Financial Overview",
    level=1,
    formatting={"font_name": "Calibri", "font_color": "1F4E78"}
)
await add_table(
    document_id=doc["document_id"],
    rows=3,
    columns=3,
    data=[["Metric", "FY24", "FY25"], ["Revenue", "$391B", "$416B"], ["Net Income", "$94B", "$112B"]],
    column_widths=[2.5, 1.5, 1.5],
    header_formatting={"bold": True, "bg_color": "1F4E78", "font_color": "FFFFFF"},
    alternating_row_color="F2F6FA",
    repeat_header=True
)
await save_document(document_id=doc["document_id"], file_path="~/Desktop/Report.docx")
```

---

## Testing

Run the unit and integration test suite with the project virtual environment:

```bash
.venv/bin/pytest tests/test_basic.py tests/test_word_excel_rw.py -v
```

This runs **29 tests** covering:
- Excel file-level read/write, batch formatting, named styles, tables, data validation, formula sheet-reference auto-quoting & auditing, `OFFICE_MCP_FORMULA_GUARD` (`literals_only`/`strict`/`off`), and `OFFICE_MCP_LIVE_SYNC` view-restore / unsaved-edits protection.
- Word file-level creation, rich paragraph/table styling, search & replace, and `.docx` XML attributes (`w:tblHeader`, borders, shading, widths).
- Composable `OFFICE_MCP_PROFILE` filtering, compact JSON responses, stripped schema noise, sparse A1 formula maps, trimmed blank cells, and automated validation of all 7 Finance `skills/*/SKILL.md` playbooks.
- Hardened AppleScript generation and parsing for Excel `16.115` and Word `16.115.1`.

---

## Troubleshooting

### Permission Issues (`-1743` / "Not authorized to send Apple events")
1. Open **System Settings** → **Privacy & Security** → **Automation** and enable Excel, Word, and PowerPoint for your MCP host app.
2. Open each Office app manually at least once to finish initial licensing/first-run dialogs.
3. Restart your terminal or MCP client after granting permissions.

### Live Excel Calls Time Out or Hang
- Check if Excel has an open modal dialog (e.g., **Save As**, **Grant Access**, or an alert) or if a cell is actively in **Edit Mode** (cursor blinking inside the formula bar). Press `Esc` in Excel and retry.
- `create_workbook(..., open_in_excel=True)` and `create_document(..., open_in_word=True)` automatically skip opening unnamed `/tmp` UUID files to avoid macOS sandbox **Grant Access** prompts; save the file to a normal folder (like `~/Desktop` or your workspace) before opening in the live app.

### Formula Cells Show `null` After `save_workbook`
- File-level `openpyxl` writes formulas to disk without running Excel's calculation engine, so cached evaluated values are empty until Excel recalculates them. Use `excel_live_open` + `excel_live_read_range` (or `OFFICE_MCP_LIVE_SYNC=auto` with the file open in Excel) to evaluate and read live calculated values.

### Import / Environment Errors
- Ensure your MCP client config points directly to `.venv/bin/python` rather than system `python3`.
- Verify dependencies with `.venv/bin/pip install -r requirements.txt`.

---

## Project Structure

```text
m365-mac-mcp/
├── src/
│   ├── office365_mcp_server.py       # Main MCP server, tool registration & profiles
│   ├── controllers/
│   │   ├── excel_controller.py       # File-level Excel (openpyxl) + live auto-sync
│   │   ├── word_controller.py        # File-level Word (python-docx) + rich styling
│   │   └── powerpoint_controller.py  # File-level PowerPoint (python-pptx)
│   ├── integrations/
│   │   └── applescript_bridge.py     # Serialized, hardened macOS AppleScript bridge
│   └── utils/
│       ├── config.py                 # Env & JSON configuration
│       ├── logger.py                 # stderr-only structured logging
│       └── validators.py             # Input & path validation
├── tests/
│   ├── test_basic.py
│   └── test_word_excel_rw.py
├── docs/
│   ├── EXECUTIVE_COMPARISON_MATRIX.md          # Standalone 16GB/32GB/64GB M4 comparison matrix
│   └── HARDWARE_AND_PROFILE_GUIDE.md           # 16GB/32GB/64GB M4 model & OFFICE_MCP_PROFILE guide
├── Sample Model Config/
│   ├── SYSTEM_PROMPT.md                        # Recommended Bionic/LM Studio system prompt
│   ├── bionic_mcp_config_part1.jpg             # Bionic MCP server setup (Connection, Command, Args)
│   ├── bionic_mcp_config_part2.jpg             # Bionic MCP server setup (Env vars, Working dir)
│   ├── qwen3.6_35b_a3b_q8_0_config_part1.png   # LM Studio model config (Prompt, Context, Reasoning)
│   └── qwen3.6_35b_a3b_q8_0_config_part2.png   # LM Studio model config (Memory, Sampling, Advanced)
├── skills/                                     # Ready-to-use Finance Skills (SKILL.md) — see skills/README.md
│   ├── README.md
│   ├── wall-street-formatter/SKILL.md
│   ├── model-auditor/SKILL.md
│   ├── live-what-if-runner/SKILL.md
│   ├── three-statement-projector/SKILL.md
│   ├── capex-project-analyzer/SKILL.md
│   ├── dcf-investment-template-builder/SKILL.md
│   └── boardroom-brief-generator/SKILL.md
├── requirements.txt
├── LICENSE
├── NOTICE
└── README.md
```

## Author & License

- **Author:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)
- **Non-Commercial / Personal & Academic Use:** Free under the **Custom Non-Commercial / Source-Available License** — see [LICENSE](LICENSE).
- **Commercial Use:** Requires a separate commercial license — contact **Sterling Jimenez, Ph.D.** at `stejimenez@gmail.com`.
- **Upstream Attribution:** Preserved MIT attribution for the original `macos-office365-mcp-server` foundation is in [NOTICE](NOTICE).

## Acknowledgments

- Built with the Model Context Protocol (MCP)
- Uses `openpyxl`, `python-docx`, and `python-pptx` for file-level Office manipulation
- Uses macOS AppleScript (`osascript`) for native Microsoft Excel and Word automation
- Based on the original `macos-office365-mcp-server` proof-of-concept by Adam Bluhm (https://www.linkedin.com/in/bluhmadam/)
