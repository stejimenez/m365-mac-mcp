# Finance Skills Guide — Microsoft 365 for Mac MCP (`m365-mac-mcp`)

> **Status: Work in Progress (Active Development)**  
> These Finance Skills are actively evolving as we test new workflows on local Apple Silicon models and refine accounting/modeling edge cases. Feedback, real-world test cases, and suggestions are welcome (`stejimenez@gmail.com`).

**Author:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)

This folder contains **9 ready-to-use**, pure-Markdown (**2.5–4.0 KB** each) **Finance Skills (`SKILL.md`)** supporting **natural-language prompts in English, French, and Spanish (EN / FR / ES)** so **non-technical finance professionals** can run institutional Excel, Word, and PowerPoint workflows—powered by **`m365-mac-mcp`**.

Every skill follows a strict **5-block low-token blueprint** (*Trigger & Scope → Pre-Flight / HALT → Exact Tool Sequence → IFRS / FASB / SOX 404 Guardrails → Verification Gate*), keeping local LLMs (such as **`Qwen3.6 35B A3B — Q8_0`** in Bionic / LM Studio) fast, auditable, and hallucination-free.

---

## 1. Prerequisites & Recommended MCP Environment

Before using these skills, make sure **`m365-mac-mcp`** is installed and connected in your MCP client (see the root [`README.md`](../README.md)).

Recommended MCP environment variables *(all optional — if omitted or left blank in Bionic, the server automatically applies its safe defaults: `full`, `auto`, `literals_only`; legacy `OFFICE_MCP_*` names remain supported)*:
- `M365_MCP_PROFILE`:
  - **`EXCEL,COMPACT`** *(25 tools, ~2.4K schema tokens — recommended)* or **`EXCEL`** *(50 tools, ~4.8K)* — covers **all 8 Excel skills** (`financial-statement-extractor`, `historical-3stmt-rebuilder`, `wall-street-formatter`, `model-auditor`, `live-what-if-runner`, `three-statement-projector`, `capex-project-analyzer`, `dcf-investment-template-builder`), including `create_chart`, `delete_worksheet`, and `excel_live_list_workbooks`.
  - **`FULL`** *(82 tools, ~7.4K schema tokens)* or **`EXCEL,WORD`** — required for **`boardroom-brief-generator`** (Excel → Word/PDF/PowerPoint).
- `M365_MCP_LIVE_SYNC=AUTO` — automatically reloads your open Excel window and restores your exact sheet, zoom, and selection after `save_workbook` (and refuses to overwrite if you have unsaved edits).
- `M365_MCP_FORMULA_GUARD=LITERALS_ONLY` *(default)* — blocks any accidental overwrite of an existing formula with a hardcoded number or blank (`STRICT` blocks even Formula → Formula edits on finalized models).

---

## 2. How to Install & Load the Skills

### Option A: Bionic / LM Studio (Recommended Local Apple Silicon Setup)
1. Open **Bionic → Settings → Skills**.
2. Click **Add / Import Skill** and select the `skills/` folder (or individual `skills/<skill-name>/SKILL.md` files).
3. **Recommended Finance System Prompt (`V3.4`):**  
   Copy [`Sample Model Config/SYSTEM_PROMPT.md`](../Sample%20Model%20Config/SYSTEM_PROMPT.md) (`UNIVERSAL STERLING ENGINE V3.4 — FINANCE & M365 SPECIALIST`) into your Bionic **Finance Project** System Prompt. It blocks Bionic's built-in Office viewer, shell/Python bypasses, and `.internal/skills/.../SKILL.md` while letting your custom Finance skills load cleanly, enforces `M365_MCP_FORMULA_GUARD`, and prevents mental calculation of uncached `openpyxl` formulas.
   > **Disclosure:** Applying `Sample Model Config/SYSTEM_PROMPT.md` **highly customizes the model for financial modeling, corporate valuation, Excel/VBA, and Microsoft 365 work only**. If you also use Bionic / LM Studio for general-purpose chat or non-finance software development, apply it to a dedicated Finance Project instead of your global default prompt.
4. Toggle your `m365-mac-mcp` MCP connection **OFF → ON** if you changed any environment variables.

### Option B: Claude Desktop / Cline (VS Code)
- **Claude Desktop Projects:** Upload the relevant `SKILL.md` file(s) into your Claude Project's **Project Knowledge** (or reference them directly when working in a local workspace).
- **Cline / VS Code / Antigravity:** Point your workspace skills directory to `skills/`—each subfolder contains a standard YAML-frontmatter `SKILL.md` (`name`, `description`) that is auto-discovered.

---

## 3. Skill Catalog & Copy-Paste Starter Prompts

| Skill | Profile Needed | Accounting / SOX Focus | File |
| --- | --- | --- | --- |
| **1. Financial Statement Extractor** *(EN/FR/ES)* | `EXCEL,COMPACT` | Read-only discovery, verbatim accounts, IFRS/FASB blueprint | [`financial-statement-extractor/SKILL.md`](financial-statement-extractor/SKILL.md) |
| **2. Historical 3-Stmt Rebuilder** *(EN/FR/ES)* | `EXCEL,COMPACT` | New formula model, verbatim accounts, live source reconciliation + Fresh-Chat Handoff | [`historical-3stmt-rebuilder/SKILL.md`](historical-3stmt-rebuilder/SKILL.md) |
| **3. Wall Street Formatter** | `EXCEL,COMPACT` | `IAS 1` / `ASC 210` presentation, SOX 404 zero-mutation | [`wall-street-formatter/SKILL.md`](wall-street-formatter/SKILL.md) |
| **4. Model Auditor** | `EXCEL,COMPACT` | `IAS 1`/`7`, `IFRS 16`/`ASC 842`, SOX 404 read-only review | [`model-auditor/SKILL.md`](model-auditor/SKILL.md) |
| **5. Live What-If Runner** | `EXCEL,COMPACT` | Live Excel recalc, Balance Check identity, SOX revert trail | [`live-what-if-runner/SKILL.md`](live-what-if-runner/SKILL.md) |
| **6. 3-Statement Projector** | `EXCEL,COMPACT` | `IAS 7`/`ASC 230`, `IAS 12`/`ASC 740`, `IAS 16`/`ASC 360` + Fresh-Chat Handoff | [`three-statement-projector/SKILL.md`](three-statement-projector/SKILL.md) |
| **7. CapEx Project Analyzer** | `EXCEL,COMPACT` | `IAS 16`/`ASC 360` capitalization, after-tax salvage, Year-0 NPV rule | [`capex-project-analyzer/SKILL.md`](capex-project-analyzer/SKILL.md) |
| **8. DCF & Valuation Template** | `EXCEL,COMPACT` | `IFRS 16`/`ASC 842` EV Bridge parity, CAPM WACC, Dual TV | [`dcf-investment-template-builder/SKILL.md`](dcf-investment-template-builder/SKILL.md) |
| **9. Boardroom Brief Generator** | `EXCEL,WORD` or `FULL` | IFRS/GAAP disclosure labeling, SOX live-cell traceability | [`boardroom-brief-generator/SKILL.md`](boardroom-brief-generator/SKILL.md) |

---

### Recommended Consecutive Pipeline (Raw Company File → Boardroom)
1. **Same Chat — Step 1 → Step 2:** Run **`financial-statement-extractor`** on a raw company Excel file → approve the Account Mapping Blueprint → reply *"Build it"* / *"Construis-le"* / *"Constrúyelo"* to run **`historical-3stmt-rebuilder`** (reuses the blueprint without re-reading the raw file, saves `<Company>_Clean_Model.xlsx`, and proves live that rebuilt totals match source totals).
2. **Fresh-Chat Handoff (Built Into Builder Skills):** After saving a model, if Bionic's bottom-right **Context** meter is above **~50%** of your limit (`>8K` on 16 GB, `>16K` on 32 GB, `>32K` on 64 GB), start a **New Chat** pointing to the saved file before running **`wall-street-formatter`**, **`three-statement-projector`**, **`dcf-investment-template-builder`**, or **`boardroom-brief-generator`**.

### New Skill: `financial-statement-extractor` (EN / FR / ES)
- **What it does:** Read-only tab discovery (`list_worksheets` → targeted `read_range` on IS/BS/CF only) across any number of historical years, keeping every company-specific account name verbatim in its original language.
- **Starter Prompts:**
  - *(EN)* *"Extract all historical financial statements from `/Users/you/Desktop/Company_Raw.xlsx` and map the accounts under IFRS/US GAAP without editing the file."*
  - *(ES)* *"Extrae los estados financieros históricos de `/Users/you/Desktop/Empresa_Raw.xlsx` y prepara el mapeo de cuentas bajo IFRS sin modificar el archivo original."*
  - *(FR)* *"Extrais les états financiers historiques de `/Users/you/Desktop/Entreprise_Raw.xlsx` et prépare le mapping comptable IFRS/GAAP sans modifier le fichier."*

### New Skill: `historical-3stmt-rebuilder` (EN / FR / ES)
- **What it does:** Rebuilds the approved blueprint (or a clean source workbook) into a **brand-new** formula-driven `.xlsx` 3-statement workbook preserving verbatim company account names, cross-statement links, data validations, live reconciliation against source reported totals, and the Fresh-Chat Handoff Note.
- **Starter Prompts:**
  - *(EN)* *"Looks good — build the clean formula model and verify live in Excel that all totals match the raw file."*
  - *(ES)* *"Se ve perfecto — construye el nuevo modelo IFRS con todas las fórmulas, vínculos y validaciones, y verifica en vivo que los totales cuadren."*
  - *(FR)* *"C'est bon — construis le nouveau modèle 3 états avec toutes les formules et vérifie en direct dans Excel que les totaux correspondent."*

---

### Skill 1: `wall-street-formatter`
- **What it does:** Applies institutional Wall Street styling in one `batch_format_workbook` call (blue inputs, black same-sheet formulas, green cross-sheet links, navy `#1F4E79` headers, `#D9E1F2` bold totals with thin/double borders, contiguous currency & `%`/CAGR formatting, freeze panes, column widths) without touching any values or formulas.
- **Starter Prompts:**
  - *"Format `/Users/you/Desktop/Model.xlsx` Wall Street style and keep it board-ready."*
  - *"Apply institutional formatting to all three financial statements in `AAPL_Model.xlsx` without changing any numbers or formulas."*

### Skill 2: `model-auditor`
- **What it does:** Runs a strictly read-only SOX/IFRS/FASB audit checking for `#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, broken/external links, hardcoded numbers inside calculated rows, wrong denominators, and Balance Sheet imbalances.
- **Starter Prompts:**
  - *"Audit `/Users/you/Desktop/Model.xlsx` under US GAAP (FASB) for formula errors, hardcoded totals, and balance-check breaks before I send it to the CFO."*
  - *"Run a read-only IFRS model audit on my open Excel workbook and circle any invalid data."*

### Skill 3: `live-what-if-runner`
- **What it does:** Shocks driver inputs live in Microsoft Excel (`recalculate=True`), captures exact recalculated deltas across linked statements, and **automatically reverts** every input back to baseline.
- **Starter Prompts:**
  - *"Run a live sensitivity test on `/Users/you/Desktop/Model.xlsx`: shock FY2025 Revenue (`Income Statement!F2`) by +10% and -10%, report the exact deltas on Net Income, CFO, FCF, and Balance Check, then revert to baseline without saving."*
  - *"Test WACC at 8.5%, 9.5%, and 10.5% live in Excel and show the implied share price for each case, then restore the original WACC."*

### Skill 4: `three-statement-projector`
- **What it does:** Extends a historical 3-statement workbook with an `Assumptions` tab, a `Base`/`Bull`/`Bear` dropdown (`add_data_validation`), dynamic projected IS/BS/CF formulas, and live Excel verification that `Balance Check == 0` across all projected years.
- **Pre-Flight Inputs Confirmed:** IFRS vs. US GAAP, horizon (e.g., 5 years), and operating drivers (or approval to use trailing historical averages).
- **Starter Prompts:**
  - *"Using `/Users/you/Desktop/Historical_3Stmt.xlsx`, project FY2026–FY2030 under US GAAP with Base, Bull, and Bear scenarios and verify live in Excel that Balance Check = 0 for every year."*

### Skill 5: `capex-project-analyzer`
- **What it does:** Builds a 3-tab Capital Budgeting workbook (`Executive Summary`, `Project Cash Flows`, `Sensitivity & Charts`) with Year-0-safe `=NPV(...)`, `=IRR(...)`, `=MIRR(...)`, Profitability Index, Simple & Discounted Payback, 2-way WACC vs. Revenue sensitivity, and cash flow chart.
- **Starter Prompts:**
  - *"Build a CapEx analysis under US GAAP (Straight-Line depreciation) for a \$20M plant expansion over 7 years: \$8M Year-1 revenue growing 10%/yr, 35% EBITDA margin, 25% tax rate, 10% WACC, 5% NWC of revenue, and \$3M salvage value. Save to `/Users/you/Desktop/Plant_CapEx.xlsx`."*

### Skill 6: `dcf-investment-template-builder`
- **What it does:** Creates a reusable 5-tab DCF & Multi-Stream Cash Flow Valuation template (`Dashboard`, `WACC & Inputs`, `Segment Cash Flows`, `DCF Valuation`, `Sensitivity`) with CAPM WACC, Gordon Growth + Exit Multiple Terminal Value, EV-to-Equity Bridge (consistent with `IFRS 16`/`ASC 842`), and Base/Bull/Bear controls.
- **Starter Prompts:**
  - *"Build a reusable 5-year US GAAP DCF valuation template for a 2-segment SaaS + Hardware company with CAPM WACC, Gordon Growth and Exit EBITDA Multiple terminal value, EV-to-Equity Bridge, and WACC vs. Terminal Growth sensitivity. Save to `/Users/you/Desktop/DCF_Template.xlsx`."*

### Skill 7: `boardroom-brief-generator`
- **What it does:** Audits and reads live-recalculated KPIs from Excel and produces a styled executive Word memo (`.docx` + `.pdf`) and/or PowerPoint deck (`.pptx`) with 100% numeric traceability.
- **Starter Prompts:**
  - *"Read the live recalculated FY2021–FY2025 KPIs from `/Users/you/Desktop/AAPL_FullExcel_Test.xlsx` and generate a 1-page executive Word memo (`.docx` and `.pdf`) for the CFO."*

---

## 4. Safety, Compliance & Troubleshooting

- **What happens when `M365_MCP_FORMULA_GUARD` pauses?**  
  If the AI tries to overwrite an existing formula cell with a hardcoded number or blank, the server blocks the write (workbook stays untouched) and returns `status: "confirmation_required"`. The assistant will stop, explain what it wants to change and why, and only proceed if you explicitly reply *"Approved — go ahead."*
- **Model didn't ask for missing assumptions or guessed a driver?**  
  Make sure your `SYSTEM_PROMPT.md` includes the **Mandatory Halt Triggers** from [`Sample Model Config/SYSTEM_PROMPT.md`](../Sample%20Model%20Config/SYSTEM_PROMPT.md).
- **Excel file didn't show changes while already open?**  
  Verify `M365_MCP_LIVE_SYNC=AUTO` is set and that you don't have unsaved manual edits open in Excel (the server intentionally refuses to overwrite unsaved manual work).
- **Word/PowerPoint tools unavailable in `boardroom-brief-generator`?**  
  Switch `M365_MCP_PROFILE` to `EXCEL,WORD` or `FULL` (or `@`-mention your secondary `m365-mac-mcp` Word & PPT server in Bionic).
