# Changelog

All notable changes to **Microsoft 365 for Mac MCP (`m365-mac-mcp`)** are documented in this file.

This project follows [Semantic Versioning (`MAJOR.MINOR.PATCH`)](https://semver.org/):
- **MAJOR (`X.0.0`)**: Breaking architectural or API changes
- **MINOR (`2.X.0`)**: Backward-compatible new features, tools, or skills
- **PATCH (`2.1.X`)**: Bug fixes, hardening, and documentation updates

---

## [v2.3.1] - 2026-11-23

### Fixed
- **OOXML Data Validation List Quoting (`add_data_validation`):** Automatically double-quotes literal comma-separated list strings (e.g., `"Base,Upside,Downside"` → `'"Base,Upside,Downside"'`) while preserving range references (`Sheet1!A1:A3`, `$A$1:$A$3`), fixing Microsoft Excel `Parameter error (-50)` / corrupt `.xlsx` open failures on string-passed dropdowns.
- **Excel Table Empty-Header Guard (`_ensure_valid_excel_table_headers`):** Validates and auto-repairs empty, formula, or duplicate OOXML Table header cells right before `wb.save()` so saved workbooks never trigger Excel `-50` file-open corruption.
- **`batch_format_workbook` `column_widths` List Fallback:** Accepts `column_widths` as either a dict (`{"A": 15, "B:F": 12}`) or a list (`[15, 10, ...]`, mapped sequentially from the op's range start column or `A`), eliminating raw `AttributeError` crashes and returning a clear `ValueError` hint on invalid types.
- **Anti-Loop Guards & Actionable Error Hints:**
  - `excel_live_*` timeouts now explicitly instruct stopping to let the user dismiss/approve any Excel modal or **Grant Access** dialog instead of suggesting file-level `read_range` fallback.
  - `read_range(evaluate_formulas=True)` now returns `uncached_formula_count` + an explicit anti-loop `warning` when formula cells are uncached (`null`).
  - `batch_write_ranges` redirects formatting-only ops (`number_format`, `styles`, `column_widths`, etc.) to `batch_format_workbook`, and auto-wraps scalar `"value"` / `"values"` / `"data"` entries as 1×1 `[[val]]`.
  - `batch_format_workbook` supports singular `column_width` + `range` (`"B:F"`, `"B1:F5"`) and skips empty no-op formatting rules.
  - `add_excel_table` accepts `range` as a fallback alias for `data_range`, defaults `table_name` to `"Table1"`, and names exact required parameters on empty-range errors.
  - `create_chart` accepts common chart-type aliases (`column`, `column_clustered`, `bar_clustered`, `bar_stacked`, `line_markers`, `pie_3d`, `doughnut`, etc.) and auto-detects text-only first columns as X-axis categories instead of plotting them as blank numeric series.
  - `excel_live_batch_read_ranges` accepts `cell_ref`/`cell` aliases for `range_ref` and `sheet`/`worksheet`/`worksheet_name`/`tab` aliases for `sheet_name`.
  - `create_document`, `create_workbook`, and `create_presentation` docstrings and `Sample Model Config/SYSTEM_PROMPT.md` clarify that parent directories are auto-created on save and anchor MCP routing via `check_office_status` across both wrapper clients (Bionic) and native MCP clients/CLIs.
- **Excel 16.115 & Word 16.115 Live AppleScript Hardening:**
  - Fixed `excel_live_list_worksheets`, `excel_live_add_worksheet`, and `excel_live_activate_worksheet` `-50`/`-1700` errors on modern Microsoft Excel for Mac (`16.115`).
  - Fixed `save_document(format="pdf"|"rtf"|"txt"|"doc")` `-1708` errors in Microsoft Word 16.115 by opening export documents read-write (`read_only=False`), guarding post-`save as` document name reads, and auto-converting Word table `column_widths` passed in points (`>15`) or twips (`>500`) to inches.

### Added
- **Model-Agnostic Stress-Test Suite (`Model-MCP Stress Test/STRESS_TEST_SUITE.md`):** 9-Phase / 12-Chat stress-test battery covering all 82 MCP tools (`EXCEL,COMPACT`, `EXCEL`, and `WORD,POWERPOINT`) and all 7 non-import Finance Skills.
- **Local Model Compatibility Documentation:** Documented tested local models (`Qwen3.6 35B A3B` GGUF Q6_K/Q8_0 primary, `Qwen3.8 27B` MLX `Compatible (Slow)`, `Gemma 4 26B A4B` MLX on `oMLX 0.5.0` `Not Compatible`, `Nemotron 3 Nano Omni` GGUF `Not Compatible`) with root-cause reasons and recommended Bionic/llama.cpp settings (`Preserve Thinking = ON`, `Presence Penalty = 0.2`, `Repeat Penalty = 1.0`, avoiding `&` in MCP server names) across `README.md`, `docs/HARDWARE_AND_PROFILE_GUIDE.md`, `docs/EXECUTIVE_COMPARISON_MATRIX.md`, `skills/README.md`, and `AGENTS.md`.
- **Automated Regression Tests:** Expanded suite to **38 passing tests** covering OOXML list quoting, `column_widths` list/singular fallback, `create_chart` auto-categories & aliases, Excel Table header repair, uncached-formula warnings, live timeout hints, scalar `value`/`values`, and `add_excel_table` hints.

---

## [v2.3.0] - 2026-11-22

### Added
- **Excel-Style Formula Drag/Fill (`fill_formula`) in `batch_write_ranges`:** Supports 1D horizontal/vertical rows and 2D sensitivity grids (`{"sheet_name": "...", "range": "C4:H4", "fill_formula": "=C2*C3"}`) using `openpyxl.formula.translate.Translator`, cutting formula output tokens by 60–80%, eliminating off-by-one column typos, and respecting `$` anchors, spaced sheet quoting, and `M365_MCP_FORMULA_GUARD`.
- **Opt-In MCP Wire Debug Logger (`M365_MCP_DEBUG_LOG`):** Optional local JSONL logger (off by default) recording tool name, arguments, latency (ms), status, warnings, errors, and response previews for real-time debugging without polluting `stdout`/`stderr`.
- **Dedicated Gemma 4 System Prompt (`Sample Model Config/SYSTEM_PROMPT_GEMMA4_31B.md`):** Separate Finance & M365 system prompt tailored for Gemma 4 (`31B` / `26B A4B` MLX) with strict JSON, `fill_formula`, `SKILL-FIRST`, and `HARD STOP` rules.
- **Dirty In-Memory Warning on Live Excel Reads:** `excel_live_read_range` and `excel_live_batch_read_ranges` now emit a 1-line `warning` **only** when a matching in-memory `openpyxl` workbook has unsaved file-level edits not yet synced via `save_workbook` (zero extra tokens when clean).
- **Automated Test Coverage:** Expanded suite from **30 → 37 passing tests** covering `fill_formula` 1D/2D translation, `workbook_id` auto-resolution, LLM parameter aliases & shape-aware orientation, internal temp-file Grant Access guards, live `#VALUE!`/`#DIV/0!` surfacing, boolean `save_changes` coercion, dirty-state warnings, and `M365_MCP_DEBUG_LOG`.

### Changed
- **Universal `workbook_id` Auto-Resolution:** Made `workbook_id: Optional[str] = None` across all 12 compact Excel workbook-handle tools (`list_worksheets`, `add_worksheet`, `delete_worksheet`, `read_range`, `write_cell`, `batch_write_ranges`, `batch_format_workbook`, `add_data_validation`, `add_excel_table`, `create_chart`, `save_workbook`, `close_workbook`), auto-resolving when exactly 1 workbook is active in memory while remaining 100% backwards-compatible for explicit positional/keyword callers.
- **Universal LLM Parameter Aliases & Shape-Aware Orientation:**
  - `batch_write_ranges` now resolves `start_cell`/`range`/`cell_range`/`cell`, `sheet_name`/`sheet`/`worksheet`/`worksheet_name`/`tab`, and `data`/`values`, wraps flat 1D lists, and auto-orients transposed `N×1`/`1×N` arrays to match explicit ranges (e.g., `B3:F3`).
  - `read_range` now accepts `range` alongside `cell_range` and `range_ref`; `excel_live_close` and `word_live_close` coerce boolean `save_changes` (`true`/`false` → `"yes"`/`"no"`).
- **Hardened Live Excel Recalculation & macOS Sandbox Safety:**
  - Added live `calculate` in `open_excel_file` and `batch_read_live_excel_ranges`, plus `string value` fallback so `#VALUE!`, `#DIV/0!`, `#REF!`, `#NAME?`, `#N/A`, and `#NUM!` surface verbatim instead of blank strings.
  - Blocked `excel_live_open` and `word_live_open` on internal `/tmp` and macOS `/var/folders/.../T/` UUID temp files to eliminate macOS Office **Grant Access** modal hangs.
- **Hardened Prompts & Builder Skills:** Updated `SYSTEM_PROMPT.md` (`HARD STOP (ZERO TOOLS)`, `SKILL-FIRST`, no sandbox/Python bypass, `save_workbook` sync, no mental pre-calculation) and builder skills (`capex-project-analyzer`, `dcf-investment-template-builder`, `three-statement-projector`) for 1-call-per-sheet `fill_formula` builds; updated `README.md` reference config to `Repeat Penalty = 1.0 (Off)`.

---

## [v2.2.0] - 2026-11-21

### Added
- **2 New Multilingual (EN / FR / ES) Financial Statement Normalization Skills (`skills/`):**
  - **`financial-statement-extractor`:** Read-only discovery (`list_worksheets` → targeted `read_range`) of any raw company Excel file across any number of historical years, preserving every unique company account name verbatim and producing an IFRS/FASB Account Mapping Blueprint for 1-click approval.
  - **`historical-3stmt-rebuilder`:** Rebuilds raw company financials into a brand-new, formula-driven 3-statement `.xlsx` workbook (preserving verbatim account names, adding dynamic subtotals/totals/margins/CAGRs, cross-statement links, data validation, and live Excel proof that rebuilt totals equal raw reported totals).
- **Built-In Fresh-Chat Handoff Notes:** Added deterministic phase-boundary Fresh-Chat Handoff guidance (tied to Bionic's bottom-right `Context` meter) across `historical-3stmt-rebuilder`, `three-statement-projector`, `capex-project-analyzer`, and `dcf-investment-template-builder` without risky LLM token-count guessing.
- **Multilingual System Prompt (`V3.4`):** Updated `Sample Model Config/SYSTEM_PROMPT.md` to respond in the user's language (English, French, or Spanish) and register all 9 Finance Skills.
- **Automated `excel,compact` Skill-Coverage Test:** Extended `test_all_finance_skills_valid_schema_size_and_compliance` to verify all 9 skills (2–4 KB, 5-block, IFRS/FASB/SOX) and guarantee zero missing tools in `excel,compact` across all 8 Excel skills.
- **Targeted Multi-Tab Inspection (`inspect_workbook(sheet_names=...)`) & `read_range` Guardrails:** Added optional `sheet_names` filtering to `inspect_workbook` (returning both evaluated values and sparse A1 formulas for only the requested tabs in 1 call) plus `max_cols=50` capping and trailing-blank trimming on explicit `read_range` calls.

### Changed
- **Standardized `m365-mac-mcp` Naming Across Code, Skills & Docs (Backwards-Compatible):** Renamed main entrypoint to `src/m365_mac_mcp_server.py` (with a thin deprecation shim at `src/office365_mcp_server.py`), standardized the MCP server wire name (`m365-mac-mcp`), resource URI (`m365-mac-mcp://status`), temp directories (`/tmp/m365-mac-mcp`), and all environment variables (`M365_MCP_PROFILE`, `M365_MCP_LIVE_SYNC`, `M365_MCP_FORMULA_GUARD`, `M365_MCP_LOG_LEVEL`, `M365_MCP_APPLESCRIPT_TIMEOUT`, etc.) with silent deprecation-warning fallback for legacy `OFFICE_MCP_*` / `OFFICE365_MCP_*` keys, plus updated System Prompt (`V3.4`), all 9 Finance Skills, and documentation.
- **High-Precision SOX 404 Formula Audit:** Refined `audit_workbook_formulas` `suspicious_literals` detection to flag literal `'""'`, horizontal formula-row breaks, and hardcoded values on calculated rows (EN/FR/ES keywords) — eliminating ~95% of false-positive warnings on legitimate historical input rows.
- **Deferred Disk Persistence & Full `STRICT` Formula Guard:** Deferred `.xlsx` disk serialization until `save_workbook` (preventing redundant full-zip rewrites and ensuring intermediate writes never bypass `M365_MCP_LIVE_SYNC` unsaved-edits protection), staged PDF exports via temp files, extended `M365_MCP_FORMULA_GUARD=STRICT` to `add_formula`, and added `close_workbook` cleanup to all builder skills.

---

## [v2.1.2] - 2026-11-21

### Added
- **Standalone Executive Comparison Matrix (`docs/EXECUTIVE_COMPARISON_MATRIX.md`):** Added a dedicated, quick-reference Markdown matrix comparing 16 GB, 32 GB, and 64 GB Apple Silicon M4/M4 Pro/M4 Max configurations for `m365-mac-mcp`.

---

## [v2.1.1] - 2026-11-21

### Added
- **Hardware, Local Model & MCP Profile Guide (`docs/HARDWARE_AND_PROFILE_GUIDE.md`):** Comprehensive Apple Silicon (16 GB / 32 GB / 64 GB M4, M4 Pro, M4 Max) local model recommendations (MLX & GGUF) and a full 8-profile `OFFICE_MCP_PROFILE` performance, skill-compatibility, and hallucination-risk matrix.
- **Main README Summary:** Added a compact Hardware Tier & MCP Profile quick-reference table in `README.md` linking to `docs/HARDWARE_AND_PROFILE_GUIDE.md`.

---

## [v2.1.0] - 2026-11-21

### Added
- **`OFFICE_MCP_FORMULA_GUARD` (`literals_only` default | `strict` | `off`):** Atomic pre-write formula protection across all 6 Excel write tools (`write_cell`, `write_range`, `batch_write_ranges`, `excel_live_write_cell`, `excel_live_write_range`, `excel_live_batch_write_cells`), returning `status: "confirmation_required"` before modifying any cell and reporting `"formulas_updated"` counts.
- **One-Call Multi-Tab Creation:** Added optional `sheet_names: List[str]` parameter to `create_workbook` to create multiple tabs in a single call with zero leftover blank `Sheet1`.
- **7 IFRS / FASB / SOX 404 Finance Skills (`skills/`):** Added pure-Markdown (`2.5–3.4 KB` each) Finance Skills (`wall-street-formatter`, `model-auditor`, `live-what-if-runner`, `three-statement-projector`, `capex-project-analyzer`, `dcf-investment-template-builder`, `boardroom-brief-generator`) plus a comprehensive WIP user guide in `skills/README.md`.
- **`Sample Model Config/`:** Added verified Mac mini M4 Pro (64 GB) + `Qwen3.6 35B A3B — Q8_0` LM Studio & Bionic MCP screenshots and hardened **Finance & M365 Specialist System Prompt (`V3.4`)**.
- **Automated Test Coverage:** Expanded suite to **29 passing tests** including formula guard, LLM style aliases, multi-sheet `create_workbook`, and automated Finance Skills schema/compliance verification.

### Changed
- **Token & Latency Optimization (-43.3% Turn-1 Static Prefix):**
  - Rebalanced `OFFICE_MCP_PROFILE=excel,compact` to **25 non-overlapping tools (~2.4K tokens)** covering 100% of all 6 Excel Finance Skills (`create_chart`, `delete_worksheet`, and `excel_live_list_workbooks` promoted to compact).
  - Compacted MCP wire responses (`separators=(",", ":")`, `structured_output=False`), sparse A1 formula maps in `inspect_workbook` & live Excel reads, deduplicated Word `read_document` (`include_full_text=False` default), and stripped redundant Pydantic schema noise (`title`, nullable `anyOf`, `default: null`, `additionalProperties: true`).
- **Hardened `batch_format_workbook`:** Added support for flat per-range operation items, style-list aliases (`rules`, `formats`, `number_formats`), nested `font`/`fill`/`alignment`/`style` dicts, and explicit unknown-key `"warnings"`.
- **Blank Env Var Safety:** Hardened `src/utils/config.py` so blank/whitespace environment variable values safely fall back to documented defaults.
- **Licensing:** Switched to **Custom Non-Commercial / Source-Available License** (free for personal & academic use; commercial license available via `stejimenez@gmail.com`) with preserved upstream MIT attribution in `NOTICE`.

---

## [v2.0.0] - 2026-11-20

### Added
- Rebranded to **Microsoft 365 for Mac MCP (`m365-mac-mcp`)** by **Sterling Jimenez, Ph.D.** (`stejimenez@gmail.com`).
- Dual-mode architecture combining file-level manipulation (`openpyxl`, `python-docx`, `python-pptx`) with serialized, `argv`-safe AppleScript automation verified on **macOS 27.2.1** and **Microsoft Office for Mac 16.115 / 16.115.1**.
- Composable `OFFICE_MCP_PROFILE` tool filtering (`full`, `compact`, `excel`, `excel,compact`, `word`, `powerpoint`, `excel,word`, `word,powerpoint`).
- Live Excel Auto-Sync (`OFFICE_MCP_LIVE_SYNC=auto|off`) with window/sheet/zoom/scroll/selection restoration and unsaved-edits protection.
- Batch Excel & live AppleScript operations (`batch_write_ranges`, `batch_format_workbook`, `excel_live_batch_read_ranges`, `excel_live_batch_write_cells`), formula auditing (`audit_workbook_formulas`), one-call inspection (`inspect_workbook`), native Excel tables, named cell styles, data validation, invalid-data circles, and rich `.docx` table/paragraph styling.
