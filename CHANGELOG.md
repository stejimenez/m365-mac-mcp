# Changelog

All notable changes to **Microsoft 365 for Mac MCP (`m365-mac-mcp`)** are documented in this file.

This project follows [Semantic Versioning (`MAJOR.MINOR.PATCH`)](https://semver.org/):
- **MAJOR (`X.0.0`)**: Breaking architectural or API changes
- **MINOR (`2.X.0`)**: Backward-compatible new features, tools, or skills
- **PATCH (`2.1.X`)**: Bug fixes, hardening, and documentation updates

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
