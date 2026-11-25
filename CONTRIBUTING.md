# Contributing Guidelines

Thank you for your interest in contributing to **Microsoft 365 for Mac MCP** (`m365-mac-mcp`)!

**Author & Maintainer:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)

## How to Contribute

1. **Fork the repository**
2. **Create a feature branch** (`git checkout -b feature/amazing-feature`)
3. **Make your changes**
4. **Add tests** for new functionality in `tests/test_word_excel_rw.py`
5. **Ensure all tests pass:**
   ```bash
   .venv/bin/pytest tests/test_basic.py tests/test_word_excel_rw.py -v
   ```
6. **Commit your changes** (`git commit -m 'Add amazing feature'`)
7. **Push to the branch** (`git push origin feature/amazing-feature`)
8. **Submit a Pull Request**

## Code Style & Architecture Rules

- Follow PEP 8 Python style guidelines and include type hints on all public functions and methods.
- **Never write to `stdout`** — MCP uses `stdio` JSON-RPC; all logging must go through `utils.logger.setup_logger` (`stderr` only).
- **Register tools with `@_tool(app=..., compact=...)`** in `src/m365_mac_mcp_server.py` so `M365_MCP_PROFILE`, compact JSON serialization, and schema compaction apply automatically.
- **Keep modular facades intact:** Add Excel file-level code to the matching domain mixin under `src/controllers/excel/` (`io_lifecycle.py`, `reader.py`, `writer.py`, `charts_tables.py`, `formatter.py`) and live Office AppleScript code under `src/integrations/applescript/` (`base.py`, `word_bridge.py`, `excel_core.py`, `excel_io.py`, `ppt_bridge.py`), keeping `ExcelController` and `AppleScriptBridge` as thin, backwards-compatible facades.
- **Safe writes & sparse scans:** Always validate write paths via `resolve_file_path(..., for_write=True)` / `assert_safe_write_path()`, persist Office archives atomically via `atomic_save_via_temp()`, and iterate populated Excel cells via `ws._cells` (never dense `1..ws.max_row` loops) so 1M-row "ghost" sheets never stall or OOM.
- **Keep AppleScript calls safe:** Always pass user inputs via `osascript -` + `argv` (never string interpolation) and respect the per-event-loop `asyncio.Lock` and timeouts in `src/integrations/applescript/base.py`.
- Keep tool docstrings concise and include key-name guides for structured dict arguments.
- **Finance Skills (`skills/*/SKILL.md`):** Keep each skill pure Markdown (**2–4 KB**), structured with the 5-block blueprint (`## 1.`–`## 5.`), IFRS/FASB/SOX compliant, token/loop-optimized, referencing only registered `m365-mac-mcp` tool/parameter names (validated automatically by `pytest`), and strictly **finance-native in user-facing output** (assume the user knows finance, never internal MCP tools, skill slugs, or code internals).

## Testing

- Add unit/integration tests for any new functionality.
- Ensure all existing tests pass without writing to `stdout`.
- Verify live-app changes on macOS with Microsoft Excel / Word installed.

## Questions?

Open an issue for discussion or contact **Sterling Jimenez, Ph.D.** at `stejimenez@gmail.com`.
