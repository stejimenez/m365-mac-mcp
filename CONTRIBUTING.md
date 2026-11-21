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
- **Register tools with `@_tool(app=..., compact=...)`** in `src/office365_mcp_server.py` so `OFFICE_MCP_PROFILE`, compact JSON serialization, and schema compaction apply automatically.
- **Keep AppleScript calls safe:** Always pass user inputs via `osascript -` + `argv` (never string interpolation) and respect the per-event-loop `asyncio.Lock` and timeouts in `src/integrations/applescript_bridge.py`.
- Keep tool docstrings concise and include key-name guides for structured dict arguments.
- **Finance Skills (`skills/*/SKILL.md`):** Keep each skill pure Markdown (**2–4 KB**), structured with the 5-block blueprint (`## 1.`–`## 5.`), IFRS/FASB/SOX compliant, and referencing only registered `m365-mac-mcp` tool/parameter names (validated automatically by `pytest`).

## Testing

- Add unit/integration tests for any new functionality.
- Ensure all existing tests pass without writing to `stdout`.
- Verify live-app changes on macOS with Microsoft Excel / Word installed.

## Questions?

Open an issue for discussion or contact **Sterling Jimenez, Ph.D.** at `stejimenez@gmail.com`.
