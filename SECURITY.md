# Security Policy — Microsoft 365 for Mac MCP (`m365-mac-mcp`)

**Author & Maintainer:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)

## Reporting Security Vulnerabilities

If you discover a security vulnerability, please email **Sterling Jimenez, Ph.D.** directly at `stejimenez@gmail.com`.

**Please do not report security vulnerabilities through public GitHub issues.**

## Supported Versions

Security updates are maintained for the current `2.x` release line.

## Security Architecture & Mitigations

This MCP server interacts with local files and macOS Microsoft Office applications. Key built-in safeguards include:

- **Parameterized AppleScript Execution:** All live Excel and Word AppleScript invocations run via `osascript -` with arguments passed strictly through `argv` (never shell or AppleScript string interpolation), preventing command/script injection.
- **Serialized, Timeout-Guarded Automation:** `osascript` calls are serialized per event loop via `asyncio.Lock` and bounded by `M365_MCP_APPLESCRIPT_TIMEOUT` (default `30s`).
- **Formula & External Link Guard:** Formula writes automatically quote spaced sheet names and block/flag references to non-existent sheets or external workbooks (`[...]`) before touching live Excel, preventing unexpected modal file pickers and stray external links.
- **Formula Overwrite Guard (`M365_MCP_FORMULA_GUARD`):** Defaults to `literals_only` (atomically blocks Formula → Hardcoded Literal / Blank overwrites across all 6 file-level and live Excel write tools with `awaiting_user_action: true` + `HARD STOP` until the user explicitly approves via `allow_formula_overwrite=true`; `strict` blocks any formula overwrite).
- **Write-Path Protection & Session-Scoped Write Fences (`validators.py`):**
  - **MCP Project-Folder Write Block:** `assert_safe_write_path()` permanently blocks writing user Office deliverables inside the MCP server's own installation directory.
  - **Session-Scoped `working_dir` Isolation:** Concurrent MCP sessions isolate `working_dir` via `contextvars.ContextVar` so multi-client / multi-chat paths never cross-contaminate.
  - **Optional Strict Write Fences:** `M365_MCP_BLOCKED_DIRS` denies writes to sensitive directories, and `M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR=true` restricts writes strictly to the active Project Working Folder (or `M365_MCP_OUTPUT_DIR`).
- **Atomic File Writes & Dirty-Save Protection:** All `.xlsx`, `.docx`, and `.pptx` writes commit atomically via `os.replace()` (preventing truncated/corrupt files on mid-save failure). In `M365_MCP_LIVE_SYNC=auto`, `save_workbook` refuses to overwrite unsaved manual edits in Excel, and both `excel_live_save` and `word_live_save` refuse to clobber unsaved in-memory edits (`UnsavedInMemoryEditsError`).
- **Away-From-Keyboard (AFK) Hard-Stop Guards & macOS Sandbox Hygiene:** Unnamed temporary `/tmp` files are never auto-opened in live Word/Excel; LaunchServices auto-open is scoped strictly to the trusted Project Working Folder; and modal/Grant Access timeouts, Automation `-1743` denials, OS Files & Folders `PermissionError`s, write-fence blocks, unsaved manual edits, and Formula Guard blocks all emit `awaiting_user_action: true` with an explicit `HARD STOP`.
- **Focus-Safe Background Multitasking (`M365_MCP_BACKGROUND_MODE`):** Defaults to `no_focus` (`open -g -j`, ignores LLM `activate=True`, minimizes newly opened MCP Office windows, and restores any previously active user workbook/document inside Excel/Word) so live MCP calls never steal OS focus or capture accidental user keystrokes while multitasking (`strict` hides MCP windows; `off` restores legacy focus behavior).

## Deployment Notes

- Only grant macOS **Automation** permissions (`System Settings → Privacy & Security → Automation`) to trusted MCP host applications.
- Review tool profiles (`M365_MCP_PROFILE`) to restrict the exposed tool surface to only the Office apps required for your workflow.
