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
- **Serialized, Timeout-Guarded Automation:** `osascript` calls are serialized per event loop via `asyncio.Lock` and bounded by `OFFICE365_MCP_APPLESCRIPT_TIMEOUT` (default `30s`).
- **Formula & External Link Guard:** Formula writes automatically quote spaced sheet names and block/flag references to non-existent sheets or external workbooks (`[...]`) before touching live Excel, preventing unexpected modal file pickers and stray external links.
- **Unsaved-Edits Protection (`OFFICE_MCP_LIVE_SYNC`):** In `auto` sync mode, `save_workbook` refuses to overwrite a workbook that has unsaved manual edits in live Microsoft Excel.
- **macOS Sandbox Hygiene:** Unnamed temporary `/tmp` files are never auto-opened in live Word/Excel, avoiding macOS sandbox **Grant Access** prompts.

## Deployment Notes

- Only grant macOS **Automation** permissions (`System Settings → Privacy & Security → Automation`) to trusted MCP host applications.
- Review tool profiles (`OFFICE_MCP_PROFILE`) to restrict the exposed tool surface to only the Office apps required for your workflow.
