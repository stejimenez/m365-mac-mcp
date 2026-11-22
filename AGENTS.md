# Repository Agent Rules (`m365-mac-mcp`)

## GitHub Release & Semantic Versioning Policy

Whenever updating GitHub for this repository, **always** cut a versioned release using industry-standard **Semantic Versioning (`SemVer 2.0.0` — `vMAJOR.MINOR.PATCH`)**:

- **`vX.Y.Z` — PATCH (`Z+1`)**: Bug fixes, minor hardening, test additions, or documentation updates.
- **`vX.Y.0` — MINOR (`Y+1`, reset `Z=0`)**: Backward-compatible new features, MCP tools, profiles, or skills.
- **`vX.0.0` — MAJOR (`X+1`, reset `Y=0, Z=0`)**: Breaking API, tool-schema, environment-variable, or architecture changes.

### Mandatory Checklist on Every GitHub Push

1. **Determine Version Bump**: Compare changes since the latest Git tag (`git describe --tags --abbrev=0`) and select `PATCH`, `MINOR`, or `MAJOR`.
2. **Update All README & Documentation Files**: Review and update every affected README and doc file so documentation is 100% in sync with the code before publishing:
   - `README.md` (features, tool/profile counts, setup, version badge)
   - `skills/README.md` (skill catalog, usage notes)
   - `docs/*.md` (`docs/HARDWARE_AND_PROFILE_GUIDE.md`, `docs/EXECUTIVE_COMPARISON_MATRIX.md`, and any other guides)
   - `CONTRIBUTING.md`, `SECURITY.md`, and `Sample Model Config/SYSTEM_PROMPT.md` (when relevant)
   - `CHANGELOG.md` → dated `## [vX.Y.Z] - YYYY-MM-DD` release entry
3. **Sync All Version References**:
   - `src/__init__.py` → `__version__`
   - `src/m365_mac_mcp_server.py` → `server_version`
   - `README.md` → version badge
   - `CHANGELOG.md` → dated `## [vX.Y.Z] - YYYY-MM-DD` release entry
4. **Verify Before Push**: Run `.venv/bin/pytest tests/test_basic.py tests/test_word_excel_rw.py -v` and confirm all tests pass.
5. **Commit, Tag & Push**:
   - Commit the release changes cleanly.
   - Create an annotated tag: `git tag -a vX.Y.Z -m "Release vX.Y.Z: <summary>"`
   - Push commit and tag together: `git push origin main --follow-tags`
6. **Publish GitHub Release**:
   - Create the matching GitHub Release automatically via `/opt/homebrew/bin/gh release create vX.Y.Z --title "vX.Y.Z — <Short Title>" --notes "<Release Notes>"`.

> **Note**: Do not commit, tag, push, or publish a release until the user explicitly requests pushing/updating GitHub.

---

## Privacy & Personal Configuration Policy

- **Never Upload Personal Config**: Never commit, push, or upload the user's personal configuration, username, local file paths, device names, or private environment details to GitHub unless you first **ask for explicit approval and explain why it needs to be shared**, or unless the user explicitly asks you to share it.
- Keep all personal/local environment tracking strictly in gitignored local files (such as `.agents/rules/local_environment_private.md`, matched by `*_private*` in `.gitignore`) for optimization and troubleshooting only.

---

## Engineering Standards (Efficiency, Tokens & Reliability)

When writing or updating **any** server code, MCP tool schema, system prompt, or Finance `SKILL.md`:
- **Efficiency & Low Latency**: Minimize tool round-trips (prefer 1 batch call per sheet over micro-calls or oversized multi-sheet payloads), keep operations atomic, and eliminate unnecessary disk/AppleScript overhead.
- **Token Utilization**: Keep tool schemas compact, JSON wire payloads tight (no boilerplate/duplicates), and skill playbooks concise so local and cloud models stay well within context limits.
- **Zero Hallucination**: Enforce ground-truth reads (live Excel recalculation for uncached formulas, exact verbatim values, hard stops when required numeric drivers are missing — never invent assumptions).
- **Zero Loops**: Return actionable, self-correcting errors and accept safe, universal parameter/orientation fallbacks so models recover in 1 turn instead of retrying blindly.

---

## Bionic & Local Model Quirks Policy

Whenever a quirk or issue **specifically related to Bionic or the local LLM model** is discovered and fixed (e.g., tool-calling behavior, schema parsing, context-window/loop quirks, MCP client transport/formatting issues), **always** record it in the `### Known Bionic & Model Quirks` section below (anonymized — no personal usernames, paths, or private config) so every future agent accounts for it when authoring code, tool schemas, or skills.

- **Strict Scope**: ONLY log issues specific to **Bionic** or the **model**. Do not log general Python, Excel/Office, macOS, or repository bugs here.
- **Cross-Platform / Universal Fix Requirement**: `m365-mac-mcp` supports multiple MCP clients and inference platforms (Bionic, LM Studio, Claude Desktop, Cline, VS Code, Antigravity, etc.). Whenever fixing a Bionic- or local-model-specific quirk, **always** design the fix as a **universal, backwards-compatible enhancement or graceful fallback** (never break standard MCP callers, positional/keyword APIs, or other clients).
- **Verify Changes on Any Issue**: Whenever an issue arises, assume the user's local baseline in `.agents/rules/local_environment_private.md` is active (always paired with the latest `Sample Model Config/SYSTEM_PROMPT.md`), **ask the user** if there was any change in the three items below, and **confirm updating** the local baseline if anything changed:
  1. Bionic Version
  2. Model Used
  3. Configuration being used
- **Entry Format**: Keep each entry concise, anonymized, and actionable:
  - `**Environment**: Bionic <version> | Model: <model> | Config: <anonymized relevant settings>`
  - `**Quirk**: <specific Bionic or model behavior observed>`
  - `**Rule / Mitigation**: <how future code, tool schemas, or skills must handle it>`

### Known Bionic & Model Quirks

1. **Dropped `workbook_id` in Large `params_json` Payloads**
   - **Environment**: Bionic `1.2.2+3` | Model: `Qwen3.6 35B A3B` (`GGUF Q8_0`) | Config: Context `65536`, Temp `0.3`, Reasoning Budget `2048`, `M365_MCP_PROFILE=EXCEL,COMPACT`
   - **Quirk**: Bionic wraps MCP tool invocations in `call_mcp_tool(mcp_name, tool_name, params_json)` where `params_json` is a double-escaped JSON string. On large batch calls (such as `batch_write_ranges`), the model serializes keys alphabetically (`"operations"` before `"workbook_id"`), emits ~2,000+ characters of nested escaped arrays, and then closes the JSON object while omitting `"workbook_id"`, causing Bionic's JSON schema validator to fail with `params requires property "workbook_id"`.
   - **Rule / Mitigation**: Keep `workbook_id: Optional[str] = None` consistently across **all** compact Excel workbook-handle tools (`list_worksheets`, `add_worksheet`, `delete_worksheet`, `read_range`, `write_cell`, `batch_write_ranges`, `batch_format_workbook`, `add_data_validation`, `add_excel_table`, `create_chart`, `save_workbook`, `close_workbook`) and auto-resolve the single active workbook in `ExcelController._resolve_workbook_id()` when `len(active_workbooks) == 1` (still raising a clear error when 0 or >1 workbooks are active).

2. **Soft "Confirm before building" Wording Ignored by Local Model**
   - **Environment**: Bionic `1.2.2+3` | Model: `Qwen3.6 35B A3B` (`GGUF Q8_0`) | Config: Context `65536`, Temp `0.3`, Reasoning Budget `2048`, `SYSTEM_PROMPT V3.4` + Finance `SKILL.md`
   - **Quirk**: When a skill's Pre-Flight section says *"Confirm concisely before building"* and the user provides a file path + partial request (e.g., CapEx outlay and horizon only), the model still calls `create_workbook` and fabricates "standard" assumptions instead of stopping to ask.
   - **Rule / Mitigation**: Use explicit **`HARD STOP (ZERO TOOLS)`** wording in both `SYSTEM_PROMPT.md` and builder `SKILL.md` Pre-Flight sections forbidding `create_workbook` or any tool call when required numeric drivers are missing.

3. **Double-Quote Triple-Escaping Failure in Bionic `params_json`**
   - **Environment**: Bionic `1.2.2+3` | Model: `Qwen3.6 35B A3B` (`GGUF Q8_0`) | Config: Context `65536`, Temp `0.3`, Reasoning Budget `2048`, `M365_MCP_PROFILE=EXCEL,COMPACT`
   - **Quirk**: Because Bionic passes MCP arguments inside a double-escaped `params_json` string, Excel formulas or number formats containing literal double quotes (e.g. `=IF(C16>0,"ACCEPT","REJECT")` or `$#,##0;($#,##0);"-"`) require triple escaping (`\\\"`). Local models frequently emit single-escaped `\"`, causing Bionic to reject the call with `params_json must be valid JSON.` before it reaches the MCP server.
   - **Rule / Mitigation**: Keep batch payloads compact (prefer 1 `batch_write_ranges` call per sheet on large models rather than stuffing 3–5 sheets into one giant `params_json`), and avoid unnecessary nested quoted literals in skill examples while keeping standard Excel syntax valid across all MCP clients.
