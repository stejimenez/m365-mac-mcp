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
- **Grouped by Model**: Organize all entries under model-specific subheadings (`#### <Model Name> (<Format/Quant>)`) so quirks and universal fixes are tracked cleanly per model.
- **Entry Format**: Keep each entry concise, anonymized, and actionable:
  - `**Environment**: Bionic <version> | Model: <model> | Config: <anonymized relevant settings>`
  - `**Quirk**: <specific Bionic or model behavior observed>`
  - `**Rule / Mitigation**: <how future code, tool schemas, or skills must handle it>`

### Known Bionic & Model Quirks

#### `Qwen3.6 35B A3B` (`GGUF Q8_0`)

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

4. **Bionic Top-Level `create_` → `create_folder` Collision Loop (`Exists folder` ×50–300)**
   - **Environment**: Bionic `1.2.2+3` | Models: `Qwen3.6 35B A3B` (`GGUF Q6_K`), `Nemotron 3 Nano Omni` (`GGUF Q8_0`), `Gemma 4 26B A4B` (`MLX 8bit`) | Config: `Preserve Thinking = Off`
   - **Quirk**: In Bionic, Office MCP tools (`create_document`, `create_workbook`, `create_presentation`) are wrapped inside `call_mcp_tool`, while Bionic's ONLY top-level tool starting with `create_` is `create_folder`. When a fresh chat tries to call `create_document`/`create_workbook` *before* any `call_mcp_tool` call exists in context, the model literally thinks `"Good, folder exists. Now let me create the document."` yet its tool-call prefix `create_` gets captured by top-level `create_folder` (`Exists folder`), returning `"Folder ... already exists."` and trapping the model in an infinite loop (especially with `Preserve Thinking = Off`). Additionally, naming an MCP server with `&` (e.g. `Word & PPT 365 MCP`) can trigger `&amp;` HTML escaping.
   - **Rule / Mitigation**: (1) Explicitly anchor `call_mcp_tool` in `SYSTEM_PROMPT.md` by invoking `check_office_status` via `call_mcp_tool` first in fresh creation chats and forbidding top-level `create_folder` + Bionic's built-in `Documents, presentations, graphics & Canvas diagrams` skill; (2) enable `Preserve Thinking = On` on Qwen3.6 (with `Reasoning Budget = 2048`) and eject/reload the model after changing Model Defaults; (3) avoid `&` in MCP server names (`Word and PPT 365 MCP`).

#### `Gemma 4 26B A4B` (`MLX 8bit MoE`)

1. **`read_range(evaluate_formulas=true)` Retry Loop & Mental-Math Hallucination After Live Timeout**
   - **Environment**: Bionic `1.2.2+3` (`oMLX 0.5.0`) | Model: `Gemma 4 26B A4B` (`MLX 8bit`) | Config: Context `65536`, Temp `0.2`, Top K `40`, Top P `0.95`, Min P `0.05`, Repeat Penalty `1`, `M365_MCP_PROFILE=EXCEL,COMPACT`
   - **Quirk**: When `excel_live_open` timed out on a macOS modal/Grant Access dialog and returned a generic fallback hint suggesting file-level tools, the model looped 5 times on `read_range(evaluate_formulas=true)` across slightly varied ranges (`F5`, `A1:F5`, `F2:F5`, `B2:F5`) trying to coax evaluated numbers out of uncached `null` cells, then fabricated an incorrect mental sum (`1,340` instead of `1,195`) and leaked internal terms (`AppleScript`, `MCP`).
   - **Rule / Mitigation**: Keep `excel_live_*` timeout hints in `_error_response()` strictly instructing the model to **stop and ask the user to dismiss/approve the Excel dialog** (never suggest file-level fallback for formula evaluation), and return `uncached_formula_count` + an explicit anti-loop `warning` in `ExcelController.read_range(evaluate_formulas=True)` whenever formula cells are uncached (`null`).

2. **`column_widths` Passed as List + Destructive `write_cell(value=null)` Fallback & Thought-Only Exit**
   - **Environment**: Bionic `1.2.2+3` (`oMLX 0.5.0`) | Model: `Gemma 4 26B A4B` (`MLX 8bit`) | Config: Context `65536`, Temp `0.2`, Top K `40`, Top P `0.95`, Min P `0.05`, Repeat Penalty `1`, `M365_MCP_PROFILE=EXCEL,COMPACT`
   - **Quirk**: The model recovers in 1 turn when MCP errors include clear hints, but when `batch_format_workbook` received `column_widths` as a JSON array (`[15, 10, ...]`) and threw a raw `AttributeError` without a hint, the model attempted `write_cell(..., cell="A1", value=null, formatting={"column_width": 15})`, wiping an Excel Table header cell (`A1`) and causing Excel to reject the saved `.xlsx` with `Parameter error (-50)`. On hitting the unrecoverable `-50` error, Bionic + `oMLX` Gemma 4 terminated with a blank/collapsed `thought` block and zero visible text.
   - **Rule / Mitigation**: Accept `column_widths` in `batch_format_workbook` as **either** a dict (`{"A": 15, "B:F": 12}`) **or** a list (`[15, 10, ...]`, mapped sequentially from the op's range start column or `A`) with a clear `ValueError` hint for invalid types, and run `ExcelController._ensure_valid_excel_table_headers()` before `wb.save()` so OOXML Table headers are always non-empty, non-formula, and unique.

3. **Parameter Alias Loops on `add_excel_table` & `batch_write_ranges` (Formatting-Only / Scalar `value`)**
   - **Environment**: Bionic `1.2.2+3` (`oMLX 0.5.0`) | Model: `Gemma 4 26B A4B` (`MLX 8bit`) | Config: Context `65536`, Temp `0.2`, Top K `40`, Top P `0.95`, Min P `0.05`, Repeat Penalty `1`, `M365_MCP_PROFILE=EXCEL,COMPACT`
   - **Quirk**: Repeated `add_excel_table` 7× when an aliased key (`range`) was stripped leaving `data_range=""`, repeated `batch_write_ranges` 5× with formatting-only `number_format` ops (no `data`), and passed scalar `{"cell": "B2", "value": "Base"}` inside `batch_write_ranges`.
   - **Rule / Mitigation**: Accept `range` alias in `add_excel_table` with explicit parameter-name error hints, redirect formatting-only `batch_write_ranges` ops to `batch_format_workbook` in the error message, and auto-wrap scalar `"value"` as `[[value]]` in `batch_write_ranges`.

4. **Random Premature Bare-`thought` Stop in Multi-Tool Chains (`oMLX 0.5.0`) & Pre-Tool Narration Roleplay**
   - **Environment**: Bionic `1.2.2+3` (`oMLX 0.5.0`) | Model: `Gemma 4 26B A4B` (`MLX 8bit`) | Config: Context `65536`, Temp `0.2` (and `0.6` test), Top K `40`/`64`, `M365_MCP_PROFILE=EXCEL,COMPACT`
   - **Quirk**: (a) Across multiple runs, `oMLX 0.5.0` randomly terminated mid-workflow right after a *successful* MCP tool call (Call #1, #10, #13) with a bare `thought` block and zero user-facing output, forcing the user to manually type `"continue"`. (b) Instructing the model to emit short progress narration before tool calls caused it to either loop verbatim in `Reasoning` (at `Temp 0.2`) or roleplay tool calls in prose without calling tools (at `Temp 0.6`).
   - **Rule / Mitigation**: Keep `ZERO CHATTER` (no pre-tool narration) and `Temp 0.2` for Gemma 4 26B A4B; minimize round-trips via universal parameter aliases so multi-step tasks finish in fewer turns; treat random bare-`thought` stops as an `oMLX` engine UX defect when evaluating model validation.

---

## Model Validation Policy & Tracker

- **Goal**: Make `m365-mac-mcp` work flawlessly across every tested model via universal, backwards-compatible fixes while logging all quirks and fixes **by model** above.
- **Validation Criterion**: A model is marked **`Validated`** **only** when **100% of the tests in `Model-MCP Stress Test/STRESS_TEST_SUITE.md` pass as explicitly confirmed by the user**. Never mark a model validated based on partial passes or assumptions.
- **Strict Sequential Gate (100% Pass Before Advancing)**: When debugging or running the stress-test suite, **always** re-run the exact same test/prompt after any fix until it passes **100%** (confirmed by the user); **never** suggest moving on to the next test or phase until the current test passes 100%.
- **Keep Updated**: Update the tracker below as each phase is confirmed by the user (keep entries anonymized — no personal paths or private config).

| Model | Format / Quant | Dedicated Prompt | Status | Confirmed Passed Phases / Rejection Reason | Validated Date |
| :--- | :--- | :--- | :---: | :--- | :---: |
| `Qwen3.6 35B A3B` | `GGUF Q6_K` / `Q8_0` | `Sample Model Config/SYSTEM_PROMPT.md` | `In Progress` | `0A`, `0B` *(all 50 Excel tools passed, 0 errors, 1-shot)* | — |
| `Gemma 4 26B A4B` | `MLX 8bit MoE` (`oMLX 0.5.0`) | `Sample Model Config/SYSTEM_PROMPT_GEMMA4_26B_A4B.md` | **`Not Compatible`** | **Rejected:** `oMLX 0.5.0` Gemma 4 token-parser bug — randomly terminates mid-workflow on bare `thought` (forcing manual `continue` with zero feedback) and leaks raw `call:call_mcp_tool{...<\|"\|>...}` control tokens into chat text. | — |
| `Nemotron 3 Nano Omni` | `GGUF Q8_0` (`llama.cpp`) | `Sample Model Config/SYSTEM_PROMPT_NEMOTRON3_NANO_OMNI.md` | **`Not Compatible`** | **Rejected:** Weak multi-step MCP planning — looped 15× on built-in folder `exists` checks, wrote tables cell-by-cell (`write_cell`), hallucinated file paths instead of calling `create_workbook`, and ignored `.internal/skills` ban. | — |
| `Qwen3.8 27B` | `MLX 8bit` (`oMLX 0.5.0`, dense) | `Sample Model Config/SYSTEM_PROMPT.md` | **`Compatible (Slow)`** | **Compatible, but slow for interactive use:** Passed 0A-1 with 10/10 accuracy (11 calls, 0 errors, exact `$1,195` live), but dense 27B 8-bit on `oMLX` (no `Reasoning Budget` cap) took **16m 50s** for a 2-tab smoke test *(not yet full-suite validated)*. | — |
| `Gemma 4 31B` | `MLX 8bit` | `Sample Model Config/SYSTEM_PROMPT_GEMMA4_31B.md` | `Not Tested` | *(Untested; shares `oMLX 0.5.0` Gemma 4 runtime)* | — |
