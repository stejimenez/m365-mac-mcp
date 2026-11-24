# Hardware, Local Model & MCP Profile Guide — Microsoft 365 for Mac MCP (`m365-mac-mcp`)

**Author:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)  
**Target Stack:** Apple Silicon (M4 / M4 Pro / M4 Max) · Bionic / LM Studio (MLX & GGUF) · Microsoft 365 for Mac (`16.x`) · `m365-mac-mcp` (`v2.1+`)

This guide provides realistic, benchmark-grounded recommendations for pairing **Apple Silicon Unified Memory tiers (16 GB, 32 GB, 64 GB)** with the right **local LLM**, **System Prompt**, and **`M365_MCP_PROFILE`** — balancing speed, token efficiency, and minimal hallucination/loop risk.

---

## Part 1 — Apple Silicon Hardware & Model Recommendations

### Executive Comparison Matrix

| Spec / Metric | **16 GB M4** *(Entry)* | **32 GB M4 / M4 Pro** *(Sweet Spot)* | **64 GB M4 Pro / Max** *(Verified Reference)* |
| --- | --- | --- | --- |
| **Safe LLM RAM Budget** *(after macOS + Excel + Bionic)* | **~9.0 – 10.5 GB** | **~24 – 26 GB** | **~54 – 56 GB** |
| **#1 Recommended Model** | **`Qwen3.5-9B`** *(GGUF `Q4_K_M` or MLX 4/6-bit)* | **`Qwen3.6 35B A3B`** *(GGUF `Q4_K_M`)* | **`Qwen3.6 35B A3B — GGUF Q6_K` / `Q8_0`** |
| **Strong Alternatives** | **`Qwen3-14B`** *(4-bit, tight ctx)* · **`Gemma 3/4 12B`** | **`gpt-oss-20b`** *(MoE 4-bit)* · **`Devstral 24B`** | **`Qwen3.8 27B`** *(GGUF `Q5_K_M`/`Q6_K`; MLX 8-bit is accurate but slow)* |
| **Recommended MCP Profile** | **`EXCEL,COMPACT`** *(mandatory, 25 tools)* | **`EXCEL,COMPACT`** *(daily)* or `EXCEL` / `EXCEL,WORD` | **`EXCEL,COMPACT`** *(fastest)* or `FULL` *(all 82 tools)* |
| **Safe Context Length** | `16,384` *(9B)* / `12,288` *(14B)* | `32,768` | `65,536` – `131,072` |
| **KV Cache Quantization** | **Required:** `K=Q8_0, V=Q8_0` | **Recommended:** `K=Q8_0, V=Q8_0` | Optional (`FP16` or `Q8_0`) |
| **Reasoning Budget** | `1024` (`Preserve Thinking = OFF`) | `1536–2048` (`Preserve Thinking = OFF`) | `2048` (`Preserve Thinking = OFF`, `Presence Penalty = 0.2`) |
| **System Prompt** | **V3.4 + 2-Phase Chunking** | **V3.4 Full** ([`SYSTEM_PROMPT.md`](../Sample%20Model%20Config/SYSTEM_PROMPT.md)) | **V3.4 Full** ([`SYSTEM_PROMPT.md`](../Sample%20Model%20Config/SYSTEM_PROMPT.md)) |
| **Expected Decode Speed** | **~20–32 tok/s** *(9B)* / ~12–16 *(14B)* | **~20–30 tok/s** *(35B MoE)* / **35–50** *(gpt-oss-20b)* | **~22–32 tok/s** *(`Q6_K`)* / **~18–26** *(`Q8_0`)* |
| **Hallucination / Loop Risk** | **Moderate** *(mitigated by chunking + Formula Guard)* | **Low–Moderate** | **Very Low** *(`Q6_K`/`Q8_0` + MCP `cell_map` & live skill loader)* |

---

### Tier 1 — 16 GB Unified Memory (M4 or Better)
- **Memory Reality:** macOS (~3.5–4.5 GB) + Excel + Bionic + MCP (~1.5–2.5 GB) leave **~9.0–10.5 GB** for weights + KV cache. Avoid 24B–35B models (swap thrashing).
- **Best Setup:** **`Qwen3.5-9B`** (MLX 4/6-bit or `Q4_K_M`, ~5.5–7.2 GB) + `M365_MCP_PROFILE=EXCEL,COMPACT` + `Context = 16384` + `Flash Attention = ON` + `K/V = Q8_0` + `Reasoning Budget = 1024` + `Preserve Thinking = OFF`.
- **Power-User Option:** **`Qwen3-14B` (MLX 4-bit / `Q4_K_M`, ~8.8 GB)** offers stronger financial reasoning, but leaves ~1–1.5 GB for KV cache—cap context at `12288–16384` and keep heavy apps closed.
- **What to Expect:** Smooth and fast (~20–30 tok/s) on `wall-street-formatter`, `model-auditor`, `live-what-if-runner`, `capex-project-analyzer`, and `boardroom-brief-generator`. Run `three-statement-projector` and `dcf-investment-template-builder` in **2 chunked phases** (Phase 1: Build + Audit → Save; Phase 2: Format + Live Recalc Check).
- **Hallucination / Loop Profile (Moderate):** Smaller 8B–9B models occasionally slip on denominators or sign conventions; `M365_MCP_FORMULA_GUARD=LITERALS_ONLY`, `audit_workbook_formulas`, and the 25-tool `excel,compact` profile prevent silent damage and tool-name guessing.

### Tier 2 — 32 GB Unified Memory (M4 / M4 Pro or Better)
- **Memory Reality:** **~24–26 GB** available for weights + KV cache (~273 GB/s bandwidth on M4 Pro).
- **Best Setup:** **`Qwen3.6 35B A3B`** (MLX 4-bit or `Q4_K_M`, ~19.5–20.5 GB) + `M365_MCP_PROFILE=EXCEL,COMPACT` (or `EXCEL`) + `Context = 32768` + `K/V = Q8_0` + `Reasoning Budget = 1536–2048` + **V3.4 Full**.
- **Alternatives:**
  - **`gpt-oss-20b`** (21B total / 3.6B active MoE, ~12–14 GB): Fastest tool caller (~35–50 tok/s) with huge KV headroom (32K–64K context).
  - **`Devstral Small 24B`** (4-bit, ~14–15 GB): Strongest specialist for heavy Excel VBA / Python automation.
  *(Avoid `Gemma 4 26B A4B` MLX on `oMLX 0.5.0` — see Compatibility Notes below.)*
- **What to Expect:** All 9 Finance Skills run cleanly in 1 shot at ~20–30 tok/s.
- **Hallucination / Loop Profile (Low–Moderate):** Occasional 4-bit quantization slips on multi-sheet accounting logic—caught reliably by `model-auditor`, `M365_MCP_FORMULA_GUARD`, and live `Balance Check == 0` checks.

### Tier 3 — 64 GB Unified Memory (M4 Pro / M4 Max or Better) — *Verified Reference*
- **Memory Reality:** **~54–56 GB** available; zero swap pressure even with Excel, Word, PowerPoint, Bionic, and 65K–128K context open together.
- **Best Setup (Verified on Mac mini M4 Pro 64 GB):** **`Qwen3.6 35B A3B — GGUF Q6_K` (~29 GB, fastest) or `Q8_0` (~38 GB, lossless)** + `M365_MCP_PROFILE=EXCEL,COMPACT` (or `FULL`) + `Context = 65536` + `Flash Attention = ON` + `Unified + GPU KV Cache = ON` + `Reasoning Budget = 2048` + `Preserve Thinking = OFF` + `Repeat Penalty = Off` + `Presence Penalty = 0.2` + **V3.4 Full** (see [`Sample Model Config/`](../Sample%20Model%20Config/)).
- **Hallucination / Loop Profile (Very Low):** `Q6_K`/`Q8_0` MoE eliminates 4-bit drift on nested JSON (`batch_format_workbook`, `batch_write_ranges`) and multi-sheet formulas while keeping interactive MoE speed.

> **GGUF (`llama.cpp`) vs. MLX (`oMLX`) Tip:** For **multi-turn MCP tool-calling in Bionic / LM Studio**, prefer **GGUF (`llama.cpp`)** — it delivers more reliable tool-call parsing, faster multi-turn prompt-cache prefill, and unlocks **`Reasoning Budget`** caps (`1024–2048`) so models stay concise and responsive.

### Tested Local Model Compatibility Notes

The following local models were stress-tested against `Model-MCP Stress Test/STRESS_TEST_SUITE.md`:

| Model | Format / Runtime | Status | Notes / Reason |
| :--- | :--- | :---: | :--- |
| **`Gemma 4 26B A4B`** | `MLX 8bit MoE` (`oMLX 0.5.0`) | **Not Compatible** | Unstable Gemma 4 special-token / tool-call parsing under `oMLX 0.5.0`: randomly stops mid-workflow with bare `thought` (zero user feedback, requiring manual `"continue"`) and leaks raw `call:call_mcp_tool{...<\|"\|>...}` control tokens into chat text. |
| **`Nemotron 3 Nano Omni`** | `GGUF Q8_0` (`llama.cpp 2.80.1`) | **Not Compatible** | Insufficient multi-step MCP planning and instruction-following discipline: looped 15× on built-in folder existence checks, wrote tables cell-by-cell (`write_cell`), tried to open non-existent/hallucinated file paths instead of calling `create_workbook`, and ignored system-prompt restrictions against probing `.internal/skills`. |
| **`Qwen3.8 27B`** | `MLX 8bit` (`oMLX 0.5.0`, dense) | **Compatible *(Slow)*** | Passed Phase 0A-1 with 10/10 institutional accuracy (11 calls, 0 errors, exact `$1,195` live recalc), but dense 27B 8-bit on `oMLX` (no `Reasoning Budget` cap) took **16m 50s** for a 2-tab smoke test — fully compatible and accurate, just slow for interactive multi-phase workflows. |

---

## Part 2 — `M365_MCP_PROFILE` Performance, Skills & Risk Matrix

Every `M365_MCP_PROFILE` value is **case-insensitive** and composable (comma-separated). Choosing the leanest profile that covers your workflow directly reduces **Turn-1 prompt processing time (TTFT)**, **memory pressure**, and **tool-selection hallucinations**.

### Complete Profile Comparison (Measured on `m365-mac-mcp` v2.1+)

| `M365_MCP_PROFILE` | Tools | Schema Tokens | Turn-1 Static Prefix *(+ V3.4 + 9 Skills)* | Supported Finance Skills | Min Recommended RAM | Tool-Confusion / Loop Risk |
| --- | --- | --- | --- | --- | --- | --- |
| **`excel,compact`** *(Recommended)* | **25** | **~2.4K** | **~3.7K** | **All 8 Excel Skills** (`financial-statement-extractor`, `historical-3stmt-rebuilder`, `wall-street-formatter`, `model-auditor`, `live-what-if-runner`, `three-statement-projector`, `capex-project-analyzer`, `dcf-investment-template-builder`) | **16 GB+** | **Lowest** — 1 unambiguous batch tool per job |
| **`excel`** | **50** | **~4.8K** | **~6.1K** | All 8 Excel Skills + single-op/legacy tools (`format_worksheet`, `format_range`, `apply_cell_style`, `excel_live_custom_views`, `excel_view_capabilities`) | **32 GB+** | **Low–Moderate** — overlapping single vs. batch tools |
| **`compact`** | **38** | **~3.5K** | **~4.8K** | All 8 Excel Skills + Core Word (`boardroom-brief-generator` for `.docx`/`.pdf`) | **16 GB+** *(9B)* / **32 GB+** | **Low** — lean Excel + Word without PowerPoint |
| **`word`** | **26** | **~2.2K** | **~3.5K** | Word creation, `.docx` styling, `.doc` conversion, PDF/RTF/TXT export, live Word | **16 GB+** | **Very Low** |
| **`powerpoint`** | **8** | **~0.5K** | **~1.8K** | PowerPoint slide deck creation & speaker notes | **16 GB+** | **Lowest** |
| **`word,powerpoint`** | **33** | **~2.7K** | **~4.0K** | Secondary on-demand MCP entry for `boardroom-brief-generator` | **16 GB+** | **Low** |
| **`excel,word`** | **75** | **~7.0K** | **~8.3K** | All 8 Excel Skills + Full Word (`boardroom-brief-generator` `.docx`/`.pdf`) | **32 GB+** *(best on 64 GB)* | **Moderate** on <32B models; **Low** on 35B Q8 |
| **`full`** *(default if unset)* | **82** | **~7.4K** | **~8.7K** | **All 9 Finance Skills** (Excel + Word + PowerPoint) | **32 GB+** *(best on 64 GB)* | **Moderate** on 8B–14B; **Low** on 35B Q8 / Cloud |

---

### Which Profile Should You Choose?

1. **Best Setup for Almost Everyone (16 GB, 32 GB, and 64 GB) — Dual-Server Bionic Pattern:**
   - **Primary MCP Server (Always ON):** `M365_MCP_PROFILE=EXCEL,COMPACT` (**25 tools, ~2.4K tokens**). Covers 100% of Excel modeling, Wall Street formatting, charts, tables, data validation, auditing, live recalculation, and open-workbook discovery (`excel_live_list_workbooks`).
   - **Secondary MCP Server (Set to *Hide from agent unless mentioned*):** `M365_MCP_PROFILE=WORD,POWERPOINT` (**33 tools, ~2.7K tokens** — **0 tokens** until you `@`-mention it for `boardroom-brief-generator`!).
   - **Disable Shell Commands in Bionic:** Set **Command approvals → `Off` (*Don't allow shell commands*)** in the Bionic chat bar so local models cannot bypass `m365-mac-mcp` with ad-hoc Python/bash scripts.
2. **When to Use `M365_MCP_PROFILE=EXCEL` (50 tools, ~4.8K tokens):**
   - Only if you specifically need Excel custom views (`excel_live_custom_views`), named style inspection (`list_cell_styles`), or prefer single-cell/single-range tools alongside batch tools on a **32 GB / 64 GB** Mac.
3. **When to Use `M365_MCP_PROFILE=FULL` (82 tools, ~7.4K tokens):**
   - On **64 GB M4 Pro / Max** (with `Qwen3.6 35B A3B — Q8_0`) or cloud clients (Claude Desktop / Cline) when you want seamless Excel + Word + PowerPoint in a single MCP entry without `@`-mentioning a second server.
4. **Why `excel,compact` Actually *Reduces* Hallucinations vs. `excel` or `full`:**
   - Exposing 50–82 overlapping tools (`format_range` + `format_worksheet` + `batch_format_workbook`, `add_formula` + `write_cell` + `write_range` + `batch_write_ranges`) increases tool-selection entropy on local models and leads to invented hybrid names (like `batch_write_cells`).
   - `excel,compact` keeps **one canonical tool per action** (`create_workbook(sheet_names=...)`, `batch_write_ranges`, `batch_format_workbook`, `excel_live_batch_read_ranges`, `excel_live_batch_write_cells`), cutting Turn-1 latency by **~43%** while improving first-try tool accuracy.
