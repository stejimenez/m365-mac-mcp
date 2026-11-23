# Executive Comparison Matrix (M4 / M4 Pro / M4 Max)

**Project:** Microsoft 365 for Mac MCP (`m365-mac-mcp`)  
**Author:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)

| Spec / Metric | **16 GB M4** *(Entry)* | **32 GB M4 / M4 Pro** *(Sweet Spot)* | **64 GB M4 Pro / Max** *(Reference / Pro)* |
| --- | --- | --- | --- |
| **Safe LLM RAM Budget** *(after macOS + Excel + Bionic)* | **~9.0 – 10.5 GB** | **~24 – 26 GB** | **~54 – 56 GB** |
| **#1 Recommended Model** | **`Qwen3.5-9B`** *(GGUF `Q4_K_M` or MLX 4/6-bit)* | **`Qwen3.6 35B A3B`** *(GGUF `Q4_K_M`)* | **`Qwen3.6 35B A3B — GGUF Q6_K` / `Q8_0`** |
| **Runner-Up / Alternative** | **`Qwen3-14B`** *(4-bit, tight ctx)* or **`Gemma 3/4 12B`** | **`gpt-oss-20b`** *(MoE, 4-bit)* or **`Devstral 24B`** | **`Qwen3.8 27B`** *(GGUF `Q5_K_M`/`Q6_K`; MLX 8-bit is slow)* |
| **Recommended MCP Profile** | **`EXCEL,COMPACT`** *(mandatory, 25 tools)* | **`EXCEL,COMPACT`** *(daily)* or **`EXCEL`** / **`FULL`** | **`EXCEL,COMPACT`** *(fastest)* or **`FULL`** *(all 82 tools)* |
| **Safe Context Length** | `16,384` *(9B)* / `12,288` *(14B)* | `32,768` | `65,536` – `131,072` |
| **KV Cache Quantization** | **Required:** `K=Q8_0, V=Q8_0` | **Recommended:** `K=Q8_0, V=Q8_0` | Optional (`FP16` or `Q8_0` both fit easily) |
| **Reasoning Budget** | `1024` (`Preserve Thinking = ON`) | `1536–2048` (`Preserve Thinking = ON`) | `2048` (`Preserve Thinking = ON`, `Presence Penalty = 0.2`) |
| **System Prompt** | **V3.4 Lite** *(~450t + 2-phase chunking)* | **V3.4 Full** *(~840t)* | **V3.4 Full** *(~840t)* |
| **Expected Decode Speed** | **~20–32 tok/s** *(9B)* / ~12–16 *(14B)* | **~20–30 tok/s** *(35B MoE)* / **35–50** *(gpt-oss-20b)* | **~22–32 tok/s** *(`Q6_K`)* / **~18–26** *(`Q8_0`)* |
| **Hallucination / Loop Risk** | **Moderate** *(mitigated by chunking + Formula Guard)* | **Low–Moderate** | **Very Low** *(`Q6_K`/`Q8_0` + `Preserve Thinking = ON`)* |

---

### Tested Local Model Compatibility Notes

| Model | Format / Runtime | Status | Notes / Reason |
| :--- | :--- | :---: | :--- |
| **`Gemma 4 26B A4B`** | `MLX 8bit MoE` (`oMLX 0.5.0`) | **Not Compatible** | `oMLX 0.5.0` tool-call parser bug: random bare-`thought` mid-workflow stops (forcing manual `"continue"`) and leaked raw `call:call_mcp_tool` control tokens in chat. |
| **`Nemotron 3 Nano Omni`** | `GGUF Q8_0` (`llama.cpp`) | **Not Compatible** | Weak multi-step MCP planning: 15× folder `exists` loop, cell-by-cell `write_cell` writes, hallucinated paths instead of `create_workbook`, and `.internal/skills` ban violation. |
| **`Qwen3.8 27B`** | `MLX 8bit` (`oMLX`, dense) | **Compatible *(Slow)*** | Passed 0A-1 with 10/10 accuracy (11 calls, 0 errors, exact `$1,195` live), but dense 27B 8-bit on `oMLX` (no `Reasoning Budget` cap) took **16m 50s** for a 2-tab smoke test — accurate, just slow for interactive use. |

---

For the full tier-by-tier breakdown and all 8 `M365_MCP_PROFILE` benchmarks, see [`HARDWARE_AND_PROFILE_GUIDE.md`](HARDWARE_AND_PROFILE_GUIDE.md).
