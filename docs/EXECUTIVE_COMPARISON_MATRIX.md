# Executive Comparison Matrix (M4 / M4 Pro / M4 Max)

**Project:** Microsoft 365 for Mac MCP (`m365-mac-mcp`)  
**Author:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)

| Spec / Metric | **16 GB M4** *(Entry)* | **32 GB M4 / M4 Pro** *(Sweet Spot)* | **64 GB M4 Pro / Max** *(Reference / Pro)* |
| --- | --- | --- | --- |
| **Safe LLM RAM Budget** *(after macOS + Excel + Bionic)* | **~9.0 – 10.5 GB** | **~24 – 26 GB** | **~54 – 56 GB** |
| **#1 Recommended Model** | **`Qwen3.5-9B`** *(MLX 4/6-bit or `Q4_K_M`)* | **`Qwen3.6 35B A3B`** *(MLX 4-bit or `Q4_K_M`)* | **`Qwen3.6 35B A3B — Q8_0`** *(or MLX 8-bit)* |
| **Runner-Up / Alternative** | **`Qwen3-14B`** *(MLX 4-bit, tight ctx)* or **`Gemma 3/4 12B`** | **`gpt-oss-20b`** *(MoE, 4-bit)* or **`Gemma 4 26B MoE`** / **`Devstral 24B`** | **`Qwen3.5/3.6 27B Dense — Q8_0`** or **`Gemma 4 31B Dense — Q8_0`** |
| **Recommended MCP Profile** | **`EXCEL,COMPACT`** *(mandatory, 25 tools)* | **`EXCEL,COMPACT`** *(daily)* or **`EXCEL`** / **`FULL`** | **`EXCEL,COMPACT`** *(fastest)* or **`FULL`** *(all 82 tools)* |
| **Safe Context Length** | `16,384` *(9B)* / `12,288` *(14B)* | `32,768` | `65,536` – `131,072` |
| **KV Cache Quantization** | **Required:** `K=Q8_0, V=Q8_0` | **Recommended:** `K=Q8_0, V=Q8_0` | Optional (`FP16` or `Q8_0` both fit easily) |
| **Reasoning Budget** | `1024` (`Preserve Thinking = OFF`) | `1536–2048` (`Preserve Thinking = OFF`) | `2048` (`Preserve Thinking = OFF`) |
| **System Prompt** | **V3.4 Lite** *(~450t + 2-phase chunking)* | **V3.4 Full** *(~840t)* | **V3.4 Full** *(~840t)* |
| **Expected Decode Speed** | **~20–32 tok/s** *(9B)* / ~12–16 *(14B)* | **~20–30 tok/s** *(35B MoE)* / **35–50** *(gpt-oss-20b)* | **~18–26 tok/s** *(M4 Pro Q8)* / **35–50+** *(M4 Max)* |
| **Hallucination / Loop Risk** | **Moderate** *(mitigated by chunking + Formula Guard)* | **Low–Moderate** | **Very Low** *(near-lossless Q8_0)* |

---

For the full tier-by-tier breakdown and all 8 `OFFICE_MCP_PROFILE` benchmarks, see [`HARDWARE_AND_PROFILE_GUIDE.md`](HARDWARE_AND_PROFILE_GUIDE.md).
