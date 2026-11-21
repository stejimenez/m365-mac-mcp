# Finance Skills for Microsoft 365 for Mac MCP (`m365-mac-mcp`)

**Author:** Sterling Jimenez, Ph.D. (stejimenez@gmail.com)

Ready-to-use, pure-Markdown (**~2.5–3.9 KB** each) **Finance Skills** (`SKILL.md`) engineered for **Microsoft 365 for Mac MCP (`m365-mac-mcp`)** on Apple Silicon local models (**Bionic / LM Studio + `Qwen3.6 35B A3B — Q8_0`**), Claude Desktop, and Cline.

Every skill follows the **5-block low-token blueprint** (*Trigger & Scope → Pre-Flight / HALT → Exact Tool Sequence → IFRS / FASB / SOX & Formula Guardrails → Verification Gate*) and enforces:
- **IFRS & FASB (US GAAP) alignment** (`IAS 1`/`ASC 210`, `IAS 7`/`ASC 230`, `IAS 12`/`ASC 740`, `IAS 16`/`ASC 360`, `IFRS 16`/`ASC 842`).
- **SOX Section 404 spreadsheet controls** (zero hardcoded calculated cells, zero silent plugs, audit traceability, and strict `OFFICE_MCP_FORMULA_GUARD` compliance).

---

## Available Skills

| Skill | Folder | Plain-English Example |
| --- | --- | --- |
| **Wall Street Formatter** | [`wall-street-formatter/SKILL.md`](wall-street-formatter/SKILL.md) | *"Format this 3-statement workbook Wall Street style."* |
| **Model Auditor** | [`model-auditor/SKILL.md`](model-auditor/SKILL.md) | *"Audit this model for formula errors, SOX breaks, and IFRS/GAAP issues."* |
| **Live What-If Runner** | [`live-what-if-runner/SKILL.md`](live-what-if-runner/SKILL.md) | *"Test Revenue ±10% and WACC ±100 bps live in Excel without saving changes."* |
| **3-Statement Projector** | [`three-statement-projector/SKILL.md`](three-statement-projector/SKILL.md) | *"Project FY26–FY30 with Base/Bull/Bear scenarios and live Balance Check = 0."* |
| **CapEx Project Analyzer** | [`capex-project-analyzer/SKILL.md`](capex-project-analyzer/SKILL.md) | *"Evaluate a \$20M plant expansion over 7 years at 10% WACC with NPV, IRR, and MIRR."* |
| **DCF & Investment Template Builder** | [`dcf-investment-template-builder/SKILL.md`](dcf-investment-template-builder/SKILL.md) | *"Build a reusable multi-segment DCF valuation template with WACC, Terminal Value, and Equity Bridge."* |
| **Boardroom Brief Generator** | [`boardroom-brief-generator/SKILL.md`](boardroom-brief-generator/SKILL.md) | *"Turn this live Excel model into a styled 1-page CFO Word/PDF memo and 3-slide PowerPoint deck."* |

---

## How to Use in Bionic / LM Studio

1. Open **Bionic → Settings → Skills** and import/point to the `skills/` folder (or individual `SKILL.md` files).
2. In your System Prompt (`Sample Model Config/SYSTEM_PROMPT.md`), narrow the Office skill rule so only Bionic's built-in generic Office skill is blocked:
   > *"NEVER use Bionic's built-in Open/viewer tool and NEVER read `.internal/skills/.../SKILL.md` for Office documents. Use `m365-mac-mcp` tools and your custom `skills/` playbooks directly."*
3. Keep `OFFICE_MCP_PROFILE=EXCEL` (or `EXCEL,WORD` / `FULL` if using `boardroom-brief-generator`), `OFFICE_MCP_LIVE_SYNC=AUTO`, and `OFFICE_MCP_FORMULA_GUARD=LITERALS_ONLY`.
