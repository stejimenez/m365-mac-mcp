---
name: dcf-investment-template-builder
version: 2.3.1F
description: "[v2.3.1F] Build an IFRS/FASB/SOX multi-segment DCF valuation template (CAPM WACC, dual TV, EV Bridge, Smart Magnitude headers, charts)."
---

# DCF & Valuation Template Builder (IFRS / FASB / SOX) (v2.3.1F)

## 1. Trigger & Scope
- **Triggers:** *"Build a DCF template"*, *"Multi-segment DCF."*
- **Deliverable:** 5 tabs (`Dashboard`, `Assumptions`, `Segment Cash Flows`, `DCF Valuation`, `Sensitivity`), R1–R4 headers, Smart Magnitude, `Base`/`Bull`/`Bear`, dual TV, chart, live Share Price.

## 2. Mandatory Pre-Flight / HALT Inputs
**HARD STOP (ZERO TOOLS):** If ANY driver, **currency**, or **magnitude** is missing, ask 1–3 questions with **ZERO tools**: Basis (IFRS `IFRS 16` / US GAAP `ASC 842`), Company, units, horizon, segments, CAPM WACC (`Rf`, `ERP`, `Beta`, `Kd`, `Tax`, `D/C`), `g`/Exit Mult, EV Bridge (Debt, Leases, Cash, Minority, Shares).
- **IMPOSSIBLE-MATH HARD STOP (`g >= WACC`):** If `g >= WACC`, **HARD STOP with ZERO writes** — Gordon Growth breaks when `WACC - g <= 0`.

## 3. Exact Office MCP Tool Sequence (Use Returned `cell_map`)
1. `create_workbook(title=..., sheet_names=["Dashboard", "Assumptions", "Segment Cash Flows", "DCF Valuation", "Sensitivity"], header_metadata={"entity": "<Company>", "period": "Years 1-5", "currency": "USD", "magnitude": "auto"})` → tables start **Row 6**.
2. `add_data_validation(validation_type="list", formula1="Base,Bull,Bear")` on `Dashboard!B7`; write `"Base"` to `Dashboard!B7`.
3. **1 `batch_write_ranges`/Sheet** (split `data` & `fill_formula`; use exact `cell_map` rows; **STORE ALL DOLLARS IN EXACT USD `120000000`, `35000000` AND SHARES IN EXACT COUNT `50000000` — NEVER mix `$M` with exact shares**):
   - **`Assumptions` (`A6`+):** **RAW INPUTS ONLY:** `B7` Rf, `B8` ERP, `B9` Beta, `B10` Kd, `B11` Tax, `B12` D/C; Segment Yr-1 Rev/growth/EBITDA%/D&A%/CapEx%/ΔNWC%; Bull/Bear deltas; `g`, Exit Mult; EV Bridge (Debt, IFRS 16 Leases, Cash, Minority, Shares).
   - **`Segment Cash Flows` (`A6`+):** Per-segment Rev, EBITDA, D&A, CapEx, ΔNWC linked to `Assumptions!$B$...` & `Dashboard!$B$7`, Consolidated (`=SUM(...)`).
   - **`DCF Valuation` (`A6`+):** CAPM WACC (`Ke=Assumptions!$B$7+Assumptions!$B$9*Assumptions!$B$8`, `Kd_at=Assumptions!$B$10*(1-Assumptions!$B$11)`, `WACC=(1-Assumptions!$B$12)*Ke+Assumptions!$B$12*Kd_at`), Consolidated UFCF (`R18`: `NOPAT+D&A-CapEx-ΔNWC`), **Mid-Year Disc. Factors (`t=0.5..4.5` via `=1/(1+$B$10)^(COLUMN()-1.5)`, NEVER `1.5..5.5`)**, PV of UFCF (`=SUMPRODUCT(...)`), **Dual TV** (Gordon `=F18*(1+g)/(WACC-g)` & Exit `=Yr5EBITDA*Mult`, both `/((1+WACC)^5)`), **EV→Equity Bridge** (`EV-Debt-Leases+Cash-Minority`; `Share Price=Equity/Shares`).
   - **`Sensitivity` (1 `batch_write_ranges`, 2 `fill_formula` grids matching `cell_map` rows):**
     - `A6:F11` (Gordon Price: WACC `A7:A11`, `g` `B6:F6`): `B7:F11` `"=(NPV($A7,'DCF Valuation'!$B$18:$F$18)*(1+$A7)^0.5+('DCF Valuation'!$F$18*(1+B$6)/($A7-B$6))/((1+$A7)^5)-'DCF Valuation'!$B$29)/Assumptions!$B$39"`.
     - `A14:F19` (Exit Price: WACC `A15:A19`, Mult `B14:F14`): `B15:F19` `"=(NPV($A15,'DCF Valuation'!$B$18:$F$18)*(1+$A15)^0.5+('Segment Cash Flows'!$F$32*B$14)/((1+$A15)^5)-'DCF Valuation'!$B$29)/Assumptions!$B$39"`.
   - **`Dashboard` (`A6`+):** **OUTPUTS ONLY** linked to `'DCF Valuation'!`, `add_excel_table`, **mandatory `create_chart`**.
4. `batch_format_workbook(operations=[{"auto_color_financial_model": true, "auto_fit_columns": true, "freeze_panes": "auto"}])`, `save_workbook`, `close_workbook`.

## 4. Accounting (IFRS / FASB), SOX & Guardrails
- **IFRS `IFRS 16` / FASB `ASC 842`:** Subtract IFRS 16 leases in EV Bridge; **HARD STOP on `g >= WACC`**. **SOX 404:** Exact USD + shares on `Assumptions`; zero hardcoded KPIs.

## 5. Verification & Output Gate
Run `audit_workbook_formulas` + `excel_live_open` + `excel_live_batch_read_ranges` on `Dashboard` & `Sensitivity`; output live Valuation Summary + *Saved to `<path>`.*
