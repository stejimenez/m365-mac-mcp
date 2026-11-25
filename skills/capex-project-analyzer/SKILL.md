---
name: capex-project-analyzer
version: 2.3.1J
description: "[v2.3.1J] IFRS/FASB/SOX CapEx workbook (NPV, IRR, MIRR, Payback, sensitivity, chart, Smart Magnitude)."
---

# CapEx Project Analyzer (IFRS / FASB / SOX) (v2.3.1J)

## 1. Scope
4 tabs (`Executive Summary`, `Assumptions`, `Project Cash Flows`, `Sensitivity & Charts`), R1–R4 headers, chart, live KPIs.

## 2. Pre-Flight HALT
**HARD STOP (ZERO TOOLS):** Missing driver, currency, or magnitude → ask 1–3 questions with **ZERO tools** (no `create_workbook`): Basis (IFRS `IAS 16` / FASB `ASC 360`), depr, CapEx, life, Rev/EBITDA, Tax, WACC, NWC, Salvage.

## 3. MCP Sequence
1. `create_workbook(title=..., sheet_names=["Executive Summary", "Assumptions", "Project Cash Flows", "Sensitivity & Charts"], header_metadata={"entity":"...","period":"Y0-Y6","currency":"USD","magnitude":"auto"})` (Row 6 start).
2. **1 `batch_write_ranges`/Sheet** (exact USD `25000000`; split `data` & `fill_formula`, e.g. `{"range":"A7:C7","data":[["Rev",0,"=Assumptions!$B$12"]]}`, `{"range":"D7:H7","fill_formula":"=C7*(1+Assumptions!$B$13)"}`):
   - **`Assumptions` (`A6:B18`):** `["Parameter","Value"]`; **RAW INPUTS ONLY:** `B7` Basis, `B8` Depr, `B9` Life, `B10` CapEx, `B11` Sunk (excl), `B12` Yr1 Rev, `B13` Growth%, `B14` EBITDA%, `B15` Tax%, `B16` WACC%, `B17` NWC%, `B18` Salvage.
   - **`Project Cash Flows` (`A6:H6`=`["Line Item",0,1,2,3,4,5,6]`; MUST write `A7:A22` labels in Col A):**
     - `R7` Revenue: `B7=0`, `C7` `=Assumptions!$B$12`, `D7:H7` `=C7*(1+Assumptions!$B$13)` | `R8` OpEx: `B8=0`, `C8:H8` `=C7*(1-Assumptions!$B$14)` | `R9` EBITDA: `B9=0`, `C9:H9` `=C7-C8`
     - `R10` Depr: `B10=0`, `C10:H10` `=IF(COLUMN()-2<=Assumptions!$B$9,Assumptions!$B$10/Assumptions!$B$9,0)` | `R11` EBIT: `B11=0`, `C11:H11` `=C9-C10` | `R12` Tax: `B12=0`, `C12:H12` `=C11*Assumptions!$B$15`
     - `R13` NOPAT: `B13=0`, `C13:H13` `=C11-C12` | `R14` +D&A: `B14=0`, `C14:H14` `=C10` | `R15` -ΔNWC: `B15=0`, `C15` `=-C7*Assumptions!$B$17`, `D15:G15` `=-(D7-C7)*Assumptions!$B$17`, `H15` `=G7*Assumptions!$B$17`
     - `R16` -CapEx: `B16` `=-Assumptions!$B$10`, `C16:H16=0` | `R17` +Salvage: `B17:G17=0`, `H17` `=Assumptions!$B$18-(Assumptions!$B$18-MAX(0,Assumptions!$B$10-SUM(C10:H10)))*Assumptions!$B$15`
     - `R18` **UFCF**: `B18` `=B16`, `C18:H18` `=C13+C14+C15+C16+C17` | `R19` Cum UFCF: `B19` `=B18`, `C19:H19` `=B19+C18` | `R20` Disc: `B20=1`, `C20:H20` `=1/(1+Assumptions!$B$16)^(COLUMN()-2)` | `R21` PV: `B21` `=B18`, `C21:H21` `=C18*C20` | `R22` Cum PV: `B22` `=B21`, `C22:H22` `=B22+C21`
   - **`Executive Summary` (`A6:B13`):** `B7` NPV `=NPV(Assumptions!$B$16,'Project Cash Flows'!C18:H18)+'Project Cash Flows'!B18`; `B8` IRR `=IRR('Project Cash Flows'!B18:H18)`; `B9` MIRR `=MIRR('Project Cash Flows'!B18:H18,Assumptions!$B$16,Assumptions!$B$16)`; `B10` PI `=(B7+Assumptions!$B$10)/Assumptions!$B$10`; `B11` Payback `=IF(MAX('Project Cash Flows'!B19:H19)<0,">6 yrs",COUNTIF('Project Cash Flows'!B19:H19,"<0")-1-INDEX('Project Cash Flows'!B19:H19,COUNTIF('Project Cash Flows'!B19:H19,"<0"))/INDEX('Project Cash Flows'!B18:H18,COUNTIF('Project Cash Flows'!B19:H19,"<0")+1))`; `B12` Disc Payback (same on `'Project Cash Flows'!B22:H22` & `'Project Cash Flows'!B21:H21`); `B13` Decision `=IF(B7>0,"Go","No-Go")`.
   - **`Sensitivity & Charts` (`A6:F11`):** `B6:F6` WACC, `A7:A11` CF Scale, `B7:F11` `fill_formula`: `=NPV(B$6,'Project Cash Flows'!$C$18:$H$18)*$A7+'Project Cash Flows'!$B$18`.
3. `create_chart(sheet_name="Sensitivity & Charts", chart_type="column", data_range="A6:F11", position="A13")`, `batch_format_workbook(operations=[{"auto_color_financial_model":true,"auto_fit_columns":true,"freeze_panes":"auto"}])`, `save_workbook`, `close_workbook`.

## 4. IFRS / FASB / SOX
**IAS 16 / ASC 360:** Exclude sunk costs/interest; tax salvage vs BV. **SOX 404:** Exact USD inputs; zero hardcoded KPIs.

## 5. Verification
`audit_workbook_formulas` + `excel_live_open` + `excel_live_batch_read_ranges` (`Executive Summary!A6:B13`); output live KPIs + *Saved to `<path>`.*
