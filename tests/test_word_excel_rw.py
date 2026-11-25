#!/usr/bin/env python3
"""
Comprehensive automated test suite for Word & Excel Read-Write workflows,
JSON serialization safety (Bionic 1.2.1+10 / Antigravity), stdio cleanliness,
and AppleScriptBridge hardening.
"""

import asyncio
import io
import json
import sys
from contextlib import redirect_stdout
from pathlib import Path

import pytest

SRC_DIR = str(Path(__file__).resolve().parent.parent / "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

import m365_mac_mcp_server as server
from controllers.excel_controller import ExcelController
from controllers.word_controller import WordController
from integrations.applescript_bridge import AppleScriptBridge, AppleScriptTimeoutError
from utils.logger import setup_logger
from utils.validators import sanitize_for_json


def _assert_strict_json(obj) -> None:
    """Verify object serializes to strict JSON (no NaN/Inf/non-serializable objects)."""
    encoded = json.dumps(obj, allow_nan=False)
    assert isinstance(encoded, str)


def test_logger_never_writes_to_stdout():
    """MCP stdio transport requires stdout to remain completely free of log messages."""
    buf = io.StringIO()
    with redirect_stdout(buf):
        log = setup_logger("test_stdout_isolation")
        log.info("This must go to stderr, never stdout")
        log.error("Error message also goes to stderr")
    assert buf.getvalue() == ""


@pytest.mark.asyncio
async def test_check_office_status_non_intrusive():
    """check_office_status should detect installed Word & Excel without errors."""
    status = await server.check_office_status()
    _assert_strict_json(status)
    assert status["status"] == "running"
    assert isinstance(status["word_available"], bool)
    assert isinstance(status["excel_available"], bool)
    assert "applications" in status
    assert "word" in status["applications"]
    assert "excel" in status["applications"]


@pytest.mark.asyncio
async def test_word_full_read_write_roundtrip(tmp_path: Path):
    """Full Word .docx create -> edit -> save (with spaces/special chars) -> close -> reopen -> read."""
    wc = WordController()

    created = await wc.create_document(title="Q4 Executive Brief", add_title_heading=True)
    _assert_strict_json(created)
    doc_id = created["document_id"]

    await wc.add_heading(doc_id, "Financial Overview", level=2)
    await wc.add_paragraph(
        doc_id,
        "Revenue grew by 25% in Q4 with ACME Corp.",
        formatting={
            "bold": True,
            "italic": False,
            "underline": True,
            "font_size": 12,
            "font_name": "Calibri",
            "color": "#1F4E79",
            "alignment": "left",
            "space_after": 6,
        },
    )
    await wc.insert_paragraph(doc_id, index=1, text="Inserted executive callout.")
    await wc.add_list(doc_id, ["Item One", "Item Two"], list_type="bullet")
    await wc.add_table(
        doc_id,
        rows=2,
        columns=3,
        data=[["Metric", "Q3", "Q4"], ["ACME Revenue", "$100k", "$125k"]],
        style="Table Grid",
    )
    await wc.update_table_cell(doc_id, table_index=0, row=1, column=2, text="$130k")

    rep = await wc.replace_text(doc_id, "ACME", "Contoso", match_case=True, include_tables=True)
    assert rep["replacements"] == 2

    await wc.set_document_properties(
        doc_id, author="Antigravity", subject="Quarterly Review", comments="Verified"
    )

    out_file = tmp_path / "Q4 Report (Final) & Notes.docx"
    saved = await wc.save_document(doc_id, str(out_file), format="docx")
    assert Path(saved["file_path"]).exists()

    closed = await wc.close_document(doc_id)
    assert closed["status"] == "success"

    # Reopen from disk and verify every element
    reopened = await wc.open_document(str(out_file))
    doc_id2 = reopened["document_id"]
    content = await wc.read_document(doc_id2, include_tables=True, include_styles=True, include_metadata=True)
    _assert_strict_json(content)

    # Default read_document omits duplicated full_text/headings and default 'Normal' style
    assert "full_text" not in content
    assert "headings" not in content
    joined_text = "\n".join(p["text"] for p in content["paragraphs"])
    assert "Q4 Executive Brief" in joined_text
    assert "Inserted executive callout." in joined_text
    assert "Contoso Corp." in joined_text
    assert "ACME" not in joined_text
    assert content["paragraphs"][0].get("style") == "Heading 1"
    assert "style" not in content["paragraphs"][1]
    assert content["table_count"] == 1
    assert content["tables"][0]["data"][1][0] == "Contoso Revenue"
    assert content["tables"][0]["data"][1][2] == "$130k"
    assert content["properties"]["author"] == "Antigravity"
    assert len(content["available_styles"]) > 0

    # Explicit include_full_text=True still returns full_text on demand
    content_full = await wc.read_document(doc_id2, include_tables=False, include_metadata=False, include_full_text=True)
    assert "Q4 Executive Brief" in content_full["full_text"]

    # Delete inserted paragraph and save in-place
    del_res = await wc.delete_paragraph(doc_id2, 1)
    assert del_res["deleted_text"] == "Inserted executive callout."
    await wc.save_document(doc_id2)
    await wc.close_document(doc_id2)


@pytest.mark.asyncio
async def test_excel_full_read_write_roundtrip_and_types(tmp_path: Path):
    """Full Excel .xlsx & .csv create -> write typed values/formulas/table -> save -> reopen -> read."""
    ec = ExcelController()

    created = await ec.create_workbook(title="Sales Model", sheet_name="Summary")
    _assert_strict_json(created)
    wb_id = created["workbook_id"]

    await ec.add_worksheet(wb_id, "Raw Data")
    await ec.rename_worksheet(wb_id, "Raw Data", "Sales")
    await ec.set_active_worksheet(wb_id, "Sales")

    rows = [
        ["Product", "Q1", "Q2", "Active", "ZIP"],
        ["Laptops", 100, 150.5, True, "01234"],
        ["Tablets", 80, 90.25, False, "02139"],
    ]
    wr = await ec.write_range(
        wb_id,
        "Sales",
        "A1",
        rows,
        header_formatting={"bold": True, "bg_color": "#D9E1F2"},
    )
    assert wr["range"] == "A1:E3"

    # Test NaN/Inf/None sanitization and number formatting
    nan_res = await ec.write_cell(wb_id, "Sales", "F1", float("nan"))
    assert nan_res["value"] is None
    _assert_strict_json(nan_res)

    fml = await ec.add_formula(
        wb_id, "Sales", "F2", "SUM(B2:C2)", number_format="$#,##0.00"
    )
    assert fml["formula"] == "=SUM(B2:C2)"

    tbl = await ec.add_table(wb_id, "Sales", "Sales Table 1", "A1:E3")
    assert tbl["table_name"] == "Sales_Table_1"

    await ec.create_chart(wb_id, "Sales", "bar", "A1:C3", chart_title="H1 Sales", position="H2")

    out_xlsx = tmp_path / "Sales Model (2026) & Forecast.xlsx"
    saved = await ec.save_workbook(wb_id, str(out_xlsx), format="xlsx")
    assert Path(saved["file_path"]).exists()

    out_csv = tmp_path / "sales_export.csv"
    csv_saved = await ec.save_workbook(wb_id, str(out_csv), format="csv", sheet_name="Sales")
    assert Path(csv_saved["file_path"]).exists()
    assert csv_saved["rows_exported"] == 3

    await ec.close_workbook(wb_id)

    # Reopen .xlsx and verify types, formula, and table
    reopened = await ec.open_workbook(str(out_xlsx))
    wb_id2 = reopened["workbook_id"]
    sheets = await ec.list_worksheets(wb_id2)
    assert sheets == ["Summary", "Sales"]

    cell_bool = await ec.read_cell(wb_id2, "Sales", "D2")
    assert cell_bool["value"] is True

    cell_fml = await ec.read_cell(wb_id2, "Sales", "F2")
    assert cell_fml["is_formula"] is True
    assert cell_fml["formula"] == "=SUM(B2:C2)"
    assert cell_fml["number_format"] == "$#,##0.00"

    rng = await ec.read_range(wb_id2, "Sales", "A1:E3")
    _assert_strict_json(rng)
    assert rng["data"][1][4] == "01234"
    assert rng["data"][2][3] is False

    tables = await ec.read_tables(wb_id2, "Sales")
    _assert_strict_json(tables)
    assert tables["table_count"] == 1
    assert tables["tables"][0]["table_name"] == "Sales_Table_1"
    await ec.close_workbook(wb_id2)

    # Reopen .csv and verify ZIP leading zero + bool/int parsing
    csv_opened = await ec.open_workbook(str(out_csv))
    csv_id = csv_opened["workbook_id"]
    csv_rng = await ec.read_range(csv_id, None, "A1:E3")
    _assert_strict_json(csv_rng)
    assert csv_rng["data"][1][1] == 100
    assert csv_rng["data"][1][3] is True
    assert csv_rng["data"][1][4] == "01234"
    await ec.close_workbook(csv_id)


@pytest.mark.asyncio
async def test_applescript_bridge_argv_safety_and_timeout():
    """AppleScriptBridge must pass special chars via argv safely and enforce timeouts."""
    bridge = AppleScriptBridge(default_timeout=5.0)

    # Verify quotes, backslashes, and spaces roundtrip without injection
    tricky = 'Hello "World" \\ path/with spaces & $pecial!'
    script = """
    on run argv
        return item 1 of argv
    end run
    """
    out = await bridge.execute_applescript(script, args=[tricky], timeout=5.0)
    assert out == tricky

    # Verify timeout kills hanging osascript cleanly
    hang_script = "delay 10"
    with pytest.raises(AppleScriptTimeoutError):
        await bridge.execute_applescript(hang_script, timeout=0.5)


@pytest.mark.asyncio
async def test_word_live_list_documents_not_running_and_empty(monkeypatch):
    """word_live_list_documents returns a clean empty list when Word is not running or has no docs."""
    called = False

    async def fake_is_running(app_name: str) -> bool:
        return False

    async def fake_execute(script: str, args=None, timeout=None) -> str:
        nonlocal called
        called = True
        return ""

    monkeypatch.setattr(server.applescript, "is_app_running", fake_is_running)
    monkeypatch.setattr(server.applescript, "execute_applescript", fake_execute)

    res = await server.word_live_list_documents()
    _assert_strict_json(res)
    assert res == {"status": "success", "count": 0, "documents": []}
    assert called is False

    # Now simulate Word running with zero open documents
    async def fake_is_running_true(app_name: str) -> bool:
        return True

    monkeypatch.setattr(server.applescript, "is_app_running", fake_is_running_true)

    res_empty = await server.word_live_list_documents()
    _assert_strict_json(res_empty)
    assert res_empty == {"status": "success", "count": 0, "documents": []}
    assert called is True


@pytest.mark.asyncio
async def test_word_live_list_documents_script_and_parsing(monkeypatch):
    """list_open_word_documents avoids `repeat with ... in documents`, compiles cleanly, and parses output."""
    captured = {}

    async def fake_is_running(app_name: str) -> bool:
        assert app_name == "Microsoft Word"
        return True

    async def fake_execute(script: str, args=None, timeout=None) -> str:
        captured["script"] = script
        captured["timeout"] = timeout
        us = "\x1f"
        rs = "\x1e"
        row1 = us.join([
            "Q4 Report, Draft.docx",
            "Macintosh HD:Users:test:Q4 Report, Draft.docx",
            "Macintosh HD:Users:test",
            "true",
            "42",
            "true",
        ])
        # Duplicate row (e.g., split window) should be deduplicated cleanly
        row1_dup = row1
        row2 = us.join(["Document1", "", "", "false", "1", "false"])
        return rs.join([row1, row1_dup, row2])

    monkeypatch.setattr(server.applescript, "is_app_running", fake_is_running)
    monkeypatch.setattr(server.applescript, "execute_applescript", fake_execute)

    res = await server.word_live_list_documents()
    _assert_strict_json(res)
    assert res["status"] == "success"
    assert res["count"] == 2
    assert res["documents"] == [
        {
            "name": "Q4 Report, Draft.docx",
            "full_name": "Macintosh HD:Users:test:Q4 Report, Draft.docx",
            "path": "Macintosh HD:Users:test",
            "saved": True,
            "paragraph_count": 42,
            "is_active": True,
        },
        {
            "name": "Document1",
            "full_name": "",
            "path": "",
            "saved": False,
            "paragraph_count": 1,
            "is_active": False,
        },
    ]

    # Regression guard against `every document doesn't understand the "count" message` and empty count=0
    script = captured["script"]
    assert "repeat with d in documents" not in script
    assert "length of docNames" not in script
    assert "count of windows" in script
    assert "document of window i" in script
    assert "count of documents" in script
    assert "document i" in script
    assert "name of every document" in script
    assert "active document" in script
    assert captured["timeout"] == 15.0

    # Verify the AppleScript compiles cleanly with osacompile on macOS (skip if sandbox blocks LaunchServices XPC)
    proc = await asyncio.create_subprocess_exec(
        "osacompile",
        "-o",
        "/dev/null",
        stdin=asyncio.subprocess.PIPE,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    _, stderr = await proc.communicate(input=script.encode("utf-8"))
    stderr_text = stderr.decode("utf-8", errors="replace")
    if "hiservices-xpcservice" not in stderr_text:
        assert proc.returncode == 0, f"osacompile failed: {stderr_text}"


@pytest.mark.asyncio
async def test_excel_and_word_live_save_applescript_syntax(monkeypatch, tmp_path: Path):
    """Verify excel_live_save and word_live_save use valid Excel.sdef / Word.sdef AppleScript syntax."""
    recorded = []

    async def fake_execute(script: str, args=None, timeout=None) -> str:
        recorded.append({"script": script, "args": list(args or []), "timeout": timeout})
        return "TestDoc"

    monkeypatch.setattr(server.applescript, "execute_applescript", fake_execute)

    # 1. excel_live_save in-place and with file_path
    out_xlsx = tmp_path / "Save Test.xlsx"
    res_xl_inplace = await server.excel_live_save()
    _assert_strict_json(res_xl_inplace)
    assert res_xl_inplace["status"] == "success"

    res_xl_path = await server.excel_live_save(file_path=str(out_xlsx), file_format="xlsx")
    _assert_strict_json(res_xl_path)
    assert res_xl_path["status"] == "success"

    xl_script = recorded[-1]["script"]
    # Must NOT contain invalid Excel.sdef syntax
    assert "overwrite yes" not in xl_script
    assert "macro enabled workbook file format" not in xl_script
    assert "save active sheet of wb as PDF" not in xl_script
    # Must contain valid Excel.sdef XlFileFormat enumerators and POSIX file outside tell block
    assert "macro enabled XML file format" in xl_script
    assert "CSV file format" in xl_script
    assert "PDF file format" in xl_script
    assert "workbook normal file format" in xl_script
    assert xl_script.index("POSIX file posixOut") < xl_script.index('tell application "Microsoft Excel"')

    # 2. word_live_save in-place and with file_path
    out_docx = tmp_path / "Save Test.docx"
    res_wd_inplace = await server.word_live_save()
    _assert_strict_json(res_wd_inplace)
    assert res_wd_inplace["status"] == "success"

    res_wd_path = await server.word_live_save(file_path=str(out_docx), file_format="pdf")
    _assert_strict_json(res_wd_path)
    assert res_wd_path["status"] == "success"

    wd_script = recorded[-1]["script"]
    assert wd_script.index("POSIX file posixOut") < wd_script.index('tell application "Microsoft Word"')
    assert "if (count of documents) is 0" not in wd_script
    assert "save as targetDoc file name hfsPath file format format PDF" in wd_script
    assert "save as targetDoc file name hfsPath file format format document" in wd_script


def test_live_excel_formula_sheet_ref_sanitizer():
    """Verify live Excel formula sanitizer auto-quotes spaced sheet names and blocks invalid refs."""
    sheets = ["Income Statement", "Balance Sheet", "Cash Flow"]
    assert (
        AppleScriptBridge._sanitize_live_excel_formula("=B2 + Cash Flow!C15", sheets)
        == "=B2 + 'Cash Flow'!C15"
    )
    assert (
        AppleScriptBridge._sanitize_live_excel_formula("='Cash Flow'!F15", sheets)
        == "='Cash Flow'!F15"
    )
    with pytest.raises(ValueError, match="unknown/external sheet"):
        AppleScriptBridge._sanitize_live_excel_formula("=Cash_Flow!F15", sheets)


@pytest.mark.asyncio
async def test_excel_live_batch_write_and_read(monkeypatch):
    """Verify excel_live_batch_write_cells and excel_live_batch_read_ranges payload encoding and parsing."""
    captured = []

    async def fake_list_sheets(workbook_name=None):
        return {"status": "success", "worksheets": ["Income Statement", "Balance Sheet", "Cash Flow"]}

    async def fake_execute(script: str, args=None, timeout=None) -> str:
        captured.append({"script": script, "args": list(args or []), "timeout": timeout})
        us, rs, gs, fs = "\x1f", "\x1e", "\x1d", "\x1c"
        if "matrix_to_text" in script:
            blk1 = f"Income Statement{us}B2:C2{gs}365817{us}394328{gs}"
            blk2 = f"Balance Sheet{us}F23:F24{gs}11433{rs}500{gs}"
            return f"AAPL.xlsx{fs}{blk1}{fs}{blk2}"
        r1 = us.join(["Income Statement", "G2", "0.0328"])
        r2 = us.join(["Income Statement", "G11", "N/A"])
        return f"AAPL.xlsx{gs}{r1}{rs}{r2}"

    monkeypatch.setattr(server.applescript, "list_live_excel_worksheets", fake_list_sheets)
    monkeypatch.setattr(server.applescript, "execute_applescript", fake_execute)

    bw = await server.excel_live_batch_write_cells(
        cells=[
            {"cell": "G2", "value": "=Cash Flow!F15/B2", "number_format": "0.0%"},
            {"cell": "G11", "value": "N/A"},
        ],
        sheet_name="Income Statement",
        recalculate=True,
    )
    _assert_strict_json(bw)
    assert bw["status"] == "success"
    assert bw["cells_written"] == 2
    assert bw["results"][0] == {"sheet_name": "Income Statement", "cell": "G2", "evaluated_value": "0.0328"}
    # Auto-quoted cross-sheet formula inside payload
    assert "='Cash Flow'!F15/B2" in captured[0]["args"][1]

    br = await server.excel_live_batch_read_ranges(
        ranges=[
            {"sheet_name": "Income Statement", "range_ref": "B2:C2"},
            {"sheet_name": "Balance Sheet", "range_ref": "F23:F24"},
        ]
    )
    _assert_strict_json(br)
    assert br["status"] == "success"
    assert br["ranges_read"] == 2
    assert br["results"][0]["values"] == [["365817", "394328"]]
    assert br["results"][1]["values"] == [["11433"], ["500"]]


@pytest.mark.asyncio
async def test_excel_formatting_column_widths_borders_and_no_external_links(tmp_path: Path):
    """Verify explicit column widths, fills, fonts, per-side borders, number formats, freeze panes, and zero external links."""
    import zipfile
    from openpyxl import load_workbook

    buf = io.StringIO()
    with redirect_stdout(buf):
        created = await server.create_workbook(title="Financial Model", sheet_name="Balance Sheet")
        _assert_strict_json(created)
        wb_id = created["workbook_id"]

        await server.add_worksheet(wb_id, "Income Statement")
        await server.add_worksheet(wb_id, "Summary")

        await server.write_range(
            wb_id,
            "Balance Sheet",
            "A1",
            [
                ["Line Item", "FY2024", "FY2025", "FY2026"],
                ["Cash & Equivalents", 1250000, 1500000, 1850000],
                ["Total Assets", 5200000, 6100000, 7250000],
            ],
        )
        await server.write_range(
            wb_id,
            "Income Statement",
            "A1",
            [
                ["Line Item", "FY2024", "FY2025", "FY2026"],
                ["Revenue", 3400000, 4100000, 4950000],
                ["EBITDA Margin", 0.245, 0.268, 0.291],
            ],
        )

        # Cross-sheet formulas: both pre-quoted and unquoted with spaces
        await server.write_range(
            wb_id,
            "Summary",
            "A1",
            [
                ["Metric", "FY2024", "FY2025", "FY2026"],
                ["Cash (Pre-quoted)", "='Balance Sheet'!B2", "='Balance Sheet'!C2", "='Balance Sheet'!D2"],
                ["Revenue (Auto-quoted)", "=Income Statement!B2", "=Income Statement!C2", "=Income Statement!D2"],
                ["EBITDA Margin", "='Income Statement'!B3", "='Income Statement'!C3", "='Income Statement'!D3"],
            ],
        )
        fml_res = await server.add_formula(
            wb_id,
            "Summary",
            "B5",
            "=Balance Sheet!B2 + Income Statement!B2",
            number_format='$#,##0;($#,##0);"-"',
            formatting={"bold": True, "font_color": "#008000"},
        )
        _assert_strict_json(fml_res)
        assert fml_res["formula"] == "='Balance Sheet'!B2 + 'Income Statement'!B2"

        # One-shot batch worksheet formatting
        fmt_res = await server.format_worksheet(
            wb_id,
            "Summary",
            column_widths={"A": 42, "B:D": 16},
            row_heights={"1": 28, "2:5": 20},
            freeze_panes="B2",
            auto_fit_columns=True,
            show_gridlines=True,
            styles=[
                {
                    "range": "A1:D1",
                    "formatting": {
                        "font_name": "Calibri",
                        "font_size": 11,
                        "bold": True,
                        "font_color": "#FFFFFF",
                        "bg_color": "#1F4E79",
                        "horizontal": "center",
                        "vertical": "center",
                        "wrap_text": True,
                    },
                },
                {
                    "range": "B2:D3",
                    "number_format": '$#,##0;($#,##0);"-"',
                    "formatting": {"font_color": "#0000FF", "horizontal": "right"},
                },
                {
                    "range": "B4:D4",
                    "number_format": "0.0%",
                    "formatting": {"italic": True, "horizontal": "right"},
                },
                {
                    "range": "A5:D5",
                    "formatting": {
                        "bold": True,
                        "bg_color": "#D9E1F2",
                        "border": {"top": "thin", "bottom": "double", "color": "#1F4E79"},
                    },
                },
            ],
        )
        _assert_strict_json(fmt_res)
        assert fmt_res["status"] == "success"
        assert fmt_res["freeze_panes"] == "B2"
        assert fmt_res["column_widths"]["A"] == 42.0
        assert fmt_res["column_widths"]["B"] == 16.0
        assert fmt_res["column_widths"]["D"] == 16.0

        # Verify standalone set_column_widths & format_range tools too
        cw_res = await server.set_column_widths(
            wb_id, "Balance Sheet", widths={"A": 36}, auto_fit=True, min_width=14.0
        )
        _assert_strict_json(cw_res)
        assert cw_res["column_widths"]["A"] == 36.0
        assert cw_res["column_widths"]["B"] >= 14.0

        fr_res = await server.format_range(
            wb_id,
            "Balance Sheet",
            "B2:D3",
            formatting={"font_color": "#000000"},
            number_format='$#,##0;($#,##0);"-"',
        )
        _assert_strict_json(fr_res)
        assert fr_res["cells_formatted"] == 6

        out_path = tmp_path / "Formatted_Financial_Model.xlsx"
        saved = await server.save_workbook(wb_id, str(out_path), format="xlsx")
        _assert_strict_json(saved)
        await server.close_workbook(wb_id)

    assert buf.getvalue() == ""

    # Reload with openpyxl and inspect persisted styles, widths, formulas, and external links
    wb_reloaded = load_workbook(str(out_path))
    assert len(wb_reloaded._external_links) == 0
    ws = wb_reloaded["Summary"]
    assert ws.freeze_panes == "B2"
    assert ws.column_dimensions["A"].width == 42.0
    assert ws.column_dimensions["B"].width == 16.0
    assert ws.column_dimensions["C"].width == 16.0
    assert ws.column_dimensions["D"].width == 16.0
    assert ws.row_dimensions[1].height == 28.0
    assert ws["A1"].font.bold is True
    assert ws["A1"].font.color.rgb.endswith("FFFFFF")
    assert ws["A1"].fill.fill_type == "solid"
    assert ws["A1"].fill.start_color.rgb.endswith("1F4E79")
    assert ws["B2"].number_format == '$#,##0;($#,##0);"-"'
    assert ws["B2"].font.color.rgb.endswith("0000FF")
    assert ws["B4"].number_format == "0.0%"
    assert ws["B5"].value == "='Balance Sheet'!B2 + 'Income Statement'!B2"
    assert ws["B5"].font.bold is True
    assert ws["B5"].font.color.rgb.endswith("008000")
    assert ws["B5"].fill.start_color.rgb.endswith("D9E1F2")
    assert ws["B5"].border.top.style == "thin"
    assert ws["B5"].border.bottom.style == "double"
    assert ws["B3"].value == "='Income Statement'!B2"
    wb_reloaded.close()

    # Verify zero externalLinks XML parts in the .xlsx ZIP archive
    with zipfile.ZipFile(out_path, "r") as zf:
        ext_parts = [n for n in zf.namelist() if "externalLink" in n]
        assert ext_parts == []


@pytest.mark.asyncio
async def test_word_explicit_formatting_and_table_styling(tmp_path: Path):
    """Verify Word heading/paragraph formatting, table shading/borders/padding/widths, and reload."""
    from docx import Document as DocxDocument
    from docx.oxml.ns import qn

    buf = io.StringIO()
    with redirect_stdout(buf):
        created = await server.create_document(title="Executive Memo")
        _assert_strict_json(created)
        doc_id = created["document_id"]

        h1 = await server.add_heading(
            doc_id,
            "FY2026 Financial Summary",
            level=1,
            formatting={
                "font_name": "Calibri",
                "font_size": 16,
                "bold": True,
                "color": "#1F4E79",
                "space_after": 8,
                "border_bottom": {"style": "single", "color": "#1F4E79", "size": 6},
            },
        )
        _assert_strict_json(h1)

        p1 = await server.add_paragraph(
            doc_id,
            "Key results exceeded guidance across all segments.",
            formatting={
                "font_name": "Calibri",
                "font_size": 11,
                "font_color": "#222222",
                "line_spacing": 1.15,
                "space_after": 6,
                "bg_color": "#F2F5F9",
            },
        )
        _assert_strict_json(p1)

        fp = await server.format_paragraph(
            doc_id,
            index=p1["paragraph_index"],
            formatting={"italic": True, "keep_with_next": True},
        )
        _assert_strict_json(fp)

        tbl = await server.add_table(
            doc_id,
            rows=3,
            columns=3,
            data=[
                ["Metric", "FY2025", "FY2026"],
                ["Revenue", "$4,100,000", "$4,950,000"],
                ["EBITDA Margin", "26.8%", "29.1%"],
            ],
            style="Table Grid",
            column_widths=[2.8, 1.6, 1.6],
            header_formatting={
                "bg_color": "#1F4E79",
                "font_color": "#FFFFFF",
                "bold": True,
                "font_name": "Calibri",
                "font_size": 10,
                "vertical_alignment": "center",
            },
            body_formatting={
                "font_name": "Calibri",
                "font_size": 10,
                "font_color": "#000000",
            },
            alternating_row_color="#D9E1F2",
            borders={"style": "single", "color": "#BFBFBF", "size": 4},
            cell_padding={"top": 4, "bottom": 4, "left": 6, "right": 6},
            repeat_header=True,
            table_alignment="center",
            column_alignments=["left", "right", "right"],
        )
        _assert_strict_json(tbl)
        assert tbl["status"] == "success"

        ft = await server.format_table(
            doc_id,
            table_index=0,
            cell_padding={"top": 5, "bottom": 5, "left": 8, "right": 8},
        )
        _assert_strict_json(ft)

        out_docx = tmp_path / "Styled_Executive_Memo.docx"
        saved = await server.save_document(doc_id, str(out_docx), format="docx")
        _assert_strict_json(saved)
        await server.close_document(doc_id)

    assert buf.getvalue() == ""

    # Reload with python-docx and verify persisted XML and formatting
    reloaded = DocxDocument(str(out_docx))
    assert len(reloaded.paragraphs) >= 2
    assert reloaded.paragraphs[0].runs[0].font.color.rgb == (0x1F, 0x4E, 0x79)
    assert reloaded.paragraphs[0]._p.pPr.find(qn("w:pBdr")) is not None
    assert reloaded.paragraphs[1].runs[0].font.italic is True
    assert reloaded.paragraphs[1]._p.pPr.find(qn("w:shd")).get(qn("w:fill")) == "F2F5F9"

    assert len(reloaded.tables) == 1
    t = reloaded.tables[0]
    assert t.autofit is False
    # Header row shading, white bold text, repeat header flag
    hdr_shd = t.rows[0].cells[0]._tc.tcPr.find(qn("w:shd"))
    assert hdr_shd is not None and hdr_shd.get(qn("w:fill")) == "1F4E79"
    assert t.rows[0].cells[0].paragraphs[0].runs[0].font.bold is True
    assert t.rows[0].cells[0].paragraphs[0].runs[0].font.color.rgb == (0xFF, 0xFF, 0xFF)
    assert t.rows[0]._tr.trPr.find(qn("w:tblHeader")) is not None
    # Alternating row shading on row 2
    alt_shd = t.rows[2].cells[0]._tc.tcPr.find(qn("w:shd"))
    assert alt_shd is not None and alt_shd.get(qn("w:fill")) == "D9E1F2"
    # Table borders and cell margins
    assert t._tbl.tblPr.find(qn("w:tblBorders")) is not None
    tc_mar = t.rows[0].cells[0]._tc.tcPr.find(qn("w:tcMar"))
    assert tc_mar is not None and tc_mar.find(qn("w:left")).get(qn("w:w")) == "160"  # 8pt * 20 dxa


@pytest.mark.asyncio
async def test_mcp_server_registers_new_formatting_tools():
    """Verify all new Excel & Word formatting tools are registered cleanly on the MCP server."""
    tools = await server.mcp.list_tools()
    tool_names = {t.name for t in tools}
    for expected in (
        "set_column_widths",
        "format_range",
        "format_worksheet",
        "batch_write_ranges",
        "batch_format_workbook",
        "update_excel_table_style",
        "list_cell_styles",
        "apply_cell_style",
        "add_data_validation",
        "list_data_validations",
        "clear_data_validations",
        "excel_live_circle_invalid",
        "excel_live_custom_views",
        "excel_live_update_table_style",
        "excel_view_capabilities",
        "format_paragraph",
        "format_table",
        "word_live_list_documents",
    ):
        assert expected in tool_names, f"Expected MCP tool '{expected}' not registered"


@pytest.mark.asyncio
async def test_mcp_server_compact_profile_reduces_tools():
    """Verify M365_MCP_PROFILE=compact (and legacy OFFICE_MCP_PROFILE + shim) exposes the lean core Word+Excel toolset cleanly."""
    import os
    env = dict(os.environ)
    env.pop("OFFICE_MCP_PROFILE", None)
    env["M365_MCP_PROFILE"] = "compact"
    proc = await asyncio.create_subprocess_exec(
        sys.executable,
        "-c",
        (
            "import asyncio, json, sys; sys.path.insert(0, 'src'); "
            "import m365_mac_mcp_server as s; "
            "tools = asyncio.run(s.mcp.list_tools()); "
            "print(json.dumps([t.name for t in tools]))"
        ),
        env=env,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    stdout, stderr = await proc.communicate()
    assert proc.returncode == 0, stderr.decode("utf-8", errors="replace")
    names = set(json.loads(stdout.decode("utf-8")))
    assert 30 <= len(names) <= 45
    for core in (
        "open_workbook",
        "batch_write_ranges",
        "batch_format_workbook",
        "excel_live_open",
        "excel_live_write_cell",
        "excel_live_save",
        "create_document",
        "add_table",
        "word_live_open",
        "word_live_list_documents",
        "word_live_save",
    ):
        assert core in names
    assert "create_presentation" not in names

    # Verify composable app profiles: excel,compact (25), word,powerpoint (33), invalid fallback, and legacy shim + OFFICE_MCP_PROFILE fallback
    for prof, must_have, must_not_have in (
        (
            "excel,compact",
            {
                "inspect_workbook",
                "audit_workbook_formulas",
                "excel_live_batch_write_cells",
                "create_chart",
                "delete_worksheet",
                "excel_live_list_workbooks",
            },
            {"create_document", "create_presentation", "format_worksheet", "add_formula"},
        ),
        ("word,powerpoint", {"create_document", "word_live_list_documents", "create_presentation"}, {"open_workbook", "excel_live_open"}),
        ("invalid_profile_xyz", {"open_workbook", "create_document", "create_presentation"}, set()),
    ):
        env_p = dict(os.environ)
        env_p.pop("OFFICE_MCP_PROFILE", None)
        env_p["M365_MCP_PROFILE"] = prof
        p2 = await asyncio.create_subprocess_exec(
            sys.executable,
            "-c",
            (
                "import asyncio, json, sys; sys.path.insert(0, 'src'); "
                "import m365_mac_mcp_server as s; "
                "tools = asyncio.run(s.mcp.list_tools()); "
                "print(json.dumps({'profile': s._ACTIVE_PROFILE_LABEL, 'names': [t.name for t in tools]}))"
            ),
            env=env_p,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        out2, err2 = await p2.communicate()
        assert p2.returncode == 0, err2.decode("utf-8", errors="replace")
        parsed = json.loads(out2.decode("utf-8"))
        name_set = set(parsed["names"])
        for req in must_have:
            assert req in name_set, f"{req} missing in profile {prof}"
        for exc in must_not_have:
            assert exc not in name_set, f"{exc} unexpectedly present in profile {prof}"
        if prof == "invalid_profile_xyz":
            assert parsed["profile"] == "full"
            assert "Unknown M365_MCP_PROFILE" in err2.decode("utf-8", errors="replace")

    # Verify legacy office365_mcp_server.py shim + OFFICE_MCP_PROFILE fallback still works cleanly
    env_legacy = dict(os.environ)
    env_legacy.pop("M365_MCP_PROFILE", None)
    env_legacy["OFFICE_MCP_PROFILE"] = "excel,compact"
    p_leg = await asyncio.create_subprocess_exec(
        sys.executable,
        "-c",
        (
            "import asyncio, json, sys; sys.path.insert(0, 'src'); "
            "import office365_mcp_server as s; "
            "tools = asyncio.run(s.mcp.list_tools()); "
            "print(json.dumps({'profile': s._ACTIVE_PROFILE_LABEL, 'count': len(tools)}))"
        ),
        env=env_legacy,
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.PIPE,
    )
    out_leg, err_leg = await p_leg.communicate()
    assert p_leg.returncode == 0, err_leg.decode("utf-8", errors="replace")
    parsed_leg = json.loads(out_leg.decode("utf-8"))
    assert parsed_leg["profile"] == "excel,compact"
    assert parsed_leg["count"] == 25
    err_leg_text = err_leg.decode("utf-8", errors="replace")
    assert "deprecated" in err_leg_text.lower()

    # Verify case-insensitivity (ALLCAPS, MixedCaps, smallcaps, and aliases like XLS)
    for raw_val in ("EXCEL,COMPACT", "Excel,Compact", "excel,compact", "XLS, Compact"):
        label, apps, is_compact = server._parse_mcp_profile(raw_val)
        assert label == "excel,compact"
        assert apps == {"excel"}
        assert is_compact is True

    # Verify check_office_status surfaces profile metadata cleanly
    status = await server.check_office_status()
    assert "profile" in status
    assert "active_profile" in status["profile"]
    assert "registered_tool_count" in status["profile"]
    assert "excel,compact" in status["profile"]["available_profiles"]


@pytest.mark.asyncio
async def test_inspect_workbook_audit_formulas_and_param_aliases(tmp_path: Path):
    """Verify inspect_workbook, audit_workbook_formulas (by id & file_path), and range_ref/cell_range aliases."""
    created = await server.create_workbook(title="Audit Test", sheet_name="Income Statement")
    wb_id = created["workbook_id"]
    await server.add_worksheet(wb_id, "Balance Sheet")

    await server.batch_write_ranges(
        wb_id,
        operations=[
            {
                "sheet_name": "Income Statement",
                "start_cell": "A1",
                "data": [
                    ["Line Item", "FY2024", "FY2025", "CAGR"],
                    ["Revenue", 100, 120, "=(C2/B2)-1"],
                    ["COGS", 40, 45, "=(C3/B3)-1"],
                    ["Gross Profit", 60, 75, "=(C4/B4)-1"],
                    ["Other Income", -5, 10, '""'],
                    ["Broken Link", 0, 0, "=Balance_Sheet!B2"],
                    ["Error Cell", 0, 0, "#REF!"],
                ],
            },
            {
                "sheet_name": "Balance Sheet",
                "start_cell": "A1",
                "data": [["Line Item", "FY2024"], ["Cash", 500]],
            },
        ],
    )

    # Test read_range with range_ref + include_formulas aliases (consistent with excel_live_read_range)
    rr = await server.read_range(
        wb_id, sheet_name="Income Statement", range_ref="A1:D3", include_formulas=True
    )
    _assert_strict_json(rr)
    assert rr["status"] == "success"
    assert rr["data"][1][3] == "=(C2/B2)-1"

    # Test format_range with range_ref alias
    fr = await server.format_range(
        wb_id, sheet_name="Income Statement", range_ref="A1:D1", formatting={"bold": True}
    )
    _assert_strict_json(fr)
    assert fr["status"] == "success"

    out_path = tmp_path / "Audit_Sample.xlsx"
    await server.save_workbook(wb_id, str(out_path), format="xlsx")

    # 1. inspect_workbook by workbook_id
    insp = await server.inspect_workbook(wb_id)
    _assert_strict_json(insp)
    assert insp["status"] == "success"
    assert insp["sheet_names"] == ["Income Statement", "Balance Sheet"]
    assert len(insp["sheets"]) == 2
    assert insp["sheets"][0]["formulas"]["D2"] == "=(C2/B2)-1"
    assert insp["sheets"][0]["values"][1][0] == "Revenue"
    assert "formulas_or_values" not in insp["sheets"][0]
    assert "cached_values" not in insp["sheets"][0]

    # 2. audit_workbook_formulas by workbook_id
    aud = await server.audit_workbook_formulas(wb_id)
    _assert_strict_json(aud)
    assert aud["status"] == "success"
    assert aud["clean"] is False
    assert aud["error_count"] == 1
    assert aud["error_cells"][0]["error"] == "#REF!"
    assert aud["sheet_ref_warning_count"] == 1
    assert "Balance_Sheet" in aud["sheet_ref_warnings"][0]
    assert any(item["cell"] == "D5" for item in aud["suspicious_literals"])

    await server.close_workbook(wb_id)

    # 3. inspect_workbook & audit_workbook_formulas directly via file_path (1-call natural-language flow)
    insp_path = await server.inspect_workbook(file_path=str(out_path))
    _assert_strict_json(insp_path)
    assert insp_path["sheet_count"] == 2

    aud_path = await server.audit_workbook_formulas(file_path=str(out_path))
    _assert_strict_json(aud_path)
    assert aud_path["error_count"] == 1
    assert aud_path["sheet_ref_warning_count"] == 1


@pytest.mark.asyncio
async def test_create_temp_files_skip_live_open_to_avoid_sandbox_grant_access(monkeypatch):
    """create_workbook(open_in_excel=True) and create_document(open_in_word=True) must not open /tmp UUID files."""
    excel_opened = False
    word_opened = False

    async def fail_excel_open(*args, **kwargs):
        nonlocal excel_opened
        excel_opened = True
        raise AssertionError("Should not open temp workbook in live Excel")

    async def fail_word_open(*args, **kwargs):
        nonlocal word_opened
        word_opened = True
        raise AssertionError("Should not open temp document in live Word")

    monkeypatch.setattr(server.excel.applescript, "open_excel_file", fail_excel_open)
    monkeypatch.setattr(server.word.applescript, "open_word_file", fail_word_open)

    wb = await server.create_workbook(title="Temp Guard Test", open_in_excel=True)
    _assert_strict_json(wb)
    assert wb["opened_in_excel"] is False
    assert wb["is_temporary"] is True
    assert "workspace folder" in wb["note"]
    assert excel_opened is False
    await server.close_workbook(wb["workbook_id"])

    doc = await server.create_document(title="Temp Guard Doc", open_in_word=True)
    _assert_strict_json(doc)
    assert doc["opened_in_word"] is False
    assert doc["is_temporary"] is True
    assert "workspace folder" in doc["note"]
    assert word_opened is False
    await server.close_document(doc["document_id"])


@pytest.mark.asyncio
async def test_comma_separated_ranges_and_auto_color_financial_model(tmp_path: Path):
    """Verify comma-separated multi-ranges, flat style rules, borders alias, and auto_color_financial_model."""
    from openpyxl import load_workbook

    wb = await server.create_workbook(title="AutoColor", sheet_name="IS")
    wb_id = wb["workbook_id"]
    await server.add_worksheet(wb_id, "BS")
    await server.batch_write_ranges(
        wb_id,
        operations=[
            {
                "sheet_name": "IS",
                "start_cell": "A1",
                "data": [
                    ["Line Item", "FY24", "FY25"],
                    ["Revenue", 100, 120],
                    ["COGS", 40, 50],
                    ["Gross Profit", "=B2-B3", "=C2-C3"],
                    ["Link", "='BS'!B2", "='BS'!C2"],
                ],
            },
            {
                "sheet_name": "BS",
                "start_cell": "A1",
                "data": [["Item", "FY24", "FY25"], ["Cash", 500, 600]],
            },
        ],
    )

    res = await server.batch_format_workbook(
        wb_id,
        operations=[
            {
                "sheet_name": "IS",
                "auto_color_financial_model": True,
                "styles": [
                    {"range": "A1:C1", "bold": True, "bg_color": "#1F4E79", "font_color": "#FFFFFF"},
                    {"range": "B2:C2, B4:C4", "number_format": "$#,##0", "borders": {"top": "thin", "bottom": "double"}},
                ],
            }
        ],
    )
    _assert_strict_json(res)
    out = tmp_path / "AutoColor.xlsx"
    await server.save_workbook(wb_id, str(out), format="xlsx")
    await server.close_workbook(wb_id)

    reloaded = load_workbook(str(out))
    ws = reloaded["IS"]
    assert ws["B2"].font.color.rgb.endswith("0000FF")  # blue input
    assert ws["B4"].font.color.rgb.endswith("000000")  # black same-sheet formula
    assert ws["B5"].font.color.rgb.endswith("008000")  # green cross-sheet link
    assert ws["B2"].number_format == "$#,##0"
    assert ws["B4"].border.bottom.style == "double"
    assert ws["A1"].fill.start_color.rgb.endswith("1F4E79")
    reloaded.close()


@pytest.mark.asyncio
async def test_office_mcp_live_sync_modes_and_auto_reload(monkeypatch, tmp_path: Path):
    """Verify M365_MCP_LIVE_SYNC (and legacy OFFICE_MCP_LIVE_SYNC fallback) auto-reloads saved files and guards unsaved edits."""
    # 1. Case-insensitive parsing & default='auto'
    monkeypatch.delenv("M365_MCP_LIVE_SYNC", raising=False)
    monkeypatch.delenv("OFFICE_MCP_LIVE_SYNC", raising=False)
    assert ExcelController.get_live_sync_mode() == "auto"
    monkeypatch.setenv("OFFICE_MCP_LIVE_SYNC", "OFF")
    assert ExcelController.get_live_sync_mode() == "off"
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "AUTO")
    assert ExcelController.get_live_sync_mode() == "auto"
    monkeypatch.delenv("OFFICE_MCP_LIVE_SYNC", raising=False)
    for raw in ("AUTO", "Auto", "auto"):
        assert ExcelController.get_live_sync_mode(raw) == "auto"
    for raw in ("OFF", "Off", "off"):
        assert ExcelController.get_live_sync_mode(raw) == "off"
    assert ExcelController.get_live_sync_mode("unknown_val") == "auto"

    # 2. Auto mode: file open in Excel with saved=True -> capture -> close -> save -> reopen -> restore
    events = []
    out_file = tmp_path / "LiveSync_Test.xlsx"

    wb = await server.create_workbook(title="SyncTest", sheet_name="Sheet1")
    wb_id = wb["workbook_id"]
    await server.write_cell(wb_id, "Sheet1", "A1", "Hello")

    async def fake_find_saved(path: str):
        events.append("find")
        return {"name": out_file.name, "full_name": str(out_file), "saved": True}

    async def fake_capture(wb_name: str):
        events.append("capture")
        return {"active_sheet": "Sheet1", "bounds": [100, 100, 1200, 800], "zoom": 115, "scroll_row": 5, "scroll_column": 2, "selection": "C7"}

    async def fake_close(workbook_name=None, save_changes="no"):
        events.append(f"close:{save_changes}")
        return {"status": "success"}

    async def fake_open(file_path: str, read_only=False, activate=False):
        events.append(f"open:activate={activate}")
        assert Path(file_path).exists()
        return {"status": "success", "workbook_name": out_file.name}

    async def fake_restore(wb_name: str, state: dict):
        events.append(f"restore:{state.get('zoom')}")
        return True

    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "Auto")
    monkeypatch.setattr(server.excel.applescript, "find_open_excel_workbook_by_path", fake_find_saved)
    monkeypatch.setattr(server.excel.applescript, "capture_excel_workbook_view_state", fake_capture)
    monkeypatch.setattr(server.excel.applescript, "close_excel_workbook", fake_close)
    monkeypatch.setattr(server.excel.applescript, "open_excel_file", fake_open)
    monkeypatch.setattr(server.excel.applescript, "restore_excel_workbook_view_state", fake_restore)

    saved = await server.save_workbook(wb_id, str(out_file), format="xlsx")
    _assert_strict_json(saved)
    assert saved["status"] == "success"
    assert saved["reloaded_in_excel"] is True
    assert saved["view_restored"] is True
    assert events == ["find", "capture", "close:no", "open:activate=False", "restore:115"]

    # 3. Auto mode: file open in Excel with saved=False -> refuses to overwrite/close
    events.clear()

    async def fake_find_unsaved(path: str):
        events.append("find_unsaved")
        return {"name": out_file.name, "full_name": str(out_file), "saved": False}

    monkeypatch.setattr(server.excel.applescript, "find_open_excel_workbook_by_path", fake_find_unsaved)
    err = await server.save_workbook(wb_id, str(out_file), format="xlsx")
    _assert_strict_json(err)
    assert err["status"] == "error"
    assert "unsaved manual changes" in err["error"]
    assert events == ["find_unsaved"]

    # 4. OFF mode: never queries live Excel
    events.clear()
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "OFF")
    saved_off = await server.save_workbook(wb_id, str(out_file), format="xlsx")
    _assert_strict_json(saved_off)
    assert saved_off["status"] == "success"
    assert "reloaded_in_excel" not in saved_off
    assert events == []

    # 5. check_office_status reports live_sync_mode
    status = await server.check_office_status()
    assert status["profile"]["live_sync_mode"] == "off"

    await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_excel_batch_ops_sheet_ref_guard_styles_tables_and_validation(tmp_path: Path):
    """Verify batch multi-sheet write/format, sheet-ref guard, table styles, named cell styles, and data validation."""
    from openpyxl import load_workbook

    buf = io.StringIO()
    with redirect_stdout(buf):
        created = await server.create_workbook(title="Model Suite", sheet_name="Balance Sheet")
        wb_id = created["workbook_id"]
        await server.add_worksheet(wb_id, "Summary")

        # 1. Batch write across two sheets in one call, including a misspelled sheet ref to trigger guard
        bw = await server.batch_write_ranges(
            wb_id,
            operations=[
                {
                    "sheet_name": "Balance Sheet",
                    "start_cell": "A1",
                    "data": [
                        ["Line Item", "FY2025", "FY2026", "Status"],
                        ["Cash", 1000000, 1400000, "High"],
                        ["Debt", 500000, 450000, "Low"],
                    ],
                },
                {
                    "sheet_name": "Summary",
                    "start_cell": "A1",
                    "data": [
                        ["Metric", "Value", "Check"],
                        ["Cash (Auto-quoted)", "=Balance Sheet!C2", "=Balance_Sheet!C2"],
                        ["External Guard", "=[Other.xlsx]Sheet1!A1", "OK"],
                    ],
                },
            ],
        )
        _assert_strict_json(bw)
        assert bw["status"] == "success"
        assert bw["operations_completed"] == 2
        warnings = bw.get("formula_warnings", [])
        assert any("Balance_Sheet" in w and "Did you mean 'Balance Sheet'!?" in w for w in warnings)
        assert any("external workbook reference" in w for w in warnings)

        # 2. Batch format both sheets in one call (widths, freeze panes, named cell style + overrides)
        bf = await server.batch_format_workbook(
            wb_id,
            operations=[
                {
                    "sheet_name": "Balance Sheet",
                    "column_widths": {"A": 28, "B:D": 16},
                    "freeze_panes": "B2",
                    "styles": [
                        {"range": "A1:D1", "cell_style": "Heading 1"},
                        {"range": "B2:C3", "cell_style": "Input", "number_format": '$#,##0;($#,##0);"-"'},
                    ],
                },
                {
                    "sheet_name": "Summary",
                    "column_widths": {"A": 30, "B:C": 18},
                    "freeze_panes": "A2",
                    "styles": [
                        {"range": "A1:C1", "cell_style": "Title"},
                        {"range": "B2:B2", "cell_style": "Linked Cell", "number_format": '$#,##0;($#,##0);"-"'},
                    ],
                },
            ],
        )
        _assert_strict_json(bf)
        assert bf["sheets_formatted"] == 2

        # 3. Named cell styles listing & direct apply_cell_style
        style_list = await server.list_cell_styles(wb_id)
        _assert_strict_json(style_list)
        assert "Good" in style_list["built_in_presets"]
        assert "Total" in style_list["built_in_presets"]

        applied_good = await server.apply_cell_style(
            wb_id, "Summary", "C3", "Good", overrides={"bold": True}
        )
        _assert_strict_json(applied_good)
        assert applied_good["style_name"] == "Good"

        # 4. Table styles (TableStyleDark2 -> update to TableStyleLight9) + banded toggles
        tbl = await server.add_excel_table(
            wb_id,
            "Balance Sheet",
            "BS_Table",
            "A1:D3",
            style_name="TableStyleDark2",
            show_first_column=True,
            show_last_column=False,
            show_row_stripes=False,
            show_column_stripes=True,
        )
        _assert_strict_json(tbl)
        assert tbl["style"] == "TableStyleDark2"
        assert tbl["show_column_stripes"] is True

        upd_tbl = await server.update_excel_table_style(
            wb_id,
            "BS_Table",
            sheet_name="Balance Sheet",
            style_name="TableStyleLight9",
            show_row_stripes=True,
            show_column_stripes=False,
        )
        _assert_strict_json(upd_tbl)
        assert upd_tbl["style"] == "TableStyleLight9"
        assert upd_tbl["show_row_stripes"] is True
        assert upd_tbl["show_column_stripes"] is False

        # 5. Data validation rules (list, whole, custom) + list + partial clear
        dv_list = await server.add_data_validation(
            wb_id,
            "Balance Sheet",
            "D2:D10",
            validation_type="list",
            formula1=["Low", "Medium", "High"],
            prompt_title="Status",
            prompt="Choose Low, Medium, or High",
        )
        _assert_strict_json(dv_list)
        assert dv_list["formula1"] == '"Low,Medium,High"'

        dv_whole = await server.add_data_validation(
            wb_id,
            "Balance Sheet",
            "B2:C10",
            validation_type="whole",
            operator="greaterThanOrEqual",
            formula1=0,
            error_title="Invalid Amount",
            error="Must be >= 0",
        )
        _assert_strict_json(dv_whole)

        dv_custom = await server.add_data_validation(
            wb_id,
            "Summary",
            "B2:B5",
            validation_type="custom",
            formula1="=B2>0",
        )
        _assert_strict_json(dv_custom)

        listed_dv = await server.list_data_validations(wb_id, "Balance Sheet")
        _assert_strict_json(listed_dv)
        assert listed_dv["count"] == 2

        # Save, reload, and verify persistence
        out_path = tmp_path / "Full_Excel_Capabilities.xlsx"
        await server.save_workbook(wb_id, str(out_path), format="xlsx")

        # Test clear_data_validations on Summary
        cleared = await server.clear_data_validations(wb_id, "Summary", cell_range="B2:B5")
        _assert_strict_json(cleared)
        assert cleared["removed_count"] == 1
        assert cleared["remaining_count"] == 0

        # 6. Actionable error hint & view capabilities
        err = await server.read_cell(wb_id, "A1", sheet_name="DoesNotExist")
        _assert_strict_json(err)
        assert err["status"] == "error"
        assert "hint" in err and "list_worksheets" in err["hint"]

        caps = await server.excel_view_capabilities()
        _assert_strict_json(caps)
        assert "focus_cell" in caps["manual_ui_only"]
        assert "invalid_data_circles" in caps["supported"]

        await server.close_workbook(wb_id)

    assert buf.getvalue() == ""

    # Reload persisted workbook with openpyxl and verify everything
    wb = load_workbook(str(out_path))
    bs = wb["Balance Sheet"]
    assert bs.freeze_panes == "B2"
    assert bs.column_dimensions["A"].width == 28.0
    assert bs["A1"].style == "Heading 1"
    assert bs["B2"].style == "Input"
    assert bs["B2"].number_format == '$#,##0;($#,##0);"-"'
    assert bs.tables["BS_Table"].tableStyleInfo.name == "TableStyleLight9"
    assert bs.tables["BS_Table"].tableStyleInfo.showRowStripes is True
    assert bs.tables["BS_Table"].tableStyleInfo.showColumnStripes is False
    assert len(bs.data_validations.dataValidation) == 2

    sm = wb["Summary"]
    assert sm["B2"].value == "='Balance Sheet'!C2"
    assert sm["C3"].style == "Good"
    assert sm["C3"].font.bold is True
    wb.close()


@pytest.mark.asyncio
async def test_token_optimizations_compact_json_schema_and_sparse_formulas(monkeypatch):
    """Verify compact JSON MCP output, stripped schema noise, sparse formula maps, and trimmed blanks."""
    import json

    # 1. Tool schemas have no Pydantic 'title' noise, no 'default': null, and no 2-item nullable anyOf,
    #    while preserving legitimate parameters named 'title' (e.g. create_workbook/create_document/create_presentation)
    tools_map = server.mcp._tool_manager._tools
    for name, tool in tools_map.items():
        params = tool.parameters
        assert "title" not in params, f"{name} root schema still has 'title'"
        for prop_name, prop_schema in (params.get("properties") or {}).items():
            assert "title" not in prop_schema, f"{name}.{prop_name} still has Pydantic 'title'"
        schema_str = json.dumps(params)
        assert '"default": null' not in schema_str, f"{name} still has redundant 'default': null"
        assert '{"type": "null"}' not in schema_str, f"{name} still has verbose nullable anyOf"
        assert '"additionalProperties": true' not in schema_str, f"{name} still has redundant additionalProperties: true"
    assert "title" in tools_map["create_workbook"].parameters["properties"]
    assert "title" in tools_map["create_document"].parameters["properties"]
    assert "title" in tools_map["create_presentation"].parameters["properties"]

    # 2. MCP tool call returns compact single-line JSON with structured_content=None (no duplication)
    res = await server.mcp._tool_manager.call_tool(
        "excel_view_capabilities", {}, context=None, convert_result=True
    )
    raw_text = res.content[0].text
    assert "\n" not in raw_text
    assert getattr(res, "structured_content", None) is None
    parsed = json.loads(raw_text)
    assert parsed["status"] == "success"

    # 3. Runtime validation still accepts explicit None for Optional[T] params despite simplified schema
    created = await server.create_workbook(title="Compact Test", sheet_name="Model")
    wb_id = created["workbook_id"]
    await server.write_range(
        wb_id,
        sheet_name="Model",
        start_cell="A1",
        data=[
            ["Metric", "FY24", "FY25", "Growth", None],
            ["Revenue", 100, 120, "=(C2/B2)-1", None],
            [None, None, None, None, None],
        ],
    )
    insp_res = await server.mcp._tool_manager.call_tool(
        "inspect_workbook",
        {"workbook_id": wb_id, "file_path": None},
        context=None,
        convert_result=True,
    )
    insp_text = insp_res.content[0].text
    assert "\n" not in insp_text
    insp = json.loads(insp_text)
    sh = insp["sheets"][0]
    # Trailing blank row and column trimmed
    assert len(sh["values"]) == 2
    assert len(sh["values"][0]) == 4
    # Sparse formula map only contains formula cells
    assert sh["formulas"] == {"D2": "=(C2/B2)-1"}
    # Empty boilerplate keys omitted
    assert "tables" not in sh
    assert "freeze_panes" not in sh
    assert "truncated" not in sh
    assert "external_links_count" not in insp
    await server.close_workbook(wb_id)

    # 4. Live Excel read_range returns sparse A1->formula map and omits 'method': 'applescript'
    us, rs, gs = chr(31), chr(30), chr(29)
    fake_raw = (
        f"Model.xlsx{us}Income Statement{gs}"
        f"Revenue{us}100{us}120{us}0.2{rs}Margin{us}50{us}60{us}0.5{gs}"
        f"Revenue{us}100{us}120{us}=(C2/B2)-1{rs}Margin{us}50{us}60{us}=C3/C2"
    )

    async def _fake_osascript(script, args=None, timeout=30.0):
        return fake_raw

    monkeypatch.setattr(server.applescript, "execute_applescript", _fake_osascript)
    live = await server.excel_live_read_range(
        range_ref="A2:D3",
        sheet_name="Income Statement",
        workbook_name="Model.xlsx",
        include_formulas=True,
    )
    assert "method" not in live
    assert live["values"] == [["Revenue", "100", "120", "0.2"], ["Margin", "50", "60", "0.5"]]
    assert live["formulas"] == {"D2": "=(C2/B2)-1", "D3": "=C3/C2"}


@pytest.mark.asyncio
async def test_batch_format_workbook_accepts_llm_style_aliases_and_warns_unknown(tmp_path: Path):
    """Verify batch_format_workbook handles rules/formats/number_formats, nested font/fill/style, and warns on unknown keys."""
    from openpyxl import load_workbook

    created = await server.create_workbook(title="Alias Format Test", sheet_name="IS")
    wb_id = created["workbook_id"]
    await server.write_range(
        wb_id,
        sheet_name="IS",
        start_cell="A1",
        data=[
            ["Line Item", "FY2025"],
            ["Revenue", 416161],
            ["Gross Margin %", 0.469],
        ],
    )
    res = await server.batch_format_workbook(
        wb_id,
        operations=[
            {
                "sheet_name": "IS",
                "column_widths": {"A": 44, "B": 15},
                "freeze_panes": "B2",
                "rules": [
                    {
                        "range": "A1:B1",
                        "style": {
                            "font": {"bold": True, "color": "#FFFFFF"},
                            "fill": {"fgColor": "#1F4E79"},
                            "alignment": {"horizontal": "center"},
                        },
                    }
                ],
                "number_formats": {
                    "B2": '$#,##0;($#,##0);"-"',
                    "B3": "0.0%",
                },
                "bogus_key": True,
            },
            # Flat per-range operation item fallback
            {
                "sheet_name": "IS",
                "range": "A2:A3",
                "italic": True,
                "bg_color": "#D9E1F2",
            },
        ],
    )
    assert res["status"] == "success"
    assert res["total_cells_formatted"] == 6
    assert "warnings" in res and "bogus_key" in res["warnings"][0]

    out_path = tmp_path / "AliasFormat.xlsx"
    await server.save_workbook(wb_id, str(out_path), format="xlsx")
    await server.close_workbook(wb_id)

    wb = load_workbook(str(out_path))
    ws = wb["IS"]
    assert ws["A1"].font.bold is True
    assert ws["A1"].font.color.rgb.endswith("FFFFFF")
    assert ws["A1"].fill.start_color.rgb.endswith("1F4E79")
    assert ws["A1"].alignment.horizontal == "center"
    assert ws["A2"].font.italic is True
    assert ws["A2"].fill.start_color.rgb.endswith("D9E1F2")
    assert ws["B2"].number_format == '$#,##0;($#,##0);"-"'
    assert ws["B3"].number_format == "0.0%"
    wb.close()


@pytest.mark.asyncio
async def test_office_mcp_formula_guard_file_and_live(monkeypatch):
    """Verify M365_MCP_FORMULA_GUARD (and legacy OFFICE_MCP_FORMULA_GUARD fallback) across all 6 Excel write tools."""
    from controllers.excel_controller import ExcelController

    # 1. Case-insensitive mode parsing + fallback
    monkeypatch.delenv("M365_MCP_FORMULA_GUARD", raising=False)
    monkeypatch.delenv("OFFICE_MCP_FORMULA_GUARD", raising=False)
    assert ExcelController.get_formula_guard_mode() == "literals_only"
    monkeypatch.setenv("OFFICE_MCP_FORMULA_GUARD", "STRICT")
    assert ExcelController.get_formula_guard_mode() == "strict"
    assert AppleScriptBridge.get_formula_guard_mode() == "strict"
    monkeypatch.setenv("M365_MCP_FORMULA_GUARD", "literals_only")
    assert ExcelController.get_formula_guard_mode() == "literals_only"
    assert AppleScriptBridge.get_formula_guard_mode() == "literals_only"
    monkeypatch.delenv("OFFICE_MCP_FORMULA_GUARD", raising=False)
    assert ExcelController.get_formula_guard_mode("STRICT") == "strict"
    assert ExcelController.get_formula_guard_mode("Off") == "off"
    assert ExcelController.get_formula_guard_mode("bogus") == "literals_only"

    # 2. File-level guard in default 'literals_only' mode
    created = await server.create_workbook(title="Guard Test", sheet_name="Model")
    wb_id = created["workbook_id"]
    await server.write_range(
        wb_id,
        sheet_name="Model",
        start_cell="A1",
        data=[
            ["Item", "FY24", "FY25"],
            ["Rev", 100, 120],
            ["Total", "=B2+C2", "=C2*2"],
        ],
    )

    # Formula -> hardcoded literal BLOCKED (workbook untouched)
    blk1 = await server.write_cell(wb_id, "Model", "B3", 220)
    assert blk1["status"] == "confirmation_required"
    assert blk1["guard_mode"] == "literals_only"
    assert blk1["blocked_count"] == 1
    assert blk1["blocked_cells"][0]["existing_formula"] == "=B2+C2"
    assert (await server.read_cell(wb_id, "B3", sheet_name="Model"))["formula"] == "=B2+C2"

    # Formula -> None (clearing) ALSO BLOCKED in literals_only
    blk_none = await server.write_cell(wb_id, "Model", "B3", None)
    assert blk_none["status"] == "confirmation_required"

    # Formula -> new Formula ALLOWED in literals_only + reports formulas_updated=1
    ok_fml = await server.write_cell(wb_id, "Model", "B3", "=SUM(B2:C2)")
    assert ok_fml["status"] == "success"
    assert ok_fml["formulas_updated"] == 1

    # Atomic range block: writing A3:C3 with a literal over B3 leaves A3 untouched too
    blk_rng = await server.write_range(wb_id, "Model", "A3", [["CHANGED_LABEL", 999, "=C2*3"]])
    assert blk_rng["status"] == "confirmation_required"
    assert (await server.read_cell(wb_id, "A3", sheet_name="Model"))["value"] == "Total"

    # Atomic batch_write_ranges block
    blk_batch = await server.batch_write_ranges(
        wb_id,
        operations=[{"sheet_name": "Model", "start_cell": "C3", "data": [[240]]}],
    )
    assert blk_batch["status"] == "confirmation_required"
    assert (await server.read_cell(wb_id, "C3", sheet_name="Model"))["formula"] == "=C2*2"

    # Explicit allow_formula_overwrite=True permits overwrite
    ok_override = await server.write_cell(wb_id, "Model", "B3", 220, allow_formula_overwrite=True)
    assert ok_override["status"] == "success"
    assert ok_override["formulas_updated"] == 1

    # 3. 'strict' mode also blocks Formula -> new Formula
    monkeypatch.setenv("M365_MCP_FORMULA_GUARD", "STRICT")
    blk_strict = await server.write_cell(wb_id, "Model", "C3", "=C2*3")
    assert blk_strict["status"] == "confirmation_required"
    assert blk_strict["guard_mode"] == "strict"

    # 4. 'off' mode allows everything
    monkeypatch.setenv("M365_MCP_FORMULA_GUARD", "off")
    ok_off = await server.write_cell(wb_id, "Model", "C3", 240)
    assert ok_off["status"] == "success"
    await server.close_workbook(wb_id)

    # 5. Live Excel writers parse BLOCKED responses & pass guardMode arg cleanly
    monkeypatch.setenv("M365_MCP_FORMULA_GUARD", "literals_only")
    us, rs, gs = chr(31), chr(30), chr(29)
    captured_args = []

    async def _fake_live(script, args=None, timeout=30.0):
        captured_args.append(list(args or []))
        if "BLOCKED" in script and len(args or []) == 8:
            return f"BLOCKED{us}Model.xlsx{us}Balance Sheet{us}=B23+'Income Statement'!C15-C22"
        if "BLOCKED" in script and len(args or []) == 7:
            return f"BLOCKED{gs}Balance Sheet{us}C23{us}=B23+1{us}-5034"
        return f"BLOCKED{gs}Balance Sheet{us}C23{us}=B23+1{us}-5034"

    monkeypatch.setattr(server.applescript, "execute_applescript", _fake_live)

    l1 = await server.excel_live_write_cell("C23", -5034, sheet_name="Balance Sheet")
    assert l1["status"] == "confirmation_required"
    assert l1["blocked_cells"][0]["cell"] == "C23"
    assert captured_args[-1][7] == "literals_only"

    l2 = await server.excel_live_write_range("C23", [[-5034]], sheet_name="Balance Sheet")
    assert l2["status"] == "confirmation_required"
    assert captured_args[-1][6] == "literals_only"

    l3 = await server.excel_live_batch_write_cells({"C23": -5034}, sheet_name="Balance Sheet")
    assert l3["status"] == "confirmation_required"
    assert captured_args[-1][3] == "literals_only"

    # Verify check_office_status exposes formula_guard_mode
    st = await server.check_office_status()
    assert st["profile"]["formula_guard_mode"] == "literals_only"
    assert st["profile"]["available_formula_guard_modes"] == ["literals_only", "strict", "off"]


def test_all_finance_skills_valid_schema_size_and_compliance():
    """Verify every skills/*/SKILL.md is pure Markdown (2-4 KB), 5-block, IFRS/FASB/SOX compliant, and references valid tools/params."""
    import json
    import re
    import subprocess

    skills_root = Path(__file__).resolve().parent.parent / "skills"
    skill_files = sorted(skills_root.glob("*/SKILL.md"))
    assert len(skill_files) >= 9

    tools_map = server.mcp._tool_manager._tools
    valid_tools = set(tools_map.keys())

    # Fetch exact excel,compact tool set so we guarantee all Excel skills work under excel,compact
    compact_out = subprocess.check_output(
        [
            sys.executable,
            "-c",
            (
                "import json, os, sys; os.environ['M365_MCP_PROFILE']='excel,compact'; "
                "sys.path.insert(0, 'src'); import m365_mac_mcp_server as s; "
                "print(json.dumps(sorted(s.mcp._tool_manager._tools.keys())))"
            ),
        ],
        text=True,
    )
    excel_compact_tools = set(json.loads(compact_out))

    skills_readme = (skills_root / "README.md").read_text(encoding="utf-8")
    assert "**Finance Skills Suite Version:**" in skills_readme, "skills/README.md missing Finance Skills Suite Version"

    for p in skill_files:
        txt = p.read_text(encoding="utf-8")
        size_bytes = len(txt.encode("utf-8"))
        assert 2048 <= size_bytes <= 4096, f"{p.parent.name} size {size_bytes}B outside 2-4 KB target"
        assert txt.startswith("---\nname:") and "\ndescription:" in txt, f"{p.parent.name} missing YAML frontmatter"
        ver_match = re.search(r"^version:\s*(\d+\.\d+\.\d+[A-Z])\s*$", txt, re.MULTILINE)
        assert ver_match, f"{p.parent.name} missing YAML 'version: X.Y.Z[A-Z]' (e.g. 2.3.1A)"
        skill_ver = ver_match.group(1)
        assert f"[v{skill_ver}]" in txt.split("---", 2)[1], (
            f"{p.parent.name} description must start with '[v{skill_ver}]' for Bionic UI visibility"
        )
        assert f"(v{skill_ver})" in txt, f"{p.parent.name} H1 heading must include '(v{skill_ver})'"
        assert f"`{skill_ver}`" in skills_readme and f"{p.parent.name}/SKILL.md" in skills_readme, (
            f"skills/README.md catalog must list {p.parent.name} at version `{skill_ver}`"
        )
        for idx in range(1, 6):
            assert f"## {idx}." in txt, f"{p.parent.name} missing section ## {idx}."
        assert "IFRS" in txt and "FASB" in txt and "SOX" in txt, f"{p.parent.name} missing IFRS/FASB/SOX compliance"

        mentioned_tools = {w for w in re.findall(r"\b[a-z][a-z0-9_]+\b", txt) if w in valid_tools}
        if p.parent.name != "boardroom-brief-generator":
            missing_in_compact = mentioned_tools - excel_compact_tools
            assert not missing_in_compact, f"{p.parent.name} uses tools missing from excel,compact: {missing_in_compact}"

        for fn_name, arg_str in re.findall(r"([a-z][a-z0-9_]+)\(([^)\n]*)\)", txt):
            if fn_name in valid_tools:
                sig_params = set(tools_map[fn_name].parameters.get("properties", {}).keys())
                for kw in re.findall(r"([a-z_]+)\s*=", arg_str):
                    assert kw in sig_params, f"{p.parent.name}: invalid param '{kw}' on tool '{fn_name}'"


@pytest.mark.asyncio
async def test_create_workbook_multi_sheet_names_one_call():
    """Verify create_workbook(sheet_names=[...]) creates all requested tabs in ONE call with no blank Sheet1."""
    res = await server.create_workbook(
        title="3-Statement Model",
        sheet_names=["Income Statement", "Balance Sheet", "Cash Flow"],
    )
    assert res["status"] if "status" in res else res["workbook_id"]
    assert res["worksheets"] == ["Income Statement", "Balance Sheet", "Cash Flow"]
    assert res["worksheet_count"] == 3
    assert res["active_sheet"] == "Income Statement"
    await server.close_workbook(res["workbook_id"])


@pytest.mark.asyncio
async def test_v22_efficiency_and_safety_improvements(monkeypatch, tmp_path: Path):
    """Verify sheet_names filtering in inspect_workbook, read_range trimming/max_cols, smart SOX audit, deferred disk save, and add_formula STRICT guard."""
    # 1. inspect_workbook(sheet_names=...) + read_range trimming & max_cols
    wb = await server.create_workbook(
        title="MultiTab Raw",
        sheet_names=["Cover", "Income Statement", "Balance Sheet", "Footnotes", "Charts"],
    )
    wb_id = wb["workbook_id"]
    await server.batch_write_ranges(
        wb_id,
        operations=[
            {
                "sheet_name": "Income Statement",
                "start_cell": "A1",
                "data": [
                    ["Line Item", "FY23", "FY24", "FY25"],
                    ["Revenue", 100, 110, 125],
                    ["COGS", 40, 44, 50],
                    ["Gross Profit", "=B2-B3", "=C2-C3", 75],  # D4: horizontal formula-row break!
                    ["Operating Expenses", 20, 22, 25],
                    ["EBITDA", "=B4-B5", "=C4-C5", "=D4-D5"],
                    ["Total Margin %", "=B6/B2", "=C6/C2", "=D6/D2"],
                    ["Net Income", "=B6*0.75", "=C6*0.75", "=D6*0.75"],
                ],
            },
            {
                "sheet_name": "Balance Sheet",
                "start_cell": "A1",
                "data": [["Item", "FY23", "FY24"], ["Cash", 50, 65]],
            },
        ],
    )

    insp = await server.inspect_workbook(
        wb_id, sheet_names=["income statement", "Balance Sheet", "NoSuchTab"]
    )
    assert insp["status"] == "success"
    assert insp["sheet_count"] == 5
    assert insp["inspected_sheet_count"] == 2
    assert [s["sheet_name"] for s in insp["sheets"]] == ["Income Statement", "Balance Sheet"]
    assert insp["missing_sheets"] == ["NoSuchTab"]
    assert insp["sheets"][0]["formulas"]["B4"] == "=B2-B3"

    # Explicit read_range trims trailing blanks and respects max_cols
    rr_trim = await server.read_range(wb_id, sheet_name="Balance Sheet", cell_range="A1:Z20")
    assert rr_trim["rows"] == 2 and rr_trim["columns"] == 3
    rr_cap = await server.read_range(wb_id, sheet_name="Income Statement", cell_range="A1:D8", max_cols=2)
    assert rr_cap["columns"] == 2 and rr_cap.get("truncated") is True

    # 2. High-precision SOX 404 suspicious_literals: flags D4 override, NOT Revenue/COGS/OpEx inputs
    aud = await server.audit_workbook_formulas(wb_id)
    assert aud["status"] == "success"
    flagged_cells = {item["cell"] for item in aud.get("suspicious_literals", [])}
    assert "D4" in flagged_cells
    for legit_input in ("B2", "C2", "D2", "B3", "C3", "D3", "B5", "C5", "D5"):
        assert legit_input not in flagged_cells, f"False positive on input cell {legit_input}"

    # 3. Deferred disk save: intermediate edits on an opened file must NOT mutate disk until save_workbook
    disk_file = tmp_path / "Deferred_Save_Test.xlsx"
    await server.save_workbook(wb_id, str(disk_file), format="xlsx")
    await server.close_workbook(wb_id)

    orig_bytes = disk_file.read_bytes()
    opened = await server.open_workbook(str(disk_file))
    wb2 = opened["workbook_id"]
    await server.write_cell(wb2, "Balance Sheet", "B2", 999999)
    assert disk_file.read_bytes() == orig_bytes, "Intermediate write mutated disk before save_workbook!"

    # 4. add_formula respects M365_MCP_FORMULA_GUARD=STRICT
    monkeypatch.delenv("OFFICE_MCP_FORMULA_GUARD", raising=False)
    monkeypatch.setenv("M365_MCP_FORMULA_GUARD", "STRICT")
    blk = await server.add_formula(wb2, "Income Statement", "B4", "=B2-B3+1")
    assert blk["status"] == "confirmation_required"
    assert blk["guard_mode"] == "strict"
    ok_fml = await server.add_formula(
        wb2, "Income Statement", "B4", "=B2-B3+1", allow_formula_overwrite=True
    )
    assert ok_fml["status"] == "success"
    assert ok_fml["formulas_updated"] == 1

    # Explicit save_workbook finally persists changes to disk
    await server.save_workbook(wb2, str(disk_file), format="xlsx")
    assert disk_file.read_bytes() != orig_bytes
    await server.close_workbook(wb2)


@pytest.mark.asyncio
async def test_optional_workbook_id_auto_resolution(tmp_path: Path):
    """Verify batch/save/close tools auto-resolve workbook_id when 1 workbook is active and reject 0 or >1."""
    # Ensure clean active workbook map
    for existing_id in list(server.excel.active_workbooks.keys()):
        await server.close_workbook(existing_id)

    # 0 active workbooks -> clear error
    err0 = await server.batch_write_ranges(
        operations=[{"sheet_name": "Sheet", "start_cell": "A1", "data": [["x"]]}]
    )
    assert err0["status"] == "error"
    assert "No active workbooks" in err0["error"]

    # 1 active workbook -> auto-resolves across batch_write_ranges, batch_format_workbook, save_workbook, close_workbook
    created = await server.create_workbook(title="AutoResolve", sheet_names=["Summary"])
    wb_id = created["workbook_id"]

    bw = await server.batch_write_ranges(
        operations=[{"sheet_name": "Summary", "start_cell": "A1", "data": [["Metric", "Val"], ["NPV", 12.5]]}]
    )
    assert bw["status"] == "success"
    assert bw["workbook_id"] == wb_id

    bf = await server.batch_format_workbook(
        operations=[{"sheet_name": "Summary", "freeze_panes": "B2"}]
    )
    assert bf["status"] == "success"
    assert bf["workbook_id"] == wb_id

    out_path = tmp_path / "auto_resolve.xlsx"
    sv = await server.save_workbook(file_path=str(out_path))
    assert sv["status"] == "success"
    assert out_path.exists()

    # 2 active workbooks -> omitting workbook_id must error clearly
    created2 = await server.create_workbook(title="SecondWB", sheet_names=["S1"])
    wb2_id = created2["workbook_id"]
    err2 = await server.batch_write_ranges(
        operations=[{"sheet_name": "S1", "start_cell": "A1", "data": [["y"]]}]
    )
    assert err2["status"] == "error"
    assert "multiple workbooks" in err2["error"].lower()

    await server.close_workbook(wb2_id)
    # Back to 1 active workbook -> close_workbook(None) succeeds
    cl = await server.close_workbook()
    assert cl["status"] == "success"
    assert cl["workbook_id"] == wb_id


@pytest.mark.asyncio
async def test_batch_write_ranges_accepts_range_cell_and_sheet_aliases():
    """Verify batch_write_ranges resolves range/cell_range/cell/sheet/values aliases instead of silently defaulting to A1."""
    created = await server.create_workbook(title="AliasTest", sheet_names=["Cash Flows"])
    wb_id = created["workbook_id"]
    try:
        res = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {"sheet": "Cash Flows", "range": "A1:B1", "values": [["Header", "Y1"]]},
                {"sheet_name": "Cash Flows", "cell_range": "A5:B5", "data": [["Rev", 100]]},
                {"sheet_name": "Cash Flows", "cell": "'Cash Flows'!C10", "data": [[250]]},
            ],
        )
        assert res["status"] == "success"
        ranges = [op["range"] for op in res["results"]]
        assert ranges == ["A1:B1", "A5:B5", "C10:C10"]

        r = await server.read_range(wb_id, "Cash Flows", "A1:C10")
        assert r["data"][0][:2] == ["Header", "Y1"]
        assert r["data"][4][:2] == ["Rev", 100]
        assert r["data"][9][2] == 250

        # Auto-orient transposed Nx1 data into explicit 1xN horizontal range, 1xN into Nx1 vertical, and flat 1D list
        orient = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {"sheet_name": "Cash Flows", "range": "B12:E12", "data": [[0.8], [0.9], [1.0], [1.1]]},
                {"sheet_name": "Cash Flows", "range": "A13:A15", "data": [["WACC1", "WACC2", "WACC3"]]},
                {"sheet_name": "Cash Flows", "range": "B13:D13", "data": [10, 20, 30]},
            ],
        )
        assert orient["status"] == "success"
        r2 = await server.read_range(wb_id, "Cash Flows", "A12:E15")
        assert r2["data"][0][1:5] == [0.8, 0.9, 1.0, 1.1]
        assert [r2["data"][1][0], r2["data"][2][0], r2["data"][3][0]] == ["WACC1", "WACC2", "WACC3"]
        assert r2["data"][1][1:4] == [10, 20, 30]
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_live_open_blocks_internal_temp_uuid_files():
    """Verify excel_live_open and word_live_open reject internal /tmp and macOS /var/folders temp UUID paths."""
    for p in [
        "/tmp/m365-mac-mcp/excel/2dd3c267-a440-4415-ac4a-86d21acc80a1.xlsx",
        str(server.excel.temp_dir / "ea1bbebc-87b9-4dfa-8d7c-b61c6924305a.xlsx"),
    ]:
        res_xl = await server.excel_live_open(p)
        assert res_xl["status"] == "error"
        assert "Grant Access" in res_xl["error"]
        assert "save_workbook" in res_xl["error"]

    for p in [
        "/tmp/m365-mac-mcp/word/2dd3c267-a440-4415-ac4a-86d21acc80a1.docx",
        str(server.word.temp_dir / "ea1bbebc-87b9-4dfa-8d7c-b61c6924305a.docx"),
    ]:
        res_wd = await server.word_live_open(p)
        assert res_wd["status"] == "error"
        assert "Grant Access" in res_wd["error"]
        assert "save_document" in res_wd["error"]


@pytest.mark.asyncio
async def test_live_close_bool_coercion_read_range_alias_and_live_error_surfacing(monkeypatch):
    """Verify excel_live_close/word_live_close accept bool save_changes, read_range accepts range=, and live batch reads surface #VALUE!/#DIV/0!."""
    # 1. read_range accepts range= alias + auto-resolves workbook_id
    for existing_id in list(server.excel.active_workbooks.keys()):
        await server.close_workbook(existing_id)
    created = await server.create_workbook(title="ErrSurface", sheet_names=["S1"])
    wb_id = created["workbook_id"]
    try:
        await server.batch_write_ranges(
            operations=[{"sheet_name": "S1", "start_cell": "A1", "data": [["Rev", 100], ["Cost", 60]]}]
        )
        r = await server.read_range(sheet_name="S1", range="A1:B2")
        assert r["status"] == "success"
        assert r["data"] == [["Rev", 100], ["Cost", 60]]
    finally:
        await server.close_workbook(wb_id)

    # 2. excel_live_close & word_live_close coerce bool save_changes -> 'yes'/'no'
    seen: list = []

    async def fake_close_excel(workbook_name=None, save_changes="no"):
        seen.append(("excel", workbook_name, save_changes))
        return {"status": "success", "save_changes": save_changes}

    async def fake_close_word(document_name=None, save_changes="no"):
        seen.append(("word", document_name, save_changes))
        return {"status": "success", "save_changes": save_changes}

    monkeypatch.setattr(server.applescript, "close_excel_workbook", fake_close_excel)
    monkeypatch.setattr(server.applescript, "close_word_document", fake_close_word)

    assert (await server.excel_live_close("Model.xlsx", save_changes=False))["save_changes"] == "no"
    assert (await server.excel_live_close("Model.xlsx", save_changes=True))["save_changes"] == "yes"
    assert (await server.word_live_close("Memo.docx", save_changes=False))["save_changes"] == "no"
    assert seen == [("excel", "Model.xlsx", "no"), ("excel", "Model.xlsx", "yes"), ("word", "Memo.docx", "no")]

    # 3. batch_read_live_excel_ranges surfaces #VALUE! / #DIV/0! from string value when value is blank
    bridge = AppleScriptBridge()
    us, rs, gs, fs = "\x1f", "\x1e", "\x1d", "\x1c"
    fake_raw = (
        f"Model.xlsx{fs}"
        f"Sheet1{us}A1:B2{gs}"
        f"10{us}{rs}{us}20{gs}"
        f"{gs}"
        f"10{us}#VALUE!{rs}#DIV/0!{us}20"
    )

    async def fake_exec(script, args=None, timeout=None):
        assert "string value of rng" in script
        assert "calculate" in script
        return fake_raw

    monkeypatch.setattr(bridge, "execute_applescript", fake_exec)
    out = await bridge.batch_read_live_excel_ranges(
        ranges=[{"sheet_name": "Sheet1", "range_ref": "A1:B2"}],
        workbook_name="Model.xlsx",
    )
    assert out["status"] == "success"
    assert out["results"][0]["values"] == [["10", "#VALUE!"], ["#DIV/0!", "20"]]


@pytest.mark.asyncio
async def test_live_read_warns_when_in_memory_workbook_has_unsaved_edits(tmp_path: Path, monkeypatch):
    """Verify excel_live_batch_read_ranges warns ONLY when a matching in-memory workbook has unsaved dirty edits."""
    for existing_id in list(server.excel.active_workbooks.keys()):
        await server.close_workbook(existing_id)

    target = tmp_path / "DirtySync.xlsx"
    created = await server.create_workbook(title="DirtySync", sheet_names=["S1"])
    wb_id = created["workbook_id"]
    try:
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "S1", "start_cell": "A1", "data": [["Val", 1]]}],
        )
        await server.save_workbook(workbook_id=wb_id, file_path=str(target))

        async def fake_batch_read(ranges, workbook_name=None, include_formulas=False):
            return {"status": "success", "workbook_name": "DirtySync.xlsx", "ranges_read": 1, "results": []}

        monkeypatch.setattr(server.applescript, "batch_read_live_excel_ranges", fake_batch_read)

        # Clean (just saved) -> NO warning
        clean_res = await server.excel_live_batch_read_ranges(
            ranges=[{"sheet_name": "S1", "range_ref": "A1:B1"}], workbook_name="DirtySync.xlsx"
        )
        assert "warning" not in clean_res

        # Mutate in-memory without calling save_workbook -> warning MUST appear
        await server.write_cell(workbook_id=wb_id, sheet_name="S1", cell="B1", value=99)
        dirty_res = await server.excel_live_batch_read_ranges(
            ranges=[{"sheet_name": "S1", "range_ref": "A1:B1"}], workbook_name="DirtySync.xlsx"
        )
        assert "warning" in dirty_res
        assert "save_workbook" in dirty_res["warning"]
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_opt_in_debug_log_writes_jsonl(tmp_path: Path, monkeypatch):
    """Verify M365_MCP_DEBUG_LOG writes structured JSONL entries when set and is a no-op when unset."""
    log_file = tmp_path / "mcp_debug_private.jsonl"
    monkeypatch.setenv("M365_MCP_DEBUG_LOG", str(log_file))

    wrapped = server._wrap_compact_json(server.check_office_status)
    out = await wrapped()
    assert isinstance(out, str)
    assert log_file.exists()
    lines = [json.loads(line) for line in log_file.read_text(encoding="utf-8").splitlines() if line.strip()]
    assert len(lines) == 1
    assert lines[0]["tool"] == "check_office_status"
    assert lines[0]["ms"] >= 0
    assert "response_preview" in lines[0]


@pytest.mark.asyncio
async def test_batch_write_ranges_fill_formula_1d_and_2d_with_formula_guard():
    """Verify batch_write_ranges `fill_formula` translates relative refs across 1D rows, 1D cols, and 2D grids and obeys Formula Guard."""
    created = await server.create_workbook(title="FillFormulaTest", sheet_names=["Cash Flows", "Sens"])
    wb_id = created["workbook_id"]
    try:
        res = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {"sheet_name": "Cash Flows", "range": "A1:D3", "data": [
                    ["Item", "Y1", "Y2", "Y3"],
                    ["Rev", 10, 12, 15],
                    ["Margin", 0.4, 0.4, 0.4],
                ]},
                {"sheet_name": "Cash Flows", "range": "A4", "data": [["EBITDA"]]},
                {"sheet_name": "Cash Flows", "range": "B4:D4", "fill_formula": "=B2*B3"},
                {"sheet_name": "Sens", "range": "A1:D3", "data": [
                    ["WACC/CapEx", 20, 25, 30],
                    [0.09, None, None, None],
                    [0.10, None, None, None],
                ]},
                {
                    "sheet_name": "Sens",
                    "range": "B2:D3",
                    "fill_formula": "=NPV($A2,'Cash Flows'!$B$4:$D$4)-B$1",
                },
            ],
        )
        assert res["status"] == "success"

        cf = await server.read_range(wb_id, "Cash Flows", "B4:D4", include_formulas=True)
        assert cf["data"][0] == ["=B2*B3", "=C2*C3", "=D2*D3"]

        sens = await server.read_range(wb_id, "Sens", "B2:D3", include_formulas=True)
        assert sens["data"] == [
            [
                "=NPV($A2,'Cash Flows'!$B$4:$D$4)-B$1",
                "=NPV($A2,'Cash Flows'!$B$4:$D$4)-C$1",
                "=NPV($A2,'Cash Flows'!$B$4:$D$4)-D$1",
            ],
            [
                "=NPV($A3,'Cash Flows'!$B$4:$D$4)-B$1",
                "=NPV($A3,'Cash Flows'!$B$4:$D$4)-C$1",
                "=NPV($A3,'Cash Flows'!$B$4:$D$4)-D$1",
            ],
        ]

        # Formula Guard still protects fill_formula cells from literal overwrite
        blk = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "Cash Flows", "range": "B4:D4", "data": [[4, 5, 6]]}],
        )
        assert blk["status"] == "confirmation_required"
        assert blk["blocked_count"] == 3
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_column_widths_list_table_header_guard_and_uncached_formula_warning(tmp_path: Path):
    """Verify column_widths list fallback, Excel Table header repair before save,
    uncached formula anti-loop warning on read_range, and live timeout hint.
    """
    created = await server.create_workbook(title="GemmaHardening", sheet_names=["Sales"])
    wb_id = created["workbook_id"]
    out_path = tmp_path / "gemma_hardening.xlsx"
    try:
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {
                    "sheet_name": "Sales",
                    "range": "A1:C3",
                    "values": [
                        ["Region", "Q1", "Total"],
                        ["North", 100, "=SUM(B2:B2)"],
                        ["South", 80, "=SUM(B3:B3)"],
                    ],
                }
            ],
        )
        await server.add_excel_table(wb_id, "Sales", "SalesTbl", "A1:C3")

        # 1. column_widths passed as a list maps sequentially from range start col (or A)
        bf = await server.batch_format_workbook(
            workbook_id=wb_id,
            operations=[
                {"sheet_name": "Sales", "range": "A1:C1", "column_widths": [18, 12, 14]}
            ],
        )
        assert bf["status"] == "success"
        ws = server.excel.active_workbooks[wb_id]["workbook_object"]["Sales"]
        assert ws.column_dimensions["A"].width == 18
        assert ws.column_dimensions["B"].width == 12
        assert ws.column_dimensions["C"].width == 14

        # 2. Wipe table header A1 to None -> save_workbook must auto-repair so Excel never rejects (-50)
        await server.write_cell(wb_id, "Sales", "A1", None)
        sv = await server.save_workbook(wb_id, str(out_path))
        assert sv["status"] == "success"
        assert ws["A1"].value == "Column1"

        # 3. read_range(evaluate_formulas=True) on uncached formulas returns anti-loop warning
        rr = await server.read_range(wb_id, "Sales", "A1:C3", evaluate_formulas=True)
        assert rr["status"] == "success"
        assert rr.get("uncached_formula_count", 0) == 2
        assert "do NOT retry read_range" in rr.get("warning", "")

        # 4. excel_live_* timeout hint instructs stopping for Excel dialog, NOT file-level fallback
        err = server._error_response("excel_live_open", TimeoutError("timed out"))
        assert "Grant Access" in err.get("hint", "")
        assert "do NOT fall back to file-level read_range" in err.get("hint", "")

        # 5. batch_write_ranges accepts scalar 'value' and redirects formatting-only ops
        bw_scalar = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "Sales", "cell": "E1", "value": "Base"}],
        )
        assert bw_scalar["status"] == "success"
        assert ws["E1"].value == "Base"

        bw_fmt_err = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "Sales", "range": "B2:C3", "number_format": "$#,##0"}],
        )
        assert bw_fmt_err["status"] == "error"
        assert "batch_format_workbook" in bw_fmt_err["error"]

        # 6. add_excel_table accepts 'range' alias and names required params on empty range
        tbl_err = await server.add_excel_table(workbook_id=wb_id, sheet_name="Sales")
        assert tbl_err["status"] == "error"
        assert "data_range" in tbl_err["error"]

        # 7. add_data_validation auto-quotes comma-separated literal lists for valid OOXML
        dv = await server.add_data_validation(
            workbook_id=wb_id,
            sheet_name="Sales",
            cell_range="E2",
            validation_type="list",
            formula1="Base,Upside,Downside",
        )
        assert dv["status"] == "success"
        assert dv["formula1"] == '"Base,Upside,Downside"'

        # 8. Scalar 'values' in batch_write_ranges and 'COLUMN_CLUSTERED' chart alias
        bw_scalar_vals = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "Sales", "range": "E3", "values": "Upside"}],
        )
        assert bw_scalar_vals["status"] == "success"
        assert ws["E3"].value == "Upside"

        ch = await server.create_chart(
            workbook_id=wb_id,
            sheet_name="Sales",
            chart_type="COLUMN_CLUSTERED",
            data_range="A1:C3",
            chart_title="Alias Test",
            position="G1",
        )
        assert ch["status"] == "success"
        # First column has text labels ("North", "South") -> 2 numeric series (Q1, Total), not 3
        assert len(ws._charts[-1].series) == 2

        # 9. Singular column_width + range alias in batch_format_workbook
        bf_single_w = await server.batch_format_workbook(
            workbook_id=wb_id,
            operations=[{"sheet": "Sales", "range": "B:C", "column_width": 16}],
        )
        assert bf_single_w["status"] == "success"
        assert ws.column_dimensions["B"].width == 16
        assert ws.column_dimensions["C"].width == 16

        # 10. batch_write_ranges parses 'Sheet!A1' prefixes AND top-level sheet_name default
        await server.add_worksheet(wb_id, "Project Cash Flows")
        bw_prefixed = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {"range": "Project Cash Flows!A1", "value": "PCF Header"},
                {"range": "'Project Cash Flows'!B1", "value": 123},
            ],
        )
        assert bw_prefixed["status"] == "success"
        assert bw_prefixed["results"][0]["sheet_name"] == "Project Cash Flows"
        assert bw_prefixed["results"][1]["sheet_name"] == "Project Cash Flows"
        bw_default = await server.batch_write_ranges(
            workbook_id=wb_id,
            sheet_name="Project Cash Flows",
            operations=[{"range": "C1", "value": 456}],
        )
        assert bw_default["status"] == "success"
        assert bw_default["results"][0]["sheet_name"] == "Project Cash Flows"
        pcf_ws = server.excel.active_workbooks[wb_id]["workbook_object"]["Project Cash Flows"]
        assert pcf_ws["A1"].value == "PCF Header"
        assert pcf_ws["B1"].value == 123
        assert pcf_ws["C1"].value == 456
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_corporate_template_env_vars_and_ppt_widescreen_double_bullet_fix(tmp_path, monkeypatch):
    """Verify M365_MCP_TEMPLATES_DIR, M365_MCP_PPT_TEMPLATE, M365_MCP_WORD_TEMPLATE, 16:9 default & double-bullet strip."""
    from docx import Document as DocxDocument
    from pptx import Presentation as PptxPresentation
    from pptx.util import Inches
    from controllers.powerpoint_controller import PowerPointController
    from controllers.word_controller import WordController

    tpl_dir = tmp_path / "Corporate_Templates"
    tpl_dir.mkdir()

    # Create sample corporate Word & PPT templates
    word_tpl = tpl_dir / "Corp_Memo.docx"
    d = DocxDocument()
    d.add_paragraph("CORPORATE HEADER TEMPLATE")
    d.save(str(word_tpl))

    ppt_tpl = tpl_dir / "Corp_Deck.pptx"
    p = PptxPresentation()
    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)
    s0 = p.slides.add_slide(p.slide_layouts[0])
    s0.shapes.title.text = "Template Title"
    p.save(str(ppt_tpl))

    # 1. Without env vars: PPT defaults to 16:9 widescreen, registers first_slide_id,
    #    doesn't overwrite title on placeholder='content', and strips duplicate bullets.
    monkeypatch.delenv("M365_MCP_TEMPLATES_DIR", raising=False)
    monkeypatch.delenv("M365_MCP_PPT_TEMPLATE", raising=False)
    monkeypatch.delenv("M365_MCP_WORD_TEMPLATE", raising=False)

    ppt_ctrl = PowerPointController()
    meta = await ppt_ctrl.create_presentation(title="Q4 Overview")
    assert meta["first_slide_id"]
    assert "template_used" not in meta
    prs = ppt_ctrl.active_presentations[meta["presentation_id"]]["pptx_object"]
    assert abs(prs.slide_width.inches - 13.333) < 0.01
    assert abs(prs.slide_height.inches - 7.5) < 0.01

    s2 = await ppt_ctrl.add_slide(meta["presentation_id"], layout="Title and Content")
    await ppt_ctrl.add_text(s2["slide_id"], "Key Regional Totals", placeholder="title")
    await ppt_ctrl.add_text(s2["slide_id"], "• North: $510M\n- South: $375M\n* West: $320M", placeholder="content")
    slide2_obj = ppt_ctrl.active_presentations[meta["presentation_id"]]["slides"][s2["slide_id"]]["slide_object"]
    assert slide2_obj.shapes.title.text == "Key Regional Totals"
    body_text = [sh.text_frame.text for sh in slide2_obj.placeholders if sh != slide2_obj.shapes.title][0]
    assert body_text == "North: $510M\nSouth: $375M\nWest: $320M"

    # 2. With M365_MCP_TEMPLATES_DIR + M365_MCP_PPT_TEMPLATE + M365_MCP_WORD_TEMPLATE
    monkeypatch.setenv("M365_MCP_TEMPLATES_DIR", str(tpl_dir))
    monkeypatch.setenv("M365_MCP_PPT_TEMPLATE", "Corp_Deck.pptx")
    monkeypatch.setenv("M365_MCP_WORD_TEMPLATE", "Corp_Memo.docx")

    meta_tpl = await ppt_ctrl.create_presentation(title="Board Deck")
    assert meta_tpl["template_used"] == str(ppt_tpl.resolve())

    word_ctrl = WordController()
    w_meta = await word_ctrl.create_document(title="CFO Brief")
    assert w_meta["template_used"] == str(word_tpl.resolve())
    w_read = await word_ctrl.read_document(w_meta["document_id"], include_full_text=True)
    assert "CORPORATE HEADER TEMPLATE" in w_read["full_text"]
    await word_ctrl.close_document(w_meta["document_id"])

    # 3. Malformed numeric env vars safely fall back to ServerConfig defaults
    from utils.config import Config
    monkeypatch.setenv("M365_MCP_APPLESCRIPT_TIMEOUT", "30s")
    monkeypatch.setenv("M365_MCP_MAX_WORKBOOKS", "not_a_number")
    cfg = Config(config_file=str(tmp_path / "nonexistent_config.json"))
    assert cfg.settings.applescript_timeout == 30.0
    assert cfg.settings.max_workbooks == 20


@pytest.mark.asyncio
async def test_universal_write_coercion_cell_map_range_mismatch_and_live_save_guard(tmp_path, monkeypatch):
    """Verify Fixes A, B, C: formatted-number & empty-blank coercion, cell_map echo, range spill warning, and excel_live_save dirty guard."""
    wb = await server.create_workbook(
        title="Universal Fixes Test",
        sheet_names=["Assumptions", "Cash Flows"],
    )
    wb_id = wb["workbook_id"]
    out_xlsx = tmp_path / "CapEx_Universal.xlsx"
    try:
        # 1. Fix B + Fix A (cell_map): Formatted strings auto-coerced to numeric + formats, "" -> None, cell_map returned
        res_assump = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{
                "sheet_name": "Assumptions",
                "range": "A1:B9",
                "data": [
                    ["Parameter", "Value"],
                    ["Project Life", "6"],
                    ["Year-0 CapEx", "$25,000,000"],
                    ["Salvage Adjustment", "($4,000,000)"],
                    ["WACC", "10.5%"],
                    ["Tax Rate", "25%"],
                    ["ZIP Code", "00123"],
                    ["Unit Label", "$510M"],
                    ["Blank Note", ""],
                ],
            }],
        )
        assert res_assump["status"] == "success"
        ws_a = server.excel.active_workbooks[wb_id]["workbook_object"]["Assumptions"]
        assert ws_a["B2"].value == 6 and isinstance(ws_a["B2"].value, int)
        assert ws_a["B3"].value == 25000000 and "$" in ws_a["B3"].number_format
        assert ws_a["B4"].value == -4000000 and "$" in ws_a["B4"].number_format
        assert ws_a["B5"].value == 0.105 and ws_a["B5"].number_format == "0.0%"
        assert ws_a["B6"].value == 0.25 and ws_a["B6"].number_format == "0.0%"
        # Preserves leading-zero IDs and unit-suffixed strings as text
        assert ws_a["B7"].value == "00123" and isinstance(ws_a["B7"].value, str)
        assert ws_a["B8"].value == "$510M" and isinstance(ws_a["B8"].value, str)
        # Empty string normalized to true blank None (prevents #VALUE! in arithmetic)
        assert ws_a["B9"].value is None
        # Ground-truth cell_map echoed for compact Assumptions sheet
        cmap = res_assump["results"][0].get("cell_map") or {}
        assert "B3" in cmap and "Year-0 CapEx" in cmap["B3"]
        assert "B5" in cmap and "WACC" in cmap["B5"]

        # 2. Fix A (Range-shape mismatch warning): 8 cols written to B1:H2 (7 cols) -> immediate warning
        res_spill = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{
                "sheet_name": "Cash Flows",
                "range": "B1:H2",
                "data": [
                    ["Year", 0, 1, 2, 3, 4, 5, 6],
                    ["Revenue", 0, 100, 112, 125, 140, 157, 176],
                ],
            }],
        )
        assert res_spill["status"] == "success"
        assert res_spill["results"][0]["range"] == "B1:I2"
        assert res_spill.get("warnings"), "Expected range shape mismatch warning on B1:H2 -> B1:I2"
        assert "B1:I2" in res_spill["warnings"][0]

        # 3. Fix C: excel_live_save blocks when matching in-memory workbook has unsaved dirty edits
        monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
        await server.save_workbook(wb_id, file_path=str(out_xlsx))
        # Make an unsaved in-memory edit
        await server.write_cell(wb_id, "Assumptions", "B2", "7")
        assert server.excel.active_workbooks[wb_id]["dirty"] is True

        live_save_res = await server.excel_live_save(
            workbook_name=out_xlsx.name,
            file_path=str(out_xlsx),
        )
        assert live_save_res["status"] == "error"
        assert live_save_res["error_type"] == "UnsavedInMemoryEditsError"
        assert "save_workbook" in live_save_res["hint"]
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_hardening_status_skill_loader_save_linter_and_fill_formulas_expansion(tmp_path, monkeypatch):
    """Verify Upgrades #1, #2, #3: check_office_status(skill_name), save_workbook linter, fill_formulas expansion & cell_formats redirect."""
    # 1. Upgrade #1: check_office_status returns registered_tools + live finance_skills catalog and loads full SKILL.md on demand
    st_menu = await server.check_office_status()
    assert st_menu["status"] == "running"
    assert "create_workbook" in st_menu["profile"]["registered_tools"]
    assert len(st_menu["finance_skills"]["available"]) >= 9
    assert any(s["name"] == "capex-project-analyzer" and s["version"] for s in st_menu["finance_skills"]["available"])

    st_loaded = await server.check_office_status(skill_name="capex-project-analyzer")
    assert st_loaded["loaded_skill"]["name"] == "capex-project-analyzer"
    assert st_loaded["loaded_skill"]["version"]
    assert "Project Cash Flows" in st_loaded["loaded_skill"]["content"]

    st_missing = await server.check_office_status(skill_name="nonexistent-skill-xyz")
    assert "skill_error" in st_missing and len(st_missing["finance_skills"]) >= 9

    # 2. Upgrade #3: plural fill_formulas expands cleanly without wiping formulas to None, and cell_formats redirects cleanly
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
    wb = await server.create_workbook(title="Hardening Test", sheet_names=["Assumptions", "EmptyTab"])
    wb_id = wb["workbook_id"]
    out_xlsx = tmp_path / "Hardening_Linter.xlsx"
    try:
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{
                "sheet_name": "Assumptions",
                "range": "A1:B6",
                "data": [
                    ["Item", "Value"],
                    ["CapEx", 25000000],
                    ["Rev", 10000000],
                    ["Growth", 0.12],
                    ["WACC", 0.105],
                    ["Calc", "=B2+B3"],
                ],
            }],
        )
        # Plural fill_formulas alongside dummy blank values should expand cleanly to Formula -> Formula update
        ff_res = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{
                "range": "Assumptions!B6",
                "values": [[""]],
                "fill_formulas": [{"range": "Assumptions!B6", "formula": "=B2*B4"}],
            }],
        )
        assert ff_res["status"] == "success"
        ws = server.excel.active_workbooks[wb_id]["workbook_object"]["Assumptions"]
        assert ws["B6"].value == "=B2*B4"

        # Formatting-only op with 'cell_formats' redirects cleanly to batch_format_workbook
        cf_res = await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"range": "Assumptions!B2", "cell_formats": [{"number_format": "$#,##0"}]}],
        )
        assert cf_res["status"] == "error"
        assert "batch_format_workbook" in cf_res["error"]

        # 3. Upgrade #2: save_workbook compile-time linter flags empty tab + unformatted numbers in 'General'
        save_res = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert save_res["status"] == "success"
        warns = save_res.get("sanity_warnings") or []
        assert any("EmptyTab" in w and "empty" in w.lower() for w in warns)
        assert any("General" in w for w in warns)

        # After formatting and populating EmptyTab, sanity_warnings disappears completely (0 extra tokens when clean)
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "EmptyTab", "range": "A1", "value": "Populated"}],
        )
        await server.batch_format_workbook(
            workbook_id=wb_id,
            operations=[{
                "sheet_name": "Assumptions",
                "number_formats": {"B2:B3": "$#,##0", "B4:B5": "0.0%"},
            }],
        )
        save_clean = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert save_clean["status"] == "success"
        assert "sanity_warnings" not in save_clean
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_audit_workbook_formulas_detects_silent_wrong_month_r1c1_inconsistency(tmp_path, monkeypatch):
    """ModeLoop audit test: catch a formula returning a valid number (0 Excel errors) that references the wrong month."""
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
    wb = await server.create_workbook(title="Monthly P&L Audit", sheet_names=["Monthly P&L"])
    wb_id = wb["workbook_id"]
    out_xlsx = tmp_path / "Monthly_WrongMonth.xlsx"
    try:
        # 6 months (Jan..Jun in B..G) + Total column (H):
        # Row 2: Revenue, Row 3: Gross Profit, Row 4: Gross Margin (=GP/Rev)
        # Plant a SILENT wrong-month logic error in Apr (E4): =E3/D2 (divides Apr GP by Mar Revenue -> valid number, wrong logic!)
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{
                "sheet_name": "Monthly P&L",
                "range": "A1:H4",
                "data": [
                    ["Metric", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Total"],
                    ["Revenue", 1000, 1100, 1200, 1500, 1400, 1600, "=SUM(B2:G2)"],
                    ["Gross Profit", 400, 440, 480, 600, 560, 640, "=SUM(B3:G3)"],
                    [
                        "Gross Margin",
                        "=B3/B2",
                        "=C3/C2",
                        "=D3/D2",
                        "=E3/D2",  # <-- SILENT WRONG-MONTH BUG (0.50 instead of 0.40, zero Excel #ERR!)
                        "=F3/F2",
                        "=G3/G2",
                        "=H3/H2",
                    ],
                ],
            }],
        )
        await server.save_workbook(wb_id, file_path=str(out_xlsx))

        # 1. Audit BEFORE fix: zero Excel errors, yet clean=False and E4 caught deterministically with exact expected_formula!
        audit_before = await server.audit_workbook_formulas(workbook_id=wb_id)
        assert audit_before["error_count"] == 0
        assert audit_before["row_inconsistency_count"] == 1
        assert audit_before["clean"] is False
        inc = audit_before["row_inconsistencies"][0]
        assert inc["sheet_name"] == "Monthly P&L"
        assert inc["cell"] == "E4"
        assert inc["col_header"] == "Apr"
        assert inc["formula"] == "=E3/D2"
        assert inc["expected_formula"] == "=E3/E2"
        assert inc["r1c1_signature"] == "=R[-1]C/R[-2]C[-1]"
        assert inc["dominant_r1c1_pattern"] == "=R[-1]C/R[-2]C"

        # Independent ground-truth check (Apr GP 600 / Apr Rev 1500 = 0.40 vs buggy 600 / Mar Rev 1200 = 0.50)
        expected_apr_margin = 600 / 1500
        buggy_apr_margin = 600 / 1200
        assert abs(expected_apr_margin - 0.40) < 1e-9
        assert abs(buggy_apr_margin - 0.50) < 1e-9

        # 2. Apply fix (=E3/E2) and re-audit -> 100% clean (and Total column H is NOT falsely flagged!)
        await server.write_cell(wb_id, "Monthly P&L", "E4", inc["expected_formula"])
        audit_after = await server.audit_workbook_formulas(workbook_id=wb_id)
        assert audit_after["error_count"] == 0
        assert audit_after["row_inconsistency_count"] == 0
        assert audit_after["clean"] is True
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_project_dir_write_guard_and_audit_circular_npv_commutative_checks(tmp_path, monkeypatch):
    """Verify MCP server blocks writes inside its own project folder and catches circular/1-arg NPV without commutative/Year-0 noise."""
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
    project_root = Path(__file__).resolve().parent.parent

    # 1. Excel save_workbook blocks both bare relative filename (when CWD is project root) and explicit project-dir paths
    monkeypatch.chdir(project_root)
    wb = await server.create_workbook(title="GuardTest", sheet_names=["Project Cash Flows", "Executive Summary"])
    wb_id = wb["workbook_id"]
    try:
        res_rel = await server.save_workbook(wb_id, file_path="Accidental_In_Project.xlsx")
        assert res_rel["status"] == "error"
        assert res_rel["error_type"] == "PermissionError"
        assert "project folder" in res_rel["error"].lower()
        assert "desktop" in res_rel.get("hint", "").lower()
        assert not (project_root / "Accidental_In_Project.xlsx").exists()

        res_abs = await server.save_workbook(wb_id, file_path=str(project_root / "nested" / "Accidental.xlsx"))
        assert res_abs["status"] == "error"
        assert res_abs["error_type"] == "PermissionError"
        assert not (project_root / "nested").exists()

        # Word & PowerPoint also block writes into project root
        doc = await server.create_document(title="GuardDoc")
        doc_id = doc["document_id"]
        try:
            d_res = await server.save_document(doc_id, file_path="Accidental_Doc.docx")
            assert d_res["status"] == "error"
            assert d_res["error_type"] == "PermissionError"
            assert not (project_root / "Accidental_Doc.docx").exists()
        finally:
            await server.close_document(doc_id)

        prs = await server.create_presentation(title="GuardPPT")
        prs_id = prs["presentation_id"]
        p_res = await server.save_presentation(prs_id, file_path="Accidental_Deck.pptx")
        assert p_res["status"] == "error"
        assert p_res["error_type"] == "PermissionError"
        assert not (project_root / "Accidental_Deck.pptx").exists()

        # 2. Audit catches circular self-ref & 1-arg NPV in error_cells, ignores commutative '+' order & Year-0 zeros
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {
                    "sheet_name": "Project Cash Flows",
                    "range": "A1:E4",
                    "data": [
                        ["Metric", 0, 1, 2, 3],
                        ["UFCF", -25, 10, 12, 14],
                        ["Cum. UFCF", "=B2", "=B3+C2", "=D2+C3", "=E2+D3"],  # commutative =B3+C2 vs =D2+C3 -> 0 R1C1 noise!
                        ["Cum. PV", "=B2", "=B4+C2", "=D4+E2", "=E4+F2"],    # D4/E4 self-reference -> #CIRCULAR!
                    ],
                },
                {
                    "sheet_name": "Project Cash Flows",
                    "range": "A5:E5",
                    "data": [["Revenue", 0, 10, "=C5*1.1", "=D5*1.1"]],      # Year-0 literal 0 -> 0 suspicious_literals noise!
                },
                {
                    "sheet_name": "Executive Summary",
                    "range": "A1:B5",
                    "data": [
                        ["KPI", "Value"],
                        ["PI", "=(NPV(B1)-25)/-25"],                         # 1-arg NPV() -> #SYNTAX!
                        ["Spacer1", 1],
                        ["Spacer2", 2],
                        ["CrossSheetPI", "=(B2+'Project Cash Flows'!$B$5)/'Project Cash Flows'!$B$5"],  # cross-sheet $B$5 on B5 must NOT be circular!
                    ],
                },
            ],
        )

        # Save to safe tmp_path -> succeeds and surfaces save-time sanity warnings for circular + 1-arg NPV
        out_xlsx = tmp_path / "Safe_Model.xlsx"
        saved = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert saved["status"] == "success"
        joined_warns = " ".join(saved.get("sanity_warnings", []))
        assert "circular" in joined_warns.lower()
        assert "npv" in joined_warns.lower()

        audit = await server.audit_workbook_formulas(workbook_id=wb_id)
        assert audit["clean"] is False
        err_codes = {(e["cell"], e["error"]) for e in audit.get("error_cells", [])}
        assert ("D4", "#CIRCULAR!") in err_codes
        assert ("E4", "#CIRCULAR!") in err_codes
        assert ("B2", "#SYNTAX!") in err_codes
        assert ("B5", "#CIRCULAR!") not in err_codes
        # Commutative Cum. UFCF row 3 must NOT be flagged as an R1C1 inconsistency
        inc_rows = {(i["sheet_name"], i["row_label"]) for i in audit.get("row_inconsistencies", [])}
        assert ("Project Cash Flows", "Cum. UFCF") not in inc_rows
        # Year-0 Revenue=0 must NOT be flagged as a suspicious literal
        lit_cells = {(s["sheet_name"], s["cell"]) for s in audit.get("suspicious_literals", [])}
        assert ("Project Cash Flows", "B5") not in lit_cells
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_default_output_dir_env_var_and_explicit_path_precedence(tmp_path, monkeypatch):
    """Verify exact 4-tier path order: (1) Explicit user path > (2) Client/Bionic Project Folder > (3) M365_MCP_OUTPUT_DIR; (4) MCP project root ALWAYS blocked."""
    from utils.validators import set_client_project_dir

    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
    project_root = Path(__file__).resolve().parent.parent
    monkeypatch.chdir(project_root)
    set_client_project_dir(None)

    try:
        # 1. When unset, check_office_status omits active_project_dir & default_output_dir
        monkeypatch.delenv("M365_MCP_OUTPUT_DIR", raising=False)
        monkeypatch.delenv("OFFICE365_MCP_OUTPUT_DIR", raising=False)
        status_unset = await server.check_office_status()
        assert "default_output_dir" not in status_unset["profile"]
        assert "active_project_dir" not in status_unset["profile"]

        # 2. Setting M365_MCP_OUTPUT_DIR or working_dir inside MCP project root is ignored/blocked
        monkeypatch.setenv("M365_MCP_OUTPUT_DIR", str(project_root))
        status_bad = await server.check_office_status(working_dir=str(project_root))
        assert "default_output_dir" not in status_bad["profile"]
        assert "active_project_dir" not in status_bad["profile"]

        # 3. Priority #3 (M365_MCP_OUTPUT_DIR) resolves bare relative filenames when no Project Folder is set
        env_fallback_dir = tmp_path / "GlobalEnvFallback"
        bionic_proj_a = tmp_path / "BionicProjectA"
        bionic_proj_b = tmp_path / "BionicProjectB"
        user_explicit_dir = tmp_path / "UserExplicitFolder"
        monkeypatch.setenv("M365_MCP_OUTPUT_DIR", str(env_fallback_dir))

        status_env = await server.check_office_status()
        assert status_env["profile"]["default_output_dir"] == str(env_fallback_dir.resolve())
        assert "active_project_dir" not in status_env["profile"]

        wb = await server.create_workbook(title="RelTest", sheet_names=["Sheet1"])
        wb_id = wb["workbook_id"]
        try:
            await server.write_cell(wb_id, "Sheet1", "A1", 123)
            saved_env = await server.save_workbook(wb_id, file_path="Env_Fallback.xlsx")
            assert saved_env["status"] == "success"
            assert Path(saved_env["file_path"]) == (env_fallback_dir / "Env_Fallback.xlsx").resolve()
            assert (env_fallback_dir / "Env_Fallback.xlsx").exists()

            # 4. Priority #2 (Bionic Project Folder via check_office_status(working_dir=...)) BEATS M365_MCP_OUTPUT_DIR!
            status_proj_a = await server.check_office_status(working_dir=str(bionic_proj_a))
            assert status_proj_a["profile"]["active_project_dir"] == str(bionic_proj_a.resolve())
            saved_proj_a = await server.save_workbook(wb_id, file_path="CapEx_Stress.xlsx")
            assert saved_proj_a["status"] == "success"
            expected_a = (bionic_proj_a / "CapEx_Stress.xlsx").resolve()
            assert Path(saved_proj_a["file_path"]) == expected_a
            assert expected_a.exists()
            assert not (env_fallback_dir / "CapEx_Stress.xlsx").exists()
            assert not (project_root / "CapEx_Stress.xlsx").exists()

            # Bare relative audit/open ALSO resolves from Priority #2 (Bionic Project Folder) first
            audit_a = await server.audit_workbook_formulas(file_path="CapEx_Stress.xlsx")
            assert audit_a["status"] == "success"
            assert Path(audit_a["file_path"]) == expected_a

            # Switching Bionic projects updates Priority #2 immediately
            await server.check_office_status(working_dir=str(bionic_proj_b))
            saved_proj_b = await server.save_workbook(wb_id, file_path="CapEx_Stress.xlsx")
            assert Path(saved_proj_b["file_path"]) == (bionic_proj_b / "CapEx_Stress.xlsx").resolve()

            # 5. Priority #1 (User explicit path) BEATS BOTH Bionic Project Folder AND M365_MCP_OUTPUT_DIR!
            explicit_other = user_explicit_dir / "Other_Model.xlsx"
            saved_other = await server.save_workbook(wb_id, file_path=str(explicit_other))
            assert saved_other["status"] == "success"
            assert Path(saved_other["file_path"]) == explicit_other.resolve()
            assert explicit_other.exists()

            # 6. Priority #4: Inside MCP project folder is ALWAYS blocked
            blocked = await server.save_workbook(wb_id, file_path=str(project_root / "Sneaky.xlsx"))
            assert blocked["status"] == "error"
            assert blocked["error_type"] == "PermissionError"
            assert not (project_root / "Sneaky.xlsx").exists()
        finally:
            await server.close_workbook(wb_id)
    finally:
        set_client_project_dir(None)


@pytest.mark.asyncio
async def test_launchservices_grant_access_strictly_scoped_to_project_dir_and_external_notice(tmp_path, monkeypatch):
    """Verify LaunchServices auto-open is strictly scoped to Bionic's Project Folder ONLY, external files still prompt with 45s timeout + save notice, and timeouts emit HARD STOP."""
    from integrations.applescript_bridge import AppleScriptBridge, AppleScriptTimeoutError
    from utils.validators import set_client_project_dir

    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
    project_dir = tmp_path / "BionicProject"
    external_dir = tmp_path / "ExternalFolder"
    project_dir.mkdir()
    external_dir.mkdir()
    monkeypatch.setenv("M365_MCP_OUTPUT_DIR", str(external_dir))  # Even M365_MCP_OUTPUT_DIR must NOT waive Grant Access!
    set_client_project_dir(project_dir)

    try:
        # 1. Save inside vs outside Project Folder -> grant_access_notice ONLY on external saves
        wb = await server.create_workbook(title="ScopeTest", sheet_names=["Sheet1"])
        wb_id = wb["workbook_id"]
        in_proj_xlsx = project_dir / "Inside.xlsx"
        ext_xlsx = external_dir / "Outside.xlsx"
        try:
            res_in = await server.save_workbook(wb_id, file_path=str(in_proj_xlsx))
            assert res_in["status"] == "success"
            assert "grant_access_notice" not in res_in

            res_ext = await server.save_workbook(wb_id, file_path=str(ext_xlsx))
            assert res_ext["status"] == "success"
            assert "grant_access_notice" in res_ext
            assert "outside" in res_ext["grant_access_notice"].lower()
        finally:
            await server.close_workbook(wb_id)

        # 2. open_excel_file INSIDE Bionic Project Folder uses LaunchServices (waives Grant Access)
        bridge = AppleScriptBridge()
        ls_calls = []
        as_timeouts = []
        opened_in_excel = {"opened": False}

        async def fake_ls(app_name, resolved, activate=False):
            ls_calls.append((app_name, str(resolved), activate))
            opened_in_excel["opened"] = True

        async def fake_find(fp):
            if opened_in_excel["opened"]:
                return {"name": Path(fp).name, "active_sheet": "Sheet1", "saved": True}
            return None

        async def fake_exec(script, args=None, timeout=30.0):
            as_timeouts.append(timeout)
            if "name of active sheet of workbook" in script and "open workbook" not in script:
                return "Sheet1"
            return f"{Path(args[0]).name}\x1fSheet1"

        monkeypatch.setattr(bridge, "_open_via_launchservices", fake_ls)
        monkeypatch.setattr(bridge, "find_open_excel_workbook_by_path", fake_find)
        monkeypatch.setattr(bridge, "execute_applescript", fake_exec)

        res_open_in = await bridge.open_excel_file(str(in_proj_xlsx), read_only=False, activate=False)
        assert res_open_in["method"] == "launchservices+applescript"
        assert res_open_in["trusted_project_dir"] is True
        assert len(ls_calls) == 1

        # 3. Already-open recovery returns immediately without calling LaunchServices again
        res_already = await bridge.open_excel_file(str(in_proj_xlsx), read_only=False, activate=False)
        assert res_already["already_open"] is True
        assert res_already["method"] == "already_open+applescript"
        assert len(ls_calls) == 1

        # 4. open_excel_file OUTSIDE Bionic Project Folder NEVER calls LaunchServices and uses 45s Grant Access timeout
        opened_in_excel["opened"] = False
        as_timeouts.clear()
        res_open_ext = await bridge.open_excel_file(str(ext_xlsx), read_only=False, activate=False)
        assert res_open_ext["method"] == "applescript"
        assert len(ls_calls) == 1  # unchanged — LaunchServices was NOT called!
        assert as_timeouts[-1] == 45.0

        # 5. Timeout error payload signals awaiting_user_action=True and away-from-keyboard HARD STOP
        err_payload = server._error_response("excel_live_open", AppleScriptTimeoutError("timed out"))
        assert err_payload["awaiting_user_action"] is True
        assert "hard stop" in err_payload["hint"].lower()
        assert "continue" in err_payload["hint"].lower()
    finally:
        set_client_project_dir(None)


@pytest.mark.asyncio
async def test_afk_hardening_all_seven_guards(tmp_path, monkeypatch):
    """Verify all 7 Away-From-Keyboard (AFK) hardening guards across PDF staging, live-sync reload timeout, unsaved edits, -1743, non-live timeouts, and formula guard."""
    from integrations.applescript_bridge import AppleScriptTimeoutError
    from utils.validators import set_client_project_dir

    project_dir = tmp_path / "TrustedProject"
    external_dir = tmp_path / "ExternalDir"
    project_dir.mkdir()
    external_dir.mkdir()
    set_client_project_dir(project_dir)

    try:
        # Fix #1: Excel & Word PDF export inside trusted Project Folder stage inside save_path.parent (and auto-clean!)
        wb = await server.create_workbook(title="PdfStageTest", sheet_names=["Sheet1"])
        wb_id = wb["workbook_id"]
        opened_excel_pdf_sources = []

        async def fake_open_excel_pdf(fp, read_only=False, activate=False):
            opened_excel_pdf_sources.append((fp, read_only))
            return {"status": "success", "workbook_name": Path(fp).name, "active_sheet": "Sheet1"}

        async def fake_save_excel_live(workbook_name=None, file_path=None, file_format="pdf"):
            Path(file_path).write_bytes(b"%PDF-1.4 mock")
            return {"status": "success", "workbook_name": workbook_name, "file_path": file_path, "format": file_format}

        async def fake_close_excel_live(workbook_name=None, save_changes="no"):
            return {"status": "success", "closed_workbook": workbook_name}

        monkeypatch.setattr(server.excel.applescript, "open_excel_file", fake_open_excel_pdf)
        monkeypatch.setattr(server.excel.applescript, "save_excel_workbook", fake_save_excel_live)
        monkeypatch.setattr(server.excel.applescript, "close_excel_workbook", fake_close_excel_live)
        try:
            in_pdf = project_dir / "Brief.pdf"
            res_pdf_in = await server.save_workbook(wb_id, file_path=str(in_pdf), format="pdf")
            assert res_pdf_in["status"] == "success"
            src_in, ro_in = opened_excel_pdf_sources[-1]
            assert Path(src_in).parent == project_dir  # staged inside trusted project dir -> LaunchServices waives Grant Access!
            assert ro_in is False
            assert not Path(src_in).exists()  # hidden staging file cleaned up in finally

            ext_pdf = external_dir / "External.pdf"
            res_pdf_ext = await server.save_workbook(wb_id, file_path=str(ext_pdf), format="pdf")
            assert res_pdf_ext["status"] == "success"
            src_ext, ro_ext = opened_excel_pdf_sources[-1]
            assert Path(src_ext).parent != project_dir
            assert ro_ext is True  # external PDF keeps read_only/AppleScript so macOS still asks Grant Access
        finally:
            await server.close_workbook(wb_id)

        # Word PDF export inside trusted Project Folder also stages inside save_path.parent and cleans up
        doc = await server.create_document(title="WordPdfStage")
        doc_id = doc["document_id"]
        opened_word_pdf_sources = []

        async def fake_open_word_pdf(fp, read_only=False, activate=False):
            opened_word_pdf_sources.append((fp, read_only))
            return {"status": "success", "document_name": Path(fp).name}

        async def fake_save_word_live(document_name=None, file_path=None, file_format="pdf"):
            Path(file_path).write_bytes(b"%PDF-1.4 word mock")
            return {"status": "success", "document_name": document_name, "file_path": file_path, "format": file_format}

        async def fake_close_word_live(document_name=None, save_changes="no"):
            return {"status": "success", "closed_document": document_name}

        monkeypatch.setattr(server.word.applescript, "open_word_file", fake_open_word_pdf)
        monkeypatch.setattr(server.word.applescript, "save_word_document", fake_save_word_live)
        monkeypatch.setattr(server.word.applescript, "close_word_document", fake_close_word_live)
        try:
            w_pdf = project_dir / "Memo.pdf"
            res_w_pdf = await server.save_document(doc_id, file_path=str(w_pdf), format="pdf")
            assert res_w_pdf["status"] == "success"
            w_src, _ = opened_word_pdf_sources[-1]
            assert Path(w_src).parent == project_dir
            assert not Path(w_src).exists()
        finally:
            await server.close_document(doc_id)

        # Fix #2: M365_MCP_LIVE_SYNC=AUTO reload timeout surfaces awaiting_user_action=True + HARD STOP
        monkeypatch.setenv("M365_MCP_LIVE_SYNC", "auto")
        wb2 = await server.create_workbook(title="ReloadTimeoutTest", sheet_names=["Sheet1"])
        wb2_id = wb2["workbook_id"]

        async def fake_find_saved(fp):
            return {"name": Path(fp).name, "saved": True}

        async def fake_capture_view(wb_name):
            return {}

        async def fake_open_timeout(fp, read_only=False, activate=False):
            raise AppleScriptTimeoutError("Excel Grant Access timed out")

        monkeypatch.setattr(server.excel.applescript, "find_open_excel_workbook_by_path", fake_find_saved)
        monkeypatch.setattr(server.excel.applescript, "capture_excel_workbook_view_state", fake_capture_view)
        monkeypatch.setattr(server.excel.applescript, "open_excel_file", fake_open_timeout)
        try:
            res_reload_to = await server.save_workbook(wb2_id, file_path=str(external_dir / "Reloaded.xlsx"))
            assert res_reload_to["status"] == "success"
            assert res_reload_to["awaiting_user_action"] is True
            assert "hard stop" in res_reload_to["hint"].lower()
            assert "grant_access_notice" in res_reload_to
        finally:
            await server.close_workbook(wb2_id)

        # Fix #3: Unsaved manual changes collision emits awaiting_user_action=True + HARD STOP
        err_unsaved = server._error_response(
            "save_workbook",
            RuntimeError("File 'Model.xlsx' is currently open in Microsoft Excel with unsaved manual changes."),
        )
        assert err_unsaved["awaiting_user_action"] is True
        assert "hard stop" in err_unsaved["hint"].lower()
        assert "excel_live_close" in err_unsaved["hint"].lower()

        # Fix #4: Automation -1743 PermissionError emits awaiting_user_action=True + HARD STOP (no file-level fallback)
        err_1743 = server._error_response("excel_live_open", RuntimeError("Not authorized to send Apple events (-1743)"))
        assert err_1743["awaiting_user_action"] is True
        assert "hard stop" in err_1743["hint"].lower()
        assert "instead of *_live_*" not in err_1743["hint"].lower()

        # Fix #5: Non-live TimeoutError (e.g. save_workbook/save_document/open_presentation) ALSO emits awaiting_user_action=True + HARD STOP
        err_nonlive_to = server._error_response("save_document", AppleScriptTimeoutError("timed out"))
        assert err_nonlive_to["awaiting_user_action"] is True
        assert "hard stop" in err_nonlive_to["hint"].lower()
        assert "then retry" not in err_nonlive_to["hint"].lower()

        # Fix #7: Formula Guard response also carries unified awaiting_user_action=True
        fg = server.excel._formula_guard_response("literals_only", [{"sheet_name": "S", "cell": "B2"}])
        assert fg["status"] == "confirmation_required"
        assert fg["awaiting_user_action"] is True

        # Bonus universal check: dict-shaped fill_formula (standalone & combined with value) expands cleanly
        wb3 = await server.create_workbook(title="DictFillTest", sheet_names=["CF"])
        wb3_id = wb3["workbook_id"]
        try:
            res_df = await server.batch_write_ranges(
                workbook_id=wb3_id,
                operations=[
                    {
                        "sheet": "CF",
                        "range": "B2:C2",
                        "value": [[0, 10]],
                        "fill_formula": {"range": "D2:F2", "formula": "=C2*1.12"},
                    },
                    {
                        "sheet": "CF",
                        "range": "C3:E3",
                        "value": None,
                        "fill_formula": {"range": "C3:E3", "formula": "=C2*0.38"},
                    },
                ],
            )
            assert res_df["status"] == "success"
            r2 = await server.read_range(wb3_id, "CF", "B2:F2", evaluate_formulas=False)
            assert r2["data"][0] == [0, 10, "=C2*1.12", "=D2*1.12", "=E2*1.12"]
            r3 = await server.read_range(wb3_id, "CF", "C3:E3", evaluate_formulas=False)
            assert r3["data"][0] == ["=C2*0.38", "=D2*0.38", "=E2*0.38"]
        finally:
            await server.close_workbook(wb3_id)
    finally:
        set_client_project_dir(None)


@pytest.mark.asyncio
async def test_wall_street_auto_styler_unicode_unescape_cross_sheet_chart_and_save_linters(tmp_path, monkeypatch):
    """Verify definitive formatting/chart/width/Unicode fixes across all financial models."""
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
    wb = await server.create_workbook(
        title="CapEx AutoStyle Test",
        sheet_names=["Executive Summary", "Assumptions", "Project Cash Flows", "Sensitivity & Charts"],
    )
    wb_id = wb["workbook_id"]
    out_xlsx = tmp_path / "CapEx_AutoStyle.xlsx"
    try:
        # 1. Write model including double-escaped \uXXXX display labels (must decode) & formulas (must NOT touch)
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {
                    "sheet_name": "Assumptions",
                    "range": "A1:B6",
                    "data": [
                        ["Parameter", "Value"],
                        ["Project Life (Years)", 6],
                        ["Year-0 CapEx ($)", 25000000],
                        ["WACC", 0.105],
                        ["Tax Rate", 0.25],
                        ["EBITDA Margin", 0.38],
                    ],
                },
                {
                    "sheet_name": "Project Cash Flows",
                    "range": "A1:E8",
                    "data": [
                        ["Line Item ($)", 0, 1, 2, 3],
                        ["Revenue", 0, 10000000, "=C2*1.12", "=D2*1.12"],
                        ["EBITDA", 0, "=C2*Assumptions!$B$6", "=D2*Assumptions!$B$6", "=E2*Assumptions!$B$6"],
                        ["-\\u0394NWC", 0, -800000, -96000, -107520],
                        ["+After-Tax Salvage", 0, 0, 0, 3000000],
                        ["UFCF", "=-Assumptions!$B$3", "=C3+C4", "=D3+D4", "=E3+E4+E5"],
                        ["Disc. Factor", 1, "=1/(1+Assumptions!$B$4)^C$1", "=1/(1+Assumptions!$B$4)^D$1", "=1/(1+Assumptions!$B$4)^E$1"],
                        ["Cum. PV", "=B6", "=B8+C6*C7", "=C8+D6*D7", "=D8+E6*E7"],
                    ],
                },
                {
                    "sheet_name": "Executive Summary",
                    "range": "A1:B6",
                    "data": [
                        ["KPI", "Value"],
                        ["Net Present Value (NPV)", "=NPV(Assumptions!$B$4,'Project Cash Flows'!C6:E6)+'Project Cash Flows'!B6"],
                        ["Internal Rate of Return (IRR)", "=IRR('Project Cash Flows'!B6:E6)"],
                        ["Profitability Index", "=(B2+Assumptions!$B$3)/Assumptions!$B$3"],
                        ["Simple Payback (Years)", 5.29],
                        ["Go / No-Go", '=IF(B2>0,"Go","No-Go")'],
                    ],
                },
                {
                    "sheet_name": "Sensitivity & Charts",
                    "range": "A1:E5",
                    "data": [
                        ["NPV Sensitivity Analysis", None, None, None, None],
                        [None, None, None, None, None],
                        ["WACC \\u2193 / CF Scale \\u2192", None, 0.085, 0.105, 0.125],
                        [None, 0.8, "=NPV(C$3,'Project Cash Flows'!$C$6:$E$6)*$B4+'Project Cash Flows'!$B$6", "=NPV(D$3,'Project Cash Flows'!$C$6:$E$6)*$B4+'Project Cash Flows'!$B$6", "=NPV(E$3,'Project Cash Flows'!$C$6:$E$6)*$B4+'Project Cash Flows'!$B$6"],
                        [None, 1.0, "=NPV(C$3,'Project Cash Flows'!$C$6:$E$6)*$B5+'Project Cash Flows'!$B$6", "=NPV(D$3,'Project Cash Flows'!$C$6:$E$6)*$B5+'Project Cash Flows'!$B$6", "=NPV(E$3,'Project Cash Flows'!$C$6:$E$6)*$B5+'Project Cash Flows'!$B$6"],
                    ],
                },
            ],
        )

        wb_obj = server.excel.active_workbooks[wb_id]["workbook_object"]
        # Unicode unescape check on non-formula display labels
        assert wb_obj["Project Cash Flows"]["A4"].value == "-ΔNWC"
        assert wb_obj["Sensitivity & Charts"]["A3"].value == "WACC ↓ / CF Scale →"

        # 2. Save BEFORE formatting & chart -> linter catches unformatted formulas, missing chart, AND missing column widths!
        pre_save = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert pre_save["status"] == "success"
        pre_warns = " | ".join(pre_save.get("sanity_warnings") or [])
        assert "general" in pre_warns.lower()
        assert "0 charts" in pre_warns.lower()
        assert "column widths" in pre_warns.lower()

        # 3. Cross-sheet create_chart places chart on 'Sensitivity & Charts' sourced from 'Project Cash Flows'
        chart_res = await server.create_chart(
            workbook_id=wb_id,
            sheet_name="Sensitivity & Charts",
            chart_type="column",
            data_range="'Project Cash Flows'!A1:E2",
            chart_title="Project Revenue Ramp",
            position="A10",
        )
        assert chart_res["status"] == "success"
        assert chart_res["source_sheet"] == "Project Cash Flows"
        assert len(wb_obj["Sensitivity & Charts"]._charts) == 1

        # 4. batch_format_workbook(auto_color_financial_model=True) WITHOUT explicit column_widths or number_formats
        #    automatically applies institutional Wall Street formats, headers, totals, AND auto-fitted widths!
        fmt_res = await server.batch_format_workbook(
            workbook_id=wb_id,
            operations=[
                {"sheet_name": "Executive Summary", "auto_color_financial_model": True},
                {"sheet_name": "Assumptions", "auto_color_financial_model": True},
                {"sheet_name": "Project Cash Flows", "auto_color_financial_model": True},
                {"sheet_name": "Sensitivity & Charts", "auto_color_financial_model": True},
            ],
        )
        assert fmt_res["status"] == "success"
        for sheet_res in fmt_res["sheets"]:
            assert sheet_res["column_widths_set"] > 0

        # Verify smart Wall Street number formats, headers, totals, and unclipped column widths
        ws_cf = wb_obj["Project Cash Flows"]
        assert ws_cf["A1"].fill.start_color.rgb.endswith("1F4E79")
        assert ws_cf["C2"].number_format == '$#,##0;($#,##0);"-"'
        assert ws_cf["C6"].number_format == '$#,##0;($#,##0);"-"'
        assert ws_cf["C6"].font.bold is True and ws_cf["C6"].fill.start_color.rgb.endswith("D9E1F2")
        assert ws_cf["C7"].number_format == "0.000"
        assert ws_cf.column_dimensions["A"].width >= 20.0

        ws_es = wb_obj["Executive Summary"]
        assert ws_es["B2"].number_format == '$#,##0;($#,##0);"-"'
        assert ws_es["B3"].number_format == "0.0%"
        assert ws_es["B4"].number_format == '0.00"x"'
        assert ws_es["B5"].number_format == "0.00"

        ws_as = wb_obj["Assumptions"]
        assert ws_as["B2"].number_format == "#,##0"
        assert ws_as["B3"].number_format == '$#,##0;($#,##0);"-"'
        assert ws_as["B4"].number_format == "0.0%"
        # Assumptions % row (EBITDA Margin) must NOT get shaded as a total row
        assert ws_as["A6"].fill.fill_type is None

        ws_sc = wb_obj["Sensitivity & Charts"]
        assert ws_sc["C3"].number_format == "0.0%"
        assert ws_sc["B4"].number_format == "0.0%"
        assert ws_sc["C4"].number_format == '$#,##0;($#,##0);"-"'
        assert ws_sc.column_dimensions["A"].width >= 20.0

        # 5. Re-save AFTER formatting + chart -> ZERO sanity_warnings!
        post_save = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert post_save["status"] == "success"
        assert "sanity_warnings" not in post_save

        # 6. Universal formula quote sanitizer repairs Bionic triple-escaped \" and standalone '...' literals
        #    while preserving 'Sheet Name'!, escaped sheet apostrophes ('Q1''26 P&L'!), and "Bob's Store"
        from integrations.applescript_bridge import AppleScriptBridge
        from utils.validators import sanitize_excel_formula_quotes

        await server.write_cell(
            wb_id,
            "Executive Summary",
            "B6",
            "=IF(B2>0,'Go','No-Go')",
        )
        assert wb_obj["Executive Summary"]["B6"].value == '=IF(B2>0,"Go","No-Go")'

        await server.write_cell(
            wb_id,
            "Executive Summary",
            "B5",
            '=IF(MAX(\'Project Cash Flows\'!$B$8:$E$8)<0,\\">6 yrs\\",COUNTIF(\'Project Cash Flows\'!$B$8:$E$8,\\"<0\\"))',
        )
        assert (
            wb_obj["Executive Summary"]["B5"].value
            == '=IF(MAX(\'Project Cash Flows\'!$B$8:$E$8)<0,">6 yrs",COUNTIF(\'Project Cash Flows\'!$B$8:$E$8,"<0"))'
        )
        assert (
            sanitize_excel_formula_quotes("=IF(A1=\"Bob's Store\",'Q1''26 P&L'!B2,'No')")
            == '=IF(A1="Bob\'s Store",\'Q1\'\'26 P&L\'!B2,"No")'
        )
        assert (
            AppleScriptBridge._sanitize_live_excel_formula("=IF(B2>0,'Go','No-Go')", ["Executive Summary"])
            == '=IF(B2>0,"Go","No-Go")'
        )

        # 7. Exact reproduction of real-world Bionic/Qwen formatting, chart & outflow-sign failure modes:
        #    (a) {"operation": ...} alias + sheet_name=None broadcasts across ALL 4 sheets
        op_alias_res = await server.batch_format_workbook(
            workbook_id=wb_id,
            operations=[
                {"operation": "auto_color_financial_model", "sheet_name": None},
                {"operation": "auto_fit_columns", "sheet_name": None},
                {"operation": "freeze_panes", "sheet_name": "B2"},
            ],
        )
        assert op_alias_res["status"] == "success"
        assert op_alias_res["sheets_formatted"] == 12  # 3 workbook-wide ops x 4 tabs
        assert wb_obj["Project Cash Flows"].freeze_panes == "B2"

        #    (b) Multi-sheet range op WITHOUT sheet_name MUST fail loudly (never silently clobber Executive Summary!)
        bad_fmt = await server.batch_format_workbook(
            workbook_id=wb_id,
            operations=[{"range": "B4:B8", "formatting": {"number_format": "0.0%"}}],
        )
        assert bad_fmt["status"] == "error"
        assert "sheet_name" in bad_fmt["error"]

        #    (c) create_chart on formula-only matrix C4:E5 auto-picks C3:E3 headers + B4:B5 categories & solid neg bars
        sens_chart_res = await server.create_chart(
            workbook_id=wb_id,
            sheet_name="Sensitivity & Charts",
            chart_type="column",
            data_range="C4:E5",
            chart_title="NPV Sensitivity",
            position="A25",
        )
        assert sens_chart_res["status"] == "success"
        sens_chart = wb_obj["Sensitivity & Charts"]._charts[-1]
        assert len(sens_chart.series) == 3  # 3 WACC columns C..E, NOT eating Row 4 as legend titles!
        assert all(ser.invertIfNegative is False for ser in sens_chart.series)
        assert sens_chart.legend.position == "b"
        assert sens_chart.x_axis.tickLblPos == "low"

        #    (d) Outflow sign linter catches positive '=Assumptions!$B$3' on '- CapEx' row
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "Project Cash Flows", "range": "A9:B9", "data": [["- CapEx", "=Assumptions!$B$3"]]}],
        )
        sign_save = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert any("outflow sign" in w.lower() for w in (sign_save.get("sanity_warnings") or []))
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_imp002_live_display_values_and_clipping_detection(monkeypatch):
    """Verify IMP-002: live formatted display_values on compact tables, ### clipping warning, and large-range token guard."""
    from integrations.applescript_bridge import AppleScriptBridge, _FS, _GS, _RS, _US

    bridge = AppleScriptBridge()
    # Mock AppleScript output with raw floats + formatted display strings (including a '###' clipped cell!)
    val_blk = _RS.join([
        _US.join(["NPV ($)", "-3224566.54"]),
        _US.join(["IRR", "0.066322"]),
        _US.join(["Wide Val", "99999999999"]),
    ])
    str_blk = _RS.join([
        _US.join(["NPV ($)", "($3,225)"]),
        _US.join(["IRR", "6.6%"]),
        _US.join(["Wide Val", "########"]),
    ])
    fake_batch_raw = f"CapEx_Stress.xlsx{_FS}" + _GS.join([f"Executive Summary{_US}A6:B8", val_blk, "", str_blk])

    async def fake_exec(script, args=None, timeout=45.0):
        return fake_batch_raw

    monkeypatch.setattr(bridge, "execute_applescript", fake_exec)
    res = await bridge.batch_read_live_excel_ranges(
        ranges=[{"sheet_name": "Executive Summary", "range_ref": "A6:B8"}],
        workbook_name="CapEx_Stress.xlsx",
    )
    assert res["status"] == "success"
    r0 = res["results"][0]
    assert r0["display_values"][0][1] == "($3,225)"
    assert r0["display_values"][1][1] == "6.6%"
    assert "B8" in r0["clipped_cells"]
    assert "Executive Summary!B8" in res["clipped_cells"]
    assert "###" in res["warning"]

    # Large unclipped range (>50 cells) omits display_values in 'auto' mode to save tokens
    big_vals = [[str(i), str(i * 1000)] for i in range(30)]  # 60 cells
    big_str = _RS.join([_US.join([str(i), f"${i},000"]) for i in range(30)])
    disp_big, clipped_big = bridge._process_live_display_strings(big_vals, big_str, "A1:B30", "auto")
    assert disp_big is None and clipped_big == []


@pytest.mark.asyncio
async def test_imp003_layout_agnostic_multi_table_formatting_and_audit(tmp_path, monkeypatch):
    """Verify IMP-003: arbitrary (row, col) headers, stacked & side-by-side tables, ERP code col, multilingual totals, and block-scoped audit."""
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
    wb = await server.create_workbook(title="MultiTable Layout", sheet_names=["MultiTab", "ERP_ES"])
    wb_id = wb["workbook_id"]
    out_xlsx = tmp_path / "MultiTable_IMP003.xlsx"
    try:
        # Sheet 1: Indented (blank Col A), 4-line title in B1:B4, Table 1 (B6:F10) side-by-side with Table 2 (H6:K9),
        # AND a stacked Table 3 (B13:F17) below! Plant a silent wrong-period bug in Table 1 E9 (=E8/D7 instead of =E8/E7).
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {
                    "sheet_name": "MultiTab",
                    "range": "B1:B4",
                    "data": [
                        ["Apex Global Industrial Holdings Corporation — Very Long Title That Must Not Bloat Col B"],
                        ["Consolidated Segment Schedule"],
                        ["Fiscal Years: 2024 – 2027"],
                        ["In USD $"],
                    ],
                },
                {
                    "sheet_name": "MultiTab",
                    "range": "B6:F10",
                    "data": [
                        ["Line Item", 2024, 2025, 2026, 2027],
                        ["Revenue", 100000, 110000, 121000, 133100],
                        ["Gross Profit", 40000, 44000, 48400, 53240],
                        ["Gross Margin %", "=C8/C7", "=D8/D7", "=E8/D7", "=F8/F7"],  # <-- E9 wrong-period bug!
                        ["Net Income", "=C8-10000", "=D8-10000", "=E8-10000", "=F8-10000"],
                    ],
                },
                {
                    # Side-by-side Table 2 on SAME rows (6..9) in Columns H..K with totally different formulas!
                    "sheet_name": "MultiTab",
                    "range": "H6:K9",
                    "data": [
                        ["KPI", 2024, 2025, 2026],
                        ["Debt", 50000, 45000, 40000],
                        ["Cash", 10000, 15000, 20000],
                        ["Net Debt", "=I7-I8", "=J7-J8", "=K7-K8"],
                    ],
                },
                {
                    # Stacked Table 3 (2-way sensitivity matrix on a non-'Sensitivity' tab!) in B13:E16
                    "sheet_name": "MultiTab",
                    "range": "B13:E16",
                    "data": [
                        ["WACC / Scale", 0.08, 0.10, 0.12],
                        [0.9, "=C10*0.9", "=D10*0.9", "=E10*0.9"],
                        [1.0, "=C10*1.0", "=D10*1.0", "=E10*1.0"],
                        [1.1, "=C10*1.1", "=D10*1.1", "=E10*1.1"],
                    ],
                },
                {
                    # Sheet 2: ERP trial balance starting at D5 with Account Code in Col D, Spanish names in Col E,
                    # mid-table section subheader, and vertical =SUM(...) total with no English keywords!
                    "sheet_name": "ERP_ES",
                    "range": "D5:G10",
                    "data": [
                        ["Codigo", "Cuenta", 2024, 2025],
                        [None, "ACTIVO CORRIENTE", None, None],
                        [1010, "Efectivo", 50000, 60000],
                        [1020, "Cuentas por Cobrar", 30000, 35000],
                        [1030, "Inventarios", 20000, 25000],
                        [1099, "Suma Activo Circulante", "=SUM(F7:F9)", "=SUM(G7:G9)"],
                    ],
                },
            ],
        )

        fmt_res = await server.batch_format_workbook(
            workbook_id=wb_id,
            operations=[{"auto_color_financial_model": True, "auto_fit_columns": True, "freeze_panes": "auto"}],
        )
        assert fmt_res["status"] == "success"
        wb_obj = server.excel.active_workbooks[wb_id]["workbook_object"]
        ws1 = wb_obj["MultiTab"]
        ws2 = wb_obj["ERP_ES"]

        # Row-6 year headers styled white-on-navy with integer '0' (NOT blue '$2,024'!)
        assert ws1["C6"].fill.start_color.rgb.endswith("1F4E79")
        assert ws1["C6"].number_format == "0"
        assert ws1["I6"].fill.start_color.rgb.endswith("1F4E79")
        assert ws1["I6"].number_format == "0"
        # Overflow 85-char title in B1 must NOT bloat Col B width to 60!
        assert ws1.column_dimensions["B"].width < 35.0
        # Dynamic freeze_panes='auto' froze right below Row-6 header and right of Col-B label -> C7!
        assert ws1.freeze_panes == "C7"
        # Matrix table detected structurally on 'MultiTab' -> axes formatted as %!
        assert ws1["C13"].number_format == "0.0%"
        assert ws1["B14"].number_format == "0.0%"
        # ERP_ES: account code col D, label col E -> freeze_panes='auto' -> F6, Spanish SUM row auto-styled as Total!
        assert ws2.freeze_panes == "F6"
        assert ws2["F10"].font.bold is True and ws2["F10"].fill.start_color.rgb.endswith("D9E1F2")
        # Mid-table section subheader 'ACTIVO CORRIENTE' bold navy
        assert ws2["E6"].font.bold is True

        # inspect_workbook surfaces detected_tables for both tabs
        insp = await server.inspect_workbook(workbook_id=wb_id)
        assert len(insp["sheets"][0]["detected_tables"]) == 3
        assert insp["sheets"][1]["detected_tables"][0]["code_col"] == "D"
        assert insp["sheets"][1]["detected_tables"][0]["label_col"] == "E"

        # audit_workbook_formulas: catches E9 with exact col_header=2026 & row_label='Gross Margin %',
        # zero side-by-side collision with Table 2 (H6:K9), and zero false-positive suspicious_literals on year headers/axes!
        audit = await server.audit_workbook_formulas(workbook_id=wb_id)
        assert audit["row_inconsistency_count"] == 1
        inc = audit["row_inconsistencies"][0]
        assert inc["cell"] == "E9"
        assert inc["col_header"] == 2026
        assert inc["row_label"] == "Gross Margin %"
        assert inc["expected_formula"] == "=E8/E7"
        assert audit["suspicious_literal_count"] == 0
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_imp001_smart_magnitude_4line_headers_mixed_assumptions_and_jump_linter(tmp_path, monkeypatch):
    """Verify IMP-001: 4-line header_metadata, Option-A Smart Magnitude ($1,000's & $ Millions whole numbers), mixed Assumptions, per-share protection, and 1,000x jump linter."""
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")

    # 1. Millions-scale model ($25M CapEx, $350K maint, $125/unit) -> Smart Magnitude auto-resolves to Thousands ($#,##0,)
    wb = await server.create_workbook(
        title="Mexico Line CapEx",
        sheet_names=["Executive Summary", "Assumptions", "Project Cash Flows"],
        header_metadata={
            "entity": "Apex Industrials — Mexico Packaging Line",
            "period": "Projected Years: Year 0 to Year 6",
            "currency": "USD",
            "magnitude": "auto",
        },
    )
    wb_id = wb["workbook_id"]
    assert wb["table_start_row"] == 6
    assert wb["recommended_freeze_panes"] == "B7"
    out_xlsx = tmp_path / "SmartMag_CapEx.xlsx"
    try:
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[
                {
                    "sheet_name": "Assumptions",
                    "range": "A6:C11",
                    "data": [
                        ["Parameter", "Value", "Unit"],
                        ["Year-0 CapEx", 25000000, "USD exact ($25M)"],
                        ["Annual Maintenance", 350000, "USD exact ($350K)"],
                        ["Unit Selling Price", 125.50, "USD / unit"],
                        ["WACC", 0.105, "%"],
                        ["Diluted Shares", 10000000, "Shares"],
                    ],
                },
                {
                    "sheet_name": "Project Cash Flows",
                    "range": "A6:D9",
                    "data": [
                        ["Line Item", 0, 1, 2],
                        ["Revenue", 0, 10000000, "=C7*1.12"],
                        ["- CapEx", "=-Assumptions!$B$7", 0, 0],
                        ["UFCF", "=B8", "=C7-Assumptions!$B$8", "=D7-Assumptions!$B$8"],
                    ],
                },
                {
                    "sheet_name": "Executive Summary",
                    "range": "A6:B9",
                    "data": [
                        ["KPI", "Value"],
                        ["Net Present Value (NPV)", "=NPV(Assumptions!$B$10,'Project Cash Flows'!C9:D9)+'Project Cash Flows'!B9"],
                        ["Implied Share Price", "=B7/Assumptions!$B$11"],
                        ["IRR", "=IRR('Project Cash Flows'!B9:D9)"],
                    ],
                },
            ],
        )

        fmt_res = await server.batch_format_workbook(
            workbook_id=wb_id,
            operations=[{"auto_color_financial_model": True, "auto_fit_columns": True, "freeze_panes": "auto"}],
        )
        assert fmt_res["status"] == "success"
        wb_obj = server.excel.active_workbooks[wb_id]["workbook_object"]

        # Row 4 auto-synced to Smart Magnitude ('In USD $1,000's') on statements, and transparent exact-basis on Assumptions!
        ws_cf = wb_obj["Project Cash Flows"]
        ws_es = wb_obj["Executive Summary"]
        ws_as = wb_obj["Assumptions"]
        assert "1,000's" in ws_cf["A4"].value
        assert "except per-share" in ws_es["A4"].value.lower()
        assert "Exact USD $" in ws_as["A4"].value and "1,000's" in ws_as["A4"].value

        # Option-A whole-number trailing-comma scaling on statements; exact currency on Assumptions & per-unit/per-share rows!
        assert ws_cf["C7"].number_format == '$#,##0,;($#,##0,);"-"'
        assert ws_es["B7"].number_format == '$#,##0,;($#,##0,);"-"'
        assert ws_es["B8"].number_format == '$#,##0.00;($#,##0.00);"-"'  # Implied Share Price protected!
        assert ws_es["B9"].number_format == "0.0%"
        assert ws_as["B7"].number_format == '$#,##0;($#,##0);"-"'        # Assumptions CapEx exact $25,000,000
        assert ws_as["B9"].number_format == '$#,##0.00;($#,##0.00);"-"'  # Unit Selling Price exact $125.50
        assert ws_cf.freeze_panes == "B7"

        # Save clean (0 warnings!)
        saved = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert saved["status"] == "success"
        assert "sanity_warnings" not in saved

        # 2. 1,000x magnitude-jump linter catches accidental $M literal (-25) mixed into a $25,000,000 workbook!
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "Executive Summary", "range": "A10:B10", "data": [["Manual Outlay", 25]]}],
        )
        # Write a separate small-scale sheet to trigger cross-sheet jump check cleanly
        await server.add_worksheet(wb_id, "BadScale")
        await server.batch_write_ranges(
            workbook_id=wb_id,
            operations=[{"sheet_name": "BadScale", "range": "A1:B2", "data": [["Item", "Val"], ["Revenue", 10]]}],
        )
        jump_save = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert any("magnitude mismatch" in w.lower() for w in (jump_save.get("sanity_warnings") or []))
    finally:
        await server.close_workbook(wb_id)

    # 3. Billions-scale EUR model (€2.5B) -> auto-resolves to EUR € Millions (€#,##0,,)
    wb_b = await server.create_workbook(
        title="MegaCap EUR DCF",
        sheet_names=["DCF Valuation"],
        header_metadata={"entity": "EuroCorp SA", "period": "FY2025 – FY2030", "currency": "EUR", "magnitude": "auto"},
    )
    wb_b_id = wb_b["workbook_id"]
    try:
        await server.batch_write_ranges(
            workbook_id=wb_b_id,
            operations=[{
                "sheet_name": "DCF Valuation",
                "range": "A6:C8",
                "data": [
                    ["Line Item", 2025, 2026],
                    ["Revenue", 2500000000, "=B7*1.08"],
                    ["EBITDA", "=B7*0.30", "=C7*0.30"],
                ],
            }],
        )
        await server.batch_format_workbook(
            workbook_id=wb_b_id,
            operations=[{"auto_color_financial_model": True, "auto_fit_columns": True, "freeze_panes": "auto"}],
        )
        ws_b = server.excel.active_workbooks[wb_b_id]["workbook_object"]["DCF Valuation"]
        assert ws_b["A4"].value == "In EUR € Millions"
        assert ws_b["B7"].number_format == '€#,##0,,;(€#,##0,,);"-"'
    finally:
        await server.close_workbook(wb_b_id)


@pytest.mark.asyncio
async def test_audit_hardening_sparse_ghost_sheet_afk_live_guard_error_hints_and_atomic_save(tmp_path, monkeypatch):
    """Verify the 5 architectural hardening fixes: sparse 1M-row ghost sheet speed, live Formula Guard AFK, PermissionError vs -1743 split, profile-aware hints, and atomic saves."""
    import time
    from integrations.applescript_bridge import AppleScriptBridge
    from utils.sheet_layout import analyze_worksheet_layout

    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")

    # 1. Live Formula Guard emits awaiting_user_action=True + HARD STOP (AFK parity with ExcelController)
    fg_live = AppleScriptBridge._formula_guard_response("literals_only", [{"sheet": "S1", "cell": "B2"}])
    assert fg_live["status"] == "confirmation_required"
    assert fg_live["awaiting_user_action"] is True
    assert "hard stop" in fg_live["instruction"].lower()

    # 2. _error_response cleanly separates OS filesystem PermissionError (Files & Folders) from AppleScript -1743
    fs_perm = server._error_response("save_workbook", PermissionError("[Errno 1] Operation not permitted: '~/Desktop/X.xlsx'"))
    assert fs_perm["awaiting_user_action"] is True
    assert "files & folders" in fs_perm["hint"].lower()
    assert "-1743" not in fs_perm["hint"]

    auto_1743 = server._error_response("excel_live_open", RuntimeError("Not authorized to send Apple events (-1743)"))
    assert auto_1743["awaiting_user_action"] is True
    assert "-1743" in auto_1743["hint"]

    # 3. Profile-aware hints NEVER suggest unregistered ghost tools in compact profiles
    orig_reg = list(server._REGISTERED_TOOL_NAMES)
    try:
        server._REGISTERED_TOOL_NAMES[:] = ["create_workbook", "open_workbook", "batch_write_ranges"]
        nf_err = server._error_response("batch_write_ranges", ValueError("Workbook 'bad-id' not found."))
        assert "list_active_workbooks" not in nf_err.get("hint", "")
        assert "create_workbook" in nf_err.get("hint", "")
    finally:
        server._REGISTERED_TOOL_NAMES[:] = orig_reg

    # 4. Sparse ws._cells iteration handles a ghost 500,000-row worksheet in < 1s without allocating empty cells!
    wb = await server.create_workbook(title="GhostRowTest", sheet_names=["Ghost"])
    wb_id = wb["workbook_id"]
    out_xlsx = tmp_path / "Atomic_Ghost.xlsx"
    try:
        ws = server.excel.active_workbooks[wb_id]["workbook_object"]["Ghost"]
        ws["A1"] = "Item"
        ws["B1"] = "Val"
        ws["A2"] = "Rev"
        ws["B2"] = 10000
        # Plant a ghost formatted blank cell way down at row 500,000 (common in corporate Excel files!)
        ws.cell(row=500000, column=1).value = None
        assert ws.max_row == 500000

        t0 = time.perf_counter()
        lay = analyze_worksheet_layout(ws)
        audit = await server.audit_workbook_formulas(workbook_id=wb_id)
        elapsed = time.perf_counter() - t0
        assert elapsed < 2.0, f"Sparse scan took too long on 500k ghost sheet: {elapsed:.2f}s"
        assert len(lay.data_blocks()) == 1
        assert audit["clean"] is True
        # Ensure we did NOT allocate 500,000 empty cells in ws._cells!
        assert len(ws._cells) < 50

        # 5. Atomic save leaves no orphan .tmp files and attaches attention_required when warnings exist
        saved = await server.save_workbook(wb_id, file_path=str(out_xlsx))
        assert saved["status"] == "success"
        assert out_xlsx.exists()
        assert not list(tmp_path.glob("*.tmp.*"))
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_imp005_session_scoped_working_dir_and_write_fences(tmp_path, monkeypatch):
    """Verify IMP-005: async session-isolated working_dir, M365_MCP_BLOCKED_DIRS, and M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR."""
    import asyncio
    from utils.validators import (
        clear_session_project_dir,
        get_client_project_dir,
        reset_current_session_id,
        resolve_file_path,
        set_client_project_dir,
        set_current_session_id,
    )

    proj_a = tmp_path / "ProjectA"
    proj_b = tmp_path / "ProjectB"
    blocked_dir = tmp_path / "BlockedVault"
    external_dir = tmp_path / "ExternalOther"
    for d in (proj_a, proj_b, blocked_dir, external_dir):
        d.mkdir()

    # 1. Concurrent async sessions have strictly isolated working_dir contexts
    async def run_session(sess_id: str, folder: Path, filename: str) -> Path:
        tok = set_current_session_id(sess_id)
        try:
            set_client_project_dir(folder)
            await asyncio.sleep(0.02)
            assert get_client_project_dir() == folder.resolve()
            return resolve_file_path(filename, for_write=True)
        finally:
            clear_session_project_dir(sess_id)
            reset_current_session_id(tok)

    path_a, path_b = await asyncio.gather(
        run_session("sess_alpha", proj_a, "Alpha.xlsx"),
        run_session("sess_beta", proj_b, "Beta.xlsx"),
    )
    assert path_a == (proj_a / "Alpha.xlsx").resolve()
    assert path_b == (proj_b / "Beta.xlsx").resolve()

    # Unregistered session never inherits another session's folder
    tok_c = set_current_session_id("sess_clean")
    try:
        assert get_client_project_dir() is None
    finally:
        reset_current_session_id(tok_c)

    # 2. M365_MCP_BLOCKED_DIRS blocks writes cleanly with awaiting_user_action + HARD STOP
    monkeypatch.setenv("M365_MCP_BLOCKED_DIRS", str(blocked_dir))
    monkeypatch.setenv("M365_MCP_LIVE_SYNC", "off")
    wb = await server.create_workbook(title="FenceTest", sheet_names=["S1"])
    wb_id = wb["workbook_id"]
    try:
        res_blocked = await server.save_workbook(wb_id, file_path=str(blocked_dir / "Secret.xlsx"))
        assert res_blocked["status"] == "error"
        assert res_blocked["error_type"] == "PermissionError"
        assert res_blocked["awaiting_user_action"] is True
        assert "hard stop" in res_blocked["hint"].lower()
        assert "m365_mcp_blocked_dirs" in res_blocked["error"].lower()
        assert "files & folders" not in res_blocked["hint"].lower()

        # 3. M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR restricts writes strictly to active Project Working Folder
        monkeypatch.delenv("M365_MCP_BLOCKED_DIRS", raising=False)
        monkeypatch.setenv("M365_MCP_RESTRICT_WRITE_TO_PROJECT_DIR", "true")
        set_client_project_dir(proj_a)
        try:
            res_ext = await server.save_workbook(wb_id, file_path=str(external_dir / "Escape.xlsx"))
            assert res_ext["status"] == "error"
            assert res_ext["error_type"] == "PermissionError"
            assert res_ext["awaiting_user_action"] is True
            assert "m365_mcp_restrict_write_to_project_dir" in res_ext["error"].lower()

            res_ok = await server.save_workbook(wb_id, file_path=str(proj_a / "Allowed.xlsx"))
            assert res_ok["status"] == "success"
            assert (proj_a / "Allowed.xlsx").exists()

            # 4. check_office_status surfaces fence config cleanly
            monkeypatch.setenv("M365_MCP_BLOCKED_DIRS", str(blocked_dir))
            status = await server.check_office_status()
            assert status["profile"]["restrict_write_to_project_dir"] is True
            assert str(blocked_dir.resolve()) in status["profile"]["blocked_write_dirs"]
        finally:
            set_client_project_dir(None)
    finally:
        await server.close_workbook(wb_id)


@pytest.mark.asyncio
async def test_imp004_word_and_ppt_ergonomics_parity(tmp_path):
    """Verify IMP-004: optional Word/PPT IDs, Word deferred save & word_live_save dirty guard, and PPT batch_add_slides."""
    from pptx import Presentation

    # Ensure clean controller state
    for did in list(server.word.active_documents):
        await server.close_document(did)
    server.powerpoint.active_presentations.clear()

    # 1. Word optional document_id auto-resolution + deferred dirty tracking + word_live_save guard
    doc_meta = await server.create_document(title="CFO Brief")
    doc_id = doc_meta["document_id"]
    out_docx = tmp_path / "CFO_Brief.docx"
    try:
        assert server.word.active_documents[doc_id]["dirty"] is False
        h_res = await server.add_heading(text="Executive Summary", level=1)
        assert h_res["status"] == "success"
        assert h_res["document_id"] == doc_id
        assert server.word.active_documents[doc_id]["dirty"] is True

        p_res = await server.add_paragraph(text="NPV is $25.0M at 10.5% WACC.")
        assert p_res["status"] == "success"

        t_res = await server.add_table(rows=2, columns=2, data=[["KPI", "Value"], ["IRR", "18.4%"]])
        assert t_res["status"] == "success"

        # word_live_save refuses to overwrite while in-memory edits are dirty (omitted doc name)
        guard_res = await server.word_live_save()
        assert guard_res["status"] == "error"
        assert guard_res["error_type"] == "UnsavedInMemoryEditsError"
        assert "save_document" in guard_res["hint"]

        # Save via omitted document_id clears dirty flag and persists atomically
        save_res = await server.save_document(file_path=str(out_docx))
        assert save_res["status"] == "success"
        assert out_docx.exists()
        assert server.word.active_documents[doc_id]["dirty"] is False

        # Subsequent edit re-dirties CFO_Brief.docx -> word_live_save(document_name="CFO_Brief.docx") is ALSO blocked!
        await server.add_paragraph(text="Appendix: Sensitivity checked.")
        assert server.word.active_documents[doc_id]["dirty"] is True
        guard_named = await server.word_live_save(document_name="CFO_Brief.docx")
        assert guard_named["status"] == "error"
        assert guard_named["error_type"] == "UnsavedInMemoryEditsError"
        await server.save_document()
        assert server.word.active_documents[doc_id]["dirty"] is False

        read_res = await server.read_document()
        assert read_res["status"] == "success"
        assert any("NPV is $25.0M" in p["text"] for p in read_res["paragraphs"])
        assert read_res["table_count"] == 1
    finally:
        close_res = await server.close_document()
        assert close_res["status"] == "success"

    # 2. PowerPoint optional presentation_id auto-resolution + 1-call batch_add_slides
    deck = await server.create_presentation(title="Q4 Boardroom Review")
    pres_id = deck["presentation_id"]
    out_pptx = tmp_path / "Boardroom_Deck.pptx"
    batch_res = await server.batch_add_slides(
        slides=[
            {
                "layout": "title_and_content",
                "title": "Executive Summary",
                "bullets": ["• Revenue grew 14.2% YoY", "- EBITDA margin reached 38.0%"],
                "notes": "Reconciled verbatim to live Excel KPI tab.",
            },
            {
                "layout": "Title and Content",
                "title": "Capital Allocation",
                "content": "• Year-0 CapEx: $25.0M\n• Payback: 3.8 years",
            },
            {
                "layout": "section_header",
                "title": "Risks & Sensitivities",
                "subtitle": "WACC ±200 bps and Volume ±20%",
            },
        ]
    )
    assert batch_res["status"] == "success"
    assert batch_res["presentation_id"] == pres_id
    assert batch_res["slides_added"] == 3
    assert batch_res["total_slides"] == 4

    save_ppt = await server.save_presentation(file_path=str(out_pptx))
    assert save_ppt["status"] == "success"
    assert out_pptx.exists()

    # Verify saved PPTX structure: 16:9 widescreen, 4 slides, single bullets (no double "• •"), speaker notes preserved
    reloaded = Presentation(str(out_pptx))
    assert len(reloaded.slides) == 4
    assert round(reloaded.slide_width.inches, 2) == 13.33
    s1_body = "\n".join(sh.text for sh in reloaded.slides[1].shapes if sh.has_text_frame)
    assert "Revenue grew 14.2% YoY" in s1_body
    assert "• Revenue" not in s1_body  # leading bullet glyph stripped cleanly
    assert "Reconciled verbatim" in reloaded.slides[1].notes_slide.notes_text_frame.text


def test_imp006_modular_facade_parity():
    """Verify IMP-006: ExcelController & AppleScriptBridge compose modular domain mixins cleanly behind 100%-compatible facades."""
    from controllers.excel import (
        ExcelChartsTablesMixin,
        ExcelFormatterMixin,
        ExcelIOMixin,
        ExcelReaderMixin,
        ExcelWriterMixin,
    )
    from controllers.excel_controller import (
        ExcelController,
        _NAMED_CELL_STYLE_PRESETS,
        _TABLE_STYLE_RE,
    )
    from integrations.applescript import (
        AppleScriptBaseMixin,
        AppleScriptExcelCoreMixin,
        AppleScriptExcelIOMixin,
        AppleScriptPowerPointMixin,
        AppleScriptWordMixin,
    )
    from integrations.applescript_bridge import (
        AppleScriptBridge,
        AppleScriptTimeoutError,
        _FS,
        _GS,
        _RS,
        _US,
    )

    for mixin in (ExcelIOMixin, ExcelReaderMixin, ExcelWriterMixin, ExcelChartsTablesMixin, ExcelFormatterMixin):
        assert issubclass(ExcelController, mixin)
    assert "Good" in _NAMED_CELL_STYLE_PRESETS
    assert _TABLE_STYLE_RE.match("TableStyleMedium9")

    for mixin in (
        AppleScriptWordMixin,
        AppleScriptExcelCoreMixin,
        AppleScriptExcelIOMixin,
        AppleScriptPowerPointMixin,
        AppleScriptBaseMixin,
    ):
        assert issubclass(AppleScriptBridge, mixin)
    assert issubclass(AppleScriptTimeoutError, TimeoutError)
    assert (_US, _RS, _GS, _FS) == ("\x1f", "\x1e", "\x1d", "\x1c")


@pytest.mark.asyncio
async def test_imp008_background_mode_focus_protection(tmp_path, monkeypatch):
    """Verify IMP-008: M365_MCP_BACKGROUND_MODE (default no_focus, strict, off) suppresses OS focus stealing and protects same-app user windows."""
    import asyncio
    from integrations.applescript_bridge import AppleScriptBridge
    from utils.validators import get_background_mode, set_client_project_dir

    # 1. Default & env-var parsing (including aliases)
    monkeypatch.delenv("M365_MCP_BACKGROUND_MODE", raising=False)
    monkeypatch.delenv("OFFICE_MCP_BACKGROUND_MODE", raising=False)
    assert get_background_mode() == "no_focus"

    for raw, expected in [
        ("no_focus", "no_focus"),
        ("NO_FOCUS", "no_focus"),
        ("on", "no_focus"),
        ("true", "no_focus"),
        ("strict", "strict"),
        ("HIDDEN", "strict"),
        ("off", "off"),
        ("false", "off"),
        ("0", "off"),
        ("invalid_garbage", "no_focus"),
    ]:
        monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", raw)
        assert get_background_mode() == expected

    # 2. check_office_status surfaces background_mode & available_background_modes
    monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", "no_focus")
    st = await server.check_office_status()
    assert st["profile"]["background_mode"] == "no_focus"
    assert st["profile"]["available_background_modes"] == ["no_focus", "strict", "off"]

    # 3. LaunchServices passes `-g -j` in no_focus/strict even when activate=True, and honors activate=True only in off
    bridge = AppleScriptBridge()
    spawned_cmds = []

    class _FakeProc:
        returncode = 0

        async def communicate(self):
            return b"", b""

    async def fake_subproc(*cmd, **kwargs):
        spawned_cmds.append(list(cmd))
        return _FakeProc()

    monkeypatch.setattr(asyncio, "create_subprocess_exec", fake_subproc)
    dummy_xlsx = tmp_path / "dummy.xlsx"
    dummy_xlsx.write_bytes(b"PK")

    monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", "no_focus")
    await bridge._open_via_launchservices("Microsoft Excel", dummy_xlsx, activate=True)
    assert spawned_cmds[-1][:4] == ["/usr/bin/open", "-g", "-j", "-a"]

    monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", "strict")
    await bridge._open_via_launchservices("Microsoft Excel", dummy_xlsx, activate=True)
    assert spawned_cmds[-1][:4] == ["/usr/bin/open", "-g", "-j", "-a"]

    monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", "off")
    await bridge._open_via_launchservices("Microsoft Excel", dummy_xlsx, activate=True)
    assert spawned_cmds[-1][:2] == ["/usr/bin/open", "-a"]
    await bridge._open_via_launchservices("Microsoft Excel", dummy_xlsx, activate=False)
    assert spawned_cmds[-1][:3] == ["/usr/bin/open", "-g", "-a"]

    # 4. launch_excel / launch_word / launch_powerpoint suppress `activate` when mode != off
    recorded = []

    async def fake_exec(script, args=None, timeout=30.0):
        recorded.append((script, list(args) if args else []))
        if "open workbook" in script or "make new workbook" in script:
            return "dummy.xlsx\x1fSheet1"
        if "open file name" in script or "make new document" in script:
            return "dummy.docx"
        return "Sheet1"

    monkeypatch.setattr(bridge, "execute_applescript", fake_exec)
    monkeypatch.setattr(bridge, "find_open_excel_workbook_by_path", lambda *a, **kw: asyncio.sleep(0, result=None))
    monkeypatch.setattr(bridge, "is_app_running", lambda *a, **kw: asyncio.sleep(0, result=False))

    monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", "no_focus")
    await bridge.launch_excel(activate=True)
    await bridge.launch_word(activate=True)
    await bridge.launch_powerpoint(activate=True)
    assert all("to launch" in s and "to activate" not in s for s, _ in recorded[-3:])

    # 5. open_excel_file, create_excel_workbook, open_word_file, create_word_document, open_powerpoint_file
    #    force shouldActivate="false", pass bg_mode, and include window-isolation + prevFront restoration
    set_client_project_dir(None)
    dummy_docx = tmp_path / "dummy.docx"
    dummy_docx.write_bytes(b"PK")
    dummy_pptx = tmp_path / "dummy.pptx"
    dummy_pptx.write_bytes(b"PK")

    recorded.clear()
    res_ex = await bridge.open_excel_file(str(dummy_xlsx), activate=True)
    assert res_ex["background_mode"] == "no_focus"
    assert recorded[-1][1][2] == "false"  # shouldActivate suppressed
    assert recorded[-1][1][3] == "no_focus"
    assert "set collapsed of (window 1 of wb) to true" in recorded[-1][0]
    assert "wasFrontmost and prevFront" in recorded[-1][0]
    assert "set screen updating to false" in recorded[-1][0]
    assert 'set visible of process "Microsoft Excel" to false' in recorded[-1][0]

    res_ex_new = await bridge.create_excel_workbook(activate=True)
    assert res_ex_new["background_mode"] == "no_focus"
    assert recorded[-1][1] == ["false", "no_focus"]

    monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", "strict")
    res_wd = await bridge.open_word_file(str(dummy_docx), activate=True)
    assert res_wd["background_mode"] == "strict"
    assert recorded[-1][1][2] == "false"
    assert recorded[-1][1][3] == "strict"
    assert "set visible of targetWin to false" in recorded[-1][0]
    assert "if wasFrontmost and prevFront" in recorded[-1][0]
    assert 'set visible of process "Microsoft Word" to false' in recorded[-1][0]

    res_wd_new = await bridge.create_word_document(title="Test", activate=True)
    assert res_wd_new["background_mode"] == "strict"
    assert recorded[-1][1] == ["Test", "false", "strict"]

    await bridge.open_powerpoint_file(str(dummy_pptx), activate=True)
    assert recorded[-1][1][1] == "false"
    assert recorded[-1][1][2] == "strict"
    assert "set frontmost of process prevFrontApp to true" in recorded[-1][0]

    # 5B. Verify view-state restore, worksheet activation, add worksheet, and PDF saves never un-minimize or steal focus
    monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", "no_focus")
    recorded.clear()
    await bridge.activate_live_excel_worksheet("Summary", workbook_name="dummy.xlsx")
    assert "set screen updating to false" in recorded[-1][0]
    assert "set frontmost of process prevFrontApp to true" in recorded[-1][0]

    await bridge.add_live_excel_worksheet("Notes", workbook_name="dummy.xlsx")
    assert "set screen updating to false" in recorded[-1][0]
    assert "set collapsed of wTarget to true" in recorded[-1][0]

    await bridge.restore_excel_workbook_view_state(
        "dummy.xlsx",
        {"active_sheet": "Summary", "bounds": [0, 25, 1920, 1080], "zoom": 100, "collapsed": True, "visible": True},
    )
    assert "set keepIsolated to" in recorded[-1][0]
    assert "if (not keepIsolated) and bStr" in recorded[-1][0]
    assert "if (not keepIsolated) and selAddr" in recorded[-1][0]
    assert recorded[-1][1][7:10] == ["no_focus", "true", "true"]

    await bridge.save_excel_workbook(workbook_name="dummy.xlsx", file_path=str(tmp_path / "out.pdf"), file_format="pdf")
    assert "set frontmost of process prevFrontApp to true" in recorded[-1][0]

    await bridge.save_word_document(document_name="dummy.docx", file_path=str(tmp_path / "out.pdf"), file_format="pdf")
    assert "set frontmost of process prevFrontApp to true" in recorded[-1][0]

    # 5C. Verify live Excel cell/range/batch writes wrap atomically in screen updating = false + enable events = false
    recorded.clear()
    await bridge.write_live_excel_cell("B2", 123, sheet_name="Summary", workbook_name="dummy.xlsx")
    assert "set screen updating to false" in recorded[-1][0]
    assert "set enable events to false" in recorded[-1][0]
    assert "set enable events to prevEvents" in recorded[-1][0]

    await bridge.write_live_excel_range("A1", [[1, 2]], sheet_name="Summary", workbook_name="dummy.xlsx")
    assert "set screen updating to false" in recorded[-1][0]
    assert "set enable events to false" in recorded[-1][0]

    await bridge.batch_write_live_excel_cells({"B2": 456}, sheet_name="Summary", workbook_name="dummy.xlsx")
    assert "set screen updating to false" in recorded[-1][0]
    assert "set enable events to false" in recorded[-1][0]

    # 6. Legacy `off` mode honors activate=True again
    monkeypatch.setenv("M365_MCP_BACKGROUND_MODE", "off")
    recorded.clear()
    await bridge.launch_excel(activate=True)
    assert "to activate" in recorded[-1][0]
    res_ex_off = await bridge.open_excel_file(str(dummy_xlsx), activate=True)
    assert res_ex_off["background_mode"] == "off"
    assert recorded[-1][1][2] == "true"
    assert recorded[-1][1][3] == "off"


@pytest.mark.asyncio
async def test_audit_fixes_sdk_v2_hook_ghost_rows_optional_ids_and_context_caps(tmp_path, monkeypatch):
    """Verify full-codebase audit fixes: SDK v2 _M365MCPServer hook, all 14 extended Excel optional workbook_id,
    ghost 1,048,576-row sparse auto_fit/format/audit, and context-window payload caps across Word & Excel."""
    import time
    from openpyxl import Workbook
    from controllers.excel_controller import ExcelController
    from controllers.word_controller import WordController
    from integrations.applescript_bridge import AppleScriptBridge

    # 1. MCP SDK v2 _M365MCPServer subclass overrides _handle_call_tool at class level
    assert isinstance(server.mcp, server._M365MCPServer)
    assert "_handle_call_tool" in server._M365MCPServer.__dict__

    # 2. ZERO registered tools require workbook_id, document_id, presentation_id, or slide_id
    tm = server.mcp._tool_manager
    for t in tm.list_tools():
        req = (t.parameters or {}).get("required", [])
        assert not [r for r in req if r.endswith("_id")], f"{t.name} still requires ID param: {req}"

    # 3. Ghost 1,048,576-row sheet: _compute_auto_fit_widths, column-only 'B:C' format bounds, and audit_workbook_formulas wb_eval
    ghost_xlsx = tmp_path / "ghost_audit.xlsx"
    wb_g = Workbook()
    ws_g = wb_g.active
    ws_g.title = "ERP"
    ws_g["A1"] = "Metric"
    ws_g["B1"] = "FY25"
    ws_g["A2"] = "Revenue"
    ws_g["B2"] = 1000
    ws_g["A3"] = "Total"
    ws_g["B3"] = "=B2*2"
    wb_g.save(ghost_xlsx)

    ec = ExcelController()
    meta = await ec.open_workbook(str(ghost_xlsx))
    wb_id = meta["workbook_id"]
    ws_mem = ec.active_workbooks[wb_id]["workbook_object"]["ERP"]
    # Simulate a ghost empty formatted cell at row 1,048,576 (classic ERP export bug)
    ws_mem.cell(row=1_048_576, column=2).number_format = "#,##0"
    assert ws_mem.max_row == 1_048_576

    t0 = time.perf_counter()
    widths = ec._compute_auto_fit_widths(ws_mem)
    _, _, _, max_r_bound, norm_ref = ec._resolve_format_bounds(ws_mem, "B:B")
    aud = await ec.audit_workbook_formulas(workbook_id=wb_id)
    elapsed_ms = (time.perf_counter() - t0) * 1000
    assert elapsed_ms < 500.0, f"Ghost-row operations took {elapsed_ms:.1f}ms (expected <500ms)"
    assert "A" in widths and "B" in widths
    assert max_r_bound == 3 and norm_ref == "B1:B3"
    assert aud["total_formulas"] == 1
    await ec.close_workbook(wb_id)

    # 4. Context-window payload caps:
    # 4a. read_range default 150x30 + pagination_hint
    big_meta = await ec.create_workbook("BigRange", sheet_names=["Data"])
    big_id = big_meta["workbook_id"]
    ws_big = ec.active_workbooks[big_id]["workbook_object"]["Data"]
    for r in range(1, 181):
        ws_big.cell(row=r, column=1, value=f"R{r}")
        ws_big.cell(row=r, column=2, value=r)
    rng_res = await ec.read_range(big_id, "Data")
    assert rng_res["rows"] == 150
    assert rng_res["truncated"] is True
    assert rng_res["total_rows"] == 180
    assert "A151" in rng_res["pagination_hint"]

    # 4b. read_tables capped at max_rows_per_table
    await ec.add_table(big_id, "Data", "BigTbl", "A1:B180")
    tbl_res = await ec.read_tables(big_id, "Data", max_rows_per_table=25)
    assert tbl_res["tables"][0]["row_count"] == 179
    assert tbl_res["tables"][0]["returned_rows"] == 25
    assert tbl_res["tables"][0]["truncated"] is True

    # 4c. audit_workbook_formulas caps error_cells at 50 while reporting exact error_count
    for r in range(1, 66):
        ws_big.cell(row=r, column=3, value="=1/0")
        ws_big.cell(row=r, column=4, value="#REF!")
    aud_big = await ec.audit_workbook_formulas(workbook_id=big_id)
    assert aud_big["error_count"] >= 65
    assert len(aud_big["error_cells"]) == 50
    assert aud_big["error_cells_truncated"] is True
    await ec.close_workbook(big_id)

    # 4d. inspect_workbook auto-scales to 15 rows/sheet when >6 sheets unfiltered
    many_meta = await ec.create_workbook("ManyTabs", sheet_names=[f"Tab{i}" for i in range(1, 9)])
    many_id = many_meta["workbook_id"]
    ws_t1 = ec.active_workbooks[many_id]["workbook_object"]["Tab1"]
    for r in range(1, 31):
        ws_t1.cell(row=r, column=1, value=f"Row {r}")
    insp_many = await ec.inspect_workbook(workbook_id=many_id)
    assert len(insp_many["sheets"][0]["values"]) == 15
    assert "preview_note" in insp_many
    await ec.close_workbook(many_id)

    # 4e. Word read_document pagination + table row cap
    wc = WordController()
    d_meta = await wc.create_document("LongDoc")
    d_id = d_meta["document_id"]
    for i in range(12):
        await wc.add_paragraph(d_id, f"Para {i}")
    r_p1 = await wc.read_document(d_id, max_paragraphs=5, start_paragraph=0)
    assert r_p1["paragraph_count"] == 12
    assert len(r_p1["paragraphs"]) == 5
    assert r_p1["truncated"] is True
    assert "start_paragraph=5" in r_p1["pagination_hint"]
    r_p2 = await wc.read_document(d_id, max_paragraphs=5, start_paragraph=5)
    assert r_p2["paragraphs"][0]["index"] == 5
    await wc.close_document(d_id)

    # 4f. Live Word read omits duplicate `text` by default & paginates cleanly
    bridge = AppleScriptBridge()
    fake_paras = "\r".join(f"Line {i}" for i in range(30))
    fake_raw = f"Memo.docx\x1f/tmp/Memo.docx\x1ftrue\x1f0\x1f0\x1f\x1d{fake_paras}\x1d"
    monkeypatch.setattr(bridge, "execute_applescript", lambda *a, **kw: __import__("asyncio").sleep(0, result=fake_raw))
    live_w = await bridge.read_live_word_document(max_paragraphs=10, start_paragraph=0)
    assert "text" not in live_w  # no 2x duplication by default!
    assert live_w["paragraph_count"] == 30
    assert len(live_w["paragraphs"]) == 10
    assert live_w["truncated"] is True
    live_w_full = await bridge.read_live_word_document(include_full_text=True, max_paragraphs=10)
    assert "Line 0" in live_w_full["text"]

    # 4g. Live Excel read trims trailing empty rows/cols & caps >3000 cells
    trimmed, _, trunc_small, _ = bridge._trim_and_cap_live_grid(
        [["A", "B", "", ""], ["1", "2", "", ""], ["", "", "", ""]], None
    )
    assert trimmed == [["A", "B"], ["1", "2"]] and trunc_small is False
    big_live = [["x"] * 10 for _ in range(400)]  # 4,000 cells
    capped, _, trunc_big, orig_r = bridge._trim_and_cap_live_grid(big_live, None, max_cells=3000)
    assert trunc_big is True and orig_r == 400 and len(capped) == 300


def test_mcp_sdk_v2_connection_scoped_session_key():
    """Verify _M365MCPServer._extract_session_key stays stable across per-request ServerSession wrappers."""
    import m365_mac_mcp_server as srv

    class FakeConn:
        pass

    class FakeSessionV2:
        def __init__(self, conn):
            self._connection = conn

    class FakeSessionV1:
        pass

    class FakeCtx:
        def __init__(self, session):
            self.session = session

    conn_a = FakeConn()
    conn_b = FakeConn()
    k1 = srv._M365MCPServer._extract_session_key(FakeCtx(FakeSessionV2(conn_a)))
    k2 = srv._M365MCPServer._extract_session_key(FakeCtx(FakeSessionV2(conn_a)))
    k3 = srv._M365MCPServer._extract_session_key(FakeCtx(FakeSessionV2(conn_b)))
    assert k1 is not None and k1 == k2, "Same connection must share one session key across requests"
    assert k1 != k3, "Distinct connections must remain isolated"

    s_v1 = FakeSessionV1()
    assert srv._M365MCPServer._extract_session_key(FakeCtx(s_v1)) == f"mcp_session_{id(s_v1)}"
    assert srv._M365MCPServer._extract_session_key(FakeCtx(None)) is None


@pytest.mark.asyncio
async def test_phase1_formula_and_linter_guards(tmp_path):
    """Verify missing '=' formula auto-repair, save_workbook formula linters, and Word add_table shape inference."""
    from controllers.excel_controller import ExcelController
    from controllers.word_controller import WordController

    ec = ExcelController()
    wb_res = await ec.create_workbook(title="GuardTest", sheet_names=["Assumptions", "Calc", "Sens"])
    wid = wb_res["workbook_id"]

    # 1. Missing '=' auto-repair for unambiguous Excel formula strings (without touching plain prose)
    await ec.batch_write_ranges(wid, [
        {"sheet_name": "Assumptions", "range": "A1:B2", "data": [["Param", "Value"], ["CapEx", 25000000]]},
        {"sheet_name": "Calc", "range": "A1:C4", "data": [
            ["Line", "Year 0", "Year 1"],
            ["CapEx", 0, "=$B$2/Assumptions!$B$2"],
            ["Disc", 1, "=1/(1+0.1)^C$1"],
            ["Payback", "IF(MAX(B2:C2)<0,\">6 yrs\",COUNTIF(B2:C2,\"<0\"))", "IF(not a formula) plain prose"],
        ]},
        {"sheet_name": "Sens", "range": "A1:C2", "data": [
            ["Scale", 0.1, None],
            [1.0, "=NPV(B$1,Calc!$C$2:$C$2)", "=NPV(C$1,Calc!$C$2:$C$2)"],
        ]},
    ])
    wb_obj = ec.active_workbooks[wid]["workbook_object"]
    assert wb_obj["Calc"]["B4"].value.startswith("=IF(")
    assert wb_obj["Calc"]["C4"].value == "IF(not a formula) plain prose"

    # 2. save_workbook compile-time linter catches all 3 subtle formula bugs
    out_xlsx = tmp_path / "guard_test.xlsx"
    save_res = await ec.save_workbook(wid, str(out_xlsx))
    warns = " | ".join(save_res.get("sanity_warnings") or [])
    assert "missing 'Assumptions!' prefix" in warns, f"Expected bare $B$2 warning, got: {warns}"
    assert "references text 'Year 1'" in warns, f"Expected text exponent warning, got: {warns}"
    assert "NPV rate C1 is blank" in warns, f"Expected blank NPV rate warning, got: {warns}"
    await ec.close_workbook(wid)

    # 3. WordController.add_table infers rows/columns from data when omitted
    wc = WordController()
    d = await wc.create_document(title="TblTest")
    t_res = await wc.add_table(document_id=d["document_id"], data=[["A", "B"], ["1", "2"], ["3", "4"]])
    assert t_res["rows"] == 3 and t_res["columns"] == 2
    await wc.close_document(d["document_id"])

    # 4. batch_write_ranges rejects range-as-fill_formula and combined data+fill_formula atomically
    wb2 = await ec.create_workbook(title="FillRejectTest", sheet_names=["CF"])
    wid2 = wb2["workbook_id"]
    with pytest.raises(ValueError, match="NOT a cell range"):
        await ec.batch_write_ranges(wid2, [
            {"sheet_name": "CF", "range": "A8:H8", "data": [["OpEx", 0, "=C7-C9"]], "fill_formula": "D8:H8"}
        ])
    with pytest.raises(ValueError, match="SEPARATE operations"):
        await ec.batch_write_ranges(wid2, [
            {"sheet_name": "CF", "range": "A8:H8", "data": [["OpEx", 0]], "fill_formula": "=C7-C9"}
        ])
    # Verify atomic rejection left sheet untouched
    assert ec.active_workbooks[wid2]["workbook_object"]["CF"]["A8"].value is None
    # 5. write_cell reports value_type='formula' (not 'str') and cell_map formats formulas cleanly without truncation
    wc_fml = await ec.write_cell(wid2, "CF", "B2", "=1+1")
    assert wc_fml["value_type"] == "formula"
    bw_map = await ec.batch_write_ranges(wid2, [
        {"sheet_name": "CF", "range": "A1:B3", "data": [
            ["KPI", "Value"],
            ["NPV", "=NPV(0.1,C1:H1)+B1"],
            ["IRR", "=IRR(B1:H1)"],
        ]}
    ], allow_formula_overwrite=True)
    cmap = bw_map["results"][0].get("cell_map") or {}
    assert cmap.get("B2") == "NPV: <formula>"
    assert cmap.get("B3") == "IRR: <formula>"
    await ec.close_workbook(wid2)


@pytest.mark.asyncio
async def test_phase2_dcf_formatting_cellmap_and_linters(tmp_path):
    """Verify parenthetical EV/Share Price formatting, bare sheet-link '=' repair, 60-row cell_map, and DCF linters."""
    from controllers.excel_controller import ExcelController

    ec = ExcelController()
    wb = await ec.create_workbook(
        title="DCFGuard",
        sheet_names=["Dashboard", "Assumptions", "Sensitivity"],
        header_metadata={"entity": "Apex", "period": "Y1-Y5", "currency": "USD", "magnitude": "auto"},
    )
    wid = wb["workbook_id"]

    # 1. Bare sheet-link missing '=' ('-Assumptions!$B$7') auto-repairs to '=-Assumptions!$B$7'
    # 2. 35-row Assumptions table (>30 rows) still returns cell_map
    assump_rows = [["Input", "Value"], ["Revenue ($M)", 120], ["Financial Debt ($M)", 120]] + [
        [f"Driver {i}", 0.05] for i in range(32)
    ] + [["Diluted Shares", 50000000]]
    bw = await ec.batch_write_ranges(wid, [
        {"sheet_name": "Assumptions", "range": f"A6:B{5 + len(assump_rows)}", "data": assump_rows},
        {"sheet_name": "Dashboard", "range": "A6:B11", "data": [
            ["Metric", "Value"],
            ["Enterprise Value (Gordon Growth)", "=Assumptions!$B$7*10"],
            ["Equity Value (Exit Multiple)", "=Assumptions!$B$7*8"],
            ["Share Price (Gordon Growth)", "=Assumptions!$B$7/5"],
            ["Less: Financial Debt", "-Assumptions!$B$8"],
            ["WACC", 0.09],
        ]},
        {"sheet_name": "Sensitivity", "range": "A6:C8", "data": [
            ["WACC / g", 0.02, 0.03],
            [0.08, 10, 12],
            [0.09, 9, 11],
        ]},
    ])
    assert bw["results"][0].get("cell_map") is not None and "B7" in bw["results"][0]["cell_map"]
    wb_obj = ec.active_workbooks[wid]["workbook_object"]
    assert wb_obj["Dashboard"]["B10"].value == "=-Assumptions!$B$8"

    # 3. Parenthetical '(Gordon Growth)' and '(Exit Multiple)' do NOT override currency/per-share formats
    await ec.batch_format_workbook(wid, operations=[{"auto_color_financial_model": True, "auto_fit_columns": True}])
    assert "%" not in wb_obj["Dashboard"]["B7"].number_format, f"EV got pct: {wb_obj['Dashboard']['B7'].number_format}"
    assert "x" not in wb_obj["Dashboard"]["B8"].number_format, f"Equity got mult: {wb_obj['Dashboard']['B8'].number_format}"
    assert "$#,##0.00" in wb_obj["Dashboard"]["B9"].number_format, f"Share Price format wrong: {wb_obj['Dashboard']['B9'].number_format}"
    assert wb_obj["Dashboard"]["B11"].number_format == "0.0%"

    # 4. save_workbook warns on 0-formula Sensitivity sheet AND Exact-Shares vs $M mismatch
    sv = await ec.save_workbook(wid, str(tmp_path / "dcf_guard.xlsx"))
    warns = " | ".join(sv.get("sanity_warnings") or [])
    assert "Sensitivity" in warns and "0 formulas" in warns
    assert "Magnitude mismatch" in warns and "50,000,000" in warns
    await ec.close_workbook(wid)


@pytest.mark.asyncio
async def test_imp010_mandatory_playbook_and_workbook_progress():
    """Verify check_office_status wraps skill_playbook in <mandatory_playbook> and batch_write_ranges returns workbook_progress."""
    st = await server.check_office_status(skill_name="capex-project-analyzer")
    _assert_strict_json(st)
    assert st["status"] == "running"
    sp = st["loaded_skill"]
    assert sp["name"] == "capex-project-analyzer"
    assert sp["content"].startswith('<mandatory_playbook name="capex-project-analyzer"')
    assert sp["content"].endswith("</mandatory_playbook>")
    assert "precedence_rule" in sp and "strictly override" in sp["precedence_rule"]

    wb = await server.create_workbook(
        title="ProgressCheck",
        sheet_names=["Executive Summary", "Assumptions", "Project Cash Flows", "Sensitivity & Charts"],
    )
    wid = wb["workbook_id"]
    bw = await server.batch_write_ranges(
        wid,
        operations=[
            {
                "sheet_name": "Assumptions",
                "range": "A1:B3",
                "data": [["Parameter", "Value"], ["CapEx", 25000000], ["WACC", 0.10]],
            }
        ],
    )
    _assert_strict_json(bw)
    prog = bw.get("workbook_progress")
    assert isinstance(prog, dict)
    assert prog["Executive Summary"] == "empty"
    assert "cells" in prog["Assumptions"]
    assert prog["Project Cash Flows"] == "empty"
    assert prog["Sensitivity & Charts"] == "empty"
    await server.close_workbook(wid)







