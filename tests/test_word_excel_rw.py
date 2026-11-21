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

import office365_mcp_server as server
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
    """Verify OFFICE_MCP_PROFILE=compact exposes the lean core Word+Excel toolset cleanly."""
    import os
    env = dict(os.environ)
    env["OFFICE_MCP_PROFILE"] = "compact"
    proc = await asyncio.create_subprocess_exec(
        sys.executable,
        "-c",
        (
            "import asyncio, json, sys; sys.path.insert(0, 'src'); "
            "import office365_mcp_server as s; "
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

    # Verify composable app profiles: excel,compact (~31), word,powerpoint (~33), and invalid fallback
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
        env_p = dict(os.environ, OFFICE_MCP_PROFILE=prof)
        p2 = await asyncio.create_subprocess_exec(
            sys.executable,
            "-c",
            (
                "import asyncio, json, sys; sys.path.insert(0, 'src'); "
                "import office365_mcp_server as s; "
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
            assert "Unknown OFFICE_MCP_PROFILE" in err2.decode("utf-8", errors="replace")

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
    """Verify OFFICE_MCP_LIVE_SYNC (AUTO/Auto/auto/OFF/Off/off) auto-reloads saved files and guards unsaved edits."""
    # 1. Case-insensitive parsing & default='auto'
    monkeypatch.delenv("OFFICE_MCP_LIVE_SYNC", raising=False)
    assert ExcelController.get_live_sync_mode() == "auto"
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

    monkeypatch.setenv("OFFICE_MCP_LIVE_SYNC", "Auto")
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
    monkeypatch.setenv("OFFICE_MCP_LIVE_SYNC", "OFF")
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
    """Verify OFFICE_MCP_FORMULA_GUARD (literals_only default | strict | off) across all 6 Excel write tools."""
    from controllers.excel_controller import ExcelController

    # 1. Case-insensitive mode parsing + fallback
    monkeypatch.delenv("OFFICE_MCP_FORMULA_GUARD", raising=False)
    assert ExcelController.get_formula_guard_mode() == "literals_only"
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
    monkeypatch.setenv("OFFICE_MCP_FORMULA_GUARD", "STRICT")
    blk_strict = await server.write_cell(wb_id, "Model", "C3", "=C2*3")
    assert blk_strict["status"] == "confirmation_required"
    assert blk_strict["guard_mode"] == "strict"

    # 4. 'off' mode allows everything
    monkeypatch.setenv("OFFICE_MCP_FORMULA_GUARD", "off")
    ok_off = await server.write_cell(wb_id, "Model", "C3", 240)
    assert ok_off["status"] == "success"
    await server.close_workbook(wb_id)

    # 5. Live Excel writers parse BLOCKED responses & pass guardMode arg cleanly
    monkeypatch.setenv("OFFICE_MCP_FORMULA_GUARD", "literals_only")
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
    import re

    skills_root = Path(__file__).resolve().parent.parent / "skills"
    skill_files = sorted(skills_root.glob("*/SKILL.md"))
    assert len(skill_files) >= 7

    tools_map = server.mcp._tool_manager._tools
    valid_tools = set(tools_map.keys())

    for p in skill_files:
        txt = p.read_text(encoding="utf-8")
        size_bytes = len(txt.encode("utf-8"))
        assert 2048 <= size_bytes <= 4096, f"{p.parent.name} size {size_bytes}B outside 2-4 KB target"
        assert txt.startswith("---\nname:") and "\ndescription:" in txt, f"{p.parent.name} missing YAML frontmatter"
        for idx in range(1, 6):
            assert f"## {idx}." in txt, f"{p.parent.name} missing section ## {idx}."
        assert "IFRS" in txt and "FASB" in txt and "SOX" in txt, f"{p.parent.name} missing IFRS/FASB/SOX compliance"

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
