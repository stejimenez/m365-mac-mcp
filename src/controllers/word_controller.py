"""
Word Controller for Microsoft 365 for Mac MCP (m365-mac-mcp).

Provides full, non-blocking file-level read-write operations for Microsoft Word
(.docx, and .doc via live Word conversion) using python-docx, plus optional
live Word integration (PDF/RTF/TXT export & open-in-app) via AppleScriptBridge.
"""

import asyncio
import atexit
import os
import re
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from integrations.applescript_bridge import AppleScriptBridge
from utils.logger import setup_logger
from utils.validators import resolve_file_path, sanitize_for_json

logger = setup_logger(__name__)


class WordController:
    """Controller for Word file-level read-write operations with optional live Word hooks."""

    def __init__(self, max_documents: int = 20):
        self.applescript = AppleScriptBridge()
        self.active_documents: Dict[str, Dict[str, Any]] = {}
        self.max_documents = max_documents
        self.temp_dir = Path(tempfile.gettempdir()) / "m365-mac-mcp" / "word"
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self._owned_temp_files: set[Path] = set()
        atexit.register(self._cleanup_temp_files)

    def _cleanup_temp_files(self) -> None:
        """Remove controller-owned temporary files on shutdown."""
        for path in list(self._owned_temp_files):
            try:
                if path.exists():
                    path.unlink()
            except Exception:
                pass
            self._owned_temp_files.discard(path)

    def _evict_if_needed(self) -> None:
        """Keep active document count bounded to prevent unbounded memory growth."""
        while len(self.active_documents) >= self.max_documents:
            oldest_id = next(iter(self.active_documents))
            entry = self.active_documents.pop(oldest_id)
            temp_path = entry.get("temp_file")
            if temp_path and temp_path in self._owned_temp_files:
                try:
                    if temp_path.exists():
                        temp_path.unlink()
                except Exception:
                    pass
                self._owned_temp_files.discard(temp_path)
            logger.info(f"Evicted oldest in-memory document {oldest_id} to stay within limit ({self.max_documents})")

    def _get_entry(self, document_id: str) -> Dict[str, Any]:
        if document_id not in self.active_documents:
            raise ValueError(f"Document '{document_id}' not found. Open or create a document first.")
        return self.active_documents[document_id]

    async def _save_working_copy(self, entry: Dict[str, Any]) -> None:
        """Persist working copy asynchronously without blocking the event loop."""
        doc = entry["docx_object"]
        target = entry["working_path"]
        await asyncio.to_thread(doc.save, str(target))

    # ------------------------------------------------------------------
    # Create / Open / Close
    # ------------------------------------------------------------------

    async def create_document(
        self,
        title: str = "New Document",
        template_path: Optional[str] = None,
        add_title_heading: bool = False,
        open_in_word: bool = False,
    ) -> Dict[str, Any]:
        """Create a new Word document (.docx)."""
        self._evict_if_needed()
        document_id = str(uuid.uuid4())

        if template_path:
            tpl = resolve_file_path(template_path, must_exist=True, extensions=[".docx", ".dotx"])
            doc = await asyncio.to_thread(Document, str(tpl))
        else:
            doc = Document()

        if title:
            doc.core_properties.title = title
            if add_title_heading and title != "New Document":
                doc.add_heading(title, level=1)

        temp_file = self.temp_dir / f"{document_id}.docx"
        await asyncio.to_thread(doc.save, str(temp_file))
        self._owned_temp_files.add(temp_file)

        opened_in_app = False
        open_note: Optional[str] = None
        if open_in_word:
            # Opening unnamed /tmp UUID files in sandboxed Microsoft Word triggers macOS
            # "Grant Access" dialogs. Keep the temp file as working fallback and instruct the
            # model to save into the chat's configured workspace folder before opening live.
            open_note = (
                "Ignored open_in_word=True on unnamed temporary file to avoid macOS Word "
                "sandbox 'Grant Access' prompts. Save the document first with save_document into "
                "the chat's configured workspace folder (or ~/Desktop / ~/Documents), then call "
                "word_live_open."
            )
            logger.info(open_note)

        metadata = {
            "document_id": document_id,
            "title": title,
            "file_path": str(temp_file),
            "source_path": None,
            "is_temporary": True,
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables),
            "opened_in_word": opened_in_app,
            "created_at": time.time(),
        }
        if open_note:
            metadata["note"] = open_note

        self.active_documents[document_id] = {
            "metadata": metadata,
            "docx_object": doc,
            "working_path": temp_file,
            "temp_file": temp_file,
        }
        logger.info(f"Created Word document '{title}' ({document_id})")
        return sanitize_for_json(metadata)

    async def open_document(self, file_path: str) -> Dict[str, Any]:
        """Open an existing Word document (.docx, or .doc converted via Word)."""
        self._evict_if_needed()
        resolved = resolve_file_path(file_path, must_exist=True, extensions=[".docx", ".doc", ".dotx"])
        document_id = str(uuid.uuid4())
        temp_file: Optional[Path] = None

        if resolved.suffix.lower() == ".doc":
            # python-docx only reads OOXML (.docx); convert legacy .doc via live Word AppleScript
            converted = self.temp_dir / f"{document_id}_converted.docx"
            open_res = await self.applescript.open_word_file(str(resolved), read_only=True, activate=False)
            live_name = open_res["document_name"]
            try:
                await self.applescript.save_word_document(
                    document_name=live_name,
                    file_path=str(converted),
                    file_format="docx",
                )
            finally:
                try:
                    await self.applescript.close_word_document(document_name=live_name, save_changes="no")
                except Exception:
                    pass
            doc = await asyncio.to_thread(Document, str(converted))
            working_path = converted
            temp_file = converted
            self._owned_temp_files.add(converted)
        else:
            doc = await asyncio.to_thread(Document, str(resolved))
            working_path = resolved

        title = doc.core_properties.title or resolved.stem
        metadata = {
            "document_id": document_id,
            "title": title,
            "file_path": str(resolved),
            "source_path": str(resolved),
            "is_temporary": resolved.suffix.lower() == ".doc",
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables),
            "opened_at": time.time(),
        }

        self.active_documents[document_id] = {
            "metadata": metadata,
            "docx_object": doc,
            "working_path": working_path,
            "temp_file": temp_file,
        }
        logger.info(f"Opened Word document '{resolved}' ({document_id})")
        return sanitize_for_json(metadata)

    async def close_document(
        self,
        document_id: str,
        save_before_close: bool = False,
    ) -> Dict[str, Any]:
        """Close an in-memory Word document and clean up any temporary file."""
        entry = self._get_entry(document_id)
        saved_path = None
        if save_before_close and not entry["metadata"].get("is_temporary", True):
            await self._save_working_copy(entry)
            saved_path = str(entry["working_path"])

        self.active_documents.pop(document_id, None)
        temp_file = entry.get("temp_file")
        if temp_file and temp_file in self._owned_temp_files:
            try:
                if temp_file.exists():
                    temp_file.unlink()
            except Exception:
                pass
            self._owned_temp_files.discard(temp_file)

        return {
            "status": "success",
            "document_id": document_id,
            "saved_before_close": bool(saved_path),
            "saved_path": saved_path,
        }

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------

    async def read_document(
        self,
        document_id: str,
        include_tables: bool = True,
        include_styles: bool = False,
        include_metadata: bool = True,
        include_full_text: bool = False,
    ) -> Dict[str, Any]:
        """Read indexed paragraphs, tables, styles, and non-empty metadata from a Word document."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        paragraphs = []
        for idx, para in enumerate(doc.paragraphs):
            style_name = para.style.name if para.style else "Normal"
            item: Dict[str, Any] = {
                "index": idx,
                "text": para.text,
            }
            if style_name and style_name != "Normal":
                item["style"] = style_name
            paragraphs.append(item)

        result: Dict[str, Any] = {
            "status": "success",
            "document_id": document_id,
            "file_path": entry["metadata"]["file_path"],
            "paragraph_count": len(paragraphs),
            "paragraphs": paragraphs,
        }
        if include_full_text:
            result["full_text"] = "\n".join(p["text"] for p in paragraphs)

        if include_tables:
            tables = []
            for t_idx, table in enumerate(doc.tables):
                rows_data = [
                    [cell.text for cell in row.cells]
                    for row in table.rows
                ]
                tbl_entry: Dict[str, Any] = {
                    "table_index": t_idx,
                    "rows": len(table.rows),
                    "columns": len(table.columns),
                    "data": rows_data,
                }
                if table.style and table.style.name:
                    tbl_entry["style"] = table.style.name
                tables.append(tbl_entry)
            result["table_count"] = len(tables)
            if tables:
                result["tables"] = tables

        if include_styles:
            result["available_styles"] = sorted(
                {s.name for s in doc.styles if s.name}
            )

        if include_metadata:
            cp = doc.core_properties
            raw_props = {
                "title": cp.title or "",
                "author": cp.author or "",
                "subject": cp.subject or "",
                "keywords": cp.keywords or "",
                "comments": cp.comments or "",
                "created": cp.created,
                "modified": cp.modified,
                "last_modified_by": cp.last_modified_by or "",
            }
            props = {k: v for k, v in raw_props.items() if v not in (None, "")}
            if props:
                result["properties"] = props

        return sanitize_for_json(result)

    # ------------------------------------------------------------------
    # Writing & Editing
    # ------------------------------------------------------------------

    async def add_heading(
        self,
        document_id: str,
        text: str,
        level: int = 1,
        style: Optional[str] = None,
        formatting: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Add a heading to a document with optional explicit formatting."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]
        clamped_level = max(0, min(9, int(level)))

        heading = doc.add_heading(text, level=clamped_level)
        if style:
            try:
                heading.style = style
            except Exception as exc:
                logger.warning(f"Could not apply heading style '{style}': {exc}")

        if formatting:
            self._apply_paragraph_formatting(heading, formatting)

        entry["metadata"]["paragraph_count"] = len(doc.paragraphs)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "document_id": document_id,
            "paragraph_index": len(doc.paragraphs) - 1,
            "type": "heading",
            "text": text,
            "level": clamped_level,
            "style": heading.style.name if heading.style else style,
        }

    async def add_paragraph(
        self,
        document_id: str,
        text: str,
        style: Optional[str] = None,
        formatting: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Append a formatted paragraph to a document."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        paragraph = doc.add_paragraph(text)
        if style:
            try:
                paragraph.style = style
            except Exception as exc:
                logger.warning(f"Could not apply paragraph style '{style}': {exc}")

        if formatting:
            self._apply_paragraph_formatting(paragraph, formatting)

        entry["metadata"]["paragraph_count"] = len(doc.paragraphs)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "document_id": document_id,
            "paragraph_index": len(doc.paragraphs) - 1,
            "type": "paragraph",
            "text": text,
            "style": paragraph.style.name if paragraph.style else style,
        }

    async def insert_paragraph(
        self,
        document_id: str,
        index: int,
        text: str,
        style: Optional[str] = None,
        formatting: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Insert a paragraph before the given 0-based paragraph index (or at end)."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        if index < 0 or index >= len(doc.paragraphs):
            paragraph = doc.add_paragraph(text)
            inserted_index = len(doc.paragraphs) - 1
        else:
            paragraph = doc.paragraphs[index].insert_paragraph_before(text)
            inserted_index = index

        if style:
            try:
                paragraph.style = style
            except Exception as exc:
                logger.warning(f"Could not apply style '{style}': {exc}")

        if formatting:
            self._apply_paragraph_formatting(paragraph, formatting)

        entry["metadata"]["paragraph_count"] = len(doc.paragraphs)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "document_id": document_id,
            "paragraph_index": inserted_index,
            "text": text,
            "style": paragraph.style.name if paragraph.style else style,
        }

    async def replace_text(
        self,
        document_id: str,
        find_text: str,
        replace_text: str,
        match_case: bool = True,
        include_tables: bool = True,
    ) -> Dict[str, Any]:
        """Replace occurrences of text across paragraphs and (optionally) tables."""
        if not find_text:
            raise ValueError("find_text cannot be empty")
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        flags = 0 if match_case else re.IGNORECASE
        pattern = re.compile(re.escape(find_text), flags)
        replacements = 0

        def _replace_in_paragraph(para) -> int:
            nonlocal pattern
            count = 0
            # Fast path: replace within individual runs to preserve run-level formatting
            run_hits = False
            for run in para.runs:
                if run.text and pattern.search(run.text):
                    new_txt, n = pattern.subn(replace_text, run.text)
                    run.text = new_txt
                    count += n
                    run_hits = True
            # Fallback: text spans multiple runs
            if not run_hits and para.text and pattern.search(para.text):
                new_txt, n = pattern.subn(replace_text, para.text)
                if para.runs:
                    para.runs[0].text = new_txt
                    for r in para.runs[1:]:
                        r.text = ""
                else:
                    para.text = new_txt
                count += n
            return count

        for para in doc.paragraphs:
            replacements += _replace_in_paragraph(para)

        if include_tables:
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            replacements += _replace_in_paragraph(para)

        if replacements > 0:
            await self._save_working_copy(entry)

        return {
            "status": "success",
            "document_id": document_id,
            "find_text": find_text,
            "replace_text": replace_text,
            "replacements": replacements,
        }

    async def delete_paragraph(self, document_id: str, index: int) -> Dict[str, Any]:
        """Delete a paragraph by its 0-based index."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        if index < 0 or index >= len(doc.paragraphs):
            raise IndexError(
                f"Paragraph index {index} out of range (0..{len(doc.paragraphs) - 1})"
            )

        para = doc.paragraphs[index]
        deleted_text = para.text
        p_elem = para._element
        p_elem.getparent().remove(p_elem)

        entry["metadata"]["paragraph_count"] = len(doc.paragraphs)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "document_id": document_id,
            "deleted_index": index,
            "deleted_text": deleted_text,
            "remaining_paragraphs": len(doc.paragraphs),
        }

    async def format_paragraph(
        self,
        document_id: str,
        index: int,
        formatting: Dict[str, Any],
        style: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Apply explicit style and formatting to an existing paragraph by 0-based index."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        if index < 0 or index >= len(doc.paragraphs):
            raise IndexError(
                f"Paragraph index {index} out of range (0..{len(doc.paragraphs) - 1})"
            )

        para = doc.paragraphs[index]
        if style:
            try:
                para.style = style
            except Exception as exc:
                logger.warning(f"Could not apply style '{style}': {exc}")

        if formatting:
            self._apply_paragraph_formatting(para, formatting)

        await self._save_working_copy(entry)
        return {
            "status": "success",
            "document_id": document_id,
            "paragraph_index": index,
            "style": para.style.name if para.style else style,
        }

    async def add_list(
        self,
        document_id: str,
        items: List[str],
        list_type: str = "bullet",
        style: Optional[str] = None,
        formatting: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Add a bulleted or numbered list to a document."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        target_style = style or ("List Number" if list_type == "number" else "List Bullet")
        for item in items:
            try:
                para = doc.add_paragraph(str(item), style=target_style)
            except Exception:
                prefix = "• " if list_type != "number" else ""
                para = doc.add_paragraph(f"{prefix}{item}")
            if formatting:
                self._apply_paragraph_formatting(para, formatting)

        entry["metadata"]["paragraph_count"] = len(doc.paragraphs)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "document_id": document_id,
            "type": "list",
            "list_type": list_type,
            "item_count": len(items),
        }

    async def add_table(
        self,
        document_id: str,
        rows: int,
        columns: int,
        data: Optional[List[List[Any]]] = None,
        style: Optional[str] = None,
        column_widths: Optional[List[float]] = None,
        header_formatting: Optional[Dict[str, Any]] = None,
        body_formatting: Optional[Dict[str, Any]] = None,
        alternating_row_color: Optional[str] = None,
        borders: Optional[Dict[str, Any]] = None,
        cell_padding: Optional[Dict[str, float]] = None,
        repeat_header: bool = False,
        table_alignment: Optional[str] = None,
        column_alignments: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Add a table to a document and optionally populate and style it."""
        if rows <= 0 or columns <= 0:
            raise ValueError("rows and columns must be positive integers")

        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        table = doc.add_table(rows=rows, cols=columns)
        if style:
            try:
                table.style = style
            except Exception as exc:
                logger.warning(f"Could not apply table style '{style}': {exc}")

        if data:
            for r_idx, row_data in enumerate(data[:rows]):
                for c_idx, cell_val in enumerate(row_data[:columns]):
                    table.cell(r_idx, c_idx).text = "" if cell_val is None else str(cell_val)

        if any(
            x is not None and x is not False
            for x in (
                column_widths,
                header_formatting,
                body_formatting,
                alternating_row_color,
                borders,
                cell_padding,
                repeat_header,
                table_alignment,
                column_alignments,
            )
        ):
            self._apply_table_styling(
                table,
                column_widths=column_widths,
                header_formatting=header_formatting,
                body_formatting=body_formatting,
                alternating_row_color=alternating_row_color,
                borders=borders,
                cell_padding=cell_padding,
                repeat_header=repeat_header,
                table_alignment=table_alignment,
                column_alignments=column_alignments,
            )

        entry["metadata"]["table_count"] = len(doc.tables)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "document_id": document_id,
            "table_index": len(doc.tables) - 1,
            "rows": rows,
            "columns": columns,
            "has_data": bool(data),
            "styled": True,
        }

    async def format_table(
        self,
        document_id: str,
        table_index: int,
        style: Optional[str] = None,
        column_widths: Optional[List[float]] = None,
        header_formatting: Optional[Dict[str, Any]] = None,
        body_formatting: Optional[Dict[str, Any]] = None,
        alternating_row_color: Optional[str] = None,
        borders: Optional[Dict[str, Any]] = None,
        cell_padding: Optional[Dict[str, float]] = None,
        repeat_header: bool = False,
        table_alignment: Optional[str] = None,
        column_alignments: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Apply explicit styling, shading, borders, padding, and column widths to an existing table."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        if table_index < 0 or table_index >= len(doc.tables):
            raise IndexError(f"table_index {table_index} out of range (0..{len(doc.tables) - 1})")
        table = doc.tables[table_index]

        if style:
            try:
                table.style = style
            except Exception as exc:
                logger.warning(f"Could not apply table style '{style}': {exc}")

        self._apply_table_styling(
            table,
            column_widths=column_widths,
            header_formatting=header_formatting,
            body_formatting=body_formatting,
            alternating_row_color=alternating_row_color,
            borders=borders,
            cell_padding=cell_padding,
            repeat_header=repeat_header,
            table_alignment=table_alignment,
            column_alignments=column_alignments,
        )

        await self._save_working_copy(entry)
        return {
            "status": "success",
            "document_id": document_id,
            "table_index": table_index,
            "rows": len(table.rows),
            "columns": len(table.columns),
        }

    async def update_table_cell(
        self,
        document_id: str,
        table_index: int,
        row: int,
        column: int,
        text: str,
        formatting: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Update the text (and optional formatting) of a specific table cell (0-based indices)."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]

        if table_index < 0 or table_index >= len(doc.tables):
            raise IndexError(f"table_index {table_index} out of range (0..{len(doc.tables) - 1})")
        table = doc.tables[table_index]
        if row < 0 or row >= len(table.rows):
            raise IndexError(f"row {row} out of range (0..{len(table.rows) - 1})")
        if column < 0 or column >= len(table.columns):
            raise IndexError(f"column {column} out of range (0..{len(table.columns) - 1})")

        cell = table.cell(row, column)
        cell.text = "" if text is None else str(text)
        if formatting:
            self._apply_cell_formatting(cell, formatting)
        await self._save_working_copy(entry)

        return {
            "status": "success",
            "document_id": document_id,
            "table_index": table_index,
            "row": row,
            "column": column,
            "text": text,
        }

    async def set_document_properties(
        self,
        document_id: str,
        title: Optional[str] = None,
        author: Optional[str] = None,
        subject: Optional[str] = None,
        comments: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Update document core metadata properties."""
        entry = self._get_entry(document_id)
        cp = entry["docx_object"].core_properties
        if title is not None:
            cp.title = title
            entry["metadata"]["title"] = title
        if author is not None:
            cp.author = author
        if subject is not None:
            cp.subject = subject
        if comments is not None:
            cp.comments = comments
        await self._save_working_copy(entry)
        return {
            "status": "success",
            "document_id": document_id,
            "properties": {
                "title": cp.title or "",
                "author": cp.author or "",
                "subject": cp.subject or "",
                "comments": cp.comments or "",
            },
        }

    # ------------------------------------------------------------------
    # Save / Export
    # ------------------------------------------------------------------

    async def save_document(
        self,
        document_id: str,
        file_path: Optional[str] = None,
        format: str = "docx",
    ) -> Dict[str, Any]:
        """Save a Word document in-place or to a new path (.docx, .pdf, .rtf, .txt, .doc)."""
        entry = self._get_entry(document_id)
        doc = entry["docx_object"]
        fmt = format.lower().lstrip(".")

        if file_path:
            save_path = resolve_file_path(file_path, must_exist=False)
        else:
            if entry["metadata"].get("is_temporary", True):
                raise ValueError(
                    "Document has no saved path yet; provide file_path to save."
                )
            save_path = Path(entry["working_path"])

        if not save_path.suffix:
            save_path = save_path.with_suffix(f".{fmt}")
        save_path.parent.mkdir(parents=True, exist_ok=True)

        if fmt == "docx":
            await asyncio.to_thread(doc.save, str(save_path))
            entry["working_path"] = save_path
            entry["metadata"]["file_path"] = str(save_path)
            entry["metadata"]["is_temporary"] = False
            logger.info(f"Saved Word document to {save_path}")
            return {
                "status": "success",
                "document_id": document_id,
                "file_path": str(save_path),
                "format": "docx",
            }

        if fmt in ("pdf", "rtf", "txt", "doc"):
            # Persist current docx first, then export cleanly via live Microsoft Word AppleScript.
            # NOTE: Microsoft Word 16.x rejects `save as` on read_only=True documents.
            await self._save_working_copy(entry)
            open_res = await self.applescript.open_word_file(
                str(entry["working_path"]), read_only=False, activate=False
            )
            live_name = open_res["document_name"]
            try:
                await self.applescript.save_word_document(
                    document_name=live_name,
                    file_path=str(save_path),
                    file_format=fmt,
                )
            finally:
                try:
                    await self.applescript.close_word_document(
                        document_name=live_name, save_changes="no"
                    )
                except Exception:
                    pass
            return {
                "status": "success",
                "document_id": document_id,
                "file_path": str(save_path),
                "format": fmt,
                "method": "applescript_export",
            }

        raise ValueError(f"Unsupported format: '{format}'. Supported: docx, pdf, rtf, txt, doc.")

    # ------------------------------------------------------------------
    # Formatting & Metadata Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _normalize_hex_color(color: Any) -> Optional[str]:
        """Normalize a hex color ('#1F4E79' or '1f4e79') to 6-char uppercase RGB."""
        if color is None:
            return None
        raw = str(color).strip().lstrip("#").upper()
        if len(raw) == 3 and all(ch in "0123456789ABCDEF" for ch in raw):
            return "".join(ch * 2 for ch in raw)
        if len(raw) == 6 and all(ch in "0123456789ABCDEF" for ch in raw):
            return raw
        logger.warning(f"Invalid hex color: {color}")
        return None

    @classmethod
    def _set_cell_shading(cls, cell, hex_color: str) -> None:
        norm = cls._normalize_hex_color(hex_color)
        if not norm:
            return
        tcPr = cell._tc.get_or_add_tcPr()
        for existing in tcPr.findall(qn("w:shd")):
            tcPr.remove(existing)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), norm)
        tcPr.append(shd)

    @staticmethod
    def _set_cell_margins(
        cell,
        top: Optional[float] = None,
        bottom: Optional[float] = None,
        left: Optional[float] = None,
        right: Optional[float] = None,
    ) -> None:
        """Set cell padding in points (converted to twips: 1pt = 20 dxa)."""
        tcPr = cell._tc.get_or_add_tcPr()
        for existing in tcPr.findall(qn("w:tcMar")):
            tcPr.remove(existing)
        tcMar = OxmlElement("w:tcMar")
        for edge, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
            if val is not None:
                node = OxmlElement(f"w:{edge}")
                node.set(qn("w:w"), str(int(round(float(val) * 20))))
                node.set(qn("w:type"), "dxa")
                tcMar.append(node)
        tcPr.append(tcMar)

    @classmethod
    def _set_table_borders(cls, table, borders: Dict[str, Any]) -> None:
        """Set XML borders on a Word table (supports uniform or per-edge definitions)."""
        tblPr = table._tbl.tblPr
        for existing in tblPr.findall(qn("w:tblBorders")):
            tblPr.remove(existing)
        tblBorders = OxmlElement("w:tblBorders")

        default_style = str(borders.get("style", "single"))
        default_color = cls._normalize_hex_color(borders.get("color")) or "BFBFBF"
        default_size = str(int(borders.get("size", 4)))

        edges = ("top", "left", "bottom", "right", "insideH", "insideV")
        for edge in edges:
            spec = borders.get(edge)
            if spec is None:
                edge_style, edge_color, edge_size = default_style, default_color, default_size
            elif isinstance(spec, str):
                edge_style, edge_color, edge_size = spec, default_color, default_size
            elif isinstance(spec, dict):
                edge_style = str(spec.get("style", default_style))
                edge_color = cls._normalize_hex_color(spec.get("color")) or default_color
                edge_size = str(int(spec.get("size", default_size)))
            elif spec is False:
                edge_style, edge_color, edge_size = "none", "auto", "0"
            else:
                edge_style, edge_color, edge_size = default_style, default_color, default_size

            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), edge_style)
            if edge_style != "none":
                el.set(qn("w:sz"), edge_size)
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), edge_color)
            tblBorders.append(el)
        tblPr.append(tblBorders)

    @staticmethod
    def _set_row_repeat_header(row) -> None:
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))

    def _apply_cell_formatting(self, cell, formatting: Dict[str, Any]) -> None:
        """Apply fill, vertical alignment, and paragraph/run formatting to a Word table cell."""
        raw_bg = formatting.get("bg_color", formatting.get("fill_color", formatting.get("shading")))
        if raw_bg:
            self._set_cell_shading(cell, raw_bg)

        if "vertical_alignment" in formatting:
            valign_map = {
                "top": WD_ALIGN_VERTICAL.TOP,
                "center": WD_ALIGN_VERTICAL.CENTER,
                "bottom": WD_ALIGN_VERTICAL.BOTTOM,
            }
            v_key = str(formatting["vertical_alignment"]).lower()
            if v_key in valign_map:
                cell.vertical_alignment = valign_map[v_key]

        para_fmt = {k: v for k, v in formatting.items() if k not in ("bg_color", "fill_color", "shading", "vertical_alignment")}
        if para_fmt:
            for para in cell.paragraphs:
                self._apply_paragraph_formatting(para, para_fmt)

    def _apply_table_styling(
        self,
        table,
        column_widths: Optional[List[float]] = None,
        header_formatting: Optional[Dict[str, Any]] = None,
        body_formatting: Optional[Dict[str, Any]] = None,
        alternating_row_color: Optional[str] = None,
        borders: Optional[Dict[str, Any]] = None,
        cell_padding: Optional[Dict[str, float]] = None,
        repeat_header: bool = False,
        table_alignment: Optional[str] = None,
        column_alignments: Optional[List[str]] = None,
    ) -> None:
        """Apply comprehensive styling to a python-docx Table."""
        if table_alignment:
            t_align_map = {
                "left": WD_TABLE_ALIGNMENT.LEFT,
                "center": WD_TABLE_ALIGNMENT.CENTER,
                "right": WD_TABLE_ALIGNMENT.RIGHT,
            }
            t_key = str(table_alignment).lower()
            if t_key in t_align_map:
                table.alignment = t_align_map[t_key]

        if column_widths:
            table.autofit = False
            numeric_widths = [float(v) for v in column_widths if v is not None and float(v) > 0]
            max_w = max(numeric_widths) if numeric_widths else 0.0
            # Universal unit guard: Word max page width is 22". LLMs often pass points (e.g. [120, 100, 80])
            # or twips (e.g. [2400, 1800, 1400]) instead of inches; auto-convert so Word never overflows twips or fails PDF export.
            if max_w > 500.0:
                divisor = 1440.0  # twips -> inches
            elif max_w > 15.0:
                divisor = 72.0    # points -> inches
            else:
                divisor = 1.0     # already inches
            for c_idx, width_raw in enumerate(column_widths):
                if c_idx < len(table.columns) and width_raw is not None:
                    width_in = max(0.3, min(10.0, float(width_raw) / divisor))
                    w = Inches(width_in)
                    table.columns[c_idx].width = w
                    for row in table.rows:
                        row.cells[c_idx].width = w

        if borders is not None:
            self._set_table_borders(table, borders)

        if repeat_header and len(table.rows) > 0:
            self._set_row_repeat_header(table.rows[0])

        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                if cell_padding:
                    self._set_cell_margins(
                        cell,
                        top=cell_padding.get("top"),
                        bottom=cell_padding.get("bottom"),
                        left=cell_padding.get("left"),
                        right=cell_padding.get("right"),
                    )

                if r_idx == 0 and header_formatting:
                    fmt = dict(header_formatting)
                    if column_alignments and c_idx < len(column_alignments) and "alignment" not in fmt:
                        fmt["alignment"] = column_alignments[c_idx]
                    self._apply_cell_formatting(cell, fmt)
                elif r_idx > 0:
                    fmt = dict(body_formatting) if body_formatting else {}
                    if alternating_row_color and (r_idx % 2 == 0) and "bg_color" not in fmt:
                        fmt["bg_color"] = alternating_row_color
                    if column_alignments and c_idx < len(column_alignments) and "alignment" not in fmt:
                        fmt["alignment"] = column_alignments[c_idx]
                    if fmt:
                        self._apply_cell_formatting(cell, fmt)

    def _apply_paragraph_formatting(self, paragraph, formatting: Dict[str, Any]) -> None:
        """Apply alignment, spacing, indents, shading, borders, font, and hex RGB color."""
        if "alignment" in formatting:
            align_map = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            }
            align_key = str(formatting["alignment"]).lower()
            if align_key in align_map:
                paragraph.alignment = align_map[align_key]

        pf = paragraph.paragraph_format
        if "space_before" in formatting:
            pf.space_before = Pt(float(formatting["space_before"]))
        if "space_after" in formatting:
            pf.space_after = Pt(float(formatting["space_after"]))
        if "line_spacing" in formatting:
            ls = float(formatting["line_spacing"])
            pf.line_spacing = Pt(ls) if ls > 5.0 else ls
        if "keep_with_next" in formatting:
            pf.keep_with_next = bool(formatting["keep_with_next"])
        if "left_indent" in formatting:
            pf.left_indent = Inches(float(formatting["left_indent"]))
        if "right_indent" in formatting:
            pf.right_indent = Inches(float(formatting["right_indent"]))
        if "first_line_indent" in formatting:
            pf.first_line_indent = Inches(float(formatting["first_line_indent"]))

        raw_bg = formatting.get("bg_color", formatting.get("shading"))
        if raw_bg:
            norm_bg = self._normalize_hex_color(raw_bg)
            if norm_bg:
                pPr = paragraph._p.get_or_add_pPr()
                for existing in pPr.findall(qn("w:shd")):
                    pPr.remove(existing)
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), norm_bg)
                pPr.append(shd)

        if "border_bottom" in formatting and formatting["border_bottom"]:
            bb = formatting["border_bottom"]
            b_style = "single"
            b_color = "1F4E79"
            b_size = "6"
            if isinstance(bb, dict):
                b_style = str(bb.get("style", "single"))
                b_color = self._normalize_hex_color(bb.get("color")) or "1F4E79"
                b_size = str(int(bb.get("size", 6)))
            pPr = paragraph._p.get_or_add_pPr()
            pBdr = pPr.find(qn("w:pBdr"))
            if pBdr is None:
                pBdr = OxmlElement("w:pBdr")
                pPr.append(pBdr)
            for existing in pBdr.findall(qn("w:bottom")):
                pBdr.remove(existing)
            bottom_el = OxmlElement("w:bottom")
            bottom_el.set(qn("w:val"), b_style)
            bottom_el.set(qn("w:sz"), b_size)
            bottom_el.set(qn("w:space"), "4")
            bottom_el.set(qn("w:color"), b_color)
            pBdr.append(bottom_el)

        rgb_color = None
        raw_color = formatting.get("color", formatting.get("font_color"))
        if isinstance(raw_color, str):
            norm_c = self._normalize_hex_color(raw_color)
            if norm_c:
                rgb_color = RGBColor(
                    int(norm_c[0:2], 16),
                    int(norm_c[2:4], 16),
                    int(norm_c[4:6], 16),
                )

        font_name = formatting.get("font_name", formatting.get("font_family"))
        for run in paragraph.runs:
            if "font_size" in formatting:
                run.font.size = Pt(float(formatting["font_size"]))
            if font_name:
                run.font.name = str(font_name)
            if "bold" in formatting:
                run.font.bold = bool(formatting["bold"])
            if "italic" in formatting:
                run.font.italic = bool(formatting["italic"])
            if "underline" in formatting:
                run.font.underline = bool(formatting["underline"])
            if rgb_color is not None:
                run.font.color.rgb = rgb_color

    async def get_document_info(self, document_id: str) -> Dict[str, Any]:
        entry = self._get_entry(document_id)
        return sanitize_for_json(entry["metadata"])

    async def list_documents(self) -> List[Dict[str, Any]]:
        return [sanitize_for_json(data["metadata"]) for data in self.active_documents.values()]
