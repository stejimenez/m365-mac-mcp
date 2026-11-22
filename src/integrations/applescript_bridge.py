"""
AppleScript Bridge for Microsoft 365 for Mac MCP (m365-mac-mcp — macOS 27.x / Office 16.x).

Provides hardened, injection-safe, timeout-protected live automation for
Microsoft Word, Microsoft Excel, and Microsoft PowerPoint on macOS.

Key design principles:
1. All dynamic inputs (paths, names, text, formulas) are passed via `osascript`
   `on run argv` arguments — never interpolated into AppleScript source strings.
2. Every subprocess call enforces a strict timeout and kills/reaps timed-out
   child processes cleanly so modal dialogs never hang the MCP server.
3. Unambiguous ASCII unit/record separators (\x1f / \x1e) are used for
   structured output parsing so commas, quotes, tabs, and newlines inside
   documents or cells never corrupt results.
4. App installation & version checks read Info.plist directly via plistlib
   (never launching Office apps just to check status), and running-state checks
   use `application "<App>" is running` (no System Events permission needed).
"""

import asyncio
import json
import os
import plistlib
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from utils.logger import setup_logger
from utils.validators import resolve_file_path, sanitize_for_json, validate_cell_ref, validate_range_ref

logger = setup_logger(__name__)

# Unambiguous ASCII delimiters for structured AppleScript return values
_US = "\x1f"  # Unit separator (fields / cells)
_RS = "\x1e"  # Record separator (rows / items)
_GS = "\x1d"  # Group separator (sections)
_FS = "\x1c"  # File separator (multi-range blocks)

_OFFICE_APPS = {
    "word": "Microsoft Word",
    "excel": "Microsoft Excel",
    "powerpoint": "Microsoft PowerPoint",
}


class AppleScriptTimeoutError(TimeoutError):
    """Raised when an AppleScript call exceeds its timeout (e.g., modal dialog)."""


class AppleScriptBridge:
    """Hardened bridge for communicating with Microsoft Office apps via AppleScript."""

    def __init__(self, default_timeout: float = 30.0):
        self.powerpoint_app = _OFFICE_APPS["powerpoint"]
        self.word_app = _OFFICE_APPS["word"]
        self.excel_app = _OFFICE_APPS["excel"]
        self.default_timeout = default_timeout
        self._lock_by_loop: Dict[int, asyncio.Lock] = {}

    def _get_loop_lock(self) -> asyncio.Lock:
        """Return an asyncio.Lock bound to the active event loop to serialize AppleEvents safely."""
        loop = asyncio.get_running_loop()
        loop_id = id(loop)
        lock = self._lock_by_loop.get(loop_id)
        if lock is None:
            lock = asyncio.Lock()
            self._lock_by_loop[loop_id] = lock
        return lock

    # ------------------------------------------------------------------
    # Core safe osascript execution
    # ------------------------------------------------------------------

    async def execute_applescript(
        self,
        script: str,
        args: Optional[Sequence[str]] = None,
        timeout: Optional[float] = None,
    ) -> str:
        """Execute an AppleScript safely with argv parameter passing and timeout.

        Args:
            script: AppleScript source (passed via stdin using `osascript -`).
            args: Positional string arguments delivered to `on run argv`.
            timeout: Timeout in seconds (defaults to self.default_timeout).

        Returns:
            Stripped stdout string from osascript.
        """
        effective_timeout = timeout if timeout is not None else self.default_timeout
        cmd = ["osascript", "-"]
        if args:
            cmd.extend(str(a) for a in args)

        async with self._get_loop_lock():
            process = await asyncio.create_subprocess_exec(
                *cmd,
                stdin=asyncio.subprocess.PIPE,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
            )

            try:
                stdout, stderr = await asyncio.wait_for(
                    process.communicate(input=script.encode("utf-8")),
                    timeout=effective_timeout,
                )
            except asyncio.TimeoutError:
                try:
                    process.kill()
                except ProcessLookupError:
                    pass
                try:
                    await asyncio.wait_for(process.wait(), timeout=5.0)
                except Exception:
                    pass
                raise AppleScriptTimeoutError(
                    f"AppleScript timed out after {effective_timeout}s. "
                    "Check if Microsoft Word/Excel is showing a modal dialog or "
                    "waiting for macOS Automation permission."
                )
            except asyncio.CancelledError:
                try:
                    process.kill()
                except ProcessLookupError:
                    pass
                raise

        if process.returncode != 0:
            error_msg = stderr.decode("utf-8", errors="replace").strip()
            if "-1743" in error_msg or "Not authorized to send Apple events" in error_msg:
                raise PermissionError(
                    f"macOS Automation permission denied (-1743): {error_msg}. "
                    "Enable access in System Settings > Privacy & Security > Automation."
                )
            raise RuntimeError(f"AppleScript execution failed: {error_msg}")

        return stdout.decode("utf-8", errors="replace").strip()

    # ------------------------------------------------------------------
    # App installation, running state & version inspection (non-intrusive)
    # ------------------------------------------------------------------

    def get_app_bundle_info(self, app_name: str) -> Dict[str, Any]:
        """Inspect an Office app's Info.plist without launching it."""
        candidates = [
            Path("/Applications") / f"{app_name}.app",
            Path.home() / "Applications" / f"{app_name}.app",
        ]
        for app_path in candidates:
            plist_path = app_path / "Contents" / "Info.plist"
            if plist_path.exists():
                try:
                    with open(plist_path, "rb") as f:
                        plist = plistlib.load(f)
                    return {
                        "installed": True,
                        "path": str(app_path),
                        "version": str(plist.get("CFBundleShortVersionString", "unknown")),
                        "build": str(plist.get("CFBundleVersion", "unknown")),
                        "bundle_id": str(plist.get("CFBundleIdentifier", "")),
                    }
                except Exception as exc:
                    return {
                        "installed": True,
                        "path": str(app_path),
                        "version": "unknown",
                        "build": "unknown",
                        "error": str(exc),
                    }
        return {"installed": False, "path": None, "version": None, "build": None}

    async def is_app_running(self, app_name: str) -> bool:
        """Check if an app is currently running without launching it or needing System Events."""
        script = """
        on run argv
            set appName to item 1 of argv
            return (application appName is running) as text
        end run
        """
        try:
            res = await self.execute_applescript(script, args=[app_name], timeout=5.0)
            return res.lower() == "true"
        except Exception as exc:
            logger.warning(f"Could not check running state for {app_name}: {exc}")
            return False

    async def check_word_status(self) -> Dict[str, Any]:
        """Check Microsoft Word installation, version, and running status."""
        info = self.get_app_bundle_info(self.word_app)
        info["running"] = await self.is_app_running(self.word_app) if info["installed"] else False
        return info

    async def check_excel_status(self) -> Dict[str, Any]:
        """Check Microsoft Excel installation, version, and running status."""
        info = self.get_app_bundle_info(self.excel_app)
        info["running"] = await self.is_app_running(self.excel_app) if info["installed"] else False
        return info

    async def check_powerpoint_status(self) -> Dict[str, Any]:
        """Check Microsoft PowerPoint installation, version, and running status."""
        info = self.get_app_bundle_info(self.powerpoint_app)
        info["running"] = await self.is_app_running(self.powerpoint_app) if info["installed"] else False
        return info

    async def get_office_version_info(self) -> Dict[str, Any]:
        """Return non-intrusive version & status info for Word, Excel, and PowerPoint."""
        word_info, excel_info, ppt_info = await asyncio.gather(
            self.check_word_status(),
            self.check_excel_status(),
            self.check_powerpoint_status(),
        )
        return {
            "word": word_info,
            "excel": excel_info,
            "powerpoint": ppt_info,
        }

    # ------------------------------------------------------------------
    # Microsoft Word (16.115.1) Live App Automation
    # ------------------------------------------------------------------

    async def launch_word(self, activate: bool = False) -> bool:
        """Launch Microsoft Word without stealing focus unless requested."""
        script = (
            'tell application "Microsoft Word" to activate'
            if activate
            else 'tell application "Microsoft Word" to launch'
        )
        await self.execute_applescript(script, timeout=15.0)
        return True

    async def list_open_word_documents(self) -> List[Dict[str, Any]]:
        """List all currently open documents in Microsoft Word.

        Avoids `repeat with d in documents` (which triggers `every document
        doesn't understand the "count" message` on Word 16.115.1 / macOS 27.2)
        and cascades cleanly across:
          1. `count of windows` -> `document of window i`
          2. `count of documents` -> `document i`
          3. `name of every document` (iterated outside Word's tell block)
          4. `active document` direct fallback
        Returns an empty list when Word is not running or has no open documents.
        """
        if not await self.is_app_running(self.word_app):
            return []

        script = """
        on run argv
            set us to character id 31
            set rs to character id 30
            set outRows to {}
            set hasRows to false
            set activeName to ""
            set activeFull to ""
            set fallbackNames to {}

            tell application "Microsoft Word"
                try
                    set rawActName to name of active document
                    if rawActName is not missing value then set activeName to rawActName as text
                end try
                try
                    set rawActFull to full name of active document
                    if rawActFull is not missing value then set activeFull to rawActFull as text
                end try

                -- Strategy 1: Enumerate via document of each open window (most reliable on Word 16.115.1)
                set winCount to 0
                try
                    set winCount to count of windows
                on error
                    set winCount to 0
                end try
                if winCount > 0 then
                    repeat with i from 1 to winCount
                        try
                            set d to document of window i
                            set dName to ""
                            set dFull to ""
                            set dPath to ""
                            set dSaved to "false"
                            set pCount to 0
                            try
                                set rawName to name of d
                                if rawName is not missing value then set dName to rawName as text
                            end try
                            try
                                set rawFull to full name of d
                                if rawFull is not missing value then set dFull to rawFull as text
                            end try
                            try
                                set rawPath to path of d
                                if rawPath is not missing value then set dPath to rawPath as text
                            end try
                            try
                                set rawSaved to saved of d
                                if rawSaved is not missing value then set dSaved to rawSaved as text
                            end try
                            try
                                set rawCount to count of paragraphs of d
                                if rawCount is not missing value then set pCount to rawCount
                            end try
                            if dName is not "" or dFull is not "" then
                                set isAct to "false"
                                if (activeFull is not "" and dFull is activeFull) or (activeName is not "" and dName is activeName) then
                                    set isAct to "true"
                                end if
                                set end of outRows to (dName & us & dFull & us & dPath & us & dSaved & us & (pCount as text) & us & isAct)
                                set hasRows to true
                            end if
                        end try
                    end repeat
                end if

                -- Strategy 2: Indexed `count of documents` + `document i`
                if not hasRows then
                    set docCount to 0
                    try
                        set docCount to count of documents
                    on error
                        set docCount to 0
                    end try
                    if docCount > 0 then
                        repeat with i from 1 to docCount
                            try
                                set d to document i
                                set dName to ""
                                set dFull to ""
                                set dPath to ""
                                set dSaved to "false"
                                set pCount to 0
                                try
                                    set rawName to name of d
                                    if rawName is not missing value then set dName to rawName as text
                                end try
                                try
                                    set rawFull to full name of d
                                    if rawFull is not missing value then set dFull to rawFull as text
                                end try
                                try
                                    set rawPath to path of d
                                    if rawPath is not missing value then set dPath to rawPath as text
                                end try
                                try
                                    set rawSaved to saved of d
                                    if rawSaved is not missing value then set dSaved to rawSaved as text
                                end try
                                try
                                    set rawCount to count of paragraphs of d
                                    if rawCount is not missing value then set pCount to rawCount
                                end try
                                if dName is not "" or dFull is not "" then
                                    set isAct to "false"
                                    if (activeFull is not "" and dFull is activeFull) or (activeName is not "" and dName is activeName) then
                                        set isAct to "true"
                                    end if
                                    set end of outRows to (dName & us & dFull & us & dPath & us & dSaved & us & (pCount as text) & us & isAct)
                                    set hasRows to true
                                end if
                            end try
                        end repeat
                    end if
                end if

                -- Strategy 3: Capture `name of every document` for iteration outside Word's tell block
                if not hasRows then
                    try
                        set rawNames to name of every document
                        if rawNames is not missing value then
                            if class of rawNames is list then
                                set fallbackNames to rawNames
                            else
                                set fallbackNames to {rawNames}
                            end if
                        end if
                    end try
                end if
            end tell

            if (not hasRows) and ((count of fallbackNames) > 0) then
                repeat with i from 1 to (count of fallbackNames)
                    set dName to (item i of fallbackNames) as text
                    set dFull to ""
                    set dPath to ""
                    set dSaved to "false"
                    set pCount to 0
                    tell application "Microsoft Word"
                        try
                            set d to document dName
                            try
                                set rawFull to full name of d
                                if rawFull is not missing value then set dFull to rawFull as text
                            end try
                            try
                                set rawPath to path of d
                                if rawPath is not missing value then set dPath to rawPath as text
                            end try
                            try
                                set rawSaved to saved of d
                                if rawSaved is not missing value then set dSaved to rawSaved as text
                            end try
                            try
                                set rawCount to count of paragraphs of d
                                if rawCount is not missing value then set pCount to rawCount
                            end try
                        end try
                    end tell
                    if dName is not "" or dFull is not "" then
                        set isAct to "false"
                        if (activeFull is not "" and dFull is activeFull) or (activeName is not "" and dName is activeName) then
                            set isAct to "true"
                        end if
                        set end of outRows to (dName & us & dFull & us & dPath & us & dSaved & us & (pCount as text) & us & isAct)
                        set hasRows to true
                    end if
                end repeat
            end if

            -- Strategy 4: Direct `active document` backstop
            if not hasRows then
                tell application "Microsoft Word"
                    try
                        set d to active document
                        set dName to ""
                        set dFull to ""
                        set dPath to ""
                        set dSaved to "false"
                        set pCount to 0
                        try
                            set rawName to name of d
                            if rawName is not missing value then set dName to rawName as text
                        end try
                        try
                            set rawFull to full name of d
                            if rawFull is not missing value then set dFull to rawFull as text
                        end try
                        try
                            set rawPath to path of d
                            if rawPath is not missing value then set dPath to rawPath as text
                        end try
                        try
                            set rawSaved to saved of d
                            if rawSaved is not missing value then set dSaved to rawSaved as text
                        end try
                        try
                            set rawCount to count of paragraphs of d
                            if rawCount is not missing value then set pCount to rawCount
                        end try
                        if dName is not "" or dFull is not "" then
                            set end of outRows to (dName & us & dFull & us & dPath & us & dSaved & us & (pCount as text) & us & "true")
                        end if
                    end try
                end tell
            end if

            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set joinedOut to outRows as text
            set AppleScript's text item delimiters to prevDelims
            return joinedOut
        end run
        """
        raw = await self.execute_applescript(script, timeout=15.0)
        if not raw:
            return []
        docs: List[Dict[str, Any]] = []
        seen: set[tuple[str, str]] = set()
        for row in raw.split(_RS):
            if not row:
                continue
            parts = row.split(_US)
            if len(parts) >= 6:
                name, full_name, path_val, saved_s, pcount_s, active_s = parts[:6]
                key = (name, full_name)
                if key in seen:
                    continue
                seen.add(key)
                docs.append({
                    "name": name,
                    "full_name": full_name,
                    "path": path_val or full_name,
                    "saved": saved_s.lower() == "true",
                    "paragraph_count": int(pcount_s) if pcount_s.isdigit() else 0,
                    "is_active": active_s.lower() == "true",
                })
            elif len(parts) >= 4:
                name, full_name, saved_s, pcount_s = parts[:4]
                key = (name, full_name)
                if key in seen:
                    continue
                seen.add(key)
                docs.append({
                    "name": name,
                    "full_name": full_name,
                    "path": full_name,
                    "saved": saved_s.lower() == "true",
                    "paragraph_count": int(pcount_s) if pcount_s.isdigit() else 0,
                    "is_active": False,
                })
        return sanitize_for_json(docs)

    async def open_word_file(
        self,
        file_path: str,
        read_only: bool = False,
        activate: bool = False,
    ) -> Dict[str, Any]:
        """Open an existing Word document (.docx/.doc/.rtf) in live Microsoft Word."""
        resolved = resolve_file_path(file_path, must_exist=True)
        script = """
        on run argv
            set posixPath to item 1 of argv
            set isReadOnly to (item 2 of argv) is "true"
            set shouldActivate to (item 3 of argv) is "true"
            set targetFile to POSIX file posixPath
            tell application "Microsoft Word"
                if shouldActivate then activate
                open file name (targetFile as text) read only isReadOnly
                set docName to name of active document
                return docName
            end tell
        end run
        """
        doc_name = await self.execute_applescript(
            script,
            args=[str(resolved), "true" if read_only else "false", "true" if activate else "false"],
            timeout=30.0,
        )
        return {
            "status": "success",
            "document_name": doc_name,
            "file_path": str(resolved),
            "read_only": read_only,
            "method": "applescript",
        }

    async def create_word_document(
        self,
        title: Optional[str] = None,
        activate: bool = False,
    ) -> Dict[str, Any]:
        """Create a new blank document in live Microsoft Word."""
        script = """
        on run argv
            set initialTitle to item 1 of argv
            set shouldActivate to (item 2 of argv) is "true"
            tell application "Microsoft Word"
                if shouldActivate then activate
                set newDoc to make new document
                if initialTitle is not "" then
                    insert text (initialTitle & return) at beginning of text object of newDoc
                end if
                return name of newDoc
            end tell
        end run
        """
        doc_name = await self.execute_applescript(
            script,
            args=[title or "", "true" if activate else "false"],
            timeout=20.0,
        )
        return {
            "status": "success",
            "document_name": doc_name,
            "title": title or "",
            "method": "applescript",
        }

    async def read_live_word_document(
        self,
        document_name: Optional[str] = None,
        include_tables: bool = True,
    ) -> Dict[str, Any]:
        """Read text, paragraphs, selection state, and optionally tables from a live Word doc."""
        script = """
        on run argv
            set targetName to item 1 of argv
            set incTables to (item 2 of argv) is "true"
            set us to character id 31
            set rs to character id 30
            set gs to character id 29
            tell application "Microsoft Word"
                if (count of documents) is 0 then
                    error "No open Word documents"
                end if
                if targetName is "" then
                    set targetDoc to active document
                else
                    set targetDoc to document targetName
                end if
                set docName to name of targetDoc
                set docFull to ""
                try
                    set docFull to full name of targetDoc
                end try
                set docSaved to (saved of targetDoc) as text
                set fullText to content of text object of targetDoc

                set selText to ""
                set selStart to -1
                set selEnd to -1
                try
                    set selText to content of text object of selection
                    set selStart to selection start of selection
                    set selEnd to selection end of selection
                end try

                set tableBlocks to {}
                if incTables then
                    set tCount to count of tables of targetDoc
                    repeat with ti from 1 to tCount
                        set tbl to table ti of targetDoc
                        set rCount to count of rows of tbl
                        set cCount to count of columns of tbl
                        set rowStrings to {}
                        repeat with ri from 1 to rCount
                            set cellStrings to {}
                            repeat with ci from 1 to cCount
                                set cText to ""
                                try
                                    set cText to content of text object of cell ci of row ri of tbl
                                end try
                                set end of cellStrings to cText
                            end repeat
                            set AppleScript's text item delimiters to us
                            set end of rowStrings to (cellStrings as text)
                        end repeat
                        set AppleScript's text item delimiters to rs
                        set end of tableBlocks to (rowStrings as text)
                    end repeat
                end if

                set AppleScript's text item delimiters to gs
                set tablesJoined to tableBlocks as text
                set metaJoined to docName & us & docFull & us & docSaved & us & (selStart as text) & us & (selEnd as text) & us & selText
                return metaJoined & gs & fullText & gs & tablesJoined
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[document_name or "", "true" if include_tables else "false"],
            timeout=30.0,
        )
        sections = raw.split(_GS)
        meta_parts = sections[0].split(_US) if sections else []
        full_text = sections[1] if len(sections) > 1 else ""
        # Word uses \r for paragraph breaks; normalize to \n
        normalized_text = full_text.replace("\r\n", "\n").replace("\r", "\n")
        paragraphs = [p for p in normalized_text.split("\n")]
        if paragraphs and paragraphs[-1] == "":
            paragraphs.pop()

        tables: List[List[List[str]]] = []
        if include_tables and len(sections) > 2:
            for tbl_raw in sections[2:]:
                if not tbl_raw:
                    continue
                rows_parsed: List[List[str]] = []
                for row_raw in tbl_raw.split(_RS):
                    # Strip Word cell-end marker (\r\x07)
                    cells = [
                        c.replace("\r\x07", "").replace("\x07", "").strip()
                        for c in row_raw.split(_US)
                    ]
                    rows_parsed.append(cells)
                tables.append(rows_parsed)

        return sanitize_for_json({
            "status": "success",
            "document_name": meta_parts[0] if len(meta_parts) > 0 else "",
            "full_name": meta_parts[1] if len(meta_parts) > 1 else "",
            "saved": (meta_parts[2].lower() == "true") if len(meta_parts) > 2 else False,
            "selection": {
                "start": int(meta_parts[3]) if len(meta_parts) > 3 and meta_parts[3].lstrip("-").isdigit() else -1,
                "end": int(meta_parts[4]) if len(meta_parts) > 4 and meta_parts[4].lstrip("-").isdigit() else -1,
                "text": (meta_parts[5] if len(meta_parts) > 5 else "").replace("\r", "\n"),
            },
            "paragraph_count": len(paragraphs),
            "paragraphs": paragraphs,
            "text": normalized_text,
            "table_count": len(tables),
            "tables": tables,
            "method": "applescript",
        })

    async def write_live_word_text(
        self,
        text: str,
        document_name: Optional[str] = None,
        mode: str = "append",
        add_newline: bool = True,
    ) -> Dict[str, Any]:
        """Write, append, insert at cursor/selection, or replace all text in a live Word doc.

        Args:
            text: Text to write.
            document_name: Exact open document name (defaults to active document).
            mode: 'append' | 'insert_at_selection' | 'prepend' | 'replace_all'.
            add_newline: Append a trailing paragraph break when mode is append/prepend.
        """
        valid_modes = {"append", "insert_at_selection", "prepend", "replace_all"}
        if mode not in valid_modes:
            raise ValueError(f"mode must be one of {sorted(valid_modes)}")

        script = """
        on run argv
            set targetName to item 1 of argv
            set bodyText to item 2 of argv
            set writeMode to item 3 of argv
            set addBreak to (item 4 of argv) is "true"
            tell application "Microsoft Word"
                if (count of documents) is 0 then
                    error "No open Word documents"
                end if
                if targetName is "" then
                    set targetDoc to active document
                else
                    set targetDoc to document targetName
                end if
                if writeMode is "replace_all" then
                    set content of text object of targetDoc to bodyText
                else if writeMode is "insert_at_selection" then
                    if addBreak then
                        insert text (bodyText & return) at selection
                    else
                        insert text bodyText at selection
                    end if
                else if writeMode is "prepend" then
                    if addBreak then
                        insert text (bodyText & return) at beginning of text object of targetDoc
                    else
                        insert text bodyText at beginning of text object of targetDoc
                    end if
                else
                    if addBreak then
                        insert text (bodyText & return) at end of text object of targetDoc
                    else
                        insert text bodyText at end of text object of targetDoc
                    end if
                end if
                return name of targetDoc
            end tell
        end run
        """
        doc_name = await self.execute_applescript(
            script,
            args=[document_name or "", text, mode, "true" if add_newline else "false"],
            timeout=20.0,
        )
        return {
            "status": "success",
            "document_name": doc_name,
            "mode": mode,
            "characters_written": len(text),
            "method": "applescript",
        }

    async def replace_live_word_text(
        self,
        find_text: str,
        replace_text: str,
        document_name: Optional[str] = None,
        match_case: bool = False,
    ) -> Dict[str, Any]:
        """Find and replace all occurrences in a live Word document using native Find."""
        if not find_text:
            raise ValueError("find_text cannot be empty")
        script = """
        on run argv
            set targetName to item 1 of argv
            set fText to item 2 of argv
            set rText to item 3 of argv
            set mCase to (item 4 of argv) is "true"
            tell application "Microsoft Word"
                if (count of documents) is 0 then
                    error "No open Word documents"
                end if
                if targetName is "" then
                    set targetDoc to active document
                else
                    set targetDoc to document targetName
                end if
                set findObj to find object of text object of targetDoc
                clear formatting findObj
                set foundAny to execute find findObj find text fText replace with rText replace replace all match case mCase
                return (name of targetDoc) & (character id 31) & (foundAny as text)
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[document_name or "", find_text, replace_text, "true" if match_case else "false"],
            timeout=20.0,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "document_name": parts[0] if parts else (document_name or ""),
            "replaced": (parts[1].lower() == "true") if len(parts) > 1 else False,
            "find_text": find_text,
            "replace_text": replace_text,
            "method": "applescript",
        }

    async def save_word_document(
        self,
        document_name: Optional[str] = None,
        file_path: Optional[str] = None,
        file_format: str = "docx",
    ) -> Dict[str, Any]:
        """Save a live Word document in-place or Save As (.docx, .pdf, .rtf, .txt, .doc)."""
        resolved_path = ""
        if file_path:
            p = resolve_file_path(file_path, must_exist=False)
            p.parent.mkdir(parents=True, exist_ok=True)
            resolved_path = str(p)

        fmt_key = file_format.lower().lstrip(".")
        valid_formats = {"docx", "pdf", "rtf", "txt", "doc"}
        if fmt_key not in valid_formats:
            raise ValueError(f"Unsupported Word save format '{file_format}'. Use one of {sorted(valid_formats)}.")

        script = """
        on run argv
            set targetName to item 1 of argv
            set posixOut to item 2 of argv
            set fmtKey to item 3 of argv
            set hfsPath to ""
            if posixOut is not "" then
                set hfsPath to (POSIX file posixOut) as text
            end if
            tell application "Microsoft Word"
                if targetName is "" then
                    try
                        set targetDoc to active document
                    on error
                        error "No open Word documents"
                    end try
                else
                    set targetDoc to document targetName
                end if
                if hfsPath is "" then
                    save targetDoc
                    return name of targetDoc
                else
                    if fmtKey is "pdf" then
                        save as targetDoc file name hfsPath file format format PDF
                    else if fmtKey is "rtf" then
                        save as targetDoc file name hfsPath file format format rtf
                    else if fmtKey is "txt" then
                        save as targetDoc file name hfsPath file format format text
                    else if fmtKey is "doc" then
                        save as targetDoc file name hfsPath file format format document97
                    else
                        save as targetDoc file name hfsPath file format format document
                    end if
                    return name of targetDoc
                end if
            end tell
        end run
        """
        saved_name = await self.execute_applescript(
            script,
            args=[document_name or "", resolved_path, fmt_key],
            timeout=30.0,
        )
        return {
            "status": "success",
            "document_name": saved_name,
            "file_path": resolved_path or "(saved in place)",
            "format": fmt_key,
            "method": "applescript",
        }

    async def close_word_document(
        self,
        document_name: Optional[str] = None,
        save_changes: str = "no",
    ) -> Dict[str, Any]:
        """Close a live Word document cleanly without triggering an unsaved-changes dialog."""
        if save_changes not in ("no", "yes"):
            raise ValueError("save_changes must be 'no' or 'yes'")

        script = """
        on run argv
            set targetName to item 1 of argv
            set saveOpt to item 2 of argv
            tell application "Microsoft Word"
                if (count of documents) is 0 then
                    return ""
                end if
                if targetName is "" then
                    set targetDoc to active document
                else
                    set targetDoc to document targetName
                end if
                set closedName to name of targetDoc
                if saveOpt is "yes" then
                    close targetDoc saving yes
                else
                    close targetDoc saving no
                end if
                return closedName
            end tell
        end run
        """
        closed_name = await self.execute_applescript(
            script,
            args=[document_name or "", save_changes],
            timeout=20.0,
        )
        return {
            "status": "success",
            "closed_document": closed_name,
            "save_changes": save_changes,
            "method": "applescript",
        }

    # ------------------------------------------------------------------
    # Microsoft Excel (16.115) Live App Automation
    # ------------------------------------------------------------------

    async def launch_excel(self, activate: bool = False) -> bool:
        """Launch Microsoft Excel without stealing focus unless requested."""
        script = (
            'tell application "Microsoft Excel" to activate'
            if activate
            else 'tell application "Microsoft Excel" to launch'
        )
        await self.execute_applescript(script, timeout=15.0)
        return True

    async def list_open_excel_workbooks(self) -> List[Dict[str, Any]]:
        """List all currently open workbooks in Microsoft Excel with their sheets.

        Uses indexed `count of workbooks` / `workbook i` and `count of worksheets` /
        `worksheet j` (with `active workbook` fallback) to avoid `every workbook
        doesn't understand the "count" message` on Excel 16.115 / macOS 27.2.
        """
        if not await self.is_app_running(self.excel_app):
            return []

        script = """
        on run argv
            set us to character id 31
            set rs to character id 30
            set gs to character id 29
            set outRows to {}
            set hasRows to false
            tell application "Microsoft Excel"
                set wbCount to 0
                try
                    set wbCount to count of workbooks
                on error
                    set wbCount to 0
                end try
                if wbCount > 0 then
                    repeat with i from 1 to wbCount
                        try
                            set wb to workbook i
                            set wbName to ""
                            set wbFull to ""
                            set wbSaved to "false"
                            set activeSh to ""
                            set shNames to {}
                            try
                                set wbName to (name of wb) as text
                            end try
                            try
                                set wbFull to (full name of wb) as text
                            end try
                            try
                                set wbSaved to (saved of wb) as text
                            end try
                            try
                                set activeSh to (name of active sheet of wb) as text
                            end try
                            try
                                set shCount to count of worksheets of wb
                                repeat with j from 1 to shCount
                                    try
                                        set end of shNames to (name of worksheet j of wb) as text
                                    end try
                                end repeat
                            end try
                            set prevDelims to AppleScript's text item delimiters
                            set AppleScript's text item delimiters to gs
                            set shJoined to shNames as text
                            set AppleScript's text item delimiters to prevDelims
                            if wbName is not "" or wbFull is not "" then
                                set end of outRows to (wbName & us & wbFull & us & wbSaved & us & activeSh & us & shJoined)
                                set hasRows to true
                            end if
                        end try
                    end repeat
                end if

                if not hasRows then
                    try
                        set wb to active workbook
                        set wbName to ""
                        set wbFull to ""
                        set wbSaved to "false"
                        set activeSh to ""
                        try
                            set wbName to (name of wb) as text
                        end try
                        try
                            set wbFull to (full name of wb) as text
                        end try
                        try
                            set wbSaved to (saved of wb) as text
                        end try
                        try
                            set activeSh to (name of active sheet of wb) as text
                        end try
                        if wbName is not "" or wbFull is not "" then
                            set end of outRows to (wbName & us & wbFull & us & wbSaved & us & activeSh & us & activeSh)
                        end if
                    end try
                end if
            end tell
            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set joinedOut to outRows as text
            set AppleScript's text item delimiters to prevDelims
            return joinedOut
        end run
        """
        raw = await self.execute_applescript(script, timeout=15.0)
        if not raw:
            return []
        workbooks = []
        for row in raw.split(_RS):
            parts = row.split(_US)
            if len(parts) >= 5:
                sheets = [s for s in parts[4].split(_GS) if s]
                workbooks.append({
                    "name": parts[0],
                    "full_name": parts[1],
                    "saved": parts[2].lower() == "true",
                    "active_sheet": parts[3],
                    "worksheets": sheets,
                })
        return workbooks

    async def open_excel_file(
        self,
        file_path: str,
        read_only: bool = False,
        activate: bool = False,
    ) -> Dict[str, Any]:
        """Open an existing Excel file (.xlsx/.xlsm/.csv) in live Microsoft Excel."""
        resolved = resolve_file_path(file_path, must_exist=True)
        script = """
        on run argv
            set posixPath to item 1 of argv
            set isReadOnly to (item 2 of argv) is "true"
            set shouldActivate to (item 3 of argv) is "true"
            set targetFile to POSIX file posixPath
            tell application "Microsoft Excel"
                set prevAlerts to display alerts
                set display alerts to false
                try
                    if shouldActivate then activate
                    set wb to open workbook workbook file name (targetFile as text) read only isReadOnly
                    try
                        calculate
                    end try
                    set wbName to name of wb
                    set activeSh to name of active sheet of wb
                    set display alerts to prevAlerts
                    return wbName & (character id 31) & activeSh
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[str(resolved), "true" if read_only else "false", "true" if activate else "false"],
            timeout=30.0,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "active_sheet": parts[1] if len(parts) > 1 else "",
            "file_path": str(resolved),
            "read_only": read_only,
            "method": "applescript",
        }

    async def create_excel_workbook(self, activate: bool = False) -> Dict[str, Any]:
        """Create a new blank workbook in live Microsoft Excel."""
        script = """
        on run argv
            set shouldActivate to (item 1 of argv) is "true"
            tell application "Microsoft Excel"
                if shouldActivate then activate
                set wb to make new workbook
                return (name of wb) & (character id 31) & (name of active sheet of wb)
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=["true" if activate else "false"],
            timeout=20.0,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "active_sheet": parts[1] if len(parts) > 1 else "",
            "method": "applescript",
        }

    async def list_live_excel_worksheets(
        self,
        workbook_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """List all worksheets in the specified (or active) live Excel workbook."""
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set us to character id 31
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set shNames to {}
                repeat with sh in worksheets of wb
                    set end of shNames to name of sh
                end repeat
                set AppleScript's text item delimiters to us
                return (name of wb) & (character id 30) & (name of active sheet of wb) & (character id 30) & (shNames as text)
            end tell
        end run
        """
        raw = await self.execute_applescript(script, args=[workbook_name or ""], timeout=15.0)
        parts = raw.split(_RS)
        sheets = [s for s in parts[2].split(_US) if s] if len(parts) > 2 else []
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "active_sheet": parts[1] if len(parts) > 1 else "",
            "worksheets": sheets,
            "method": "applescript",
        }

    async def add_live_excel_worksheet(
        self,
        sheet_name: str,
        workbook_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Create a new worksheet in the targeted live Excel workbook."""
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set newName to item 2 of argv
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set newSh to make new worksheet at end of worksheets of wb with properties {name:newName}
                return (name of wb) & (character id 31) & (name of newSh)
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", sheet_name],
            timeout=15.0,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "sheet_name": parts[1] if len(parts) > 1 else sheet_name,
            "method": "applescript",
        }

    async def rename_live_excel_worksheet(
        self,
        old_name: str,
        new_name: str,
        workbook_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Rename a worksheet in the targeted live Excel workbook."""
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set oldSh to item 2 of argv
            set newSh to item 3 of argv
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set name of worksheet oldSh of wb to newSh
                return (name of wb) & (character id 31) & newSh
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", old_name, new_name],
            timeout=15.0,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "old_name": old_name,
            "new_name": parts[1] if len(parts) > 1 else new_name,
            "method": "applescript",
        }

    async def activate_live_excel_worksheet(
        self,
        sheet_name: str,
        workbook_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Switch the active worksheet in the targeted live Excel workbook."""
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                activate object worksheet shTarget of wb
                return (name of wb) & (character id 31) & (name of active sheet of wb)
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", sheet_name],
            timeout=15.0,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "active_sheet": parts[1] if len(parts) > 1 else sheet_name,
            "method": "applescript",
        }

    async def read_live_excel_range(
        self,
        range_ref: str,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        include_formulas: bool = False,
    ) -> Dict[str, Any]:
        """Read cell or range values (and optionally formulas) from a specific live Excel workbook & sheet."""
        clean_ref = validate_range_ref(range_ref)
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set refStr to item 3 of argv
            set incFormulas to (item 4 of argv) is "true"
            set us to character id 31
            set rs to character id 30
            set gs to character id 29

            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                if shTarget is "" then
                    set sh to active sheet of wb
                else
                    set sh to worksheet shTarget of wb
                end if
                set wbName to name of wb
                set shName to name of sh
                set rng to range refStr of sh
                set rawVals to value of rng
                set rawFmls to missing value
                if incFormulas then
                    set rawFmls to formula of rng
                end if
            end tell

            -- Normalize scalar / 1D list / 2D list outside Excel's tell block (fast & avoids sub-range indexing quirks)
            set valBlock to my matrix_to_text(rawVals, us, rs)
            set fmlBlock to ""
            if incFormulas then
                set fmlBlock to my matrix_to_text(rawFmls, us, rs)
            end if
            return wbName & us & shName & gs & valBlock & gs & fmlBlock
        end run

        on matrix_to_text(rawData, us, rs)
            if rawData is missing value then
                return ""
            end if
            if class of rawData is not list then
                return (rawData as text)
            end if
            if (count of rawData) is 0 then
                return ""
            end if
            if class of (item 1 of rawData) is not list then
                set matrixRows to {rawData}
            else
                set matrixRows to rawData
            end if
            set outRows to {}
            repeat with rowItem in matrixRows
                set rowList to contents of rowItem
                if class of rowList is not list then
                    set rowList to {rowList}
                end if
                set cellTexts to {}
                repeat with cellItem in rowList
                    set v to contents of cellItem
                    if v is missing value then
                        set end of cellTexts to ""
                    else
                        set end of cellTexts to (v as text)
                    end if
                end repeat
                set prevDelims to AppleScript's text item delimiters
                set AppleScript's text item delimiters to us
                set end of outRows to (cellTexts as text)
                set AppleScript's text item delimiters to prevDelims
            end repeat
            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set joined to outRows as text
            set AppleScript's text item delimiters to prevDelims
            return joined
        end matrix_to_text
        """
        raw = await self.execute_applescript(
            script,
            args=[
                workbook_name or "",
                sheet_name or "",
                clean_ref,
                "true" if include_formulas else "false",
            ],
            timeout=30.0,
        )
        sections = raw.split(_GS)
        header = sections[0].split(_US) if sections else ["", ""]
        val_block = sections[1] if len(sections) > 1 else ""
        fml_block = sections[2] if len(sections) > 2 else ""

        values = [row.split(_US) for row in val_block.split(_RS)] if val_block else []
        result: Dict[str, Any] = {
            "status": "success",
            "workbook_name": header[0],
            "sheet_name": header[1] if len(header) > 1 else "",
            "range": clean_ref,
            "rows": len(values),
            "columns": len(values[0]) if values else 0,
            "values": values,
        }
        if include_formulas:
            result["formulas"] = self._extract_sparse_formulas(fml_block, clean_ref)
        return result

    @staticmethod
    def _extract_sparse_formulas(fml_block: str, range_ref: str) -> Dict[str, str]:
        """Extract only formula cells ('=...') from a serialized formula matrix into a sparse A1->formula map."""
        if not fml_block:
            return {}
        start_part = (range_ref or "A1").split(":")[0].replace("$", "").strip().upper()
        m = re.match(r"^([A-Z]+)([0-9]+)$", start_part)
        if m:
            col_letters, row_str = m.group(1), m.group(2)
            start_col = 0
            for ch in col_letters:
                start_col = start_col * 26 + (ord(ch) - ord("A") + 1)
            start_row = int(row_str)
        else:
            start_col, start_row = 1, 1

        def _idx_to_col(idx: int) -> str:
            out = ""
            while idx > 0:
                idx, rem = divmod(idx - 1, 26)
                out = chr(65 + rem) + out
            return out

        sparse: Dict[str, str] = {}
        for r_off, row_str_val in enumerate(fml_block.split(_RS)):
            for c_off, cell_val in enumerate(row_str_val.split(_US)):
                if cell_val.startswith("="):
                    coord = f"{_idx_to_col(start_col + c_off)}{start_row + r_off}"
                    sparse[coord] = cell_val
        return sparse

    @staticmethod
    def _sanitize_live_excel_formula(formula: str, sheet_names: Sequence[str]) -> str:
        """Auto-quote same-workbook sheet names with spaces and block non-existent sheet refs.

        Prevents live Microsoft Excel from popping up a modal 'Cannot find <Sheet>. Copy from:'
        file picker when a formula has an unquoted sheet name with spaces (e.g. `=Cash Flow!F15`)
        or references a sheet that does not exist in the open workbook.
        """
        if not formula or not formula.startswith("=") or "!" not in formula:
            return formula

        existing = list(sheet_names)
        existing_set = set(existing)
        sheets_needing_quotes = sorted(
            [
                name
                for name in existing
                if not re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", name)
                or re.match(r"^[A-Za-z]{1,3}[0-9]+$", name)
            ],
            key=len,
            reverse=True,
        )

        # Auto-quote unquoted same-workbook sheet names outside string literals and already-quoted refs
        if sheets_needing_quotes:
            token_re = re.compile(r'("(?:""|[^"])*"|\'(?:\'\'|[^\'])*\'!)')
            parts = token_re.split(formula)
            for idx in range(0, len(parts), 2):
                seg = parts[idx]
                if not seg or "!" not in seg:
                    continue
                for sname in sheets_needing_quotes:
                    escaped = sname.replace("'", "''")
                    pat = re.compile(r"(^|[=(,+\-*/:; &<>^])" + re.escape(sname) + r"!")
                    seg = pat.sub(rf"\1'{escaped}'!", seg)
                parts[idx] = seg
            formula = "".join(parts)

        # Guard against non-existent sheet references that would trigger Excel's modal file picker
        if existing:
            stripped = re.sub(r'"(?:""|[^"])*"', '""', formula)
            ref_re = re.compile(r"(?:'((?:''|[^'])+)'|([A-Za-z0-9_.\[\]]+))!")
            for match in ref_re.finditer(stripped):
                q, u = match.group(1), match.group(2)
                ref = q.replace("''", "'") if q is not None else u
                if not ref:
                    continue
                if "[" in ref or "]" in ref or ref not in existing_set:
                    raise ValueError(
                        f"Formula '{formula}' references unknown/external sheet '{ref}' "
                        f"(would trigger Excel's modal file picker). Available sheets: {existing}. "
                        "Use single quotes for sheet names with spaces, e.g. ='Cash Flow'!F15."
                    )

        return formula

    _warned_legacy_formula_guard: bool = False

    @classmethod
    def get_formula_guard_mode(cls, raw: Optional[str] = None) -> str:
        """Parse M365_MCP_FORMULA_GUARD (or legacy OFFICE_MCP_FORMULA_GUARD): 'literals_only' (default) | 'strict' | 'off'."""
        if raw is None:
            env_val = os.environ.get("M365_MCP_FORMULA_GUARD")
            if env_val is None or not env_val.strip():
                legacy = os.environ.get("OFFICE_MCP_FORMULA_GUARD")
                if legacy is not None and legacy.strip():
                    if not cls._warned_legacy_formula_guard:
                        logger.warning(
                            "OFFICE_MCP_FORMULA_GUARD is deprecated; rename it to M365_MCP_FORMULA_GUARD."
                        )
                        cls._warned_legacy_formula_guard = True
                    env_val = legacy
            val = (env_val or "literals_only").strip().lower()
        else:
            val = raw.strip().lower()
        if not val or val in ("literals_only", "literals", "literal", "default"):
            return "literals_only"
        if val in ("strict", "all", "on"):
            return "strict"
        if val in ("off", "false", "0", "no", "none", "disabled"):
            return "off"
        return "literals_only"

    @staticmethod
    def _formula_guard_response(mode: str, blocked: List[Dict[str, Any]]) -> Dict[str, Any]:
        return sanitize_for_json({
            "status": "confirmation_required",
            "guard_mode": mode,
            "blocked_count": len(blocked),
            "blocked_cells": blocked[:10],
            "instruction": (
                "STOP and ask the user for approval before retrying with allow_formula_overwrite=true. "
                "Briefly explain what you want to change and why."
            ),
        })

    async def write_live_excel_cell(
        self,
        cell_ref: str,
        value: Any,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        is_formula: bool = False,
        number_format: Optional[str] = None,
        recalculate: bool = True,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write a typed value or formula into a specific cell of a targeted live Excel workbook/sheet."""
        clean_cell = validate_cell_ref(cell_ref)
        guard_mode = "off" if allow_formula_overwrite else self.get_formula_guard_mode()

        # Determine explicit type tag so AppleScript preserves booleans, numbers, nulls, formulas, strings
        if value is None:
            val_type = "null"
            val_str = ""
        elif is_formula or (isinstance(value, str) and value.startswith("=")):
            val_type = "formula"
            val_str = value if str(value).startswith("=") else f"={value}"
            if "!" in val_str:
                try:
                    ws_info = await self.list_live_excel_worksheets(workbook_name=workbook_name)
                    val_str = self._sanitize_live_excel_formula(
                        val_str, ws_info.get("worksheets", [])
                    )
                except ValueError:
                    raise
                except Exception:
                    pass
        elif isinstance(value, bool):
            val_type = "bool"
            val_str = "true" if value else "false"
        elif isinstance(value, (int, float)):
            val_type = "number"
            val_str = str(value)
        else:
            val_type = "string"
            val_str = str(value)

        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set cellRef to item 3 of argv
            set valType to item 4 of argv
            set valStr to item 5 of argv
            set numFmt to item 6 of argv
            set doCalc to (item 7 of argv) is "true"
            set guardMode to item 8 of argv
            set us to character id 31
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                if shTarget is "" then
                    set sh to active sheet of wb
                else
                    set sh to worksheet shTarget of wb
                end if
                set prevAlerts to display alerts
                set display alerts to false
                try
                    set targetCell to range cellRef of sh
                    set fmlUpdated to 0
                    if guardMode is not "off" then
                        set existingFml to ""
                        try
                            set rawExisting to formula of targetCell
                            if rawExisting is not missing value then
                                set existingFml to rawExisting as text
                            end if
                        end try
                        if (length of existingFml) > 0 and (character 1 of existingFml is "=") then
                            if guardMode is "strict" or valType is not "formula" then
                                set display alerts to prevAlerts
                                return "BLOCKED" & us & (name of wb) & us & (name of sh) & us & existingFml
                            else
                                set fmlUpdated to 1
                            end if
                        end if
                    end if
                    if valType is "null" then
                        clear contents targetCell
                    else if valType is "formula" then
                        set formula of targetCell to valStr
                    else if valType is "bool" then
                        if valStr is "true" then
                            set value of targetCell to true
                        else
                            set value of targetCell to false
                        end if
                    else if valType is "number" then
                        set value of targetCell to (valStr as number)
                    else
                        set value of targetCell to valStr
                    end if
                    if numFmt is not "" then
                        set number format of targetCell to numFmt
                    end if
                    if doCalc then
                        calculate
                    end if
                    set evalVal to value of targetCell
                    if evalVal is missing value then
                        set evalText to ""
                    else
                        set evalText to evalVal as text
                    end if
                    set display alerts to prevAlerts
                    return (name of wb) & us & (name of sh) & us & evalText & us & (fmlUpdated as text)
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[
                workbook_name or "",
                sheet_name or "",
                clean_cell,
                val_type,
                val_str,
                number_format or "",
                "true" if recalculate else "false",
                guard_mode,
            ],
            timeout=20.0,
        )
        parts = raw.split(_US)
        if parts and parts[0] == "BLOCKED":
            return self._formula_guard_response(
                guard_mode,
                [{
                    "sheet": parts[2] if len(parts) > 2 else (sheet_name or ""),
                    "cell": clean_cell,
                    "existing_formula": parts[3] if len(parts) > 3 else "",
                    "proposed_value": value,
                }],
            )
        out: Dict[str, Any] = {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "sheet_name": parts[1] if len(parts) > 1 else "",
            "cell": clean_cell,
            "written_value": value,
            "evaluated_value": parts[2] if len(parts) > 2 else "",
            "number_format": number_format,
        }
        if len(parts) > 3 and parts[3] == "1":
            out["formulas_updated"] = 1
        return sanitize_for_json(out)

    async def write_live_excel_range(
        self,
        start_cell: str,
        data: List[List[Any]],
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        number_format: Optional[str] = None,
        recalculate: bool = True,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write a 2D array into a targeted live Excel workbook & worksheet."""
        if not data:
            raise ValueError("data must be a non-empty 2D list")
        clean_start = validate_cell_ref(start_cell)
        guard_mode = "off" if allow_formula_overwrite else self.get_formula_guard_mode()

        # If any cell contains a cross-sheet formula, fetch sheet names once to auto-quote & guard
        has_cross_sheet_formula = any(
            isinstance(v, str) and v.startswith("=") and "!" in v
            for row in data
            for v in row
        )
        live_sheets: List[str] = []
        if has_cross_sheet_formula:
            try:
                ws_info = await self.list_live_excel_worksheets(workbook_name=workbook_name)
                live_sheets = list(ws_info.get("worksheets", []))
            except Exception:
                live_sheets = []

        # Encode rows with type tags: <type>:<value> separated by _US / _RS
        encoded_rows: List[str] = []
        max_cols = 0
        for row in data:
            max_cols = max(max_cols, len(row))
            encoded_cells: List[str] = []
            for cell_val in row:
                if cell_val is None:
                    encoded_cells.append("N:")
                elif isinstance(cell_val, bool):
                    encoded_cells.append(f"B:{'true' if cell_val else 'false'}")
                elif isinstance(cell_val, (int, float)):
                    encoded_cells.append(f"D:{cell_val}")
                elif isinstance(cell_val, str) and cell_val.startswith("="):
                    safe_fml = (
                        self._sanitize_live_excel_formula(cell_val, live_sheets)
                        if "!" in cell_val and live_sheets
                        else cell_val
                    )
                    encoded_cells.append(f"F:{safe_fml}")
                else:
                    encoded_cells.append(f"S:{cell_val}")
            encoded_rows.append(_US.join(encoded_cells))
        payload = _RS.join(encoded_rows)

        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set startRef to item 3 of argv
            set payload to item 4 of argv
            set numFmt to item 5 of argv
            set doCalc to (item 6 of argv) is "true"
            set guardMode to item 7 of argv
            set us to character id 31
            set rs to character id 30
            set gs to character id 29

            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set rowList to text items of payload
            set AppleScript's text item delimiters to prevDelims

            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                if shTarget is "" then
                    set sh to active sheet of wb
                else
                    set sh to worksheet shTarget of wb
                end if

                set anchorCell to range startRef of sh
                set startR to first row index of anchorCell
                set startC to first column index of anchorCell

                set fmlUpdated to 0
                if guardMode is not "off" then
                    set blockedRows to {}
                    set rIdx to 0
                    repeat with rStr in rowList
                        set AppleScript's text item delimiters to us
                        set cellList to text items of (rStr as text)
                        set AppleScript's text item delimiters to prevDelims
                        set cIdx to 0
                        repeat with entry in cellList
                            set entryStr to entry as text
                            set tagChar to ""
                            set rawVal to ""
                            if (length of entryStr) >= 2 then
                                set tagChar to character 1 of entryStr
                                if (length of entryStr) > 2 then
                                    set rawVal to text 3 thru -1 of entryStr
                                end if
                            end if
                            set targetCell to cell (startC + cIdx) of row (startR + rIdx) of sh
                            set existingFml to ""
                            try
                                set rawEx to formula of targetCell
                                if rawEx is not missing value then
                                    set existingFml to rawEx as text
                                end if
                            end try
                            if (length of existingFml) > 0 and (character 1 of existingFml is "=") then
                                if guardMode is "strict" or tagChar is not "F" then
                                    set cellAddr to get address of targetCell without row absolute and column absolute
                                    set end of blockedRows to ((name of sh) & us & cellAddr & us & existingFml & us & rawVal)
                                else
                                    set fmlUpdated to fmlUpdated + 1
                                end if
                            end if
                            set cIdx to cIdx + 1
                        end repeat
                        set rIdx to rIdx + 1
                    end repeat
                    if (count of blockedRows) > 0 then
                        set AppleScript's text item delimiters to rs
                        set joinedBlocked to blockedRows as text
                        set AppleScript's text item delimiters to prevDelims
                        return "BLOCKED" & gs & joinedBlocked
                    end if
                end if

                set rIdx to 0
                repeat with rStr in rowList
                    set AppleScript's text item delimiters to us
                    set cellList to text items of (rStr as text)
                    set AppleScript's text item delimiters to prevDelims
                    set cIdx to 0
                    repeat with entry in cellList
                        set entryStr to entry as text
                        set targetCell to cell (startC + cIdx) of row (startR + rIdx) of sh
                        if (length of entryStr) >= 2 then
                            set tagChar to character 1 of entryStr
                            if (length of entryStr) > 2 then
                                set rawVal to text 3 thru -1 of entryStr
                            else
                                set rawVal to ""
                            end if
                            if tagChar is "N" then
                                clear contents targetCell
                            else if tagChar is "B" then
                                if rawVal is "true" then
                                    set value of targetCell to true
                                else
                                    set value of targetCell to false
                                end if
                            else if tagChar is "D" then
                                set value of targetCell to (rawVal as number)
                            else if tagChar is "F" then
                                set formula of targetCell to rawVal
                            else
                                set value of targetCell to rawVal
                            end if
                            if numFmt is not "" then
                                set number format of targetCell to numFmt
                            end if
                        end if
                        set cIdx to cIdx + 1
                    end repeat
                    set rIdx to rIdx + 1
                end repeat

                if doCalc then
                    calculate
                end if
                return (name of wb) & us & (name of sh) & us & (fmlUpdated as text)
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[
                workbook_name or "",
                sheet_name or "",
                clean_start,
                payload,
                number_format or "",
                "true" if recalculate else "false",
                guard_mode,
            ],
            timeout=45.0,
        )
        if raw.startswith(f"BLOCKED{_GS}"):
            blocked_raw = raw.split(_GS, 1)[1]
            blocked_list: List[Dict[str, Any]] = []
            for row_s in blocked_raw.split(_RS):
                if not row_s:
                    continue
                p = row_s.split(_US)
                if len(p) >= 3:
                    blocked_list.append({
                        "sheet": p[0],
                        "cell": p[1],
                        "existing_formula": p[2],
                        "proposed_value": p[3] if len(p) > 3 else "",
                    })
            return self._formula_guard_response(guard_mode, blocked_list)
        parts = raw.split(_US)
        out: Dict[str, Any] = {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "sheet_name": parts[1] if len(parts) > 1 else "",
            "start_cell": clean_start,
            "rows_written": len(data),
            "columns_written": max_cols,
        }
        if len(parts) > 2 and parts[2].isdigit() and int(parts[2]) > 0:
            out["formulas_updated"] = int(parts[2])
        return out

    async def batch_write_live_excel_cells(
        self,
        cells: Any,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        recalculate: bool = True,
        allow_formula_overwrite: bool = False,
    ) -> Dict[str, Any]:
        """Write multiple non-contiguous cells/formulas in ONE osascript call and recalculate once.

        `cells` may be either:
          - A dict mapping cell ref to value: `{'G2': '=...', 'G11': 'N/A'}`
          - A list of dicts: `[{'cell': 'G2', 'value': '=...', 'sheet_name': 'Income Statement', 'number_format': '0.0%'}]`
        """
        normalized_items: List[Dict[str, Any]] = []
        if isinstance(cells, dict):
            for cell_ref, val in cells.items():
                normalized_items.append({
                    "sheet_name": sheet_name or "",
                    "cell": validate_cell_ref(str(cell_ref)),
                    "value": val,
                    "number_format": "",
                    "is_formula": False,
                })
        elif isinstance(cells, list):
            for item in cells:
                if not isinstance(item, dict):
                    raise ValueError("Each item in cells list must be a dict with 'cell' and 'value'")
                raw_cell = item.get("cell") or item.get("cell_ref")
                if not raw_cell:
                    raise ValueError("Each cell update must include 'cell' (e.g. 'G2')")
                normalized_items.append({
                    "sheet_name": item.get("sheet_name") or sheet_name or "",
                    "cell": validate_cell_ref(str(raw_cell)),
                    "value": item.get("value"),
                    "number_format": str(item.get("number_format") or ""),
                    "is_formula": bool(item.get("is_formula", False)),
                })
        else:
            raise ValueError("cells must be a dict ({'G2': ...}) or non-empty list of dicts")

        if not normalized_items:
            raise ValueError("cells cannot be empty")

        # Sanitize any cross-sheet formulas once against the live workbook's sheet list
        has_cross_sheet = any(
            (it["is_formula"] or (isinstance(it["value"], str) and it["value"].startswith("=")))
            and "!" in str(it["value"] or "")
            for it in normalized_items
        )
        live_sheets: List[str] = []
        if has_cross_sheet:
            try:
                ws_info = await self.list_live_excel_worksheets(workbook_name=workbook_name)
                live_sheets = list(ws_info.get("worksheets", []))
            except Exception:
                live_sheets = []

        encoded_rows: List[str] = []
        for it in normalized_items:
            val = it["value"]
            if val is None:
                vtype, vstr = "N", ""
            elif it["is_formula"] or (isinstance(val, str) and val.startswith("=")):
                vtype = "F"
                vstr = val if str(val).startswith("=") else f"={val}"
                if "!" in vstr and live_sheets:
                    vstr = self._sanitize_live_excel_formula(vstr, live_sheets)
            elif isinstance(val, bool):
                vtype, vstr = "B", ("true" if val else "false")
            elif isinstance(val, (int, float)):
                vtype, vstr = "D", str(val)
            else:
                vtype, vstr = "S", str(val)
            it["_encoded_value"] = vstr
            encoded_rows.append(_US.join([it["sheet_name"], it["cell"], vtype, vstr, it["number_format"]]))

        guard_mode = "off" if allow_formula_overwrite else self.get_formula_guard_mode()
        payload = _RS.join(encoded_rows)
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set payload to item 2 of argv
            set doCalc to (item 3 of argv) is "true"
            set guardMode to item 4 of argv
            set us to character id 31
            set rs to character id 30
            set gs to character id 29

            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set rowList to text items of payload
            set AppleScript's text item delimiters to prevDelims

            set targetRefs to {}
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set wbName to name of wb
                set prevAlerts to display alerts
                set display alerts to false
                try
                    set fmlUpdated to 0
                    if guardMode is not "off" then
                        set blockedRows to {}
                        repeat with rStr in rowList
                            set AppleScript's text item delimiters to us
                            set parts to text items of (rStr as text)
                            set AppleScript's text item delimiters to prevDelims
                            if (count of parts) >= 5 then
                                set shName to item 1 of parts
                                set cellRef to item 2 of parts
                                set tagChar to item 3 of parts
                                set rawVal to item 4 of parts
                                if shName is "" then
                                    set sh to active sheet of wb
                                else
                                    set sh to worksheet shName of wb
                                end if
                                set targetCell to range cellRef of sh
                                set existingFml to ""
                                try
                                    set rawEx to formula of targetCell
                                    if rawEx is not missing value then
                                        set existingFml to rawEx as text
                                    end if
                                end try
                                if (length of existingFml) > 0 and (character 1 of existingFml is "=") then
                                    if guardMode is "strict" or tagChar is not "F" then
                                        set end of blockedRows to ((name of sh) & us & cellRef & us & existingFml & us & rawVal)
                                    else
                                        set fmlUpdated to fmlUpdated + 1
                                    end if
                                end if
                            end if
                        end repeat
                        if (count of blockedRows) > 0 then
                            set display alerts to prevAlerts
                            set AppleScript's text item delimiters to rs
                            set joinedBlocked to blockedRows as text
                            set AppleScript's text item delimiters to prevDelims
                            return "BLOCKED" & gs & joinedBlocked
                        end if
                    end if

                    repeat with rStr in rowList
                        set AppleScript's text item delimiters to us
                        set parts to text items of (rStr as text)
                        set AppleScript's text item delimiters to prevDelims
                        if (count of parts) >= 5 then
                            set shName to item 1 of parts
                            set cellRef to item 2 of parts
                            set tagChar to item 3 of parts
                            set rawVal to item 4 of parts
                            set numFmt to item 5 of parts
                            if shName is "" then
                                set sh to active sheet of wb
                            else
                                set sh to worksheet shName of wb
                            end if
                            set actualSh to name of sh
                            set targetCell to range cellRef of sh
                            if tagChar is "N" then
                                clear contents targetCell
                            else if tagChar is "F" then
                                set formula of targetCell to rawVal
                            else if tagChar is "B" then
                                if rawVal is "true" then
                                    set value of targetCell to true
                                else
                                    set value of targetCell to false
                                end if
                            else if tagChar is "D" then
                                set value of targetCell to (rawVal as number)
                            else
                                set value of targetCell to rawVal
                            end if
                            if numFmt is not "" then
                                set number format of targetCell to numFmt
                            end if
                            set end of targetRefs to {actualSh, cellRef, targetCell}
                        end if
                    end repeat

                    if doCalc then
                        calculate
                    end if

                    set outRows to {}
                    repeat with entry in targetRefs
                        set actualSh to item 1 of entry
                        set cellRef to item 2 of entry
                        set cObj to item 3 of entry
                        set evalVal to value of cObj
                        if evalVal is missing value then
                            set evalText to ""
                        else
                            set evalText to evalVal as text
                        end if
                        set end of outRows to (actualSh & us & cellRef & us & evalText)
                    end repeat

                    set display alerts to prevAlerts
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell

            set AppleScript's text item delimiters to rs
            set joinedRows to outRows as text
            set AppleScript's text item delimiters to prevDelims
            return wbName & gs & joinedRows & gs & (fmlUpdated as text)
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", payload, "true" if recalculate else "false", guard_mode],
            timeout=45.0,
        )
        if raw.startswith(f"BLOCKED{_GS}"):
            blocked_raw = raw.split(_GS, 1)[1]
            blocked_list: List[Dict[str, Any]] = []
            for row_s in blocked_raw.split(_RS):
                if not row_s:
                    continue
                p = row_s.split(_US)
                if len(p) >= 3:
                    blocked_list.append({
                        "sheet": p[0],
                        "cell": p[1],
                        "existing_formula": p[2],
                        "proposed_value": p[3] if len(p) > 3 else "",
                    })
            return self._formula_guard_response(guard_mode, blocked_list)
        sections = raw.split(_GS)
        wb_name = sections[0] if sections else (workbook_name or "")
        rows_raw = sections[1] if len(sections) > 1 else ""
        fml_updated_str = sections[2] if len(sections) > 2 else "0"

        results: List[Dict[str, Any]] = []
        if rows_raw:
            for row in rows_raw.split(_RS):
                if not row:
                    continue
                parts = row.split(_US)
                if len(parts) >= 3:
                    results.append({
                        "sheet_name": parts[0],
                        "cell": parts[1],
                        "evaluated_value": parts[2],
                    })

        out: Dict[str, Any] = {
            "status": "success",
            "workbook_name": wb_name,
            "cells_written": len(results),
            "recalculated": recalculate,
            "results": results,
        }
        if fml_updated_str.isdigit() and int(fml_updated_str) > 0:
            out["formulas_updated"] = int(fml_updated_str)
        return sanitize_for_json(out)

    async def batch_read_live_excel_ranges(
        self,
        ranges: List[Dict[str, Any]],
        workbook_name: Optional[str] = None,
        include_formulas: bool = False,
    ) -> Dict[str, Any]:
        """Read multiple ranges across one or more live Excel sheets in ONE osascript call.

        Each item in `ranges` should be `{'sheet_name': 'Income Statement', 'range_ref': 'B2:F16'}`
        (`range` and `cell_range` are also accepted as aliases for `range_ref`).
        """
        if not ranges or not isinstance(ranges, list):
            raise ValueError("ranges must be a non-empty list of {'sheet_name', 'range_ref'} dicts")

        req_rows: List[str] = []
        cleaned_specs: List[Dict[str, str]] = []
        for item in ranges:
            if not isinstance(item, dict):
                raise ValueError("Each range spec must be a dict with 'range_ref'")
            raw_ref = item.get("range_ref") or item.get("range") or item.get("cell_range")
            if not raw_ref:
                raise ValueError("Each range spec must include 'range_ref' (e.g. 'B2:F16')")
            clean_ref = validate_range_ref(str(raw_ref))
            sh_name = str(item.get("sheet_name") or "")
            cleaned_specs.append({"sheet_name": sh_name, "range_ref": clean_ref})
            req_rows.append(_US.join([sh_name, clean_ref]))

        payload = _RS.join(req_rows)
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set payload to item 2 of argv
            set incFormulas to (item 3 of argv) is "true"
            set us to character id 31
            set rs to character id 30
            set gs to character id 29
            set fs to character id 28

            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set reqList to text items of payload
            set AppleScript's text item delimiters to prevDelims

            set rawResults to {}
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set wbName to name of wb
                try
                    calculate
                end try
                repeat with reqStr in reqList
                    set AppleScript's text item delimiters to us
                    set parts to text items of (reqStr as text)
                    set AppleScript's text item delimiters to prevDelims
                    if (count of parts) >= 2 then
                        set shName to item 1 of parts
                        set refStr to item 2 of parts
                        if shName is "" then
                            set sh to active sheet of wb
                        else
                            set sh to worksheet shName of wb
                        end if
                        set actualSh to name of sh
                        set rng to range refStr of sh
                        set rVals to value of rng
                        set rStrs to missing value
                        try
                            set rStrs to string value of rng
                        end try
                        set rFmls to missing value
                        if incFormulas then
                            set rFmls to formula of rng
                        end if
                        set end of rawResults to {actualSh, refStr, rVals, rFmls, rStrs}
                    end if
                end repeat
            end tell

            set outBlocks to {}
            repeat with itemRec in rawResults
                set actualSh to item 1 of itemRec
                set refStr to item 2 of itemRec
                set rVals to item 3 of itemRec
                set rFmls to item 4 of itemRec
                set rStrs to item 5 of itemRec
                set valBlock to my matrix_to_text(rVals, us, rs)
                set fmlBlock to ""
                if incFormulas then
                    set fmlBlock to my matrix_to_text(rFmls, us, rs)
                end if
                set strBlock to my matrix_to_text(rStrs, us, rs)
                set end of outBlocks to (actualSh & us & refStr & gs & valBlock & gs & fmlBlock & gs & strBlock)
            end repeat

            set AppleScript's text item delimiters to fs
            set joinedBlocks to outBlocks as text
            set AppleScript's text item delimiters to prevDelims
            return wbName & fs & joinedBlocks
        end run

        on matrix_to_text(rawData, us, rs)
            if rawData is missing value then
                return ""
            end if
            if class of rawData is not list then
                return (rawData as text)
            end if
            if (count of rawData) is 0 then
                return ""
            end if
            if class of (item 1 of rawData) is not list then
                set matrixRows to {rawData}
            else
                set matrixRows to rawData
            end if
            set outRows to {}
            repeat with rowItem in matrixRows
                set rowList to contents of rowItem
                if class of rowList is not list then
                    set rowList to {rowList}
                end if
                set cellTexts to {}
                repeat with cellItem in rowList
                    set v to contents of cellItem
                    if v is missing value then
                        set end of cellTexts to ""
                    else
                        set end of cellTexts to (v as text)
                    end if
                end repeat
                set prevDelims to AppleScript's text item delimiters
                set AppleScript's text item delimiters to us
                set end of outRows to (cellTexts as text)
                set AppleScript's text item delimiters to prevDelims
            end repeat
            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rs
            set joined to outRows as text
            set AppleScript's text item delimiters to prevDelims
            return joined
        end matrix_to_text
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", payload, "true" if include_formulas else "false"],
            timeout=45.0,
        )
        blocks = raw.split(_FS)
        wb_name = blocks[0] if blocks else (workbook_name or "")
        results: List[Dict[str, Any]] = []
        for blk in blocks[1:]:
            if not blk:
                continue
            secs = blk.split(_GS)
            hdr = secs[0].split(_US) if secs else ["", ""]
            val_block = secs[1] if len(secs) > 1 else ""
            fml_block = secs[2] if len(secs) > 2 else ""
            str_block = secs[3] if len(secs) > 3 else ""
            values = [row.split(_US) for row in val_block.split(_RS)] if val_block else []
            if str_block and values:
                str_rows = [row.split(_US) for row in str_block.split(_RS)]
                for r_i, row in enumerate(values):
                    if r_i < len(str_rows):
                        for c_i, cell_val in enumerate(row):
                            if (cell_val == "" or cell_val is None) and c_i < len(str_rows[r_i]):
                                sv = str_rows[r_i][c_i].strip()
                                if sv.startswith("#"):
                                    row[c_i] = sv
            range_str = hdr[1] if len(hdr) > 1 else ""
            entry: Dict[str, Any] = {
                "sheet_name": hdr[0],
                "range": range_str,
                "rows": len(values),
                "columns": len(values[0]) if values else 0,
                "values": values,
            }
            if include_formulas:
                entry["formulas"] = self._extract_sparse_formulas(fml_block, range_str)
            results.append(entry)

        return sanitize_for_json({
            "status": "success",
            "workbook_name": wb_name,
            "ranges_read": len(results),
            "results": results,
        })

    @staticmethod
    def _normalize_excel_full_name(raw_full_name: str) -> str:
        """Convert macOS HFS path ('Macintosh HD:Users:...') to POSIX ('/Users/...') for path matching."""
        if not raw_full_name:
            return ""
        s = raw_full_name.strip()
        if s.startswith("/") or s.startswith("http://") or s.startswith("https://"):
            return s
        # HFS paths use ':' as separator and start with volume name (e.g. 'Macintosh HD:Users:...')
        if ":" in s:
            parts = s.split(":")
            if len(parts) >= 2:
                return "/" + "/".join(parts[1:])
        return s

    async def find_open_excel_workbook_by_path(self, file_path: str) -> Optional[Dict[str, Any]]:
        """Return matching open live Excel workbook metadata if `file_path` is currently open in Excel."""
        if not await self.is_app_running(self.excel_app):
            return None
        target = Path(file_path).expanduser().resolve()
        target_str = str(target)
        target_name = target.name.lower()

        open_wbs = await self.list_open_excel_workbooks()
        for wb in open_wbs:
            norm_full = self._normalize_excel_full_name(wb.get("full_name", ""))
            if norm_full.startswith("/"):
                try:
                    if str(Path(norm_full).resolve()) == target_str:
                        return wb
                except Exception:
                    if norm_full == target_str:
                        return wb
            # CloudStorage / OneDrive URLs or basename-only full_name fallback
            if wb.get("name", "").lower() == target_name:
                if not norm_full or norm_full.startswith(("http://", "https://")) or ":" not in wb.get("full_name", ""):
                    return wb
        return None

    async def capture_excel_workbook_view_state(self, workbook_name: str) -> Dict[str, Any]:
        """Capture an open Excel workbook's window bounds, active sheet, selection, zoom, and scroll."""
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set us to character id 31
            set shName to ""
            set bStr to ""
            set zoomStr to ""
            set scrollR to ""
            set scrollC to ""
            set selAddr to ""
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then return ""
                set wb to workbook wbTarget
                try
                    set shName to name of active sheet of wb
                end try
                try
                    set w to window 1 of wb
                    try
                        set b to bounds of w
                        set bStr to ((item 1 of b) as text) & "," & ((item 2 of b) as text) & "," & ((item 3 of b) as text) & "," & ((item 4 of b) as text)
                    end try
                    try
                        set zoomStr to (zoom of w) as text
                    end try
                    try
                        set scrollR to (scroll row of active pane of w) as text
                    end try
                    try
                        set scrollC to (scroll column of active pane of w) as text
                    end try
                end try
                try
                    if (name of active workbook) is (name of wb) then
                        set selAddr to get address of selection without row absolute and column absolute
                    end if
                end try
            end tell
            return shName & us & bStr & us & zoomStr & us & scrollR & us & scrollC & us & selAddr
        end run
        """
        try:
            raw = await self.execute_applescript(script, args=[workbook_name], timeout=10.0)
        except Exception as exc:
            logger.warning(f"Could not capture Excel view state for '{workbook_name}': {exc}")
            return {}
        if not raw:
            return {}
        parts = raw.split(_US)
        state: Dict[str, Any] = {}
        if len(parts) > 0 and parts[0]:
            state["active_sheet"] = parts[0]
        if len(parts) > 1 and parts[1]:
            coords = [p.strip() for p in parts[1].split(",")]
            if len(coords) == 4 and all(c.lstrip("-").isdigit() for c in coords):
                state["bounds"] = [int(c) for c in coords]
        if len(parts) > 2 and parts[2]:
            try:
                state["zoom"] = int(float(parts[2]))
            except ValueError:
                pass
        if len(parts) > 3 and parts[3].isdigit():
            state["scroll_row"] = int(parts[3])
        if len(parts) > 4 and parts[4].isdigit():
            state["scroll_column"] = int(parts[4])
        if len(parts) > 5 and parts[5]:
            state["selection"] = parts[5]
        return state

    async def restore_excel_workbook_view_state(
        self,
        workbook_name: str,
        state: Dict[str, Any],
    ) -> bool:
        """Restore window bounds, active sheet, zoom, scroll position, and selection in live Excel."""
        if not state:
            return False
        bounds = state.get("bounds")
        b_arg = ",".join(str(int(x)) for x in bounds) if isinstance(bounds, list) and len(bounds) == 4 else ""
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shName to item 2 of argv
            set bStr to item 3 of argv
            set zoomStr to item 4 of argv
            set scrollR to item 5 of argv
            set scrollC to item 6 of argv
            set selAddr to item 7 of argv
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then return "false"
                set wb to workbook wbTarget
                set prevAlerts to display alerts
                set display alerts to false
                try
                    if shName is not "" then
                        try
                            activate object worksheet shName of wb
                        end try
                    end if
                    try
                        set w to window 1 of wb
                        if bStr is not "" then
                            set prevDelims to AppleScript's text item delimiters
                            set AppleScript's text item delimiters to ","
                            set bParts to text items of bStr
                            set AppleScript's text item delimiters to prevDelims
                            if (count of bParts) is 4 then
                                try
                                    set bounds of w to {(item 1 of bParts) as integer, (item 2 of bParts) as integer, (item 3 of bParts) as integer, (item 4 of bParts) as integer}
                                end try
                            end if
                        end if
                        if zoomStr is not "" then
                            try
                                set zoom of w to (zoomStr as integer)
                            end try
                        end if
                        if scrollR is not "" then
                            try
                                set scroll row of active pane of w to (scrollR as integer)
                            end try
                        end if
                        if scrollC is not "" then
                            try
                                set scroll column of active pane of w to (scrollC as integer)
                            end try
                        end if
                    end try
                    if selAddr is not "" then
                        try
                            select (range selAddr of active sheet of wb)
                        end try
                    end if
                    set display alerts to prevAlerts
                    return "true"
                on error
                    set display alerts to prevAlerts
                    return "false"
                end try
            end tell
        end run
        """
        try:
            res = await self.execute_applescript(
                script,
                args=[
                    workbook_name,
                    str(state.get("active_sheet") or ""),
                    b_arg,
                    str(state.get("zoom") or ""),
                    str(state.get("scroll_row") or ""),
                    str(state.get("scroll_column") or ""),
                    str(state.get("selection") or ""),
                ],
                timeout=10.0,
            )
            return res.lower() == "true"
        except Exception as exc:
            logger.warning(f"Could not restore Excel view state for '{workbook_name}': {exc}")
            return False

    async def save_excel_workbook(
        self,
        workbook_name: Optional[str] = None,
        file_path: Optional[str] = None,
        file_format: str = "xlsx",
    ) -> Dict[str, Any]:
        """Save a live Excel workbook in-place or Save As (.xlsx, .xlsm, .csv, .pdf)."""
        resolved_path = ""
        if file_path:
            p = resolve_file_path(file_path, must_exist=False)
            p.parent.mkdir(parents=True, exist_ok=True)
            resolved_path = str(p)

        fmt_key = file_format.lower().lstrip(".")
        valid_formats = {"xlsx", "xlsm", "csv", "pdf"}
        if fmt_key not in valid_formats:
            raise ValueError(f"Unsupported Excel save format '{file_format}'. Use one of {sorted(valid_formats)}.")

        script = """
        on run argv
            set wbTarget to item 1 of argv
            set posixOut to item 2 of argv
            set fmtKey to item 3 of argv
            set hfsPath to ""
            if posixOut is not "" then
                set hfsPath to (POSIX file posixOut) as text
            end if
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set prevAlerts to display alerts
                set display alerts to false
                try
                    if hfsPath is "" then
                        save wb
                        set savedName to name of wb
                    else
                        if fmtKey is "xlsm" then
                            save workbook as wb filename hfsPath file format macro enabled XML file format
                        else if fmtKey is "csv" then
                            save workbook as wb filename hfsPath file format CSV file format
                        else if fmtKey is "pdf" then
                            save workbook as wb filename hfsPath file format PDF file format
                        else
                            save workbook as wb filename hfsPath file format workbook normal file format
                        end if
                        set savedName to name of wb
                    end if
                    set display alerts to prevAlerts
                    return savedName
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        saved_name = await self.execute_applescript(
            script,
            args=[workbook_name or "", resolved_path, fmt_key],
            timeout=30.0,
        )
        return {
            "status": "success",
            "workbook_name": saved_name,
            "file_path": resolved_path or "(saved in place)",
            "format": fmt_key,
            "method": "applescript",
        }

    async def close_excel_workbook(
        self,
        workbook_name: Optional[str] = None,
        save_changes: str = "no",
    ) -> Dict[str, Any]:
        """Close a live Excel workbook cleanly without unsaved-change dialog hangs."""
        if save_changes not in ("no", "yes"):
            raise ValueError("save_changes must be 'no' or 'yes'")

        script = """
        on run argv
            set wbTarget to item 1 of argv
            set saveOpt to item 2 of argv
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    return ""
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set closedName to name of wb
                set prevAlerts to display alerts
                set display alerts to false
                try
                    if saveOpt is "yes" then
                        close wb saving yes
                    else
                        close wb saving no
                    end if
                    set display alerts to prevAlerts
                    return closedName
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        closed_name = await self.execute_applescript(
            script,
            args=[workbook_name or "", save_changes],
            timeout=20.0,
        )
        return {
            "status": "success",
            "closed_workbook": closed_name,
            "save_changes": save_changes,
            "method": "applescript",
        }

    async def excel_live_circle_invalid(
        self,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        clear: bool = False,
    ) -> Dict[str, Any]:
        """Trigger (`circle invalid`) or clear (`clear circles`) invalid-data circles in live Excel."""
        if not await self.is_app_running(self.excel_app):
            raise RuntimeError("Microsoft Excel is not running. Open a workbook in Excel first.")

        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set shouldClear to (item 3 of argv) is "true"
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                if shTarget is "" then
                    set sh to active sheet of wb
                else
                    set sh to worksheet shTarget of wb
                end if
                set prevAlerts to display alerts
                set display alerts to false
                try
                    if shouldClear then
                        clear circles sh
                    else
                        circle invalid sh
                    end if
                    set display alerts to prevAlerts
                    return (name of wb) & (character id 31) & (name of sh)
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", sheet_name or "", "true" if clear else "false"],
            timeout=15.0,
        )
        parts = raw.split(_US)
        return sanitize_for_json({
            "status": "success",
            "action": "clear_circles" if clear else "circle_invalid",
            "workbook_name": parts[0] if parts else (workbook_name or ""),
            "sheet_name": parts[1] if len(parts) > 1 else (sheet_name or ""),
            "method": "applescript",
        })

    async def excel_live_custom_views(
        self,
        action: str = "list",
        view_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """List or display (`show custom view`) Excel custom workbook views in live Excel."""
        act = action.strip().lower()
        if act not in ("list", "show"):
            raise ValueError("action must be 'list' or 'show'")
        if act == "show" and not view_name:
            raise ValueError("view_name is required when action is 'show'")

        if not await self.is_app_running(self.excel_app):
            return sanitize_for_json({
                "status": "success",
                "action": act,
                "running": False,
                "custom_views": [],
                "note": "Microsoft Excel is not running; Custom Views are only accessible in live Excel.",
            })

        if act == "list":
            script = """
            on run argv
                set wbTarget to item 1 of argv
                set us to character id 31
                set rs to character id 30
                set outRows to {}
                tell application "Microsoft Excel"
                    if (count of workbooks) is 0 then
                        return ""
                    end if
                    if wbTarget is "" then
                        set wb to active workbook
                    else
                        set wb to workbook wbTarget
                    end if
                    set vCount to 0
                    try
                        set vCount to count of custom views of wb
                    end try
                    if vCount > 0 then
                        repeat with i from 1 to vCount
                            set cv to custom view i of wb
                            set cvName to ""
                            set cvPrint to "false"
                            set cvRowCol to "false"
                            try
                                set cvName to (name of cv) as text
                            end try
                            try
                                set cvPrint to (custom view print settings of cv) as text
                            end try
                            try
                                set cvRowCol to (row col settings of cv) as text
                            end try
                            set end of outRows to (cvName & us & cvPrint & us & cvRowCol)
                        end repeat
                    end if
                end tell
                set prevDelims to AppleScript's text item delimiters
                set AppleScript's text item delimiters to rs
                set joinedOut to outRows as text
                set AppleScript's text item delimiters to prevDelims
                return joinedOut
            end run
            """
            raw = await self.execute_applescript(script, args=[workbook_name or ""], timeout=15.0)
            views: List[Dict[str, Any]] = []
            if raw:
                for row in raw.split(_RS):
                    if not row:
                        continue
                    parts = row.split(_US)
                    if len(parts) >= 3:
                        views.append({
                            "name": parts[0],
                            "print_settings": parts[1].lower() == "true",
                            "row_col_settings": parts[2].lower() == "true",
                        })
            return sanitize_for_json({
                "status": "success",
                "action": "list",
                "count": len(views),
                "custom_views": views,
                "method": "applescript",
            })

        # action == "show"
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set targetView to item 2 of argv
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set prevAlerts to display alerts
                set display alerts to false
                try
                    show custom view (custom view targetView of wb)
                    set display alerts to prevAlerts
                    return (name of wb) & (character id 31) & targetView
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script, args=[workbook_name or "", view_name or ""], timeout=15.0
        )
        parts = raw.split(_US)
        return sanitize_for_json({
            "status": "success",
            "action": "show",
            "workbook_name": parts[0] if parts else (workbook_name or ""),
            "view_name": parts[1] if len(parts) > 1 else view_name,
            "method": "applescript",
        })

    async def excel_live_update_table_style(
        self,
        table_name: str,
        sheet_name: Optional[str] = None,
        workbook_name: Optional[str] = None,
        style_name: Optional[str] = None,
        show_row_stripes: Optional[bool] = None,
        show_column_stripes: Optional[bool] = None,
        show_first_column: Optional[bool] = None,
        show_last_column: Optional[bool] = None,
    ) -> Dict[str, Any]:
        """Update a live Excel Table (list object) style and banded row/column toggles."""
        if not await self.is_app_running(self.excel_app):
            raise RuntimeError("Microsoft Excel is not running. Open a workbook in Excel first.")

        def _tri(val: Optional[bool]) -> str:
            if val is None:
                return ""
            return "true" if val else "false"

        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set tblName to item 3 of argv
            set styleName to item 4 of argv
            set rowStr to item 5 of argv
            set colStr to item 6 of argv
            set firstStr to item 7 of argv
            set lastStr to item 8 of argv
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                if shTarget is "" then
                    set sh to active sheet of wb
                else
                    set sh to worksheet shTarget of wb
                end if
                set prevAlerts to display alerts
                set display alerts to false
                try
                    set lo to list object tblName of sh
                    if styleName is not "" then
                        set table style of lo to styleName
                    end if
                    if rowStr is not "" then
                        set show table style row stripes of lo to (rowStr is "true")
                    end if
                    if colStr is not "" then
                        set show table style column stripes of lo to (colStr is "true")
                    end if
                    if firstStr is not "" then
                        set show table style first column of lo to (firstStr is "true")
                    end if
                    if lastStr is not "" then
                        set show table style last column of lo to (lastStr is "true")
                    end if
                    set display alerts to prevAlerts
                    return (name of wb) & (character id 31) & (name of sh) & (character id 31) & (name of lo)
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[
                workbook_name or "",
                sheet_name or "",
                table_name,
                style_name or "",
                _tri(show_row_stripes),
                _tri(show_column_stripes),
                _tri(show_first_column),
                _tri(show_last_column),
            ],
            timeout=15.0,
        )
        parts = raw.split(_US)
        return sanitize_for_json({
            "status": "success",
            "workbook_name": parts[0] if parts else (workbook_name or ""),
            "sheet_name": parts[1] if len(parts) > 1 else (sheet_name or ""),
            "table_name": parts[2] if len(parts) > 2 else table_name,
            "style_name": style_name,
            "method": "applescript",
        })

    @staticmethod
    def get_excel_view_capabilities() -> Dict[str, Any]:
        """Return factual capability metadata for Excel 16.115 & openpyxl UI/workbook features."""
        return {
            "status": "success",
            "excel_version_target": "16.115 (macOS)",
            "supported": {
                "table_styles": "Supported via openpyxl (add_excel_table, update_excel_table_style) and live Excel AppleScript (excel_live_update_table_style) with TableStyleLight*/Medium*/Dark* and banded row/column toggles.",
                "named_cell_styles": "Supported via openpyxl (apply_cell_style, list_cell_styles, or cell_style in formatting dicts) with auto-registered built-in presets (Normal, Good, Bad, Neutral, Input, Calculation, Check Cell, Linked Cell, Note, Output, Warning Text, Title, Heading 1-4, Total).",
                "data_validation": "Supported via openpyxl (add_data_validation, list_data_validations, clear_data_validations) for list, whole, decimal, date, time, textLength, and custom rules.",
                "invalid_data_circles": "Supported in live Excel 16.115 via AppleScript (excel_live_circle_invalid with clear=False/True using `circle invalid` and `clear circles`).",
                "custom_views": "Supported for listing and displaying existing workbook Custom Views in live Excel 16.115 via AppleScript (excel_live_custom_views). Not exposed in openpyxl file-level API.",
                "freeze_panes_and_gridlines": "Supported via openpyxl (format_worksheet, batch_format_workbook).",
            },
            "manual_ui_only": {
                "focus_cell": "Not exposed in openpyxl or Microsoft Excel 16.115 Excel.sdef. Toggle manually via View > Focus Cell in Excel.",
                "focus_cell_color": "Not exposed in openpyxl or Microsoft Excel 16.115 Excel.sdef. Change manually via View > Focus Cell > Focus Cell Color in Excel.",
                "show_auto_highlight": "Not exposed in openpyxl or Microsoft Excel 16.115 Excel.sdef. Toggle manually via View > Focus Cell > Show Auto-Highlight in Excel.",
            },
        }

    # ------------------------------------------------------------------
    # Microsoft PowerPoint (hardened existing methods)
    # ------------------------------------------------------------------

    async def launch_powerpoint(self, activate: bool = False) -> bool:
        """Launch Microsoft PowerPoint."""
        script = (
            'tell application "Microsoft PowerPoint" to activate'
            if activate
            else 'tell application "Microsoft PowerPoint" to launch'
        )
        await self.execute_applescript(script, timeout=15.0)
        return True

    async def open_powerpoint_file(self, file_path: str) -> bool:
        """Open a PowerPoint file safely via argv."""
        resolved = resolve_file_path(file_path, must_exist=True)
        script = """
        on run argv
            set targetFile to POSIX file (item 1 of argv)
            tell application "Microsoft PowerPoint"
                open targetFile
                activate
            end tell
        end run
        """
        await self.execute_applescript(script, args=[str(resolved)], timeout=30.0)
        return True
