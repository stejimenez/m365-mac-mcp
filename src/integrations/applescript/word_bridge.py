"""Live Microsoft Word AppleScript automation."""

import asyncio
from pathlib import Path
from typing import (
    Any,
    Dict,
    List,
    Optional,
)

from utils.logger import setup_logger
from utils.validators import (
    get_background_mode,
    get_grant_access_timeout,
    is_within_trusted_project_dir,
    resolve_file_path,
    sanitize_for_json,
)

logger = setup_logger(__name__)
from integrations.applescript.base import _GS, _RS, _US

class AppleScriptWordMixin:
    """Live Microsoft Word AppleScript automation."""

    async def launch_word(self, activate: bool = False) -> bool:
        """Launch Microsoft Word without stealing focus unless allowed by M365_MCP_BACKGROUND_MODE."""
        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"
        script = (
            'tell application "Microsoft Word" to activate'
            if effective_activate
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
            set uSep to character id 31
            set rSep to character id 30
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
                                set end of outRows to (dName & uSep & dFull & uSep & dPath & uSep & dSaved & uSep & (pCount as text) & uSep & isAct)
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
                                    set end of outRows to (dName & uSep & dFull & uSep & dPath & uSep & dSaved & uSep & (pCount as text) & uSep & isAct)
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
                        set end of outRows to (dName & uSep & dFull & uSep & dPath & uSep & dSaved & uSep & (pCount as text) & uSep & isAct)
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
                            set end of outRows to (dName & uSep & dFull & uSep & dPath & uSep & dSaved & uSep & (pCount as text) & uSep & "true")
                        end if
                    end try
                end tell
            end if

            set prevDelims to AppleScript's text item delimiters
            set AppleScript's text item delimiters to rSep
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

    async def _preflight_word_background_state(self, prev_front: str = "") -> str:
        """Single osascript pre-flight: capture frontmost app and hide idle Word if not frontmost with 0 user docs."""
        try:
            script = """
            on run argv
                set prevDoc to item 1 of argv
                set prevApp to ""
                tell application "System Events"
                    try
                        set prevApp to (name of first application process whose frontmost is true) as text
                    end try
                    if prevApp is not "Microsoft Word" and prevDoc is "" then
                        try
                            if exists process "Microsoft Word" then
                                set visible of process "Microsoft Word" to false
                            end if
                        end try
                    end if
                end tell
                return prevApp
            end run
            """
            raw = await self.execute_applescript(script, args=[prev_front], timeout=4.0)
            return raw.strip() if isinstance(raw, str) else ""
        except Exception:
            return ""

    async def _find_open_word_doc_name_fast(self, target_name_low: str) -> Optional[str]:
        """Lightweight poll helper checking only open Word document names (avoids full paragraph/table scan every 50ms)."""
        try:
            script = """
            on run argv
                set rSep to character id 30
                set outNames to {}
                tell application "System Events"
                    if not (exists process "Microsoft Word") then return ""
                end tell
                tell application "Microsoft Word"
                    try
                        set wCount to count of windows
                        repeat with i from 1 to wCount
                            try
                                set dName to (name of document of window i) as text
                                if dName is not "" then set end of outNames to dName
                            end try
                        end repeat
                    end try
                    if (count of outNames) is 0 then
                        try
                            set actN to (name of active document) as text
                            if actN is not "" then set end of outNames to actN
                        end try
                    end if
                end tell
                set prevDelims to AppleScript's text item delimiters
                set AppleScript's text item delimiters to rSep
                set joined to outNames as text
                set AppleScript's text item delimiters to prevDelims
                return joined
            end run
            """
            raw = await self.execute_applescript(script, args=[""], timeout=5.0)
            for name in (raw or "").split(_RS):
                if name and name.strip().lower() == target_name_low:
                    return name.strip()
        except Exception:
            pass
        return None

    async def open_word_file(
        self,
        file_path: str,
        read_only: bool = False,
        activate: bool = False,
    ) -> Dict[str, Any]:
        """Open an existing Word document (.docx/.doc/.rtf) in live Microsoft Word."""
        resolved = resolve_file_path(file_path, must_exist=True)
        target_str = str(resolved)
        target_name_low = resolved.name.lower()
        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"
        prev_front = ""

        # 1. Instant return if already open in Word (seamless recovery after Grant Access click)
        # Never minimize/hide a document the user already had open — also capture active user doc as prev_front.
        try:
            if await self.is_app_running(self.word_app):
                open_docs = await self.list_open_word_documents()
                for doc in open_docs:
                    doc_full = str(doc.get("full_name") or "")
                    doc_name = str(doc.get("name") or "")
                    if doc.get("is_active") and doc_name and not prev_front:
                        prev_front = doc_name
                    if doc_name.lower() == target_name_low or (
                        doc_full.startswith("/") and str(Path(doc_full).resolve()) == target_str
                    ):
                        if effective_activate:
                            await self.execute_applescript('tell application "Microsoft Word" to activate', timeout=10.0)
                        return {
                            "status": "success",
                            "document_name": doc_name or resolved.name,
                            "file_path": target_str,
                            "read_only": read_only,
                            "already_open": True,
                            "background_mode": bg_mode,
                            "method": "already_open+applescript",
                        }
        except Exception:
            pass

        # 2. Strictly scoped to active client/Bionic Project Working Folder ONLY (waives Grant Access inside project)
        if not read_only and is_within_trusted_project_dir(resolved):
            try:
                prev_front_app = (
                    await self._preflight_word_background_state(prev_front)
                    if bg_mode != "off"
                    else ""
                )
                await self._open_via_launchservices(self.word_app, resolved, activate=effective_activate)
                start_t = asyncio.get_running_loop().time()
                deadline = start_t + 8.0
                while asyncio.get_running_loop().time() < deadline:
                    elapsed = asyncio.get_running_loop().time() - start_t
                    await asyncio.sleep(0.05 if elapsed < 2.0 else 0.2)
                    doc_name = await self._find_open_word_doc_name_fast(target_name_low)
                    if doc_name:
                            if bg_mode != "off":
                                isolate_script = """
                                on run argv
                                    set targetDocName to item 1 of argv
                                    set bgMode to item 2 of argv
                                    set prevFront to item 3 of argv
                                    set prevFrontApp to item 4 of argv
                                    set wasFrontmost to (prevFrontApp is "Microsoft Word")
                                    tell application "Microsoft Word"
                                        try
                                            set targetWin to active window of document targetDocName
                                            if bgMode is "strict" then
                                                try
                                                    set visible of targetWin to false
                                                    if visible of targetWin is true then
                                                        set collapsed of targetWin to true
                                                    end if
                                                on error
                                                    try
                                                        set collapsed of targetWin to true
                                                    on error
                                                        try
                                                            set window state of targetWin to window state minimize
                                                        end try
                                                    end try
                                                end try
                                            else if bgMode is "no_focus" then
                                                try
                                                    set collapsed of targetWin to true
                                                on error
                                                    try
                                                        set window state of targetWin to window state minimize
                                                    end try
                                                end try
                                            end if
                                        end try
                                        if wasFrontmost and prevFront is not "" and prevFront is not targetDocName then
                                            try
                                                activate object document prevFront
                                            end try
                                        end if
                                    end tell
                                    if not wasFrontmost then
                                        tell application "System Events"
                                            if prevFront is "" then
                                                try
                                                    if exists process "Microsoft Word" then
                                                        set visible of process "Microsoft Word" to false
                                                    end if
                                                end try
                                            end if
                                            if prevFrontApp is not "" then
                                                try
                                                    if (name of first application process whose frontmost is true) is "Microsoft Word" then
                                                        set frontmost of process prevFrontApp to true
                                                    end if
                                                end try
                                            end if
                                        end tell
                                    end if
                                end run
                                """
                                try:
                                    await self.execute_applescript(
                                        isolate_script,
                                        args=[doc_name, bg_mode, prev_front, prev_front_app],
                                        timeout=10.0,
                                    )
                                except Exception:
                                    pass
                            return {
                                "status": "success",
                                "document_name": doc_name,
                                "file_path": target_str,
                                "read_only": False,
                                "trusted_project_dir": True,
                                "background_mode": bg_mode,
                                "method": "launchservices+applescript",
                            }
            except Exception as exc:
                logger.warning(f"LaunchServices open_word_file fallback to AppleScript: {exc}")

        # 3. Everywhere outside the Project Working Folder (or read_only=True): standard AppleScript so macOS asks Grant Access
        # Keep fallback <= 15s inside Project Folder (worst-case 8+15=23s, well below 60s client timeout)
        open_timeout = get_grant_access_timeout() if not is_within_trusted_project_dir(resolved) else 15.0
        script = """
        on run argv
            set posixPath to item 1 of argv
            set isReadOnly to (item 2 of argv) is "true"
            set shouldActivate to (item 3 of argv) is "true"
            set bgMode to item 4 of argv
            set prevFront to item 5 of argv
            set targetFile to POSIX file posixPath
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Word" then
                            set wasFrontmost to true
                        end if
                    end try
                end tell
            end if
            tell application "Microsoft Word"
                if bgMode is not "off" and prevFront is "" then
                    try
                        if (count of documents) > 0 then
                            set prevFront to (name of active document) as text
                        end if
                    end try
                end if
                if shouldActivate then activate
                open file name (targetFile as text) read only isReadOnly
                set docName to name of active document
                if bgMode is not "off" then
                    try
                        set targetWin to active window of document docName
                        if bgMode is "strict" then
                            try
                                set visible of targetWin to false
                                if visible of targetWin is true then
                                    set collapsed of targetWin to true
                                end if
                            on error
                                try
                                    set collapsed of targetWin to true
                                on error
                                    try
                                        set window state of targetWin to window state minimize
                                    end try
                                end try
                            end try
                        else if bgMode is "no_focus" then
                            try
                                set collapsed of targetWin to true
                            on error
                                try
                                    set window state of targetWin to window state minimize
                                end try
                            end try
                        end if
                    end try
                    if wasFrontmost and prevFront is not "" and prevFront is not docName then
                        try
                            activate object document prevFront
                        end try
                    end if
                end if
            end tell
            if bgMode is not "off" and (not wasFrontmost) then
                tell application "System Events"
                    if prevFront is "" then
                        try
                            if exists process "Microsoft Word" then
                                set visible of process "Microsoft Word" to false
                            end if
                        end try
                    end if
                    if prevFrontApp is not "" then
                        try
                            if (name of first application process whose frontmost is true) is "Microsoft Word" then
                                set frontmost of process prevFrontApp to true
                            end if
                        end try
                    end if
                end tell
            end if
            return docName
        end run
        """
        doc_name = await self.execute_applescript(
            script,
            args=[
                target_str,
                "true" if read_only else "false",
                "true" if effective_activate else "false",
                bg_mode,
                prev_front,
            ],
            timeout=open_timeout,
        )
        return {
            "status": "success",
            "document_name": doc_name,
            "file_path": target_str,
            "read_only": read_only,
            "background_mode": bg_mode,
            "method": "applescript",
        }

    async def create_word_document(
        self,
        title: Optional[str] = None,
        activate: bool = False,
    ) -> Dict[str, Any]:
        """Create a new blank document in live Microsoft Word."""
        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"
        script = """
        on run argv
            set initialTitle to item 1 of argv
            set shouldActivate to (item 2 of argv) is "true"
            set bgMode to item 3 of argv
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Word" then
                            set wasFrontmost to true
                        end if
                    end try
                end tell
            end if
            tell application "Microsoft Word"
                set prevFront to ""
                if bgMode is not "off" then
                    try
                        if (count of documents) > 0 then
                            set prevFront to (name of active document) as text
                        end if
                    end try
                end if
                if shouldActivate then activate
                set newDoc to make new document
                if initialTitle is not "" then
                    insert text (initialTitle & return) at beginning of text object of newDoc
                end if
                set docName to name of newDoc
                if bgMode is not "off" then
                    try
                        set targetWin to active window of newDoc
                        if bgMode is "strict" then
                            try
                                set visible of targetWin to false
                                if visible of targetWin is true then
                                    set collapsed of targetWin to true
                                end if
                            on error
                                try
                                    set collapsed of targetWin to true
                                end try
                            end try
                        else if bgMode is "no_focus" then
                            try
                                set collapsed of targetWin to true
                            on error
                                try
                                    set window state of targetWin to window state minimize
                                end try
                            end try
                        end if
                    end try
                    if wasFrontmost and prevFront is not "" and prevFront is not docName then
                        try
                            activate object document prevFront
                        end try
                    end if
                end if
            end tell
            if bgMode is not "off" and (not wasFrontmost) then
                tell application "System Events"
                    if prevFront is "" then
                        try
                            if exists process "Microsoft Word" then
                                set visible of process "Microsoft Word" to false
                            end if
                        end try
                    end if
                    if prevFrontApp is not "" then
                        try
                            if (name of first application process whose frontmost is true) is "Microsoft Word" then
                                set frontmost of process prevFrontApp to true
                            end if
                        end try
                    end if
                end tell
            end if
            return docName
        end run
        """
        doc_name = await self.execute_applescript(
            script,
            args=[title or "", "true" if effective_activate else "false", bg_mode],
            timeout=20.0,
        )
        return {
            "status": "success",
            "document_name": doc_name,
            "title": title or "",
            "background_mode": bg_mode,
            "method": "applescript",
        }

    async def read_live_word_document(
        self,
        document_name: Optional[str] = None,
        include_tables: bool = True,
        include_full_text: bool = False,
        max_paragraphs: int = 200,
        start_paragraph: int = 0,
    ) -> Dict[str, Any]:
        """Read paragraphs, selection state, and optionally tables from a live Word doc (paginated; omits duplicate full text by default)."""
        script = """
        on run argv
            set targetName to item 1 of argv
            set incTables to (item 2 of argv) is "true"
            set uSep to character id 31
            set rSep to character id 30
            set gSep to character id 29
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
                            set AppleScript's text item delimiters to uSep
                            set end of rowStrings to (cellStrings as text)
                        end repeat
                        set AppleScript's text item delimiters to rSep
                        set end of tableBlocks to (rowStrings as text)
                    end repeat
                end if

                set AppleScript's text item delimiters to gSep
                set tablesJoined to tableBlocks as text
                set metaJoined to docName & uSep & docFull & uSep & docSaved & uSep & (selStart as text) & uSep & (selEnd as text) & uSep & selText
                return metaJoined & gSep & fullText & gSep & tablesJoined
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

        total_paras = len(paragraphs)
        start_idx = max(0, int(start_paragraph))
        para_cap = max(1, int(max_paragraphs))
        end_idx = min(total_paras, start_idx + para_cap)
        sliced_paragraphs = paragraphs[start_idx:end_idx]
        truncated = end_idx < total_paras

        tables: List[List[List[str]]] = []
        tables_truncated = False
        max_tbl_rows = 100
        if include_tables and len(sections) > 2:
            for tbl_raw in sections[2:]:
                if not tbl_raw:
                    continue
                rows_parsed: List[List[str]] = []
                raw_rows = tbl_raw.split(_RS)
                if len(raw_rows) > max_tbl_rows:
                    tables_truncated = True
                    raw_rows = raw_rows[:max_tbl_rows]
                for row_raw in raw_rows:
                    # Strip Word cell-end marker (\r\x07)
                    cells = [
                        c.replace("\r\x07", "").replace("\x07", "").strip()
                        for c in row_raw.split(_US)
                    ]
                    rows_parsed.append(cells)
                tables.append(rows_parsed)

        payload: Dict[str, Any] = {
            "status": "success",
            "document_name": meta_parts[0] if len(meta_parts) > 0 else "",
            "full_name": meta_parts[1] if len(meta_parts) > 1 else "",
            "saved": (meta_parts[2].lower() == "true") if len(meta_parts) > 2 else False,
            "selection": {
                "start": int(meta_parts[3]) if len(meta_parts) > 3 and meta_parts[3].lstrip("-").isdigit() else -1,
                "end": int(meta_parts[4]) if len(meta_parts) > 4 and meta_parts[4].lstrip("-").isdigit() else -1,
                "text": (meta_parts[5] if len(meta_parts) > 5 else "").replace("\r", "\n"),
            },
            "paragraph_count": total_paras,
            "paragraphs": sliced_paragraphs,
            "table_count": len(tables),
            "tables": tables,
            "method": "applescript",
        }
        if include_full_text:
            payload["text"] = "\n".join(sliced_paragraphs)
        if truncated or start_idx > 0:
            payload["returned_paragraphs"] = len(sliced_paragraphs)
            payload["start_paragraph"] = start_idx
        if truncated:
            payload["truncated"] = True
            payload["pagination_hint"] = (
                f"Showing paragraphs {start_idx}..{end_idx - 1} of {total_paras}. "
                f"Pass start_paragraph={end_idx} to read the next slice."
            )
        if tables_truncated:
            payload["tables_truncated"] = True
        return sanitize_for_json(payload)

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
        mode_aliases = {
            "insert": "insert_at_selection",
            "replace": "replace_all",
            "overwrite": "replace_all",
        }
        mode = mode_aliases.get((mode or "append").strip().lower(), (mode or "append").strip().lower())
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
        target_path_obj: Optional[Path] = None
        created_placeholder = False
        if file_path:
            p = resolve_file_path(file_path, must_exist=False, for_write=True)
            p.parent.mkdir(parents=True, exist_ok=True)
            target_path_obj = p
            resolved_path = str(p)

        fmt_key = file_format.lower().lstrip(".")
        valid_formats = {"docx", "pdf", "rtf", "txt", "doc"}
        if fmt_key not in valid_formats:
            raise ValueError(f"Unsupported Word save format '{file_format}'. Use one of {sorted(valid_formats)}.")

        # macOS 27.x / Word 16.115.1 sandbox quirk: `save as ... file format format PDF`
        # raises -1708 on CloudStorage (Dropbox/iCloud/OneDrive) paths unless the destination
        # file node already exists before AppleScript invokes `save as`.
        if target_path_obj is not None and fmt_key == "pdf" and not target_path_obj.exists():
            try:
                target_path_obj.touch()
                created_placeholder = True
            except Exception:
                pass

        bg_mode = get_background_mode()
        script = """
        on run argv
            set targetName to item 1 of argv
            set posixOut to item 2 of argv
            set fmtKey to item 3 of argv
            set bgMode to item 4 of argv
            set hfsPath to ""
            if posixOut is not "" then
                set hfsPath to (POSIX file posixOut) as text
            end if
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Word" then set wasFrontmost to true
                    end try
                end tell
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
                set prevCol to false
                try
                    set prevCol to collapsed of (active window of targetDoc)
                end try
                if hfsPath is "" then
                    save targetDoc
                    set savedName to name of targetDoc
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
                    try
                        set savedName to name of targetDoc
                    on error
                        set savedName to name of active document
                    end try
                end if
                if bgMode is not "off" and (prevCol or (not wasFrontmost)) then
                    try
                        set collapsed of (active window of targetDoc) to true
                    end try
                end if
            end tell
            if bgMode is not "off" and (not wasFrontmost) and prevFrontApp is not "" then
                tell application "System Events"
                    try
                        if (name of first application process whose frontmost is true) is "Microsoft Word" then
                            set frontmost of process prevFrontApp to true
                        end if
                    end try
                end tell
            end if
            return savedName
        end run
        """
        try:
            saved_name = await self.execute_applescript(
                script,
                args=[document_name or "", resolved_path, fmt_key, bg_mode],
                timeout=30.0,
            )
        except Exception:
            if created_placeholder and target_path_obj is not None and target_path_obj.exists():
                try:
                    if target_path_obj.stat().st_size == 0:
                        target_path_obj.unlink()
                except Exception:
                    pass
            raise
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
