"""Live Microsoft Excel workbook, worksheet, view-state, and table-style automation."""

import asyncio
from pathlib import Path
from typing import (
    Any,
    Dict,
    List,
    Optional,
    Tuple,
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

class AppleScriptExcelCoreMixin:
    """Live Microsoft Excel workbook, worksheet, view-state, and table-style automation."""

    async def _preflight_excel_background_state(self) -> Tuple[str, str]:
        """Single osascript pre-flight: capture (prev_front_app, prev_front_wb) and hide idle Excel if not frontmost."""
        try:
            script = """
            on run argv
                set us to character id 31
                set prevApp to ""
                set exRunning to false
                tell application "System Events"
                    try
                        set prevApp to (name of first application process whose frontmost is true) as text
                    end try
                    try
                        set exRunning to exists process "Microsoft Excel"
                    end try
                end tell
                set prevWb to ""
                if exRunning then
                    tell application "Microsoft Excel"
                        try
                            if (count of workbooks) > 0 then
                                set prevWb to (name of active workbook) as text
                            end if
                        end try
                    end tell
                    if prevApp is not "Microsoft Excel" and prevWb is "" then
                        tell application "System Events"
                            try
                                set visible of process "Microsoft Excel" to false
                            end try
                        end tell
                    end if
                end if
                return prevApp & us & prevWb
            end run
            """
            raw = await self.execute_applescript(script, args=[""], timeout=5.0)
            if isinstance(raw, str) and _US in raw:
                p_app, p_wb = raw.split(_US, 1)
                return p_app.strip(), p_wb.strip()
            return "", ""
        except Exception:
            return "", ""

    async def _get_front_excel_workbook_name(self) -> str:
        """Return the currently active Excel workbook name if Excel is running (best-effort, non-intrusive)."""
        try:
            if not await self.is_app_running(self.excel_app):
                return ""
            script = """
            on run argv
                tell application "Microsoft Excel"
                    try
                        if (count of workbooks) > 0 then
                            return (name of active workbook) as text
                        end if
                    end try
                end tell
                return ""
            end run
            """
            raw = await self.execute_applescript(script, args=[""], timeout=5.0)
            return raw.strip() if isinstance(raw, str) else ""
        except Exception:
            return ""

    async def launch_excel(self, activate: bool = False) -> bool:
        """Launch Microsoft Excel without stealing focus unless allowed by M365_MCP_BACKGROUND_MODE."""
        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"
        script = (
            'tell application "Microsoft Excel" to activate'
            if effective_activate
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
        target_str = str(resolved)
        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"

        # 1. Instant return if already open in Excel (seamless recovery after user clicks Grant Access)
        # Never minimize/hide a workbook the user already had open — only suppress focus stealing.
        try:
            existing = await self.find_open_excel_workbook_by_path(target_str)
            if existing is not None:
                wb_name = str(existing.get("name") or resolved.name)
                calc_script = """
                on run argv
                    set wbTarget to item 1 of argv
                    set shouldActivate to (item 2 of argv) is "true"
                    tell application "Microsoft Excel"
                        if shouldActivate then activate
                        try
                            calculate
                        end try
                        return name of active sheet of workbook wbTarget
                    end tell
                end run
                """
                active_sh = ""
                try:
                    active_sh = await self.execute_applescript(
                        calc_script,
                        args=[wb_name, "true" if effective_activate else "false"],
                        timeout=15.0,
                    )
                except Exception:
                    active_sh = str(existing.get("active_sheet") or "")
                return {
                    "status": "success",
                    "workbook_name": wb_name,
                    "active_sheet": active_sh,
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
                if bg_mode != "off":
                    prev_front_app, prev_front = await self._preflight_excel_background_state()
                else:
                    prev_front_app, prev_front = "", ""
                await self._open_via_launchservices(self.excel_app, resolved, activate=effective_activate)
                start_t = asyncio.get_running_loop().time()
                deadline = start_t + 10.0
                while asyncio.get_running_loop().time() < deadline:
                    elapsed = asyncio.get_running_loop().time() - start_t
                    await asyncio.sleep(0.05 if elapsed < 2.0 else 0.2)
                    found = await self.find_open_excel_workbook_by_path(target_str)
                    if found is not None:
                        wb_name = str(found.get("name") or resolved.name)
                        calc_script = """
                        on run argv
                            set wbTarget to item 1 of argv
                            set bgMode to item 2 of argv
                            set prevFront to item 3 of argv
                            set prevFrontApp to item 4 of argv
                            set wasFrontmost to (prevFrontApp is "Microsoft Excel")
                            tell application "Microsoft Excel"
                                try
                                    calculate
                                end try
                                if bgMode is "strict" then
                                    try
                                        set wTarget to window 1 of workbook wbTarget
                                        set visible of wTarget to false
                                        if visible of wTarget is true then
                                            set collapsed of wTarget to true
                                        end if
                                    on error
                                        try
                                            set collapsed of (window 1 of workbook wbTarget) to true
                                        end try
                                    end try
                                else if bgMode is "no_focus" then
                                    try
                                        set collapsed of (window 1 of workbook wbTarget) to true
                                    on error
                                        try
                                            set window state of (window 1 of workbook wbTarget) to window state minimized
                                        end try
                                    end try
                                end if
                                if bgMode is not "off" and wasFrontmost and prevFront is not "" and prevFront is not wbTarget then
                                    try
                                        activate object (window 1 of workbook prevFront)
                                    on error
                                        try
                                            activate object workbook prevFront
                                        end try
                                    end try
                                end if
                                set activeSh to name of active sheet of workbook wbTarget
                            end tell
                            if bgMode is not "off" and (not wasFrontmost) then
                                tell application "System Events"
                                    if prevFront is "" then
                                        try
                                            if exists process "Microsoft Excel" then
                                                set visible of process "Microsoft Excel" to false
                                            end if
                                        end try
                                    end if
                                    if prevFrontApp is not "" then
                                        try
                                            if (name of first application process whose frontmost is true) is "Microsoft Excel" then
                                                set frontmost of process prevFrontApp to true
                                            end if
                                        end try
                                    end if
                                end tell
                            end if
                            return activeSh
                        end run
                        """
                        active_sh = ""
                        try:
                            active_sh = await self.execute_applescript(
                                calc_script,
                                args=[wb_name, bg_mode, prev_front, prev_front_app],
                                timeout=15.0,
                            )
                        except Exception:
                            active_sh = str(found.get("active_sheet") or "")
                        return {
                            "status": "success",
                            "workbook_name": wb_name,
                            "active_sheet": active_sh,
                            "file_path": target_str,
                            "read_only": False,
                            "trusted_project_dir": True,
                            "background_mode": bg_mode,
                            "method": "launchservices+applescript",
                        }
            except Exception as exc:
                logger.warning(f"LaunchServices open_excel_file fallback to AppleScript: {exc}")

        # 3. Everywhere outside the Project Working Folder (or read_only=True): standard AppleScript so macOS asks Grant Access
        # Keep fallback <= 15s inside Project Folder (worst-case 10+15=25s, well below 60s/90s client timeouts)
        open_timeout = get_grant_access_timeout() if not is_within_trusted_project_dir(resolved) else 15.0
        script = """
        on run argv
            set posixPath to item 1 of argv
            set isReadOnly to (item 2 of argv) is "true"
            set shouldActivate to (item 3 of argv) is "true"
            set bgMode to item 4 of argv
            set targetFile to POSIX file posixPath
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Excel" then
                            set wasFrontmost to true
                        end if
                    end try
                end tell
            end if
            tell application "Microsoft Excel"
                set prevAlerts to display alerts
                set display alerts to false
                set prevScreen to true
                try
                    set prevScreen to screen updating
                end try
                set prevFront to ""
                if bgMode is not "off" then
                    try
                        if (count of workbooks) > 0 then
                            set prevFront to (name of active workbook) as text
                        end if
                    end try
                end if
                try
                    if bgMode is not "off" then
                        try
                            set screen updating to false
                        end try
                    end if
                    if shouldActivate then activate
                    set wb to open workbook workbook file name (targetFile as text) read only isReadOnly
                    try
                        calculate
                    end try
                    set wbName to name of wb
                    set activeSh to name of active sheet of wb
                    if bgMode is "strict" then
                        try
                            set wTarget to window 1 of wb
                            set visible of wTarget to false
                            if visible of wTarget is true then
                                set collapsed of wTarget to true
                            end if
                        on error
                            try
                                set collapsed of (window 1 of wb) to true
                            end try
                        end try
                    else if bgMode is "no_focus" then
                        try
                            set collapsed of (window 1 of wb) to true
                        on error
                            try
                                set window state of (window 1 of wb) to window state minimized
                            end try
                        end try
                    end if
                    if bgMode is not "off" and wasFrontmost and prevFront is not "" and prevFront is not wbName then
                        try
                            activate object (window 1 of workbook prevFront)
                        on error
                            try
                                activate object workbook prevFront
                            end try
                        end try
                    end if
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                on error errMsg number errNum
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
            if bgMode is not "off" and (not wasFrontmost) then
                tell application "System Events"
                    if prevFront is "" then
                        try
                            if exists process "Microsoft Excel" then
                                set visible of process "Microsoft Excel" to false
                            end if
                        end try
                    end if
                    if prevFrontApp is not "" then
                        try
                            if (name of first application process whose frontmost is true) is "Microsoft Excel" then
                                set frontmost of process prevFrontApp to true
                            end if
                        end try
                    end if
                end tell
            end if
            return wbName & (character id 31) & activeSh
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[target_str, "true" if read_only else "false", "true" if effective_activate else "false", bg_mode],
            timeout=open_timeout,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "active_sheet": parts[1] if len(parts) > 1 else "",
            "file_path": target_str,
            "read_only": read_only,
            "background_mode": bg_mode,
            "method": "applescript",
        }

    async def create_excel_workbook(self, activate: bool = False) -> Dict[str, Any]:
        """Create a new blank workbook in live Microsoft Excel."""
        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"
        script = """
        on run argv
            set shouldActivate to (item 1 of argv) is "true"
            set bgMode to item 2 of argv
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Excel" then
                            set wasFrontmost to true
                        end if
                    end try
                end tell
            end if
            tell application "Microsoft Excel"
                set prevFront to ""
                if bgMode is not "off" then
                    try
                        if (count of workbooks) > 0 then
                            set prevFront to (name of active workbook) as text
                        end if
                    end try
                end if
                if shouldActivate then activate
                set wb to make new workbook
                set wbName to name of wb
                set activeSh to name of active sheet of wb
                if bgMode is "strict" then
                    try
                        set wTarget to window 1 of wb
                        set visible of wTarget to false
                        if visible of wTarget is true then
                            set collapsed of wTarget to true
                        end if
                    on error
                        try
                            set collapsed of (window 1 of wb) to true
                        end try
                    end try
                else if bgMode is "no_focus" then
                    try
                        set collapsed of (window 1 of wb) to true
                    on error
                        try
                            set window state of (window 1 of wb) to window state minimized
                        end try
                    end try
                end if
                if bgMode is not "off" and wasFrontmost and prevFront is not "" and prevFront is not wbName then
                    try
                        activate object (window 1 of workbook prevFront)
                    on error
                        try
                            activate object workbook prevFront
                        end try
                    end try
                end if
            end tell
            if bgMode is not "off" and (not wasFrontmost) then
                tell application "System Events"
                    if prevFront is "" then
                        try
                            if exists process "Microsoft Excel" then
                                set visible of process "Microsoft Excel" to false
                            end if
                        end try
                    end if
                    if prevFrontApp is not "" then
                        try
                            if (name of first application process whose frontmost is true) is "Microsoft Excel" then
                                set frontmost of process prevFrontApp to true
                            end if
                        end try
                    end if
                end tell
            end if
            return wbName & (character id 31) & activeSh
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=["true" if effective_activate else "false", bg_mode],
            timeout=20.0,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "active_sheet": parts[1] if len(parts) > 1 else "",
            "background_mode": bg_mode,
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
            set rs to character id 30
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set wbName to (name of wb) as text
                set activeSh to ""
                try
                    set activeSh to (name of active sheet of wb) as text
                on error
                    try
                        set activeSh to (name of active sheet) as text
                    end try
                end try
                set shNames to {}
                set shCount to count of worksheets of wb
                repeat with j from 1 to shCount
                    set end of shNames to (name of worksheet j of wb) as text
                end repeat
                set prevDelims to AppleScript's text item delimiters
                set AppleScript's text item delimiters to us
                set joinedSheets to shNames as text
                set AppleScript's text item delimiters to prevDelims
                return wbName & rs & activeSh & rs & joinedSheets
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
        """Create a new worksheet in the targeted live Excel workbook without stealing OS focus."""
        bg_mode = get_background_mode()
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set newName to item 2 of argv
            set bgMode to item 3 of argv
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Excel" then set wasFrontmost to true
                    end try
                end tell
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
                set prevCol to false
                set prevVis to true
                set prevScreen to true
                try
                    set prevScreen to screen updating
                end try
                try
                    set wTarget to window 1 of wb
                    set prevCol to collapsed of wTarget
                    set prevVis to visible of wTarget
                end try
                try
                    if bgMode is not "off" then
                        try
                            set screen updating to false
                        end try
                    end if
                    set newSh to make new worksheet at end of wb
                    set name of newSh to newName
                    set wbNameOut to (name of wb) as text
                    if bgMode is not "off" then
                        try
                            set wTarget to window 1 of wb
                            if (not prevVis) or bgMode is "strict" then
                                set visible of wTarget to false
                            else if prevCol or (not wasFrontmost) then
                                set collapsed of wTarget to true
                            end if
                        end try
                    end if
                    try
                        set screen updating to prevScreen
                    end try
                on error errMsg number errNum
                    try
                        set screen updating to prevScreen
                    end try
                    error errMsg number errNum
                end try
            end tell
            if bgMode is not "off" and (not wasFrontmost) and prevFrontApp is not "" then
                tell application "System Events"
                    try
                        if (name of first application process whose frontmost is true) is "Microsoft Excel" then
                            set frontmost of process prevFrontApp to true
                        end if
                    end try
                end tell
            end if
            return wbNameOut & (character id 31) & newName
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", sheet_name, bg_mode],
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
        """Switch the active worksheet in the targeted live Excel workbook without un-minimizing or stealing focus."""
        bg_mode = get_background_mode()
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shTarget to item 2 of argv
            set bgMode to item 3 of argv
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Excel" then set wasFrontmost to true
                    end try
                end tell
            end if
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then
                    error "No open Excel workbooks"
                end if
                set prevFrontWb to ""
                try
                    set prevFrontWb to (name of active workbook) as text
                end try
                if wbTarget is "" then
                    set wb to active workbook
                else
                    set wb to workbook wbTarget
                end if
                set wbNameOut to (name of wb) as text
                set prevCol to false
                set prevVis to true
                set prevScreen to true
                try
                    set prevScreen to screen updating
                end try
                try
                    set wTarget to window 1 of wb
                    set prevCol to collapsed of wTarget
                    set prevVis to visible of wTarget
                end try
                try
                    if bgMode is not "off" then
                        try
                            set screen updating to false
                        end try
                    end if
                    activate object worksheet shTarget of wb
                    set activeSh to shTarget
                    try
                        set activeSh to (name of active sheet of wb) as text
                    on error
                        try
                            set activeSh to (name of active sheet) as text
                        end try
                    end try
                    if bgMode is not "off" then
                        try
                            set wTarget to window 1 of wb
                            if (not prevVis) or bgMode is "strict" then
                                set visible of wTarget to false
                            else if prevCol or (not wasFrontmost) then
                                set collapsed of wTarget to true
                            end if
                        end try
                        if wasFrontmost and prevFrontWb is not "" and prevFrontWb is not wbNameOut then
                            try
                                activate object (window 1 of workbook prevFrontWb)
                            end try
                        end if
                    end if
                    try
                        set screen updating to prevScreen
                    end try
                on error errMsg number errNum
                    try
                        set screen updating to prevScreen
                    end try
                    error errMsg number errNum
                end try
            end tell
            if bgMode is not "off" and (not wasFrontmost) and prevFrontApp is not "" then
                tell application "System Events"
                    try
                        if (name of first application process whose frontmost is true) is "Microsoft Excel" then
                            set frontmost of process prevFrontApp to true
                        end if
                    end try
                end tell
            end if
            return wbNameOut & (character id 31) & activeSh
        end run
        """
        raw = await self.execute_applescript(
            script,
            args=[workbook_name or "", sheet_name, bg_mode],
            timeout=15.0,
        )
        parts = raw.split(_US)
        return {
            "status": "success",
            "workbook_name": parts[0] if parts else "",
            "active_sheet": parts[1] if len(parts) > 1 else sheet_name,
            "method": "applescript",
        }

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
        """Capture an open Excel workbook's window bounds, collapsed/visible state, active sheet, selection, zoom, and scroll."""
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
            set colStr to "false"
            set visStr to "true"
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then return ""
                set wb to workbook wbTarget
                try
                    set shName to name of active sheet of wb
                end try
                try
                    set w to window 1 of wb
                    try
                        set colStr to (collapsed of w) as text
                    end try
                    try
                        set visStr to (visible of w) as text
                    end try
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
            return shName & us & bStr & us & zoomStr & us & scrollR & us & scrollC & us & selAddr & us & colStr & us & visStr
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
        if len(parts) > 6 and parts[6]:
            state["collapsed"] = parts[6].strip().lower() == "true"
        if len(parts) > 7 and parts[7]:
            state["visible"] = parts[7].strip().lower() != "false"
        return state

    async def restore_excel_workbook_view_state(
        self,
        workbook_name: str,
        state: Dict[str, Any],
    ) -> bool:
        """Restore active sheet, zoom, scroll (and bounds/selection ONLY when safe under M365_MCP_BACKGROUND_MODE)."""
        if not state:
            return False
        bg_mode = get_background_mode()
        bounds = state.get("bounds")
        b_arg = ",".join(str(int(x)) for x in bounds) if isinstance(bounds, list) and len(bounds) == 4 else ""
        was_collapsed = "true" if state.get("collapsed") else "false"
        was_visible = "false" if state.get("visible") is False else "true"
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set shName to item 2 of argv
            set bStr to item 3 of argv
            set zoomStr to item 4 of argv
            set scrollR to item 5 of argv
            set scrollC to item 6 of argv
            set selAddr to item 7 of argv
            set bgMode to item 8 of argv
            set wasCollapsed to (item 9 of argv) is "true"
            set wasVisible to (item 10 of argv) is "true"
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Excel" then set wasFrontmost to true
                    end try
                end tell
            end if
            set keepIsolated to (bgMode is not "off") and ((not wasFrontmost) or wasCollapsed or (not wasVisible))
            tell application "Microsoft Excel"
                if (count of workbooks) is 0 then return "false"
                set wb to workbook wbTarget
                set prevAlerts to display alerts
                set display alerts to false
                set prevScreen to true
                try
                    set prevScreen to screen updating
                end try
                try
                    if keepIsolated then
                        try
                            set screen updating to false
                        end try
                    end if
                    if shName is not "" then
                        try
                            activate object worksheet shName of wb
                        end try
                    end if
                    try
                        set w to window 1 of wb
                        if (not keepIsolated) and bStr is not "" then
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
                        if keepIsolated then
                            if (not wasVisible) or bgMode is "strict" then
                                try
                                    set visible of w to false
                                on error
                                    try
                                        set collapsed of w to true
                                    end try
                                end try
                            else
                                try
                                    set collapsed of w to true
                                end try
                            end if
                        end if
                    end try
                    if (not keepIsolated) and selAddr is not "" then
                        try
                            select (range selAddr of active sheet of wb)
                        end try
                    end if
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                    set restoreOk to "true"
                on error
                    try
                        set screen updating to prevScreen
                    end try
                    set display alerts to prevAlerts
                    set restoreOk to "false"
                end try
            end tell
            if bgMode is not "off" and (not wasFrontmost) and prevFrontApp is not "" then
                tell application "System Events"
                    try
                        if (name of first application process whose frontmost is true) is "Microsoft Excel" then
                            set frontmost of process prevFrontApp to true
                        end if
                    end try
                end tell
            end if
            return restoreOk
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
                    bg_mode,
                    was_collapsed,
                    was_visible,
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
            p = resolve_file_path(file_path, must_exist=False, for_write=True)
            p.parent.mkdir(parents=True, exist_ok=True)
            resolved_path = str(p)

        fmt_key = file_format.lower().lstrip(".")
        valid_formats = {"xlsx", "xlsm", "csv", "pdf"}
        if fmt_key not in valid_formats:
            raise ValueError(f"Unsupported Excel save format '{file_format}'. Use one of {sorted(valid_formats)}.")

        bg_mode = get_background_mode()
        script = """
        on run argv
            set wbTarget to item 1 of argv
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
                        if prevFrontApp is "Microsoft Excel" then set wasFrontmost to true
                    end try
                end tell
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
                set prevCol to false
                set prevVis to true
                try
                    set wTarget to window 1 of wb
                    set prevCol to collapsed of wTarget
                    set prevVis to visible of wTarget
                end try
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
                    if bgMode is not "off" then
                        try
                            set wTarget to window 1 of wb
                            if (not prevVis) or bgMode is "strict" then
                                set visible of wTarget to false
                            else if prevCol then
                                set collapsed of wTarget to true
                            end if
                        end try
                    end if
                    set display alerts to prevAlerts
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
            if bgMode is not "off" and (not wasFrontmost) and prevFrontApp is not "" then
                tell application "System Events"
                    try
                        if (name of first application process whose frontmost is true) is "Microsoft Excel" then
                            set frontmost of process prevFrontApp to true
                        end if
                    end try
                end tell
            end if
            return savedName
        end run
        """
        saved_name = await self.execute_applescript(
            script,
            args=[workbook_name or "", resolved_path, fmt_key, bg_mode],
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
        bg_mode = get_background_mode()
        script = """
        on run argv
            set wbTarget to item 1 of argv
            set targetView to item 2 of argv
            set bgMode to item 3 of argv
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft Excel" then set wasFrontmost to true
                    end try
                end tell
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
                set prevCol to false
                set prevVis to true
                try
                    set wTarget to window 1 of wb
                    set prevCol to collapsed of wTarget
                    set prevVis to visible of wTarget
                end try
                set prevAlerts to display alerts
                set display alerts to false
                try
                    show custom view (custom view targetView of wb)
                    set wbNameOut to (name of wb) as text
                    if bgMode is not "off" and (not wasFrontmost) then
                        try
                            set wTarget to window 1 of wb
                            if (not prevVis) or bgMode is "strict" then
                                set visible of wTarget to false
                            else if prevCol then
                                set collapsed of wTarget to true
                            end if
                        end try
                    end if
                    set display alerts to prevAlerts
                on error errMsg number errNum
                    set display alerts to prevAlerts
                    error errMsg number errNum
                end try
            end tell
            if bgMode is not "off" and (not wasFrontmost) and prevFrontApp is not "" then
                tell application "System Events"
                    try
                        if (name of first application process whose frontmost is true) is "Microsoft Excel" then
                            set frontmost of process prevFrontApp to true
                        end if
                    end try
                end tell
            end if
            return wbNameOut & (character id 31) & targetView
        end run
        """
        raw = await self.execute_applescript(
            script, args=[workbook_name or "", view_name or "", bg_mode], timeout=15.0
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
