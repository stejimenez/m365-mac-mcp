"""Live Microsoft PowerPoint AppleScript automation."""


from utils.logger import setup_logger
from utils.validators import get_background_mode, resolve_file_path

logger = setup_logger(__name__)

class AppleScriptPowerPointMixin:
    """Live Microsoft PowerPoint AppleScript automation."""

    async def launch_powerpoint(self, activate: bool = False) -> bool:
        """Launch Microsoft PowerPoint without stealing focus unless allowed by M365_MCP_BACKGROUND_MODE."""
        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"
        script = (
            'tell application "Microsoft PowerPoint" to activate'
            if effective_activate
            else 'tell application "Microsoft PowerPoint" to launch'
        )
        await self.execute_applescript(script, timeout=15.0)
        return True

    async def open_powerpoint_file(self, file_path: str, activate: bool = False) -> bool:
        """Open a PowerPoint file safely via argv, honoring M365_MCP_BACKGROUND_MODE."""
        resolved = resolve_file_path(file_path, must_exist=True)
        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"
        script = """
        on run argv
            set targetFile to POSIX file (item 1 of argv)
            set shouldActivate to (item 2 of argv) is "true"
            set bgMode to item 3 of argv
            set prevFrontApp to ""
            set wasFrontmost to false
            if bgMode is not "off" then
                tell application "System Events"
                    try
                        set prevFrontApp to (name of first application process whose frontmost is true) as text
                        if prevFrontApp is "Microsoft PowerPoint" then set wasFrontmost to true
                    end try
                end tell
            end if
            set prevCount to 0
            tell application "Microsoft PowerPoint"
                try
                    set prevCount to count of presentations
                end try
                open targetFile
                if shouldActivate then
                    activate
                else if bgMode is "strict" then
                    try
                        set visible of window 1 to false
                        if visible of window 1 is true then
                            set collapsed of window 1 to true
                        end if
                    on error
                        try
                            set collapsed of window 1 to true
                        end try
                    end try
                else if bgMode is "no_focus" then
                    try
                        set collapsed of window 1 to true
                    on error
                        try
                            set miniaturized of window 1 to true
                        end try
                    end try
                end if
            end tell
            if bgMode is not "off" and (not wasFrontmost) then
                tell application "System Events"
                    if prevCount is 0 then
                        try
                            if exists process "Microsoft PowerPoint" then
                                set visible of process "Microsoft PowerPoint" to false
                            end if
                        end try
                    end if
                    if prevFrontApp is not "" then
                        try
                            if (name of first application process whose frontmost is true) is "Microsoft PowerPoint" then
                                set frontmost of process prevFrontApp to true
                            end if
                        end try
                    end if
                end tell
            end if
        end run
        """
        await self.execute_applescript(
            script,
            args=[str(resolved), "true" if effective_activate else "false", bg_mode],
            timeout=30.0,
        )
        return True
