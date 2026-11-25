"""Core osascript execution, delimiters, timeouts, and Office bundle inspection."""

import asyncio
import plistlib
from pathlib import Path
from typing import (
    Any,
    Dict,
    Optional,
    Sequence,
)

from utils.logger import setup_logger

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


class AppleScriptBaseMixin:
    """Core osascript runner and macOS Office bundle/status helpers."""

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

    async def _get_frontmost_app_name(self) -> str:
        """Return the name of the currently frontmost macOS application process (best-effort)."""
        try:
            script = """
            on run argv
                tell application "System Events"
                    try
                        return (name of first application process whose frontmost is true) as text
                    end try
                end tell
                return ""
            end run
            """
            res = await self.execute_applescript(script, args=[""], timeout=4.0)
            return str(res).strip() if res else ""
        except Exception:
            return ""

    async def _is_app_frontmost(self, app_name: str) -> bool:
        """Return True ONLY when `app_name` is currently the frontmost macOS application."""
        try:
            if not await self.is_app_running(app_name):
                return False
            front = await self._get_frontmost_app_name()
            return front == app_name
        except Exception:
            return False

    async def _hide_app_process(self, app_name: str) -> None:
        """Hide an Office process in the background when no user files are open and another app is frontmost."""
        try:
            if not await self.is_app_running(app_name):
                return
            script = """
            on run argv
                set targetApp to item 1 of argv
                tell application "System Events"
                    try
                        if exists process targetApp then
                            set visible of process targetApp to false
                        end if
                    end try
                end tell
            end run
            """
            await self.execute_applescript(script, args=[app_name], timeout=4.0)
        except Exception:
            pass

    async def _open_via_launchservices(
        self,
        app_name: str,
        resolved: Path,
        activate: bool = False,
    ) -> None:
        """Open a trusted Project Working Folder file via macOS LaunchServices (`/usr/bin/open`).

        LaunchServices grants the sandboxed Office app a security-scoped token for that exact file
        so macOS does not show a Grant Access dialog. ONLY invoked for files strictly inside the
        active client/Bionic Project Working Folder.
        """
        from utils.validators import get_background_mode

        bg_mode = get_background_mode()
        effective_activate = bool(activate) and bg_mode == "off"
        cmd = ["/usr/bin/open"]
        if not effective_activate:
            cmd.append("-g")
            if bg_mode != "off":
                cmd.append("-j")
        cmd.extend(["-a", app_name, str(resolved)])
        proc = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        _, stderr_bytes = await asyncio.wait_for(proc.communicate(), timeout=15.0)
        if proc.returncode != 0:
            err_txt = stderr_bytes.decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"LaunchServices open failed ({proc.returncode}): {err_txt}")
