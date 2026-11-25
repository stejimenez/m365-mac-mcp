"""
Configuration management for Microsoft 365 for Mac MCP (m365-mac-mcp).
"""

import io
import json
import os
import tempfile
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Optional, Sequence, Union


@dataclass
class ServerConfig:
    """Server configuration settings."""

    log_level: str = "INFO"
    temp_dir: str = str(Path(tempfile.gettempdir()) / "m365-mac-mcp")
    max_presentations: int = 20
    max_documents: int = 20
    max_workbooks: int = 20
    applescript_timeout: float = 30.0
    enable_applescript: bool = True
    enable_cloud_api: bool = False
    templates_dir: Optional[str] = None
    ppt_template: Optional[str] = None
    word_template: Optional[str] = None
    output_dir: Optional[str] = None


class Config:
    """Configuration manager for the MCP server."""

    def __init__(self, config_file: Optional[str] = None):
        self.config_file = config_file or self._get_default_config_path()
        self.settings = self._load_config()

    def _get_default_config_path(self) -> str:
        """Get the default configuration file path."""
        return str(Path(__file__).parent.parent.parent / "config.json")

    def _load_config(self) -> ServerConfig:
        """Load configuration from file and environment variable overrides."""
        config_data: dict[str, Any] = {}

        if Path(self.config_file).exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                    if isinstance(loaded, dict):
                        valid_keys = set(ServerConfig.__dataclass_fields__.keys())
                        config_data = {
                            k: v for k, v in loaded.items() if k in valid_keys
                        }
            except Exception as exc:
                import logging as _logging
                _logging.getLogger(__name__).warning(
                    f"Ignoring invalid config file '{self.config_file}': {exc}"
                )

        env_overrides = {
            "log_level": os.getenv("M365_MCP_LOG_LEVEL") or os.getenv("OFFICE365_MCP_LOG_LEVEL"),
            "temp_dir": os.getenv("M365_MCP_TEMP_DIR") or os.getenv("OFFICE365_MCP_TEMP_DIR"),
            "max_presentations": os.getenv("M365_MCP_MAX_PRESENTATIONS") or os.getenv("OFFICE365_MCP_MAX_PRESENTATIONS"),
            "max_documents": os.getenv("M365_MCP_MAX_DOCUMENTS") or os.getenv("OFFICE365_MCP_MAX_DOCUMENTS"),
            "max_workbooks": os.getenv("M365_MCP_MAX_WORKBOOKS") or os.getenv("OFFICE365_MCP_MAX_WORKBOOKS"),
            "applescript_timeout": os.getenv("M365_MCP_APPLESCRIPT_TIMEOUT") or os.getenv("OFFICE365_MCP_APPLESCRIPT_TIMEOUT"),
            "enable_applescript": os.getenv("M365_MCP_ENABLE_APPLESCRIPT") or os.getenv("OFFICE365_MCP_ENABLE_APPLESCRIPT"),
            "enable_cloud_api": os.getenv("M365_MCP_ENABLE_CLOUD_API") or os.getenv("OFFICE365_MCP_ENABLE_CLOUD_API"),
            "templates_dir": os.getenv("M365_MCP_TEMPLATES_DIR") or os.getenv("OFFICE365_MCP_TEMPLATES_DIR"),
            "ppt_template": os.getenv("M365_MCP_PPT_TEMPLATE") or os.getenv("OFFICE365_MCP_PPT_TEMPLATE"),
            "word_template": os.getenv("M365_MCP_WORD_TEMPLATE") or os.getenv("OFFICE365_MCP_WORD_TEMPLATE"),
            "output_dir": os.getenv("M365_MCP_OUTPUT_DIR") or os.getenv("OFFICE365_MCP_OUTPUT_DIR"),
        }

        for key, value in env_overrides.items():
            if value is None or not value.strip():
                continue
            value = value.strip()
            if key in ("max_presentations", "max_documents", "max_workbooks"):
                try:
                    config_data[key] = int(value)
                except ValueError:
                    continue
            elif key == "applescript_timeout":
                try:
                    config_data[key] = float(value)
                except ValueError:
                    continue
            elif key in ("enable_applescript", "enable_cloud_api"):
                config_data[key] = value.lower() in ("true", "1", "yes")
            else:
                config_data[key] = value

        if "temp_dir" in config_data:
            config_data["temp_dir"] = str(Path(config_data["temp_dir"]).expanduser())

        return ServerConfig(**config_data)

    def save_config(self) -> None:
        """Save current configuration to file."""
        config_path = Path(self.config_file).expanduser()
        config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(asdict(self.settings), f, indent=2)

    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value."""
        return getattr(self.settings, key, default)

    def set(self, key: str, value: Any) -> None:
        """Set a configuration value."""
        if hasattr(self.settings, key):
            setattr(self.settings, key, value)


def resolve_template_path(
    explicit_path: Optional[str],
    default_env_var: str,
    allowed_extensions: Sequence[str] = (".pptx", ".potx", ".docx", ".dotx"),
) -> Optional[Path]:
    """Resolve an explicit or default corporate template path.

    Resolution order:
    1. ``explicit_path`` (if provided): treated as a direct path if it exists, or looked up
       inside ``M365_MCP_TEMPLATES_DIR`` when given as a bare filename/relative path.
    2. ``os.getenv(default_env_var)`` (e.g. ``M365_MCP_PPT_TEMPLATE`` or
       ``M365_MCP_WORD_TEMPLATE``): resolved the same way against ``M365_MCP_TEMPLATES_DIR``.
    3. ``None`` when neither is configured (blank default).
    """
    raw = (explicit_path or "").strip()
    is_explicit = bool(raw)
    if not raw:
        raw = (os.getenv(default_env_var) or "").strip()
    if not raw:
        return None

    templates_dir_raw = (
        os.getenv("M365_MCP_TEMPLATES_DIR") or os.getenv("OFFICE365_MCP_TEMPLATES_DIR") or ""
    ).strip()
    templates_dir = Path(templates_dir_raw).expanduser().resolve() if templates_dir_raw else None

    candidate = Path(raw).expanduser()
    resolved: Optional[Path] = None
    if candidate.is_absolute():
        if candidate.exists():
            resolved = candidate.resolve()
    else:
        if templates_dir is not None and (templates_dir / candidate).exists():
            resolved = (templates_dir / candidate).resolve()
        elif candidate.exists():
            resolved = candidate.resolve()

    if resolved is None:
        if is_explicit:
            hint = f" (also checked M365_MCP_TEMPLATES_DIR='{templates_dir}')" if templates_dir else ""
            raise FileNotFoundError(f"Template file not found: '{raw}'{hint}")
        # Default env var points to missing file — fall back gracefully rather than breaking creation
        return None

    ext = resolved.suffix.lower()
    allowed_lower = {e.lower() for e in allowed_extensions}
    if allowed_lower and ext not in allowed_lower:
        raise ValueError(
            f"Unsupported template extension '{ext}' for '{resolved.name}'. "
            f"Allowed: {sorted(allowed_lower)}"
        )
    return resolved


def open_office_template(path: Path) -> Union[str, io.BytesIO]:
    """Return a python-pptx/python-docx compatible source for .pptx/.docx/.potx/.dotx templates.

    python-pptx and python-docx reject .potx/.dotx strictly because of the OOXML main part
    content-type string in ``[Content_Types].xml``. For .potx/.dotx, rewrite that single
    content-type in an in-memory BytesIO copy so corporate templates load seamlessly.
    """
    ext = path.suffix.lower()
    if ext not in (".potx", ".dotx"):
        return str(path)

    replacements = {
        ".potx": (
            b"application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
            b"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        ),
        ".dotx": (
            b"application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
            b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        ),
    }
    old_ct, new_ct = replacements[ext]
    buf = io.BytesIO()
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(old_ct, new_ct)
            zout.writestr(item, data)
    buf.seek(0)
    return buf
