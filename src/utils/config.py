"""
Configuration management for Microsoft 365 for Mac MCP (m365-mac-mcp).
"""

import json
import os
import tempfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Optional


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
            except Exception:
                pass

        env_overrides = {
            "log_level": os.getenv("M365_MCP_LOG_LEVEL") or os.getenv("OFFICE365_MCP_LOG_LEVEL"),
            "temp_dir": os.getenv("M365_MCP_TEMP_DIR") or os.getenv("OFFICE365_MCP_TEMP_DIR"),
            "max_presentations": os.getenv("M365_MCP_MAX_PRESENTATIONS") or os.getenv("OFFICE365_MCP_MAX_PRESENTATIONS"),
            "max_documents": os.getenv("M365_MCP_MAX_DOCUMENTS") or os.getenv("OFFICE365_MCP_MAX_DOCUMENTS"),
            "max_workbooks": os.getenv("M365_MCP_MAX_WORKBOOKS") or os.getenv("OFFICE365_MCP_MAX_WORKBOOKS"),
            "applescript_timeout": os.getenv("M365_MCP_APPLESCRIPT_TIMEOUT") or os.getenv("OFFICE365_MCP_APPLESCRIPT_TIMEOUT"),
            "enable_applescript": os.getenv("M365_MCP_ENABLE_APPLESCRIPT") or os.getenv("OFFICE365_MCP_ENABLE_APPLESCRIPT"),
            "enable_cloud_api": os.getenv("M365_MCP_ENABLE_CLOUD_API") or os.getenv("OFFICE365_MCP_ENABLE_CLOUD_API"),
        }

        for key, value in env_overrides.items():
            if value is None or not value.strip():
                continue
            value = value.strip()
            if key in ("max_presentations", "max_documents", "max_workbooks"):
                config_data[key] = int(value)
            elif key == "applescript_timeout":
                config_data[key] = float(value)
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
