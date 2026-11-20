"""
Logging utilities for Microsoft 365 for Mac MCP (m365-mac-mcp).

IMPORTANT: MCP stdio servers must NEVER write log output to stdout, because
stdout is reserved exclusively for JSON-RPC protocol frames. All console logs
are routed to sys.stderr.
"""

import logging
import os
import sys
from pathlib import Path
from typing import Optional


def setup_logger(
    name: str,
    level: Optional[str] = None,
    log_file: Optional[str] = None,
) -> logging.Logger:
    """Set up a logger that writes safely to stderr (never stdout).

    Args:
        name: Logger name.
        level: Logging level (defaults to OFFICE365_MCP_LOG_LEVEL or INFO).
        log_file: Optional log file path.

    Returns:
        Configured logger.
    """
    logger = logging.getLogger(name)

    if logger.handlers:
        return logger

    effective_level = (
        level
        or os.getenv("OFFICE365_MCP_LOG_LEVEL", "INFO")
    ).upper()
    logger.setLevel(getattr(logging, effective_level, logging.INFO))
    logger.propagate = False

    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # Always use stderr so MCP stdio JSON-RPC on stdout is never corrupted.
    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setFormatter(formatter)
    logger.addHandler(stderr_handler)

    if log_file:
        log_path = Path(log_file).expanduser().resolve()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_path)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger
