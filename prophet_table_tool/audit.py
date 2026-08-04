"""Audit trail logging for every tool run."""

from __future__ import annotations

import logging
from datetime import datetime, timezone
from pathlib import Path


def setup_audit_logger(
    output_path: Path,
    run_id: str,
    mode: str,
) -> tuple[logging.Logger, Path]:
    """
    Create a dedicated audit logger writing to ``Output/Audit/<run_id>_<mode>.log``.

    Returns (logger, log_path).
    """
    audit_dir = Path(output_path) / "Audit"
    audit_dir.mkdir(parents=True, exist_ok=True)
    log_path = audit_dir / f"{run_id}_{mode}.log"

    logger = logging.getLogger(f"prophet_table_tool.audit.{run_id}.{mode}")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    logger.propagate = False

    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(message)s"))
    logger.addHandler(fh)

    # Also echo warnings+ to stderr
    sh = logging.StreamHandler()
    sh.setLevel(logging.WARNING)
    sh.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
    logger.addHandler(sh)

    logger.info("timestamp=%s", datetime.now(timezone.utc).isoformat())
    logger.info("run_id=%s", run_id)
    logger.info("mode=%s", mode)
    return logger, log_path


def close_audit_logger(logger: logging.Logger) -> None:
    """Flush and close all handlers on the audit logger."""
    for handler in list(logger.handlers):
        handler.flush()
        handler.close()
        logger.removeHandler(handler)
