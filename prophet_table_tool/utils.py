"""Shared utilities: hashing, run IDs, path helpers."""

from __future__ import annotations

import hashlib
from datetime import datetime
from pathlib import Path


def generate_run_id() -> str:
    """Return a run identifier as YYYYMMDD_HHMMSS."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def file_hash(path: Path, algorithm: str = "sha256") -> str:
    """Compute a hex digest of a file's contents."""
    h = hashlib.new(algorithm)
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def resolve_path(base: Path, path_str: str) -> Path:
    """Resolve a path that may be relative to *base* or absolute."""
    p = Path(path_str)
    if p.is_absolute():
        return p
    return (base / p).resolve()


def yn_to_bool(value: object) -> bool:
    """Convert Control.xlsx Y/N (or bool-like) cells to bool."""
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    return str(value).strip().upper() in {"Y", "YES", "TRUE", "1"}


def table_stem(filename: str | Path) -> str:
    """Return the table name (filename without .csv extension)."""
    return Path(filename).stem
