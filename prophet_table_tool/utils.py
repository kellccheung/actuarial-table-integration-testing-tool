"""Shared utilities: hashing, run IDs, path helpers, value comparison."""

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


def normalize_key_part(value: object) -> str:
    """
    Canonicalize one key cell for matching.

    Numeric-looking values become a stable form so ``1.10`` and ``1.1`` join
    as the same key; whole numbers stay without a trailing ``.0`` (``20`` not
    ``20.0``). Non-numeric values stay as text.
    """
    if value is None:
        return ""
    s = str(value)
    if s == "":
        return ""
    try:
        f = float(s)
    except (TypeError, ValueError):
        return s
    if f.is_integer() and abs(f) < 1e15:
        return str(int(f))
    return str(f)


def values_equal(a: object, b: object) -> bool:
    """
    Compare two cell values: exact string match, or numeric equality when both
    parse as floats (``1.10`` == ``1.1``).
    """
    sa = "" if a is None else str(a)
    sb = "" if b is None else str(b)
    if sa == sb:
        return True
    if sa == "" or sb == "":
        return False
    try:
        return float(sa) == float(sb)
    except (TypeError, ValueError):
        return False


def format_normalized_key(values: list | tuple | None) -> str:
    """Pipe-join key parts after numeric normalization."""
    if values is None:
        return ""
    return "|".join(normalize_key_part(v) for v in values)
