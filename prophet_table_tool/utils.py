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


def unwrap_wrapping_quotes(value: str) -> str:
    """Drop one pair of double quotes that wraps the whole field.

    Spaces outside that pair are dropped with it. Unquoted text is unchanged.
    shortcut: one wrapping pair only; upgrade if a field holds a comma or an escaped quote.
    """
    trimmed = value.strip()
    if len(trimmed) >= 2 and trimmed.startswith('"') and trimmed.endswith('"'):
        return trimmed[1:-1]
    return value


def normalize_key_part(value: object) -> str:
    """
    Canonicalize one key cell for matching.

    One wrapping pair of double quotes is ignored first (``"PROD_A"`` joins
    ``PROD_A``). Numeric-looking values then become a stable form so ``1.10``
    and ``1.1`` join as the same key; whole numbers stay without a trailing
    ``.0`` (``20`` not ``20.0``). Other text stays as it is.
    """
    if value is None:
        return ""
    s = unwrap_wrapping_quotes(str(value))
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
    Compare two cell values after dropping one wrapping pair of double quotes.

    Exact string match, or numeric equality when both parse as floats
    (``1.10`` == ``1.1``, ``"PROD_A"`` == ``PROD_A``).
    """
    sa = unwrap_wrapping_quotes("" if a is None else str(a))
    sb = unwrap_wrapping_quotes("" if b is None else str(b))
    if sa == sb:
        return True
    if sa == "" or sb == "":
        return False
    try:
        return float(sa) == float(sb)
    except (TypeError, ValueError):
        return False

