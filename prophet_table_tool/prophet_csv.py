"""Prophet-style CSV / FAC table read/write helpers (Polars-backed)."""

from __future__ import annotations

import io
import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

import polars as pl

# Prophet exports are often named .FAC; change-request snapshots may use .csv.
TABLE_EXTENSIONS = (".csv", ".fac")

_READ_ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
_UTF8_BOM = b"\xef\xbb\xbf"

_HEADER_MARKER_RE = re.compile(rb"^!\d+$")

logger = logging.getLogger(__name__)


@dataclass
class ProphetTable:
    """In-memory representation of a Prophet CSV/FAC table."""

    n_keys: int
    columns: list[str]  # data column headers (excludes the marker column)
    data: pl.DataFrame  # columns == self.columns; all Utf8; no marker col
    source_path: Path | None = None
    # Raw lines before the !N header / after the last * data row (not compared).
    # Stored as bytes so trailing junk with mixed/odd encodings round-trips.
    leading_dummy_lines: list[bytes] = field(default_factory=list)
    trailing_dummy_lines: list[bytes] = field(default_factory=list)
    source_encoding: str | None = None

    @property
    def key_columns(self) -> list[str]:
        """Key columns among the data headers (first n_keys - 1)."""
        return self.columns[: self.n_keys - 1]

    @property
    def value_columns(self) -> list[str]:
        """Non-key data columns."""
        return self.columns[self.n_keys - 1 :]

    @property
    def source_suffix(self) -> str:
        """Preferred output extension from the source file (default ``.csv``)."""
        if self.source_path is not None and self.source_path.suffix:
            return self.source_path.suffix
        return ".csv"

    def with_key_tuple(self) -> pl.DataFrame:
        """Return data with synthetic ``_key_str`` (normalized, pipe-joined)."""
        return with_normalized_keys(self.data, self.key_columns)


def normalized_key_expr(col: str) -> pl.Expr:
    """Polars expression: numeric-looking cells → canonical key text."""
    s = pl.col(col).cast(pl.Utf8).fill_null("")
    num = s.cast(pl.Float64, strict=False)
    is_numeric = num.is_not_null() & (s != "")
    is_int = is_numeric & (num == num.floor()) & (num.abs() < 1e15)
    return (
        pl.when(is_int)
        .then(num.cast(pl.Int64).cast(pl.Utf8))
        .when(is_numeric)
        .then(num.cast(pl.Utf8))
        .otherwise(s)
    )


def key_str_expr(key_cols: list[str]) -> pl.Expr:
    """Pipe-join normalized key columns; empty key list → empty string."""
    if not key_cols:
        return pl.lit("")
    return pl.concat_str([normalized_key_expr(c) for c in key_cols], separator="|")


def with_normalized_keys(df: pl.DataFrame, key_cols: list[str]) -> pl.DataFrame:
    """Add ``_key_str`` (pipe-joined) with numeric key normalization."""
    return df.with_columns(key_str_expr(key_cols).alias("_key_str"))


def read_prophet_csv(path: Path) -> ProphetTable:
    """
    Parse a Prophet-format CSV / FAC file.

    Content block:
      - Starts at the first row whose first cell is ``!N`` (N = key count).
      - Continues through subsequent rows whose first cell is ``*``.

    Any lines before that ``!N`` row and any lines after the last ``*`` data
    row are treated as dummy lines: stored as raw bytes for round-trip
    write-back, and never parsed as CSV (so odd encodings in the trailer
    cannot break the data block).

    Encoding is tried in order on the header + data block only:
    utf-8-sig, utf-8, cp1252, latin-1.
    """
    path = Path(path)
    raw_bytes = path.read_bytes()
    if raw_bytes.startswith(_UTF8_BOM):
        raw_bytes = raw_bytes[len(_UTF8_BOM) :]
    if not raw_bytes:
        raise ValueError(f"Empty CSV: {path}")

    lines = raw_bytes.splitlines()
    header_idx = _find_header_index(lines, path)
    leading_dummy_lines = list(lines[:header_idx])

    data_line_bytes: list[bytes] = []
    trailing_dummy_lines: list[bytes] = []
    pending_blanks: list[bytes] = []
    in_trailing = False

    for line in lines[header_idx + 1 :]:
        if in_trailing:
            trailing_dummy_lines.append(line)
            continue
        if not line.strip():
            pending_blanks.append(line)
            continue
        if _first_field_bytes(line) == b"*":
            pending_blanks.clear()
            data_line_bytes.append(line)
            continue
        in_trailing = True
        trailing_dummy_lines.extend(pending_blanks)
        pending_blanks.clear()
        trailing_dummy_lines.append(line)

    trailing_dummy_lines.extend(pending_blanks)

    block = b"\n".join([lines[header_idx], *data_line_bytes])
    text, encoding = _decode_with_fallback(block, path)
    decoded = text.split("\n")
    header_text = decoded[0] if decoded else ""
    data_text_lines = decoded[1:] if data_line_bytes else []

    header_cells = _split_csv_line(header_text)
    try:
        n_keys = int(header_cells[0][1:])
    except (ValueError, IndexError) as exc:
        marker = header_cells[0] if header_cells else ""
        raise ValueError(f"Invalid key-count marker in {path}: {marker!r}") from exc

    columns = header_cells[1:]
    if n_keys < 1:
        raise ValueError(f"n_keys must be >= 1 in {path}, got {n_keys}")
    if n_keys - 1 > len(columns):
        raise ValueError(
            f"n_keys={n_keys} exceeds available columns ({len(columns)}) in {path}"
        )

    data = _read_data_block(columns, data_text_lines)

    return ProphetTable(
        n_keys=n_keys,
        columns=columns,
        data=data,
        source_path=path,
        leading_dummy_lines=leading_dummy_lines,
        trailing_dummy_lines=trailing_dummy_lines,
        source_encoding=encoding,
    )


def write_prophet_csv(table: ProphetTable, path: Path) -> None:
    """Write a ProphetTable preserving dummy bytes and ``!N`` / ``*`` format."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    chunks: list[bytes] = []
    for line in table.leading_dummy_lines:
        chunks.append(line)
        chunks.append(b"\n")

    header = ",".join([f"!{table.n_keys}", *table.columns])
    chunks.append(header.encode("utf-8"))
    chunks.append(b"\n")

    if table.data.height:
        starred = table.data.select(table.columns).with_columns(
            pl.lit("*").alias("_marker")
        ).select(["_marker", *table.columns])
        buf = io.BytesIO()
        starred.write_csv(
            buf,
            include_header=False,
            quote_style="never",
            null_value="",
        )
        body = buf.getvalue()
        if body:
            chunks.append(body)
            if not body.endswith(b"\n"):
                chunks.append(b"\n")

    for line in table.trailing_dummy_lines:
        chunks.append(line)
        chunks.append(b"\n")

    path.write_bytes(b"".join(chunks))


def discover_csv_tables(folder: Path) -> dict[str, Path]:
    """
    Map table identity -> path for Prophet table files under *folder* (recursive).

    Identity is the posix-relative path without suffix (e.g. ``SubA/MORT_TABLE``).
    Recognizes ``*.csv`` and ``*.FAC`` / ``*.fac``. If both exist for the same
    identity, ``.FAC`` wins (native Prophet export name).
    """
    folder = Path(folder)
    if not folder.is_dir():
        return {}

    found: dict[str, Path] = {}
    for p in sorted(folder.rglob("*")):
        if not p.is_file() or p.suffix.lower() not in TABLE_EXTENSIONS:
            continue
        rel = p.relative_to(folder)
        identity = rel.with_suffix("").as_posix()
        existing = found.get(identity)
        if existing is None:
            found[identity] = p
        elif existing.suffix.lower() == ".csv" and p.suffix.lower() == ".fac":
            found[identity] = p
    return found


def _read_data_block(columns: list[str], data_lines: list[str]) -> pl.DataFrame:
    """Parse ``*`` data lines with Polars; dummy lines never reach here."""
    if not columns:
        return pl.DataFrame()
    if not data_lines:
        return pl.DataFrame({c: [] for c in columns}).cast(pl.Utf8)

    csv_text = ",".join(["_marker", *columns]) + "\n" + "\n".join(data_lines)
    schema = {"_marker": pl.Utf8, **{c: pl.Utf8 for c in columns}}
    df = pl.read_csv(
        io.BytesIO(csv_text.encode("utf-8")),
        has_header=True,
        schema=schema,
        truncate_ragged_lines=True,
        quote_char=None,
    )
    if "_marker" in df.columns:
        df = df.drop("_marker")
    for c in columns:
        if c not in df.columns:
            df = df.with_columns(pl.lit("").cast(pl.Utf8).alias(c))
    return df.select(columns).cast(pl.Utf8).fill_null("")


def _decode_with_fallback(data: bytes, path: Path) -> tuple[str, str]:
    last_error: Exception | None = None
    for encoding in _READ_ENCODINGS:
        try:
            text = data.decode(encoding)
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
        logger.debug("read %s data block using encoding=%s", path, encoding)
        return text, encoding
    raise ValueError(
        f"Unable to decode {path} with encodings {_READ_ENCODINGS}: {last_error}"
    )


def _find_header_index(lines: list[bytes], path: Path) -> int:
    for i, line in enumerate(lines):
        if not line.strip():
            continue
        if _HEADER_MARKER_RE.match(_first_field_bytes(line)):
            return i
    raise ValueError(
        f"Invalid Prophet table in {path}: no !N header row found "
        f"(content must start with a !N marker row)"
    )


def _first_field_bytes(line: bytes) -> bytes:
    return line.split(b",", 1)[0].strip()


def _split_csv_line(line: str) -> list[str]:
    """Simple comma-split (Prophet tables are plain CSV without quoted commas)."""
    return [c.strip() for c in line.rstrip("\r\n").split(",")]
