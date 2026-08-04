"""Prophet-style CSV / FAC table read/write helpers (Polars-backed)."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import polars as pl

# Prophet exports are often named .FAC; change-request snapshots may use .csv.
TABLE_EXTENSIONS = (".csv", ".fac")

_HEADER_MARKER_RE = re.compile(r"^!\d+$")


@dataclass
class ProphetTable:
    """In-memory representation of a Prophet CSV/FAC table."""

    n_keys: int
    columns: list[str]  # data column headers (excludes the marker column)
    data: pl.DataFrame  # columns == self.columns; all Utf8; no marker col
    source_path: Path | None = None
    # Raw lines before the !N header / after the last * data row (not compared).
    leading_dummy_lines: list[str] = field(default_factory=list)
    trailing_dummy_lines: list[str] = field(default_factory=list)

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
        """Return data with a synthetic `_key_tuple` list column for matching."""
        key_cols = self.key_columns
        if not key_cols:
            # Only the marker column is a key (n_keys == 1) — rare but valid.
            return self.data.with_columns(
                pl.lit([]).cast(pl.List(pl.Utf8)).alias("_key_tuple")
            )
        return self.data.with_columns(
            pl.concat_list([pl.col(c).cast(pl.Utf8) for c in key_cols]).alias(
                "_key_tuple"
            )
        )


def read_prophet_csv(path: Path) -> ProphetTable:
    """
    Parse a Prophet-format CSV / FAC file.

    Content block:
      - Starts at the first row whose first cell is ``!N`` (N = key count).
      - Continues through subsequent rows whose first cell is ``*``.

    Any lines before that ``!N`` row and any lines after the last ``*`` data
    row are treated as dummy lines: stored for round-trip write-back, but not
    included in ``data`` (so Stage 1 Change Log comparison ignores them).
    """
    path = Path(path)
    raw = path.read_text(encoding="utf-8-sig").splitlines()
    if not raw:
        raise ValueError(f"Empty CSV: {path}")

    header_idx = _find_header_index(raw, path)
    leading_dummy_lines = list(raw[:header_idx])

    header_cells = _split_csv_line(raw[header_idx])
    try:
        n_keys = int(header_cells[0][1:])
    except ValueError as exc:
        raise ValueError(f"Invalid key-count marker in {path}: {header_cells[0]!r}") from exc

    columns = header_cells[1:]
    if n_keys < 1:
        raise ValueError(f"n_keys must be >= 1 in {path}, got {n_keys}")
    if n_keys - 1 > len(columns):
        raise ValueError(
            f"n_keys={n_keys} exceeds available columns ({len(columns)}) in {path}"
        )

    rows: list[list[str]] = []
    trailing_dummy_lines: list[str] = []
    pending_blanks: list[str] = []
    in_trailing = False

    for i, line in enumerate(raw[header_idx + 1 :], start=header_idx + 2):
        if in_trailing:
            trailing_dummy_lines.append(line)
            continue

        if not line.strip():
            # Defer blanks: they belong to trailing dummies if no further * row.
            pending_blanks.append(line)
            continue

        cells = _split_csv_line(line)
        if cells and cells[0] == "*":
            pending_blanks.clear()
            values = cells[1:]
            # Pad / trim to header width for robustness
            if len(values) < len(columns):
                values = values + [""] * (len(columns) - len(values))
            elif len(values) > len(columns):
                values = values[: len(columns)]
            rows.append(values)
            continue

        # First non-blank, non-* row after the header → trailing dummies.
        in_trailing = True
        trailing_dummy_lines.extend(pending_blanks)
        pending_blanks.clear()
        trailing_dummy_lines.append(line)

    # Trailing blank lines with no following dummy text still count as trailing.
    trailing_dummy_lines.extend(pending_blanks)

    if rows:
        data = pl.DataFrame(rows, schema=columns, orient="row").cast(pl.Utf8)
    else:
        data = pl.DataFrame({c: [] for c in columns}).cast(pl.Utf8)

    return ProphetTable(
        n_keys=n_keys,
        columns=columns,
        data=data,
        source_path=path,
        leading_dummy_lines=leading_dummy_lines,
        trailing_dummy_lines=trailing_dummy_lines,
    )


def write_prophet_csv(table: ProphetTable, path: Path) -> None:
    """Write a ProphetTable preserving dummy lines and ``!N`` / ``*`` format."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    lines: list[str] = []
    lines.extend(table.leading_dummy_lines)

    header = [f"!{table.n_keys}"] + list(table.columns)
    lines.append(",".join(header))

    for row in table.data.select(table.columns).iter_rows():
        cells = ["*"] + ["" if v is None else str(v) for v in row]
        lines.append(",".join(cells))

    lines.extend(table.trailing_dummy_lines)

    path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")


def discover_csv_tables(folder: Path) -> dict[str, Path]:
    """
    Map table_name -> path for Prophet table files in *folder*.

    Recognizes ``*.csv`` and ``*.FAC`` / ``*.fac``. If both exist for the same
    stem, ``.FAC`` wins (native Prophet export name).
    """
    folder = Path(folder)
    if not folder.is_dir():
        return {}

    found: dict[str, Path] = {}
    for p in sorted(folder.iterdir()):
        if not p.is_file() or p.suffix.lower() not in TABLE_EXTENSIONS:
            continue
        stem = p.stem
        existing = found.get(stem)
        if existing is None:
            found[stem] = p
        elif existing.suffix.lower() == ".csv" and p.suffix.lower() == ".fac":
            found[stem] = p
    return found


def _find_header_index(raw: list[str], path: Path) -> int:
    for i, line in enumerate(raw):
        if not line.strip():
            continue
        cells = _split_csv_line(line)
        if cells and _HEADER_MARKER_RE.match(cells[0]):
            return i
    raise ValueError(
        f"Invalid Prophet table in {path}: no !N header row found "
        f"(content must start with a !N marker row)"
    )


def _split_csv_line(line: str) -> list[str]:
    """Simple comma-split (Prophet tables are plain CSV without quoted commas)."""
    return [c.strip() for c in line.rstrip("\r\n").split(",")]
