"""Human-friendly per-table review sheets for the Stage 1 Change Log."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import polars as pl
import xlsxwriter

from .diff import ChangeRow, STRUCTURAL_TYPES, empty_cell_changes, rows_to_cell_df
from .prophet_csv import ProphetTable, with_normalized_keys

RESERVED_SHEET_NAMES = frozenset(
    {"Summary", "ChangeLog_Detail", "Conflicts", "ReviewFiles"}
)

ROOT_GROUP = "root"

_CHANGE_COL = "_change"

_INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")
_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')


def _changed_keys_from_cells(cell_changes: pl.DataFrame) -> set[str]:
    """Keys that appear in row_add / row_delete / value_update (for sparse reviews)."""
    if cell_changes.is_empty():
        return set()
    keys = (
        cell_changes.filter(
            pl.col("change_type").is_in(["row_add", "row_delete", "value_update"])
            & (pl.col("key_tuple") != "")
        )
        .get_column("key_tuple")
        .unique()
        .to_list()
    )
    return {str(k) for k in keys}


@dataclass
class TableContribution:
    change_request_id: str
    before: ProphetTable | None
    after: ProphetTable | None
    structural_rows: list[ChangeRow] = field(default_factory=list)
    cell_changes: pl.DataFrame = field(default_factory=empty_cell_changes)


@dataclass
class TableReviewModel:
    table_name: str
    contributions: list[TableContribution] = field(default_factory=list)


@dataclass
class _CellAnnotation:
    old_value: str
    new_value: str
    change_request_id: str


@dataclass
class _ReviewRow:
    key_str: str
    status: str  # "", "ADD", "DELETE"
    values: dict[str, str]
    cell_updates: dict[str, list[_CellAnnotation]] = field(default_factory=dict)
    source_crs: list[str] = field(default_factory=list)


@dataclass
class ReviewGrid:
    table_name: str
    columns: list[str]
    structural_notes: list[str]
    rows: list[_ReviewRow]
    multi_cr: bool


@dataclass
class _ReviewFormats:
    bold: Any
    header: Any
    note: Any
    add: Any
    delete: Any
    update: Any


def add_contribution(
    reviews: dict[str, TableReviewModel],
    table_name: str,
    change_request_id: str,
    before: ProphetTable | None,
    after: ProphetTable | None,
    change_rows: list[ChangeRow],
    cell_changes: pl.DataFrame | None = None,
) -> None:
    """Accumulate one CR's before/after snapshot for a table into the review map."""
    structural = [r for r in change_rows if r.table_name == table_name and r.change_type in STRUCTURAL_TYPES]
    if cell_changes is None:
        cells = rows_to_cell_df(
            [r for r in change_rows if r.table_name == table_name and r.change_type in {"value_update", "row_add", "row_delete"}]
        )
    else:
        cells = cell_changes
        if cells.height and "table_name" in cells.columns:
            cells = cells.filter(pl.col("table_name") == table_name)

    if not structural and cells.is_empty() and before is None and after is None:
        return

    wanted = _changed_keys_from_cells(cells)
    if table_name not in reviews:
        reviews[table_name] = TableReviewModel(table_name=table_name)
    reviews[table_name].contributions.append(
        TableContribution(
            change_request_id=change_request_id,
            before=_slice_table(before, wanted),
            after=_slice_table(after, wanted),
            structural_rows=structural,
            cell_changes=cells,
        )
    )


def _slice_table(table: ProphetTable | None, wanted: set[str]) -> ProphetTable | None:
    """Keep schema plus changed keys only (full table is not needed for review)."""
    if table is None:
        return None
    if not wanted:
        return ProphetTable(
            n_keys=table.n_keys,
            columns=list(table.columns),
            data=table.data.head(0),
            source_path=table.source_path,
            source_encoding=table.source_encoding,
        )
    keyed = with_normalized_keys(table.data, table.key_columns)
    sliced = keyed.filter(pl.col("_key_str").is_in(list(wanted))).drop("_key_str")
    return ProphetTable(
        n_keys=table.n_keys,
        columns=list(table.columns),
        data=sliced.select(table.columns),
        source_path=table.source_path,
        source_encoding=table.source_encoding,
    )


def review_group_for_table(table_name: str) -> str:
    """First path segment of a table identity, or ``root`` if it has no folder."""
    name = str(table_name).replace("\\", "/").strip("/")
    if not name or "/" not in name:
        return ROOT_GROUP
    return name.split("/", 1)[0]


def sanitize_group_filename(group: str) -> str:
    """Make a first-level group name safe as an ``.xlsx`` stem."""
    cleaned = _INVALID_FILENAME_CHARS.sub("_", group).strip(" .") or ROOT_GROUP
    return cleaned


def group_table_reviews(
    reviews: dict[str, TableReviewModel],
) -> dict[str, dict[str, TableReviewModel]]:
    """Partition review models by first-level production-table folder."""
    grouped: dict[str, dict[str, TableReviewModel]] = {}
    for table_name, model in reviews.items():
        grouped.setdefault(review_group_for_table(table_name), {})[table_name] = model
    return grouped


def build_review_grid(model: TableReviewModel) -> ReviewGrid | None:
    """Build a wide Prophet-shaped review grid from accumulated CR contributions."""
    if not model.contributions:
        return None

    cr_ids = {c.change_request_id for c in model.contributions}
    multi_cr = len(cr_ids) > 1

    columns = _resolve_columns(model)
    if not columns:
        return None

    structural_notes = _structural_notes(model)
    row_map = _seed_rows(model, columns)
    _apply_change_annotations(model, row_map, columns)

    # Stable order: ADD rows, then unchanged/updated, then DELETE
    status_rank = {"": 1, "ADD": 0, "DELETE": 2}
    rows = sorted(
        row_map.values(),
        key=lambda r: (status_rank.get(r.status, 1), r.key_str),
    )
    return ReviewGrid(
        table_name=model.table_name,
        columns=columns,
        structural_notes=structural_notes,
        rows=rows,
        multi_cr=multi_cr,
    )


def write_review_workbook(
    path: Path,
    reviews: dict[str, TableReviewModel],
    *,
    conflict_headers: list[str] | None = None,
    conflict_rows: list[list] | None = None,
) -> None:
    """Write Conflicts (optional) plus one review sheet per table via xlsxwriter."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(str(path), {"constant_memory": True})
    try:
        formats = _make_formats(wb)
        used_names: set[str] = set(RESERVED_SHEET_NAMES)
        if conflict_headers is not None:
            _write_conflicts_sheet(
                wb, formats, conflict_headers, conflict_rows or []
            )
        for table_name in sorted(reviews):
            grid = build_review_grid(reviews[table_name])
            if grid is None:
                continue
            sheet_name = _unique_sheet_name(table_name, used_names)
            used_names.add(sheet_name)
            ws = wb.add_worksheet(sheet_name)
            _write_review_sheet(ws, grid, formats)
    finally:
        wb.close()


def write_review_sheets(
    path: Path,
    reviews: dict[str, TableReviewModel],
) -> None:
    """Write per-table review sheets to a standalone workbook (no Conflicts tab)."""
    write_review_workbook(path, reviews)


def _make_formats(wb: xlsxwriter.Workbook) -> _ReviewFormats:
    return _ReviewFormats(
        bold=wb.add_format({"bold": True}),
        header=wb.add_format({"bold": True, "bg_color": "D9D9D9"}),
        note=wb.add_format({"bg_color": "DDEBF7"}),
        add=wb.add_format({"bg_color": "C6EFCE"}),
        delete=wb.add_format({"bg_color": "FFC7CE"}),
        update=wb.add_format({"bg_color": "FFF2CC"}),
    )


def _write_conflicts_sheet(
    wb: xlsxwriter.Workbook,
    formats: _ReviewFormats,
    headers: list[str],
    rows: list[list],
) -> None:
    ws = wb.add_worksheet("Conflicts")
    for j, header in enumerate(headers):
        ws.write(0, j, header, formats.header)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            ws.write(i, j, val)


def _resolve_columns(model: TableReviewModel) -> list[str]:
    columns: list[str] = []
    for contrib in model.contributions:
        src = contrib.after or contrib.before
        if src is None:
            continue
        for col in src.columns:
            if col not in columns:
                columns.append(col)
    # Also pick up columns mentioned only in change rows (e.g. column_add with empty table)
    extra: list[str] = []
    for contrib in model.contributions:
        for row in contrib.structural_rows:
            if row.change_type == "column_add" and row.column_name and row.column_name not in columns:
                extra.append(row.column_name)
        if contrib.cell_changes.height:
            for col in contrib.cell_changes.get_column("column_name").unique().to_list():
                name = str(col or "")
                if name and name not in columns and "->" not in name:
                    extra.append(name)
    for col in extra:
        if col not in columns:
            columns.append(col)
    return columns


def _structural_notes(model: TableReviewModel) -> list[str]:
    notes: list[str] = []
    seen: set[str] = set()
    for contrib in model.contributions:
        cr = contrib.change_request_id
        for row in contrib.structural_rows:
            text = _format_structural_note(row, cr)
            if text not in seen:
                seen.add(text)
                notes.append(text)
    return notes


def _format_structural_note(row: ChangeRow, cr_id: str) -> str:
    prefix = f"[{cr_id}] "
    if row.change_type == "table_add":
        return f"{prefix}table_add - {row.notes or 'Table only present in after/'}"
    if row.change_type == "key_count_change":
        return (
            f"{prefix}key_count_change - {row.old_value} -> {row.new_value}"
            + (f" ({row.notes})" if row.notes else "")
        )
    if row.change_type == "column_rename":
        return f"{prefix}column_rename - {row.old_value} -> {row.new_value}"
    if row.change_type == "column_add":
        return f"{prefix}column_add - {row.column_name}"
    if row.change_type == "column_delete":
        return f"{prefix}column_delete - {row.column_name}"
    return f"{prefix}{row.change_type} - {row.column_name or row.notes}"


def _table_row_dict(
    table: ProphetTable,
    wanted: set[str] | None = None,
) -> dict[str, dict[str, str]]:
    """Map key_str → {col: value}, optionally restricted to ``wanted`` keys."""
    if wanted is not None and not wanted:
        return {}
    keyed = with_normalized_keys(table.data, table.key_columns)
    if wanted is not None:
        keyed = keyed.filter(pl.col("_key_str").is_in(list(wanted)))
    out: dict[str, dict[str, str]] = {}
    for row in keyed.iter_rows(named=True):
        key_str = row["_key_str"]
        values = {
            col: "" if row[col] is None else str(row[col]) for col in table.columns
        }
        out[key_str] = values
    return out


def _seed_rows(
    model: TableReviewModel,
    columns: list[str],
) -> dict[str, _ReviewRow]:
    """
    Seed review rows for changed keys only (sparse grid for large tables).

    Prefer after values; fill gaps from before. Keys are taken from change rows
    (row_add / row_delete / value_update), not the full table.
    """
    row_map: dict[str, _ReviewRow] = {}
    wanted: set[str] = set()
    for contrib in model.contributions:
        wanted |= _changed_keys_from_cells(contrib.cell_changes)

    # Prefer after values for wanted keys
    for contrib in model.contributions:
        if contrib.after is None:
            continue
        for key_str, values in _table_row_dict(contrib.after, wanted).items():
            if key_str not in row_map:
                row_map[key_str] = _ReviewRow(
                    key_str=key_str,
                    status="",
                    values={c: values.get(c, "") for c in columns},
                )
            else:
                for c in columns:
                    if c in values:
                        row_map[key_str].values[c] = values[c]

    # Before values for wanted keys still missing (or DELETE candidates)
    for contrib in model.contributions:
        if contrib.before is None:
            continue
        for key_str, values in _table_row_dict(contrib.before, wanted).items():
            if key_str not in row_map:
                row_map[key_str] = _ReviewRow(
                    key_str=key_str,
                    status="",
                    values={c: values.get(c, "") for c in columns},
                )

    return row_map


def _record_source_cr(review_row: _ReviewRow, cr: str) -> None:
    if cr and cr not in review_row.source_crs:
        review_row.source_crs.append(cr)


def _format_change_cell(review_row: _ReviewRow) -> str:
    """ADD/DELETE prefix plus contributing [change_request_id] values."""
    cr_part = " | ".join(f"[{cr}]" for cr in review_row.source_crs)
    if review_row.status in {"ADD", "DELETE"}:
        if cr_part:
            return f"{review_row.status} {cr_part}"
        return review_row.status
    return cr_part


def _apply_change_annotations(
    model: TableReviewModel,
    row_map: dict[str, _ReviewRow],
    columns: list[str],
) -> None:
    col_set = set(columns)
    for contrib in model.contributions:
        cr = contrib.change_request_id
        cells = contrib.cell_changes
        if cells.is_empty():
            continue

        add_vals: dict[str, dict[str, str]] = {}
        delete_vals: dict[str, dict[str, str]] = {}
        updates: list[tuple[str, str, str, str]] = []

        for rec in cells.select(
            ["change_type", "key_tuple", "column_name", "old_value", "new_value"]
        ).iter_rows(named=True):
            key_str = str(rec["key_tuple"] or "")
            col = str(rec["column_name"] or "")
            ct = rec["change_type"]
            if ct == "row_add" and key_str:
                add_vals.setdefault(key_str, {})
                if col:
                    add_vals[key_str][col] = "" if rec["new_value"] is None else str(rec["new_value"])
            elif ct == "row_delete" and key_str:
                delete_vals.setdefault(key_str, {})
                if col:
                    delete_vals[key_str][col] = "" if rec["old_value"] is None else str(rec["old_value"])
            elif ct == "value_update" and key_str and col:
                updates.append(
                    (
                        key_str,
                        col,
                        "" if rec["old_value"] is None else str(rec["old_value"]),
                        "" if rec["new_value"] is None else str(rec["new_value"]),
                    )
                )

        for key_str, col_vals in add_vals.items():
            review_row = row_map.get(key_str)
            if review_row is None:
                values = {c: col_vals.get(c, "") for c in columns}
                review_row = _ReviewRow(key_str=key_str, status="ADD", values=values)
                row_map[key_str] = review_row
            else:
                if review_row.status != "DELETE":
                    review_row.status = "ADD"
                for c, v in col_vals.items():
                    if c in review_row.values:
                        review_row.values[c] = v
            _record_source_cr(review_row, cr)

        for key_str, col_vals in delete_vals.items():
            review_row = row_map.get(key_str)
            if review_row is None:
                values = {c: col_vals.get(c, "") for c in columns}
                review_row = _ReviewRow(key_str=key_str, status="DELETE", values=values)
                row_map[key_str] = review_row
            else:
                review_row.status = "DELETE"
                for c, v in col_vals.items():
                    if c in review_row.values:
                        review_row.values[c] = v
            _record_source_cr(review_row, cr)

        for key_str, col, old_v, new_v in updates:
            if col not in col_set:
                continue
            review_row = row_map.get(key_str)
            if review_row is None:
                review_row = _ReviewRow(
                    key_str=key_str,
                    status="",
                    values={c: "" for c in columns},
                )
                row_map[key_str] = review_row
            _record_source_cr(review_row, cr)
            review_row.cell_updates.setdefault(col, []).append(
                _CellAnnotation(
                    old_value=old_v,
                    new_value=new_v,
                    change_request_id=cr,
                )
            )
            review_row.values[col] = new_v


def _display_cell(review_row: _ReviewRow, column: str, multi_cr: bool) -> str:
    updates = review_row.cell_updates.get(column)
    if not updates:
        return review_row.values.get(column, "")
    parts: list[str] = []
    for u in updates:
        text = f"{u.old_value} -> {u.new_value}"
        if multi_cr:
            text = f"{text} [{u.change_request_id}]"
        parts.append(text)
    return " | ".join(parts)


def _sanitize_sheet_name(name: str) -> str:
    cleaned = _INVALID_SHEET_CHARS.sub("_", name).strip() or "Table"
    # Excel sheet names cannot start/end with apostrophe in a problematic way;
    # also cap at 31 chars.
    return cleaned[:31]


def _unique_sheet_name(table_name: str, used: set[str]) -> str:
    base = _sanitize_sheet_name(table_name)
    if base not in used and base not in RESERVED_SHEET_NAMES:
        return base
    # Truncate to leave room for suffix _N
    for i in range(2, 1000):
        suffix = f"_{i}"
        candidate = f"{base[: 31 - len(suffix)]}{suffix}"
        if candidate not in used and candidate not in RESERVED_SHEET_NAMES:
            return candidate
    raise ValueError(f"Unable to allocate unique sheet name for table {table_name!r}")


def _write_review_sheet(ws: Any, grid: ReviewGrid, formats: _ReviewFormats) -> None:
    row_idx = 0

    ws.write(
        row_idx,
        0,
        f"Review: {grid.table_name} (human review only - not used by apply)",
        formats.bold,
    )
    row_idx += 1

    if grid.structural_notes:
        ws.write(row_idx, 0, "Structural changes:", formats.bold)
        row_idx += 1
        for note in grid.structural_notes:
            ws.write(row_idx, 0, note, formats.note)
            row_idx += 1

    # Blank spacer when we had notes or title
    row_idx += 1

    header_row = row_idx
    # Freeze below the header and after the _change column (set before data for
    # xlsxwriter constant_memory).
    ws.freeze_panes(header_row + 1, 1)

    headers = [_CHANGE_COL, *grid.columns]
    for j, header in enumerate(headers):
        ws.write(header_row, j, header, formats.header)
    row_idx = header_row + 1

    for review_row in grid.rows:
        status = review_row.status
        if status == "ADD":
            row_fmt = formats.add
        elif status == "DELETE":
            row_fmt = formats.delete
        else:
            row_fmt = None

        ws.write(row_idx, 0, _format_change_cell(review_row), row_fmt)

        for j, col in enumerate(grid.columns, start=1):
            value = _display_cell(review_row, col, grid.multi_cr)
            cell_fmt = row_fmt
            if cell_fmt is None and col in review_row.cell_updates:
                cell_fmt = formats.update
            ws.write(row_idx, j, value, cell_fmt)

        row_idx += 1
