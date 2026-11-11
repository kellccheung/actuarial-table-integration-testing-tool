"""External paths, key-column pairs, structure scan, and spilled change logs."""

from __future__ import annotations

import csv
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from prophet_table_tool.changelog import (
    StructureScanError,
    generate_change_log,
    run_structure_scan,
    scan_structure,
)
from prophet_table_tool.control import (
    ControlConfig,
    KeyCountApproval,
    read_control,
)
from prophet_table_tool.diff import diff_table
from prophet_table_tool.prophet_csv import read_prophet_csv, read_prophet_header


def _write_control(
    path: Path,
    *,
    working_root: Path,
    change_requests: list[tuple],
    key_rows: list[tuple] | None = None,
    run_id: str = "STAGE1",
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Config"
    ws.append(["key", "value"])
    ws.append(["working_root", str(working_root)])
    ws.append(["mode", "generate_changelog"])
    ws.append(["production_tables_path", "Production_Tables"])
    ws.append(["output_path", "Output"])
    ws.append(["run_id", run_id])
    ws_cr = wb.create_sheet("ChangeRequests")
    ws_cr.append(
        [
            "change_request_id",
            "order",
            "include",
            "approved",
            "description",
            "notes",
            "after_path",
            "before_source",
            "before_path",
        ]
    )
    for row in change_requests:
        ws_cr.append(list(row))
    ws_kc = wb.create_sheet("KeyCountApprovals")
    ws_kc.append(
        ["change_request_id", "table_name", "approved", "before_column", "after_column"]
    )
    for row in key_rows or []:
        ws_kc.append(list(row))
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _write_table(path: Path, n_keys: int, columns: list[str], rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [",".join([f"!{n_keys}", *columns])]
    for row in rows:
        lines.append(",".join(["*", *row]))
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _detail(path: Path) -> list[dict]:
    with path.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def test_read_prophet_header_ignores_body(tmp_path: Path):
    path = tmp_path / "T.csv"
    path.write_bytes(b"dummy\n!2,Age,Rate\n*\xff\xfe not utf8\n")
    n_keys, columns = read_prophet_header(path)
    assert n_keys == 2
    assert columns == ["Age", "Rate"]


def test_key_approval_requires_every_row():
    control = ControlConfig(
        control_path=Path("Control.xlsx"),
        working_root=Path("."),
        mode="apply",
        production_tables_path=Path("Production_Tables"),
        output_path=Path("Output"),
        run_id="T",
        key_count_approvals=[
            KeyCountApproval("CR", "MORT_TABLE", True, "Age", "Age"),
            KeyCountApproval("CR", "MORT_TABLE", False, "Duration", "Term"),
        ],
    )
    assert control.is_key_count_approved("CR", "MORT_TABLE") is False


def test_key_name_mismatch_is_fatal_until_declared(tmp_path: Path):
    before = tmp_path / "ChangeRequests" / "CR_KEY" / "before"
    after = tmp_path / "ChangeRequests" / "CR_KEY" / "after"
    _write_table(before / "MORT_TABLE.csv", 3, ["Age", "Duration", "Rate"], [["20", "1", "0.1"]])
    _write_table(after / "MORT_TABLE.csv", 3, ["Age", "Term", "Rate"], [["20", "1", "0.2"]])
    (tmp_path / "Production_Tables").mkdir()
    control_path = tmp_path / "Control.xlsx"
    _write_control(
        control_path,
        working_root=tmp_path,
        change_requests=[("CR_KEY", 1, "Y", "Y", "rename key", "")],
    )
    findings = scan_structure(read_control(control_path))
    assert any(
        row["severity"] == "FATAL" and "KeyCountApprovals" in row["message"] for row in findings
    )
    with pytest.raises(StructureScanError):
        generate_change_log(control_path)

    before_tbl = read_prophet_csv(before / "MORT_TABLE.csv")
    after_tbl = read_prophet_csv(after / "MORT_TABLE.csv")
    with pytest.raises(ValueError, match="KeyCountApprovals"):
        diff_table("CR_KEY", "MORT_TABLE", before_tbl, after_tbl, read_control(control_path))

    _write_control(
        control_path,
        working_root=tmp_path,
        change_requests=[("CR_KEY", 1, "Y", "Y", "rename key", "")],
        key_rows=[
            ("CR_KEY", "MORT_TABLE", "Y", "Age", "Age"),
            ("CR_KEY", "MORT_TABLE", "Y", "Duration", "Term"),
        ],
    )
    clog = generate_change_log(control_path)
    rows = _detail(clog.with_name(f"{clog.stem}_Detail.csv"))
    renames = [row for row in rows if row["change_type"] == "column_rename"]
    updates = [row for row in rows if row["change_type"] == "value_update"]
    assert renames[0]["old_value"] == "Duration"
    assert renames[0]["new_value"] == "Term"
    assert updates[0]["key_tuple"] == "20|1"
    assert updates[0]["column_name"] == "Rate"
    assert not (tmp_path / "Output" / ".stage1_STAGE1").exists()


def test_production_before_ignores_other_tables(tmp_path: Path):
    _write_table(
        tmp_path / "Production_Tables" / "MORT_TABLE.csv",
        3,
        ["Age", "Duration", "Rate"],
        [["20", "1", "0.1"]],
    )
    _write_table(
        tmp_path / "Production_Tables" / "EXTRA.csv",
        2,
        ["Code", "Value"],
        [["A", "1"]],
    )
    after = tmp_path / "incoming" / "MORT_TABLE.csv"
    _write_table(after, 3, ["Age", "Duration", "Rate"], [["20", "1", "0.2"]])
    control_path = tmp_path / "Control.xlsx"
    _write_control(
        control_path,
        working_root=tmp_path,
        change_requests=[
            ("CR_PROD", 1, "Y", "Y", "from production", "", str(after.parent), "production", ""),
        ],
    )
    clog = generate_change_log(control_path)
    rows = _detail(clog.with_name(f"{clog.stem}_Detail.csv"))
    assert {row["table_name"] for row in rows} == {"MORT_TABLE"}
    assert any(row["change_type"] == "value_update" for row in rows)


def test_two_requests_spill_still_reports_overlap(tmp_path: Path):
    for cr_id, rate in (("CR_A", "0.2"), ("CR_B", "0.3")):
        _write_table(
            tmp_path / "ChangeRequests" / cr_id / "before" / "MORT_TABLE.csv",
            3,
            ["Age", "Duration", "Rate"],
            [["20", "1", "0.1"]],
        )
        _write_table(
            tmp_path / "ChangeRequests" / cr_id / "after" / "MORT_TABLE.csv",
            3,
            ["Age", "Duration", "Rate"],
            [["20", "1", rate]],
        )
    (tmp_path / "Production_Tables").mkdir()
    control_path = tmp_path / "Control.xlsx"
    _write_control(
        control_path,
        working_root=tmp_path,
        change_requests=[
            ("CR_A", 1, "Y", "Y", "a", ""),
            ("CR_B", 2, "Y", "Y", "b", ""),
        ],
    )
    clog = generate_change_log(control_path)
    rows = _detail(clog.with_name(f"{clog.stem}_Detail.csv"))
    assert {row["change_request_id"] for row in rows if row["change_type"] == "value_update"} == {
        "CR_A",
        "CR_B",
    }
    wb = load_workbook(clog, data_only=True, read_only=True)
    conflicts = list(wb["Conflicts"].iter_rows(values_only=True))
    wb.close()
    assert any(row and row[0] == "cell_overlap" for row in conflicts[1:])
    review = tmp_path / "Output" / "ChangeLog_STAGE1_reviews" / "root.xlsx"
    assert review.is_file()
    review_wb = load_workbook(review, data_only=True, read_only=True)
    assert "MORT_TABLE" in review_wb.sheetnames
    review_wb.close()
    assert not (tmp_path / "Output" / ".stage1_STAGE1").exists()


def test_scan_reports_duplicate_index_key(tmp_path: Path):
    before = tmp_path / "ChangeRequests" / "CR_DUP" / "before"
    after = tmp_path / "ChangeRequests" / "CR_DUP" / "after"
    _write_table(
        before / "MORT_TABLE.csv",
        2,
        ["Age", "Rate"],
        [["20", "0.1"], ["21", "0.2"]],
    )
    _write_table(
        after / "MORT_TABLE.csv",
        2,
        ["Age", "Rate"],
        [["20", "0.1"], ["20", "0.3"]],
    )
    _write_table(before / "DUP_COL.csv", 2, ["Age", "Age"], [["20", "0.1"]])
    _write_table(after / "DUP_COL.csv", 2, ["Age", "Age"], [["20", "0.2"]])
    (tmp_path / "Production_Tables").mkdir()
    control_path = tmp_path / "Control.xlsx"
    _write_control(
        control_path,
        working_root=tmp_path,
        change_requests=[("CR_DUP", 1, "Y", "Y", "dup key", "")],
        run_id="SCAN",
    )
    control = read_control(control_path)
    gate = scan_structure(control)
    assert not any("duplicate key" in row["message"] for row in gate)

    scan_path, findings = run_structure_scan(control)
    assert scan_path.is_file()
    log_text = (tmp_path / "Output" / "Audit" / "SCAN_scan.log").read_text(encoding="utf-8")
    assert "scanning MORT_TABLE" in log_text
    assert "finished MORT_TABLE" in log_text
    assert any(
        row["severity"] == "FATAL" and "duplicate key" in row["message"]
        for row in findings
    )
    assert any(
        row["table_name"] == "DUP_COL" and "duplicate column names" in row["message"]
        for row in findings
    )
