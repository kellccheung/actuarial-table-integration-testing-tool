"""Tests for CLI Change Log resolution."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from prophet_table_tool.__main__ import ChangeLogResolveError, _resolve_change_log, main
from prophet_table_tool.changelog import generate_change_log
from prophet_table_tool.control import ControlConfig
from test_sequential_apply import _cr_dirs, _write_control, _write_table


def _control(tmp_path: Path) -> ControlConfig:
    return ControlConfig(
        control_path=tmp_path / "Control.xlsx",
        working_root=tmp_path,
        mode="validate_only",
        production_tables_path=tmp_path / "Production_Tables",
        output_path=tmp_path / "Output",
        run_id="TEST",
    )


def _touch_changelog(path: Path, mtime: int | None = None) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"xlsx")
    if mtime is not None:
        os.utime(path, (mtime, mtime))
    return path


def test_resolve_omitted_one_file(tmp_path: Path):
    control = _control(tmp_path)
    only = _touch_changelog(control.output_path / "ChangeLog_ONLY.xlsx")
    assert _resolve_change_log(control, None) == only


def test_resolve_omitted_picks_newest(tmp_path: Path):
    control = _control(tmp_path)
    older = _touch_changelog(control.output_path / "ChangeLog_OLD.xlsx", mtime=1)
    newer = _touch_changelog(
        control.output_path / "ChangeLog_NEW.xlsx", mtime=2_000_000_000
    )
    assert older != newer
    assert _resolve_change_log(control, None) == newer


def test_resolve_explicit_existing_path(tmp_path: Path):
    control = _control(tmp_path)
    _touch_changelog(control.output_path / "ChangeLog_NEW.xlsx", mtime=2_000_000_000)
    pinned = _touch_changelog(control.output_path / "ChangeLog_PINNED.xlsx", mtime=1)
    assert _resolve_change_log(control, pinned) == pinned


def test_resolve_bare_filename_under_output(tmp_path: Path):
    control = _control(tmp_path)
    _touch_changelog(control.output_path / "ChangeLog_NEW.xlsx", mtime=2_000_000_000)
    pinned = _touch_changelog(control.output_path / "ChangeLog_PINNED.xlsx", mtime=1)
    assert _resolve_change_log(control, Path("ChangeLog_PINNED.xlsx")) == pinned


def test_resolve_omitted_missing_output_dir(tmp_path: Path):
    control = _control(tmp_path)
    with pytest.raises(ChangeLogResolveError, match="Output folder not found"):
        _resolve_change_log(control, None)


def test_resolve_omitted_no_changelogs(tmp_path: Path):
    control = _control(tmp_path)
    control.output_path.mkdir()
    with pytest.raises(ChangeLogResolveError, match="No ChangeLog_"):
        _resolve_change_log(control, None)


def test_resolve_bare_filename_missing(tmp_path: Path):
    control = _control(tmp_path)
    control.output_path.mkdir()
    with pytest.raises(ChangeLogResolveError, match="Change Log not found"):
        _resolve_change_log(control, Path("ChangeLog_MISSING.xlsx"))


def test_apply_failed_report_exits_nonzero(tmp_path: Path):
    b, a = _cr_dirs(tmp_path, "CR")
    cols = ["Age", "Rate"]
    prod = tmp_path / "Production_Tables" / "MORT_TABLE.csv"
    _write_table(prod, cols, 2, [["20", "0.1"]])
    _write_table(b / "MORT_TABLE.csv", cols, 2, [["20", "0.1"]])
    _write_table(a / "MORT_TABLE.csv", cols, 2, [["20", "0.2"]])
    control = _write_control(tmp_path, [("CR", 1, "Y", "Y", "missing prod", "")], "EXIT1")
    clog = generate_change_log(control)
    prod.unlink()
    assert main([str(control), "--mode", "apply", "--change-log", str(clog)]) == 1


def test_resolve_missing_explicit_path(tmp_path: Path):
    control = _control(tmp_path)
    control.output_path.mkdir()
    missing = tmp_path / "elsewhere" / "ChangeLog_GONE.xlsx"
    with pytest.raises(ChangeLogResolveError, match="Change Log not found"):
        _resolve_change_log(control, missing)
