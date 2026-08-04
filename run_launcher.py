"""Launcher used by Run.bat.

Uses the mode from Control.xlsx. For Stage 2 modes, auto-picks the newest
ChangeLog_*.xlsx under the Control output folder.
"""

from __future__ import annotations

import sys
from pathlib import Path

from prophet_table_tool.__main__ import main
from prophet_table_tool.control import read_control


def launch(control_path: Path) -> int:
    control = read_control(control_path)
    mode = control.mode
    print(f"Mode (from Control): {mode}")

    argv: list[str] = [str(control_path)]
    if mode in {"validate_only", "apply"}:
        if not control.output_path.is_dir():
            print(
                f"Output folder not found: {control.output_path}\n"
                "Run Stage 1 (generate_changelog) first, or check output_path in Control.xlsx.",
                file=sys.stderr,
            )
            return 1
        logs = sorted(
            control.output_path.glob("ChangeLog_*.xlsx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not logs:
            print(
                f"No ChangeLog_*.xlsx found in {control.output_path}\n"
                "Run Stage 1 (generate_changelog) first, or place a Change Log in Output.",
                file=sys.stderr,
            )
            return 1
        change_log = logs[0]
        print(f"Using Change Log: {change_log}")
        argv += ["--change-log", str(change_log)]

    return main(argv)


def _cli(argv: list[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args:
        print("Usage: python run_launcher.py path\\to\\Control.xlsx", file=sys.stderr)
        return 2
    control_path = Path(args[0])
    if not control_path.is_file():
        print(f"Control file not found: {control_path}", file=sys.stderr)
        return 1
    return launch(control_path)


if __name__ == "__main__":
    raise SystemExit(_cli())
