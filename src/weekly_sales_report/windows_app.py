"""Small, native Windows entry point for the portable delivery folder."""

from __future__ import annotations

import ctypes
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Callable

from weekly_sales_report.report import build_report


REPORT_NAME = "weekly_sales_report.xlsx"
LOG_NAME = "last_run.log"
INPUT_FORMATS = {".csv", ".xlsx", ".xls", ".xlsm", ".xlsb"}


def delivery_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def _message(text: str, *, error: bool = False, warning: bool = False) -> None:
    icon = 0x10 if error else 0x30 if warning else 0x40
    ctypes.windll.user32.MessageBoxW(None, text, "Weekly Sales Report", icon)


def _write_log(path: Path, lines: list[str]) -> bool:
    try:
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return True
    except OSError:
        return False


def _log_lines(status: str) -> list[str]:
    return [f"timestamp={datetime.now().astimezone().isoformat(timespec='seconds')}",
            f"status={status}"]


def _exception_metadata(exc: BaseException) -> list[str]:
    lines = [f"exception_class={type(exc).__name__}"]
    for name in ("errno", "winerror", "filename", "filename2"):
        value = getattr(exc, name, None)
        if value is not None:
            lines.append(f"{name}={value}")
    return lines


def run(root: Path | None = None, *,
        notify: Callable[..., None] = _message,
        open_report: Callable[[str], None] = os.startfile) -> int:
    root = delivery_root() if root is None else Path(root)
    input_dir = root / "input"
    output_dir = root / "output"
    report_path = output_dir / REPORT_NAME
    log_path = output_dir / LOG_NAME

    try:
        input_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)
        if not any(path.is_file() and not path.name.startswith("~$") and path.suffix.lower() in INPUT_FORMATS
                   for path in input_dir.iterdir()):
            written = _write_log(log_path, _log_lines("FAILURE") + [
                "exception_class=EmptyInput",
                "error=No CSV or XLSX files found in input",
            ])
            note = "" if written else "\nThe support log could not be written."
            notify("No new report was created. Add CSV or XLSX files to the input folder and run again." + note,
                   error=True)
            return 1
        metrics = build_report(input_dir, report_path)
    except PermissionError as exc:
        written = _write_log(log_path, _log_lines("FAILURE") + [
            "error=Access denied while creating the report",
        ] + _exception_metadata(exc))
        note = f"\nDetails: {log_path}" if written else "\nThe support log could not be written."
        notify("The report could not be written. It may be open in Excel, read-only, or the output folder may not be writable." + note,
               error=True)
        return 1
    except Exception as exc:
        written = _write_log(log_path, _log_lines("FAILURE") + [
            "error=Report creation failed",
        ] + _exception_metadata(exc))
        note = f"Check {log_path} for details." if written else "The support log could not be written."
        notify(f"Could not create the report. {note}", error=True)
        return 1

    try:
        open_report(str(report_path))
        opened = True
    except Exception:
        opened = False
    no_data = metrics["clean_row_count"] == 0
    _write_log(log_path, _log_lines("CREATED_NO_ACCEPTED_DATA" if no_data else "SUCCESS") + [
        f"accepted_rows={metrics['clean_row_count']}",
        f"review_rows={metrics['row_exception_count']}",
        f"rejected_files={metrics['rejected_file_count']}",
        f"source_rows={metrics['total_source_row_count']}",
        f"auto_open={'SUCCESS' if opened else 'FAILED'}",
    ])
    message = (("Report created, but no sales rows were accepted. Review Exceptions and Run_Info."
                if no_data else "Report created.")
               + f"\n\nAccepted rows: {metrics['clean_row_count']}\n"
               f"Rows needing review: {metrics['row_exception_count']}\n"
               f"Rejected files: {metrics['rejected_file_count']}\n\n"
               f"Report: {report_path}")
    if not opened:
        message += "\n\nCould not open automatically. Open the report from the output folder."
    notify(message, warning=True) if no_data else notify(message)
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
