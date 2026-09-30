from pathlib import Path

from openpyxl import load_workbook

from weekly_sales_report import windows_app
from test_report import HEADERS, rows, write_csv


METRICS = {
    "clean_row_count": 3,
    "row_exception_count": 2,
    "rejected_file_count": 1,
    "total_source_row_count": 7,
}


def test_frozen_delivery_root_uses_executable_location(monkeypatch, tmp_path):
    executable = tmp_path / "portable" / "Weekly Sales Report.exe"
    monkeypatch.setattr(windows_app.sys, "frozen", True, raising=False)
    monkeypatch.setattr(windows_app.sys, "executable", str(executable))
    assert windows_app.delivery_root() == executable.parent


def test_empty_input_creates_folders_and_gives_actionable_feedback(tmp_path):
    messages = []
    result = windows_app.run(tmp_path, notify=lambda text, **kw: messages.append((text, kw)),
                             open_report=lambda path: None)
    assert result == 1
    assert (tmp_path / "input").is_dir() and (tmp_path / "output").is_dir()
    assert "Add CSV or XLSX files" in messages[0][0]
    assert messages[0][1] == {"error": True}
    assert "status=FAILURE" in (tmp_path / "output" / "last_run.log").read_text(encoding="utf-8")
    assert not (tmp_path / "output" / "weekly_sales_report.xlsx").exists()


def _one_input(root: Path) -> None:
    (root / "input").mkdir()
    (root / "input" / "sales.csv").write_text("sample", encoding="utf-8")


def test_success_feedback_receives_report_metrics(monkeypatch, tmp_path):
    _one_input(tmp_path)
    called = []
    messages = []

    def build(input_dir, report_path):
        called.append((input_dir, report_path))
        report_path.write_bytes(b"report")
        return METRICS

    monkeypatch.setattr(windows_app, "build_report", build)
    result = windows_app.run(tmp_path, notify=lambda text, **kw: messages.append(text),
                             open_report=lambda path: called.append(path))
    report = tmp_path / "output" / "weekly_sales_report.xlsx"
    assert result == 0
    assert called == [(tmp_path / "input", report), str(report)]
    assert all(text in messages[0] for text in ("Accepted rows: 3", "Rows needing review: 2",
                                                "Rejected files: 1", str(report)))
    log = (tmp_path / "output" / "last_run.log").read_text(encoding="utf-8")
    assert "status=SUCCESS" in log and "review_rows=2" in log and "rejected_files=1" in log


def test_auto_open_failure_keeps_success(monkeypatch, tmp_path):
    _one_input(tmp_path)
    report = tmp_path / "output" / "weekly_sales_report.xlsx"
    def build(_input_dir, report_path):
        report_path.write_bytes(b"report")
        return METRICS
    monkeypatch.setattr(windows_app, "build_report", build)
    def cannot_open(_path):
        raise OSError("association unavailable")
    messages = []
    assert windows_app.run(tmp_path, notify=lambda text, **kw: messages.append(text),
                           open_report=cannot_open) == 0
    assert report.read_bytes() == b"report"
    assert "Could not open automatically" in messages[0]
    assert "status=SUCCESS" in (tmp_path / "output" / "last_run.log").read_text(encoding="utf-8")


def test_fatal_failure_is_concise_and_does_not_log_row_text(monkeypatch, tmp_path):
    _one_input(tmp_path)
    def fail(_input_dir, _report_path):
        raise RuntimeError("private customer row contents")
    monkeypatch.setattr(windows_app, "build_report", fail)
    messages = []
    assert windows_app.run(tmp_path, notify=lambda text, **kw: messages.append((text, kw)),
                           open_report=lambda path: None) == 1
    log = (tmp_path / "output" / "last_run.log").read_text(encoding="utf-8")
    assert "exception_class=RuntimeError" in log and "status=FAILURE" in log
    assert "private customer row contents" not in log + messages[0][0]
    assert "Traceback" not in messages[0][0]
    assert messages[0][1] == {"error": True}


def test_permission_error_preserves_existing_report(monkeypatch, tmp_path):
    _one_input(tmp_path)
    output = tmp_path / "output"
    output.mkdir()
    report = output / "weekly_sales_report.xlsx"
    report.write_bytes(b"existing report")
    def locked(_input_dir, _report_path):
        raise PermissionError("locked")
    monkeypatch.setattr(windows_app, "build_report", locked)
    messages = []
    assert windows_app.run(tmp_path, notify=lambda text, **kw: messages.append(text),
                           open_report=lambda path: None) == 1
    assert report.read_bytes() == b"existing report"
    assert "read-only" in messages[0] and "output folder" in messages[0]
    assert "exception_class=PermissionError" in (output / "last_run.log").read_text(encoding="utf-8")


def test_real_run_isolates_bad_file_and_warns_when_all_rejected(tmp_path):
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    write_csv(input_dir / "good.csv", HEADERS, [[
        "2026-09-21", "B1", "O1", 1, "P1", "Food", 1, 2, "COMPLETED",
    ]])
    (input_dir / "bad.xlsx").write_bytes(b"not an xlsx")
    messages = []
    assert windows_app.run(tmp_path, notify=lambda text, **kw: messages.append((text, kw)),
                           open_report=lambda path: None) == 0
    assert "Accepted rows: 1" in messages[0][0]
    workbook = load_workbook(tmp_path / "output" / "weekly_sales_report.xlsx", read_only=True)
    try:
        assert rows(workbook["Exceptions"])[0]["source_file"] == "bad.xlsx"
        assert rows(workbook["Run_Info"])[0]["status"] == "REJECTED"
    finally:
        workbook.close()

    (input_dir / "good.csv").unlink()
    messages.clear()
    assert windows_app.run(tmp_path, notify=lambda text, **kw: messages.append((text, kw)),
                           open_report=lambda path: None) == 0
    assert "no sales rows were accepted" in messages[0][0]
    assert messages[0][1] == {"warning": True}
    assert "status=CREATED_NO_ACCEPTED_DATA" in (tmp_path / "output" / "last_run.log").read_text(encoding="utf-8")


def test_real_run_permission_error_logs_safe_os_metadata(monkeypatch, tmp_path):
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    write_csv(input_dir / "good.csv", HEADERS, [[
        "2026-09-21", "B1", "O1", 1, "P1", "private", 1, 2, "COMPLETED",
    ]])
    report = tmp_path / "output" / "weekly_sales_report.xlsx"
    report.parent.mkdir()
    report.write_bytes(b"previous")
    def denied(source, target):
        raise PermissionError(13, "private row contents", str(source), None, str(target))
    monkeypatch.setattr("weekly_sales_report.report.os.replace", denied)
    messages = []
    assert windows_app.run(tmp_path, notify=lambda text, **kw: messages.append(text),
                           open_report=lambda path: None) == 1
    assert report.read_bytes() == b"previous"
    assert "read-only" in messages[0]
    log = (tmp_path / "output" / "last_run.log").read_text(encoding="utf-8")
    assert "exception_class=PermissionError" in log and "errno=13" in log
    assert "filename=" in log and f"filename2={report}" in log
    assert "private row contents" not in log + messages[0]
