# Validation evidence

This page separates checks performed for this portfolio package from earlier accepted delivery checks. It describes the demonstrated fictional sample, not production readiness for arbitrary customer files.

## Current package checks

| Check | Evidence and result |
| --- | --- |
| Full automated suite | `.venv/Scripts/python.exe -m pytest -q`: **22 passed** on the unchanged product code. Tests are in `tests/test_report.py` and `tests/test_windows_app.py`. |
| Normal report generation | The normal CLI generated `docs/demo/weekly_sales_report_demo.xlsx` from `sample_data/input`. CLI metrics: 3 discovered files, 2 processed, 1 rejected, 75 readable source rows, 53 accepted, 7 row-level exceptions, 15 rows in a rejected file, 1 duplicate, and 987.75 total sales. |
| Row reconciliation | `75 = 53 + 7 + 15`. The rejected file also has one file-level exception record; that record is not counted as another source row. |
| Workbook structure | The tracked workbook was inspected for `Summary`, `Clean_Data`, `Exceptions`, `Run_Info`, the expected metrics, and two native Summary charts. |
| Microsoft Excel | The tracked demo workbook opened in real Microsoft Excel without a repair prompt. Summary charts, Exceptions, and Run_Info rendered and were captured in `docs/assets/`. |

## Regression coverage in the 22 tests

- A mixed CSV/XLSX run checks fixed expected totals, provenance, source SHA-256 equality before and after two runs, repeat-run data equivalence, duplicate handling, and all four sheets.
- A fixture saved by Microsoft Excel checks that formatting residue beyond the data is ignored. Separate tests cover stale XLSX dimensions, blank CSV records, duplicate headers, and missing columns.
- Malformed XLSX XML, invalid ZIP content, unsupported spreadsheet formats, and unsupported CSV encoding are tested alongside a valid file so the valid contribution survives.
- CP949 text and formula-like customer text are checked for literal round-trip behavior. Illegal Excel control characters are visible in Exceptions while a valid neighboring row survives.
- Windows-wrapper tests cover empty input, successful counts, automatic-open failure, all-rejected input, locked/denied output, preservation of the previous report, and concise error logging without customer row text.

## Earlier accepted Windows delivery evidence

The accepted Windows delivery checkpoint documented a portable EXE launched **outside this repository**, with Python environment variables neutralized. That verified the demonstrated frozen route used its bundled `_internal` runtime rather than a local Python installation. The delivery exercise also included malformed and valid files together, empty input, locked output, and input SHA-256 preservation. These manual delivery checks were **not rerun for this documentation package**; the current gate changed no product or build code.

## Boundaries

The fixed demo schema and identity rule are client-specific. An unreadable file has no invented row count. The 22 passing tests and one Excel open show the tested routes, not every workbook variant or every Windows security policy. The EXE is unsigned. Client accounting and correction rules would need to be agreed using representative customer files before a real deployment.
