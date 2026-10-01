# Weekly Sales Report Automation

Turn a week's branch sales files into one Excel report that shows both the sales result and the records needing attention.

**A portfolio demo with fictional sales data, not commissioned client work.** It combines CSV/XLSX exports from multiple branches, validates sales rows, keeps the first valid copy of each order line, and records excluded rows or files with reasons. The output is one workbook with totals, accepted data, exceptions, and a per-file processing record. Source files stay untouched.

Open the [demo workbook](docs/demo/weekly_sales_report_demo.xlsx) to inspect the result, or [run it from source](#run-from-source) with the included sample files.

![Demo Summary sheet](docs/assets/weekly-sales-summary.png)

## The problem

A small multi-branch business may receive separate CSV and XLSX exports every week. Preparing a report by hand means combining files, checking invalid rows and duplicate transactions, excluding unreliable data from totals, and rebuilding branch and category summaries. The report keeps the reasons for exclusion alongside the final numbers.

## The workflow

Branch CSV/XLSX files → validation → duplicate checks → accepted rows and exceptions → summaries → Excel workbook.

In a built Windows package: place this week's files in `input` → double-click `Weekly Sales Report.exe` → open `output/weekly_sales_report.xlsx`. Valid rows and files can continue when others are rejected. The repository includes source and a demo workbook; it does not include an EXE or ZIP.

## What the report gives you

| Sheet | Why it matters |
| --- | --- |
| **Summary** | Sales total, branch and category charts, review counts, and processing health. |
| **Clean_Data** | Accepted rows with normalized values and their source file and row. |
| **Exceptions** | Row or file problems, reasons, scope, and available original values. |
| **Run_Info** | One status record per discovered file, including its contribution or rejection reason. |

![Demo Exceptions and Run_Info views](docs/assets/exceptions-and-run-info.png)

## Example result

The three fictional files in `sample_data/input` contain **75 readable source rows: 53 accepted, 7 row-level review items, and 15 rows excluded with one rejected file**. The report identifies **1 duplicate** and totals accepted sales at **987.75**. A rejected file produces one file-level exception record; its rows are counted separately from row-level review items.

## How excluded data is handled

- Inputs are read without modification; accepted rows retain source file and row provenance.
- File order and the first valid occurrence of a transaction determine duplicate handling. Later copies, including conflicts, remain visible in Exceptions.
- A malformed file can be isolated while valid files continue, and rejected files remain visible in Run_Info.
- The workbook is written to a temporary file and replaced atomically. If replacement fails, the previous valid report remains.
- Missing XLSX cells are row exceptions. Rows containing actual XLSX formulas are rejected explicitly; export calculated values before processing. Formula-like customer text remains literal text in the output.
- Excel lock files are ignored. CSV decoding tries UTF-8 (with or without BOM), then CP949. This is a fallback, not encoding detection: some other encodings can decode as garbled text and still be accepted. Prefer UTF-8 and check exported text.

## Windows delivery

The build script packages `input`, `output`, `examples`, instructions, `Weekly Sales Report.exe`, and its `_internal` runtime. Keep the **entire folder** together. This is an unsigned build, so an organization's security policy may block or warn about it. The current source validation does not establish compatibility on a Windows machine without Python; test a built package in the intended delivery environment.

## Validation evidence

Tests cover row and file validation, sparse XLSX rows, formulas with missing or stale caches, duplicates, source byte preservation, repeat runs, and output replacement failures. See [validation details](docs/VALIDATION.md) for commands, results, and limits, or the [case study](docs/CASE_STUDY.md) for design choices. The screenshots show the included demo; the current checks read generated workbooks with openpyxl and do not certify rendering in Microsoft Excel.

## Tech

Python · openpyxl · PyInstaller · pytest · PowerShell

## Limitations and adaptation points

Customer column schemas need adaptation when they differ. The demo assumes `(order_id, line_no)` identifies a transaction and that `order_id` is unique across branches; a client's correction or re-export rule must be agreed. Only the first XLSX worksheet is read, and formula rows are unsupported. `.xls`, `.xlsm`, and `.xlsb` are surfaced as unsupported, not processed. Refunds, cancellations, taxes, currency, high-precision amounts, and accounting treatment need client-specific rules. See [input requirements](INPUT_REQUIREMENTS.txt) for the exact demo contract.

## Run from source

From the repository root in PowerShell with Python 3.12 installed (the package requires Python 3.10+):

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[test]"
.\.venv\Scripts\python.exe scripts\generate_sample_data.py
.\.venv\Scripts\python.exe -m weekly_sales_report --input sample_data\input --output sample_data\output\weekly_sales_report.xlsx
.\.venv\Scripts\python.exe -m pytest
```

## Build the Windows package

From the repository root on Windows, after the source setup above:

```powershell
.\.venv\Scripts\python.exe -m pip install -e ".[test,build]"
.\scripts\build_windows.ps1
```

The script creates `dist\Weekly Sales Report\` and `dist\WeeklySalesReport.zip`. Deliver the ZIP or the whole folder, including `_internal`.

This repair used AI assistance. The files and tests are available for inspection; no measured client saving or production deployment is claimed.
