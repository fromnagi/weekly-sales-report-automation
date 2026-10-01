# Case study: weekly sales reporting from messy branch files

This is a portfolio project built around a fictional small multi-branch retailer. The sample data and business are fictional; no paying client commissioned this implementation, and no measured customer saving or production deployment is claimed. The workflow is realistic: each week, staff receive separate sales exports and need one trustworthy Excel report.

## Problem

Manual preparation starts with collecting CSV and XLSX files from several branches. Someone must line up columns, check dates and amounts, spot duplicate transactions, exclude unusable rows, calculate branch and category sales, and then explain why source and reported counts differ. A workbook containing only the final total hides the data-quality work. A manual process also makes a repeated file or a missing column easy to overlook.

## Constraints

The program leaves branch exports unchanged and writes a new workbook outside the input folder. CSV and XLSX are the supported inputs, with a fixed nine-column contract for this fictional business. Only completed sales enter totals. The demo has a client-specific identity assumption: an order line is `(order_id, line_no)`, with order IDs unique across branches. The first valid occurrence wins; later copies need review rather than an invented correction. A Windows wrapper provides a double-click entry point for a built package; the source route requires Python.

A practical constraint was that malformed files should remain visible while valid files still contribute. An unreadable file cannot honestly be assigned a known row count. The report therefore separates row-level exceptions, file-level exception records, and readable rows inside rejected files.

## Engineering decisions

Processing is deterministic: top-level filenames are sorted and rows are handled in source order. The code validates a row before it may claim a transaction identity. For a later valid duplicate, Exceptions names the conflict and the first accepted row remains the one counted. Each accepted record retains its source filename and row position. Summary accounts for readable source rows as accepted rows, row exceptions, or rows in rejected files.

Writing uses a temporary workbook in the output directory followed by atomic replacement. When replacement fails, the earlier report remains available and the Windows wrapper reports failure. The wrapper also distinguishes empty input from a report that was created with zero accepted rows. It records concise support status without putting customer row text into a fatal-error message. Customer text beginning with `=` is written as literal text in Excel. The source exports are read only; tests compare their SHA-256 hashes before and after processing.

Sparse XLSX rows are padded to the header width before validation, so an unstored final cell becomes a missing-value exception. Actual XLSX formulas are rejected at row level rather than trusting a missing or stale cached value. Exceptions retains the formula as literal text and names the affected columns. This does not require an Excel calculation engine.

## Using the report

The recipient places the week's files in `input` and double-clicks `Weekly Sales Report.exe`. A completion message shows accepted rows, review rows, and rejected files. The generated workbook opens when file association permits. Clean_Data holds accepted normalized records; Exceptions shows row and file problems with reasons and provenance; Run_Info shows what happened to each discovered file. The original inputs stay where they were.

The tracked [demo workbook](demo/weekly_sales_report_demo.xlsx) lets a prospect inspect the result without running the program. The [Summary screenshot](assets/weekly-sales-summary.png) and [Exceptions / Run_Info screenshot](assets/exceptions-and-run-info.png) provide previews. [Validation notes](VALIDATION.md) give current source-run results and their limits.

## Known limitations

This exact input contract is not a promise to process arbitrary customer workbooks. It reads the first XLSX worksheet and rejects formula rows. CSV decoding tries UTF-8 (with or without BOM), then CP949; it cannot reliably identify other encodings, which may produce garbled text. It surfaces `.xls`, `.xlsm`, and `.xlsb` as unsupported. Different column names, transaction identity, re-export precedence, refunds, cancellations, taxes, currencies, and precision rules require discussion with a client and representative files. The portable EXE is unsigned, and some Windows environments may block it. Current source tests do not establish Python-free delivery or Excel repair-free compatibility. No real customer-file integration, production deployment, or continuous reliability claim follows from this portfolio demonstration.
