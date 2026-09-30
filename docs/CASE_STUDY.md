# Case study: weekly sales reporting from messy branch files

This is a portfolio project built around a fictional small multi-branch retailer. The sample data and business are fictional; no paying client commissioned this implementation, and no measured customer saving or production deployment is claimed. The workflow is realistic: each week, staff receive separate sales exports and need one trustworthy Excel report.

## 1. Problem

Manual preparation starts with collecting CSV and XLSX files from several branches. Someone must line up columns, check dates and amounts, spot duplicate transactions, exclude unusable rows, calculate branch and category sales, and then explain why source and reported counts differ. A workbook containing only the final total hides the data-quality work. A manual process also makes a repeated file or a missing column easy to overlook.

The demo models a weekly task that could take roughly 1–2 hours by hand. That is a scenario estimate, not measured time saved by this tool. The useful outcome is a report in which the sales result and the reasons for exclusion are visible together.

## 2. Constraints

The recipient should be able to use a portable Windows folder without installing Python. The program must leave branch exports unchanged and write a new workbook outside the input folder. CSV and XLSX are the supported inputs, with a fixed nine-column contract for this fictional business. Only completed sales enter totals. The demo has a client-specific identity assumption: an order line is `(order_id, line_no)`, with order IDs unique across branches. The first valid occurrence wins; later copies need review rather than an invented correction.

A practical constraint was that malformed files should remain visible while valid files still contribute. An unreadable file cannot honestly be assigned a known row count. The report therefore separates row-level exceptions, file-level exception records, and readable rows inside rejected files.

## 3. First vertical slice

The project first established the complete path from input files to one workbook: discover files, validate records, split accepted and rejected data, aggregate accepted sales, and write Summary, Clean_Data, Exceptions, and Run_Info. This made the business rule observable before visual polish. The next step improved the Summary layout, charts, filters, number formats, and review sheets. Windows delivery then wrapped the same reporting engine in a double-click workflow with a portable PyInstaller folder and concise completion or failure feedback.

That order mattered. A polished dashboard would have been weak proof if duplicate or invalid rows could still enter totals. The initial end-to-end report made those errors testable, and the later design made them understandable to an ordinary recipient.

## 4. Important failure modes discovered

An independent clean-room audit challenged the accepted path with less tidy spreadsheet exports and delivery conditions. Among the issues it exposed were XLSX sheets whose used range was inflated by formatting, stale worksheet dimensions that could hide data, numeric identifiers that Excel had already rounded or represented ambiguously, and CSV encoding differences. The audit also pushed on malformed workbook streams, Excel-illegal control characters, formula-like customer text, conflicting duplicate rows, and behavior when the destination workbook was open or unwritable.

These were material because a report could appear successful while silently omitting source rows, changing identifiers, or losing the previous valid output. The repairs were targeted to the established contract. The project did not turn into a general-purpose spreadsheet parser.

## 5. Engineering decisions

Processing is deterministic: top-level filenames are sorted and rows are handled in source order. The code validates a row before it may claim a transaction identity. For a later valid duplicate, Exceptions names the conflict and the first accepted row remains the one counted. Each accepted record retains its source filename and row position. In the demo, the row accounting is explicit: `75 = 53 accepted + 7 row-level review items + 15 readable rows in one rejected file`.

Writing uses a temporary workbook in the output directory followed by atomic replacement. When replacement fails, the earlier report remains available and the Windows wrapper reports failure. The wrapper also distinguishes empty input from a report that was created with zero accepted rows. It records concise support status without putting customer row text into a fatal-error message. Customer text beginning with `=` is written as literal text in Excel. The source exports are read only; tests compare their SHA-256 hashes before and after processing.

## 6. Independent audit

The independent audit was a challenge to the workflow, not a claim of certification. It checked the engine, workbook, and portable delivery against realistic files and failure paths. Its findings provided specific cases to reproduce and test.

## 7. Audit-driven repairs

Focused fixes and regression tests addressed styled blank cells, stale dimensions, numeric identifier limits, malformed XLSX content, CP949 text, illegal characters, and output failure paths. An Excel-saved fixture checked compatibility with a file actually saved by Microsoft Excel, alongside synthetic edge cases. The Windows package was exercised outside the repository with the Python environment neutralized, testing that the bundled `_internal` runtime was sufficient for the demonstrated route.

## 8. Final demonstrated workflow

The recipient places the week's files in `input` and double-clicks `Weekly Sales Report.exe`. A completion message shows accepted rows, review rows, and rejected files. The generated workbook opens when file association permits. Clean_Data holds accepted normalized records; Exceptions shows row and file problems with reasons and provenance; Run_Info shows what happened to each discovered file. The original inputs stay where they were.

The tracked [demo workbook](demo/weekly_sales_report_demo.xlsx) lets a prospect inspect the result without running the program. The [Summary screenshot](assets/weekly-sales-summary.png) and [Exceptions / Run_Info screenshot](assets/exceptions-and-run-info.png) were captured from that workbook in Microsoft Excel, rather than drawn as mockups.

## 9. Evidence

At this checkpoint, the full pytest suite has 22 passing tests. The deterministic sample has three files, 75 readable source rows, 53 accepted rows, seven row-level review items, 15 rows from a rejected file, one duplicate, and sales of 987.75. The generated workbook has four expected sheets and two native Excel charts. It opened in Microsoft Excel without a repair prompt. Earlier accepted delivery verification also covered a frozen EXE launched outside the repository, mixed malformed and valid files, empty input, locked output, and unchanged input hashes. [Validation notes](VALIDATION.md) distinguish these checks and their limits.

## 10. Known limitations

This exact input contract is not a promise to process arbitrary customer workbooks. It reads the first XLSX worksheet, supports CSV encoded as UTF-8 (with or without BOM) or CP949, and surfaces `.xls`, `.xlsm`, and `.xlsb` as unsupported. Different column names, transaction identity, re-export precedence, refunds, cancellations, taxes, currencies, and precision rules require discussion with a client and representative files. The portable EXE is unsigned, and some Windows environments may block it. No real customer-file integration, production deployment, or continuous reliability claim follows from this portfolio demonstration.
