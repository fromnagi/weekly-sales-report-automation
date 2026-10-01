# Validation notes

Source checks run on 2026-10-01 used fictional sample data. They establish the tested input and output behavior, not compatibility with arbitrary client exports.

## Environment and commands

Windows 11 (build 26200), Python **3.12.10**, openpyxl **3.1.5**, pytest **9.1.1**, and pip **25.0.1**. Other installed test/runtime dependencies: et_xmlfile 2.0.0, colorama 0.4.6, iniconfig 2.3.0, packaging 26.3, pluggy 1.6.0, and Pygments 2.21.0. PyInstaller was not installed or exercised for these checks.

The [README source commands](../README.md#run-from-source) were run in a fresh local clone with the candidate changes applied and a newly created virtual environment. The clone imported its own source, not the owner checkout's installed module. Candidate code and tests matched the owner checkout after Git's CRLF/LF normalization. This used the existing Windows host, not a clean Windows machine.

## Results

| Check | Direct result |
| --- | --- |
| Baseline suite | 22 tests passed before product changes. |
| New regressions on old code | Both sparse-row cases failed with `KeyError: 'status'`; 22 formula cases failed their intended assertions. |
| Final suite | `python -m pytest -q`: **49 passed** in the owner checkout and **49 passed** in the fresh clone. |
| Documented normal path | Fresh environment installation, sample generation, CLI report generation, and tests completed successfully. |
| Demo result | 3 files; 75 readable rows = 53 accepted + 7 row exceptions + 15 rows in a rejected file; 1 duplicate; total 987.75. The rejected file also has one file-level exception record, not another source row. |
| Workbook | Four expected sheets and two Summary charts. All sheet cell values matched the tracked demo after regeneration in both checkouts. ZIP integrity was checked in the owner checkout. |
| Input preservation | SHA-256 values before and after reporting matched in regression tests and direct demo checks. Sample generation deliberately creates inputs; report generation only reads them. |
| Static checks | `git diff --check` and `python -m compileall -q src scripts tests` completed successfully. |

The tracked demo workbook and screenshots were retained: the sample contains no sparse invalid row or actual formula, so these repairs do not change its report values. No expected demo values were updated.

## Focused coverage

- Sparse trailing status cells created with openpyxl and by removing a cell from the existing Excel-saved fixture become `MISSING_STATUS` row exceptions. Neighboring valid rows and a separate valid CSV still contribute; source counts reconcile and input bytes stay unchanged. The fixture variant is an XML mutation, not a newly saved Excel file.
- Formula cases cover each required column, an extra named column, and a row containing only formulas, with absent caches and deliberately inconsistent numeric caches. A direct old-reader reproduction accepted cached `999` for `=1+1`; another dropped a formula-only row entirely. Formula rows now become `UNSUPPORTED_FORMULA` exceptions, retain literal formula text, and do not reserve a duplicate key. Array and data-table formula objects also produce inspectable exceptions.
- Formula-like literal text in both CSV and XLSX stays text in the output. Existing tests cover duplicate conflicts, malformed files, stale dimensions, styled blank residue, repeat runs, CP949 text, and Excel-illegal control characters.
- Windows-wrapper tests simulate output replacement denial and check preservation of the previous report and concise error messages. They also cover empty input, all-rejected input, and automatic-open failure.

## Encoding and delivery boundaries

CSV decoding attempts UTF-8 (with or without BOM), then CP949 after a decode failure. Bytes invalid in both are rejected. This does not identify the original encoding: a Latin-1 category encoded as bytes `A4 A1` failed UTF-8 decoding, decoded differently under CP949, and was accepted. UTF-8 exports and inspection of text are recommended; automatic encoding detection was not added.

The current checks did not build or launch a packaged EXE, test a machine without Python, or open the generated workbook in Microsoft Excel. Existing screenshots are previews, not current proof of repair-free Excel compatibility. Prior audit or delivery labels are not used as evidence for this repair.

The first-worksheet, fixed-schema, and transaction-identity rules remain specific to the demo. Accounting policy, arbitrary workbook support, real client integration, deployment, and continuous reliability are outside this evidence.
