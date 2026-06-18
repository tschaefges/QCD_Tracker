# Python Code Review Notes

Date reviewed: 2026-06-18

Scope: README.md plus the Python scripts in this project. No source code changes were made during this review.

## Project Summary

This project is a Windows-oriented QCD charity lookup workflow. The main tool, `lookup_charity.py`, reads a filtered IRS EO BMF CSV, searches charity names, and writes the selected charity name, EIN, and address into the next available row of `QCD_Tracker_YYYY.xlsx`.

The monthly refresh workflow, `refresh_bmf.py`, downloads IRS regional EO BMF files, filters them to likely QCD-eligible organizations, and writes both a shared server CSV and a local backup CSV.

The one-time scripts, `modify_spreadsheet.py` and `add_date_columns.py`, prepare the workbook layout expected by the lookup tool.

## Read-Only Verification Performed

Using the bundled Codex Python runtime, I inspected the current workbook and local backup CSV without modifying them.

Confirmed facts:

- `QCD_Tracker_2026.xlsx` contains the expected sheet: `QCD Transaction Log`.
- Row 3 headers match the README layout, including `Charity Name`, `Charity EIN`, `Charity Address`, `Amount ($)`, and the three check date columns.
- Rows 4 through 23 in the Charity Name column are currently empty.
- `eo_bmf_backup.csv` has the expected columns: `EIN`, `NAME`, `STREET`, `CITY`, `STATE`, `ZIP`.

## Findings

### P1: Partial IRS Downloads Can Overwrite Good Data With Incomplete Data

File: `refresh_bmf.py`

Relevant code:

- `download_regional()` returns `None` when one regional file fails.
- `main()` appends only successful downloads to `frames`.
- The script aborts only if all downloads fail.
- It writes the filtered output even when only one, two, or three regional files were downloaded successfully.

Why this matters:

The README says the refresh process downloads all four regional EO BMF files. If one regional IRS file fails temporarily, the script can still overwrite both `eo_bmf_backup.csv` and the server `eo_bmf.csv` with a dataset missing an entire region. That would make valid charities disappear from lookup results until the next successful refresh.

This is a data integrity issue because the script reports success after producing an incomplete dataset.

Recommended fix:

- Treat any missing regional file as a failed refresh.
- Do not write either output CSV unless all expected regional files download and parse successfully.
- Print the list of failed files and leave the existing server and backup files untouched.

Suggested implementation shape:

```python
frames = []
failed = []

for filename in REGIONAL_FILES:
    df = download_regional(filename)
    if df is None:
        failed.append(filename)
    else:
        frames.append(df)

if failed:
    print("\nERROR: Refresh aborted. These regional files failed:")
    for filename in failed:
        print(f"  - {filename}")
    print("Existing CSV files were not changed.")
    return
```

Optional hardening:

- Validate that each downloaded CSV contains all required columns before combining.
- Write to a temporary file first, then replace the target file only after a complete successful write.

### P1: Charity Search Treats User Input as a Regex Instead of a Literal Substring

File: `lookup_charity.py`

Relevant code:

```python
results = df[df["NAME"].str.contains(term, case=False, na=False)]
```

Why this matters:

The README says search is case-insensitive and matches any substring of the charity name. However, pandas `str.contains()` defaults to `regex=True`. That means user input is interpreted as a regular expression.

Possible effects:

- A search like `st. mary` can match more than literal `ST. MARY`, because `.` means any character in regex.
- A search containing `[` or another incomplete regex construct can raise an exception and crash the tool.
- Users may see confusing matches that do not correspond to the text they typed.

Recommended fix:

Use literal substring matching:

```python
results = df[df["NAME"].str.contains(term, case=False, na=False, regex=False)]
```

Suggested verification:

- Search for a normal term such as `humane society`.
- Search for a term containing regex punctuation, such as `st. mary`.
- Search for `[` and confirm the tool does not crash.

### P2: Spreadsheet Save Failures Are Not Handled

File: `lookup_charity.py`

Relevant code:

```python
try:
    wb = openpyxl.load_workbook(xlsx_path)
except PermissionError:
    ...

...

wb.save(xlsx_path)
return True
```

Why this matters:

The code handles `PermissionError` when opening the workbook, but it does not handle failures when saving. Excel, OneDrive, antivirus software, or a network/storage hiccup could allow the workbook to open but fail during save.

Based on available evidence, this is a plausible failure mode rather than one reproduced during review.

Recommended fix:

- Wrap `wb.save(xlsx_path)` in `try/except`.
- Catch at least `PermissionError` and `OSError`.
- Return `False` with a user-friendly message if save fails.
- Close the workbook in a `finally` block.

Suggested implementation shape:

```python
try:
    wb.save(xlsx_path)
except PermissionError:
    print(f"\nERROR: Could not save {xlsx_path.name}.")
    print("Please close it in Excel and try again.")
    return False
except OSError as exc:
    print(f"\nERROR: Could not save {xlsx_path.name}: {exc}")
    return False
finally:
    wb.close()
```

Important detail:

If save fails after cell values were assigned in memory, the workbook object should still be closed. The file on disk should remain unchanged unless openpyxl partially wrote before the failure. Using a temp-file-and-replace pattern would provide stronger safety, but may be more than this project needs.

### P2: One-Time Spreadsheet Scripts Can Modify the Wrong Tracker When Multiple Trackers Exist

Files:

- `modify_spreadsheet.py`
- `add_date_columns.py`

Relevant code:

```python
matches = list(SCRIPT_DIR.glob("QCD_Tracker*.xlsx"))
...
xlsx_path = matches[0]
```

Why this matters:

The README describes yearly files named like `QCD_Tracker_2026.xlsx`. Over time, the directory may contain multiple trackers. The main lookup tool handles this by prompting the user to choose a workbook. The one-time scripts do not. They silently choose the first glob result.

This suggests a user could accidentally alter the wrong year if more than one tracker exists in the folder.

Recommended fix:

- Reuse the interactive selection behavior from `lookup_charity.py`.
- Exclude Excel lock files whose names start with `~$`.
- If there is exactly one match, use it.
- If there are multiple matches, print a numbered list and ask which file to modify.

Suggested implementation shape:

```python
matches = [
    p for p in SCRIPT_DIR.glob("QCD_Tracker*.xlsx")
    if not p.name.startswith("~$")
]

if len(matches) == 1:
    xlsx_path = matches[0]
else:
    print("Multiple QCD_Tracker spreadsheets found:")
    for i, p in enumerate(matches, 1):
        print(f"  {i}. {p.name}")
    ...
```

Optional improvement:

Extract this small workbook-selection helper into a shared module only if future maintenance needs it. For now, duplicating a tiny helper may be simpler than adding a shared abstraction for two one-time scripts.

### P3: Minor Unused Code

Files:

- `lookup_charity.py`
- `modify_spreadsheet.py`

Relevant code:

- `lookup_charity.py` sets `warned = False` and later `warned = True`, but never uses the variable.
- `modify_spreadsheet.py` imports `glob as _glob`, but never uses it.

Why this matters:

This is not a behavior bug. It is small maintenance noise.

Recommended fix:

- Remove `warned`.
- Remove the unused `_glob` import.

## Suggested Implementation Order

1. Fix literal search matching in `lookup_charity.py`.
2. Add save-error handling and workbook closing in `lookup_charity.py`.
3. Make `refresh_bmf.py` abort unless all four regional files succeed.
4. Add workbook selection prompts to the one-time spreadsheet scripts.
5. Remove unused variables/imports.

This order fixes the highest user-facing risks first while keeping each change small and independently testable.

## Suggested Manual Test Plan

### Lookup Tool

Run:

```text
python lookup_charity.py
```

Test cases:

- Search for a common literal term, for example `humane society`.
- Search for a term containing a period, for example `st. mary`.
- Search for `[`, and confirm the program does not crash.
- Select a result, confirm write, and verify columns D, E, and F in the next empty row.
- Repeat with the workbook open in Excel, and confirm the tool reports a friendly save/open error.

### Refresh Script

Because the refresh script depends on IRS network downloads, use one of these approaches:

- Temporarily set one regional filename to an invalid name and confirm the script aborts without writing output files.
- Or mock `download_regional()` in a small test harness so one file returns `None`.

Expected behavior after the fix:

- Any failed regional file aborts the refresh.
- Existing `eo_bmf.csv` and `eo_bmf_backup.csv` are not overwritten.
- The console output names the failed regional files.

### Spreadsheet Setup Scripts

Use copies of test workbooks, not the production tracker.

Test cases:

- Folder contains one `QCD_Tracker*.xlsx`: script selects it automatically.
- Folder contains two `QCD_Tracker*.xlsx` files: script prompts for selection.
- Folder contains an Excel lock file such as `~$QCD_Tracker_2026.xlsx`: script ignores it.

## Open Questions

- Should `refresh_bmf.py` preserve the previous backup if the server write fails after a successful full download? Current behavior updates the local backup first, then attempts the server write. That seems reasonable for Tom's machine, but it means local and server data can differ after a server failure.
- Should the lookup tool consider a row occupied if any of columns D, E, or F are filled, instead of only checking Charity Name in column D? The current behavior is consistent with the workflow, but checking all three could protect against partially edited rows.
- Should the fixed row range, rows 4 through 23, remain hard-coded? The README describes that range, so this is acceptable today. If the workbook grows, `LAST_DATA_ROW` will need to change or become dynamic.
