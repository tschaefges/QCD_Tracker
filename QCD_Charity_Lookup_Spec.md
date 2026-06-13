# QCD Charity Lookup System — Complete Specification

**Purpose:** This document is a complete implementation specification for Claude Code. It describes a charity lookup system that searches the IRS Exempt Organizations Business Master File (EO BMF) and writes charity data (name, EIN, address) into a QCD tracking spreadsheet. No prior context from the design conversation is required — everything needed to build the system is contained here.

**Inference flags:** Any detail marked `[INFERRED]` is an assumption that should be verified against actual data before finalizing the implementation.

---

## 1. System Overview

### 1.1 Purpose

Two users (husband and wife) each maintain their own QCD tracking spreadsheet on their own Windows computers. Both computers access a shared IRS data file stored on a Windows home server via UNC path. A standalone Windows executable (`lookup_charity.exe`) allows each user to search for a charity by name and write the charity's name, EIN, and address directly into the next available row of their local spreadsheet.

### 1.2 Components

| Component | Location | Who uses it |
|---|---|---|
| `eo_bmf.csv` | `\\House\Documents\IRS_Data\` | Both (read-only) |
| `refresh_bmf.py` | Tom's computer, same folder as his spreadsheet | Tom only (monthly) |
| `lookup_charity.exe` | Each person's computer, same folder as their spreadsheet | Both |
| `config.txt` | Each person's computer, same folder as their spreadsheet | Both |
| `eo_bmf_backup.csv` | Each person's computer, same folder as their spreadsheet | Automatic fallback |
| `QCD_Tracker_YYYY.xlsx` | Each person's computer | Each person |

---

## 2. IRS EO BMF Data File

### 2.1 Source

The EO BMF is a free public dataset published monthly by the IRS. The download page is:

https://www.irs.gov/charities-non-profits/exempt-organizations-business-master-file-extract-eo-bmf

The IRS publishes regional CSV files (EO1 through EO4, plus a Puerto Rico file). For a national lookup, all regional files should be downloaded and concatenated into a single `eo_bmf.csv`.

**[INFERRED]** The regional files are named `eo1.csv`, `eo2.csv`, `eo3.csv`, `eo4.csv`. Verify the actual filenames and URLs on the IRS download page before implementing `refresh_bmf.py`. The IRS has changed these URLs in the past.

### 2.2 Key Columns (expected)

**[INFERRED]** Based on IRS EO BMF documentation, the relevant columns are expected to be:

| Column name | Content |
|---|---|
| `EIN` | Employer Identification Number (9 digits, no dashes) |
| `NAME` | Organization name |
| `STREET` | Street address |
| `CITY` | City |
| `STATE` | Two-letter state code |
| `ZIP` | ZIP code |
| `SUBSECTION` | Tax-exempt subsection code (3 = 501(c)(3)) |
| `STATUS` | Revocation status (`0` = active) **[INFERRED — verify]** |
| `DEDUCTIBILITY` | Deductibility code (1 = deductible) **[INFERRED — verify]** |

**Critical:** Before implementing, download one regional file and inspect the actual column names. The IRS uses all-caps column names and the exact names above should be verified.

### 2.3 Filtering

When loading the EO BMF for search, filter to rows where:
- `SUBSECTION` == `3` (501(c)(3) organizations only)
- Organization is not revoked (verify the correct column and value for this)
- `DEDUCTIBILITY` indicates contributions are deductible (verify the correct value)

This reduces the file size significantly and excludes organizations that are ineligible for QCD purposes.

### 2.4 Formatting EIN for Display

The IRS stores EINs as 9-digit strings without dashes (e.g., `371234567`). Format for display and spreadsheet entry as `XX-XXXXXXX` (e.g., `37-1234567`).

### 2.5 Formatting Address

Combine street, city, state, and ZIP into a single string for display and spreadsheet entry:

```
{STREET}, {CITY}, {STATE} {ZIP}
```

Example: `123 Main St, Champaign, IL 61820`

---

## 3. File Structure

On each user's computer, all files live in the same folder. The folder location is up to each user.

```
[User's chosen folder]/
    QCD_Tracker_2026.xlsx       ← The user's spreadsheet (must contain "QCD_Tracker" in filename)
    lookup_charity.exe          ← The lookup tool
    config.txt                  ← One-line config: UNC path to server data folder
    eo_bmf_backup.csv           ← Local backup of EO BMF data (written by refresh_bmf.py)
```

On Tom's computer only, `refresh_bmf.py` also lives in this same folder.

On the home server:
```
\\House\Documents\IRS_Data\
    eo_bmf.csv                  ← Primary data file, updated monthly
```

---

## 4. config.txt Format

`config.txt` is a plain text file containing exactly one line: the UNC path to the folder on the server that contains `eo_bmf.csv`. No key, no label — just the path.

```
\\House\Documents\IRS_Data
```

- No trailing backslash
- The EXE appends `\eo_bmf.csv` to this path to locate the data file
- If `config.txt` is missing or empty, the EXE must display a clear error message:

```
ERROR: config.txt not found in this folder.
Please create a file named config.txt containing one line:
the UNC path to the shared IRS data folder.
Example:  \\House\Documents\IRS_Data
Press Enter to exit.
```

---

## 5. refresh_bmf.py

This script runs on Tom's computer only. Tom has Python installed. It does not need to be compiled to an EXE.

### 5.1 Behavior

1. Downloads all regional EO BMF CSV files from the IRS website.
2. Concatenates them into a single DataFrame, keeping only the columns needed for search (NAME, EIN, STREET, CITY, STATE, ZIP, SUBSECTION, and any revocation/deductibility columns).
3. Filters to active 501(c)(3) organizations with deductible contributions.
4. Saves the result as `eo_bmf.csv` to the server path read from `config.txt` in the script's own folder.
5. Also saves a copy as `eo_bmf_backup.csv` to the script's own folder (same folder as the spreadsheet).
6. Prints progress as it goes, and a final summary showing how many organizations are in the file and when it was saved.

### 5.2 Error Handling

- If the server path is unreachable, save `eo_bmf_backup.csv` locally and print a warning that the server copy could not be updated.
- If any individual regional file fails to download, print a warning and continue with the others rather than aborting entirely.
- If `config.txt` is missing, print the same error message described in Section 4.

### 5.3 Dependencies

- `requests` (for downloading)
- `pandas` (for concatenation and filtering)

Include a `requirements.txt` for Tom's computer:
```
requests
pandas
openpyxl
```

### 5.4 Recommended Run Schedule

Add a comment at the top of `refresh_bmf.py` reminding Tom to run it on the first of each month. The IRS updates the EO BMF monthly.

---

## 6. lookup_charity.py / lookup_charity.exe

This is the primary user-facing tool. It is compiled to a standalone Windows EXE using PyInstaller so it runs on machines without Python installed.

### 6.1 Startup Sequence

1. Read `config.txt` from the script's own folder. If missing, show error (Section 4) and exit.
2. Attempt to load `eo_bmf.csv` from the server path in `config.txt`.
3. If the server file is not found or the server is unreachable:
   - Look for `eo_bmf_backup.csv` in the script's own folder.
   - If found, load it and display a warning:
     ```
     WARNING: Could not reach server at \\House\Documents\IRS_Data
     Using local backup copy instead.
     Backup last updated: [date from file modification time]
     Data may not be current.
     ```
   - If not found either, display an error and exit:
     ```
     ERROR: Could not find IRS data file.
     Tried: \\House\Documents\IRS_Data\eo_bmf.csv
     Also tried: [local backup path]
     Please check that the server is running and try again,
     or run refresh_bmf.py to create a local backup.
     Press Enter to exit.
     ```
4. Display a startup banner:
   ```
   ─────────────────────────────────────────────
    QCD Charity Lookup — IRS EO BMF Search
    [N] organizations loaded  |  Data: [date]
   ─────────────────────────────────────────────
   ```
5. Enter the search loop (Section 6.2).

### 6.2 Search Loop

The tool stays open and repeats this loop until the user types `Q` or `QUIT` at any prompt.

**Step 1 — Search prompt:**
```
Enter charity name to search (or Q to quit):
>
```

- Search is case-insensitive
- Partial name matches are supported (substring match on the NAME column)
- A blank entry re-displays the prompt without searching

**Step 2 — Display results:**

If no matches found:
```
No matches found for "red cross". Try a shorter search term.
```

If matches found, display up to 10, numbered:
```
Found 3 matches for "humane society":

  #   Name                                EIN          City, State
  ─   ────────────────────────────────    ───────────  ───────────
  1   CHAMPAIGN COUNTY HUMANE SOCIETY     37-1234567   Champaign, IL
  2   CHAMPAIGN COUNTY HUMANE SOCIETY     37-9876543   Urbana, IL
  3   HEARTLAND HUMANE SOCIETY            37-1111111   Bloomington, IL

Enter number to select, S to search again, or Q to quit:
>
```

If more than 10 matches exist:
```
Showing 10 of 47 matches. Try a more specific search term for better results.
```

**Step 3 — Selection:**

User enters a number. Display the full details and the target row:

```
─────────────────────────────────────────────
  Selected organization:

  Name:    CHAMPAIGN COUNTY HUMANE SOCIETY
  EIN:     37-1234567
  Address: 123 Main St, Champaign, IL 61820

  Will write to row 7 of QCD_Tracker_2026.xlsx

  Columns to be filled:
    D (Charity Name)    → CHAMPAIGN COUNTY HUMANE SOCIETY
    E (Charity EIN)     → 37-1234567
    F (Charity Address) → 123 Main St, Champaign, IL 61820
─────────────────────────────────────────────
Confirm? (Y/N):
>
```

**Step 4 — Write or cancel:**

- If Y: write to the spreadsheet, display `Done. Written to row 7.`, then return to Step 1.
- If N: display `Cancelled. Nothing was written.`, then return to Step 1.
- If anything else: re-display the confirmation prompt.

### 6.3 Spreadsheet Targeting

The EXE locates the spreadsheet by:
1. Looking in its own folder for any `.xlsx` file whose name contains `QCD_Tracker` (case-insensitive).
2. If exactly one match is found, use it.
3. If multiple matches are found, display a numbered list and ask the user to choose.
4. If no match is found:
   ```
   ERROR: No spreadsheet found in this folder.
   Looking for: any .xlsx file with "QCD_Tracker" in the name.
   Folder checked: [path]
   Press Enter to exit.
   ```

### 6.4 Finding the Next Empty Row

The EXE should write to the first row in the QCD Transaction Log sheet where column D (Charity Name) is empty, starting from row 4 (the first data row, after the header rows). If all 20 data rows are full, warn the user:

```
WARNING: All data rows in the spreadsheet appear to be filled.
The spreadsheet supports 20 rows. You may need to add more rows manually,
or start a new spreadsheet for additional entries.
```

### 6.5 Dependencies (for compilation)

- `pandas`
- `openpyxl`

The EXE is compiled with PyInstaller. See Section 8 for build instructions.

---

## 7. Spreadsheet Modification — QCD_Tracker_2026.xlsx

The existing spreadsheet (already built) needs one new column inserted into the QCD Transaction Log sheet: **Charity Address**, to be placed between the existing EIN column (column E) and the Amount column (column F).

### 7.1 New Column Layout (QCD Transaction Log sheet)

After modification, the column order in the data area should be:

| Col | Header | Content |
|-----|--------|---------|
| A | Date of Distribution | Date the QCD was made |
| B | IRA Custodian | Name of custodian |
| C | IRA Account # (last 4 digits) | Account identifier |
| D | Charity Name | **Written by lookup_charity.exe** |
| E | Charity EIN | **Written by lookup_charity.exe** |
| F | Charity Address | **Written by lookup_charity.exe** ← NEW COLUMN |
| G | Amount ($) | Distribution amount |
| H | Check / Wire Reference # | Transaction reference |
| I | Check Payable to Charity? (Y/N) | Compliance checkbox |
| J | Written Acknowledgment Received? (Y/N) | Compliance checkbox |
| K | Acknowledgment States No Goods/Services? (Y/N) | Compliance checkbox |
| L | Custodian Distribution Request Copy Saved? (Y/N) | Compliance checkbox |
| M | 501(c)(3) Status Verified? (Y/N) | Compliance checkbox |
| N | 1099-R Code Y Received? (Y/N) | Compliance checkbox |
| O | Notes | Free text |

The existing total formula in row 24 (column F, previously Amount) must be updated to reference the new Amount column (G). Verify all formulas in the sheet after inserting the column.

### 7.2 Formatting

The new Charity Address column (F) should match the style of the adjacent columns: same header fill color, same body font, same border, width of approximately 28 characters, wrap text enabled.

---

## 8. Build Instructions — Compiling the EXE

These steps are performed on Tom's computer (which has Python).

### 8.1 Install PyInstaller

```
pip install pyinstaller
```

### 8.2 Compile

From the folder containing `lookup_charity.py`:

```
pyinstaller --onefile --console --name lookup_charity lookup_charity.py
```

The resulting `lookup_charity.exe` will be in the `dist\` subfolder. Copy it to the working folder alongside the spreadsheet and `config.txt`.

### 8.3 Distribution to Wife's Computer

Copy two files to the same folder as her `QCD_Tracker_2026.xlsx`:
- `lookup_charity.exe`
- `config.txt` (containing `\\House\Documents\IRS_Data`)

That is all she needs. No Python, no installation, no other dependencies.

---

## 9. Deployment Checklist

### 9.1 Home Server

- [ ] Create folder `\\House\Documents\IRS_Data\` if it does not exist
- [ ] Confirm both personal computers can reach `\\House\Documents\IRS_Data\` and have read/write access

### 9.2 Tom's Computer

- [ ] Install dependencies: `pip install -r requirements.txt`
- [ ] Create `config.txt` containing `\\House\Documents\IRS_Data`
- [ ] Run `refresh_bmf.py` to populate the server file and create the local backup
- [ ] Verify `eo_bmf.csv` appears at `\\House\Documents\IRS_Data\eo_bmf.csv`
- [ ] Compile `lookup_charity.exe` per Section 8
- [ ] Run `lookup_charity.exe` and confirm it finds the data and the spreadsheet
- [ ] Test a search, confirm the confirmation prompt, confirm the write

### 9.3 Wife's Computer

- [ ] Copy `lookup_charity.exe` to the folder with her spreadsheet
- [ ] Create `config.txt` in that same folder containing `\\House\Documents\IRS_Data`
- [ ] Run `lookup_charity.exe` and confirm it finds the server data and her spreadsheet
- [ ] Test a search end-to-end

---

## 10. Monthly Maintenance

On the first of each month, Tom runs:

```
python refresh_bmf.py
```

This updates `\\House\Documents\IRS_Data\eo_bmf.csv` (primary, used by both) and `eo_bmf_backup.csv` (local fallback on Tom's computer only). Wife's computer does not have a local backup; if the server is unreachable on her machine, she should wait until the server is available or ask Tom to run a refresh so the backup exists on her machine as well.

**Optional improvement (not in initial scope):** `refresh_bmf.py` could also write `eo_bmf_backup.csv` to a second path — her computer's folder — if her machine is online and reachable. This would give her a local fallback too.

---

## 11. Notes for Claude Code

- The `lookup_charity.py` script must use `__file__` (or `sys.executable` when compiled) to determine its own folder location, not `os.getcwd()`, since the working directory may differ from the EXE's location when the user double-clicks it.
- When opening the spreadsheet with openpyxl, use `keep_vba=False` and do not use `data_only=True` (which would destroy formulas on save).
- The EO BMF can be large (several hundred thousand rows). Load it once at startup, not on each search.
- pandas `str.contains()` with `case=False, na=False` is appropriate for the name search.
- Test with the server path both reachable and unreachable to verify fallback behavior.
- The spreadsheet sheet name to write to is `QCD Transaction Log`. If the sheet name is not found, display a clear error rather than crashing.
