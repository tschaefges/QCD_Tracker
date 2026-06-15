# QCD Charity Lookup

A standalone Windows tool that searches the IRS Exempt Organizations Business Master File (EO BMF) and writes charity data (name, EIN, address) directly into a QCD tracking spreadsheet.

---

## Overview

Two users (husband and wife) each maintain their own `QCD_Tracker_YYYY.xlsx` on separate Windows computers. Both computers read a shared IRS data file stored on a home server. The lookup tool lets each user search for a charity by name and write the result into the next available row of their spreadsheet with a single keypress.

```
[Home Server]
  \\House\Documents\IRS_Data\
    eo_bmf.csv               ← shared, updated monthly by Tom

[Tom's Computer]             [Wife's Computer]
  QCD_Tracker_2026.xlsx        QCD_Tracker_2026.xlsx
  lookup_charity.exe           lookup_charity.exe
  config.txt                   config.txt
  eo_bmf_backup.csv            (no local backup)
  refresh_bmf.py
  build.bat
```

---

## Components

| File | Purpose |
|------|---------|
| `QCD_Tracker_YYYY.xlsx` | The QCD tracking spreadsheet (one per user) |
| `config.txt` | One-line config: UNC path to the server IRS data folder |
| `lookup_charity.exe` | The interactive charity search tool (compiled EXE, no Python needed) |
| `lookup_charity.py` | Source for the EXE (compiled with PyInstaller) |
| `refresh_bmf.py` | Monthly script to download fresh IRS data (Tom's computer only) |
| `eo_bmf_backup.csv` | Local fallback copy of IRS data (written by `refresh_bmf.py`) |
| `modify_spreadsheet.py` | One-time setup script that inserts the Charity Address column |
| `add_date_columns.py` | One-time setup script that inserts the three check date columns |
| `build.bat` | Compiles `lookup_charity.py` into `lookup_charity.exe` |
| `requirements.txt` | Python dependencies for Tom's computer |

---

## Prerequisites

- **Tom's computer:** Python 3.x installed. Run `pip install -r requirements.txt` once.
- **Wife's computer:** Nothing. She only needs the compiled `lookup_charity.exe` and `config.txt`.
- **Home server:** A shared folder accessible via UNC path (e.g., `\\House\Documents\IRS_Data`). Both computers need read access; Tom's computer needs write access.

---

## First-Time Setup

### 1. Install Python dependencies (Tom's computer only)

```
pip install -r requirements.txt
```

### 2. Prepare the spreadsheet (one time)

Close `QCD_Tracker_2026.xlsx` in Excel, then run:

```
python modify_spreadsheet.py
```

This inserts a new **Charity Address** column (F) between EIN and Amount, applies matching formatting, and updates the SUM formulas. Run it once and discard it (or keep it for reference).

Open the spreadsheet in Excel afterward to verify the new column looks correct before proceeding.

### 3. Create the server folder

Create `\\House\Documents\IRS_Data\` on the home server if it does not already exist. Confirm both computers can reach it.

### 4. Download IRS data

```
python refresh_bmf.py
```

This downloads all four regional EO BMF files from the IRS, filters to 501(c)(3) organizations with deductible contributions, and saves the result to:
- `\\House\Documents\IRS_Data\eo_bmf.csv` (server, used by both computers)
- `eo_bmf_backup.csv` (local backup on Tom's computer)

Expect this to take a few minutes. The filtered file typically contains several hundred thousand organizations.

### 5. Test before compiling

```
python lookup_charity.py
```

Search for a charity, select a result, and confirm it writes correctly to the spreadsheet. Fix any issues before compiling.

### 6. Compile the EXE

```
build.bat
```

The resulting EXE will be at `dist\lookup_charity.exe`. Copy it to Tom's working folder alongside `config.txt`.

### 7. Set up wife's computer

Copy two files to the same folder as her `QCD_Tracker_2026.xlsx`:
- `lookup_charity.exe`
- `config.txt`

That is all she needs. Run `lookup_charity.exe` and test a search end-to-end.

---

## Monthly Maintenance

On the first of each month, Tom runs:

```
python refresh_bmf.py
```

This updates `\\House\Documents\IRS_Data\eo_bmf.csv` (used by both computers) and refreshes the local backup on Tom's machine.

---

## Using lookup_charity.exe

Double-click `lookup_charity.exe` (or run from the command line). The tool:

1. Reads `config.txt` to locate the IRS data on the server.
2. Falls back to `eo_bmf_backup.csv` in the same folder if the server is unreachable.
3. Displays a startup banner with the organization count and data date.

Then it enters a search loop:

```
Enter charity name to search (or Q to quit):
> humane society

Found 3 matches for "humane society":

  #    Name                                EIN          City, State
  ─    ────────────────────────────────    ───────────  ───────────
  1    CHAMPAIGN COUNTY HUMANE SOCIETY     37-1234567   Champaign, IL
  2    HEARTLAND HUMANE SOCIETY            37-9999999   Bloomington, IL
  3    CENTRAL IL HUMANE SOCIETY           37-8888888   Springfield, IL

Enter number to select, S to search again, or Q to quit:
> 1

─────────────────────────────────────────────
  Selected organization:

  Name:    CHAMPAIGN COUNTY HUMANE SOCIETY
  EIN:     37-1234567
  Address: 123 Main St, Champaign, IL 61820

  Will write to row 5 of QCD_Tracker_2026.xlsx

  Columns to be filled:
    D (Charity Name)    → CHAMPAIGN COUNTY HUMANE SOCIETY
    E (Charity EIN)     → 37-1234567
    F (Charity Address) → 123 Main St, Champaign, IL 61820
─────────────────────────────────────────────
Confirm? (Y/N):
> Y

Done. Written to row 5.
```

Search is case-insensitive and matches any substring of the charity name. If more than 10 results are returned, the tool shows the first 10 and suggests a more specific search term.

Type `Q` or `QUIT` at any prompt to exit.

---

## Spreadsheet Column Layout (after modify_spreadsheet.py)

Sheet: **QCD Transaction Log**

| Col | Header | Filled by |
|-----|--------|-----------|
| A | Date of Distribution | User |
| B | IRA Custodian | User |
| C | IRA Account # (last 4 digits) | User |
| D | Charity Name | `lookup_charity.exe` |
| E | Charity EIN | `lookup_charity.exe` |
| F | Charity Address | `lookup_charity.exe` |
| G | Amount ($) | User |
| H | Date Check Requested from IRA | User |
| I | Date Check Received from IRA | User |
| J | Date Check Mailed to Charity | User |
| K | Check / Wire Reference # | User |
| L | Check Payable to Charity? (Y/N) | User |
| M | Written Acknowledgment Received? (Y/N) | User |
| N | Acknowledgment States No Goods/Services? (Y/N) | User |
| O | Custodian Distribution Request Copy Saved? (Y/N) | User |
| P | 501(c)(3) Status Verified? (Y/N) | User |
| Q | 1099-R Code Y Received? (Y/N) | User |
| R | Notes | User |

---

## IRS Data Source

The EO BMF is a free public dataset published monthly by the IRS:

https://www.irs.gov/charities-non-profits/exempt-organizations-business-master-file-extract-eo-bmf

`refresh_bmf.py` downloads regional files EO1 through EO4, concatenates them, and filters to organizations that are:
- 501(c)(3) (`SUBSECTION = 03`)
- Contributions deductible (`DEDUCTIBILITY = 1`)
- Actively exempt (`STATUS = 01` or `02`)

If the IRS changes the download URLs (it has happened before), update `IRS_BASE_URL` in `refresh_bmf.py`. The correct URLs are listed on the IRS download page linked above.

---

## config.txt Format

Plain text file, one line, no trailing backslash:

```
\\House\Documents\IRS_Data
```

The tool appends `\eo_bmf.csv` to this path to locate the data file.

---

## Fallback Behavior

If the server is unreachable when `lookup_charity.exe` starts, it looks for `eo_bmf_backup.csv` in its own folder and displays a warning with the backup date. Wife's computer does not have a local backup, so if the server is unreachable she should wait until it is available again.
