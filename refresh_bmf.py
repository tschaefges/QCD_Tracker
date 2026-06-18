# Run on the first of each month to refresh the IRS EO BMF data.
# The IRS updates the EO BMF monthly; stale data may include revoked organizations
# or miss newly registered ones.
#
# Usage:  python refresh_bmf.py
#
# Download URLs verified 2026-06-13 at:
#   https://www.irs.gov/charities-non-profits/exempt-organizations-business-master-file-extract-eo-bmf
# If downloads fail, re-check that page — the IRS has changed these URLs in the past.

import tempfile
from datetime import datetime
from io import StringIO
from pathlib import Path

import requests
import pandas as pd

SCRIPT_DIR = Path(__file__).parent

IRS_BASE_URL    = "https://www.irs.gov/pub/irs-soi/"
REGIONAL_FILES  = ["eo1.csv", "eo2.csv", "eo3.csv", "eo4.csv"]
KEEP_COLS       = ["EIN", "NAME", "STREET", "CITY", "STATE", "ZIP"]
READ_DTYPE      = {"EIN": str, "SUBSECTION": str, "STATUS": str, "DEDUCTIBILITY": str, "ZIP": str}


def read_config():
    config_path = SCRIPT_DIR / "config.txt"
    if not config_path.exists():
        print("ERROR: config.txt not found in this folder.")
        print("Please create a file named config.txt containing one line:")
        print("the UNC path to the shared IRS data folder.")
        print(r"Example:  \\House\Documents\IRS_Data")
        return None
    text = config_path.read_text().strip()
    if not text:
        print("ERROR: config.txt is empty. It must contain the UNC path to the IRS data folder.")
        return None
    return text


def download_regional(filename):
    url = IRS_BASE_URL + filename
    print(f"  Downloading {filename}...", end=" ", flush=True)
    try:
        r = requests.get(url, timeout=120)
        r.raise_for_status()
        df = pd.read_csv(
            StringIO(r.content.decode("latin-1")),
            dtype=READ_DTYPE,
            low_memory=False,
        )
        print(f"{len(df):,} rows")
        return df
    except Exception as exc:
        print(f"FAILED ({exc})")
        return None


def filter_qcd_eligible(df):
    # 501(c)(3) organizations with deductible contributions and active exemption status
    mask = (
        (df["SUBSECTION"] == "03") &
        (df["DEDUCTIBILITY"] == "1") &
        (df["STATUS"].isin(["01", "02"]))
    )
    return df.loc[mask, KEEP_COLS].copy()


def _safe_write_csv(df, target_path):
    target_path = Path(target_path)
    tmp_fd, tmp_name = tempfile.mkstemp(dir=target_path.parent, suffix=".tmp")
    try:
        with open(tmp_fd, "w", newline="", encoding="utf-8") as f:
            df.to_csv(f, index=False)
        Path(tmp_name).replace(target_path)
    except Exception:
        Path(tmp_name).unlink(missing_ok=True)
        raise


def main():
    print(f"QCD Charity Lookup — Refresh IRS EO BMF")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print()

    server_path = read_config()
    if server_path is None:
        return

    print("Downloading regional EO BMF files from IRS...")
    frames = []
    failed = []
    for filename in REGIONAL_FILES:
        df = download_regional(filename)
        if df is None:
            failed.append(filename)
        else:
            frames.append(df)

    if failed:
        print("\nERROR: Refresh aborted. These regional files failed to download:")
        for filename in failed:
            print(f"  - {filename}")
        print("Existing CSV files were not changed.")
        return

    print(f"\nCombining {len(frames)} regional files and filtering...")
    combined = pd.concat(frames, ignore_index=True)
    filtered = filter_qcd_eligible(combined)
    print(f"  {len(filtered):,} QCD-eligible organizations (501(c)(3), deductible, active)")

    # Save local backup (write to temp file, then replace to avoid partial writes)
    backup_path = SCRIPT_DIR / "eo_bmf_backup.csv"
    _safe_write_csv(filtered, backup_path)
    print(f"\nLocal backup saved: {backup_path}")

    # Save to server
    server_dir = Path(server_path)
    server_file = server_dir / "eo_bmf.csv"
    try:
        _safe_write_csv(filtered, server_file)
        print(f"Server copy saved:  {server_file}")
    except Exception as exc:
        print(f"\nWARNING: Could not write to server ({exc})")
        print(f"  The local backup at {backup_path} is current.")
        print("  Check that the server is running and the share is accessible.")

    print(f"\nDone. {datetime.now().strftime('%Y-%m-%d %H:%M')}")


if __name__ == "__main__":
    main()
