"""Distill the RAW Master Net Change Validation Report (100k+ rows, ~110MB,
split across "Data_1"/"Data_2" sheets by SharePoint/Excel) down to the small
3-column format `price_impact.load_master()` actually needs, written as a
gzip-compressed CSV (sku, final, old) for the bundled default -- much smaller
and faster to parse in-browser than .xlsx (no openpyxl overhead at all).
Loaded back via `price_impact.load_master_csv()`.

Usage:
    .venv/bin/python3 distill_master_report.py <raw_xlsx_path> <original_filename>

Writes/overwrites:
    master_net_change_validation_report.csv.gz
    master_net_change_validation_report.meta.json
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import pandas as pd

REQUIRED_COLS = ["Service SKU", "Final Price", "Old Price"]


def _read_sheet(path: str, sheet_name: str) -> pd.DataFrame | None:
    try:
        return pd.read_excel(path, sheet_name=sheet_name, engine="calamine",
                              header=0, usecols=REQUIRED_COLS)
    except Exception:
        try:
            return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl",
                                  header=0, usecols=REQUIRED_COLS)
        except Exception:
            return None


def distill(raw_path: str, original_filename: str, out_path: Path, meta_path: Path) -> int:
    frames = []
    for sheet in ("Data_1", "Data_2", "Data"):
        df = _read_sheet(raw_path, sheet)
        if df is not None and not df.empty:
            frames.append(df)
    if not frames:
        raise ValueError("No Data_1/Data_2/Data sheet with the expected columns was found.")

    combined = pd.concat(frames, ignore_index=True)
    combined.columns = ["sku", "final", "old"]
    combined["sku"] = combined["sku"].astype(str).str.strip()
    combined["old"] = pd.to_numeric(combined["old"], errors="coerce")
    combined["final"] = pd.to_numeric(combined["final"], errors="coerce")
    bad = {"", "nan", "service sku", "sku"}
    combined = combined[~combined["sku"].str.lower().isin(bad)]
    combined = combined.dropna(subset=["old", "final"])
    combined = combined[combined["old"] > 0]
    # last-write-wins for duplicate SKUs, same semantics as load_master()
    combined = combined.drop_duplicates(subset=["sku"], keep="last")

    out = combined[["sku", "final", "old"]]
    out.to_csv(out_path, index=False, compression="gzip")

    m = re.search(r"(\d{2}[-_][A-Za-z]{3}[-_]\d{2,4})", original_filename)
    date_suffix = m.group(1) if m else "unknown-date"
    meta_path.write_text(json.dumps({
        "date_suffix": date_suffix,
        "original_filename": original_filename,
    }, indent=1))

    return len(out)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <raw_xlsx_path> <original_filename>")
        sys.exit(1)
    raw_path, original_filename = sys.argv[1], sys.argv[2]
    here = Path(__file__).parent
    n = distill(
        raw_path, original_filename,
        here / "master_net_change_validation_report.csv.gz",
        here / "master_net_change_validation_report.meta.json",
    )
    print(f"Wrote {n} unique SKUs to master_net_change_validation_report.csv.gz")
