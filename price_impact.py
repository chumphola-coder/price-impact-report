"""
price_impact.py
===============
Price Change Impact Analysis engine.

Given:
  1. Master Net Change Validation Report (first sheet / "Data" tab):
       col C (idx 2) = Service SKU
       col G (idx 6) = Final Price   (new unit list price)
       col H (idx 7) = Old Price     (previous unit list price)

  2. A quote file to check (two supported layouts — auto-detected):
       Layout A (e.g. All Quote-CAT Thaipak.xlsx):  header at row 2 (0-based)
         SKU              = col AE (30)
         Start Date       = col AJ (35)
         End Date         = col AK (36)
         Quantity         = col AL (37)
         Unit List Price  = col AN (39)

       Layout B (e.g. 332285614.xlsx):  header at row 31 (0-based)
         SKU              = col K  (10)
         Start Date       = col P  (15)
         End Date         = col Q  (16)
         Quantity         = col R  (17)
         Unit List Price  = col V  (21)

The engine auto-detects which layout by scanning for a header row
that contains "SKU" in the first 40 rows.

Prorated formula:
  days        = (end_date - start_date).days
  orig_prorated = unit_list * qty * (days + 1) / 365
  new_unit      = unit_list * (final / old)   [ratio from master]
  new_prorated  = new_unit  * qty * (days + 1) / 365
"""

from __future__ import annotations

from io import BytesIO

import openpyxl
import pandas as pd
import xlsxwriter


# ── Master loader ─────────────────────────────────────────────────────────────
def load_master(file) -> dict[str, dict]:
    """
    Load the Master Net Change Validation Report.
    Returns {service_sku: {'old': float, 'final': float, 'ratio': float}}.

    Reads only the three needed columns (C=SKU, G=Final Price, H=Old Price).
    Uses calamine engine (~5s for the 680K-row file) with openpyxl fallback.
    Always tries the 'Data' sheet first, then sheet 0.
    """
    if hasattr(file, "seek"):
        file.seek(0)

    def _read(engine):
        if hasattr(file, "seek"):
            file.seek(0)
        try:
            return pd.read_excel(file, sheet_name="Data",
                                 engine=engine, header=0, usecols=[2, 6, 7])
        except Exception:
            if hasattr(file, "seek"):
                file.seek(0)
            return pd.read_excel(file, sheet_name=0,
                                 engine=engine, header=0, usecols=[2, 6, 7])

    try:
        df = _read("calamine")
    except Exception:
        try:
            df = _read("openpyxl")
        except (IndexError, ValueError) as e:
            raise ValueError("Master file is missing required columns (needs at least columns A through H).") from e

    # Standardise column names regardless of exact header text
    df.columns = ["sku", "final", "old"]

    # Vectorised build — much faster than iterrows() for 680K rows
    df["sku"]   = df["sku"].astype(str).str.strip()
    df["old"]   = pd.to_numeric(df["old"],   errors="coerce")
    df["final"] = pd.to_numeric(df["final"], errors="coerce")
    bad = {"", "nan", "service sku", "sku"}
    df = df[~df["sku"].str.lower().isin(bad)]
    df = df.dropna(subset=["old", "final"])
    df = df[df["old"] > 0]
    df["ratio"] = df["final"] / df["old"]
    # last-write-wins for duplicate SKUs (same behaviour as before)
    df = df.drop_duplicates(subset=["sku"], keep="last")
    records = df[["sku", "old", "final", "ratio"]].to_dict("records")
    return {r["sku"]: {"old": r["old"], "final": r["final"], "ratio": r["ratio"]}
            for r in records}


def load_master_csv(file) -> dict[str, dict]:
    """
    Load the bundled-default Master Net Change Validation Report from its
    DISTILLED form: a small gzip-compressed CSV (columns sku, final, old)
    produced by distill_master_report.py from the real ~100MB+ multi-sheet
    SharePoint report. Only used for the bundled default -- user-uploaded
    files always go through load_master() (real .xlsx), never this path.

    Rows are already cleaned/deduped by the distill step, so no filtering is
    re-applied here; this stays intentionally tiny/fast (plain pandas CSV
    parse, no openpyxl/xlsx overhead at all -- the whole point of this format
    for in-browser/Pyodide loads where speed matters most).
    """
    df = pd.read_csv(file, compression="infer")
    df.columns = ["sku", "final", "old"]
    df["ratio"] = df["final"] / df["old"]
    records = df.to_dict("records")
    return {r["sku"]: {"old": r["old"], "final": r["final"], "ratio": r["ratio"]}
            for r in records}


# ── Quote loader ──────────────────────────────────────────────────────────────
_REQUIRED_KEYWORDS = {"sku", "start date", "end date", "unit list price"}

def _detect_header_row(raw: pd.DataFrame) -> int:
    """Return the 0-based row index of the header row, or raise ValueError."""
    for i, row in raw.iterrows():
        lowered = {str(v).strip().lower() for v in row if pd.notna(v)}
        if "sku" in lowered and "unit list price" in lowered:
            return i
    raise ValueError(
        "Could not find a header row with 'SKU' and 'Unit List Price'. "
        "Make sure this is a Cisco renewal quote export."
    )


def _col(headers: pd.Index, *candidates: str) -> str | None:
    """Return the first column name matching any candidate (case-insensitive)."""
    lower_map = {str(c).strip().lower(): c for c in headers}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def _find_col_pos(header_values: list[str], *candidates: str) -> int | None:
    """
    Return the 0-based position of the first header matching any candidate
    (case-insensitive), keeping the FIRST occurrence when a header name is
    duplicated. Used instead of `_col()` + `list.index()` because that
    two-step lookup can point at the wrong column when a workbook has
    duplicate header text.
    """
    lower_map: dict[str, int] = {}
    for idx, v in enumerate(header_values):
        key = v.strip().lower()
        if key and key not in lower_map:
            lower_map[key] = idx
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def _extract_quote_metadata(raw: pd.DataFrame, hdr_row: int) -> tuple[str, str]:
    """
    Scan the pre-header block (rows 0..hdr_row-1) for quote number and name.
    Format 2: 'Quote Number' in col B (idx 1), value in col C (idx 2).
              'Quote Name'   in col B (idx 1), value in col C (idx 2).
    Returns (quote_name, quote_number) or ("", "") if not found.
    """
    q_name, q_number = "", ""
    for i in range(hdr_row):
        row = raw.iloc[i]
        for j in range(len(row) - 1):
            label = str(row.iloc[j]).strip().lower() if pd.notna(row.iloc[j]) else ""
            val   = str(row.iloc[j + 1]).strip() if pd.notna(row.iloc[j + 1]) else ""
            if val.lower() in ("nan", ""):
                # try col j+2 if j+1 is empty
                if j + 2 < len(row):
                    val = str(row.iloc[j + 2]).strip() if pd.notna(row.iloc[j + 2]) else ""
            if "quote number" in label and val and val.lower() != "nan":
                q_number = val
            if "quote name" in label and val and val.lower() != "nan":
                q_name = val
    return q_name, q_number


def _build_quote_frame(raw: pd.DataFrame, hdr_row: int, source_name: str,
                       sheet_name: str) -> pd.DataFrame:
    """Turn one sheet's raw grid + detected header row into a normalised quote
    frame. Raises ValueError if the sheet is missing required columns."""
    # Extract quote name/number from metadata rows above the table header
    # (fallback for formats that don't have them as per-row columns)
    meta_name, meta_number = _extract_quote_metadata(raw, hdr_row)

    df = raw.iloc[hdr_row + 1 :].copy()
    df.columns = [str(v).strip() if pd.notna(v) else f"__col{j}" for j, v in enumerate(raw.iloc[hdr_row])]

    # locate required columns
    sku_col   = _col(df.columns, "SKU", "Service SKU")
    ulp_col   = _col(df.columns, "Unit List Price")
    qty_col   = _col(df.columns, "Quantity")
    start_col = _col(df.columns, "Start Date")
    end_col   = _col(df.columns, "End Date")

    missing = [n for n, c in [("SKU", sku_col), ("Unit List Price", ulp_col),
                               ("Quantity", qty_col), ("Start Date", start_col),
                               ("End Date", end_col)] if c is None]
    if missing:
        raise ValueError(f"Quote sheet '{sheet_name}' is missing columns: {', '.join(missing)}")

    out = pd.DataFrame()
    out["sku"]             = df[sku_col].astype(str).str.strip()
    out["unit_list_price"] = pd.to_numeric(df[ulp_col], errors="coerce")
    out["quantity"]        = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    out["start_date"]      = pd.to_datetime(df[start_col], errors="coerce")
    out["end_date"]        = pd.to_datetime(df[end_col], errors="coerce")

    # also carry quote/product info if present
    for extra, col in [("product_number", _col(df.columns, "Product Number")),
                       ("product_description", _col(df.columns, "Product Description")),
                       ("service_level", _col(df.columns, "Service Level")),
                       ("serial_number", _col(df.columns, "PAK/Serial Number")),
                       ("instance_number", _col(df.columns, "Instance Number")),
                       ("quote_name",   _col(df.columns, "Quote Name")),
                       ("quote_number", _col(df.columns, "Quote Number")),]:
        out[extra] = df[col].astype(str).str.strip() if col else ""

    # Fill quote_name / quote_number from this sheet's metadata block where the
    # per-row column is absent
    if meta_name and (out["quote_name"].eq("").all() or out["quote_name"].str.lower().eq("nan").all()):
        out["quote_name"] = meta_name
    if meta_number and (out["quote_number"].eq("").all() or out["quote_number"].str.lower().eq("nan").all()):
        out["quote_number"] = meta_number

    out["source_file"]  = source_name
    out["source_sheet"] = sheet_name

    # drop rows with no SKU or no price
    out = out[out["sku"].ne("") & out["sku"].str.lower().ne("nan")]
    out = out[out["unit_list_price"].notna() & out["unit_list_price"].gt(0)]
    return out.reset_index(drop=True)


def load_quote(file, source_name: str = "") -> pd.DataFrame:
    """
    Load a quote file, auto-detecting the table on EVERY worksheet that has one.

    A single workbook may contain several quote tabs (one quote per tab) plus
    non-table tabs (cover sheet, BOM, notes, ...). Every tab that carries a
    usable table is read and tagged with its own sheet name and quote
    name/number, then all are concatenated — so the report can present the price
    change per quote. Non-table tabs are skipped harmlessly.

    Returns a DataFrame with normalised columns:
      sku, unit_list_price, quantity, start_date, end_date, quote_name,
      quote_number, source_file, source_sheet, ...
    Rows with no SKU or zero/null unit_list_price are dropped.
    """
    if hasattr(file, "seek"):
        file.seek(0)
    xl = pd.ExcelFile(file, engine="openpyxl")

    frames: list[pd.DataFrame] = []
    col_errors: list[str] = []
    for sheet in xl.sheet_names:
        try:
            candidate = xl.parse(sheet, header=None, dtype=object)
        except Exception:
            continue  # unreadable tab — skip
        try:
            hdr_row = _detect_header_row(candidate.astype(str))
        except ValueError:
            continue  # no quote header on this tab — skip
        try:
            frame = _build_quote_frame(candidate, hdr_row, source_name, sheet)
        except ValueError as ce:
            col_errors.append(str(ce))  # table-like tab missing columns — skip
            continue
        if not frame.empty:
            frames.append(frame)

    if not frames:
        if col_errors:
            raise ValueError(col_errors[0])
        raise ValueError(
            "Could not find a usable quote table with 'SKU' and 'Unit List Price' "
            "in any worksheet. Make sure this is a Cisco renewal quote export."
        )

    return pd.concat(frames, ignore_index=True)


# ── Quote regeneration (write new prices back into the original workbook) ────
_METADATA_TOTAL_LABELS = {"quote extended list price", "extended list price", "total list price"}
_METADATA_NAME_LABELS = {"quote name"}
_ESTIMATE_SUFFIX = " - Estimate"


def _scan_metadata_labels(raw: pd.DataFrame, header_row: int) -> dict[str, tuple[int, int]]:
    """
    Scan the rows above `header_row` for label/value pairs (label text in one
    cell, value in a cell shortly to its right) matching known metadata field
    names. Returns 0-based {"quote_name": (row, col), "total": (row, col)},
    omitting keys that weren't found. Layouts with no such metadata block
    (e.g. one flat multi-quote table) simply yield an empty dict.
    """
    found: dict[str, tuple[int, int]] = {}
    for r in range(min(header_row, len(raw))):
        row = raw.iloc[r]
        for c in range(len(row) - 1):
            label = str(row.iloc[c]).strip().lower() if pd.notna(row.iloc[c]) else ""
            if not label:
                continue
            val_col = None
            for cc in range(c + 1, min(c + 3, len(row))):
                if pd.notna(row.iloc[cc]) and str(row.iloc[cc]).strip() != "":
                    val_col = cc
                    break
            if val_col is None:
                continue
            if "quote_name" not in found and label in _METADATA_NAME_LABELS:
                found["quote_name"] = (r, val_col)
            if "total" not in found and label in _METADATA_TOTAL_LABELS:
                found["total"] = (r, val_col)
    return found


def regenerate_quote_with_new_price(file, master: dict[str, dict]) -> tuple[bytes, dict]:
    """
    Return (new_workbook_bytes, stats). The uploaded quote workbook is
    reproduced with every other tab/column/formatting left untouched, except:
      - Unit List Price (and Prorated List Price, if that column exists) on
        matched quote lines are scaled by the master's price-change ratio --
        scaling both by the same ratio preserves the file's own proration
        exactly, without needing to re-derive it from quantity/dates.
      - Each quote's "Quote Extended List Price" metadata total (if present
        above the line-item table) is recalculated as the new sum of that
        quote's Prorated List Price column.
      - Each quote's "Quote Name" metadata field (if present) gets an
        " - Estimate" suffix so the regenerated file is clearly distinguishable
        from the original quote.
      - A "Summary" tab (if present) listing quote-number -> total-value pairs
        keyed by sheet name is updated to match each quote's new total.
    Sheets where no SKU + Unit List Price table can be detected (cover pages,
    BOM, notes, a "Summary" tab, ...) are left completely untouched.

    stats = {"matched_lines": int, "sheets_updated": list[str]}
    """
    if hasattr(file, "seek"):
        file.seek(0)
    file_bytes = file.read() if hasattr(file, "read") else file

    # Detection pass (pandas, read-only) -- reuses the exact same header/column
    # detection as load_quote() so the two stay consistent.
    xl = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    plans: dict[str, dict] = {}
    for sheet in xl.sheet_names:
        try:
            raw = xl.parse(sheet, header=None, dtype=object)
        except Exception:
            continue
        try:
            hdr_row = _detect_header_row(raw.astype(str))
        except ValueError:
            continue  # not a quote table tab -- leave completely untouched

        header_values = [str(v).strip() if pd.notna(v) else "" for v in raw.iloc[hdr_row]]
        sku_pos = _find_col_pos(header_values, "SKU", "Service SKU")
        ulp_pos = _find_col_pos(header_values, "Unit List Price")
        if sku_pos is None or ulp_pos is None:
            continue
        prorated_pos = _find_col_pos(header_values, "Prorated List Price", "Prorated List")

        metadata = _scan_metadata_labels(raw, hdr_row)

        plans[sheet] = {
            "header_excel_row": hdr_row + 1,
            "sku_col": sku_pos + 1,
            "ulp_col": ulp_pos + 1,
            "prorated_col": (prorated_pos + 1) if prorated_pos is not None else None,
            "quote_name_cell": (metadata["quote_name"][0] + 1, metadata["quote_name"][1] + 1) if "quote_name" in metadata else None,
            "total_cell": (metadata["total"][0] + 1, metadata["total"][1] + 1) if "total" in metadata else None,
        }

    if not plans:
        raise ValueError(
            "Could not find a usable quote table with 'SKU' and 'Unit List Price' "
            "in any worksheet. Make sure this is a Cisco renewal quote export."
        )

    # Case-insensitive fallback index -- some quote exports don't match the
    # master's SKU casing exactly.
    master_upper = {k.upper(): v for k, v in master.items()}

    # A read-only, formula-EVALUATED snapshot (Excel's last-saved cached
    # values) used only as a fallback when a cell we need to read for the
    # running total turns out to be a live formula rather than a static
    # number (openpyxl can't evaluate formulas itself).
    wb_cached = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)

    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=False)
    matched_lines = 0
    sheet_totals: dict[str, float] = {}

    for sheet, plan in plans.items():
        ws = wb[sheet]
        ws_cached = wb_cached[sheet]
        first_data_row = plan["header_excel_row"] + 1
        new_total = 0.0
        for row in range(first_data_row, ws.max_row + 1):
            sku_val = ws.cell(row=row, column=plan["sku_col"]).value
            sku = str(sku_val).strip() if sku_val is not None else ""
            prorated_cell = ws.cell(row=row, column=plan["prorated_col"]) if plan["prorated_col"] else None

            m = master.get(sku) or (master_upper.get(sku.upper()) if sku else None)
            if m is not None:
                ulp_cell = ws.cell(row=row, column=plan["ulp_col"])
                try:
                    old_ulp = float(ulp_cell.value)
                except (TypeError, ValueError):
                    old_ulp = None
                if old_ulp is not None and old_ulp > 0:
                    ulp_cell.value = old_ulp * m["ratio"]
                    matched_lines += 1
                    if prorated_cell is not None:
                        try:
                            old_prorated = float(prorated_cell.value)
                        except (TypeError, ValueError):
                            old_prorated = None
                        if old_prorated is not None:
                            prorated_cell.value = old_prorated * m["ratio"]

            if prorated_cell is not None:
                try:
                    new_total += float(prorated_cell.value or 0)
                except (TypeError, ValueError):
                    # Formula cell we didn't touch (unmatched line) -- fall
                    # back to Excel's last cached result so the recalculated
                    # total doesn't silently undercount that line.
                    cached_val = ws_cached.cell(row=row, column=plan["prorated_col"]).value
                    try:
                        new_total += float(cached_val or 0)
                    except (TypeError, ValueError):
                        pass

        if plan["quote_name_cell"]:
            r, c = plan["quote_name_cell"]
            cell = ws.cell(row=r, column=c)
            if cell.data_type != "f":
                current = str(cell.value or "").strip()
                if current and not current.endswith(_ESTIMATE_SUFFIX):
                    cell.value = current + _ESTIMATE_SUFFIX

        if plan["total_cell"]:
            r, c = plan["total_cell"]
            total_cell = ws.cell(row=r, column=c)
            if total_cell.data_type != "f":
                total_cell.value = new_total

        sheet_totals[sheet] = new_total

    # ── Update a "Summary" tab (quote-number -> total-value pairs), if present ──
    # Some templates make this tab a LIVE formula (e.g. ='<quote>'!C7) pointing
    # back at each quote tab's own total -- those auto-recalculate in Excel once
    # the source cell is updated above, so touching them would only strip the
    # formula. Only overwrite genuinely static (non-formula) Summary rows.
    for sheet in wb.sheetnames:
        if sheet.strip().lower() != "summary":
            continue
        ws = wb[sheet]
        for row in range(1, ws.max_row + 1):
            key_cell = ws.cell(row=row, column=1)
            if key_cell.data_type == "f":
                continue
            key = str(key_cell.value).strip() if key_cell.value is not None else ""
            if key not in sheet_totals:
                continue
            val_cell = ws.cell(row=row, column=2)
            if val_cell.data_type != "f":
                val_cell.value = sheet_totals[key]

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue(), {"matched_lines": matched_lines, "sheets_updated": list(plans.keys())}


# ── Impact computation ────────────────────────────────────────────────────────
def compute_impact(master: dict[str, dict], quote: pd.DataFrame) -> pd.DataFrame:
    """
    Match quote lines to master SKUs and compute prorated impact.
    Returns a DataFrame with one row per quote line (matched and unmatched).
    """
    rows = []
    for _, r in quote.iterrows():
        sku = r["sku"]
        ulp = float(r["unit_list_price"])
        qty = float(r["quantity"])
        sd = r["start_date"]
        ed = r["end_date"]

        if pd.isna(sd) or pd.isna(ed):
            days = None
            factor = None
        else:
            # Clip to 0: guards against start_date > end_date (data entry error)
            # which would otherwise invert the prorated price silently.
            days = max(0, int((ed - sd).days))
            factor = (days + 1) / 365.0

        m = master.get(sku)
        matched = m is not None

        if factor is not None:
            orig_prorated = ulp * qty * factor
        else:
            orig_prorated = None

        if matched and factor is not None:
            new_ulp = ulp * m["ratio"]
            new_prorated = new_ulp * qty * factor
            change = new_prorated - orig_prorated
        else:
            new_ulp = None
            new_prorated = None
            change = None

        rows.append({
            "matched": matched,
            "quote_name":   str(r.get("quote_name",   "") or ""),
            "quote_number": str(r.get("quote_number", "") or ""),
            "source_file":  str(r.get("source_file",  "") or ""),
            "source_sheet": str(r.get("source_sheet", "") or ""),
            "sku": sku,
            "product_number": r.get("product_number", ""),
            "product_description": r.get("product_description", ""),
            "service_level": r.get("service_level", ""),
            "serial_number": r.get("serial_number", ""),
            "instance_number": r.get("instance_number", ""),
            "start_date": sd,
            "end_date": ed,
            "days": days,
            "prorate_factor": factor,
            "quantity": qty,
            "orig_unit_list": ulp,
            "new_unit_list": new_ulp,
            "master_old": m["old"] if matched else None,
            "master_final": m["final"] if matched else None,
            "pct_change": (m["ratio"] - 1) if matched else None,
            "orig_prorated": orig_prorated,
            "new_prorated": new_prorated,
            "prorated_change": change,
        })
    return pd.DataFrame(rows)


# ── Excel report writer ───────────────────────────────────────────────────────
def write_impact_report(detail: pd.DataFrame, customer_name: str = "") -> bytes:
    """
    Build a 3-tab Excel report:
      1. Summary by SKU
      2. Line Detail
      3. Totals
    """
    buf = BytesIO()
    wb = xlsxwriter.Workbook(buf, {"nan_inf_to_errors": True})

    # formats
    hdr_fmt  = wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF",
                               "border": 1, "align": "center"})
    money_fmt = wb.add_format({"num_format": "#,##0.00", "border": 1})
    pct_fmt   = wb.add_format({"num_format": "0.00%", "border": 1, "align": "center"})
    int_fmt   = wb.add_format({"num_format": "#,##0", "border": 1, "align": "center"})
    text_fmt  = wb.add_format({"border": 1})
    date_fmt  = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1, "align": "center"})
    pos_money = wb.add_format({"num_format": "#,##0.00", "border": 1, "font_color": "#006100"})
    neg_money = wb.add_format({"num_format": "#,##0.00", "border": 1, "font_color": "#9C0006"})
    total_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 2,
                               "num_format": "#,##0.00"})
    total_pct = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 2,
                               "num_format": "0.00%", "align": "center"})
    total_lbl = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 2})

    matched = detail[detail["matched"]].copy()

    # ── Tab 1b: Summary by Quote ──────────────────────────────────────────────
    ws0 = wb.add_worksheet("Summary by Quote")
    ws0.set_tab_color("#2E75B6")
    quote_cols = [
        ("Quote Name", 28), ("Quote Number", 18), ("Source File", 30), ("Sheet / Tab", 22),
        ("Total Lines", 12), ("Matched Lines", 13),
        ("Full Quote Prorated ($)", 22), ("Matched Prorated ($)", 20),
        ("New Prorated ($)", 18), ("Price Impact ($)", 18), ("Impact %", 12),
    ]
    for c, (h, w) in enumerate(quote_cols):
        ws0.write(0, c, h, hdr_fmt)
        ws0.set_column(c, c, w)

    # group by (quote_name, quote_number, source_file, source_sheet) so each
    # distinct quote — including one-quote-per-tab workbooks — is its own row
    gkey = []
    for col in ("quote_name", "quote_number", "source_file", "source_sheet"):
        if col in detail.columns:
            gkey.append(col)
    if not gkey:
        gkey = ["source_file"] if "source_file" in detail.columns else []

    if gkey:
        q_summary_rows = []
        for keys, grp_df in detail.groupby(gkey, sort=False, dropna=False):
            if not isinstance(keys, tuple):
                keys = (keys,)
            key_dict = dict(zip(gkey, keys))
            matched_g = grp_df[grp_df["matched"]]
            full_p  = grp_df["orig_prorated"].sum()
            match_p = matched_g["orig_prorated"].sum()
            new_p   = matched_g["new_prorated"].sum()
            chg     = new_p - match_p
            pct     = chg / match_p if match_p else 0.0
            q_summary_rows.append({
                "quote_name":   key_dict.get("quote_name", ""),
                "quote_number": key_dict.get("quote_number", ""),
                "source_file":  key_dict.get("source_file", ""),
                "source_sheet": key_dict.get("source_sheet", ""),
                "total_lines":  len(grp_df),
                "matched_lines": len(matched_g),
                "full_prorated": full_p,
                "match_prorated": match_p,
                "new_prorated": new_p,
                "change": chg,
                "pct": pct,
            })
        # sort by impact descending
        q_summary_rows.sort(key=lambda r: r["change"], reverse=True)
        for row_i, r in enumerate(q_summary_rows, start=1):
            ch_fmt2 = pos_money if r["change"] >= 0 else neg_money
            int_f = wb.add_format({"num_format": "#,##0", "border": 1, "align": "center"})
            ws0.write(row_i, 0, r["quote_name"],   text_fmt)
            ws0.write(row_i, 1, r["quote_number"],  text_fmt)
            ws0.write(row_i, 2, r["source_file"],   text_fmt)
            ws0.write(row_i, 3, r["source_sheet"],  text_fmt)
            ws0.write(row_i, 4, int(r["total_lines"]),   int_f)
            ws0.write(row_i, 5, int(r["matched_lines"]), int_f)
            ws0.write(row_i, 6, r["full_prorated"],  money_fmt)
            ws0.write(row_i, 7, r["match_prorated"], money_fmt)
            ws0.write(row_i, 8, r["new_prorated"],   money_fmt)
            ws0.write(row_i, 9, r["change"],          ch_fmt2)
            ws0.write(row_i, 10, r["pct"],            pct_fmt)
        # totals row
        if q_summary_rows:
            tr = len(q_summary_rows) + 1
            ws0.write(tr, 0, "TOTAL", total_lbl)
            for c in range(1, 6): ws0.write(tr, c, "", total_lbl)
            ws0.write(tr, 6, sum(r["full_prorated"]  for r in q_summary_rows), total_fmt)
            ws0.write(tr, 7, sum(r["match_prorated"] for r in q_summary_rows), total_fmt)
            ws0.write(tr, 8, sum(r["new_prorated"]   for r in q_summary_rows), total_fmt)
            total_chg = sum(r["change"] for r in q_summary_rows)
            ws0.write(tr, 9, total_chg,
                      wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 2,
                                     "num_format": "#,##0.00",
                                     "font_color": "#006100" if total_chg >= 0 else "#9C0006"}))
            ws0.write(tr, 10, "", total_lbl)

    # ── Tab 2: Summary by SKU ──────────────────────────────────────────────────
    ws1 = wb.add_worksheet("Summary by SKU")
    ws1.set_tab_color("#1F3864")
    summary_cols = [
        ("Service SKU", 22),
        ("Master Old Price", 16),
        ("Master Final Price", 16),
        ("Price Change %", 14),
        ("Matched Lines", 13),
        ("Total Qty", 12),
        ("Orig Prorated Total", 18),
        ("New Prorated Total", 18),
        ("Prorated Change", 16),
    ]
    for c, (h, w) in enumerate(summary_cols):
        ws1.write(0, c, h, hdr_fmt)
        ws1.set_column(c, c, w)

    if not matched.empty:
        grp = matched.groupby("sku", sort=False).agg(
            master_old=("master_old", "first"),
            master_final=("master_final", "first"),
            pct_change=("pct_change", "first"),
            lines=("sku", "count"),
            qty=("quantity", "sum"),
            orig=("orig_prorated", "sum"),
            new=("new_prorated", "sum"),
            change=("prorated_change", "sum"),
        ).reset_index().sort_values("change", ascending=False)

        for row_i, (_, r) in enumerate(grp.iterrows(), start=1):
            ch_fmt = pos_money if r["change"] >= 0 else neg_money
            ws1.write(row_i, 0, r["sku"], text_fmt)
            ws1.write(row_i, 1, r["master_old"], money_fmt)
            ws1.write(row_i, 2, r["master_final"], money_fmt)
            ws1.write(row_i, 3, r["pct_change"], pct_fmt)
            ws1.write(row_i, 4, int(r["lines"]), int_fmt)
            ws1.write(row_i, 5, int(r["qty"]), int_fmt)
            ws1.write(row_i, 6, r["orig"], money_fmt)
            ws1.write(row_i, 7, r["new"], money_fmt)
            ws1.write(row_i, 8, r["change"], ch_fmt)

        # Totals row for G / H / I
        total_row = len(grp) + 1
        total_row_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 2,
                                       "num_format": "#,##0.00"})
        total_lbl_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 2})
        ch_total = grp["change"].sum()
        ch_total_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 2,
                                      "num_format": "#,##0.00",
                                      "font_color": "#006100" if ch_total >= 0 else "#9C0006"})
        for c in range(6):  # A-F blank
            ws1.write(total_row, c, "" if c > 0 else "TOTAL", total_lbl_fmt)
        ws1.write(total_row, 6, grp["orig"].sum(), total_row_fmt)
        ws1.write(total_row, 7, grp["new"].sum(), total_row_fmt)
        ws1.write(total_row, 8, ch_total, ch_total_fmt)

    # ── Tab 2: Line Detail ────────────────────────────────────────────────────
    ws2 = wb.add_worksheet("Line Detail")
    ws2.set_tab_color("#375623")
    detail_cols = [
        ("Quote Name", 26), ("Quote Number", 16), ("Source File", 26),
        ("SKU", 22), ("Product Number", 20), ("Product Description", 30),
        ("Service Level", 18), ("Serial Number", 20), ("Instance Number", 20),
        ("Start Date", 13), ("End Date", 13), ("Days", 8), ("Prorate Factor", 13),
        ("Quantity", 10), ("Orig Unit List", 16), ("New Unit List", 16),
        ("Price Change %", 14), ("Orig Prorated", 16), ("New Prorated", 16),
        ("Prorated Change", 16), ("Matched", 10),
    ]
    for c, (h, w) in enumerate(detail_cols):
        ws2.write(0, c, h, hdr_fmt)
        ws2.set_column(c, c, w)

    for row_i, (_, r) in enumerate(detail.iterrows(), start=1):
            ch_fmt = (pos_money if r["prorated_change"] and r["prorated_change"] >= 0
                      else neg_money) if r["matched"] else money_fmt
            ws2.write(row_i, 0, str(r.get("quote_name",   "") or ""), text_fmt)
            ws2.write(row_i, 1, str(r.get("quote_number", "") or ""), text_fmt)
            ws2.write(row_i, 2, str(r.get("source_file",  "") or ""), text_fmt)
            ws2.write(row_i, 3, r["sku"], text_fmt)
            ws2.write(row_i, 4, str(r["product_number"]), text_fmt)
            ws2.write(row_i, 5, str(r["product_description"]), text_fmt)
            ws2.write(row_i, 6, str(r["service_level"]), text_fmt)
            ws2.write(row_i, 7, str(r["serial_number"]), text_fmt)
            ws2.write(row_i, 8, str(r["instance_number"]), text_fmt)
            if pd.notna(r["start_date"]):
                ws2.write_datetime(row_i, 9, r["start_date"].to_pydatetime(), date_fmt)
            else:
                ws2.write_blank(row_i, 9, None, date_fmt)
            if pd.notna(r["end_date"]):
                ws2.write_datetime(row_i, 10, r["end_date"].to_pydatetime(), date_fmt)
            else:
                ws2.write_blank(row_i, 10, None, date_fmt)
            ws2.write(row_i, 11, r["days"] if r["days"] is not None else "", int_fmt)
            ws2.write(row_i, 12, r["prorate_factor"] if r["prorate_factor"] is not None else "", money_fmt)
            ws2.write(row_i, 13, int(r["quantity"]), int_fmt)
            ws2.write(row_i, 14, r["orig_unit_list"], money_fmt)
            ws2.write(row_i, 15, r["new_unit_list"] if r["new_unit_list"] is not None else "", money_fmt)
            ws2.write(row_i, 16, r["pct_change"] if r["pct_change"] is not None else "", pct_fmt)
            ws2.write(row_i, 17, r["orig_prorated"] if r["orig_prorated"] is not None else "", money_fmt)
            ws2.write(row_i, 18, r["new_prorated"] if r["new_prorated"] is not None else "", money_fmt)
            ws2.write(row_i, 19, r["prorated_change"] if r["prorated_change"] is not None else "", ch_fmt)
            ws2.write(row_i, 20, "Yes" if r["matched"] else "No", text_fmt)

    ws2.autofilter(0, 0, len(detail), len(detail_cols) - 1)
    ws2.freeze_panes(1, 0)

    # ── Tab 3: Totals ─────────────────────────────────────────────────────────
    ws3 = wb.add_worksheet("Totals")
    ws3.set_tab_color("#C00000")
    ws3.set_column(0, 0, 34)
    ws3.set_column(1, 1, 18)

    total_lines   = len(detail)
    matched_lines = int(matched.shape[0])
    total_skus    = int(matched["sku"].nunique()) if not matched.empty else 0
    quote_orig_total  = float(detail["orig_prorated"].sum()) if not detail.empty else 0.0
    orig_total    = float(matched["orig_prorated"].sum()) if not matched.empty else 0.0
    new_total     = float(matched["new_prorated"].sum()) if not matched.empty else 0.0
    change_total  = new_total - orig_total
    pct_total     = change_total / orig_total if orig_total else 0.0

    rows_t = [
        ("Customer / Quote", customer_name or "—"),
        ("Total Lines in Quote", total_lines),
        ("Matched Lines (SKU in master)", matched_lines),
        ("Unmatched Lines (no price change)", total_lines - matched_lines),
        ("Unique SKUs Impacted", total_skus),
        (None, None),
        ("Full Quote Prorated Total — all lines ($)", quote_orig_total),
        ("  of which: Matched Lines Prorated Total ($)", orig_total),
        ("  of which: Unmatched Lines Prorated Total ($)", quote_orig_total - orig_total),
        (None, None),
        ("New Prorated Total — matched lines ($)", new_total),
        ("Price Impact / Prorated Change ($)", change_total),
        ("Price Impact / Prorated Change (%)", pct_total),
    ]
    ws3.write(0, 0, "Metric", hdr_fmt)
    ws3.write(0, 1, "Value", hdr_fmt)
    for row_i, (label, value) in enumerate(rows_t, start=1):
        if label is None:
            continue
        ws3.write(row_i, 0, label, total_lbl)
        if isinstance(value, str):
            ws3.write(row_i, 1, value, total_lbl)
        elif isinstance(value, int):
            ws3.write(row_i, 1, value, wb.add_format({"bold": True, "bg_color": "#D9E1F2",
                                                       "border": 2, "num_format": "#,##0",
                                                       "align": "center"}))
        elif label.endswith("(%)"):
            ws3.write(row_i, 1, value, total_pct)
        else:
            ws3.write(row_i, 1, value, total_fmt)

    wb.close()
    return buf.getvalue()
