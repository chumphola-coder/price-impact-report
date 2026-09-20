"""
tier_impact_engine.py
======================
Service Tier Change Impact Analysis engine.

Detects when a matched renewal line (same serial number + instance number,
last year vs this year) moved its "Allocated Service Program" to one of
Cisco's consolidated support tiers -- CX L1, CX L2, Partner ST L1, or
Partner ST L2 -- and calculates the resulting list price impact.

Two reference files drive the classification:
  - egrid.xlsx  ("Service Level Hierarchy"): maps a quote line's Service Level
    code (its CONTRACT TYPE) to an Allocated Service Program. Primary source.
  - glasia.xlsx (ASIA-PAC Global Price List): fallback lookup by Service SKU
    (its Product column) when a Service Level code isn't found in egrid,
    using its own Service Program column.

All four target tiers only ever appear under Service Category = "TECHNICAL
SUPPORT SERVICES" in egrid.xlsx (verified against the reference file), so
checking the resulting tier against the target set alone is sufficient -- no
separate category filter is required.
"""

from __future__ import annotations

from io import BytesIO

import numpy as np
import pandas as pd
import xlsxwriter

from compare_engine import TS_LINE_MATCH_KEYS, _aggregate_lines, _is_sw, add_calculated_fields
from service_tier_classifier import (
    DEFAULT_EGRID_PATH,
    DEFAULT_GLASIA_PATH,
    TARGET_TIERS,
    abbreviate_tier,
    classify_line as _classify_line_full,
    load_egrid,
    load_glasia,
)


# ── Impact computation ────────────────────────────────────────────────────────
def compute_tier_impact(last_df: pd.DataFrame, this_df: pd.DataFrame,
                        egrid_lookup: dict, glasia_lookup: dict) -> dict[str, pd.DataFrame]:
    """
    Match last-year and this-year TS quote lines by serial number + instance
    number (same identity rule as the main Renewal Comparison tool), classify
    each side's tier, and split the matched lines into:
      - tier_changes:     SKU changed, landed on a target tier (CX L1/L2,
                          Partner ST L1/L2), tier actually changed.
      - sku_price_changes: SKU unchanged, unit list price changed.
      - sla_changes:      SKU changed, but not a qualifying tier change.
      - unresolved:       SKU changed, but one/both sides couldn't be
                          classified by either reference file.
    """
    empty_cols = TS_LINE_MATCH_KEYS + ["service_sku_last", "service_sku_this"]
    empty = pd.DataFrame(columns=empty_cols)
    result = {"tier_changes": empty, "sku_price_changes": empty, "sla_changes": empty, "unresolved": empty}

    last_calc = add_calculated_fields(last_df)
    this_calc = add_calculated_fields(this_df)
    last_ts = last_calc[~_is_sw(last_calc)].copy()
    this_ts = this_calc[~_is_sw(this_calc)].copy()
    if last_ts.empty or this_ts.empty:
        return result

    last_lines = _aggregate_lines(last_ts)
    this_lines = _aggregate_lines(this_ts)
    merged = last_lines.merge(this_lines, on=TS_LINE_MATCH_KEYS, how="inner", suffixes=("_last", "_this"))
    if merged.empty:
        return result

    tier_last, sla_last, source_last, tier_this, sla_this, source_this = [], [], [], [], [], []
    for _, r in merged.iterrows():
        t, sla, s = _classify_line_full(r.get("service_level_last"), r.get("service_sku_last"), egrid_lookup, glasia_lookup)
        tier_last.append(t); sla_last.append(sla); source_last.append(s)
        t, sla, s = _classify_line_full(r.get("service_level_this"), r.get("service_sku_this"), egrid_lookup, glasia_lookup)
        tier_this.append(t); sla_this.append(sla); source_this.append(s)
    merged["tier_last"] = tier_last
    merged["sla_last"] = sla_last
    merged["tier_source_last"] = source_last
    merged["tier_this"] = tier_this
    merged["sla_this"] = sla_this
    merged["tier_source_this"] = source_this

    for col in ("unit_list_price_last", "unit_list_price_this",
               "calculated_extended_list_price_last", "calculated_extended_list_price_this"):
        merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0)

    merged["unit_price_change"] = merged["unit_list_price_this"] - merged["unit_list_price_last"]
    merged["unit_price_change_pct"] = np.where(
        merged["unit_list_price_last"].ne(0),
        merged["unit_price_change"] / merged["unit_list_price_last"],
        np.nan,
    )
    merged["extended_value_change"] = (
        merged["calculated_extended_list_price_this"] - merged["calculated_extended_list_price_last"]
    )

    sku_changed = (
        merged["service_sku_last"].fillna("").astype(str).str.strip()
        != merged["service_sku_this"].fillna("").astype(str).str.strip()
    )
    moved_to_target = merged["tier_this"].isin(TARGET_TIERS)
    tier_changed = merged["tier_last"].ne(merged["tier_this"]) & merged["tier_last"].ne("") & merged["tier_this"].ne("")
    unresolved_mask = sku_changed & (merged["tier_source_last"].eq("unresolved") | merged["tier_source_this"].eq("unresolved"))

    is_tier_change = sku_changed & moved_to_target & tier_changed & ~unresolved_mask
    is_sku_price_change = ~sku_changed & merged["unit_price_change"].ne(0)
    is_sla_change_other = sku_changed & ~is_tier_change & ~unresolved_mask

    def _label(sla: str, tier: str) -> str:
        tier_abbrev = abbreviate_tier(tier)
        return f"{sla} {tier_abbrev}".strip() if sla else tier_abbrev

    merged["tier_change_label"] = [
        f"{_label(sla_l, tier_l)} \u2192 {_label(sla_t, tier_t)}" if change else ""
        for change, sla_l, tier_l, sla_t, tier_t in zip(
            is_tier_change, merged["sla_last"], merged["tier_last"], merged["sla_this"], merged["tier_this"]
        )
    ]

    common_cols = TS_LINE_MATCH_KEYS + [
        "service_sku_last", "service_sku_this",
        "service_level_last", "service_level_this",
        "quantity_last", "quantity_this",
        "unit_list_price_last", "unit_list_price_this",
        "unit_price_change", "unit_price_change_pct",
        "calculated_extended_list_price_last", "calculated_extended_list_price_this",
        "extended_value_change",
    ]

    result["tier_changes"] = merged.loc[is_tier_change, common_cols + [
        "tier_last", "tier_this", "sla_last", "sla_this", "tier_change_label", "tier_source_last", "tier_source_this",
    ]].reset_index(drop=True)
    result["sku_price_changes"] = merged.loc[is_sku_price_change, common_cols].reset_index(drop=True)
    result["sla_changes"] = merged.loc[is_sla_change_other, common_cols + ["tier_last", "tier_this", "sla_last", "sla_this"]].reset_index(drop=True)
    result["unresolved"] = merged.loc[unresolved_mask, common_cols + [
        "tier_last", "tier_this", "tier_source_last", "tier_source_this",
    ]].reset_index(drop=True)
    return result


# ── Excel report writer ───────────────────────────────────────────────────────
_COLUMN_LABELS = {
    "serial_number": ("Serial Number", 20), "instance_number": ("Instance Number", 16),
    "service_sku_last": ("SKU (Last Year)", 22), "service_sku_this": ("SKU (This Year)", 22),
    "service_level_last": ("Service Level (Last Year)", 20), "service_level_this": ("Service Level (This Year)", 20),
    "quantity_last": ("Qty (Last Year)", 13), "quantity_this": ("Qty (This Year)", 13),
    "unit_list_price_last": ("Unit List Price (Last Year)", 18), "unit_list_price_this": ("Unit List Price (This Year)", 18),
    "unit_price_change": ("Unit Price Change", 16), "unit_price_change_pct": ("Unit Price Change %", 16),
    "calculated_extended_list_price_last": ("Extended List (Last Year)", 18),
    "calculated_extended_list_price_this": ("Extended List (This Year)", 18),
    "extended_value_change": ("Extended Value Change", 18),
    "tier_last": ("Tier (Last Year)", 20), "tier_this": ("Tier (This Year)", 20),
    "sla_last": ("SLA (Last Year)", 16), "sla_this": ("SLA (This Year)", 16),
    "tier_change_label": ("Tier Change", 28),
    "tier_source_last": ("Tier Source (Last Year)", 14), "tier_source_this": ("Tier Source (This Year)", 14),
}


def _write_sheet(wb, name: str, df: pd.DataFrame, formats: dict) -> None:
    ws = wb.add_worksheet(name)
    if df.empty:
        ws.write(0, 0, "No lines in this category.", formats["text"])
        return
    for c, col in enumerate(df.columns):
        label, width = _COLUMN_LABELS.get(col, (col, 16))
        ws.write(0, c, label, formats["hdr"])
        ws.set_column(c, c, width)
    for r, (_, row) in enumerate(df.iterrows(), start=1):
        for c, col in enumerate(df.columns):
            val = row[col]
            if col.startswith("unit_list_price") or col == "unit_price_change" or col.startswith("calculated_extended") or col == "extended_value_change":
                ws.write(r, c, float(val) if pd.notna(val) else 0.0, formats["money"])
            elif col == "unit_price_change_pct":
                ws.write(r, c, float(val) if pd.notna(val) else 0.0, formats["pct"])
            elif col.startswith("quantity"):
                ws.write(r, c, float(val) if pd.notna(val) else 0.0, formats["int"])
            else:
                ws.write(r, c, "" if pd.isna(val) else str(val), formats["text"])
    ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    ws.freeze_panes(1, 0)


def write_tier_impact_report(results: dict[str, pd.DataFrame], customer_name: str = "") -> bytes:
    buf = BytesIO()
    wb = xlsxwriter.Workbook(buf, {"nan_inf_to_errors": True})
    formats = {
        "hdr": wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF", "border": 1, "align": "center"}),
        "money": wb.add_format({"num_format": "#,##0.00", "border": 1}),
        "pct": wb.add_format({"num_format": "0.00%", "border": 1, "align": "center"}),
        "int": wb.add_format({"num_format": "#,##0", "border": 1, "align": "center"}),
        "text": wb.add_format({"border": 1}),
    }

    ws0 = wb.add_worksheet("Summary")
    ws0.set_column(0, 0, 32)
    ws0.set_column(1, 1, 18)
    ws0.write(0, 0, "Service Tier Change Impact Analysis", formats["hdr"])
    ws0.write(0, 1, customer_name, formats["hdr"])
    tc = results["tier_changes"]
    row = 2
    ws0.write(row, 0, "Tier Change lines", formats["text"])
    ws0.write(row, 1, len(tc), formats["int"])
    row += 1
    ws0.write(row, 0, "Total extended value change from tier changes ($)", formats["text"])
    ws0.write(row, 1, float(tc["extended_value_change"].sum()) if not tc.empty else 0.0, formats["money"])
    row += 1
    ws0.write(row, 0, "SKU Price Change lines (no tier change)", formats["text"])
    ws0.write(row, 1, len(results["sku_price_changes"]), formats["int"])
    row += 1
    ws0.write(row, 0, "Other SLA Change lines (SKU changed, not a target tier)", formats["text"])
    ws0.write(row, 1, len(results["sla_changes"]), formats["int"])
    row += 1
    ws0.write(row, 0, "Unresolved lines (could not classify tier)", formats["text"])
    ws0.write(row, 1, len(results["unresolved"]), formats["int"])

    _write_sheet(wb, "Tier Changes", results["tier_changes"], formats)
    _write_sheet(wb, "SKU Price Changes", results["sku_price_changes"], formats)
    _write_sheet(wb, "SLA Changes", results["sla_changes"], formats)
    _write_sheet(wb, "Unresolved", results["unresolved"], formats)

    wb.close()
    return buf.getvalue()
