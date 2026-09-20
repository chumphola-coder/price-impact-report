from __future__ import annotations

from datetime import date

import numpy as np
import pandas as pd


MATCH_KEYS = ["serial_number", "instance_number", "service_sku"]
# TS line identity for year-over-year matching: serial# + instance# only.
# service_sku is deliberately excluded here because Cisco encodes the SLA/
# service-level in the SKU itself (e.g. CON-SNTP-... -> CON-L24HR-...) — the
# same physical instance keeps its instance_number across an SLA change even
# though the SKU changes. Matching on the 3-key MATCH_KEYS (used elsewhere for
# SNIFF enrichment) would misclassify an SLA upgrade/downgrade as a removed
# old-SKU line plus an added new-SKU line instead of one matched line with an
# SLA change.
TS_LINE_MATCH_KEYS = ["serial_number", "instance_number"]
SNIFF_ENRICH_COLUMNS = [
    "asset_status",
    "sniff_start_date",
    "sniff_end_date",
    "parent_serial_number",
    "installed_base_status",
    "replacement_serial_number",
    "replacement_instance_number",
    "product_family",
    "product_group",
    "product_sub_type",
    "item_type",
    "warranty_type",
    "warranty_status",
    "warranty_end_date",
    "termination_date",
    "renewed_from_instance",
    "renewed_to_instance",
    "price_list",
]

SUMMARY_COLUMNS = ["section", "metric", "value svc ($)", "value SW ($)", "Total $", "date", "formula / note"]


def add_calculated_fields(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    # Ensure all date columns are datetime so .dt operations never fail on
    # object/string dtype (varies by source quote export).
    for _dcol in ("start_date", "end_date", "ldos_date"):
        if _dcol in out.columns:
            out[_dcol] = pd.to_datetime(out[_dcol], errors="coerce")
    effective_end = out["end_date"].copy()
    has_ldos = out["ldos_date"].notna()
    effective_end.loc[has_ldos] = pd.concat([out.loc[has_ldos, "end_date"], out.loc[has_ldos, "ldos_date"]], axis=1).min(axis=1)

    out["effective_end_date"] = effective_end
    out["ldos_capped"] = out["ldos_date"].notna() & out["end_date"].notna() & (out["end_date"] > out["ldos_date"])

    full_days = (out["end_date"] - out["start_date"]).dt.days.clip(lower=0)
    eff_days = (out["effective_end_date"] - out["start_date"]).dt.days.clip(lower=0)
    out["duration_days"] = eff_days

    ulp = pd.to_numeric(out["unit_list_price"], errors="coerce").fillna(0)
    qty = pd.to_numeric(out["quantity"], errors="coerce").fillna(0)
    prorated = pd.to_numeric(out.get("prorated_list_price", pd.Series(index=out.index, dtype=float)), errors="coerce")

    # Extended list price basis:
    #   Prefer the quote's own prorated_list_price so totals match the quote
    #   exactly. If coverage is capped at LDOS, scale the prorated value down by
    #   the capped/full day ratio. Fall back to Cisco's proration formula
    #   unit_list_price × quantity × (inclusive coverage days / 365) when
    #   prorated is unavailable — inclusive = (end − start) + 1, so a full year
    #   (364 exclusive days) prorates to exactly 1.0.
    cap_ratio = np.divide(
        eff_days.to_numpy(),
        full_days.to_numpy(),
        out=np.ones_like(eff_days.to_numpy(), dtype=float),
        where=full_days.to_numpy() > 0
    )
    cap_ratio = np.clip(cap_ratio, 0.0, 1.0)
    value_from_prorated = prorated.fillna(0).to_numpy() * cap_ratio
    inclusive_days = np.where(eff_days.to_numpy() > 0, eff_days.to_numpy() + 1, 0)
    value_from_formula = (ulp * qty).to_numpy() * (inclusive_days / 365.0)
    use_prorated = (prorated.notna() & (prorated != 0)).to_numpy()
    out["calculated_extended_list_price"] = np.where(use_prorated, value_from_prorated, value_from_formula)

    # Effective annual duration implied by the value, used by the effect
    # decomposition so that value = unit_list_price × quantity × duration_years.
    denom = (ulp * qty).to_numpy()
    safe_denom = np.where(denom > 0, denom, 1.0)
    out["duration_years"] = np.where(
        denom > 0,
        out["calculated_extended_list_price"].to_numpy() / safe_denom,
        inclusive_days / 365.0,
    )
    out["duration_months"] = out["duration_years"] * 12
    return out


def _first_non_empty(series: pd.Series):
    for value in series:
        if pd.notna(value) and str(value).strip() != "":
            return value
    return pd.NA


def _duration_months(start: pd.Series, end: pd.Series) -> pd.Series:
    start = pd.to_datetime(start, errors="coerce")
    end = pd.to_datetime(end, errors="coerce")
    return ((end - start).dt.days.clip(lower=0) / 30).round(4)


def _normalize_change_details(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["duration_months"] = out.get("duration_months", pd.Series(index=out.index, dtype=float))
    if "duration_months" not in out or out["duration_months"].isna().all():
        date_start = "sniff_start_date" if "sniff_start_date" in out else "start_date"
        date_end = "sniff_end_date" if "sniff_end_date" in out else "effective_end_date"
        if date_start in out and date_end in out:
            out["duration_months"] = _duration_months(out[date_start], out[date_end])
    return out


# ── SW Subscription helpers ───────────────────────────────────────────────────

def _is_sw(df: pd.DataFrame) -> pd.Series:
    """Boolean mask: True for Sub-TnC software subscription lines."""
    if "service_type" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["service_type"].fillna("").astype(str).str.strip().str.upper().eq("SUB-TNC")


def _agg_sw_group(grp: pd.DataFrame) -> dict:
    """Reduce a group of SW sub rows to a single scalar-value dict.

    calculated_extended_list_price is summed across ALL lines (including
    zero-price) so the total dollar value is correct.

    qty and unit_list_price are derived from PRICED lines only (ulp > 0).
    Zero-price wrapper/housekeeping SKUs (e.g. C1-N9K-RENEW-T with ulp=0)
    carry no dollar value and must not inflate the aggregated quantity used
    in the Shapley decomposition.

    duration_years is back-calculated from SUM(ext) / (priced_qty × mean_ulp)
    so the Shapley bridge identity  qty × ulp × dur = ext  holds exactly.
    Falls back to mean(duration_years) for fully zero-price groups.
    """
    result: dict = {}

    # Total extended list price — sum ALL lines
    if "calculated_extended_list_price" in grp.columns:
        result["calculated_extended_list_price"] = float(grp["calculated_extended_list_price"].sum())

    # qty and ulp — priced lines only
    if "unit_list_price" in grp.columns:
        priced = grp[pd.to_numeric(grp["unit_list_price"], errors="coerce").fillna(0) > 0]
    else:
        priced = grp.iloc[0:0]
    if not priced.empty:
        result["quantity"]        = float(priced["quantity"].sum()) if "quantity" in priced.columns else 0.0
        result["unit_list_price"] = float(priced["unit_list_price"].mean())
    else:
        result["quantity"]        = float(grp["quantity"].sum()) if "quantity" in grp.columns else 0.0
        result["unit_list_price"] = 0.0

    # Effective duration: back-calculated so qty * ulp * dur == ext exactly
    ext   = result.get("calculated_extended_list_price", 0.0) or 0.0
    qty   = result.get("quantity", 0.0) or 0.0
    ulp   = result.get("unit_list_price", 0.0) or 0.0
    denom = qty * ulp
    if denom > 0.0:
        result["duration_years"] = ext / denom
    elif "duration_years" in grp.columns:
        result["duration_years"] = float(grp["duration_years"].mean())
    else:
        result["duration_years"] = 0.0
    result["duration_months"] = result["duration_years"] * 12
    for col in [
        "service_sku", "product_number", "product_description",
        "service_level", "service_level_description",
        "instance_number", "serial_number",
        "subscription_id", "reference_instance_number", "reference_serial_number",
        "start_date", "end_date", "sw_key_type",
    ]:
        if col in grp.columns:
            result[col] = _first_non_empty(grp[col])
    result["line_count"] = len(grp)
    return result


def _resolve_sw_match_key(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add sw_match_key and sw_key_type columns.
    Priority: instance_number → serial_number → subscription_id →
              reference_instance_number → reference_serial_number → "" (unresolved)

    Placeholder values (e.g. "NEW", "Not Applicable") are treated as no key, so
    brand-new subscription lines fall through to the unresolved path and are
    grouped by product number instead of collapsing under a single fake key.
    """
    out = df.copy()
    placeholders = {"", "new", "n/a", "na", "none", "not applicable", "tbd", "-", "nan"}

    def _clean_key(series: pd.Series) -> pd.Series:
        s = series.fillna("").astype(str).str.strip()
        return s.where(~s.str.lower().isin(placeholders), "")

    key = pd.Series("", index=out.index, dtype=str)
    key_type = pd.Series("none", index=out.index, dtype=str)
    for col, ktype in [
        ("instance_number", "instance"),
        ("serial_number", "serial"),
        ("subscription_id", "subscription_id"),
        ("reference_instance_number", "ref_instance"),
        ("reference_serial_number", "ref_serial"),
    ]:
        if col in out.columns:
            cleaned = _clean_key(out[col])
            populated = key.eq("") & cleaned.ne("")
            key = key.where(~populated, cleaned)
            key_type = key_type.where(~populated, ktype)
    out["sw_match_key"] = key
    out["sw_key_type"] = key_type
    return out


def _classify_sw_subs(last_sw: pd.DataFrame, this_sw: pd.DataFrame) -> pd.DataFrame:
    """
    Match and classify Sub-TnC lines between last year and this year.

    Pass 1 — key-resolved lines:
      Each line's match key is resolved via the priority chain in _resolve_sw_match_key.
      Lines sharing the same key are compared:
        • both years, same service_sku  → SW Matched
        • both years, diff service_sku  → SW Tier Change
        • only this year               → SW Added (key matched)
        • only last year               → SW Removed (key matched)

    Pass 2 — unresolved lines (no key found):
      Grouped by product_number and compared against the other year's unresolved set:
        • product not in other year    → SW Added (new product) / SW Removed (product gone)
        • product in both years        → SW Added/Removed (same product - review)

    Returns one row per matched group / unresolved product group with effect columns.
    """
    if last_sw.empty and this_sw.empty:
        return pd.DataFrame()

    last_r = _resolve_sw_match_key(last_sw.reset_index(drop=True))
    this_r = _resolve_sw_match_key(this_sw.reset_index(drop=True))

    rows: list[dict] = []

    # ── pass 1: key-resolved ─────────────────────────────────
    last_keyed = last_r[last_r["sw_match_key"].ne("")]
    this_keyed = this_r[this_r["sw_match_key"].ne("")]
    last_key_map = {k: g for k, g in last_keyed.groupby("sw_match_key", sort=False)}
    this_key_map = {k: g for k, g in this_keyed.groupby("sw_match_key", sort=False)}

    # ── Sub-ID rescue pass ────────────────────────────────────
    # Cisco historically used instance numbers as the primary identifier for some
    # SW TnC products and has since migrated to subscription IDs.  When a
    # last-year line was resolved to an instance key and the same-numbered
    # instance does not appear this year, it would normally become an orphan
    # "SW Removed" row.  But if the line's subscription_id matches a this-year
    # sub_id key, the subscription clearly continues — Cisco just stopped
    # exposing the instance number on the renewal.  In that case we merge the
    # orphan last-year rows into the sub_id-keyed match group so the
    # subscription is treated as "SW Matched" (with correct Shapley attribution)
    # rather than a spurious remove + re-add pair.
    _PLACEHOLDER_IDS = {"", "new", "n/a", "na", "none", "not applicable", "tbd", "-", "nan"}
    _rescued_instance_keys: set[str] = set()      # instance keys absorbed into a sub_id group
    _sub_id_extra_last: dict[str, pd.DataFrame] = {}  # sub_id → absorbed last-year rows

    for key, l_grp in last_key_map.items():
        if key in this_key_map:
            continue  # already matched by the same key — no rescue needed
        key_type = l_grp["sw_key_type"].iloc[0] if not l_grp.empty else ""
        if key_type not in ("instance", "serial"):
            continue
        if "subscription_id" not in l_grp.columns:
            continue
        sids = (
            l_grp["subscription_id"]
            .fillna("").astype(str).str.strip()
            .loc[lambda s: ~s.str.lower().isin(_PLACEHOLDER_IDS)]
            .unique()
        )
        for sid in sids:
            if sid in this_key_map:
                _rescued_instance_keys.add(key)
                if sid in _sub_id_extra_last:
                    _sub_id_extra_last[sid] = pd.concat(
                        [_sub_id_extra_last[sid], l_grp], ignore_index=True
                    )
                else:
                    _sub_id_extra_last[sid] = l_grp.copy()
                break  # only rescue to the first matching sub_id

    for key in set(last_key_map) | set(this_key_map):
        if key in _rescued_instance_keys:
            continue  # rows merged into a sub_id group above — skip standalone

        l_grp = last_key_map.get(key)
        t_grp = this_key_map.get(key)

        # Augment last-year group with any rescued instance/serial rows
        if key in _sub_id_extra_last:
            extra = _sub_id_extra_last[key]
            l_grp = pd.concat([l_grp, extra], ignore_index=True) if l_grp is not None else extra.copy()

        l_agg = _agg_sw_group(l_grp) if l_grp is not None else {}
        t_agg = _agg_sw_group(t_grp) if t_grp is not None else {}

        val_l = float(l_agg.get("calculated_extended_list_price", 0) or 0)
        val_t = float(t_agg.get("calculated_extended_list_price", 0) or 0)
        qty_l = float(l_agg.get("quantity", 0) or 0)
        qty_t = float(t_agg.get("quantity", 0) or 0)
        dur_l = float(l_agg.get("duration_years", 0) or 0)
        dur_t = float(t_agg.get("duration_years", 0) or 0)
        ulp_l = float(l_agg.get("unit_list_price", 0) or 0)
        ulp_t = float(t_agg.get("unit_list_price", 0) or 0)
        l_sku = str(l_agg.get("service_sku", ""))
        t_sku = str(t_agg.get("service_sku", ""))

        if l_grp is not None and t_grp is not None:
            status = "SW Tier Change" if l_sku and t_sku and l_sku != t_sku else "SW Matched"
            # Shapley (symmetric) decomposition: qty × duration × unit price
            shap_q = (qty_t - qty_l) * (dur_l * ulp_l / 3.0 + (dur_l * ulp_t + dur_t * ulp_l) / 6.0 + dur_t * ulp_t / 3.0)
            shap_d = (dur_t - dur_l) * (qty_l * ulp_l / 3.0 + (qty_l * ulp_t + qty_t * ulp_l) / 6.0 + qty_t * ulp_t / 3.0)
            shap_u = (ulp_t - ulp_l) * (qty_l * dur_l / 3.0 + (qty_l * dur_t + qty_t * dur_l) / 6.0 + qty_t * dur_t / 3.0)
            qty_eff = shap_q
            dur_eff = shap_d
            tier_eff = shap_u if status == "SW Tier Change" else 0.0
            price_eff = 0.0 if status == "SW Tier Change" else shap_u
            sw_add = 0.0
            sw_rm = 0.0
        elif t_grp is not None:
            status = "SW Added (key matched)"
            qty_eff = dur_eff = tier_eff = price_eff = 0.0
            sw_add = val_t
            sw_rm = 0.0
        else:
            status = "SW Removed (key matched)"
            qty_eff = dur_eff = tier_eff = price_eff = 0.0
            sw_add = 0.0
            sw_rm = -val_l

        row: dict = {
            "sw_match_key": key,
            "sw_key_type": l_agg.get("sw_key_type") or t_agg.get("sw_key_type") or "",
            "sw_status": status,
            "sw_added_value": sw_add,
            "sw_removed_value": sw_rm,
            "sw_qty_effect": qty_eff,
            "sw_duration_effect": dur_eff,
            "sw_tier_effect": tier_eff,
            "sw_unit_price_effect": price_eff,
            "sw_total_change": val_t - val_l,
        }
        for c, v in l_agg.items():
            row[f"{c}_last"] = v
        for c, v in t_agg.items():
            row[f"{c}_this"] = v
        rows.append(row)

    # ── pass 2: unresolved → product_number fallback ──────────
    last_unres = last_r[last_r["sw_match_key"].eq("")]
    this_unres = this_r[this_r["sw_match_key"].eq("")]
    last_unres_prods = set(last_unres["product_number"].fillna("").astype(str).str.strip()) - {""}
    this_unres_prods = set(this_unres["product_number"].fillna("").astype(str).str.strip()) - {""}

    # All last-year product_numbers (key-resolved AND unresolved), used so that
    # this-year lines with blank/NEW sub IDs whose product exists ANYWHERE in
    # last year are classified as "potential expired & buy as new" (D5) rather
    # than as a genuinely new subscription (D4).  A blank/NEW sub ID means the
    # renewal was placed early (still "signed" state) and hasn't been assigned a
    # proper subscription ID yet — it is almost certainly the continuation of an
    # existing contract, not a brand-new one.
    last_all_prods = set(last_r["product_number"].fillna("").astype(str).str.strip()) - {""}

    # Heuristic: (qty, unit_list_price) fingerprint of every last-year line
    # (only where ulp > 0 to avoid false matches on zero-price housekeeping
    # lines).  When a this-year unresolved line has a different product_number
    # but identical qty AND unit price, it is very likely the same physical
    # configuration renewed under a different contract term suffix (e.g.
    # ESS2-100G-SIA-1M → ESS2-100G-SIA-ST).  Cisco's quote export does not
    # expose the parent-child sub-TnC relationship, so qty+price is the most
    # reliable proxy available.
    last_all_qty_ulp: set[tuple] = set()
    for _, lr in last_r.iterrows():
        ulp = float(lr.get("unit_list_price") or 0)
        if ulp > 0:
            last_all_qty_ulp.add((
                round(float(lr.get("quantity") or 0), 0),
                round(ulp, 2),
            ))

    for prod, grp in this_unres.groupby("product_number", dropna=False, sort=False):
        prod_s = str(prod).strip() if pd.notna(prod) else ""
        t_agg = _agg_sw_group(grp)
        val_t = float(t_agg.get("calculated_extended_list_price", 0) or 0)
        t_ulp = float(t_agg.get("unit_list_price", 0) or 0)
        t_qty = float(t_agg.get("quantity", 0) or 0)
        qty_ulp_match = (
            t_ulp > 0
            and (round(t_qty, 0), round(t_ulp, 2)) in last_all_qty_ulp
        )
        if prod_s in last_unres_prods or prod_s in last_all_prods:
            status = "SW Added (same product - review)"   # exact product match last year
        elif qty_ulp_match:
            status = "SW Added (same product - review)"   # same qty+price → likely same config, different term suffix
        else:
            status = "SW Added (new product)"
        row = {
            "sw_match_key": prod_s or "(no key)",
            "sw_key_type": "product_number_fallback",
            "sw_status": status,
            "sw_added_value": val_t,
            "sw_removed_value": 0.0,
            "sw_qty_effect": 0.0,
            "sw_duration_effect": 0.0,
            "sw_tier_effect": 0.0,
            "sw_unit_price_effect": 0.0,
            "sw_total_change": val_t,
        }
        for c, v in t_agg.items():
            row[f"{c}_this"] = v
        rows.append(row)

    for prod, grp in last_unres.groupby("product_number", dropna=False, sort=False):
        prod_s = str(prod).strip() if pd.notna(prod) else ""
        l_agg = _agg_sw_group(grp)
        val_l = float(l_agg.get("calculated_extended_list_price", 0) or 0)
        status = "SW Removed (product gone)" if prod_s not in this_unres_prods else "SW Removed (same product - review)"
        row = {
            "sw_match_key": prod_s or "(no key)",
            "sw_key_type": "product_number_fallback",
            "sw_status": status,
            "sw_added_value": 0.0,
            "sw_removed_value": -val_l,
            "sw_qty_effect": 0.0,
            "sw_duration_effect": 0.0,
            "sw_tier_effect": 0.0,
            "sw_unit_price_effect": 0.0,
            "sw_total_change": -val_l,
        }
        for c, v in l_agg.items():
            row[f"{c}_last"] = v
        rows.append(row)

    return pd.DataFrame(rows) if rows else pd.DataFrame()


def _entity_changes(last_df: pd.DataFrame, this_df: pd.DataFrame, column: str, label: str) -> pd.DataFrame:
    detail_cols = [
        column,
        "product_number",
        "product_description",
        "quantity",
        "serial_number",
        "instance_number",
        "parent_instance_number",
        "parent_serial_number",
        "ldos_date",
        "unit_list_price",
        "duration_months",
        "service_type",
        "service_level",
        "service_level_description",
        "item_type",
        "product_family",
        "product_group",
        "replacement_serial_number",
        "asset_status",
    ]
    rows = []
    last_values = set(last_df[column].dropna().astype(str)) - {""}
    this_values = set(this_df[column].dropna().astype(str)) - {""}

    def add_rows(source_df: pd.DataFrame, values: set[str], change_type: str, period: str) -> None:
        subset = _normalize_change_details(source_df[source_df[column].astype(str).isin(values)])
        available_cols = [col for col in detail_cols if col in subset.columns]
        for value, group in subset.groupby(column, dropna=False):
            row = {label: value, "change_type": change_type, "period": period, "line_count": len(group)}
            for col in available_cols:
                if col == column:
                    continue
                row[col] = group[col].sum() if col == "quantity" else _first_non_empty(group[col])
            rows.append(row)

    add_rows(this_df, this_values - last_values, "Added", "This Year")
    add_rows(last_df, last_values - this_values, "Removed", "Last Year")
    add_rows(this_df, last_values & this_values, "Existing", "This Year")
    return pd.DataFrame(rows)


def _sniff_lookup(sniff_df: pd.DataFrame | None, suffix: str) -> pd.DataFrame | None:
    if sniff_df is None or sniff_df.empty:
        return None
    agg = {col: _first_non_empty for col in SNIFF_ENRICH_COLUMNS if col in sniff_df.columns}
    for col in ["sniff_start_date", "sniff_end_date", "ldos_date"]:
        if col in sniff_df.columns:
            agg[col] = _first_non_empty
    cols = MATCH_KEYS + [col for col in agg if col not in MATCH_KEYS]
    lookup = sniff_df[cols].groupby(MATCH_KEYS, dropna=False, as_index=False).agg(agg)
    return lookup.rename(columns={col: f"{col}_{suffix}" for col in lookup.columns if col not in MATCH_KEYS})


def enrich_quote_with_sniff(quote_df: pd.DataFrame, sniff_df: pd.DataFrame | None) -> pd.DataFrame:
    if sniff_df is None or sniff_df.empty:
        return quote_df.copy()
    lookup = _sniff_lookup(sniff_df, "sniff")
    if lookup is None:
        return quote_df.copy()
    out = quote_df.merge(lookup, on=MATCH_KEYS, how="left")
    for col in SNIFF_ENRICH_COLUMNS + ["sniff_start_date", "sniff_end_date"]:
        sniff_col = f"{col}_sniff"
        if sniff_col in out.columns:
            out[col] = out[sniff_col]
    # groupby.agg(_first_non_empty) can yield object dtype (Timestamp mixed with
    # pd.NA); coerce date columns back to datetime so downstream .dt is safe.
    for col in ("sniff_start_date", "sniff_end_date", "ldos_date"):
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce")
    return out


def _build_replacement_pairs(last_sniff_df: pd.DataFrame | None) -> pd.DataFrame:
    if last_sniff_df is None or last_sniff_df.empty or "replacement_serial_number" not in last_sniff_df.columns:
        return pd.DataFrame()
    wanted = [
        "serial_number",
        "replacement_serial_number",
        "instance_number",
        "replacement_instance_number",
        "product_number",
        "product_description",
    ]
    cols = [c for c in wanted if c in last_sniff_df.columns]
    return last_sniff_df[
        last_sniff_df["replacement_serial_number"].notna()
        & last_sniff_df["replacement_serial_number"].astype(str).ne("")
    ][cols].drop_duplicates()


def _apply_replacement_matching(this_df: pd.DataFrame, replacement_pairs: pd.DataFrame) -> pd.DataFrame:
    if replacement_pairs.empty:
        return this_df.copy()
    out = this_df.copy()
    serial_map = dict(zip(replacement_pairs["replacement_serial_number"].astype(str), replacement_pairs["serial_number"].astype(str)))
    instance_pairs = replacement_pairs[
        replacement_pairs["replacement_instance_number"].notna()
        & replacement_pairs["replacement_instance_number"].astype(str).ne("")
    ]
    instance_map = dict(zip(instance_pairs["replacement_instance_number"].astype(str), instance_pairs["instance_number"].astype(str)))

    out["actual_serial_number"] = out["serial_number"]
    out["actual_instance_number"] = out["instance_number"]
    out["replacement_note"] = ""
    serial_match = out["serial_number"].astype(str).isin(serial_map)
    out.loc[serial_match, "serial_number"] = out.loc[serial_match, "serial_number"].astype(str).map(serial_map)
    out.loc[serial_match, "replacement_note"] = "Replacement SN matched to original SN due to equipment failure"

    if instance_map:
        instance_match = out["instance_number"].astype(str).isin(instance_map)
        out.loc[instance_match, "instance_number"] = out.loc[instance_match, "instance_number"].astype(str).map(instance_map)
        out.loc[instance_match & out["replacement_note"].eq(""), "replacement_note"] = "Replacement instance matched to original instance due to equipment failure"
    return out


def _detect_overlaps(df: pd.DataFrame, period_label: str) -> pd.DataFrame:
    """Find SN#+instance#+SKU groups within one quote file whose coverage periods
    overlap. Split orders must be sequential (non-overlapping).

    Exact-duplicate periods (identical start AND end) are treated as data
    duplicates and de-duplicated silently — they are NOT reported here. Only
    genuine partial overlaps are flagged as errors for the user to fix.
    """
    if df.empty:
        return pd.DataFrame()
    end_col = "effective_end_date" if "effective_end_date" in df.columns else "end_date"
    rows = []
    for keys, group in df.groupby(MATCH_KEYS, dropna=False, sort=False):
        if len(group) < 2:
            continue
        sn, inst, sku = (keys if isinstance(keys, tuple) else (keys, "", ""))
        # Skip groups with no usable identity (all keys blank)
        if not any(str(k).strip() for k in (sn, inst, sku)):
            continue
        # Collect unique (start, end) intervals; exact duplicates collapse via set()
        intervals = set()
        for _, r in group.iterrows():
            s, e = r.get("start_date"), r.get(end_col)
            if pd.isna(s) or pd.isna(e):
                continue
            intervals.add((pd.Timestamp(s), pd.Timestamp(e)))
        ordered = sorted(intervals)
        for i in range(1, len(ordered)):
            prev_s, prev_e = ordered[i - 1]
            cur_s, cur_e = ordered[i]
            if cur_s < prev_e:  # starts before the previous period ends → overlap
                rows.append({
                    "period": period_label,
                    "serial_number": sn,
                    "instance_number": inst,
                    "service_sku": sku,
                    "period_1": f"{prev_s.date()} → {prev_e.date()}",
                    "period_2": f"{cur_s.date()} → {cur_e.date()}",
                    "issue": "Overlapping coverage periods (split orders must be sequential, not overlapping)",
                })
    return pd.DataFrame(rows)


def _union_duration_years(starts, ends) -> float:
    """Total covered duration in years from the UNION of [start, end] intervals,
    using the 30-day-month convention.

    Overlapping or duplicate periods are merged so they are not double-counted;
    sequential (split-order) periods are summed. This lets multiple lines for the
    same SN#+instance#+SKU be treated as one asset with a combined coverage span.
    """
    intervals = []
    for s, e in zip(starts, ends):
        if pd.isna(s) or pd.isna(e):
            continue
        s = pd.Timestamp(s)
        e = pd.Timestamp(e)
        if e < s:
            continue
        intervals.append((s, e))
    if not intervals:
        return 0.0
    intervals.sort()
    total_days = 0.0
    cur_s, cur_e = intervals[0]
    for s, e in intervals[1:]:
        if s <= cur_e:  # overlapping or contiguous → merge
            if e > cur_e:
                cur_e = e
        else:  # gap → close the current span and start a new one
            total_days += (cur_e - cur_s).days
            cur_s, cur_e = s, e
    total_days += (cur_e - cur_s).days
    return max(total_days, 0.0) / 30.0 / 12.0


def _aggregate_lines(df: pd.DataFrame) -> pd.DataFrame:
    group_cols = TS_LINE_MATCH_KEYS
    first_cols = [
        "service_sku",
        "product_number",
        "product_description",
        "service_level",
        "service_level_description",
        "service_type",
        "ldos_date",
        "start_date",
        "end_date",
        "effective_end_date",
        "ldos_capped",
        "parent_serial_number",
        "asset_status",
        "sniff_start_date",
        "sniff_end_date",
        "installed_base_status",
        "replacement_serial_number",
        "replacement_instance_number",
        "product_family",
        "product_group",
        "product_sub_type",
        "item_type",
        "warranty_type",
        "warranty_status",
        "warranty_end_date",
        "termination_date",
        "renewed_from_instance",
        "renewed_to_instance",
        "price_list",
        "actual_serial_number",
        "actual_instance_number",
        "replacement_note",
    ]
    if df.empty:
        return df.copy()

    end_col = "effective_end_date" if "effective_end_date" in df.columns else "end_date"

    def _combine(group: pd.DataFrame) -> dict:
        # De-duplicate: same SN#+instance#+SKU counts once. Use the representative
        # (max) quantity rather than summing split/duplicate lines.
        qty = float(pd.to_numeric(group["quantity"], errors="coerce").fillna(0).max())
        ulp = float(pd.to_numeric(group["unit_list_price"], errors="coerce").fillna(0).mean())
        # Value = sum of the (prorated-based) extended list price over UNIQUE
        # coverage periods. Exact-duplicate rows (same start+end) are counted
        # once; sequential split-order periods are summed. This keeps A/B totals
        # equal to the quote's prorated list price.
        uniq = group.drop_duplicates(subset=["start_date", end_col])
        value = float(pd.to_numeric(uniq["calculated_extended_list_price"], errors="coerce").fillna(0).sum())
        if ulp * qty > 0:
            dur_years = value / (ulp * qty)
        else:
            dur_years = _union_duration_years(group["start_date"], group[end_col])
        rec: dict = {
            "quantity": qty,
            "unit_list_price": ulp,
            "duration_years": dur_years,
            "duration_months": dur_years * 12.0,
            "calculated_extended_list_price": value,
            "source_excel_row": ", ".join(
                str(int(v)) for v in pd.to_numeric(group["source_excel_row"], errors="coerce").dropna().unique()[:20]
            ),
        }
        for col in first_cols:
            if col in group.columns:
                rec[col] = group[col].iloc[0]
        # Span the combined coverage window and preserve LDOS capping if any split was capped
        rec["start_date"] = pd.to_datetime(group["start_date"], errors="coerce").min()
        combined_end = pd.to_datetime(group[end_col], errors="coerce").max()
        rec["end_date"] = combined_end
        rec["effective_end_date"] = combined_end
        if "ldos_capped" in group.columns:
            rec["ldos_capped"] = bool(group["ldos_capped"].any())
        return rec

    records = []
    for keys, group in df.groupby(group_cols, dropna=False, sort=False):
        rec = _combine(group)
        key_tuple = keys if isinstance(keys, tuple) else (keys,)
        for col, val in zip(group_cols, key_tuple):
            rec[col] = val
        records.append(rec)

    result = pd.DataFrame(records)
    # Put the match-key columns first for readability
    ordered = group_cols + [c for c in result.columns if c not in group_cols]
    return result[ordered]



def _build_product_item_type_map(*sniff_dfs: pd.DataFrame | None) -> dict[str, str]:
    """product_number -> item_type (most common), built from SNIFF file(s)."""
    frames = [df for df in sniff_dfs if df is not None and not df.empty and "item_type" in df.columns and "product_number" in df.columns]
    if not frames:
        return {}
    both = pd.concat(frames, ignore_index=True)
    m = both[
        both["item_type"].astype(str).str.strip().ne("")
        & both["item_type"].astype(str).str.strip().str.lower().ne("nan")
        & both["product_number"].astype(str).str.strip().ne("")
    ]
    if m.empty:
        return {}
    grouped = m.groupby(m["product_number"].astype(str).str.strip())["item_type"]
    return {prod: (vals.mode().iloc[0] if not vals.mode().empty else "") for prod, vals in grouped}


def _fill_item_type_by_product(df: pd.DataFrame, prod_item_type: dict[str, str]) -> pd.DataFrame:
    """Fill blank/NaN item_type using the product_number -> item_type map."""
    if df.empty or not prod_item_type or "product_number" not in df.columns:
        return df
    out = df.copy()
    if "item_type" not in out.columns:
        out["item_type"] = ""
    s = out["item_type"]
    # Normalize nulls to "" BEFORE astype(str) — a float64 all-NaN column does not
    # convert to the string "nan" reliably, which would break the blank mask.
    text = s.where(s.notna(), "").astype(str).str.strip()
    blank = text.eq("") | text.str.lower().isin(["nan", "none", "<na>"])
    mapped = out["product_number"].astype(str).str.strip().map(prod_item_type)
    out["item_type"] = np.where(blank & mapped.notna(), mapped, text)
    return out


def _build_product_family_map(*sniff_dfs: pd.DataFrame | None) -> dict[str, str]:
    """product_number -> product_family (most common), built from raw SNIFF file(s).

    Uses product_number as the key so that SW subscription products that fail
    the MATCH_KEYS enrichment join (no matching serial/instance/SKU) can still
    get their product_family label from the SNIFF's Product Family column (AY).
    """
    frames = [df for df in sniff_dfs if df is not None and not df.empty
              and "product_family" in df.columns and "product_number" in df.columns]
    if not frames:
        return {}
    both = pd.concat(frames, ignore_index=True)
    m = both[
        both["product_family"].astype(str).str.strip().ne("")
        & both["product_family"].astype(str).str.strip().str.lower().ne("nan")
        & both["product_number"].astype(str).str.strip().ne("")
    ]
    if m.empty:
        return {}
    grouped = m.groupby(m["product_number"].astype(str).str.strip())["product_family"]
    return {prod: (vals.mode().iloc[0] if not vals.mode().empty else "") for prod, vals in grouped}


def _count_chassis(df: pd.DataFrame) -> pd.DataFrame:
    """Count items with a service price (unit_list_price > 0), grouped by product
    ID, regardless of Major/Minor. A priced Minor line (line card, module,
    software license) is a serviceable item and is counted. $0 items are excluded.
    Returns QTY (count of unique lines), qty_sum (sum of quantity field), item_type, ldos_date.
    """
    if df.empty:
        return pd.DataFrame(columns=["product_number", "QTY", "qty_sum", "item_type", "ldos_date"])
    ulp = pd.to_numeric(df.get("unit_list_price", pd.Series(0, index=df.index)), errors="coerce").fillna(0)
    df = df[ulp.gt(0)].copy()
    if df.empty:
        return pd.DataFrame(columns=["product_number", "QTY", "qty_sum", "item_type", "ldos_date"])
    if "item_type" not in df.columns:
        df["item_type"] = ""
    if "ldos_date" not in df.columns:
        df["ldos_date"] = pd.NaT
    return (
        df
        .groupby("product_number", dropna=False)
        .agg(
            QTY=("product_number", "size"),
            qty_sum=("quantity", lambda x: int(pd.to_numeric(x, errors="coerce").fillna(0).sum())),
            item_type=("item_type", lambda x: next((v for v in x if pd.notna(v) and str(v).strip() and str(v).strip().lower() != "nan"), "")),
            ldos_date=("ldos_date", lambda x: x.dropna().min() if x.notna().any() else pd.NaT),
        )
        .reset_index()
    )


def _removed_chassis_by_reason(last_calc: pd.DataFrame, serial_changes: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if serial_changes.empty:
        empty = pd.DataFrame(columns=["product_number", "QTY", "item_type", "ldos_date"])
        return empty, empty
    removed_sns = set(serial_changes.loc[serial_changes["change_type"].eq("Removed"), "serial_number"].astype(str))
    removed = last_calc[last_calc["serial_number"].astype(str).isin(removed_sns)].copy()
    if removed.empty:
        empty = pd.DataFrame(columns=["product_number", "QTY", "item_type", "ldos_date"])
        return empty, empty
    # LDOS removal = item whose contract ended exactly at its LDOS date (end_date == ldos_date)
    # This is consistent with _ldos_products and does not depend on today's date.
    ldos_mask = removed["ldos_date"].notna() & (removed["end_date"] == removed["ldos_date"])
    return _count_chassis(removed[ldos_mask]), _count_chassis(removed[~ldos_mask])


def _build_summary(
    last_ts_calc: pd.DataFrame,
    this_ts_calc: pd.DataFrame,
    last_sw_calc: pd.DataFrame,
    this_sw_calc: pd.DataFrame,
    merged_ts: pd.DataFrame,
    sw_result: pd.DataFrame,
    serial_changes: pd.DataFrame,
    instance_changes: pd.DataFrame,
    replacement_pairs: pd.DataFrame,
) -> pd.DataFrame:

    def srow(section: str, metric: str, svc: object = "", sw: object = "", dt: str = "", note: str = "") -> dict:
        return {"section": section, "metric": metric, "value svc ($)": svc,
                "value SW ($)": sw, "Total $": "", "date": dt, "formula / note": note}

    # ── TS effects ────────────────────────────────────────────
    if not merged_ts.empty:
        has_sn = merged_ts["serial_number"].fillna("").astype(str).ne("")
        ts_added = merged_ts["line_status"].eq("Added")
        ts_removed = merged_ts["line_status"].eq("Removed")
        ts_existing = merged_ts["line_status"].eq("Matched")

        ts_add_sn = float(merged_ts.loc[ts_added & has_sn, "calculated_extended_list_price_this"].sum())
        ts_add_inst = float(merged_ts.loc[ts_added & ~has_sn, "calculated_extended_list_price_this"].sum())
        ts_longer_dur = float(merged_ts.loc[ts_existing & merged_ts["duration_effect"].gt(0), "duration_effect"].sum())
        ts_higher_sla = float(merged_ts.loc[ts_existing & merged_ts["sla_change_effect"].gt(0), "sla_change_effect"].sum())
        ts_higher_price = float(merged_ts.loc[ts_existing & merged_ts["unit_price_effect"].gt(0), "unit_price_effect"].sum())
        ts_higher_qty = float(merged_ts.loc[ts_existing & merged_ts["quantity_effect"].gt(0), "quantity_effect"].sum())
        ts_rm_sn = float(merged_ts.loc[ts_removed & has_sn, "removed_value"].sum())
        ts_rm_inst = float(merged_ts.loc[ts_removed & ~has_sn, "removed_value"].sum())
        # Split removed TS values by LDOS (end_date == ldos_date) vs customer-removed
        # Use last_ts_calc as the join key since merged_ts may not carry ldos_date
        def _rm_ldos_split(rm_mask: pd.Series) -> tuple[float, float]:
            """Return (ldos_value, customer_value) for the removed rows matching rm_mask."""
            sub = merged_ts.loc[rm_mask].copy()
            if sub.empty:
                return 0.0, 0.0
            # Try to get ldos_date from the last year calc via serial/instance key
            sn_ldos: set = set()
            inst_ldos: set = set()
            for src in (last_ts_calc,):
                if src.empty or "ldos_date" not in src.columns or "end_date" not in src.columns:
                    continue
                ldos_rows = src[src["ldos_date"].notna() & (src["end_date"] == src["ldos_date"])]
                sn_ldos  |= set(ldos_rows["serial_number"].dropna().astype(str).str.strip())
                inst_ldos |= set(ldos_rows["instance_number"].dropna().astype(str).str.strip())
            # Mark each row as LDOS if its SN# or Instance# is in the ldos sets
            sn_col   = merged_ts["serial_number"].fillna("").astype(str).str.strip()
            inst_col = merged_ts["instance_number"].fillna("").astype(str).str.strip()
            is_ldos = (sn_col.isin(sn_ldos) | inst_col.isin(inst_ldos))
            ldos_val = float(merged_ts.loc[rm_mask & is_ldos,   "removed_value"].sum())
            cust_val = float(merged_ts.loc[rm_mask & ~is_ldos,  "removed_value"].sum())
            return ldos_val, cust_val
        ts_rm_sn_ldos,   ts_rm_sn_cust   = _rm_ldos_split(ts_removed & has_sn)
        ts_rm_inst_ldos, ts_rm_inst_cust = _rm_ldos_split(ts_removed & ~has_sn)
        ts_shorter_dur = float(merged_ts.loc[ts_existing & merged_ts["duration_effect"].lt(0), "duration_effect"].sum())
        ts_lower_sla = float(merged_ts.loc[ts_existing & merged_ts["sla_change_effect"].lt(0), "sla_change_effect"].sum())
        ts_lower_price = float(merged_ts.loc[ts_existing & merged_ts["unit_price_effect"].lt(0), "unit_price_effect"].sum())
        ts_lower_qty = float(merged_ts.loc[ts_existing & merged_ts["quantity_effect"].lt(0), "quantity_effect"].sum())
    else:
        ts_add_sn = ts_add_inst = ts_longer_dur = ts_higher_sla = ts_higher_price = ts_higher_qty = 0.0
        ts_rm_sn = ts_rm_inst = ts_shorter_dur = ts_lower_sla = ts_lower_price = ts_lower_qty = 0.0
        ts_rm_sn_ldos = ts_rm_sn_cust = ts_rm_inst_ldos = ts_rm_inst_cust = 0.0

    ts_add_value = ts_add_sn + ts_add_inst + ts_longer_dur + ts_higher_sla + ts_higher_price + ts_higher_qty
    ts_rm_value = ts_rm_sn + ts_rm_inst + ts_shorter_dur + ts_lower_sla + ts_lower_price + ts_lower_qty

    # ── SW effects ────────────────────────────────────────────
    if not sw_result.empty:
        sw_added_key_m = sw_result["sw_status"].eq("SW Added (key matched)")
        sw_added_new_m = sw_result["sw_status"].eq("SW Added (new product)")
        sw_added_rev_m = sw_result["sw_status"].eq("SW Added (same product - review)")
        sw_rm_key_m = sw_result["sw_status"].eq("SW Removed (key matched)")
        sw_rm_gone_m = sw_result["sw_status"].eq("SW Removed (product gone)")
        sw_rm_rev_m = sw_result["sw_status"].eq("SW Removed (same product - review)")
        sw_match_m = sw_result["sw_status"].isin(["SW Matched", "SW Tier Change"])

        sw_add_key_val = float(sw_result.loc[sw_added_key_m, "sw_added_value"].sum())
        sw_add_new_val = float(sw_result.loc[sw_added_new_m, "sw_added_value"].sum())
        sw_add_rev_val = float(sw_result.loc[sw_added_rev_m, "sw_added_value"].sum())
        sw_rm_key_val = float(sw_result.loc[sw_rm_key_m, "sw_removed_value"].sum())
        sw_rm_gone_val = float(sw_result.loc[sw_rm_gone_m, "sw_removed_value"].sum())
        sw_rm_rev_val = float(sw_result.loc[sw_rm_rev_m, "sw_removed_value"].sum())
        sw_longer_dur = float(sw_result.loc[sw_match_m & sw_result["sw_duration_effect"].gt(0), "sw_duration_effect"].sum())
        sw_higher_tier = float(sw_result.loc[sw_match_m & sw_result["sw_tier_effect"].gt(0), "sw_tier_effect"].sum())
        sw_higher_price = float(sw_result.loc[sw_match_m & sw_result["sw_unit_price_effect"].gt(0), "sw_unit_price_effect"].sum())
        sw_higher_qty = float(sw_result.loc[sw_match_m & sw_result["sw_qty_effect"].gt(0), "sw_qty_effect"].sum())
        sw_shorter_dur = float(sw_result.loc[sw_match_m & sw_result["sw_duration_effect"].lt(0), "sw_duration_effect"].sum())
        sw_lower_tier = float(sw_result.loc[sw_match_m & sw_result["sw_tier_effect"].lt(0), "sw_tier_effect"].sum())
        sw_lower_price = float(sw_result.loc[sw_match_m & sw_result["sw_unit_price_effect"].lt(0), "sw_unit_price_effect"].sum())
        sw_lower_qty = float(sw_result.loc[sw_match_m & sw_result["sw_qty_effect"].lt(0), "sw_qty_effect"].sum())

        sw_add_value = sw_add_key_val + sw_add_new_val + sw_add_rev_val + sw_longer_dur + sw_higher_tier + sw_higher_price + sw_higher_qty
        sw_rm_value = sw_rm_key_val + sw_rm_gone_val + sw_rm_rev_val + sw_shorter_dur + sw_lower_tier + sw_lower_price + sw_lower_qty

        # Section F counts are LINE counts (not unique-product counts) so they
        # stay consistent with Section G, which sums line_count per product.
        def _sw_lcnt(mask, col):
            if col in sw_result.columns:
                return int(pd.to_numeric(sw_result.loc[mask, col], errors="coerce").fillna(1).sum())
            return int(mask.sum())

        sw_add_key_cnt = _sw_lcnt(sw_added_key_m, "line_count_this")
        sw_add_new_cnt = _sw_lcnt(sw_added_new_m, "line_count_this")
        sw_add_rev_cnt = _sw_lcnt(sw_added_rev_m, "line_count_this")
        sw_rm_key_cnt = _sw_lcnt(sw_rm_key_m, "line_count_last")
        sw_rm_gone_cnt = _sw_lcnt(sw_rm_gone_m, "line_count_last")
        sw_rm_rev_cnt = _sw_lcnt(sw_rm_rev_m, "line_count_last")
        sw_tier_cnt = _sw_lcnt(sw_result["sw_status"].eq("SW Tier Change"), "line_count_this")
    else:
        sw_add_key_val = sw_add_new_val = sw_add_rev_val = 0.0
        sw_rm_key_val = sw_rm_gone_val = sw_rm_rev_val = 0.0
        sw_longer_dur = sw_higher_tier = sw_higher_price = sw_higher_qty = 0.0
        sw_shorter_dur = sw_lower_tier = sw_lower_price = sw_lower_qty = 0.0
        sw_add_value = sw_rm_value = 0.0
        sw_add_key_cnt = sw_add_new_cnt = sw_add_rev_cnt = 0
        sw_rm_key_cnt = sw_rm_gone_cnt = sw_rm_rev_cnt = sw_tier_cnt = 0

    # A/B totals use the de-duplicated aggregated line values (from merged_ts) so
    # that split/duplicate SN# lines are counted once, consistent with the effects.
    if not merged_ts.empty and "calculated_extended_list_price_last" in merged_ts.columns:
        ts_last_total = float(merged_ts["calculated_extended_list_price_last"].sum())
        ts_this_total = float(merged_ts["calculated_extended_list_price_this"].sum())
    else:
        ts_last_total = float(last_ts_calc["calculated_extended_list_price"].sum()) if not last_ts_calc.empty else 0.0
        ts_this_total = float(this_ts_calc["calculated_extended_list_price"].sum()) if not this_ts_calc.empty else 0.0
    sw_last_total = float(last_sw_calc["calculated_extended_list_price"].sum()) if not last_sw_calc.empty else 0.0
    sw_this_total = float(this_sw_calc["calculated_extended_list_price"].sum()) if not this_sw_calc.empty else 0.0

    ts_total_change = ts_add_value + ts_rm_value
    sw_total_change = sw_add_value + sw_rm_value

    # Build product_number → product_family map from sw_result (already enriched
    # in compare_renewals from the raw SNIFF data by product_number, bypassing
    # the serial/instance/SKU join that fails for SW subscriptions).
    sw_prod_family: dict[str, str] = {}
    if not sw_result.empty and "sw_product_family" in sw_result.columns:
        for _, _row in sw_result.iterrows():
            for _col in ("product_number_this", "product_number_last"):
                _pn = str(_row.get(_col) or "").strip()
                _pf = str(_row.get("sw_product_family") or "").strip()
                if _pn and _pf and _pf.lower() not in ("nan", "none", ""):
                    sw_prod_family[_pn] = _pf

    def _qty(col: str) -> int:
        """Sum quantity from serial_changes / instance_changes for a given change_type."""
        return 0  # placeholder; defined per call below

    def _sc_qty(df: pd.DataFrame, change_type: str) -> int:
        if df.empty or "quantity" not in df.columns:
            return 0
        return int(pd.to_numeric(df.loc[df["change_type"].eq(change_type), "quantity"], errors="coerce").fillna(0).sum())

    added_sn_count = int(serial_changes["change_type"].eq("Added").sum()) if not serial_changes.empty else 0
    added_sn_qty = _sc_qty(serial_changes, "Added")
    removed_sn_count = int(serial_changes["change_type"].eq("Removed").sum()) if not serial_changes.empty else 0
    removed_sn_qty = _sc_qty(serial_changes, "Removed")
    added_inst_count = int(instance_changes["change_type"].eq("Added").sum()) if not instance_changes.empty else 0
    added_inst_qty = _sc_qty(instance_changes, "Added")
    removed_inst_count = int(instance_changes["change_type"].eq("Removed").sum()) if not instance_changes.empty else 0
    removed_inst_qty = _sc_qty(instance_changes, "Removed")
    all_ts = pd.concat([last_ts_calc, this_ts_calc], ignore_index=True) if not (last_ts_calc.empty and this_ts_calc.empty) else pd.DataFrame()
    # ldos_count: items in either year where end_date == ldos_date (priced equipment at LDOS)
    def _count_ldos(df: pd.DataFrame) -> int:
        if df.empty or "ldos_date" not in df.columns or "end_date" not in df.columns:
            return 0
        has = df["ldos_date"].notna() & df["end_date"].notna() & (df["end_date"] == df["ldos_date"])
        priced = pd.to_numeric(df.get("unit_list_price", 0), errors="coerce").fillna(0).gt(0)
        return int((has & priced).sum())
    ldos_count = _count_ldos(last_ts_calc) + _count_ldos(this_ts_calc)

    rows = [
        srow("A",   "Last year calculated extended list price", svc=ts_last_total,      sw=sw_last_total,
             note="Sum of prorated list price from last year quote (matches the quote exactly)"),
        srow("B",   "This year calculated extended list price", svc=ts_this_total,      sw=sw_this_total,
             note="Sum of prorated list price from this year quote (matches the quote exactly)"),
        srow("C",   "Total increase / decrease",               svc=ts_total_change,    sw=sw_total_change,   note="C = D + E"),
        srow("",    "",                                         ""),
        srow("D",   "Add value",                               svc=ts_add_value,       sw=sw_add_value),
        srow("1",   "TS - Add SN#",                            svc=ts_add_sn,
             note="TS: Line Added and SN# has value"),
        srow("2",   "TS - Add Instance#",                      svc=ts_add_inst,
             note="TS: Line Added and SN# is blank"),
        srow("3",   "Sub-TnC - Add (key matched)",                             sw=sw_add_key_val,
             note="SW Added; resolved key (instance#/SN#/subscription ID) not in last year"),
        srow("4",   "Sub-TnC - Add (new product)",                             sw=sw_add_new_val,
             note="SW: no resolvable key and product ID not in last year SW subscriptions"),
        srow("5",   "Sub-TnC - potential expired & buy as New",                sw=sw_add_rev_val,
             note="SW: no key found; same product exists in last year. High probability this is a Sub-TnC that expired and was forced to re-order as new (no key link because system treats it as a fresh order)"),
        srow("6",   "Existing TS/Sub-TnC positive change effect",
             note="Shapley attribution for matched lines where one or more of quantity, duration, unit price increased"),
        srow("6.1", "Longer Duration value",                   svc=ts_longer_dur,      sw=sw_longer_dur,
             note="Existing lines (Matched) only \u2014 Shapley duration effect (positive)"),
        srow("6.2", "Higher SLA / Tier value",                 svc=ts_higher_sla,      sw=sw_higher_tier,
             note="Existing lines (Matched) only \u2014 TS: SLA level increased | SW: product ID (tier) upgraded; Shapley unit-price effect (positive)"),
        srow("6.3", "Higher Unit Price value",                 svc=ts_higher_price,    sw=sw_higher_price,
             note="Existing lines (Matched) only, same SLA/tier \u2014 Shapley unit-price effect (positive)"),
        srow("6.4", "Higher QTY value",                        svc=ts_higher_qty,      sw=sw_higher_qty,
             note="Existing lines (Matched) only \u2014 Shapley quantity effect (positive)"),
        srow("",    "",                                         ""),
        srow("E",   "Removal Value",                           svc=ts_rm_value,        sw=sw_rm_value),
        srow("1",   "TS - Remove SN#",                         svc=ts_rm_sn,
             note="TS: Line Removed and SN# has value"),
        srow("1.1", "TS - Remove SN# - LDOS",                  svc=ts_rm_sn_ldos,
             note="TS: SN# removed; contract ended at LDOS date (end_date == ldos_date)"),
        srow("1.2", "TS - Remove SN# - By Customer",            svc=ts_rm_sn_cust,
             note="TS: SN# removed; decommission, tech refresh, or LDOS date not reached"),
        srow("2",   "TS - Remove Instance#",                   svc=ts_rm_inst,
             note="TS: Line Removed and SN# is blank"),
        srow("2.1", "TS - Remove Instance# - LDOS",             svc=ts_rm_inst_ldos,
             note="TS: Instance# removed; contract ended at LDOS date (end_date == ldos_date)"),
        srow("2.2", "TS - Remove Instance# - By Customer",      svc=ts_rm_inst_cust,
             note="TS: Instance# removed; decommission, tech refresh, or LDOS date not reached"),
        srow("3",   "Sub-TnC - Remove (key matched)",                          sw=sw_rm_key_val,
             note="SW Removed; resolved key not present in this year"),
        srow("4",   "Sub-TnC - Remove (new product)",                          sw=sw_rm_gone_val,
             note="SW: no resolvable key and product ID not in this year SW subscriptions"),
        srow("5",   "Sub-TnC - Remove (data error)",                           sw=sw_rm_rev_val,
             note="SW: last-year line has no resolvable key AND same product appears unresolved in this year. A non-zero value indicates incomplete SNIFF \u2014 existing contracts should always have a subscription ID or instance number in the SNIFF file"),
        srow("6",   "Existing TS/Sub-TnC negative change effect",
             note="Shapley attribution for matched lines where one or more of quantity, duration, unit price decreased"),
        srow("6.1", "Shorter Duration value",                  svc=ts_shorter_dur,     sw=sw_shorter_dur,
             note="Existing lines (Matched) only \u2014 Shapley duration effect (negative)"),
        srow("6.2", "Lower SLA / Tier value",                  svc=ts_lower_sla,       sw=sw_lower_tier,
             note="Existing lines (Matched) only \u2014 TS: SLA level decreased | SW: product ID (tier) downgraded; Shapley unit-price effect (negative)"),
        srow("6.3", "Lower Unit Price value",                  svc=ts_lower_price,     sw=sw_lower_price,
             note="Existing lines (Matched) only, same SLA/tier \u2014 Shapley unit-price effect (negative)"),
        srow("6.4", "Lower QTY value",                         svc=ts_lower_qty,       sw=sw_lower_qty,
             note="Existing lines (Matched) only \u2014 Shapley quantity effect (negative)"),
        srow("",    "",                                         ""),
    ]

    # ── Section F: All-count summary (C=COUNT, D=QTY Sum) ───────────────────
    r_f = srow("F", "Quantity & Item Counts - All",
               note="COUNT of unique assets (de-duplicated); not sum of quantity column")
    r_f["value svc ($)"] = "COUNT"
    r_f["value SW ($)"] = "QTY Sum"
    rows.append(r_f)

    def _frow(section, metric, count, qty_sum, note=""):
        """Count row for section F: C=count, D=qty_sum (via sw), E=blank."""
        r = srow(section, metric, svc=count, sw=qty_sum, note=note)
        return r

    def _qrow(section, metric, count, qty_sum, note=""):
        """Count row helper (backward-compat alias)."""
        return _frow(section, metric, count, qty_sum, note)

    rows += [
        _frow("1", "TS - Replacement SN#",             len(replacement_pairs), len(replacement_pairs),
              note="Replacement treated as existing; value only from duration/qty/SLA/price change"),
        _frow("2", "TS - Add SN#",                     added_sn_count,    added_sn_qty,    "COUNT(TS SN# in this year not in last year)"),
        _frow("3", "TS - Add Instance#",               added_inst_count,  added_inst_qty,  "COUNT(TS instance# in this year not in last year)"),
        _frow("4", "TS - Remove SN#",                  removed_sn_count,  removed_sn_qty,  "COUNT(TS SN# in last year not in this year)"),
        _frow("5", "TS - Remove Instance#",            removed_inst_count,removed_inst_qty,"COUNT(TS instance# in last year not in this year)"),
    ]

    # Compute SW qty sums from sw_result
    def _sw_qty_sum(mask: pd.Series, q_col: str) -> int:
        if sw_result.empty or q_col not in sw_result.columns:
            return 0
        return int(pd.to_numeric(sw_result.loc[mask, q_col], errors="coerce").fillna(0).sum())

    sw_add_key_qty  = _sw_qty_sum(sw_added_key_m, "quantity_this") if not sw_result.empty else 0
    sw_add_new_qty  = _sw_qty_sum(sw_added_new_m, "quantity_this") if not sw_result.empty else 0
    sw_add_rev_cnt  = sw_add_rev_cnt if not sw_result.empty else 0
    sw_rm_qty       = _sw_qty_sum(sw_rm_key_m | sw_rm_gone_m, "quantity_last") if not sw_result.empty else 0
    sw_tier_qty     = _sw_qty_sum(sw_result["sw_status"].eq("SW Tier Change"), "quantity_this") if not sw_result.empty else 0

    sw_clearly_added = sw_add_key_cnt + sw_add_new_cnt
    sw_clearly_added_qty = sw_add_key_qty + sw_add_new_qty
    r_f6 = _frow("6", "Sub-TnC - Add", sw_clearly_added, sw_clearly_added_qty,
                 note="COUNT of clearly added SW subscription lines (key-matched new + new product)")
    rows.append(r_f6)

    r_f7 = _frow("7", "Sub-TnC - potential expired & buy as New", sw_add_rev_cnt, sw_add_rev_cnt,
                 note="Same product in both years but no key link; high probability expired sub-TnC re-ordered as new")
    rows.append(r_f7)

    sw_clearly_removed = sw_rm_key_cnt + sw_rm_gone_cnt
    r_f8 = _frow("8", "Sub-TnC - Remove (key matched)", sw_clearly_removed, sw_rm_qty,
                 note="Clearly removed SW subscriptions (key-matched gone + product ID absent this year)")
    rows.append(r_f8)

    r_f9 = _frow("9", "Sub-TnC - Tier Changes", sw_tier_cnt, sw_tier_qty,
                 note="Same key matched across years but product ID (tier) changed")
    rows.append(r_f9)
    rows.append(srow("", "", ""))

    # ── Section G: Items-with-service-price (C=QTY, D=QTY Sum or item_type, E=item_type or QTY Sum, F=date) ──
    r_g = srow("G", "Quantity & Item Counts - Only price items", svc="QTY", dt="Date",
               note="Rows below are quantities / item counts, not dollar values")
    r_g["value SW ($)"] = "QTY Sum"
    r_g["Total $"] = "Item Types"
    rows.append(r_g)

    def _item_type_str(r: object) -> str:
        v = getattr(r, "item_type", None)
        return str(v).strip() if v and str(v).strip() not in ("", "nan", "None") else ""

    def _item_ldos_str(r: object) -> str:
        ldos = getattr(r, "ldos_date", None)
        if ldos is not None and pd.notna(ldos):
            try:
                return pd.Timestamp(ldos).strftime("%d-%b-%Y")
            except Exception:
                pass
        return ""

    def _qty_str(r: object) -> int:
        v = getattr(r, "qty_sum", None)
        if v is not None:
            try:
                return int(v)
            except Exception:
                pass
        return int(r.QTY)

    # G.1: TS - Add (C=QTY, D=qty_sum, E=item_type)
    added_chassis = _count_chassis(
        this_ts_calc[this_ts_calc["serial_number"].astype(str).isin(
            set(serial_changes.loc[serial_changes["change_type"].eq("Added"), "serial_number"].astype(str))
        )] if not serial_changes.empty and not this_ts_calc.empty else pd.DataFrame()
    )

    # _ldos_products: items with end_date == ldos_date and unit list price > 0
    def _ldos_products(df: pd.DataFrame) -> pd.DataFrame:
        cols = ["product_number", "QTY", "qty_sum", "item_type", "ldos_date"]
        if df.empty:
            return pd.DataFrame(columns=cols)
        has_ldos = df["ldos_date"].notna()
        end_eq = df["end_date"].notna() & (df["end_date"] == df["ldos_date"])
        ulp_pos = pd.to_numeric(df.get("unit_list_price", 0), errors="coerce").fillna(0).gt(0)
        sub = df[has_ldos & end_eq & ulp_pos].copy()
        if "item_type" not in sub.columns:
            sub["item_type"] = ""
        if sub.empty:
            return pd.DataFrame(columns=cols)
        return (
            sub.groupby("product_number", dropna=False)
            .agg(
                QTY=("product_number", "size"),
                qty_sum=("quantity", lambda x: int(pd.to_numeric(x, errors="coerce").fillna(0).sum())),
                item_type=("item_type", lambda x: next((v for v in x if pd.notna(v) and str(v).strip()), "")),
                ldos_date=("ldos_date", "first"),
            )
            .reset_index()
        )

    # Compute LDOS products early so G.2.1 and G.3.1 share the same source
    last_ldos_products = _ldos_products(last_ts_calc)
    this_ldos_products = _ldos_products(this_ts_calc)

    removed_chassis_ldos, removed_chassis_customer = _removed_chassis_by_reason(last_ts_calc, serial_changes)
    removed_chassis_total = int(removed_chassis_ldos["QTY"].sum() + removed_chassis_customer["QTY"].sum())
    removed_chassis_total_qty = int(removed_chassis_ldos.get("qty_sum", removed_chassis_ldos["QTY"]).sum() +
                                    removed_chassis_customer.get("qty_sum", removed_chassis_customer["QTY"]).sum())

    r_g1 = srow("1", "TS - Add",
                svc=int(added_chassis["QTY"].sum()) if not added_chassis.empty else 0,
                sw=int(added_chassis["qty_sum"].sum()) if not added_chassis.empty and "qty_sum" in added_chassis.columns else 0,
                note="Added TS SN# lines with unit list price > 0, grouped by Product ID")
    rows.append(r_g1)
    for idx, r in enumerate(added_chassis.itertuples(index=False), start=1):
        ri = srow(f"1.{idx}", r.product_number or "Product ID", svc=int(r.QTY), sw=_qty_str(r))
        ri["Total $"] = _item_type_str(r)
        rows.append(ri)

    # G.2: TS - Remove
    # G.2: TS - Remove (G.2.1=LDOS items from last year, G.2.2=customer-removed)
    g21_qty = int(last_ldos_products["QTY"].sum()) if not last_ldos_products.empty else 0
    g21_qty_sum = int(last_ldos_products["qty_sum"].sum()) if not last_ldos_products.empty and "qty_sum" in last_ldos_products.columns else 0
    g22_qty = int(removed_chassis_customer["QTY"].sum()) if not removed_chassis_customer.empty else 0
    g22_qty_sum = int(removed_chassis_customer["qty_sum"].sum()) if not removed_chassis_customer.empty and "qty_sum" in removed_chassis_customer.columns else 0
    r_g2 = srow("2", "TS - Remove", svc=g21_qty + g22_qty, sw=g21_qty_sum + g22_qty_sum,
                note="Removed TS SN# items grouped by removal reason")
    rows.append(r_g2)

    # G.2.1: LDOS-expired items from last year (end_date == ldos_date, price > 0)
    # Same data as G.3.1 — G.3 is the full informational LDOS section
    r_g21 = srow("2.1", "TS - Remove due to LDOS",
                 svc=g21_qty, sw=g21_qty_sum,
                 note="Last year items whose contract ended at LDOS date (end_date == ldos_date)")
    rows.append(r_g21)
    for idx, r in enumerate(last_ldos_products.itertuples(index=False), start=1):
        row = srow(f"2.1.{idx}", r.product_number or "Product ID",
                   svc=int(r.QTY), sw=_item_type_str(r), dt=_item_ldos_str(r))
        row["Total $"] = _qty_str(r)
        rows.append(row)

    # G.2.2: Removed by customer
    r_g22 = srow("2.2", "TS - Remove by customer",
                 svc=g22_qty, sw=g22_qty_sum,
                 note="SN# removed but LDOS not reached or LDOS is blank — decommission/tech refresh")
    rows.append(r_g22)
    for idx, r in enumerate(removed_chassis_customer.itertuples(index=False), start=1):
        ri = srow(f"2.2.{idx}", r.product_number or "Product ID", svc=int(r.QTY), sw=_qty_str(r))
        ri["Total $"] = _item_type_str(r)
        rows.append(ri)

    # G.3: LDOS Items (this year only — last year LDOS removals are now shown in G.2.1)
    # last_ldos_products and this_ldos_products already computed before G.2
    this_ldos_qty = int(this_ldos_products["QTY"].sum()) if not this_ldos_products.empty else 0
    this_ldos_qty_sum = int(this_ldos_products["qty_sum"].sum()) if not this_ldos_products.empty and "qty_sum" in this_ldos_products.columns else 0

    r_g3 = srow("3", "This Year Renew to LDOS Date",
                svc=this_ldos_qty, sw=this_ldos_qty_sum,
                note="This year quote items with unit price > 0 whose end date = LDOS date "
                     "(last year LDOS removals are shown in G.2.1)")
    rows.append(r_g3)
    for idx, r in enumerate(this_ldos_products.itertuples(index=False), start=1):
        ri = srow(f"3.{idx}", r.product_number or "Product ID",
                  svc=int(r.QTY), sw=_qty_str(r), dt=_item_ldos_str(r))
        ri["Total $"] = _item_type_str(r)
        rows.append(ri)

    # G.4-6: SW subscription items (C=count, D=qty_sum, E=product_family)
    # Section G lists ONLY items with unit list price > 0 (F holds the full
    # counts including zero-price lines).
    def _sw_g_products(statuses, prod_col, cnt_col, qty_col="", price_col=""):
        """Return ordered {product: {count, qty_sum}} for priced SW lines."""
        prod_data: dict[str, dict] = {}
        if sw_result.empty:
            return prod_data
        subset = sw_result[sw_result["sw_status"].isin(statuses)]
        if price_col and price_col in subset.columns:
            price_vals = pd.to_numeric(subset[price_col], errors="coerce").fillna(0)
            subset = subset[price_vals > 0]
        if subset.empty:
            return prod_data
        prod_series = subset[prod_col].fillna("").astype(str) if prod_col in subset.columns else subset["sw_match_key"].fillna("").astype(str)
        cnt_series = subset[cnt_col].fillna(1).astype(int) if cnt_col in subset.columns else pd.Series(1, index=subset.index)
        qty_series = pd.to_numeric(subset[qty_col], errors="coerce").fillna(0) if qty_col and qty_col in subset.columns else cnt_series.astype(float)
        for prod, cnt, qty_val in zip(prod_series, cnt_series, qty_series):
            key = prod.strip() or "(unknown product)"
            entry = prod_data.setdefault(key, {"count": 0, "qty_sum": 0.0})
            entry["count"] += int(cnt)
            entry["qty_sum"] += float(qty_val)
        return prod_data

    def _emit_sw_g_subrows(prod_data, parent_section):
        for sub_idx, (prod, data) in enumerate(prod_data.items(), start=1):
            pf = sw_prod_family.get(prod.strip(), "")
            qsum = data["qty_sum"]
            r = srow(f"{parent_section}.{sub_idx}", f"  {prod}", svc=data["count"],
                     sw=int(qsum) if qsum == int(qsum) else qsum)
            r["Total $"] = pf
            rows.append(r)

    def _totals(prod_data):
        cnt = sum(d["count"] for d in prod_data.values())
        qty = sum(d["qty_sum"] for d in prod_data.values())
        return cnt, (int(qty) if qty == int(qty) else qty)

    # G.4 Sub-TnC Add = 4.1 (key matched) + 4.2 (potential expired), priced items only
    g41 = _sw_g_products(["SW Added (key matched)"], "product_number_this", "line_count_this", "quantity_this", "unit_list_price_this")
    g42 = _sw_g_products(["SW Added (same product - review)"], "product_number_this", "line_count_this", "quantity_this", "unit_list_price_this")
    g41_cnt, g41_qty = _totals(g41)
    g42_cnt, g42_qty = _totals(g42)
    g4_cnt = g41_cnt + g42_cnt
    g4_qty_raw = (g41_qty if isinstance(g41_qty, int) else g41_qty) + (g42_qty if isinstance(g42_qty, int) else g42_qty)
    g4_qty = int(g4_qty_raw) if g4_qty_raw == int(g4_qty_raw) else g4_qty_raw

    r_g4 = srow("4", "Sub-TnC - Add", svc=g4_cnt, sw=g4_qty,
                note="Additional SW subscriptions with unit list price > 0 (clearly added)")
    rows.append(r_g4)
    r_g41 = srow("4.1", "Sub-TnC - Add", svc=g41_cnt, sw=g41_qty,
                 note="Resolved key not in last year (unit list price > 0)")
    rows.append(r_g41)
    _emit_sw_g_subrows(g41, "4.1")
    r_g42 = srow("4.2", "Sub-TnC - potential expired & buy as New", svc=g42_cnt, sw=g42_qty,
                 note="No key; likely expired and re-ordered as new (unit list price > 0)")
    rows.append(r_g42)
    _emit_sw_g_subrows(g42, "4.2")

    # G.5 Sub-TnC Remove, priced items only
    g5 = _sw_g_products(["SW Removed (key matched)", "SW Removed (product gone)"], "product_number_last", "line_count_last", "quantity_last", "unit_list_price_last")
    g5_cnt, g5_qty = _totals(g5)
    r_g5 = srow("5", "Sub-TnC - Remove", svc=g5_cnt, sw=g5_qty,
                note="Clearly removed SW subscriptions with unit list price > 0")
    rows.append(r_g5)
    _emit_sw_g_subrows(g5, "5")

    r_g6 = srow("6", "Sub-TnC - Tier Changes", svc=sw_tier_cnt, sw=sw_tier_qty,
                note="Same key matched across years but product ID (tier) changed")
    rows.append(r_g6)

    # ── Total $ post-processing ─────────────────────────────────────────────────
    # Total $ column = value svc ($) + value SW ($), only for the dollar-bridge
    # rows (sections A–E and their 1.x / 2.x sub-items). Count rows stay blank.
    # Preserve any text already written to Total $ (e.g. LDOS dates on 9.1.x rows).
    dollar_secs = {"A", "B", "C", "D", "E"}
    for r in rows:
        sec = str(r.get("section", ""))
        is_dollar = sec in dollar_secs or sec.startswith("1.") or sec.startswith("2.")
        existing = r.get("Total $", "")
        # Don't overwrite any non-empty value already placed by the builder
        # (covers text headers like "QTY Sum", integer qty sums, date strings, etc.)
        if existing != "":
            continue
        total: object = ""
        if is_dollar:
            acc = 0.0
            has = False
            for k in ("value svc ($)", "value SW ($)"):
                x = r.get(k, "")
                if isinstance(x, (int, float)) and not isinstance(x, bool):
                    acc += float(x)
                    has = True
            total = acc if has else ""
        r["Total $"] = total

    rows.append(srow("", "", ""))
    rows.append(srow("NOTE", "", note=(
        "Note: For matched lines where more than one factor changed (quantity, duration, unit price), "
        "each effect is the average marginal contribution across all possible orderings of the three factors. "
        "This is order-independent and splits any interaction term fairly. "
        "Effects always sum exactly to the total line change (bridge identity). "
        "Formula: E_factor = \u0394factor \u00d7 [1/3\u00b7(last other two) + 1/6\u00b7(mixed) + 1/3\u00b7(this other two)]."
    )))

    return pd.DataFrame(rows, columns=SUMMARY_COLUMNS)


def _build_line_detail_summary(merged: pd.DataFrame) -> pd.DataFrame:
    if merged.empty:
        return pd.DataFrame()
    pivot = merged.pivot_table(
        index=["line_status", "service_level_last", "service_level_this"],
        values=[
            "calculated_extended_list_price_last",
            "calculated_extended_list_price_this",
            "added_value",
            "removed_value",
            "quantity_effect",
            "duration_effect",
            "unit_price_effect",
            "sla_change_effect",
            "total_change",
        ],
        aggfunc="sum",
        fill_value=0,
        dropna=False,
    ).reset_index()
    counts = merged.groupby(["line_status", "service_level_last", "service_level_this"], dropna=False).size().reset_index(name="line_count")
    pivot = counts.merge(pivot, on=["line_status", "service_level_last", "service_level_this"], how="left")
    return pivot.sort_values(["line_status", "service_level_last", "service_level_this"]).reset_index(drop=True)


def compare_renewals(
    last_df: pd.DataFrame,
    this_df: pd.DataFrame,
    last_sniff_df: pd.DataFrame | None = None,
    this_sniff_df: pd.DataFrame | None = None,
) -> dict[str, pd.DataFrame]:
    replacement_pairs = _build_replacement_pairs(last_sniff_df)
    last_calc = add_calculated_fields(enrich_quote_with_sniff(last_df, last_sniff_df))
    this_calc = add_calculated_fields(_apply_replacement_matching(enrich_quote_with_sniff(this_df, this_sniff_df), replacement_pairs))

    # Fill item_type by product ID from SNIFF (a product's type is the same
    # regardless of serial). This gives removed/retired items an item_type even
    # when their specific serial is not present in the current SNIFF.
    prod_item_type = _build_product_item_type_map(last_sniff_df, this_sniff_df)
    last_calc = _fill_item_type_by_product(last_calc, prod_item_type)
    this_calc = _fill_item_type_by_product(this_calc, prod_item_type)

    # ── Separate TS service and SW Sub-TnC lines ──────────────
    last_sw_mask = _is_sw(last_calc)
    this_sw_mask = _is_sw(this_calc)
    last_ts_calc = last_calc[~last_sw_mask].copy()
    this_ts_calc = this_calc[~this_sw_mask].copy()
    last_sw_calc = last_calc[last_sw_mask].copy()
    this_sw_calc = this_calc[this_sw_mask].copy()

    # ── TS comparison (serial# + instance# match; SKU is a data field) ────
    last_ts_lines = _aggregate_lines(last_ts_calc)
    this_ts_lines = _aggregate_lines(this_ts_calc)

    merged = last_ts_lines.merge(this_ts_lines, on=TS_LINE_MATCH_KEYS, how="outer", suffixes=("_last", "_this"), indicator=True)
    merged["line_status"] = np.select(
        [merged["_merge"].eq("left_only"), merged["_merge"].eq("right_only")],
        ["Removed", "Added"],
        default="Matched",
    )
    merged["replacement_adjusted"] = merged.get("replacement_note_this", pd.Series(index=merged.index, dtype=object)).fillna("").astype(str).ne("")

    for col in ["quantity", "unit_list_price", "duration_years", "duration_months", "calculated_extended_list_price"]:
        merged[f"{col}_last"] = merged[f"{col}_last"].fillna(0)
        merged[f"{col}_this"] = merged[f"{col}_this"].fillna(0)

    merged["unit_price_change"] = merged["unit_list_price_this"] - merged["unit_list_price_last"]
    merged["unit_price_change_pct"] = np.where(
        (merged["line_status"].eq("Added")) | (merged["unit_list_price_last"].eq(0) & merged["unit_list_price_this"].gt(0)),
        1.0,
        np.where(merged["unit_list_price_last"].ne(0), merged["unit_price_change"] / merged["unit_list_price_last"], np.nan),
    )
    merged["note"] = np.select(
        [merged["replacement_adjusted"], merged["line_status"].eq("Added"), merged["line_status"].eq("Removed")],
        ["Replacement SN treated as existing; value impact only from duration, quantity, or SLA/unit price changes", "Additional SKU", "Removed SKU"],
        default="",
    )
    merged["service_level_changed"] = (merged["line_status"].eq("Matched")) & (
        (merged.get("service_level_last", "").fillna("").astype(str) != merged.get("service_level_this", "").fillna("").astype(str))
        | (merged.get("service_sku_last", "").fillna("").astype(str) != merged.get("service_sku_this", "").fillna("").astype(str))
    )

    last_value = merged["calculated_extended_list_price_last"]
    this_value = merged["calculated_extended_list_price_this"]
    ts_matched = merged["line_status"].eq("Matched")

    merged["added_value"] = np.where(merged["line_status"].eq("Added"), this_value, 0.0)
    merged["replacement_value"] = 0.0
    merged["removed_value"] = np.where(merged["line_status"].eq("Removed"), -last_value, 0.0)

    # Shapley (symmetric) decomposition of the matched-line value change into
    # quantity, duration and unit-price effects. value = qty × duration × ulp.
    # Each effect is the average marginal contribution over all factor orderings,
    # so interaction terms are split fairly and the three effects sum exactly to
    # the line's total change regardless of ordering.
    q0 = merged["quantity_last"]
    q1 = merged["quantity_this"]
    d0 = merged["duration_years_last"]
    d1 = merged["duration_years_this"]
    u0 = merged["unit_list_price_last"]
    u1 = merged["unit_list_price_this"]
    shap_q = (q1 - q0) * (d0 * u0 / 3.0 + (d0 * u1 + d1 * u0) / 6.0 + d1 * u1 / 3.0)
    shap_d = (d1 - d0) * (q0 * u0 / 3.0 + (q0 * u1 + q1 * u0) / 6.0 + q1 * u1 / 3.0)
    shap_u = (u1 - u0) * (q0 * d0 / 3.0 + (q0 * d1 + q1 * d0) / 6.0 + q1 * d1 / 3.0)

    merged["quantity_effect"] = np.where(ts_matched, shap_q, 0.0)
    merged["duration_effect"] = np.where(ts_matched, shap_d, 0.0)
    merged["unit_price_effect"] = np.where(ts_matched & ~merged["service_level_changed"], shap_u, 0.0)
    merged["sla_change_effect"] = np.where(ts_matched & merged["service_level_changed"], shap_u, 0.0)
    merged["total_change"] = this_value - last_value
    merged["bridge_total"] = merged[["added_value", "replacement_value", "removed_value", "quantity_effect", "duration_effect", "unit_price_effect", "sla_change_effect"]].sum(axis=1)

    # ── SW subscription comparison ────────────────────────────
    sw_result = _classify_sw_subs(last_sw_calc, this_sw_calc)

    # Enrich sw_result with product_family from raw SNIFF by product_number.
    # The standard SNIFF enrichment joins on serial/instance/SKU and misses SW
    # subs; the product_family map joins on product_number instead.
    sniff_family_map = _build_product_family_map(last_sniff_df, this_sniff_df)
    if not sw_result.empty and sniff_family_map:
        def _get_sw_family(row: pd.Series) -> str:
            for col in ("product_number_this", "product_number_last"):
                pn = str(row.get(col) or "").strip()
                if pn in sniff_family_map:
                    return sniff_family_map[pn]
            return sniff_family_map.get(str(row.get("sw_match_key") or "").strip(), "")
        sw_result = sw_result.copy()
        sw_result["sw_product_family"] = sw_result.apply(_get_sw_family, axis=1)

    # ── Entity change sheets (TS only) ────────────────────────
    serial_changes = _entity_changes(last_ts_calc, this_ts_calc, "serial_number", "serial_number")
    instance_changes = _entity_changes(last_ts_calc, this_ts_calc, "instance_number", "instance_number")

    sku_changes = merged[TS_LINE_MATCH_KEYS + [
        "service_sku_last",
        "service_sku_this",
        "line_status",
        "service_level_last",
        "service_level_this",
        "service_level_changed",
        "unit_list_price_last",
        "unit_list_price_this",
        "unit_price_change",
        "unit_price_change_pct",
        "note",
    ]].copy()

    sla_changes = merged[merged["service_level_changed"]][TS_LINE_MATCH_KEYS + [
        "service_sku_last",
        "service_sku_this",
        "service_level_last",
        "service_level_this",
        "unit_list_price_last",
        "unit_list_price_this",
        "sla_change_effect",
        "total_change",
    ]].copy()

    ldos_exceptions = pd.concat([last_ts_calc.assign(period="Last Year"), this_ts_calc.assign(period="This Year")], ignore_index=True)
    if "ldos_capped" in ldos_exceptions.columns:
        ldos_exceptions = ldos_exceptions[ldos_exceptions["ldos_capped"]].copy()
    else:
        ldos_exceptions = pd.DataFrame()

    # ── Data quality checks ───────────────────────────────────
    # A genuine data issue is only raised for PRICED items (unit_list_price > 0),
    # counted on DE-DUPLICATED assets (split-order / duplicate lines counted
    # once) and reported both by asset count AND by quantity sum.
    # Missing serial number is NOT an issue (software has an instance but no
    # serial; cables have neither). Blank serial/SKU on $0 items is normal and is
    # shown only in the informational 'Blank Field Summary' sheet.
    # Nothing sourced from SNIFF is ever treated as a data error.
    quality_rows = []
    for label, df in [("Last Year", last_calc), ("This Year", this_calc)]:
        qdf = df.drop_duplicates(subset=MATCH_KEYS).copy()
        qty = pd.to_numeric(qdf["quantity"], errors="coerce").fillna(0)
        priced_mask = pd.to_numeric(qdf["unit_list_price"], errors="coerce").fillna(0) > 0
        priced = qdf[priced_mask]

        def _issue(mask: pd.Series, issue: str) -> None:
            if mask.any():
                quality_rows.append({
                    "period": label,
                    "issue": issue,
                    "asset_count": int(mask.sum()),
                    "qty_sum": float(qty[mask].sum()),
                })

        # Priced item missing service SKU (genuine issue)
        sku_missing = priced_mask & qdf["service_sku"].astype(str).eq("")
        _issue(sku_missing, "Priced item missing service SKU")
        # Missing dates
        _issue(qdf["start_date"].isna(), "Missing start date")
        _issue(qdf["end_date"].isna(), "Missing end date")
        # Quantity / price sanity
        _issue(qty <= 0, "Quantity is zero or negative")
        _issue(pd.to_numeric(qdf["unit_list_price"], errors="coerce").fillna(0) < 0, "Unit list price is negative")
        # Priced TS line missing instance
        if "service_type" in qdf.columns:
            ts_missing_inst = qdf["service_type"].astype(str).eq("TS") & qdf["instance_number"].astype(str).eq("") & priced_mask
            _issue(ts_missing_inst, "Priced TS service line missing instance number")
        # Sub-TnC with no resolvable key at all
        sw_mask = _is_sw(qdf)
        if sw_mask.any():
            sw_r = _resolve_sw_match_key(qdf[sw_mask])
            no_key_no_prod = sw_r["sw_match_key"].eq("") & sw_r["product_number"].fillna("").astype(str).eq("")
            if no_key_no_prod.any():
                q_sw = pd.to_numeric(qdf.loc[sw_mask, "quantity"], errors="coerce").fillna(0)
                quality_rows.append({
                    "period": label,
                    "issue": "Sub-TnC line with no resolvable key and no product number",
                    "asset_count": int(no_key_no_prod.sum()),
                    "qty_sum": float(q_sw[no_key_no_prod].sum()),
                })

    # Informational only (NOT data issues): de-duplicated blank serial / SKU,
    # reported both by asset count and by quantity sum.
    info_rows = []
    for label, df in [("Last Year", last_calc), ("This Year", this_calc)]:
        qdf = df.drop_duplicates(subset=MATCH_KEYS)
        qty = pd.to_numeric(qdf["quantity"], errors="coerce").fillna(0)
        sn_blank = qdf["serial_number"].astype(str).eq("")
        sku_blank = qdf["service_sku"].astype(str).eq("")
        info_rows.append({
            "period": label,
            "missing serial (count)": int(sn_blank.sum()),
            "missing serial (qty sum)": float(qty[sn_blank].sum()),
            "missing service SKU (count)": int(sku_blank.sum()),
            "missing service SKU (qty sum)": float(qty[sku_blank].sum()),
            "total unique assets": int(len(qdf)),
            "note": "Blank serial/SKU is normal (SW has no serial; cables have neither; $0 minor items have no SKU) — not a data issue",
        })

    sniff_status_rows = []
    for label, sniff_df, quote_df in [
        ("Last Year SNIFF", last_sniff_df, last_calc),
        ("This Year SNIFF", this_sniff_df, this_calc),
    ]:
        if sniff_df is None or sniff_df.empty:
            continue
        status = sniff_df.get("asset_status", pd.Series(dtype=object)).astype(str).str.strip().str.upper()
        # Capture both EXPIRED and NEVER COVERED (out-of-coverage assets)
        flagged = sniff_df[status.isin(["EXPIRED", "NEVER COVERED"])].copy()
        if flagged.empty:
            continue
        flagged_status = status[flagged.index]
        # Determine whether each flagged SNIFF asset is actually present in the
        # matching quote (actionable) or is SNIFF-only extra config the customer
        # removed (informational noise, never counted in the value analysis).
        quote_keys: set = set()
        if quote_df is not None and not quote_df.empty:
            quote_keys = set(map(tuple, quote_df[MATCH_KEYS].astype(str).values))
        sniff_keys = flagged[MATCH_KEYS].astype(str)
        in_quote = [tuple(row) in quote_keys for row in sniff_keys.values]
        flagged = flagged.assign(
            period=label,
            asset_status_norm=flagged_status.values,
            in_current_quote=in_quote,
            coverage_note=[
                (f"{st.title()}; asset IS in the renewal quote — review coverage"
                 if inq else
                 f"{st.title()}; SNIFF-only extra config not in the quote — excluded from value analysis")
                for st, inq in zip(flagged_status.values, in_quote)
            ],
        )
        sniff_status_rows.append(flagged)
    sniff_status = pd.concat(sniff_status_rows, ignore_index=True) if sniff_status_rows else pd.DataFrame()

    # ── Overlapping-duration check (split orders must be sequential) ──────
    overlap_errors = pd.concat(
        [
            _detect_overlaps(last_ts_calc, "Last Year"),
            _detect_overlaps(this_ts_calc, "This Year"),
            _detect_overlaps(last_sw_calc, "Last Year"),
            _detect_overlaps(this_sw_calc, "This Year"),
        ],
        ignore_index=True,
    )

    # ── Summary ───────────────────────────────────────────────
    summary = _build_summary(
        last_ts_calc, this_ts_calc, last_sw_calc, this_sw_calc,
        merged, sw_result, serial_changes, instance_changes, replacement_pairs,
    )

    # ── Calculation explanation ───────────────────────────────
    explanation = pd.DataFrame([
        {"row_type": "Last / This year calculated extended list price",
         "meaning": "Extended list value taken from the quote's prorated list price so it matches the quote exactly. TS service in 'value svc ($)'; SW Sub-TnC in 'value SW ($)'.",
         "formula": "sum(quote prorated_list_price), scaled by capped/full days when LDOS caps coverage"},
        {"row_type": "Additional SN# / Instance Value (TS)",
         "meaning": "Value of new TS lines not matched to last year.",
         "formula": "sum(this_value for Added TS lines)"},
        {"row_type": "Additional Sub-TnC (key matched)",
         "meaning": "SW sub line this year whose resolved key (instance#/SN#/subscription ID) was absent from last year.",
         "formula": "sum(this_value)"},
        {"row_type": "Additional Sub-TnC (new product)",
         "meaning": "SW sub with no resolvable key; product ID not seen in last year SW subs.",
         "formula": "sum(this_value)"},
        {"row_type": "Sub-TnC same product - review",
         "meaning": "SW sub with no resolvable key; same product ID exists in last year but lines cannot be linked. May be new or continued.",
         "formula": "sum(this_value)"},
        {"row_type": "Longer / Shorter Duration value",
         "meaning": "Shapley duration effect on matched lines (TS and SW). Order-independent; interaction with qty/price split fairly.",
         "formula": "Δduration × [ (q0·u0)/3 + (q0·u1 + q1·u0)/6 + (q1·u1)/3 ]"},
        {"row_type": "Higher / Lower SLA value (TS)",
         "meaning": "TS matched lines where service level changed — Shapley unit-price effect.",
         "formula": "Δunit_price × [ (q0·d0)/3 + (q0·d1 + q1·d0)/6 + (q1·d1)/3 ]"},
        {"row_type": "Higher / Lower Tier value (SW)",
         "meaning": "SW matched lines where product ID (tier) changed on same resolved key — Shapley unit-price effect.",
         "formula": "Δunit_price × [ (q0·d0)/3 + (q0·d1 + q1·d0)/6 + (q1·d1)/3 ]"},
        {"row_type": "Higher / Lower Unit Price value",
         "meaning": "Matched lines where price changed without SLA/tier change — Shapley unit-price effect.",
         "formula": "Δunit_price × [ (q0·d0)/3 + (q0·d1 + q1·d0)/6 + (q1·d1)/3 ]"},
        {"row_type": "Higher / Lower QTY value",
         "meaning": "Shapley quantity effect on matched lines.",
         "formula": "Δqty × [ (d0·u0)/3 + (d0·u1 + d1·u0)/6 + (d1·u1)/3 ]"},
        {"row_type": "Bridge total",
         "meaning": "Sum of all TS effects at line level. Equals total_change exactly (Shapley effects sum to the total change).",
         "formula": "added + replacement + removed + qty + duration + unit_price + sla"},
        {"row_type": "Shapley decomposition",
         "meaning": "Duration/quantity/unit-price effects use a symmetric (Shapley) split so the attribution is order-independent and interaction terms are shared fairly. Subscripts 0 = last year, 1 = this year.",
         "formula": "each effect = average marginal contribution over all factor orderings"},
    ])

    result = {
        "Summary": summary,
        "SW Subscription Analysis": sw_result if not sw_result.empty else pd.DataFrame(),
        "Line Detail Summary": _build_line_detail_summary(merged),
        "Serial Number Changes": serial_changes,
        "Instance Number Changes": instance_changes,
        "SKU Price Changes": sku_changes,
        "SLA Changes": sla_changes,
        "Line Comparison Detail": merged.drop(columns=["_merge"]),
        "Price Change Bridge": merged[TS_LINE_MATCH_KEYS + ["service_sku_last", "service_sku_this", "line_status", "added_value", "replacement_value", "removed_value", "quantity_effect", "duration_effect", "unit_price_effect", "sla_change_effect", "total_change", "bridge_total"]],
        "LDOS Exceptions": ldos_exceptions,
        "Data Quality Issues": pd.DataFrame(quality_rows),
        "Blank Field Summary": pd.DataFrame(info_rows),
        "Replacement SN Mapping": replacement_pairs,
        "Calculation Explanation": explanation,
        "Last Year Normalized": last_calc,
        "This Year Normalized": this_calc,
    }
    if not overlap_errors.empty:
        result["Overlapping Duration Errors"] = overlap_errors
    if not sniff_status.empty:
        result["SNIFF Coverage Exceptions"] = sniff_status
    if last_sniff_df is not None:
        result["Last Year SNIFF Normalized"] = last_sniff_df
    if this_sniff_df is not None:
        result["This Year SNIFF Normalized"] = this_sniff_df
    return result
