"""
service_tier_classifier.py
===========================
Shared Service Level Hierarchy (egrid.xlsx) + ASIA-PAC Global Price List
(glasia.xlsx) classification helpers.

Kept in its own module -- separate from both compare_engine.py and
tier_impact_engine.py -- so neither of those two needs to import the other
(tier_impact_engine already imports helpers FROM compare_engine, so the
reverse import would be circular).

Two reference files:
  - egrid.xlsx  ("Service Level Hierarchy"): one row per Service Level code
    (its CONTRACT TYPE, e.g. "SNTP", "L1NBD"), giving the Allocated Service
    Program (the "tier", e.g. "CX L2", "SMART NET TOTAL CARE") and the
    Service Level Group (the "SLA", e.g. "24X7X4", "NBD", "4HR").
  - glasia.xlsx (ASIA-PAC Global Price List): one row per Service SKU, giving
    its Service Program and Price in USD. Cisco Service SKUs follow the
    pattern CON-<CONTRACT TYPE code>-<product number>, so swapping just the
    contract-type segment lets us look up a hypothetical "what would this
    product cost under a different tier/SLA" price for the SAME hardware.
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


TARGET_TIERS = {"CX L1", "CX L2", "PARTNER ST L1", "PARTNER ST L2"}

# Display abbreviations for long Allocated Service Program names.
TIER_ABBREVIATIONS = {
    "SMART NET TOTAL CARE": "SNTC",
}

DEFAULT_EGRID_PATH = Path(__file__).parent / "egrid.xlsx"
DEFAULT_GLASIA_PATH = Path(__file__).parent / "glasia.xlsx"
DEFAULT_PRICE_CHANGE_LOG_PATH = Path(__file__).parent / "price_change_log.json"


def _norm(value: object) -> str:
    return " ".join(str(value).strip().upper().split()) if pd.notna(value) else ""


def abbreviate_tier(name: str) -> str:
    """Return the display abbreviation for a tier/service-program name, if known."""
    return TIER_ABBREVIATIONS.get(name, name)


# ── Reference file loaders ────────────────────────────────────────────────────
def load_egrid(file=None) -> dict:
    """
    Load egrid.xlsx ("Service Level Hierarchy"). Returns:
      {"by_code": {contract_type: {"category", "tier", "sla_group"}},
       "by_tier_sla": {(tier, sla_group): contract_type}}   # reverse lookup,
       first match wins when several contract types share a (tier, sla_group).
    Falls back to the bundled copy next to this module when ``file`` is None.
    """
    source = file if file is not None else DEFAULT_EGRID_PATH
    if hasattr(source, "seek"):
        source.seek(0)
    df = pd.read_excel(source, sheet_name=0, header=0, engine="openpyxl", dtype=object)
    df.columns = [str(c).strip() for c in df.columns]

    by_code: dict[str, dict] = {}
    by_tier_sla: dict[tuple[str, str], str] = {}
    for _, row in df.iterrows():
        code = _norm(row.get("CONTRACT TYPE"))
        if not code:
            continue
        category = _norm(row.get("SERVICE CATEGORY"))
        tier = _norm(row.get("ALLOCATED SERVICE PROGRAM"))
        sla_group = _norm(row.get("SERVICE LEVEL GROUP"))
        by_code[code] = {"category": category, "tier": tier, "sla_group": sla_group}
        if tier and sla_group and (tier, sla_group) not in by_tier_sla:
            by_tier_sla[(tier, sla_group)] = code
    return {"by_code": by_code, "by_tier_sla": by_tier_sla}


def load_glasia(file=None, apply_price_change_log: bool = True) -> dict[str, dict]:
    """
    Load glasia.xlsx (ASIA-PAC Global Price List). Returns a lookup dict keyed
    by Product/SKU (upper-stripped): {"tier": ..., "price": float | None}.
    Falls back to the bundled copy next to this module when ``file`` is None.

    Reads all worksheets (the list is split across multiple tabs) via calamine
    (falls back to openpyxl) since the file can be 100+ MB / 1M+ rows/sheet.

    When ``apply_price_change_log`` is True (default) and ``file`` was not
    explicitly passed (i.e. we're using the bundled default glasia.xlsx), any
    SKUs recorded in the monthly price-change log (see
    ``load_price_change_log``) have their price overlaid on top -- this lets
    the bundled snapshot stay current between full glasia.xlsx refreshes by
    layering each month's "Master Net Change Validation Report" deltas on
    top instead of re-downloading the ~120MB catalog every time.
    """
    source = file if file is not None else DEFAULT_GLASIA_PATH

    def _read(engine):
        if hasattr(source, "seek"):
            source.seek(0)
        xl = pd.ExcelFile(source, engine=engine)
        frames = [xl.parse(sheet, header=1, dtype=object, usecols=["Product", "Service Program", "Price in USD"])
                 for sheet in xl.sheet_names]
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=["Product", "Service Program", "Price in USD"])

    try:
        df = _read("calamine")
    except Exception:
        df = _read("openpyxl")

    df["Product"] = df["Product"].map(_norm)
    df["Service Program"] = df["Service Program"].map(_norm)
    df["Price in USD"] = pd.to_numeric(
        df["Price in USD"].astype(str).str.replace(r"[^0-9.\-]", "", regex=True), errors="coerce"
    )
    df = df[df["Product"] != ""]
    df = df.drop_duplicates(subset=["Product"], keep="last")
    lookup = {r["Product"]: {"tier": r["Service Program"], "price": r["Price in USD"]} for r in df.to_dict("records")}

    if apply_price_change_log and file is None:
        for sku, change in load_price_change_log().items():
            entry = lookup.setdefault(sku, {"tier": "", "price": None})
            entry["price"] = change["final"]
    return lookup


# ── Monthly price-change log ──────────────────────────────────────────────────
# Tracks incremental price changes from each month's "Master Net Change
# Validation Report" so the bundled glasia.xlsx snapshot can stay current
# without re-downloading the full ASIA-PAC price list every month. Known
# limitation (accepted): this only overlays PRICE for SKUs already present
# (or newly added), it can't detect a SKU/service-program being fully
# discontinued -- a full glasia.xlsx refresh is still needed periodically to
# catch those edge cases.
def load_price_change_log(log_path=None) -> dict[str, dict]:
    """Return {sku: {"old": float, "final": float, "start_date": str, "source": str}}."""
    path = Path(log_path) if log_path is not None else DEFAULT_PRICE_CHANGE_LOG_PATH
    if not path.exists():
        return {}
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def record_price_changes(records: list[dict], report_date: str, source_name: str, log_path=None) -> int:
    """
    Merge parsed Master Net Change Validation Report rows into the persistent
    price-change log (later reports override earlier ones per-SKU).

    ``records`` items need keys: "sku", "old", "final" (numeric or numeric-like
    strings). Returns the number of SKUs written/updated.
    """
    path = Path(log_path) if log_path is not None else DEFAULT_PRICE_CHANGE_LOG_PATH
    log = load_price_change_log(log_path)

    written = 0
    for r in records:
        sku = _norm(r.get("sku"))
        if not sku:
            continue
        try:
            old = float(r["old"])
            final = float(r["final"])
        except (KeyError, TypeError, ValueError):
            continue
        log[sku] = {"old": old, "final": final, "start_date": report_date, "source": source_name}
        written += 1

    with open(path, "w", encoding="utf-8") as f:
        json.dump(log, f, indent=1, sort_keys=True)
    return written


# ── Classification ────────────────────────────────────────────────────────────
def classify_line(service_level: object, service_sku: object,
                  egrid: dict, glasia: dict) -> tuple[str, str, str]:
    """Return (tier, sla_group, source) for one quote line.

    source is 'egrid' (primary, via the quote's Service Level code), 'glasia'
    (fallback, via the quote's SKU -- sla_group is unknown from this path), or
    'unresolved'.
    """
    hit = egrid["by_code"].get(_norm(service_level))
    if hit is not None and (hit["tier"] or hit["sla_group"]):
        return hit["tier"], hit["sla_group"], "egrid"
    hit2 = glasia.get(_norm(service_sku))
    if hit2 is not None and hit2["tier"]:
        return hit2["tier"], "", "glasia"
    return "", "", "unresolved"


def _split_sku(sku: object) -> tuple[str, str] | None:
    """Split 'CON-<contract type code>-<product>' into (code, product)."""
    s = _norm(sku)
    if not s.startswith("CON-"):
        return None
    rest = s[4:]
    if "-" not in rest:
        return None
    code, product = rest.split("-", 1)
    return code, product


def hypothetical_price(tier: str, sla_group: str, sample_sku: object,
                       egrid: dict, glasia: dict) -> float | None:
    """
    Look up the list price of a hypothetical SKU combining ``tier`` + ``sla_group``
    for the SAME hardware as ``sample_sku`` (e.g. old tier + this year's SLA
    group). Returns None if it can't be resolved from the reference data.
    """
    code = egrid.get("by_tier_sla", {}).get((tier, sla_group))
    if not code:
        return None
    split = _split_sku(sample_sku)
    if split is None:
        return None
    _, product = split
    hit = glasia.get(_norm(f"CON-{code}-{product}"))
    if hit is None or hit.get("price") is None or pd.isna(hit["price"]):
        return None
    return float(hit["price"])


def split_sla_tier_effect(service_level_last: object, service_level_this: object,
                          service_sku_last: object, service_sku_this: object,
                          unit_price_last: object, unit_price_this: object,
                          total_effect: float, egrid: dict, glasia: dict) -> tuple[float, float]:
    """
    Split ``total_effect`` (a dollar amount already computed from the raw
    unit-price delta of a matched line) into (sla_effect, tier_effect) using
    the sequential rule requested:
      1. SLA effect  = value of moving old tier + old SLA -> old tier + new SLA
                       (e.g. NBD -> 24x7x4, same service program)
      2. Tier effect = value of moving old tier + new SLA -> new tier + new SLA
                       (e.g. 4HR SNTC -> 4HR CX L2, same SLA)
    Falls back to attributing the whole effect to whichever single dimension
    actually changed, and to an even 50/50 split when a hypothetical price
    can't be resolved -- either way sla_effect + tier_effect == total_effect
    always (the split never drops or invents value).
    """
    tier_last, sla_last, _ = classify_line(service_level_last, service_sku_last, egrid, glasia)
    tier_this, sla_this, _ = classify_line(service_level_this, service_sku_this, egrid, glasia)

    tier_changed = bool(tier_last) and bool(tier_this) and tier_last != tier_this
    sla_changed = bool(sla_last) and bool(sla_this) and sla_last != sla_this

    if not tier_changed and not sla_changed:
        # Neither egrid dimension moved (e.g. unresolved codes) but the SKU/
        # service_level text still differed -- keep the legacy combined
        # behaviour (all as SLA) so nothing is silently dropped.
        return total_effect, 0.0
    if tier_changed and not sla_changed:
        return 0.0, total_effect
    if sla_changed and not tier_changed:
        return total_effect, 0.0

    # Both changed: split proportionally to the SLA-first / Tier-second
    # sequence using real (u0/u1) vs. hypothetical (old tier + new SLA)
    # unit list prices.
    u0 = 0.0 if pd.isna(unit_price_last) else float(unit_price_last)
    u1 = 0.0 if pd.isna(unit_price_this) else float(unit_price_this)
    if u1 == u0:
        return total_effect / 2.0, total_effect / 2.0

    hyp = hypothetical_price(tier_last, sla_this, service_sku_this, egrid, glasia)
    if hyp is None:
        hyp = hypothetical_price(tier_last, sla_this, service_sku_last, egrid, glasia)
    if hyp is None:
        return total_effect / 2.0, total_effect / 2.0

    ratio = (hyp - u0) / (u1 - u0)
    ratio = max(0.0, min(1.0, ratio))
    return total_effect * ratio, total_effect * (1.0 - ratio)
