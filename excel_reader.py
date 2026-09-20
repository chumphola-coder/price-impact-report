from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO

import pandas as pd


REQUIRED_HEADERS = {
    "serial_number": ["pak/serial number", "serial number", "pak serial number"],
    "instance_number": ["instance number"],
    "service_sku": ["sku", "service sku"],
    "start_date": ["start date"],
    "end_date": ["end date"],
    "quantity": ["quantity", "qty"],
    "unit_list_price": ["unit list price"],
}

OPTIONAL_HEADERS = {
    "quote_name": ["quote name"],
    "quote_number": ["quote number"],
    "product_number": ["product number"],
    "product_description": ["product description"],
    "ldos_date": ["last date of support", "ldos"],
    "service_level": ["service level"],
    "service_level_description": ["service level description"],
    "service_type": ["service type"],
    "parent_instance_number": ["parent instance number"],
    "major_minor": ["major/minor", "major minor"],
    "prorated_list_price": ["prorated list price"],
    "extended_net_price": ["extended net price"],
    "currency": ["currency"],
    "status": ["status"],
    # SW subscription specific columns
    "subscription_id": ["subscription id"],
    "reference_serial_number": ["reference serial number"],
    "reference_instance_number": ["reference instance number"],
    "reference_subscription_id": ["reference subscription id"],
    "previous_instance_number": ["previous instance number"],
    "host_id": ["host id/mac id", "host id", "mac id"],
    "migrated_instances": ["migrated / consolidated instances", "migrated/consolidated instances"],
}

STANDARD_COLUMNS = list(REQUIRED_HEADERS) + list(OPTIONAL_HEADERS)

SNIFF_HEADERS = {
    "product_number": ["product /offer name", "product offer name", "product number"],
    "product_description": ["description"],
    "contract_number": ["subscription id/contract number"],
    "asset_status": ["status"],
    "sniff_start_date": ["start date"],
    "sniff_end_date": ["end date"],
    "serial_number": ["pak/serial number", "serial number"],
    "parent_serial_number": ["parent pak/serial number", "parent serial number"],
    "major_minor": ["major/minor", "major minor"],
    "instance_number": ["instance number"],
    "parent_instance_number": ["parent instance number"],
    "service_level": ["service level/offer type", "service level"],
    "service_sku": ["service sku", "sku"],
    "service_level_description": ["service/offer description", "service level description"],
    "quantity": ["quantity"],
    "installed_base_status": ["installed base status"],
    "replacement_serial_number": ["replacement serial number"],
    "replacement_instance_number": ["replacement instance number"],
    "product_family": ["product family"],
    "product_group": ["product group"],
    "product_sub_type": ["product sub type"],
    "ldos_date": ["last date of support"],
    "last_date_of_sale": ["last date of sale"],
    "item_type": ["item type"],
    "warranty_type": ["warranty type"],
    "warranty_status": ["warranty status"],
    "warranty_end_date": ["warranty end date"],
    "termination_date": ["termination date"],
    "renewed_from_instance": ["renewed from instance"],
    "renewed_to_instance": ["renewed to instance"],
    "total_net_price": ["total net price"],
    "monthly_cost": ["monthly cost"],
    "currency": ["currency"],
    "price_list": ["price list"],
}

SNIFF_COLUMNS = list(SNIFF_HEADERS)


@dataclass(frozen=True)
class TableDetection:
    sheet_name: str
    header_row: int
    score: int
    column_map: dict[str, str]


def _clean_header(value: object) -> str:
    return " ".join(str(value).replace("\n", " ").strip().lower().split()) if pd.notna(value) else ""


def _clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        return text[:-2]
    return text


def _find_column_index(headers: list[str], aliases: list[str]) -> int | None:
    for alias in aliases:
        for idx, header in enumerate(headers):
            if header == alias:
                return idx
    for alias in aliases:
        for idx, header in enumerate(headers):
            if header and alias in header:
                return idx
    return None


def _score_row(values: list[object]) -> tuple[int, dict[str, int]]:
    headers = [_clean_header(value) for value in values]
    column_map: dict[str, int] = {}
    score = 0

    for standard, aliases in {**REQUIRED_HEADERS, **OPTIONAL_HEADERS}.items():
        idx = _find_column_index(headers, aliases)
        if idx is not None:
            column_map[standard] = idx
            score += 10 if standard in REQUIRED_HEADERS else 2

    required_hits = sum(1 for key in REQUIRED_HEADERS if key in column_map)
    if required_hits >= 5:
        score += required_hits * 20
    return score, column_map


# Columns that appear ONLY in quote exports (never in SNIFF)
_QUOTE_ONLY_TOKENS = ["unit list price", "service type", "prorated list price", "extended net price"]
# Columns that appear ONLY in SNIFF exports (never in quotes)
_SNIFF_ONLY_TOKENS = ["installed base status", "product /offer name", "product offer name",
                      "subscription id/contract number", "replacement serial number",
                      "subscription and/or service contract entitlement"]


def _rewind(workbook: str | Path | BinaryIO) -> None:
    """Seek a file-like workbook back to the start; no-op for paths."""
    if hasattr(workbook, "seek"):
        try:
            workbook.seek(0)
        except Exception:
            pass


def _iter_sheet_frames(workbook: str | Path | BinaryIO, nrows: int | None = None):
    """Yield ``(sheet_name, raw_dataframe)`` for every readable worksheet.

    Sheets are parsed one at a time and any that fail to parse are skipped, so a
    workbook can carry any number of non-table tabs (cover sheet, BOM, notes,
    pivot, ...) alongside the real renewal table without aborting the read.
    """
    _rewind(workbook)
    xl = pd.ExcelFile(workbook, engine="openpyxl")
    for sheet_name in xl.sheet_names:
        try:
            raw = xl.parse(sheet_name, header=None, dtype=object, nrows=nrows)
        except Exception:
            continue
        yield sheet_name, raw


def detect_file_kind(workbook: str | Path | BinaryIO) -> str:
    """
    Inspect a workbook's header rows and decide whether it is a 'quote' or a
    'sniff' export. Returns 'quote', 'sniff', or 'unknown'.

    Uses columns that are unique to each export type so a quote uploaded into
    a SNIFF slot (or vice versa) can be caught before analysis. Extra/non-table
    tabs are scanned harmlessly and simply contribute no hits.
    """
    quote_hits = 0
    sniff_hits = 0
    for _, raw in _iter_sheet_frames(workbook, nrows=60):
        scan_limit = min(len(raw), 40)
        for row_idx in range(scan_limit):
            cells = [_clean_header(v) for v in raw.iloc[row_idx].tolist()]
            joined = " | ".join(cells)
            for token in _QUOTE_ONLY_TOKENS:
                if any(token == c for c in cells):
                    quote_hits += 1
            for token in _SNIFF_ONLY_TOKENS:
                if token in joined:
                    sniff_hits += 1
    if quote_hits == 0 and sniff_hits == 0:
        return "unknown"
    return "quote" if quote_hits >= sniff_hits else "sniff"


def detect_table(workbook: str | Path | BinaryIO) -> TableDetection:
    best: TableDetection | None = None
    scanned: list[str] = []

    for sheet_name, raw in _iter_sheet_frames(workbook):
        scanned.append(sheet_name)
        scan_limit = min(len(raw), 120)
        for row_idx in range(scan_limit):
            score, idx_map = _score_row(raw.iloc[row_idx].tolist())
            if not best or score > best.score:
                header_values = raw.iloc[row_idx].tolist()
                column_map = {
                    standard: str(header_values[col_idx]).strip()
                    for standard, col_idx in idx_map.items()
                    if pd.notna(header_values[col_idx])
                }
                best = TableDetection(sheet_name, row_idx, score, column_map)

    if not best or best.score < 100:
        tabs = ", ".join(scanned) if scanned else "(none)"
        raise ValueError(
            "Could not detect a renewal table header row in any worksheet. "
            f"Tabs scanned: {tabs}. Check that the workbook has serial, instance, "
            "SKU, dates, quantity, and unit price columns."
        )
    return best


def detect_sniff_table(workbook: str | Path | BinaryIO) -> TableDetection:
    best: TableDetection | None = None
    scanned: list[str] = []

    for sheet_name, raw in _iter_sheet_frames(workbook):
        scanned.append(sheet_name)
        scan_limit = min(len(raw), 120)
        for row_idx in range(scan_limit):
            headers = [_clean_header(value) for value in raw.iloc[row_idx].tolist()]
            idx_map: dict[str, int] = {}
            score = 0
            for standard, aliases in SNIFF_HEADERS.items():
                idx = _find_column_index(headers, aliases)
                if idx is not None:
                    idx_map[standard] = idx
                    score += 10 if standard in {"serial_number", "instance_number", "service_sku", "asset_status"} else 2
            if not best or score > best.score:
                header_values = raw.iloc[row_idx].tolist()
                column_map = {
                    standard: str(header_values[col_idx]).strip()
                    for standard, col_idx in idx_map.items()
                    if pd.notna(header_values[col_idx])
                }
                best = TableDetection(sheet_name, row_idx, score, column_map)

    if not best or best.score < 60:
        tabs = ", ".join(scanned) if scanned else "(none)"
        raise ValueError(
            "Could not detect a SNIFF LineDetails table header row in any worksheet. "
            f"Tabs scanned: {tabs}."
        )
    return best


def load_renewal_table(workbook: str | Path | BinaryIO, source_label: str) -> tuple[pd.DataFrame, TableDetection, pd.DataFrame]:
    detection = detect_table(workbook)
    _rewind(workbook)
    raw = pd.read_excel(
        workbook,
        sheet_name=detection.sheet_name,
        header=detection.header_row,
        dtype=object,
        engine="openpyxl",
    )
    raw = raw.dropna(how="all").copy()
    raw.columns = [str(col).strip() for col in raw.columns]

    normalized = pd.DataFrame(index=raw.index)
    for standard in STANDARD_COLUMNS:
        source_col = detection.column_map.get(standard)
        normalized[standard] = raw[source_col] if source_col in raw.columns else pd.NA

    for col in [
        "serial_number", "instance_number", "service_sku", "product_number",
        "parent_instance_number", "quote_name", "quote_number",
        "subscription_id", "reference_serial_number", "reference_instance_number",
        "reference_subscription_id", "previous_instance_number", "host_id", "migrated_instances",
    ]:
        normalized[col] = normalized[col].map(_clean_text)

    for col in ["start_date", "end_date", "ldos_date"]:
        normalized[col] = pd.to_datetime(normalized[col], errors="coerce")

    for col in ["quantity", "unit_list_price", "prorated_list_price", "extended_net_price"]:
        normalized[col] = pd.to_numeric(normalized[col], errors="coerce")

    normalized["quantity"] = normalized["quantity"].fillna(0)
    normalized["unit_list_price"] = normalized["unit_list_price"].fillna(0)
    normalized["source"] = source_label
    normalized["source_sheet"] = detection.sheet_name
    normalized["source_excel_row"] = normalized.index + detection.header_row + 2

    has_id = (
        (normalized["serial_number"] != "")
        | (normalized["instance_number"] != "")
        | (normalized["service_sku"] != "")
    )
    # Also keep genuine quote lines that carry a product number and a start date
    # but no serial/instance/SKU — e.g. NEW SUB-TNC software subscriptions, which
    # are identified by product number + subscription id rather than a serial.
    if "product_number" in normalized.columns:
        prod = normalized["product_number"].fillna("").astype(str).str.strip().str.lower()
        has_product = (prod != "") & (prod != "nan")
    else:
        has_product = pd.Series(False, index=normalized.index)
    start_ok = normalized["start_date"].notna()
    has_line = has_product & start_ok

    valid = normalized[has_id | has_line].copy()

    return valid.reset_index(drop=True), detection, raw


def load_sniff_table(workbook: str | Path | BinaryIO, source_label: str) -> tuple[pd.DataFrame, TableDetection, pd.DataFrame]:
    detection = detect_sniff_table(workbook)
    _rewind(workbook)
    raw = pd.read_excel(
        workbook,
        sheet_name=detection.sheet_name,
        header=detection.header_row,
        dtype=object,
        engine="openpyxl",
    )
    raw = raw.dropna(how="all").copy()
    raw.columns = [str(col).strip() for col in raw.columns]

    normalized = pd.DataFrame(index=raw.index)
    for standard in SNIFF_COLUMNS:
        source_col = detection.column_map.get(standard)
        normalized[standard] = raw[source_col] if source_col in raw.columns else pd.NA

    text_cols = [
        "product_number",
        "product_description",
        "contract_number",
        "asset_status",
        "serial_number",
        "parent_serial_number",
        "instance_number",
        "parent_instance_number",
        "service_level",
        "service_sku",
        "service_level_description",
        "installed_base_status",
        "replacement_serial_number",
        "replacement_instance_number",
        "product_family",
        "product_group",
        "product_sub_type",
        "item_type",
        "warranty_type",
        "warranty_status",
        "renewed_from_instance",
        "renewed_to_instance",
        "currency",
        "price_list",
    ]
    for col in text_cols:
        normalized[col] = normalized[col].map(_clean_text)

    for col in ["sniff_start_date", "sniff_end_date", "ldos_date", "last_date_of_sale", "warranty_end_date", "termination_date"]:
        normalized[col] = pd.to_datetime(normalized[col], errors="coerce")

    for col in ["quantity", "total_net_price", "monthly_cost"]:
        normalized[col] = pd.to_numeric(normalized[col], errors="coerce")

    normalized["source"] = source_label
    normalized["source_sheet"] = detection.sheet_name
    normalized["source_excel_row"] = normalized.index + detection.header_row + 2

    has_id = (
        (normalized["serial_number"] != "")
        | (normalized["instance_number"] != "")
        | (normalized["service_sku"] != "")
    )
    valid = normalized[has_id].copy()
    return valid.reset_index(drop=True), detection, raw
