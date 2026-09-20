from __future__ import annotations

import io
import json
import re
from datetime import date as _date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from price_impact import load_master_csv, load_quote, compute_impact, write_impact_report, regenerate_quote_with_new_price

# Dedicated web-only entrypoint: Price Change Impact only (Tab 2 of
# tier_impact_app.py), with none of the Tab 1 dependencies (glasia.xlsx,
# egrid.xlsx, python-pptx/lxml) that made the full app too slow to load
# under Pyodide. price_impact.py itself only needs openpyxl/pandas/xlsxwriter.
# There is no local filesystem to save to in a browser, so this file only
# ever offers Download buttons -- no "Save to Folder" UI at all.

# Bundled default is a DISTILLED gzip CSV (sku, final, old), not the real
# ~100MB+ multi-sheet .xlsx report -- far smaller/faster to parse in-browser
# (no openpyxl overhead). See distill_master_report.py. Kept up to date
# automatically by auto_refresh_master.py -- users never upload this file.
DEFAULT_MASTER_PATH = Path(__file__).parent / "master_net_change_validation_report.csv.gz"
DEFAULT_MASTER_META_PATH = Path(__file__).parent / "master_net_change_validation_report.meta.json"

st.set_page_config(page_title="Price Change Impact Estimator", layout="wide")

# Same design system as the companion "Cisco Renewals MMM Dashboard"
# (https://wwwin-github.cisco.com/pages/camornpi/MMM/) for visual consistency
# across the two internal tools.
st.markdown("""
<style>
:root {
    --cisco-blue: #049fd9; --cisco-dark: #0b2545; --cisco-navy: #132840;
    --bg-light: #f8fafc; --card-bg: #ffffff; --text-main: #0f172a; --text-muted: #64748b;
    --border: #e2e8f0; --accent-green: #10b981; --accent-indigo: #6366f1;
    --radius: 12px; --shadow: 0 4px 14px rgba(0,0,0,0.06);
}
.block-container { padding-top: 1.4rem; padding-bottom: 1rem; max-width: 1200px; }
section[data-testid="stSidebar"] .block-container { padding-top: 1.4rem; }
section[data-testid="stSidebar"] { min-width: 320px !important; width: 320px !important; }

/* Tighten Streamlit's generous default spacing so more fits above the fold */
div[data-testid="stVerticalBlock"] { gap: 0.45rem; }
div[data-testid="stElementContainer"] { margin-bottom: 0 !important; }
.stMarkdown p { margin-bottom: 0.3rem; }

.top-header {
    background: linear-gradient(135deg, #0b2545 0%, #1e3a8a 100%);
    color: white; padding: 10px 18px; border-radius: var(--radius); margin-bottom: 8px;
    box-shadow: 0 8px 24px rgba(11,37,69,0.2);
}
.top-header h1 { font-size: 18px; font-weight: 800; display: flex; align-items: center; gap: 8px; margin: 0; flex-wrap: wrap; }
.top-header .badge {
    background: rgba(4,159,217,0.3); border: 1px solid var(--cisco-blue); color: #38bdf8;
    font-size: 10px; text-transform: uppercase; letter-spacing: 0.8px; white-space: nowrap;
    padding: 2px 7px; border-radius: 4px; font-weight: 700; vertical-align: middle;
}
.top-header p { font-size: 11.5px; color: #94a3b8; margin: 3px 0 0 0; }

.section-heading {
    font-size: 13px; font-weight: 700; color: var(--cisco-dark); text-transform: uppercase;
    letter-spacing: 0.4px; margin: 2px 0 4px 0; display: flex; align-items: center; gap: 6px;
}
.step-badge {
    display: inline-flex; align-items: center; justify-content: center; background: var(--cisco-blue);
    color: white; border-radius: 50%; width: 18px; height: 18px; font-size: 0.7rem; font-weight: 700;
}

div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: var(--radius) !important; box-shadow: var(--shadow); border-color: var(--border) !important;
}
div[data-testid="stVerticalBlockBorderWrapper"] > div > div[data-testid="stVerticalBlock"] { gap: 0.35rem; }
div[data-testid="stFileUploader"], div[data-testid="stTextInput"] input { border-radius: 8px; }
div[data-testid="stFileUploaderDropzone"] { padding: 0.6rem !important; min-height: 0 !important; }
button[kind="primary"] {
    background-color: var(--cisco-blue) !important; border-color: var(--cisco-blue) !important;
    box-shadow: 0 2px 8px rgba(4,159,217,0.35);
}
button[kind="primary"]:hover { background-color: #0387ba !important; border-color: #0387ba !important; }
div[data-testid="stDownloadButton"] > button {
    background-color: var(--accent-green); color: white; border: 1px solid var(--accent-green);
}
div[data-testid="stDownloadButton"] > button:hover {
    background-color: #0d9668; border-color: #0d9668; color: white;
}
/* Compact metric tiles so a 2x2 grid stays short */
div[data-testid="stMetric"] {
    background: var(--bg-light); border: 1px solid var(--border); border-radius: 8px;
    padding: 6px 10px;
}
div[data-testid="stMetricLabel"] p {
    font-size: 10px; color: var(--text-muted); white-space: normal !important; overflow: visible !important;
}
div[data-testid="stMetricValue"] {
    font-size: 14px; white-space: normal !important; overflow: visible !important; word-break: break-word;
}
div[data-testid="stMetricValue"] p {
    font-size: 14px; white-space: normal !important; overflow: visible !important;
    text-overflow: clip !important; word-break: break-word;
}
</style>
""", unsafe_allow_html=True)


def _extract_date_suffix(filename: str) -> str:
    """
    Pull a trailing date token out of a filename (e.g. the Master Net Change
    Validation Report is usually named "...12 Sep 26.xlsx") and return it
    formatted as "12-Sep-26". Falls back to today's date if none can be found
    or parsed, so the resulting suffix is always meaningful.
    """
    stem = Path(filename).stem
    match = re.search(r"(\d{1,2})[\s\-_]+([A-Za-z]{3,9}|\d{1,2})[\s\-_]+(\d{2,4})\s*$", stem)
    if match:
        day, month, year = match.groups()
        for fmt in ("%d %b %y", "%d %B %y", "%d %b %Y", "%d %B %Y", "%d %m %y", "%d %m %Y"):
            try:
                return datetime.strptime(f"{day} {month} {year}", fmt).strftime("%d-%b-%y")
            except ValueError:
                continue
    return _date.today().strftime("%d-%b-%y")


def _default_master_date_suffix() -> str:
    if DEFAULT_MASTER_META_PATH.exists():
        try:
            meta = json.loads(DEFAULT_MASTER_META_PATH.read_text())
            if meta.get("date_suffix"):
                return meta["date_suffix"]
        except (json.JSONDecodeError, OSError):
            pass
    return _extract_date_suffix(DEFAULT_MASTER_PATH.name)


@st.cache_resource(show_spinner=False)
def _cached_master(mtime: float):
    """Cache the bundled (distilled, gzip CSV) Master Net Change Validation Report by file mtime -- reloads only if the file on disk changes."""
    return load_master_csv(DEFAULT_MASTER_PATH)


st.markdown("""
<div class="top-header">
  <h1>\U0001f4b0 Price Change Impact Estimator <span class="badge">Browser-based</span></h1>
  <p>See how a Cisco service price change affects your renewal quotes \u2014 nothing is uploaded anywhere.</p>
</div>
""", unsafe_allow_html=True)
with st.expander("ℹ️ How it works"):
    st.write(
        "1. The bundled **Master Net Change Validation Report** (Cisco's old/new price per Service SKU) "
        "is kept up to date automatically \u2014 no upload needed.\n"
        "2. Upload one or more **quote files** in the sidebar \u2014 the tool matches each line's SKU against the master.\n"
        "3. For every matched line, it recalculates the prorated list price at the new rate "
        "(unit price \u00d7 qty \u00d7 (days + 1) / 365) and shows the total dollar impact."
    )

with st.sidebar:
    st.header("📄 Quote file(s) to check")
    quote_files = st.file_uploader(
        "Quote file(s) to check", type=["xlsx"], key="pi_quote",
        accept_multiple_files=True, label_visibility="collapsed",
        help="Drop one or more renewal quote files. Both multi-quote and single-quote formats are auto-detected and combined."
    )
    if not quote_files:
        st.caption("Drop one or more quote files here to get started.")

    # Duplicate filenames can't be told apart once bundled into one download,
    # so block rather than silently rename -- ask the user to fix it instead.
    _dupe_names = sorted({
        n for n in (f.name for f in (quote_files or []))
        if list(f.name for f in quote_files).count(n) > 1
    })
    if _dupe_names:
        st.error(
            "\u26d4 Duplicate file name(s): " + ", ".join(f"**{n}**" for n in _dupe_names) + ". "
            "Rename or remove the duplicate \u2014 the buttons below are disabled until this is fixed."
        )

    _master_available = DEFAULT_MASTER_PATH.exists()
    if _master_available:
        st.caption(f"✅ Master price list: **{_default_master_date_suffix()}** "
                  f"(updated {pd.Timestamp(DEFAULT_MASTER_PATH.stat().st_mtime, unit='s'):%Y-%m-%d %H:%M}, "
                  "kept current automatically).")
    else:
        st.warning("Bundled Master Net Change Validation Report is missing — contact the app owner.")

    pi_can_run = _master_available and bool(quote_files) and not _dupe_names

    st.divider()
    st.header("⚙️ Output file names")
    _pi_sidebar_name = st.text_input(
        "Impact report", value="Price_Change_Impact.xlsx", key="tab2_filename",
        help="Output file name for the Price Change Impact report."
    )
    _pi_sidebar_name = _pi_sidebar_name.strip().strip('"').strip("'").strip()
    if not _pi_sidebar_name.lower().endswith(".xlsx"):
        _pi_sidebar_name += ".xlsx"
    _pi_newprice_sidebar_name = st.text_input(
        "Repriced quote", value="Quote_New_Price.xlsx", key="tab2_newprice_filename",
        help="Output file name for the regenerated quote(s) with new prices. Multiple quote files are bundled into a .zip."
    )
    _pi_newprice_sidebar_name = _pi_newprice_sidebar_name.strip().strip('"').strip("'").strip()

    st.divider()
    pi_clicked = st.button(
        "📊 Generate Price Impact Report", key="pi_generate",
        type="primary", disabled=not pi_can_run, use_container_width=True,
        help="Upload quote file(s) above to enable." if not pi_can_run else None,
    )
    newprice_clicked = st.button(
        "📝 Generate Quote(s) with New Price", key="pi_newprice_generate",
        type="primary", disabled=not pi_can_run, use_container_width=True,
        help="Upload quote file(s) above to enable." if not pi_can_run else None,
    )

result_col1, result_col2 = st.columns(2)

with result_col1:
    st.markdown('<div class="section-heading"><span class="step-badge">1</span>Price impact report</div>', unsafe_allow_html=True)
    with st.container(border=True):
        st.write("Summarize the total dollar impact across your quote(s).")

        if pi_clicked:
            try:
                with st.spinner("Loading master price-change list..."):
                    master_lookup = _cached_master(DEFAULT_MASTER_PATH.stat().st_mtime)
                if not master_lookup:
                    st.error("Master file loaded 0 SKUs \u2014 the bundled master file may be corrupt.")
                    st.stop()
                st.caption(f"Master loaded: {len(master_lookup):,} SKUs")

                with st.spinner("Loading quote file(s)..."):
                    frames = []
                    for qf in quote_files:
                        qf.seek(0)
                        frames.append(load_quote(io.BytesIO(qf.read()), source_name=qf.name))
                    quote_df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
                st.caption(f"Quote loaded: {len(quote_df):,} lines from {len(quote_files)} file(s)")

                with st.spinner("Computing price impact..."):
                    detail_df = compute_impact(master_lookup, quote_df)
                    report_bytes = write_impact_report(detail_df, "")

                st.session_state["pi_report_bytes"] = report_bytes

                matched = detail_df[detail_df["matched"]]
                orig  = matched["orig_prorated"].sum()
                new_  = matched["new_prorated"].sum()
                chg   = new_ - orig
                pct   = chg / orig * 100 if orig else 0.0
                skus  = int(matched["sku"].nunique())
                full_orig = detail_df["orig_prorated"].sum()

                st.success(f"✅ {len(matched)} matched lines across {skus} SKU(s)")

                m1, m2 = st.columns(2)
                m1.metric("Full Quote", f"${full_orig:,.2f}")
                m2.metric("Matched",    f"${orig:,.2f}")
                m3, m4 = st.columns(2)
                m3.metric("New Price",  f"${new_:,.2f}")
                m4.metric("Impact",     f"${chg:+,.2f}", delta=f"{pct:+.2f}%")

            except Exception as exc:
                st.error(str(exc))
                st.exception(exc)

        if "pi_report_bytes" in st.session_state:
            st.download_button(
                "\u2b07 Download Price Impact Report (.xlsx)",
                data=st.session_state["pi_report_bytes"],
                file_name=_pi_sidebar_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pi_download",
                use_container_width=True,
            )

with result_col2:
    st.markdown('<div class="section-heading"><span class="step-badge">2</span>Regenerate quote(s) with new pricing</div>', unsafe_allow_html=True)
    with st.container(border=True):
        st.write(
            "Reproduces your quote(s) as-is except each matched line's Unit List Price is updated "
            "to the new rate. Quote total/name updated too."
        )

        if newprice_clicked:
            try:
                if _dupe_names:
                    st.error("Duplicate quote file name(s) uploaded -- fix that first (see the sidebar message).")
                    st.stop()

                with st.spinner("Loading master price-change list..."):
                    master_lookup = _cached_master(DEFAULT_MASTER_PATH.stat().st_mtime)
                if not master_lookup:
                    st.error("Master file loaded 0 SKUs \u2014 the bundled master file may be corrupt.")
                    st.stop()

                with st.spinner("Regenerating quote(s) with new prices..."):
                    results = []
                    for qf in quote_files:
                        qf.seek(0)
                        out_bytes, stats = regenerate_quote_with_new_price(qf, master_lookup)
                        results.append((qf.name, out_bytes, stats))

                total_matched = sum(r[2]["matched_lines"] for r in results)
                st.success(f"✅ Regenerated {len(results)} file(s), {total_matched} line(s) repriced.")
                for name, _bytes, stats in results:
                    st.caption(f"**{name}**: {stats['matched_lines']} line(s) repriced across {len(stats['sheets_updated'])} quote tab(s).")

                date_suffix = _default_master_date_suffix()
                named_results = [
                    (f"{Path(name).stem} - Estimate-Price-{date_suffix}.xlsx", out_bytes, stats)
                    for name, out_bytes, stats in results
                ]

                if len(named_results) == 1:
                    filename, out_bytes, _ = named_results[0]
                    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                else:
                    import zipfile
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                        for name, out_bytes, _ in named_results:
                            zf.writestr(name, out_bytes)
                    out_bytes = zip_buf.getvalue()
                    stem = _pi_newprice_sidebar_name.rsplit(".", 1)[0]
                    filename = f"{stem} - Estimate-Price-{date_suffix}.zip"
                    mime = "application/zip"

                st.session_state["pi_newprice_bytes"] = out_bytes
                st.session_state["pi_newprice_filename"] = filename
                st.session_state["pi_newprice_mime"] = mime

            except Exception as exc:
                st.error(str(exc))
                st.exception(exc)

        if "pi_newprice_bytes" in st.session_state:
            st.download_button(
                "\u2b07 Download Quote(s) with New Price",
                data=st.session_state["pi_newprice_bytes"],
                file_name=st.session_state["pi_newprice_filename"],
                mime=st.session_state["pi_newprice_mime"],
                key="pi_newprice_download",
                use_container_width=True,
            )

