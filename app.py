from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from compare_engine import compare_renewals
from excel_reader import detect_file_kind, load_renewal_table, load_sniff_table
from ppt_writer import write_ppt
from report_writer import write_report
from price_impact import load_master, load_quote, compute_impact, write_impact_report


st.set_page_config(page_title="Renewal Quote Comparator", layout="wide")

# Reduce Streamlit's large default top padding so content sits near the top edge.
st.markdown("""
<style>
.block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
section[data-testid="stSidebar"] .block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

def _next_available_path(path: Path, sibling_suffixes: tuple[str, ...] = ()) -> Path:
    """Return a non-colliding save path, browser-style.

    If ``path`` (and, when given, its siblings sharing the same stem but the
    ``sibling_suffixes`` extensions) does not exist, return it unchanged.
    Otherwise append " (n)" before the suffix using the smallest n >= 1 for
    which the target and all its siblings are free — e.g. ``abc.xlsx`` ->
    ``abc (1).xlsx`` -> ``abc (2).xlsx``. ``sibling_suffixes`` keeps a paired
    file (like a matching ``.pptx``) on the same running number.
    """
    parent, stem, suffix = path.parent, path.stem, path.suffix

    def _all_free(candidate_stem: str) -> bool:
        for suf in (suffix, *sibling_suffixes):
            if (parent / f"{candidate_stem}{suf}").exists():
                return False
        return True

    if _all_free(stem):
        return path
    n = 1
    while True:
        candidate_stem = f"{stem} ({n})"
        if _all_free(candidate_stem):
            return parent / f"{candidate_stem}{suffix}"
        n += 1


st.title("Renewal Quote Comparator")
st.write("Compare last year and this year renewal Excel files locally. No AI API or cloud upload is used.")

tab1, tab2 = st.tabs(["📊 Renewal Comparison", "💰 Price Change Impact"])

with tab1:

    # ── Color theme for last year (blue) vs this year (orange) ─────────────────────
    st.markdown("""
    <style>
    /* Last year uploaders → blue accent */
    div[data-testid="stFileUploader"]:has(#lastyear_marker),
    section[data-testid="stSidebar"] .lastyear-box {
        border-left: 5px solid #1f77b4;
        padding-left: 8px;
        background: rgba(31, 119, 180, 0.06);
        border-radius: 6px;
    }
    .lastyear-title { color: #1f77b4; font-weight: 700; font-size: 0.95rem; margin: 4px 0; }
    .thisyear-title { color: #d9660a; font-weight: 700; font-size: 0.95rem; margin: 4px 0; }
    .lastyear-box { border-left: 5px solid #1f77b4; padding: 6px 10px; background: rgba(31,119,180,0.06);
                    border-radius: 6px; margin-bottom: 6px; }
    .thisyear-box { border-left: 5px solid #d9660a; padding: 6px 10px; background: rgba(217,102,10,0.06);
                    border-radius: 6px; margin-bottom: 6px; }
    </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.header("Output File")
        default_folder = str(Path.home() / "Desktop")
        save_folder = st.text_input("Save folder", value=default_folder,
                                    help="Folder where reports will be saved. Used by both tabs.")
        # Sanitize: copy-pasted paths often carry surrounding quotes/whitespace.
        save_folder = save_folder.strip().strip('"').strip("'").strip()
        save_folder = str(Path(save_folder).expanduser())
        save_filename = st.text_input(
            "Tab 1 file name", value="renewal_comparison.xlsx", key="tab1_filename",
            help="Output file name for the Renewal Comparison report."
        )
        save_filename = save_filename.strip().strip('"').strip("'").strip()
        if not save_filename.lower().endswith(".xlsx"):
            save_filename += ".xlsx"
        _pi_sidebar_name = st.text_input(
            "Tab 2 file name", value="Price_Change_Impact.xlsx", key="tab2_filename",
            help="Output file name for the Price Change Impact report."
        )
        _pi_sidebar_name = _pi_sidebar_name.strip().strip('"').strip("'").strip()
        if not _pi_sidebar_name.lower().endswith(".xlsx"):
            _pi_sidebar_name += ".xlsx"

        st.markdown("**PowerPoint summary**")
        ppt_customer = st.text_input("Customer name", value="",
                                     help="Shown on the PPT cover slide. Also appended as suffix to output filenames.")
        ppt_period = st.text_input("Period label", value="",
                                   help="e.g. 'Renewal FY26 vs FY25'. Shown on the PPT cover slide.")

        # Compute effective filenames: append customer name as suffix when set
        _cust_safe = ppt_customer.strip().replace(" ", "_") if ppt_customer.strip() else ""
        def _with_cust(fname: str) -> str:
            if not _cust_safe:
                return fname
            stem = fname[:-5] if fname.lower().endswith(".xlsx") else fname
            return f"{stem}_{_cust_safe}.xlsx"
        save_filename_eff = _with_cust(save_filename)
        _pi_name_eff = _with_cust(_pi_sidebar_name)
        if _cust_safe:
            st.caption(f"Tab 1 will save as: **{save_filename_eff}**")
            st.caption(f"Tab 2 will save as: **{_pi_name_eff}**")

    # ── File uploaders (top of the main page, 4 boxes in one row) ──────────────────
    st.subheader("Upload Files")
    up_col1, up_col2, up_col3, up_col4 = st.columns(4)

    with up_col1:
        st.markdown('<div class="lastyear-title">🔵 LAST YEAR — Quote</div>', unsafe_allow_html=True)
        last_files = st.file_uploader("Last year CCWR quote / All quotes",
                                      type=["xlsx", "xlsm"], key="last",
                                      accept_multiple_files=True,
                                      help="You can drop several quote files here; they will be combined.")

    with up_col2:
        st.markdown('<div class="thisyear-title">🟠 THIS YEAR — Quote</div>', unsafe_allow_html=True)
        this_files = st.file_uploader("This year CCWR quote / All quotes",
                                      type=["xlsx", "xlsm"], key="this",
                                      accept_multiple_files=True,
                                      help="You can drop several quote files here; they will be combined.")

    with up_col3:
        st.markdown('<div class="lastyear-title">🔵 LAST YEAR — SNIFF (optional)</div>', unsafe_allow_html=True)
        last_sniff_files = st.file_uploader("Last year SNIFF Excel",
                                            type=["xlsx", "xlsm"], key="last_sniff",
                                            accept_multiple_files=True,
                                            help="You can drop several SNIFF files here; they will be combined.")

    with up_col4:
        st.markdown('<div class="thisyear-title">🟠 THIS YEAR — SNIFF (optional)</div>', unsafe_allow_html=True)
        this_sniff_files = st.file_uploader("This year SNIFF Excel",
                                            type=["xlsx", "xlsm"], key="this_sniff",
                                            accept_multiple_files=True,
                                            help="You can drop several SNIFF files here; they will be combined.")

    if not last_files or not this_files:
        st.caption("Upload Last year and This year quote files to enable the report.")

    # ── Validate uploads (used to enable/disable the Generate button) ──────────────
    def _seek0(f):
        try:
            f.seek(0)
        except Exception:
            pass

    def _detect_kind_cached(f) -> str:
        """detect_file_kind with a per-session cache keyed by file name+size, so we
        don't re-read the workbook on every Streamlit rerun."""
        cache = st.session_state.setdefault("_kind_cache", {})
        key = (getattr(f, "name", ""), getattr(f, "size", None))
        if key not in cache:
            _seek0(f)
            try:
                cache[key] = detect_file_kind(f)
            except Exception:
                cache[key] = "unknown"
            _seek0(f)
        return cache[key]

    upload_issues: list[str] = []
    if not last_files:
        upload_issues.append("Upload at least one **Last year** quote file.")
    if not this_files:
        upload_issues.append("Upload at least one **This year** quote file.")

    for files, expected, label in [
        (last_files, "quote", "Last year quote"),
        (this_files, "quote", "This year quote"),
        (last_sniff_files, "sniff", "Last year SNIFF"),
        (this_sniff_files, "sniff", "This year SNIFF"),
    ]:
        for f in (files or []):
            kind = _detect_kind_cached(f)
            if kind != "unknown" and kind != expected:
                upload_issues.append(
                    f"**{label}** — '{f.name}' looks like a **{kind.upper()}** file, "
                    f"not a {expected.upper()}. Move it to the correct box."
                )

    can_generate = not upload_issues

    # ── Green style for the download button ───────────────────────────────────────
    st.markdown("""
    <style>
    div[data-testid="stDownloadButton"] > button {
        background-color: #28a745;
        color: white;
        border: 1px solid #28a745;
    }
    div[data-testid="stDownloadButton"] > button:hover {
        background-color: #218838;
        border-color: #1e7e34;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Action bar: Generate + Save-to-folder + Download Excel + Download PPT ───────
    action_col1, action_col2, action_col3, action_col4 = st.columns([1, 1, 1, 1])
    generate_clicked = action_col1.button(
        "Generate Comparison", type="primary", disabled=not can_generate,
        help=None if can_generate else "Upload valid Last year and This year quote files to enable.",
    )
    save_slot = action_col2.empty()          # "Save to Folder" button
    download_slot = action_col3.empty()      # "Download Excel" button
    ppt_download_slot = action_col4.empty()  # "Download PowerPoint" button

    if upload_issues:
        st.warning("**Generate Comparison** is disabled. Fix the following:\n\n"
                   + "\n".join(f"- {m}" for m in upload_issues))

    def _render_actions() -> None:
        """Render Save-to-Folder and Download buttons using the LIVE folder/filename.

        Reading save_folder / save_filename here (not from session_state) means the
        user can change the file name or folder any time — before or after running
        the analysis — and the very next click uses the new value.

        Guarded so it renders at most once per script run (it may be called both at
        the top of the script and again right after generation).
        """
        if st.session_state.get("_actions_rendered"):
            return
        if "report_bytes" not in st.session_state:
            return
        st.session_state["_actions_rendered"] = True
        report_bytes = st.session_state["report_bytes"]

        if save_slot.button("💾 Save to Folder", key="save_to_folder"):
            out_path = Path(save_folder) / save_filename_eff
            has_ppt = bool(st.session_state.get("ppt_bytes"))
            try:
                out_path.parent.mkdir(parents=True, exist_ok=True)
                out_path = _next_available_path(out_path, (".pptx",) if has_ppt else ())
                out_path.write_bytes(report_bytes)
                saved_msg = f"Excel report saved to: {out_path}"
                if has_ppt:
                    ppt_path = out_path.with_suffix(".pptx")
                    ppt_path.write_bytes(st.session_state["ppt_bytes"])
                    saved_msg += f"\n\nPowerPoint saved to: {ppt_path}"
                st.success(saved_msg)
            except Exception as save_err:
                st.error(f"Could not save to '{out_path}': {save_err}")

        download_slot.download_button(
            "⬇ Download Excel Report",
            data=report_bytes,
            file_name=save_filename_eff,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_report",
        )

        if "ppt_bytes" in st.session_state:
            ppt_name = save_filename_eff.rsplit(".", 1)[0] + ".pptx"
            ppt_download_slot.download_button(
                "⬇ Download PowerPoint",
                data=st.session_state["ppt_bytes"],
                file_name=ppt_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="download_ppt",
            )

    # Show Save/Download if a report already exists from a previous run
    st.session_state["_actions_rendered"] = False  # reset the once-per-run guard
    _render_actions()

    if generate_clicked:
        try:
            # ── Guard 1: verify each file is the right kind for its slot ───────
            def _seek0(f):
                try:
                    f.seek(0)
                except Exception:
                    pass

            wrong_slot = []
            slot_groups = [
                (last_files, "quote", "Last year quote"),
                (this_files, "quote", "This year quote"),
                (last_sniff_files, "sniff", "Last year SNIFF"),
                (this_sniff_files, "sniff", "This year SNIFF"),
            ]
            for files, expected, label in slot_groups:
                for f in (files or []):
                    _seek0(f)
                    kind = detect_file_kind(f)
                    _seek0(f)
                    if kind != "unknown" and kind != expected:
                        wrong_slot.append(f"• **{label}** — '{f.name}' looks like a **{kind.upper()}** file, not a {expected.upper()}.")

            if wrong_slot:
                st.error(
                    "The uploaded files appear to be in the wrong slots:\n\n"
                    + "\n".join(wrong_slot)
                    + "\n\nPlease check the blue (last year) and orange (this year) upload boxes and try again."
                )
                st.stop()

            # ── Loaders that combine multiple files into one table ─────────────
            def _load_quotes(files, label):
                dfs, detections = [], []
                for f in files:
                    _seek0(f)
                    df, det, _ = load_renewal_table(f, label)
                    dfs.append(df)
                    detections.append((f.name, det))
                return pd.concat(dfs, ignore_index=True), detections

            def _load_sniffs(files, label):
                dfs, detections = [], []
                for f in (files or []):
                    _seek0(f)
                    df, det, _ = load_sniff_table(f, label)
                    dfs.append(df)
                    detections.append((f.name, det))
                combined = pd.concat(dfs, ignore_index=True) if dfs else None
                return combined, detections

            with st.spinner("Reading last year quote file(s)..."):
                last_df, last_dets = _load_quotes(last_files, "last_year")
            with st.spinner("Reading this year quote file(s)..."):
                this_df, this_dets = _load_quotes(this_files, "this_year")

            # ── Guard 2: year sanity check (are last/this swapped?) ────────────
            def _median_start(df: pd.DataFrame):
                if "start_date" in df.columns:
                    s = pd.to_datetime(df["start_date"], errors="coerce").dropna()
                    if not s.empty:
                        return s.median()
                return None

            last_med = _median_start(last_df)
            this_med = _median_start(this_df)
            if last_med is not None and this_med is not None and last_med > this_med:
                st.warning(
                    f"⚠️ The **last year** file's typical start date ({last_med.date()}) is **later** than "
                    f"the **this year** file ({this_med.date()}). The two quote files may be swapped. "
                    "Double-check the blue (last year) and orange (this year) boxes."
                )

            with st.spinner("Reading last year SNIFF file(s)..."):
                last_sniff_df, last_sniff_dets = _load_sniffs(last_sniff_files, "last_year_sniff")
            with st.spinner("Reading this year SNIFF file(s)..."):
                this_sniff_df, this_sniff_dets = _load_sniffs(this_sniff_files, "this_year_sniff")

            with st.spinner("Calculating comparison..."):
                sheets = compare_renewals(last_df, this_df, last_sniff_df, this_sniff_df)
                report = write_report(sheets)

            # Store in session state so the buttons persist across reruns
            st.session_state["report_bytes"] = report

            with st.spinner("Building PowerPoint summary..."):
                try:
                    st.session_state["ppt_bytes"] = write_ppt(
                        sheets, customer_name=ppt_customer, period_label=ppt_period
                    )
                except Exception as ppt_err:
                    st.session_state.pop("ppt_bytes", None)
                    st.warning(f"PowerPoint could not be generated: {ppt_err}")

            # Fill the top Save/Download slots immediately on this same run
            _render_actions()

            # ── Overlapping-duration error (split orders must be sequential) ───
            overlap_df = sheets.get("Overlapping Duration Errors")
            if overlap_df is not None and not overlap_df.empty:
                st.error(
                    f"⛔ Found {len(overlap_df)} overlapping coverage period(s) in the quote files. "
                    "Split orders for the same SN#/instance#/SKU must be sequential (not overlapping). "
                    "See the 'Overlapping Duration Errors' sheet in the report."
                )
                st.dataframe(overlap_df, use_container_width=True)

            st.success("Comparison report generated.")

            col1, col2 = st.columns(2)
            col1.metric("Last year rows", len(last_df))
            col2.metric("This year rows", len(this_df))

            st.subheader("Detected Tables")
            detected = []
            for fname, det in last_dets:
                detected.append({"slot": "Last year quote", "file": fname, "sheet": det.sheet_name,
                                 "header_row_excel": det.header_row + 1, "score": det.score})
            for fname, det in this_dets:
                detected.append({"slot": "This year quote", "file": fname, "sheet": det.sheet_name,
                                 "header_row_excel": det.header_row + 1, "score": det.score})
            for fname, det in last_sniff_dets:
                detected.append({"slot": "Last year SNIFF", "file": fname, "sheet": det.sheet_name,
                                 "header_row_excel": det.header_row + 1, "score": det.score})
            for fname, det in this_sniff_dets:
                detected.append({"slot": "This year SNIFF", "file": fname, "sheet": det.sheet_name,
                                 "header_row_excel": det.header_row + 1, "score": det.score})
            st.dataframe(detected, use_container_width=True)

            st.subheader("Summary")
            # Streamlit/pyarrow can't render columns that mix text (COUNT/QTY) with
            # numbers, so show a stringified copy for display only.
            summary_display = sheets["Summary"].copy()
            for _col in summary_display.columns:
                summary_display[_col] = summary_display[_col].map(
                    lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
                )
            st.dataframe(summary_display, use_container_width=True)

            # ── Auto-save to disk on generation (using current folder/filename) ──
            out_path = Path(save_folder) / save_filename_eff
            has_ppt = bool(st.session_state.get("ppt_bytes"))
            try:
                out_path.parent.mkdir(parents=True, exist_ok=True)
                out_path = _next_available_path(out_path, (".pptx",) if has_ppt else ())
                out_path.write_bytes(report)
                saved_msg = f"Excel report saved to: {out_path}"
                if has_ppt:
                    ppt_path = out_path.with_suffix(".pptx")
                    ppt_path.write_bytes(st.session_state["ppt_bytes"])
                    saved_msg += f"\n\nPowerPoint saved to: {ppt_path}"
                st.success(saved_msg)
                st.caption("To save under a different name or folder, change the fields in the sidebar "
                           "and click **💾 Save to Folder** above — no need to regenerate.")
            except Exception as save_err:
                st.warning(f"Could not auto-save ({save_err}). Change the folder in the sidebar and click "
                           "**💾 Save to Folder**, or use the **⬇ Download** button above.")

        except Exception as exc:
            st.error(str(exc))
            st.exception(exc)

    # ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Price Change Impact Analysis
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Price Change Impact Analysis")
    st.write(
        "Upload the Master Net Change Validation Report and one quote file. "
        "The engine matches Service SKUs, applies the master's price-change ratio "
        "to each line's Unit List Price, and recomputes the prorated list using "
        "unit × qty × (days + 1) / 365."
    )

    pc_col1, pc_col2 = st.columns(2)
    with pc_col1:
        master_file = st.file_uploader(
            "Master Net Change Validation Report (.xlsx)",
            type=["xlsx"], key="pi_master",
            help="The Cisco master price-change file. First sheet, col C = SKU, col G = Final Price, col H = Old Price."
        )
    with pc_col2:
        quote_files = st.file_uploader(
            "Quote file(s) to check (.xlsx)",
            type=["xlsx"], key="pi_quote",
            accept_multiple_files=True,
            help="Drop one or more renewal quote files. Both multi-quote and single-quote formats are auto-detected and combined."
        )

    pi_can_run = master_file is not None and bool(quote_files)
    pi_clicked = st.button(
        "Generate Price Impact Report", key="pi_generate",
        type="primary", disabled=not pi_can_run,
        help="Upload both files to enable." if not pi_can_run else None,
    )

    if "pi_report_bytes" in st.session_state:
        _pi_btn_col1, _pi_btn_col2 = st.columns(2)
        if _pi_btn_col1.button("💾 Save to Folder", key="pi_save_to_folder"):
            _pi_save_path = Path(save_folder) / _pi_name_eff
            try:
                _pi_save_path.parent.mkdir(parents=True, exist_ok=True)
                _pi_save_path = _next_available_path(_pi_save_path)
                _pi_save_path.write_bytes(st.session_state["pi_report_bytes"])
                st.success(f"Saved to: {_pi_save_path}")
            except Exception as _e:
                st.error(f"Could not save: {_e}")
        _pi_btn_col2.download_button(
            "⬇ Download Price Impact Report (.xlsx)",
            data=st.session_state["pi_report_bytes"],
            file_name=_pi_name_eff,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="pi_download",
        )

    if pi_clicked:
        try:
            import io as _io

            @st.cache_data(show_spinner=False)
            def _cached_load_master(file_bytes: bytes) -> dict:
                """Cache master by file content so repeat runs are instant."""
                return load_master(_io.BytesIO(file_bytes))

            with st.spinner("Loading master price-change list (first load ~5 s, cached after)..."):
                master_file.seek(0)
                master_lookup = _cached_load_master(master_file.read())
            if not master_lookup:
                st.error("Master file loaded 0 SKUs — check that you uploaded the "
                         "correct Master Net Change Validation Report file.")
                st.stop()
            st.caption(f"Master loaded: {len(master_lookup):,} SKUs")

            with st.spinner("Loading quote file(s)..."):
                import pandas as _pd
                frames = []
                for qf in quote_files:
                    qf.seek(0)
                    frames.append(load_quote(_io.BytesIO(qf.read()), source_name=qf.name))
                quote_df = _pd.concat(frames, ignore_index=True) if frames else _pd.DataFrame()
            st.caption(f"Quote loaded: {len(quote_df):,} lines from {len(quote_files)} file(s)")

            with st.spinner("Computing price impact..."):
                detail_df = compute_impact(master_lookup, quote_df)
                report_bytes = write_impact_report(detail_df, ppt_customer)

            st.session_state["pi_report_bytes"] = report_bytes

            matched = detail_df[detail_df["matched"]]
            orig  = matched["orig_prorated"].sum()
            new_  = matched["new_prorated"].sum()
            chg   = new_ - orig
            pct   = chg / orig * 100 if orig else 0.0
            skus  = int(matched["sku"].nunique())
            full_orig = detail_df["orig_prorated"].sum()

            st.success(f"Report generated: {len(matched)} matched lines across {skus} SKU(s)")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Full Quote Prorated Total", f"${full_orig:,.2f}")
            m2.metric("Matched Lines Prorated",    f"${orig:,.2f}")
            m3.metric("New Prorated (matched)",    f"${new_:,.2f}")
            m4.metric("Price Impact",              f"${chg:+,.2f}", delta=f"{pct:+.2f}%")

            # auto-save using the sidebar Tab 2 filename
            out_path = Path(save_folder) / _pi_name_eff
            try:
                out_path.parent.mkdir(parents=True, exist_ok=True)
                out_path = _next_available_path(out_path)
                out_path.write_bytes(report_bytes)
                st.success(f"Saved to: {out_path}")
                st.caption("To save under a different name, change 'Tab 2 — Price Change Impact file name' "
                           "in the sidebar and use the ⬇ Download button above.")
            except Exception as save_err:
                st.warning(f"Could not auto-save ({save_err}). Use the download button.")

            st.download_button(
                "⬇ Download Price Impact Report (.xlsx)",
                data=report_bytes,
                file_name=_pi_name_eff,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pi_download_inline",
            )

            with st.expander("Top impacted SKUs (matched lines only)"):
                grp = (
                    matched.groupby("sku")
                    .agg(lines=("sku","count"), change=("prorated_change","sum"))
                    .reset_index()
                    .sort_values("change", ascending=False)
                )
                grp["change"] = grp["change"].map(lambda v: f"${v:+,.2f}")
                st.dataframe(grp.rename(columns={"sku":"SKU","lines":"Lines","change":"Prorated Change"}),
                             use_container_width=True, hide_index=True)

        except Exception as exc:
            st.error(str(exc))
            st.exception(exc)
