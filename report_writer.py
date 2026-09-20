from __future__ import annotations

from io import BytesIO

import pandas as pd


def write_report(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
        workbook = writer.book
        money_fmt = workbook.add_format({"num_format": '$#,##0.00;[Red]-$#,##0.00;"-"'})
        integer_fmt = workbook.add_format({"num_format": '#,##0;-#,##0;"-"'})
        pct_fmt = workbook.add_format({"num_format": "0.00%"})
        date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd"})
        header_fmt = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        group_fmt = workbook.add_format({"bold": True, "bg_color": "#E2F0D9", "border": 1})
        group_center_fmt = workbook.add_format({"bold": True, "bg_color": "#E2F0D9", "border": 1, "align": "center", "valign": "vcenter"})
        group_money_fmt = workbook.add_format({"bold": True, "bg_color": "#E2F0D9", "border": 1, "num_format": "$#,##0;[Red]($#,##0)", "align": "center", "valign": "vcenter"})
        indent_fmt = workbook.add_format({"indent": 1})
        indent_section_fmt = workbook.add_format({"indent": 1, "border": 1})
        indent_section_deep_fmt = workbook.add_format({"indent": 2, "border": 1})
        indent_metric_fmt = workbook.add_format({"indent": 1, "border": 1})
        indent_metric_deep_fmt = workbook.add_format({"indent": 2, "border": 1})
        summary_text_fmt = workbook.add_format({"border": 1})
        summary_money_fmt = workbook.add_format({"num_format": "$#,##0;[Red]($#,##0)", "border": 1})
        summary_integer_fmt = workbook.add_format({"num_format": "#,##0", "border": 1})
        summary_center_fmt = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})
        summary_money_center_fmt = workbook.add_format({"num_format": "$#,##0;[Red]($#,##0)", "border": 1, "align": "center", "valign": "vcenter"})
        summary_integer_center_fmt = workbook.add_format({"num_format": "#,##0", "border": 1, "align": "center", "valign": "vcenter"})

        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            table = df.copy()
            table.to_excel(writer, sheet_name=safe_name, index=False)
            worksheet = writer.sheets[safe_name]
            if safe_name in ("Summary", "SW Subscription Analysis", "Line Detail Summary"):
                worksheet.set_tab_color("#00B050")  # green
            elif safe_name == "Line Comparison Detail":
                worksheet.set_tab_color("#FFFF00")  # yellow

            for col_idx, col_name in enumerate(table.columns):
                worksheet.write(0, col_idx, col_name, header_fmt)
                width = min(max(len(str(col_name)) + 2, 12), 45)
                if not table.empty:
                    sample_width = table[col_name].astype(str).str.len().replace(float("inf"), 0).max()
                    if pd.notna(sample_width):
                        width = min(max(width, int(sample_width) + 2), 45)
                fmt = None
                lower = str(col_name).lower()
                if any(token in lower for token in ["price", "value", "effect", "change", "total"]):
                    fmt = money_fmt
                if "pct" in lower or "percent" in lower:
                    fmt = pct_fmt
                if "date" in lower:
                    fmt = date_fmt
                worksheet.set_column(col_idx, col_idx, width, fmt)

            if len(table.columns) > 0:
                worksheet.autofilter(0, 0, max(len(table), 1), len(table.columns) - 1)
                worksheet.freeze_panes(1, 0)

            if safe_name == "Summary" and "metric" in table.columns:
                worksheet.set_zoom(120)
                cols = list(table.columns)
                section_col = cols.index("section") if "section" in cols else None
                metric_col = cols.index("metric") if "metric" in cols else None
                value_cols = [cols.index(c) for c in ["value svc ($)", "value SW ($)", "Total $"] if c in cols]
                date_col = cols.index("date") if "date" in cols else None
                formula_col = cols.index("formula / note") if "formula / note" in cols else None
                ncols = len(cols)

                # Format cache: build xlsxwriter formats on demand
                _fmt_cache: dict = {}
                def _fmt(**props):
                    key = tuple(sorted((k, str(v)) for k, v in props.items()))
                    if key not in _fmt_cache:
                        _fmt_cache[key] = workbook.add_format(dict(props))
                    return _fmt_cache[key]

                # Header row (medium outer box, thin internal grid, bold, C-F centered)
                for ci in range(ncols):
                    hp = {"bold": True, "bg_color": "#D9EAF7", "valign": "vcenter",
                          "top": 2, "bottom": 2,
                          "left": 2 if ci == 0 else 1,
                          "right": 2 if ci == ncols - 1 else 1}
                    if ci in (2, 3, 4, 5):
                        hp["align"] = "center"
                    label = cols[ci]
                    if ci == metric_col:
                        label = "Total List Price Changes"
                    elif ci == date_col:
                        label = "% change"
                    worksheet.write(0, ci, label, workbook.add_format(hp))

                # Blank-row flags (0-based table rows) for table-box border logic
                def _row_is_blank(rr):
                    return (str(rr.get("section", "") or "").strip() == ""
                            and str(rr.get("metric", "") or "").strip() == "")
                blank_flags = [_row_is_blank(table.iloc[k]) for k in range(len(table))]

                current_parent = None  # track D/E/F/G to decide count vs dollar format

                for row_idx in range(1, len(table) + 1):
                    r = table.iloc[row_idx - 1]
                    section = str(r.get("section", "") or "").strip()
                    top = section.split(".")[0] if section else ""
                    depth = section.count(".")

                    # Update parent section tracker
                    if section in ("A", "B", "C", "D", "E", "F", "G"):
                        current_parent = section

                    # NOTE rows: merge all columns, wrap text, thick border
                    if section == "NOTE":
                        note_text = str(r.get("formula / note", "") or "")
                        note_fmt = _fmt(text_wrap=True, valign="vcenter", border=2, italic=True)
                        worksheet.merge_range(row_idx, 0, row_idx, ncols - 1, note_text, note_fmt)
                        worksheet.set_row(row_idx, 75)
                        continue

                    # Count rows are sub-rows of F or G (numeric section under those headers)
                    is_count = current_parent in ("F", "G") and top.isdigit()
                    is_section_header = section in ("D", "E", "F", "G")

                    # Blank separator rows: no borders, no fill (clean gap between tables)
                    if blank_flags[row_idx - 1]:
                        for col_idx in range(ncols):
                            worksheet.write_blank(row_idx, col_idx, None, _fmt())
                        continue

                    # Background fill for coloured section headers
                    row_style: dict = {}
                    if section == "E":
                        row_style = {"bg_color": "#FFC000"}    # orange
                    elif section in ("F", "G"):
                        row_style = {"bg_color": "#FFFF00"}     # yellow
                    elif section == "D":
                        row_style = {"bg_color": "#92D050"}     # green (template)

                    # Neighbour context for thick table-box borders
                    prev_blank = row_idx - 2 >= 0 and blank_flags[row_idx - 2]
                    next_blank = row_idx < len(table) and blank_flags[row_idx]
                    is_last = row_idx == len(table)
                    prev_sec = (str(table.iloc[row_idx - 2].get("section", "") or "").strip()
                                if row_idx - 2 >= 0 else "")
                    prev_thick_bottom = (row_idx == 1) or prev_sec in ("C", "D", "E", "F", "G")

                    # Bold columns (A..F only; the formula/note column is never bold)
                    if is_section_header:
                        bold_cols = set(range(ncols))
                    elif current_parent in ("D", "E") and section == "6":
                        bold_cols = set(range(ncols))          # D.6 / E.6 grouper only
                    elif current_parent == "G" and section.isdigit():
                        bold_cols = set(range(ncols))          # G.1 .. G.6 (not sub-items)
                    elif section == "C":
                        bold_cols = set(range(ncols))          # Total increase/decrease row
                    elif section in ("A", "B"):
                        bold_cols = {0}                        # only the section letter
                    else:
                        bold_cols = set()

                    for col_idx in range(ncols):
                        value = r.iloc[col_idx]
                        cell = dict(row_style)

                        # Bold (never the formula/note column)
                        if col_idx in bold_cols and col_idx != formula_col:
                            cell["bold"] = True

                        # Thick table-box border: medium outer edges, thin internal grid
                        cell["left"] = 2 if col_idx == 0 else 1
                        cell["right"] = 2 if col_idx == ncols - 1 else 1
                        if is_section_header or prev_blank:
                            cell["top"] = 2
                        elif prev_thick_bottom:
                            cell["top"] = 0
                        else:
                            cell["top"] = 1
                        if is_section_header or section == "C" or next_blank or is_last:
                            cell["bottom"] = 2
                        else:
                            cell["bottom"] = 1

                        # Value columns: money for dollar rows, integer for count rows
                        if col_idx in value_cols:
                            cell["align"] = "center"
                            cell["valign"] = "vcenter"
                            is_total_col = cols[col_idx] == "Total $"
                            if is_count:
                                if isinstance(value, (int, float)) and not isinstance(value, bool):
                                    cell["num_format"] = "#,##0" if is_total_col else '#,##0;-#,##0;"-"'
                            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                                cell["num_format"] = '$#,##0;[Red]($#,##0);"-"'
                        # Center the date/% column (F) for the F section header and Section G rows
                        if col_idx == date_col and (
                            current_parent == "G" or (is_section_header and section == "F")
                        ):
                            cell["align"] = "center"
                            cell["valign"] = "vcenter"
                        # Indent sub-items in section/metric columns
                        if depth and col_idx in {section_col, metric_col}:
                            cell["indent"] = 2 if depth > 1 else 1

                        # % change cell (F4) = Total change / last-year total, as a percentage
                        if section == "C" and col_idx == date_col and date_col is not None:
                            pct_cell = dict(cell)
                            pct_cell["align"] = "center"
                            pct_cell["valign"] = "vcenter"
                            pct_cell["num_format"] = "0.0%"
                            worksheet.write_formula(
                                row_idx, col_idx, f"=E{row_idx + 1}/E2", _fmt(**pct_cell))
                            continue

                        if pd.notna(value) and value != "":
                            worksheet.write(row_idx, col_idx, value, _fmt(**cell))
                        else:
                            worksheet.write_blank(row_idx, col_idx, None, _fmt(**cell))

                worksheet.set_column(0, 0, 9.3)
                worksheet.set_column(1, 1, 36.5)
                worksheet.set_column(2, 2, 16.7)
                worksheet.set_column(3, 3, 16)
                if "Total $" in cols:
                    worksheet.set_column(cols.index("Total $"), cols.index("Total $"), 16)
                if date_col is not None:
                    worksheet.set_column(date_col, date_col, 13)
                if formula_col is not None:
                    worksheet.set_column(formula_col, formula_col, 110.7)
                # Suppress green "number stored as text" triangles in section column (A)
                worksheet.ignore_errors({"number_stored_as_text": f"A1:A{len(table) + 2}"})

    return output.getvalue()
