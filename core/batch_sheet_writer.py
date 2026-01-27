# batch_sheet_writer.py
# Writes a formatted "pair comparison" worksheet into an output workbook.

from __future__ import annotations

import os
import time
from typing import Optional

import pandas as pd

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def write_formatted_pair_sheet(
    wb,
    sheet_name: str,
    df_pair: pd.DataFrame,
    *,
    expandable_issue_view: bool = True,
) -> None:
    """
    Writes one comparison sheet in the style you prefer:
      - groups by Resulting Issue (LimViolID)
      - sorts groups by max percent (worst first)
      - within group sorts by CTG max percent
      - bolds the highest line(s) within each group
      - optional "expandable" layout for Excel +/- controls
    """
    ws = wb.create_sheet(sheet_name)

    # Styles
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(bold=True, color="FFFFFF")
    subheader_fill = PatternFill("solid", fgColor="D9E1F2")
    group_fill = PatternFill("solid", fgColor="F2F2F2")
    bold_font = Font(bold=True)
    normal_font = Font(bold=False)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    thin = Side(style="thin", color="A6A6A6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Columns to output
    # NOTE: These internal names come from comparator.build_case_type_comparison
    # and match your existing formatting.
    col_casetype = "CaseType"
    col_ctg = "CTGLabel"
    col_issue = "LimViolID"
    col_left = "LimViolPct (Left)"
    col_right = "LimViolPct (Right)"
    col_delta = "Delta"
    col_max = "MaxPct"

    # If any are missing, fail gracefully
    for req in [col_casetype, col_ctg, col_issue, col_left, col_right, col_delta, col_max]:
        if req not in df_pair.columns:
            raise ValueError(f"Missing required column in df_pair: {req}")

    df = df_pair.copy()

    # Group max for ordering
    df["__group_max"] = df.groupby([col_issue])[col_max].transform("max")
    df = df.sort_values(by=["__group_max", col_max], ascending=[False, False], na_position="last")

    # Header row
    headers = [
        "Contingency Event",
        "Resulting Issue",
        "Left %",
        "Right %",
        "Δ (Right-Left)",
    ]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    ws.row_dimensions[1].height = 22

    # Write rows
    start_row = 2
    rows = []
    for _, r in df.iterrows():
        rows.append(
            (
                str(r[col_ctg]) if r[col_ctg] is not None else "",
                str(r[col_issue]) if r[col_issue] is not None else "",
                r[col_left],
                r[col_right],
                r[col_delta],
                r[col_max],  # for bolding
            )
        )

    # Expandable layout: group header rows + indented children
    current_issue = None
    group_start_row = None
    group_rows = []

    def flush_group():
        nonlocal group_rows, group_start_row
        if not group_rows:
            return

        # Determine max row(s) in this group for bolding
        vals = [gr[5] for gr in group_rows if gr[5] is not None]
        vmax = max(vals) if vals else None

        if expandable_issue_view:
            # Add a group header row (issue name at top)
            header_row = group_rows[0][1]
            ws.cell(row=group_start_row, column=1, value="").alignment = left
            ws.cell(row=group_start_row, column=2, value=header_row).alignment = left
            for c in range(1, 6):
                ws.cell(row=group_start_row, column=c).fill = group_fill
                ws.cell(row=group_start_row, column=c).border = border
            ws.cell(row=group_start_row, column=2).font = bold_font

            # Child rows start after group header
            child_start = group_start_row + 1
            for i, gr in enumerate(group_rows):
                rr = child_start + i
                ctg, issue, l, r, d, m = gr
                ws.cell(row=rr, column=1, value=ctg).alignment = left
                ws.cell(row=rr, column=2, value="").alignment = left  # keep issue blank under group
                ws.cell(row=rr, column=3, value=l)
                ws.cell(row=rr, column=4, value=r)
                ws.cell(row=rr, column=5, value=d)

                # formats
                for c in [3, 4, 5]:
                    ws.cell(row=rr, column=c).alignment = center
                    ws.cell(row=rr, column=c).number_format = "0.000"

                # bold max within group
                if vmax is not None and m is not None and abs(m - vmax) < 1e-9:
                    for c in range(1, 6):
                        ws.cell(row=rr, column=c).font = bold_font
                else:
                    for c in range(1, 6):
                        ws.cell(row=rr, column=c).font = normal_font

                for c in range(1, 6):
                    ws.cell(row=rr, column=c).border = border

            # Outline grouping for Excel +/- at TOP of group
            # Children are grouped under the header row.
            ws.row_dimensions[group_start_row].outlineLevel = 0
            for rr in range(child_start, child_start + len(group_rows)):
                ws.row_dimensions[rr].outlineLevel = 1
                ws.row_dimensions[rr].hidden = True

        else:
            # Non-expandable: just write normal rows with issue repeated
            for i, gr in enumerate(group_rows):
                rr = group_start_row + i
                ctg, issue, l, r, d, m = gr
                ws.cell(row=rr, column=1, value=ctg).alignment = left
                ws.cell(row=rr, column=2, value=issue).alignment = left
                ws.cell(row=rr, column=3, value=l)
                ws.cell(row=rr, column=4, value=r)
                ws.cell(row=rr, column=5, value=d)

                for c in [3, 4, 5]:
                    ws.cell(row=rr, column=c).alignment = center
                    ws.cell(row=rr, column=c).number_format = "0.000"

                if vmax is not None and m is not None and abs(m - vmax) < 1e-9:
                    for c in range(1, 6):
                        ws.cell(row=rr, column=c).font = bold_font
                else:
                    for c in range(1, 6):
                        ws.cell(row=rr, column=c).font = normal_font

                for c in range(1, 6):
                    ws.cell(row=rr, column=c).border = border

        group_rows = []
        group_start_row = None

    r_idx = start_row
    for row in rows:
        # Yield periodically to keep UI responsive while writing large sheets
        if (r_idx - start_row) % 300 == 0:
            time.sleep(0)

        ctg, issue, l, r, d, m = row
        if issue != current_issue:
            flush_group()
            current_issue = issue
            group_start_row = r_idx
            group_rows = []
            if expandable_issue_view:
                # reserve one row for group header
                r_idx += 1

        group_rows.append((ctg, issue, l, r, d, m))
        r_idx += 1

    flush_group()

    # Sheet settings
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "C2"  # keeps header and first 2 columns

    # Column widths
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14

    # Enable outline symbols
    ws.sheet_properties.outlinePr.summaryBelow = True
    ws.sheet_properties.outlinePr.summaryRight = True