# core/batch_sheet_writer.py
#
# Sheet formatting helpers for the batch comparison workbook.
# This file writes nicely formatted Excel sheets using openpyxl.
#

from __future__ import annotations

from typing import Optional

import math
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")  # dark blue
TITLE_FILL = PatternFill("solid", fgColor="D9E1F2")   # light blue

HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FONT = Font(color="1F4E79", bold=True)

CELL_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_ALIGN_WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)

THIN = Side(style="thin", color="9AA0A6")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ---------------------------------------------------------------------------
# Existing batch pair sheet writer (unchanged)
# ---------------------------------------------------------------------------

def write_formatted_pair_sheet(
    wb: Workbook,
    sheet_name: str,
    df: pd.DataFrame,
    *,
    expandable_issue_view: bool = True,
) -> None:
    """
    Write one formatted pair-comparison sheet.
    df columns expected:
      CaseType, Contingency, ResultingIssue, LeftPct, RightPct, DeltaDisplay
    """
    ws = wb.create_sheet(title=sheet_name)

    # Headers
    headers = ["CaseType", "Contingency", "ResultingIssue", "LeftPct", "RightPct", "Delta"]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = CELL_ALIGN_CENTER
        c.border = THIN_BORDER

    # Body
    if df is None or df.empty:
        df = pd.DataFrame([{
            "CaseType": "",
            "Contingency": "No rows above threshold.",
            "ResultingIssue": "",
            "LeftPct": None,
            "RightPct": None,
            "DeltaDisplay": "",
        }])

    for r, row in enumerate(df.itertuples(index=False), start=2):
        ws.cell(row=r, column=1, value=row[0]).alignment = CELL_ALIGN_CENTER
        ws.cell(row=r, column=2, value=row[1]).alignment = CELL_ALIGN_WRAP
        ws.cell(row=r, column=3, value=row[2]).alignment = CELL_ALIGN_WRAP

        # numbers
        c4 = ws.cell(row=r, column=4, value=row[3])
        c5 = ws.cell(row=r, column=5, value=row[4])
        c6 = ws.cell(row=r, column=6, value=row[5])

        for c in (c4, c5):
            c.alignment = CELL_ALIGN_CENTER
            try:
                float(c.value)
                c.number_format = "0.000"
            except Exception:
                pass

        c6.alignment = CELL_ALIGN_CENTER

        # borders
        for col in range(1, 7):
            ws.cell(row=r, column=col).border = THIN_BORDER

    # Widths
    widths = {
        1: 14,
        2: 70,
        3: 70,
        4: 12,
        5: 12,
        6: 14,
    }
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 22


# ---------------------------------------------------------------------------
# Straight comparison sheet (all original sheets side-by-side)
# ---------------------------------------------------------------------------

def write_straight_comparison_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame) -> None:
    """
    Write a single "straight comparison" sheet where each row is a unique
    (Category, Resulting Issue, Contingency Events), and each original sheet
    gets its own Percent column.

    Expected df columns:
      - Category
      - Limit
      - Contingency Events
      - Resulting Issue
      - one or more per-sheet percent columns (typically named after the sheet)

    Formatting:
      - Blue header bar
      - Frozen header row + auto-filter
      - Wrapped text for long fields
      - Bold the highest percent value in each row across the percent columns
    """
    ws = wb.create_sheet(title=sheet_name)

    if df is None or df.empty:
        df = pd.DataFrame([{
            "Category": "",
            "Limit": "",
            "Contingency Events": "No rows above threshold.",
            "Resulting Issue": "",
        }])

    headers = list(df.columns)

    fixed_cols = {"Category", "Limit", "Contingency Events", "Resulting Issue"}
    percent_col_idxs = [i for i, h in enumerate(headers, start=1) if h not in fixed_cols]

    # Header
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CELL_ALIGN_CENTER
        cell.border = THIN_BORDER

    # Body
    for r_offset, row in enumerate(df.itertuples(index=False), start=0):
        r = 2 + r_offset

        max_pct = None
        max_col = None
        for col_idx in percent_col_idxs:
            v = row[col_idx - 1]
            try:
                fv = float(v)
                if math.isnan(fv):
                    continue
                if (max_pct is None) or (fv > max_pct):
                    max_pct = fv
                    max_col = col_idx
            except Exception:
                continue

        for col_idx, header in enumerate(headers, start=1):
            v = row[col_idx - 1]
            cell = ws.cell(row=r, column=col_idx, value=v)
            cell.border = THIN_BORDER

            if header in ("Contingency Events", "Resulting Issue"):
                cell.alignment = CELL_ALIGN_WRAP
            else:
                cell.alignment = CELL_ALIGN_CENTER

            if header == "Limit":
                try:
                    float(v)
                    cell.number_format = "0.0"
                except Exception:
                    pass

            if col_idx in percent_col_idxs:
                try:
                    float(v)
                    cell.number_format = "0.000"
                except Exception:
                    pass

            if max_col is not None and col_idx == max_col:
                cell.font = Font(bold=True)

    # Widths
    for col_idx, header in enumerate(headers, start=1):
        if header == "Category":
            width = 14
        elif header == "Limit":
            width = 10
        elif header == "Contingency Events":
            width = 65
        elif header == "Resulting Issue":
            width = 65
        else:
            width = 12
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 22