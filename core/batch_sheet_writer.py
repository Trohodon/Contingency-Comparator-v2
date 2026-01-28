"""core/batch_sheet_writer.py

Writes one formatted sheet for a left/right comparison pair.

UPDATE (Limit column support):
  Pair sheets now show:
    B Contingency Events | C Resulting Issue | D Limit | E Left % | F Right % | G Δ% (Right - Left) / Status

  When expandable_issue_view=True, rows are grouped by ResultingIssue with
  the summary row showing the worst (highest max(left,right)) percent.
"""

from __future__ import annotations

import math
from typing import Callable, Optional

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


PRETTY_CASE_NAMES = {
    "ACCA_LongTerm": "ACCA LongTerm",
    "ACCA_P1,2,4,7": "ACCA",
    "DCwACver_P1-7": "DCwAC",
}


def _is_nan(x) -> bool:
    return isinstance(x, float) and math.isnan(x)


def _fmt_pct(x) -> str:
    if x is None:
        return ""
    if isinstance(x, str):
        return x
    try:
        if _is_nan(float(x)):
            return ""
        return f"{float(x):.2f}"
    except Exception:
        return str(x)


def _to_float(x) -> float:
    if x is None:
        return float("nan")
    if isinstance(x, (int, float)):
        try:
            return float(x)
        except Exception:
            return float("nan")
    s = str(x).strip().replace("%", "")
    try:
        return float(s)
    except Exception:
        return float("nan")


# --------------------------- styles --------------------------- #

title_fill = PatternFill(fill_type="solid", fgColor="305496")
title_font = Font(color="FFFFFF", bold=True, size=12)

header_fill = PatternFill(fill_type="solid", fgColor="305496")
header_font = Font(color="FFFFFF", bold=True)

data_font = Font(color="000000")
data_bold_font = Font(color="000000", bold=True)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _setup_columns(ws):
    # Columns are B..G
    widths = {
        2: 55,  # Contingency
        3: 55,  # Resulting Issue
        4: 18,  # Limit
        5: 12,  # Left %
        6: 12,  # Right %
        7: 22,  # Delta / Status
    }
    for col_idx, w in widths.items():
        ws.column_dimensions[chr(ord("A") + col_idx - 1)].width = w


def _write_title_row(ws, row: int, title: str) -> int:
    # Merge B..G
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=7)
    c = ws.cell(row=row, column=2)
    c.value = title
    c.fill = title_fill
    c.font = title_font
    c.alignment = center
    for col in range(2, 8):
        ws.cell(row=row, column=col).border = thin_border
    return row + 1


def _write_header_row(ws, row: int, left_label: str, right_label: str) -> int:
    headers = [
        ("B", "Contingency Events"),
        ("C", "Resulting Issue"),
        ("D", "Limit"),
        ("E", f"{left_label}"),
        ("F", f"{right_label}"),
        ("G", "Δ% (Right - Left) /\nStatus"),
    ]
    for col_letter, text in headers:
        col_idx = ord(col_letter) - ord("A") + 1
        c = ws.cell(row=row, column=col_idx)
        c.value = text
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = thin_border
    return row + 1


def _write_data_row(ws, row: int, cont: str, issue: str, limit: str, left_pct, right_pct, delta_text: str, bold: bool) -> int:
    font = data_bold_font if bold else data_font

    values = [cont, issue, limit, _fmt_pct(left_pct), _fmt_pct(right_pct), delta_text]
    for i, v in enumerate(values, start=2):  # B=2
        c = ws.cell(row=row, column=i)
        c.value = v
        c.font = font
        c.alignment = left_align if i in (2, 3) else center
        c.border = thin_border
    return row + 1


def _write_blank_row(ws, row: int) -> int:
    # keep borders off for blank row, but advance
    return row + 1


# --------------------------- public --------------------------- #

def write_formatted_pair_sheet(
    ws,
    df: pd.DataFrame,
    left_label: str,
    right_label: str,
    expandable_issue_view: bool = True,
    log_func: Optional[Callable[[str], None]] = None,
) -> None:
    """Write a formatted comparison sheet into an openpyxl worksheet."""
    _setup_columns(ws)

    # Excel outline behavior: summary rows ABOVE details
    ws.sheet_properties.outlinePr.summaryBelow = False

    current_row = 2  # blank row 1 (matches the builder workbook convention)

    # Expect df columns: CaseType, Contingency, ResultingIssue, Limit, LeftPct, RightPct, DeltaDisplay
    if df is None or df.empty:
        # still write a small notice
        current_row = _write_title_row(ws, current_row, f"{left_label} vs {right_label}")
        current_row = _write_header_row(ws, current_row, left_label, right_label)
        _write_data_row(ws, current_row, "No rows", "", "", "", "", "", True)
        return

    for case_type, block in df.groupby("CaseType", sort=False):
        pretty = PRETTY_CASE_NAMES.get(case_type, str(case_type))

        current_row = _write_title_row(ws, current_row, pretty)
        current_row = _write_header_row(ws, current_row, left_label, right_label)

        block = block.copy()
        block["_left_num"] = block["LeftPct"].apply(_to_float)
        block["_right_num"] = block["RightPct"].apply(_to_float)
        block["_max"] = block.apply(
            lambda r: max([v for v in [r["_left_num"], r["_right_num"]] if not _is_nan(v)] or [float("nan")]),
            axis=1,
        )

        if expandable_issue_view and "ResultingIssue" in block.columns:
            # Group rows by ResultingIssue; summary row is max(_max)
            for issue, g in (
                block.sort_values(by=["ResultingIssue", "_max"], ascending=[True, False], kind="mergesort")
                .groupby("ResultingIssue", sort=True)
            ):
                if g.empty:
                    continue

                # summary row = worst row
                g = g.sort_values(by=["_max"], ascending=[False], kind="mergesort")
                r0 = g.iloc[0]

                summary_row = current_row
                current_row = _write_data_row(
                    ws,
                    current_row,
                    str(r0.get("Contingency", "")),
                    str(r0.get("ResultingIssue", "")),
                    str(r0.get("Limit", "")),
                    r0.get("LeftPct", ""),
                    r0.get("RightPct", ""),
                    str(r0.get("DeltaDisplay", "")),
                    True,
                )

                # detail rows
                detail_start = None
                detail_end = None
                for _, rr in g.iloc[1:].iterrows():
                    if detail_start is None:
                        detail_start = current_row

                    current_row = _write_data_row(
                        ws,
                        current_row,
                        str(rr.get("Contingency", "")),
                        "",  # keep issue blank in details to visually group
                        str(rr.get("Limit", "")),
                        rr.get("LeftPct", ""),
                        rr.get("RightPct", ""),
                        str(rr.get("DeltaDisplay", "")),
                        False,
                    )
                    detail_end = current_row - 1

                if detail_start is not None and detail_end is not None:
                    ws.row_dimensions.group(
                        detail_start,
                        detail_end,
                        outline_level=1,
                        hidden=True,
                    )
                    ws.row_dimensions[summary_row].collapsed = True

        else:
            # No grouping: dump sorted by max desc
            block = block.sort_values(by=["_max"], ascending=[False], kind="mergesort")
            for _, rr in block.iterrows():
                current_row = _write_data_row(
                    ws,
                    current_row,
                    str(rr.get("Contingency", "")),
                    str(rr.get("ResultingIssue", "")),
                    str(rr.get("Limit", "")),
                    rr.get("LeftPct", ""),
                    rr.get("RightPct", ""),
                    str(rr.get("DeltaDisplay", "")),
                    False,
                )

        current_row = _write_blank_row(ws, current_row)