# core/comparator.py
#
# Helpers for working with the formatted
# Combined_ViolationCTG_Comparison.xlsx workbook and for building
# batch comparison workbooks in a nicely formatted style.
#
# Public functions used by the GUI:
#   - list_sheets(workbook_path)
#   - build_case_type_comparison(...)
#   - compare_scenarios(...)
#   - build_pair_comparison_df(...)
#   - build_batch_comparison_workbook(...)
#
# NOTE: This module intentionally does not depend on GUI code.

from __future__ import annotations

import os
import re
import math
from dataclasses import dataclass
from typing import Callable, Dict, List, Optional, Tuple

import pandas as pd

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Canonical case type handling
# ---------------------------------------------------------------------------

# Reverse mapping so we can label rows nicely when exporting
CANONICAL_TO_PRETTY = {
    "ACCA_LongTerm": "ACCA LongTerm",
    "ACCA_P1,2,4,7": "ACCA",
    "DCwACver_P1-7": "DCwAC",
}


# ---------------------------------------------------------------------------
# Basic workbook helpers
# ---------------------------------------------------------------------------

def list_sheets(workbook_path: str) -> List[str]:
    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        return wb.sheetnames
    finally:
        wb.close()


def _sanitize_sheet_name(name: str) -> str:
    """
    Make a string safe to use as an Excel sheet name.
    """
    invalid = set(r'[]:*?/\\')
    cleaned = "".join(ch if ch not in invalid else "_" for ch in name)
    cleaned = cleaned.strip()
    if not cleaned:
        cleaned = "Sheet"

    # Excel limit
    return cleaned[:31]


# ===== Formatting helpers for batch workbook =================================

def _apply_table_styles(ws: Worksheet):
    """
    Set reasonable column widths for a formatted comparison sheet.
    """
    widths = {
        2: 45,  # B: Contingency Events
        3: 45,  # C: Resulting Issue
        4: 15,  # D: Left %
        5: 15,  # E: Right %
        6: 22,  # F: Delta
    }
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # No freeze panes: scroll normally
    # ws.freeze_panes = "B4"


# Styles (approximate the blue style from the first tab)
HEADER_FILL = PatternFill("solid", fgColor="305496")  # dark blue
HEADER_FONT = Font(color="FFFFFF", bold=True)
TITLE_FILL = HEADER_FILL
TITLE_FONT = Font(color="FFFFFF", bold=True, size=12)

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

CELL_ALIGN_WRAP = Alignment(wrap_text=True, vertical="top")
CELL_ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _is_nan(x) -> bool:
    return isinstance(x, float) and math.isnan(x)


def _max_pct(left: Optional[float], right: Optional[float]) -> float:
    vals = []
    if left is not None and not _is_nan(left):
        vals.append(float(left))
    if right is not None and not _is_nan(right):
        vals.append(float(right))
    return max(vals) if vals else float("-inf")


def _normalize_issue_series(series: pd.Series) -> pd.Series:
    """
    Forward-fill blanks so that blank ResultingIssue inherits the issue above.

    Avoid pandas Series.replace(...) due to pandas version quirks that can throw:
      "'regex' must be a string ... you passed a 'bool'"
    """
    s = series.copy()

    def is_blank(v) -> bool:
        if v is None:
            return True
        if isinstance(v, float) and math.isnan(v):
            return True
        if isinstance(v, str) and v.strip() == "":
            return True
        return False

    mask = s.apply(is_blank)
    s = s.mask(mask)
    s = s.ffill()
    return s.fillna("")


def _write_title_row(ws: Worksheet, row: int, title: str):
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
    cell = ws.cell(row=row, column=2)
    cell.value = title
    cell.fill = TITLE_FILL
    cell.font = TITLE_FONT
    cell.alignment = CELL_ALIGN_CENTER


def _write_header_row(ws: Worksheet, row: int):
    headers = [
        "Contingency Events",
        "Resulting Issue",
        "Left %",
        "Right %",
        "Δ% (Right - Left) / Status",
    ]
    for col_offset, header in enumerate(headers):
        cell = ws.cell(row=row, column=2 + col_offset)
        cell.value = header
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CELL_ALIGN_CENTER
        cell.border = THIN_BORDER


def _write_data_row(
    ws: Worksheet,
    row: int,
    cont: str,
    issue: str,
    left_pct,
    right_pct,
    delta,
    outline_level: int = 0,
    hidden: bool = False,
    bold: bool = False,
):
    values = [cont, issue, left_pct, right_pct, delta]

    for col_offset, val in enumerate(values):
        cell = ws.cell(row=row, column=2 + col_offset)
        cell.value = val
        cell.border = THIN_BORDER

        if bold:
            base = cell.font or Font()
            cell.font = Font(
                name=base.name,
                size=base.size,
                bold=True,
                italic=base.italic,
                vertAlign=base.vertAlign,
                underline=base.underline,
                strike=base.strike,
                color=base.color,
            )

        if col_offset in (0, 1):
            cell.alignment = CELL_ALIGN_WRAP
        else:
            cell.alignment = Alignment(horizontal="right", vertical="top")

        if col_offset in (2, 3) and isinstance(val, (float, int)):
            cell.number_format = "0.00"

    try:
        ws.row_dimensions[row].outlineLevel = int(outline_level)
        ws.row_dimensions[row].hidden = bool(hidden)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Case-type comparison builder (single pair)
# ---------------------------------------------------------------------------

def build_case_type_comparison(
    workbook_path: str,
    left_sheet: str,
    right_sheet: str,
    threshold: float,
    log_func: Optional[Callable[[str], None]] = None,
) -> pd.DataFrame:
    """
    Build a consolidated comparison DF between two sheets in the formatted workbook.
    Output columns:
      CaseType, Contingency, ResultingIssue, LeftPct, RightPct, DeltaDisplay
    """
    if log_func:
        log_func(f"Loading workbook: {workbook_path}")

    wb = load_workbook(workbook_path, data_only=True)
    try:
        if left_sheet not in wb.sheetnames:
            raise ValueError(f"Left sheet not found: {left_sheet}")
        if right_sheet not in wb.sheetnames:
            raise ValueError(f"Right sheet not found: {right_sheet}")

        ws_left = wb[left_sheet]
        ws_right = wb[right_sheet]

        df_left = _extract_sheet_as_df(ws_left, threshold, log_func=log_func, side="Left")
        df_right = _extract_sheet_as_df(ws_right, threshold, log_func=log_func, side="Right")

        df_pair = build_pair_comparison_df(df_left, df_right, threshold=threshold)
        return df_pair
    finally:
        wb.close()


def _extract_sheet_as_df(
    ws: Worksheet,
    threshold: float,
    log_func: Optional[Callable[[str], None]] = None,
    side: str = "",
) -> pd.DataFrame:
    """
    Extract rows from the formatted sheet into a DF.

    Expected sheet structure:
      Blocks per case type, with a blue title row, then a blue header row, then data rows.
      Columns:
        B: Contingency Events
        C: Resulting Issue
        D: Contingency Value (MVA)
        E: Percent Loading

    We only keep rows where Percent Loading >= threshold.
    """
    rows = []
    current_case_type = ""

    max_row = ws.max_row
    for r in range(1, max_row + 1):
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value
        d = ws.cell(row=r, column=4).value
        e = ws.cell(row=r, column=5).value

        if isinstance(b, str) and b.strip() in ("ACCA LongTerm", "ACCA", "DCwAC"):
            current_case_type = b.strip()
            continue

        if current_case_type and isinstance(b, str) and b.strip() == "Contingency Events":
            continue

        if not current_case_type:
            continue

        if b is None and c is None and d is None and e is None:
            continue

        try:
            pct = float(e) if e is not None else None
        except Exception:
            pct = None

        if pct is None:
            continue

        if pct < threshold:
            continue

        cont = str(b) if b is not None else ""
        issue = str(c) if c is not None else ""

        rows.append(
            {
                "CaseType": current_case_type,
                "Contingency": cont,
                "ResultingIssue": issue,
                "Pct": pct,
                "MVA": float(d) if d is not None else None,
            }
        )

    df = pd.DataFrame(rows)

    if log_func:
        log_func(f"  Extracted {len(df)} rows >= {threshold:.2f}% from {side} sheet.")

    if not df.empty:
        df["ResultingIssue"] = _normalize_issue_series(df["ResultingIssue"])

    return df


# ---------------------------------------------------------------------------
# Pair DF merge logic (left/right)
# ---------------------------------------------------------------------------

def build_pair_comparison_df(
    df_left: pd.DataFrame,
    df_right: pd.DataFrame,
    *,
    threshold: float,
) -> pd.DataFrame:
    """
    Merge left/right extracted DFs into a single DF.

    Output columns:
      CaseType, Contingency, ResultingIssue, LeftPct, RightPct, DeltaDisplay
    """
    if df_left is None:
        df_left = pd.DataFrame()
    if df_right is None:
        df_right = pd.DataFrame()

    key_cols = ["CaseType", "Contingency", "ResultingIssue"]

    left = df_left.copy()
    right = df_right.copy()

    if not left.empty:
        left = left.rename(columns={"Pct": "LeftPct", "MVA": "LeftMVA"})
    else:
        left = pd.DataFrame(columns=key_cols + ["LeftPct", "LeftMVA"])

    if not right.empty:
        right = right.rename(columns={"Pct": "RightPct", "MVA": "RightMVA"})
    else:
        right = pd.DataFrame(columns=key_cols + ["RightPct", "RightMVA"])

    merged = pd.merge(left, right, on=key_cols, how="outer")

    def make_delta(r):
        lp = r.get("LeftPct", None)
        rp = r.get("RightPct", None)

        if pd.isna(lp) and pd.isna(rp):
            return ""
        if pd.isna(lp):
            return "Only in right"
        if pd.isna(rp):
            return "Only in left"
        try:
            d = float(rp) - float(lp)
            if abs(d) < 1e-9:
                return "No Change"
            return f"{d:+.2f}"
        except Exception:
            return ""

    merged["DeltaDisplay"] = merged.apply(make_delta, axis=1)

    def row_sort_key(r):
        return _max_pct(r.get("LeftPct", None), r.get("RightPct", None))

    merged["_SortKey"] = merged.apply(row_sort_key, axis=1)
    merged = merged.sort_values(by=["CaseType", "_SortKey"], ascending=[True, False])

    merged = merged.drop(columns=["_SortKey"])

    for col in ["LeftPct", "RightPct", "LeftMVA", "RightMVA"]:
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce")

    return merged


# ---------------------------------------------------------------------------
# Scenario compare sheet (one sheet inside same workbook)
# ---------------------------------------------------------------------------

def compare_scenarios(
    workbook_path: str,
    base_sheet: str,
    new_sheet: str,
    output_path: str,
    threshold: float,
    mode: str,
    log_func: Optional[Callable[[str], None]] = None,
):
    """
    Compare two sheets (base/new) and write a comparison sheet into a workbook.

    mode:
      - "all": keep all merged rows above threshold
      - "only_changed": keep only where Delta is nonzero or row exists in one side
    """
    if log_func:
        log_func(f"Comparing '{base_sheet}' vs '{new_sheet}' (threshold={threshold}, mode={mode})")

    wb = load_workbook(workbook_path)
    try:
        if base_sheet not in wb.sheetnames:
            raise ValueError(f"Base sheet not found: {base_sheet}")
        if new_sheet not in wb.sheetnames:
            raise ValueError(f"New sheet not found: {new_sheet}")

        df_base = _extract_sheet_as_df(wb[base_sheet], threshold, log_func=log_func, side="Base")
        df_new = _extract_sheet_as_df(wb[new_sheet], threshold, log_func=log_func, side="New")

        df_pair = build_pair_comparison_df(df_base, df_new, threshold=threshold)

        merged = df_pair.copy()

        status_col = []
        for _, r in merged.iterrows():
            lp = r.get("LeftPct", None)
            rp = r.get("RightPct", None)
            if pd.isna(lp) and not pd.isna(rp):
                status_col.append("Only in new")
            elif not pd.isna(lp) and pd.isna(rp):
                status_col.append("Only in base")
            else:
                try:
                    d = float(rp) - float(lp)
                    if abs(d) < 1e-9:
                        status_col.append("Unchanged")
                    else:
                        status_col.append("Changed")
                except Exception:
                    status_col.append("Changed")

        merged["Status"] = status_col

        def keep_row(r) -> bool:
            base_pct = r["LeftPct"]
            new_pct = r["RightPct"]
            status = r["Status"]

            if status in ("Only in new", "Only in base"):
                return True

            if pd.isna(base_pct) or pd.isna(new_pct):
                return True

            delta = new_pct - base_pct

            if mode == "only_changed":
                return abs(delta) > 1e-9
            return True

        merged = merged[merged.apply(keep_row, axis=1)].copy()

        # Nicer sheet name (spaces, no _vs_)
        base_short = _sanitize_sheet_name(base_sheet).replace(" ", "")
        new_short = _sanitize_sheet_name(new_sheet).replace(" ", "")
        comp_name = f"Compare {base_short} vs {new_short}"
        comp_name = _sanitize_sheet_name(comp_name)

        if comp_name in wb.sheetnames:
            wb.remove(wb[comp_name])

        ws = wb.create_sheet(title=comp_name)

        cols = [
            "CaseType",
            "Contingency",
            "ResultingIssue",
            "LeftMVA",
            "LeftPct",
            "RightMVA",
            "RightPct",
            "DeltaDisplay",
            "Status",
        ]
        ws.append(cols)

        for _, r in merged.iterrows():
            ws.append([r.get(c, "") for c in cols])

        wb.save(output_path)

        if log_func:
            log_func(f"Wrote scenario compare sheet: {output_path}")

    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Batch comparison workbook builder
# ---------------------------------------------------------------------------

def build_batch_comparison_workbook(
    src_workbook_path: str,
    pairs: List[Tuple[str, str]],
    output_path: str,
    threshold: float,
    expandable_issue_view: bool,
    log_func: Optional[Callable[[str], None]] = None,
):
    """
    Build a brand-new .xlsx workbook with one sheet per (left_sheet, right_sheet) pair,
    Each sheet is grouped into ACCA LongTerm / ACCA / DCwAC blocks, with columns:
      Contingency Events | Resulting Issue | Left % | Right % | Δ% / Status

    If expandable_issue_view=True:
      - group by ResultingIssue
      - sort each issue group's rows by max(LeftPct, RightPct) desc
      - top row (max) is visible & bold; remaining rows hidden with outlineLevel=1
      - Resulting Issue is blank for hidden rows
      - summary is ABOVE detail rows (so +/- appears at TOP like the comparison builder output)
    """
    if log_func:
        log_func(f"Building batch workbook: {output_path}")
        log_func(f"  Source workbook: {src_workbook_path}")
        log_func(f"  Pairs: {len(pairs)}")
        log_func(f"  Threshold: {threshold}")
        log_func(f"  Expandable issue view: {expandable_issue_view}")

    wb = Workbook()

    # Remove the default sheet; we'll create our own
    default_sheet = wb.active
    wb.remove(default_sheet)

    used_names = set()

    for idx, (left_sheet, right_sheet) in enumerate(pairs, start=1):
        if log_func:
            log_func(f"Processing pair {idx}: '{left_sheet}' vs '{right_sheet}'...")

        try:
            df_pair = build_case_type_comparison(
                src_workbook_path, left_sheet, right_sheet, threshold, log_func=log_func
            )
        except Exception as e:
            if log_func:
                log_func(f"  ERROR: Failed to build DF for pair {idx}: {e}")
            df_pair = pd.DataFrame(
                [
                    {
                        "CaseType": "",
                        "Contingency": f"ERROR building comparison for: {left_sheet} vs {right_sheet}",
                        "ResultingIssue": "",
                        "LeftPct": None,
                        "RightPct": None,
                        "DeltaDisplay": "",
                    }
                ]
            )

        if df_pair is None or df_pair.empty:
            df_pair = pd.DataFrame(
                [
                    {
                        "CaseType": "",
                        "Contingency": "No rows above threshold.",
                        "ResultingIssue": "",
                        "LeftPct": None,
                        "RightPct": None,
                        "DeltaDisplay": "",
                    }
                ]
            )

        # Sheet name: keep it clean and readable (no numbering)
        # Example: "Base Case vs Breaker Test 1"
        base_name = f"{left_sheet} vs {right_sheet}"
        base_name = _sanitize_sheet_name(base_name)

        name = base_name
        counter = 2
        while name in used_names:
            suffix = f"_{counter}"
            name = _sanitize_sheet_name(base_name[: (31 - len(suffix))] + suffix)
            counter += 1

        used_names.add(name)

        _write_formatted_pair_sheet(
            wb, name, df_pair, expandable_issue_view=expandable_issue_view
        )

    wb.save(output_path)

    if log_func:
        log_func(f"Batch workbook saved: {output_path}")


def _write_formatted_pair_sheet(
    wb: Workbook,
    ws_name: str,
    df_pair: pd.DataFrame,
    *,
    expandable_issue_view: bool,
):
    ws = wb.create_sheet(title=ws_name)

    _apply_table_styles(ws)

    # Outline behavior: summary ABOVE details => +/- appears at the TOP
    try:
        ws.sheet_properties.outlinePr.summaryBelow = False
        ws.sheet_properties.outlinePr.applyStyles = True
    except Exception:
        pass

    current_row = 2

    for case_type_pretty in ["ACCA LongTerm", "ACCA", "DCwAC"]:
        sub = df_pair[df_pair["CaseType"] == case_type_pretty].copy()
        if sub.empty:
            continue

        if "ResultingIssue" not in sub.columns:
            sub["ResultingIssue"] = ""
        sub["ResultingIssue"] = _normalize_issue_series(sub["ResultingIssue"])

        _write_title_row(ws, current_row, case_type_pretty)
        current_row += 1
        _write_header_row(ws, current_row)
        current_row += 1

        if not expandable_issue_view:
            for _, r in sub.iterrows():
                cont = str(r.get("Contingency", "") or "")
                issue = str(r.get("ResultingIssue", "") or "")
                left_pct = r.get("LeftPct", None)
                right_pct = r.get("RightPct", None)
                delta = str(r.get("DeltaDisplay", "") or "")

                _write_data_row(
                    ws,
                    current_row,
                    cont,
                    issue,
                    left_pct,
                    right_pct,
                    delta,
                    outline_level=0,
                    hidden=False,
                    bold=False,
                )
                current_row += 1

            current_row += 1
            continue

        # Expandable issue view
        sub["_SortKey"] = sub.apply(
            lambda r: _max_pct(r.get("LeftPct", None), r.get("RightPct", None)),
            axis=1,
        )

        group_max = (
            sub.groupby("ResultingIssue")["_SortKey"]
            .max()
            .sort_values(ascending=False)
        )
        ordered_issues = list(group_max.index)

        for issue_key in ordered_issues:
            g = sub[sub["ResultingIssue"] == issue_key].copy()
            if g.empty:
                continue

            g = g.sort_values(by="_SortKey", ascending=False, na_position="last")

            first = True
            for _, r in g.iterrows():
                cont = str(r.get("Contingency", "") or "")
                issue = str(r.get("ResultingIssue", "") or "")
                issue_display = issue if first else ""

                left_pct = r.get("LeftPct", None)
                right_pct = r.get("RightPct", None)
                delta = str(r.get("DeltaDisplay", "") or "")

                _write_data_row(
                    ws,
                    current_row,
                    cont,
                    issue_display,
                    left_pct,
                    right_pct,
                    delta,
                    outline_level=0 if first else 1,
                    hidden=False if first else True,
                    bold=True if first else False,
                )
                current_row += 1
                first = False

        current_row += 1
