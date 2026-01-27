
# core/straight_comparison.py
#
# Builds a "Straight Comparison" sheet that compares ALL original scenario sheets
# side-by-side (Base Case, Breaker 1, Breaker 2, ...), grouped by Resulting Issue.
#
# This is intentionally separate from batch pair comparison so we don't disturb
# existing workbook styles/behavior.
#
# Output style matches the blue-block look used elsewhere:
# - Case-type blocks (ACCA LongTerm / ACCA / DCwAC)
# - Excel outline (+/-) grouping by Resulting Issue
# - Summary row ABOVE details (so +/- appears at the TOP)
# - First (max) row per Resulting Issue is bolded

from __future__ import annotations

from typing import Dict, List, Optional, Sequence

import math
import os

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# --- Case type mappings (kept consistent with comparator.py) -----------------

CANONICAL_CASE_TYPES = {
    "ACCA LongTerm": "ACCA_LongTerm",
    "ACCA Long Term": "ACCA_LongTerm",
    "ACCA": "ACCA_P1,2,4,7",
    "DCwAC": "DCwACver_P1-7",
}

CASE_TYPES_CANONICAL: List[str] = [
    "ACCA_LongTerm",
    "ACCA_P1,2,4,7",
    "DCwACver_P1-7",
]

CANONICAL_TO_PRETTY = {
    "ACCA_LongTerm": "ACCA LongTerm",
    "ACCA_P1,2,4,7": "ACCA",
    "DCwACver_P1-7": "DCwAC",
}


# --- Formatting (same vibe as batch_sheet_writer.py) -------------------------

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


def apply_straight_table_styles(ws: Worksheet, num_cases: int):
    # Column widths
    ws.column_dimensions[get_column_letter(2)].width = 45  # B
    ws.column_dimensions[get_column_letter(3)].width = 45  # C

    # Case % columns (D..)
    for i in range(num_cases):
        ws.column_dimensions[get_column_letter(4 + i)].width = 14

    # Outline: +/- visible, summary row ABOVE details
    try:
        ws.sheet_properties.outlinePr.summaryBelow = False
        ws.sheet_properties.outlinePr.summaryRight = False
        ws.sheet_properties.outlinePr.applyStyles = True
    except Exception:
        pass

    try:
        ws.sheet_view.showOutlineSymbols = True
    except Exception:
        pass


def _write_title_row(ws: Worksheet, row: int, title: str, last_col: int):
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
    cell = ws.cell(row=row, column=2)
    cell.value = title
    cell.fill = TITLE_FILL
    cell.font = TITLE_FONT
    cell.alignment = CELL_ALIGN_CENTER


def _write_header_row(ws: Worksheet, row: int, case_labels: Sequence[str]):
    headers = ["Contingency Events", "Resulting Issue"] + list(case_labels)
    for col_offset, header in enumerate(headers):
        cell = ws.cell(row=row, column=2 + col_offset)
        cell.value = header
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CELL_ALIGN_CENTER
        cell.border = THIN_BORDER


def _is_nan(x) -> bool:
    return isinstance(x, float) and math.isnan(x)


def _row_max(values: Sequence[Optional[float]]) -> float:
    v = []
    for x in values:
        if x is None:
            continue
        if _is_nan(x):
            continue
        try:
            v.append(float(x))
        except Exception:
            pass
    return max(v) if v else float("-inf")


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def _parse_scenario_sheet(ws, log_func=None) -> pd.DataFrame:
    """
    Parse one formatted scenario sheet into:
      CaseType, CTGLabel, LimViolID, LimViolValue, LimViolPct

    Forward-fills LimViolID (Resulting Issue) within each CaseType block.
    """
    records: List[Dict] = []

    max_row = ws.max_row or 1
    row_idx = 1

    while row_idx <= max_row:
        title_val = ws.cell(row=row_idx, column=2).value  # column B

        if isinstance(title_val, str) and title_val.strip():
            pretty_name = title_val.strip()
            case_type = CANONICAL_CASE_TYPES.get(pretty_name, pretty_name)

            header_row = row_idx + 1
            data_row = header_row + 1

            last_issue = None
            r = data_row
            while r <= max_row:
                b = ws.cell(row=r, column=2).value
                c = ws.cell(row=r, column=3).value
                d = ws.cell(row=r, column=4).value
                e = ws.cell(row=r, column=5).value

                if _is_blank(b) and _is_blank(c) and _is_blank(d) and _is_blank(e):
                    break

                if _is_blank(c) and last_issue is not None:
                    c = last_issue
                else:
                    if not _is_blank(c):
                        last_issue = c

                records.append(
                    {
                        "CaseType": case_type,
                        "CTGLabel": b,
                        "LimViolID": c,
                        "LimViolValue": d,
                        "LimViolPct": e,
                    }
                )
                r += 1

            row_idx = r + 1
        else:
            row_idx += 1

    return pd.DataFrame.from_records(records)


def _load_sheet_as_df(workbook_path: str, sheet_name: str, log_func=None) -> pd.DataFrame:
    wb = load_workbook(workbook_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
    ws = wb[sheet_name]
    return _parse_scenario_sheet(ws, log_func=log_func)


def discover_scenario_sheets(workbook_path: str, log_func=None) -> List[str]:
    """
    Return scenario sheet names in their existing order.

    We filter out non-scenario tabs by checking for known case-type titles
    somewhere in column B near the top of the sheet.
    """
    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    titles = set(CANONICAL_CASE_TYPES.keys())

    scenario_names: List[str] = []
    for name in wb.sheetnames:
        ws = wb[name]
        is_scenario = False
        # Look in column B for a title row ("ACCA LongTerm", "ACCA", "DCwAC")
        for r in range(1, min(250, (ws.max_row or 1)) + 1):
            v = ws.cell(row=r, column=2).value
            if isinstance(v, str) and v.strip() in titles:
                is_scenario = True
                break
        if is_scenario:
            scenario_names.append(name)

    if log_func:
        log_func(f"Detected {len(scenario_names)} scenario sheets: {scenario_names}")
    return scenario_names


def build_straight_comparison_df(
    workbook_path: str,
    sheet_names: Sequence[str],
    threshold: float,
    *,
    log_func=None,
) -> Dict[str, pd.DataFrame]:
    """
    Returns a dict:
      pretty_case_type -> DataFrame with columns:
        Contingency, ResultingIssue, <sheet1>, <sheet2>, ...

    Rows are sorted by row max percent desc.
    """
    out: Dict[str, pd.DataFrame] = {}

    # Pre-load all sheets once
    loaded: Dict[str, pd.DataFrame] = {}
    for s in sheet_names:
        df = _load_sheet_as_df(workbook_path, s, log_func=log_func)
        loaded[s] = df

    for case_type in CASE_TYPES_CANONICAL:
        pretty = CANONICAL_TO_PRETTY.get(case_type, case_type)

        per_sheet: List[pd.DataFrame] = []
        for s in sheet_names:
            df = loaded[s]
            sub = df[df["CaseType"] == case_type].copy()
            if sub.empty:
                # Keep shape via empty df; we'll merge outer anyway
                sub = pd.DataFrame(columns=["CTGLabel", "LimViolID", "LimViolPct"])
            sub = sub.rename(columns={"LimViolPct": s})
            per_sheet.append(sub[["CTGLabel", "LimViolID", s]])

        # Outer merge across all sheets on (CTGLabel, LimViolID)
        merged = None
        for d in per_sheet:
            merged = d if merged is None else pd.merge(
                merged, d, on=["CTGLabel", "LimViolID"], how="outer"
            )

        if merged is None or merged.empty:
            out[pretty] = pd.DataFrame(
                columns=["Contingency", "ResultingIssue"] + list(sheet_names)
            )
            continue

        merged = merged.rename(columns={"CTGLabel": "Contingency", "LimViolID": "ResultingIssue"})

        # Threshold: keep rows where max across all sheet columns >= threshold
        pct_cols = list(sheet_names)
        merged["_RowMax"] = merged[pct_cols].max(axis=1, skipna=True)

        if threshold and threshold > 0:
            merged = merged[merged["_RowMax"].fillna(float("-inf")) >= float(threshold)]

        merged = merged.sort_values(by="_RowMax", ascending=False, na_position="last").drop(columns=["_RowMax"])

        out[pretty] = merged

    return out


def _write_data_row(
    ws: Worksheet,
    row: int,
    cont: str,
    issue: str,
    pct_values: Sequence[Optional[float]],
    *,
    outline_level: int = 0,
    hidden: bool = False,
    bold: bool = False,
):
    values = [cont, issue] + list(pct_values)

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

        if col_offset >= 2 and isinstance(val, (float, int)):
            cell.number_format = "0.00"

    try:
        ws.row_dimensions[row].outlineLevel = int(outline_level)
        ws.row_dimensions[row].hidden = bool(hidden)
    except Exception:
        pass


def _normalize_issue_series(series: pd.Series) -> pd.Series:
    """Forward-fill blanks (safety net)."""
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


def write_straight_comparison_sheet(
    wb: Workbook,
    ws_name: str,
    workbook_path: str,
    *,
    threshold: float = 0.0,
    sheet_names: Optional[Sequence[str]] = None,
    log_func=None,
    expandable_issue_view: bool = True,
):
    """
    Adds a final sheet comparing ALL scenario sheets side-by-side.

    If sheet_names is None, we auto-detect scenario sheets from workbook_path.
    """
    if sheet_names is None:
        sheet_names = discover_scenario_sheets(workbook_path, log_func=log_func)

    case_labels = list(sheet_names)
    num_cases = len(case_labels)
    last_col = 3 + num_cases  # B..(2+1+num_cases)

    ws = wb.create_sheet(title=ws_name)
    apply_straight_table_styles(ws, num_cases=num_cases)

    if num_cases == 0:
        ws.cell(row=2, column=2).value = "No scenario sheets detected."
        return

    case_type_to_df = build_straight_comparison_df(
        workbook_path, sheet_names=sheet_names, threshold=threshold, log_func=log_func
    )

    current_row = 2

    for pretty_case_type in ["ACCA LongTerm", "ACCA", "DCwAC"]:
        df = case_type_to_df.get(pretty_case_type)
        if df is None or df.empty:
            continue

        # Safety-net normalize blanks
        df = df.copy()
        df["ResultingIssue"] = _normalize_issue_series(df["ResultingIssue"])

        _write_title_row(ws, current_row, pretty_case_type, last_col=last_col)
        current_row += 1
        _write_header_row(ws, current_row, case_labels=case_labels)
        current_row += 1

        if not expandable_issue_view:
            for _, r in df.iterrows():
                cont = str(r.get("Contingency", "") or "")
                issue = str(r.get("ResultingIssue", "") or "")
                pcts = [r.get(s, None) for s in sheet_names]
                _write_data_row(ws, current_row, cont, issue, pcts, outline_level=0, hidden=False, bold=False)
                current_row += 1
            current_row += 1
            continue

        # Expandable: group by issue, sort each group by row max
        df["_SortKey"] = df.apply(lambda r: _row_max([r.get(s, None) for s in sheet_names]), axis=1)

        group_max = df.groupby("ResultingIssue")["_SortKey"].max().sort_values(ascending=False)
        ordered_issues = list(group_max.index)

        for issue_key in ordered_issues:
            g = df[df["ResultingIssue"] == issue_key].copy()
            if g.empty:
                continue

            g = g.sort_values(by="_SortKey", ascending=False, na_position="last")

            summary_row_index = None
            first = True

            for _, r in g.iterrows():
                cont = str(r.get("Contingency", "") or "")
                issue = str(r.get("ResultingIssue", "") or "")
                issue_display = issue if first else ""

                pcts = [r.get(s, None) for s in sheet_names]

                if first:
                    summary_row_index = current_row

                _write_data_row(
                    ws,
                    current_row,
                    cont,
                    issue_display,
                    pcts,
                    outline_level=0 if first else 1,
                    hidden=False if first else True,
                    bold=True if first else False,
                )
                current_row += 1
                first = False

            try:
                if summary_row_index is not None and len(g) > 1:
                    ws.row_dimensions[summary_row_index].collapsed = True
            except Exception:
                pass

        current_row += 1
