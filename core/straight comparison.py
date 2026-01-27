# core/straight_comparison.py
#
# Builds and writes a "Straight Comparison" sheet that compares N original
# scenario sheets side-by-side (no pairwise deltas).
#
# Output is styled to match the blue-block look of the batch comparison sheets:
# - CaseType blocks (ACCA LongTerm / ACCA / DCwAC)
# - Columns: Contingency Events | Resulting Issue | <one % column per scenario>
# - Optional Excel +/- outline grouped by Resulting Issue (summary row ABOVE details)
# - The top (max) row per Resulting Issue is bolded (like the batch sheets)

from __future__ import annotations

from typing import Dict, List, Sequence, Tuple
import math
import os

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# CaseType mappings (same as comparator.py)
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Parsing the formatted scenario sheets
# ---------------------------------------------------------------------------

def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def _parse_scenario_sheet(ws, log_func=None) -> pd.DataFrame:
    """
    Parse one formatted scenario sheet into rows with:
        CaseType, CTGLabel, LimViolID, LimViolPct

    NOTE:
      Resulting Issue cells may be blank for visual grouping. Those blanks mean
      "same as above" within the current CaseType block; we forward-fill LimViolID.
    """
    records: List[Dict] = []
    max_row = ws.max_row or 1
    row_idx = 1

    while row_idx <= max_row:
        title_val = ws.cell(row=row_idx, column=2).value  # column B
        if isinstance(title_val, str) and title_val.strip():
            pretty = title_val.strip()
            case_type = CANONICAL_CASE_TYPES.get(pretty, pretty)

            header_row = row_idx + 1
            data_row = header_row + 1

            last_issue = None
            r = data_row

            while r <= max_row:
                b = ws.cell(row=r, column=2).value  # CTGLabel
                c = ws.cell(row=r, column=3).value  # LimViolID
                e = ws.cell(row=r, column=5).value  # LimViolPct

                d = ws.cell(row=r, column=4).value
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
                        "LimViolPct": e,
                    }
                )
                r += 1

            row_idx = r + 1
        else:
            row_idx += 1

    df = pd.DataFrame.from_records(records)
    if log_func:
        log_func(f"Parsed {len(df)} rows from sheet '{ws.title}' for straight comparison.")
    return df


def load_sheet_as_df(workbook_path: str, sheet_name: str, log_func=None) -> pd.DataFrame:
    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(workbook_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
    ws = wb[sheet_name]
    return _parse_scenario_sheet(ws, log_func=log_func)


# ---------------------------------------------------------------------------
# Build a straight comparison dataframe
# ---------------------------------------------------------------------------

def _safe_float(x):
    try:
        if x is None:
            return None
        if isinstance(x, float) and math.isnan(x):
            return None
        return float(x)
    except Exception:
        return None


def build_straight_comparison_df(
    workbook_path: str,
    sheet_names: Sequence[str],
    threshold: float,
    log_func=None,
) -> Tuple[pd.DataFrame, List[str]]:
    """
    Returns (df, ordered_case_labels).

    df columns:
      CaseType, Contingency, ResultingIssue, <case_label_1>, <case_label_2>, ...

    Threshold:
      keep rows where max across cases >= threshold.
    """
    if not sheet_names:
        return pd.DataFrame(columns=["CaseType", "Contingency", "ResultingIssue"]), []

    # Create short-but-unique labels for column headers
    used: set[str] = set()
    labels: List[str] = []

    def make_label(name: str) -> str:
        base = str(name).strip()
        lab = base if len(base) <= 14 else base[:14]
        candidate = lab
        k = 2
        while candidate in used or candidate == "":
            suffix = f"_{k}"
            candidate = (lab[: max(1, 14 - len(suffix))] + suffix)
            k += 1
        used.add(candidate)
        return candidate

    for s in sheet_names:
        labels.append(make_label(s))

    master = None

    for sheet_name, col_label in zip(sheet_names, labels):
        df = load_sheet_as_df(workbook_path, sheet_name, log_func=log_func)
        if df.empty:
            continue

        df["CaseTypePretty"] = df["CaseType"].map(CANONICAL_TO_PRETTY).fillna(df["CaseType"])
        df = df.rename(columns={"CTGLabel": "Contingency", "LimViolID": "ResultingIssue"})
        df[col_label] = df["LimViolPct"].apply(_safe_float)

        # Collapse duplicates by max within this scenario
        df = (
            df.groupby(["CaseTypePretty", "Contingency", "ResultingIssue"], as_index=False)[col_label]
            .max()
        )

        if master is None:
            master = df
        else:
            master = pd.merge(
                master,
                df,
                on=["CaseTypePretty", "Contingency", "ResultingIssue"],
                how="outer",
            )

    if master is None or master.empty:
        out = pd.DataFrame(columns=["CaseType", "Contingency", "ResultingIssue"] + labels)
        return out, labels

    master = master.rename(columns={"CaseTypePretty": "CaseType"})
    case_cols = [c for c in labels if c in master.columns]

    if case_cols:
        max_series = master[case_cols].max(axis=1, skipna=True)
        master = master[max_series.fillna(float("-inf")) >= float(threshold)].copy()

        master["_SortKey"] = max_series
        master = master.sort_values(
            by=["CaseType", "_SortKey"], ascending=[True, False], na_position="last"
        ).drop(columns=["_SortKey"])

    for c in labels:
        if c not in master.columns:
            master[c] = None

    master = master[["CaseType", "Contingency", "ResultingIssue"] + labels].copy()
    return master, labels


# ---------------------------------------------------------------------------
# Worksheet writer (blue-block style + optional +/-)
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="305496")
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


def _apply_table_styles(ws: Worksheet, num_cases: int):
    ws.column_dimensions[get_column_letter(2)].width = 45  # B
    ws.column_dimensions[get_column_letter(3)].width = 45  # C
    for i in range(num_cases):
        ws.column_dimensions[get_column_letter(4 + i)].width = 12

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


def _write_row(
    ws: Worksheet,
    row: int,
    values: Sequence,
    *,
    outline_level: int = 0,
    hidden: bool = False,
    bold: bool = False,
):
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


def _max_across_cases(row: pd.Series, case_cols: Sequence[str]) -> float:
    vals = []
    for c in case_cols:
        v = row.get(c, None)
        if v is None:
            continue
        if isinstance(v, float) and math.isnan(v):
            continue
        try:
            vals.append(float(v))
        except Exception:
            continue
    return max(vals) if vals else float("-inf")


def write_formatted_straight_sheet(
    wb: Workbook,
    ws_name: str,
    df: pd.DataFrame,
    case_labels: Sequence[str],
    *,
    expandable_issue_view: bool = True,
):
    ws = wb.create_sheet(title=ws_name)
    _apply_table_styles(ws, num_cases=len(case_labels))

    if df is None or df.empty:
        ws.cell(row=2, column=2).value = "No rows above threshold."
        return

    current_row = 2
    case_cols = list(case_labels)
    last_col = 2 + 1 + len(case_cols)

    for case_type_pretty in ["ACCA LongTerm", "ACCA", "DCwAC"]:
        sub = df[df["CaseType"] == case_type_pretty].copy()
        if sub.empty:
            continue

        _write_title_row(ws, current_row, case_type_pretty, last_col=last_col)
        current_row += 1
        _write_header_row(ws, current_row, case_cols)
        current_row += 1

        if not expandable_issue_view:
            for _, r in sub.iterrows():
                cont = str(r.get("Contingency", "") or "")
                issue = str(r.get("ResultingIssue", "") or "")
                vals = [cont, issue] + [r.get(c, None) for c in case_cols]
                _write_row(ws, current_row, vals)
                current_row += 1
            current_row += 1
            continue

        sub["_SortKey"] = sub.apply(lambda rr: _max_across_cases(rr, case_cols), axis=1)
        group_max = sub.groupby("ResultingIssue")["_SortKey"].max().sort_values(ascending=False)
        ordered_issues = list(group_max.index)

        for issue_key in ordered_issues:
            g = sub[sub["ResultingIssue"] == issue_key].copy()
            if g.empty:
                continue

            g = g.sort_values(by="_SortKey", ascending=False, na_position="last")

            summary_row_index = None
            first = True

            for _, r in g.iterrows():
                cont = str(r.get("Contingency", "") or "")
                issue = str(r.get("ResultingIssue", "") or "")
                issue_display = issue if first else ""
                vals = [cont, issue_display] + [r.get(c, None) for c in case_cols]

                if first:
                    summary_row_index = current_row

                _write_row(
                    ws,
                    current_row,
                    vals,
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