"""core/straight comparison.py

Builds a *single* sheet workbook where each row is a unique:
  (Contingency Events, Resulting Issue)
and each column is a selected scenario (e.g., Base Case, Breaker 1, Breaker 2...).

UPDATE (Limit column support):
  Adds a Limit column (associated with the Resulting Issue) as column D.

The source scenario sheets may be in either layout:
  Old:  B=Contingency, C=Issue, D=Value, E=Percent
  New:  B=Contingency, C=Issue, D=Limit, E=Value, F=Percent

This module detects header columns dynamically.
"""

from __future__ import annotations

import math
import os
from typing import Callable, Dict, Optional, Sequence

import pandas as pd

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False


# ----------------------------- helpers ----------------------------- #

def _norm_header(x) -> str:
    if x is None:
        return ""
    return str(x).strip().lower().replace("\n", " ")


def _detect_columns(ws, header_row: int) -> Dict[str, int]:
    cols = {}
    for col_idx in range(2, 9):  # B..H
        v = ws.cell(row=header_row, column=col_idx).value
        h = _norm_header(v)
        if not h:
            continue

        if ("contingency" in h and "value" not in h) or "ctglabel" in h:
            cols["ctg"] = col_idx
            continue

        if "resulting issue" in h or "limviolid" in h or (("issue" in h) and ("resulting" in h)):
            cols["issue"] = col_idx
            continue

        if "limit" in h and "value" not in h:
            cols["limit"] = col_idx
            continue

        if "mva" in h or ("contingency value" in h) or (h == "value") or ("limviolvalue" in h):
            cols["value"] = col_idx
            continue

        if "percent" in h or "loading" in h or "pct" in h or "limviolpct" in h:
            cols["pct"] = col_idx
            continue

    return cols


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


def _is_nan(x) -> bool:
    return isinstance(x, float) and math.isnan(x)


def _parse_scenario_sheet(workbook_path: str, sheet_name: str, log_func: Optional[Callable[[str], None]] = None) -> pd.DataFrame:
    """Parse one scenario sheet into normalized rows.

    Output columns:
      CaseType, CTGLabel, LimViolID, LimViolLimit, LimViolPct
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to read Excel workbooks in this build.")

    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        ws = wb[sheet_name]
        rows = []
        r = 1
        max_r = ws.max_row or 1

        while r <= max_r:
            title_cell = ws.cell(row=r, column=2).value  # B
            title = str(title_cell).strip() if title_cell is not None else ""
            if title in ("ACCA LongTerm", "ACCA", "DCwAC"):
                header_row = r + 1
                colmap = _detect_columns(ws, header_row)

                ctg_col = colmap.get("ctg", 2)
                issue_col = colmap.get("issue", 3)
                limit_col = colmap.get("limit")
                value_col = colmap.get("value")
                pct_col = colmap.get("pct")

                # Old layout fallback
                if value_col is None and pct_col is None:
                    value_col = 4
                    pct_col = 5

                # New layout fallback
                if limit_col is None and value_col == 5 and pct_col is None:
                    pct_col = 6

                case_type = {
                    "ACCA LongTerm": "ACCA_LongTerm",
                    "ACCA": "ACCA_P1,2,4,7",
                    "DCwAC": "DCwACver_P1-7",
                }.get(title, title)

                r = header_row + 1
                while r <= max_r:
                    ctg = ws.cell(row=r, column=ctg_col).value
                    issue = ws.cell(row=r, column=issue_col).value

                    if (ctg is None or str(ctg).strip() == "") and (issue is None or str(issue).strip() == ""):
                        break

                    lim = ws.cell(row=r, column=limit_col).value if limit_col else None
                    pct = ws.cell(row=r, column=pct_col).value if pct_col else None

                    rows.append(
                        {
                            "CaseType": case_type,
                            "CTGLabel": "" if ctg is None else str(ctg).strip(),
                            "LimViolID": "" if issue is None else str(issue).strip(),
                            "LimViolLimit": "" if lim is None else lim,
                            "LimViolPct": "" if pct is None else pct,
                        }
                    )
                    r += 1

                r += 1
                continue

            r += 1

        df = pd.DataFrame(rows)
        if df.empty:
            return df

        df["LimViolID"] = df["LimViolID"].replace("", pd.NA)
        df["LimViolID"] = df.groupby("CaseType")["LimViolID"].ffill()
        df["LimViolID"] = df["LimViolID"].fillna("")

        return df

    finally:
        wb.close()


def build_straight_comparison_df(
    src_workbook: str,
    left_sheet: str,
    right_sheet: str,
    case_types: Optional[Sequence[str]] = None,
    log_func: Optional[Callable[[str], None]] = None,
) -> pd.DataFrame:
    """Build the straight comparison table as a DataFrame.

    Columns:
      CaseType, Contingency, ResultingIssue, Limit, <left_sheet>, <right_sheet>
    """
    if case_types is None:
        case_types = ["ACCA_LongTerm", "ACCA_P1,2,4,7", "DCwACver_P1-7"]

    left_df = _parse_scenario_sheet(src_workbook, left_sheet, log_func)
    right_df = _parse_scenario_sheet(src_workbook, right_sheet, log_func)

    out_blocks = []

    for ct in case_types:
        l = left_df[left_df["CaseType"] == ct].copy()
        r = right_df[right_df["CaseType"] == ct].copy()

        # numeric percent for aggregations
        l["_pct_num"] = l["LimViolPct"].apply(_to_float)
        r["_pct_num"] = r["LimViolPct"].apply(_to_float)

        key_cols = ["CTGLabel", "LimViolID"]

        # aggregate duplicates: take max pct, and first non-empty limit
        def agg_limit(s):
            s = [x for x in s if x not in (None, "", "nan")]
            return s[0] if s else ""

        l_agg = (
            l.groupby(key_cols, dropna=False)
            .agg({"_pct_num": "max", "LimViolLimit": agg_limit})
            .reset_index()
        )
        r_agg = (
            r.groupby(key_cols, dropna=False)
            .agg({"_pct_num": "max", "LimViolLimit": agg_limit})
            .reset_index()
        )

        l_agg = l_agg.rename(columns={"_pct_num": left_sheet, "LimViolLimit": "Limit_Left"})
        r_agg = r_agg.rename(columns={"_pct_num": right_sheet, "LimViolLimit": "Limit_Right"})

        merged = pd.merge(l_agg, r_agg, on=key_cols, how="outer")
        merged["CaseType"] = ct
        merged["Contingency"] = merged["CTGLabel"].fillna("").astype(str)
        merged["ResultingIssue"] = merged["LimViolID"].fillna("").astype(str)

        # Choose Limit (prefer right, else left; if mismatch and both exist, show both)
        def choose_limit_row(row):
            llim = row.get("Limit_Left", "")
            rlim = row.get("Limit_Right", "")
            ltxt = "" if llim is None else str(llim).strip()
            rtxt = "" if rlim is None else str(rlim).strip()
            if rtxt and ltxt:
                lf = _to_float(ltxt)
                rf = _to_float(rtxt)
                if not _is_nan(lf) and not _is_nan(rf) and abs(lf - rf) < 1e-9:
                    return f"{rf:g}"
                if ltxt == rtxt:
                    return rtxt
                return f"L:{ltxt} | R:{rtxt}"
            return rtxt or ltxt or ""

        merged["Limit"] = merged.apply(choose_limit_row, axis=1)

        out_blocks.append(
            merged[["CaseType", "Contingency", "ResultingIssue", "Limit", left_sheet, right_sheet]]
        )

    if not out_blocks:
        return pd.DataFrame()

    out_df = pd.concat(out_blocks, ignore_index=True)

    # sort by worst percent (max of the two columns)
    out_df["_max"] = out_df.apply(
        lambda r: max([v for v in [r.get(left_sheet, float("nan")), r.get(right_sheet, float("nan"))] if not _is_nan(v)] or [float("nan")]),
        axis=1,
    )
    out_df = out_df.sort_values(by=["CaseType", "_max"], ascending=[True, False], na_position="last")
    out_df = out_df.drop(columns=["_max"])

    return out_df


def write_straight_comparison_workbook(
    src_workbook: str,
    left_sheet: str,
    right_sheet: str,
    output_path: str,
    log_func: Optional[Callable[[str], None]] = None,
) -> str:
    """Create the straight comparison workbook (formatted)."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to write Excel workbooks in this build.")

    df = build_straight_comparison_df(src_workbook, left_sheet, right_sheet, log_func=log_func)

    wb = Workbook()
    ws = wb.active
    ws.title = "Straight Comparison"

    # Styles
    title_fill = PatternFill(fill_type="solid", fgColor="305496")
    title_font = Font(color="FFFFFF", bold=True, size=12)

    header_fill = PatternFill(fill_type="solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True)

    data_font = Font(color="000000")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Column widths (table starts at B)
    ws.column_dimensions["B"].width = 55  # Contingency
    ws.column_dimensions["C"].width = 55  # Issue
    ws.column_dimensions["D"].width = 18  # Limit
    ws.column_dimensions["E"].width = 12  # Left
    ws.column_dimensions["F"].width = 12  # Right

    # Title row
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=6)
    c = ws.cell(row=1, column=2)
    c.value = f"{left_sheet} vs {right_sheet}"
    c.fill = title_fill
    c.font = title_font
    c.alignment = center
    for col in range(2, 7):
        ws.cell(row=1, column=col).border = thin_border

    # Headers
    headers = [
        ("B", "Contingency Events"),
        ("C", "Resulting Issue"),
        ("D", "Limit"),
        ("E", left_sheet),
        ("F", right_sheet),
    ]
    for col_letter, text in headers:
        col_idx = ord(col_letter) - ord("A") + 1
        hc = ws.cell(row=2, column=col_idx)
        hc.value = text
        hc.fill = header_fill
        hc.font = header_font
        hc.alignment = center
        hc.border = thin_border

    # Data
    r = 3
    for _, row in df.iterrows():
        ws.cell(row=r, column=2).value = row.get("Contingency", "")
        ws.cell(row=r, column=3).value = row.get("ResultingIssue", "")
        ws.cell(row=r, column=4).value = row.get("Limit", "")
        ws.cell(row=r, column=5).value = row.get(left_sheet, "")
        ws.cell(row=r, column=6).value = row.get(right_sheet, "")

        for col in range(2, 7):
            cell = ws.cell(row=r, column=col)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = left_align if col in (2, 3) else center

        r += 1

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)

    if log_func:
        log_func(f"Saved straight comparison workbook: {output_path}")

    return output_path