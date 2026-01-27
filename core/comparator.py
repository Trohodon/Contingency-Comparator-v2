# core/comparator.py
#
# Helpers for working with the formatted Combined_ViolationCTG_Comparison workbook
# and for building batch comparison workbooks in a nicely formatted style.
#
# Public functions used by the GUI:
#   - list_sheets(workbook_path)
#   - build_case_type_comparison(...)
#   - build_pair_comparison_df(...)
#   - build_batch_comparison_workbook(...)
#
# NEW:
#   - Final "Straight Comparison" sheet is appended to the batch workbook.
#     It compares ALL original scenario sheets found in the SOURCE workbook
#     (not just the selected batch pairs).

from __future__ import annotations

from typing import List, Dict, Optional, Sequence, Tuple
import math
import os
import pandas as pd

try:
    from openpyxl import load_workbook, Workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

from core.batch_sheet_writer import write_formatted_pair_sheet
from core.straight_comparison import (
    build_straight_comparison_df,
    write_formatted_straight_sheet,
    discover_scenario_sheets,
)

# Mapping from the pretty case-type titles in the formatted sheet
# to the canonical internal case-type names.
CANONICAL_CASE_TYPES = {
    "ACCA LongTerm": "ACCA_LongTerm",
    "ACCA Long Term": "ACCA_LongTerm",  # just in case
    "ACCA": "ACCA_P1,2,4,7",
    "DCwAC": "DCwACver_P1-7",
}

# Canonical list (used by the GUI)
CASE_TYPES_CANONICAL: List[str] = [
    "ACCA_LongTerm",
    "ACCA_P1,2,4,7",
    "DCwACver_P1-7",
]

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
    """
    Return a list of sheet names from the given Excel workbook.
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for sheet listing and comparison.")

    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    return list(wb.sheetnames)


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False


def _parse_scenario_sheet(ws, log_func=None) -> pd.DataFrame:
    """
    Parse one formatted scenario sheet into a DataFrame with columns:

        CaseType, CTGLabel, LimViolID, LimViolValue, LimViolPct

    IMPORTANT:
      Resulting Issue cells may be blank on subsequent rows to visually group
      multiple contingencies under the same issue. Those blanks mean
      "same as the issue above" within that CaseType block.
    """
    records: List[Dict] = []

    max_row = ws.max_row or 1
    row_idx = 1

    while row_idx <= max_row:
        title_val = ws.cell(row=row_idx, column=2).value  # column B

        if isinstance(title_val, str) and title_val.strip():
            pretty_name = title_val.strip()
            case_type = CANONICAL_CASE_TYPES.get(pretty_name, pretty_name)

            data_row = (row_idx + 1) + 1
            last_issue = None

            r = data_row
            while r <= max_row:
                b = ws.cell(row=r, column=2).value  # CTGLabel
                c = ws.cell(row=r, column=3).value  # LimViolID
                d = ws.cell(row=r, column=4).value  # LimViolValue
                e = ws.cell(row=r, column=5).value  # LimViolPct

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

    df = pd.DataFrame.from_records(records)
    if log_func:
        log_func(
            f"Parsed {len(df)} rows from sheet '{ws.title}'. "
            f"Columns: {list(df.columns)}"
        )
    return df


def _load_sheet_as_df(workbook_path: str, sheet_name: str, log_func=None) -> pd.DataFrame:
    """
    Load a scenario sheet from the formatted workbook into a normalized DataFrame.
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for comparison.")

    wb = load_workbook(workbook_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
    ws = wb[sheet_name]
    return _parse_scenario_sheet(ws, log_func=log_func)


# ---------------------------------------------------------------------------
# Per–case-type comparison for the split-screen GUI
# ---------------------------------------------------------------------------

def build_case_type_comparison(
    workbook_path: str,
    base_sheet: str,
    new_sheet: str,
    case_type: str,
    max_rows: Optional[int] = None,
    log_func=None,
) -> pd.DataFrame:
    """
    Build a left/right comparison DataFrame for a single case type.

    Returns columns:
        Contingency, ResultingIssue, LeftPct, RightPct, DeltaPct

    Sorted by new/right pct (fallback left pct).
    """
    if case_type not in CASE_TYPES_CANONICAL:
        raise ValueError(f"Unknown case type: {case_type}")

    base_df = _load_sheet_as_df(workbook_path, base_sheet, log_func=log_func)
    new_df = _load_sheet_as_df(workbook_path, new_sheet, log_func=log_func)

    base_df = base_df[base_df["CaseType"] == case_type].copy()
    new_df = new_df[new_df["CaseType"] == case_type].copy()

    if base_df.empty and new_df.empty:
        return pd.DataFrame(
            columns=["Contingency", "ResultingIssue", "LeftPct", "RightPct", "DeltaPct"]
        )

    base_df = base_df.rename(columns={"LimViolPct": "Left_Pct"})
    new_df = new_df.rename(columns={"LimViolPct": "Right_Pct"})

    key_cols = ["CTGLabel", "LimViolID"]
    left_cols = key_cols + ["Left_Pct"]
    right_cols = key_cols + ["Right_Pct"]

    merged = pd.merge(base_df[left_cols], new_df[right_cols], on=key_cols, how="outer")
    merged["Delta_Pct"] = merged["Right_Pct"] - merged["Left_Pct"]

    result = merged.rename(
        columns={
            "CTGLabel": "Contingency",
            "LimViolID": "ResultingIssue",
            "Left_Pct": "LeftPct",
            "Right_Pct": "RightPct",
            "Delta_Pct": "DeltaPct",
        }
    )

    sort_series = result["RightPct"]
    result["_SortPct"] = sort_series if sort_series.notna().any() else result["LeftPct"]
    result = result.sort_values(by="_SortPct", ascending=False, na_position="last").drop(
        columns=["_SortPct"]
    )

    if max_rows is not None and max_rows > 0:
        result = result.head(max_rows)

    return result


# ---------------------------------------------------------------------------
# Batch export helpers
# ---------------------------------------------------------------------------

def _is_nan(x) -> bool:
    return isinstance(x, float) and math.isnan(x)


def build_pair_comparison_df(
    workbook_path: str,
    left_sheet: str,
    right_sheet: str,
    threshold: float,
    log_func=None,
) -> pd.DataFrame:
    """
    Build a flat DataFrame for one (left_sheet, right_sheet) pair, combining all case types.

    Columns:
      CaseType, Contingency, ResultingIssue, LeftPct, RightPct, DeltaDisplay
    """
    records: List[Dict] = []

    for case_type in CASE_TYPES_CANONICAL:
        pretty = CANONICAL_TO_PRETTY.get(case_type, case_type)

        df = build_case_type_comparison(
            workbook_path,
            base_sheet=left_sheet,
            new_sheet=right_sheet,
            case_type=case_type,
            max_rows=None,
            log_func=log_func,
        )

        if df.empty:
            continue

        for _, row in df.iterrows():
            cont = str(row.get("Contingency", "") or "")
            issue = row.get("ResultingIssue", "")
            issue = "" if issue is None else str(issue)

            left_pct = row.get("LeftPct", math.nan)
            right_pct = row.get("RightPct", math.nan)
            delta_pct = row.get("DeltaPct", math.nan)

            values = []
            if not _is_nan(left_pct):
                values.append(float(left_pct))
            if not _is_nan(right_pct):
                values.append(float(right_pct))

            if not values:
                continue
            if max(values) < threshold:
                continue

            if _is_nan(left_pct) and not _is_nan(right_pct):
                delta_text = "Only in right"
            elif not _is_nan(left_pct) and _is_nan(right_pct):
                delta_text = "Only in left"
            elif _is_nan(left_pct) and _is_nan(right_pct):
                delta_text = ""
            else:
                try:
                    delta_text = f"{float(delta_pct):.2f}"
                except Exception:
                    delta_text = str(delta_pct)

            records.append(
                {
                    "CaseType": pretty,
                    "Contingency": cont,
                    "ResultingIssue": issue,
                    "LeftPct": float(left_pct) if not _is_nan(left_pct) else None,
                    "RightPct": float(right_pct) if not _is_nan(right_pct) else None,
                    "DeltaDisplay": delta_text,
                }
            )

    df_all = pd.DataFrame.from_records(records)

    if not df_all.empty:
        sort_vals = df_all[["LeftPct", "RightPct"]].max(axis=1)
        df_all["_SortKey"] = sort_vals
        df_all = df_all.sort_values(
            by=["CaseType", "_SortKey"], ascending=[True, False], na_position="last"
        ).drop(columns=["_SortKey"])

    return df_all


def _sanitize_sheet_name(name: str) -> str:
    invalid = set(r'[]:*?/\\')
    cleaned = "".join(ch if ch not in invalid else "_" for ch in name)
    cleaned = cleaned.strip()
    if not cleaned:
        cleaned = "Sheet"
    return cleaned[:31]


def build_batch_comparison_workbook(
    src_workbook: Optional[str] = None,
    pairs: Sequence[Tuple[str, str]] = (),
    threshold: float = 0.0,
    output_path: str = "",
    log_func=None,
    *,
    expandable_issue_view: bool = True,
    workbook_path: Optional[str] = None,
    **kwargs,
) -> str:
    """
    Build a brand-new .xlsx workbook with:
      - one sheet per (left_sheet, right_sheet) pair (UNCHANGED behavior)
      - FINAL sheet: "Straight Comparison" comparing ALL original scenario sheets
        found in the SOURCE workbook (not just the selected pairs).
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to build the batch workbook.")

    # Backward/forward compatibility for callers
    if src_workbook is None:
        src_workbook = workbook_path
    if src_workbook is None:
        src_workbook = kwargs.get("src_workbook") or kwargs.get("workbook") or kwargs.get("path")

    if not src_workbook:
        raise ValueError("Missing source workbook path (src_workbook / workbook_path).")
    if not pairs:
        raise ValueError("No comparison pairs provided.")

    if log_func:
        log_func(
            f"\n=== Building batch comparison workbook ===\n"
            f"Source: {src_workbook}\n"
            f"Output: {output_path}\n"
            f"Threshold: {threshold:.2f}%\n"
            f"Expandable issue view (+/-): {expandable_issue_view}\n"
        )

    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    used_names: set[str] = set()

    # Pair sheets (UNCHANGED)
    for (left_sheet, right_sheet) in pairs:
        if log_func:
            log_func(f"Processing pair: '{left_sheet}' vs '{right_sheet}'...")

        df_pair = build_pair_comparison_df(
            src_workbook, left_sheet, right_sheet, threshold, log_func=log_func
        )

        if df_pair.empty:
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

        base_name = _sanitize_sheet_name(f"{left_sheet} vs {right_sheet}")
        name = base_name
        counter = 2
        while name in used_names:
            suffix = f" ({counter})"
            name = _sanitize_sheet_name(base_name[: (31 - len(suffix))] + suffix)
            counter += 1
        used_names.add(name)

        write_formatted_pair_sheet(
            wb,
            name,
            df_pair,
            expandable_issue_view=expandable_issue_view,
        )

    # FINAL: Straight Comparison (ALL source scenario sheets)
    try:
        all_originals = discover_scenario_sheets(src_workbook, log_func=log_func)

        if not all_originals:
            # fallback: at least something predictable
            all_originals = list_sheets(src_workbook)

        df_straight, case_labels = build_straight_comparison_df(
            src_workbook,
            all_originals,
            threshold=threshold,
            log_func=log_func,
        )

        sc_base = _sanitize_sheet_name("Straight Comparison")
        sc_name = sc_base
        k = 2
        while sc_name in used_names:
            suffix = f" ({k})"
            sc_name = _sanitize_sheet_name(sc_base[: (31 - len(suffix))] + suffix)
            k += 1
        used_names.add(sc_name)

        write_formatted_straight_sheet(
            wb,
            sc_name,
            df_straight,
            case_labels,
            expandable_issue_view=expandable_issue_view,
        )

        if log_func:
            log_func(f"Added final sheet: '{sc_name}' (Straight Comparison of {len(all_originals)} originals)")
    except Exception as e:
        if log_func:
            log_func(f"WARNING: Straight Comparison sheet failed: {e}")

    wb.save(output_path)

    if log_func:
        log_func(f"Batch comparison workbook written to:\n{output_path}")

    return output_path