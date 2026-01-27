# core/comparator.py
#
# Batch comparison builder + GUI helpers.
#
# Speed fix:
# - Avoid repeated workbook loads + repeated parsing (major freeze cause).
# - Parse each sheet ONCE per build and reuse those DataFrames.

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


def list_sheets(workbook_path: str) -> List[str]:
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
    FAST iter_rows parser.
    Outputs:
      CaseType, CTGLabel, LimViolID, LimViolValue, LimViolPct
    """
    records: List[Dict] = []

    current_case_type = None
    skip_rows = 0
    last_issue = None

    # columns B..E
    for (b, c, d, e) in ws.iter_rows(min_row=1, min_col=2, max_col=5, values_only=True):
        if current_case_type is None:
            if isinstance(b, str) and b.strip():
                pretty = b.strip()
                current_case_type = CANONICAL_CASE_TYPES.get(pretty, pretty)
                skip_rows = 1  # skip header row
                last_issue = None
            continue

        if skip_rows > 0:
            skip_rows -= 1
            continue

        if _is_blank(b) and _is_blank(c) and _is_blank(d) and _is_blank(e):
            current_case_type = None
            skip_rows = 0
            last_issue = None
            continue

        if _is_blank(c) and last_issue is not None:
            c = last_issue
        else:
            if not _is_blank(c):
                last_issue = c

        records.append(
            {
                "CaseType": current_case_type,
                "CTGLabel": b,
                "LimViolID": c,
                "LimViolValue": d,
                "LimViolPct": e,
            }
        )

    df = pd.DataFrame.from_records(records)
    if log_func:
        log_func(f"Parsed {len(df)} rows from sheet '{ws.title}'.")
    return df


def _is_nan(x) -> bool:
    return isinstance(x, float) and math.isnan(x)


def _case_type_comparison_from_dfs(
    base_df_full: pd.DataFrame,
    new_df_full: pd.DataFrame,
    case_type: str,
    max_rows: Optional[int] = None,
) -> pd.DataFrame:
    """
    Same output as build_case_type_comparison, but uses pre-parsed sheet dataframes.
    """
    base_df = base_df_full[base_df_full["CaseType"] == case_type].copy()
    new_df = new_df_full[new_df_full["CaseType"] == case_type].copy()

    if base_df.empty and new_df.empty:
        return pd.DataFrame(columns=["Contingency", "ResultingIssue", "LeftPct", "RightPct", "DeltaPct"])

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
    result = result.sort_values(by="_SortPct", ascending=False, na_position="last").drop(columns=["_SortPct"])

    if max_rows is not None and max_rows > 0:
        result = result.head(max_rows)

    return result


def build_case_type_comparison(
    workbook_path: str,
    base_sheet: str,
    new_sheet: str,
    case_type: str,
    max_rows: Optional[int] = None,
    log_func=None,
) -> pd.DataFrame:
    """
    GUI helper (single case type). Still works the same, but loads workbook once internally.
    """
    if case_type not in CASE_TYPES_CANONICAL:
        raise ValueError(f"Unknown case type: {case_type}")

    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for comparison.")
    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(workbook_path, read_only=True, data_only=True)

    if base_sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{base_sheet}' not found in workbook.")
    if new_sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{new_sheet}' not found in workbook.")

    base_df_full = _parse_scenario_sheet(wb[base_sheet], log_func=log_func)
    new_df_full = _parse_scenario_sheet(wb[new_sheet], log_func=log_func)

    return _case_type_comparison_from_dfs(base_df_full, new_df_full, case_type, max_rows=max_rows)


def build_pair_comparison_df(
    workbook_path: str,
    left_sheet: str,
    right_sheet: str,
    threshold: float,
    log_func=None,
) -> pd.DataFrame:
    """
    FAST batch helper:
    - parse left sheet once
    - parse right sheet once
    - do all case types from cached dfs
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for comparison.")
    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(workbook_path, read_only=True, data_only=True)

    if left_sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{left_sheet}' not found in workbook.")
    if right_sheet not in wb.sheetnames:
        raise ValueError(f"Sheet '{right_sheet}' not found in workbook.")

    left_df_full = _parse_scenario_sheet(wb[left_sheet], log_func=log_func)
    right_df_full = _parse_scenario_sheet(wb[right_sheet], log_func=log_func)

    records: List[Dict] = []

    for case_type in CASE_TYPES_CANONICAL:
        pretty = CANONICAL_TO_PRETTY.get(case_type, case_type)

        df = _case_type_comparison_from_dfs(
            left_df_full,
            right_df_full,
            case_type,
            max_rows=None,
        )

        if df.empty:
            continue

        for _, row in df.iterrows():
            cont = str(row.get("Contingency", "") or "")
            issue = "" if row.get("ResultingIssue", None) is None else str(row.get("ResultingIssue"))

            left_pct = row.get("LeftPct", math.nan)
            right_pct = row.get("RightPct", math.nan)
            delta_pct = row.get("DeltaPct", math.nan)

            values = []
            if not _is_nan(left_pct):
                values.append(float(left_pct))
            if not _is_nan(right_pct):
                values.append(float(right_pct))

            if not values or max(values) < threshold:
                continue

            if _is_nan(left_pct) and not _is_nan(right_pct):
                delta_text = "Only in right"
            elif not _is_nan(left_pct) and _is_nan(right_pct):
                delta_text = "Only in left"
            elif _is_nan(left_pct) and _is_nan(right_pct):
                delta_text = ""
            else:
                delta_text = f"{float(delta_pct):.2f}" if not _is_nan(delta_pct) else ""

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
    cleaned = "".join(ch if ch not in invalid else "_" for ch in name).strip()
    return (cleaned or "Sheet")[:31]


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
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to build the batch workbook.")

    if src_workbook is None:
        src_workbook = workbook_path
    if src_workbook is None:
        src_workbook = kwargs.get("src_workbook") or kwargs.get("workbook") or kwargs.get("path")
    if not src_workbook:
        raise ValueError("Missing source workbook path (src_workbook / workbook_path).")
    if not pairs:
        raise ValueError("No comparison pairs provided.")

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    used_names: set[str] = set()

    # Pair sheets
    for (left_sheet, right_sheet) in pairs:
        if log_func:
            log_func(f"Processing pair: {left_sheet} vs {right_sheet}")

        df_pair = build_pair_comparison_df(src_workbook, left_sheet, right_sheet, threshold, log_func=log_func)

        if df_pair.empty:
            df_pair = pd.DataFrame([{
                "CaseType": "",
                "Contingency": "No rows above threshold.",
                "ResultingIssue": "",
                "LeftPct": None,
                "RightPct": None,
                "DeltaDisplay": "",
            }])

        base_name = _sanitize_sheet_name(f"{left_sheet} vs {right_sheet}")
        name = base_name
        counter = 2
        while name in used_names:
            suffix = f" ({counter})"
            name = _sanitize_sheet_name(base_name[: (31 - len(suffix))] + suffix)
            counter += 1
        used_names.add(name)

        write_formatted_pair_sheet(wb_out, name, df_pair, expandable_issue_view=expandable_issue_view)

    # Straight Comparison (ALL originals from SOURCE workbook)
    try:
        originals = discover_scenario_sheets(src_workbook, log_func=log_func)

        if not originals:
            # fallback (should rarely happen)
            originals = list_sheets(src_workbook)

        df_straight, case_labels = build_straight_comparison_df(
            src_workbook,
            originals,
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
            wb_out,
            sc_name,
            df_straight,
            case_labels,
            expandable_issue_view=expandable_issue_view,
        )

        if log_func:
            log_func(f"Added Straight Comparison for {len(originals)} original source sheets.")
    except Exception as e:
        if log_func:
            log_func(f"WARNING: Straight Comparison failed: {e}")

    wb_out.save(output_path)
    return output_path