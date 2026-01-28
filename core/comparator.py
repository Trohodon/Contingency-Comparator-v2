# core/comparator.py
#
# Helpers for working with the formatted
# Combined_ViolationCTG_Comparison.xlsx workbook and for building
# batch comparison workbooks in a nicely formatted style.
#
# Public functions used by the GUI:
#   - list_sheets(workbook_path)
#   - build_case_type_comparison(... )
#   - build_pair_comparison_df(... )
#   - build_batch_comparison_workbook(... )
#
# UPDATED:
#   - build_batch_comparison_workbook now allows pairs=[] so users can build a
#     workbook containing ONLY the "Straight Comparison" sheet (all originals).

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
from core.straight_comparison import build_straight_comparison_df, write_formatted_straight_sheet


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
    records: List[Dict] = []

    max_row = ws.max_row or 1
    row_idx = 1

    while row_idx <= max_row:
        title_cell = ws.cell(row=row_idx, column=2)
        title_val = title_cell.value

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

    df = pd.DataFrame.from_records(records)
    if log_func:
        log_func(
            f"Parsed {len(df)} rows from sheet '{ws.title}'. "
            f"Columns: {list(df.columns)}"
        )
    return df


def _load_sheet_as_df(workbook_path: str, sheet_name: str, log_func=None) -> pd.DataFrame:
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required for comparison.")
    wb = load_workbook(workbook_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in workbook.")
    ws = wb[sheet_name]
    return _parse_scenario_sheet(ws, log_func=log_func)


def build_case_type_comparison(
    workbook_path: str,
    base_sheet: str,
    new_sheet: str,
    case_type: str,
    max_rows: Optional[int] = None,
    log_func=None,
) -> pd.DataFrame:
    if case_type not in CASE_TYPES_CANONICAL:
        raise ValueError(f"Unknown case type: {case_type}")

    base_df = _load_sheet_as_df(workbook_path, base_sheet, log_func=log_func)
    new_df = _load_sheet_as_df(workbook_path, new_sheet, log_func=log_func)

    base_df = base_df[base_df["CaseType"] == case_type].copy()
    new_df = new_df[new_df["CaseType"] == case_type].copy()

    if log_func:
        log_func(f"  [{case_type}] base rows={len(base_df)}, new rows={len(new_df)}")

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


def _is_nan(x) -> bool:
    return isinstance(x, float) and math.isnan(x)


def build_pair_comparison_df(
    workbook_path: str,
    left_sheet: str,
    right_sheet: str,
    threshold: float,
    log_func=None,
) -> pd.DataFrame:
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
            by=["CaseType", "_SortKey"],
            ascending=[True, False],
            na_position="last",
        ).drop(columns=["_SortKey"])
    return df_all


def _sanitize_sheet_name(name: str) -> str:
    invalid = set(r'[]:*?/\\')
    cleaned = "".join(ch if ch not in invalid else "_" for ch in name).strip()
    return (cleaned or "Sheet")[:31]


def _looks_like_output_sheet(name: str) -> bool:
    """
    Filter out sheets that are clearly generated outputs.
    We want originals from the source workbook for Straight Comparison.
    """
    n = (name or "").strip().lower()
    if not n:
        return True
    if " vs " in n:
        return True
    if n.startswith("straight comparison"):
        return True
    if n.startswith("batch comparison"):
        return True
    if n.startswith("comparison"):
        return True
    return False


def _ordered_original_sheets(src_workbook: str, pairs: Sequence[Tuple[str, str]]) -> List[str]:
    """
    If pairs are provided: return sheets involved in pairs in workbook order.
    If pairs is empty: return ALL likely-original scenario sheets in workbook order.
    """
    try:
        wb = load_workbook(src_workbook, read_only=True, data_only=True)
        sheetnames = list(wb.sheetnames)

        # If no pairs, include all "original-looking" sheets
        if not pairs:
            originals = [s for s in sheetnames if not _looks_like_output_sheet(s)]
            return originals

        wanted: set[str] = set()
        for a, b in pairs:
            wanted.add(a)
            wanted.add(b)

        ordered = [s for s in sheetnames if s in wanted]
        for s in wanted:
            if s not in ordered:
                ordered.append(s)
        return ordered

    except Exception:
        # Fallback behavior if workbook can't be opened
        if not pairs:
            return []

        ordered: List[str] = []
        seen: set[str] = set()
        for a, b in pairs:
            for s in (a, b):
                if s not in seen:
                    seen.add(s)
                    ordered.append(s)
        return ordered


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

    # IMPORTANT UPDATE:
    # pairs can be empty -> build a workbook with only Straight Comparison.

    wb = Workbook()
    wb.remove(wb.active)

    used_names: set[str] = set()

    # Pair sheets (only if pairs exist)
    if pairs:
        for (left_sheet, right_sheet) in pairs:
            df_pair = build_pair_comparison_df(
                src_workbook, left_sheet, right_sheet, threshold, log_func=log_func
            )
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

            write_formatted_pair_sheet(wb, name, df_pair, expandable_issue_view=expandable_issue_view)
    else:
        if log_func:
            log_func("No queued pairs provided; building Straight Comparison only.")

    # Straight Comparison (always attempt)
    try:
        originals = _ordered_original_sheets(src_workbook, pairs)
        if not originals:
            if log_func:
                log_func("WARNING: No original sheets detected for Straight Comparison.")
        else:
            df_straight, case_labels = build_straight_comparison_df(
                src_workbook,
                originals,
                threshold=threshold,
                log_func=log_func
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
    except Exception as e:
        if log_func:
            log_func(f"WARNING: Straight Comparison sheet failed: {e}")

    wb.save(output_path)
    return output_path
