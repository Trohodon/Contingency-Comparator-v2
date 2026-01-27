# core/comparator.py

from __future__ import annotations

import os
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# We prefer using your formatting writer for batch sheets (outline +/- etc.)
try:
    from batch_sheet_writer import write_formatted_pair_sheet, apply_table_styles
except Exception:
    write_formatted_pair_sheet = None
    apply_table_styles = None


# Canonical case-type ordering (used for sorting / grouping)
CASE_TYPES_CANONICAL = [
    "Branch MVA",
    "Interface MW",
    "Interface MVA",
    "Voltage (PU)",
    "Bus Voltage (PU)",
    "Transformer MVA",
    "DC Line",
    "Other",
]


# -------------------------
# Public API used by GUI
# -------------------------

def list_sheets(workbook_path: str) -> List[str]:
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def build_case_type_comparison(
    workbook_path: str,
    left_sheet: str,
    right_sheet: str,
    case_type: str,
    max_rows: int = 5000,
) -> pd.DataFrame:
    """
    Used by the GUI "live view" for a single case type.
    """
    df_pair = build_pair_comparison_df(
        workbook_path=workbook_path,
        left_sheet=left_sheet,
        right_sheet=right_sheet,
        threshold=0.0,
        log_func=None,
    )

    df = df_pair[df_pair["CaseType"] == case_type].copy()
    if len(df) > max_rows:
        df = df.head(max_rows).copy()
    return df.reset_index(drop=True)


def compare_scenarios(
    workbook_path: str,
    left_sheet: str,
    right_sheet: str,
    threshold: float = 0.0,
    mode: str = "expanded",
) -> pd.DataFrame:
    """
    Convenience: full-sheet comparison (all case types). The GUI may or may not use this.
    """
    return build_pair_comparison_df(
        workbook_path=workbook_path,
        left_sheet=left_sheet,
        right_sheet=right_sheet,
        threshold=threshold,
        log_func=None,
    )


def build_pair_comparison_df(
    workbook_path: str,
    left_sheet: str,
    right_sheet: str,
    threshold: float,
    log_func=None,
) -> pd.DataFrame:
    """
    Returns a pairwise comparison dataframe with columns expected by batch_sheet_writer:
      CaseType, Contingency, ResultingIssue, LeftPct, RightPct, DeltaPct, DeltaDisplay
    """
    log = log_func or (lambda _m: None)

    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        df_left = _parse_scenario_sheet(wb[left_sheet])
        df_right = _parse_scenario_sheet(wb[right_sheet])
    finally:
        wb.close()

    df = _compare_two(df_left, df_right)

    if threshold and threshold > 0:
        keep = (df["LeftPct"].fillna(-1) >= threshold) | (df["RightPct"].fillna(-1) >= threshold)
        df = df.loc[keep].copy()

    # Delta display
    df["DeltaDisplay"] = df["DeltaPct"].apply(_fmt_delta)

    # Sort (case type order, then max pct desc)
    df["_max"] = df[["LeftPct", "RightPct"]].max(axis=1, skipna=True)
    df["_case_rank"] = df["CaseType"].apply(_case_type_rank)

    df.sort_values(
        by=["_case_rank", "_max", "DeltaPct", "Contingency", "ResultingIssue"],
        ascending=[True, False, False, True, True],
        inplace=True,
        kind="mergesort",
    )

    df.drop(columns=["_max", "_case_rank"], inplace=True)
    return df.reset_index(drop=True)


def build_batch_comparison_workbook(
    src_workbook: str,
    pairs: Sequence[Tuple[str, str]],
    threshold: float,
    output_path: str,
    log_func=None,
    **_kwargs,   # extra safety if older callers pass unexpected keywords
) -> str:
    """
    Builds your batch workbook:
      - queued pairwise sheets
      - ALWAYS appends a final sheet: "Straight Comparison" (all original sheets wide)
    """
    log = log_func or (lambda _m: None)
    pairs = list(pairs or [])

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    # Load source sheet order once
    wb_src = load_workbook(src_workbook, read_only=True, data_only=True)
    try:
        src_sheet_order = list(wb_src.sheetnames)
    finally:
        wb_src.close()

    wb_out = Workbook()
    # remove default
    if wb_out.worksheets:
        wb_out.remove(wb_out.worksheets[0])

    # -------------------------
    # 1) queued pairwise sheets
    # -------------------------
    log(f"Building queued workbook with {len(pairs)} pair sheet(s)...")

    for left_sheet, right_sheet in pairs:
        log(f"Comparing: {left_sheet} vs {right_sheet}")
        df_pair = build_pair_comparison_df(
            workbook_path=src_workbook,
            left_sheet=left_sheet,
            right_sheet=right_sheet,
            threshold=threshold,
            log_func=None,  # keep logging here, not inside parser loop
        )

        desired_name = f"{left_sheet} vs {right_sheet}"
        ws_name = _safe_unique_sheet_name(wb_out, desired_name)

        if write_formatted_pair_sheet is not None:
            # batch_sheet_writer creates & formats the sheet
            write_formatted_pair_sheet(
                wb=wb_out,
                ws_name=ws_name,
                df_pair=df_pair,
                expandable_issue_view=True,
            )
        else:
            # fallback (still works)
            ws = wb_out.create_sheet(ws_name)
            _write_df_basic(ws, df_pair)

        log(f"  Added sheet: {ws_name}")

    # -------------------------
    # 2) Straight Comparison (ALWAYS LAST)
    # -------------------------
    log("Building final sheet: Straight Comparison (all originals together)...")

    # "Original sheets" = all sheets in the source workbook, in their original order.
    # (This guarantees “big picture” no matter how the user queued pairs.)
    straight_sheet_order = src_sheet_order[:]

    df_straight = _build_straight_comparison_df(
        workbook_path=src_workbook,
        sheets=straight_sheet_order,
        threshold=threshold,
    )

    ws_compare_name = _safe_unique_sheet_name(wb_out, "Straight Comparison")
    ws_compare = wb_out.create_sheet(ws_compare_name)
    _write_straight_comparison_sheet(
        ws=ws_compare,
        df=df_straight,
        sheet_order=straight_sheet_order,
    )

    log(f"  Added sheet: {ws_compare_name}")

    wb_out.save(output_path)
    log(f"Saved: {output_path}")
    return output_path


# -------------------------
# Internals
# -------------------------

def _parse_scenario_sheet(ws) -> pd.DataFrame:
    """
    Normalizes a scenario sheet into:
      CaseType, Contingency, ResultingIssue, Limit, Percent

    Assumes:
      - 2 header rows
      - case-type headers appear as [None, "Branch MVA", None, None, None]
      - data columns:
          A: Contingency
          B: ResultingIssue
          D: Limit (MVA)   (optional)
          E: Percent       (optional)
    """
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) <= 2:
        return pd.DataFrame(columns=["CaseType", "Contingency", "ResultingIssue", "Limit", "Percent"])

    data = rows[2:]  # skip 2 header rows
    cur_case = None
    out = []

    for r in data:
        a = r[0] if len(r) > 0 else None
        b = r[1] if len(r) > 1 else None
        d = r[3] if len(r) > 3 else None
        e = r[4] if len(r) > 4 else None

        # case-type header row
        if a is None and isinstance(b, str) and b.strip() and (d is None and e is None):
            cur_case = b.strip()
            continue

        out.append(
            {
                "CaseType": cur_case,
                "Contingency": a,
                "ResultingIssue": b,
                "Limit": d,
                "Percent": e,
            }
        )

    df = pd.DataFrame(out)

    # Forward-fill continuity rows (so keys stay stable)
    for c in ["Contingency", "ResultingIssue", "CaseType"]:
        df[c] = df[c].astype("string")

    df["Contingency"] = df["Contingency"].where(df["Contingency"].notna() & (df["Contingency"].str.strip() != ""), pd.NA)
    df["ResultingIssue"] = df["ResultingIssue"].where(df["ResultingIssue"].notna() & (df["ResultingIssue"].str.strip() != ""), pd.NA)

    df["Contingency"] = df["Contingency"].ffill()
    df["ResultingIssue"] = df["ResultingIssue"].ffill()

    df["Limit"] = pd.to_numeric(df["Limit"], errors="coerce")
    df["Percent"] = pd.to_numeric(df["Percent"], errors="coerce")

    # drop totally empty
    df = df.loc[~(df["Contingency"].isna() & df["ResultingIssue"].isna())].copy()
    return df.reset_index(drop=True)


def _compare_two(df_left: pd.DataFrame, df_right: pd.DataFrame) -> pd.DataFrame:
    key = ["CaseType", "Contingency", "ResultingIssue"]

    L = df_left[key + ["Percent"]].copy().rename(columns={"Percent": "LeftPct"})
    R = df_right[key + ["Percent"]].copy().rename(columns={"Percent": "RightPct"})

    # Collapse duplicates (messy source protection)
    L = L.groupby(key, dropna=False, as_index=False).agg({"LeftPct": "max"})
    R = R.groupby(key, dropna=False, as_index=False).agg({"RightPct": "max"})

    df = pd.merge(L, R, on=key, how="outer")
    df["DeltaPct"] = df["RightPct"] - df["LeftPct"]
    return df


def _build_straight_comparison_df(workbook_path: str, sheets: List[str], threshold: float) -> pd.DataFrame:
    """
    Wide compare across ALL original sheets:
      Category (CaseType), Contingency, ResultingIssue, Limit,
      then one column per sheet with Percent.
    """
    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        parsed: Dict[str, pd.DataFrame] = {}
        for s in sheets:
            if s not in wb.sheetnames:
                continue
            parsed[s] = _parse_scenario_sheet(wb[s])
    finally:
        wb.close()

    # Build union key set
    key_cols = ["CaseType", "Contingency", "ResultingIssue"]

    # base key frame
    keys = None
    limits_max = None
    pct_series: Dict[str, pd.Series] = {}

    for s in sheets:
        df = parsed.get(s)
        if df is None or df.empty:
            continue

        g = df.groupby(key_cols, dropna=False, as_index=True).agg(
            Limit=("Limit", "max"),
            Percent=("Percent", "max"),
        )

        if keys is None:
            keys = g.index.to_frame(index=False)
            limits_max = g["Limit"]
        else:
            keys = pd.concat([keys, g.index.to_frame(index=False)], ignore_index=True).drop_duplicates()
            if limits_max is not None:
                limits_max = pd.concat([limits_max, g["Limit"]]).groupby(level=[0, 1, 2]).max()

        pct_series[s] = g["Percent"]

    if keys is None or keys.empty:
        return pd.DataFrame(columns=["Category", "Contingency", "ResultingIssue", "Limit"])

    out = keys.copy()
    out.rename(columns={"CaseType": "Category"}, inplace=True)

    # attach limit
    mi = pd.MultiIndex.from_frame(keys[key_cols])
    out["Limit"] = limits_max.reindex(mi).to_numpy() if limits_max is not None else pd.NA

    # attach each sheet's percent
    for s in sheets:
        ser = pct_series.get(s)
        if ser is None:
            out[s] = pd.NA
        else:
            out[s] = ser.reindex(mi).to_numpy()

    # threshold filter: keep if ANY >= threshold
    pct_cols = [s for s in sheets if s in out.columns]
    if threshold and threshold > 0 and pct_cols:
        maxpct = out[pct_cols].apply(pd.to_numeric, errors="coerce").max(axis=1, skipna=True)
        out = out.loc[maxpct.fillna(-1) >= threshold].copy()

    # sort: case type rank, then max percent desc
    if pct_cols:
        out["_maxpct"] = out[pct_cols].apply(pd.to_numeric, errors="coerce").max(axis=1, skipna=True)
    else:
        out["_maxpct"] = pd.NA
    out["_case_rank"] = out["Category"].astype("string").apply(_case_type_rank)

    out.sort_values(
        by=["_case_rank", "_maxpct", "Contingency", "ResultingIssue"],
        ascending=[True, False, True, True],
        inplace=True,
        kind="mergesort",
    )
    out.drop(columns=["_maxpct", "_case_rank"], inplace=True)
    return out.reset_index(drop=True)


def _write_straight_comparison_sheet(ws, df: pd.DataFrame, sheet_order: List[str]) -> None:
    fixed = ["Contingency", "ResultingIssue", "Category", "Limit"]
    pct_cols = [s for s in sheet_order if s in df.columns]
    cols = fixed + pct_cols

    # header
    for j, c in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=j, value=c)
        cell.font = cell.font.copy(bold=True)

    # rows
    for i, row in enumerate(df[cols].itertuples(index=False), start=2):
        for j, v in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=_excel_safe(v))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"

    # widths
    _set_width(ws, 1, 55)  # contingency
    _set_width(ws, 2, 55)  # resulting issue
    _set_width(ws, 3, 18)  # category
    _set_width(ws, 4, 12)  # limit
    for k in range(5, len(cols) + 1):
        _set_width(ws, k, min(26, max(10, len(str(cols[k - 1])) + 2)))

    # formats
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=4).number_format = "0.000"  # limit
    for col_idx in range(5, len(cols) + 1):
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=col_idx).number_format = "0.000"  # percent

    # optional styling
    if apply_table_styles is not None:
        try:
            apply_table_styles(ws)
        except Exception:
            pass


def _safe_unique_sheet_name(wb: Workbook, desired: str) -> str:
    """
    Excel rules:
      - max 31 chars
      - cannot contain: : \ / ? * [ ]
      - must be unique
    """
    name = (desired or "Sheet").strip()
    for ch in [":", "\\", "/", "?", "*", "[", "]"]:
        name = name.replace(ch, " ")
    name = " ".join(name.split())  # collapse whitespace
    name = name[:31]

    if name not in wb.sheetnames:
        return name

    # suffix with (2), (3)...
    for n in range(2, 999):
        suffix = f" ({n})"
        base = name[: 31 - len(suffix)]
        cand = (base + suffix)[:31]
        if cand not in wb.sheetnames:
            return cand

    return (name[:28] + "_dup")[:31]


def _write_df_basic(ws, df: pd.DataFrame) -> None:
    # minimal fallback writer
    for j, c in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=j, value=c)
        cell.font = cell.font.copy(bold=True)
    for i, row in enumerate(df.itertuples(index=False), start=2):
        for j, v in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=_excel_safe(v))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(df.shape[1])}{ws.max_row}"


def _fmt_delta(v) -> Optional[str]:
    if pd.isna(v):
        return None
    try:
        v = float(v)
    except Exception:
        return None
    sign = "+" if v > 0 else ""
    return f"{sign}{v:.3f}"


def _case_type_rank(ct: Optional[str]) -> int:
    if ct is None:
        return 999
    ct = str(ct).strip()
    try:
        return CASE_TYPES_CANONICAL.index(ct)
    except ValueError:
        return 999


def _excel_safe(v):
    if pd.isna(v):
        return None
    try:
        import numpy as np
        if isinstance(v, np.generic):
            return v.item()
    except Exception:
        pass
    return v


def _set_width(ws, col_idx: int, width: float) -> None:
    ws.column_dimensions[get_column_letter(col_idx)].width = float(width)