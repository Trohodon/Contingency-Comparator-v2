# comparator.py
# Core comparison/build logic used by Tab Compare
# - Builds "pair" sheets (A vs B) into a single output workbook
# - NEW: Adds a final "Compare" sheet that straight-compares all involved source sheets side-by-side

from __future__ import annotations

import os
import time
import re
from typing import Callable, Dict, List, Optional, Sequence, Tuple

import pandas as pd

# Optional dependency: openpyxl
OPENPYXL_AVAILABLE = True
try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter
except Exception:
    OPENPYXL_AVAILABLE = False
    Workbook = object  # type: ignore

# Import sheet writer used for formatted pair sheets
from batch_sheet_writer import write_formatted_pair_sheet


# -------------------------
# Helpers
# -------------------------

def _safe_float(x) -> Optional[float]:
    try:
        if x is None or (isinstance(x, float) and pd.isna(x)):
            return None
        return float(x)
    except Exception:
        return None


def _norm_str(x) -> str:
    if x is None:
        return ""
    try:
        s = str(x)
    except Exception:
        return ""
    return s.strip()


def _sanitize_sheet_name(name: str) -> str:
    """
    Excel sheet name constraints:
      - max 31 chars
      - cannot contain: : \ / ? * [ ]
    """
    name = _norm_str(name)
    name = re.sub(r"[:\\\/\?\*\[\]]", "-", name)
    name = name.strip()
    if len(name) > 31:
        name = name[:31]
    if not name:
        name = "Sheet"
    return name


def _ws_rows_to_df(ws) -> pd.DataFrame:
    """
    Reads a worksheet into a DataFrame, using the first non-empty row as header.
    """
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame()

    # Find first non-empty row to use as header
    header_idx = None
    for i, r in enumerate(rows):
        if r and any(v is not None and str(v).strip() != "" for v in r):
            header_idx = i
            break
    if header_idx is None:
        return pd.DataFrame()

    header = [_norm_str(v) for v in rows[header_idx]]
    data = rows[header_idx + 1 :]
    df = pd.DataFrame(data, columns=header)

    # Drop fully empty rows
    df = df.dropna(how="all")
    return df


def _parse_scenario_sheet(ws) -> pd.DataFrame:
    """
    Parses a formatted "scenario" sheet produced by Tab Case.

    Expected columns include:
      - CaseType
      - CTGLabel
      - LimViolID
      - LimViolValue
      - LimViolPct

    Handles the "blank resulting issues" problem by forward-filling LimViolID
    within each CaseType block.
    """
    df = _ws_rows_to_df(ws)
    if df.empty:
        return df

    # Normalize col names
    rename_map = {}
    for c in df.columns:
        c_norm = _norm_str(c).lower()
        if c_norm in ["casetype", "case type"]:
            rename_map[c] = "CaseType"
        elif c_norm in ["ctglabel", "contingency", "contingency event", "name"]:
            rename_map[c] = "CTGLabel"
        elif c_norm in ["limviolid", "resulting issue", "element", "issue", "lim viol id"]:
            rename_map[c] = "LimViolID"
        elif c_norm in ["limviolvalue", "limit", "lim viol value"]:
            rename_map[c] = "LimViolValue"
        elif c_norm in ["limviolpct", "percent", "percent loading", "lim viol pct"]:
            rename_map[c] = "LimViolPct"

    df = df.rename(columns=rename_map)

    required = ["CaseType", "CTGLabel", "LimViolID", "LimViolPct"]
    for r in required:
        if r not in df.columns:
            return pd.DataFrame()

    # Clean types
    df["CaseType"] = df["CaseType"].fillna("").astype(str).str.strip()
    df["CTGLabel"] = df["CTGLabel"].fillna("").astype(str).str.strip()
    df["LimViolID"] = df["LimViolID"].fillna("").astype(str).str.strip()

    # Percent numeric
    df["LimViolPct"] = df["LimViolPct"].apply(_safe_float)

    # Forward-fill LimViolID within each CaseType block (blank ID means "same as above")
    # We also forward fill CTGLabel sometimes, but typically CTGLabel is always present.
    df["LimViolID"] = df.groupby("CaseType")["LimViolID"].apply(
        lambda s: s.replace("", pd.NA).ffill().fillna("")
    )

    # Drop rows without essential fields
    df = df[(df["CTGLabel"] != "") & (df["LimViolID"] != "")]
    df = df.dropna(subset=["LimViolPct"])
    return df


def build_case_type_comparison(
    ws_left,
    ws_right,
    *,
    left_name: str,
    right_name: str,
) -> pd.DataFrame:
    """
    Builds a row-aligned comparison DataFrame for one pair of source sheets.
    """
    df_left = _parse_scenario_sheet(ws_left)
    df_right = _parse_scenario_sheet(ws_right)

    if df_left.empty and df_right.empty:
        return pd.DataFrame()

    # Reduce to max pct per (CaseType, CTGLabel, LimViolID)
    def reduce_df(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        g = (
            df.groupby(["CaseType", "CTGLabel", "LimViolID"], dropna=False)["LimViolPct"]
            .max()
            .reset_index()
        )
        return g

    df_left = reduce_df(df_left)
    df_right = reduce_df(df_right)

    # Outer merge to keep all issues
    df = pd.merge(
        df_left,
        df_right,
        on=["CaseType", "CTGLabel", "LimViolID"],
        how="outer",
        suffixes=(f" ({left_name})", f" ({right_name})"),
    )

    # Rename pct columns to stable internal names
    left_col = "LimViolPct (Left)"
    right_col = "LimViolPct (Right)"

    # Figure out actual column names after merge
    left_pct_cols = [c for c in df.columns if c.startswith("LimViolPct") and left_name in c]
    right_pct_cols = [c for c in df.columns if c.startswith("LimViolPct") and right_name in c]

    if left_pct_cols:
        df = df.rename(columns={left_pct_cols[0]: left_col})
    else:
        df[left_col] = None

    if right_pct_cols:
        df = df.rename(columns={right_pct_cols[0]: right_col})
    else:
        df[right_col] = None

    df[left_col] = df[left_col].apply(_safe_float)
    df[right_col] = df[right_col].apply(_safe_float)

    df["MaxPct"] = df[[left_col, right_col]].max(axis=1, skipna=True)
    df["Delta"] = df[right_col] - df[left_col]

    # Sort biggest first
    df = df.sort_values(by=["MaxPct"], ascending=False, na_position="last").reset_index(drop=True)

    return df


# -------------------------
# NEW: Straight compare writer
# -------------------------

def write_straight_compare_sheet(
    wb_out: "Workbook",
    src_wb: "Workbook",
    sheet_names: Sequence[str],
    *,
    sheet_title: str = "Compare",
    log_func: Optional[Callable[[str], None]] = None,
) -> None:
    """Create a side-by-side 'straight compare' sheet.

    Columns:
      Contingency Event | Resulting Issue | <Sheet 1 Percent> | <Sheet 2 Percent> | ...

    - Uses the union of (CTGLabel, LimViolID) found in the given sheet_names.
    - For each sheet, if a (CTGLabel, LimViolID) appears multiple times (e.g., across categories),
      we keep the *maximum* Percent Loading for that key in that sheet.
    - The highest percent in each row is bolded to make it easy to spot the worst case.
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to build the comparison workbook.")

    # Defensive: keep order, remove duplicates
    ordered_names: List[str] = []
    seen: set = set()
    for n in sheet_names:
        if n and n not in seen and n in src_wb.sheetnames:
            ordered_names.append(n)
            seen.add(n)

    if not ordered_names:
        if log_func:
            log_func("No source sheets found for straight comparison; skipping Compare sheet.")
        return

    # Remove existing sheet if re-building
    if sheet_title in wb_out.sheetnames:
        del wb_out[sheet_title]
    ws = wb_out.create_sheet(sheet_title)

    header_fill = PatternFill("solid", fgColor="1F4E79")  # dark blue
    header_font = Font(bold=True, color="FFFFFF")
    subheader_fill = PatternFill("solid", fgColor="D9E1F2")  # light blue-gray
    subheader_font = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    thin = Side(style="thin", color="A6A6A6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Two-row header (like your screenshot)
    ws["A1"] = "Contingency Event"
    ws["B1"] = "Resulting Issue"
    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    for addr in ("A1", "B1"):
        ws[addr].fill = header_fill
        ws[addr].font = header_font
        ws[addr].alignment = center
        ws[addr].border = border

    # Data gather
    key_labels: Dict[Tuple[str, str], Tuple[str, str]] = {}
    sheet_map: Dict[str, Dict[Tuple[str, str], float]] = {}

    for idx, sheet_name in enumerate(ordered_names, start=1):
        if log_func:
            log_func(f"Parsing sheet for straight compare: {sheet_name} ({idx}/{len(ordered_names)})")
        src_ws = src_wb[sheet_name]
        df = _parse_scenario_sheet(src_ws)
        if df.empty:
            sheet_map[sheet_name] = {}
            continue

        df = df.copy()
        df["CTGLabel"] = df["CTGLabel"].fillna("").astype(str)
        df["LimViolID"] = df["LimViolID"].fillna("").astype(str)

        g = (
            df.groupby(["CTGLabel", "LimViolID"], dropna=False)["LimViolPct"]
            .max()
            .reset_index()
        )

        mp: Dict[Tuple[str, str], float] = {}
        for _, r in g.iterrows():
            key = (str(r["CTGLabel"]), str(r["LimViolID"]))
            pct = _safe_float(r["LimViolPct"])
            if pct is None:
                continue
            mp[key] = pct
            if key not in key_labels:
                key_labels[key] = key
        sheet_map[sheet_name] = mp

        # Yield briefly so the UI stays responsive
        time.sleep(0)

    all_keys = set()
    for mp in sheet_map.values():
        all_keys.update(mp.keys())

    def row_max(k: Tuple[str, str]) -> float:
        vals = [sheet_map[s].get(k) for s in ordered_names]
        vals = [v for v in vals if v is not None]
        return max(vals) if vals else -1.0

    ordered_keys = sorted(all_keys, key=row_max, reverse=True)

    # Percent columns headers
    start_col = 3
    for i, sheet_name in enumerate(ordered_names):
        col = start_col + i
        cell1 = ws.cell(row=1, column=col, value=sheet_name)
        cell2 = ws.cell(row=2, column=col, value="Percent")

        cell1.fill = header_fill
        cell1.font = header_font
        cell1.alignment = center
        cell1.border = border

        cell2.fill = subheader_fill
        cell2.font = subheader_font
        cell2.alignment = center
        cell2.border = border

    # Data rows start at row 3
    r_out = 3
    for k in ordered_keys:
        ctg, issue = key_labels.get(k, k)
        ws.cell(row=r_out, column=1, value=ctg).alignment = left
        ws.cell(row=r_out, column=2, value=issue).alignment = left

        vals: List[Optional[float]] = []
        for i, sheet_name in enumerate(ordered_names):
            v = sheet_map[sheet_name].get(k)
            vals.append(v)
            c = ws.cell(row=r_out, column=start_col + i, value=v)
            if v is not None:
                c.number_format = "0.000"
            c.alignment = center

        # Bold the max value(s)
        present = [v for v in vals if v is not None]
        if present:
            vmax = max(present)
            for i, v in enumerate(vals):
                if v is not None and abs(v - vmax) < 1e-9:
                    ws.cell(row=r_out, column=start_col + i).font = Font(bold=True)

        # Borders
        for c in range(1, start_col + len(ordered_names)):
            ws.cell(row=r_out, column=c).border = border

        r_out += 1
        if (r_out % 250) == 0:
            time.sleep(0)

    # Column sizing
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 65
    for i in range(len(ordered_names)):
        col_letter = get_column_letter(start_col + i)
        ws.column_dimensions[col_letter].width = 16

    # Freeze headers + first 2 columns
    ws.freeze_panes = "C3"


# -------------------------
# Public API: build batch workbook
# -------------------------

def build_batch_comparison_workbook(
    tasks: List[Dict],
    output_path: str,
    *,
    src_workbook: Optional[str] = None,
    workbook_path: Optional[str] = None,
    expandable_issue_view: bool = True,
    log_func: Optional[Callable[[str], None]] = None,
) -> None:
    """
    Build a batch comparison workbook containing one sheet per queued comparison.

    tasks: List of dicts containing at least:
      - "LeftSheet"
      - "RightSheet"
      - "LeftName"
      - "RightName"
      - "SheetName"

    src_workbook/workbook_path: path to the source workbook that contains the sheets.
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to build the comparison workbook.")

    src_path = src_workbook or workbook_path
    if not src_path or not os.path.isfile(src_path):
        raise FileNotFoundError("Source workbook path not provided or does not exist.")

    if log_func:
        log_func(f"Loading source workbook: {src_path}")

    src_wb = load_workbook(src_path, data_only=True)
    wb = Workbook()
    # Remove default sheet
    try:
        default = wb.active
        wb.remove(default)
    except Exception:
        pass

    total = len(tasks)
    for idx, task in enumerate(tasks, start=1):
        left_sheet = task.get("LeftSheet", "")
        right_sheet = task.get("RightSheet", "")
        left_name = task.get("LeftName", "Left")
        right_name = task.get("RightName", "Right")
        out_sheet_name = task.get("SheetName") or f"{left_name} vs {right_name}"
        out_sheet_name = _sanitize_sheet_name(out_sheet_name)

        if log_func:
            log_func(f"[{idx}/{total}] Building: {out_sheet_name}")

        if left_sheet not in src_wb.sheetnames or right_sheet not in src_wb.sheetnames:
            if log_func:
                log_func(f"Skipping '{out_sheet_name}' (missing sheet in source workbook).")
            continue

        df_pair = build_case_type_comparison(
            src_wb[left_sheet],
            src_wb[right_sheet],
            left_name=left_name,
            right_name=right_name,
        )

        if df_pair.empty:
            if log_func:
                log_func(f"Skipping '{out_sheet_name}' (no comparable data).")
            continue

        # Ensure unique output sheet name
        name = out_sheet_name
        suffix = 2
        while name in wb.sheetnames:
            name = _sanitize_sheet_name(f"{out_sheet_name} ({suffix})")
            suffix += 1

        write_formatted_pair_sheet(
            wb,
            name,
            df_pair,
            expandable_issue_view=expandable_issue_view,
        )

        # Yield briefly so the Tk UI doesn't appear "Not Responding" while we build large workbooks
        time.sleep(0)

    # Final "straight compare" sheet (side-by-side Percent Loading across all source sheets used)
    try:
        compare_sheet_names: List[str] = []
        _seen_sheets: set = set()
        for t in tasks:
            for s in (t.get("LeftSheet"), t.get("RightSheet")):
                if s and s not in _seen_sheets:
                    compare_sheet_names.append(s)
                    _seen_sheets.add(s)

        if compare_sheet_names:
            write_straight_compare_sheet(
                wb,
                src_wb,
                compare_sheet_names,
                sheet_title="Compare",
                log_func=log_func,
            )
    except Exception as e:
        if log_func:
            log_func(f"Warning: could not build straight Compare sheet: {e}")

    wb.save(output_path)
    if log_func:
        log_func(f"Saved batch workbook: {output_path}")