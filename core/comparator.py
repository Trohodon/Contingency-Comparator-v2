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

from typing import Iterable, List, Dict, Optional, Sequence, Tuple

import math
import os
import pandas as pd

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


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


def _parse_scenario_sheet(ws: "Worksheet", log_func=None) -> pd.DataFrame:
    """
    Parse one formatted scenario sheet into a DataFrame with columns:

        CaseType, CTGLabel, LimViolID, LimViolValue, LimViolPct

    Sheet structure (from comparison_builder):
      - Title row: merged B:E, text is pretty case name ("ACCA LongTerm", "ACCA", "DCwAC")
      - Next row: headers in B..E
      - Following rows: data until blank row; then repeat for next case type.
    """
    records: List[Dict] = []

    max_row = ws.max_row or 1
    row_idx = 1

    while row_idx <= max_row:
        title_cell = ws.cell(row=row_idx, column=2)  # column B
        title_val = title_cell.value

        # Identify a title row by non-empty text in B
        if isinstance(title_val, str) and title_val.strip():
            pretty_name = title_val.strip()
            case_type = CANONICAL_CASE_TYPES.get(pretty_name, pretty_name)

            # Header row is next; data row follows
            header_row = row_idx + 1
            data_row = header_row + 1

            r = data_row
            while r <= max_row:
                b = ws.cell(row=r, column=2).value  # CTGLabel / contingency
                c = ws.cell(row=r, column=3).value  # LimViolID / resulting issue
                d = ws.cell(row=r, column=4).value  # LimViolValue (MVA)
                e = ws.cell(row=r, column=5).value  # LimViolPct (% loading)

                # Stop when B..E are all blank
                if (
                    (b is None or str(b).strip() == "")
                    and (c is None or str(c).strip() == "")
                    and (d is None or str(d).strip() == "")
                    and (e is None or str(e).strip() == "")
                ):
                    break

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

            # Jump to the row after the blank separator
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
    df = _parse_scenario_sheet(ws, log_func=log_func)
    return df


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

    Returns a DataFrame with columns:

        Contingency, ResultingIssue, LeftPct, RightPct, DeltaPct

    - case_type must be the canonical string, e.g. "ACCA_LongTerm".
    - If BOTH sheets have zero rows for this case_type, the DataFrame is empty
      (meaning 'no issues' for that case type).
    - If only one sheet has rows, the other side will be NaN (shown as blank
      in the GUI) – effectively "no issues" on that side.

    Rows are sorted by Percent Loading (highest to lowest). The primary
    sort key is RightPct (new sheet). If all RightPct values are NaN,
    we fall back to LeftPct.
    """
    if case_type not in CASE_TYPES_CANONICAL:
        raise ValueError(f"Unknown case type: {case_type}")

    # Load full scenario sheets
    base_df = _load_sheet_as_df(workbook_path, base_sheet, log_func=log_func)
    new_df = _load_sheet_as_df(workbook_path, new_sheet, log_func=log_func)

    # Filter to this case type
    base_df = base_df[base_df["CaseType"] == case_type].copy()
    new_df = new_df[new_df["CaseType"] == case_type].copy()

    if log_func:
        log_func(
            f"  [{case_type}] base rows={len(base_df)}, new rows={len(new_df)}"
        )

    # If both are empty, just return an empty DF (no contingencies for this type)
    if base_df.empty and new_df.empty:
        if log_func:
            log_func(f"  [{case_type}] No contingencies in either sheet.")
        return pd.DataFrame(
            columns=["Contingency", "ResultingIssue", "LeftPct", "RightPct", "DeltaPct"]
        )

    # Rename for Left / Right
    base_df = base_df.rename(
        columns={"LimViolValue": "Left_MVA", "LimViolPct": "Left_Pct"}
    )
    new_df = new_df.rename(
        columns={"LimViolValue": "Right_MVA", "LimViolPct": "Right_Pct"}
    )

    key_cols = ["CTGLabel", "LimViolID"]
    left_cols = key_cols + ["Left_Pct"]
    right_cols = key_cols + ["Right_Pct"]

    merged = pd.merge(
        base_df[left_cols],
        new_df[right_cols],
        on=key_cols,
        how="outer",
    )

    # Delta = Right - Left
    merged["Delta_Pct"] = merged["Right_Pct"] - merged["Left_Pct"]

    # Rename into user-friendly column names
    result = merged.rename(
        columns={
            "CTGLabel": "Contingency",
            "LimViolID": "ResultingIssue",
            "Left_Pct": "LeftPct",
            "Right_Pct": "RightPct",
            "Delta_Pct": "DeltaPct",
        }
    )

    # Choose a sort key so items are numbered highest-to-lowest by Percent Loading.
    # Prefer RightPct (new scenario). If all RightPct are NaN, fall back to LeftPct.
    sort_series = result["RightPct"]
    if sort_series.notna().any():
        result["_SortPct"] = sort_series
    else:
        result["_SortPct"] = result["LeftPct"]

    # Sort descending (highest loading first). NaNs go last.
    result = result.sort_values(
        by="_SortPct", ascending=False, na_position="last"
    ).drop(columns=["_SortPct"])

    if max_rows is not None and max_rows > 0:
        result = result.head(max_rows)

    return result


# ---------------------------------------------------------------------------
# Workbook-level comparison sheet (existing feature)
# ---------------------------------------------------------------------------


def compare_scenarios(
    workbook_path: str,
    base_sheet: str,
    new_sheet: str,
    case_types_to_include: Optional[Iterable[str]] = None,
    pct_threshold: float = 0.0,
    mode: str = "all",  # "all", "worse", "better"
    log_func=None,
) -> str:
    """
    Compare two scenario sheets inside the same workbook and write the result
    into a new comparison sheet.

    Returns the workbook_path (same file) when successful.
    """
    if log_func:
        log_func(
            f"\n=== Compare scenarios ===\n"
            f"Workbook: {workbook_path}\n"
            f"Base: {base_sheet}\n"
            f"New:  {new_sheet}\n"
            f"Mode: {mode}, Threshold: {pct_threshold} %"
        )

    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    # Load base & new
    base_df = _load_sheet_as_df(workbook_path, base_sheet, log_func=log_func)
    new_df = _load_sheet_as_df(workbook_path, new_sheet, log_func=log_func)

    # Optional case-type filter
    if case_types_to_include:
        case_set = set(case_types_to_include)
        base_df = base_df[base_df["CaseType"].isin(case_set)]
        new_df = new_df[new_df["CaseType"].isin(case_set)]

    # Normalize column names
    base_df = base_df.rename(
        columns={
            "LimViolValue": "Base_MVA",
            "LimViolPct": "Base_Pct",
        }
    )
    new_df = new_df.rename(
        columns={
            "LimViolValue": "New_MVA",
            "LimViolPct": "New_Pct",
        }
    )

    # Merge on key
    key_cols = ["CaseType", "CTGLabel", "LimViolID"]
    merged = pd.merge(
        base_df,
        new_df,
        on=key_cols,
        how="outer",
        suffixes=("", "_y"),
    )

    # Compute deltas
    merged["Delta_Pct"] = merged["New_Pct"] - merged["Base_Pct"]
    merged["Delta_MVA"] = merged["New_MVA"] - merged["Base_MVA"]

    # Status classification
    status_col: List[str] = []
    for _, row in merged.iterrows():
        base_pct = row["Base_Pct"]
        new_pct = row["New_Pct"]

        if pd.isna(base_pct) and pd.isna(new_pct):
            status_col.append("Unknown")
        elif pd.isna(base_pct):
            status_col.append("Only in new")
        elif pd.isna(new_pct):
            status_col.append("Only in base")
        else:
            delta = new_pct - base_pct
            if abs(delta) < pct_threshold:
                status_col.append("Same (within threshold)")
            elif delta > 0:
                status_col.append("Worse in new")
            elif delta < 0:
                status_col.append("Better in new")
            else:
                status_col.append("Same")

    merged["Status"] = status_col

    # Apply mode filter
    def keep_row(r) -> bool:
        base_pct = r["Base_Pct"]
        new_pct = r["New_Pct"]
        status = r["Status"]

        # Always keep unmatched rows
        if status in ("Only in new", "Only in base"):
            return True

        if pd.isna(base_pct) or pd.isna(new_pct):
            return True  # weird, but keep

        delta = new_pct - base_pct
        if abs(delta) < pct_threshold:
            return False  # below threshold

        if mode == "worse":
            return delta > 0
        elif mode == "better":
            return delta < 0
        else:
            # "all"
            return True

    filtered = merged[merged.apply(keep_row, axis=1)].copy()

    # Sort filtered results: highest New_Pct first (worse loading at top).
    if "New_Pct" in filtered.columns:
        filtered = filtered.sort_values(
            by="New_Pct", ascending=False, na_position="last"
        )

    if log_func:
        log_func(
            f"Matched/merged rows: {len(merged)}; rows after filters: {len(filtered)}"
        )

    # Write result into the same workbook as a new sheet
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to write comparison sheet.")

    from openpyxl import load_workbook as _load_wb  # alias to avoid confusion

    wb = _load_wb(workbook_path)
    base_short = base_sheet[:10]
    new_short = new_sheet[:10]
    comp_name = f"Compare_{base_short}_vs_{new_short}"
    # Trim to Excel limit
    comp_name = comp_name[:31]

    # If sheet already exists, delete it so we replace with fresh results
    if comp_name in wb.sheetnames:
        del wb[comp_name]

    ws = wb.create_sheet(title=comp_name)

    # Order & names for output columns
    cols = [
        "CaseType",
        "CTGLabel",
        "LimViolID",
        "Base_MVA",
        "Base_Pct",
        "New_MVA",
        "New_Pct",
        "Delta_Pct",
        "Delta_MVA",
        "Status",
    ]

    # Header
    for col_idx, col_name in enumerate(cols, start=1):
        ws.cell(row=1, column=col_idx).value = col_name

    # Data rows
    for row_idx, (_, row) in enumerate(filtered.iterrows(), start=2):
        for col_idx, col_name in enumerate(cols, start=1):
            ws.cell(row=row_idx, column=col_idx).value = row.get(col_name)

    wb.save(workbook_path)

    if log_func:
        log_func(f"Comparison sheet '{comp_name}' written to workbook.")

    return workbook_path


# ---------------------------------------------------------------------------
# Batch export helpers for the build-list queue
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
    Build a single flat DataFrame for one (left_sheet, right_sheet) pair,
    combining all case types.

    Columns:
      CaseType, Contingency, ResultingIssue, LeftPct, RightPct, DeltaDisplay

    Threshold logic matches the GUI:
      - if BOTH Left% and Right% exist and are < threshold -> row is skipped
      - if only Left% exists -> kept only if Left% >= threshold
      - if only Right% exists -> kept only if Right% >= threshold

    DeltaDisplay:
      - "Only in left"   (only left present)
      - "Only in right"  (only right present)
      - ""               (neither present, but that should be rare)
      - numeric string   (Right - Left, 2 decimals)
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

        if log_func:
            log_func(
                f"  Pair {left_sheet} vs {right_sheet} | {pretty}: raw rows={len(df)}"
            )

        for _, row in df.iterrows():
            cont = str(row.get("Contingency", "") or "")
            issue = str(row.get("ResultingIssue", "") or "")
            left_pct = row.get("LeftPct", math.nan)
            right_pct = row.get("RightPct", math.nan)
            delta_pct = row.get("DeltaPct", math.nan)

            values = []
            if not _is_nan(left_pct):
                values.append(float(left_pct))
            if not _is_nan(right_pct):
                values.append(float(right_pct))

            if not values:
                # truly empty
                continue

            max_val = max(values)
            if max_val < threshold:
                # below threshold on both sides
                continue

            # Decide DeltaDisplay text
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

    # To keep ordering nice: sort by CaseType groups, and within each group,
    # by max(LeftPct, RightPct) descending
    if not df_all.empty:
        sort_vals = df_all[["LeftPct", "RightPct"]].max(axis=1)
        df_all["_SortKey"] = sort_vals
        df_all = df_all.sort_values(
            by=["CaseType", "_SortKey"], ascending=[True, False], na_position="last"
        ).drop(columns=["_SortKey"])

    return df_all


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
    (No frozen panes – normal scrolling.)
    """
    widths = {
        2: 45,  # B: Contingency
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


def _write_formatted_pair_sheet(
    wb: Workbook,
    ws_name: str,
    df_pair: pd.DataFrame,
):
    """
    Create one sheet in the batch workbook using the same style/structure
    as the Combined_ViolationCTG_Comparison.xlsx sheet:

        [Title row]   merged B:F, e.g. 'ACCA'
        [Header row]  B: Contingency Events
                      C: Resulting Issue
                      D: Left %
                      E: Right %
                      F: Δ% (Right - Left) / Status
        [Data rows]   grouped by case type, with a blank row between blocks.
    """
    ws = wb.create_sheet(title=ws_name)
    _apply_table_styles(ws)

    # Enable Excel outline symbols for expand/collapse groups
    try:
        ws.sheet_view.showOutlineSymbols = True
        ws.sheet_properties.outlinePr.summaryBelow = True
    except Exception:
        pass


    if df_pair.empty:
        ws.cell(row=2, column=2).value = "No rows above threshold."
        return

    current_row = 2

    # Group by pretty CaseType name (ACCA LongTerm / ACCA / DCwAC)
    for case_type_pretty in ["ACCA LongTerm", "ACCA", "DCwAC"]:
        sub = df_pair[df_pair["CaseType"] == case_type_pretty].copy()
        if sub.empty:
            continue

        # Title row (merged B:F)
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=6)
        title_cell = ws.cell(row=current_row, column=2)
        title_cell.value = case_type_pretty
        title_cell.fill = TITLE_FILL
        title_cell.font = TITLE_FONT
        title_cell.alignment = CELL_ALIGN_CENTER
        current_row += 1

        # Header row
        headers = [
            "Contingency Events",
            "Resulting Issue",
            "Left %",
            "Right %",
            "Δ% (Right - Left) / Status",
        ]
        for col_offset, header in enumerate(headers):
            cell = ws.cell(row=current_row, column=2 + col_offset)
            cell.value = header
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = CELL_ALIGN_CENTER
            cell.border = THIN_BORDER
        current_row += 1


# Data rows (grouped by Resulting Issue with Excel outline)
#
# Parent row = worst contingency for a given Resulting Issue.
# Child rows (other contingencies) are written underneath and hidden by default.
if "ResultingIssue" in sub.columns and len(sub) > 0:
    work = sub.copy()

    # Safe numeric versions for ordering
    work["_L"] = pd.to_numeric(work["LeftPct"], errors="coerce")
    work["_R"] = pd.to_numeric(work["RightPct"], errors="coerce")
    work["_RowMax"] = work[["_L", "_R"]].max(axis=1)

    grouped = []
    for issue_key, g in work.groupby("ResultingIssue", dropna=False):
        g_worst = g["_RowMax"].max()
        grouped.append((issue_key, g, g_worst))

    # Sort groups by worst loading desc (NaNs last)
    grouped.sort(key=lambda t: (-(t[2]) if pd.notna(t[2]) else float("inf")))

    for issue_key, g, _g_worst in grouped:
        # Pick parent row (prefer highest RightPct if any, else LeftPct)
        if g["_R"].notna().any():
            parent_idx = g["_R"].idxmax()
            g_sorted = g.drop(index=[parent_idx]).sort_values(by="_R", ascending=False, na_position="last")
        else:
            parent_idx = g["_L"].idxmax()
            g_sorted = g.drop(index=[parent_idx]).sort_values(by="_L", ascending=False, na_position="last")

        parent = g.loc[parent_idx]

        # ---- Parent row ----
        p_cont = str(parent.get("Contingency", "") or "")
        p_issue = str(parent.get("ResultingIssue", "") or "")
        p_left = parent.get("LeftPct", None)
        p_right = parent.get("RightPct", None)
        p_delta = parent.get("DeltaDisplay", "")

        values = [p_cont, p_issue, p_left, p_right, p_delta]
        for col_offset, val in enumerate(values):
            cell = ws.cell(row=current_row, column=2 + col_offset)
            cell.value = val
            cell.border = THIN_BORDER

            if col_offset in (0, 1):  # text columns
                cell.alignment = CELL_ALIGN_WRAP
            else:
                cell.alignment = Alignment(horizontal="right", vertical="top")

            # number formatting for percentages
            if col_offset in (2, 3) and isinstance(val, (float, int)):
                cell.number_format = "0.00"

        # Mark parent as collapsed if it has children
        if len(g) > 1:
            ws.row_dimensions[current_row].collapsed = True

        current_row += 1

        # ---- Child rows (hidden) ----
        for _, row in g_sorted.iterrows():
            cont = str(row.get("Contingency", "") or "")
            # leave issue blank for child rows (cleaner)
            issue = ""
            left_pct = row.get("LeftPct", None)
            right_pct = row.get("RightPct", None)
            delta = row.get("DeltaDisplay", "")

            values = [cont, issue, left_pct, right_pct, delta]
            for col_offset, val in enumerate(values):
                cell = ws.cell(row=current_row, column=2 + col_offset)
                cell.value = val
                cell.border = THIN_BORDER

                if col_offset in (0, 1):  # text columns
                    cell.alignment = CELL_ALIGN_WRAP
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="top")

                if col_offset in (2, 3) and isinstance(val, (float, int)):
                    cell.number_format = "0.00"

            # Outline / hidden for Excel collapse
            ws.row_dimensions[current_row].outlineLevel = 1
            ws.row_dimensions[current_row].hidden = True

            current_row += 1
else:
    # Fallback: flat rows
    for _, row in sub.iterrows():
        cont = str(row.get("Contingency", "") or "")
        issue = str(row.get("ResultingIssue", "") or "")
        left_pct = row.get("LeftPct", None)
        right_pct = row.get("RightPct", None)
        delta = row.get("DeltaDisplay", "")

        values = [cont, issue, left_pct, right_pct, delta]
        for col_offset, val in enumerate(values):
            cell = ws.cell(row=current_row, column=2 + col_offset)
            cell.value = val
            cell.border = THIN_BORDER

            if col_offset in (0, 1):  # text columns
                cell.alignment = CELL_ALIGN_WRAP
            else:
                cell.alignment = Alignment(horizontal="right", vertical="top")

            # number formatting for percentages
            if col_offset in (2, 3) and isinstance(val, (float, int)):
                cell.number_format = "0.00"

        current_row += 1
        # Blank row between blocks
        current_row += 1


# ---------------------------------------------------------------------------
# Build full batch comparison workbook (pretty formatting)
# ---------------------------------------------------------------------------


def build_batch_comparison_workbook(
    src_workbook: str,
    pairs: Sequence[Tuple[str, str]],
    threshold: float,
    output_path: str,
    log_func=None,
) -> str:
    """
    Build a brand-new .xlsx workbook with one sheet per (left_sheet, right_sheet) pair,
    using the same blue-block style as Combined_ViolationCTG_Comparison.xlsx.

    Each sheet is grouped into ACCA LongTerm / ACCA / DCwAC blocks, with columns:
      Contingency Events | Resulting Issue | Left % | Right % | Δ% (Right - Left) / Status

    threshold is the same loading threshold used by the GUI.

    Returns output_path.
    """
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to build the batch workbook.")

    if log_func:
        log_func(
            f"\n=== Building batch comparison workbook ===\n"
            f"Source: {src_workbook}\n"
            f"Output: {output_path}\n"
            f"Threshold: {threshold:.2f}%\n"
        )

    if not pairs:
        raise ValueError("No comparison pairs provided.")

    wb = Workbook()
    # Remove the default sheet; we'll create our own
    default_sheet = wb.active
    wb.remove(default_sheet)

    used_names: set[str] = set()

    for idx, (left_sheet, right_sheet) in enumerate(pairs, start=1):
        if log_func:
            log_func(f"Processing pair {idx}: '{left_sheet}' vs '{right_sheet}'...")

        df_pair = build_pair_comparison_df(
            src_workbook, left_sheet, right_sheet, threshold, log_func=log_func
        )

        if df_pair.empty:
            # Create a tiny DF with a message so the sheet isn't totally blank
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

        # Sheet name: prefix with index so it always matches the queue/log order
        # Example: "01_Base Case_vs_Breaker Test 1"
        base_name = f"{idx:02d}_{left_sheet}_vs_{right_sheet}"
        base_name = _sanitize_sheet_name(base_name)

        name = base_name
        counter = 2
        while name in used_names:
            # If a duplicate somehow occurs, append a small numeric suffix
            suffix = f"_{counter}"
            name = _sanitize_sheet_name(base_name[: (31 - len(suffix))] + suffix)
            counter += 1

        used_names.add(name)
        _write_formatted_pair_sheet(wb, name, df_pair)

    wb.save(output_path)

    if log_func:
        log_func(f"Batch comparison workbook written to:\n{output_path}")

    return output_path
