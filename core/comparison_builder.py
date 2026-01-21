import os
from typing import Dict, Optional, List, Tuple

import pandas as pd

from .case_finder import TARGET_PATTERNS

# Optional: nicer formatting if openpyxl is installed
try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False


# Pretty display names for the three case types
CASE_TYPE_DISPLAY = {
    "ACCA_LongTerm": "ACCA LongTerm",
    "ACCA_P1,2,4,7": "ACCA",
    "DCwACver_P1-7": "DCwAC",
}


def _safe_filename_component(name: str) -> str:
    """
    Make a string safe to use inside a Windows filename.
    """
    if not name:
        return ""
    bad = '<>:"/\\|?*'
    out = "".join("_" if c in bad else c for c in name)
    out = out.strip().strip(".")
    return out


def _output_workbook_path(
    root_folder: str,
    include_branch_mva: bool,
    include_bus_low_volts: bool,
) -> str:
    """
    Naming rules requested:

    - Branch MVA only  -> {main folder name}_BranchMVA_CTG_Comparison.xlsx
    - Bus Low Volts only -> {main folder name}_BusLowVolts_CTG_Comarison.xlsx  (spelling per request)
    - Both -> {main folder name}_CombinedCTG_Comparison.xlsx

    If neither is checked, we fall back to: {main folder name}_CTG_Comparison.xlsx
    """
    base = _safe_filename_component(os.path.basename(os.path.normpath(root_folder))) or "CTG"

    if include_branch_mva and include_bus_low_volts:
        filename = f"{base}_CombinedCTG_Comparison.xlsx"
    elif include_branch_mva:
        filename = f"{base}_BranchMVA_CTG_Comparison.xlsx"
    elif include_bus_low_volts:
        filename = f"{base}_BusLowVolts_CTG_Comarison.xlsx"
    else:
        filename = f"{base}_CTG_Comparison.xlsx"

    return os.path.join(root_folder, filename)


def _pick_first_existing(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def _coerce_float(x):
    try:
        if pd.isna(x):
            return float("nan")
        return float(x)
    except Exception:
        return float("nan")


def _build_simple_workbook(
    root_folder: str,
    folder_to_case_csvs: Dict[str, Dict[str, str]],
    include_branch_mva: bool,
    include_bus_low_volts: bool,
    log_func=None,
) -> Optional[str]:
    """
    Fallback if openpyxl isn't available:
      - Creates one sheet per scenario folder
      - Writes raw CSV rows (no styling)
    """
    if not folder_to_case_csvs:
        if log_func:
            log_func("No data to build combined workbook.")
        return None

    workbook_path = _output_workbook_path(
        root_folder, include_branch_mva=include_branch_mva, include_bus_low_volts=include_bus_low_volts
    )

    try:
        with pd.ExcelWriter(workbook_path, engine="openpyxl") as writer:
            for folder_name, case_map in folder_to_case_csvs.items():
                combined = []
                for label in TARGET_PATTERNS:
                    csv_path = case_map.get(label)
                    if not csv_path:
                        continue
                    try:
                        df = pd.read_csv(csv_path)
                        df.insert(0, "CaseType", label)
                        combined.append(df)
                    except Exception as e:
                        if log_func:
                            log_func(f"WARNING: Could not read CSV [{csv_path}]: {e}")
                if combined:
                    out = pd.concat(combined, ignore_index=True)
                    sheet = (folder_name or "Scenario")[:31]
                    out.to_excel(writer, sheet_name=sheet, index=False)

        if log_func:
            log_func(f"Workbook created: {workbook_path}")
        return workbook_path
    except Exception as e:
        if log_func:
            log_func(f"ERROR: Failed to create workbook: {e}")
        return None


def build_workbook(
    root_folder: str,
    folder_to_case_csvs: Dict[str, Dict[str, str]],
    include_branch_mva: bool = True,
    include_bus_low_volts: bool = False,
    group_details: bool = True,
    log_func=None,
) -> Optional[str]:
    """
    Build a combined comparison workbook.

    Parameters
    ----------
    root_folder:
        Main root folder (output workbook goes here)
    folder_to_case_csvs:
        {scenario_folder_name: {case_type_label: filtered_csv_path}}
    include_branch_mva / include_bus_low_volts:
        Used only for output filename rule.
    group_details:
        If True, create Excel +/- row groups:
          - show the max percent row per Resulting Issue
          - collapse the other contingency rows under it
        If False, just list all rows.
    """
    if not folder_to_case_csvs:
        if log_func:
            log_func("No data to build combined workbook.")
        return None

    if not OPENPYXL_AVAILABLE:
        return _build_simple_workbook(
            root_folder,
            folder_to_case_csvs,
            include_branch_mva=include_branch_mva,
            include_bus_low_volts=include_bus_low_volts,
            log_func=log_func,
        )

    workbook_path = _output_workbook_path(
        root_folder, include_branch_mva=include_branch_mva, include_bus_low_volts=include_bus_low_volts
    )

    # Styles
    header_fill = PatternFill("solid", fgColor="1F4E79")  # dark blue
    header_font = Font(color="FFFFFF", bold=True)
    title_fill = PatternFill("solid", fgColor="0E2A47")   # darker blue
    title_font = Font(color="FFFFFF", bold=True, size=12)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    thin = Side(style="thin", color="9E9E9E")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def style_range(ws, row: int, col_start: int, col_end: int, fill=None, font=None, align=None):
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=c)
            if fill:
                cell.fill = fill
            if font:
                cell.font = font
            if align:
                cell.alignment = align
            cell.border = border

    def fmt_number_cell(cell, value):
        cell.value = value
        cell.number_format = "0.000000"
        cell.alignment = right
        cell.border = border

    wb = Workbook()
    # remove default sheet
    if wb.worksheets:
        wb.remove(wb.worksheets[0])

    # We leave row 1 blank to "shift down by 1 row" per your note
    START_ROW = 2

    for folder_name, case_map in folder_to_case_csvs.items():
        sheet_name = (folder_name or "Scenario")[:31]
        ws = wb.create_sheet(title=sheet_name)

        # Column widths to match your screenshot style (B-E)
        ws.column_dimensions["A"].width = 2    # spacer
        ws.column_dimensions["B"].width = 60   # Contingency Events
        ws.column_dimensions["C"].width = 72   # Resulting Issue
        ws.column_dimensions["D"].width = 18   # Value
        ws.column_dimensions["E"].width = 18   # Percent

        current_row = START_ROW

        for case_type in TARGET_PATTERNS:
            csv_path = case_map.get(case_type)
            if not csv_path or not os.path.exists(csv_path):
                continue

            # Title row
            title = CASE_TYPE_DISPLAY.get(case_type, case_type)
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
            tcell = ws.cell(row=current_row, column=2, value=title)
            tcell.fill = title_fill
            tcell.font = title_font
            tcell.alignment = center
            style_range(ws, current_row, 2, 5, fill=title_fill, font=title_font, align=center)
            ws.row_dimensions[current_row].height = 22
            current_row += 1

            # Header row
            headers = ["Contingency Events", "Resulting Issue", "Contingency Value\n(MVA)", "Percent Loading"]
            for j, h in enumerate(headers, start=2):
                cell = ws.cell(row=current_row, column=j, value=h)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center
                cell.border = border
            ws.row_dimensions[current_row].height = 36
            current_row += 1

            try:
                df = pd.read_csv(csv_path)
            except Exception as e:
                if log_func:
                    log_func(f"WARNING: Could not read CSV for [{folder_name}] [{case_type}]: {e}")
                current_row += 2
                continue

            # Identify columns
            cont_col = _pick_first_existing(df, ["CTGLabel", "Contingency Events", "Contingency", "CTG"])
            issue_col = _pick_first_existing(df, ["LimViolID", "Resulting Issue", "LimViolID:1", "Violation"])
            val_col = _pick_first_existing(df, ["LimViolValue:1", "LimViolValue", "Value", "MVA", "CTGValue"])
            pct_col = _pick_first_existing(df, ["LimViolPct", "Percent Loading", "Percent", "LimViolPct:1"])

            # If required columns aren't present, just dump what we can
            if issue_col is None or pct_col is None:
                if log_func:
                    log_func(
                        f"WARNING: Expected columns not found in [{csv_path}]. "
                        f"Found columns: {list(df.columns)}"
                    )
                dump_cols = [c for c in [cont_col, issue_col, val_col, pct_col] if c] or list(df.columns)[:4]
                for _, r in df[dump_cols].iterrows():
                    vals = list(r.values)
                    while len(vals) < 4:
                        vals.append("")
                    ws.cell(row=current_row, column=2, value=str(vals[0]) if len(vals) > 0 else "").alignment = left_wrap
                    ws.cell(row=current_row, column=3, value=str(vals[1]) if len(vals) > 1 else "").alignment = left_wrap
                    ws.cell(row=current_row, column=4, value=vals[2] if len(vals) > 2 else "")
                    ws.cell(row=current_row, column=5, value=vals[3] if len(vals) > 3 else "")
                    style_range(ws, current_row, 2, 5, align=left_wrap)
                    current_row += 1

                current_row += 2
                continue

            # Normalize columns for sort/group
            d = df.copy()
            if cont_col is None:
                d["_cont"] = ""
                cont_col = "_cont"
            if val_col is None:
                d["_val"] = float("nan")
                val_col = "_val"

            d["_pct"] = d[pct_col].apply(_coerce_float)

            # Group by "Resulting Issue" and sort groups by max percent (descending)
            # Within each issue group, sort rows by percent (descending)
            import math

            groups: List[Tuple[str, pd.DataFrame, float]] = []
            for issue, g in d.groupby(issue_col, dropna=False):
                g2 = g.sort_values("_pct", ascending=False, na_position="last")
                max_pct = _coerce_float(g2["_pct"].iloc[0]) if len(g2) else float("nan")
                groups.append((str(issue) if not pd.isna(issue) else "", g2, max_pct))

            groups.sort(key=lambda t: (-(t[2] if not math.isnan(t[2]) else -1e18)))

            # Write rows
            for issue, g2, _max_pct in groups:
                # first row = max percent (always visible)
                top = g2.iloc[0]
                ws.cell(row=current_row, column=2, value=str(top[cont_col])).alignment = left_wrap
                ws.cell(row=current_row, column=3, value=str(top[issue_col])).alignment = left_wrap
                fmt_number_cell(ws.cell(row=current_row, column=4), _coerce_float(top[val_col]))
                fmt_number_cell(ws.cell(row=current_row, column=5), _coerce_float(top[pct_col]))
                style_range(ws, current_row, 2, 3, align=left_wrap)
                current_row += 1

                # detail rows (collapsed under +/-)
                if group_details and len(g2) > 1:
                    for idx in range(1, len(g2)):
                        row = g2.iloc[idx]
                        ws.cell(row=current_row, column=2, value=str(row[cont_col])).alignment = left_wrap
                        ws.cell(row=current_row, column=3, value=str(row[issue_col])).alignment = left_wrap
                        fmt_number_cell(ws.cell(row=current_row, column=4), _coerce_float(row[val_col]))
                        fmt_number_cell(ws.cell(row=current_row, column=5), _coerce_float(row[pct_col]))
                        style_range(ws, current_row, 2, 3, align=left_wrap)

                        ws.row_dimensions[current_row].outlineLevel = 1
                        ws.row_dimensions[current_row].hidden = True
                        current_row += 1

                    ws.sheet_properties.outlinePr.summaryBelow = True
                    ws.sheet_properties.outlinePr.applyStyles = True

            # Spacing between blocks
            current_row += 2

        # Freeze panes under the first block header if you scroll
        ws.freeze_panes = ws["B3"]

    try:
        wb.save(workbook_path)
        if log_func:
            log_func(f"Workbook created: {workbook_path}")
        return workbook_path
    except Exception as e:
        if log_func:
            log_func(f"ERROR: Failed to save workbook: {e}")
        return None