# core/comparison_builder.py
import os
import pandas as pd

from .case_finder import TARGET_PATTERNS

# Try to import openpyxl for nice formatting / grouping
try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False


# Pretty display names for the three case types
PRETTY_CASE_NAMES = {
    "ACCA_LongTerm": "ACCA LongTerm",
    "ACCA_P1,2,4,7": "ACCA",
    "DCwACver_P1-7": "DCwAC",
}


def _safe_filename_component(name: str) -> str:
    """Make a Windows-safe filename chunk."""
    if not name:
        return "Output"
    s = str(name).strip().replace(" ", "_")
    illegal = '<>:"/\\|?*'
    for ch in illegal:
        s = s.replace(ch, "")
    return s or "Output"


def _compute_workbook_path(root_folder: str, include_branch_mva: bool, include_bus_lv: bool) -> str:
    """
    Naming rules (per v2 request):
      - Branch MVA only -> {main folder name}_BranchMVA_CTG_Comparison.xlsx
      - Bus Low Volts only -> {main folder name}_BusLowVolts_CTG_Comarison.xlsx  (note spelling per request)
      - Both -> {main folder name}_CombinedCTG_Comparison.xlsx
    """
    root_name = _safe_filename_component(os.path.basename(os.path.normpath(root_folder)))

    if include_branch_mva and include_bus_lv:
        suffix = "CombinedCTG_Comparison"
    elif include_branch_mva:
        suffix = "BranchMVA_CTG_Comparison"
    elif include_bus_lv:
        suffix = "BusLowVolts_CTG_Comarison"
    else:
        suffix = "CTG_Comparison"

    return os.path.join(root_folder, f"{root_name}_{suffix}.xlsx")


def _to_float_series(series: pd.Series) -> pd.Series:
    """Convert a LimViolPct-like series to float safely."""
    if series is None:
        return pd.Series(dtype="float64")
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)
    cleaned = series.astype(str).str.replace("%", "", regex=False).str.strip()
    return pd.to_numeric(cleaned, errors="coerce")


def _build_simple_workbook(
    root_folder: str,
    folder_to_case_csvs: dict,
    include_branch_mva: bool,
    include_bus_lv: bool,
    log_func=None,
) -> str | None:
    """
    Fallback: simple one-sheet-per-scenario workbook, no fancy formatting,
    and no outline dropdown grouping.
    """
    if not folder_to_case_csvs:
        if log_func:
            log_func("No data to build combined workbook.")
        return None

    workbook_path = _compute_workbook_path(root_folder, include_branch_mva, include_bus_lv)

    if log_func:
        log_func(f"\nBuilding SIMPLE combined workbook:\n  {workbook_path}")

    try:
        writer = pd.ExcelWriter(workbook_path)
    except Exception as e:
        if log_func:
            log_func(f"ERROR: Could not create ExcelWriter: {e}")
        return None

    try:
        with writer as xls_writer:
            for folder_name, case_map in folder_to_case_csvs.items():
                dfs = []
                for label in TARGET_PATTERNS:
                    csv_path = case_map.get(label)
                    if not csv_path:
                        continue
                    try:
                        df = pd.read_csv(csv_path)
                        df.insert(0, "CaseType", label)
                        dfs.append(df)
                    except Exception as e:
                        if log_func:
                            log_func(f"  [{folder_name}] WARNING: Failed to read {csv_path}: {e}")

                if not dfs:
                    continue

                combined = pd.concat(dfs, ignore_index=True)
                sheet_name = (folder_name or "Sheet").strip()[:31] or "Sheet"
                combined.to_excel(xls_writer, sheet_name=sheet_name, index=False)

    except Exception as e:
        if log_func:
            log_func(f"ERROR while building simple workbook: {e}")
        return None

    if log_func:
        log_func("Simple combined workbook build complete.")
    return workbook_path


def build_workbook(
    root_folder: str,
    folder_to_case_csvs: dict,
    group_details: bool = True,
    include_branch_mva: bool = True,
    include_bus_lv: bool = False,
    log_func=None,
) -> str | None:
    """
    Build a combined Excel workbook with one sheet per subfolder.

    Parameters:
      - root_folder: the selected "main" folder.
      - folder_to_case_csvs: dict of {scenario_name: {case_label: filtered_csv_path}}
      - group_details:
            True  -> expandable issue view using Excel outline (+/-):
                     show max row per LimViolID, hide the rest underneath.
            False -> dump all rows as-is (no grouping).
      - include_branch_mva / include_bus_lv:
            Only used for output filename (filters are already applied upstream).

    v2 behavior:
      - Summary rows (main Resulting Issues) are ordered by highest Percent Loading (worst first).
      - All output is shifted down by 1 row (blank Row 1).
    """
    if not folder_to_case_csvs:
        if log_func:
            log_func("No data to build combined workbook.")
        return None

    # If openpyxl is not available, build a simple workbook without formatting
    if not OPENPYXL_AVAILABLE:
        if log_func:
            log_func("openpyxl not available; building simple combined workbook without special formatting.")
        return _build_simple_workbook(root_folder, folder_to_case_csvs, include_branch_mva, include_bus_lv, log_func)

    # ---------------------------
    # Build scenario DataFrames
    # ---------------------------
    scenario_data = {}  # folder_name -> combined DataFrame

    for folder_name, case_map in folder_to_case_csvs.items():
        dfs = []
        for label in TARGET_PATTERNS:
            csv_path = case_map.get(label)
            if not csv_path:
                continue
            try:
                df = pd.read_csv(csv_path)
                df.insert(0, "CaseType", label)
                dfs.append(df)
            except Exception as e:
                if log_func:
                    log_func(f"  [{folder_name}] WARNING: Failed to read {csv_path}: {e}")

        if dfs:
            scenario_data[folder_name] = pd.concat(dfs, ignore_index=True)

    if not scenario_data:
        if log_func:
            log_func("No scenario data to write into workbook.")
        return None

    # ---------------------------
    # Create formatted workbook
    # ---------------------------
    workbook_path = _compute_workbook_path(root_folder, include_branch_mva, include_bus_lv)

    if log_func:
        log_func(f"\nBuilding FORMATTED combined workbook:\n  {workbook_path}")
        log_func(f"Expandable dropdown grouping is {'ON' if group_details else 'OFF'}.")
        log_func("Sorting Resulting Issues by highest Percent Loading (worst first).")
        log_func("Shifting output down by 1 row (blank Row 1).")

    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    # Styles
    title_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    title_font = Font(color="FFFFFF", bold=True, size=12)

    header_fill = PatternFill(fill_type="solid", fgColor="305496")
    header_font = Font(color="FFFFFF", bold=True)

    data_font = Font(color="000000")
    data_bold_font = Font(color="000000", bold=True)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    required_cols = ["CTGLabel", "LimViolValue", "LimViolPct"]

    for folder_name, df in scenario_data.items():
        sheet_name = (folder_name or "Sheet").strip()[:31] or "Sheet"
        ws = wb.create_sheet(title=sheet_name)

        # Excel outline behavior: summary rows ABOVE details (so dropdown is on the max row)
        ws.sheet_properties.outlinePr.summaryBelow = False

        # Set column widths – contiguous columns B to E
        ws.column_dimensions["B"].width = 55  # Contingency Events
        ws.column_dimensions["C"].width = 55  # Resulting Issue
        ws.column_dimensions["D"].width = 18  # Contingency Value (MVA)
        ws.column_dimensions["E"].width = 18  # Percent Loading

        # Start on row 2 to leave a blank row 1
        current_row = 2

        for label in TARGET_PATTERNS:
            block_df = df[df["CaseType"] == label].copy()
            if block_df.empty:
                continue

            pretty_name = PRETTY_CASE_NAMES.get(label, label)

            # ===== Title row =====
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=5)
            c = ws.cell(row=current_row, column=2)
            c.value = pretty_name
            c.fill = title_fill
            c.font = title_font
            c.alignment = center
            for col in range(2, 6):
                ws.cell(row=current_row, column=col).border = thin_border
            current_row += 1

            # ===== Header row =====
            headers = [
                ("B", "Contingency Events"),
                ("C", "Resulting Issue"),
                ("D", "Contingency Value (MVA)"),
                ("E", "Percent Loading"),
            ]
            for col_letter, text in headers:
                col_idx = ord(col_letter) - ord("A") + 1
                hc = ws.cell(row=current_row, column=col_idx)
                hc.value = text
                hc.fill = header_fill
                hc.font = header_font
                hc.alignment = center
                hc.border = thin_border
            current_row += 1

            # Validate columns
            for col in required_cols:
                if col not in block_df.columns and log_func:
                    log_func(f"  [{folder_name} / {label}] WARNING: column '{col}' missing.")

            has_limviolid = "LimViolID" in block_df.columns

            # Grouping mode (expandable issue view)
            if group_details and has_limviolid:
                # Create numeric percent column for ordering
                if "LimViolPct" in block_df.columns:
                    block_df["_pct_num"] = _to_float_series(block_df["LimViolPct"])
                else:
                    block_df["_pct_num"] = pd.Series([float("nan")] * len(block_df), index=block_df.index)

                # Sort within each LimViolID so the max row comes first
                sort_cols = ["LimViolID", "_pct_num"]
                asc = [True, False]
                if "CTGLabel" in block_df.columns:
                    sort_cols.append("CTGLabel")
                    asc.append(True)

                block_df = block_df.sort_values(
                    by=sort_cols,
                    ascending=asc,
                    na_position="last",
                    kind="mergesort",
                )

                # NEW: order the groups themselves by their max % (worst first)
                group_max = (
                    block_df.groupby("LimViolID", dropna=False)["_pct_num"]
                    .max()
                    .reset_index()
                    .rename(columns={"_pct_num": "_group_max_pct"})
                )
                group_max["_lim_str"] = group_max["LimViolID"].astype(str)
                group_max = group_max.sort_values(
                    by=["_group_max_pct", "_lim_str"],
                    ascending=[False, True],
                    na_position="last",
                    kind="mergesort",
                )

                ordered_limviolid_values = list(group_max["LimViolID"])

                # Write each group (summary row + collapsed details)
                for limviolid in ordered_limviolid_values:
                    g = block_df[block_df["LimViolID"].eq(limviolid)].copy()
                    if g.empty:
                        continue

                    rows = list(g.itertuples(index=False))
                    if not rows:
                        continue

                    # Summary row (max within group)
                    r0 = rows[0]

                    cB = ws.cell(row=current_row, column=2)
                    cB.value = getattr(r0, "CTGLabel", "")
                    cB.font = data_bold_font
                    cB.alignment = left_align
                    cB.border = thin_border

                    cC = ws.cell(row=current_row, column=3)
                    cC.value = getattr(r0, "LimViolID", "")
                    cC.font = data_bold_font
                    cC.alignment = left_align
                    cC.border = thin_border

                    cD = ws.cell(row=current_row, column=4)
                    cD.value = getattr(r0, "LimViolValue", "")
                    cD.font = data_bold_font
                    cD.alignment = center
                    cD.border = thin_border

                    cE = ws.cell(row=current_row, column=5)
                    cE.value = getattr(r0, "LimViolPct", "")
                    cE.font = data_bold_font
                    cE.alignment = center
                    cE.border = thin_border

                    summary_row = current_row
                    current_row += 1

                    # Detail rows (collapsed)
                    detail_start = None
                    detail_end = None

                    for r in rows[1:]:
                        if detail_start is None:
                            detail_start = current_row

                        cB = ws.cell(row=current_row, column=2)
                        cB.value = getattr(r, "CTGLabel", "")
                        cB.font = data_font
                        cB.alignment = left_align
                        cB.border = thin_border

                        cC = ws.cell(row=current_row, column=3)
                        # Blank so it's visually "under" the same Resulting Issue
                        cC.value = ""
                        cC.font = data_font
                        cC.alignment = left_align
                        cC.border = thin_border

                        cD = ws.cell(row=current_row, column=4)
                        cD.value = getattr(r, "LimViolValue", "")
                        cD.font = data_font
                        cD.alignment = center
                        cD.border = thin_border

                        cE = ws.cell(row=current_row, column=5)
                        cE.value = getattr(r, "LimViolPct", "")
                        cE.font = data_font
                        cE.alignment = center
                        cE.border = thin_border

                        detail_end = current_row
                        current_row += 1

                    # Apply outline grouping (hide details by default)
                    if detail_start is not None and detail_end is not None:
                        ws.row_dimensions.group(
                            detail_start,
                            detail_end,
                            outline_level=1,
                            hidden=True,
                        )
                        ws.row_dimensions[summary_row].collapsed = True

            else:
                # Non-grouped mode: just dump the rows
                for _, row in block_df.iterrows():
                    c = ws.cell(row=current_row, column=2)
                    c.value = row.get("CTGLabel", "")
                    c.font = data_font
                    c.alignment = left_align
                    c.border = thin_border

                    c = ws.cell(row=current_row, column=3)
                    c.value = row.get("LimViolID", "") if has_limviolid else ""
                    c.font = data_font
                    c.alignment = left_align
                    c.border = thin_border

                    c = ws.cell(row=current_row, column=4)
                    c.value = row.get("LimViolValue", "")
                    c.font = data_font
                    c.alignment = center
                    c.border = thin_border

                    c = ws.cell(row=current_row, column=5)
                    c.value = row.get("LimViolPct", "")
                    c.font = data_font
                    c.alignment = center
                    c.border = thin_border

                    current_row += 1

            # One blank row between blocks
            current_row += 1

    # Save workbook
    try:
        wb.save(workbook_path)
    except Exception as e:
        if log_func:
            log_func(f"ERROR saving formatted workbook: {e}")
        return None

    if log_func:
        log_func("Formatted combined workbook build complete.")
    return workbook_path
