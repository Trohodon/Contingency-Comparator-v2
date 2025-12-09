# core/comparator.py

import os
from typing import Iterable, List, Dict, Optional

import pandas as pd

try:
    from openpyxl import load_workbook
    from openpyxl.worksheet.worksheet import Worksheet
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# These labels match what we used when building the formatted workbook
CANONICAL_CASE_TYPES = {
    "ACCA LongTerm": "ACCA_LongTerm",
    "ACCA": "ACCA_P1,2,4,7",
    "DCwAC": "DCwACver_P1-7",
}


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

    Sheet structure we produced in comparison_builder:
      - Title row: merged B:E, text is pretty case name ("ACCA LongTerm", "ACCA", "DCwAC")
      - Next row: headers in B..E
      - Following rows: data until blank row; then repeat block.
    """
    records: List[Dict] = []

    max_row = ws.max_row or 1
    row_idx = 1

    while row_idx <= max_row:
        title_cell = ws.cell(row=row_idx, column=2)  # B
        title_val = title_cell.value

        # Identify a title row by having some text in B
        if isinstance(title_val, str) and title_val.strip():
            pretty_name = title_val.strip()
            case_type = CANONICAL_CASE_TYPES.get(pretty_name, pretty_name)

            # Header row is next
            header_row = row_idx + 1
            # Data starts on the row after header
            data_row = header_row + 1

            # Read data rows until we hit a blank B..E
            r = data_row
            while r <= max_row:
                b = ws.cell(row=r, column=2).value
                c = ws.cell(row=r, column=3).value
                d = ws.cell(row=r, column=4).value
                e = ws.cell(row=r, column=5).value

                # Stop when whole row is empty
                if (b is None or str(b).strip() == "") and \
                   (c is None or str(c).strip() == "") and \
                   (d is None or str(d).strip() == "") and \
                   (e is None or str(e).strip() == ""):
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
            f"  Parsed {len(df)} rows from sheet '{ws.title}'. "
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

    mode:
      - "all"    : show all matched rows (subject to abs(Δ%) >= threshold) plus unmatched
      - "worse" : show only rows where New% > Base% by at least threshold, plus unmatched
      - "better": show only rows where New% < Base% by at least threshold, plus unmatched
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

    if log_func:
        log_func(
            f"Matched/merged rows: {len(merged)}; rows after filters: {len(filtered)}"
        )

    # Write result into the same workbook as a new sheet
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl is required to write comparison sheet.")

    from openpyxl import load_workbook  # already imported above

    wb = load_workbook(workbook_path)
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