import os
import pandas as pd

from .case_finder import TARGET_PATTERNS

Try to import openpyxl for nice formatting

try:
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
OPENPYXL_AVAILABLE = True
except ImportError:
OPENPYXL_AVAILABLE = False

Pretty display names for the three case types

PRETTY_CASE_NAMES = {
"ACCA_LongTerm": "ACCA LongTerm",
"ACCA_P1,2,4,7": "ACCA",
"DCwACver_P1-7": "DCwAC",
}

def _build_simple_workbook(root_folder, folder_to_case_csvs, log_func=None):
"""
Fallback: simple one-sheet-per-scenario workbook, no fancy formatting.
"""
if not folder_to_case_csvs:
if log_func:
log_func("No data to build combined workbook.")
return None

workbook_path = os.path.join(  
    root_folder, "Combined_ViolationCTG_Comparison.xlsx"  
)  

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
                        log_func(  
                            f"  [{folder_name}] WARNING: Failed to read {csv_path}: {e}"  
                        )  

            if not dfs:  
                continue  

            combined = pd.concat(dfs, ignore_index=True)  

            sheet_name = (folder_name or "Sheet").strip()[:31]  
            if not sheet_name:  
                sheet_name = "Sheet"  

            combined.to_excel(xls_writer, sheet_name=sheet_name, index=False)  

except Exception as e:  
    if log_func:  
        log_func(f"ERROR while building simple workbook: {e}")  
    return None  

if log_func:  
    log_func("Simple combined workbook build complete.")  
return workbook_path

def build_workbook(root_folder, folder_to_case_csvs, log_func=None):
"""
Build a combined Excel workbook with one sheet per subfolder, formatted
like the manual comparison sheet.

If openpyxl is not available, falls back to a simple table layout.  
"""  
if not folder_to_case_csvs:  
    if log_func:  
        log_func("No data to build combined workbook.")  
    return None  

if not OPENPYXL_AVAILABLE:  
    if log_func:  
        log_func(  
            "openpyxl not available; building simple combined workbook "  
            "without special formatting."  
        )  
    return _build_simple_workbook(root_folder, folder_to_case_csvs, log_func)  

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
                log_func(  
                    f"  [{folder_name}] WARNING: Failed to read {csv_path}: {e}"  
                )  

    if dfs:  
        scenario_data[folder_name] = pd.concat(dfs, ignore_index=True)  

if not scenario_data:  
    if log_func:  
        log_func("No scenario data to write into workbook.")  
    return None  

# ---------------------------  
# Create formatted workbook  
# ---------------------------  
workbook_path = os.path.join(  
    root_folder, "Combined_ViolationCTG_Comparison.xlsx"  
)  

if log_func:  
    log_func(f"\nBuilding FORMATTED combined workbook:\n  {workbook_path}")  

wb = Workbook()  
# Remove the default sheet; we'll create our own  
default_sheet = wb.active  
wb.remove(default_sheet)  

# Styles  
title_fill = PatternFill(fill_type="solid", fgColor="1F4E78")  
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

# Build each scenario sheet  
for folder_name, df in scenario_data.items():  
    # Sanity check columns  
    required = ["CaseType", "CTGLabel", "LimViolValue", "LimViolPct"]  
    for col in required:  
        if col not in df.columns and log_func:  
            log_func(f"  [{folder_name}] WARNING: column '{col}' missing.")  

    has_limviolid = "LimViolID" in df.columns  

    # Create sheet  
    sheet_name = (folder_name or "Sheet").strip()[:31]  
    if not sheet_name:  
        sheet_name = "Sheet"  
    ws = wb.create_sheet(title=sheet_name)  

    # Set column widths – contiguous columns B to E  
    ws.column_dimensions["B"].width = 55  # Contingency Events  
    ws.column_dimensions["C"].width = 55  # Resulting Issue  
    ws.column_dimensions["D"].width = 18  # Contingency Value (MVA)  
    ws.column_dimensions["E"].width = 18  # Percent Loading  

    current_row = 1  

    # Process blocks in fixed order: ACCA_LongTerm, ACCA_P1,2,4,7, DCwACver_P1-7  
    for label in TARGET_PATTERNS:  
        block_df = df[df["CaseType"] == label]  
        if block_df.empty:  
            continue  

        pretty_name = PRETTY_CASE_NAMES.get(label, label)  

        # ===== Title row =====  
        # Merge B:E for the title (no gaps between columns)  
        ws.merge_cells(  
            start_row=current_row,  
            start_column=2,  # B  
            end_row=current_row,  
            end_column=5,    # E  
        )  
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

        # ===== Data rows =====  
        for _, row in block_df.iterrows():  
            # Contingency Events (B)  
            c = ws.cell(row=current_row, column=2)  
            c.value = row.get("CTGLabel", "")  
            c.font = data_font  
            c.alignment = left_align  
            c.border = thin_border  

            # Resulting Issue (C)  
            c = ws.cell(row=current_row, column=3)  
            c.value = row.get("LimViolID", "") if has_limviolid else ""  
            c.font = data_font  
            c.alignment = left_align  
            c.border = thin_border  

            # Contingency Value (MVA) (D)  
            c = ws.cell(row=current_row, column=4)  
            c.value = row.get("LimViolValue", "")  
            c.font = data_font  
            c.alignment = center  
            c.border = thin_border  

            # Percent Loading (E)  
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