# core/comparison_builder.py

import os
import pandas as pd

from .case_finder import TARGET_PATTERNS


def build_workbook(root_folder, folder_to_case_csvs, log_func=None):
    """
    Build a combined Excel workbook with one sheet per subfolder.

    Args:
        root_folder: top-level folder user selected.
        folder_to_case_csvs: dict mapping
            folder_name -> { case_type_label -> filtered_csv_path }
        log_func: optional logging function.

    Returns:
        Full path to the created workbook, or None if nothing was written.
    """
    if not folder_to_case_csvs:
        if log_func:
            log_func("No data to build combined workbook.")
        return None

    workbook_path = os.path.join(
        root_folder, "Combined_ViolationCTG_Comparison.xlsx"
    )

    if log_func:
        log_func(f"\nBuilding combined workbook:\n  {workbook_path}")

    # Use Excel writer so we can create multiple sheets.
    with pd.ExcelWriter(workbook_path, engine="xlsxwriter") as writer:
        for folder_name, case_map in folder_to_case_csvs.items():
            dfs = []

            # Preserve fixed order of case types
            for label in TARGET_PATTERNS:
                csv_path = case_map.get(label)
                if not csv_path:
                    if log_func:
                        log_func(
                            f"  [{folder_name}] Missing case type '{label}' "
                            f"(no filtered CSV found)."
                        )
                    continue

                try:
                    df = pd.read_csv(csv_path)
                    # Optional: tag which case each row came from
                    df.insert(0, "CaseType", label)
                    dfs.append(df)
                except Exception as e:
                    if log_func:
                        log_func(
                            f"  [{folder_name}] WARNING: Failed to read {csv_path}: {e}"
                        )

            if not dfs:
                if log_func:
                    log_func(f"  [{folder_name}] No CSVs to combine; skipping sheet.")
                continue

            combined = pd.concat(dfs, ignore_index=True)

            # Excel sheet names must be <= 31 chars and non-empty
            sheet_name = (folder_name or "Sheet").strip()[:31]
            if not sheet_name:
                sheet_name = "Sheet"

            combined.to_excel(writer, sheet_name=sheet_name, index=False)

    if log_func:
        log_func("Combined workbook build complete.")
    return workbook_path