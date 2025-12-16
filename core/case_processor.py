# core/case_processor.py

import os
import pandas as pd

from .pwb_exporter import export_violation_ctg
from .column_blacklist import (
    apply_blacklist,
    apply_row_filter,
    apply_limviolid_max_filter,
)


def _make_filtered_path(original_csv: str) -> str:
    base, ext = os.path.splitext(original_csv)
    if not ext:
        ext = ".csv"
    return f"{base}_Filtered{ext}"


def post_process_csv(csv_path: str, dedup_enabled: bool, keep_categories, log_func=None) -> str:
    """
    Apply:
      1) Row filter (LimViolCat) using keep_categories
      2) Optional LimViolID max filter (keep row with highest LimViolPct per LimViolID)
      3) Column blacklist
    Returns:
        path to filtered CSV (or None on failure)
    """
    if log_func:
        log_func("\nReading CSV to detect headers...")

    try:
        # Skip the first row because it only has "ViolationCTG" in one column.
        raw = pd.read_csv(csv_path, header=None, skiprows=1)

        if raw.shape[0] < 1:
            if log_func:
                log_func("Not enough rows in CSV to extract headers (need at least 1).")
            return None

        header_row = list(raw.iloc[0])
        if log_func:
            log_func(f"Detected {len(header_row)} headers from row 2.")

        if raw.shape[0] <= 1:
            if log_func:
                log_func("No data rows found after header row; nothing to filter.")
            return None

        # Data rows are index >= 1
        data = raw.iloc[1:].copy()
        data.columns = header_row

        # 1) Row filter with chosen categories
        if log_func:
            cats_txt = ", ".join(sorted(keep_categories)) if keep_categories else "NONE"
            log_func(f"\nApplying row filter for LimViolCat categories: {cats_txt}")
        filtered_data, removed_rows = apply_row_filter(
            data, keep_values=keep_categories, log_func=log_func
        )
        if log_func:
            log_func(f"Rows removed by row filter: {removed_rows}")

        # 2) Optional dedup by LimViolID
        if dedup_enabled:
            if log_func:
                log_func(
                    "\nApplying LimViolID max filter "
                    "(keep row with highest LimViolPct per LimViolID)..."
                )
            filtered_data, removed_max = apply_limviolid_max_filter(
                filtered_data, log_func=log_func
            )
            if log_func:
                log_func(f"Rows removed by LimViolID max filter: {removed_max}")
        else:
            if log_func:
                log_func("\nLimViolID max filter is disabled; skipping.")

        # 3) Column blacklist
        if log_func:
            log_func("\nApplying column blacklist...")
        filtered_data, removed_cols = apply_blacklist(filtered_data)

        if log_func:
            if removed_cols:
                log_func("Columns removed by blacklist:")
                for c in removed_cols:
                    log_func(f"  - {c}")
            else:
                log_func("No columns matched blacklist; no columns removed.")

        # Save filtered CSV
        filtered_csv = _make_filtered_path(csv_path)
        filtered_data.to_csv(filtered_csv, index=False)

        if log_func:
            log_func(f"Filtered CSV saved to:\n  {filtered_csv}")
            log_func("\nPreview of first few filtered data rows:")
            preview = filtered_data.head(10).to_string(index=False)
            log_func(preview)

        return filtered_csv

    except Exception as e:
        if log_func:
            log_func(f"(Could not read CSV for header inspection: {e})")
        return None


def process_case(
    pwb_path: str,
    dedup_enabled: bool,
    keep_categories,
    delete_original: bool,
    log_func=None,
) -> str:
    """
    Full pipeline for a single .pwb:
      - Export ViolationCTG to CSV via SimAuto
      - Run post_process_csv on it
      - Optionally delete the original (unfiltered) CSV
    Returns:
      path to filtered CSV (or None on error)
    """
    if log_func:
        log_func("\nConnecting to PowerWorld and exporting ViolationCTG...")

    csv_out = export_violation_ctg(pwb_path, log_func)

    if log_func:
        log_func(f"Exported CSV path: {csv_out}")

    filtered_csv = post_process_csv(csv_out, dedup_enabled, keep_categories, log_func)

    if delete_original and filtered_csv and os.path.exists(csv_out):
        try:
            os.remove(csv_out)
            if log_func:
                log_func(f"Deleted original (unfiltered) CSV: {csv_out}")
        except Exception as e:
            if log_func:
                log_func(f"WARNING: Failed to delete original CSV: {e}")

    return filtered_csv
