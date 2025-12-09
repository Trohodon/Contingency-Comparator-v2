# core/column_blacklist.py

"""
Central place for:
- Removing unwanted columns (column blacklist)
- Filtering unwanted ROWS (row filter)
- Optional dedup filter on LimViolID (keep highest LimViolPct)
"""

# ───────────────────────────────────────
# COLUMN BLACKLIST
# ───────────────────────────────────────

BLACKLIST_BASE_NAMES = {
    # Your base-name blacklist:
    "BusNum",
    "BusName",
    "BusNomVolt",
    "AreaNum",
    "AreaName",
    "ZoneNum",
    # ...add the rest of your base names here...
}

BLACKLIST_EXACT_NAMES = {
    # If you added any exact-name items, they go here
}


def is_blacklisted(col_name: str) -> bool:
    name = str(col_name)
    base = name.split(":", 1)[0]

    if base in BLACKLIST_BASE_NAMES:
        return True

    if name in BLACKLIST_EXACT_NAMES:
        return True

    return False


def apply_blacklist(df):
    """Remove columns from the DataFrame."""
    original_cols = list(df.columns)
    keep_cols = [c for c in original_cols if not is_blacklisted(c)]
    removed_cols = [c for c in original_cols if c not in keep_cols]

    filtered_df = df[keep_cols].copy()
    return filtered_df, removed_cols


# ───────────────────────────────────────
# ROW FILTER LOGIC (LimViolCat)
# ───────────────────────────────────────

ROW_FILTER_ENABLED = True
ROW_FILTER_COLUMN = "LimViolCat"
# Default if GUI doesn't specify anything:
ROW_FILTER_KEEP_VALUES = {"Branch MVA"}


def apply_row_filter(df, keep_values=None, log_func=None):
    """
    Filter out rows that don't match keep_values for LimViolCat.
    keep_values: iterable of category strings (e.g. {"Branch MVA", "Bus Low Volts"})
                 If None, falls back to ROW_FILTER_KEEP_VALUES.
                 If empty set/list, row filter is skipped.
    Return (filtered_df, removed_count).
    """
    if not ROW_FILTER_ENABLED:
        return df, 0

    if ROW_FILTER_COLUMN not in df.columns:
        if log_func:
            log_func(f"WARNING: Row filter column '{ROW_FILTER_COLUMN}' not found.")
        return df, 0

    if keep_values is None:
        keep_values = ROW_FILTER_KEEP_VALUES

    keep_values = set(keep_values)

    if not keep_values:
        if log_func:
            log_func("Row filter disabled: no LimViolCat categories selected.")
        return df, 0

    before = len(df)
    filtered_df = df[df[ROW_FILTER_COLUMN].isin(keep_values)].copy()
    after = len(filtered_df)
    removed = before - after

    return filtered_df, removed


# ───────────────────────────────────────
# DEDUP FILTER LOGIC (LimViolID / LimViolPct)
# ───────────────────────────────────────

DEDUP_ID_COLUMN = "LimViolID"
DEDUP_VALUE_COLUMN = "LimViolPct"


def apply_limviolid_max_filter(df, log_func=None):
    """
    For each LimViolID, keep only the row(s) with the highest LimViolPct.
    Return (filtered_df, removed_count).
    """
    if DEDUP_ID_COLUMN not in df.columns or DEDUP_VALUE_COLUMN not in df.columns:
        if log_func:
            log_func(
                "WARNING: Cannot apply LimViolID max filter because "
                f"'{DEDUP_ID_COLUMN}' or '{DEDUP_VALUE_COLUMN}' is missing."
            )
        return df, 0

    before = len(df)

    try:
        max_per_id = df.groupby(DEDUP_ID_COLUMN)[DEDUP_VALUE_COLUMN].transform("max")
    except Exception as e:
        if log_func:
            log_func(f"WARNING: Failed to compute max per LimViolID: {e}")
        return df, 0

    filtered_df = df[df[DEDUP_VALUE_COLUMN] == max_per_id].copy()
    after = len(filtered_df)
    removed = before - after

    return filtered_df, removed