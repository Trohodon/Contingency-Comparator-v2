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
    # Your existing base-name blacklist goes here:
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

# Only keep rows where LimViolCat == "Branch MVA"
ROW_FILTER_ENABLED = True
ROW_FILTER_COLUMN = "LimViolCat"
ROW_FILTER_KEEP_VALUES = {"Branch MVA"}  # everything else is dropped


def apply_row_filter(df, log_func=None):
    """
    Filter out rows that don't match ROW_FILTER_KEEP_VALUES.
    Return (filtered_df, removed_count).
    """
    if not ROW_FILTER_ENABLED:
        return df, 0

    if ROW_FILTER_COLUMN not in df.columns:
        if log_func:
            log_func(f"WARNING: Row filter column '{ROW_FILTER_COLUMN}' not found.")
        return df, 0

    before = len(df)
    filtered_df = df[df[ROW_FILTER_COLUMN].isin(ROW_FILTER_KEEP_VALUES)].copy()
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

    If LimViolID or LimViolPct is missing, does nothing.
    """
    if DEDUP_ID_COLUMN not in df.columns or DEDUP_VALUE_COLUMN not in df.columns:
        if log_func:
            log_func(
                "WARNING: Cannot apply LimViolID max filter because "
                f"'{DEDUP_ID_COLUMN}' or '{DEDUP_VALUE_COLUMN}' is missing."
            )
        return df, 0

    before = len(df)

    # Compute max LimViolPct per LimViolID
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