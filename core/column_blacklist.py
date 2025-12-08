# core/column_blacklist.py

"""
Central place for:
- Removing unwanted columns (column blacklist)
- Filtering unwanted ROWS (row filter)
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
    # Add whatever else you've typed
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
# ROW FILTER LOGIC
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