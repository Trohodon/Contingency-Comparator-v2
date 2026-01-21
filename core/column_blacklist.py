"""
Central place for:
- Removing unwanted columns (column blacklist)
- Filtering unwanted ROWS (row filter)
- OPTIONAL LimViolID grouping (keep max as primary, preserve others as expandable "children")

IMPORTANT CHANGE (v2 request):
- We NO LONGER delete rows for duplicate LimViolID.
- We rank/group them instead, so the highest LimViolPct is the "main" row,
  and the other rows remain available for GUI expansion and Excel outlining.

Outputs added by grouping:
  __LimViolPct_num : numeric version of LimViolPct for sorting (can be hidden later)
  __GroupRank      : 1 = primary (max pct), 2..N = additional contingencies
  __GroupSize      : number of rows in this LimViolID group
  __IsChild        : True if __GroupRank > 1
"""

import pandas as pd

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

# Helper columns introduced by grouping logic
GROUP_HELPER_COLUMNS = {
    "__LimViolPct_num",
    "__GroupRank",
    "__GroupSize",
    "__IsChild",
}


def is_blacklisted(col_name: str) -> bool:
    name = str(col_name)
    base = name.split(":", 1)[0]

    if base in BLACKLIST_BASE_NAMES:
        return True

    if name in BLACKLIST_EXACT_NAMES:
        return True

    return False


def apply_blacklist(df: pd.DataFrame):
    """
    Remove columns from the DataFrame using the blacklist.

    NOTE:
      We DO NOT automatically remove helper columns here because:
      - GUI might need them for expand/collapse behavior
      - Excel exporter might need them to build outline groups

    If you want helper columns hidden from the GUI table, do that in the GUI layer
    (e.g., omit them when building columns).
    """
    if df is None or df.empty:
        return df, []

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
ROW_FILTER_KEEP_VALUES = {"Branch MVA"}  # Default if GUI doesn't specify anything


def apply_row_filter(df: pd.DataFrame, keep_values=None, log_func=None):
    """
    Filter out rows that don't match keep_values for LimViolCat.

    keep_values: iterable of category strings (e.g. {"Branch MVA", "Bus Low Volts"})
                 If None, falls back to ROW_FILTER_KEEP_VALUES.
                 If empty set/list, row filter is skipped.

    Return (filtered_df, removed_count).
    """
    if df is None or df.empty:
        return df, 0

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
# LimViolID GROUPING (no deletion)
# ───────────────────────────────────────

GROUPING_ENABLED = True

DEDUP_ID_COLUMN = "LimViolID"
DEDUP_VALUE_COLUMN = "LimViolPct"


def _to_float_series(series: pd.Series) -> pd.Series:
    """
    Convert LimViolPct values to float safely.
    Supports:
      - numeric types
      - strings
      - strings with '%' sign
    Non-convertible values become NaN.
    """
    if series is None:
        return pd.Series(dtype="float64")

    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float)

    cleaned = (
        series.astype(str)
        .str.replace("%", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce")


def apply_limviolid_grouping(df: pd.DataFrame, log_func=None):
    """
    Group/rank ONLY by LimViolID:
      - Keep ALL rows
      - Sort so that highest LimViolPct is first per LimViolID (rank 1)
      - Add helper columns for GUI expansion and Excel outlining

    Return (grouped_df, info_dict)
    """
    if df is None or df.empty:
        return df, {"groups": 0, "groups_with_children": 0, "child_rows": 0}

    if not GROUPING_ENABLED:
        return df, {"groups": 0, "groups_with_children": 0, "child_rows": 0}

    if DEDUP_ID_COLUMN not in df.columns:
        if log_func:
            log_func(f"WARNING: '{DEDUP_ID_COLUMN}' not found; skipping LimViolID grouping.")
        return df, {"groups": 0, "groups_with_children": 0, "child_rows": 0}

    work = df.copy()

    # Numeric pct for sorting
    if DEDUP_VALUE_COLUMN in work.columns:
        work["__LimViolPct_num"] = _to_float_series(work[DEDUP_VALUE_COLUMN])
    else:
        work["__LimViolPct_num"] = float("nan")
        if log_func:
            log_func(f"WARNING: '{DEDUP_VALUE_COLUMN}' not found; ranking within LimViolID will be arbitrary.")

    # Sort so max pct appears first per LimViolID
    work = work.sort_values(
        by=[DEDUP_ID_COLUMN, "__LimViolPct_num"],
        ascending=[True, False],
        na_position="last",
    ).reset_index(drop=True)

    # Rank inside each LimViolID group
    work["__GroupRank"] = work.groupby(DEDUP_ID_COLUMN).cumcount() + 1
    work["__GroupSize"] = work.groupby(DEDUP_ID_COLUMN)[DEDUP_ID_COLUMN].transform("size")
    work["__IsChild"] = work["__GroupRank"] > 1

    # Logging stats
    total_groups = int(work[DEDUP_ID_COLUMN].nunique())
    groups_with_children = int((work.groupby(DEDUP_ID_COLUMN)["__GroupSize"].first() > 1).sum())
    child_rows = int(work["__IsChild"].sum())

    if log_func:
        log_func("Grouping key: LimViolID (NO ROWS DELETED)")
        log_func("Primary row per LimViolID = highest LimViolPct")
        log_func(f"Total LimViolID groups: {total_groups}")
        log_func(f"Groups with additional contingencies: {groups_with_children}")
        log_func(f"Child rows (expandable): {child_rows}")

    return work, {
        "groups": total_groups,
        "groups_with_children": groups_with_children,
        "child_rows": child_rows,
    }


def get_primary_rows_only(df_grouped: pd.DataFrame) -> pd.DataFrame:
    """
    Returns ONLY the primary rows (rank 1) from a grouped DataFrame.
    This is what you show in the GUI table by default.
    """
    if df_grouped is None or df_grouped.empty:
        return df_grouped

    if "__GroupRank" not in df_grouped.columns:
        return df_grouped

    return df_grouped[df_grouped["__GroupRank"] == 1].copy()


def strip_group_helper_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Optional helper:
    Remove grouping helper columns from a dataframe copy.
    Useful if you want a clean CSV export, but keep them for XLSX outlining.
    """
    if df is None or df.empty:
        return df
    return df.drop(columns=[c for c in GROUP_HELPER_COLUMNS if c in df.columns], errors="ignore").copy()


# ───────────────────────────────────────
# BACKWARDS-COMPAT ALIAS (so existing code won't crash)
# ───────────────────────────────────────

def apply_limviolid_max_filter(df: pd.DataFrame, log_func=None):
    """
    BACKWARDS COMPATIBILITY:
    Old behavior deleted duplicates. New requirement is to preserve them.

    So this function now:
      - applies grouping
      - returns ONLY primary rows (rank 1) as the filtered_df
      - and reports removed_count as 0 (because nothing is deleted anymore)

    This lets existing code keep working while you migrate GUI/export
    to use the full grouped dataframe.
    """
    grouped, info = apply_limviolid_grouping(df, log_func=log_func)
    primary = get_primary_rows_only(grouped)

    # We do NOT delete rows anymore, so removed_count is conceptually 0.
    # (If your logs want a number, you can compute "child_rows", but they're preserved.)
    removed_count = 0

    if log_func and info:
        log_func("NOTE: LimViolID 'max filter' is now non-destructive (grouping mode).")
        log_func("Primary rows returned for main table; child rows preserved for expansion/export.")
        log_func(f"Primary rows: {len(primary)} | Child rows preserved: {info.get('child_rows', 0)}")

    return primary, removed_count