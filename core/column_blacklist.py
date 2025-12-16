# core/column_blacklist.py

import pandas as pd


# ------------------------------------------------------------
# Column Blacklist
# ------------------------------------------------------------
# Edit this list as needed. Anything in here will be removed
# AFTER filters run, right before saving the filtered CSV.
BLACKLIST_COLUMNS = {
    # Example columns you may want removed (add/remove freely):
    "ViolationCTG",
    "CaseName",
    "CaseID",
    "TimeStamp",
    "Timestamp",
    "DateTime",
    "Memo",
    "Notes",
    "ContingencyID",
    "MonitoredElementID",
    "MonitoredElementType",
}


def apply_blacklist(df: pd.DataFrame):
    """
    Remove blacklisted columns. Returns (filtered_df, removed_cols_list).
    """
    if df is None or df.empty:
        return df, []

    removed = []
    keep_cols = []
    for c in df.columns:
        if c in BLACKLIST_COLUMNS:
            removed.append(c)
        else:
            keep_cols.append(c)

    return df[keep_cols].copy(), removed


# ------------------------------------------------------------
# Row Filter (LimViolCat)
# ------------------------------------------------------------
def apply_row_filter(df: pd.DataFrame, keep_values, log_func=None):
    """
    Keep rows where LimViolCat is in keep_values.
    If LimViolCat doesn't exist, do nothing.
    Returns: (filtered_df, removed_count)
    """
    if df is None or df.empty:
        return df, 0

    if not keep_values:
        # If user selects nothing, treat that as "keep everything"
        if log_func:
            log_func("No categories selected; skipping LimViolCat row filter.")
        return df, 0

    if "LimViolCat" not in df.columns:
        if log_func:
            log_func("LimViolCat column not found; skipping category row filter.")
        return df, 0

    before = len(df)
    mask = df["LimViolCat"].astype(str).isin(set(keep_values))
    out = df.loc[mask].copy()
    removed = before - len(out)

    return out, removed


# ------------------------------------------------------------
# LimViolID max filter (FIXED)
# ------------------------------------------------------------
def _to_float_series(s: pd.Series) -> pd.Series:
    """
    Convert LimViolPct column to float safely.
    Handles numeric, strings, and strings with %.
    Non-convertible -> NaN.
    """
    if s is None:
        return pd.Series(dtype="float64")

    if pd.api.types.is_numeric_dtype(s):
        return s.astype(float)

    cleaned = (
        s.astype(str)
        .str.replace("%", "", regex=False)
        .str.strip()
    )
    return pd.to_numeric(cleaned, errors="coerce")


def apply_limviolid_max_filter(df: pd.DataFrame, log_func=None):
    """
    Dedup ONLY by LimViolID:
      - If multiple rows share the same LimViolID, keep the single row
        with the highest LimViolPct.
      - Do NOT dedup by CTGLabel.
      - If there are ties, keep the first after sorting (so only one row kept).
    Returns: (filtered_df, removed_count)
    """
    if df is None or df.empty:
        return df, 0

    if "LimViolID" not in df.columns:
        if log_func:
            log_func("LimViolID not found; skipping LimViolID max filter.")
        return df, 0

    if "LimViolPct" not in df.columns:
        # No pct -> keep first row per LimViolID
        before = len(df)
        out = df.drop_duplicates(subset=["LimViolID"], keep="first")
        removed = before - len(out)
        if log_func:
            log_func("LimViolPct not found. Keeping first row per LimViolID.")
            log_func(f"Rows removed by LimViolID dedup: {removed}")
        return out, removed

    work = df.copy()
    work["_LimViolPct_num"] = _to_float_series(work["LimViolPct"])

    before = len(work)

    # Sort so the max pct appears first within each LimViolID
    work = work.sort_values(
        by=["LimViolID", "_LimViolPct_num"],
        ascending=[True, False],
        na_position="last",
    )

    # Keep exactly ONE row per LimViolID (the one with highest pct)
    out = work.drop_duplicates(subset=["LimViolID"], keep="first").copy()

    out = out.drop(columns=["_LimViolPct_num"], errors="ignore")

    removed = before - len(out)

    if log_func:
        log_func("Dedup key: LimViolID")
        log_func("Keeping ONLY the highest LimViolPct row per LimViolID.")
        log_func(f"Rows before: {before}")
        log_func(f"Rows after:  {len(out)}")
        log_func(f"Rows removed by LimViolID max filter: {removed}")

    return out, removed
