"""
Central place for:
- Removing unwanted columns (column blacklist)
- Filtering unwanted ROWS (row filter)
- Optional dedup filter on LimViolID (keep highest LimViolPct)
"""

import pandas as pd

# ───────────────────────────────────────
# COLUMN BLACKLIST
# ───────────────────────────────────────

BLACKLIST_BASE_NAMES = {
    # >>> Add / edit these based on what you logged <<<
    "BusNum",
    "BusName",
    "BusNomVolt",
    "AreaNum",
    "AreaName",
    "ZoneNum",
    "ZoneName",
    "BGLoadMW",
    "BGGenMW",
    "LoadMW",
    "GenMW",
    "LineCircuit",
    "LinePTDF",
    "IslandTotalBus",
    "CTGSkip",
    "CTCProc",
    "CTGSolved",
    "CTGViol",
    "CTGViolMaxLine",
    "CTGViolMinVolt",
    "CTGViolMaxVolt",
    "CTGViol",
    "CTGProc",
    "LimViolCat",
    "LimViolLimit",
    "LineLength",
    "OwnerName",
    "CustomFloat",
    "CTGViolCompare",
    "CTGViolMaxLineCompare",
    "CTGViolMinVoltCompare",
    "CTGViolMaxVoltCompare",
    "CTGViolDiff",
    "CTGViolMaxLineDiff",
    "CTGViolMinVoltDiff",
    "CTGViolMaxVoltDiff",
    "LimViolLimitCompare",
    "LimViolValueCompare",
    "LimViolPctCompare",
    "LimViolLimitDiff",
    "LimViolValueDiff",
    "LimViolPctDiff",
    "CustomExpression",
    "LineMonEle",
    "SubNum",
    "SubName",
    "Selected",
    "CTGNItr",
    "CtgFileName",
    "CTG_CalculationMethod",
    "CTGRANK",
    "CTGViolMaxInterface",
    "CTGViolMaxInterfaceCompare",
    "CTGViolMaxInterfaceDiff",
    "AllLabels",
    "Label",
    "CTGNBranchViol",
    "CTGNInterfaceViol",
    "CTGNVoltViol",
    "AggrPercentOverload",
    "ObjectMemo",
    "PVCritical",
    "QVAutoplot",
    "CustomString",
    "AggrMVAOverload",
    "SymbolType",
    "CalcField",
    "Latitude",
    "Longitude",
    "LatitudeString",
    "LongitudeString",
    "CTGViolMaxdVdQ",
    "SODashed",
    "EMSDeviceID",
    "CustomInteger",
    "PLVisible",
    "PLThickness",
    "PLColor",
    "CTGUseMonExcept",
    "CTGIgnoreSolutionOptions",
    "ContainedInDiffFlowsBC",
    "CTGSolutionOptions",
    "TSCTGElementCount",
    "CustomExpressionStr",
    "CTGSolvedComparison",
    "LabelAppend",
    "Include",
    "CTGCustMonViol",
    "CTGWhatOccurredCount",
    "Category",
    "ObjectID",
    "LimViolCTGSpecifiedLimit",
    "EMSType",
    "CustomIntegerOther",
    "CustomFloatOther",
    "CustomStringOther",
    "BAName",
    "BANumber",
    "DataMaintainer",
    "DataMaintainerAssign",
    "SourceList",
    "LimitScaled",
    "LimitCompareScaled",
    "LimitDiffScaled",
    "PercentScaled",
    "PercentCompareScaled",
    "PercentDiffScaled",
    "DataMaintainerInherit",
    "Note",
    "CustomExpressionOther",
    "CustomExpressionStrOther",
    "ScreenAllow",
    "ScreenRank",
    "CTGViolMaxBusPairAngle",
    "CTGViolMaxBusPairAngleCompare",
    "CTGViolMaxBusPairAngleDiff",
    "CTGNBusPairAngleViol",
    "EMSViolID",
    "NormalRatingNoAction",
    "DataCheck",
    "DataCheckAggr",
    "InjectorMax",
    "InjectorMin",
    "InjectorInc",
    "InjectorDec",
    "MWEffectInc",
    "MWEffectDec",
    "MWInjSensMax",
    "MWInjSensMin",
    "CalcFieldExtra",
    "CTGUseSolutionOptions",
    "CTGSolutionTimeSeconds",
    "CTGAltPFPossible",
    "CTGAltPFCheckAllow",
    "CTGAltPFBusCount",
    "CTGRemedialActionApplied",
    "ReferenceDistance",
    "FixedNumBus",
    "SubNodeNum",
}

# Full header names to remove exactly (including suffixes like ':1', ':2', etc.).
# Leave empty if you're only using base names.
BLACKLIST_EXACT_NAMES = {
    "LimViolID:1",
    "LimViolID:2",
    "LimViolValue:1",
    "LimViolValue:2",
    "LimViolValue:3",
    "LimViolPct:1",
    "LimViolPct:2",
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


def apply_limviolid_max_filter(df, log_func=None):
    """
    Dedup ONLY by LimViolID:
      - If multiple rows share the same LimViolID, keep exactly ONE row:
        the row with the highest LimViolPct.
      - Do NOT dedup by CTGLabel.
      - If LimViolPct is missing, keep first row per LimViolID.
    Return (filtered_df, removed_count).
    """
    if df is None or df.empty:
        return df, 0

    if DEDUP_ID_COLUMN not in df.columns:
        if log_func:
            log_func(f"WARNING: '{DEDUP_ID_COLUMN}' not found; skipping LimViolID max filter.")
        return df, 0

    before = len(df)

    # If pct column missing, just drop duplicates by LimViolID (keep first)
    if DEDUP_VALUE_COLUMN not in df.columns:
        out = df.drop_duplicates(subset=[DEDUP_ID_COLUMN], keep="first").copy()
        removed = before - len(out)
        if log_func:
            log_func(f"WARNING: '{DEDUP_VALUE_COLUMN}' not found; keeping first row per LimViolID.")
            log_func(f"Rows removed by LimViolID dedup: {removed}")
        return out, removed

    work = df.copy()
    work["_LimViolPct_num"] = _to_float_series(work[DEDUP_VALUE_COLUMN])

    # Sort so the max pct appears first per LimViolID
    work = work.sort_values(
        by=[DEDUP_ID_COLUMN, "_LimViolPct_num"],
        ascending=[True, False],
        na_position="last",
    )

    # Keep exactly one per LimViolID
    out = work.drop_duplicates(subset=[DEDUP_ID_COLUMN], keep="first").copy()
    out = out.drop(columns=["_LimViolPct_num"], errors="ignore")

    removed = before - len(out)

    if log_func:
        log_func("Dedup key: LimViolID")
        log_func("Keeping ONLY the highest LimViolPct row per LimViolID.")
        log_func(f"Rows before: {before}")
        log_func(f"Rows after:  {len(out)}")
        log_func(f"Rows removed by LimViolID max filter: {removed}")

    return out, removed
