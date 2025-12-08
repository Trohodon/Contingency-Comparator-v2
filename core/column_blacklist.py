# core/column_blacklist.py

"""
Central place to define which columns should be removed (blacklisted)
from the ViolationCTG export.

You can edit BLACKLIST_BASE_NAMES and BLACKLIST_EXACT_NAMES as needed.

- If a column name is like 'BusNum:1', we treat 'BusNum' as the BASE name.
- Any column whose BASE is in BLACKLIST_BASE_NAMES will be removed.
- You can also blacklist specific full names in BLACKLIST_EXACT_NAMES.
"""


# Anything whose "base name" (before ':') is in here will be removed.
BLACKLIST_BASE_NAMES = {
    # >>> Add / edit these based on what you logged <<<
    "BusNum",
    "BusName",
    "BusNomVolt",
    "AreaNum",
    "AreaName",
    "ZoneNum",
    # Example of others you might want:
    # "OwnerNum",
    # "OwnerName",
    # "SubNum",
    # "SubName",
}

# Full header names to remove exactly (including suffixes like ':1', ':2', etc.).
# Leave empty if you're only using base names.
BLACKLIST_EXACT_NAMES = {
    # "SomeColumn:1",
    # "SomeColumn:2",
}


def is_blacklisted(col_name: str) -> bool:
    """
    Return True if a column name should be removed according to the blacklist.
    """
    name = str(col_name)
    base = name.split(":", 1)[0]  # "BusNum:1" -> "BusNum"

    if base in BLACKLIST_BASE_NAMES:
        return True

    if name in BLACKLIST_EXACT_NAMES:
        return True

    return False


def apply_blacklist(df):
    """
    Given a DataFrame whose columns are the ViolationCTG headers, return:
        (filtered_df, removed_columns)

    - filtered_df: new DataFrame with blacklisted columns removed
    - removed_columns: list of column names that were removed
    """
    original_cols = list(df.columns)
    keep_cols = [c for c in original_cols if not is_blacklisted(c)]
    removed_cols = [c for c in original_cols if c not in keep_cols]

    filtered_df = df[keep_cols].copy()
    return filtered_df, removed_cols