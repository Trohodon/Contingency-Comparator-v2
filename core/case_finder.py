# core/case_finder.py

import os

# Map human-friendly labels to filename substrings we look for
TARGET_PATTERNS = {
    "ACCA_LongTerm": "ACCA_LongTerm",
    "ACCA_P1,2,4,7": "ACCA_P1,2,4,7",
    "DCwACver_P1-7": "DCwACver_P1-7",
}


def _find_pwb_files(folder: str):
    """Return a list of .pwb filenames (no paths) in the folder."""
    return [f for f in os.listdir(folder) if f.lower().endswith(".pwb")]


def _classify_case(filename: str) -> str:
    """Return the case type label based on TARGET_PATTERNS, or 'Other'."""
    for label, pattern in TARGET_PATTERNS.items():
        if pattern in filename:
            return label
    return "Other"


def scan_folder(folder: str, log_func=None):
    """
    Scan a folder for .pwb files.

    Returns:
        cases: list of dicts with keys:
            - filename
            - path
            - type  (e.g. 'ACCA_LongTerm', 'ACCA_P1,2,4,7', 'DCwACver_P1-7', or 'Other')
            - is_target (True if type != 'Other')

        target_cases: dict mapping type -> full path (first one found for each type)
    """
    cases = []
    target_cases = {}

    if log_func:
        log_func(f"\nScanning folder for .pwb cases:\n{folder}")

    pwb_files = _find_pwb_files(folder)
    if not pwb_files:
        if log_func:
            log_func("No .pwb files found in folder.")
        return cases, target_cases

    for fname in sorted(pwb_files):
        fpath = os.path.join(folder, fname)
        ctype = _classify_case(fname)
        is_target = ctype != "Other"

        if is_target:
            if ctype not in target_cases:
                target_cases[ctype] = fpath
            else:
                if log_func:
                    log_func(
                        f"WARNING: Multiple cases found for type '{ctype}'. "
                        f"Using first: {target_cases[ctype]}"
                    )

        cases.append(
            {
                "filename": fname,
                "path": fpath,
                "type": ctype,
                "is_target": is_target,
            }
        )

    if log_func:
        log_func("Folder scan complete.")
        for label in TARGET_PATTERNS:
            if label in target_cases:
                log_func(f"  Found target case [{label}]: {target_cases[label]}")
            else:
                log_func(f"  WARNING: No case found for type [{label}] in this folder.")

    return cases, target_cases
