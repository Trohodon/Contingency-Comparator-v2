from __future__ import annotations


# Ordered canonical case types. The order controls processing and workbook blocks.
CASE_TYPE_DEFINITIONS = (
    ("ACCA_LongTerm", "ACCA LongTerm"),
    ("ACCA_P1,2,4,7", "ACCA"),
    ("DCwACver_P1-7", "DCwAC"),
    ("AUXapplied", "AUXapplied"),
)

CASE_TYPES_CANONICAL = [canonical for canonical, _pretty in CASE_TYPE_DEFINITIONS]
CANONICAL_TO_PRETTY = dict(CASE_TYPE_DEFINITIONS)
PRETTY_TO_CANONICAL = {
    pretty: canonical for canonical, pretty in CASE_TYPE_DEFINITIONS
}

# Backward-compatible title used by older workbooks.
PRETTY_TO_CANONICAL["ACCA Long Term"] = "ACCA_LongTerm"

PRETTY_CASE_TITLES = tuple(PRETTY_TO_CANONICAL)

# Filename matching is case-insensitive.
TARGET_PATTERNS = {
    canonical: canonical for canonical in CASE_TYPES_CANONICAL
}
