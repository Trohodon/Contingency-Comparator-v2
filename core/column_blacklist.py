# core/column_blacklist.py

"""
List of column names to drop from ViolationCTG exports.

Copy the names from your log (without the leading "- " and spaces)
and put them in this list as plain strings.

Example:
    "BusNum",
    "BusNum:1",
    ...

Make sure each name matches exactly what appears in the header row
of the CSV.
"""

BLACKLISTED_COLUMNS = [
    # TODO: paste your full list here, e.g.:
    # "BusNum",
    # "BusNum:1",
    # "BusNum:2",
    # "BusName",
    # "BusName:1",
    # "BusNomVolt",
    # "BusNomVolt:1",
    # ...
]