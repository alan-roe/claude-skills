"""
OOXML future function compatibility.

openpyxl doesn't add the _xlfn. prefix required by OOXML for "future
functions" (MAXIFS, IFS, etc.). Without it, LibreOffice shows #NAME?
errors. Excel and Google Sheets silently add the prefix on open.

Dynamic-array / spilling functions (XLOOKUP, XMATCH, SORT, FILTER,
UNIQUE, SEQUENCE, SORTBY, RANDARRAY) are deliberately excluded: an
openpyxl-written file has no spill-range metadata, so LibreOffice either
errors unpredictably (#NAME?/#VALUE!, function-dependent) or — worse —
silently writes only the top-left cell of the result and reports zero
errors. Auto-adding the prefix for these would trade an honest #NAME?
for a silently wrong workbook. See skills/xlsx/SKILL.md, "Choosing
formulas that survive verification".
"""

import re

# Sorted longest-first so regex matches greedily (FORECAST.ETS.CONFINT before FORECAST.ETS).
FUTURE_FUNCTIONS = sorted(
    [
        # Excel 2019 / Office 365 (scalar — safe to prefix)
        "CONCAT", "IFS", "MAXIFS", "MINIFS", "SWITCH", "TEXTJOIN",
        # Advanced (365) — scalar, don't spill
        "LET", "LAMBDA",
        # Forecast (2016+)
        "FORECAST.ETS", "FORECAST.ETS.CONFINT", "FORECAST.ETS.SEASONALITY",
        "FORECAST.ETS.STAT", "FORECAST.LINEAR",
        # Math & trig (2013+)
        "ACOT", "ACOTH", "ARABIC", "BASE", "CEILING.MATH", "CEILING.PRECISE",
        "COMBINA", "COT", "COTH", "CSC", "CSCH", "DECIMAL",
        "FLOOR.MATH", "FLOOR.PRECISE", "GAMMA", "GAUSS", "MUNIT",
        "PERMUTATIONA", "PHI", "RRI", "SEC", "SECH",
        # Logical / info (2013+)
        "BITAND", "BITOR", "BITXOR", "BITLSHIFT", "BITRSHIFT",
        "IFNA", "ISFORMULA", "NUMBERVALUE", "PDURATION",
        "SHEET", "SHEETS", "SKEW.P", "UNICHAR", "UNICODE", "XOR",
        # Date (2013+)
        "DAYS", "ISO.CEILING", "ISOWEEKNUM",
        # Web / text (2013+)
        "ENCODEURL", "FILTERXML", "FORMULATEXT", "WEBSERVICE",
    ],
    key=len,
    reverse=True,
)

_PATTERN = re.compile(
    r"(?<![A-Za-z0-9_.])"
    r"(" + "|".join(re.escape(f) for f in FUTURE_FUNCTIONS) + r")"
    r"(?=\s*\()",
    re.IGNORECASE,
)


def add_xlfn_prefixes(wb):
    """Add _xlfn. prefix to future functions in all formulas.

    Returns count of modified cells.
    """
    modified = 0
    for ws in wb.worksheets:
        if not hasattr(ws, "iter_rows"):
            continue
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.startswith("="):
                    continue
                new_v = _PATTERN.sub(r"_xlfn.\1", v)
                if new_v != v:
                    cell.value = new_v
                    modified += 1
    return modified
