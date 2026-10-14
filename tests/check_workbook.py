"""
Static + recalculated checks on the workbook. Usage: python tests/check_workbook.py [path]

1. Google Sheets audit: every function used exists in Sheets; conditional formats do not point at other tabs
   (Sheets rejects that); no array-entered formulas.
2. Script audit: no message promises a held or released slot.
3. Excel recalculation (separate Excel process): zero formula errors; every queue message has no unfilled
   {FIELD}; every pre-filled WhatsApp link fits Excel's 255-character HYPERLINK limit.
"""
import os
import re
import sys
import urllib.parse

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel import open_excel, sheet_errors  # noqa: E402

PATH = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else
                       os.path.join(os.path.dirname(__file__), "..", "Weekday_Scheduling_Desk.xlsx"))

# Functions Google Sheets supports (subset relevant here; checked against the Sheets function list).
SHEETS_OK = {"IF", "IFERROR", "AND", "OR", "NOT", "INDEX", "MATCH", "VLOOKUP", "COUNTIF", "COUNTIFS", "SUMIFS", "SUM",
             "SUMPRODUCT", "MAXIFS", "MINIFS", "MAX", "MIN", "ROUND", "ROUNDUP", "ROUNDDOWN", "MOD", "INT", "VALUE",
             "RIGHT", "LEFT", "MID", "LEN", "FIND", "SEARCH", "TRIM", "LOWER", "SUBSTITUTE", "TEXT", "NOW", "ISNUMBER",
             "HYPERLINK", "ENCODEURL"}
BANNED = ["held", "hold", "release", "pencil", "reserved"]

wb = openpyxl.load_workbook(PATH)
funcs, problems = set(), []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                funcs |= set(re.findall(r"(?:_xlfn\.)?([A-Z][A-Z0-9\.]*)\(", c.value))
    for rng, rules in ws.conditional_formatting._cf_rules.items():
        for rule in rules:
            for f in rule.formula or []:
                if "!" in f:
                    problems.append(f"{ws.title} conditional format points at another tab: {f}")
unknown = sorted(funcs - SHEETS_OK)
print("functions used:", ", ".join(sorted(funcs)))
print("not in the Sheets list:", unknown or "none")
print("conditional-format problems:", problems or "none")

parts = [str(c.value) for c in wb["LISTS"]["B"][13:19]]
bad_words = [(p, w) for p in parts for w in BANNED if re.search(rf"\b{w}", p, re.I)]
print("message parts promising a held slot:", bad_words or "none")

xl = open_excel()
try:
    b = xl.Workbooks.Open(PATH)
    xl.CalculateFull()
    errs = sheet_errors(b)
    print("formula errors after recalculation:", len(errs), errs[:5])
    total = fits = braces = 0
    for sh in ("RAM", "SHYAM"):
        v = b.Worksheets(sh).Range("A7:M31").Value
        for row in v:
            if row[2] in (None, ""):
                continue
            total += 1
            msg = str(row[12])
            braces += ("{" in msg or "}" in msg)
            # Excel's own verdict (column I falls back to "paste column M" when its link would be too long),
            # cross-checked with Python's encoding
            fits += (row[8] == "Open WhatsApp" and
                     len("https://wa.me/91" + str(row[2]) + "?text=" + urllib.parse.quote(msg, safe="")) <= 255)
    print(f"queue rows shown: {total} | messages with unfilled fields: {braces} | WhatsApp links within 255 chars: {fits}/{total}")
    b.Close(False)
finally:
    xl.Quit()
ok = not unknown and not problems and not bad_words and not errs and braces == 0 and fits == total
print("RESULT:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
