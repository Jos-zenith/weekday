"""
Sanity checks on the workbook. Run: python tests/check_workbook.py [path]

1. Will it work in Google Sheets? Every function used has to exist there, and no conditional format may point at
   another tab (Sheets doesn't allow that).
2. No message promises a held slot.
3. Recalculate in Excel (in its own process): no formula errors, no unfilled {FIELD} in any message, and every
   pre-filled WhatsApp link fits in 255 characters.
"""
import os
import re
import sys
import urllib.parse

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from excel import open_excel, sheet_errors  # noqa: E402

PATH = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else
                       os.path.join(os.path.dirname(__file__), "..", "Weekday Scheduling Desk.xlsx"))

# Functions Google Sheets supports (subset relevant here; checked against the Sheets function list).
SHEETS_OK = {"IF", "IFERROR", "AND", "OR", "NOT", "INDEX", "MATCH", "VLOOKUP", "COUNTIF", "COUNTIFS", "SUMIFS", "SUM",
             "SUMPRODUCT", "MAXIFS", "MINIFS", "MAX", "MIN", "ROUND", "ROUNDUP", "ROUNDDOWN", "MOD", "INT", "VALUE",
             "RIGHT", "LEFT", "MID", "LEN", "FIND", "SEARCH", "TRIM", "LOWER", "SUBSTITUTE", "TEXT", "NOW", "ISNUMBER",
             "HYPERLINK", "ENCODEURL", "ROW"}
BANNED = ["held", "hold", "release", "pencil", "reserved"]

def paren_balance(f):
    """Net open brackets outside string literals (Excel refuses to open a file with an unbalanced formula)."""
    depth, in_str = 0, False
    for ch in f:
        if ch == '"':
            in_str = not in_str
        elif not in_str:
            depth += (ch == "(") - (ch == ")")
    return depth


wb = openpyxl.load_workbook(PATH)
funcs, problems, unbalanced = set(), [], []
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                funcs |= set(re.findall(r"(?:_xlfn\.)?([A-Z][A-Z0-9\.]*)\(", c.value))
                if paren_balance(c.value) and len(unbalanced) < 5:
                    unbalanced.append(f"{ws.title}!{c.coordinate}")
print("formulas with unbalanced brackets:", unbalanced or "none")
if unbalanced:
    sys.exit(1)
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
