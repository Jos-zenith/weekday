"""
Request-level tracking test. Usage: python tests/test_request_tracking.py [workbook]

Loads hand-made requests into a copy of the workbook, writes LOG lines in stages exactly as interns would, recalculates
in Excel and checks every request ID. Cases: repeated rounds at one company, one candidate with three requests,
same company and round with two interviewers, a re-sent duplicate, a link that turns into a booking later, an
explicit-ID line beating a blank line, wrong IDs, NO_SLOTS by request ID, and an ID used twice in INPUT.
"""
import datetime as dt
import os
import shutil
import sys

from openpyxl.utils import column_index_from_string as ci

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from excel import open_excel, sheet_errors  # noqa: E402

SRC = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "Weekday_Scheduling_Desk.xlsx"))
DST = os.path.join(HERE, "output", "request_tracking_test.xlsx")
shutil.copyfile(SRC, DST)

EPOCH = dt.datetime(1899, 12, 30)
serial = lambda t: (t - EPOCH).total_seconds() / 86400.0
ADDED = dt.datetime(2026, 10, 13, 19, 51)
DAY1 = dt.datetime(2026, 10, 14)
at = lambda h, m, s=0: DAY1 + dt.timedelta(hours=h, minutes=m, seconds=s)

YES1, YES2, YES3, NO1 = "Partner 002", "Partner 004", "Partner 005", "Partner 000"   # booking consent YES / NO in PARTNERS
ASHA, BALA, CHITRA, DEV, ESHA = "9111111111", "9222222222", "9333333333", "9444444444", "9555555555"


def person(company, intv, name, phone, rnd, added=ADDED):
    first = name.split()[0].lower()
    return [company, f"Interviewer {intv}", f"i{intv}@{company.replace(' ', '').lower()}.example.com", name,
            f"{first}@example.com", phone, "", rnd, serial(added), "Opened"]


INPUT = [
    person(YES1, 1, "Asha K.", ASHA, "Round1"),                                   # R001  repeated rounds,
    person(YES1, 1, "Asha K.", ASHA, "Round2"),                                   # R002  same company
    person(YES2, 1, "Bala M.", BALA, "Round1"),                                   # R003  one candidate,
    person(YES3, 1, "Bala M.", BALA, "Round1"),                                   # R004  three requests
    person(YES1, 2, "Bala M.", BALA, "Round2"),                                   # R005
    person(YES2, 1, "Chitra S.", CHITRA, "Round1"),                               # R006  same company + round,
    person(YES2, 2, "Chitra S.", CHITRA, "Round1"),                               # R007  different interviewer
    person(YES3, 1, "Dev P.", DEV, "Round1"),                                     # R008
    person(YES3, 1, "Dev P.", DEV, "Round1", ADDED + dt.timedelta(minutes=2)),    # R009  partner re-sent it
    person(NO1, 1, "Esha R.", ESHA, "Round1"),                                    # R010  partner wants its own link
]

results = []


def check(name, got, want):
    ok = got == want
    results.append(ok)
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + ("" if ok else f"   got {got!r}, want {want!r}"))


xl = open_excel()
wb = xl.Workbooks.Open(DST)
try:
    S = wb.Worksheets
    S("INPUT").Range("A2:J601").ClearContents()
    S("INPUT").Range(f"A2:J{1 + len(INPUT)}").Value = tuple(tuple(r) for r in INPUT)
    S("LOG").Range("A2:F2001").ClearContents()
    next_row = [2]

    def log(*lines):
        vals = tuple((serial(t), who, ph, out, rid, "") for t, who, ph, out, rid in lines)
        r0 = next_row[0]
        S("LOG").Range(f"A{r0}:F{r0 + len(vals) - 1}").Value = vals
        next_row[0] += len(vals)

    def calc(clock):
        S("SETTINGS").Range("B3").Value2 = serial(clock)
        xl.CalculateFull()

    def desk():
        v = S("DESK").Range("A2:AY12").Value
        col = lambda row, c: row[ci(c) - 1]
        return {col(r, "AW"): dict(unique=col(r, "L"), booked=col(r, "M"), declined=col(r, "N"), state=col(r, "AL"),
                                   tries=col(r, "Q"), by_us=col(r, "AP"), after_link=col(r, "AX"), hours=col(r, "AN"),
                                   in24=col(r, "AO"), problem=col(r, "AY"))
                for r in v if col(r, "AW")}

    def ram_list():
        return {str(r[2]): str(r[3]) for r in S("RAM").Range("A7:P31").Value if r[2]}

    dash = lambda cell: S("DASHBOARD").Range(cell).Value
    log_check = lambda row: S("LOG").Range(f"J{row}").Value

    # ---- stage 0: nothing logged
    calc(at(10, 30))
    d = desk()
    check("10 rows get IDs R001-R010", sorted(d), [f"R{k:03d}" for k in range(1, 11)])
    check("re-sent copy R009 is removed; R008 is kept", (d["R009"]["unique"], d["R008"]["unique"]), (0, 1))
    check("unique requests = 9", dash("B6"), 9)
    check("no request ID problems", dash("B9"), 0)
    lst = ram_list()
    check("Asha's row lists both rounds by ID", ("R001" in lst[ASHA], "R002" in lst[ASHA]), (True, True))
    check("Chitra's row lists both interviewers' requests", ("R006" in lst[CHITRA], "R007" in lst[CHITRA]), (True, True))
    check("Dev's row never shows the duplicate ID", "R009" in lst[DEV], False)

    # ---- stage 1: a morning of tries
    log((at(10, 5), "Ram", ASHA, "BOOKED", "R001"),                  # books Round1 only
        (at(10, 10), "Ram", BALA, "BOOKED", ""),                     # blank = all three of Bala's requests
        (at(10, 15), "Ram", CHITRA, "BOOKED", "R006"),               # one conversation, two answers
        (at(10, 15, 30), "Ram", CHITRA, "NOT_INTERESTED", "R007"),
        (at(10, 20), "Ram", ESHA, "SENT_LINK", ""))                  # a link is not a booking
    calc(at(10, 40))
    d, lst = desk(), ram_list()
    check("R001 booked by us, R002 (next round, same company) still open",
          (d["R001"]["booked"], d["R001"]["by_us"], d["R002"]["booked"], d["R002"]["state"]), (1, 1, 0, "Open"))
    check("Asha stays on the list, now with only R002", ("R001" in lst.get(ASHA, ""), "R002" in lst.get(ASHA, "")),
          (False, True))
    check("R001 time to schedule = 14.2 h (from its own booking line)", round(d["R001"]["hours"], 2), 14.23)
    check("blank BOOKED books all three of Bala's requests", [d[k]["booked"] for k in ("R003", "R004", "R005")], [1, 1, 1])
    check("R006 booked, R007 declined (same company, same round)",
          (d["R006"]["booked"], d["R007"]["declined"], d["R007"]["state"]), (1, 1, "Not interested"))
    check("Chitra's two lines are one try", d["R006"]["tries"], 1)
    check("R010 link sent is NOT booked", (d["R010"]["booked"], d["R010"]["state"]), (0, "Open: link sent, not booked yet"))
    check("dashboard: 5 scheduled, all by interns, 0 after a link", (dash("B10"), dash("B11"), dash("B12")), (5, 5, 0))
    check("dashboard: 1 link sent, 1 open with link not booked", (dash("B18"), dash("B15")), (1, 1))
    check("try 1 confirmed rate = 3 of 4 (the link does not count)", (dash("E38"), dash("D38")), (4, 0.75))

    # ---- stage 2: the link turns into a booking; a blank line does not override an explicit one
    log((at(12, 0), "Lead", ESHA, "SELF_BOOKED", "R010"),
        (at(12, 30), "Ram", ASHA, "NOT_INTERESTED", ""))             # blank: applies to R002 only, R001 stays booked
    calc(at(12, 40))
    d = desk()
    check("R010 booked by the candidate after our link",
          (d["R010"]["booked"], d["R010"]["by_us"], d["R010"]["after_link"]), (1, 0, 1))
    check("dashboard: 1 booked after a link, 0 links pending", (dash("B12"), dash("B15")), (1, 0))
    check("try 1 confirmed rate = 4 of 4 once the booking is logged", dash("D38"), 1.0)
    check("explicit BOOKED R001 survives a later blank NOT_INTERESTED", (d["R001"]["booked"], d["R002"]["declined"]), (1, 1))

    # ---- stage 3: wrong IDs are flagged; NO_SLOTS works by request ID
    r_dup, r_wrong = next_row[0], next_row[0] + 1
    log((at(13, 0), "Ram", DEV, "NO_ANSWER", "R009"),
        (at(13, 1), "Ram", DEV, "NO_ANSWER", "R001"),
        (at(13, 10), "Ram", DEV, "NO_SLOTS", "R008"))
    calc(at(13, 20))
    d = desk()
    check("ID of a removed duplicate is flagged", log_check(r_dup), "That ID is a removed duplicate: use the ID on your list")
    check("someone else's ID is flagged", log_check(r_wrong), "Request ID is not one of theirs: copy it from column D")
    check("NO_SLOTS on R008 pauses that partner's open request", d["R008"]["state"], "Partner has no slots")
    check("...but not R004, already booked at the same partner", d["R004"]["booked"], 1)

    # ---- stage 4: after the deadline
    calc(at(20, 0))
    d = desk()
    check("all 6 bookings count within 24h", sum(d[k]["in24"] for k in d), 6)
    check("final share = 6 of 9", round(dash("B3"), 4), round(6 / 9, 4))

    # ---- stage 5: the same ID typed twice in INPUT is caught
    S("INPUT").Range("K11").Value = "R001"
    calc(at(20, 0))
    check("an ID used twice is reported", (dash("B9"), desk()["R001"]["problem"]), (2, "ID used twice"))
    check("no formula errors", len(sheet_errors(wb)), 0)
    wb.Save()
finally:
    wb.Close(False)
    xl.Quit()
print(f"\n{sum(results)} of {len(results)} checks passed")
sys.exit(0 if all(results) else 1)
