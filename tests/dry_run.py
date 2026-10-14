"""
Dry run: drive the workbook through one simulated working day (14 Oct 10:00 -> deadline 19:51).

A bot plays Ram and Shyam. It reads ONLY what the RAM / SHYAM tabs show, follows column G, and writes LOG lines
exactly as START HERE tells an intern to. Candidate behaviour comes from hidden truth (tests/output/truth.json),
drawn from the analysis model's ASSUMED rates. So this tests the sheet's mechanics; its booking rate is NOT
evidence about real candidates and must not be quoted as one.

Usage: python tests/dry_run.py [workbook]   (default: Weekday_Scheduling_Desk.xlsx; writes tests/output/)
"""
import collections
import datetime as dt
import json
import os
import shutil
import sys

import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(ROOT, "analysis"))
import sim2  # noqa: E402
from excel import open_excel, sheet_errors  # noqa: E402

SRC = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "Weekday_Scheduling_Desk.xlsx"))
OUTD = os.path.join(HERE, "output")
DST = os.path.join(OUTD, "dry_run_day1.xlsx")
shutil.copyfile(SRC, DST)

truth = json.load(open(os.path.join(OUTD, "truth.json")))
DAY0 = dt.datetime.fromisoformat(truth["day0"])
cls = {p: v["cls"] for p, v in truth["cand"].items()}
consent = truth["consent"]
self_min = {(r["phone"], r["company"], r["round"]): r["self_min"] for r in truth["req"]}
ANS = np.array(sim2.BASE["ans_p"])
CONV, EXTRA, OPT = sim2.BASE["conv_p"], sim2.BASE["extra_conv"], sim2.BASE["opt_p"]
P_LINK, LINK_DELAY = 0.6, (15, 180)        # ASSUMED: share who book after we send a link, and how long they take
rng = np.random.default_rng(424242)

START = DAY0 + dt.timedelta(days=1, hours=10)
END = DAY0 + dt.timedelta(days=1, hours=19, minutes=51)
CYCLE = 30
SHIFT = {"Ram": (10 * 60, 18 * 60), "Shyam": (12 * 60, 20 * 60)}
EPOCH = dt.datetime(1899, 12, 30)
serial = lambda t: (t - EPOCH).total_seconds() / 86400.0
to_abs = lambda t: (t - DAY0).total_seconds() / 60.0

xl = open_excel()
wb = xl.Workbooks.Open(DST)
xl.Calculation = -4135            # manual; recalculated after each batch of lines
S = wb.Worksheets
log, settings = S("LOG"), S("SETTINGS")
next_row = 2
booked = set()                    # (phone, company, round) booked so far
pending_link = []                 # (time, phone, company, round): candidate books from a link we sent
stats = collections.Counter()
codes_seen = collections.Counter()


def recalc(clock):
    settings.Range("B3").Value2 = serial(clock)
    xl.CalculateFull()


def write_log(lines):
    global next_row
    if not lines:
        return
    lines.sort(key=lambda x: x[0])
    vals = tuple(tuple([serial(x[0])] + list(x[1:])) for x in lines)
    log.Range(log.Cells(next_row, 1), log.Cells(next_row + len(lines) - 1, 6)).Value = vals
    next_row += len(lines)


try:
    clock = START
    recalc(clock)
    while True:
        lines = []
        # 1. bookings candidates made themselves (the lead copies Calendly confirmations into LOG)
        for (ph, comp, rd), sm in self_min.items():
            if sm is not None and (ph, comp, rd) not in booked and DAY0 + dt.timedelta(minutes=sm) <= clock:
                booked.add((ph, comp, rd))
                lines.append([DAY0 + dt.timedelta(minutes=sm), "Lead", ph, "SELF_BOOKED", comp, "calendly confirmation"])
        for item in list(pending_link):
            t, ph, comp, rd = item
            if t <= clock:
                pending_link.remove(item)
                if (ph, comp, rd) not in booked:
                    booked.add((ph, comp, rd))
                    lines.append([t, "Lead", ph, "SELF_BOOKED", comp, "booked from our link"])
        # 2. interns work their list top to bottom
        for who, sheet in (("Ram", "RAM"), ("Shyam", "SHYAM")):
            lo, hi = SHIFT[who]
            if not (lo <= clock.hour * 60 + clock.minute < hi):
                continue
            t_local = to_abs(clock)
            budget_end = min(t_local + CYCLE, to_abs(DAY0 + dt.timedelta(days=1, minutes=hi)))
            for row in S(sheet).Range("A7:P31").Value:
                phone = row[2]
                if phone in (None, ""):
                    continue
                phone, code = str(phone), row[15]
                codes_seen[code] += 1
                items = []
                for x in str(row[3]).split("; "):
                    comp, rd = x.rsplit(" ", 1)
                    if (phone, comp, rd) not in booked:
                        items.append((comp, rd))
                if not items:
                    continue
                cost = float(rng.uniform(3, 6)) + 1.5 * (len(items) - 1)
                if t_local + cost > budget_end:
                    break
                t_local += cost
                when = DAY0 + dt.timedelta(minutes=t_local)
                stats["tries"] += 1
                stats["rows_with_last_try_label"] += str(row[6]).startswith("LAST TRY")
                n_before = len(lines)
                answered = rng.random() < min(1.0, ANS[cls[phone]] * sim2.tod_factor(t_local))
                if answered and rng.random() < CONV:
                    keep = [it for j, it in enumerate(items) if j == 0 or rng.random() < EXTRA]
                    drop = [it for it in items if it not in keep]
                    link_items = [it for it in keep if consent[it[0]] != "YES"]
                    book_items = [it for it in keep if consent[it[0]] == "YES"]
                    if link_items:          # the try line comes first; same-conversation lines follow within a minute
                        lines.append([when, who, phone, "SENT_LINK", "", "sent booking link"])
                        for comp, rd in link_items:
                            if rng.random() < P_LINK:
                                pending_link.append((when + dt.timedelta(minutes=float(rng.uniform(*LINK_DELAY))), phone, comp, rd))
                    if book_items and not link_items and not drop:
                        lines.append([when, who, phone, "BOOKED", "", ""])
                    else:
                        for comp, rd in book_items:
                            lines.append([when + dt.timedelta(seconds=20), who, phone, "BOOKED", comp, ""])
                    for comp, rd in book_items:
                        booked.add((phone, comp, rd))
                    for comp, rd in drop:
                        lines.append([when + dt.timedelta(seconds=40), who, phone, "NOT_INTERESTED", comp, "declined this one"])
                elif rng.random() < OPT:
                    lines.append([when, who, phone, "STOP", "", "asked to stop"])
                elif answered:
                    lines.append([when, who, phone, "CALL_BACK", "", "busy"])
                else:
                    lines.append([when, who, phone, "NO_ANSWER", "", ""])
                stats["intern_lines"] += len(lines) - n_before
                stats["multi_line_conversations"] += (len(lines) - n_before) > 1
        write_log(lines)
        if clock >= END:
            break
        clock = min(clock + dt.timedelta(minutes=CYCLE), END)
        recalc(clock)
        d = S("DASHBOARD")
        print(f"{clock:%H:%M}  in-24h={d.Range('B3').Value:.1%}  tries={d.Range('B14').Value:.0f}", flush=True)

    # read-out one minute after the deadline
    recalc(END + dt.timedelta(minutes=1))
    d, lg = S("DASHBOARD"), S("LOG")
    out = {"note": "Mechanics test. Candidate behaviour drawn from ASSUMED rates; the booking share is not evidence."}
    out["dashboard"] = {str(d.Cells(r, 1).Value).strip(): d.Cells(r, 2).Value for r in list(range(6, 17)) + list(range(19, 28))}
    out["dashboard"]["Scheduled within 24h (share)"] = d.Range("B3").Value
    out["capacity_and_cost"] = {str(d.Cells(r, 1).Value).strip(): d.Cells(r, 2).Value for r in range(30, 51) if d.Cells(r, 1).Value}
    logv = lg.Range(lg.Cells(2, 1), lg.Cells(max(2, next_row - 1), 10)).Value2     # Value2: dates as serial numbers
    checks = collections.Counter(str(r[9]) for r in logv if r[3])
    per_person = collections.Counter(str(r[2]) for r in logv if r[7] == 1)
    times = [dt.datetime(1899, 12, 30) + dt.timedelta(days=float(r[6])) for r in logv if r[7] == 1]
    out["log"] = {
        "lines": next_row - 2,
        "intern_lines": stats["intern_lines"], "bot_tries": stats["tries"], "sheet_tries": sum(per_person.values()),
        "intern_lines_per_try": round(stats["intern_lines"] / max(1, stats["tries"]), 3),
        "conversations_needing_more_than_one_line": stats["multi_line_conversations"],
        "check_column": dict(checks),
        "max_tries_per_person": max(per_person.values()) if per_person else 0,
        "tries_outside_10_to_21": sum(1 for t in times if not (10 <= t.hour < 21)),
    }
    out["list_codes_shown"] = dict(codes_seen)
    out["formula_errors_after_day"] = len(sheet_errors(wb))
    wb.Save()
finally:
    wb.Close(False)
    xl.Quit()
json.dump(out, open(os.path.join(OUTD, "dry_run_result.json"), "w"), indent=1, default=str)
print(json.dumps(out, indent=1, default=str))
