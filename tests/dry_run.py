"""
Dry run: drive the workbook through one simulated working day (14 Oct 10:00 -> deadline 19:51).

A bot plays Ram and Shyam. It reads ONLY what the RAM / SHYAM tabs show, follows column G, and writes LOG lines
exactly as START HERE tells an intern to. Candidate behaviour comes from hidden truth (tests/output/truth.json),
drawn from the analysis model's ASSUMED rates. So this tests the sheet's mechanics; its booking rate is NOT
evidence about real candidates and must not be quoted as one.

Usage: python tests/dry_run.py [workbook]   (default: Weekday_Scheduling_Desk.xlsx; writes tests/output/)
       EVENING_MSG=0 python tests/dry_run.py   runs the day without the evening WhatsApp pilot
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
TAG = "" if os.environ.get("EVENING_MSG", "1") == "1" else "_no_message"
DST = os.path.join(OUTD, f"dry_run_day1{TAG}.xlsx")
shutil.copyfile(SRC, DST)

truth = json.load(open(os.path.join(OUTD, "truth.json")))
DAY0 = dt.datetime.fromisoformat(truth["day0"])
cls = {p: v["cls"] for p, v in truth["cand"].items()}
consent = truth["consent"]
self_min = {(r["phone"], r["company"], r["round"]): r["self_min"] for r in truth["req"]}
ANS = np.array(sim2.BASE["ans_p"])
CONV, EXTRA, OPT = sim2.BASE["conv_p"], sim2.BASE["extra_conv"], sim2.BASE["opt_p"]
P_LINK, LINK_DELAY = 0.6, (15, 180)        # ASSUMED: share who book after we send a link, and how long they take
# Evening WhatsApp pilot (MESSAGES tab). Reply rates are INVENTED for the mechanics test, by candidate class:
EVENING = os.environ.get("EVENING_MSG", "1") == "1"
REPLY_YES = [0.35, 0.15, 0.03]             # replies YES (template A) or books from the link (template B)
REPLY_STOP = 0.01
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
yes_phones, stop_phones = set(), set()
yes_first_position = {}           # list position where a YES replier first appeared


def evening_message():
    """Lead confirms opt-in, exports the send list, and pastes the provider's report back (day 0, ~20:00-22:00)."""
    settings.Range("B22").Value = "YES"
    xl.CalculateFull()
    report = []
    for row in S("MESSAGES").Range(f"A9:G{8 + 600}").Value:
        if row[1] in (None, ""):
            continue
        ph, template, company = str(row[1])[-10:], row[2], str(row[4]).split(" and ")[0]
        stats["sent"] += 1
        stats[f"template_{template}"] += 1
        stats["template_matches_consent"] += (template == "A") == (consent[company] == "YES")
        stats["template_B_has_link"] += template == "B" and str(row[5]).startswith("https://")
        u, reply = rng.random(), ""
        if u < REPLY_STOP:
            reply = "STOP"
            stop_phones.add(ph)
        elif u < REPLY_STOP + REPLY_YES[cls[ph]]:
            if template == "A":
                reply = "Yes please"
                yes_phones.add(ph)
            else:
                for (p2, comp, rd) in self_min:
                    if p2 == ph and comp == company:
                        pending_link.append((DAY0 + dt.timedelta(hours=20, minutes=float(rng.uniform(5, 120))), ph, comp, rd))
        # alternate phone formats to test that the sheet matches "+91 ..." and plain numbers alike
        report.append(("+91 " + ph if len(report) % 2 else ph, "DELIVERED", reply, ""))
    if report:
        S("MESSAGES").Range(f"I9:L{8 + len(report)}").Value = tuple(report)


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
    if EVENING:
        evening_message()
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
            for pos, row in enumerate(S(sheet).Range("A7:P31").Value, start=1):
                phone = row[2]
                if phone in (None, ""):
                    continue
                phone, code = str(phone), row[15]
                codes_seen[code] += 1
                replied_yes = str(row[6]).startswith("REPLIED YES")
                stats["stop_repliers_listed"] += phone in stop_phones
                stats["yes_label_without_yes_reply"] += replied_yes and phone not in yes_phones
                if phone in yes_phones:
                    yes_first_position.setdefault(phone, pos)
                items = []
                for x in str(row[3]).split("; "):
                    comp, rd = x.rsplit(" ", 1)
                    if (phone, comp, rd) not in booked:
                        items.append((comp, rd))
                if not items:
                    continue
                cost = (2.0 if replied_yes else float(rng.uniform(3, 6))) + 1.5 * (len(items) - 1)
                if t_local + cost > budget_end:
                    break
                t_local += cost
                when = DAY0 + dt.timedelta(minutes=t_local)
                stats["tries"] += 1
                stats["rows_with_last_try_label"] += str(row[6]).startswith("LAST TRY")
                n_before = len(lines)
                answered = replied_yes or rng.random() < min(1.0, ANS[cls[phone]] * sim2.tod_factor(t_local))
                if replied_yes or (answered and rng.random() < CONV):
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
    out["readout_used_planning_measured_n"] = {str(d.Cells(r, 1).Value): [d.Cells(r, c).Value for c in (2, 3, 4, 5)]
                                               for r in range(31, 36)}
    out["capacity_and_cost"] = {str(d.Cells(r, 1).Value).strip(): d.Cells(r, 2).Value
                                for r in list(range(38, 41)) + list(range(45, 51)) + list(range(53, 64)) if d.Cells(r, 1).Value}
    m = S("MESSAGES")
    out["evening_message"] = {
        "enabled": EVENING, "sent": stats["sent"], "template_A": stats["template_A"], "template_B": stats["template_B"],
        "template_matches_partner_consent": stats["template_matches_consent"],
        "template_B_rows_with_link": stats["template_B_has_link"],
        "replied_yes": len(yes_phones), "replied_stop": len(stop_phones),
        "stop_repliers_ever_listed": stats["stop_repliers_listed"],
        "yes_label_shown_without_yes_reply": stats["yes_label_without_yes_reply"],
        "yes_repliers_first_list_position": sorted(yes_first_position.values()),
        "pilot_readout": [[m.Cells(r, c).Value for c in range(1, 9)] for r in (6, 7)],
        "note": "Reply rates are invented; this checks that replies flow into the lists, not how many will reply.",
    }
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
json.dump(out, open(os.path.join(OUTD, f"dry_run_result{TAG}.json"), "w"), indent=1, default=str)
print(json.dumps(out, indent=1, default=str))
