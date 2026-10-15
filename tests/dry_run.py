"""
Dry run: plays one full working day (14 Oct, 10:00 to the 19:51 deadline) through the workbook.

A small bot stands in for Ram and Shyam. It only looks at what their tabs show, does what column G says, and writes
LOG lines the way START HERE asks, with request IDs when an answer only covers some of a person's requests. How each
candidate behaves comes from tests/results/demo-candidates.json, which the simulation drew from assumed rates.

The result file has two parts that shouldn't be mixed up:
- mechanics_checks: pass/fail checks that the sheet does what it should. These are the point of the exercise.
- simulated_outcomes: booking shares produced by assumed behaviour. Fine as a sanity check, but they say nothing
  about real candidates.

Run: python tests/dry_run.py [workbook]      (writes to tests/results/)
     EVENING_MSG=0 python tests/dry_run.py   same day without the evening WhatsApp
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
import simulation  # noqa: E402
from excel import open_excel, sheet_errors  # noqa: E402

SRC = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "Weekday Scheduling Desk.xlsx"))
OUTD = os.path.join(HERE, "results")
EVENING = os.environ.get("EVENING_MSG", "1") == "1"
TAG = "" if EVENING else "-no-whatsapp"
DST = os.path.join(OUTD, f"dry-run{TAG}.xlsx")
shutil.copyfile(SRC, DST)

truth = json.load(open(os.path.join(OUTD, "demo-candidates.json")))
DAY0 = dt.datetime.fromisoformat(truth["day0"])
cls = {p: v["cls"] for p, v in truth["cand"].items()}
consent = truth["consent"]
req = {r["id"]: r for r in truth["req"]}                       # canonical request ID -> phone, company, round, self_min
ANS = np.array(simulation.BASE["ans_p"])
CONV, EXTRA, OPT = simulation.BASE["conv_p"], simulation.BASE["extra_conv"], simulation.BASE["opt_p"]
P_LINK, LINK_DELAY = 0.6, (15, 180)        # assumed: share who book after we send a link, and how long they take
REPLY_YES = [0.35, 0.15, 0.03]             # made up for the test: evening YES reply / booking from the link, by class
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
booked = {}                       # request ID -> (time, "us" | "candidate")
link_logged = {}                  # phone -> time of the first SENT_LINK line in LOG
pending_link = []                 # (time, phone, request ID): candidate books from a link we sent
stats = collections.Counter()
codes_seen = collections.Counter()
yes_phones, stop_phones = set(), set()
yes_first_position = {}


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


def evening_message():
    """Lead confirms opt-in, exports the send list, and pastes the provider's report back (day 0, ~20:00-22:00)."""
    settings.Range("B22").Value = "YES"
    xl.CalculateFull()
    report = []
    for row in S("MESSAGES").Range("A9:H608").Value:
        if row[1] in (None, ""):
            continue
        ph, template, rid = str(row[1])[-10:], row[2], row[7]
        stats["sent"] += 1
        stats[f"template_{template}"] += 1
        stats["template_matches_consent"] += (template == "A") == (consent[req[rid]["company"]] == "YES")
        stats["template_B_has_link"] += template == "B" and str(row[5]).startswith("https://")
        u, reply = rng.random(), ""
        if u < REPLY_STOP:
            reply = "STOP"
            stop_phones.add(ph)
        elif u < REPLY_STOP + REPLY_YES[cls[ph]]:
            if template == "A":
                reply = "Yes please"
                yes_phones.add(ph)
            else:                     # books the request the link is for, by its ID
                pending_link.append((DAY0 + dt.timedelta(hours=20, minutes=float(rng.uniform(5, 120))), ph, rid))
        # alternate phone formats to test that the sheet matches "+91 ..." and plain numbers alike
        report.append(("+91 " + ph if len(report) % 2 else ph, "DELIVERED", reply, ""))
    if report:
        S("MESSAGES").Range(f"I9:L{8 + len(report)}").Value = tuple(report)


def dashboard():
    """Column A label -> column B value, for every labelled row on DASHBOARD."""
    v = S("DASHBOARD").Range("A1:E90").Value
    return {str(r[0]).strip(): r[1] for r in v if r[0]}


try:
    if EVENING:
        evening_message()
    clock = START
    recalc(clock)
    while True:
        lines = []
        # 1. bookings candidates made themselves (the lead copies Calendly confirmations into LOG, by request ID)
        for rid, r in req.items():
            sm = r["self_min"]
            if sm is not None and rid not in booked and DAY0 + dt.timedelta(minutes=sm) <= clock:
                t = DAY0 + dt.timedelta(minutes=sm)
                booked[rid] = (t, "candidate")
                lines.append([t, "Lead", r["phone"], "SELF_BOOKED", rid, "calendly confirmation"])
        for item in list(pending_link):
            t, ph, rid = item
            if t <= clock:
                pending_link.remove(item)
                if rid not in booked:
                    booked[rid] = (t, "candidate")
                    lines.append([t, "Lead", ph, "SELF_BOOKED", rid, "booked from our link"])
        write_log(lines)              # the lead logs confirmations first, so the interns' lists are current
        recalc(clock)
        lines = []
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
                ids = [x.split(" ", 1)[0] for x in str(row[3]).split("; ")]
                stats["listed_ids_not_this_persons"] += sum(req.get(i, {}).get("phone") != phone for i in ids)
                items = [i for i in ids if i not in booked]
                stats["listed_ids_already_booked"] += len(ids) - len(items)
                if not items:
                    continue
                cost = (2.0 if replied_yes else float(rng.uniform(3, 6))) + 1.5 * (len(items) - 1)
                if t_local + cost > budget_end:
                    break
                t_local += cost
                when = DAY0 + dt.timedelta(minutes=t_local)
                stats["tries"] += 1
                n_before = len(lines)
                answered = replied_yes or rng.random() < min(1.0, ANS[cls[phone]] * simulation.tod_factor(t_local))
                if replied_yes or (answered and rng.random() < CONV):
                    keep = [i for j, i in enumerate(items) if j == 0 or rng.random() < EXTRA]
                    drop = [i for i in items if i not in keep]
                    link_items = [i for i in keep if consent[req[i]["company"]] != "YES"]
                    book_items = [i for i in keep if consent[req[i]["company"]] == "YES"]
                    if link_items:          # the try line comes first; same-conversation lines follow within a minute
                        link_logged.setdefault(phone, when)
                        lines.append([when, who, phone, "SENT_LINK", "", "sent booking link"])
                        for i in link_items:
                            if rng.random() < P_LINK:
                                pending_link.append((when + dt.timedelta(minutes=float(rng.uniform(*LINK_DELAY))), phone, i))
                    if book_items and not link_items and not drop:
                        lines.append([when, who, phone, "BOOKED", "", ""])        # blank ID = all their open requests
                    else:
                        for i in book_items:
                            lines.append([when + dt.timedelta(seconds=20), who, phone, "BOOKED", i, ""])
                    for i in book_items:
                        booked[i] = (when, "us")
                    for i in drop:
                        lines.append([when + dt.timedelta(seconds=40), who, phone, "NOT_INTERESTED", i, "declined this one"])
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
        print(f"{clock:%H:%M}  in-24h={d.Range('B3').Value:.1%}  tries={d.Range('B17').Value:.0f}", flush=True)

    # ---- read-out one minute after the deadline
    recalc(END + dt.timedelta(minutes=1))
    dash = dashboard()
    dv = S("DESK").Range("A2:AY601").Value
    # DESK columns (0-based): L=11 unique, M=12 booked, AP=41 by us, AW=48 request ID, AX=49 by candidate after our link
    sheet_booked = {r[48]: ("us" if r[41] == 1 else ("link" if r[49] == 1 else "self"))
                    for r in dv if r[48] and r[11] == 1 and r[12] == 1}
    after_link = lambda i, t: req[i]["phone"] in link_logged and link_logged[req[i]["phone"]] < t
    bot_booked = {i: ("us" if how == "us" else ("link" if after_link(i, t) else "self"))
                  for i, (t, how) in booked.items() if t <= END}
    lg = S("LOG")
    logv = lg.Range(lg.Cells(2, 1), lg.Cells(max(2, next_row - 1), 13)).Value2     # Value2: dates as serial numbers
    checks = collections.Counter(str(r[9]) for r in logv if r[3])
    per_person = collections.Counter(str(r[2]) for r in logv if r[7] == 1)
    times = [EPOCH + dt.timedelta(days=float(r[6])) for r in logv if r[7] == 1]
    errors = len(sheet_errors(wb))
    mech = {
        "every_log_line_checks_OK": dict(checks),
        "requests_booked_in_sheet_equal_requests_the_bot_booked": sheet_booked.keys() == bot_booked.keys(),
        "who_booked_matches_for_every_request": all(sheet_booked.get(i) == h for i, h in bot_booked.items()),
        "requests_booked": len(bot_booked),
        "sent_links_counted_as_bookings": sum(1 for i in sheet_booked if i not in bot_booked),
        "list_showed_another_persons_request_ID": stats["listed_ids_not_this_persons"],
        "list_showed_an_already_booked_request": stats["listed_ids_already_booked"],
        "request_ID_problems_on_dashboard": dash.get("Request ID problems (missing, or used twice)"),
        "sheet_tries_equal_bot_tries": sum(per_person.values()) == stats["tries"],
        "max_tries_per_person": max(per_person.values()) if per_person else 0,
        "tries_outside_10_to_21": sum(1 for t in times if not (10 <= t.hour < 21)),
        "log_lines_per_conversation": round(stats["intern_lines"] / max(1, stats["tries"]), 3),
        "formula_errors_after_day": errors,
    }
    if EVENING:
        mech.update({
            "evening_send_list_rows": stats["sent"],
            "template_matches_partner_consent": f'{stats["template_matches_consent"]} of {stats["sent"]}',
            "template_B_rows_with_link": f'{stats["template_B_has_link"]} of {stats["template_B"]}',
            "STOP_repliers_ever_listed": stats["stop_repliers_listed"],
            "YES_label_without_a_YES_reply": stats["yes_label_without_yes_reply"],
            "YES_repliers_first_list_positions": sorted(yes_first_position.values()),
        })
    mech["ALL_PASS"] = bool(
        list(checks) == ["OK"] and mech["requests_booked_in_sheet_equal_requests_the_bot_booked"]
        and mech["who_booked_matches_for_every_request"] and mech["sent_links_counted_as_bookings"] == 0
        and mech["list_showed_another_persons_request_ID"] == 0 and mech["list_showed_an_already_booked_request"] == 0
        and mech["request_ID_problems_on_dashboard"] == 0 and mech["sheet_tries_equal_bot_tries"]
        and mech["max_tries_per_person"] <= 3 and mech["tries_outside_10_to_21"] == 0 and errors == 0
        and mech.get("STOP_repliers_ever_listed", 0) == 0 and mech.get("YES_label_without_a_YES_reply", 0) == 0)
    sim = {k: dash.get(k) for k in (
        "LIVE  Scheduled within 24h (share of all unique requests)", "Scheduled within 24h", "booked by an intern (BOOKED)",
        "booked by the candidate after we sent a link", "booked by the candidate on their own",
        "Links sent (SENT_LINK lines; not bookings)", "Tries made", "Expected share with today's staff",
        "Best case with unlimited interns (3 tries each)", "Interns needed (same shift length as today)", "Verdict")}
    rd = S("DASHBOARD").Range("A37:E42").Value
    sim["planning_vs_measured (used, planning, measured, n)"] = {r[0]: list(r[1:]) for r in rd}
    if EVENING:
        m = S("MESSAGES")
        sim["evening_pilot_readout (group, requests, scheduled, share, delivered, failed, YES, STOP)"] = \
            [[m.Cells(r, c).Value for c in range(1, 9)] for r in (6, 7)]
    out = {
        "evidence_type": "Synthetic dry run on made-up data. mechanics_checks show whether the sheet works. "
                         "simulated_outcomes come from assumed candidate behaviour, so they say nothing about real candidates.",
        "evening_message": EVENING,
        "mechanics_checks": mech,
        "simulated_outcomes": sim,
    }
    wb.Save()
finally:
    wb.Close(False)
    xl.Quit()
json.dump(out, open(os.path.join(OUTD, f"dry-run{TAG}.json"), "w"), indent=1, default=str)
print(json.dumps(out, indent=1, default=str))
