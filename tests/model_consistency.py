"""
Do DASHBOARD's MODEL block and analysis/scenarios.py give the same answers on the same batch?

1. The interns' ceiling, rebuilt from the raw demo batch (tests/results/demo-candidates.json): the request mix, the
   requests each successful call settles on average, and that no more tries are used than the interns have. Also the
   two outer bounds (one request per call; most requests first, all settled).
2. The workbook's own model, recalculated in Excel at the test clock, against scenarios.schedule() fed the workbook's
   exact inputs (reachable people, requests per successful call, unique requests, tries and rates), for the plan's six
   scenarios plus one where later tries book better than earlier ones.
Differences are reported, not hidden: the run passes only if every gap is within rounding (0.05 points, same interns).
Run from the repo root after build/workbook.py: python tests/model_consistency.py   (needs Excel on Windows)
"""
import collections
import json
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "analysis"))
import scenarios as sc  # noqa: E402
from excel import open_excel, sheet_errors  # noqa: E402

OUTD = os.path.join(HERE, "results")
SRC = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "Weekday Scheduling Desk.xlsx"))
DST = os.path.join(OUTD, "model-consistency.xlsx")
TOL = 0.0005
report = {"ceiling": {}, "dashboard_vs_scenarios": [], "notes": []}
ok = True

# ---- 1. the ceiling, from the raw batch
truth = json.load(open(os.path.join(OUTD, "demo-candidates.json")))
per_person = collections.Counter(r["phone"] for r in truth["req"])
usable = collections.Counter(k for p, k in per_person.items() if truth["cand"][p]["valid"])
unusable = collections.Counter(k for p, k in per_person.items() if not truth["cand"][p]["valid"])
n_req, people = len(truth["req"]), sum(usable.values())
per_success = sum(k * n for k, n in usable.items()) / people
cap = sc.tries(2)
called = min(cap, people)
left, most_first = cap, 0.0
for k in sorted(usable, reverse=True):
    t = min(left, usable[k])
    most_first += t * k
    left -= t
report["ceiling"] = dict(
    unique_requests=n_req, people=len(per_person), reachable_people=people, usable_mix=dict(sorted(usable.items())),
    unusable_mix=dict(sorted(unusable.items())), mix_matches_scenarios_py=dict(usable) == sc.USABLE and dict(unusable) == sc.UNUSABLE,
    requests_per_successful_call=round(per_success, 4), tries_available=round(cap, 1), tries_used=round(called, 1),
    ceiling=round(called * per_success / n_req, 4),
    outer_bound_one_request_per_call=round(called / n_req, 4),
    outer_bound_most_requests_first=round(most_first / n_req, 4),
    assumption="Every try books and settles the batch average of requests per reachable person. Outer bounds: one "
               "request per call, or calls reaching the people with the most requests first and settling all of them.")
c = report["ceiling"]
ok &= (c["mix_matches_scenarios_py"] and called <= cap
       and abs(called * per_success / n_req - sc.ceiling_if_every_try_books(cap)) < 1e-9
       and abs(most_first / n_req - sc.bound_most_requests_first(cap)) < 1e-9)

# ---- 2. the workbook's model vs scenarios.py, same batch, same inputs
CASES = [(0.25, 0.40, 1.0), (0.25, 0.40, 0.7), (0.25, 0.40, 0.4), (0.40, 0.40, 1.0), (0.40, 0.40, 0.7), (0.40, 0.40, 0.4),
         (0.25, 0.40, 1.2)]       # last: later tries book BETTER; tries must still happen in order
shutil.copyfile(SRC, DST)
xl = open_excel()
wb = xl.Workbooks.Open(DST)
try:
    st, d = wb.Worksheets("SETTINGS"), wb.Worksheets("DASHBOARD")
    v = lambda a: d.Range(a).Value
    for s, b1, dec in CASES:
        st.Range("B18").Value, st.Range("B19").Value, st.Range("B20").Value = b1, dec, s
        xl.CalculateFull()
        ppl = v("B47") + v("B48") + v("B49")
        r_, s_used, k_used = v("B6"), v("B41"), v("B58")
        rates = [v("B38"), v("B39"), v("B40")]
        tries, mins, per_try = v("B45"), v("B44"), v("B37")
        py_now, py_all = sc.schedule(tries, s_used, rates, ppl, k_used, r_)
        per_intern = (mins / 2) / per_try
        py_interns = next((n for n in range(1, 41)
                           if sc.schedule(n * per_intern, s_used, rates, ppl, k_used, r_)[0] >= sc.TARGET), None)
        dash_interns = v("B64") if isinstance(v("B64"), float) else None
        row = dict(
            case=dict(no_call=s, try1=b1, later_keep=dec),
            inputs_read=dict(unique_requests=r_, reachable_people=ppl, requests_per_successful_call=round(k_used, 4),
                             rates=[round(x, 4) for x in rates], tries=tries),
            expected_today=dict(dashboard=round(v("B59"), 4), scenarios_py=round(py_now, 4)),
            best_case=dict(dashboard=round(v("B60"), 4), scenarios_py=round(py_all, 4)),
            interns_for_80=dict(dashboard=dash_interns and int(dash_interns), scenarios_py=py_interns),
            ceiling=dict(dashboard=round(v("B65"), 4), scenarios_py=round(sc.ceiling_if_every_try_books(tries, ppl, k_used, r_), 4)),
            bound_most_first=dict(dashboard=round(v("B66"), 4), scenarios_py=round(sc.bound_most_requests_first(tries, sc.USABLE, r_), 4)),
            bound_one_per_call=dict(dashboard=round(v("B67"), 4), scenarios_py=round(sc.bound_one_request_per_call(tries, ppl, r_), 4)),
        )
        keys = ("expected_today", "best_case", "ceiling", "bound_most_first", "bound_one_per_call")
        gaps = [abs(row[k]["dashboard"] - row[k]["scenarios_py"]) for k in keys]
        row["max_gap_points"] = round(100 * max(gaps), 3)
        row["match"] = max(gaps) <= TOL and row["interns_for_80"]["dashboard"] == row["interns_for_80"]["scenarios_py"]
        ok &= row["match"]
        report["dashboard_vs_scenarios"].append(row)
    errs = sheet_errors(wb)
    report["formula_errors"] = len(errs)
    ok &= not errs
finally:
    wb.Close(False)
    xl.Quit()

# the plan's table uses scenarios.py with 475 minutes per intern (211.1 tries); the workbook rounds tries down to 211
plan = sc.outcome(2, 0.25, 0.40, 0.7)
dash = next(r for r in report["dashboard_vs_scenarios"] if r["case"]["later_keep"] == 0.7 and r["case"]["no_call"] == 0.25)
report["notes"].append(dict(
    what="Plan table (scenarios.py, 211.1 tries) vs workbook (211 whole tries), planning case",
    plan=round(plan, 4), workbook=dash["expected_today"]["dashboard"],
    gap_points=round(100 * abs(plan - dash["expected_today"]["dashboard"]), 3)))
report["ALL_MATCH"] = bool(ok)
json.dump(report, open(os.path.join(OUTD, "model-consistency.json"), "w"), indent=1)
print(json.dumps(report, indent=1))
sys.exit(0 if ok else 1)
