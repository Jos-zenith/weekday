"""
Path-to-80% scenarios and cost sensitivity. Deterministic arithmetic, no simulation, so every number can be
checked by hand. Usage: python analysis/scenarios.py  (writes analysis/scenarios.json)

Model of one evening batch worked the next day:
  * P people (R requests). A share `s` of people book by themselves (helped by the evening message) and need no try.
  * Everyone else with valid contacts gets try 1, then try 2 if not booked, then try 3 (max 3).
  * Try k ends in a CONFIRMED booking with probability b1 * d**(k-1) (booked by us, or by the candidate after our link;
    a link sent is not a booking). `d` says how fast later tries lose power (d=1: no loss).
  * Interns supply tries in order: all first tries, then second tries, then third (best tries first).
  * Requests booked = people booked x requests per person.
The three inputs that decide the outcome (s, b1, d) are all measured on day 1 by the DASHBOARD.
"""
import itertools
import json
import os

P, R = 426, 500                 # demo batch: people and unique requests
VALID = 0.97                    # share of people with usable contact details
MIN_PER_INTERN = 475            # minutes before the 19:51 deadline (10:00-18:00 and 12:00-19:51 average)
MIN_PER_TRY = 4.5
TARGET = 0.80
INTERN_DAY = 20000 / 22         # Rs


def outcome(interns, s, b1, d, min_per_try=MIN_PER_TRY):
    capacity = interns * MIN_PER_INTERN / min_per_try
    waiting = P * (1 - s) * VALID
    booked, tries = 0.0, 0.0
    for k in range(3):
        b = b1 * d ** k
        t = min(waiting, capacity - tries)
        if t <= 0:
            break
        tries += t
        booked += t * b
        waiting -= waiting * b          # everyone in this round who was not booked moves to the next round
        if t < waiting / (1 - b) - 1e-9:  # ran out of capacity inside this round
            break
    return min(1.0, (P * s + booked) / P), tries


def ceiling(s, b1, d):
    miss = 1.0
    for k in range(3):
        miss *= 1 - b1 * d ** k
    return min(1.0, s + (1 - s) * VALID * (1 - miss))


def interns_needed(s, b1, d):
    for n in range(2, 41):
        if outcome(n, s, b1, d)[0] >= TARGET:
            return n
    return None


S = [0.15, 0.25, 0.40]          # share of people who book alone (evening message pushes this up)
B1 = [0.30, 0.40]               # first-try confirmed-booking rate (planning value 0.40)
D = [1.0, 0.7, 0.4]             # later-try decay
rows = []
for s, b1, d in itertools.product(S, B1, D):
    rows.append(dict(s=s, b1=b1, d=d, two_interns=round(outcome(2, s, b1, d)[0], 3),
                     four=round(outcome(4, s, b1, d)[0], 3), six=round(outcome(6, s, b1, d)[0], 3),
                     ceiling=round(ceiling(s, b1, d), 3), interns_for_80=interns_needed(s, b1, d)))

print("Share scheduled within 24h (people). s = self-book share, b1 = try-1 confirmed-booking rate, d = later-try decay")
print(f"{'s':>5} {'b1':>5} {'d':>4} | {'2 int':>6} {'4 int':>6} {'6 int':>6} | {'ceiling':>7} | interns for 80%")
for r in rows:
    print(f"{r['s']:5.0%} {r['b1']:5.2f} {r['d']:4.1f} | {r['two_interns']:6.0%} {r['four']:6.0%} {r['six']:6.0%} | "
          f"{r['ceiling']:7.0%} | {r['interns_for_80'] or 'not reachable'}")

# ---- cost sensitivity: which decisions survive wrong assumptions?
cost_per_try = INTERN_DAY / (MIN_PER_INTERN / MIN_PER_TRY)
fin = []
for hire, ctc in itertools.product([0.01, 0.03, 0.05], [1_200_000, 1_800_000, 2_400_000]):
    value = 0.15 * ctc * hire
    fin.append(dict(hire=hire, ctc=ctc, value=round(value), breakeven_booking_per_try=round(cost_per_try / (value * R / P), 4)))
print(f"\ncost per try Rs {cost_per_try:.1f}")
print("hire  CTC(L)  value/interview  extra tries pay if more than 1 in N tries books")
for f in fin:
    print(f"{f['hire']:4.0%} {f['ctc']/1e5:6.0f}  Rs {f['value']:>7,}       1 in {1 / f['breakeven_booking_per_try']:,.0f}")

out = dict(inputs=dict(P=P, R=R, valid=VALID, min_per_intern=MIN_PER_INTERN, min_per_try=MIN_PER_TRY,
                       target=TARGET, intern_day=round(INTERN_DAY)),
           scenarios=rows, cost_per_try=round(cost_per_try, 2), finance=fin)
json.dump(out, open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "scenarios.json"), "w"), indent=1)
