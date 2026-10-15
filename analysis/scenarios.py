"""
How far can the desk get, and what would 80% take? Plain arithmetic with no randomness, so any number can be checked
by hand. Run: python analysis/scenarios.py  (writes analysis/scenarios.json)

One set of assumptions, used everywhere (the DASHBOARD MODEL block does the same arithmetic in formulas). Every share is
a share of the 500 unique requests, never of people. For one evening batch worked the next day:
- a share s of people book on their own (the evening message should push this up) and never need a call. It doesn't
  depend on how many requests they have, so s is also the share of requests booked without a call;
- everyone else with usable contacts gets up to three tries, all first tries before any second try. Try k ends in a
  confirmed booking with probability b1 * d**(k-1). Confirmed means we booked it, or they booked after we sent the
  link; the link alone doesn't count. d is how much each later try keeps (1 = all);
- each successful call settles PER_SUCCESS requests. The planning value is the batch average among people we can
  reach (489 requests / 418 people = 1.17). The simulated day settled 1.15, close to it. After day 1 the dashboard
  uses the measured value.
s, b1, d and PER_SUCCESS are exactly what the dashboard measures in the first two days, so real numbers replace these.

Outer bounds on the interns' ceiling, for reference only: 56% if calls reached the people with the most requests first
and settled all of them; 42% if each successful call settled only one request.
"""
import itertools
import json
import os

R = 500                                     # demo batch: unique requests
USABLE = {1: 356, 2: 53, 3: 9}              # people with usable contacts, by number of requests (489 requests)
UNUSABLE = {1: 6, 2: 1, 3: 1}               # people with a broken email or phone (11 requests): never called
P = sum(USABLE.values()) + sum(UNUSABLE.values())   # 426 people
PEOPLE = sum(USABLE.values())                        # 418 people we can call
PER_SUCCESS = sum(k * n for k, n in USABLE.items()) / PEOPLE   # 1.17 requests settled per successful call
MIN_PER_INTERN = 475            # minutes before the 19:51 deadline: Ram 10:00-18:00, Shyam 12:00 to the deadline
MIN_PER_TRY = 4.5
TARGET = 0.80
INTERN_DAY = 20000 / 22         # Rs
LOST = 0.25                     # share of requests that miss 24h and are then never interviewed (SETTINGS B29)
assert sum(k * n for k, n in USABLE.items()) + sum(k * n for k, n in UNUSABLE.items()) == R


def tries(interns):
    return interns * MIN_PER_INTERN / MIN_PER_TRY


def schedule(capacity, s, rates, people=PEOPLE, per_success=PER_SUCCESS, r=R):
    """(share scheduled with this many tries, share with unlimited tries). rates: chance tries 1, 2 and 3 book."""
    waiting, left, succ, succ_all = people * (1 - s), capacity, 0.0, 0.0
    for b in rates:
        x = min(waiting, max(0.0, left))
        succ += x * b
        succ_all += waiting * b
        waiting *= 1 - b                    # everyone not booked by this try is waiting for the next one
        left -= x
    return min(1.0, s + succ * per_success / r), min(1.0, s + succ_all * per_success / r)


def ceiling_if_every_try_books(capacity, people=PEOPLE, per_success=PER_SUCCESS, r=R):
    return min(capacity, people) * per_success / r


def bound_most_requests_first(capacity, groups=USABLE, r=R):
    """Outer bound: every try books, calls reach the people with the most requests first and settle all of them."""
    left, booked = capacity, 0.0
    for k in sorted(groups, reverse=True):
        t = min(left, groups[k])
        booked += t * k
        left -= t
    return booked / r


def bound_one_request_per_call(capacity, people=PEOPLE, r=R):
    return min(capacity, people) / r


def outcome(interns, s, b1, d):
    return schedule(tries(interns), s, [b1, b1 * d, b1 * d * d])[0]


def ceiling(s, b1, d):
    return schedule(0, s, [b1, b1 * d, b1 * d * d])[1]


def interns_needed(s, b1, d):
    for n in range(1, 41):
        if outcome(n, s, b1, d) >= TARGET:
            return n
    return None


if __name__ == "__main__":
    t2 = tries(2)
    print(f"Tries two interns can make: {t2:.1f}. Requests settled per successful call (planning): {PER_SUCCESS:.3f}.")
    print(f"Interns' ceiling if every try books: {ceiling_if_every_try_books(t2):.1%}. Outer bounds: "
          f"{bound_one_request_per_call(t2):.1%} (one request per call) to {bound_most_requests_first(t2):.1%} "
          f"(most requests first, all settled).\n")

    S = [0.15, 0.25, 0.40]          # share booked with no call (evening message pushes this up)
    B1 = [0.30, 0.40]               # first-try confirmed-booking rate (planning value 0.40)
    D = [1.0, 0.7, 0.4]             # later-try decay
    rows = []
    for s, b1, d in itertools.product(S, B1, D):
        rows.append(dict(s=s, b1=b1, d=d, two_interns=round(outcome(2, s, b1, d), 3),
                         four=round(outcome(4, s, b1, d), 3), six=round(outcome(6, s, b1, d), 3),
                         ceiling=round(ceiling(s, b1, d), 3), interns_for_80=interns_needed(s, b1, d)))
    print("Share of requests scheduled within 24h. s = booked with no call, b1 = try-1 confirmed-booking rate, d = later-try decay")
    print(f"{'s':>5} {'b1':>5} {'d':>4} | {'2 int':>6} {'4 int':>6} {'6 int':>6} | {'ceiling':>7} | interns for 80%")
    for r in rows:
        print(f"{r['s']:5.0%} {r['b1']:5.2f} {r['d']:4.1f} | {r['two_interns']:6.0%} {r['four']:6.0%} {r['six']:6.0%} | "
              f"{r['ceiling']:7.0%} | {r['interns_for_80'] or 'not reachable'}")

    # ---- cost sensitivity: which decisions survive wrong assumptions? A successful try settles PER_SUCCESS interviews.
    cost_per_try = INTERN_DAY / (MIN_PER_INTERN / MIN_PER_TRY)
    fin = []
    for hire, ctc in itertools.product([0.01, 0.03, 0.05], [1_200_000, 1_800_000, 2_400_000]):
        value = 0.15 * ctc * hire
        fin.append(dict(hire=hire, ctc=ctc, value=round(value),
                        breakeven_booking_per_try=round(cost_per_try / (value * PER_SUCCESS), 5)))
    print(f"\ncost per try Rs {cost_per_try:.1f}")
    print("hire  CTC(L)  value/interview  extra tries pay if more than 1 in N tries books")
    for f in fin:
        print(f"{f['hire']:4.0%} {f['ctc']/1e5:6.0f}  Rs {f['value']:>7,}       1 in {1 / f['breakeven_booking_per_try']:,.0f}")

    # ---- fee scenario: the gap between the two-intern planning estimate and the target
    base = outcome(2, 0.25, 0.40, 0.7)
    gap = (TARGET - base) * R
    risk = []
    for hire in (0.01, 0.03, 0.05):
        value = 0.15 * 1_800_000 * hire
        risk.append(dict(hire=hire, value=round(value), if_quarter_lost=round(gap * LOST * value),
                         if_all_lost=round(gap * value)))
    print(f"\nPlanning case (s 25%, b1 40%, d 0.7), two interns: {base:.1%}. Gap to 80%: {gap:.0f} requests a day.")
    print(f"SCENARIO, not a measured loss. Rs 18 lakh CTC, 15% fee; 'lost' = late requests never interviewed (planning {LOST:.0%}):")
    for x in risk:
        print(f"  hire {x['hire']:.0%}: value Rs {x['value']:,} / interview -> Rs {x['if_quarter_lost']:,} if {LOST:.0%} lost, "
              f"Rs {x['if_all_lost']:,} if all lost (upper case)")

    out = dict(inputs=dict(R=R, P=P, people_reachable=PEOPLE, usable=USABLE, unusable=UNUSABLE,
                           per_success=round(PER_SUCCESS, 4), min_per_intern=MIN_PER_INTERN, min_per_try=MIN_PER_TRY,
                           target=TARGET, intern_day=round(INTERN_DAY), lost=LOST),
               tries_two_interns=round(t2, 1), ceiling_two=round(ceiling_if_every_try_books(t2), 3),
               bounds_two=dict(one_request_per_call=round(bound_one_request_per_call(t2), 3),
                               most_requests_first=round(bound_most_requests_first(t2), 3)),
               scenarios=rows, cost_per_try=round(cost_per_try, 2), finance=fin,
               fees_at_risk=dict(two_intern_share=round(base, 3), gap_requests=round(gap), by_hire_rate=risk))
    json.dump(out, open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "scenarios.json"), "w"), indent=1)
