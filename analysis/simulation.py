"""
Simulation of a day at the scheduling desk (second version).

What it models:
- Requests land in one batch around 19:51 and the first automated message goes out straight away. Candidates only
  act between 07:00 and 23:00, so nobody self-books overnight. Interns work 10-18 and 12-20. Each request's
  deadline is its own arrival time + 24h.
- A request is (candidate, company, round, requisition). The same key twice is a partner re-sending a record and is
  dropped before any policy runs. The same candidate at two companies is two real requests.
- One denominator for the headline: all unique requests, including ones with bad contact details. A "reachable only"
  figure is reported separately.
- "Low-reach" candidates are just hard to reach (12% pick-up per try), not impossible. Only bad contact details are
  impossible. The ceiling comes out of a run with unlimited interns rather than being assumed.
- Booking several companies in one conversation costs extra minutes, and each booking can still fail.
- Opt-outs and attempts are counted.
- There is also a within-day randomised trial (by candidate, stratified by partner) for sizing a pilot.

In this folder it is only used to generate the demo data and the dry run's hidden truth.
"""
import json
import math
import sys
import time

import numpy as np

DAY = 1440
MAXA = 3
GAP = 120

BASE = dict(
    N=500, n_comp=100, zipf=1.1,
    dup_p=[0.85, 0.12, 0.03],        # candidate holds 1/2/3 requests, always at DIFFERENT companies
    resend_p=0.04,                   # partners re-sending an identical record (true duplicate)
    invalid_p=0.03,
    warm_spread=0.0,                 # >0: partners differ in how warm their candidates are
    tiebreak="random",              # FCFS order among simultaneous arrivals: random|sheet|partner_big|partner_small|partner_rand                  # candidates whose contact details are unusable (impossible)
    prior=[0.35, 0.40, 0.25],        # engaged / passive / low-reach (probabilistic, NOT impossible)
    round_p=[0.975, 0.016, 0.009],
    arrive="batch", arrive_clock=19 * 60 + 51,   # "spread": uniform over 10:00-20:00
    first_touch_office=None,         # e.g. (480,1080): auto-touch only inside 08:00-18:00
    awake=(420, 1380),               # candidates act only 07:00-23:00
    self_p=[0.60, 0.20, 0.03], self_mult=1.0, self_median=300, self_sigma=0.9,   # awake-minutes
    ans_p=[0.70, 0.45, 0.12], conv_p=0.75, extra_conv=0.85,
    sig=[[0.60, 0.30, 0.10], [0.10, 0.50, 0.40], [0.00, 0.20, 0.80]],
    cost=(3.0, 6.0), extra_cost=1.5, opt_p=0.015,
    shifts=[(600, 1080), (720, 1200)],            # Ram 10-18, Shyam 12-20, on day 0 and day 1
    rbonus=[1.0, 1.3, 1.6],
)

POLICIES = ["random", "fcfs", "round", "quick", "batched", "composite", "oracle", "anti"]
GROUPED = {"batched", "composite", "oracle"}
LABEL = {
    "random": "Random order",
    "fcfs": "FCFS (baseline)",
    "round": "Round-first",
    "quick": "Quick-win first (signals)",
    "batched": "FCFS + candidate batching",
    "composite": "Composite (batching + value/min)",
    "oracle": "Oracle ceiling",
    "anti": "Worst-first floor",
}


def tod_factor(t):
    c = t % DAY
    if 600 <= c < 720:
        return 0.8
    if 720 <= c < 840:
        return 1.0
    if 840 <= c < 1020:
        return 0.8
    if 1020 <= c < 1200:
        return 1.15
    if 1200 <= c < 1260:
        return 1.1
    return 0.0


def awake_advance(start, d, lo, hi):
    """Move `d` awake-minutes forward from `start` (absolute minutes), skipping the sleep gap."""
    t, rem = start.astype(float).copy(), d.astype(float).copy()
    for _ in range(8):
        act = rem > 1e-9
        if not act.any():
            break
        day = np.floor(t / DAY)
        c = t - day * DAY
        t = np.where(act & (c < lo), day * DAY + lo, t)
        t = np.where(act & (c >= hi), (day + 1) * DAY + lo, t)
        c = t - np.floor(t / DAY) * DAY
        use = np.where(act, np.minimum(rem, hi - c), 0.0)
        t = t + use
        rem = rem - use
    return t


def make_world(rng, cfg):
    N, C = cfg["N"], cfg["n_comp"]
    z = 1.0 / np.arange(1, C + 1) ** cfg["zipf"]
    z /= z.sum()
    sizes, tot = [], 0
    while tot < N:
        k = min(int(rng.choice([1, 2, 3], p=cfg["dup_p"])), N - tot)
        sizes.append(k)
        tot += k
    nc = len(sizes)
    cand, comp = [], []
    for c, k in enumerate(sizes):
        cand += [c] * k
        comp += list(rng.choice(C, size=k, replace=False, p=z))
    perm = rng.permutation(N)
    cand, comp = np.array(cand)[perm], np.array(comp)[perm]
    rnd = rng.choice(3, size=N, p=cfg["round_p"])
    req = rng.integers(0, 3, size=N)

    # raw rows incl. partner re-sends, then key-based removal of true duplicates
    resend = np.flatnonzero(rng.random(N) < cfg["resend_p"])
    rc = np.concatenate([cand, cand[resend]])
    rp = np.concatenate([comp, comp[resend]])
    rr = np.concatenate([rnd, rnd[resend]])
    rq = np.concatenate([req, req[resend]])
    key = ((rc * C + rp) * 3 + rr) * 3 + rq
    hygiene = dict(raw=int(len(key)), unique=int(len(np.unique(key))),
                   removed=int(len(key) - len(np.unique(key))),
                   multi_company_candidates=int(sum(1 for s in sizes if s > 1)))

    lead = np.zeros(nc, dtype=int)
    seen = np.zeros(nc, dtype=bool)
    for r_ in range(N):
        c_ = cand[r_]
        if not seen[c_]:
            seen[c_] = True
            lead[c_] = comp[r_]
    if cfg["warm_spread"] > 0:
        wv = rng.uniform(-1, 1, size=C) * cfg["warm_spread"]
        pr = np.array(cfg["prior"])[None, :] * np.exp(wv[lead][:, None] * np.array([1.0, 0.0, -1.0])[None, :])
        pr /= pr.sum(axis=1, keepdims=True)
        cls = (rng.random(nc)[:, None] > np.cumsum(pr, axis=1)).sum(axis=1).clip(0, 2)
    else:
        cls = rng.choice(3, size=nc, p=cfg["prior"])
    sig_cum = np.cumsum(np.array(cfg["sig"])[cls], axis=1)
    sig = (rng.random(nc)[:, None] > sig_cum).sum(axis=1).clip(0, 2)
    valid = rng.random(nc) >= cfg["invalid_p"]

    if cfg["arrive"] == "batch":
        arr = np.full(N, float(cfg["arrive_clock"]))
    else:
        arr = rng.uniform(600, 1200, size=N)
    start = arr.copy()
    if cfg["first_touch_office"] is not None:
        lo, hi = cfg["first_touch_office"]
        clk = arr % DAY
        day = np.floor(arr / DAY)
        start = np.where((clk >= lo) & (clk < hi), arr,
                         np.where(clk < lo, day * DAY + lo, (day + 1) * DAY + lo))
    d = rng.lognormal(math.log(cfg["self_median"]), cfg["self_sigma"], size=N)
    t_self = awake_advance(start, d, *cfg["awake"])
    sp = np.array(cfg["self_p"])[cls[cand]] * cfg["self_mult"]
    flag = (rng.random(N) < np.clip(sp, 0, 1)) & valid[cand]
    self_time = np.where(flag & (t_self <= arr + DAY), t_self, np.inf)

    w = dict(
        n=N, nc=nc, cand_of=cand, comp=comp, cls=cls, sig=sig, valid=valid, rnd=rnd,
        arr=arr, self_time=self_time, hygiene=hygiene,
        u_ans_r=rng.random((N, MAXA)), u_conv_r=rng.random((N, MAXA)), u_opt_r=rng.random((N, MAXA)),
        u_ans_c=rng.random((nc, MAXA)), u_conv_c=rng.random((nc, MAXA)), u_opt_c=rng.random((nc, MAXA)),
        u_extra=rng.random(N), u_rand=rng.random(N), cost=rng.uniform(*cfg["cost"], size=800),
    )
    w["rank"] = make_rank(rng, w, cfg)
    return w


def make_rank(rng, w, cfg):
    """Tie-break rank among simultaneous arrivals (0 = served first). Reproducible via the seed."""
    n, tb = w["n"], cfg["tiebreak"]
    jit = rng.random(n)
    cnt = np.bincount(w["comp"])
    if tb == "random":
        key = jit
    elif tb == "sheet":
        key = np.arange(n, dtype=float)
    elif tb == "partner_big":
        key = -cnt[w["comp"]] * 10.0 + w["comp"] * 1e-3 + jit * 1e-4
    elif tb == "partner_small":
        key = cnt[w["comp"]] * 10.0 + w["comp"] * 1e-3 + jit * 1e-4
    else:  # partner_rand
        key = rng.permutation(len(cnt))[w["comp"]] * 10.0 + jit
    return np.argsort(np.argsort(key)).astype(float)


def run(w, policy, cfg, detail=False):
    n, nc = w["n"], w["nc"]
    grouped = policy in GROUPED
    ans_p, sigm, rb = np.array(cfg["ans_p"]), np.array(cfg["sig"]), np.array(cfg["rbonus"])
    cand_of = w["cand_of"]
    if grouped:
        unit, nu, ucand = cand_of, nc, np.arange(nc)
        u_ans, u_conv, u_opt = w["u_ans_c"], w["u_conv_c"], w["u_opt_c"]
    else:
        unit, nu, ucand = np.arange(n), n, cand_of
        u_ans, u_conv, u_opt = w["u_ans_r"], w["u_conv_r"], w["u_opt_r"]
    ucls, usig, uvalid = w["cls"][ucand], w["sig"][ucand], w["valid"][ucand]
    arr = w["arr"]
    dl = arr + DAY
    urow = np.full(nu, float(n))
    np.minimum.at(urow, unit, w["rank"])
    uarr = np.full(nu, np.inf)
    np.minimum.at(uarr, unit, arr)
    uround = np.zeros(nu, dtype=int)
    np.maximum.at(uround, unit, w["rnd"])
    urand = np.zeros(nu)
    np.maximum.at(urand, unit, w["u_rand"])

    sched = w["self_time"].copy()
    attempts = np.zeros(nu, dtype=int)
    next_el = np.zeros(nu)
    enq = np.zeros(nu)
    dead = np.zeros(nc, dtype=bool)
    ctouch = np.zeros(nc, dtype=int)
    post = np.array(cfg["prior"])[None, :] * sigm[:, usig].T
    post /= post.sum(axis=1, keepdims=True)
    post0 = post.copy()

    iv = [[day * DAY + s, day * DAY + e] for day in (0, 1) for s, e in cfg["shifts"]]
    touches, minutes = 0, 0.0
    while True:
        open_iv = [v for v in iv if v[0] < v[1]]
        if not open_iv:
            break
        v = min(open_iv, key=lambda x: x[0])
        t = v[0]
        live = (sched > t) & (arr <= t) & (dl > t)
        krem = np.bincount(unit[live], minlength=nu)
        ok_u = (attempts < MAXA) & (~dead[ucand])
        pend = (krem > 0) & ok_u
        elig = pend & (next_el <= t)
        if not elig.any():
            wake = []
            if pend.any():
                wake.append(float(next_el[pend].min()))
            fut = (sched > t) & (arr > t) & (dl > t) & ok_u[unit]
            if fut.any():
                wake.append(float(arr[fut].min()))
            v[0] = v[1] if not wake else max(t + 1e-6, min(wake))
            continue
        ready = np.maximum(enq, uarr)
        if policy == "random":
            sc = urand
        elif policy in ("fcfs", "batched"):
            sc = -(ready * 1e4 + urow)
        elif policy == "round":
            sc = uround * 1e12 - ready * 1e4 - urow
        elif policy == "quick":
            sc = post0 @ ans_p - 1e-9 * urow
        elif policy == "composite":
            sc = krem * (post @ ans_p) * rb[uround] - 1e-9 * urow
        elif policy == "oracle":
            sc = krem * ans_p[ucls] * rb[uround] - 1e-9 * urow
        else:
            sc = -ans_p[ucls] - 1e-9 * urow
        u = int(np.argmax(np.where(elig, sc, -np.inf)))
        c = float(w["cost"][touches % len(w["cost"])])
        if grouped:
            c += cfg["extra_cost"] * (int(krem[u]) - 1)
        te = t + c
        if te > v[1]:
            v[0] = v[1]
            continue
        touches += 1
        minutes += c
        ctouch[ucand[u]] += 1
        a = attempts[u]
        attempts[u] += 1
        f = tod_factor(t)
        if not uvalid[u]:
            answered = False
            attempts[u] = MAXA          # bounce / wrong number: moves to failed-contacts queue
        else:
            answered = u_ans[u, a] < min(1.0, ans_p[ucls[u]] * f)
        next_el[u] = te + GAP
        enq[u] = te
        if answered and u_conv[u, a] < cfg["conv_p"]:
            members = np.flatnonzero((unit == u) & live)
            for j, r in enumerate(members):
                if (j == 0 or w["u_extra"][r] < cfg["extra_conv"]) and te <= dl[r] and sched[r] > te:
                    sched[r] = te
        else:
            if u_opt[u, a] < cfg["opt_p"]:
                dead[ucand[u]] = True
            if not answered:
                post[u] *= (1 - ans_p * f)
                post[u] /= post[u].sum()
        v[0] = te

    ok = sched <= dl
    human = ok & (sched < w["self_time"])
    vr = w["valid"][cand_of]
    tc = ctouch[ctouch > 0]
    res = dict(
        pct=ok.mean(),
        pct_reach=ok[vr].mean(),
        auto_pct=np.isfinite(w["self_time"]).mean(),
        med_t=float(np.median((sched - arr)[ok])) if ok.any() else float("nan"),
        touches=float(touches),
        mins_per_book=minutes / max(int(human.sum()), 1),
        human=float(human.sum()),
        att_per_cand=float(tc.mean()) if len(tc) else 0.0,
        optout=float(dead.mean()),
        r23=float(ok[w["rnd"] > 0].mean()) if (w["rnd"] > 0).any() else float("nan"),
        lowreach=float(ok[w["cls"][cand_of] == 2].mean()),
        engaged=float(ok[w["cls"][cand_of] == 0].mean()),
        passive=float(ok[w["cls"][cand_of] == 1].mean()),
        invalid_share=float((~vr).mean()),
    )
    if detail:
        res["cand_s"] = np.bincount(cand_of[ok], minlength=nc)
        res["cand_k"] = np.bincount(cand_of, minlength=nc)
    return res


def study(cfg, policies, reps, seed):
    out = {p: {} for p in policies}
    hyg = []
    for r in range(reps):
        w = make_world(np.random.default_rng([seed, r]), cfg)
        hyg.append(w["hygiene"])
        for p in policies:
            for k, val in run(w, p, cfg).items():
                out[p].setdefault(k, []).append(val)
    res = {p: {k: np.array(v) for k, v in d.items()} for p, d in out.items()}
    return res, hyg


def paired(a, b, rng, B=3000):
    d = a - b
    bs = rng.choice(d, size=(B, len(d)), replace=True).mean(axis=1)
    lo, hi = np.percentile(bs, [2.5, 97.5])
    return dict(mean=float(d.mean()), lo=float(lo), hi=float(hi), win=float((d > 0).mean()),
                sd=float(d.std(ddof=1)))


# ---- trial simulation: within-day, candidate-clustered, partner-stratified randomisation -------
def subset(w, cmask):
    rmask = cmask[w["cand_of"]]
    cmap = -np.ones(w["nc"], dtype=int)
    cmap[cmask] = np.arange(cmask.sum())
    s = {k: v for k, v in w.items()}
    s.update(n=int(rmask.sum()), nc=int(cmask.sum()), cand_of=cmap[w["cand_of"][rmask]])
    for k in ("cls", "sig", "valid", "u_ans_c", "u_conv_c", "u_opt_c"):
        s[k] = w[k][cmask]
    for k in ("rnd", "arr", "self_time", "u_ans_r", "u_conv_r", "u_opt_r", "u_extra", "u_rand", "comp", "rank"):
        s[k] = w[k][rmask]
    return s


def assign_arms(w, rng):
    """Cluster = candidate. Stratify by (size tercile of the candidate's lead partner, has R2/R3)."""
    nc = w["nc"]
    comp_cnt = np.bincount(w["comp"], minlength=w["comp"].max() + 1)
    lead = np.full(nc, -1)
    for r in range(w["n"] - 1, -1, -1):
        lead[w["cand_of"][r]] = w["comp"][r]
    size = comp_cnt[lead]
    q = np.quantile(size, [1 / 3, 2 / 3])
    terc = (size > q[0]).astype(int) + (size > q[1]).astype(int)
    hi = np.zeros(nc, dtype=int)
    np.maximum.at(hi, w["cand_of"], (w["rnd"] > 0).astype(int))
    stratum = terc * 2 + hi
    arm = np.zeros(nc, dtype=bool)
    for s in np.unique(stratum):
        idx = np.flatnonzero(stratum == s)
        idx = idx[rng.permutation(len(idx))]
        arm[idx[::2]] = True
    return arm


def trial(cfg, days, seed):
    half = {**cfg, "shifts": [(600, 1080)]}      # each arm gets the same intern-minutes
    rows = []
    for d in range(days):
        rng = np.random.default_rng([seed, d])
        w = make_world(rng, cfg)
        arm = assign_arms(w, rng)
        a = run(subset(w, ~arm), "fcfs", half)
        b = run(subset(w, arm), "composite", half)
        rows.append((a, b))
    return rows


if __name__ == "__main__":
    reps = int(sys.argv[1]) if len(sys.argv) > 1 else 400
    sreps = int(sys.argv[2]) if len(sys.argv) > 2 else 150
    tdays = int(sys.argv[3]) if len(sys.argv) > 3 else 300
    t0 = time.time()
    rng = np.random.default_rng(11)
    out = {"main": {}, "scen": {}, "ceiling": {}, "trial": {}}

    base, hyg = study(BASE, POLICIES, reps, seed=2026)
    out["hygiene"] = {k: float(np.mean([h[k] for h in hyg])) for k in hyg[0]}
    fc = base["fcfs"]["pct"]
    for p in POLICIES:
        d = base[p]
        row = {k: float(np.nanmean(v)) for k, v in d.items()}
        row["pct_lo"], row["pct_hi"] = (float(x) for x in np.percentile(d["pct"], [2.5, 97.5]))
        if p != "fcfs":
            row["vs_fcfs"] = paired(d["pct"], fc, rng)
        out["main"][p] = row
    print(f"main {time.time()-t0:.0f}s", flush=True)

    many = [(600, 1080), (720, 1200)] * 20
    ceil, _ = study({**BASE, "shifts": many}, ["composite"], 80, seed=77)
    out["ceiling"] = {k: float(np.nanmean(v)) for k, v in ceil["composite"].items()}
    print(f"ceiling {time.time()-t0:.0f}s", flush=True)

    SCEN = {
        "base (batch 19:51, awake 07-23)": {},
        "OLD assumption: candidates act 24h": dict(awake=(0, 1440)),
        "first touch held to 08:00": dict(first_touch_office=(480, 1080)),
        "batch arrives 10:00": dict(arrive_clock=600),
        "batch arrives 14:00": dict(arrive_clock=840),
        "requests spread 10:00-20:00": dict(arrive="spread"),
        "1 intern": dict(shifts=[(600, 1080)]),
        "4 interns": dict(shifts=[(600, 1080), (720, 1200)] * 2),
        "10% invalid contacts": dict(invalid_p=0.10),
        "40% low-reach pool": dict(prior=[0.30, 0.30, 0.40]),
        "multi-company heavy (60/30/10)": dict(dup_p=[0.60, 0.30, 0.10]),
        "opt-out 4% per failed touch": dict(opt_p=0.04),
        "batching costs more (+3 min)": dict(extra_cost=3.0),
    }
    ps = ["fcfs", "round", "batched", "composite", "oracle"]
    for name, ov in SCEN.items():
        r, _ = study({**BASE, **ov}, ps, sreps, seed=99)
        e = {"fcfs": float(r["fcfs"]["pct"].mean()), "fcfs_optout": float(r["fcfs"]["optout"].mean())}
        for p in ps[1:]:
            e[p] = {"abs": float(r[p]["pct"].mean()), **paired(r[p]["pct"], r["fcfs"]["pct"], rng, 1500),
                    "optout": float(r[p]["optout"].mean())}
        out["scen"][name] = e
        print(f"{name} {time.time()-t0:.0f}s", flush=True)

    rows = trial(BASE, tdays, seed=5150)
    def col(i, k):
        return np.array([r[i][k] for r in rows])
    diff = col(1, "pct") - col(0, "pct")
    out["trial"] = dict(
        days=tdays, mean_diff=float(diff.mean()), sd_diff=float(diff.std(ddof=1)),
        days_80power=float(math.ceil((2.8016 * diff.std(ddof=1) / diff.mean()) ** 2)) if diff.mean() > 0 else None,
        fcfs=float(col(0, "pct").mean()), comp=float(col(1, "pct").mean()),
        mins_per_book_fcfs=float(col(0, "mins_per_book").mean()), mins_per_book_comp=float(col(1, "mins_per_book").mean()),
        att_fcfs=float(col(0, "att_per_cand").mean()), att_comp=float(col(1, "att_per_cand").mean()),
        optout_fcfs=float(col(0, "optout").mean()), optout_comp=float(col(1, "optout").mean()),
        medt_fcfs=float(col(0, "med_t").mean()), medt_comp=float(col(1, "med_t").mean()),
    )
    out["meta"] = dict(reps=reps, sreps=sreps, base=BASE, label=LABEL)
    with open(__file__.replace("simulation.py", "simulation_results.json"), "w") as f:
        json.dump(out, f, indent=1)
    print(f"saved {time.time()-t0:.0f}s", flush=True)
