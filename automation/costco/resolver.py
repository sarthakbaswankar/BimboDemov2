"""Costco variance resolution engine.

Replaces the manual "mark red / sort / try +1, +2 / search ship-to and date /
add rows until the variance is 0" work. Every document it attaches is logged
with the rule that justified it, and an open item is never used twice.
"""
from datetime import timedelta
from itertools import combinations


class Resolver:
    def __init__(self, oi, used, P):
        self.oi, self.used = oi, used
        pk = P["COSTCO"]
        self.eps = P["COMMON"]["AMOUNT_EPSILON"]
        self.tol, self.incl = pk["WO_TOLERANCE"], pk["WO_INCLUSIVE"]
        self.offsets = pk["NEIGHBOUR_OFFSETS"]
        self.max_comb = pk["MAX_COMBINATION"]
        self.max_cand = pk["MAX_CANDIDATES"]
        self.day_before = pk["SEARCH_DAY_BEFORE"]

    def within_tol(self, amount):
        return abs(amount) <= self.tol + 1e-9 if self.incl else abs(amount) < self.tol - 1e-9

    def _pool(self, ship_tos, dt):
        out = []
        for s in ship_tos:
            out += [d for d in self.oi.by_shipto_date.get((s, dt), []) if d["id"] not in self.used]
        return out

    def _subsets(self, cands, target):
        """Smallest set of candidates whose open amounts sum to target; (solution, count)."""
        cands = cands[: self.max_cand]
        for k in range(1, min(self.max_comb, len(cands)) + 1):
            found = []
            for combo in combinations(cands, k):
                if abs(sum(d["due"] for d in combo) - target) < self.eps:
                    found.append(combo)
                    if len(found) > 1:
                        break
            if found:
                return list(found[0]), len(found)
        return None, 0

    def resolve(self, key, paid, docs, ship_tos, dt):
        """Attach open items until the variance is 0 (or nothing helps).

        key: R12 of the paid invoice; docs: documents already attached (primary);
        ship_tos: set of candidate ship-tos; dt: date to search.
        Returns (added_docs, log, final_variance, notes).
        """
        var = round(paid - sum(d["due"] for d in docs), 2)
        added, log, notes = [], [], []

        def take(d, rule):
            nonlocal var
            before = var
            var = round(var - d["due"], 2)
            self.used.add(d["id"])
            added.append(d)
            log.append({"Rule": rule, "Doc Trx": d["trx"], "Doc Date": d["date"], "Doc Ship-to": d["ship_to"],
                        "Doc Open Amount": d["due"], "Variance Before": before, "Variance After": var})

        if abs(var) < self.eps or dt is None:
            return added, log, var, notes

        # C-11 credit memo with the same ship-to, date and amount
        for d in self._pool(ship_tos, dt):
            if d["due"] < 0 and abs(var - d["due"]) < self.eps:
                take(d, "C-11 Credit memo: same ship-to, date and amount")
                return added, log, var, notes

        # C-12 neighbour invoice numbers (x+1, x+2, x-1); must reduce |variance|
        if key.isdigit():
            for _ in range(2):
                best = None
                for off in self.offsets:
                    nk = str(int(key) + off).zfill(12)
                    for d in self.oi.by_key.get(nk, []):
                        if d["id"] in self.used or d["date"] != dt:
                            continue
                        if ship_tos and d["ship_to"] not in ship_tos:
                            continue
                        new = round(var - d["due"], 2)
                        if abs(new) < abs(var) - self.eps:
                            score = (0 if abs(new) < self.eps else 1, abs(new))
                            if best is None or score < best[0]:
                                best = (score, d, off)
                if not best:
                    break
                take(best[1], f"C-12 Neighbour invoice x{best[2]:+d}: same ship-to and date")
                if abs(var) < self.eps:
                    return added, log, var, notes

        # C-13 combination of open items for the ship-to (same day, then day before)
        dates = [(dt, "same day")] + ([(dt - timedelta(days=1), "day before")] if self.day_before else [])
        for search_dt, label in dates:
            sol, n = self._subsets(self._pool(ship_tos, search_dt), var)
            if sol:
                for d in sol:
                    take(d, f"C-13 {len(sol)} open item(s) totalling the variance: same ship-to, {label}")
                if n > 1:
                    notes.append("More than one combination fits; first one used, please verify")
                return added, log, var, notes

        # C-14 partial: one same-day document that leaves only a write-off-sized residual
        if not self.within_tol(var):
            best = None
            for d in self._pool(ship_tos, dt):
                new = round(var - d["due"], 2)
                if self.within_tol(new) and abs(new) < abs(var) - self.eps:
                    if best is None or abs(new) < best[0]:
                        best = (abs(new), d)
            if best:
                take(best[1], "C-14 Partial: same ship-to/date document leaves a residual within tolerance")
                notes.append("Partial match (C-14): residual written off, please verify")
        return added, log, var, notes
