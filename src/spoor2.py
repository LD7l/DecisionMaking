"""
VNS voor het verbeteren van cyclische 4-wekenroosters, voor alle afdelingen
samen, met uitlenen tussen afdelingen en daarna aanwerven/ontslaan.

Voorstelling (één grote matrix voor alle verpleegkundigen):
    x[i][d] = dienst van verpleegkundige i op dag d
              0 = E (06-15u), 1 = L (15-24u), 2 = N (21-06u), 3 = F (vrij)
    a[i][d] = afdeling waar die dienst gewerkt wordt (thuisafdeling, tenzij uitgeleend)
Dag 0 is een maandag; week w bestaat uit de dagen 7w .. 7w+6.

Volgorde van voorkeur bij tekorten en overschotten:
    1. uitlenen naar een andere afdeling (in de VNS, kleine strafkost w_float)
    2. pas wat daarna nog overblijft: aanwerven of ontslaan (hire_fire)

Cycliciteit: een cel wijkt af als de DIENST verschilt van het cyclische rooster.
Uitlenen verandert het uur niet en telt dus niet als afwijking (aparte KPI).

Niveaus (welke soorten stappen toegelaten zijn), zie LEVELS onderaan.
"""

import random
import time
from dataclasses import dataclass, replace

E, L, N, F = range(4)
WORK = (E, L, N)
NAMES = "ELNF"
NS = len(NAMES)

# Soorten stappen
FLOAT = "float"            # werkdienst naar een andere afdeling, zelfde uur
SWAP_WORK = "swap_work"    # twee personen (zelfde afdeling) ruilen werkdiensten op één dag
CHANGE_WORK = "change_work"  # werkdienst naar ander uur (E/L/N), vrije dagen blijven
MOVE_OFF = "move_off"      # binnen één week een werkdag en een vrije dag van plaats wisselen
SWAP = "swap"              # ruil op één dag, ook met vrije dagen
CHANGE = "change"          # één cel vrij wijzigen (ook werk <-> vrij)
WEEKEND = "weekend"        # twee personen ruilen hun weekend
DP = "dp"                  # één persoon volledig herplannen (horizontale fase)

LEVELS = {
    0: [],
    1: [FLOAT],
    2: [FLOAT, SWAP_WORK],
    3: [FLOAT, SWAP_WORK, CHANGE_WORK],
    4: [FLOAT, SWAP_WORK, CHANGE_WORK, MOVE_OFF],
    5: [FLOAT, SWAP_WORK, CHANGE_WORK, MOVE_OFF, SWAP, CHANGE, WEEKEND, DP],
}


# ---------------------------------------------------------------------------
# Invoer
# ---------------------------------------------------------------------------
@dataclass
class Instance:
    pref: list          # pref[i][d][s]     voorkeurskost (matrix 9.20), s = E, L, N, F
    typ: list           # typ[i]            0 = type 1, 1 = type 2
    dept: list          # dept[i]           thuisafdeling 0..K-1
    demand: list        # demand[k][d][s][t] vereist aantal per afdeling, s in WORK
    cyclic: list        # cyclic[i][d]      cyclisch rooster; None = geen cyclische lijn
    target: list        # target[i]         gewenst aantal diensten in 4 weken
    wage: list          # wage[i][d][s]     loonkost (tabel 9.16), 0 voor F
    can_float: list = None  # can_float[i]  verzameling afdelingen waar i mag werken
    fixed: list = None      # fixed[i][d]   True = cel ligt vast (bovenop verlof)
    prev_last: list = None  # dienst op de dag vóór de planningsperiode
    prev_run: list = None   # aaneengesloten werkdagen op het einde van de vorige periode

    # harde regels
    max_work_week: int = 6
    max_consecutive: int = 6                 # max_cons_wrk
    min_consecutive: int = 2                 # min_cons_wrk: geen losse werkdagen
    # verboden opeenvolgingen (gisteren, vandaag): N->E 0u rust, N->L 9u, L->E 6u
    forbidden_pairs: frozenset = frozenset({(N, E), (N, L), (L, E)})
    type1_covers_type2: bool = True
    leave_threshold: float = 100             # voorkeur >= dit = verboden dienst (verlof)

    # gewichten (tabel 9.18; w_over, w_float en w_dev zelf te kiezen)
    w_pref: float = 1.0
    w_short: float = 100.0
    w_over: float = 10.0
    w_fair: float = 5.0
    w_float: float = 3.0    # per uitgeleende dienst; klein, zodat uitlenen een tekort altijd verslaat
    w_wage: float = 0.0
    w_dev: float = 0.0

    def __post_init__(self):
        self.n = len(self.pref)
        self.days = len(self.pref[0])
        self.weeks = self.days // 7
        self.K = len(self.demand)
        if self.can_float is None:
            self.can_float = [set(range(self.K)) for _ in range(self.n)]
        for i in range(self.n):
            self.can_float[i] = set(self.can_float[i]) | {self.dept[i]}
        if self.fixed is None:
            self.fixed = [[False] * self.days for _ in range(self.n)]
        if self.prev_last is None:
            self.prev_last = [F] * self.n
        if self.prev_run is None:
            self.prev_run = [0] * self.n
        self.banned = [[{s for s in WORK if self.pref[i][d][s] >= self.leave_threshold}
                        for d in range(self.days)] for i in range(self.n)]
        for i in range(self.n):
            for d in range(self.days):
                if len(self.banned[i][d]) == len(WORK):
                    self.fixed[i][d] = True


# ---------------------------------------------------------------------------
# Harde regels controleren en het startrooster herstellen
# ---------------------------------------------------------------------------
def check_roster(inst, x):
    """Lijst van overtredingen: (verpleegkundige, [dagen], uitleg). Leeg = in orde."""
    errors = []
    for i in range(inst.n):
        run, start = inst.prev_run[i], 0
        for d in range(inst.days):
            s = x[i][d]
            prev = x[i][d - 1] if d else inst.prev_last[i]
            if (prev, s) in inst.forbidden_pairs:
                errors.append((i, [d], f"{NAMES[prev]} gevolgd door {NAMES[s]}"))
            if s in inst.banned[i][d]:
                errors.append((i, [d], "dienst op verlof"))
            if s != F:
                if run == 0:
                    start = d
                run += 1
                if run > inst.max_consecutive:
                    errors.append((i, [d], f"{run} werkdagen na elkaar"))
            else:
                if 0 < run < inst.min_consecutive:
                    errors.append((i, list(range(start, d)), f"reeks van {run} werkdag(en)"))
                run = 0
        for w in range(inst.weeks):
            wdays = [d for d in range(7 * w, 7 * w + 7) if x[i][d] != F]
            if len(wdays) > inst.max_work_week:
                errors.append((i, [wdays[-1]], f"{len(wdays)} werkdagen in week {w + 1}"))
    return errors


def prepare_start(inst, x, verbose=False):
    """Elke cel die een harde regel breekt wordt F, tot alles in orde is."""
    x = [row[:] for row in x]
    while True:
        errors = check_roster(inst, x)
        if not errors:
            return x
        if verbose:
            print(f"herstel: {len(errors)} overtredingen, bv. {errors[0]}")
        for i, days, _ in errors:
            for d in days:
                x[i][d] = F


# ---------------------------------------------------------------------------
# Rooster met tellers (delta-evaluatie)
# ---------------------------------------------------------------------------
class Roster:
    def __init__(self, inst, x, a=None, budget=None, kinds=None, floating=True):
        self.I = inst
        self.x = [row[:] for row in x]
        self.a = ([row[:] for row in a] if a is not None
                  else [[inst.dept[i]] * inst.days for i in range(inst.n)])
        ref = sum(c is not None for row in inst.cyclic for c in row)
        self.budget = ref if budget is None else budget
        self.kinds = list(LEVELS[5] if kinds is None else kinds)
        self.floating = floating
        if not floating and FLOAT in self.kinds:
            self.kinds.remove(FLOAT)
        self._rebuild()

    def depts_for(self, i):
        return self.I.can_float[i] if self.floating else {self.I.dept[i]}

    # --- alles één keer volledig uitrekenen --------------------------------
    def _rebuild(self):
        I, x = self.I, self.x
        self.cnt = [[[[0, 0] for _ in WORK] for _ in range(I.days)] for _ in range(I.K)]
        self.work = [0] * I.n
        self.wk = [[0] * I.weeks for _ in range(I.n)]
        self.pref_tot = self.wage_tot = 0.0
        self.dev = self.floats = 0
        for i in range(I.n):
            for d in range(I.days):
                s, k = x[i][d], self.a[i][d]
                if s == F:
                    self.a[i][d] = k = I.dept[i]
                self.pref_tot += I.pref[i][d][s]
                self.wage_tot += I.wage[i][d][s]
                self.dev += self._is_dev(i, d, s)
                if s != F:
                    self.cnt[k][d][s][I.typ[i]] += 1
                    self.work[i] += 1
                    self.wk[i][d // 7] += 1
                    self.floats += k != I.dept[i]
        self.staff_tot = sum(self._cell(k, d, s) for k in range(I.K)
                             for d in range(I.days) for s in WORK)
        self.fair_tot = sum(abs(self.work[i] - I.target[i]) for i in range(I.n))

    def _is_dev(self, i, d, s):
        cyc = self.I.cyclic[i][d]
        return cyc is not None and s != cyc

    # --- tekort en overschot per afdeling, dag en dienst ---------------------
    def _short_over_types(self, k, d, s):
        """(tekort type 1, tekort type 2, overschot)"""
        c1, c2 = self.cnt[k][d][s]
        d1, d2 = self.I.demand[k][d][s]
        if self.I.type1_covers_type2:
            extra1 = max(0, c1 - d1)
            need2 = max(0, d2 - c2)
            use = min(extra1, need2)
            return max(0, d1 - c1), need2 - use, extra1 - use + max(0, c2 - d2)
        return max(0, d1 - c1), max(0, d2 - c2), max(0, c1 - d1) + max(0, c2 - d2)

    def _cell(self, k, d, s):
        s1, s2, over = self._short_over_types(k, d, s)
        return self.I.w_short * (s1 + s2) + self.I.w_over * over

    def objective(self):
        I = self.I
        return (I.w_pref * self.pref_tot + self.staff_tot + I.w_fair * self.fair_tot
                + I.w_float * self.floats + I.w_wage * self.wage_tot + I.w_dev * self.dev)

    # --- één cel wijzigen (dienst s, afdeling k) en tellers bijwerken --------
    def _set(self, i, d, s, k):
        I = self.I
        if s == F:
            k = I.dept[i]
        old, ok_ = self.x[i][d], self.a[i][d]
        if old == s and ok_ == k:
            return
        t, home = I.typ[i], I.dept[i]
        self.pref_tot += I.pref[i][d][s] - I.pref[i][d][old]
        self.wage_tot += I.wage[i][d][s] - I.wage[i][d][old]
        self.dev += self._is_dev(i, d, s) - self._is_dev(i, d, old)
        self.fair_tot -= abs(self.work[i] - I.target[i])
        if old != F:
            self.staff_tot -= self._cell(ok_, d, old)
            self.cnt[ok_][d][old][t] -= 1
            self.staff_tot += self._cell(ok_, d, old)
            self.work[i] -= 1
            self.wk[i][d // 7] -= 1
            self.floats -= ok_ != home
        if s != F:
            self.staff_tot -= self._cell(k, d, s)
            self.cnt[k][d][s][t] += 1
            self.staff_tot += self._cell(k, d, s)
            self.work[i] += 1
            self.wk[i][d // 7] += 1
            self.floats += k != home
        self.fair_tot += abs(self.work[i] - I.target[i])
        self.x[i][d], self.a[i][d] = s, k

    # --- harde regels rond één cel controleren ------------------------------
    def _ok(self, i, d):
        I, row = self.I, self.x[i]
        s = row[d]
        if s in I.banned[i][d]:
            return False
        if s != F and self.a[i][d] not in self.depts_for(i):
            return False
        prev = row[d - 1] if d > 0 else I.prev_last[i]
        if (prev, s) in I.forbidden_pairs:
            return False
        if d + 1 < I.days and (s, row[d + 1]) in I.forbidden_pairs:
            return False
        if s != F and self.wk[i][d // 7] > I.max_work_week:
            return False
        return all(self._run_ok(i, q) for q in (d - 1, d, d + 1) if 0 <= q < I.days)

    def _run_ok(self, i, d):
        """Reeks werkdagen rond dag d: niet langer dan het maximum en niet korter
        dan het minimum (behalve als ze doorloopt tot het einde van de periode)."""
        I, row = self.I, self.x[i]
        if row[d] == F:
            return True
        lo = d
        while lo > 0 and row[lo - 1] != F:
            lo -= 1
        hi = d
        while hi + 1 < I.days and row[hi + 1] != F:
            hi += 1
        length = hi - lo + 1 + (I.prev_run[i] if lo == 0 else 0)
        if length > I.max_consecutive:
            return False
        return hi == I.days - 1 or length >= I.min_consecutive

    # --- een stap proberen ---------------------------------------------------
    def try_move(self, changes, accept=lambda delta: delta < -1e-9):
        """changes = lijst van (i, d, dienst, afdeling). True = aanvaard."""
        if any(self.I.fixed[i][d] for i, d, _, _ in changes):
            return False
        before = self.objective()
        undo = [(i, d, self.x[i][d], self.a[i][d]) for i, d, _, _ in changes]
        for i, d, s, k in changes:
            self._set(i, d, s, k)
        feasible = (self.dev <= self.budget
                    and all(self._ok(i, d) for i, d, _, _ in changes))
        if feasible and accept(self.objective() - before):
            return True
        for i, d, s, k in reversed(undo):
            self._set(i, d, s, k)
        return False

    def changes_something(self, changes):
        return any(self.x[i][d] != s or (s != F and self.a[i][d] != k)
                   for i, d, s, k in changes)

    # ---------------------------------------------------------------------
    # Stappen per soort: alle kandidaten (voor VND) en één willekeurige (schudden)
    # ---------------------------------------------------------------------
    def _groups(self):
        g = {}
        for i in range(self.I.n):
            g.setdefault(self.I.dept[i], []).append(i)
        return list(g.values())

    def candidates(self, kind):
        I, x, a = self.I, self.x, self.a
        out = []
        if kind == FLOAT:
            for i in range(I.n):
                for d in range(I.days):
                    if x[i][d] != F:
                        out += [[(i, d, x[i][d], k)] for k in self.depts_for(i)
                                if k != a[i][d]]
        elif kind in (SWAP_WORK, SWAP):
            for grp in self._groups():
                for d in range(I.days):
                    for p in range(len(grp)):
                        for q in range(p + 1, len(grp)):
                            i, j = grp[p], grp[q]
                            if kind == SWAP_WORK and (x[i][d] == F or x[j][d] == F):
                                continue
                            if (x[i][d], a[i][d]) != (x[j][d], a[j][d]):
                                out.append([(i, d, x[j][d], a[j][d]),
                                            (j, d, x[i][d], a[i][d])])
        elif kind == CHANGE_WORK:
            for i in range(I.n):
                for d in range(I.days):
                    if x[i][d] != F:
                        out += [[(i, d, s, a[i][d])] for s in WORK if s != x[i][d]]
        elif kind == MOVE_OFF:
            for i in range(I.n):
                for w in range(I.weeks):
                    wd = range(7 * w, 7 * w + 7)
                    for d1 in wd:
                        if x[i][d1] == F:
                            continue
                        for d2 in wd:
                            if x[i][d2] == F:
                                out.append([(i, d1, F, I.dept[i]),
                                            (i, d2, x[i][d1], a[i][d1])])
        elif kind == CHANGE:
            for i in range(I.n):
                for d in range(I.days):
                    for s in range(NS):
                        if s == F:
                            if x[i][d] != F:
                                out.append([(i, d, F, I.dept[i])])
                        elif x[i][d] != F:
                            if s != x[i][d]:
                                out.append([(i, d, s, a[i][d])])
                        else:   # van vrij naar werk: in de eigen of een andere afdeling
                            out += [[(i, d, s, k)] for k in self.depts_for(i)]
        elif kind == WEEKEND:
            for grp in self._groups():
                for w in range(I.weeks):
                    sa, su = 7 * w + 5, 7 * w + 6
                    for p in range(len(grp)):
                        for q in range(p + 1, len(grp)):
                            i, j = grp[p], grp[q]
                            if ((x[i][sa], x[i][su]) != (x[j][sa], x[j][su])
                                    or (a[i][sa], a[i][su]) != (a[j][sa], a[j][su])):
                                out.append([(i, sa, x[j][sa], a[j][sa]),
                                            (i, su, x[j][su], a[j][su]),
                                            (j, sa, x[i][sa], a[i][sa]),
                                            (j, su, x[i][su], a[i][su])])
        return out

    def random_move(self, kind):
        """Eén willekeurige stap van deze soort (kan ongeldig zijn; try_move beslist)."""
        I, x, a = self.I, self.x, self.a
        i, d = random.randrange(I.n), random.randrange(I.days)
        if kind == FLOAT:
            if x[i][d] == F:
                return None
            return [(i, d, x[i][d], random.choice(list(self.depts_for(i))))]
        if kind in (SWAP_WORK, SWAP, WEEKEND):
            grp = [j for j in range(I.n) if I.dept[j] == I.dept[i] and j != i]
            if not grp:
                return None
            j = random.choice(grp)
            days = [d] if kind != WEEKEND else [7 * random.randrange(I.weeks) + 5,
                                                7 * random.randrange(I.weeks) + 6]
            if kind == WEEKEND:
                days = [days[0], days[0] + 1]
            if kind == SWAP_WORK and (x[i][d] == F or x[j][d] == F):
                return None
            ch = []
            for q in days:
                ch += [(i, q, x[j][q], a[j][q]), (j, q, x[i][q], a[i][q])]
            return ch
        if kind == CHANGE_WORK:
            if x[i][d] == F:
                return None
            return [(i, d, random.choice(WORK), a[i][d])]
        if kind == MOVE_OFF:
            w = d // 7
            d2 = 7 * w + random.randrange(7)
            if x[i][d] == F or x[i][d2] != F:
                return None
            return [(i, d, F, I.dept[i]), (i, d2, x[i][d], a[i][d])]
        if kind == CHANGE:
            s = random.randrange(NS)
            k = a[i][d] if x[i][d] != F else random.choice(list(self.depts_for(i)))
            return [(i, d, s, k)]
        return None

    def improve(self, kind):
        """Eén volledige pas over een soort stap (first improvement)."""
        if kind == DP:
            return self.n_dp()
        moves = self.candidates(kind)
        random.shuffle(moves)
        improved = False
        for ch in moves:
            if self.changes_something(ch) and self.try_move(ch):
                improved = True
        return improved

    # ---------------------------------------------------------------------
    # Horizontale fase: één persoon herplannen met dynamisch programmeren
    # ---------------------------------------------------------------------
    def n_dp(self, nurses=None):
        I = self.I
        improved = False
        nurses = list(range(I.n)) if nurses is None else list(nurses)
        random.shuffle(nurses)
        for i in nurses:
            own = sum(self._is_dev(i, d, self.x[i][d]) for d in range(I.days))
            allowed = self.budget - (self.dev - own)
            for lam in (0.0, 20.0, 100.0, 1000.0):
                plan = self._best_row(i, lam)
                if plan is None:
                    break
                if sum(self._is_dev(i, d, s) for d, (s, _) in enumerate(plan)) <= allowed:
                    ch = [(i, d, s, k) for d, (s, k) in enumerate(plan)]
                    if self.changes_something(ch) and self._try_row(i, ch):
                        improved = True
                    break
        return improved

    def _try_row(self, i, changes):
        before = self.objective()
        undo = [(i, d, self.x[i][d], self.a[i][d]) for _, d, _, _ in changes]
        for _, d, s, k in changes:
            self._set(i, d, s, k)
        ok = self.dev <= self.budget and all(self._ok(i, d) for d in range(self.I.days))
        if ok and self.objective() - before < -1e-9:
            return True
        for _, d, s, k in reversed(undo):
            self._set(i, d, s, k)
        return False

    def _best_row(self, i, lam=0.0):
        I, t, home = self.I, self.I.typ[i], self.I.dept[i]
        # 1. bijdrage van i tijdelijk uit de tellers halen
        for d in range(I.days):
            if self.x[i][d] != F:
                self.cnt[self.a[i][d]][d][self.x[i][d]][t] -= 1
        # 2. per dag en dienst: kost én beste afdeling (uitlenen kost w_float)
        c = [[0.0] * NS for _ in range(I.days)]
        best_k = [[home] * NS for _ in range(I.days)]
        for d in range(I.days):
            for s in range(NS):
                v = (I.w_pref * I.pref[i][d][s] + I.w_wage * I.wage[i][d][s]
                     + (I.w_dev + lam) * self._is_dev(i, d, s))
                if s != F:
                    best = None
                    ks = [self.a[i][d]] if I.fixed[i][d] else self.depts_for(i)
                    for k in ks:
                        base = self._cell(k, d, s)
                        self.cnt[k][d][s][t] += 1
                        m = self._cell(k, d, s) - base + I.w_float * (k != home)
                        self.cnt[k][d][s][t] -= 1
                        if best is None or m < best:
                            best, best_k[d][s] = m, k
                    v += best
                c[d][s] = v
        # 3. tellers herstellen
        for d in range(I.days):
            if self.x[i][d] != F:
                self.cnt[self.a[i][d]][d][self.x[i][d]][t] += 1
        # 4. DP. Toestand = (dienst gisteren, werkdagen deze week, totaal, reeks)
        layer = {(I.prev_last[i], 0, 0, I.prev_run[i]): (0.0, None)}
        history = []
        for d in range(I.days):
            new = {}
            allowed = ([self.x[i][d]] if I.fixed[i][d]
                       else [s for s in range(NS) if s not in I.banned[i][d]])
            for key_old, (cost, _) in layer.items():
                prev, wk, tot, run = key_old
                if d % 7 == 0:
                    wk = 0
                for s in allowed:
                    if (prev, s) in I.forbidden_pairs:
                        continue
                    if s == F:
                        if 0 < run < I.min_consecutive:
                            continue
                        key = (s, wk, tot, 0)
                    else:
                        if wk + 1 > I.max_work_week or run + 1 > I.max_consecutive:
                            continue
                        key = (s, wk + 1, tot + 1, run + 1)
                    val = cost + c[d][s]
                    if key not in new or val < new[key][0]:
                        new[key] = (val, key_old)
            if not new:
                return None
            history.append(new)
            layer = new
        key = min(layer, key=lambda q: layer[q][0] + I.w_fair * abs(q[2] - I.target[i]))
        row = [F] * I.days
        for d in range(I.days - 1, -1, -1):
            row[d] = key[0]
            key = history[d][key][1]
        return [(s, best_k[d][s] if s != F else home) for d, s in enumerate(row)]

    # ---------------------------------------------------------------------
    # Schudden: k willekeurige toegelaten stappen die echt iets veranderen
    # ---------------------------------------------------------------------
    def shake(self, k):
        kinds = [q for q in self.kinds if q != DP]
        touched = set()
        if not kinds:
            return touched
        done = tries = 0
        while done < k and tries < 200 * k:
            tries += 1
            ch = self.random_move(random.choice(kinds))
            if not ch or not self.changes_something(ch):
                continue
            if self.try_move(ch, accept=lambda delta: True):
                done += 1
                touched.update(c[0] for c in ch)
        return touched

    # ---------------------------------------------------------------------
    # KPI's
    # ---------------------------------------------------------------------
    def kpis(self):
        I = self.I
        short = over = 0
        per_dept = [0] * I.K
        for k in range(I.K):
            for d in range(I.days):
                for s in WORK:
                    s1, s2, o = self._short_over_types(k, d, s)
                    short += s1 + s2
                    per_dept[k] += s1 + s2
                    over += o
        ref = sum(c is not None for row in I.cyclic for c in row)
        return {
            "objective": round(self.objective(), 1),
            "nurse_pref": round(self.pref_tot, 1),
            "fairness_dev": self.fair_tot,
            "shortage": short,
            "shortage_per_dept": per_dept,
            "overstaffing": over,
            "floated_shifts": self.floats,
            "wage_cost": round(self.wage_tot, 1),
            "deviation_cells": self.dev,
            "deviation_pct": round(100 * self.dev / ref, 1) if ref else 0.0,
            "nurses": I.n,
        }


# ---------------------------------------------------------------------------
# VND en VNS
# ---------------------------------------------------------------------------
def vnd(r, dp_nurses=None):
    """Stappen in de volgorde van r.kinds; DP als laatste, eventueel beperkt
    tot bepaalde personen (bv. wie geschud werd)."""
    kinds = [q for q in r.kinds if q != DP]
    hoods = [lambda q=q: r.improve(q) for q in kinds]
    if DP in r.kinds:
        hoods.append(lambda: r.n_dp(dp_nurses))
    j = 0
    while j < len(hoods):
        j = 0 if hoods[j]() else j + 1


def vns(inst, start_x, start_a=None, budget=None, kinds=None, floating=True,
        time_limit=60, k_max=10, seed=0, verbose=False):
    random.seed(seed)
    r = Roster(inst, prepare_start(inst, start_x), start_a, budget, kinds, floating)
    vnd(r)
    cur = ([row[:] for row in r.x], [row[:] for row in r.a])
    cur_obj = best_obj = r.objective()
    best = cur
    t0, k = time.time(), 1
    while time.time() - t0 < time_limit:
        touched = r.shake(k)
        if not touched:
            break                                   # niets meer te schudden (bv. niveau 0)
        vnd(r, dp_nurses=touched)
        if r.objective() < cur_obj - 1e-9:
            cur, cur_obj, k = ([row[:] for row in r.x], [row[:] for row in r.a]), r.objective(), 1
            if cur_obj < best_obj:
                best, best_obj = cur, cur_obj
                if verbose:
                    print(f"{time.time()-t0:6.1f}s  beste = {best_obj:.1f}")
        else:
            r = Roster(inst, cur[0], cur[1], budget, kinds, floating)
            k = k + 1 if k < k_max else 1
    return Roster(inst, best[0], best[1], budget, kinds, floating)


# ---------------------------------------------------------------------------
# Fase 3: aanwerven en ontslaan, pas na uitlenen
# ---------------------------------------------------------------------------
def _add_nurse(inst, k, t, target, wage_row):
    """Nieuwe verpleegkundige: voorkeur 5 voor elke dienst (case), geen cyclische lijn."""
    def add(lst, v):
        return list(lst) + [v]
    return replace(
        inst,
        pref=add(inst.pref, [[5, 5, 5, 0] for _ in range(inst.days)]),
        typ=add(inst.typ, t), dept=add(inst.dept, k),
        cyclic=add(inst.cyclic, [None] * inst.days),
        target=add(inst.target, target), wage=add(inst.wage, wage_row),
        can_float=add(inst.can_float, set(range(inst.K))),
        fixed=add(inst.fixed, [False] * inst.days),
        prev_last=add(inst.prev_last, F), prev_run=add(inst.prev_run, 0))


def _remove_nurse(inst, i):
    def rm(lst):
        return [v for j, v in enumerate(lst) if j != i]
    return replace(inst, pref=rm(inst.pref), typ=rm(inst.typ), dept=rm(inst.dept),
                   cyclic=rm(inst.cyclic), target=rm(inst.target), wage=rm(inst.wage),
                   can_float=rm(inst.can_float), fixed=rm(inst.fixed),
                   prev_last=rm(inst.prev_last), prev_run=rm(inst.prev_run))


def hire_fire(inst, r, kinds=None, time_limit=20, hire_threshold=10,
              fire_threshold=20, new_target=20, max_changes=10, verbose=False):
    """Na de VNS (mét uitlenen):
    - aanwerven: per afdeling en type, zolang het resterende tekort over de
      4 weken >= hire_threshold diensten is, één persoon bijnemen en opnieuw VNS;
    - ontslaan: per afdeling en type, als het overschot >= fire_threshold is,
      de persoon met de hoogste gemiddelde aversiescore (case) proberen weg te
      laten; behouden als het tekort daardoor niet stijgt.
    Uitlenen gebeurt dus altijd eerst. Geeft (instance, roster, logboek)."""
    log = []
    for _ in range(max_changes):
        need = {}
        for k in range(inst.K):
            for d in range(inst.days):
                for s in WORK:
                    s1, s2, _ = r._short_over_types(k, d, s)
                    need[(k, 0)] = need.get((k, 0), 0) + s1
                    need[(k, 1)] = need.get((k, 1), 0) + s2
        (k, t), amount = max(need.items(), key=lambda q: q[1])
        if amount < hire_threshold:
            break
        ref = next(j for j in range(inst.n) if inst.typ[j] == t)
        inst = _add_nurse(inst, k, t, min(new_target, amount), [row[:] for row in inst.wage[ref]])
        r = vns(inst, r.x + [[F] * inst.days], r.a + [[k] * inst.days],
                kinds=kinds, time_limit=time_limit)
        log.append(f"aanwerving afdeling {k}, type {t + 1} (resterend tekort {amount})")
        if verbose:
            print(log[-1], r.kpis()["shortage"])
    for _ in range(max_changes):
        surplus = {}
        for k in range(inst.K):
            for d in range(inst.days):
                for s in WORK:
                    surplus[k] = surplus.get(k, 0) + r._short_over_types(k, d, s)[2]
        k, amount = max(surplus.items(), key=lambda q: q[1])
        if amount < fire_threshold:
            break
        cands = [i for i in range(inst.n) if inst.dept[i] == k and inst.cyclic[i][0] is not None]
        aversion = lambda i: sum(min(inst.pref[i][d][s] for s in WORK) for d in range(inst.days))
        i = max(cands, key=aversion)
        new_inst = _remove_nurse(inst, i)
        new_r = vns(new_inst, [row for j, row in enumerate(r.x) if j != i],
                    [row for j, row in enumerate(r.a) if j != i],
                    kinds=kinds, time_limit=time_limit)
        if new_r.kpis()["shortage"] > r.kpis()["shortage"]:
            break
        log.append(f"ontslag verpleegkundige {i} uit afdeling {k} (overschot {amount})")
        inst, r = new_inst, new_r
        if verbose:
            print(log[-1])
    return inst, r, log


# ---------------------------------------------------------------------------
# Gevoeligheidsanalyse over de niveaus, met en zonder uitlenen
# ---------------------------------------------------------------------------
def run_levels(inst, levels=(0, 1, 2, 3, 4, 5), time_limit=20, seed=0):
    results = []
    for floating in (False, True):
        x, a = inst.cyclic, None
        for lv in levels:
            r = vns(inst, x, a, kinds=LEVELS[lv], floating=floating,
                    time_limit=time_limit, seed=seed)
            k = r.kpis()
            k.update(level=lv, floating=floating)
            results.append(k)
            x, a = r.x, r.a                      # warme start voor het volgende niveau
    return results


# ---------------------------------------------------------------------------
# Demo met verzonnen data (vervang door jullie eigen invoer)
# ---------------------------------------------------------------------------
def demo_instance(n_per_dept=(32, 28, 30, 26), seed=1):
    random.seed(seed)
    days, K = 28, len(n_per_dept)
    base = ("EEELLFF" "LLNNFFF" "EEELLFF" "NNNFFFF")
    base = [NAMES.index(c) for c in base]
    dept = [k for k, m in enumerate(n_per_dept) for _ in range(m)]
    n = len(dept)
    cyclic = [base[7 * (i % 4):] + base[:7 * (i % 4)] for i in range(n)]
    typ = [0 if i % 3 == 0 else 1 for i in range(n)]
    pref = [[[random.randint(1, 9) for _ in WORK] + [0] for _ in range(days)]
            for _ in range(n)]
    for i in range(n):
        if random.random() < 0.3:
            pref[i][random.randrange(days)] = [100, 100, 100, 0]
        if random.random() < 0.3:
            pref[i][random.randrange(days)][N] = 100
    # vraag rond het cyclische aanbod, met een afdeling met tekorten (0)
    # en een afdeling met overschotten (3)
    bias = [1, 0, 0, -1]
    demand = [[[[0, 0] for _ in WORK] for _ in range(days)] for _ in range(K)]
    for i in range(n):
        for d in range(days):
            if cyclic[i][d] != F:
                demand[dept[i]][d][cyclic[i][d]][typ[i]] += 1
    for k in range(K):
        for d in range(days):
            for s in WORK:
                for t in (0, 1):
                    v = demand[k][d][s][t] + random.choice((-1, 0, 0, 1))
                    if random.random() < 0.4:
                        v += bias[k]
                    demand[k][d][s][t] = max(0, v)
    target = [20 if i % 5 else 15 for i in range(n)]
    wage = [[[100 + 20 * (s == N) + 30 * (d % 7 >= 5) if s != F else 0
              for s in range(NS)] for d in range(days)] for _ in range(n)]
    prev_last = [row[-1] for row in cyclic]
    prev_run = []
    for row in cyclic:
        k = 0
        while k < days and row[days - 1 - k] != F:
            k += 1
        prev_run.append(k)
    return Instance(pref=pref, typ=typ, dept=dept, demand=demand, cyclic=cyclic,
                    target=target, wage=wage, prev_last=prev_last, prev_run=prev_run)


if _name_ == "_main_":
    inst = demo_instance()
    print("Cyclisch rooster:", Roster(inst, prepare_start(inst, inst.cyclic)).kpis())
    r = vns(inst, inst.cyclic, time_limit=30, verbose=True)
    print("Na VNS (alles, met uitlenen):", r.kpis())
    print("Overtredingen:", check_roster(inst, r.x))
    inst2, r2, log = hire_fire(inst, r, time_limit=15, verbose=True)
    print("Na aanwerven/ontslaan:", r2.kpis())
    print("\nNiveaus, zonder en met uitlenen:")
    for row in run_levels(inst, time_limit=10):
        print(row["floating"], row["level"], {k: row[k] for k in
              ("objective", "shortage", "overstaffing", "floated_shifts",
               "nurse_pref", "deviation_pct")})