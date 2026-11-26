#!/usr/bin/env python3
"""
CASE E - SPOOR 1: het cyclisch rooster met namen (afdelingen A t/m D)
====================================================================

IDEE
  Case D zegt hoeveel lijnen er nodig zijn: voor afdeling A 52 lijnen = 52 FTE.
  Een lijn is een standaardweek: één shift (Early, Late of Night) op 5 vaste dagen.
  Dit script vult die lijnen in:
    1. de eigen verpleegkundigen krijgen elk één lijn (zo goed mogelijk passend bij
       hun cyclische voorkeur);
    2. het tekort in FTE (52 - wat het eigen personeel samen werkt) wordt aangeworven.
  Omdat deeltijdsen minder werken dan hun lijn, zijn er meer mensen dan lijnen.

INPUT  (uit de Excel van de groep)
  - 'Case D fix (2)'                 : de lijnen (3 x 9h, 5-2 non-identical), de werkpatronen,
                                       de starturen van de shiften en de minimale bezetting
  - 'Department X - Nurse Char.'     : per verpleegkundige: type, tewerkstelling, cyclische voorkeur
  - 'Department X - Preferences'     : per verpleegkundige: aversie per dag en shift (28 dagen)

PLAATS IN HET PROJECT
  project/
    data/Excel_Group11_Case_A_B_C_D.xlsx   <- input (wordt hier gelezen)
    src/spoor1.py                          <- dit script (naast main.cpp)
    files/                                 <- output: invoer voor roster.exe
    results/                               <- output: overzicht-Excels
  Het script zoekt zelf de projectmap (de map met 'data/' erin), dus het maakt niet uit
  vanuit welke map je het start.

OUTPUT
  - map files/ met de 4 invoerbestanden voor de C++-template:
      Shift_system_dpt_X.txt   shiften (starturen, lengte) en minimale bezetting per shift
      Cyclic_roster_dpt_X.txt  het cyclisch rooster: per persoon 28 dagen (standaardweek x 4)
      Personnel_dpt_X.txt      per persoon: nummer, 28 x 5 voorkeurscores, tewerkstelling, type
      Constraints_dpt_X.txt    roosterregels (min/max diensten, opeenvolgende dagen, ...)
  - map results/ met:
      spoor1_overzicht_dpt_X.xlsx              leesbaar overzicht van het resultaat
      Cyclic_roster_dpt_X_leesbaar.txt         hetzelfde cyclisch rooster, maar met kolomkoppen
                                               (dagen), rijkoppen (personen) en de bezetting per dag
  (Het bestand in files/ blijft zonder koppen: roster.exe leest het met fscanf en verwacht
   exact dat formaat.)

GEBRUIK
  pip install openpyxl
  Open src/spoor1.py in VS Code en druk op Run (of in de terminal: python src/spoor1.py).
  Dan worden alle afdelingen (DEPARTMENTS) doorgerekend.
  Eén afdeling of een andere Excel: python src/spoor1.py --dept A --excel pad/naar/bestand.xlsx
"""

import argparse
import os
import re

import openpyxl
from openpyxl.styles import Font, PatternFill

# ============================================================================
# INSTELLINGEN
# ============================================================================

EXCEL_FILE = "Excel_Group11_Case_A_B_C_D.xlsx"   # bestandsnaam van de Excel in data/
DEPARTMENTS = "ABCD"                             # afdelingen die doorgerekend worden
DATA_DIR = "data"                                # map met de Excel
FILES_DIR = "files"                              # map waar roster.exe zijn input leest
RESULTS_DIR = "results"                          # map voor de overzicht-Excels

SHEET_CASE_D = "Case D fix (2)"
SCENARIO = "3 x 9h"
POLICY = "5-2 non-identical"
WEEKS = 4                    # de standaardweek wordt 4 keer herhaald (28 dagen)

# Strafpunten voor de cyclische voorkeur. Contractbepalingen ("No ...", "Only ...") zijn hard.
PEN_SHIFT_PREF = 5           # "Preference for X shifts" en de lijn heeft een ander shifttype
PEN_DAY_OFF = 1              # per dag van de lijn die ze liever vrij heeft
PEN_DAY_ON = 1               # per gewenste werkdag die niet in de lijn zit
TYPE1_ON_TYPE2 = False       # mag een type 1 op een type II-lijn? False = nee: elk type alleen op zijn eigen lijnen
PEN_TYPE1_ON_TYPE2 = 3       # alleen gebruikt als TYPE1_ON_TYPE2 = True: strafpunten voor type 1 op een type II-lijn

HIRE_CONTRACTS = [0.5, 0.75, 1.0]   # contracten voor aanwervingen (zoveel mogelijk voltijds)
HIRE_PREF = 5                       # voorkeurscore van een aanwerving (zoals de template voorstelt)

REQ_WITH_BUFFER = False      # Shift_system-bestand: minimum Case C (False) of eis met buffer (True)

CONSTRAINTS = {              # Constraints-bestand; de template vermenigvuldigt min/max_ass met de tewerkstelling
    "min_ass": 20, "max_ass": 20,              # doel: 20 diensten per 28 dagen voor een voltijdse
    "min_cons_wrk": 1, "max_cons_wrk": 6,      # opeenvolgende werkdagen
    "min_cons_shift": 1, "max_cons_shift": 6,  # opeenvolgende dagen met dezelfde shift
    "min_shift": 0, "max_shift": 28,           # aantal keer dezelfde shift per maand
    "identical": "N",                          # non-identical weekends
}

DAYS_NL = ["Ma", "Di", "Wo", "Do", "Vr", "Za", "Zo"]
DAY_NAMES = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
WEEKEND = (5, 6)
LETTER = {"Early": "E", "Day": "D", "Late": "L", "Night": "N"}
ID_RE = r"30[12][A-D]\d{3}"
EPS = 1e-9


# ============================================================================
# STAP 0 - INLEZEN
# ============================================================================

def find_row(ws, col, starts_with, first=1):
    for r in range(first, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(starts_with):
            return r
    raise ValueError(f"'{starts_with}' niet gevonden in '{ws.title}'")


def read_patterns(ws):
    """De 7 werkpatronen (ma-vr, di-za, ...) uit 'Work patterns used by the model'."""
    pats = []
    for r in range(1, ws.max_row + 1):
        label = ws.cell(r, 2).value
        if isinstance(label, str) and label.startswith(f"{POLICY[:3]} pattern"):
            pats.append({"label": label, "days": [int(ws.cell(r, c).value) for c in range(3, 10)]})
    assert len(pats) == 7, "verwachtte 7 werkpatronen"
    return pats


def read_shifts(ws, dept):
    """Shiften met startuur en minimale bezetting (Case C) en eis na buffer (Case D)."""
    d = "ABCD".index(dept)
    r0 = find_row(ws, 2, "Staffing requirements per shift from Case C")
    shifts = []
    for r in range(r0 + 2, r0 + 12):
        name, start = ws.cell(r, 5).value, ws.cell(r, 4).value
        if ws.cell(r, 2).value != SCENARIO or name in (None, "-") or start in (None, "-"):
            continue
        hour = start.hour if hasattr(start, "hour") else int(str(start).split(":")[0])
        shifts.append({"name": name, "start": hour,
                       "min": int(ws.cell(r, 6 + d).value), "min_buffer": int(ws.cell(r, 10 + d).value)})
    return sorted(shifts, key=lambda s: s["start"])


def read_lines(ws, dept):
    """De lijnen van Case D per (shift, patroon, type). Elk zo'n groepje heet een component."""
    r1 = find_row(ws, 2, f"Model {SCENARIO}")
    r2 = find_row(ws, 2, f"Total number of nurses ({SCENARIO})", r1)
    comps, req = [], {}
    for r in range(r1, r2):
        if ws.cell(r, 2).value != POLICY or ws.cell(r, 3).value != f"Dept {dept}":
            continue
        shift = ws.cell(r, 5).value
        req[shift] = {"req": round(ws.cell(r, 22).value), "min_type1": round(ws.cell(r, 23).value)}
        for k in range(7):
            for typ, col in ((1, 6 + k), (2, 13 + k)):      # F..L = type I, M..S = type II
                lines = round(ws.cell(r, col).value or 0)
                if lines:
                    comps.append({"shift": shift, "pattern": k, "type": typ, "lines": lines, "members": []})
    return comps, req


def read_nurses(wb, dept):
    ws = wb[f"Department {dept} - Nurse Char."]
    nurses = []
    for r in range(1, ws.max_row + 1):
        nid = ws.cell(r, 1).value
        if isinstance(nid, str) and re.fullmatch(ID_RE, nid.strip()):
            text = str(ws.cell(r, 18).value or "").strip()
            nurses.append({"id": nid.strip(), "rate": float(ws.cell(r, 16).value), "type": int(ws.cell(r, 17).value),
                           "text": text, "pref": parse_preference(text), "new": False})
    return nurses


def read_monthly_preferences(wb, dept):
    ws = wb[f"Department {dept} - Preferences"]
    head = find_row(ws, 2, "E")
    assert [ws.cell(head, c).value for c in range(2, 7)] == ["E", "D", "L", "N", "F"]
    prefs = {}
    for r in range(head + 1, ws.max_row + 1):
        nid = ws.cell(r, 1).value
        if isinstance(nid, str) and re.fullmatch(ID_RE, nid.strip()):
            vals = [ws.cell(r, c).value for c in range(2, 2 + 140)]
            prefs[nid.strip()] = [int(v) if isinstance(v, (int, float)) else 0 for v in vals]
    return prefs


# ============================================================================
# STAP 1 - STRAFSCORE: hoe slecht past een verpleegkundige bij een component?
# ============================================================================

def parse_preference(text):
    t = text.lower()
    p = {"no_shift": set(), "only_shift": None, "no_weekend": "no weekend" in t,
         "pref_shift": None, "days_off": set(), "days_on": set()}
    for s in ("early", "day", "late", "night"):
        if f"no {s} shift" in t:
            p["no_shift"].add(s.capitalize())
    m = re.search(r"only (early|day|late|night) shifts", t)
    p["only_shift"] = m.group(1).capitalize() if m else None
    m = re.search(r"preference for (early|day|late|night) shifts", t)
    p["pref_shift"] = m.group(1).capitalize() if m else None
    for clause in t.split(","):
        days = {i for i, name in enumerate(DAY_NAMES) if name in clause}
        if "weekend" in clause and "no weekend" not in clause:
            days |= set(WEEKEND)
        if "day off" in clause:
            p["days_off"] |= days
        elif "day on" in clause:
            p["days_on"] |= days
    return p


def score(nurse, comp, patterns):
    """None = mag niet (type of contract). Anders: aantal strafpunten (0 = past perfect)."""
    if nurse["type"] == 2 and comp["type"] == 1:
        return None
    if nurse["type"] == 1 and comp["type"] == 2 and not TYPE1_ON_TYPE2:
        return None
    p, days = nurse["pref"], patterns[comp["pattern"]]["days"]
    if comp["shift"] in p["no_shift"] or (p["only_shift"] and comp["shift"] != p["only_shift"]):
        return None
    if p["no_weekend"] and any(days[d] for d in WEEKEND):
        return None
    s = PEN_SHIFT_PREF if p["pref_shift"] and comp["shift"] != p["pref_shift"] else 0
    s += PEN_DAY_OFF * sum(days[d] for d in p["days_off"])
    s += PEN_DAY_ON * sum(1 - days[d] for d in p["days_on"])
    if nurse["type"] == 1 and comp["type"] == 2:
        s += PEN_TYPE1_ON_TYPE2
    return s


# ============================================================================
# STAP 2 - DE EIGEN VERPLEEGKUNDIGEN KRIJGEN ELK ÉÉN LIJN
# ============================================================================

def assign_own(nurses, comps, patterns):
    """Herhaal tot iedereen een lijn heeft: kies het koppel (verpleegkundige, vrije lijn)
    met de laagste strafscore. Bij gelijke score gaat wie het minst perfect passende
    lijnen heeft eerst, zodat flexibele mensen de rest opvullen. Type I-lijnen eerst."""
    S = {(n["id"], i): score(n, c, patterns) for n in nurses for i, c in enumerate(comps)}
    free = {n["id"]: n for n in nurses}

    def place(nid, i, status):
        n = free.pop(nid)
        comps[i]["members"].append(n)
        n["comp"], n["score"], n["status"] = i, S[nid, i], status

    for phase in (1, 2):
        while free:
            open_i = [i for i, c in enumerate(comps) if c["type"] == phase and len(c["members"]) < c["lines"]]
            flex = {nid: sum(1 for i in open_i if S[nid, i] == 0) for nid in free}
            options = [(S[nid, i], flex[nid], nid, i) for i in open_i for nid in free if S[nid, i] is not None]
            if not options:
                break
            _, _, nid, i = min(options)
            place(nid, i, "eigen personeel")

    # wie op geen enkele vrije lijn past (bv. 'No weekend shifts' en alle ma-vr-lijnen bezet),
    # komt als extra persoon op haar best passende lijn
    for nid in sorted(free):
        opts = [(S[nid, i], i) for i in range(len(comps)) if S[nid, i] is not None]
        place(nid, min(opts)[1], "eigen personeel (extra op lijn)")


# ============================================================================
# STAP 3 - AANWERVEN VOOR HET TEKORT IN FTE
# ============================================================================

def contracts(fte):
    """24,5 FTE -> 24 voltijdse + 1 halftijdse."""
    out, rest = [], fte
    while rest > EPS:
        rate = 1.0 if rest >= 1 - EPS else min(r for r in HIRE_CONTRACTS if r >= rest - EPS)
        out.append(rate)
        rest -= rate
    return out


def fte(c):
    return sum(m["rate"] for m in c["members"])


def hire(nurses, comps, dept):
    """Tekort = lijnen - som van de tewerkstelling van het eigen personeel.
    Eerst type I (als er te weinig type I-FTE is), dan type II. Mag type 1 niet op type II-lijnen
    (TYPE1_ON_TYPE2 = False), dan wordt elk type apart geteld: een overschot aan type 1 vult geen type II-lijnen.
    Elke aanwerving gaat naar de component met het grootste FTE-tekort."""
    need = sum(c["lines"] for c in comps)
    need_1 = sum(c["lines"] for c in comps if c["type"] == 1)
    have = sum(n["rate"] for n in nurses)
    have_1 = sum(n["rate"] for n in nurses if n["type"] == 1)
    hire_1 = max(0.0, need_1 - have_1)
    if TYPE1_ON_TYPE2:
        hire_2 = max(0.0, need - have - hire_1)
    else:
        hire_2 = max(0.0, (need - need_1) - (have - have_1))

    hires = []
    for typ, amount in ((1, hire_1), (2, hire_2)):
        for rate in contracts(amount):
            n = {"id": f"NEW{dept}{len(hires) + 1:03d}", "rate": rate, "type": typ, "text": "nieuwe aanwerving",
                 "pref": parse_preference(""), "new": True, "score": 0, "status": "aanwerving"}
            allowed = [i for i, c in enumerate(comps)
                       if c["type"] == typ or (typ == 1 and c["type"] == 2 and TYPE1_ON_TYPE2)]
            # grootste FTE-tekort eerst; bij gelijk tekort: eigen type, dan een lijn die nog leeg is
            i = max(allowed, key=lambda i: (comps[i]["lines"] - fte(comps[i]), comps[i]["type"] == typ,
                                            comps[i]["lines"] - len(comps[i]["members"]), -i))
            comps[i]["members"].append(n)
            n["comp"] = i
            hires.append(n)
    return hires


# ============================================================================
# STAP 4 - UITVOER
# ============================================================================

def write_template_files(out, dept, roster, comps, patterns, shifts, monthly):
    os.makedirs(out, exist_ok=True)
    length = int(re.search(r"(\d+)h", SCENARIO).group(1))
    code = {s["name"]: k + 1 for k, s in enumerate(shifts)}     # 0 = vrij, 1..n = shift op startuur

    with open(os.path.join(out, f"Shift_system_dpt_{dept}.txt"), "w") as f:
        f.write(f"{len(shifts)}\t{length}\n")
        f.writelines(f"{s['start']}\n" for s in shifts)
        f.write("\t".join(str(s["min_buffer"] if REQ_WITH_BUFFER else s["min"]) for s in shifts) + "\n")

    with open(os.path.join(out, f"Cyclic_roster_dpt_{dept}.txt"), "w") as f:
        f.write(f"{len(roster)}\n")
        for n in roster:
            c = comps[n["comp"]]
            week = [code[c["shift"]] if patterns[c["pattern"]]["days"][d] else 0 for d in range(7)]
            f.write("\t".join(map(str, week * WEEKS)) + "\n")

    with open(os.path.join(out, f"Personnel_dpt_{dept}.txt"), "w") as f:
        for n in roster:
            prefs = [HIRE_PREF] * 140 if n["new"] else monthly[n["id"]]
            f.write(f"{n['id']}\t" + "\t".join(map(str, prefs)) + f"\t{n['rate']}\t{n['type']}\n")

    c = CONSTRAINTS
    with open(os.path.join(out, f"Constraints_dpt_{dept}.txt"), "w") as f:
        f.write(f"{c['min_ass']} {c['max_ass']}\n{c['min_cons_wrk']} {c['max_cons_wrk']}\n")
        f.writelines(f"{c['min_cons_shift']} {c['max_cons_shift']}\n" for _ in shifts)
        f.writelines(f"{c['min_shift']} {c['max_shift']}\n" for _ in shifts)
        f.write(f"{c['identical']}\n")


def add_sheet(wb, title, header, rows, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.append(header)
    for cell in ws[1]:
        cell.font, cell.fill = Font(bold=True), PatternFill("solid", fgColor="EEEEEE")
    for r in rows:
        ws.append(r)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or "")) for c in col) + 2, 48)
    ws.freeze_panes = "A2"


def write_overview(path, dept, summary, roster, comps, patterns, shifts, req):
    wb = openpyxl.Workbook()
    add_sheet(wb, "Samenvatting", ["", "Waarde"], summary, first=True)

    rows = []
    for n in roster:
        c = comps[n["comp"]]
        days = patterns[c["pattern"]]["days"]
        rows.append([n["id"], n["type"], n["rate"], n["text"], c["shift"], patterns[c["pattern"]]["label"],
                     n["score"], n["status"]] + [LETTER[c["shift"]] if days[d] else "-" for d in range(7)])
    add_sheet(wb, "Cyclisch rooster", ["Verpleegkundige", "Type", "Tewerkstelling", "Cyclische voorkeur", "Shift",
                                       "Patroon", "Strafscore", "Status"] + DAYS_NL, rows)

    rows = [[c["shift"], patterns[c["pattern"]]["label"], "I" if c["type"] == 1 else "II", c["lines"],
             len(c["members"]), fte(c), round(fte(c) - c["lines"], 2),
             ", ".join(m["id"] for m in c["members"])] for c in comps]
    add_sheet(wb, "Lijnen", ["Shift", "Patroon", "Type", "Lijnen (FTE nodig)", "Personen", "FTE ingevuld",
                             "Verschil (FTE)", "Wie"], rows)

    rows = []
    for s in shifts:
        for d in range(7):
            ppl = [n for n in roster if comps[n["comp"]]["shift"] == s["name"]
                   and patterns[comps[n["comp"]]["pattern"]]["days"][d]]
            rows.append([s["name"], DAYS_NL[d], len(ppl), sum(n["rate"] for n in ppl),
                         sum(1 for n in ppl if n["type"] == 1), req[s["name"]]["req"],
                         req[s["name"]]["min_type1"], s["min"]])
    add_sheet(wb, "Dekking", ["Shift", "Dag", "Personen (cyclisch rooster)", "FTE (na vrije dagen deeltijdsen)",
                              "Type I", "Eis Case D (met buffer)", "Min. type I", "Minimum Case C"], rows)
    wb.save(path)


# ============================================================================
# HOOFDPROGRAMMA
# ============================================================================

def write_readable_roster(path, dept, roster, comps, patterns, shifts):
    """Cyclisch rooster met koppen, om zelf te lezen. Zelfde codes als Cyclic_roster_dpt_X.txt."""
    code = {s["name"]: k + 1 for k, s in enumerate(shifts)}
    days = [f"{DAYS_NL[d % 7]}{d + 1}" for d in range(7 * WEEKS)]
    rows = []
    for n in roster:
        c = comps[n["comp"]]
        week = [code[c["shift"]] if patterns[c["pattern"]]["days"][d] else 0 for d in range(7)]
        rows.append((n, week * WEEKS))

    def line(head, values):          # 4 kolommen per dag, '|' tussen de weken
        cells = [f"{v:>4}" for v in values]
        weeks = [" ".join(cells[w * 7:(w + 1) * 7]) for w in range(WEEKS)]
        return head + " | " + " | ".join(weeks) + "\n"

    legend = ", ".join(f"{code[s['name']]} = {s['name']} ({s['start']:02d}:00)" for s in shifts)
    own = sum(1 for n in roster if not n["new"])
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"Cyclisch rooster afdeling {dept} ({SCENARIO}, {POLICY})\n")
        f.write(f"Codes: 0 = vrij, {legend}\n")
        f.write(f"{len(roster)} personen: {own} eigen personeel + {len(roster) - own} aanwervingen (NEW...)\n")
        f.write("Elke week is dezelfde standaardweek (cyclisch rooster = 1 week x 4).\n\n")
        head = f"{'Regel':>5}  {'Persoon':<9} {'Type':>4} {'FTE':>5}"
        f.write(line(head, days))
        f.write("-" * (len(head) + 3 + WEEKS * (7 * 5 + 2)) + "\n")
        for i, (n, d28) in enumerate(rows, start=2):   # regelnummer in Cyclic_roster_dpt_X.txt
            f.write(line(f"{i:>5}  {n['id']:<9} {n['type']:>4} {n['rate']:>5.2f}", d28))
        f.write("-" * (len(head) + 3 + WEEKS * (7 * 5 + 2)) + "\n")
        for s_ in shifts:
            k = code[s_["name"]]
            count = [sum(1 for _, d28 in rows if d28[d] == k) for d in range(7 * WEEKS)]
            f.write(line(f"{'':>5}  {'# ' + s_['name']:<9} {'min':>4} {s_['min']:>5}", count))
        f.write("\n'Regel' = regelnummer van die persoon in files/Cyclic_roster_dpt_" + dept + ".txt.\n")
        f.write("Onderaan: hoeveel mensen er die dag op die shift staan, tegenover het minimum uit Case C.\n")


def run_department(wb, dept, out_dir, overview_dir):
    ws = wb[SHEET_CASE_D]
    patterns = read_patterns(ws)
    shifts = read_shifts(ws, dept)
    comps, req = read_lines(ws, dept)
    nurses = read_nurses(wb, dept)
    monthly = read_monthly_preferences(wb, dept)

    assign_own(nurses, comps, patterns)          # stap 2
    hires = hire(nurses, comps, dept)            # stap 3
    roster = nurses + hires

    # controles
    need = sum(c["lines"] for c in comps)
    total = sum(n["rate"] for n in roster)
    assert total >= need - EPS, "te weinig FTE"
    assert sum(n["rate"] for n in roster if n["type"] == 1) >= sum(c["lines"] for c in comps if c["type"] == 1) - EPS
    assert all(score(n, comps[n["comp"]], patterns) is not None for n in nurses), "contractregel geschonden"
    assert TYPE1_ON_TYPE2 or all(n["type"] == comps[n["comp"]]["type"] for n in roster), "iemand staat op een lijn van een ander type"

    have = sum(n["rate"] for n in nurses)
    have_1 = sum(n["rate"] for n in nurses if n["type"] == 1)
    have_2 = sum(n["rate"] for n in nurses if n["type"] == 2)
    need_1 = sum(c["lines"] for c in comps if c["type"] == 1)
    summary = [
        ["Afdeling", dept],
        ["Lijnen uit Case D (= FTE nodig)", need],
        ["   waarvan type I", sum(c["lines"] for c in comps if c["type"] == 1)],
        ["Eigen personeel (personen)", len(nurses)],
        ["Eigen personeel (FTE)", have],
        ["Tekort (FTE; negatief = overschot)", need - have],
        ["   tekort type I (FTE)", need_1 - have_1],
        ["   tekort type II (FTE)", (need - need_1) - have_2],
        ["Aanwervingen (personen)", len(hires)],
        ["Aanwervingen (FTE)", sum(h["rate"] for h in hires)],
        ["   contracten", ", ".join(f"{sum(1 for h in hires if h['rate'] == r)} x {r}"
                                    for r in sorted({h["rate"] for h in hires}, reverse=True)) or "-"],
        ["Personen in het cyclisch rooster", len(roster)],
        ["FTE in het cyclisch rooster", total],
        ["Eigen personeel met strafscore 0", f"{sum(1 for n in nurses if n['score'] == 0)} van {len(nurses)}"],
        ["Totale strafscore eigen personeel", sum(n["score"] for n in nurses)],
        ["Eigen personeel extra op een lijn (paste nergens anders)",
         sum(1 for n in nurses if "extra" in n["status"])],
    ]
    write_template_files(out_dir, dept, roster, comps, patterns, shifts, monthly)
    overview = os.path.join(overview_dir, f"spoor1_overzicht_dpt_{dept}.xlsx")
    write_overview(overview, dept, summary, roster, comps, patterns, shifts, req)
    write_readable_roster(os.path.join(overview_dir, f"Cyclic_roster_dpt_{dept}_leesbaar.txt"),
                          dept, roster, comps, patterns, shifts)
    return summary


def project_root():
    """De projectmap: vanaf de map van dit script naar boven, de eerste map met een submap data/."""
    d = os.path.dirname(os.path.abspath(__file__))
    for _ in range(4):
        if os.path.isdir(os.path.join(d, DATA_DIR)):
            return d
        d = os.path.dirname(d)
    raise SystemExit(f"Geen map '{DATA_DIR}/' gevonden boven {os.path.abspath(__file__)}.\n"
                     f"Zet dit script in src/ van het project (naast data/, files/, results/).")


def main():
    root = project_root()
    ap = argparse.ArgumentParser(description="Case E spoor 1: cyclisch rooster met namen")
    ap.add_argument("--excel", default=os.path.join(root, DATA_DIR, EXCEL_FILE))
    ap.add_argument("--dept", default=DEPARTMENTS, help="bv. A of ABCD")
    ap.add_argument("--out", default=os.path.join(root, FILES_DIR))
    ap.add_argument("--results", default=os.path.join(root, RESULTS_DIR))
    args = ap.parse_args()

    if not os.path.exists(args.excel):
        raise SystemExit(f"Excel niet gevonden: {args.excel}\n"
                         f"Zet '{EXCEL_FILE}' in de map {DATA_DIR}/ (of pas EXCEL_FILE bovenaan aan).")
    wb = openpyxl.load_workbook(args.excel, data_only=True)
    if SHEET_CASE_D not in wb.sheetnames:
        raise SystemExit(f"Tabblad '{SHEET_CASE_D}' ontbreekt in {args.excel}. Is dit de juiste versie van de Excel?")
    os.makedirs(args.results, exist_ok=True)

    results = {d: run_department(wb, d, args.out, args.results) for d in args.dept.upper()}

    depts = list(results)
    w = 2 + max(len(str(v)) for d in depts for _, v in results[d])   # kolombreedte
    print(f"\nSpoor 1 ({SCENARIO}, {POLICY})\n" + "-" * (58 + w * len(depts)))
    print(f"{'':<58}" + "".join(f"{d:>{w}}" for d in depts))
    for k in range(1, len(results[depts[0]])):
        label = results[depts[0]][k][0]
        print(f"{label:<58}" + "".join(f"{str(results[d][k][1]):>{w}}" for d in depts))
    print(f"\nExcel gelezen uit:         {args.excel}")
    print(f"Invoer voor roster.exe in: {args.out}")
    print(f"Overzicht-Excels in:       {args.results}")


if __name__ == "__main__":
    main()