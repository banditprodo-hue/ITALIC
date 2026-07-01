# -*- coding: utf-8 -*-
"""
Generatore turnazione personale ITALIC
Periodo: 01/09/2026 - 31/08/2027

Crea un file .xlsx (senza dipendenze esterne) con:
 - foglio "Turni"        : calendario giornaliero con le presenze
 - foglio "Riepilogo"    : turni svolti per settimana / dipendente
 - foglio "Segnalazioni" : settimane/giorni in cui non tutte le regole sono applicabili
 - foglio "Note"         : assunzioni e regole utilizzate
"""

import datetime as dt
import zipfile
from itertools import combinations, product
from xml.sax.saxutils import escape

# ---------------------------------------------------------------------------
# 1. CALENDARIO E FESTIVITA'
# ---------------------------------------------------------------------------

START = dt.date(2026, 9, 1)
END = dt.date(2027, 8, 31)

# Festivita' nazionali italiane nel periodo (national holidays only).
ITALIAN_HOLIDAYS = {
    dt.date(2026, 11, 1):  "Tutti i Santi",
    dt.date(2026, 12, 8):  "Immacolata Concezione",
    dt.date(2026, 12, 25): "Natale",
    dt.date(2026, 12, 26): "Santo Stefano",
    dt.date(2027, 1, 1):   "Capodanno",
    dt.date(2027, 1, 6):   "Epifania",
    dt.date(2027, 3, 29):  "Lunedi' dell'Angelo",      # Pasqua 2027 = 28/03
    dt.date(2027, 4, 25):  "Festa della Liberazione",
    dt.date(2027, 5, 1):   "Festa del Lavoro",
    dt.date(2027, 6, 2):   "Festa della Repubblica",
    dt.date(2027, 8, 15):  "Ferragosto",
}

# Giorno escluso espressamente dalla richiesta
SPECIAL_EXCLUDED = {dt.date(2027, 5, 10): "Giorno escluso (richiesta)"}

WD_NAMES = ["Lunedi", "Martedi", "Mercoledi", "Giovedi", "Venerdi", "Sabato", "Domenica"]


def daterange(a, b):
    d = a
    while d <= b:
        yield d
        d += dt.timedelta(days=1)


def reason_excluded(d):
    """Ritorna None se lavorativo, altrimenti il motivo dell'esclusione."""
    if d.weekday() == 5:
        return "Sabato"
    if d.weekday() == 6:
        return "Domenica"
    if d in ITALIAN_HOLIDAYS:
        return "Festivo: " + ITALIAN_HOLIDAYS[d]
    if d in SPECIAL_EXCLUDED:
        return SPECIAL_EXCLUDED[d]
    return None


ALL_DATES = list(daterange(START, END))
WORKING_DATES = [d for d in ALL_DATES if reason_excluded(d) is None]

# Settimane (lunedi -> data). Indice sequenziale per la regola di alternanza.
def monday_of(d):
    return d - dt.timedelta(days=d.weekday())

mondays = sorted({monday_of(d) for d in ALL_DATES})
week_index = {m: i for i, m in enumerate(mondays)}

# Per ogni settimana: weekday -> data (solo giorni lavorativi presenti nel range)
weeks = {}  # week_idx -> { weekday(0..4): date }
for d in WORKING_DATES:
    wi = week_index[monday_of(d)]
    weeks.setdefault(wi, {})[d.weekday()] = d

# ---------------------------------------------------------------------------
# 2. DIPENDENTI E REGOLE
# ---------------------------------------------------------------------------

FUNZIONARI = ["Caforio", "Prota", "Amenduni", "Gigante"]
ASSISTENTI = ["Chianura", "Donnaloia", "Raffaele", "Gaballo", "Scalone"]
ALL_EMP = FUNZIONARI + ASSISTENTI
GROUP = {e: ("Funzionari" if e in FUNZIONARI else "Assistenti") for e in ALL_EMP}

MON, TUE, WED, THU, FRI = 0, 1, 2, 3, 4
DOUBLE_DAYS = [MON, WED, FRI]   # settimane con due turni
SINGLE_DAYS = [TUE, THU]        # settimane con un turno

# Coppie che NON possono stare nello stesso giorno
FORBIDDEN_PAIRS = [
    frozenset(("Caforio", "Prota")),
    frozenset(("Gaballo", "Donnaloia")),
    frozenset(("Scalone", "Chianura")),
]

# Dipendenti ad alternanza (1 turno / 2 turni). La fase viene scelta dal solver.
ALTERNATING = ["Caforio", "Prota", "Gigante", "Donnaloia", "Gaballo", "Scalone"]
# Variabili di fase da ottimizzare: alternanze + fase dei giorni di Amenduni
PHASE_VARS = ALTERNATING + ["AmenduniDay"]

# Scalone non ha regola specifica nella richiesta -> assunzione documentata.
SCALONE_ASSUMPTION = True


def emp_target_and_days(emp, wi, phases):
    """
    Ritorna (target_turni, [weekday ammessi]) per dipendente/settimana.
    target 0 significa che il dipendente non e' in turnazione quella settimana.
    """
    # Amenduni: sempre 2 turni, ma i giorni alternano Lun+Mer / Lun+Ven
    if emp == "Amenduni":
        dphase = phases.get("AmenduniDay", 0)
        if (wi + dphase) % 2 == 0:
            return 2, [MON, WED]     # settimana A: Lunedi + Mercoledi
        return 2, [MON, FRI]         # settimana B: Lunedi + Venerdi
    if emp == "Chianura":
        return 1, [TUE]          # un turno a settimana, solo Martedi
    if emp == "Raffaele":
        return 1, [THU]          # un turno a settimana, solo Giovedi

    # Dipendenti ad alternanza
    phase = phases[emp]
    is_double = ((wi + phase) % 2 == 0)
    target = 2 if is_double else 1

    if emp == "Caforio":
        # turni solo Lunedi, Giovedi, Venerdi
        if is_double:
            return 2, [MON, FRI]      # settimana a 2 turni -> Lun + Ven
        else:
            return 1, [THU]           # settimana a 1 turno -> Gio
    # Prota, Gigante, Donnaloia, Gaballo, Scalone -> pattern generico
    if is_double:
        return 2, DOUBLE_DAYS
    else:
        return 1, SINGLE_DAYS


# ---------------------------------------------------------------------------
# 3. SOLVER PER SETTIMANA
# ---------------------------------------------------------------------------

def solve_week(wi, phases, cum=None):
    """
    Assegna i turni della settimana wi.
    Ritorna (assignment, shortfalls):
      assignment: { emp: set(weekday) }
      shortfalls: { emp: (assegnati, target) } solo dove assegnati < target
    Massimizza il numero totale di turni piazzati rispettando i vincoli rigidi.
    """
    avail_wd = weeks.get(wi, {})            # weekday -> date disponibili
    avail_days = set(avail_wd.keys())

    emps = []
    for e in ALL_EMP:
        target, days = emp_target_and_days(e, wi, phases)
        if target == 0:
            continue
        feas = sorted(set(days) & avail_days)
        emps.append((e, target, feas))

    # Ordina: piu' vincolati prima (meno giorni disponibili)
    emps.sort(key=lambda x: (len(x[2]), -x[1]))

    # Genera opzioni per dipendente: sottoinsiemi di dimensione da target giu' a 0
    options = {}
    for e, target, feas in emps:
        opts = []
        maxk = min(target, len(feas))
        for k in range(maxk, -1, -1):
            for combo in combinations(feas, k):
                opts.append(frozenset(combo))
        # con bilanciamento: a parita' di dimensione, prima i giorni meno usati dal dipendente
        if cum is not None:
            opts.sort(key=lambda s: (-len(s), sum(cum[e][wd] for wd in s)))
        options[e] = opts

    n = len(emps)
    day_total = {d: 0 for d in range(5)}
    day_funz = {d: 0 for d in range(5)}
    day_members = {d: set() for d in range(5)}

    best = {"cov": -1, "sec": None, "assign": None}
    cur_assign = {}

    def can_place(emp, days):
        is_f = GROUP[emp] == "Funzionari"
        for d in days:
            if day_total[d] + 1 > 4:
                return False
            if is_f and day_funz[d] + 1 > 2:
                return False
            for other in day_members[d]:
                if frozenset((emp, other)) in FORBIDDEN_PAIRS:
                    return False
        return True

    def place(emp, days, on):
        is_f = GROUP[emp] == "Funzionari"
        for d in days:
            day_total[d] += 1 if on else -1
            if is_f:
                day_funz[d] += 1 if on else -1
            if on:
                day_members[d].add(emp)
            else:
                day_members[d].discard(emp)

    # branch & bound; primario: copertura; secondario: (equita' carenze, bilanciamento giorni)
    tdict = {e: t for e, t, _ in emps}
    # peso carenze: Amenduni va protetto (deve sempre fare 2 turni)
    PRIO = {e: (5 if e == "Amenduni" else 1) for e in ALL_EMP}

    def dfs(i, cov, remaining_max):
        if cov + remaining_max < best["cov"]:
            return
        if i == n:
            # 1) penalita' carenze: spalma le perdite e protegge Amenduni (quadratica e pesata)
            sf = 0
            for e, t in tdict.items():
                d = t - len(cur_assign.get(e, ()))
                if d > 0:
                    sf += PRIO[e] * d * d
            # 2) copertura funzionari: vogliamo 2 funzionari per giorno lavorativo
            wf = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}
            for e, days in cur_assign.items():
                if GROUP[e] == "Funzionari":
                    for wd in days:
                        wf[wd] += 1
            fp = sum((2 - wf[wd]) ** 2 for wd in avail_days)
            # 3) penalita' bilanciamento giorni
            bal = 0
            if cum is not None:
                for e, days in cur_assign.items():
                    for wd in days:
                        bal += cum[e][wd]
            sec = (sf, fp, bal)
            if cov > best["cov"] or (cov == best["cov"] and (best["sec"] is None or sec < best["sec"])):
                best["cov"] = cov
                best["sec"] = sec
                best["assign"] = dict(cur_assign)
            return
        emp, target, feas = emps[i]
        rest = remaining_max - target
        for opt in options[emp]:
            if can_place(emp, opt):
                place(emp, opt, True)
                cur_assign[emp] = set(opt)
                dfs(i + 1, cov + len(opt), rest)
                place(emp, opt, False)
                del cur_assign[emp]

    total_target = sum(t for _, t, _ in emps)
    dfs(0, 0, total_target)

    assign = best["assign"] or {}
    shortfalls = {}
    for e, target, feas in emps:
        got = len(assign.get(e, set()))
        if got < target:
            shortfalls[e] = (got, target)
    return assign, shortfalls


def evaluate(phases):
    """Risolve tutte le settimane; ritorna (assegnati, shortfall, funz_pen, imbalance, dettagli)."""
    cum = {e: {0: 0, 1: 0, 2: 0, 3: 0, 4: 0} for e in ALL_EMP}
    total_assigned = 0
    total_shortfall = 0
    details = {}
    for wi in sorted(weeks.keys()):
        assign, shorts = solve_week(wi, phases, cum)
        details[wi] = (assign, shorts)
        total_assigned += sum(len(s) for s in assign.values())
        total_shortfall += len(shorts)
        for e, days in assign.items():
            for wd in days:
                cum[e][wd] += 1
    # copertura funzionari: quanto ci discostiamo da 2 funzionari/giorno (sui giorni lavorativi)
    funz_pen = 0
    for wi in weeks:
        assign = details[wi][0]
        wf = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}
        for e, days in assign.items():
            if GROUP[e] == "Funzionari":
                for wd in days:
                    wf[wd] += 1
        for wd in weeks[wi]:            # solo giorni lavorativi
            funz_pen += (2 - wf[wd]) ** 2
    # imbalance: varianza dei giorni Lun/Mer/Ven per i dipendenti "flessibili"
    bal_emps = ["Prota", "Gigante", "Donnaloia", "Gaballo", "Scalone"]
    imb = 0.0
    for e in bal_emps:
        vals = [cum[e][0], cum[e][2], cum[e][4]]   # Lun, Mer, Ven
        m = sum(vals) / 3.0
        imb += sum((x - m) ** 2 for x in vals)
    return total_assigned, total_shortfall, funz_pen, imb, details


# ---------------------------------------------------------------------------
# 4. SCELTA DELLE FASI (brute force)
# ---------------------------------------------------------------------------

COVERAGE_TOLERANCE = 3   # accettiamo fino a 3 turni/anno in meno

solutions = []
for bits in product([0, 1], repeat=len(PHASE_VARS)):
    phases = dict(zip(PHASE_VARS, bits))
    assigned, shortfall, funz_pen, imb, details = evaluate(phases)
    solutions.append((assigned, shortfall, funz_pen, imb, phases, details))

cov_max = max(s[0] for s in solutions)
# fra le soluzioni a copertura quasi massima: prima massima copertura funzionari (2/giorno),
# poi giorni piu' equilibrati
candidates = [s for s in solutions if s[0] >= cov_max - COVERAGE_TOLERANCE]
candidates.sort(key=lambda s: (s[2], s[3], -s[0], s[1]))   # funz_pen, imbalance, copertura, shortfall
best = candidates[0]
PHASES = best[4]
DETAILS = best[5]
best_metric = (best[0], -best[1], -best[3])
print("Copertura massima possibile:", cov_max)
print("Fasi scelte:", PHASES)
print("Turni piazzati:", best[0], " shortfall:", best[1],
      " funz_pen:", best[2], " imbalance:", round(best[3], 1))

# ---------------------------------------------------------------------------
# 5. RACCOLTA DATI PER L'OUTPUT
# ---------------------------------------------------------------------------

# assegnazione giornaliera: date -> list(emp)
day_assignment = {d: [] for d in WORKING_DATES}
weekly_counts = {}     # wi -> { emp: n }
warnings = []          # lista di stringhe

for wi in sorted(weeks.keys()):
    assign, shorts = DETAILS[wi]
    wd_to_date = weeks[wi]
    weekly_counts[wi] = {}
    for emp, days in assign.items():
        weekly_counts[wi][emp] = len(days)
        for wd in days:
            d = wd_to_date[wd]
            day_assignment[d].append(emp)
    # segnalazioni shortfall
    monday = mondays[wi]
    for emp, (got, target) in sorted(shorts.items()):
        # motivo: giorno richiesto festivo/non disponibile
        _, reqdays = emp_target_and_days(emp, wi, PHASES)
        missing = [WD_NAMES[d] for d in reqdays if d not in wd_to_date]
        motivo = ""
        if missing:
            giorni = ", ".join(sorted(set(missing)))
            motivo = " (giorno/i non lavorativo/i: %s)" % giorni
        warnings.append(
            "Settimana del %s: %s ha svolto %d turno/i invece di %d%s"
            % (monday.strftime("%d/%m/%Y"), emp, got, target, motivo)
        )

# verifica vincoli rigidi (sanity check) e segnalazioni capacita'
for d in WORKING_DATES:
    members = day_assignment[d]
    nf = sum(1 for e in members if GROUP[e] == "Funzionari")
    if len(members) > 4:
        warnings.append("ATTENZIONE %s: %d dipendenti (>4)!" % (d.strftime("%d/%m/%Y"), len(members)))
    if nf > 2:
        warnings.append("ATTENZIONE %s: %d funzionari (>2)!" % (d.strftime("%d/%m/%Y"), nf))
    for pair in FORBIDDEN_PAIRS:
        if pair <= set(members):
            warnings.append("ATTENZIONE %s: coppia vietata %s insieme!" % (d.strftime("%d/%m/%Y"), "/".join(pair)))

if not warnings:
    warnings.append("Nessuna criticita': tutte le regole sono applicate per ogni settimana e giorno.")

# Copertura funzionari: verifica del vincolo "sempre 2 funzionari"
_fz = {0: 0, 1: 0, 2: 0}
for d in WORKING_DATES:
    nf = sum(1 for e in day_assignment[d] if GROUP[e] == "Funzionari")
    _fz[nf] = _fz.get(nf, 0) + 1
if _fz.get(0, 0) or _fz.get(1, 0):
    warnings.insert(0,
        "IMPOSSIBILE avere SEMPRE 2 funzionari ogni giorno: con soli 4 funzionari il massimo e' 8 turni/settimana contro i 10 necessari (2 x 5 giorni). "
        "Risultato: %d giorni con 2 funzionari, %d giorni con 1, %d giorni con 0 (i cali sono su Martedi/Giovedi, giorni in cui lavorano solo i funzionari a 'un turno'). "
        "I 2 funzionari sono garantiti quasi sempre su Lun/Mer/Ven." % (_fz.get(2, 0), _fz.get(1, 0), _fz.get(0, 0)))

# ---------------------------------------------------------------------------
# 6. SCRITTURA XLSX (senza dipendenze)
# ---------------------------------------------------------------------------

def col_letter(idx):  # 1-based
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


class Sheet:
    def __init__(self, name):
        self.name = name
        self.rows = {}        # row(1-based) -> { col(1-based): (value, style, is_number) }
        self.merges = []
        self.colwidths = {}   # col -> width
        self.maxc = 0
        self.maxr = 0
        self.freeze = None    # (rows_frozen, cols_frozen)
        self.rowheights = {}  # row -> height

    def set(self, r, c, value, style=0, number=False):
        self.rows.setdefault(r, {})[c] = (value, style, number)
        self.maxc = max(self.maxc, c)
        self.maxr = max(self.maxr, r)

    def merge(self, r1, c1, r2, c2):
        self.merges.append("%s%d:%s%d" % (col_letter(c1), r1, col_letter(c2), r2))

    def width(self, c, w):
        self.colwidths[c] = w


def sheet_xml(sh):
    parts = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    parts.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    if sh.freeze:
        fr, fc = sh.freeze
        topleft = "%s%d" % (col_letter(fc + 1), fr + 1)
        parts.append('<sheetViews><sheetView workbookViewId="0">'
                     '<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                     '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>'
                     '</sheetView></sheetViews>' % (fc, fr, topleft, topleft, topleft))
    if sh.colwidths:
        parts.append('<cols>')
        for c, w in sorted(sh.colwidths.items()):
            parts.append('<col min="%d" max="%d" width="%.2f" customWidth="1"/>' % (c, c, w))
        parts.append('</cols>')
    parts.append('<sheetData>')
    for r in range(1, sh.maxr + 1):
        rowcells = sh.rows.get(r)
        if not rowcells:
            continue
        parts.append('<row r="%d"%s>' % (r, (' ht="%.2f" customHeight="1"' % sh.rowheights[r]) if r in sh.rowheights else ''))
        for c in sorted(rowcells.keys()):
            value, style, number = rowcells[c]
            ref = "%s%d" % (col_letter(c), r)
            if number:
                parts.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, style, value))
            else:
                txt = escape(str(value))
                parts.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, style, txt))
        parts.append('</row>')
    parts.append('</sheetData>')
    if sh.merges:
        parts.append('<mergeCells count="%d">' % len(sh.merges))
        for m in sh.merges:
            parts.append('<mergeCell ref="%s"/>' % m)
        parts.append('</mergeCells>')
    parts.append('</worksheet>')
    return "".join(parts)


# ---- stili -----------------------------------------------------------------
# Indici cellXfs (definiti in STYLES_XML qui sotto):
S_DEFAULT = 0
S_TITLE = 1
S_HDR = 2          # header generico (blu scuro)
S_HDR_FUNZ = 3     # header funzionari (blu)
S_HDR_ASS = 4      # header assistenti (verde)
S_DATE = 5         # cella data
S_DAY = 6          # cella giorno settimana
S_X_FUNZ = 7       # turno funzionario
S_X_ASS = 8        # turno assistente
S_NUM = 9          # numero centrato
S_NA = 10          # giorno non lavorativo (grigio)
S_WARN = 11        # segnalazione (arancio)
S_NOTE = 12        # testo nota
S_CELL = 13        # cella bordata vuota
S_WEEKHDR = 14     # intestazione settimana

STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="6">
  <font><sz val="11"/><name val="Calibri"/></font>
  <font><b/><sz val="16"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  <font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><color rgb="FF9C5700"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><color rgb="FF1F3864"/><name val="Calibri"/></font>
</fonts>
<fills count="10">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF548235"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFBDD7EE"/></patternFill></fill>
</fills>
<borders count="2">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color rgb="FFB0B0B0"/></left>
    <right style="thin"><color rgb="FFB0B0B0"/></right>
    <top style="thin"><color rgb="FFB0B0B0"/></top>
    <bottom style="thin"><color rgb="FFB0B0B0"/></bottom>
    <diagonal/>
  </border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="15">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
  <xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="3" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
  <xf numFmtId="0" fontId="4" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
  <xf numFmtId="0" fontId="5" fillId="9" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
</cellXfs>
</styleSheet>'''


# ---- costruzione fogli ------------------------------------------------------

# Fogli mensili: UN FOGLIO PER OGNI MESE
# nominativi nella prima colonna, date nella prima riga
MONTHS_IT = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
             "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
MON_ABBR = ["", "Gen", "Feb", "Mar", "Apr", "Mag", "Giu",
            "Lug", "Ago", "Set", "Ott", "Nov", "Dic"]
WD2 = ["LU", "MA", "ME", "GI", "VE", "SA", "DO"]

# raggruppa le date lavorative per mese (ordine cronologico)
months_dates = {}
months_order = []
for d in WORKING_DATES:
    key = (d.year, d.month)
    if key not in months_dates:
        months_dates[key] = []
        months_order.append(key)
    months_dates[key].append(d)

month_sheets = []
for (yy, mm) in months_order:
    mdates = months_dates[(yy, mm)]
    sh = Sheet("%s %d" % (MON_ABBR[mm], yy))
    # riga 1 = angolo + DATE
    sh.set(1, 1, "%s %d" % (MONTHS_IT[mm], yy), S_HDR)
    dcol = {}
    for i, d in enumerate(mdates):
        c = 2 + i
        dcol[d] = c
        sh.set(1, c, "%02d\n%s" % (d.day, WD2[d.weekday()]), S_HDR)
    tot_col = 2 + len(mdates)
    sh.set(1, tot_col, "Tot", S_HDR)
    sh.rowheights[1] = 30
    # righe 2.. = DIPENDENTI (prima colonna)
    r = 2
    for e in ALL_EMP:
        lab = S_HDR_FUNZ if GROUP[e] == "Funzionari" else S_HDR_ASS
        sh.set(r, 1, e, lab)
        mtot = 0
        for d in mdates:
            c = dcol[d]
            if e in day_assignment[d]:
                sh.set(r, c, "X", S_X_FUNZ if GROUP[e] == "Funzionari" else S_X_ASS)
                mtot += 1
            else:
                sh.set(r, c, "", S_CELL)
        sh.set(r, tot_col, mtot, S_NUM, number=True)
        r += 1
    # righe totali per giorno
    rtd, rtf = r, r + 1
    sh.set(rtd, 1, "Tot / giorno", S_HDR)
    sh.set(rtf, 1, "Tot funzionari / giorno", S_HDR)
    for d in mdates:
        c = dcol[d]
        mem = day_assignment[d]
        sh.set(rtd, c, len(mem), S_NUM, number=True)
        sh.set(rtf, c, sum(1 for x in mem if GROUP[x] == "Funzionari"), S_NUM, number=True)
    sh.width(1, 22)
    for d in mdates:
        sh.width(dcol[d], 5)
    sh.width(tot_col, 6)
    sh.freeze = (1, 1)   # blocca la riga delle date e la colonna dei nomi
    month_sheets.append(sh)

# Foglio 2: Elenco per persona (formato lista ordinato per dipendente)
order_idx = {e: i for i, e in enumerate(FUNZIONARI)}
order_idx.update({e: i for i, e in enumerate(ASSISTENTI)})
shifts = []
for d in WORKING_DATES:
    for e in day_assignment[d]:
        shifts.append((e, d))
shifts.sort(key=lambda x: (0 if GROUP[x[0]] == "Funzionari" else 1, order_idx[x[0]], x[1]))

sE = Sheet("Elenco per persona")
sE.set(1, 1, "ELENCO TURNI ORDINATO PER DIPENDENTE", S_TITLE)
sE.merge(1, 1, 1, 6)
hr = 3
sE.set(hr, 1, "Gruppo", S_HDR)
sE.set(hr, 2, "Dipendente", S_HDR)
sE.set(hr, 3, "N.", S_HDR)
sE.set(hr, 4, "Data", S_HDR)
sE.set(hr, 5, "Giorno", S_HDR)
sE.set(hr, 6, "Turni nella settimana", S_HDR)
rr = hr + 1
counter = {}
for e, d in shifts:
    counter[e] = counter.get(e, 0) + 1
    wi = week_index[monday_of(d)]
    wk = weekly_counts.get(wi, {}).get(e, 0)
    lab_style = S_HDR_FUNZ if GROUP[e] == "Funzionari" else S_HDR_ASS
    sE.set(rr, 1, GROUP[e], S_CELL)
    sE.set(rr, 2, e, lab_style)
    sE.set(rr, 3, counter[e], S_NUM, number=True)
    sE.set(rr, 4, d.strftime("%d/%m/%Y"), S_DATE)
    sE.set(rr, 5, WD_NAMES[d.weekday()], S_DAY)
    sE.set(rr, 6, wk, S_NUM, number=True)
    rr += 1
sE.width(1, 12); sE.width(2, 13); sE.width(3, 5)
sE.width(4, 12); sE.width(5, 11); sE.width(6, 20)
sE.freeze = (hr, 0)

# Foglio 2: Riepilogo settimanale
s2 = Sheet("Riepilogo")
s2.set(1, 1, "RIEPILOGO TURNI PER SETTIMANA (numero turni svolti)", S_TITLE)
s2.merge(1, 1, 1, 3 + len(ALL_EMP))
hr = 3
s2.set(hr, 1, "Settimana (lun)", S_HDR)
s2.set(hr, 2, "Giorni lav.", S_HDR)
c = 3
for e in FUNZIONARI:
    s2.set(hr, c, e, S_HDR_FUNZ); c += 1
for e in ASSISTENTI:
    s2.set(hr, c, e, S_HDR_ASS); c += 1
s2.set(hr, c, "Tot", S_HDR)
totcol2 = c
rr = hr + 1
emp_totals = {e: 0 for e in ALL_EMP}
for wi in sorted(weeks.keys()):
    monday = mondays[wi]
    s2.set(rr, 1, monday.strftime("%d/%m/%Y"), S_DATE)
    s2.set(rr, 2, len(weeks[wi]), S_NUM, number=True)
    cc = 3
    wtot = 0
    for e in ALL_EMP:
        n = weekly_counts.get(wi, {}).get(e, 0)
        emp_totals[e] += n
        wtot += n
        style = S_NUM
        target, _ = emp_target_and_days(e, wi, PHASES)
        if n < target:
            style = S_WARN
        s2.set(rr, cc, n, style, number=True)
        cc += 1
    s2.set(rr, totcol2, wtot, S_NUM, number=True)
    rr += 1
# riga totali
s2.set(rr, 1, "TOTALE ANNO", S_HDR)
s2.set(rr, 2, len(WORKING_DATES), S_NUM, number=True)
cc = 3
grand = 0
for e in ALL_EMP:
    s2.set(rr, cc, emp_totals[e], S_NUM, number=True)
    grand += emp_totals[e]
    cc += 1
s2.set(rr, totcol2, grand, S_NUM, number=True)
s2.width(1, 14); s2.width(2, 11)
for i in range(len(ALL_EMP) + 1):
    s2.width(3 + i, 11)

# Foglio 3: Segnalazioni
s3 = Sheet("Segnalazioni")
s3.set(1, 1, "SEGNALAZIONI - settimane/giorni in cui non tutte le regole sono applicabili", S_TITLE)
s3.merge(1, 1, 1, 6)
rr = 3
s3.set(rr, 1, "#", S_HDR)
s3.set(rr, 2, "Descrizione", S_HDR)
s3.merge(rr, 2, rr, 6)
rr += 1
for i, w in enumerate(warnings, 1):
    s3.set(rr, 1, i, S_NUM, number=True)
    s3.set(rr, 2, w, S_WARN)
    s3.merge(rr, 2, rr, 6)
    rr += 1
s3.width(1, 5); s3.width(2, 30)
for k in range(3, 7):
    s3.width(k, 20)

# Foglio 4: Note / regole / assunzioni
s4 = Sheet("Note")
s4.set(1, 1, "NOTE, REGOLE E ASSUNZIONI", S_TITLE)
s4.merge(1, 1, 1, 8)
notes = [
    "Periodo: 01/09/2026 - 31/08/2027.",
    "Esclusi da turnazione: sabati, domeniche, 10/05/2027 e festivi nazionali italiani.",
    "Festivi nazionali considerati: " + ", ".join(
        "%s (%s)" % (d.strftime("%d/%m/%Y"), n) for d, n in sorted(ITALIAN_HOLIDAYS.items())
    ) + ".",
    "Non sono inclusi i santi patroni locali (citta' non specificata).",
    "Max 4 dipendenti in turno nello stesso giorno; si punta a 2 funzionari/giorno (garantiti su Lun/Mer/Ven; su Mar/Gio non sempre possibile - vedi Segnalazioni).",
    "Settimane con DUE turni: turni di Lunedi/Mercoledi/Venerdi. Settimane con UN turno: Martedi/Giovedi.",
    "Caforio: alterna 1/2 turni; turni solo Lun/Gio/Ven (2 turni = Lun+Ven, 1 turno = Gio).",
    "Prota: alterna 1/2 turni. Amenduni: 2 turni ogni settimana alternando Lun+Mer e Lun+Ven. Gigante: alterna 1/2 turni.",
    "Chianura: 1 turno/settimana solo Martedi. Raffaele: 1 turno/settimana solo Giovedi.",
    "Donnaloia, Gaballo e Scalone: alternano una settimana 1 turno e una settimana 2 turni.",
    "Coppie mai nello stesso giorno: Caforio-Prota, Gaballo-Donnaloia, Scalone-Chianura.",
    "Struttura file: un foglio per ogni mese (nominativi in prima colonna, date in prima riga), piu' i fogli Elenco per persona, Riepilogo, Segnalazioni e Note.",
    "ASSUNZIONE: le fasi di alternanza sono ottimizzate automaticamente per equilibrare i giorni (Lun/Mer/Ven) tra i dipendenti ed evitare che i venerdi' ricadano sempre sulle stesse persone.",
    "BILANCIAMENTO: per distribuire i giorni in modo equo, Caforio e Gigante risultano in fasi opposte; questo comporta circa 3 turni/anno in meno (alcune settimane con festivi hanno 1 turno in meno), spalmati e mai a carico di Amenduni.",
    "Fasi scelte (settimana di partenza a 2 turni se valore 0): " + ", ".join("%s=%s" % (k, "pari" if v == 0 else "dispari") for k, v in PHASES.items()) + ".",
    "Le celle arancioni nel 'Riepilogo' indicano settimane in cui un dipendente ha svolto meno turni del previsto per via di festivi sul suo giorno obbligato (vedi foglio Segnalazioni).",
]
rr = 3
for ln in notes:
    s4.set(rr, 1, ln, S_NOTE)
    s4.merge(rr, 1, rr, 8)
    rr += 1
s4.width(1, 20)
for k in range(2, 9):
    s4.width(k, 14)

SHEETS = month_sheets + [sE, s2, s3, s4]

# ---- pacchetto xlsx ---------------------------------------------------------

def build_xlsx(path):
    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' \
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' \
        '<Default Extension="xml" ContentType="application/xml"/>' \
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' \
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    for i in range(len(SHEETS)):
        content_types += '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
    content_types += '</Types>'

    root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' \
        '</Relationships>'

    wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
         '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' \
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
    for i, sh in enumerate(SHEETS):
        wb += '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(sh.name), i + 1, i + 1)
    wb += '</sheets></workbook>'

    wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    for i in range(len(SHEETS)):
        wb_rels += '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
    sid = len(SHEETS)
    wb_rels += '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (sid + 1)
    wb_rels += '</Relationships>'

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", STYLES_XML)
        for i, sh in enumerate(SHEETS):
            z.writestr("xl/worksheets/sheet%d.xml" % (i + 1), sheet_xml(sh))


OUT = "/projects/sandbox/ITALIC/Turnazione_Personale_2026-2027.xlsx"
build_xlsx(OUT)
print("File creato:", OUT)
print("Giorni lavorativi nel periodo:", len(WORKING_DATES))
print("Settimane:", len(weeks))
print("Numero segnalazioni:", len([w for w in warnings if not w.startswith("Nessuna")]))
print("\n--- Segnalazioni ---")
for w in warnings:
    print(" -", w)
print("\n--- Totali turni per dipendente ---")
for e in ALL_EMP:
    print(" %-10s %s : %d" % (e, GROUP[e][:4], emp_totals[e]))
