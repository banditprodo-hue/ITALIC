"""
Generatore presentazione PPTX:
"Linee Guida INAIL aggiornate sulla valutazione del rischio
 da Stress Lavoro-Correlato"
Contesto applicativo: Arsenali Militari (es. MARINARSEN Taranto)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Palette istituzionale (blu Difesa / grigio amministrativo)
BLU_DIFESA = RGBColor(0x0B, 0x2E, 0x55)
BLU_CHIARO = RGBColor(0x1F, 0x4E, 0x79)
ORO = RGBColor(0xC9, 0xA2, 0x27)
GRIGIO = RGBColor(0x55, 0x55, 0x55)
GRIGIO_CHIARO = RGBColor(0xEC, 0xEC, 0xEC)
BIANCO = RGBColor(0xFF, 0xFF, 0xFF)
NERO = RGBColor(0x1A, 0x1A, 0x1A)
ROSSO = RGBColor(0xB3, 0x1B, 0x1B)
VERDE = RGBColor(0x2E, 0x7D, 0x32)
GIALLO = RGBColor(0xF9, 0xA8, 0x25)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height

BLANK = prs.slide_layouts[6]


def add_rect(slide, x, y, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
    sh.shadow.inherit = False
    return sh


def add_text(slide, x, y, w, h, text, *, size=14, bold=False, color=NERO,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_bullets(slide, x, y, w, h, items, *, size=16, color=NERO,
                bullet="\u2022", indent_levels=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(4)
        lvl = (indent_levels[i] if indent_levels else 0)
        prefix = ("    " * lvl) + (f"{bullet} " if bullet else "")
        if isinstance(it, tuple):
            head, body = it
            r1 = p.add_run()
            r1.text = prefix + head
            r1.font.name = "Calibri"
            r1.font.size = Pt(size)
            r1.font.bold = True
            r1.font.color.rgb = color
            r2 = p.add_run()
            r2.text = " " + body
            r2.font.name = "Calibri"
            r2.font.size = Pt(size)
            r2.font.color.rgb = color
        else:
            r = p.add_run()
            r.text = prefix + str(it)
            r.font.name = "Calibri"
            r.font.size = Pt(size)
            r.font.color.rgb = color
    return tb


def header_bar(slide, title, subtitle=None, slide_no=None, total=None):
    # Fascia superiore blu Difesa
    add_rect(slide, 0, 0, SW, Inches(0.9), BLU_DIFESA)
    # Linea oro decorativa
    add_rect(slide, 0, Inches(0.9), SW, Inches(0.05), ORO)
    add_text(slide, Inches(0.4), Inches(0.12), Inches(11.5), Inches(0.55),
             title, size=24, bold=True, color=BIANCO,
             anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, Inches(0.4), Inches(0.55), Inches(11.5), Inches(0.32),
                 subtitle, size=12, color=ORO, anchor=MSO_ANCHOR.MIDDLE)
    # Fascia inferiore
    add_rect(slide, 0, SH - Inches(0.35), SW, Inches(0.35), BLU_DIFESA)
    add_text(slide, Inches(0.3), SH - Inches(0.32), Inches(8), Inches(0.3),
             "Linee Guida INAIL - Stress Lavoro-Correlato | Arsenali Militari",
             size=9, color=BIANCO, anchor=MSO_ANCHOR.MIDDLE)
    if slide_no and total:
        add_text(slide, SW - Inches(1.5), SH - Inches(0.32),
                 Inches(1.2), Inches(0.3),
                 f"{slide_no} / {total}", size=9, color=ORO,
                 align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


TOTAL = 22

# =========================================================
# SLIDE 1 - COPERTINA
# =========================================================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, BLU_DIFESA)
# banda decorativa oro
add_rect(s, 0, Inches(2.6), SW, Inches(0.08), ORO)
add_rect(s, 0, Inches(5.2), SW, Inches(0.08), ORO)

add_text(s, Inches(0.6), Inches(0.5), Inches(12), Inches(0.4),
         "MINISTERO DELLA DIFESA - MARINA MILITARE",
         size=14, bold=True, color=ORO)
add_text(s, Inches(0.6), Inches(0.9), Inches(12), Inches(0.4),
         "Servizio di Prevenzione e Protezione - Arsenale Militare",
         size=12, color=BIANCO)

add_text(s, Inches(0.6), Inches(2.9), Inches(12.1), Inches(1.2),
         "VALUTAZIONE DEL RISCHIO\nDA STRESS LAVORO-CORRELATO",
         size=40, bold=True, color=BIANCO)

add_text(s, Inches(0.6), Inches(5.4), Inches(12.1), Inches(0.6),
         "Linee Guida INAIL - Metodologia aggiornata (ed. 2017)",
         size=22, color=ORO, bold=True)
add_text(s, Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.5),
         "Applicazione al contesto degli Arsenali Militari",
         size=16, color=BIANCO)

add_text(s, Inches(0.6), Inches(6.7), Inches(12.1), Inches(0.4),
         "D.Lgs. 81/2008 art. 28 c.1-bis  -  Accordo Europeo 8 ottobre 2004"
         "  -  Circolare Min. Lavoro prot. 23692 del 18/11/2010",
         size=11, color=GRIGIO_CHIARO)

# =========================================================
# SLIDE 2 - INDICE
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Indice della presentazione",
           "Articolazione dei contenuti", 2, TOTAL)

items_l = [
    "1. Definizioni e quadro di riferimento",
    "2. Quadro normativo nazionale ed europeo",
    "3. Indicazioni della Commissione Consultiva (2010)",
    "4. Metodologia INAIL: dal 2011 all'aggiornamento 2017",
    "5. Architettura del processo valutativo",
    "6. Fase propedeutica",
    "7. Valutazione preliminare - eventi sentinella",
    "8. Contenuto del lavoro",
    "9. Contesto del lavoro",
    "10. Lista di controllo - punteggi e fasce di rischio",
]
items_r = [
    "11. Valutazione approfondita",
    "12. Strumento Indicatore (Q-IND) - HSE",
    "13. Pianificazione delle misure correttive",
    "14. Monitoraggio e aggiornamento",
    "15. Ruoli e responsabilita'",
    "16. Specificita' degli Arsenali Militari",
    "17. Fattori di rischio in ambito difensivo",
    "18. Misure di prevenzione - buone prassi",
    "19. Indicatori di efficacia",
    "20. Riferimenti normativi e bibliografici",
]
add_bullets(s, Inches(0.5), Inches(1.2), Inches(6.3), Inches(5.7),
            items_l, size=15, color=NERO, bullet="")
add_bullets(s, Inches(6.9), Inches(1.2), Inches(6.3), Inches(5.7),
            items_r, size=15, color=NERO, bullet="")

# =========================================================
# SLIDE 3 - DEFINIZIONI
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Definizioni e quadro di riferimento",
           "Cosa si intende per stress lavoro-correlato", 3, TOTAL)

add_rect(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(1.5),
         GRIGIO_CHIARO)
add_text(s, Inches(0.7), Inches(1.3), Inches(12.0), Inches(1.4),
         "Definizione (Accordo Europeo 8 ottobre 2004, recepito in Italia "
         "dall'Accordo Interconfederale 9 giugno 2008):\n"
         "\"Lo stress lavoro-correlato e' una condizione che puo' essere "
         "accompagnata da disturbi o disfunzioni di natura fisica, "
         "psicologica o sociale ed e' conseguenza del fatto che taluni "
         "individui non si sentono in grado di corrispondere alle richieste "
         "o alle aspettative riposte in loro\".",
         size=13, color=NERO, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Inches(0.5), Inches(2.95), Inches(12), Inches(0.4),
         "Elementi distintivi:", size=18, bold=True, color=BLU_DIFESA)

add_bullets(s, Inches(0.7), Inches(3.4), Inches(12), Inches(3.5), [
    ("Origine organizzativa:",
     "lo stress lavoro-correlato non coincide con lo stress "
     "individuale; deriva dalla organizzazione e dalle condizioni di "
     "lavoro."),
    ("Ambito di applicazione:",
     "tutti i lavoratori, pubblici e privati, ivi compreso il personale "
     "civile e militare in servizio presso le Forze Armate (cfr. D.M. "
     "284/2000 e D.P.R. 90/2010 - T.U.O.M.)."),
    ("Esclusioni concettuali:",
     "non sono oggetto della valutazione la violenza sul lavoro, le "
     "molestie e lo stress post-traumatico, che richiedono valutazioni "
     "dedicate."),
    ("Riferimento normativo cogente:",
     "art. 28, comma 1-bis, D.Lgs. 81/2008 - obbligo di valutazione "
     "secondo i contenuti dell'Accordo Europeo."),
], size=14, color=NERO)

# =========================================================
# SLIDE 4 - QUADRO NORMATIVO
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Quadro normativo di riferimento",
           "Fonti europee e nazionali", 4, TOTAL)

# tabella
rows = [
    ("Fonte", "Contenuto essenziale"),
    ("Accordo Europeo 8/10/2004",
     "Definizione condivisa di stress lavoro-correlato; obbligo "
     "datoriale di prevenzione; partecipazione dei lavoratori."),
    ("Accordo Interconfederale 9/6/2008",
     "Recepimento in Italia dell'Accordo Europeo."),
    ("D.Lgs. 81/2008 - art. 28 c.1-bis",
     "Obbligo di valutazione del rischio SLC secondo i contenuti "
     "dell'Accordo Europeo, nell'ambito del DVR."),
    ("Commissione Consultiva Permanente - 17/11/2010",
     "Indicazioni metodologiche minime per la valutazione del "
     "rischio SLC."),
    ("Circolare Min. Lavoro prot. 23692 del 18/11/2010",
     "Diffusione delle indicazioni; data di decorrenza dell'obbligo: "
     "31/12/2010."),
    ("Metodologia INAIL ed. 2011",
     "Modello operativo di prima applicazione (Manuale ad uso delle "
     "aziende)."),
    ("Metodologia INAIL ed. 2017",
     "AGGIORNAMENTO: revisione lista di controllo, validazione "
     "Q-IND, integrazione differenze di genere/eta'/provenienza."),
    ("D.M. 284/2000 - D.P.R. 90/2010",
     "Applicazione del T.U. 81/08 al personale militare e civile "
     "della Difesa, con adattamenti."),
]
x0 = Inches(0.5)
y0 = Inches(1.2)
col_w = [Inches(3.6), Inches(9.2)]
row_h = Inches(0.55)
for i, (a, b) in enumerate(rows):
    fill = BLU_DIFESA if i == 0 else (BIANCO if i % 2 else GRIGIO_CHIARO)
    color = BIANCO if i == 0 else NERO
    bold = (i == 0)
    add_rect(s, x0, y0 + row_h * i, col_w[0], row_h, fill,
             line=GRIGIO)
    add_rect(s, x0 + col_w[0], y0 + row_h * i, col_w[1], row_h, fill,
             line=GRIGIO)
    add_text(s, x0 + Inches(0.05), y0 + row_h * i,
             col_w[0] - Inches(0.1), row_h, a,
             size=11, bold=bold, color=color, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x0 + col_w[0] + Inches(0.05), y0 + row_h * i,
             col_w[1] - Inches(0.1), row_h, b,
             size=11, bold=bold, color=color, anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# SLIDE 5 - INDICAZIONI COMMISSIONE CONSULTIVA
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Indicazioni della Commissione Consultiva (17/11/2010)",
           "Percorso metodologico minimo - Lettera Circolare 23692/2010",
           5, TOTAL)

add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5),
         "La valutazione si articola in DUE FASI:",
         size=18, bold=True, color=BLU_DIFESA)

# Box 1 - Valutazione preliminare
add_rect(s, Inches(0.5), Inches(1.85), Inches(6.0), Inches(4.6),
         GRIGIO_CHIARO, line=BLU_DIFESA)
add_rect(s, Inches(0.5), Inches(1.85), Inches(6.0), Inches(0.5),
         BLU_DIFESA)
add_text(s, Inches(0.6), Inches(1.85), Inches(5.8), Inches(0.5),
         "FASE 1 - VALUTAZIONE PRELIMINARE (necessaria)",
         size=13, bold=True, color=BIANCO, anchor=MSO_ANCHOR.MIDDLE)
add_bullets(s, Inches(0.7), Inches(2.5), Inches(5.6), Inches(3.9), [
    "Rilevazione di indicatori OGGETTIVI e VERIFICABILI",
    "Tre famiglie di indicatori:",
    "1) Eventi sentinella",
    "2) Fattori di contenuto del lavoro",
    "3) Fattori di contesto del lavoro",
    "Coinvolgimento di Datore di Lavoro, RSPP, MC, RLS",
    "Sentire i lavoratori e/o un campione rappresentativo",
    "Esito: rischio NON RILEVANTE / MEDIO / ALTO",
], size=12)

# Box 2 - Valutazione approfondita
add_rect(s, Inches(6.8), Inches(1.85), Inches(6.0), Inches(4.6),
         GRIGIO_CHIARO, line=BLU_DIFESA)
add_rect(s, Inches(6.8), Inches(1.85), Inches(6.0), Inches(0.5),
         ROSSO)
add_text(s, Inches(6.9), Inches(1.85), Inches(5.8), Inches(0.5),
         "FASE 2 - VALUTAZIONE APPROFONDITA (eventuale)",
         size=13, bold=True, color=BIANCO, anchor=MSO_ANCHOR.MIDDLE)
add_bullets(s, Inches(7.0), Inches(2.5), Inches(5.6), Inches(3.9), [
    "Attivata SOLO se la fase preliminare evidenzia elementi di "
    "rischio e le misure correttive risultano inefficaci",
    "Rilevazione della percezione soggettiva dei lavoratori",
    "Strumenti: questionari, focus group, interviste "
    "semi-strutturate",
    "Riferimento INAIL: Strumento Indicatore (Q-IND), versione "
    "italiana del HSE Indicator Tool",
    "Nelle aziende con < 5 lavoratori: riunioni con i lavoratori",
    "Coinvolgimento del MC e dello psicologo del lavoro",
], size=12)

add_text(s, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
         "Riferimento: Lettera Circolare prot. 15/SEGR/0023692 del "
         "18/11/2010 - Allegato I.",
         size=10, color=GRIGIO, align=PP_ALIGN.CENTER)

# =========================================================
# SLIDE 6 - METODOLOGIA INAIL: dal 2011 al 2017
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Metodologia INAIL: dal 2011 all'aggiornamento 2017",
           "Cosa cambia con la nuova edizione", 6, TOTAL)

add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.5),
         "Le novita' introdotte dalla edizione 2017 della Metodologia INAIL",
         size=16, bold=True, color=BLU_DIFESA)

# tabella confronto
rows = [
    ("Aspetto", "Edizione 2011", "Edizione 2017 (aggiornata)"),
    ("Lista di controllo",
     "35 indicatori - aree A/B/C",
     "Revisione/integrazione indicatori; maggiore aderenza al contesto "
     "organizzativo reale"),
    ("Differenze di genere",
     "Non esplicitamente trattate",
     "Integrazione esplicita di genere, eta', provenienza, tipologia "
     "contrattuale (D.Lgs. 81/08 art. 28 c.1)"),
    ("Strumento Indicatore (Q-IND)",
     "Adattamento del HSE Indicator Tool",
     "Versione validata INAIL su campione italiano; piattaforma online "
     "INAIL per somministrazione e analisi"),
    ("Coinvolgimento lavoratori",
     "Genericamente previsto",
     "Modalita' definite (focus group strutturati, partecipazione "
     "preposti)"),
    ("Gestione del rischio",
     "Misure correttive a valle",
     "Approccio gestionale ciclico (PDCA) - integrazione con SGSL"),
    ("Strumenti operativi",
     "Manuale + foglio di calcolo",
     "Piattaforma INAIL on-line + manuale aggiornato"),
]
x0 = Inches(0.5)
y0 = Inches(1.7)
col_w = [Inches(2.6), Inches(4.6), Inches(5.6)]
row_h = Inches(0.7)
for i, row in enumerate(rows):
    fill = BLU_DIFESA if i == 0 else (BIANCO if i % 2 else GRIGIO_CHIARO)
    color = BIANCO if i == 0 else NERO
    x = x0
    for j, cell in enumerate(row):
        add_rect(s, x, y0 + row_h * i, col_w[j], row_h, fill, line=GRIGIO)
        add_text(s, x + Inches(0.05), y0 + row_h * i,
                 col_w[j] - Inches(0.1), row_h, cell,
                 size=10, bold=(i == 0), color=color,
                 anchor=MSO_ANCHOR.MIDDLE)
        x += col_w[j]

# =========================================================
# SLIDE 7 - ARCHITETTURA DEL PROCESSO
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Architettura del processo valutativo",
           "Le quattro fasi della Metodologia INAIL", 7, TOTAL)

phases = [
    ("FASE\nPROPEDEUTICA",
     "Costituzione Gruppo di Gestione della Valutazione (GGV);\n"
     "definizione dei gruppi omogenei;\nattivita' di sensibilizzazione.",
     BLU_CHIARO),
    ("VALUTAZIONE\nPRELIMINARE",
     "Lista di controllo:\nEventi sentinella + Contenuto + Contesto.\n"
     "Esito: NON RILEVANTE / MEDIO / ALTO.",
     ORO),
    ("PIANO DI\nAZIONE",
     "Definizione e attuazione delle misure correttive\n"
     "tecniche, organizzative, procedurali, formative.",
     VERDE),
    ("VALUTAZIONE\nAPPROFONDITA",
     "(Eventuale) Q-IND, focus group, interviste.\n"
     "Coinvolgimento Medico Competente.",
     ROSSO),
]

x = Inches(0.4)
y = Inches(1.5)
w = Inches(3.0)
h = Inches(4.2)
gap = Inches(0.15)
for title, body, col in phases:
    # blocco
    add_rect(s, x, y, w, h, GRIGIO_CHIARO, line=GRIGIO)
    add_rect(s, x, y, w, Inches(1.1), col)
    add_text(s, x + Inches(0.1), y + Inches(0.05),
             w - Inches(0.2), Inches(1.0), title,
             size=15, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.15), y + Inches(1.25),
             w - Inches(0.3), h - Inches(1.4), body,
             size=12, color=NERO,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP)
    x += w + gap

# Freccia ciclica
add_text(s, Inches(0.5), Inches(5.9), Inches(12.3), Inches(0.5),
         "MONITORAGGIO E RIESAME - Riesame periodico ed in occasione "
         "di mutamenti organizzativi significativi",
         size=14, bold=True, color=BLU_DIFESA, align=PP_ALIGN.CENTER)
add_text(s, Inches(0.5), Inches(6.4), Inches(12.3), Inches(0.4),
         "Approccio ciclico (PDCA) coerente con il Sistema di Gestione "
         "della Sicurezza sul Lavoro (UNI ISO 45001 - art. 30 D.Lgs. 81/08)",
         size=11, color=GRIGIO, align=PP_ALIGN.CENTER)

# =========================================================
# SLIDE 8 - FASE PROPEDEUTICA
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Fase propedeutica",
           "Pianificazione, partecipazione, gruppi omogenei", 8, TOTAL)

add_text(s, Inches(0.5), Inches(1.15), Inches(6.2), Inches(0.4),
         "1) Costituzione del GGV", size=18, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(0.6), Inches(1.6), Inches(6.0), Inches(2.2), [
    "Datore di Lavoro o suo delegato (presiede)",
    "RSPP / ASPP",
    "Medico Competente",
    "RLS / RLST",
    "Eventuale consulente esperto (psicologo del lavoro)",
    "Per gli arsenali: Direttore di Stabilimento, Capi Reparto",
], size=12)

add_text(s, Inches(0.5), Inches(3.85), Inches(6.2), Inches(0.4),
         "2) Definizione dei gruppi omogenei",
         size=18, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(0.6), Inches(4.3), Inches(6.0), Inches(2.7), [
    "Per mansione/lavorazione (es. saldatori carpenteria navale)",
    "Per partizione organizzativa (es. Officina Motori)",
    "Per turnazione (es. squadre H24 banchina)",
    "Tenere conto di genere, eta', anzianita' di servizio",
], size=12)

add_text(s, Inches(7.0), Inches(1.15), Inches(6.0), Inches(0.4),
         "3) Strategia di comunicazione",
         size=18, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(7.1), Inches(1.6), Inches(5.8), Inches(2.2), [
    "Informativa preventiva al personale",
    "Coinvolgimento RLS e Organizzazioni Sindacali",
    "Tutela della riservatezza dei dati (Reg. UE 679/2016)",
    "Pubblicazione esiti in forma aggregata",
], size=12)

add_text(s, Inches(7.0), Inches(3.85), Inches(6.0), Inches(0.4),
         "4) Cronoprogramma",
         size=18, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(7.1), Inches(4.3), Inches(5.8), Inches(2.7), [
    "Avvio entro 30 gg dalla nomina del GGV",
    "Conclusione fase preliminare entro 90 gg",
    "Revisione almeno biennale o al mutare delle condizioni",
    "Aggiornamento immediato in caso di evento sentinella critico",
], size=12)

# =========================================================
# SLIDE 9 - EVENTI SENTINELLA
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Valutazione preliminare - Area A: Eventi sentinella",
           "Indicatori oggettivi - dati consuntivi triennali", 9, TOTAL)

add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.5),
         "Si confronta l'andamento del triennio con i due trienni "
         "precedenti (variazione % per gruppo omogeneo)",
         size=13, color=GRIGIO)

ev = [
    "Indici infortunistici",
    "Assenze per malattia",
    "Assenze dal lavoro",
    "Ferie non godute",
    "Rotazione del personale",
    "Turnover",
    "Procedimenti / sanzioni disciplinari",
    "Richieste di visita medica straordinaria al MC",
    "Segnalazioni del MC su condizioni di stress",
    "Istanze giudiziarie per molestie/mobbing",
]
x = Inches(0.5)
y = Inches(1.85)
w = Inches(3.05)
h = Inches(0.85)
for i, v in enumerate(ev):
    col = i % 4
    row = i // 4
    add_rect(s, x + col * (w + Inches(0.05)),
             y + row * (h + Inches(0.1)), w, h,
             GRIGIO_CHIARO, line=BLU_DIFESA)
    add_text(s, x + col * (w + Inches(0.05)) + Inches(0.1),
             y + row * (h + Inches(0.1)),
             w - Inches(0.2), h, v,
             size=12, bold=True, color=NERO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

add_rect(s, Inches(0.5), Inches(5.3), Inches(12.3), Inches(1.7),
         GRIGIO_CHIARO, line=ORO)
add_text(s, Inches(0.7), Inches(5.4), Inches(12), Inches(0.4),
         "Modalita' di valutazione (Metodologia INAIL):",
         size=14, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(0.7), Inches(5.75), Inches(12), Inches(1.3), [
    "Per ciascun indicatore: 0 punti se diminuisce/stabile, 1 punto se "
    "aumenta entro soglia, 4 punti se aumento > 10%",
    "Somma dei punteggi -> punteggio dell'Area A",
    "Eventi sentinella attribuibili al gruppo omogeneo, NON al singolo",
], size=12)

# =========================================================
# SLIDE 10 - CONTENUTO DEL LAVORO
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Valutazione preliminare - Area B: Contenuto del lavoro",
           "Fattori legati alle caratteristiche tecniche e operative",
           10, TOTAL)

categories = [
    ("Ambiente di lavoro\ned attrezzature",
     "Esposizione a rumore, vibrazioni, microclima, illuminazione, "
     "agenti chimici/cancerogeni, spazi confinati, lavori in quota; "
     "adeguatezza dei DPI; manutenzione attrezzature."),
    ("Pianificazione\ndei compiti",
     "Adeguatezza dei compiti rispetto alle competenze; carenze nella "
     "definizione di mansioni e procedure; lavoro frammentato o "
     "ripetitivo."),
    ("Carico - Ritmo\ndi lavoro",
     "Sovraccarico/sottocarico; pressione temporale; scadenze rigide "
     "(es. consegna unita' navale); ritmi imposti da macchine."),
    ("Orario di lavoro",
     "Turnazione; lavoro notturno; reperibilita'; straordinari "
     "frequenti; mancanza di flessibilita'; lavoro in giorni festivi."),
]
x = Inches(0.5)
y = Inches(1.2)
w = Inches(6.2)
h = Inches(2.7)
for i, (t, b) in enumerate(categories):
    col = i % 2
    row = i // 2
    xx = x + col * (w + Inches(0.1))
    yy = y + row * (h + Inches(0.1))
    add_rect(s, xx, yy, w, h, GRIGIO_CHIARO, line=BLU_DIFESA)
    add_rect(s, xx, yy, w, Inches(0.7), BLU_CHIARO)
    add_text(s, xx + Inches(0.15), yy, w - Inches(0.3), Inches(0.7),
             t, size=14, bold=True, color=BIANCO,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, xx + Inches(0.15), yy + Inches(0.75),
             w - Inches(0.3), h - Inches(0.85), b,
             size=12, color=NERO)

add_text(s, Inches(0.5), Inches(6.85), Inches(12.3), Inches(0.3),
         "Riferimento: Lista di controllo INAIL ed. 2017 - "
         "indicatori organizzativi verificabili.",
         size=10, color=GRIGIO, align=PP_ALIGN.CENTER)

# =========================================================
# SLIDE 11 - CONTESTO DEL LAVORO
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Valutazione preliminare - Area C: Contesto del lavoro",
           "Fattori organizzativi, relazionali, gestionali", 11, TOTAL)

categories = [
    ("Funzione e cultura\norganizzativa",
     "Chiarezza degli obiettivi; sistema di comunicazione interna; "
     "coerenza tra dichiarato e agito; codice etico/disciplinare."),
    ("Ruolo nell'ambito\ndell'organizzazione",
     "Conflitto di ruolo; ambiguita' di ruolo; sovrapposizione "
     "gerarchica civile/militare tipica degli arsenali."),
    ("Evoluzione della\ncarriera",
     "Trasparenza percorsi di sviluppo; sistema premiante; mobilita' "
     "interna; precarieta' del rapporto di lavoro."),
    ("Autonomia decisionale\n- Controllo del lavoro",
     "Margini di autonomia operativa; partecipazione alle decisioni; "
     "rigidita' delle procedure."),
    ("Rapporti interpersonali\nsul lavoro",
     "Qualita' delle relazioni con colleghi e superiori; presenza di "
     "conflitti; supporto sociale."),
    ("Interfaccia\ncasa-lavoro",
     "Conciliazione tempi di vita/lavoro; pendolarismo; impatto "
     "delle turnazioni sulla vita familiare."),
]
x = Inches(0.5)
y = Inches(1.15)
w = Inches(4.1)
h = Inches(1.85)
for i, (t, b) in enumerate(categories):
    col = i % 3
    row = i // 3
    xx = x + col * (w + Inches(0.07))
    yy = y + row * (h + Inches(0.1))
    add_rect(s, xx, yy, w, h, GRIGIO_CHIARO, line=BLU_DIFESA)
    add_rect(s, xx, yy, w, Inches(0.6), BLU_CHIARO)
    add_text(s, xx + Inches(0.1), yy, w - Inches(0.2), Inches(0.6),
             t, size=11, bold=True, color=BIANCO,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
    add_text(s, xx + Inches(0.1), yy + Inches(0.65),
             w - Inches(0.2), h - Inches(0.75), b,
             size=10, color=NERO)

add_rect(s, Inches(0.5), Inches(5.1), Inches(12.3), Inches(1.85),
         GRIGIO_CHIARO, line=ORO)
add_text(s, Inches(0.7), Inches(5.2), Inches(12), Inches(0.4),
         "Specificita' del contesto militare:",
         size=14, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(0.7), Inches(5.55), Inches(12), Inches(1.4), [
    "Compresenza di personale militare (ordinamento gerarchico) e "
    "civile (CCNL Funzioni Centrali) - potenziali tensioni di "
    "ruolo e di trattamento",
    "Vincoli operativi connessi alle esigenze di prontezza operativa "
    "e segretezza",
    "Mobilita' del personale militare per esigenze di servizio",
    "Doppia catena di comando: gerarchica militare e tecnico-funzionale "
    "in ambito SPP (art. 31 D.Lgs. 81/08)",
], size=11)

# =========================================================
# SLIDE 12 - LISTA DI CONTROLLO E PUNTEGGI
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Lista di controllo - Punteggi e fasce di rischio",
           "Esito della valutazione preliminare", 12, TOTAL)

add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5),
         "Aggregazione dei punteggi nelle tre Aree (A, B, C) e "
         "attribuzione del livello di rischio",
         size=14, color=NERO)

# Tabella livelli di rischio
rows = [
    ("Fascia", "Punteggio totale (% sul max)", "Conseguenze operative"),
    ("RISCHIO\nNON RILEVANTE",
     "0% - 25%",
     "Monitoraggio biennale; non sono richieste misure correttive."),
    ("RISCHIO\nMEDIO",
     "25% - 50%",
     "Adozione di misure correttive (organizzative, formative, "
     "comunicative); verifica di efficacia entro 12 mesi."),
    ("RISCHIO\nALTO",
     "> 50%",
     "Misure correttive immediate + Valutazione approfondita "
     "(Fase 2) entro tempi congrui."),
]
colors = [BLU_DIFESA, VERDE, GIALLO, ROSSO]
x0 = Inches(0.5)
y0 = Inches(1.85)
col_w = [Inches(2.6), Inches(3.0), Inches(7.2)]
row_h = Inches(1.0)
for i, row in enumerate(rows):
    color_bg = colors[i] if i < 4 else BIANCO
    color_txt = BIANCO if i == 0 else (BIANCO if i in (1, 3) else NERO)
    x = x0
    for j, cell in enumerate(row):
        fill = color_bg if (i == 0 or j == 0) else \
            (BIANCO if i % 2 else GRIGIO_CHIARO)
        txt_color = (BIANCO if (i == 0 or j == 0) and i in (0, 1, 3)
                     else (NERO if not (i == 0 and j > 0) else BIANCO))
        if i == 0:
            fill = BLU_DIFESA
            txt_color = BIANCO
        elif j == 0:
            fill = colors[i]
            txt_color = BIANCO if i in (1, 3) else NERO
        else:
            fill = BIANCO if i % 2 else GRIGIO_CHIARO
            txt_color = NERO
        add_rect(s, x, y0 + row_h * i, col_w[j], row_h,
                 fill, line=GRIGIO)
        add_text(s, x + Inches(0.1), y0 + row_h * i,
                 col_w[j] - Inches(0.2), row_h, cell,
                 size=12, bold=(i == 0 or j == 0),
                 color=txt_color, anchor=MSO_ANCHOR.MIDDLE,
                 align=PP_ALIGN.LEFT if j == 2 else PP_ALIGN.CENTER)
        x += col_w[j]

add_text(s, Inches(0.5), Inches(6.0), Inches(12.3), Inches(1.0),
         "Le soglie di calcolo derivano dall'algoritmo della Metodologia "
         "INAIL (Manuale ed. 2017): la somma dei punteggi delle Aree A, "
         "B e C viene rapportata al massimo teorico, con pesi "
         "differenziati per Area. L'esito alimenta direttamente la "
         "sezione \"Stress lavoro-correlato\" del DVR.",
         size=11, color=GRIGIO, anchor=MSO_ANCHOR.TOP)



# =========================================================
# SLIDE 13 - VALUTAZIONE APPROFONDITA
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Valutazione approfondita (Fase 2)",
           "Quando, come, con quali strumenti", 13, TOTAL)

add_rect(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(1.05),
         GRIGIO_CHIARO, line=ROSSO)
add_text(s, Inches(0.7), Inches(1.2), Inches(12), Inches(0.45),
         "QUANDO si attiva", size=15, bold=True, color=ROSSO)
add_text(s, Inches(0.7), Inches(1.6), Inches(12), Inches(0.6),
         "Esclusivamente quando la fase preliminare evidenzia rischio "
         "ALTO oppure quando le misure correttive adottate non si "
         "rivelano efficaci entro i tempi pianificati.",
         size=12, color=NERO)

add_text(s, Inches(0.5), Inches(2.5), Inches(6.0), Inches(0.4),
         "Strumenti privilegiati", size=16, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(0.6), Inches(2.95), Inches(5.9), Inches(3.8), [
    ("Q-IND (Strumento Indicatore INAIL):",
     "questionario validato derivato dal HSE Indicator Tool, "
     "35 item su 7 dimensioni (Domanda, Controllo, Supporto dei "
     "superiori, Supporto dei colleghi, Relazioni, Ruolo, "
     "Cambiamento)."),
    ("Focus group:",
     "gruppi di 6-10 lavoratori per gruppo omogeneo, condotti da "
     "facilitatore esperto."),
    ("Interviste semi-strutturate:",
     "per gruppi ridotti o per posizioni dirigenziali."),
], size=12)

add_text(s, Inches(6.8), Inches(2.5), Inches(6.0), Inches(0.4),
         "Requisiti procedurali", size=16, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(6.9), Inches(2.95), Inches(5.9), Inches(3.8), [
    "Tutela dell'anonimato dei rispondenti (Reg. UE 679/2016 - GDPR)",
    "Soglia minima di adesione per significativita' statistica "
    "(consigliato >= 60% del gruppo omogeneo)",
    "Coinvolgimento attivo del Medico Competente (artt. 25 e 41 "
    "D.Lgs. 81/08)",
    "Restituzione degli esiti aggregati a tutti i lavoratori",
    "Per imprese < 5 lavoratori: riunioni dirette in luogo dei "
    "questionari (Indicazioni CCP 2010)",
], size=12)

# =========================================================
# SLIDE 14 - STRUMENTO INDICATORE Q-IND
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Strumento Indicatore (Q-IND)",
           "Le 7 dimensioni HSE - traduzione e validazione INAIL",
           14, TOTAL)

dims = [
    ("DOMANDA",
     "Carico di lavoro, ritmi, ore di lavoro, ambiente fisico"),
    ("CONTROLLO",
     "Margine di autonomia decisionale del lavoratore sui propri "
     "compiti e tempi"),
    ("SUPPORTO DEI\nSUPERIORI",
     "Sostegno informativo, organizzativo, emotivo da parte dei "
     "responsabili"),
    ("SUPPORTO DEI\nCOLLEGHI",
     "Aiuto e collaborazione tra colleghi nel medesimo gruppo "
     "omogeneo"),
    ("RELAZIONI",
     "Qualita' delle relazioni interpersonali; presenza di conflitti "
     "o comportamenti inaccettabili"),
    ("RUOLO",
     "Chiarezza del ruolo, compatibilita' fra i compiti, conflitti di "
     "ruolo"),
    ("CAMBIAMENTO",
     "Modalita' di gestione e comunicazione dei cambiamenti "
     "organizzativi"),
]

x = Inches(0.5)
y = Inches(1.2)
w = Inches(6.1)
h = Inches(0.7)
for i, (t, d) in enumerate(dims):
    col = i % 2
    row = i // 2
    if col == 0 and row == 3 and i == 6:
        # Ultima dimensione centrata
        xx = (SW - w) / 2
    else:
        xx = x + col * (w + Inches(0.1))
    yy = y + row * (h + Inches(0.1))
    add_rect(s, xx, yy, Inches(2.0), h, BLU_DIFESA)
    add_rect(s, xx + Inches(2.0), yy, w - Inches(2.0), h,
             GRIGIO_CHIARO, line=BLU_DIFESA)
    add_text(s, xx + Inches(0.05), yy, Inches(1.95), h, t,
             size=11, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, xx + Inches(2.05), yy, w - Inches(2.1), h, d,
             size=11, color=NERO, anchor=MSO_ANCHOR.MIDDLE)

add_rect(s, Inches(0.5), Inches(5.2), Inches(12.3), Inches(1.7),
         GRIGIO_CHIARO, line=ORO)
add_text(s, Inches(0.7), Inches(5.3), Inches(12), Inches(0.4),
         "Esito del Q-IND - rappresentazione \"a semaforo\" per "
         "dimensione:",
         size=14, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(0.7), Inches(5.7), Inches(12), Inches(1.2), [
    "VERDE: situazione adeguata - mantenere",
    "GIALLO: necessitano interventi di miglioramento",
    "ARANCIONE/ROSSO: necessita' di azioni correttive prioritarie",
], size=12)

# =========================================================
# SLIDE 15 - PIANO DELLE MISURE CORRETTIVE
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Pianificazione delle misure correttive",
           "Dalla valutazione all'azione di prevenzione", 15, TOTAL)

add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4),
         "Tassonomia delle misure (gerarchia dei controlli - "
         "art. 15 D.Lgs. 81/08):",
         size=15, bold=True, color=BLU_DIFESA)

cats = [
    ("MISURE\nORGANIZZATIVE",
     "Riprogettazione carichi e turni;\nrotazione delle mansioni "
     "particolarmente gravose;\nrevisione delle procedure;\n"
     "ottimizzazione delle pause."),
    ("MISURE\nCOMUNICATIVE",
     "Riunioni periodiche di reparto;\nsistema di feedback;\n"
     "trasparenza degli obiettivi;\ncanali per segnalazioni "
     "(whistleblowing - L. 179/2017)."),
    ("MISURE\nFORMATIVE",
     "Formazione preposti su gestione del personale;\nformazione "
     "specifica ex art. 37 D.Lgs. 81/08;\nformazione lavoratori "
     "(Accordo S-R 21/12/2011)."),
    ("MISURE\nTECNICHE",
     "Bonifica ambientale (rumore, microclima);\nergonomia delle "
     "postazioni;\nstrumenti tecnologici di supporto;\nDPI "
     "adeguati (Titolo III)."),
    ("MISURE\nGESTIONALI",
     "Sistema premiante equo;\nvalutazione della prestazione;\n"
     "percorsi di sviluppo;\nsupporto psicologico (sportello di "
     "ascolto)."),
    ("MISURE\nSANITARIE",
     "Sorveglianza sanitaria mirata (art. 41 c.2 lett. a);\n"
     "visite a richiesta del lavoratore (art. 41 c.2 lett. c);\n"
     "protocollo MC integrato con SLC."),
]
x = Inches(0.5)
y = Inches(1.7)
w = Inches(4.1)
h = Inches(2.45)
colors_cat = [BLU_CHIARO, BLU_CHIARO, BLU_CHIARO,
              BLU_CHIARO, BLU_CHIARO, BLU_CHIARO]
for i, (t, d) in enumerate(cats):
    col = i % 3
    row = i // 3
    xx = x + col * (w + Inches(0.07))
    yy = y + row * (h + Inches(0.1))
    add_rect(s, xx, yy, w, h, GRIGIO_CHIARO, line=BLU_DIFESA)
    add_rect(s, xx, yy, w, Inches(0.7), colors_cat[i])
    add_text(s, xx + Inches(0.1), yy, w - Inches(0.2), Inches(0.7),
             t, size=12, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, xx + Inches(0.15), yy + Inches(0.8),
             w - Inches(0.3), h - Inches(0.9), d,
             size=10, color=NERO)

add_text(s, Inches(0.5), Inches(6.7), Inches(12.3), Inches(0.3),
         "Per ciascuna misura: responsabile, tempistica, "
         "risorse, indicatore di efficacia. Tracciabilita' nel DVR.",
         size=10, color=GRIGIO, align=PP_ALIGN.CENTER)

# =========================================================
# SLIDE 16 - MONITORAGGIO E AGGIORNAMENTO
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Monitoraggio e aggiornamento",
           "Ciclo di Deming applicato alla valutazione SLC", 16, TOTAL)

# Disegno PDCA
cx = Inches(3.2)
cy = Inches(3.8)
size_box = Inches(2.4)
gap_box = Inches(0.3)
positions = [
    (cx, cy - size_box - gap_box, "PLAN",
     "Pianificare la valutazione, definire metodi, GGV, gruppi "
     "omogenei", BLU_CHIARO),
    (cx + size_box + gap_box, cy, "DO",
     "Eseguire la rilevazione (lista di controllo + eventuale Q-IND), "
     "raccogliere dati", VERDE),
    (cx, cy + size_box + gap_box, "CHECK",
     "Verificare risultati, efficacia delle misure, indicatori di "
     "monitoraggio", ORO),
    (cx - size_box - gap_box, cy, "ACT",
     "Riesaminare, correggere, aggiornare il DVR e ripianificare",
     ROSSO),
]
for x, y, t, d, col in positions:
    add_rect(s, x, y, size_box, size_box, col)
    add_text(s, x + Inches(0.1), y + Inches(0.1),
             size_box - Inches(0.2), Inches(0.6), t,
             size=22, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.1), y + Inches(0.75),
             size_box - Inches(0.2), size_box - Inches(0.85), d,
             size=10, color=BIANCO,
             align=PP_ALIGN.CENTER)

add_rect(s, Inches(7.5), Inches(1.3), Inches(5.3), Inches(5.6),
         GRIGIO_CHIARO, line=BLU_DIFESA)
add_rect(s, Inches(7.5), Inches(1.3), Inches(5.3), Inches(0.5),
         BLU_DIFESA)
add_text(s, Inches(7.6), Inches(1.3), Inches(5.1), Inches(0.5),
         "Eventi che impongono il riesame", size=13, bold=True,
         color=BIANCO, anchor=MSO_ANCHOR.MIDDLE)
add_bullets(s, Inches(7.7), Inches(1.95), Inches(5.0), Inches(4.9), [
    "Modifica significativa dell'organizzazione (es. "
    "riorganizzazione di reparto, accorpamento officine)",
    "Introduzione di nuove tecnologie/lavorazioni",
    "Variazione consistente degli organici",
    "Eventi infortunistici/sanitari ripetuti",
    "Segnalazioni RLS o lavoratori",
    "Esito non conforme della sorveglianza sanitaria",
    "Cadenza minima: revisione almeno biennale (best practice INAIL)",
    "Aggiornamento del DVR ai sensi dell'art. 29 c.3 D.Lgs. 81/08",
], size=11)

# =========================================================
# SLIDE 17 - RUOLI E RESPONSABILITA'
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Ruoli e responsabilita'",
           "Il sistema di prevenzione applicato all'arsenale militare",
           17, TOTAL)

ruoli = [
    ("DATORE\nDI LAVORO",
     "Direttore dell'Arsenale (o soggetto delegato ex art. 16). "
     "Obbligo di valutazione (art. 17), redazione DVR, attuazione "
     "misure, riesame."),
    ("DIRIGENTI E\nPREPOSTI",
     "Capi Reparto/Officina, Sottufficiali responsabili. Vigilanza "
     "ex artt. 18 e 19; obbligo di attuazione delle misure di "
     "prevenzione."),
    ("RSPP / ASPP",
     "Coordinamento tecnico-metodologico della valutazione SLC; "
     "elaborazione lista di controllo; tenuta documentale."),
    ("MEDICO\nCOMPETENTE",
     "Collabora alla VdR (art. 25 c.1 lett. a); sorveglianza "
     "sanitaria mirata; segnalazioni di casi indicativi di "
     "stress."),
    ("RLS",
     "Consultato preventivamente (art. 50); accesso al DVR; "
     "partecipazione al GGV; ricezione esiti aggregati."),
    ("LAVORATORI",
     "Cooperazione (art. 20); partecipazione a Q-IND e focus group; "
     "segnalazione di criticita' al preposto/RLS."),
]
x = Inches(0.5)
y = Inches(1.2)
w = Inches(4.1)
h = Inches(1.8)
for i, (t, d) in enumerate(ruoli):
    col = i % 3
    row = i // 3
    xx = x + col * (w + Inches(0.07))
    yy = y + row * (h + Inches(0.15))
    add_rect(s, xx, yy, w, h, GRIGIO_CHIARO, line=BLU_DIFESA)
    add_rect(s, xx, yy, Inches(1.2), h, BLU_DIFESA)
    add_text(s, xx + Inches(0.05), yy, Inches(1.1), h, t,
             size=11, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, xx + Inches(1.3), yy + Inches(0.1),
             w - Inches(1.4), h - Inches(0.2), d,
             size=10, color=NERO, anchor=MSO_ANCHOR.MIDDLE)

add_text(s, Inches(0.5), Inches(6.6), Inches(12.3), Inches(0.4),
         "In ambito difensivo trovano applicazione le specificita' del "
         "D.M. 14/06/2000 n. 284 e del D.P.R. 90/2010 (T.U.O.M.).",
         size=11, color=GRIGIO, align=PP_ALIGN.CENTER)

# =========================================================
# SLIDE 18 - SPECIFICITA' ARSENALI MILITARI
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Specificita' degli Arsenali Militari",
           "Fattori organizzativi caratteristici da considerare",
           18, TOTAL)

specs = [
    "Compresenza di personale militare (ordinamento gerarchico) e "
    "civile (CCNL Funzioni Centrali)",
    "Doppia catena di comando: gerarchica militare e "
    "tecnico-professionale (RSPP, MC)",
    "Lavorazioni di cantieristica navale militare ad alta "
    "complessita' (saldatura, carpenteria, allestimento, "
    "manutenzione propulsivi)",
    "Lavori in spazi confinati (D.P.R. 177/2011) a bordo unita' "
    "in costruzione/manutenzione",
    "Lavori in quota su ponteggi di carena, alberature, scali",
    "Manipolazione di materiale esplosivo (D.M. 272/1999) e "
    "munizionamento",
    "Esposizione ad agenti chimici (vernici, sgrassanti), "
    "cancerogeni (residui amianto in unita' datate)",
    "Lavorazioni in regime di prontezza operativa - tempi "
    "compressi (es. allestimento pre-deployment)",
    "Riservatezza/segretezza connessa a programmi militari "
    "(potenziale isolamento informativo)",
    "Mobilita' del personale militare per esigenze di servizio",
]
add_bullets(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.7),
            specs, size=14, color=NERO)

# =========================================================
# SLIDE 19 - MISURE DI PREVENZIONE - BUONE PRASSI
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Misure di prevenzione - Buone prassi",
           "Esempi di interventi adottabili nel contesto MARINARSEN",
           19, TOTAL)

esempi = [
    ("Officina Motori / Carpenteria",
     "Rotazione su lavorazioni a maggiore impatto (rumore, MMC); "
     "pause programmate; rotazione su mansioni a complessita' "
     "elevata; revisione layout per riduzione spostamenti."),
    ("Bacino di carenaggio",
     "Pianificazione delle lavorazioni con margini per imprevisti "
     "meteo-marini; coordinamento con ditte appaltatrici via PSC "
     "(art. 100); briefing di squadra a inizio/fine turno."),
    ("Deposito munizionamento (D.M. 272/99)",
     "Procedure operative chiare; double-check; turnazione equa; "
     "supporto psicologico per personale impiegato in operazioni "
     "ad alta criticita'."),
    ("Squadre di banchina H24",
     "Equilibrio nei turni di reperibilita'; riposi compensativi; "
     "monitoraggio ore di straordinario per gruppo omogeneo."),
    ("Tutte le aree",
     "Sportello di ascolto psicologico (anonimato garantito); "
     "formazione preposti su gestione del personale; canale "
     "segnalazioni ex L. 179/2017; riconoscimenti non monetari."),
]
x = Inches(0.5)
y = Inches(1.2)
w = Inches(12.3)
h = Inches(1.05)
for i, (t, d) in enumerate(esempi):
    yy = y + i * (h + Inches(0.05))
    add_rect(s, x, yy, w, h, GRIGIO_CHIARO, line=BLU_DIFESA)
    add_rect(s, x, yy, Inches(3.0), h, BLU_CHIARO)
    add_text(s, x + Inches(0.1), yy, Inches(2.9), h, t,
             size=12, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(3.15), yy + Inches(0.05),
             Inches(9.05), h - Inches(0.1), d,
             size=11, color=NERO, anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# SLIDE 20 - INDICATORI DI EFFICACIA
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Indicatori di efficacia",
           "Come misurare l'effetto delle misure adottate", 20, TOTAL)

ind = [
    ("INDICATORI\nOGGETTIVI (lag)",
     "Riduzione assenze per malattia (%);\n"
     "Riduzione tasso infortunistico SLC-correlato;\n"
     "Riduzione turnover/richieste di trasferimento;\n"
     "Riduzione richieste di visita straordinaria al MC;\n"
     "Riduzione procedimenti disciplinari."),
    ("INDICATORI\nDI PROCESSO (lead)",
     "% lavoratori formati su SLC (target >= 95%);\n"
     "% misure correttive attuate nei tempi pianificati;\n"
     "Numero focus group/riunioni effettuati;\n"
     "Tempo medio di risposta a segnalazioni RLS;\n"
     "Copertura sorveglianza sanitaria mirata."),
    ("INDICATORI\nDI PERCEZIONE",
     "Punteggi Q-IND per dimensione;\n"
     "Soddisfazione lavoratori (survey periodica);\n"
     "Indicatore di clima organizzativo;\n"
     "Tasso di adesione ai questionari (proxy di "
     "engagement)."),
]
x = Inches(0.5)
y = Inches(1.2)
w = Inches(4.15)
h = Inches(5.4)
for i, (t, d) in enumerate(ind):
    xx = x + i * (w + Inches(0.07))
    add_rect(s, xx, y, w, h, GRIGIO_CHIARO, line=BLU_DIFESA)
    add_rect(s, xx, y, w, Inches(0.85), BLU_DIFESA)
    add_text(s, xx + Inches(0.1), y, w - Inches(0.2), Inches(0.85),
             t, size=13, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, xx + Inches(0.15), y + Inches(0.95),
             w - Inches(0.3), h - Inches(1.05), d,
             size=12, color=NERO)

add_text(s, Inches(0.5), Inches(6.75), Inches(12.3), Inches(0.3),
         "Ogni indicatore deve avere: baseline, target, frequenza di "
         "rilevazione, responsabile.",
         size=10, color=GRIGIO, align=PP_ALIGN.CENTER)

# =========================================================
# SLIDE 21 - RIFERIMENTI NORMATIVI E BIBLIOGRAFICI
# =========================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Riferimenti normativi e bibliografici",
           "Fonti utilizzate", 21, TOTAL)

add_text(s, Inches(0.5), Inches(1.15), Inches(6.0), Inches(0.4),
         "Normativa", size=18, bold=True, color=BLU_DIFESA)
add_bullets(s, Inches(0.6), Inches(1.55), Inches(6.0), Inches(5.2), [
    "D.Lgs. 9 aprile 2008, n. 81 (artt. 17, 18, 25, 28, 29, 31, "
    "37, 41, 50)",
    "Accordo Quadro Europeo sullo stress lavoro-correlato dell'8 "
    "ottobre 2004",
    "Accordo Interconfederale di recepimento del 9 giugno 2008",
    "Commissione Consultiva Permanente - documento approvato "
    "il 17 novembre 2010",
    "Lettera Circolare del Ministero del Lavoro prot. 23692 "
    "del 18 novembre 2010",
    "D.M. 14 giugno 2000 n. 284 - applicazione T.U. al personale "
    "militare",
    "D.P.R. 15 marzo 2010 n. 90 - T.U. Ordinamento Militare",
    "Accordo Stato-Regioni 21 dicembre 2011 (art. 37)",
    "Accordo Stato-Regioni 7 luglio 2016 (revisione formazione)",
    "Reg. UE 679/2016 (GDPR) - tutela dati nei questionari",
], size=11)

add_text(s, Inches(7.0), Inches(1.15), Inches(6.0), Inches(0.4),
         "Documentazione tecnica INAIL", size=18, bold=True,
         color=BLU_DIFESA)
add_bullets(s, Inches(7.1), Inches(1.55), Inches(6.0), Inches(5.2), [
    "INAIL (2017) - \"La metodologia per la valutazione e gestione "
    "del rischio stress lavoro-correlato. Manuale ad uso delle "
    "aziende\" - edizione aggiornata",
    "INAIL (2011) - prima edizione del Manuale (Network Nazionale "
    "per la prevenzione del Disagio Psicosociale nei Luoghi di "
    "Lavoro)",
    "INAIL - Piattaforma online per la valutazione del rischio "
    "SLC (accesso via Punto Cliente)",
    "ISO 45003:2021 - Gestione della salute e sicurezza "
    "psicologica sul lavoro",
    "EU-OSHA - Calculating the cost of work-related stress and "
    "psychosocial risks",
    "HSE - Management Standards Indicator Tool (versione "
    "originale)",
], size=11)

# =========================================================
# SLIDE 22 - CONCLUSIONI / Q&A
# =========================================================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, BLU_DIFESA)
add_rect(s, 0, Inches(2.6), SW, Inches(0.08), ORO)

add_text(s, Inches(0.6), Inches(0.6), Inches(12), Inches(0.5),
         "MINISTERO DELLA DIFESA - MARINA MILITARE",
         size=14, bold=True, color=ORO)

add_text(s, Inches(0.6), Inches(2.95), Inches(12), Inches(1.0),
         "MESSAGGI CHIAVE",
         size=36, bold=True, color=BIANCO)

msgs = [
    "1. La valutazione SLC e' un OBBLIGO di legge (art. 28 c.1-bis) "
    "e parte integrante del DVR.",
    "2. La Metodologia INAIL (ed. 2017) e' lo strumento di "
    "riferimento, conforme alle indicazioni della Commissione "
    "Consultiva.",
    "3. La PARTECIPAZIONE attiva di lavoratori, RLS e MC e' un "
    "requisito di validita' del processo.",
    "4. La valutazione e' un PROCESSO CICLICO: pianificare - fare "
    "- verificare - agire.",
    "5. Negli arsenali militari il modello va calato sulle "
    "SPECIFICITA' organizzative e operative del contesto difensivo.",
]
add_bullets(s, Inches(0.6), Inches(4.0), Inches(12.1), Inches(2.7),
            msgs, size=15, color=BIANCO)

add_rect(s, 0, SH - Inches(0.6), SW, Inches(0.6), NERO)
add_text(s, Inches(0.6), SH - Inches(0.55), Inches(12), Inches(0.5),
         "Grazie per l'attenzione - Servizio di Prevenzione e "
         "Protezione, Arsenale Militare",
         size=14, bold=True, color=ORO,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# =========================================================
# Salvataggio
# =========================================================
out = "/projects/sandbox/ITALIC/INAIL_Stress_Lavoro_Correlato_" \
      "Arsenali_Militari.pptx"
prs.save(out)
print(f"OK -> {out}")
print(f"Slide totali: {len(prs.slides)}")
