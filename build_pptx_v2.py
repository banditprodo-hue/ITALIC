"""
Presentazione PPT v2 - GUIDA OPERATIVA per l'Ammiraglio
Valutazione Rischio Stress Lavoro-Correlato
Integra: Linee Guida INAIL 2017 + Monografia INAIL 2024/2025
         + Strumenti operativi (Questionario DOCX, Excel, Flowchart)
Contesto: Arsenale Militare Marittimo
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# Palette istituzionale
BLU = RGBColor(0x0B, 0x2E, 0x55)
BLU2 = RGBColor(0x1F, 0x4E, 0x79)
ORO = RGBColor(0xC9, 0xA2, 0x27)
GRIGIO = RGBColor(0x55, 0x55, 0x55)
GR_CH = RGBColor(0xEC, 0xEC, 0xEC)
BIANCO = RGBColor(0xFF, 0xFF, 0xFF)
NERO = RGBColor(0x1A, 0x1A, 0x1A)
ROSSO = RGBColor(0xB3, 0x1B, 0x1B)
VERDE = RGBColor(0x2E, 0x7D, 0x32)
GIALLO = RGBColor(0xF9, 0xA8, 0x25)
ARANCIO = RGBColor(0xE6, 0x5C, 0x00)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]
TOTAL = 28



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



def add_bullets(slide, x, y, w, h, items, *, size=14, color=NERO,
                bullet="\u2022"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(4)
        prefix = f"{bullet} " if bullet else ""
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


def header_bar(slide, title, subtitle=None, n=None):
    add_rect(slide, 0, 0, SW, Inches(0.9), BLU)
    add_rect(slide, 0, Inches(0.9), SW, Inches(0.05), ORO)
    add_text(slide, Inches(0.4), Inches(0.12), Inches(11.5), Inches(0.55),
             title, size=24, bold=True, color=BIANCO,
             anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, Inches(0.4), Inches(0.55), Inches(11.5),
                 Inches(0.32), subtitle, size=12, color=ORO,
                 anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, 0, SH - Inches(0.35), SW, Inches(0.35), BLU)
    add_text(slide, Inches(0.3), SH - Inches(0.32), Inches(8),
             Inches(0.3),
             "Progetto SLC - Arsenale Militare | Briefing Ammiraglio",
             size=9, color=BIANCO, anchor=MSO_ANCHOR.MIDDLE)
    if n:
        add_text(slide, SW - Inches(1.5), SH - Inches(0.32),
                 Inches(1.2), Inches(0.3), f"{n}/{TOTAL}",
                 size=9, color=ORO, align=PP_ALIGN.RIGHT,
                 anchor=MSO_ANCHOR.MIDDLE)



# =============================================================
# SLIDE 1 - COPERTINA
# =============================================================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, BLU)
add_rect(s, 0, Inches(2.6), SW, Inches(0.08), ORO)
add_rect(s, 0, Inches(5.2), SW, Inches(0.08), ORO)
add_text(s, Inches(0.6), Inches(0.5), Inches(12), Inches(0.4),
         "MINISTERO DELLA DIFESA - MARINA MILITARE",
         size=14, bold=True, color=ORO)
add_text(s, Inches(0.6), Inches(0.9), Inches(12), Inches(0.4),
         "Servizio di Prevenzione e Protezione - Arsenale Militare",
         size=12, color=BIANCO)
add_text(s, Inches(0.6), Inches(2.9), Inches(12.1), Inches(1.2),
         "PROGETTO DI VALUTAZIONE DEL RISCHIO\nSTRESS LAVORO-CORRELATO",
         size=38, bold=True, color=BIANCO)
add_text(s, Inches(0.6), Inches(4.4), Inches(12.1), Inches(0.6),
         "Briefing informativo per il Datore di Lavoro (Ammiraglio Direttore)",
         size=18, color=BIANCO)
add_text(s, Inches(0.6), Inches(5.4), Inches(12.1), Inches(0.6),
         "Linee Guida INAIL 2017 + Monografia INAIL 2024/2025\n"
         "Modulo Smart Working / Technostress",
         size=18, color=ORO, bold=True)
add_text(s, Inches(0.6), Inches(6.5), Inches(12.1), Inches(0.5),
         "D.Lgs. 81/2008 art. 28 c.1-bis | Accordo Europeo 8/10/2004 | "
         "Circ. Min. Lavoro 18/11/2010 | INAIL ed. 2017 + 2024/2025",
         size=11, color=GR_CH)



# =============================================================
# SLIDE 2 - PERCHE' SIAMO QUI
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Perche' siamo qui",
           "Obbligo normativo e opportunita' gestionale", 2)
add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5),
         "IL QUADRO IN SINTESI PER IL DATORE DI LAVORO",
         size=20, bold=True, color=BLU)
add_bullets(s, Inches(0.6), Inches(1.8), Inches(12), Inches(5.2), [
    ("OBBLIGO DI LEGGE:",
     "L'art. 28 comma 1-bis del D.Lgs. 81/2008 impone la valutazione "
     "del rischio Stress Lavoro-Correlato (SLC) per TUTTI i lavoratori, "
     "incluso il personale civile e militare degli arsenali."),
    ("NON E' FACOLTATIVO:",
     "La mancata valutazione e' sanzionata penalmente a carico del "
     "Datore di Lavoro (arresto da 3 a 6 mesi o ammenda da 3.071 a "
     "7.862 EUR - art. 55 c.1 lett. a) D.Lgs. 81/08)."),
    ("PRIMA VALUTAZIONE:",
     "Questo Arsenale non ha ancora effettuato una valutazione SLC "
     "formale. E' necessario avviare il processo con urgenza, "
     "documentandolo adeguatamente nel DVR."),
    ("NOVITA' 2024/2025:",
     "INAIL ha pubblicato una Monografia (aprile 2024/2025) che "
     "integra il modulo SMART WORKING e TECHNOSTRESS - obbligatorio "
     "per il personale in lavoro agile."),
    ("STRUMENTI GIA' PRONTI:",
     "Il SPP ha predisposto un kit operativo completo (Questionario "
     "Word + Foglio Excel automatizzato + Flowchart di processo) "
     "pronto per l'avvio immediato."),
    ("COSA SERVE DA LEI:",
     "1) Approvazione formale dell'avvio del progetto; "
     "2) Nomina del Gruppo di Gestione (GGV); "
     "3) Firma della comunicazione ai lavoratori."),
], size=14, color=NERO)



# =============================================================
# SLIDE 3 - COSA E' LO STRESS LAVORO-CORRELATO
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Cos'e' lo Stress Lavoro-Correlato (SLC)",
           "Definizione ufficiale - semplificata per il decisore", 3)
add_rect(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(1.6), GR_CH)
add_text(s, Inches(0.7), Inches(1.3), Inches(12), Inches(1.5),
         "DEFINIZIONE (Accordo Europeo 8/10/2004):\n"
         "Condizione in cui il lavoratore non si sente in grado di "
         "corrispondere alle richieste o aspettative lavorative.\n"
         "NON E' una malattia, ma puo' causare malattie.",
         size=14, color=NERO, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.5), Inches(3.0), Inches(12.3), Inches(0.5),
         "ATTENZIONE - cosa NON e':", size=16, bold=True, color=ROSSO)
add_bullets(s, Inches(0.6), Inches(3.5), Inches(12), Inches(1.5), [
    "NON e' lo stress personale del singolo (problemi familiari, ecc.)",
    "NON e' il mobbing (che richiede valutazione separata)",
    "NON e' lo stress post-traumatico",
], size=14, color=NERO)
add_text(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(0.5),
         "COSA SI VALUTA:", size=16, bold=True, color=VERDE)
add_bullets(s, Inches(0.6), Inches(5.5), Inches(12), Inches(1.5), [
    "L'ORGANIZZAZIONE del lavoro (carichi, turni, ruoli, comunicazione)",
    "Le CONDIZIONI di lavoro (ambiente, attrezzature, autonomia)",
    "Le RELAZIONI sul lavoro (supporto capi, colleghi, clima)",
], size=14, color=NERO)



# =============================================================
# SLIDE 4 - NOVITA' INAIL 2024/2025
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Novita' INAIL 2024/2025",
           "Cosa cambia rispetto al modello precedente", 4)
add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4),
         "EVOLUZIONE DELLA METODOLOGIA INAIL",
         size=18, bold=True, color=BLU)
# Timeline
timeline = [
    ("2011", "Prima edizione del Manuale INAIL - 35 indicatori"),
    ("2017", "Revisione: validazione Q-IND, differenze genere/eta'"),
    ("2024/2025", "MONOGRAFIA: Modulo Smart Working + Technostress "
     "(27 item aggiuntivi) - Lavoro da remoto e ICT intensivo"),
]
y = Inches(1.7)
for anno, desc in timeline:
    add_rect(s, Inches(0.5), y, Inches(1.5), Inches(0.8), BLU)
    add_text(s, Inches(0.55), y, Inches(1.4), Inches(0.8),
             anno, size=20, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(2.1), y, Inches(10.7), Inches(0.8),
             GR_CH, line=BLU)
    add_text(s, Inches(2.2), y, Inches(10.5), Inches(0.8),
             desc, size=14, color=NERO, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(1.0)

add_text(s, Inches(0.5), Inches(4.8), Inches(12.3), Inches(0.4),
         "LE 3 NUOVE AREE DELLA MONOGRAFIA 2024/2025:",
         size=16, bold=True, color=BLU)
# 3 colonne
cols = [
    ("REMOTE WORK\n(10 item)", "Postazione remota\nComunicazione\n"
     "Isolamento\nSupporto a distanza\nObiettivi remoti", BLU2),
    ("TECHNOSTRESS\n(15 item)", "Sovraccarico info\nAnsia tecnologica\n"
     "Monitoraggio invasivo\nAffaticamento ICT\nObsolescenza digitale",
     ARANCIO),
    ("WORK-LIFE\nBALANCE (2 item)", "Confusione spazi\nDiritto alla\n"
     "disconnessione", VERDE),
]
x = Inches(0.5)
for title, body, col in cols:
    add_rect(s, x, Inches(5.3), Inches(4.0), Inches(1.8), GR_CH, line=col)
    add_rect(s, x, Inches(5.3), Inches(4.0), Inches(0.6), col)
    add_text(s, x + Inches(0.1), Inches(5.3), Inches(3.8), Inches(0.6),
             title, size=12, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.15), Inches(5.95), Inches(3.7), Inches(1.1),
             body, size=11, color=NERO)
    x += Inches(4.1)



# =============================================================
# SLIDE 5 - IL PROCESSO IN 6 FASI (dal flowchart)
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Il processo di valutazione SLC in 6 Fasi",
           "Flowchart conforme alle Linee Guida INAIL 2017 + 2024/2025", 5)
phases = [
    ("1", "PROPEDEUTICA", "Costituzione GGV\nDefinizione gruppi omogenei\n"
     "Comunicazione al personale", BLU2),
    ("2", "GRUPPI\nOMOGENEI", "Identificazione per reparto,\n"
     "mansione, turno, modalita'", ORO),
    ("3", "VALUTAZIONE\nPRELIMINARE", "Lista di controllo INAIL\n"
     "Aree A + B + C\nCalcolo punteggio", BLU2),
    ("4", "DECISIONE", "Rischio Trascurabile?\n"
     "SI -> Monitoraggio\nNO -> Fase Approfondita", GIALLO),
    ("5", "APPROFONDITA\n(eventuale)", "Questionario 35+27 item\n"
     "Focus group / Interviste\nAnalisi aggregata", ROSSO),
    ("6", "PIANO +\nDVR", "Misure correttive\nInserimento nel DVR\n"
     "Data rivalutazione", VERDE),
]
x = Inches(0.3)
w = Inches(2.0)
h = Inches(4.5)
y = Inches(1.2)
for num, title, body, col in phases:
    add_rect(s, x, y, w, h, GR_CH, line=GRIGIO)
    add_rect(s, x, y, w, Inches(0.5), col)
    add_text(s, x + Inches(0.05), y, w - Inches(0.1), Inches(0.5),
             f"FASE {num}", size=12, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.05), y + Inches(0.55), w - Inches(0.1),
             Inches(0.7), title, size=11, bold=True, color=col,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.1), y + Inches(1.4), w - Inches(0.2),
             h - Inches(1.5), body, size=10, color=NERO)
    x += w + Inches(0.12)

add_rect(s, Inches(0.3), Inches(5.9), Inches(12.7), Inches(0.8),
         GR_CH, line=ORO)
add_text(s, Inches(0.5), Inches(5.95), Inches(12.5), Inches(0.7),
         "MONITORAGGIO CONTINUO: Rivalutazione ogni 2-3 anni o "
         "immediatamente dopo cambiamenti organizzativi significativi "
         "(riorganizzazione, nuove tecnologie, variazione organici)",
         size=13, bold=True, color=BLU, anchor=MSO_ANCHOR.MIDDLE)



# =============================================================
# SLIDE 6 - RUOLO DEL DATORE DI LAVORO (Ammiraglio)
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Cosa deve fare il Datore di Lavoro",
           "Ruolo dell'Ammiraglio Direttore nel processo SLC", 6)
add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4),
         "LE 7 AZIONI CHE COMPETONO AL DATORE DI LAVORO",
         size=18, bold=True, color=BLU)
actions = [
    ("1. APPROVARE l'avvio",
     "Autorizzare formalmente l'avvio della valutazione SLC "
     "(obbligo indelegabile ex art. 17 D.Lgs. 81/08)."),
    ("2. NOMINARE il GGV",
     "Costituire il Gruppo di Gestione della Valutazione "
     "(RSPP + MC + RLS + Capi Reparto)."),
    ("3. FIRMARE la comunicazione",
     "Sottoscrivere l'informativa ai lavoratori sull'avvio "
     "della valutazione."),
    ("4. GARANTIRE le risorse",
     "Assicurare tempo, strumenti e budget per il processo "
     "(stampa questionari, eventuale consulente)."),
    ("5. APPROVARE il DVR",
     "Sottoscrivere la sezione SLC del Documento di Valutazione "
     "dei Rischi con data certa."),
    ("6. ATTUARE le misure",
     "Disporre l'attuazione delle misure correttive individuate, "
     "assegnando responsabili e tempistiche."),
    ("7. MONITORARE",
     "Ricevere report periodici sull'andamento e disporre la "
     "rivalutazione nei tempi previsti."),
]
y = Inches(1.6)
for title, desc in actions:
    add_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.7), GR_CH, line=BLU)
    add_text(s, Inches(0.6), y, Inches(3.5), Inches(0.7),
             title, size=12, bold=True, color=BLU,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(4.2), y, Inches(8.5), Inches(0.7),
             desc, size=12, color=NERO, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.75)



# =============================================================
# SLIDE 7 - GLI STRUMENTI OPERATIVI PREDISPOSTI
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Strumenti operativi gia' predisposti dal SPP",
           "Kit pronto all'uso - in attesa di approvazione", 7)
add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4),
         "3 STRUMENTI COMPLEMENTARI + PRESENTAZIONE GUIDA",
         size=18, bold=True, color=BLU)

tools = [
    ("QUESTIONARIO\nWORD (.docx)",
     "Questionario_SLC_INAIL_2025.docx",
     "35 item standard + 27 item Smart/ICT\n"
     "Scala Likert 1-5 | Anonimo\n"
     "Informativa privacy integrata\n"
     "Stampabile per compilazione cartacea\n"
     "Sezione firme RSPP/MC/RLS",
     BLU2),
    ("FOGLIO EXCEL\nAUTOMATIZZATO (.xlsx)",
     "Valutazione_SLC_INAIL_2025.xlsx",
     "10 fogli di lavoro automatizzati\n"
     "Fino a 20 gruppi omogenei\n"
     "Calcolo automatico punteggi\n"
     "Classificazione rischio a semaforo\n"
     "Report pronto per il DVR",
     VERDE),
    ("FLOWCHART\nINTERATTIVO (.html)",
     "Flowchart_SLC_Interattivo.html",
     "Diagramma di flusso cliccabile\n"
     "Spiegazione di ogni fase\n"
     "Ruoli e responsabilita'\n"
     "Apribile su qualsiasi browser\n"
     "Guida visiva per il DL",
     ARANCIO),
]
x = Inches(0.5)
for title, filename, desc, col in tools:
    add_rect(s, x, Inches(1.7), Inches(4.0), Inches(5.0), GR_CH, line=col)
    add_rect(s, x, Inches(1.7), Inches(4.0), Inches(0.8), col)
    add_text(s, x + Inches(0.1), Inches(1.7), Inches(3.8), Inches(0.8),
             title, size=14, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.15), Inches(2.6), Inches(3.7), Inches(0.5),
             filename, size=10, bold=True, color=col,
             align=PP_ALIGN.CENTER)
    add_text(s, x + Inches(0.15), Inches(3.1), Inches(3.7), Inches(3.5),
             desc, size=12, color=NERO)
    x += Inches(4.15)



# =============================================================
# SLIDE 8 - FASE 1 PROPEDEUTICA (dettaglio)
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "FASE 1 - Propedeutica",
           "Cosa fare PRIMA di iniziare la valutazione", 8)
add_text(s, Inches(0.5), Inches(1.15), Inches(6.0), Inches(0.4),
         "A) Costituzione del GGV", size=16, bold=True, color=BLU)
add_bullets(s, Inches(0.6), Inches(1.55), Inches(5.9), Inches(2.2), [
    "Ammiraglio Direttore (presiede o delega)",
    "RSPP (coordina tecnicamente)",
    "Medico Competente",
    "RLS / RLST",
    "Capi Reparto / Officina interessati",
    "Eventuale psicologo del lavoro",
], size=12)
add_text(s, Inches(0.5), Inches(3.8), Inches(6.0), Inches(0.4),
         "B) Comunicazione al personale", size=16, bold=True, color=BLU)
add_bullets(s, Inches(0.6), Inches(4.2), Inches(5.9), Inches(2.0), [
    "Informativa scritta (firmata dal DL)",
    "Consultazione preventiva RLS (art. 50)",
    "Garanzia di anonimato (GDPR)",
    "Restituzione esiti in forma aggregata",
], size=12)
add_text(s, Inches(7.0), Inches(1.15), Inches(6.0), Inches(0.4),
         "C) Definizione gruppi omogenei", size=16, bold=True, color=BLU)
add_bullets(s, Inches(7.1), Inches(1.55), Inches(5.8), Inches(2.2), [
    "Per REPARTO (Officina Motori, Carpenteria...)",
    "Per MANSIONE (operai, impiegati, tecnici...)",
    "Per TURNO (giornaliero, H24, reperibilita')",
    "Per MODALITA' (presenza, smart working, ICT)",
    "Minimo 6 lavoratori per gruppo",
], size=12)
add_text(s, Inches(7.0), Inches(3.8), Inches(6.0), Inches(0.4),
         "D) Cronoprogramma", size=16, bold=True, color=BLU)
add_bullets(s, Inches(7.1), Inches(4.2), Inches(5.8), Inches(2.5), [
    "Settimana 1-2: Nomina GGV + comunicazione",
    "Settimana 3-4: Definizione gruppi + stampa questionari",
    "Settimana 5-8: Somministrazione e raccolta",
    "Settimana 9-10: Elaborazione dati (Excel)",
    "Settimana 11-12: Redazione DVR + report",
    "TOTALE: circa 90 giorni",
], size=12)



# =============================================================
# SLIDE 9 - FASE 2/3 VALUTAZIONE PRELIMINARE
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "FASE 2/3 - Valutazione Preliminare",
           "Lista di controllo INAIL: 3 Aree di indagine", 9)
areas = [
    ("AREA A\nEVENTI SENTINELLA", "Dati OGGETTIVI triennali:\n"
     "- Infortuni\n- Assenze malattia\n- Turnover\n- Ferie non godute\n"
     "- Sanzioni disciplinari\n- Visite straordinarie MC",
     "Il MC fornisce i dati sanitari\nRSPP raccoglie da Ufficio Personale",
     ROSSO),
    ("AREA B\nCONTENUTO LAVORO", "Caratteristiche del LAVORO:\n"
     "- Ambiente e attrezzature\n- Carichi e ritmi\n- Turni e orari\n"
     "- Competenze vs mansioni\n- Autonomia operativa",
     "RSPP compila con sopralluogo\nRLS contribuisce",
     BLU2),
    ("AREA C\nCONTESTO LAVORO", "Caratteristiche ORGANIZZATIVE:\n"
     "- Chiarezza del ruolo\n- Comunicazione interna\n- Supporto dei capi\n"
     "- Relazioni colleghi\n- Gestione cambiamenti",
     "RSPP + RLS compilano insieme\nConsultazione lavoratori",
     VERDE),
]
x = Inches(0.5)
for title, body, note, col in areas:
    add_rect(s, x, Inches(1.2), Inches(4.0), Inches(5.6), GR_CH, line=col)
    add_rect(s, x, Inches(1.2), Inches(4.0), Inches(0.8), col)
    add_text(s, x + Inches(0.1), Inches(1.2), Inches(3.8), Inches(0.8),
             title, size=13, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.15), Inches(2.1), Inches(3.7), Inches(3.2),
             body, size=11, color=NERO)
    add_rect(s, x + Inches(0.1), Inches(5.4), Inches(3.8), Inches(1.3),
             col)
    add_text(s, x + Inches(0.2), Inches(5.5), Inches(3.6), Inches(1.1),
             note, size=10, bold=True, color=BIANCO,
             anchor=MSO_ANCHOR.MIDDLE)
    x += Inches(4.15)



# =============================================================
# SLIDE 10 - FASCE DI RISCHIO
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Esito della valutazione - Fasce di rischio",
           "Come leggere i risultati", 10)
add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4),
         "SEMAFORO DEL RISCHIO SLC (dal foglio Excel automatizzato)",
         size=16, bold=True, color=BLU)
fasce = [
    ("VERDE", "RISCHIO\nTRASCURABILE", "Punteggio medio <= 2.3",
     "Nessuna azione correttiva necessaria.\n"
     "Monitoraggio ogni 2-3 anni.\n"
     "Inserimento nel DVR con esito favorevole.",
     VERDE),
    ("GIALLO", "RISCHIO\nMODERATO", "Punteggio medio 2.3 - 3.0",
     "Adozione misure correttive mirate.\n"
     "Verifica efficacia entro 12 mesi.\n"
     "Se inefficaci: passare a Fase Approfondita.",
     GIALLO),
    ("ROSSO", "RISCHIO\nRILEVANTE", "Punteggio medio > 3.0",
     "Misure correttive IMMEDIATE.\n"
     "Avvio Valutazione Approfondita (questionario).\n"
     "Coinvolgimento MC + eventuale psicologo.",
     ROSSO),
]
y = Inches(1.7)
for sem, title, soglia, desc, col in fasce:
    add_rect(s, Inches(0.5), y, Inches(1.5), Inches(1.7), col)
    add_text(s, Inches(0.55), y, Inches(1.4), Inches(1.7),
             sem, size=18, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(2.1), y, Inches(2.5), Inches(1.7), GR_CH, line=col)
    add_text(s, Inches(2.2), y, Inches(2.3), Inches(1.7),
             title + "\n\n" + soglia, size=12, bold=True, color=col,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(4.7), y, Inches(8.1), Inches(1.7), GR_CH, line=col)
    add_text(s, Inches(4.9), y + Inches(0.1), Inches(7.8), Inches(1.5),
             desc, size=13, color=NERO, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(1.85)



# =============================================================
# SLIDE 11 - VALUTAZIONE APPROFONDITA
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "FASE 5 - Valutazione Approfondita (eventuale)",
           "Quando e come si usa il questionario ai lavoratori", 11)
add_rect(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.9),
         GR_CH, line=ROSSO)
add_text(s, Inches(0.7), Inches(1.2), Inches(12), Inches(0.85),
         "SI ATTIVA SOLO SE: la fase preliminare da' esito RILEVANTE "
         "oppure le misure correttive adottate non funzionano entro 12 mesi.",
         size=14, bold=True, color=ROSSO, anchor=MSO_ANCHOR.MIDDLE)
add_text(s, Inches(0.5), Inches(2.3), Inches(6.0), Inches(0.4),
         "Lo strumento: Questionario 35+27 item", size=15, bold=True,
         color=BLU)
add_bullets(s, Inches(0.6), Inches(2.7), Inches(5.9), Inches(4.0), [
    "PARTE 1 - Standard (35 item): OBBLIGATORIA per tutti",
    "Contenuto del lavoro (10 item)",
    "Contesto del lavoro (9 item)",
    "Relazioni interpersonali (7 item)",
    "Cambiamento (4 item)",
    "Equilibrio vita-lavoro (5 item)",
    "",
    "PARTE 2 - Smart/ICT (27 item): SOLO per chi lavora",
    "in smart working o con uso intensivo ICT",
    "Remote Work (10 item)",
    "Technostress (15 item)",
    "Work-Life Balance (2 item)",
], size=12)
add_text(s, Inches(7.0), Inches(2.3), Inches(6.0), Inches(0.4),
         "Regole operative", size=15, bold=True, color=BLU)
add_bullets(s, Inches(7.1), Inches(2.7), Inches(5.8), Inches(4.0), [
    "Compilazione ANONIMA (no nome, no CF)",
    "Solo ID compilazione + ID gruppo omogeneo",
    "Scala Likert 1 (MAI) - 5 (SEMPRE)",
    "Soglia minima adesione: >= 60% del gruppo",
    "Consegna in busta chiusa / urna",
    "Elaborazione SOLO aggregata per gruppo",
    "",
    "RISPETTO GDPR: Reg. UE 679/2016",
    "Base giuridica: art. 9 par.2 lett. b)",
    "(obbligo ex D.Lgs. 81/08)",
    "Nessun dato nominativo conservato",
], size=12)



# =============================================================
# SLIDE 12 - IL FOGLIO EXCEL (come funziona)
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Come funziona il foglio Excel automatizzato",
           "Valutazione_SLC_INAIL_2025.xlsx - 10 fogli integrati", 12)
add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4),
         "FLUSSO DI LAVORO NEL FILE EXCEL", size=16, bold=True, color=BLU)
steps = [
    ("1\nCONFIGURA", "Inserisci nome ente,\nresponsabili,\ndata valutazione",
     BLU2),
    ("2\nGRUPPI", "Definisci fino a\n20 gruppi omogenei\n(reparto, mansione)",
     BLU2),
    ("3\nINSERISCI", "Trascrivi le risposte\ndei questionari\n(1 riga per persona)",
     ORO),
    ("4\nAUTOMATICO", "Il foglio calcola:\npunteggi per area,\ninversione item [R]",
     VERDE),
    ("5\nRISULTATI", "Classificazione\na semaforo per\nogni gruppo",
     ROSSO),
]
x = Inches(0.3)
for title, body, col in steps:
    add_rect(s, x, Inches(1.7), Inches(2.4), Inches(2.8), GR_CH, line=col)
    add_rect(s, x, Inches(1.7), Inches(2.4), Inches(0.7), col)
    add_text(s, x + Inches(0.05), Inches(1.7), Inches(2.3), Inches(0.7),
             title, size=11, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.1), Inches(2.5), Inches(2.2), Inches(2.0),
             body, size=11, color=NERO, align=PP_ALIGN.CENTER)
    x += Inches(2.55)

add_text(s, Inches(0.5), Inches(4.7), Inches(12.3), Inches(0.4),
         "DETTAGLIO DEI 10 FOGLI:", size=14, bold=True, color=BLU)
fogli = [
    "README_Istruzioni - Guida completa con riferimenti normativi",
    "Configurazione - Dati organizzazione, soglie cut-off personalizzabili",
    "Gruppi_Omogenei - Anagrafica dei gruppi (fino a 20)",
    "DataEntry_Standard - Inserimento 35 item (fino a 50 compilazioni)",
    "DataEntry_SmartICT - Inserimento 27 item aggiuntivi (solo gruppi ICT)",
    "Calcolo_Punteggi - Inversione automatica, punteggi per sezione",
    "Aggregazione_Gruppi - Media/DS per gruppo omogeneo",
    "Classificazione_Rischio - Fascia + priorita' intervento",
    "Report_Risultati - Tabelle pronte per il DVR",
    "Privacy_Log - Traccia compilazioni GDPR-compliant",
]
add_bullets(s, Inches(0.6), Inches(5.1), Inches(12), Inches(2.0),
            fogli, size=10, color=NERO)



# =============================================================
# SLIDE 13 - SPECIFICITA' ARSENALE MILITARE
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Specificita' del nostro Arsenale",
           "Fattori che rendono la valutazione SLC particolarmente rilevante",
           13)
add_bullets(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.8), [
    ("Personale misto civile/militare:",
     "ordinamento gerarchico + CCNL Funzioni Centrali - potenziali "
     "tensioni di ruolo."),
    ("Doppia catena di comando:",
     "gerarchica militare e tecnico-funzionale (SPP) - "
     "possibile ambiguita' di ruolo."),
    ("Lavorazioni ad alta complessita':",
     "cantieristica navale, saldatura, spazi confinati, lavori in "
     "quota - fattore CONTENUTO critico."),
    ("Tempi compressi:",
     "consegna unita' navali in prontezza operativa - pressione "
     "temporale elevata."),
    ("Esposizioni multiple:",
     "rumore, vibrazioni, agenti chimici, microclima - stress "
     "ambientale cumulativo."),
    ("Segretezza:",
     "vincoli di riservatezza su programmi militari - possibile "
     "isolamento informativo."),
    ("Mobilita' forzata:",
     "trasferimenti del personale militare per esigenze di "
     "servizio - rottura delle reti sociali."),
    ("Smart working limitato ma presente:",
     "personale amministrativo/tecnico in telelavoro - necessario "
     "il modulo ICT 2024/2025."),
], size=13, color=NERO)

# =============================================================
# SLIDE 14 - SANZIONI
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Conseguenze della mancata valutazione",
           "Profilo sanzionatorio - art. 55 D.Lgs. 81/08", 14)
add_rect(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(2.5),
         GR_CH, line=ROSSO)
add_text(s, Inches(0.7), Inches(1.3), Inches(12), Inches(0.4),
         "SANZIONI PER IL DATORE DI LAVORO (art. 55 c.1 lett. a):",
         size=16, bold=True, color=ROSSO)
add_bullets(s, Inches(0.8), Inches(1.75), Inches(11.8), Inches(1.8), [
    "Arresto da 3 a 6 mesi",
    "OPPURE ammenda da 3.071,27 EUR a 7.862,44 EUR",
    "In caso di piu' violazioni: cumulo delle sanzioni",
    "Prescrizione obbligatoria dell'organo di vigilanza (art. 20 D.Lgs. 758/94)",
], size=14, color=NERO)
add_text(s, Inches(0.5), Inches(4.0), Inches(12.3), Inches(0.4),
         "ALTRE CONSEGUENZE:", size=16, bold=True, color=BLU)
add_bullets(s, Inches(0.6), Inches(4.4), Inches(12), Inches(2.8), [
    ("Responsabilita' civile:",
     "risarcimento danni in caso di patologia SLC-correlata "
     "non prevenuta (art. 2087 c.c.)."),
    ("Responsabilita' dirigenziale:",
     "danno erariale per mancata adozione di misure obbligatorie "
     "(Corte dei Conti)."),
    ("Responsabilita' ex D.Lgs. 231/01:",
     "applicabile anche alle PA per reati contro la sicurezza "
     "sul lavoro."),
    ("Reputazione istituzionale:",
     "contenziosi sindacali, segnalazioni ANAC, "
     "impatto mediatico negativo."),
], size=13, color=NERO)



# =============================================================
# SLIDE 15 - CRONOPROGRAMMA VISIVO
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Cronoprogramma del progetto",
           "90 giorni dalla firma all'inserimento nel DVR", 15)
gantt = [
    ("Sett. 1-2", "Nomina GGV + Comunicazione personale", BLU2, 2),
    ("Sett. 3-4", "Definizione gruppi + Stampa questionari", ORO, 2),
    ("Sett. 5-8", "Somministrazione questionari + Raccolta", VERDE, 4),
    ("Sett. 9-10", "Data entry Excel + Elaborazione", BLU2, 2),
    ("Sett. 11-12", "Redazione sezione DVR + Piano misure", ROSSO, 2),
    ("Sett. 13", "Firma DL + Comunicazione esiti", ORO, 1),
]
y = Inches(1.3)
# Header
add_text(s, Inches(0.5), y, Inches(2.5), Inches(0.5),
         "ATTIVITA'", size=12, bold=True, color=BIANCO)
add_rect(s, Inches(0.5), y, Inches(2.5), Inches(0.5), BLU)
add_text(s, Inches(0.55), y, Inches(2.4), Inches(0.5),
         "ATTIVITA'", size=12, bold=True, color=BIANCO,
         anchor=MSO_ANCHOR.MIDDLE)
for w in range(13):
    add_rect(s, Inches(3.1 + w * 0.75), y, Inches(0.75), Inches(0.5), BLU)
    add_text(s, Inches(3.1 + w * 0.75), y, Inches(0.75), Inches(0.5),
             str(w + 1), size=9, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
y += Inches(0.55)
week_start = 0
for label, desc, col, duration in gantt:
    add_rect(s, Inches(0.5), y, Inches(2.5), Inches(0.8), GR_CH, line=GRIGIO)
    add_text(s, Inches(0.55), y, Inches(2.4), Inches(0.8),
             f"{label}\n{desc}", size=10, color=NERO,
             anchor=MSO_ANCHOR.MIDDLE)
    # barra GANTT
    bar_x = Inches(3.1 + week_start * 0.75)
    bar_w = Inches(duration * 0.75)
    add_rect(s, bar_x, y + Inches(0.2), bar_w, Inches(0.4), col)
    week_start += duration
    y += Inches(0.85)

add_rect(s, Inches(0.5), Inches(6.3), Inches(12.3), Inches(0.8),
         GR_CH, line=ORO)
add_text(s, Inches(0.7), Inches(6.35), Inches(12), Inches(0.7),
         "NOTA: Il cronoprogramma e' indicativo. Il DL puo' accelerare "
         "i tempi autorizzando le risorse necessarie fin dalla prima "
         "settimana. Il RSPP gestisce operativamente le attivita'.",
         size=12, color=GRIGIO, anchor=MSO_ANCHOR.MIDDLE)



# =============================================================
# SLIDE 16 - MISURE CORRETTIVE TIPICHE
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Misure correttive tipiche",
           "Interventi graduati in base all'esito", 16)
cats = [
    ("ORGANIZZATIVE", "Revisione turni e carichi\nRotazione mansioni\n"
     "Pause programmate\nChiarificazione ruoli", BLU2),
    ("COMUNICATIVE", "Briefing periodici\nCanali segnalazione\n"
     "Feedback strutturato\nTrasparenza obiettivi", ORO),
    ("FORMATIVE", "Formazione preposti\nGestione del personale\n"
     "Comunicazione efficace\nBenessere organizzativo", VERDE),
    ("SUPPORTO", "Sportello ascolto\nSorveglianza sanitaria\n"
     "Supporto psicologico\nMediazione conflitti", ROSSO),
]
x = Inches(0.5)
for title, body, col in cats:
    add_rect(s, x, Inches(1.2), Inches(3.0), Inches(3.5), GR_CH, line=col)
    add_rect(s, x, Inches(1.2), Inches(3.0), Inches(0.6), col)
    add_text(s, x + Inches(0.1), Inches(1.2), Inches(2.8), Inches(0.6),
             title, size=13, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, x + Inches(0.15), Inches(1.9), Inches(2.7), Inches(2.7),
             body, size=12, color=NERO)
    x += Inches(3.15)

add_rect(s, Inches(0.5), Inches(5.0), Inches(12.3), Inches(2.0),
         GR_CH, line=BLU)
add_text(s, Inches(0.7), Inches(5.1), Inches(12), Inches(0.4),
         "PRINCIPIO GUIDA (art. 15 D.Lgs. 81/08):", size=14, bold=True,
         color=BLU)
add_text(s, Inches(0.7), Inches(5.5), Inches(12), Inches(1.4),
         "Le misure devono seguire la GERARCHIA DEI CONTROLLI:\n"
         "1) Eliminare il rischio alla fonte (es. riorganizzazione)\n"
         "2) Ridurre il rischio (es. rotazione, pause)\n"
         "3) Proteggere il lavoratore (es. formazione, supporto)\n"
         "Ogni misura deve avere: RESPONSABILE + TEMPISTICA + "
         "INDICATORE DI EFFICACIA",
         size=12, color=NERO)

# =============================================================
# SLIDE 17 - PRIVACY E GARANZIE
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Privacy e garanzie per i lavoratori",
           "Conformita' GDPR - Reg. UE 679/2016", 17)
add_text(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(0.4),
         "GARANZIE INTEGRATE NEL SISTEMA", size=18, bold=True, color=BLU)
garanzie = [
    ("ANONIMATO TOTALE",
     "Nessun dato identificativo nel questionario. "
     "Solo ID compilazione (assegnato dal RSPP) e ID gruppo omogeneo."),
    ("ELABORAZIONE AGGREGATA",
     "I risultati sono elaborati e comunicati ESCLUSIVAMENTE "
     "in forma aggregata per gruppo omogeneo (mai individuali)."),
    ("SOGLIA MINIMA",
     "Gruppi con meno di 6 compilazioni non vengono elaborati "
     "(rischio di identificabilita' indiretta)."),
    ("CONSEGNA PROTETTA",
     "Questionari in busta chiusa, raccolti in urna sigillata. "
     "Nessun intermediario tra compilatore e RSPP."),
    ("INFORMATIVA PREVENTIVA",
     "Ogni lavoratore riceve informativa completa ai sensi degli "
     "artt. 13-14 GDPR prima della compilazione."),
    ("BASE GIURIDICA",
     "Art. 9 par.2 lett. b) GDPR: trattamento necessario per "
     "assolvere obblighi in materia di sicurezza sul lavoro."),
    ("CONSERVAZIONE LIMITATA",
     "I questionari cartacei sono distrutti dopo l'elaborazione. "
     "Il file Excel non contiene dati personali."),
]
y = Inches(1.6)
for title, desc in garanzie:
    add_rect(s, Inches(0.5), y, Inches(12.3), Inches(0.7), GR_CH, line=BLU)
    add_text(s, Inches(0.6), y, Inches(3.0), Inches(0.7),
             title, size=11, bold=True, color=BLU,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(3.7), y, Inches(9.0), Inches(0.7),
             desc, size=11, color=NERO, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.75)



# =============================================================
# SLIDE 18 - QUADRO NORMATIVO SINTETICO
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Quadro normativo di riferimento",
           "Le norme che governano il processo", 18)
norme = [
    ("D.Lgs. 81/2008 art. 28 c.1-bis",
     "Obbligo di valutazione SLC per tutti i lavoratori"),
    ("Accordo Europeo 8/10/2004",
     "Definizione condivisa di stress lavoro-correlato"),
    ("Circ. Min. Lavoro 18/11/2010",
     "Indicazioni metodologiche della Commissione Consultiva"),
    ("INAIL ed. 2017",
     "Metodologia validata: lista di controllo + Q-IND"),
    ("INAIL Monografia 2024/2025",
     "Modulo Smart Working + Technostress (27 item aggiuntivi)"),
    ("D.M. 284/2000 + D.P.R. 90/2010",
     "Applicazione al personale militare e civile della Difesa"),
    ("Reg. UE 679/2016 (GDPR)",
     "Tutela dati personali nei questionari"),
    ("Accordo Stato-Regioni 21/12/2011",
     "Formazione obbligatoria in materia di sicurezza"),
]
y = Inches(1.2)
for i, (norma, desc) in enumerate(norme):
    fill = BLU if i == 0 else (BIANCO if i % 2 else GR_CH)
    txt = BIANCO if i == 0 else NERO
    add_rect(s, Inches(0.5), y, Inches(4.5), Inches(0.6), fill, line=GRIGIO)
    add_rect(s, Inches(5.0), y, Inches(7.8), Inches(0.6), fill, line=GRIGIO)
    add_text(s, Inches(0.6), y, Inches(4.3), Inches(0.6),
             norma, size=11, bold=(i == 0), color=txt,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(5.1), y, Inches(7.6), Inches(0.6),
             desc, size=11, color=txt, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.65)

# =============================================================
# SLIDE 19 - PROSSIMI PASSI RICHIESTI AL DL
# =============================================================
s = prs.slides.add_slide(BLANK)
header_bar(s, "Prossimi passi - Decisioni richieste",
           "Azioni immediate per l'Ammiraglio Direttore", 19)
add_text(s, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.5),
         "RICHIESTE AL DATORE DI LAVORO - IN ORDINE DI PRIORITA'",
         size=18, bold=True, color=BLU)
richieste = [
    ("OGGI", "Approvare l'avvio del progetto SLC",
     "Disposizione scritta / Determina"),
    ("ENTRO 7 GG", "Nominare il GGV (Gruppo di Gestione)",
     "Ordine di servizio con composizione"),
    ("ENTRO 14 GG", "Firmare la comunicazione ai lavoratori",
     "Informativa predisposta dal RSPP"),
    ("ENTRO 14 GG", "Autorizzare le risorse",
     "Stampa questionari + eventuale consulente"),
    ("ENTRO 90 GG", "Firmare il DVR aggiornato",
     "Sezione SLC con data certa"),
]
y = Inches(1.8)
for quando, cosa, come in richieste:
    add_rect(s, Inches(0.5), y, Inches(2.0), Inches(0.9), BLU)
    add_text(s, Inches(0.55), y, Inches(1.9), Inches(0.9),
             quando, size=13, bold=True, color=BIANCO,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(2.6), y, Inches(5.5), Inches(0.9), GR_CH, line=BLU)
    add_text(s, Inches(2.7), y, Inches(5.3), Inches(0.9),
             cosa, size=13, bold=True, color=NERO,
             anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(8.2), y, Inches(4.6), Inches(0.9), GR_CH, line=GRIGIO)
    add_text(s, Inches(8.3), y, Inches(4.4), Inches(0.9),
             come, size=12, color=GRIGIO, anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.95)



# =============================================================
# SLIDE 20 - CONCLUSIONE
# =============================================================
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, SW, SH, BLU)
add_rect(s, 0, Inches(2.6), SW, Inches(0.08), ORO)
add_text(s, Inches(0.6), Inches(0.6), Inches(12), Inches(0.5),
         "MINISTERO DELLA DIFESA - MARINA MILITARE",
         size=14, bold=True, color=ORO)
add_text(s, Inches(0.6), Inches(2.9), Inches(12), Inches(1.0),
         "IN SINTESI", size=36, bold=True, color=BIANCO)
msgs = [
    "1. E' un OBBLIGO DI LEGGE non ancora adempiuto - rischio sanzione.",
    "2. Il SPP ha GIA' predisposto tutti gli strumenti operativi.",
    "3. Servono 90 GIORNI dall'approvazione per completare il processo.",
    "4. L'unica cosa che serve ORA e' la SUA FIRMA sull'avvio.",
    "5. Il risultato: un DVR completo e conforme, tutela per l'Ente "
    "e per Lei personalmente.",
]
add_bullets(s, Inches(0.6), Inches(4.0), Inches(12.1), Inches(2.5),
            msgs, size=16, color=BIANCO)
add_rect(s, 0, SH - Inches(0.6), SW, Inches(0.6), NERO)
add_text(s, Inches(0.6), SH - Inches(0.55), Inches(12), Inches(0.5),
         "Il RSPP resta a disposizione per ogni chiarimento.",
         size=14, bold=True, color=ORO,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# =============================================================
# Salvataggio
# =============================================================
out = "/projects/sandbox/ITALIC/INAIL_Stress_Lavoro_Correlato_Arsenali_Militari.pptx"
prs.save(out)
print(f"OK -> {out}")
print(f"Slide totali: {len(prs.slides)}")
