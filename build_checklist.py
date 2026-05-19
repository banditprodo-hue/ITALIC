"""
Generatore documento DOCX con CHECK-LIST OPERATIVE
per attivita' preliminari della Valutazione SLC INAIL.
Contesto: Arsenale Militare Marittimo - per RSPP.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BLU = RGBColor(0x0B, 0x2E, 0x55)
ORO = RGBColor(0xC9, 0xA2, 0x27)
GRIGIO = RGBColor(0x55, 0x55, 0x55)
NERO = RGBColor(0x1A, 0x1A, 0x1A)
ROSSO = RGBColor(0xB3, 0x1B, 0x1B)

doc = Document()

# Imposta margini
for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

# Stile base
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10)



def add_para(text, *, size=10, bold=False, color=NERO,
             align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4, italic=False):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    r.font.name = 'Calibri'
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return p


def shade_cell(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), hex_color)
    tc_pr.append(shd)


def set_cell_borders(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '6')
        b.set(qn('w:color'), '888888')
        borders.append(b)
    tc_pr.append(borders)


def add_checklist_header(codice, titolo, sottotitolo, riferimenti):
    """Aggiunge intestazione di una check-list."""
    # Titolo
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    shade_cell(cell, '0B2E55')
    set_cell_borders(cell)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(f"  {codice}  -  {titolo}")
    r.font.name = 'Calibri'
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    if sottotitolo:
        p2 = cell.add_paragraph()
        r2 = p2.add_run(f"  {sottotitolo}")
        r2.font.name = 'Calibri'
        r2.font.size = Pt(10)
        r2.font.color.rgb = ORO
    # riferimenti normativi
    add_para(f"Riferimento: {riferimenti}", size=9, italic=True,
             color=GRIGIO, space_after=8)



def add_id_block(area="", responsabile="RSPP",
                 tempistica="da definire"):
    """Blocco identificazione (Reparto, Data, RSPP, ecc.)"""
    tbl = doc.add_table(rows=2, cols=4)
    tbl.autofit = False
    headers = ["Reparto/Area:", "Data:", "Responsabile:", "Tempistica:"]
    values = [area, "____ / ____ / ________", responsabile, tempistica]
    for i, (h, v) in enumerate(zip(headers, values)):
        c1 = tbl.cell(0, i)
        c2 = tbl.cell(1, i)
        shade_cell(c1, 'ECECEC')
        set_cell_borders(c1)
        set_cell_borders(c2)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(h)
        r1.font.size = Pt(9)
        r1.font.bold = True
        r1.font.color.rgb = BLU
        p2 = c2.paragraphs[0]
        r2 = p2.add_run(v if v else "________________")
        r2.font.size = Pt(9)
    add_para("", size=4)


def add_check_items(items):
    """Aggiunge una lista di item con caselle barrabili."""
    for item in items:
        if isinstance(item, tuple):
            # sottocategoria
            subtitle, sub_items = item
            add_para(subtitle, size=10, bold=True, color=BLU,
                     space_after=2)
            for sub in sub_items:
                p = doc.add_paragraph()
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.left_indent = Cm(0.6)
                r1 = p.add_run("\u2610  ")
                r1.font.size = Pt(11)
                r2 = p.add_run(sub)
                r2.font.name = 'Calibri'
                r2.font.size = Pt(9.5)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.left_indent = Cm(0.3)
            r1 = p.add_run("\u2610  ")
            r1.font.size = Pt(11)
            r2 = p.add_run(item)
            r2.font.name = 'Calibri'
            r2.font.size = Pt(9.5)


def add_notes_box(rows=3):
    """Riquadro per note/osservazioni."""
    add_para("Note / Osservazioni / Criticita' rilevate:",
             size=9, bold=True, color=BLU, space_after=2)
    tbl = doc.add_table(rows=rows, cols=1)
    for i in range(rows):
        cell = tbl.cell(i, 0)
        set_cell_borders(cell)
        cell.paragraphs[0].add_run(" ")
    add_para("", size=4)


def add_signature_block():
    """Blocco firme."""
    tbl = doc.add_table(rows=2, cols=2)
    headers = ["RSPP - Firma", "Data"]
    for i, h in enumerate(headers):
        c1 = tbl.cell(0, i)
        c2 = tbl.cell(1, i)
        shade_cell(c1, 'ECECEC')
        set_cell_borders(c1)
        set_cell_borders(c2)
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(h)
        r1.font.size = Pt(9)
        r1.font.bold = True
        r1.font.color.rgb = BLU
        c2.paragraphs[0].add_run(" ")
    add_para("", size=4)


def page_break():
    doc.add_page_break()



# =====================================================
# COPERTINA
# =====================================================
add_para("MINISTERO DELLA DIFESA - MARINA MILITARE",
         size=11, bold=True, color=ORO,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
add_para("Servizio di Prevenzione e Protezione - Arsenale Militare",
         size=10, color=GRIGIO,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)

add_para("CHECK-LIST OPERATIVE", size=22, bold=True, color=BLU,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
add_para("Attivita' Preliminari della Valutazione",
         size=16, bold=True, color=BLU,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
add_para("Rischio Stress Lavoro-Correlato (SLC)",
         size=16, bold=True, color=BLU,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)

add_para("Strumento operativo per il RSPP", size=12, italic=True,
         color=GRIGIO, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
add_para("Conforme alle Linee Guida INAIL ed. 2017 + "
         "Monografia INAIL 2024/2025",
         size=10, color=GRIGIO,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24)

add_para("Riferimenti normativi:", size=10, bold=True, color=BLU,
         space_after=2)
add_para("- D.Lgs. 81/2008 - art. 28 c.1-bis (obbligo "
         "valutazione SLC)", size=9, space_after=1)
add_para("- D.Lgs. 81/2008 - art. 33 (compiti del SPP)",
         size=9, space_after=1)
add_para("- D.Lgs. 81/2008 - art. 50 (consultazione RLS)",
         size=9, space_after=1)
add_para("- Accordo Europeo 8 ottobre 2004 sullo SLC",
         size=9, space_after=1)
add_para("- Circ. Min. Lavoro prot. 23692 del 18/11/2010",
         size=9, space_after=1)
add_para("- Metodologia INAIL ed. 2017",
         size=9, space_after=1)
add_para("- Monografia INAIL 2024/2025 - Modulo Smart/ICT",
         size=9, space_after=1)
add_para("- Reg. UE 679/2016 (GDPR)", size=9, space_after=24)

# Indice
add_para("INDICE DELLE CHECK-LIST", size=14, bold=True, color=BLU,
         space_after=8)
indice = [
    ("CHK-PRE-01", "Verifica delle precondizioni"),
    ("CHK-PRE-02", "Costituzione del Gruppo di Gestione (GGV)"),
    ("CHK-PRE-03", "Definizione dei gruppi omogenei"),
    ("CHK-PRE-04", "Consultazione preventiva del RLS"),
    ("CHK-PRE-05", "Comunicazione ai lavoratori"),
    ("CHK-PRE-06", "Cronoprogramma e risorse"),
    ("CHK-PRE-07", "Area A - Raccolta eventi sentinella"),
    ("CHK-PRE-08", "Area B - Sopralluogo ambiente di lavoro"),
    ("CHK-PRE-09", "Area B - Contenuto del lavoro (organizzativo)"),
    ("CHK-PRE-10", "Area C - Contesto del lavoro"),
    ("CHK-PRE-11", "Compilazione foglio Excel e calcolo punteggi"),
    ("CHK-PRE-12", "Validazione preliminare e chiusura fase"),
]
for cod, tit in indice:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(f"{cod}  ")
    r1.font.size = Pt(10)
    r1.font.bold = True
    r1.font.color.rgb = BLU
    r2 = p.add_run(tit)
    r2.font.size = Pt(10)

page_break()



# =====================================================
# CHK-PRE-01 - VERIFICA PRECONDIZIONI
# =====================================================
add_checklist_header(
    "CHK-PRE-01",
    "Verifica delle precondizioni",
    "Documenti e dati che il RSPP deve avere PRIMA di iniziare",
    "D.Lgs. 81/08 art. 33 c.1 lett. a) - art. 28")
add_id_block(area="N/A (attivita' preliminare)",
             tempistica="Settimana 0 (prima dell'avvio)")

add_para("DOCUMENTI ISTITUZIONALI", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Atto di nomina RSPP firmato dal DL e in corso di validita'",
    "Atto di nomina del Medico Competente (MC) in corso di validita'",
    "Verbale elezione/designazione RLS aggiornato",
    "Organigramma dell'Arsenale aggiornato (max 12 mesi)",
    "Mansionario o repertorio delle mansioni",
    "DVR vigente (anche senza sezione SLC)",
    "Eventuale precedente valutazione SLC (se esistente)",
    "Procedura aziendale SLC (se gia' formalizzata)",
])

add_para("DATI STATISTICI (triennio precedente)", size=11, bold=True,
         color=BLU, space_after=4)
add_check_items([
    "Dati infortunistici per reparto (n. infortuni, gg, indici)",
    "Dati assenteismo per malattia (gg/lavoratore/anno)",
    "Dati turnover/trasferimenti per reparto",
    "Dati ferie non godute / ROL accumulati",
    "Numero sanzioni disciplinari per reparto",
    "Relazione annuale del MC (art. 25 c.1 lett. i)",
    "Numero richieste visita straordinaria al MC",
    "Eventuali segnalazioni/contenziosi pregressi (mobbing, "
    "molestie, ecc.)",
])

add_para("STRUMENTI OPERATIVI", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Questionario_SLC_INAIL_2025.docx disponibile",
    "Valutazione_SLC_INAIL_2025.xlsx disponibile",
    "Flowchart_SLC_Interattivo.html disponibile",
    "Computer/laptop per data entry",
    "Stampante e risme di carta (per stampa questionari)",
    "Buste cartacee (per consegna anonima)",
    "Urna sigillabile (per raccolta protetta)",
])

add_notes_box(rows=3)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-02 - COSTITUZIONE GGV
# =====================================================
add_checklist_header(
    "CHK-PRE-02",
    "Costituzione del Gruppo di Gestione della Valutazione (GGV)",
    "Atto formale di nomina e composizione del team",
    "D.Lgs. 81/08 art. 17 c.1 lett. a) - Metodologia INAIL")
add_id_block(tempistica="Settimana 1")

add_para("BOZZA DELL'ATTO DI NOMINA", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Bozza ordine di servizio/determina di nomina GGV redatta",
    "Oggetto chiaro (Avvio valutazione SLC ai sensi art. 28 c.1-bis)",
    "Riferimenti normativi richiamati",
    "Composizione completa indicata (nome, qualifica, ruolo nel GGV)",
    "Mandato del GGV definito (perimetro e durata)",
    "Eventuale presidente/coordinatore designato",
])

add_para("COMPOSIZIONE OBBLIGATORIA", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Datore di Lavoro o suo delegato (presiede)",
    "RSPP (coordinatore tecnico-metodologico)",
    "Eventuale ASPP",
    "Medico Competente",
    "RLS o RLST (consultato preventivamente, partecipa)",
])

add_para("COMPOSIZIONE INTEGRATIVA (raccomandata per Arsenali)",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Direttore di Stabilimento / suo delegato",
    "Capi Reparto / Officina dei reparti coinvolti",
    "Eventuale psicologo del lavoro (consulente esterno)",
    "Rappresentante Ufficio Personale",
])

add_para("FORMALIZZAZIONE", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Bozza sottoposta al DL per firma",
    "Atto firmato dal DL con data certa",
    "Atto numerato e protocollato",
    "Notifica via PEC/email a tutti i nominati",
    "Inserimento in fascicolo \"Sicurezza/SLC/Atti\"",
    "Convocazione prima riunione GGV (entro 7 gg)",
])

add_notes_box(rows=3)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-03 - GRUPPI OMOGENEI
# =====================================================
add_checklist_header(
    "CHK-PRE-03",
    "Definizione dei gruppi omogenei",
    "Suddivisione del personale ai fini della valutazione",
    "Metodologia INAIL ed. 2017 - Fase Propedeutica")
add_id_block(tempistica="Settimana 2-3")

add_para("MAPPATURA INIZIALE", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Estrazione organico aggiornato (totale lavoratori)",
    "Suddivisione per reparto/officina",
    "Suddivisione per qualifica (civili / militari)",
    "Suddivisione per livello (operai/impiegati/tecnici/dirigenti)",
    "Suddivisione per turno (giornalieri/turnisti H24/reperibili)",
    "Identificazione lavoratori in smart working/telelavoro",
    "Identificazione lavoratori con uso ICT intensivo",
])

add_para("CRITERI DI RAGGRUPPAMENTO", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Criterio principale scelto e motivato (mansione/reparto/turno)",
    "Tutti i lavoratori sono assegnati a un gruppo (no esclusioni)",
    "Verifica che ogni gruppo abbia almeno 6 componenti",
    "Per gruppi <6 lavoratori: aggregazione con gruppo simile",
    "Identificazione gruppi che richiedono PARTE 2 (Smart/ICT)",
    "Bilanciamento di genere considerato (se rilevante)",
])

add_para("VALIDAZIONE", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Tabella anagrafica gruppi compilata (foglio Excel - "
    "Gruppi_Omogenei)",
    "Per ogni gruppo: ID, reparto, mansione, turno, n.componenti, "
    "modalita'",
    "Validazione con i Capi Reparto interessati",
    "Validazione con il MC (per coerenza con sorveglianza sanitaria)",
    "Approvazione finale del GGV (verbalizzata)",
])

add_notes_box(rows=4)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-04 - CONSULTAZIONE RLS
# =====================================================
add_checklist_header(
    "CHK-PRE-04",
    "Consultazione preventiva del RLS",
    "Adempimento obbligatorio - tutela giuridica del processo",
    "D.Lgs. 81/08 art. 50 c.1 lett. b) e c)")
add_id_block(responsabile="DL su istruttoria RSPP",
             tempistica="Settimana 1 (PRIMA dell'avvio operativo)")

add_para("PREPARAZIONE", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Lettera di convocazione formale predisposta dal RSPP",
    "Convocazione firmata dal DL (o suo delegato)",
    "Trasmissione almeno 7 gg prima della riunione",
    "Allegata: bozza atto di costituzione GGV",
    "Allegata: bozza dei gruppi omogenei",
    "Allegata: bozza informativa ai lavoratori",
    "Allegata: cronoprogramma proposto",
    "Allegata: copia metodologia INAIL e questionario",
])

add_para("SVOLGIMENTO DELLA RIUNIONE", size=11, bold=True,
         color=BLU, space_after=4)
add_check_items([
    "Riunione svolta in luogo idoneo (riservato)",
    "Verbale di consultazione redatto",
    "Presa visione dei documenti registrata",
    "Eventuali osservazioni del RLS verbalizzate",
    "Eventuali proposte di modifica discusse",
    "Decisioni del DL motivate (se non recepite osservazioni RLS)",
    "Verbale firmato da DL, RSPP e RLS",
])

add_para("CONSEGUENZE OPERATIVE", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Aggiornamento documenti con osservazioni RLS recepite",
    "Comunicazione formale al RLS dell'avvio definitivo",
    "Coinvolgimento del RLS in tutte le fasi successive",
    "Diritto di accesso al DVR garantito (art. 50 c.1 lett. e)",
    "Inserimento del RLS come destinatario degli esiti aggregati",
])

add_para("ATTENZIONE", size=10, bold=True, color=ROSSO,
         space_after=2)
add_para("La mancata consultazione preventiva del RLS comporta "
         "sanzione amministrativa al DL (art. 55 c.5 lett. d) e "
         "puo' invalidare il procedimento valutativo.",
         size=9, italic=True, color=ROSSO, space_after=4)

add_notes_box(rows=3)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-05 - COMUNICAZIONE LAVORATORI
# =====================================================
add_checklist_header(
    "CHK-PRE-05",
    "Comunicazione ai lavoratori",
    "Informativa di avvio + informativa privacy GDPR",
    "D.Lgs. 81/08 art. 36 - Reg. UE 679/2016 artt. 13-14")
add_id_block(tempistica="Settimana 2")

add_para("CONTENUTI MINIMI DELL'INFORMATIVA", size=11, bold=True,
         color=BLU, space_after=4)
add_check_items([
    "Oggetto: avvio della valutazione SLC ai sensi art. 28 c.1-bis",
    "Quadro normativo richiamato",
    "Finalita' e benefici del processo",
    "Metodologia adottata (INAIL 2017 + Monografia 2024/2025)",
    "Articolazione in fasi (preliminare + eventuale approfondita)",
    "Composizione del GGV",
    "Cronoprogramma indicativo",
    "Garanzia di anonimato (compilazione anonima)",
    "Modalita' di restituzione degli esiti aggregati",
    "Indicazione referenti del SPP per chiarimenti",
])

add_para("INFORMATIVA GDPR (allegata)", size=11, bold=True,
         color=BLU, space_after=4)
add_check_items([
    "Titolare del trattamento identificato (Ministero Difesa)",
    "Finalita' del trattamento (obbligo legge)",
    "Base giuridica (art. 9 par.2 lett. b GDPR)",
    "Tipologia dati raccolti (NO dati identificativi)",
    "Durata della conservazione",
    "Diritti dell'interessato (artt. 15-22 GDPR)",
    "Contatti DPO/RPD",
])

add_para("CANALI DI DIFFUSIONE", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Affissione su bacheche di reparto",
    "Trasmissione via email a tutto il personale",
    "Invio circolare interna protocollata",
    "Briefing dei Capi Reparto al proprio personale",
    "Pubblicazione su intranet (se disponibile)",
    "Distribuzione cartacea ai lavoratori senza email",
])

add_para("DOCUMENTAZIONE", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Originale informativa firmato dal DL archiviato",
    "Copia in fascicolo \"Sicurezza/SLC/Comunicazioni\"",
    "Conferma di avvenuta diffusione (Capi Reparto)",
    "Ricezione di eventuali quesiti dei lavoratori (registrati)",
])

add_notes_box(rows=3)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-06 - CRONOPROGRAMMA
# =====================================================
add_checklist_header(
    "CHK-PRE-06",
    "Cronoprogramma e risorse",
    "Pianificazione tempistiche e approvvigionamenti",
    "Buona prassi gestionale - art. 17 D.Lgs. 81/08")
add_id_block(tempistica="Settimana 2")

add_para("PIANIFICAZIONE TEMPORALE", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "GANTT del progetto predisposto (90 giorni totali)",
    "Milestone intermedie definite",
    "Riunioni GGV calendarizzate (almeno 4: avvio, "
    "intermedia, validazione, chiusura)",
    "Date dei sopralluoghi nei reparti programmate",
    "Finestra di somministrazione questionari (4 settimane)",
    "Data di inserimento nel DVR fissata",
])

add_para("RISORSE UMANE", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Stima ore RSPP/ASPP necessarie",
    "Stima ore MC necessarie",
    "Stima ore RLS necessarie (consultazioni)",
    "Eventuale consulente esterno (psicologo del lavoro)",
    "Disponibilita' Ufficio Personale per estrazione dati",
    "Disponibilita' Capi Reparto per accompagnamento sopralluoghi",
])

add_para("RISORSE MATERIALI E LOGISTICHE", size=11, bold=True,
         color=BLU, space_after=4)
add_check_items([
    "Stampa questionari: stimate copie x lavoratore",
    "Buste cartacee per consegna anonima",
    "Urne sigillate (1 per reparto principale)",
    "Sale riunioni prenotate per GGV e focus group",
    "Computer per data entry Excel",
    "Spazio archivistico per conservazione documenti",
])

add_para("STRUMENTI DI MISURA (per Area B - ambiente)",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Fonometro disponibile (anche di tipo 2)",
    "Termo-igrometro disponibile",
    "Luxmetro disponibile",
    "Macchina fotografica autorizzata",
    "Tablet/smartphone per registrazione note in campo",
])

add_para("AUTORIZZAZIONI", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Approvazione formale del cronoprogramma da parte del DL",
    "Autorizzazione spese (se eventuali consulenti esterni)",
    "Autorizzazione fotografica nelle aree militari sensibili",
    "Autorizzazione accesso a tutti i reparti per sopralluoghi",
])

add_notes_box(rows=3)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-07 - AREA A EVENTI SENTINELLA
# =====================================================
add_checklist_header(
    "CHK-PRE-07",
    "Area A - Raccolta Eventi Sentinella",
    "Dati oggettivi triennali per ogni gruppo omogeneo",
    "Metodologia INAIL ed. 2017 - Lista di Controllo Area A")
add_id_block(tempistica="Settimana 3-4")

add_para("RICHIESTA DATI A UFFICIO PERSONALE",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Lettera di richiesta dati formalizzata",
    "Triennio di riferimento esplicitato (es. 2022-2023-2024)",
    "Estrazione per gruppo omogeneo (NON per singolo lavoratore)",
    "Dati richiesti: assenze malattia (gg totali per gruppo/anno)",
    "Dati richiesti: ferie non godute / ROL accumulati",
    "Dati richiesti: turnover (n. ingressi/uscite per gruppo)",
    "Dati richiesti: trasferimenti (n. richieste e attuati)",
    "Dati richiesti: sanzioni disciplinari (n. per gruppo)",
])

add_para("RICHIESTA DATI A SPP / INAIL", size=11, bold=True,
         color=BLU, space_after=4)
add_check_items([
    "Estrazione dati infortunistici dal registro infortuni",
    "Indici elaborati per gruppo (frequenza, gravita')",
    "Dati di near-miss (se gestiti)",
    "Stima del fenomeno sommerso (sottosegnalazione)",
])

add_para("RICHIESTA DATI AL MEDICO COMPETENTE",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Relazione annuale del MC ricevuta (art. 25 c.1 lett. i)",
    "Numero visite di sorveglianza sanitaria per gruppo",
    "Numero visite straordinarie a richiesta lavoratore "
    "(art. 41 c.2 lett. c)",
    "Numero giudizi di non idoneita' / idoneita' parziale",
    "Numero segnalazioni MC su sospetto SLC",
    "Trend rispetto agli anni precedenti",
])

add_para("ELABORAZIONE", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Tabella dati per gruppo omogeneo compilata",
    "Calcolo variazioni triennali (% di aumento/diminuzione)",
    "Confronto con benchmark INAIL (se disponibili)",
    "Identificazione dei gruppi con \"campanelli d'allarme\"",
    "Punteggi Area A inseriti nel foglio Excel "
    "(0/1/4 secondo Metodologia INAIL)",
    "Verifica congruita' con il MC",
])

add_para("ATTENZIONE PRIVACY", size=10, bold=True, color=ROSSO,
         space_after=2)
add_para("I dati sanitari del MC devono essere SEMPRE in forma "
         "AGGREGATA. Mai dati nominativi nel verbale. "
         "Conservare separatamente le richieste e le risposte.",
         size=9, italic=True, color=ROSSO, space_after=4)

add_notes_box(rows=3)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-08 - AREA B SOPRALLUOGO AMBIENTE (la piu' richiesta)
# =====================================================
add_checklist_header(
    "CHK-PRE-08",
    "Area B - Sopralluogo Ambiente di Lavoro",
    "Osservazione diretta dei reparti per Area B (Contenuto)",
    "Metodologia INAIL ed. 2017 - art. 28 c.2 lett. a")
add_id_block(tempistica="Settimana 3-5 (un sopralluogo per reparto)")

add_para("MICROCLIMA", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Temperatura rilevata e registrata (estiva/invernale)",
    "Umidita' relativa rilevata",
    "Ventilazione presente e funzionante (naturale/forzata)",
    "Correnti d'aria fastidiose presenti",
    "Differenze termiche tra zone del reparto (gradiente)",
    "Aree con temperatura percepita estrema (caldo/freddo)",
])

add_para("ILLUMINAZIONE", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Lux misurati nelle postazioni di lavoro",
    "Conformita' UNI EN 12464-1 verificata per la mansione",
    "Presenza di abbagliamenti (sole, faretti, riflessi)",
    "Aree buie o poco illuminate identificate",
    "Stato di pulizia dei corpi illuminanti",
    "Funzionamento dell'illuminazione di emergenza",
])

add_para("RUMORE", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Misurazione fonometrica indicativa effettuata",
    "Confronto con valutazione rumore vigente",
    "Presenza di rumori di fondo continui",
    "Presenza di picchi rumorosi (smerigliatrici, "
    "martellature, fischi)",
    "Lavoratori che indossano i DPI uditivi correttamente",
    "Adeguatezza dei DPI rispetto al livello di esposizione",
])

add_para("VIBRAZIONI", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Utensili a vibrazione mano-braccio presenti (smerigliatrici, "
    "scalpellatori)",
    "Tempi di esposizione monitorati (timer)",
    "Mezzi di sollevamento generatori di vibrazioni corpo intero",
    "Sedili dei mezzi adeguatamente ammortizzati",
    "Manutenzione antivibrante effettuata",
])

add_para("QUALITA' DELL'ARIA", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Fumi di saldatura aspirati alla fonte",
    "Polveri (sabbiatura, molatura) controllate",
    "Solventi e vernici stoccati correttamente",
    "Aspiratori funzionanti e manutenzionati",
    "Odori sgradevoli o irritanti percepiti",
    "Carriche di gas in spazi confinati segnalate (CO, H2S, "
    "ossigeno)",
])


add_para("SPAZI E LAYOUT", size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Larghezza percorsi pedonali adeguata (>= 1,2 m)",
    "Separazione percorsi pedoni/mezzi presente e segnalata",
    "Spazi di lavoro non sovraffollati",
    "Distanza tra postazioni adeguata",
    "Vie di esodo sgombre e visibili",
    "Aree di stoccaggio materiale ordinate",
    "Pavimenti puliti e privi di ostacoli/scivoli",
])

add_para("ATTREZZATURE E DPI", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Stato manutentivo macchine/attrezzature buono",
    "Schede macchina aggiornate (libretto manutenzione)",
    "Guasti ricorrenti segnalati (intervista al preposto)",
    "Disponibilita' ricambi adeguata (tempi sostituzione)",
    "DPI presenti, in buono stato e indossati",
    "Disponibilita' DPI in tutte le taglie",
    "Sostituzione DPI tempestiva (registro consegne)",
    "Strumentazione di misura tarata e disponibile",
])

add_para("SERVIZI E AREE COMUNI", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Spogliatoi puliti e dimensionati",
    "Spogliatoi separati civili/militari (coerente con normativa)",
    "Servizi igienici puliti e funzionanti (n. adeguato)",
    "Mensa accessibile, code gestite",
    "Aree pausa decompressione presenti e dignitose",
    "Distributori automatici funzionanti",
    "Punti di erogazione acqua potabile fresca disponibili",
])

add_para("CONDIZIONI OPERATIVE PARTICOLARI", size=11, bold=True,
         color=BLU, space_after=4)
add_check_items([
    "Spazi confinati: procedure D.P.R. 177/2011 affisse",
    "Lavori in quota: ponteggi conformi e ispezionati",
    "Lavori a caldo: permessi di lavoro presenti",
    "Esplosivi/munizioni: protocolli D.M. 272/99 visibili",
    "Lavori in mare/banchina: procedure specifiche disponibili",
    "Compresenza ditte appaltatrici: PSC/DUVRI consultati",
])

add_para("SEGNALI ORGANIZZATIVI VISIBILI", size=11, bold=True,
         color=BLU, space_after=4)
add_check_items([
    "Cartellonistica di sicurezza aggiornata",
    "Bacheca RLS con avvisi recenti",
    "Pulizia generale del reparto",
    "Atmosfera percepita (positiva/tesa/frustrata)",
    "Conversazioni informali tra lavoratori (clima)",
    "Disponibilita' del preposto a rispondere alle domande",
    "Reazione dei lavoratori al sopralluogo (apertura/chiusura)",
])

add_notes_box(rows=5)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-09 - AREA B CONTENUTO LAVORO ORGANIZZATIVO
# =====================================================
add_checklist_header(
    "CHK-PRE-09",
    "Area B - Contenuto del Lavoro (organizzativo)",
    "Aspetti organizzativi non ambientali del lavoro",
    "Metodologia INAIL ed. 2017 - Lista di Controllo Area B")
add_id_block(tempistica="Settimana 4")

add_para("PIANIFICAZIONE DEI COMPITI", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Esiste mansionario/job description per ogni mansione",
    "Compiti assegnati corrispondono alle competenze possedute",
    "Lavoratori sono adeguatamente formati per le proprie mansioni",
    "Compiti monotoni/ripetitivi identificati e gestiti (rotazione)",
    "Compiti complessi distribuiti equamente",
    "Possibilita' di sviluppo professionale presente",
])

add_para("CARICO DI LAVORO E RITMI", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Carico di lavoro misurabile (commesse, ore, output)",
    "Distribuzione del carico equa nel gruppo omogeneo",
    "Ritmi di lavoro determinati dalle macchine (se applicabile)",
    "Scadenze rigide e ricorrenti (es. consegna unita' navale)",
    "Picchi di lavoro stagionali/eventi e loro gestione",
    "Pause programmate e rispettate",
    "Possibilita' di gestione autonoma dei tempi",
])

add_para("ORARIO DI LAVORO", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Turni di lavoro analizzati (durata, alternanza)",
    "Turno notturno presente (rispetto D.Lgs. 66/2003)",
    "Riposo compensativo correttamente concesso",
    "Reperibilita': frequenza, durata, rotazione",
    "Straordinari: media mensile per gruppo",
    "Compatibilita' con vita familiare",
    "Flessibilita' oraria offerta (se compatibile col servizio)",
])

add_para("AUTONOMIA OPERATIVA", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Margini di scelta nelle modalita' di lavoro",
    "Possibilita' di partecipare a decisioni operative",
    "Procedure rigide ma comprese e accettate",
    "Spazi di autonomia per problem solving",
    "Disponibilita' delle risorse necessarie a fare il lavoro",
])

add_notes_box(rows=4)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-10 - AREA C CONTESTO LAVORO
# =====================================================
add_checklist_header(
    "CHK-PRE-10",
    "Area C - Contesto del Lavoro",
    "Aspetti relazionali, culturali, gestionali",
    "Metodologia INAIL ed. 2017 - Lista di Controllo Area C")
add_id_block(tempistica="Settimana 4-5")

add_para("FUNZIONE E CULTURA ORGANIZZATIVA",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Obiettivi dell'Arsenale comunicati al personale",
    "Obiettivi del reparto/officina chiari e condivisi",
    "Sistema di comunicazione interna funzionante",
    "Coerenza tra dichiarato (regolamenti) e agito (prassi)",
    "Codice etico/disciplinare presente e diffuso",
    "Cultura della sicurezza percepita come autentica",
])

add_para("RUOLO NELL'ORGANIZZAZIONE",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Ruoli e responsabilita' definiti per scritto",
    "Doppia catena di comando (militare/tecnica) chiarita",
    "Conflitti di ruolo identificati (richieste contraddittorie)",
    "Ambiguita' di ruolo segnalate (compiti non chiari)",
    "Sovrapposizioni gerarchiche civili/militari gestite",
    "Responsabilita' per il lavoro altrui correttamente attribuite",
])

add_para("RAPPORTI CON I SUPERIORI",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Capi Reparto formati alla gestione del personale",
    "Disponibilita' dei capi all'ascolto",
    "Feedback strutturato sul lavoro (positivo/correttivo)",
    "Riconoscimento dell'impegno del personale",
    "Stile di leadership coerente nel reparto",
    "Modalita' di gestione dei conflitti efficace",
])

add_para("RAPPORTI CON I COLLEGHI",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Clima collaborativo nei gruppi di lavoro",
    "Assenza di conflittualita' interpersonale strutturale",
    "Supporto reciproco tra colleghi",
    "Integrazione tra personale civile e militare",
    "Integrazione tra personale di diverse anzianita'",
    "Eventuali segnalazioni di mobbing/molestie (canale dedicato)",
])

add_para("EVOLUZIONE DELLA CARRIERA",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Trasparenza dei percorsi di carriera (civili e militari)",
    "Sistema premiante equo e percepito tale",
    "Mobilita' interna gestita con equita'",
    "Procedure di valutazione della prestazione applicate",
    "Tipologie contrattuali precarie (se presenti) gestite",
])

add_para("INTERFACCIA CASA-LAVORO",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Conciliazione vita-lavoro promossa",
    "Pendolarismo dei lavoratori monitorato",
    "Mobilita' forzata militare e impatto familiare",
    "Smart working strutturato (se previsto)",
    "Diritto alla disconnessione rispettato (per ICT/smart)",
])

add_notes_box(rows=4)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-11 - COMPILAZIONE EXCEL
# =====================================================
add_checklist_header(
    "CHK-PRE-11",
    "Compilazione foglio Excel e calcolo punteggi",
    "Inserimento dati e validazione automatica",
    "Strumento operativo Valutazione_SLC_INAIL_2025.xlsx")
add_id_block(tempistica="Settimana 5-6")

add_para("FOGLIO 'CONFIGURAZIONE'", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Nome organizzazione completo inserito",
    "Sede / Unita' produttiva specificata",
    "ATECO / Settore inserito (Difesa)",
    "Numero totale dipendenti aggiornato",
    "Data della valutazione fissata",
    "Data prossima rivalutazione (24-36 mesi) fissata",
    "Nome RSPP, MC, RLS inseriti",
    "Eventuale consulente esterno indicato",
    "Soglie cut-off verificate (default: 2.3 e 3.0)",
])

add_para("FOGLIO 'GRUPPI_OMOGENEI'", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Tutti i gruppi anagrafati (ID, reparto, mansione, turno)",
    "N. componenti coerente con organico",
    "Modalita' di lavoro indicata (Sede/Smart/ICT/Misto)",
    "Flag Smart/ICT attivato per gruppi che richiedono PARTE 2",
    "Note specifiche inserite (se necessarie)",
])

add_para("INSERIMENTO PUNTEGGI AREA A",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Per ogni gruppo: punteggi eventi sentinella inseriti",
    "Coerenza con i dati raccolti (CHK-PRE-07) verificata",
    "Calcolo automatico delle variazioni triennali OK",
    "Assenza di celle vuote",
])

add_para("INSERIMENTO PUNTEGGI AREA B",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Per ogni gruppo: punteggi ambiente di lavoro inseriti",
    "Per ogni gruppo: punteggi pianificazione compiti inseriti",
    "Per ogni gruppo: punteggi carico/ritmo inseriti",
    "Per ogni gruppo: punteggi orario di lavoro inseriti",
    "Coerenza con i sopralluoghi (CHK-PRE-08, CHK-PRE-09) verificata",
])

add_para("INSERIMENTO PUNTEGGI AREA C",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Per ogni gruppo: punteggi cultura organizzativa inseriti",
    "Per ogni gruppo: punteggi ruolo inseriti",
    "Per ogni gruppo: punteggi carriera inseriti",
    "Per ogni gruppo: punteggi rapporti inseriti",
    "Coerenza con CHK-PRE-10 verificata",
])

add_para("VERIFICA AUTOMATICA E REPORT",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Foglio 'Calcolo_Punteggi' aggiornato automaticamente",
    "Foglio 'Aggregazione_Gruppi' verificato",
    "Foglio 'Classificazione_Rischio' visualizzato",
    "Fascia di rischio per ogni gruppo coerente con dati",
    "Foglio 'Report_Risultati' stampato",
    "Eventuali anomalie evidenziate ai responsabili",
])

add_notes_box(rows=3)
add_signature_block()
page_break()



# =====================================================
# CHK-PRE-12 - VALIDAZIONE E CHIUSURA FASE PRELIMINARE
# =====================================================
add_checklist_header(
    "CHK-PRE-12",
    "Validazione preliminare e chiusura fase",
    "Verifica esiti, decisione, inserimento nel DVR",
    "D.Lgs. 81/08 art. 28-29 - Metodologia INAIL")
add_id_block(tempistica="Settimana 6-7")

add_para("RIUNIONE DI VALIDAZIONE GGV",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Convocazione riunione GGV per validazione esiti",
    "Presentazione esiti gruppo per gruppo",
    "Discussione delle anomalie/criticita' rilevate",
    "Coinvolgimento del MC per il parere sanitario",
    "Coinvolgimento del RLS per il punto di vista lavoratori",
    "Verbale di validazione redatto",
    "Eventuale revisione dei punteggi formalizzata",
])

add_para("DECISIONE OPERATIVA", size=11, bold=True, color=BLU,
         space_after=4)
add_check_items([
    "Per gruppi VERDE (rischio trascurabile): chiusura processo",
    "Per gruppi GIALLO (rischio moderato): definizione misure "
    "correttive",
    "Per gruppi ROSSO (rischio rilevante): avvio Fase Approfondita",
    "Decisione formalizzata in verbale",
    "Cronoprogramma misure correttive predisposto",
])

add_para("REDAZIONE SEZIONE SLC DEL DVR",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Premessa metodologica e normativa redatta",
    "Composizione GGV documentata",
    "Elenco dei gruppi omogenei valutati",
    "Esiti Area A/B/C riportati per ogni gruppo",
    "Fascia di rischio finale per ogni gruppo",
    "Piano misure correttive integrato",
    "Date di rivalutazione fissate",
    "Allegati: questionari (in bianco), foglio Excel report, "
    "verbali GGV",
])

add_para("APPROVAZIONE E ARCHIVIAZIONE",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Sezione SLC integrata nel DVR",
    "DVR aggiornato firmato dal DL con data certa",
    "Firme di RSPP, MC e RLS apposte (per ricezione)",
    "DVR comunicato a RLS (art. 50 c.5)",
    "DVR archiviato nella sede principale (art. 29 c.4)",
    "Copia digitale archiviata in modo sicuro",
    "Inserimento nel registro dei documenti SPP",
])

add_para("RESTITUZIONE AGLI STAKEHOLDER",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Comunicazione al DL degli esiti finali",
    "Comunicazione al RLS con consegna esiti aggregati",
    "Restituzione ai lavoratori (forma aggregata) tramite "
    "Capi Reparto",
    "Affissione sintesi sulle bacheche",
    "Eventuale comunicazione alle Organizzazioni Sindacali",
])

add_para("MONITORAGGIO POST-VALUTAZIONE",
         size=11, bold=True, color=BLU, space_after=4)
add_check_items([
    "Calendario di follow-up sulle misure correttive predisposto",
    "Indicatori di efficacia definiti",
    "Responsabili dell'attuazione misure assegnati",
    "Alert di rivalutazione automatica programmato (Excel)",
    "Eventuale Fase Approfondita pianificata (se necessaria)",
])

add_notes_box(rows=3)
add_signature_block()

# =====================================================
# Salvataggio
# =====================================================
out = "/projects/sandbox/ITALIC/CheckList_Operative_SLC_RSPP.docx"
doc.save(out)
print(f"OK -> {out}")
print(f"Documento generato con {len(doc.paragraphs)} paragrafi.")
