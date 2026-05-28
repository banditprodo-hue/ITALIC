"""
Genera il documento Word della Procedura FR 23/67 Rev. 2 (2026) in stile
"Master 2013" - bianco e nero, Times New Roman, intestazioni classiche,
tabelle con bordi pieni, paginazione.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "Procedura_Sicurezza_Macchine_Attrezzature_Rev2_2026.docx"

doc = Document()

# ---------------- Stili globali ----------------
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(11)
style.font.color.rgb = RGBColor(0, 0, 0)

# Margini sezione
for section in doc.sections:
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.0)
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)



# ---------------- Helper ----------------
def set_cell_borders(cell, color="000000", sz="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for border_name in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), sz)
        b.set(qn('w:color'), color)
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def set_cell_shading(cell, fill_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tc_pr.append(shd)


def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar'); fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.text = "PAGE"
    fld_sep = OxmlElement('w:fldChar'); fld_sep.set(qn('w:fldCharType'), 'separate')
    fld_end = OxmlElement('w:fldChar'); fld_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_sep); run._r.append(fld_end)
    paragraph.add_run(" di ")
    run2 = paragraph.add_run()
    fld_begin2 = OxmlElement('w:fldChar'); fld_begin2.set(qn('w:fldCharType'), 'begin')
    instr2 = OxmlElement('w:instrText'); instr2.text = "NUMPAGES"
    fld_sep2 = OxmlElement('w:fldChar'); fld_sep2.set(qn('w:fldCharType'), 'separate')
    fld_end2 = OxmlElement('w:fldChar'); fld_end2.set(qn('w:fldCharType'), 'end')
    run2._r.append(fld_begin2); run2._r.append(instr2); run2._r.append(fld_sep2); run2._r.append(fld_end2)


def H1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text.upper())
    r.bold = True; r.font.size = Pt(13); r.font.name = 'Times New Roman'


def H2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True; r.font.size = Pt(12); r.font.name = 'Times New Roman'


def H3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True; r.italic = True; r.font.size = Pt(11)


def P(text, bold=False, italic=False, justify=True, indent_first=True):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.space_after = Pt(4)
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if indent_first:
        p.paragraph_format.first_line_indent = Cm(0.6)
    r = p.add_run(text)
    r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    r.bold = bold; r.italic = italic
    return p


def BULLET(text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.runs[0] if p.runs else p.add_run("")
    p.text = ""  # reset
    r = p.add_run(text)
    r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    return p


def NUM(text):
    p = doc.add_paragraph(style='List Number')
    p.paragraph_format.space_after = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.text = ""
    r = p.add_run(text)
    r.font.name = 'Times New Roman'; r.font.size = Pt(11)
    return p


def NOTE(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("► " + text)
    r.italic = True; r.font.size = Pt(10.5)
    return p


def make_table(headers, rows, col_widths=None, header_shade='D9D9D9'):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    # header
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); r.bold = True; r.font.size = Pt(10.5); r.font.name = 'Times New Roman'
        set_cell_borders(cell)
        set_cell_shading(cell, header_shade)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # rows
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            r = p.add_run(val); r.font.size = Pt(10.5); r.font.name = 'Times New Roman'
            set_cell_borders(cell)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    if col_widths:
        for j, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[j].width = w
    return table


def add_page_break():
    doc.add_page_break()



# ---------------- Header & Footer ----------------
section = doc.sections[0]
header = section.header
ph = header.paragraphs[0]
ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
rh = ph.add_run("ARSENALE MARINA MILITARE – TARANTO\nUFFICIO PREVENZIONE E PROTEZIONE")
rh.bold = True; rh.font.size = Pt(9); rh.font.name = 'Times New Roman'

footer = section.footer
pf = footer.paragraphs[0]
pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_page_number(pf)
for r in pf.runs:
    r.font.size = Pt(9); r.font.name = 'Times New Roman'

# ---------------- FRONTESPIZIO ----------------
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_before = Pt(60)
title.paragraph_format.space_after = Pt(24)
r = title.add_run("PROCEDURA DI SICUREZZA\n«USO DI MACCHINE ED ATTREZZATURE»")
r.bold = True; r.font.size = Pt(20); r.font.name = 'Times New Roman'

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.paragraph_format.space_after = Pt(40)
r = sub.add_run("FR 23/67 – Revisione 2ª – Anno 2026")
r.bold = True; r.font.size = Pt(14); r.font.name = 'Times New Roman'

# Tabella dati documento
hdr_data = [
    ("Codice procedura", "FR 23/67"),
    ("Revisione", "2ª Revisione"),
    ("Data emissione", "Aprile 2026"),
    ("Sostituisce", "1ª Emissione del 15/04/2013 (prot. UPP/3/0133)"),
    ("Approvata da", "Capo Ufficio Prevenzione e Protezione"),
    ("Emessa da", "Nucleo Attrezzature e Impianti"),
    ("Trasmissione", "Posta elettronica certificata interna"),
    ("Riferimento", "D.Lgs. 81/08 e s.m.i."),
]
tab = doc.add_table(rows=len(hdr_data), cols=2)
tab.alignment = WD_TABLE_ALIGNMENT.CENTER
tab.autofit = False
for i, (k, v) in enumerate(hdr_data):
    c1 = tab.rows[i].cells[0]; c2 = tab.rows[i].cells[1]
    c1.text = ''; c2.text = ''
    p1 = c1.paragraphs[0]; p2 = c2.paragraphs[0]
    r1 = p1.add_run(k); r1.bold = True; r1.font.size = Pt(11); r1.font.name = 'Times New Roman'
    r2 = p2.add_run(v); r2.font.size = Pt(11); r2.font.name = 'Times New Roman'
    set_cell_borders(c1); set_cell_borders(c2)
    set_cell_shading(c1, 'F2F2F2')
    c1.width = Cm(5.5); c2.width = Cm(10.0)

add_page_break()

# ---------------- DESTINATARI ----------------
H1("Destinatari")
P("La presente procedura è trasmessa, per gli adempimenti di rispettiva competenza, ai seguenti destinatari:", indent_first=False)

destinatari = [
    "Direttore Direzione Supporto Diretto",
    "Capo Divisione Sistema Nave",
    "Capo Divisione Sistemi di Combattimento",
    "Capo Divisione Servizi Arsenale",
    "Capo Ufficio Consulente Giuridico",
    "Capo Ufficio Assicurazione Qualità",
    "Capo Ufficio Personale",
    "Capo Ufficio Servizi Tecnici",
    "Capo Ufficio Materiale",
]
for d in destinatari:
    BULLET(d)

P("E, per conoscenza:", bold=True, indent_first=False)
pc = [
    "Direttore Arsenale",
    "Vice Direttore Arsenale",
    "Direttore Lavori e Servizi",
    "Direttore Direzione Amministrazione",
    "Medico Competente",
    "Rappresentanti dei Lavoratori per la Sicurezza (RR.LL.SS.) — personale civile e militare",
]
for d in pc:
    BULLET(d)

add_page_break()



# =============== 1. OBIETTIVO ===============
H1("1. Obiettivo")
P("Scopo della presente procedura è definire le modalità di corretta gestione di macchine ed "
  "attrezzature di lavoro in dotazione all'Arsenale Marina Militare di Taranto, comprensiva delle "
  "attività di controllo periodico, manutenzione preventiva e correttiva, verifica di conformità e "
  "gestione degli interventi in sicurezza, incluse le procedure di isolamento delle fonti di energia "
  "(LOTO – Lockout/Tagout).")
P("La presente procedura richiama necessariamente i manuali e i libretti di istruzione consegnati "
  "all'atto di acquisto delle macchine e delle attrezzature, nonché la documentazione tecnica "
  "aggiornata prodotta dal fabbricante.")

# =============== 2. CAMPO DI APPLICAZIONE ===============
H1("2. Campo di applicazione")
P("La procedura si applica a tutte le attività svolte nei reparti, officine, bacini, bordo navi, "
  "laboratori e cantieri dell'Arsenale in cui vengano impiegate attrezzature di lavoro, così come "
  "definite dall'art. 69 del D.Lgs. 81/08 e s.m.i.")

# =============== 3. RIFERIMENTI NORMATIVI ===============
H1("3. Riferimenti normativi")
make_table(
    ["Norma / Fonte", "Oggetto"],
    [
        ("D.Lgs. 9 aprile 2008, n. 81 e s.m.i.",
         "Testo Unico Sicurezza sul Lavoro – Titolo III (Uso delle attrezzature di lavoro e dei DPI), "
         "Allegato V (Requisiti di sicurezza), Allegato VI (Disposizioni d'uso), Allegato VIII (DPI)."),
        ("Direttiva 2006/42/CE",
         "Direttiva Macchine – requisiti essenziali di sicurezza e salute (recepita con D.Lgs. 17/2010, "
         "oggi coordinata nel corpo del D.Lgs. 81/08)."),
        ("Regolamento (UE) 2023/1230",
         "Nuovo Regolamento Macchine – sostituirà la Direttiva 2006/42/CE dal 14 gennaio 2027. "
         "Le disposizioni in materia di cybersecurity dei sistemi di controllo e di IA devono essere "
         "considerate fin da subito nelle nuove acquisizioni."),
        ("Norme tecniche UNI EN ISO in vigore",
         "UNI EN ISO 12100 (principi generali di progettazione); UNI EN ISO 13849-1/-2 (sistemi di "
         "comando legati alla sicurezza); UNI EN ISO 14119 (interblocchi); UNI EN ISO 13850 (arresto "
         "di emergenza); UNI EN ISO 13857 (distanze di sicurezza)."),
        ("Norme CEI EN 60204-1 / CEI EN 62061",
         "Equipaggiamento elettrico delle macchine e sicurezza funzionale dei sistemi di controllo."),
        ("UNI EN ISO 14118",
         "Prevenzione dell'avviamento inatteso – procedure di Lockout/Tagout (LOTO)."),
        ("Reg. (UE) 2016/425 e D.Lgs. 17/2019",
         "Dispositivi di Protezione Individuale (DPI) – categorie I, II, III."),
        ("Accordo Stato-Regioni in vigore",
         "Formazione dei lavoratori, dei preposti e abilitazione all'uso di attrezzature particolari "
         "(art. 73, c. 5, D.Lgs. 81/08)."),
    ],
    col_widths=[Cm(5.5), Cm(10.5)],
)
NOTE("Sono da ritenersi abrogate e superate tutte le previgenti disposizioni richiamate in precedenti "
     "edizioni della presente procedura (DPR 547/55, DPR 303/56, DPR 459/96 quale normativa autonoma). "
     "Il quadro di riferimento è esclusivamente il D.Lgs. 81/08 e s.m.i. coordinato con la normativa "
     "europea sopra citata.")

add_page_break()



# =============== 4. DEFINIZIONI ===============
H1("4. Definizioni")
make_table(
    ["Termine", "Definizione"],
    [
        ("Datore di Lavoro",
         "Soggetto titolare del rapporto di lavoro o, comunque, soggetto responsabile "
         "dell'organizzazione nel cui ambito il lavoratore presta la propria attività "
         "(art. 2, c. 1, lett. b, D.Lgs. 81/08)."),
        ("Dirigente",
         "Persona che, in ragione delle competenze professionali e dei poteri gerarchici e funzionali, "
         "attua le direttive del Datore di Lavoro organizzando l'attività lavorativa e vigilando su "
         "di essa (art. 2, c. 1, lett. d)."),
        ("Preposto",
         "Persona che, in ragione delle competenze professionali e nei limiti dei poteri gerarchici e "
         "funzionali, sovrintende all'attività lavorativa, garantendo l'attuazione delle direttive "
         "ricevute, controllandone l'esecuzione ed esercitando funzionale potere d'iniziativa "
         "(art. 2, c. 1, lett. e)."),
        ("Attrezzatura di lavoro",
         "Qualsiasi macchina, apparecchio, utensile o impianto destinato ad essere usato durante il "
         "lavoro (art. 69, c. 1, lett. a)."),
        ("Uso di una attrezzatura di lavoro",
         "Qualsiasi operazione lavorativa connessa all'attrezzatura: messa in servizio o fuori "
         "servizio, impiego, trasporto, riparazione, trasformazione, manutenzione, pulizia, "
         "montaggio, smontaggio (art. 69, c. 1, lett. b)."),
        ("Zona pericolosa",
         "Qualsiasi zona all'interno o in prossimità di un'attrezzatura nella quale la presenza di "
         "un lavoratore costituisce un rischio per la salute o la sicurezza (art. 69, c. 1, lett. c)."),
        ("Lavoratore esposto",
         "Lavoratore che si trovi interamente o in parte in una zona pericolosa."),
        ("Operatore",
         "Lavoratore incaricato dell'uso di un'attrezzatura, in possesso di formazione specifica e "
         "addestramento documentato."),
        ("LOTO (Lockout/Tagout)",
         "Procedura formalizzata di isolamento, blocco e segnalazione delle fonti di energia "
         "(elettrica, pneumatica, idraulica, meccanica, termica, chimica, gravitazionale) prima di "
         "interventi di manutenzione o pulizia, finalizzata a garantire lo stato di «energia zero»."),
    ],
    col_widths=[Cm(4.5), Cm(11.5)],
)

add_page_break()

# =============== 5. RESPONSABILITÀ ===============
H1("5. Responsabilità")

H2("5.1 Dirigente")
P("Il Dirigente deve assicurare che:", indent_first=False)
NUM("Le attrezzature messe a disposizione dei lavoratori siano conformi alle specifiche disposizioni "
    "legislative e regolamentari di recepimento delle Direttive comunitarie di prodotto "
    "(art. 70, c. 1, D.Lgs. 81/08).")
NUM("Le attrezzature costruite in assenza di norme di recepimento, ovvero messe a disposizione "
    "antecedentemente alla loro emanazione, siano conformi ai requisiti generali di sicurezza di cui "
    "all'Allegato V del D.Lgs. 81/08.")
NUM("Sia effettuato e documentato il controllo periodico sull'applicazione della presente procedura "
    "presso tutti i reparti di competenza.")
NUM("Le nuove acquisizioni rispettino i requisiti del Regolamento (UE) 2023/1230 ove applicabile.")

H2("5.2 Preposto")
P("Il Preposto deve garantire:", indent_first=False)
BULLET("L'applicazione puntuale della presente procedura, vigilando sul rispetto da parte dei "
       "lavoratori delle disposizioni in materia di tutela della sicurezza e salute.")
BULLET("La fornitura al personale di una chiara definizione della corretta sistemazione dell'area di "
       "lavoro.")
BULLET("La verifica dell'efficienza e dell'idoneità delle attrezzature in dotazione al personale, "
       "con particolare attenzione ai dispositivi di protezione e di arresto di emergenza.")
BULLET("La presa visione obbligatoria della presente procedura da parte di tutto il personale "
       "assegnato, con firma per ricevuta.")
BULLET("La segnalazione immediata al Dirigente di qualsiasi non conformità riscontrata.")

H2("5.3 Lavoratori e operatori")
P("I lavoratori sono tenuti a:", indent_first=False)
BULLET("Possedere adeguata conoscenza degli aspetti esecutivi del lavoro relativi all'utilizzo delle "
       "attrezzature, avendo completato il percorso di formazione specifica e addestramento "
       "documentato di cui all'art. 73 del D.Lgs. 81/08.")
BULLET("Utilizzare correttamente le attrezzature mantenendole in perfetta efficienza, nel rispetto "
       "delle istruzioni del fabbricante.")
BULLET("Arrestare immediatamente il lavoro e segnalare senza indugio al Preposto qualsiasi anomalia, "
       "difetto o situazione di pericolo.")
BULLET("Astenersi dall'apportare modifiche di propria iniziativa alle attrezzature, ai dispositivi "
       "di sicurezza e ai ripari.")
BULLET("Prendere visione della presente procedura e sottoscriverla per ricevuta.")

H2("5.4 Nucleo Attrezzature e Impianti")
P("Il Nucleo Attrezzature garantisce:", indent_first=False)
BULLET("La pianificazione, esecuzione e registrazione della manutenzione preventiva programmata "
       "secondo i piani del fabbricante e le indicazioni del Servizio di Prevenzione e Protezione.")
BULLET("La tenuta del registro attrezzature, con aggiornamento dello stato di conformità, delle "
       "verifiche periodiche e delle certificazioni.")
BULLET("La gestione del magazzino ricambi originali o equivalenti certificati.")
BULLET("L'esecuzione delle procedure LOTO in occasione di ogni intervento manutentivo.")

H2("5.5 Servizio di Prevenzione e Protezione (UPP)")
P("Il Servizio di Prevenzione e Protezione deve essere preventivamente informato dell'acquisto di "
  "nuove attrezzature, della modifica di quelle esistenti e dell'installazione di nuovi impianti, "
  "ai fini della valutazione dei nuovi rischi introdotti.")

add_page_break()



# =============== 6. PREMESSA NORMATIVA ===============
H1("6. Premessa normativa – quadro legislativo vigente")

H2("6.1 Il Titolo III del D.Lgs. 81/08")
P("Il Titolo III del D.Lgs. 81/08 e s.m.i. regolamenta l'uso delle attrezzature di lavoro e sancisce "
  "l'obbligo per il Datore di Lavoro di:", indent_first=False)
BULLET("Mettere a disposizione dei lavoratori attrezzature adeguate al lavoro da svolgere ovvero "
       "adattate a tale scopo.")
BULLET("Attuare le misure tecniche e organizzative adeguate per ridurre al minimo i rischi connessi "
       "all'uso delle attrezzature.")
BULLET("Assicurare che le attrezzature siano installate e utilizzate in conformità alle istruzioni "
       "del fabbricante e oggetto di idonea manutenzione programmata.")
BULLET("Provvedere affinché i lavoratori incaricati ricevano formazione e addestramento adeguati e "
       "documentati, in rapporto alla sicurezza, alle condizioni di impiego, alle situazioni anormali "
       "prevedibili e alle procedure di emergenza.")
BULLET("Riservare l'uso di attrezzature che richiedono conoscenze o responsabilità particolari "
       "esclusivamente ai lavoratori allo scopo incaricati e qualificati.")

H2("6.2 Obblighi dei lavoratori")
BULLET("Sottoporsi ai programmi di formazione e addestramento organizzati dal Datore di Lavoro.")
BULLET("Utilizzare le attrezzature conformemente alla formazione, informazione e addestramento "
       "ricevuti.")
BULLET("Avere cura delle attrezzature messe a loro disposizione, senza apportare modifiche di "
       "propria iniziativa.")
BULLET("Segnalare immediatamente ai superiori gerarchici qualsiasi difetto, anomalia o "
       "inconveniente rilevato.")

H2("6.3 Direttiva Macchine e Regolamento (UE) 2023/1230")
P("La Direttiva 2006/42/CE — recepita in Italia con D.Lgs. 17/2010, oggi coordinata nel corpo del "
  "D.Lgs. 81/08 — si applica alle macchine nelle diverse fasi che vanno dalla progettazione alla "
  "fabbricazione, immissione sul mercato, vendita e messa in servizio presso l'utente finale.")
P("Il Regolamento (UE) 2023/1230 relativo alle macchine e ai prodotti correlati, pubblicato il "
  "29 giugno 2023, sostituirà integralmente la Direttiva 2006/42/CE a partire dal 14 gennaio 2027. "
  "Le disposizioni in esso contenute — in particolare quelle relative ai requisiti essenziali di "
  "sicurezza per macchine con sistemi digitali, cybersecurity dei sistemi di controllo e "
  "adeguamenti per macchine con componenti di intelligenza artificiale — devono essere considerate "
  "fin da subito in fase di nuove acquisizioni e di progettazione di modifiche sostanziali.")
P("Il fabbricante è tenuto a:", indent_first=False)
NUM("Costituire il Fascicolo Tecnico della macchina.")
NUM("Predisporre il Manuale di istruzioni per l'uso e la manutenzione in lingua italiana.")
NUM("Redigere la Dichiarazione di conformità UE ai requisiti essenziali di sicurezza.")
NUM("Apporre la Marcatura CE sulla macchina.")

H2("6.4 Macchine preesistenti (costruite prima del 21/09/1996)")
P("Le macchine usate costruite prima del 21 settembre 1996 non ricadono nel regime della Direttiva "
  "Macchine. La legislazione di riferimento per tali attrezzature è rappresentata dall'Allegato V "
  "del D.Lgs. 81/08. Il Datore di Lavoro ha l'obbligo di:", indent_first=False)
NUM("Verificare che la macchina risponda ai requisiti generali di sicurezza dell'Allegato V.")
NUM("Individuare e attuare gli adeguamenti necessari.")
NUM("Verificare la presenza e l'efficacia dell'arresto di emergenza conforme.")
NUM("Aggiornare e conservare il relativo libretto di manutenzione.")

H2("6.5 Modifiche a macchine esistenti")
P("Le macchine usate che abbiano subito modifiche sostanziali in qualsiasi momento della loro vita "
  "ricadono nel regime della Direttiva Macchine (e del futuro Regolamento UE 2023/1230). Chi "
  "effettua la modifica ne diventa giuridicamente il fabbricante e deve eseguire l'intero iter di "
  "marcatura CE.")
P("Per modifica sostanziale si intende l'insieme degli interventi che esulano dalla manutenzione "
  "ordinaria o straordinaria, finalizzati ad adeguare la produzione o i sistemi di sicurezza a "
  "nuove esigenze o tecnologie, alterando significativamente le caratteristiche descritte nel "
  "fascicolo tecnico originale.")

add_page_break()



# =============== 7. ACQUISTO MACCHINA ===============
H1("7. Avvertenze per l'acquisto di una macchina")

H2("7.1 Premessa")
P("È obbligatorio includere nel contratto/ordine d'acquisto di una macchina, tra le altre, le "
  "seguenti clausole:", indent_first=False)
NUM("Saldo di una quota (15–25%) del prezzo subordinato all'installazione completata e al collaudo "
    "avvenuto con esito positivo, previa eliminazione di eventuali vizi palesi, con particolare "
    "riferimento alla conformità ai requisiti di sicurezza.")
NUM("Garanzia scritta del pieno rispetto delle prestazioni di funzionamento dichiarate dal "
    "fabbricante.")
NUM("Consegna obbligatoria, contestuale alla macchina, di tutta la documentazione tecnica in lingua "
    "italiana (Fascicolo Tecnico, Manuale di istruzioni, Dichiarazione di conformità UE).")
NOTE("ATTENZIONE: l'acquirente che non ottenga la documentazione prescritta al momento della "
     "consegna è obbligato a provvedere in proprio al perfezionamento di quanto non fatto dal "
     "fabbricante o venditore inadempiente, prima della messa in servizio dell'attrezzatura.")

H2("7.2 Macchina nuova (costruita dopo il 21/09/1996)")
P("All'atto della consegna, la macchina deve essere dotata da parte del fabbricante di:", indent_first=False)
BULLET("Marcatura CE conforme.")
BULLET("Dichiarazione di conformità UE.")
BULLET("Manuale di istruzione (installazione, uso, manutenzione, trasporto, smaltimento) in lingua "
       "italiana.")
BULLET("Fascicolo tecnico (conservato dal fabbricante, consultabile su richiesta).")
P("Quanto sopra deve risultare anche dalla documentazione di consegna (DDT/bolla).")

H2("7.3 Macchina usata (costruita prima del 21/09/1996)")
P("Il venditore è tenuto a rilasciare dichiarazione scritta attestante che la macchina è conforme, "
  "al momento della consegna, ai requisiti dell'Allegato V del D.Lgs. 81/08. La data di costruzione "
  "deve essere attestata per iscritto. Anche la macchina usata deve essere dotata di adeguato "
  "manuale di istruzione in lingua italiana.")

add_page_break()

# =============== 8. PRINCIPALI PERICOLI ===============
H1("8. Principali pericoli delle macchine")
P("La valutazione della pericolosità di una macchina, ai sensi dell'Allegato V del D.Lgs. 81/08 e "
  "della norma UNI EN ISO 12100, non deve basarsi esclusivamente su indagini storiche degli "
  "infortuni: non è consentito definire innocua una macchina per la sola assenza di dati "
  "infortunistici pregressi. Il degrado degli impianti, la manutenzione non corretta o non "
  "tempestiva, l'obsolescenza, la familiarità acquisita nel tempo, la mutata preparazione "
  "professionale del personale e lo sviluppo tecnologico hanno un peso determinante sulla "
  "pericolosità residua.")

H2("8.1 Pericoli di natura meccanica")
for v in ["Schiacciamento", "Cesoiamento", "Taglio o sezionamento", "Impigliamento",
          "Trascinamento o intrappolamento", "Urto", "Perforazione o puntura",
          "Attrito o abrasione", "Proiezione di fluido ad alta pressione",
          "Proiezione di parti della macchina o materiali/pezzi lavorati",
          "Perdita di stabilità della macchina o di sue parti",
          "Scivolamento, inciampo e caduta in relazione alla macchina"]:
    BULLET(v)

H2("8.2 Pericoli di natura elettrica")
for v in ["Lesioni o morte per contatti diretti e indiretti con elementi in tensione (CEI EN 60204-1).",
          "Fenomeni elettrostatici.",
          "Influenze esterne sugli equipaggiamenti elettrici, in particolare sui circuiti di sicurezza.",
          "Spruzzi metallici da cortocircuiti.",
          "Radiazioni termiche e altri fenomeni correlati."]:
    BULLET(v)

H2("8.3 Pericoli di natura termica")
BULLET("Bruciature e scottature per contatto con elementi in temperatura, irraggiamento, fiamme o "
       "esplosioni.")
BULLET("Danni alla salute provocati dall'alterazione delle condizioni microclimatiche dei luoghi di "
       "lavoro.")

H2("8.4 Pericoli generati da rumore")
for v in ["Perdita permanente dell'acuità uditiva (ipoacusia professionale).",
          "Acufeni.",
          "Stanchezza, tensione, stress lavoro-correlato.",
          "Interferenze con la comunicazione verbale e con i segnali acustici di sicurezza."]:
    BULLET(v)
NOTE("Per macchine caratterizzate da elevata rumorosità (presse, torni, fresatrici, molatrici) è "
     "obbligatorio l'utilizzo di otoprotettori conformi e selezionati sulla base della valutazione "
     "del rischio rumore (Titolo VIII, Capo II, D.Lgs. 81/08).")

H2("8.5 Pericoli generati da vibrazioni")
for v in ["Disturbi vascolari (sindrome da vibrazioni mano-braccio – HAV).",
          "Disturbi neurologici.",
          "Disturbi osteo-articolari (patologie del rachide per vibrazioni al corpo intero – WBV)."]:
    BULLET(v)

H2("8.6 Pericoli generati da radiazioni (non ionizzanti e ionizzanti)")
for v in ["Archi elettrici nei processi di saldatura.",
          "Apparecchiature laser su macchine utensili CNC.",
          "Sorgenti di radiazioni ionizzanti in alcune tipologie di macchine per controlli non distruttivi."]:
    BULLET(v)

H2("8.7 Pericoli generati da materiali o sostanze")
for v in ["Contatto o inalazione di fluidi, gas, nebbie, fumi e polveri.",
          "Pericoli biologici (muffe) e microbiologici (virus, batteri).",
          "Pericoli d'incendio o d'esplosione."]:
    BULLET(v)

H2("8.8 Pericoli ergonomici")
for v in ["Posizioni errate o sforzi eccessivi e ripetitivi.",
          "Inadeguatezza dell'anatomia umana rispetto ai comandi.",
          "Inadeguatezza dell'illuminazione locale.",
          "Eccessivo o scarso impegno mentale.",
          "Mancato uso dei DPI.",
          "Errori umani."]:
    BULLET(v)

H2("8.9 Pericoli da guasti e malfunzionamenti")
for v in ["Guasti all'alimentazione di energia.",
          "Proiezione di parti di macchina o fluidi per rottura.",
          "Errori di montaggio.",
          "Riavviamento intempestivo dopo interruzione e ripristino dell'alimentazione."]:
    BULLET(v)

add_page_break()



# =============== 9. NORME GENERALI DI PROTEZIONE ===============
H1("9. Norme generali di protezione delle macchine")
P("Tutti gli organi che possono costituire pericolo (pulegge, cinghie, cremagliere, ingranaggi, "
  "parti sporgenti, alberi di trasmissione) devono essere muniti di protezioni, segregati o "
  "provvisti di idonei dispositivi di sicurezza conformi alle norme UNI EN ISO in vigore, per "
  "evitare qualsiasi possibilità di afferramento, urto, cesoiamento o contatto con gli operatori.")
P("In funzione del tipo di rischio individuato nella valutazione, deve essere predisposta una "
  "protezione di tipo fisso, mobile, regolabile o immateriale, nel rispetto della gerarchia delle "
  "misure di protezione stabilita dalla UNI EN ISO 12100.")
P("I ripari devono essere di costruzione robusta, non facilmente eludibili o disattivabili, non "
  "devono limitare la visibilità necessaria alla conduzione del ciclo lavorativo e devono essere "
  "posizionati a distanza idonea dalla zona pericolosa, calcolata secondo la UNI EN ISO 13857.")

H2("9.1 Protezioni fisse")
P("Provvedono all'isolamento permanente di singole parti delle macchine, degli organi in movimento "
  "o di aree di lavoro. Sono fissate alla struttura della macchina con viti o bulloni e richiedono "
  "l'uso di utensili per la rimozione. Possono essere rimosse esclusivamente a macchina ferma e "
  "con procedura LOTO attivata.")

H2("9.2 Ripari mobili con interblocco")
P("Conformi alla UNI EN ISO 14119, sono dispositivi interconnessi ai comandi della macchina tali che:", indent_first=False)
NUM("Finché la protezione non è inserita, la macchina non può operare.")
NUM("Finché il movimento non è cessato completamente, la protezione non può essere sbloccata.")
NUM("Disinserendo la protezione durante il funzionamento, la macchina si arresta immediatamente.")

H2("9.3 Protezioni regolabili")
P("Nel caso in cui sia tecnicamente impossibile applicare una protezione fissa, l'accesso alla "
  "zona pericolosa deve essere impedito con una protezione regolabile, progettata in modo da "
  "ridurre al minimo il rischio residuo.")

# =============== 10. DISPOSITIVI DI COMANDO ===============
H1("10. Dispositivi di comando")
P("I dispositivi di comando devono essere sicuri, affidabili e progettati secondo i principi della "
  "norma UNI EN ISO 13849-1 (livello di prestazione PL richiesto dalla valutazione del rischio).")
P("Dal posto di comando deve essere possibile la verifica dell'assenza di personale nella zona "
  "pericolosa; se ciò non è possibile, l'avviamento deve essere preceduto da segnale di "
  "avvertimento sonoro e/o visivo con tempo sufficiente all'allontanamento. Tutti i dispositivi di "
  "azionamento devono essere protetti contro gli azionamenti accidentali.")

H3("10.1 Arresto normale")
P("Consente l'arresto di tutti gli elementi mobili della macchina, ponendoli in condizione di "
  "sicurezza. Deve avere priorità rispetto al comando di avviamento.")

H3("10.2 Arresto di emergenza")
P("Conforme alla UNI EN ISO 13850. Deve essere chiaramente individuabile (pulsante a fungo rosso "
  "su fondo giallo, ben visibile e immediatamente accessibile) e provocare l'arresto nel più breve "
  "tempo possibile. Lo sblocco avviene solo mediante manovra intenzionale dell'operatore "
  "qualificato e autorizza la rimessa in funzione ma non il riavvio automatico della macchina.")

H3("10.3 Comando ad azione mantenuta")
P("Il funzionamento è consentito solo finché il comando è mantenuto attivato dall'operatore. Al "
  "rilascio, la macchina si arresta immediatamente.")

H3("10.4 Comando a due mani")
P("Conforme alla UNI EN 574, tipo IIIC. L'avvio è consentito solo se entrambi i comandi sono "
  "azionati simultaneamente. Il rilascio anche temporaneo di uno dei due comandi pone la macchina "
  "in posizione di sicurezza.")

H3("10.5 Dispositivo a pedale")
P("Deve essere protetto sopra e ai lati da una custodia che impedisca ogni azionamento "
  "accidentale. Il suo utilizzo è subordinato a valutazione del rischio specifica.")

H3("10.6 Dispositivi sensibili")
BULLET("Azionati meccanicamente: dispositivi a fune, tappeti e bordi sensibili alla pressione.")
BULLET("Azionati non meccanicamente: barriere fotoelettriche (ESPE/AOPD conformi UNI EN 61496), "
       "dispositivi a ultrasuoni, scanner laser di sicurezza.")

H3("10.7 Dispositivi di interblocco")
P("Conformi alla UNI EN ISO 14119, sono utilizzati con i ripari mobili secondo i criteri del "
  "punto 9.2.")

add_page_break()



# =============== 11. OBBLIGHI OPERATORE ===============
H1("11. Obblighi dell'operatore alle macchine")
P("L'operatore addetto alla conduzione di macchine e attrezzature deve:", indent_first=False)
NUM("Disattivare l'attrezzatura ogni qualvolta sospenda la lavorazione, anche per brevi periodi, in "
    "modo che non possa essere attivata accidentalmente da terzi.")
NUM("Effettuare la manutenzione ordinaria e straordinaria solo se in possesso di formazione e "
    "addestramento documentato; in caso contrario gli interventi devono essere affidati a personale "
    "specializzato interno (Nucleo Attrezzature) o a ditte esterne qualificate sotto contratto.")
NUM("Utilizzare i Dispositivi di Protezione Individuale (DPI) prescritti dal fabbricante, dalla "
    "valutazione dei rischi e dalla presente procedura.")
NUM("Sospendere immediatamente l'attività e informare il Preposto qualora riscontri difetti o "
    "anomalie nel funzionamento.")
NUM("Astenersi dall'apportare modifiche di propria iniziativa alle attrezzature, ai ripari o ai "
    "dispositivi di sicurezza.")
NUM("Non rimuovere, disattivare o eludere in alcun modo le protezioni e i dispositivi di sicurezza "
    "installati.")

# =============== 12. MODALITÀ OPERATIVE ===============
H1("12. Modalità operative generali – uso di macchine e attrezzature")

H2("12.1 Disposizioni generali")
BULLET("L'uso di macchine e attrezzature è riservato esclusivamente al personale appositamente "
       "incaricato, in possesso di formazione specifica e addestramento documentato.")
BULLET("Le macchine devono essere utilizzate unicamente secondo le indicazioni e le condizioni "
       "operative previste dal fabbricante nel Manuale di istruzioni.")
BULLET("Gli elementi delle macchine, quando costituiscono pericolo, devono essere protetti, "
       "segregati o provvisti di dispositivi di sicurezza conformi.")

H2("12.2 Divieti assoluti")
P("È fatto assoluto divieto di:", indent_first=False)
BULLET("Rimuovere, anche temporaneamente, le protezioni e i dispositivi di sicurezza, salvo le "
       "deroghe di cui al punto 12.3.")
BULLET("Pulire, oliare o ingrassare a mano gli organi e gli elementi in moto delle macchine.")
BULLET("Compiere su organi in movimento qualsiasi operazione di manutenzione, riparazione o "
       "registrazione.")
BULLET("Utilizzare aria compressa per la pulizia del corpo o degli indumenti dell'operatore.")
BULLET("Operare su macchine in assenza dei ripari e delle protezioni previste.")

H2("12.3 Rimozione temporanea delle protezioni")
P("Qualora, per improrogabili necessità tecniche debitamente documentate, le protezioni debbano "
  "essere rimosse:", indent_first=False)
NUM("Deve essere immediatamente redatto un Permesso di Lavoro specifico, autorizzato dal Preposto.")
NUM("Devono essere adottate misure supplementari atte a ridurre al minimo il pericolo residuo "
    "(riduzione della velocità, sorveglianza continua, uso di comandi ad azione mantenuta).")
NUM("La rimessa in posto della protezione deve avvenire immediatamente al cessare della necessità.")
NUM("L'evento deve essere registrato nel registro di manutenzione della macchina.")

H2("12.4 Manutenzione")

H3("12.4.1 Manutenzione preventiva programmata (ordinaria)")
P("Insieme delle operazioni di manutenzione pianificate dal fabbricante e indicate nel Manuale di "
  "istruzioni, che l'utilizzatore è obbligato a eseguire secondo le scadenze stabilite. Il Nucleo "
  "Attrezzature tiene traccia di ogni intervento nel registro di manutenzione.")

H3("12.4.2 Manutenzione correttiva (straordinaria)")
P("Intervento tecnico che l'utilizzatore deve eseguire allorquando si manifestino anomalie, "
  "decadimento delle prestazioni o una generale vetustà. In tali casi qualunque sostituzione o "
  "ripristino di organi o pezzi deve avvenire nel rispetto delle scelte progettuali originali, con "
  "elementi strettamente equivalenti o superiori, e deve essere documentato.")

H3("12.4.3 Procedura LOTO (Lockout/Tagout)")
P("È fatto obbligo assoluto di applicare la procedura LOTO prima di qualsiasi intervento di "
  "manutenzione, riparazione, pulizia o ispezione che comporti l'accesso a zone pericolose della "
  "macchina:", indent_first=False)
NUM("Identificare tutte le fonti di energia (elettrica, pneumatica, idraulica, meccanica, termica, "
    "gravitazionale).")
NUM("Informare tutto il personale coinvolto dell'inizio della procedura di isolamento.")
NUM("Sezionare tutte le fonti di energia mediante i dispositivi di sezionamento previsti.")
NUM("Applicare i lucchetti personali (lockout) e le etichette di segnalazione (tagout) su ogni "
    "dispositivo di sezionamento.")
NUM("Dissipare le energie residue (scaricare accumulatori pneumatici/idraulici, mettere a terra "
    "condensatori, bloccare meccanicamente parti soggette a gravità).")
NUM("Verificare l'effettivo stato di «energia zero» tentando l'avviamento della macchina (prova di "
    "riavviamento negativa).")
NUM("Eseguire l'intervento manutentivo.")
NUM("Ripristinare le protezioni e rimuovere lucchetti e tag solo dopo aver verificato che nessun "
    "lavoratore si trovi nella zona pericolosa.")
NUM("Riavviare la macchina secondo la procedura prevista dal fabbricante.")
NOTE("Ogni violazione della procedura LOTO costituisce grave infrazione disciplinare e può "
     "configurare responsabilità penali ai sensi degli artt. 20 e 59 del D.Lgs. 81/08.")

add_page_break()

# =============== 13. FORMAZIONE ===============
H1("13. Formazione, informazione e addestramento")
P("Ai sensi dell'art. 73 del D.Lgs. 81/08 e s.m.i. e dell'Accordo Stato-Regioni vigente:", indent_first=False)
NUM("Tutti i lavoratori incaricati dell'uso di attrezzature di lavoro devono ricevere una "
    "formazione specifica sufficiente e adeguata e, ove necessario, uno specifico addestramento "
    "pratico.")
NUM("La formazione e l'addestramento devono riguardare le condizioni di impiego dell'attrezzatura, "
    "le situazioni anormali prevedibili, le procedure di emergenza, le procedure LOTO applicabili e "
    "l'uso corretto dei DPI associati.")
NUM("L'addestramento deve essere documentato con verbale sottoscritto dal lavoratore e dal "
    "soggetto formatore, conservato nel fascicolo personale di formazione.")
NUM("Per le attrezzature la cui conduzione richiede conoscenze e responsabilità particolari "
    "(art. 73, c. 5), l'abilitazione è subordinata al superamento di specifico percorso formativo.")
NUM("L'uso di DPI di III categoria (otoprotettori in ambienti ad elevata rumorosità, DPI "
    "anticaduta, protezioni delle vie respiratorie) richiede addestramento specifico obbligatorio, "
    "da ripetere periodicamente e da documentare.")
NUM("Il personale preposto deve ricevere formazione aggiuntiva ai sensi dell'art. 37, c. 7, e "
    "dell'Accordo Stato-Regioni, con aggiornamento periodico obbligatorio.")

add_page_break()



# =============== 14. SUGGERIMENTI OPERATIVI ===============
H1("14. Suggerimenti operativi per l'uso delle principali macchine utensili")

# 14.1 TORNIO
H2("14.1 Tornio")
P("Il tornio è una macchina utensile che opera per asportazione di truciolo; il moto di taglio è "
  "conferito al pezzo in lavorazione (moto rotatorio) mentre il moto di avanzamento è posseduto "
  "dall'utensile (moto traslatorio).")
H3("Rischi principali")
for v in [
    "Contatti accidentali con gli attrezzi di fissaggio (mandrino, menabrida, staffe, plateu) e "
    "con parti in movimento.",
    "Accesso alla zona di alloggiamento delle cinghie, pulegge o ingranaggi di trasmissione.",
    "Elettrocuzione.",
    "Inefficienza o assenza dei sistemi di arresto di emergenza.",
    "Mancata trattenuta del pezzo in lavorazione (proiezione del pezzo).",
    "Mancata protezione della zona di operazione dell'utensile.",
    "Rottura dell'utensile con proiezione di frammenti.",
    "Proiezione di trucioli lunghi trascinati dalla rotazione del pezzo.",
    "Impigliamento di abiti, capelli, accessori personali."]:
    BULLET(v)
H3("Cautele obbligatorie")
for v in [
    "Bloccare il pezzo da lavorare in modo certo e sicuro mediante autocentrante o contropunta.",
    "Accertarsi che non vi siano chiavi o altri attrezzi sul mandrino prima della messa in moto.",
    "Chiudere sempre il riparo coprimandrino.",
    "Assicurarsi che lo schermo antiproiezione sia correttamente posizionato.",
    "Utilizzare un fioretto o uncino con paraschegge per asportare i trucioli; è fatto divieto di "
    "utilizzare le mani.",
    "Effettuare misurazioni del pezzo solo a macchina ferma e dopo aver allontanato l'utensile.",
    "È fatto assoluto divieto di fissare, registrare o misurare il pezzo durante la tornitura.",
    "È fatto divieto di effettuare pulizie con aria compressa.",
    "È fatto assoluto divieto di indossare indumenti svolazzanti, braccialetti, collane, anelli o "
    "altri accessori.",
    "Mai avvicinare le mani al pezzo in rotazione.",
    "Applicare la procedura LOTO prima di ogni intervento manutentivo."]:
    BULLET(v)

# 14.2 FRESATRICE
H2("14.2 Fresatrice")
P("La fresatrice opera per asportazione di truciolo mediante un utensile a taglienti multipli "
  "(fresa); il moto di avanzamento è posseduto dalla tavola portapezzo.")
H3("Rischi principali")
for v in [
    "Rottura dell'utensile con proiezione di frammenti.",
    "Mancata protezione della zona di lavorazione dell'utensile.",
    "Aggiustamento o misurazione con macchina in moto.",
    "Caduta di pezzi dalla tavola portapezzo.",
    "Pezzi sporgenti dalla sagoma della tavola durante la lavorazione.",
    "Mancata trattenuta del pezzo.",
    "Accesso alla zona di trasmissione del moto durante il funzionamento.",
    "Asportazione trucioli durante il moto.",
    "Elettrocuzione."]:
    BULLET(v)
H3("Cautele obbligatorie")
for v in [
    "Bloccare il pezzo in modo certo e sicuro sulla tavola portapezzo.",
    "Accertarsi che non vi siano chiavi o attrezzi sulla tavola durante la lavorazione.",
    "È fatto assoluto divieto di lavorare con la macchina sprovvista di dispositivi di sicurezza.",
    "È fatto divieto di fissare, registrare o misurare il pezzo durante la fresatura.",
    "È fatto divieto di eseguire operazioni di manutenzione e pulizia con organi in movimento.",
    "Non utilizzare aria compressa per la pulizia.",
    "È fatto assoluto divieto di indossare indumenti che possano impigliarsi.",
    "Verificare che lo schermo antiproiezione sia correttamente posizionato.",
    "Mai avvicinare le mani all'utensile in rotazione.",
    "Applicare la procedura LOTO prima di ogni intervento manutentivo."]:
    BULLET(v)

# 14.3 TRAPANO
H2("14.3 Trapano a colonna")
P("Il trapano è macchina utensile destinata all'esecuzione di fori mediante punta rotante.")
H3("Rischi principali")
for v in [
    "Rottura dell'utensile.",
    "Mancata protezione della zona di lavorazione.",
    "Elettrocuzione.",
    "Mancata trattenuta del pezzo (trascinamento in rotazione).",
    "Presenza di attrezzi o pezzi non in lavorazione sulla tavola.",
    "Aggiustamento o misurazione con macchina in moto.",
    "Proiezione di trucioli lunghi trascinati dalla punta.",
    "Impigliamento di abiti e capelli."]:
    BULLET(v)
H3("Cautele obbligatorie")
for v in [
    "È fatto assoluto divieto di indossare guanti, orologi, braccialetti, collane o qualsiasi "
    "elemento afferrabile dalla punta.",
    "Usare berretti, cuffie o retine raccoglicapelli se i capelli sono lunghi.",
    "Fissare sempre i pezzi in lavorazione, anche di piccole dimensioni, sulla tavola portapezzo "
    "con morse o staffe.",
    "È fatto divieto di fissare, registrare o misurare il pezzo durante la foratura.",
    "Limitare la lunghezza dei trucioli selezionando velocità di rotazione e avanzamento "
    "appropriati.",
    "Asportare trucioli e schegge con uncini muniti di schermo o spazzole metalliche; mai con le "
    "mani o l'aria compressa.",
    "Tenere la tavola sgombra da attrezzi o pezzi non in lavorazione.",
    "In caso di inceppamento della punta: fermare il trapano, estrarre la punta, controllarla "
    "prima di riprendere il lavoro.",
    "Applicare la procedura LOTO prima di ogni intervento manutentivo."]:
    BULLET(v)

# 14.4 MOLATRICE
H2("14.4 Molatrice")
H3("Rischi principali")
for v in [
    "Contatti accidentali con la mola in rotazione.",
    "Proiezione di materiali e polveri.",
    "Instabilità della macchina.",
    "Variazione della velocità di rotazione.",
    "Elettrocuzione.",
    "Rottura della mola con proiezione violenta di frammenti."]:
    BULLET(v)
H3("Cautele obbligatorie")
for v in [
    "È obbligatorio l'uso costante di occhiali di protezione o visiera integrale.",
    "Per uso prolungato e/o in ambienti chiusi: utilizzare cuffia antirumore, copricapo e maschera "
    "antipolvere.",
    "È fatto assoluto divieto di utilizzare mole abrasive a velocità superiore a quella garantita "
    "dal fabbricante.",
    "È fatto assoluto divieto di lavorare con la macchina sprovvista di dispositivi di sicurezza, "
    "cuffia di protezione o ripari.",
    "È fatto divieto di effettuare manutenzione con la macchina in movimento.",
    "Il montaggio della mola deve essere effettuato esclusivamente da persona competente, che "
    "verifica integrità, data di fabbricazione, velocità massima prescritta e serraggio delle flange.",
    "Verificare prima di ogni utilizzo che la distanza tra poggiapezzo e mola non superi i 2 mm.",
    "Applicare la procedura LOTO prima di ogni intervento manutentivo."]:
    BULLET(v)

add_page_break()



# 14.5 CESOIE E PRESSE
H2("14.5 Cesoie e presse")
H3("Rischi principali")
BULLET("Cesoiamento di arti (mani, dita).")
BULLET("Schiacciamento da parte del pressore o punzone.")
H3("Dispositivi di protezione obbligatori")
P("Devono essere adottati uno o più dei seguenti sistemi, conformi al livello di prestazione PL ≥ "
  "«d» ai sensi della UNI EN ISO 13849-1:", indent_first=False)
NUM("Barriera immateriale (ESPE/AOPD): fotocellule a più fasci conformi alla UNI EN 61496, "
    "collegate al sistema di comando con livello di sicurezza SIL 3 / PL «e».")
NUM("Comando a due mani simultaneo (UNI EN 574, tipo IIIC): a uomo presente. In tal caso alla "
    "pressa deve essere addetto un solo lavoratore; per più operatori contemporanei deve essere "
    "previsto un comando a due mani per ciascuno.")
NUM("Dispositivo antiripetitore del colpo: con elettrovalvole a doppio corpo ridondanti, su due "
    "circuiti elettrici alimentati separatamente e monitorati.")
P("Distanze di sicurezza per barriere immateriali (Allegato V, D.Lgs. 81/08):", indent_first=False)
make_table(
    ["Tempo di arresto totale (ms)*", "60", "75", "100", "250", "500"],
    [("Distanza minima di sicurezza (mm)", "100", "120", "160", "400", "800")],
    col_widths=[Cm(5.5), Cm(2.0), Cm(2.0), Cm(2.0), Cm(2.0), Cm(2.0)],
)
NOTE("* Tempo di arresto della pressa più tempo di risposta del sistema di comando.")
H3("Cautele obbligatorie")
for v in [
    "È fatto assoluto divieto di lavorare con la macchina sprovvista di dispositivi di sicurezza, "
    "ripari o di rimuovere/eludere gli stessi.",
    "È fatto assoluto divieto di operare in più di una persona quando la pressa è dotata di "
    "singolo comando a due mani.",
    "È fatto divieto di effettuare operazioni di pulizia o manutenzione con organi in movimento.",
    "Prima di ogni intervento nel campo di azione del punzone, applicare la procedura LOTO."]:
    BULLET(v)

# 14.6 LEGNO
H2("14.6 Macchine per la lavorazione del legno")
H3("Rischi principali")
for v in [
    "Proiezioni di schegge e frammenti.",
    "Inalazione di polvere di legno (agente cancerogeno per esposizione professionale – "
    "Titolo IX, Capo II, D.Lgs. 81/08).",
    "Contatto con utensili taglienti durante la manipolazione dei pezzi.",
    "Rigetto del pezzo (kickback)."]:
    BULLET(v)
H3("Cautele obbligatorie")
for v in [
    "Utilizzare sempre gli schermi previsti contro la proiezione di schegge.",
    "Non lasciare il pezzo impegnato sulla lama o sull'utensile in caso di interruzione; sfilarlo "
    "sempre.",
    "È fatto divieto di sfilare o spostare il pezzo durante il moto.",
    "Graduare la spinta del pezzo, prestando attenzione alla presenza di nodi.",
    "Azionare sempre i sistemi di aspirazione della segatura e dei trucioli (impianto di "
    "aspirazione localizzata, Allegato IV, punto 2, D.Lgs. 81/08).",
    "Per le seghe a nastro: i volantini di rinvio devono essere completamente protetti.",
    "Per le seghe circolari: cuffia registrabile, coltello divisore in acciaio (≤ 3 mm dalla "
    "dentatura), schermi laterali, spingitoi per pezzi piccoli o sottili.",
    "Applicare la procedura LOTO prima di ogni intervento manutentivo."]:
    BULLET(v)

# 14.7 SALDATURA E TAGLIO - GENERALI
H2("14.7 Saldatura e taglio – disposizioni generali")
P("Per la protezione dai rischi derivanti dai lavori di saldatura e taglio (esplosioni, fumi "
  "dannosi, incendi, radiazioni ottiche artificiali, ustioni) devono essere utilizzati:", indent_first=False)
BULLET("Schermi o maschere/caschi per saldatura con filtri idonei, conformi a UNI EN 175 e UNI EN "
       "169/170/171.")
BULLET("DPI completi: cuffia o copricapo ignifugo, guanti da saldatore (UNI EN 12477), grembiule "
       "in cuoio o tessuto ignifugo, ghette, calzature di sicurezza.")
H3("Divieti assoluti – non effettuare operazioni di saldatura")
for v in [
    "Su recipienti o tubi chiusi non preventivamente bonificati e certificati gas-free.",
    "Su recipienti o tubi contenenti residui che possano formare miscele esplosive sotto l'azione "
    "del calore (benzina, acetilene, nafta, olio, solventi).",
    "All'interno di locali, cunicoli o fosse non efficacemente ventilati: necessaria verifica "
    "dell'atmosfera con rilevatore multigas prima dell'ingresso."]:
    BULLET(v)
H3("Disposizioni operative durante la saldatura")
for v in [
    "Delimitare con schermi ignifughi i posti di saldatura per evitare abbagliamenti.",
    "Allontanare materiali combustibili; ove impossibile, proteggerli con schermi parascintille e "
    "predisporre estintori a portata di mano.",
    "Impedire che scintille o gocce incandescenti possano cadere su persone o materiali "
    "infiammabili sottostanti.",
    "Installare nei posti fissi un sistema di aspirazione localizzata dei fumi conforme alle "
    "norme UNI EN ISO 21904.",
    "Per saldature su materiali zincati, verniciati o contenenti piombo/cromo: aspirazione "
    "obbligatoria e DPI per le vie respiratorie (FFP3 o semimaschera con filtri A2P3)."]:
    BULLET(v)

# 14.8 OSSIACETILENICO
H2("14.8 Saldatura e taglio ossiacetilenico")
H3("Rischi specifici")
for v in [
    "Esposizione a fumi e sostanze nocive.",
    "Esposizione a radiazioni emesse dalla fiamma.",
    "Ustioni per proiezione di scorie incandescenti.",
    "Rischio di incendio o esplosione."]:
    BULLET(v)
H3("Cautele obbligatorie")
for v in [
    "Utilizzare sempre i DPI prescritti.",
    "Verificare prima dell'uso l'efficienza di manometri, riduttori, valvole di ritegno di fiamma, "
    "tubazioni e cannelli.",
    "Aprire valvole e rubinetti a mano o con chiave apposita; è fatto divieto di forzature.",
    "Per le fughe di gas: usare esclusivamente acqua saponata o prodotti specifici; mai fiamme.",
    "Usare fascette a vite per il fissaggio delle tubazioni.",
    "Distendere le tubazioni in curve ampie, lontano da passaggi e fonti di calore.",
    "Accendere i cannelli con accenditori piezoelettrici dedicati; mai con fiammiferi.",
    "Mantenere le bombole di acetilene in posizione verticale; prelievo orario mai superiore a "
    "1/5 della capacità.",
    "Estinguere la fiamma chiudendo prima la valvola dell'acetilene, poi quella dell'ossigeno.",
    "A fine lavoro: chiudere le valvole delle bombole, attendere il ritorno a zero dei manometri, "
    "allentare i riduttori."]:
    BULLET(v)
H3("Deposito e movimentazione bombole")
for v in [
    "Bombole contraddistinte con i colori EN 1089-3: bianco per ossigeno, marrone per acetilene.",
    "Valvola protetta da cappuccio metallico quando non è applicato il riduttore.",
    "È fatto divieto di esporre le bombole al sole o a fonti di calore.",
    "Depositi in locali non interrati, ventilati, con segnaletica di divieto di fumo e fiamme libere.",
    "Bombole di ossigeno e acetilene in locali separati; piene distinte da vuote; ancorate a parete.",
    "Movimentazione: esclusivamente con carrelli dedicati, senza urti o rotolamenti.",
    "È fatto assoluto divieto di contatto tra bombole/riduttori/raccordi e oli o grassi."]:
    BULLET(v)

# 14.9 SALDATURA ELETTRICA
H2("14.9 Saldatura elettrica (ad arco)")
H3("Rischi specifici")
for v in [
    "Folgorazione o disturbi da corrente elettrica.",
    "Esposizione al calore e alle radiazioni ottiche artificiali emesse dall'arco (UV, visibile, IR).",
    "Ustioni per contatto o proiezione di particelle incandescenti.",
    "Esposizione a fumi, gas e vapori generati dall'elettrodo e dal materiale base."]:
    BULLET(v)
H3("Cautele obbligatorie")
for v in [
    "Apparecchiature conformi alla CEI EN 60974, dotate di interruttore onnipolare, pinze "
    "portaelettrodi con impugnatura isolante e schermo a disco, cavo di massa in perfetto stato.",
    "DPI specifici obbligatori: casco/maschera con filtro auto-oscurante (UNI EN 379), guanti da "
    "saldatore (UNI EN 12477), grembiule in cuoio o tessuto ignifugo (UNI EN ISO 11611), ghette "
    "ignifughe, calzature S3.",
    "È fatto divieto di indossare oggetti metallici durante la saldatura.",
    "In ambienti scarsamente ventilati: aspirazione localizzata o ventilazione forzata.",
    "Prima di saldare: raschiare e pulire i pezzi verniciati, zincati o sporchi di olio e grasso.",
    "Proteggere il personale circostante con schermi opachi conformi."]:
    BULLET(v)

add_page_break()



# 14.10 UTENSILI MANUALI
H2("14.10 Uso di utensili manuali")
H3("Principali cause di infortunio")
for v in [
    "Impiego scorretto o inadeguato dell'utensile.",
    "Qualità scadente del materiale.",
    "Cattivo stato di manutenzione.",
    "Inadeguato trattamento termico superficiale.",
    "Parti taglienti o acuminate non protette.",
    "Proiezioni di schegge durante l'uso.",
    "Errori di mira."]:
    BULLET(v)
H3("Misure di prevenzione obbligatorie")
for v in [
    "Utilizzare esclusivamente attrezzi di qualità certificata e conforme alle norme UNI EN.",
    "Utilizzare attrezzi convenientemente temprati.",
    "Scegliere attrezzi con manici ergonomici.",
    "Per lavori su parti sotto tensione: solo attrezzi isolati conformi alla UNI EN 60900 (≤ 1000 V c.a.).",
    "In luoghi con pericolo di esplosione: solo attrezzi antiscintilla (lega Cu-Be o Cu-Al).",
    "Controllare sempre lo stato degli attrezzi prima dell'uso; sostituire o riparare i difettosi.",
    "Tenere puliti e ordinati gli attrezzi.",
    "Proteggere le parti pungenti o taglienti durante trasporto e deposito.",
    "È fatto divieto di portare nelle tasche attrezzi pungenti o taglienti.",
    "Nei lavori in quota: borse o cinture portautensili, assicurando contro la caduta."]:
    BULLET(v)

H3("Disposizioni specifiche per tipologia")
P("Scalpelli", bold=True, indent_first=False)
for v in [
    "Verificare affilatura e assenza di ricalcature (fungature).",
    "Utilizzare schermo paraschegge in presenza di terzi.",
    "Utilizzare porta-scalpello o proteggi-mano contro errori di mira.",
    "Scartare gli scalpelli che producono schegge.",
    "È obbligatorio l'uso di occhiali di protezione."]:
    BULLET(v)
P("Martelli", bold=True, indent_first=False)
for v in [
    "Manico con fibre parallele all'asse, superficie liscia, ben incastrato e assicurato con cuneo.",
    "Faccia e penna levigate, angoli convenientemente smussati.",
    "Movimento di battuta prevalentemente con l'articolazione del polso."]:
    BULLET(v)
P("Cacciavite", bold=True, indent_first=False)
for v in [
    "Verificare integrità e adeguatezza della lama all'intaglio della vite.",
    "È fatto divieto di tenere piccoli pezzi nel palmo della mano durante l'avvitamento; usare "
    "morsa o supporto fisso."]:
    BULLET(v)
P("Chiavi fisse e regolabili", bold=True, indent_first=False)
for v in [
    "L'apertura deve corrispondere esattamente al dado/bullone; scartare chiavi danneggiate.",
    "Tenere la chiave ad angolo retto rispetto all'asse della vite.",
    "È fatto divieto di utilizzare prolungamenti improvvisati."]:
    BULLET(v)
P("Seghetti", bold=True, indent_first=False)
for v in [
    "Lama ben fissata al telaio; pezzo bloccato in morsa.",
    "Iniziare il taglio con presa controllata, pollice lontano dalla dentatura."]:
    BULLET(v)
P("Lime", bold=True, indent_first=False)
BULLET("Manico con anello metallico e codolo correttamente inserito; è fatto divieto di utilizzare "
       "lime prive di manico.")

add_page_break()

# =============== 15. ORDINE E PULIZIA ===============
H1("15. Ordine e pulizia dei luoghi di lavoro")
P("Un laboratorio o un'officina mantenuti in perfetto ordine e pulizia prevengono una significativa "
  "quota di infortuni derivanti da cause occasionali (inciampi, scivolamenti, contatti accidentali "
  "con materiali taglienti o in temperatura).")
P("Il personale preposto deve assicurare che al termine di ogni turno di lavoro:", indent_first=False)
BULLET("Le postazioni siano sgombre da trucioli, sfridi, oli esausti e materiali di risulta.")
BULLET("Gli utensili siano riposti nei luoghi assegnati.")
BULLET("I DPI siano verificati, puliti e riposti correttamente.")
BULLET("I passaggi e le vie di fuga siano liberi da ostacoli.")
P("È fatto assoluto divieto di fumare durante le operazioni di lavoro e in tutti i locali in cui "
  "siano presenti sostanze infiammabili, esplosive o polverose.", bold=True)

add_page_break()



# =============== 16. SCHEDE DI PREVENZIONE ===============
H1("16. Schede di misure di prevenzione – macchine principali")

def scheda(titolo, righe):
    H2(titolo)
    make_table(
        ["N.", "Rischio meccanico", "Azione correttiva"],
        righe,
        col_widths=[Cm(1.0), Cm(5.0), Cm(10.0)],
    )

scheda("16.1 Tornio", [
    ("1", "Organi di trasmissione non protetti",
     "Racchiudere completamente entro carter fissi o mobili dotati di interblocco UNI EN ISO 14119 "
     "(All. V, p. 5, D.Lgs. 81/08)."),
    ("2", "Avviamento accidentale",
     "Dispositivi di avviamento azionabili solo intenzionalmente: leve a due tempi, pulsanti "
     "contornati da ghiera (All. V, p. 2.1)."),
    ("3", "Volantini di comando in moto",
     "Volantini svincolati dal sistema di trasmissione durante il moto, lisci, ad anima piena, con "
     "impugnatura ripiegabile (All. V, p. 3)."),
    ("4", "Proiezione di materiale",
     "Schermo trasparente in policarbonato resistente, fissato al carro portautensili o scorrevole "
     "su guida; protezione anche della parte posteriore."),
    ("5", "Inerzia del mandrino dopo arresto",
     "Sistema di frenatura efficace o protezione temporizzata che impedisca l'accesso fino "
     "all'arresto completo (All. V, Parte I, p. 6.4)."),
    ("6", "Arresto di emergenza",
     "Pulsante a fungo rosso/giallo o barra di arresto, conforme UNI EN ISO 13850 (All. V, "
     "Parte I, p. 6.2)."),
    ("7", "DPI e formazione",
     "Formazione specifica e addestramento documentato (art. 73); DPI prescritti (All. VIII)."),
    ("8", "Conformità normativa",
     "Per macchine nuove o modificate: Fascicolo tecnico, Manuale d'uso in italiano, Marcatura CE, "
     "Dichiarazione di conformità UE (Dir. 2006/42/CE / Reg. UE 2023/1230)."),
    ("9", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento manutentivo."),
])

scheda("16.2 Fresatrice", [
    ("1", "Organi di trasmissione non protetti",
     "Carter fissi o mobili con interblocco UNI EN ISO 14119 (All. V, p. 5)."),
    ("2", "Contatto con utensile e proiezione schegge",
     "Riparo trasparente resistente agli urti, fissato alla tavola portapezzo (All. V, p. 3)."),
    ("3", "Avviamento accidentale",
     "Dispositivi di avviamento azionabili solo intenzionalmente (All. V, p. 2.1)."),
    ("4", "Arresto di emergenza",
     "Pulsante a fungo rosso/giallo conforme UNI EN ISO 13850."),
    ("5", "Riavviamento automatico",
     "Dispositivo a bassa tensione contro il riavviamento intempestivo (CEI EN 60204-1)."),
    ("6", "DPI e formazione",
     "Formazione specifica e addestramento documentato (artt. 73 e 77)."),
    ("7", "Conformità normativa",
     "Fascicolo tecnico, Manuale d'uso, Marcatura CE, Dichiarazione di conformità UE."),
    ("8", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento manutentivo."),
])

scheda("16.3 Trapano a colonna", [
    ("1", "Organi di trasmissione non protetti",
     "Carter fissi o mobili con interblocco UNI EN ISO 14119 (All. V, p. 5)."),
    ("2", "Contatto con utensile e schegge",
     "Riparo trasparente resistente agli urti, fissato alla tavola portapezzo."),
    ("3", "Arresto di emergenza",
     "Pulsante a fungo rosso/giallo conforme UNI EN ISO 13850."),
    ("4", "Riavviamento automatico",
     "Dispositivo contro il riavviamento intempestivo (CEI EN 60204-1)."),
    ("5", "DPI e formazione",
     "Formazione specifica e DPI conformi (artt. 73 e 77)."),
    ("6", "Conformità normativa",
     "Fascicolo tecnico, Manuale d'uso, Marcatura CE, Dichiarazione di conformità UE."),
    ("7", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento manutentivo."),
])

scheda("16.4 Troncatrice", [
    ("1", "Protezione della lama",
     "Cuffia di protezione a molla o gravità, non manomissibile, completa copertura della lama in "
     "fase di riposo (All. V, Parte II, p. 5.5.4)."),
    ("2", "Comando di azionamento",
     "Per cicli manuali: comando ad azione mantenuta (UNI EN ISO 13849-1, CEI EN 60204-1)."),
    ("3", "DPI e formazione",
     "Occhiali, otoprotettori, guanti antitaglio (carico/scarico)."),
    ("4", "Conformità normativa",
     "Fascicolo tecnico, Manuale d'uso, Marcatura CE, Dichiarazione di conformità UE."),
    ("5", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento manutentivo."),
])

scheda("16.5 Molatrice", [
    ("1", "Protezione della mola",
     "Cuffia metallica che ricopre la parte non utilizzata estendendosi ai lati "
     "(All. V, Parte II, p. 5.1.3)."),
    ("2", "Proiezione di schegge",
     "Schermo trasparente in policarbonato; obbligo di occhiali (All. V, Parte II, p. 5.1.6)."),
    ("3", "Arresto di emergenza",
     "Pulsante a fungo rosso/giallo conforme UNI EN ISO 13850."),
    ("4", "DPI e formazione",
     "Occhiali/visiera, otoprotettori (LAeq > 80 dB(A)), maschera antipolvere."),
    ("5", "Conformità normativa",
     "Fascicolo tecnico, Manuale d'uso, Marcatura CE, Dichiarazione di conformità UE."),
    ("6", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento (sostituzione mola)."),
])

scheda("16.6 Rettificatrice", [
    ("1", "Contatti accidentali con la mola",
     "Cuffia metallica robusta per tutta la larghezza e periferia. Schermi mobili con interblocco "
     "UNI EN ISO 14119 (All. V, Parte I, p. 6.2)."),
    ("2", "Proiezione di schegge",
     "Come sopra. Per fissaggio magnetico: dispositivo che impedisce l'avviamento a piano "
     "magnetico disattivato (All. V, p. 3)."),
    ("3", "Arresto di emergenza",
     "Pulsante a fungo rosso/giallo conforme UNI EN ISO 13850."),
    ("4", "Urti/schiacciamenti con piano mobile",
     "Area di lavoro segregata con barriere fisiche o immateriali (All. V, p. 3)."),
    ("5", "DPI e formazione",
     "Formazione specifica, addestramento documentato e DPI conformi."),
    ("6", "Conformità normativa",
     "Fascicolo tecnico, Manuale d'uso, Marcatura CE, Dichiarazione di conformità UE."),
    ("7", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento manutentivo."),
])

scheda("16.7 Sega a nastro", [
    ("1", "Volani non protetti",
     "Riparo mobile con interblocco UNI EN ISO 14119 esteso al tratto di lama non utilizzato; "
     "apertura solo a volani fermi (All. V, Parte I, p. 6.3 e Parte II, p. 5.5.2)."),
    ("2", "Tratto di lama scoperto",
     "Riparo fisso registrabile limitato al tratto di taglio; spintoni e freno per arresto rapido "
     "del nastro (All. V, Parte II, p. 5.5.2; All. VI, p. 9)."),
    ("3", "Arresto di emergenza",
     "Pulsante a fungo rosso/giallo conforme UNI EN ISO 13850."),
    ("4", "DPI e formazione",
     "Guanti antitaglio (solo carico/scarico), occhiali, otoprotettori."),
    ("5", "Conformità normativa",
     "Fascicolo tecnico, Manuale d'uso, Marcatura CE, Dichiarazione di conformità UE."),
    ("6", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento (sostituzione nastro)."),
])

scheda("16.8 Sega circolare", [
    ("1", "Contatti accidentali con la lama",
     "(a) Cuffia registrabile, bloccabile all'altezza del pezzo, in materiale resistente, estesa "
     "fino allo spigolo anteriore (All. V, Parte II, p. 5.5.3); (b) schermi fissi ai lati sotto la "
     "tavola; (c) spingitoi per pezzi piccoli/sottili (All. VI, p. 9); (d) appoggio al dispositivo "
     "di guida longitudinale."),
    ("2", "Rifiuto del pezzo (kickback)",
     "Coltello divisore: spessore ≤ larghezza di taglio − 0,5 mm; posizionato a max 3 mm dalla "
     "dentatura; punto più alto ≤ 5 mm dalla sporgenza della lama."),
    ("3", "Arresto di emergenza",
     "Pulsante a fungo rosso/giallo conforme UNI EN ISO 13850."),
    ("4", "DPI e formazione",
     "Occhiali antiproiezione, otoprotettori, guanti antitaglio (carico/scarico)."),
    ("5", "Conformità normativa",
     "Fascicolo tecnico, Manuale d'uso, Marcatura CE, Dichiarazione di conformità UE."),
    ("6", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento manutentivo."),
])

scheda("16.9 Pialla a filo", [
    ("1", "Contatti con albero portacoltelli",
     "Copertura mobile oscillante con molle di richiamo o copertura regolabile in altezza e "
     "larghezza che copra l'intero albero (All. V, Parte II, p. 5.5.6; All. VI, p. 9)."),
    ("2", "Proiezione dei coltelli",
     "Fissaggio dei coltelli con sistemi idonei certificati; verifica del serraggio prima di ogni "
     "turno (All. V, Parte II, p. 5.1.2)."),
    ("3", "Proiezione di trucioli",
     "Dispositivi che impediscano la proiezione e/o impianto di aspirazione localizzata conforme "
     "(All. V, p. 3; All. IV, p. 2)."),
    ("4", "Arresto di emergenza",
     "Pulsante a fungo rosso/giallo conforme UNI EN ISO 13850."),
    ("5", "DPI e formazione",
     "Occhiali, otoprotettori, maschera antipolvere (polveri di legno – cancerogeno)."),
    ("6", "Conformità normativa",
     "Fascicolo tecnico, Manuale d'uso, Marcatura CE, Dichiarazione di conformità UE."),
    ("7", "Procedura LOTO",
     "Applicazione obbligatoria prima di ogni intervento (sostituzione coltelli)."),
])

add_page_break()



# =============== 17. DISPOSIZIONI FINALI ===============
H1("17. Disposizioni finali")

H2("17.1 Entrata in vigore")
P("La presente procedura entra in vigore dalla data di emissione e sostituisce integralmente la "
  "precedente edizione (FR 23/67, 1ª Emissione del 15 aprile 2013, prot. UPP/3/0133).")

H2("17.2 Diffusione e presa visione")
P("Il personale preposto deve assicurare che tutti i lavoratori assegnati ai reparti, officine, "
  "bacini, bordo navi e laboratori prendano visione della presente procedura entro 30 giorni dalla "
  "data di emissione, sottoscrivendo apposito foglio firma per ricevuta.")

H2("17.3 Conservazione")
P("L'originale della presente procedura è conservato e consultabile presso l'Ufficio Prevenzione e "
  "Protezione. Una copia in formato elettronico è disponibile nella cartella di rete condivisa "
  "«UPP» e nel sistema di gestione documentale dell'Arsenale.")

H2("17.4 Aggiornamento")
P("La presente procedura è soggetta a revisione:", indent_first=False)
BULLET("In occasione di modifiche normative significative.")
BULLET("A seguito di infortuni o «quasi-infortuni» (near miss) correlati all'uso di macchine e "
       "attrezzature.")
BULLET("In esito ad audit interni o esterni che evidenzino non conformità.")
BULLET("Con periodicità almeno triennale, anche in assenza dei precedenti eventi.")

H2("17.5 Sanzioni")
P("La violazione delle disposizioni contenute nella presente procedura costituisce infrazione "
  "disciplinare ai sensi della normativa vigente in materia di pubblico impiego e del Codice "
  "dell'Ordinamento Militare. Ferme restando le sanzioni penali e amministrative previste dal "
  "D.Lgs. 81/08 e s.m.i. (artt. 55–60), il personale trasgressore sarà sottoposto ai provvedimenti "
  "disciplinari previsti dall'ordinamento di appartenenza.")

# =============== 18. ALLEGATI ===============
H1("18. Allegati")
for v in [
    "Allegato A – Modello di Permesso di Lavoro per rimozione temporanea protezioni.",
    "Allegato B – Modello di Procedura LOTO (Lockout/Tagout) – scheda operativa.",
    "Allegato C – Registro Manutenzione Attrezzature (fac-simile).",
    "Allegato D – Verbale di addestramento e presa visione (fac-simile).",
    "Allegato E – Check-list di verifica pre-utilizzo macchina (fac-simile)."]:
    BULLET(v)

# Firma
firma = doc.add_paragraph()
firma.paragraph_format.space_before = Pt(48)
firma.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = firma.add_run("Il Capo Ufficio Prevenzione e Protezione")
r.bold = True; r.font.size = Pt(11); r.font.name = 'Times New Roman'

firma2 = doc.add_paragraph()
firma2.paragraph_format.space_before = Pt(36)
firma2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
r = firma2.add_run("__________________________________")
r.font.size = Pt(11); r.font.name = 'Times New Roman'

# Footer finale
end = doc.add_paragraph()
end.paragraph_format.space_before = Pt(36)
end.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = end.add_run("Fine documento – Procedura FR 23/67, Revisione 2ª, Anno 2026\n"
                "Arsenale Marina Militare – Taranto – Ufficio Prevenzione e Protezione")
r.italic = True; r.font.size = Pt(9); r.font.name = 'Times New Roman'

doc.save(OUT)
print(f"Salvato: {OUT}")
