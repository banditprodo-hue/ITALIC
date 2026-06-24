# -*- coding: utf-8 -*-
"""
Generatore PPTX - "Il Dolore Postoperatorio: una sfida di squadra"
Congresso Nazionale - Relatrice: Dott.ssa Antonella Cecere (Anestesista Rianimatore)
Pubblico: chirurghi, anestesisti, infermieri. Durata 30 min.
Enfasi: coordinamento dell'equipe + scale di valutazione del dolore.

Riusa il motore OOXML gia' validato (build_slide_rischio_meccanico.py).
Palette professionale: blu / grigio / bianco.
"""
import os
from build_slide_rischio_meccanico import (
    esc, run, para, shape_rect, shape_text, wrap_slide, build_pptx,
    reset_ids, SW, SH,
)

# Palette professionale
NAVY      = "1F3864"
BLUE      = "2E5496"
ACCENT    = "2E75B6"
STEEL     = "8FAADC"
LIGHTBLUE = "EAF1FA"
WHITE     = "FFFFFF"
LIGHTGRAY = "F2F4F7"
GRAY      = "595959"
DARKTEXT  = "1A2238"
MIDGRAY   = "808A9B"


def slide_cover(title, subtitle, presenter, affil, congress):
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=NAVY, name="bg"))
    s.append(shape_rect(0, 0, SW, 70000, fill=ACCENT, name="topbar"))
    # blocco diagonale accento
    s.append(shape_rect(0, 2050000, SW, 1900000, fill=BLUE, name="band"))
    s.append(shape_rect(0, 2050000, 360000, 1900000, fill=STEEL, name="lbar"))
    s.append(shape_text(820000, 1450000, 10500000, 560000,
        [para([run("CONGRESSO NAZIONALE - TERAPIA DEL DOLORE PERIOPERATORIO", 1400,
                   STEEL, bold=True)], space_before=0)]))
    s.append(shape_text(820000, 2150000, 10500000, 1150000,
        [para([run(title, 4000, WHITE, bold=True)], space_before=0, line=104000)]))
    s.append(shape_text(820000, 3320000, 10500000, 560000,
        [para([run(subtitle, 1850, STEEL)], space_before=0, line=108000)]))
    # relatore
    s.append(shape_text(820000, 4350000, 10500000, 1100000,
        [para([run(presenter, 2000, WHITE, bold=True)], space_before=0),
         para([run(affil, 1500, STEEL)], space_before=200)]))
    s.append(shape_rect(0, SH-540000, SW, 540000, fill=ACCENT, name="ft"))
    s.append(shape_text(820000, SH-540000, 10500000, 540000,
        [para([run(congress, 1300, WHITE, bold=True)], space_before=0)], anchor="ctr"))
    return wrap_slide(s)


def slide_content(num, total, title, tag, bullets, image_desc, kicker=""):
    """bullets: lista di stringhe; prefisso '>' = secondo livello."""
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=WHITE, name="bg"))
    # header
    s.append(shape_rect(0, 0, SW, 1020000, fill=NAVY, name="header"))
    s.append(shape_rect(0, 1020000, SW, 46000, fill=ACCENT, name="hacc"))
    s.append(shape_rect(0, 0, 140000, 1020000, fill=STEEL, name="htab"))
    if kicker:
        s.append(shape_text(420000, 120000, 8800000, 300000,
            [para([run(kicker.upper(), 1050, STEEL, bold=True)], space_before=0)]))
        ty, th = 360000, 600000
    else:
        ty, th = 90000, 850000
    s.append(shape_text(420000, ty, 8800000, th,
        [para([run(title, 2500, WHITE, bold=True)], space_before=0, line=102000)],
        anchor="ctr"))
    # tag normativo / fonte
    if tag:
        s.append(shape_rect(9250000, 300000, 2680000, 470000, fill=STEEL, name="tag"))
        s.append(shape_text(9270000, 300000, 2640000, 470000,
            [para([run(tag, 1050, NAVY, bold=True)], algn="ctr", space_before=0)],
            anchor="ctr"))
    # bullet
    bp = []
    for b in bullets:
        if b.startswith(">"):
            bp.append(para([run(b[1:].strip(), 1400, GRAY)], bullet=True, level=1,
                           bullet_color=STEEL, space_before=300))
        else:
            bp.append(para([run(b, 1650, DARKTEXT)], bullet=True, level=0,
                           bullet_color=ACCENT, space_before=520))
    s.append(shape_text(470000, 1360000, 7150000, 4950000, bp))
    # box immagine
    bx, by, bw, bh = 7820000, 1360000, 3950000, 4800000
    s.append(shape_rect(bx, by, bw, bh, fill=LIGHTBLUE, line=STEEL, line_w=9525,
                        name="imgbox"))
    s.append(shape_rect(bx, by, bw, 430000, fill=NAVY, name="imghdr"))
    s.append(shape_rect(bx, by, 130000, 430000, fill=ACCENT, name="imgtab"))
    s.append(shape_text(bx+200000, by, bw-220000, 430000,
        [para([run("IMMAGINE / GRAFICO", 1100, WHITE, bold=True)], space_before=0)],
        anchor="ctr"))
    s.append(shape_text(bx+170000, by+540000, bw-340000, bh-700000,
        [para([run(image_desc, 1300, DARKTEXT, italic=True)], space_before=0,
              line=116000)]))
    # footer
    s.append(shape_rect(0, SH-450000, SW, 450000, fill=NAVY, name="footer"))
    s.append(shape_rect(0, SH-450000, SW, 40000, fill=ACCENT, name="facc"))
    s.append(shape_text(420000, SH-450000, 9000000, 450000,
        [para([run("Dolore postoperatorio - Dott.ssa A. Cecere, Anestesista Rianimatore",
                   1000, WHITE)], space_before=0)], anchor="ctr"))
    s.append(shape_text(SW-2100000, SH-450000, 1680000, 450000,
        [para([run(f"{num} / {total}", 1050, STEEL, bold=True)], algn="r",
              space_before=0)], anchor="ctr"))
    return wrap_slide(s)


def slide_section(num, total, big, sub):
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=NAVY, name="bg"))
    s.append(shape_rect(820000, 3050000, 1700000, 60000, fill=ACCENT, name="acc"))
    s.append(shape_text(820000, 2350000, 10500000, 1000000,
        [para([run(big, 3600, WHITE, bold=True)], space_before=0)]))
    s.append(shape_text(820000, 3250000, 10500000, 700000,
        [para([run(sub, 1700, STEEL)], space_before=0)]))
    s.append(shape_text(SW-2100000, SH-520000, 1680000, 420000,
        [para([run(f"{num} / {total}", 1050, STEEL)], algn="r", space_before=0)]))
    return wrap_slide(s)



# ---------------------------------------------------------------------------
# CONTENUTI DELLE SLIDE (slide 2 -> 26) + backup
# Ogni voce: (titolo, tag/fonte, [bullet], descrizione immagine, kicker)
# ---------------------------------------------------------------------------
SLIDES = [
 # 2 - Obiettivi
 ("Obiettivi di apprendimento", "Learning goals",
  ["Riconoscere impatto clinico e prognostico del dolore postoperatorio non controllato",
   "Applicare in modo sistematico le scale del dolore (paziente comunicante e non)",
   "Impostare un'analgesia multimodale e opioid-sparing, procedura-specifica",
   "Definire ruoli e flussi di comunicazione dell'equipe (Acute Pain Service, handover SBAR)"],
  "Schema (BioRender) con 4 icone: valutazione, farmaco, blocco loco-regionale, team. Palette blu/grigio, 300 dpi.",
  "Obiettivi formativi"),

 # 3 - Epidemiologia
 ("Epidemiologia e impatto clinico", "CPG 2024 / Gan 2017",
  ["Una quota rilevante di pazienti riferisce dolore moderato-severo dopo l'intervento",
   "Il dolore acuto severo e' predittore di dolore cronico postchirurgico (CPSP)",
   "CPSP: incidenza variabile per tipo di chirurgia (toracica, mammella, ernia, ortopedia)",
   "Conseguenze: ritardata mobilizzazione, complicanze, degenza e costi maggiori"],
  "Grafico a barre: prevalenza dolore moderato-severo per tipo di chirurgia. Fonte: dati aggregati, ricreare in stile vettoriale 300 dpi.",
  "Epidemiologia"),

 # 4 - Fisiopatologia 1
 ("Fisiopatologia del dolore acuto", "Pain pathway",
  ["Lesione tissutale -> mediatori infiammatori -> sensibilizzazione periferica",
   "Amplificazione midollare e sensibilizzazione centrale (fenomeno del wind-up)",
   "Comparsa di iperalgesia e allodinia",
   "Componenti: nocicettiva, infiammatoria e neuropatica (lesione nervosa)"],
  "Diagramma della via del dolore (nocicettore -> midollo -> talamo -> corteccia) con punti d'azione dei farmaci. BioRender, 300 dpi.",
  "Fisiopatologia I"),

 # 5 - Fisiopatologia 2
 ("Dalla fase acuta alla cronicizzazione", "Glare 2019 / Lancet",
  ["Esiste una finestra critica: un buon controllo del dolore acuto previene la cronicizzazione",
   "Fattori di rischio CPSP: dolore preoperatorio, dolore acuto intenso, lesione nervosa",
   "Fattori psicologici: catastrofizzazione, ansia, depressione",
   "Fattori non modificabili: sesso femminile, eta' giovane, tipo di intervento"],
  "Timeline orizzontale 'dolore acuto -> transizione -> dolore cronico' con finestra d'intervento evidenziata. Vettoriale 300 dpi.",
  "Fisiopatologia II"),

 # 6 - Classificazione/tipi
 ("Classificazione del dolore postoperatorio", "IASP",
  ["Nocicettivo: somatico (ferita, osso) e viscerale (organi)",
   "Neuropatico: lesione/irritazione nervosa, spesso urente o a scossa",
   "Misto: combinazione delle componenti",
   "Valutare sempre dolore a RIPOSO e DINAMICO (movimento, tosse, mobilizzazione)"],
  "Tabella comparativa tipi di dolore (caratteristiche, esempi chirurgici, target terapeutico). Stile tabellare blu/grigio.",
  "Classificazione"),

 # 7 - Scale unidimensionali (KEY)
 ("Le scale del dolore: paziente comunicante", "NRS / VAS / VRS",
  ["NRS (0-10): scala numerica, standard raccomandato nell'adulto comunicante",
   "VAS: scala analogico-visiva (linea 0-100 mm)",
   "VRS: scala verbale (nessuno-lieve-moderato-severo)",
   "Wong-Baker FACES: utile in pediatria e barriere linguistiche",
   ">NRS, VAS e VRS mostrano forte correlazione tra loro nel postoperatorio"],
  "Pannello visivo con NRS 0-10, righello VAS e scala FACES affiancate. Grafica ad alta leggibilita', 300 dpi.",
  "Scale del dolore I"),

 # 8 - Scale pz non comunicante (KEY)
 ("Le scale del dolore: paziente non comunicante", "CPOT / BPS / PAINAD",
  ["CPOT e BPS: pazienti sedati/intubati in terapia intensiva",
   "PAINAD: pazienti con deterioramento cognitivo/demenza",
   "Si basano su indicatori comportamentali osservabili",
   ">Espressione facciale, tono muscolare, movimenti, adattamento al ventilatore",
   "Mai assumere assenza di dolore solo perche' il paziente non lo riferisce"],
  "Tabella di scoring CPOT (4 item, 0-8) con descrizione dei comportamenti. Ricreare come tabella vettoriale 300 dpi.",
  "Scale del dolore II"),

 # 9 - Algoritmo valutazione
 ("Valutazione strutturata: come e quando", "CPG 2024",
  ["Misurare a riposo e in movimento, a intervalli regolari e dopo ogni terapia",
   "Target operativi: NRS <=3 a riposo e controllo del dolore dinamico",
   "Il dolore come '5o parametro vitale': sempre documentato",
   "Rivalutare dopo l'intervento analgesico (es. entro 30-60 minuti)"],
  "Flowchart dell'algoritmo: valuta -> tratta -> rivaluta -> aggiusta. Diagramma a blocchi blu/grigio, 300 dpi.",
  "Valutazione"),

 # 10 - Equipe e ruoli (FOCUS)
 ("Una responsabilita' di squadra", "Team coordination",
  ["Chirurgo: tecnica, infiltrazione ferita, indicazioni procedura-specifiche",
   "Anestesista: piano analgesico, tecniche loco-regionali, gestione del rischio",
   "Infermiere: valutazione, somministrazione, monitoraggio, educazione del paziente",
   "Acute Pain Service: protocolli condivisi, formazione, audit e continuita'",
   ">Obiettivi e soglie d'intervento devono essere CONDIVISI da tutta l'equipe"],
  "Matrice RACI / diagramma a cerchi sovrapposti dei ruoli (chirurgo-anestesista-infermiere-APS) con il paziente al centro. 300 dpi.",
  "Coordinamento equipe"),

 # 11 - Comunicazione / handover
 ("Comunicazione e passaggio di consegne", "SBAR",
  ["Documentazione standardizzata: migliora gli esiti e riduce il consumo di oppioidi",
   "Soglie d'intervento condivise (es. rivalutare/trattare se NRS >=4)",
   "Handover strutturato SBAR tra sala operatoria, PACU e reparto",
   ">Situation - Background - Assessment - Recommendation"],
  "Template grafico SBAR applicato al dolore postoperatorio (4 riquadri). Stile infografico blu/grigio, 300 dpi.",
  "Comunicazione"),

 # 12 - Multimodale principi
 ("Analgesia multimodale: i principi", "PROSPECT / CPG 2024",
  ["Combinare farmaci con meccanismi d'azione diversi e complementari",
   "Strategia opioid-sparing: ridurre dose e effetti avversi degli oppioidi",
   "Approccio procedura-specifico (raccomandazioni PROSPECT)",
   "Integrazione nei percorsi ERAS (Enhanced Recovery After Surgery)"],
  "Piramide/colonna dei bersagli lungo la via del dolore con i farmaci che agiscono su ciascun punto. BioRender, 300 dpi.",
  "Trattamento I"),

 # 13 - Non oppioidi
 ("Analgesici non oppioidi e adiuvanti", "Multimodal review 2024",
  ["Base: paracetamolo + FANS o COX-2 selettivi (salvo controindicazioni)",
   "Adiuvanti selettivi: ketamina a basse dosi, lidocaina ev, desametasone",
   "Alfa-2 agonisti e gabapentinoidi in casi selezionati",
   ">Riducono il fabbisogno di oppioidi e migliorano il controllo dinamico"],
  "Tabella 'farmaco - meccanismo d'azione - ruolo nel multimodale'. Stile tabellare, evitare dosaggi specifici. 300 dpi.",
  "Trattamento II"),

 # 14 - Oppioidi
 ("Oppioidi: uso razionale e sicuro", "CPG 2024",
  ["Indicati nel dolore moderato-severo, alla dose minima efficace e per il tempo piu' breve",
   "PCA (analgesia controllata dal paziente) per una titolazione personalizzata",
   "Rischi: depressione respiratoria, sedazione, nausea/vomito, ileo, ritenzione",
   "Monitorare la sedazione; attenzione a iperalgesia, tolleranza e dipendenza"],
  "Foto/schema di pompa PCA con legenda dei parametri (bolo, lockout). Imaging anonimizzato o vettoriale, 300 dpi.",
  "Trattamento III"),

 # 15 - Loco-regionali
 ("Tecniche loco-regionali", "ASRA / PROSPECT",
  ["Analgesia neuroassiale (peridurale) nella chirurgia maggiore toraco-addominale",
   "Blocchi dei piani fasciali ecoguidati: TAP, ESP, PECS, quadrato dei lombi",
   "Infiltrazione della ferita e cateteri perineurali continui",
   ">Migliorano il controllo del dolore dinamico e riducono gli oppioidi"],
  "Immagine ecografica di blocco ecoguidato (es. ESP block) con annotazioni anatomiche. Imaging anonimizzato, 300 dpi.",
  "Trattamento IV"),

 # 16 - Procedure-specific / ERAS
 ("Approccio procedura-specifico ed ERAS", "PROSPECT 2024",
  ["Il miglior schema dipende dal tipo di intervento (procedura-specifico)",
   "Toracotomia: blocco paravertebrale o ESP; addome: TAP o peridurale",
   "Mammella: blocchi della parete toracica; ortopedia: blocchi periferici dedicati",
   "Integrazione nei bundle ERAS per recupero precoce della funzione"],
  "Tabella comparativa: tipo di chirurgia -> tecnica raccomandata -> livello di evidenza (PROSPECT). 300 dpi.",
  "Trattamento V"),

 # 17 - Complicanze farmaci
 ("Complicanze: gestione perioperatoria", "Safety",
  ["Oppioidi: depressione respiratoria (antagonista: naloxone), sedazione, PONV, ileo",
   "FANS: rischio gastrointestinale, renale, cardiovascolare ed emorragico",
   "Prevenzione: selezione del paziente, monitoraggio, profilassi PONV",
   ">Personalizzare in base a comorbidita' e fragilita'"],
  "Tabella 'effetto avverso - segni - gestione' con codifica a colori del rischio. Stile clinico, 300 dpi.",
  "Complicanze I"),

 # 18 - Complicanze regionali
 ("Complicanze delle tecniche regionali", "ASRA",
  ["Peridurale: ipotensione, blocco motorio, raro ematoma/ascesso (gestione anticoagulanti)",
   "Tossicita' sistemica da anestetici locali (LAST): riconoscimento precoce",
   "Trattamento della LAST con emulsione lipidica (lipid rescue)",
   ">Sorveglianza neurologica e protocolli di emergenza disponibili"],
  "Algoritmo di gestione della LAST (riconoscimento -> ABC -> lipidi) a step. Diagramma 300 dpi.",
  "Complicanze II"),

 # 19 - APS / follow-up
 ("Acute Pain Service e continuita'", "APS",
  ["Team dedicato con ronde, protocolli condivisi e formazione continua",
   "Riduce intensita' del dolore e complicanze, aumenta la soddisfazione",
   "Transizione sicura degli oppioidi alla dimissione (de-prescribing)",
   ">Audit periodici e indicatori di qualita' del controllo del dolore"],
  "Workflow dell'Acute Pain Service (24/7) con punti di contatto del percorso del paziente. Infografica 300 dpi.",
  "Follow-up I"),

 # 20 - Prevenzione CPSP
 ("Prevenire il dolore cronico postchirurgico", "Doleman 2023 / BJA",
  ["Identificare precocemente i pazienti ad alto rischio",
   "Ottimizzare l'analgesia perioperatoria multimodale e opioid-sparing",
   "Transitional Pain Service per i casi complessi",
   ">Rivalutazione strutturata a distanza dall'intervento"],
  "Modello di Transitional Pain Service (pre -> intra -> post -> follow-up). Schema a percorso, 300 dpi.",
  "Follow-up II"),

 # 21 - Casi clinici
 ("Casi clinici", "Case-based",
  ["Caso 1 - Chirurgia addominale maggiore (ERAS): peridurale/TAP + multimodale, PCA, mobilizzazione precoce",
   "Caso 2 - Paziente anziano con deterioramento cognitivo: valutazione con PAINAD, schema opioid-sparing",
   ">Discussione: scelta della scala, ruolo dell'equipe e outcome funzionale"],
  "Due pannelli affiancati con timeline del caso e box 'pre/post trattamento' (NRS o PAINAD). Imaging anonimizzato, 300 dpi.",
  "Casi clinici"),

 # 22 - Controversie
 ("Controversie e ricerca in corso", "Open questions",
  ["Gabapentinoidi di routine: benefici limitati a fronte di effetti avversi",
   "Ketamina e lidocaina ev: evidenze in evoluzione, indicazioni da definire",
   "Anestetici locali a rilascio prolungato: dibattito su costo-efficacia",
   ">Predizione del dolore con strumenti digitali e intelligenza artificiale"],
  "Slide 'research highlights' con 3-4 trial/ambiti in corso e icone. Stile editoriale blu/grigio, 300 dpi.",
  "Controversie"),

 # 23 - Take-home
 ("Take-home messages", "Key points",
  ["Il dolore postoperatorio e' prevedibile e trattabile: valutarlo e' il primo atto terapeutico",
   "Una scala condivisa (NRS) e strumenti dedicati ai non comunicanti (CPOT/PAINAD)",
   "Analgesia multimodale, opioid-sparing e procedura-specifica",
   "Controllare il dolore acuto previene il dolore cronico",
   "Il coordinamento dell'equipe (APS, handover SBAR) fa la differenza"],
  "Visual riepilogativo con 5 icone numerate dei messaggi chiave. Grafica pulita, 300 dpi.",
  "Conclusioni"),

 # 24 - Bibliografia
 ("Bibliografia essenziale (Vancouver)", "References 1-7",
  [">1. Clinical practice guidelines for postoperative pain management in adults (2024 ed.). 2024.",
   ">2. Chou R, et al. Management of postoperative pain: APS/ASRA/ASA guideline. J Pain. 2016;17(2):131-57.",
   ">3. Joshi GP, et al. PROSPECT procedure-specific recommendations. Anaesthesia (aggiornamenti 2023-2024).",
   ">4. Doleman B, et al. Non-opioid analgesics for prevention of CPSP: NMA. Br J Anaesth. 2023.",
   ">5. Glare P, Aubrey KR, Myles PS. Transition from acute to chronic pain after surgery. Lancet. 2019;393:1537-46.",
   ">6. Gan TJ. Poorly controlled postoperative pain: prevalence and prevention. J Pain Res. 2017;10:2287-98.",
   ">7. Perioperative multimodal analgesia: efficacy and safety review. 2024."],
  "Sfondo neutro; elenco riferimenti su due colonne se necessario. Verificare sempre le citazioni sulle fonti originali.",
  "Bibliografia"),

 # 25 - Bibliografia 2
 ("Bibliografia essenziale (Vancouver)", "References 8-14",
  [">8. Kehlet H, Jensen TS, Woolf CJ. Persistent postsurgical pain: risk factors and prevention. Lancet. 2006;367:1618-25.",
   ">9. Beyond measurement: pain scales for postoperative pain assessment. 2024.",
   ">10. Gelinas C, et al. Critical-Care Pain Observation Tool (CPOT): validazione e uso clinico.",
   ">11. Warden V, et al. Pain Assessment in Advanced Dementia (PAINAD) scale.",
   ">12. ERAS Society Guidelines (procedure-specifiche), aggiornamenti recenti.",
   ">13. Neal JM, et al. ASRA practice advisory on Local Anesthetic Systemic Toxicity (LAST).",
   ">14. Documentazione del dolore e consumo di oppioidi: studio d'intervento. 2022."],
  "Continua l'elenco. Formato Vancouver coerente; controllare volume/pagine sulle riviste originali.",
  "Bibliografia"),

 # 26 - Disclosure
 ("Disclosure / Conflitti di interesse", "Disclosure",
  ["La relatrice dichiara l'assenza di conflitti di interesse rilevanti per questa presentazione",
   "Eventuali rapporti con aziende (relazioni, advisory board, grant) vanno elencati qui",
   "I contenuti riflettono linee guida ed evidenze indipendenti",
   ">Aggiornare con le dichiarazioni effettive prima del congresso"],
  "Slide istituzionale sobria con logo del congresso/centro. Sfondo bianco, testo navy.",
  "Disclosure"),
]

# Slide di backup
BACKUP = [
 ("Backup - Dettaglio scala CPOT", "Backup",
  ["Espressione facciale (0-2)", "Movimenti corporei (0-2)",
   "Tensione muscolare (0-2)", "Adattamento al ventilatore o vocalizzazione (0-2)",
   ">Punteggio totale 0-8; soglia indicativa di dolore > 2"],
  "Tabella CPOT completa con descrittori per ciascun punteggio. Vettoriale 300 dpi.", "Backup"),
 ("Backup - Handover SBAR esteso", "Backup",
  ["Situation: stato attuale del dolore (scala, sede, dinamico/riposo)",
   "Background: intervento, tecnica analgesica, allergie, comorbidita'",
   "Assessment: efficacia, effetti avversi, sedazione",
   "Recommendation: piano, soglie, quando allertare l'anestesista/APS"],
  "Modulo SBAR compilabile per il dolore postoperatorio. Layout a form, 300 dpi.", "Backup"),
 ("Backup - Elementi del bundle ERAS", "Backup",
  ["Preoperatorio: educazione, ottimizzazione, no digiuno prolungato",
   "Intraoperatorio: analgesia multimodale, tecniche regionali, normotermia",
   "Postoperatorio: mobilizzazione precoce, rialimentazione, opioid-sparing",
   ">Il controllo del dolore e' un pilastro trasversale dell'intero percorso"],
  "Schema circolare del percorso ERAS con la voce 'analgesia' evidenziata. 300 dpi.", "Backup"),
]


def main():
    slides = []
    # Totale conteggiato sulle slide principali (1..N). I backup sono extra.
    total = 1 + len(SLIDES)

    # 1 - Cover
    slides.append(slide_cover(
        title="Il Dolore Postoperatorio: una sfida di squadra",
        subtitle="Valutazione strutturata, analgesia multimodale e coordinamento dell'equipe",
        presenter="Dott.ssa Antonella Cecere",
        affil="Anestesista Rianimatore",
        congress="Congresso Nazionale di Urologia / Anestesia - Sessione Terapia del Dolore"))

    # 2..N - contenuti
    for i, (title, tag, bullets, img, kicker) in enumerate(SLIDES, start=2):
        slides.append(slide_content(i, total, title, tag, bullets, img, kicker))

    # Sezione backup
    slides.append(slide_section(total, total, "Slide di backup",
                                "Materiale di supporto per la discussione"))
    n = total
    for (title, tag, bullets, img, kicker) in BACKUP:
        n += 1
        slides.append(slide_content(n, n, title, tag, bullets, img, kicker))

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "Dolore Postoperatorio - Cecere.pptx")
    build_pptx(slides, out, title="Il Dolore Postoperatorio - una sfida di squadra")
    print(f"Creato: {out}")
    print(f"Slide principali: {total} | totali (con backup): {len(slides)}")


if __name__ == "__main__":
    main()
