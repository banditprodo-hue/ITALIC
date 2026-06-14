# -*- coding: utf-8 -*-
"""
Completa il Documento Informativo sui Rischi DIV-2026-001 (Art. 26 c.1 lett. b
D.Lgs. 81/2008) aggiungendo:
  - Sezione 5: DESCRIZIONE DELLE ATTIVITA' FORMATIVE E PROGRAMMA (attivita' interne)
  - Sezione 9: COOPERAZIONE E COORDINAMENTO (Art. 26 commi 2 e 3)
e rinumerando le sezioni successive. Lavora direttamente sull'OOXML per
preservare lo stile del documento originale (python-docx non disponibile).
"""
import zipfile, shutil, os

BASE = "/projects/sandbox/ITALIC"
SRC = os.path.join(BASE, "DIV-2026-001_Documento_Informativo_Rischi_MARINARSEN.docx")
OUT = os.path.join(BASE, "DIV-2026-001_Documento_Informativo_Rischi_MARINARSEN_completo.docx")

RF = '<w:rFonts w:ascii="Arial" w:cs="Arial" w:eastAsia="Arial" w:hAnsi="Arial"/>'


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def h1(text):
    return ('<w:p><w:pPr><w:spacing w:after="120" w:before="280"/></w:pPr>'
            '<w:r><w:rPr>' + RF + '<w:b/><w:bCs/><w:color w:val="2C4770"/>'
            '<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
            '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>')


def h2(text):
    return ('<w:p><w:pPr><w:spacing w:after="120" w:before="280"/></w:pPr>'
            '<w:r><w:rPr>' + RF + '<w:b/><w:bCs/><w:color w:val="2C4770"/>'
            '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>'
            '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>')


def para(text, italic=False, grey=False, size="20"):
    rpr = RF + '<w:b w:val="false"/><w:bCs w:val="false"/>'
    if italic:
        rpr += '<w:i/><w:iCs/>'
    if grey:
        rpr += '<w:color w:val="595959"/>'
    rpr += '<w:sz w:val="%s"/><w:szCs w:val="%s"/>' % (size, size)
    return ('<w:p><w:pPr><w:spacing w:after="60" w:before="60"/><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr>' + rpr + '</w:rPr>'
            '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>')


def bullet(text):
    return ('<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
            '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>'
            '<w:spacing w:after="40" w:before="40"/><w:jc w:val="both"/></w:pPr>'
            '<w:r><w:rPr>' + RF + '<w:sz w:val="19"/><w:szCs w:val="19"/></w:rPr>'
            '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>')


def _hdr_cell(text, w):
    return ('<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="%d"/><w:gridSpan w:val="1"/>'
            '<w:tcBorders><w:top w:val="single" w:color="2C4770" w:sz="4"/>'
            '<w:left w:val="single" w:color="2C4770" w:sz="4"/>'
            '<w:bottom w:val="single" w:color="2C4770" w:sz="4"/>'
            '<w:right w:val="single" w:color="2C4770" w:sz="4"/></w:tcBorders>'
            '<w:shd w:fill="2C4770" w:val="clear"/>'
            '<w:tcMar><w:top w:type="dxa" w:w="80"/><w:left w:type="dxa" w:w="120"/>'
            '<w:bottom w:type="dxa" w:w="80"/><w:right w:type="dxa" w:w="120"/></w:tcMar>'
            '<w:vAlign w:val="center"/></w:tcPr>'
            '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr>' + RF +
            '<w:b/><w:bCs/><w:color w:val="FFFFFF"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
            '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p></w:tc>') % (w, )


def _body_cell(text, w, bold=False, fill=None):
    rpr = RF
    if bold:
        rpr += '<w:b/><w:bCs/>'
    else:
        rpr += '<w:b w:val="false"/><w:bCs w:val="false"/>'
    rpr += '<w:sz w:val="19"/><w:szCs w:val="19"/>'
    shd = '<w:shd w:fill="%s" w:val="clear"/>' % fill if fill else ''
    return ('<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="%d"/><w:gridSpan w:val="1"/>'
            '<w:tcBorders><w:top w:val="single" w:color="AAAAAA" w:sz="4"/>'
            '<w:left w:val="single" w:color="AAAAAA" w:sz="4"/>'
            '<w:bottom w:val="single" w:color="AAAAAA" w:sz="4"/>'
            '<w:right w:val="single" w:color="AAAAAA" w:sz="4"/></w:tcBorders>' + shd +
            '<w:tcMar><w:top w:type="dxa" w:w="80"/><w:left w:type="dxa" w:w="120"/>'
            '<w:bottom w:type="dxa" w:w="80"/><w:right w:type="dxa" w:w="120"/></w:tcMar>'
            '<w:vAlign w:val="center"/></w:tcPr>'
            '<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr>' + rpr + '</w:rPr>'
            '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p></w:tc>') % (w, )


def table(headers, rows, widths):
    total = sum(widths)
    borders = ('<w:tblBorders><w:top w:val="single" w:color="auto" w:sz="4"/>'
               '<w:left w:val="single" w:color="auto" w:sz="4"/>'
               '<w:bottom w:val="single" w:color="auto" w:sz="4"/>'
               '<w:right w:val="single" w:color="auto" w:sz="4"/>'
               '<w:insideH w:val="single" w:color="auto" w:sz="4"/>'
               '<w:insideV w:val="single" w:color="auto" w:sz="4"/></w:tblBorders>')
    grid = '<w:tblGrid>' + ''.join('<w:gridCol w:w="%d"/>' % w for w in widths) + '</w:tblGrid>'
    out = '<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="%d"/>%s</w:tblPr>%s' % (total, borders, grid)
    out += '<w:tr>' + ''.join(_hdr_cell(h, widths[i]) for i, h in enumerate(headers)) + '</w:tr>'
    for ri, row in enumerate(rows):
        fill = "F2F5FA" if ri % 2 == 0 else None
        out += '<w:tr>' + ''.join(
            _body_cell(v, widths[i], bold=(i == 0), fill=fill) for i, v in enumerate(row)
        ) + '</w:tr>'
    out += '</w:tbl>'
    # un paragrafo vuoto dopo la tabella (come nel resto del documento)
    out += '<w:p><w:r><w:t xml:space="preserve"></w:t></w:r></w:p>'
    return out


# ---------------------------------------------------------------------------
# COSTRUZIONE SEZIONE 5 - DESCRIZIONE DELLE ATTIVITA' E PROGRAMMA
# ---------------------------------------------------------------------------
sez5 = []
sez5.append(h1("5. DESCRIZIONE DELLE ATTIVITÀ FORMATIVE E PROGRAMMA"))
sez5.append(para(
    "Il presente capitolo descrive le attività che i frequentatori del Corso di abilitazione alla condotta di "
    "impianti ad alta/media tensione – 2ª edizione 2026 (codice corso 26MAAEI55071M02) svolgono presso le aree "
    "operative di MARINARSEN Taranto. Tutte le attività hanno natura esclusivamente didattica e di osservazione: "
    "i frequentatori sono classificati come Persona Comune (PEC) ai sensi della Norma CEI 11-27 e non dispongono "
    "di alcuna autonomia operativa sugli impianti."))
sez5.append(para(
    "Le attività si articolano in affiancamento passivo (\"tandem\") al personale esperto e in prove pratiche "
    "d'esame, sempre sotto la sorveglianza a vista del F.T. Angelo Cardellicchio o del sostituto F.T. Luigi "
    "Gammariello, che esercitano funzione di garanzia. Nessuna manovra, operazione o intervento sugli impianti "
    "è consentita ai frequentatori."))

sez5.append(h2("5.1 Attività presso la Centrale Convertitori"))
sez5.append(para(
    "Presso la Centrale Convertitori, impianto di conversione e distribuzione dell'energia elettrica ad "
    "alta/media tensione, i frequentatori svolgono le seguenti attività:"))
for b in [
    "Osservazione didattica del layout impiantistico e identificazione dei principali componenti (gruppi di conversione, quadri, sbarre, trasformatori).",
    "Affiancamento passivo al personale PES/PAV durante l'illustrazione delle manovre tipo e delle procedure di esercizio dell'impianto.",
    "Riconoscimento dei dispositivi di manovra, sezionamento e protezione, senza alcuna operazione diretta sugli stessi.",
    "Osservazione delle procedure di messa in sicurezza dell'impianto (sequenza delle \"cinque regole d'oro\" della CEI 11-27) eseguite esclusivamente dal personale abilitato.",
    "Prova pratica d'esame consistente nella descrizione corretta delle sequenze e delle procedure, mantenendo la posizione arretrata oltre la Distanza di Vicinanza (DV) dalle parti in tensione.",
]:
    sez5.append(bullet(b))

sez5.append(h2("5.2 Attività presso la Centrale Elettrica Manganecchia"))
sez5.append(para(
    "Presso la Centrale Elettrica Manganecchia, centrale di produzione e smistamento dell'energia elettrica "
    "dotata di quadri MT/AT, trasformatori di potenza e macchinari rotanti, i frequentatori svolgono le "
    "seguenti attività:"))
for b in [
    "Osservazione dei quadri MT/AT, dei trasformatori di potenza e dei macchinari rotanti in esercizio, da posizione protetta e a distanza di sicurezza.",
    "Affiancamento passivo durante le manovre di smistamento e le procedure di conduzione della centrale.",
    "Riconoscimento della segnaletica di sicurezza, delle vie di esodo e dei presidi antincendio presenti nei locali.",
    "Prova pratica d'esame di conduzione simulata, descritta verbalmente sotto sorveglianza fissa a vista, senza alcun azionamento reale di apparecchiature.",
]:
    sez5.append(bullet(b))

sez5.append(h2("5.3 Programma e articolazione temporale delle attività"))
sez5.append(para("Le attività pratiche e d'esame si articolano nelle fasi riepilogate nella tabella seguente."))
sez5.append(table(
    ["FASE", "ATTIVITÀ", "SEDE", "REGIME DI SORVEGLIANZA"],
    [
        ["Fase 1\nBriefing di sicurezza",
         "Illustrazione dei rischi specifici, dei DPI obbligatori, delle vie di esodo, dei numeri di emergenza e delle prescrizioni comportamentali.",
         "Sito di accesso",
         "A cura del F.T. Cardellicchio o del sostituto F.T. Gammariello, prima di ogni accesso"],
        ["Fase 2\nAffiancamento passivo (tandem)",
         "Osservazione didattica degli impianti AT/MT in esercizio e delle manovre/procedure eseguite dal personale PES/PAV, senza alcuna operazione diretta.",
         "Centrale Convertitori e Centrale Manganecchia",
         "Sorveglianza continuativa a vista"],
        ["Fase 3\nEsami pratici",
         "Verifica delle competenze acquisite mediante prove pratiche di descrizione e conduzione simulata, mantenendo la Distanza di Vicinanza (DV).",
         "Centrale Convertitori e Centrale Manganecchia",
         "Sorveglianza fissa a vista del personale incaricato"],
        ["Fase 4\nDebriefing",
         "Riepilogo delle attività svolte e delle eventuali criticità di sicurezza rilevate.",
         "Sito di accesso",
         "Referente dell'attività"],
    ],
    widths=[1700, 3900, 2100, 1938],
))
sez5.append(para(
    "Le attività si svolgono nel periodo 08/06/2026 – 03/07/2026. Il calendario giornaliero di dettaglio è "
    "definito dal referente dell'attività (F.T. Cardellicchio) d'intesa con Mariscuola Taranto e comunicato ai "
    "frequentatori prima dell'accesso alle aree operative.", italic=True, grey=True, size="19"))

sez5.append(h2("5.4 Modalità di svolgimento e regime di sorveglianza"))
for b in [
    "I frequentatori operano esclusivamente in modalità di affiancamento passivo (\"tandem\") e, durante gli esami pratici, sotto sorveglianza fissa a vista del personale incaricato.",
    "È fatto divieto assoluto di eseguire qualsiasi manovra, operazione o intervento su impianti, quadri, sezionatori, interruttori o componenti, anche se ritenuti fuori tensione.",
    "I frequentatori mantengono in ogni momento la posizione indicata dal personale di sorveglianza, oltre la Distanza di Vicinanza (DV) stabilita dalla CEI 11-27 in funzione del livello di tensione presente.",
    "L'accesso alle aree operative è subordinato alla previa effettuazione del briefing di sicurezza e alla firma del foglio di presa visione del presente documento.",
    "In presenza di lavori o manovre sotto tensione eseguiti dal personale MARINARSEN, le attività didattiche sono sospese e i frequentatori allontanati dall'area interessata.",
]:
    sez5.append(bullet(b))

SEZ5 = "".join(sez5)

# ---------------------------------------------------------------------------
# COSTRUZIONE SEZIONE 9 - COOPERAZIONE E COORDINAMENTO (Art. 26 c.2 e c.3)
# ---------------------------------------------------------------------------
sez9 = []
sez9.append(h1("9. COOPERAZIONE E COORDINAMENTO (ART. 26, COMMI 2 E 3, D.LGS. 81/2008)"))
sez9.append(para(
    "Ai sensi dell'art. 26, comma 2, del D.Lgs. 81/2008, MARINARSEN Taranto (datore di lavoro committente) e "
    "Mariscuola Taranto (ente esterno di appartenenza dei frequentatori) cooperano all'attuazione delle misure "
    "di prevenzione e protezione e coordinano gli interventi di protezione e prevenzione dai rischi cui sono "
    "esposti i frequentatori, informandosi reciprocamente al fine di eliminare i rischi dovuti alle interferenze."))
sez9.append(para(
    "Considerato che le attività dei frequentatori hanno natura di sola osservazione passiva e di prova pratica "
    "sotto sorveglianza a vista, senza l'esecuzione di lavorazioni autonome che si sovrappongano alle attività "
    "operative ordinarie dello Stabilimento, non si configurano lavorazioni interferenti tali da richiedere la "
    "redazione di un DUVRI ai sensi dell'art. 26, comma 3. Le misure di coordinamento sono comunque formalizzate "
    "nel presente documento e attuate attraverso le seguenti disposizioni:"))
for b in [
    "Designazione di un referente unico dell'attività (F.T. Angelo Cardellicchio, sostituto F.T. Luigi Gammariello) responsabile del coordinamento e della sorveglianza a vista dei frequentatori.",
    "Effettuazione di un briefing di sicurezza preliminare congiunto prima di ogni accesso alle aree operative, con scambio reciproco delle informazioni sui rischi specifici.",
    "Programmazione delle attività didattiche in orari e modalità che non interferiscano con le attività operative ordinarie del personale MARINARSEN.",
    "Sospensione immediata delle attività e allontanamento dei frequentatori in caso di manovre, lavori sotto tensione o situazioni di emergenza nelle aree interessate.",
    "Divieto per i frequentatori di accedere autonomamente ad aree diverse da quelle indicate e di interferire con le attività del personale dello Stabilimento.",
    "Tracciabilità della presa visione e dell'accettazione delle prescrizioni mediante il foglio firme allegato al presente documento.",
]:
    sez9.append(bullet(b))
sez9.append(para(
    "Il presente documento costituisce attuazione dell'obbligo informativo di cui all'art. 26, comma 1, lett. b) "
    "e formalizza le misure di cooperazione e coordinamento di cui all'art. 26, comma 2, del D.Lgs. 81/2008."))

SEZ9 = "".join(sez9)

# ---------------------------------------------------------------------------
# LETTURA E MODIFICA DEL document.xml
# ---------------------------------------------------------------------------
with zipfile.ZipFile(SRC) as z:
    names = z.namelist()
    data = {n: z.read(n) for n in names}

xml = data["word/document.xml"].decode("utf-8")

# 1) Completa il numero di telefono segnaposto dell'Infermeria
xml = xml.replace("099 775 2XXX (c/o Guardia Circuiti)", "099 775 2841 (c/o Sala Medica Arsenale)")

# 2) Rinumerazione delle sezioni esistenti (le due nuove sezioni 5 e 9 si inseriscono)
renumber = [
    ("9. DISPOSIZIONI FINALI E PRESCRIZIONI GENERALI", "11. DISPOSIZIONI FINALI E PRESCRIZIONI GENERALI"),
    ("8. RIEPILOGO DISPOSITIVI DI PROTEZIONE INDIVIDUALE (DPI) OBBLIGATORI", "10. RIEPILOGO DISPOSITIVI DI PROTEZIONE INDIVIDUALE (DPI) OBBLIGATORI"),
    ("7.2 Quadro di sintesi dei rischi", "8.2 Quadro di sintesi dei rischi"),
    ("7.1 Codificazione dei fattori di rischio", "8.1 Codificazione dei fattori di rischio"),
    ("7. INFORMAZIONI SUI RISCHI SPECIFICI", "8. INFORMAZIONI SUI RISCHI SPECIFICI"),
    ("6. INFORMAZIONI SUGLI ACCESSI AL COMPRENSORIO", "7. INFORMAZIONI SUGLI ACCESSI AL COMPRENSORIO"),
    ("5.2 Procedure di emergenza ed evacuazione", "6.2 Procedure di emergenza ed evacuazione"),
    ("5.1 Numeri di emergenza", "6.1 Numeri di emergenza"),
    ("5. GESTIONE DELLE EMERGENZE, EVACUAZIONE E PRIMO SOCCORSO", "6. GESTIONE DELLE EMERGENZE, EVACUAZIONE E PRIMO SOCCORSO"),
]
for old, new in renumber:
    assert xml.count(old) == 1, "Atteso 1 match per: %r (trovati %d)" % (old, xml.count(old))
    xml = xml.replace(old, new)


def insert_before_heading(xml, heading_text, block):
    idx = xml.find(heading_text)
    assert idx != -1, "Heading non trovato: %r" % heading_text
    pstart = xml.rfind("<w:p>", 0, idx)
    assert pstart != -1
    return xml[:pstart] + block + xml[pstart:]


# 3) Inserisce la nuova Sezione 5 prima dell'attuale "6. GESTIONE DELLE EMERGENZE..."
xml = insert_before_heading(xml, "6. GESTIONE DELLE EMERGENZE, EVACUAZIONE E PRIMO SOCCORSO", SEZ5)

# 4) Inserisce la nuova Sezione 9 prima di "10. RIEPILOGO DISPOSITIVI..."
xml = insert_before_heading(xml, "10. RIEPILOGO DISPOSITIVI DI PROTEZIONE INDIVIDUALE (DPI) OBBLIGATORI", SEZ9)

data["word/document.xml"] = xml.encode("utf-8")

# ---------------------------------------------------------------------------
# RISCRITTURA DEL DOCX
# ---------------------------------------------------------------------------
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    for n in names:
        z.writestr(n, data[n])

print("Documento completato salvato in:", OUT)
print("Dimensione:", os.path.getsize(OUT), "bytes")
