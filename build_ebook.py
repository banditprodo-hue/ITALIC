"""
Genera l'ebook PDF "Conoscenze Basilari su Blockchain, Bitcoin, ETH"
di PROTA DOMENICO.
Linguaggio semplice, adatto a tutti. Italiano.
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image,
    Table, TableStyle, ListFlowable, ListItem, KeepTogether
)
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

OUT_PDF = "/projects/sandbox/ITALIC/Ebook_Crypto_Prota_Domenico.pdf"
LOGO = "/projects/sandbox/ITALIC/logo_prota_domenico.png"

# --- Font ----------------------------------------------------------------
FONT_PATHS = [
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]
FONT_PATHS_BOLD = [
    "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]
def first_existing(paths):
    for p in paths:
        if os.path.exists(p):
            return p
    return None

reg = first_existing(FONT_PATHS)
bold = first_existing(FONT_PATHS_BOLD)
if reg:
    pdfmetrics.registerFont(TTFont("Body", reg))
if bold:
    pdfmetrics.registerFont(TTFont("BodyBold", bold))

BODY_FONT = "Body" if reg else "Helvetica"
BODY_BOLD = "BodyBold" if bold else "Helvetica-Bold"

# --- Colori --------------------------------------------------------------
GOLD = HexColor("#F7931A")        # Bitcoin
NAVY = HexColor("#1C235A")        # blu profondo
NAVY_LIGHT = HexColor("#3468C0")
ETH_PURPLE = HexColor("#627EEA")
LIGHT_BG = HexColor("#F5F1E6")    # crema
DARK = HexColor("#0F142D")
GREY = HexColor("#666666")

# --- Stili ---------------------------------------------------------------
styles = getSampleStyleSheet()

H1 = ParagraphStyle("H1", fontName=BODY_BOLD, fontSize=24, leading=30,
                    textColor=NAVY, spaceAfter=14, spaceBefore=10)
H2 = ParagraphStyle("H2", fontName=BODY_BOLD, fontSize=16, leading=22,
                    textColor=GOLD, spaceAfter=8, spaceBefore=14)
H3 = ParagraphStyle("H3", fontName=BODY_BOLD, fontSize=13, leading=18,
                    textColor=NAVY_LIGHT, spaceAfter=6, spaceBefore=10)
BODY = ParagraphStyle("Body", fontName=BODY_FONT, fontSize=11, leading=16,
                      textColor=DARK, alignment=TA_JUSTIFY, spaceAfter=8)
BULLET = ParagraphStyle("Bullet", parent=BODY, leftIndent=14, bulletIndent=2,
                        spaceAfter=4)
QUOTE = ParagraphStyle("Quote", parent=BODY, leftIndent=20, rightIndent=20,
                       textColor=GREY, fontSize=10, leading=14,
                       borderColor=GOLD, borderWidth=0, spaceBefore=6,
                       spaceAfter=10)
TIP = ParagraphStyle("Tip", parent=BODY, fontSize=10.5, leading=15,
                     textColor=DARK, leftIndent=10, rightIndent=10,
                     spaceBefore=4, spaceAfter=4)
TOC_ITEM = ParagraphStyle("TOC", fontName=BODY_FONT, fontSize=12,
                          leading=22, textColor=DARK)

# --- Cover, header, footer ----------------------------------------------
def draw_cover(canv, doc):
    """Pagina 1: copertina."""
    w, h = A4
    # Sfondo: gradiente verticale (simulato con strisce)
    steps = 80
    for i in range(steps):
        t = i / steps
        r = int(15 * (1 - t) + 28 * t)
        g = int(20 * (1 - t) + 35 * t)
        b = int(45 * (1 - t) + 90 * t)
        canv.setFillColorRGB(r / 255, g / 255, b / 255)
        canv.rect(0, h - (i + 1) * h / steps, w, h / steps + 1, fill=1, stroke=0)

    # Bordo dorato
    canv.setStrokeColor(GOLD)
    canv.setLineWidth(4)
    canv.rect(1.2 * cm, 1.2 * cm, w - 2.4 * cm, h - 2.4 * cm, fill=0, stroke=1)
    canv.setLineWidth(1)
    canv.rect(1.5 * cm, 1.5 * cm, w - 3.0 * cm, h - 3.0 * cm, fill=0, stroke=1)

    # Logo
    if os.path.exists(LOGO):
        logo_size = 7 * cm
        canv.drawImage(LOGO, (w - logo_size) / 2, h - 11 * cm,
                       width=logo_size, height=logo_size,
                       preserveAspectRatio=True, mask='auto')

    # Etichetta "EBOOK"
    canv.setFont(BODY_BOLD, 14)
    canv.setFillColor(GOLD)
    canv.drawCentredString(w / 2, h - 12.2 * cm, "EBOOK")

    # Titolo
    canv.setFillColor(white)
    canv.setFont(BODY_BOLD, 30)
    canv.drawCentredString(w / 2, h - 14 * cm, "CONOSCENZE BASILARI")
    canv.setFont(BODY_BOLD, 22)
    canv.setFillColor(GOLD)
    canv.drawCentredString(w / 2, h - 15.3 * cm, "su Blockchain, Bitcoin & Ethereum")

    # Sottotitolo
    canv.setFont(BODY_FONT, 13)
    canv.setFillColor(white)
    canv.drawCentredString(w / 2, h - 16.6 * cm,
                           "Come comprare e vendere crypto in sicurezza")
    canv.drawCentredString(w / 2, h - 17.3 * cm,
                           "+ Guida pratica all'Hard Wallet fai-da-te")

    # Linea decorativa
    canv.setStrokeColor(GOLD)
    canv.setLineWidth(2)
    canv.line(w / 2 - 4 * cm, h - 18 * cm, w / 2 + 4 * cm, h - 18 * cm)

    # Autore
    canv.setFont(BODY_FONT, 12)
    canv.setFillColor(white)
    canv.drawCentredString(w / 2, h - 19.5 * cm, "Autore")
    canv.setFont(BODY_BOLD, 26)
    canv.setFillColor(GOLD)
    canv.drawCentredString(w / 2, h - 21 * cm, "PROTA DOMENICO")

    # Tagline in basso
    canv.setFont(BODY_FONT, 10)
    canv.setFillColor(HexColor("#CCCCCC"))
    canv.drawCentredString(w / 2, 2.3 * cm,
                           "Una guida semplice, chiara e adatta a tutti.")
    canv.drawCentredString(w / 2, 1.9 * cm,
                           "Nessuna esperienza richiesta.")


def draw_page(canv, doc):
    """Header + footer per le pagine interne."""
    w, h = A4
    # Header: linea dorata + titolo abbreviato
    canv.setStrokeColor(GOLD)
    canv.setLineWidth(0.8)
    canv.line(2 * cm, h - 1.5 * cm, w - 2 * cm, h - 1.5 * cm)
    canv.setFont(BODY_FONT, 9)
    canv.setFillColor(GREY)
    canv.drawString(2 * cm, h - 1.2 * cm,
                    "Conoscenze Basilari su Blockchain, Bitcoin & ETH")
    canv.drawRightString(w - 2 * cm, h - 1.2 * cm, "PROTA DOMENICO")

    # Footer: numero di pagina + linea
    canv.line(2 * cm, 1.6 * cm, w - 2 * cm, 1.6 * cm)
    canv.setFont(BODY_FONT, 9)
    canv.setFillColor(GREY)
    canv.drawString(2 * cm, 1.1 * cm, "© Prota Domenico - Crypto Guide")
    canv.drawRightString(w - 2 * cm, 1.1 * cm,
                         f"Pag. {doc.page - 1}")  # -1 perché la cover è pag.1


# --- Helper per costruire i contenuti ------------------------------------
def p(text, style=BODY):
    return Paragraph(text, style)

def bullets(items, style=BULLET):
    flow = ListFlowable(
        [ListItem(Paragraph(t, style), leftIndent=10) for t in items],
        bulletType="bullet", start="•", leftIndent=18, bulletFontSize=11,
        bulletColor=GOLD,
    )
    return flow

def callout(title, body_text, color=GOLD):
    """Box colorato per consigli/avvisi."""
    table = Table([
        [Paragraph(f"<b>{title}</b>", ParagraphStyle(
            "CT", fontName=BODY_BOLD, fontSize=11, textColor=white))],
        [Paragraph(body_text, ParagraphStyle(
            "CB", fontName=BODY_FONT, fontSize=10.5, leading=15,
            textColor=DARK, alignment=TA_JUSTIFY))]
    ], colWidths=[16 * cm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), color),
        ("TEXTCOLOR", (0, 0), (0, 0), white),
        ("BACKGROUND", (0, 1), (0, 1), LIGHT_BG),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LINEABOVE", (0, 0), (-1, 0), 0, color),
        ("BOX", (0, 0), (-1, -1), 0.5, color),
    ]))
    return KeepTogether([Spacer(1, 6), table, Spacer(1, 8)])


def section_divider():
    t = Table([[""]], colWidths=[16 * cm], rowHeights=[0.06 * cm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), GOLD)]))
    return t


# --- Documento -----------------------------------------------------------
doc = SimpleDocTemplate(
    OUT_PDF, pagesize=A4,
    leftMargin=2 * cm, rightMargin=2 * cm,
    topMargin=2.2 * cm, bottomMargin=2.2 * cm,
    title="Conoscenze Basilari su Blockchain, Bitcoin & ETH",
    author="Prota Domenico",
    subject="Guida introduttiva alle criptovalute",
)

story = []

# ---- Pagina copertina (vuota: la disegna draw_cover) --------------------
story.append(PageBreak())  # forza la fine della copertina

# ---- Pagina: dedica / disclaimer ----------------------------------------
story.append(Spacer(1, 4 * cm))
story.append(p("<i>A chi vuole imparare senza paura,<br/>"
               "una passo alla volta.</i>",
               ParagraphStyle("Ded", parent=BODY, alignment=TA_CENTER,
                              fontSize=14, textColor=NAVY)))
story.append(Spacer(1, 6 * cm))
story.append(p("<b>Avviso importante</b>",
               ParagraphStyle("D1", parent=H3, alignment=TA_CENTER)))
story.append(p(
    "Questo ebook ha finalità puramente <b>informative ed educative</b>. "
    "Non costituisce consulenza finanziaria, fiscale o di investimento. "
    "Le criptovalute sono asset ad alta volatilità: investi solo somme "
    "che puoi permetterti di perdere. L'autore declina ogni responsabilità "
    "per scelte personali derivanti dalla lettura.",
    ParagraphStyle("D2", parent=BODY, alignment=TA_CENTER, fontSize=10,
                   textColor=GREY)))
story.append(PageBreak())

# ---- Indice -------------------------------------------------------------
story.append(p("Indice", H1))
story.append(section_divider())
story.append(Spacer(1, 0.4 * cm))
toc = [
    "1. Introduzione: perché le crypto?",
    "2. Cos'è la Blockchain (spiegata semplice)",
    "3. Bitcoin: il primo e il più famoso",
    "4. Ethereum (ETH): più di una moneta",
    "5. Wallet: il tuo portafoglio digitale",
    "6. Come comprare crypto (passo passo)",
    "7. Come vendere crypto e prelevare in euro",
    "8. Hardware Wallet: la massima sicurezza",
    "9. Guida fai-da-te all'Hard Wallet",
    "10. Errori comuni da evitare",
    "11. Glossario rapido",
    "12. Conclusioni",
]
for item in toc:
    story.append(p(item, TOC_ITEM))
story.append(PageBreak())

# ---- 1. Introduzione ----------------------------------------------------
story.append(p("1. Introduzione: perché le crypto?", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p(
    "Negli ultimi anni la parola <b>criptovaluta</b> è entrata nelle case "
    "di tutti: ne parlano i giornali, gli amici al bar, perfino i nonni. "
    "Eppure, quando si chiede cosa siano davvero, le risposte sono spesso "
    "confuse. Questo ebook nasce per fare chiarezza, con un linguaggio "
    "semplice e diretto, <b>senza tecnicismi inutili</b>."))
story.append(p(
    "Imparerai cos'è la blockchain, come funzionano Bitcoin ed Ethereum, "
    "come si comprano e si vendono le crypto in modo sicuro, e soprattutto "
    "come custodirle con un <b>hardware wallet</b>, il metodo più sicuro "
    "esistente oggi."))
story.append(callout("Promessa del libro",
    "Alla fine di queste pagine saprai muoverti tra blockchain, exchange e "
    "wallet con tranquillità. Non serve essere esperti di informatica: "
    "basta saper usare uno smartphone."))
story.append(p(
    "Le criptovalute non sono solo un investimento: sono una nuova "
    "tecnologia che cambia il modo in cui scambiamo valore, firmiamo "
    "contratti e custodiamo dati. Capirle oggi significa essere preparati "
    "per il domani."))
story.append(PageBreak())

# ---- 2. Blockchain ------------------------------------------------------
story.append(p("2. Cos'è la Blockchain (spiegata semplice)", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p("2.1 L'esempio del quaderno condiviso", H2))
story.append(p(
    "Immagina un <b>quaderno</b> in cui scrivi tutte le transazioni del tuo "
    "paese: «Mario dà 10€ a Luigi», «Anna paga 5€ ad Antonio», e così via. "
    "Adesso immagina che lo stesso quaderno sia <b>copiato e tenuto da migliaia "
    "di persone</b>, in tutto il mondo, contemporaneamente."))
story.append(p(
    "Ogni nuova pagina del quaderno (chiamata <b>blocco</b>) viene aggiunta "
    "solo se la maggioranza delle persone è d'accordo che le transazioni "
    "scritte sono giuste. Una volta scritta, la pagina <b>non può essere "
    "modificata</b>. Ecco, in due righe, cos'è la blockchain: un quaderno "
    "pubblico, condiviso e impossibile da falsificare."))

story.append(p("2.2 Perché si chiama \"catena di blocchi\"", H2))
story.append(p(
    "Ogni blocco contiene un riassunto unico (l'<b>hash</b>) del blocco "
    "precedente. Se qualcuno provasse a modificare una pagina vecchia, "
    "tutti i blocchi successivi si \"romperebbero\" e gli altri "
    "se ne accorgerebbero subito. Per questo si parla di <b>catena</b>: "
    "ogni anello dipende dal precedente."))

story.append(p("2.3 Le caratteristiche chiave", H2))
story.append(bullets([
    "<b>Decentralizzata</b>: nessuna banca, nessun governo la controlla.",
    "<b>Trasparente</b>: chiunque può consultare le transazioni.",
    "<b>Immutabile</b>: ciò che è scritto non si cancella.",
    "<b>Sicura</b>: protetta da crittografia matematica.",
    "<b>Pubblica</b> (in genere): aperta a tutti, 24 ore su 24.",
]))
story.append(callout("In una frase",
    "La blockchain è un registro digitale condiviso da tutti, "
    "controllato da nessuno e modificabile da nessuno."))
story.append(PageBreak())

# ---- 3. Bitcoin ---------------------------------------------------------
story.append(p("3. Bitcoin: il primo e il più famoso", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p("3.1 Le origini", H2))
story.append(p(
    "Bitcoin nasce nel <b>2009</b> grazie a una persona (o gruppo) che "
    "si firma <b>Satoshi Nakamoto</b>. La sua identità è ancora un mistero. "
    "L'obiettivo era creare una moneta digitale che non avesse bisogno di "
    "banche per funzionare."))

story.append(p("3.2 Caratteristiche principali", H2))
story.append(bullets([
    "Quantità <b>limitata</b>: ne esisteranno al massimo 21 milioni.",
    "Si può dividere fino a 8 decimali (1 satoshi = 0,00000001 BTC).",
    "Le transazioni avvengono direttamente da persona a persona (peer-to-peer).",
    "È spesso chiamato <b>oro digitale</b> per la sua scarsità.",
    "Funziona 24/7, in qualsiasi parte del mondo.",
]))

story.append(p("3.3 Cosa si può fare con Bitcoin", H2))
story.append(bullets([
    "Conservare valore nel tempo (riserva di valore).",
    "Inviare denaro all'estero senza intermediari.",
    "Comprare beni e servizi (sempre più aziende lo accettano).",
    "Investire a lungo termine (HODL).",
]))

story.append(callout("Curiosità",
    "Il primo acquisto reale fatto con Bitcoin fu nel 2010: un programmatore "
    "spese 10.000 BTC per due pizze. Oggi sarebbero milioni di euro. "
    "Il 22 maggio si celebra ogni anno il «Bitcoin Pizza Day»."))
story.append(PageBreak())

# ---- 4. Ethereum --------------------------------------------------------
story.append(p("4. Ethereum (ETH): più di una moneta", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p("4.1 Cosa cambia rispetto a Bitcoin", H2))
story.append(p(
    "Ethereum nasce nel <b>2015</b> dall'idea di <b>Vitalik Buterin</b>. "
    "Se Bitcoin è una moneta digitale, Ethereum è qualcosa di più: è una "
    "<b>piattaforma</b> su cui si possono creare programmi (smart contract) "
    "che funzionano da soli, senza intermediari."))

story.append(p("4.2 Smart contract: cosa sono?", H2))
story.append(p(
    "Uno <b>smart contract</b> è un contratto digitale che si esegue "
    "automaticamente quando si verificano certe condizioni. Esempio: «se "
    "Marco paga 100€, allora gli viene consegnato il biglietto del concerto». "
    "Il codice fa tutto da solo, senza bisogno di un notaio o di una banca."))

story.append(p("4.3 ETH, la moneta di Ethereum", H2))
story.append(p(
    "<b>ETH</b> è la valuta che si usa dentro Ethereum per pagare le "
    "operazioni (le cosiddette <i>gas fee</i>, cioè le commissioni di rete). "
    "Si compra e si vende come Bitcoin, ed è la <b>seconda criptovaluta</b> "
    "al mondo per capitalizzazione."))

story.append(p("4.4 Cosa si fa con Ethereum", H2))
story.append(bullets([
    "<b>DeFi</b> (finanza decentralizzata): prestiti, scambi senza banche.",
    "<b>NFT</b>: certificati digitali unici per arte, musica, collezionabili.",
    "<b>DApp</b>: applicazioni che girano sulla blockchain.",
    "<b>Token</b>: chiunque può creare la propria moneta sopra Ethereum.",
]))

story.append(callout("In sintesi",
    "Bitcoin = oro digitale (riserva di valore). "
    "Ethereum = computer mondiale (piattaforma per applicazioni)."))
story.append(PageBreak())

# ---- 5. Wallet ----------------------------------------------------------
story.append(p("5. Wallet: il tuo portafoglio digitale", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p("5.1 Cos'è un wallet", H2))
story.append(p(
    "Un <b>wallet</b> (portafoglio) è uno strumento che ti permette di "
    "<b>ricevere, conservare e inviare</b> criptovalute. Attenzione: il "
    "wallet non contiene le monete come un portafoglio fisico contiene "
    "le banconote. Le monete vivono sulla blockchain. Il wallet contiene "
    "le <b>chiavi</b> che dimostrano che quelle monete sono tue."))

story.append(p("5.2 Chiave pubblica e chiave privata", H2))
story.append(bullets([
    "<b>Chiave pubblica</b> (o indirizzo): è come l'IBAN della banca. "
    "La puoi condividere per ricevere fondi.",
    "<b>Chiave privata</b>: è come il PIN del bancomat. Chi la possiede "
    "controlla i fondi. <b>Non va MAI condivisa.</b>",
    "<b>Seed phrase</b> (12 o 24 parole): è il backup della chiave privata. "
    "Se la perdi, perdi tutto. Se qualcuno la legge, ti può rubare tutto.",
]))

story.append(p("5.3 Tipi di wallet", H2))
data = [
    ["Tipo", "Esempio", "Sicurezza", "Comodità"],
    ["Exchange (custodial)", "Binance, Coinbase, Kraken", "Bassa", "Alta"],
    ["App / Hot wallet", "MetaMask, Trust Wallet", "Media", "Alta"],
    ["Hardware (cold)", "Ledger, Trezor, BitBox", "Massima", "Media"],
    ["Paper wallet", "Foglio stampato", "Alta (se ben conservato)", "Bassa"],
]
tbl = Table(data, colWidths=[3.8*cm, 4.6*cm, 3.6*cm, 3*cm])
tbl.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), NAVY),
    ("TEXTCOLOR", (0, 0), (-1, 0), white),
    ("FONTNAME", (0, 0), (-1, 0), BODY_BOLD),
    ("FONTNAME", (0, 1), (-1, -1), BODY_FONT),
    ("FONTSIZE", (0, 0), (-1, -1), 10),
    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [LIGHT_BG, white]),
    ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ("TOPPADDING", (0, 0), (-1, -1), 6),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ("BOX", (0, 0), (-1, -1), 0.6, NAVY),
    ("INNERGRID", (0, 0), (-1, -1), 0.3, GREY),
]))
story.append(tbl)
story.append(Spacer(1, 0.3 * cm))
story.append(callout("Regola d'oro",
    "Se non possiedi le chiavi private, non possiedi davvero le tue "
    "crypto. <b>«Not your keys, not your coins»</b> dicono in inglese.",
    color=NAVY))
story.append(PageBreak())

# ---- 6. Comprare crypto -------------------------------------------------
story.append(p("6. Come comprare crypto (passo passo)", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p(
    "Per comprare la tua prima crypto bastano pochi passaggi. Vediamoli "
    "uno per uno con calma."))

story.append(p("Passo 1 - Scegli un exchange affidabile", H2))
story.append(p(
    "Un <b>exchange</b> è un sito o un'app dove si comprano e vendono "
    "criptovalute con euro. Scegli un exchange <b>regolamentato</b>, "
    "con buona reputazione e attivo da diversi anni. Esempi noti: "
    "Coinbase, Kraken, Binance, Bitpanda, Young Platform (italiano)."))

story.append(p("Passo 2 - Registrati e verifica l'identità (KYC)", H2))
story.append(p(
    "Dovrai caricare un documento d'identità e fare una breve "
    "videoselfie. È una procedura obbligatoria per legge, si chiama "
    "<b>KYC</b> (Know Your Customer). Serve a prevenire frodi e "
    "riciclaggio."))

story.append(p("Passo 3 - Attiva la sicurezza (2FA)", H2))
story.append(p(
    "Imposta subito l'<b>autenticazione a due fattori</b> (2FA), meglio "
    "se con un'app come Google Authenticator o Authy. Evita la 2FA via "
    "SMS: è meno sicura."))

story.append(p("Passo 4 - Deposita euro", H2))
story.append(bullets([
    "<b>Bonifico SEPA</b>: gratuito o quasi, arriva in 1-2 giorni.",
    "<b>Carta di credito/debito</b>: istantaneo ma con commissioni alte (1-3%).",
    "<b>Apple Pay / Google Pay</b>: rapido, commissioni medie.",
]))

story.append(p("Passo 5 - Compra la tua crypto", H2))
story.append(p(
    "Cerca BTC o ETH nell'app, scegli l'importo in euro, controlla le "
    "commissioni e conferma. La crypto comparirà nel tuo conto exchange "
    "in pochi secondi."))

story.append(callout("Consiglio pratico",
    "Inizia con piccoli importi (es. 20-50€) per familiarizzare con "
    "l'interfaccia. Solo dopo aver imparato bene, considera importi "
    "maggiori. <b>Non investire mai più di quanto puoi permetterti di "
    "perdere.</b>"))

story.append(p("Passo 6 - Sposta le crypto sul tuo wallet", H2))
story.append(p(
    "Lasciare grandi somme sull'exchange è rischioso (gli exchange "
    "possono essere hackerati o falliti). Una volta comprata la crypto, "
    "<b>spostala sul tuo wallet personale</b>, meglio se hardware. "
    "Vedremo come nei prossimi capitoli."))
story.append(PageBreak())

# ---- 7. Vendere crypto --------------------------------------------------
story.append(p("7. Come vendere crypto e prelevare in euro", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p("7.1 La procedura inversa", H2))
story.append(p(
    "Vendere è semplice: dall'exchange selezioni la crypto, scegli "
    "«Vendi», inserisci l'importo e confermi. Gli euro arrivano subito "
    "sul tuo saldo dell'exchange."))

story.append(p("7.2 Prelevare gli euro sul conto bancario", H2))
story.append(bullets([
    "Vai nella sezione <b>«Preleva»</b> o <b>«Withdraw»</b>.",
    "Aggiungi il tuo IBAN (la prima volta serve una breve verifica).",
    "Scegli l'importo e conferma.",
    "Il bonifico SEPA arriva di solito in 1-2 giorni lavorativi.",
]))

story.append(p("7.3 Tasse: cenni rapidi", H2))
story.append(p(
    "In Italia le plusvalenze su criptovalute sono <b>tassate al 26%</b> "
    "(aliquota in vigore al momento di stesura). Esiste una <b>franchigia "
    "di 2.000€</b> di plusvalenze annue sotto la quale non si paga. "
    "Le crypto detenute al 31/12 vanno indicate nel <b>quadro RW</b> "
    "della dichiarazione dei redditi e si paga l'<b>imposta di bollo</b> "
    "(IVAFE) dello 0,2%."))
story.append(callout("Attenzione",
    "Le regole fiscali cambiano nel tempo. Per importi rilevanti "
    "consulta sempre un <b>commercialista</b> esperto in crypto.",
    color=NAVY))

story.append(p("7.4 Quando vendere?", H2))
story.append(p(
    "Non esiste una risposta perfetta. Le strategie comuni sono:"))
story.append(bullets([
    "<b>HODL</b>: tieni la crypto a lungo termine, indipendentemente dal prezzo.",
    "<b>DCA in uscita</b>: vendi piccole quantità a intervalli regolari.",
    "<b>Take profit</b>: vendi quando la crypto raggiunge un obiettivo prefissato.",
    "<b>Stop loss</b>: vendi se il prezzo scende sotto una certa soglia, per limitare le perdite.",
]))
story.append(PageBreak())

# ---- 8. Hardware Wallet -------------------------------------------------
story.append(p("8. Hardware Wallet: la massima sicurezza", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p("8.1 Cos'è e perché serve", H2))
story.append(p(
    "Un <b>hardware wallet</b> (o <b>cold wallet</b>) è un piccolo "
    "dispositivo fisico, simile a una chiavetta USB, che custodisce "
    "le tue chiavi private <b>offline</b>. Le chiavi non escono mai "
    "dal dispositivo, nemmeno quando firma una transazione. Per questo "
    "è praticamente <b>impossibile da hackerare a distanza</b>."))

story.append(p("8.2 Vantaggi", H2))
story.append(bullets([
    "Le chiavi private non toccano mai internet.",
    "Resiste a virus, phishing e malware sul tuo PC/telefono.",
    "Lo schermo del dispositivo mostra cosa stai firmando: niente sorprese.",
    "Funziona anche se il tuo computer è compromesso.",
]))

story.append(p("8.3 Modelli più conosciuti", H2))
data = [
    ["Modello", "Produttore", "Note"],
    ["Ledger Nano S Plus", "Ledger (Francia)", "Economico, molto diffuso"],
    ["Ledger Nano X", "Ledger (Francia)", "Bluetooth, ottimo per mobile"],
    ["Trezor Model One", "SatoshiLabs (CZ)", "Open source, basilare"],
    ["Trezor Safe 3 / 5", "SatoshiLabs (CZ)", "Open source, schermo touch"],
    ["BitBox02", "Shift Crypto (CH)", "Open source, design pulito"],
    ["Coldcard Mk4", "Coinkite", "Solo Bitcoin, focalizzato sicurezza"],
]
tbl2 = Table(data, colWidths=[4.5*cm, 4.5*cm, 6*cm])
tbl2.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), NAVY),
    ("TEXTCOLOR", (0, 0), (-1, 0), white),
    ("FONTNAME", (0, 0), (-1, 0), BODY_BOLD),
    ("FONTNAME", (0, 1), (-1, -1), BODY_FONT),
    ("FONTSIZE", (0, 0), (-1, -1), 10),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [LIGHT_BG, white]),
    ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
    ("TOPPADDING", (0, 0), (-1, -1), 6),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ("BOX", (0, 0), (-1, -1), 0.6, NAVY),
    ("INNERGRID", (0, 0), (-1, -1), 0.3, GREY),
]))
story.append(tbl2)
story.append(Spacer(1, 0.3 * cm))

story.append(callout("Regola fondamentale",
    "Compra l'hardware wallet <b>SOLO dal sito ufficiale del produttore</b>. "
    "Mai da Amazon, eBay o rivenditori sconosciuti: il dispositivo "
    "potrebbe essere stato manomesso prima di arrivarti."))
story.append(PageBreak())

# ---- 9. Guida pratica all'Hard Wallet -----------------------------------
story.append(p("9. Guida fai-da-te all'Hard Wallet", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p(
    "In questo capitolo vediamo passo dopo passo come configurare e usare "
    "in modo sicuro un hardware wallet. La procedura è simile per tutti "
    "i modelli."))

story.append(p("Passo 1 - Acquisto sicuro", H2))
story.append(bullets([
    "Compra <b>solo dal sito ufficiale</b> del produttore.",
    "Verifica che la confezione sia integra: niente sigilli rotti, niente fogli stampati con la seed già scritta.",
    "Se hai dubbi sull'integrità, <b>non usarlo</b> e contatta il supporto.",
]))

story.append(p("Passo 2 - Setup iniziale", H2))
story.append(bullets([
    "Scarica l'app ufficiale (es. Ledger Live, Trezor Suite) <b>solo dal sito ufficiale</b>.",
    "Collega il dispositivo via USB.",
    "Aggiorna il firmware all'ultima versione.",
    "Scegli «Configura come nuovo dispositivo».",
]))

story.append(p("Passo 3 - Imposta il PIN", H2))
story.append(p(
    "Scegli un PIN di almeno 6 cifre. <b>Non usare</b> date di nascita, "
    "1234 o sequenze ovvie. Memorizzalo bene: senza PIN il dispositivo "
    "diventa inutilizzabile dopo qualche tentativo errato."))

story.append(p("Passo 4 - La SEED PHRASE: il momento più importante", H2))
story.append(p(
    "Il dispositivo genererà 12 o 24 parole in inglese, in un ordine "
    "preciso. Questa è la <b>seed phrase</b> (chiamata anche recovery "
    "phrase). È la chiave maestra: chi la possiede ha accesso a tutti "
    "i tuoi fondi."))
story.append(callout("REGOLE D'ORO PER LA SEED",
    "1) Scrivila <b>a mano su carta</b> (mai foto, mai screenshot, mai email, mai cloud).<br/>"
    "2) <b>Non digitarla mai</b> sul PC o sullo smartphone.<br/>"
    "3) Conservala in un posto sicuro, lontano da occhi indiscreti, fuoco e umidità.<br/>"
    "4) Considera di farne <b>2 copie</b> in luoghi diversi (es. casa e una cassetta di sicurezza).<br/>"
    "5) <b>Non condividerla con NESSUNO</b>, neanche con il «supporto tecnico»: chi te la chiede sta cercando di rubarti i fondi.",
    color=GOLD))

story.append(p("Passo 5 - Verifica della seed", H2))
story.append(p(
    "Il dispositivo ti chiederà di reinserire alcune parole per "
    "verificare che le hai trascritte correttamente. <b>Non saltare</b> "
    "questo passaggio: è la tua sola garanzia in caso di smarrimento "
    "o rottura del dispositivo."))

story.append(p("Passo 6 - Soluzioni \"a prova di catastrofe\"", H2))
story.append(p(
    "La carta brucia e si bagna. Per importi importanti valuta le "
    "<b>piastre in metallo</b> dove punzonare le parole della seed. "
    "Sono praticamente indistruttibili: resistono a fuoco, acqua e tempo."))
story.append(bullets([
    "Cryptosteel Capsule",
    "Billfodl",
    "Trezor Keep Metal",
    "Soluzioni fai-da-te con punzoni e piastre d'acciaio",
]))

story.append(p("Passo 7 - Ricevere e inviare crypto", H2))
story.append(p("<b>Per ricevere</b>:"))
story.append(bullets([
    "Apri l'app, scegli la crypto e clicca «Ricevi».",
    "Verifica che l'indirizzo mostrato a schermo coincida con quello sul display del dispositivo. <b>Sempre.</b>",
    "Manda dall'exchange un piccolo importo di prova prima di trasferire grosse somme.",
]))
story.append(p("<b>Per inviare</b>:"))
story.append(bullets([
    "Inserisci l'indirizzo del destinatario nell'app.",
    "Sul display del dispositivo verifica destinatario, importo e commissioni.",
    "Conferma fisicamente premendo il pulsante.",
]))

story.append(p("Passo 8 - La passphrase opzionale (avanzato)", H2))
story.append(p(
    "I principali wallet permettono di aggiungere una <b>passphrase</b>: "
    "una parola in più, scelta da te, che si aggiunge alle 24. "
    "Crea un secondo wallet «nascosto» a partire dalla stessa seed. "
    "È un livello di sicurezza extra contro attacchi fisici, ma "
    "<b>se la dimentichi perdi tutto</b>. Usala solo se sai cosa "
    "stai facendo."))

story.append(p("Passo 9 - Manutenzione", H2))
story.append(bullets([
    "Aggiorna periodicamente il firmware (sempre dall'app ufficiale).",
    "Verifica ogni tanto che il dispositivo si accenda e funzioni.",
    "Non lasciarlo in luoghi visibili o facilmente identificabili.",
    "Se viaggi, considera di portarlo con te o lasciarlo in cassaforte.",
]))
story.append(PageBreak())

# ---- 10. Errori comuni --------------------------------------------------
story.append(p("10. Errori comuni da evitare", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))

errors = [
    ("Lasciare grandi somme sull'exchange",
     "Gli exchange sono comodi, ma se falliscono o vengono hackerati i tuoi fondi possono sparire. Sposta sul tuo wallet ciò che non stai attivamente scambiando."),
    ("Salvare la seed sul PC o nel cloud",
     "Foto, file di testo, email, Google Drive: sono <b>tutti</b> posti vietati. La seed va su carta o metallo, mai in formato digitale."),
    ("Cliccare su link sospetti",
     "Il <b>phishing</b> è la frode più comune nel mondo crypto. Diffida di email, SMS e DM che ti chiedono di accedere al wallet o di «verificare» qualcosa."),
    ("Comprare hardware wallet usati o non ufficiali",
     "Possono essere stati compromessi. Sempre dal sito ufficiale, sempre nuovi, sempre con confezione integra."),
    ("Inseguire le \"100x\" e i progetti sconosciuti",
     "La maggior parte delle crypto piccole e sconosciute va a zero. Se non capisci di cosa si tratta, non comprarlo."),
    ("Investire più di quanto puoi permetterti",
     "Le crypto sono volatili. Investire i risparmi o, peggio, soldi presi a prestito è pericoloso."),
    ("Non verificare gli indirizzi prima di inviare",
     "Le transazioni blockchain sono <b>irreversibili</b>. Un errore di un carattere e i fondi sono persi per sempre. Fai sempre un invio di prova."),
    ("Fidarsi di chi promette guadagni garantiti",
     "Nessuno regala soldi. Trader \"mago\", gruppi Telegram VIP a pagamento, signor X su Instagram: nella stragrande maggioranza dei casi sono <b>truffe</b>."),
]
for title, desc in errors:
    story.append(p(f"<b>• {title}</b>", H3))
    story.append(p(desc, BODY))
story.append(PageBreak())

# ---- 11. Glossario ------------------------------------------------------
story.append(p("11. Glossario rapido", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
glossary = [
    ("Airdrop", "Distribuzione gratuita di token a chi possiede un determinato wallet o criptovaluta."),
    ("Altcoin", "Qualunque criptovaluta diversa da Bitcoin."),
    ("Blockchain", "Registro digitale pubblico, condiviso e immutabile."),
    ("BTC", "Sigla di Bitcoin."),
    ("Cold wallet", "Wallet offline (es. hardware wallet)."),
    ("DCA", "Dollar Cost Averaging: comprare piccole somme a intervalli regolari per ridurre il rischio del timing."),
    ("DeFi", "Finanza decentralizzata: servizi finanziari senza banche."),
    ("ETH", "Sigla di Ether, la moneta di Ethereum."),
    ("Exchange", "Piattaforma per comprare e vendere crypto (es. Binance, Coinbase)."),
    ("Fork", "Aggiornamento o divisione di una blockchain."),
    ("Gas fee", "Commissione di rete pagata per eseguire una transazione su Ethereum."),
    ("Hash", "Impronta digitale unica di un dato; alla base della blockchain."),
    ("HODL", "Tenere a lungo termine; deriva da un errore di battitura per «hold»."),
    ("Hot wallet", "Wallet connesso a internet (app, estensione browser)."),
    ("KYC", "Know Your Customer: verifica d'identità obbligatoria sugli exchange."),
    ("Mining", "Processo che convalida le transazioni in alcune blockchain (es. Bitcoin) e crea nuove monete."),
    ("NFT", "Token non fungibile: certificato digitale unico (arte, collezionabili)."),
    ("Phishing", "Truffa che imita un sito o servizio per rubarti credenziali o seed."),
    ("Satoshi", "Unità minima di Bitcoin (1 sat = 0,00000001 BTC)."),
    ("Seed phrase", "12 o 24 parole che rappresentano il backup del wallet."),
    ("Smart contract", "Programma che gira sulla blockchain ed esegue contratti automaticamente."),
    ("Stablecoin", "Crypto agganciata a una valuta tradizionale (es. USDT, USDC = 1 dollaro)."),
    ("Staking", "Bloccare crypto per ricevere ricompense, simile a un interesse."),
    ("Token", "Asset digitale creato sopra una blockchain (spesso Ethereum)."),
    ("Wallet", "Portafoglio digitale che gestisce le tue chiavi crypto."),
    ("2FA", "Autenticazione a due fattori: secondo livello di sicurezza dopo la password."),
]
for term, definition in glossary:
    story.append(p(f"<b><font color='#F7931A'>{term}</font></b> — {definition}", BODY))
story.append(PageBreak())

# ---- 12. Conclusioni ----------------------------------------------------
story.append(p("12. Conclusioni", H1))
story.append(section_divider())
story.append(Spacer(1, 0.3 * cm))
story.append(p(
    "Sei arrivato fino in fondo: complimenti. Ora hai una base solida "
    "per muoverti nel mondo delle criptovalute con consapevolezza."))
story.append(p(
    "Ricorda i tre pilastri di questa guida:"))
story.append(bullets([
    "<b>Capire</b> prima di comprare. La conoscenza è il miglior investimento.",
    "<b>Custodire</b> in modo sicuro. Le chiavi sono tutto: seed su carta o metallo, mai online.",
    "<b>Calma</b>. Il mondo crypto va veloce ma le decisioni migliori si prendono con freddezza.",
]))
story.append(p(
    "Le criptovalute non sono né una bacchetta magica né una truffa: "
    "sono una tecnologia rivoluzionaria che, come ogni rivoluzione, "
    "premia chi studia e punisce chi corre senza guardare. Vai con calma, "
    "parti piccolo, impara dagli errori (di importo limitato!) e nel "
    "tempo costruirai una conoscenza che pochi hanno."))
story.append(Spacer(1, 1 * cm))
story.append(p("Buon viaggio nel mondo crypto!", ParagraphStyle(
    "Final", parent=H2, alignment=TA_CENTER, textColor=GOLD)))
story.append(Spacer(1, 0.6 * cm))
story.append(p("— Prota Domenico", ParagraphStyle(
    "Sign", parent=BODY, alignment=TA_CENTER, fontSize=13,
    textColor=NAVY)))

# Pagina finale "About"
story.append(PageBreak())
if os.path.exists(LOGO):
    img_logo = Image(LOGO, width=4 * cm, height=4 * cm)
    img_logo.hAlign = "CENTER"
    story.append(Spacer(1, 4 * cm))
    story.append(img_logo)
story.append(Spacer(1, 0.6 * cm))
story.append(p("PROTA DOMENICO", ParagraphStyle(
    "About1", parent=H1, alignment=TA_CENTER, textColor=GOLD)))
story.append(p("Crypto Guide - Edizione Base",
               ParagraphStyle("About2", parent=BODY, alignment=TA_CENTER,
                              fontSize=12, textColor=NAVY)))
story.append(Spacer(1, 1 * cm))
story.append(p(
    "Grazie per aver letto questo ebook. Se ti è stato utile, "
    "condividilo con chi pensa che le crypto siano «troppo difficili». "
    "Forse insieme possiamo cambiare un po' di idee.",
    ParagraphStyle("About3", parent=BODY, alignment=TA_CENTER,
                   fontSize=11, textColor=DARK)))
story.append(Spacer(1, 2 * cm))
story.append(p(
    "© Prota Domenico - Tutti i diritti riservati.<br/>"
    "Questa guida ha finalità educative e non costituisce consulenza finanziaria.",
    ParagraphStyle("About4", parent=BODY, alignment=TA_CENTER,
                   fontSize=9, textColor=GREY)))


# --- Costruzione finale --------------------------------------------------
def first_page(canv, d):
    draw_cover(canv, d)

def later_pages(canv, d):
    draw_page(canv, d)

doc.build(story, onFirstPage=first_page, onLaterPages=later_pages)
print(f"PDF salvato in: {OUT_PDF}")
print(f"Dimensione: {os.path.getsize(OUT_PDF)} byte "
      f"({os.path.getsize(OUT_PDF)/1024:.1f} KB)")
