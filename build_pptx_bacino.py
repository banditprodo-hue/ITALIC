# -*- coding: utf-8 -*-
"""
Generatore PPTX "Gestione del Rischio Cantieristico e Responsabilita' di Bordo"
Convegno operativo - 60 slide, durata 2 ore.

Costruisce un file .pptx valido (OOXML) usando solo la libreria standard,
perche' python-pptx non e' installabile in questo ambiente (rete bloccata).

Palette istituzionale Marina Militare:
- blu Difesa (header/footer)
- deep blue (cover)
- oro (accenti, takeaway, tab)
- rosso (stop / divieti)
- ambra (pretendi / warning)
- verde (ok / azione)
"""
import os
import zipfile

# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------
NAVY        = "0B2E55"   # blu Difesa
DEEPNAVY    = "06203F"   # blu profondo (cover)
MIDNAVY     = "1F4E79"   # blu medio (azione)
GOLD        = "C9A227"   # oro istituzionale
GOLD_LIGHT  = "E6C661"
RED         = "B31B1B"   # stop / divieto
AMBER       = "E6A700"   # pretendi / warning
GREEN       = "2E7D32"   # ok / azione
WHITE       = "FFFFFF"
LIGHTGRAY   = "F2F4F7"
MIDGRAY     = "8C8C8C"
DARKTEXT    = "1A1A1A"
GRAYTEXT    = "4A4A4A"

# Dimensioni slide 16:9 (EMU)
SW = 12192000
SH = 6858000

# ---------------------------------------------------------------------------
# Helper XML
# ---------------------------------------------------------------------------
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def run(text, sz=1800, color=DARKTEXT, bold=False, italic=False,
        font="Calibri"):
    b = ' b="1"' if bold else ' b="0"'
    i = ' i="1"' if italic else ''
    return (
        f'<a:r><a:rPr lang="it-IT" sz="{sz}"{b}{i} dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/></a:rPr>'
        f'<a:t>{esc(text)}</a:t></a:r>'
    )


def para(runs, algn="l", bullet=False, level=0, bullet_color=GOLD,
         space_before=600, space_after=0, line=112000):
    if bullet:
        bu = (f'<a:buClr><a:srgbClr val="{bullet_color}"/></a:buClr>'
              f'<a:buFont typeface="Arial"/><a:buChar char="&#9632;"/>')
        indent = ' marL="285750" indent="-285750"'
    else:
        bu = '<a:buNone/>'
        indent = ' marL="0" indent="0"'
    sb = (f'<a:spcBef><a:spcPts val="{space_before}"/></a:spcBef>'
          f'<a:spcAft><a:spcPts val="{space_after}"/></a:spcAft>')
    ln = f'<a:lnSpc><a:spcPct val="{line}"/></a:lnSpc>'
    ppr = f'<a:pPr{indent} lvl="{level}" algn="{algn}">{ln}{sb}{bu}</a:pPr>'
    return f'<a:p>{ppr}{"".join(runs)}</a:p>'


_uid = [1]
def nid():
    _uid[0] += 1
    return _uid[0]


def reset_ids():
    _uid[0] = 1


def shape_rect(x, y, cx, cy, fill=None, line=None, line_w=12700, name="rect",
               prst="rect"):
    if fill is None:
        fill_xml = '<a:noFill/>'
    else:
        fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    if line:
        line_xml = (f'<a:ln w="{line_w}"><a:solidFill>'
                    f'<a:srgbClr val="{line}"/></a:solidFill></a:ln>')
    else:
        line_xml = '<a:ln><a:noFill/></a:ln>'
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
        f'{fill_xml}{line_xml}</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="it-IT"/>'
        f'</a:p></p:txBody></p:sp>'
    )


def shape_text(x, y, cx, cy, paras, anchor="t", name="txt", wrap="square"):
    bodypr = (f'<a:bodyPr wrap="{wrap}" lIns="91440" tIns="45720" '
              f'rIns="91440" bIns="45720" anchor="{anchor}">'
              f'<a:normAutofit/></a:bodyPr>')
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
        f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f'<p:txBody>{bodypr}<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>'
    )


# ---------------------------------------------------------------------------
# Costruttori di slide
# ---------------------------------------------------------------------------
def wrap_slide(shapes):
    sps = "".join(shapes)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{sps}'
        '</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping '
        'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" '
        'accent2="accent2" accent3="accent3" accent4="accent4" '
        'accent5="accent5" accent6="accent6" hlink="hlink" '
        'folHlink="folHlink"/></p:clrMapOvr></p:sld>'
    )


def header(title, kicker=None, norm=None):
    """Header standard slide content: barra blu profondo con titolo bianco."""
    shapes = []
    # banda blu
    shapes.append(shape_rect(0, 0, SW, 1080000, fill=NAVY, name="header"))
    # tab oro verticale a sinistra
    shapes.append(shape_rect(0, 0, 180000, 1080000, fill=GOLD, name="htab"))
    # barra oro sotto header
    shapes.append(shape_rect(0, 1080000, SW, 50000, fill=GOLD, name="hbar"))
    # kicker (modulo) sopra il titolo
    if kicker:
        shapes.append(shape_text(420000, 130000, 8800000, 280000,
            [para([run(kicker, 1000, GOLD_LIGHT, bold=True)],
                  space_before=0)]))
        ty = 380000
    else:
        ty = 230000
    # titolo
    shapes.append(shape_text(420000, ty, 8800000, 750000,
        [para([run(title, 2400, WHITE, bold=True)], space_before=0,
              line=104000)]))
    # tag riferimento normativo (in alto a destra)
    if norm:
        shapes.append(shape_rect(9380000, 320000, 2620000, 460000,
                                 fill=GOLD, name="normtag"))
        shapes.append(shape_text(9400000, 320000, 2580000, 460000,
            [para([run(norm, 1050, DEEPNAVY, bold=True)], algn="ctr",
                  space_before=0)], anchor="ctr"))
    return shapes


def footer(left_text, slide_no, total):
    shapes = []
    shapes.append(shape_rect(0, SH - 440000, SW, 440000, fill=NAVY,
                             name="footer"))
    shapes.append(shape_rect(0, SH - 440000, SW, 40000, fill=GOLD,
                             name="facc"))
    shapes.append(shape_text(420000, SH - 440000, 9000000, 440000,
        [para([run(left_text, 1000, WHITE)], space_before=0)], anchor="ctr"))
    shapes.append(shape_text(SW - 2200000, SH - 440000, 1780000, 440000,
        [para([run(f"{slide_no} / {total}", 1050, GOLD, bold=True)],
              algn="r", space_before=0)], anchor="ctr"))
    return shapes


# Colori per tipi di box
BOX_STYLES = {
    "regola":   (GOLD,  DEEPNAVY, "REGOLA D'ORO"),
    "stop":     (RED,   WHITE,    "STOP"),
    "vietato":  (RED,   WHITE,    "VIETATO"),
    "pretendi": (AMBER, DEEPNAVY, "PRETENDI"),
    "azione":   (GREEN, WHITE,    "AZIONE C.TE"),
    "rif":      (MIDNAVY, WHITE,  "RIFERIMENTO"),
    "warning":  (AMBER, DEEPNAVY, "ATTENZIONE"),
}


def side_box(bx, by, bw, bh, btype, text):
    """Box laterale colorato (regola, stop, azione, ecc.)."""
    fill, txt_color, label = BOX_STYLES.get(btype, BOX_STYLES["rif"])
    shapes = []
    # ombra/bordo
    shapes.append(shape_rect(bx, by, bw, bh, fill=WHITE, line=fill,
                             line_w=19050, name="boxbg"))
    # header colorato
    shapes.append(shape_rect(bx, by, bw, 420000, fill=fill, name="boxhdr"))
    shapes.append(shape_text(bx + 100000, by, bw - 200000, 420000,
        [para([run(label, 1150, txt_color, bold=True)], space_before=0)],
        anchor="ctr"))
    # testo
    shapes.append(shape_text(bx + 160000, by + 480000, bw - 320000,
                             bh - 580000,
        [para([run(text, 1300, DARKTEXT)], space_before=0, line=120000)]))
    return shapes


def takeaway_band(text):
    """Banda oro a piè pagina sopra il footer (TAKEAWAY)."""
    shapes = []
    by = SH - 1000000
    shapes.append(shape_rect(0, by, SW, 520000, fill=GOLD, name="tk_bg"))
    shapes.append(shape_rect(0, by, 180000, 520000, fill=DEEPNAVY,
                             name="tk_tab"))
    shapes.append(shape_text(280000, by, 1700000, 520000,
        [para([run("TAKEAWAY", 1200, DEEPNAVY, bold=True)],
              space_before=0)], anchor="ctr"))
    shapes.append(shape_text(2050000, by, SW - 2200000, 520000,
        [para([run(text, 1300, DEEPNAVY, bold=True, italic=True)],
              space_before=0)], anchor="ctr"))
    return shapes


def slide_cover(title, subtitle, footer_txt, kicker):
    reset_ids()
    s = []
    # sfondo blu profondo
    s.append(shape_rect(0, 0, SW, SH, fill=DEEPNAVY, name="bg"))
    # banda diagonale superiore (decorativa)
    s.append(shape_rect(0, 0, SW, 90000, fill=GOLD, name="topbar"))
    # banda navy in fascia centrale
    s.append(shape_rect(0, 1900000, SW, 2500000, fill=NAVY, name="band"))
    # tab oro a sinistra
    s.append(shape_rect(0, 1900000, 360000, 2500000, fill=GOLD, name="lbar"))
    # kicker
    s.append(shape_text(900000, 2050000, 10000000, 380000,
        [para([run(kicker, 1400, GOLD_LIGHT, bold=True)], space_before=0)]))
    # titolo
    s.append(shape_text(900000, 2450000, 10400000, 1450000,
        [para([run(title, 4000, WHITE, bold=True)], space_before=0,
              line=104000)]))
    # sottotitolo
    s.append(shape_text(900000, 3950000, 10400000, 500000,
        [para([run(subtitle, 1700, GOLD_LIGHT, italic=True)],
              space_before=0)]))
    # footer riferimenti
    s.append(shape_rect(0, SH - 600000, SW, 600000, fill=NAVY, name="ft"))
    s.append(shape_rect(0, SH - 600000, SW, 60000, fill=GOLD, name="ftacc"))
    s.append(shape_text(900000, SH - 600000, 10400000, 600000,
        [para([run(footer_txt, 1200, WHITE, bold=True)],
              space_before=0)], anchor="ctr"))
    # decorazione: piccolo ancora-marker (cerchio oro)
    s.append(shape_rect(11000000, 800000, 600000, 600000, fill=GOLD,
                        name="emblem", prst="ellipse"))
    s.append(shape_text(11000000, 800000, 600000, 600000,
        [para([run("M.M.", 1300, DEEPNAVY, bold=True)],
              algn="ctr", space_before=0)], anchor="ctr"))
    return wrap_slide(s)


def slide_divider(modulo, title, points, slide_no, total):
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=NAVY, name="bg"))
    # blocco numero modulo
    s.append(shape_rect(0, 0, 3600000, SH, fill=DEEPNAVY, name="numblock"))
    s.append(shape_rect(3600000, 0, 90000, SH, fill=GOLD, name="vbar"))
    s.append(shape_text(300000, 1850000, 3100000, 500000,
        [para([run(modulo, 1500, GOLD, bold=True)], algn="ctr",
              space_before=0)]))
    num = modulo.split()[-1] if modulo.split() else ""
    s.append(shape_text(300000, 2300000, 3100000, 1900000,
        [para([run(num, 12000, GOLD, bold=True)], algn="ctr",
              space_before=0)], anchor="ctr"))
    # titolo modulo
    s.append(shape_text(4050000, 1900000, 7700000, 1500000,
        [para([run(title, 3400, WHITE, bold=True)], space_before=0,
              line=104000)]))
    s.append(shape_rect(4080000, 3450000, 1500000, 50000, fill=GOLD,
                        name="acc"))
    # punti
    pp = [para([run(p, 1700, LIGHTGRAY)], bullet=True, space_before=520,
               bullet_color=GOLD)
          for p in points]
    s.append(shape_text(4050000, 3650000, 7700000, 2400000, pp))
    # footer numero slide
    s.append(shape_text(SW - 1800000, SH - 560000, 1500000, 400000,
        [para([run(f"{slide_no} / {total}", 1100, GOLD)], algn="r",
              space_before=0)], anchor="ctr"))
    return wrap_slide(s)


def slide_content(title, kicker, norm, bullets, box, takeaway,
                  footer_left, slide_no, total):
    """Slide standard 'contenuto'.
    - title, kicker (modulo), norm (tag in alto destra)
    - bullets: lista di stringhe
    - box: tuple (btype, text) o None
    - takeaway: stringa o None
    """
    reset_ids()
    s = []
    # sfondo bianco
    s.append(shape_rect(0, 0, SW, SH, fill=WHITE, name="bg"))
    # header
    s += header(title, kicker=kicker, norm=norm)
    # area corpo
    body_top = 1280000
    body_bottom = SH - 440000 - (520000 if takeaway else 0)
    # Spazio per box laterale
    if box:
        body_w = 7300000
    else:
        body_w = 11300000
    bp = [para([run(b, 1650, DARKTEXT)], bullet=True, bullet_color=GOLD,
               space_before=480, line=118000) for b in bullets]
    s.append(shape_text(470000, body_top, body_w, body_bottom - body_top, bp))
    # box laterale
    if box:
        btype, btext = box
        bx, by = 7900000, body_top
        bw = SW - bx - 380000
        bh = body_bottom - body_top
        s += side_box(bx, by, bw, bh, btype, btext)
    # takeaway band
    if takeaway:
        s += takeaway_band(takeaway)
    # footer
    s += footer(footer_left, slide_no, total)
    return wrap_slide(s)


def slide_riepilogo(modulo_num, quote, next_module, footer_left,
                    slide_no, total):
    """Slide di riepilogo modulo: quote centrale e 'PROSSIMO MODULO'."""
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=DEEPNAVY, name="bg"))
    # banda oro orizzontale
    s.append(shape_rect(0, 1100000, SW, 70000, fill=GOLD, name="band"))
    # kicker
    s.append(shape_text(700000, 600000, 10800000, 450000,
        [para([run(f"RIEPILOGO MODULO {modulo_num}", 1500, GOLD,
                   bold=True)], space_before=0)]))
    # icona virgolette
    s.append(shape_text(700000, 1500000, 700000, 800000,
        [para([run('"', 7000, GOLD, bold=True)], space_before=0)],
        anchor="t"))
    # quote
    s.append(shape_text(1500000, 1700000, 10000000, 2600000,
        [para([run(quote, 2600, WHITE, italic=True, bold=True)],
              space_before=0, line=126000)], anchor="t"))
    # banda prossimo modulo
    by = SH - 1500000
    s.append(shape_rect(0, by, SW, 500000, fill=NAVY, name="nxt_bg"))
    s.append(shape_rect(0, by, 180000, 500000, fill=GOLD, name="nxt_tab"))
    s.append(shape_text(280000, by, 2400000, 500000,
        [para([run("PROSSIMO MODULO", 1200, GOLD, bold=True)],
              space_before=0)], anchor="ctr"))
    s.append(shape_text(2700000, by, SW - 2800000, 500000,
        [para([run(next_module, 1500, WHITE, bold=True)],
              space_before=0)], anchor="ctr"))
    # footer
    s += footer(footer_left, slide_no, total)
    return wrap_slide(s)


def slide_qa(title, subtitle, stimolo, takeaway, footer_left,
             slide_no, total):
    """Slide finale Q&A."""
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=DEEPNAVY, name="bg"))
    s.append(shape_rect(0, 0, SW, 90000, fill=GOLD, name="topbar"))
    # grande Q&A
    s.append(shape_text(700000, 700000, 10800000, 1500000,
        [para([run(title, 6000, GOLD, bold=True)], space_before=0)]))
    # sottotitolo
    s.append(shape_text(700000, 2200000, 10800000, 600000,
        [para([run(subtitle, 2000, WHITE)], space_before=0)]))
    # box stimolo
    bx, by, bw, bh = 700000, 3100000, 10800000, 1700000
    s.append(shape_rect(bx, by, bw, bh, fill=NAVY, line=GOLD,
                        line_w=19050, name="stim"))
    s.append(shape_rect(bx, by, 180000, bh, fill=GOLD, name="stim_tab"))
    s.append(shape_text(bx + 300000, by + 130000, bw - 400000, 380000,
        [para([run("DOMANDA STIMOLO", 1100, GOLD, bold=True)],
              space_before=0)]))
    s.append(shape_text(bx + 300000, by + 550000, bw - 400000,
                        bh - 600000,
        [para([run(stimolo, 1500, WHITE, italic=True)],
              space_before=0, line=124000)]))
    # takeaway
    s += takeaway_band(takeaway)
    s += footer(footer_left, slide_no, total)
    return wrap_slide(s)


# ---------------------------------------------------------------------------
# Scaffolding OOXML (theme, master, layout, package)
# ---------------------------------------------------------------------------
THEME = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="MarinaConvegno">
<a:themeElements>
<a:clrScheme name="MarinaConvegno">
<a:dk1><a:srgbClr val="0B2E55"/></a:dk1>
<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="06203F"/></a:dk2>
<a:lt2><a:srgbClr val="F2F4F7"/></a:lt2>
<a:accent1><a:srgbClr val="C9A227"/></a:accent1>
<a:accent2><a:srgbClr val="0B2E55"/></a:accent2>
<a:accent3><a:srgbClr val="B31B1B"/></a:accent3>
<a:accent4><a:srgbClr val="E6A700"/></a:accent4>
<a:accent5><a:srgbClr val="2E7D32"/></a:accent5>
<a:accent6><a:srgbClr val="1F4E79"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>'''

SLIDE_MASTER = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr></p:titleStyle>
<p:bodyStyle><a:lvl1pPr marL="285750" indent="-285750"><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr>
<a:lvl2pPr marL="571500" indent="-285750"><a:defRPr sz="1600"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl2pPr></p:bodyStyle>
<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>
</p:sldMaster>'''

SLIDE_LAYOUT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
<p:cSld name="Vuoto"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld><p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>
</p:sldLayout>'''


def build_pptx(slides_xml, out_path, title):
    n = len(slides_xml)
    overrides = [
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for i in range(1, n + 1):
        overrides.append(
            f'<Override PartName="/ppt/slides/slide{i}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides) + '</Types>')

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>')

    sldid = "".join(
        f'<p:sldId id="{255 + i}" r:id="rId{i + 1}"/>' for i in range(n))
    presentation = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId_master"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{sldid}</p:sldIdLst>'
        f'<p:sldSz cx="{SW}" cy="{SH}" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>')

    pres_rels_items = [
        '<Relationship Id="rId_master" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>',
        '<Relationship Id="rId_theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
        '<Relationship Id="rId_props" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>',
    ]
    for i in range(n):
        pres_rels_items.append(
            f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i + 1}.xml"/>')
    pres_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(pres_rels_items) + '</Relationships>')

    master_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
        '</Relationships>')

    layout_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        '</Relationships>')

    presProps = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')

    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        f'<dc:title>{esc(title)}</dc:title>'
        '<dc:creator>Kiro</dc:creator>'
        '<cp:keywords>Bacino di Carenaggio; Marinarsen; SMM-PREVA-1062; PSBM001A; NAV-G-001; D.Lgs. 272/99</cp:keywords>'
        '</cp:coreProperties>')

    app = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Kiro OOXML Generator</Application>'
        f'<Slides>{n}</Slides></Properties>')

    slide_rel = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>')

    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)
        z.writestr("ppt/presentation.xml", presentation)
        z.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
        z.writestr("ppt/presProps.xml", presProps)
        z.writestr("ppt/theme/theme1.xml", THEME)
        z.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout_rels)
        for i, xml in enumerate(slides_xml, start=1):
            z.writestr(f"ppt/slides/slide{i}.xml", xml)
            z.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", slide_rel)
    return out_path



# ===========================================================================
# CONTENUTI DELLE 60 SLIDE
# ===========================================================================
TOTAL = 60
FOOT  = ("Gestione del Rischio Cantieristico e Responsabilita' di Bordo  |  "
         "SMM-PREVA-1062  -  PSBM001A  -  NAV-G-001  -  D.Lgs. 272/99")

# Ciascun elemento e' una dict:
#   t = titolo
#   k = kicker (modulo)
#   n = riferimento normativo (tag in alto a dx, opzionale)
#   b = bullet list
#   box = (tipo, testo)  oppure None
#   tk  = takeaway (stringa) oppure None

MOD1 = "MODULO 1  -  QUADRO COGENTE E MINDSET"
MOD2 = "MODULO 2  -  LA TIMELINE OPERATIVA"
MOD3 = "MODULO 3  -  IL GIORNO DELL'INGRESSO"
MOD4 = "MODULO 4  -  IN BACINO: RUBRICA E CONSEGNE"
MOD5 = "MODULO 5  -  CANTIERE VIVO, NORMATIVE E INTERFERENZE"
MOD6 = "MODULO 6  -  CHIUSURA"


SLIDES = [
    # ------------------- MODULO 1 -------------------
    # Slide 2 - Obiettivi
    {"t":"Obiettivi Formativi", "k":MOD1, "n":None,
     "b":["Saper gestire la timeline di ingresso in bacino (T-60 -> T-0)",
          "Conoscere la rubrica telefonica e i ruoli dell'Arsenale",
          "Delimitare il confine tra Alta Vigilanza e Ingerenza di fatto",
          "Proteggersi legalmente in caso di infortunio di personale esterno"],
     "box":None,
     "tk":"Zero teoria, solo operativita'."},

    # Slide 3 - Agenda
    {"t":"Agenda delle 2 Ore", "k":MOD1, "n":"2 h",
     "b":["Modulo 1: Quadro Cogente e Responsabilita'  -  20 min",
          "Modulo 2: La Timeline e l'Ingresso  -  25 min",
          "Modulo 3: In Bacino: Rubrica, Ruoli e Consegne  -  30 min",
          "Modulo 4: Cantiere Vivo, Normative e Interferenze  -  30 min",
          "Modulo 5: Emergenze, Uscita e Q&A  -  15 min"],
     "box":None, "tk":None},

    # Slide 4 - Fonti normative
    {"t":"Le Fonti Normative \"Salva-Vita\"", "k":MOD1, "n":"4 Fonti",
     "b":["D.Lgs. 272/1999: norme specifiche cantieristica navale portuale",
          "SMM-PREVA-1062 (Ed. Gen 2024, Titolo VII): istruzioni operative "
          "Marina (Art. 102-120)",
          "PSBM001A: procedura locale Marinarsen Taranto (Bacini BRIN e "
          "FERRATI)",
          "NAV-G-001 (Art. 37-43): regole d'ingaggio nave-bacino"],
     "box":("regola", "Queste 4 fonti sono la tua assicurazione sulla "
                       "vita. Tienile sempre in plancia."),
     "tk":"Queste 4 fonti sono la tua assicurazione sulla vita."},

    # Slide 5 - Ruolo Comandante
    {"t":"Il Ruolo del Comandante", "k":MOD1, "n":"Art. 113 SMM-PREVA-1062",
     "b":["NON dirige i lavori, NON da' ordini agli operai privati",
          "Esercita l'ALTA VIGILANZA sul rispetto delle norme di sicurezza",
          "Verifica che le ditte mitighino le interferenze con il bordo",
          "Fornisce informazioni sui rischi specifici (DVR, RTVR)"],
     "box":("regola", "Se intervieni operativamente, diventi "
                       "co-responsabile."),
     "tk":None},

    # Slide 6 - Confine giuridico
    {"t":"Il Confine Giuridico", "k":MOD1, "n":"Art. 113 c.8",
     "b":["ALTA VIGILANZA (Dovere): designare Preposti, fornire DVR, "
          "interrompere solo per grave pericolo",
          "INGERENZA (Errore): dire all'operaio \"come\" fare, spostare "
          "attrezzature, dare ordini diretti",
          "La linea di demarcazione e' la delega formale ai Preposti",
          "Tutto deve essere tracciato per iscritto"],
     "box":("warning", "Vigilare e' un dovere. Sostituirsi all'impresa "
                        "e' un reato di ingerenza."),
     "tk":"Vigila, non sostituire. Il confine e' la delega formale."},

    # Slide 7 - Catena di comando ditte
    {"t":"La Catena di Comando delle Ditte", "k":MOD1, "n":"D.Lgs. 272/99",
     "b":["Le ditte rispondono all'IMPRESA CAPO COMMESSA",
          "Il referente unico e' il RTL (Responsabile Tecnico dei Lavori)",
          "Il Comandante si interfaccia con RTL o Preposto ditta, MAI con "
          "l'operaio",
          "Se l'impresa sbaglia -> responsabile e' il RTL / Datore di "
          "lavoro"],
     "box":("stop", "Mai dare ordini diretti all'operaio. Chiama "
                     "sempre il RTL."),
     "tk":"Mai dare ordini diretti: chiama il RTL."},

    # Slide 8 - Documenti del Comandante
    {"t":"I Documenti del Comandante", "k":MOD1, "n":"DVR / RTVR",
     "b":["DVR (Documento Valutazione Rischi): la bibbia del Comandante",
          "RTVR (Relazione Tecnica Valutazione Rischi): viaggia con la nave",
          "Descrive rischi strutturali, ambienti confinati, interferenze",
          "Il Comandante FORNISCE questi documenti alle ditte per "
          "POS / DUVRI"],
     "box":("pretendi", "Senza RTVR aggiornata e consegnata, le ditte "
                         "non possono iniziare a lavorare."),
     "tk":"Senza RTVR aggiornata, le ditte non possono lavorare."},

    # Slide 9 - Documenti coordinamento
    {"t":"Documenti di Coordinamento", "k":MOD1, "n":"Art. 116 / DUVRI / PSC",
     "b":["DUVRI: valuta i rischi da interferenza nave-ditte",
          "PSC: obbligatorio per cantieri complessi (soste lavori cicliche)",
          "Devono essere firmati e presenti a bordo PRIMA dell'inizio "
          "lavori",
          "Il RTL aggiorna il Documento di Sicurezza (Art. 116)"],
     "box":("stop", "Nessun documento firmato = Nessun lavoro autorizzato."),
     "tk":"Nessun documento firmato = Nessun lavoro autorizzato."},

    # ------------------- MODULO 2 -------------------
    # Slide 11 - T-60
    {"t":"T-60 Giorni: La Richiesta", "k":MOD2, "n":"PSBM001A §4.1",
     "b":["Ufficio Programma -> Sezione Bacini: richiesta formale ingresso",
          "Valutazione preliminare: tempi approntamento piano di posa",
          "Verifica risorse disponibili (manovalanza, materiali)",
          "Approvazione preliminare della finestra temporale"],
     "box":("azione", "Il Comandante verifica che l'Ufficio Programma "
                       "abbia trasmesso la richiesta entro T-60."),
     "tk":None},

    # Slide 12 - T-30/T-15
    {"t":"T-30 e T-15: I Disegni e gli Assetti", "k":MOD2,
     "n":"PSBM001A §4.2 / §5.1",
     "b":["T-30: invio disegno piano nominale taccate al Capo Nucleo NBM",
          "T-15: contatto vie brevi (C.te Piattaforma / DM) con Capo "
          "Nucleo NBM per assetti",
          "Verifica anomalie scafo, sporgenze sotto chiglia, immersioni",
          "Approvazione assetto nominale longitudinale e trasversale"],
     "box":("rif", "PSBM001A paragrafi 4.2 e 5.1. Il Capo Nucleo NBM e' "
                    "il tuo interlocutore tecnico."),
     "tk":None},

    # Slide 13 - T-7
    {"t":"T-7 Giorni: Riunione Obbligatoria", "k":MOD2,
     "n":"PSBM001A §5.2",
     "b":["Si definisce il dettaglio della manovra di ingresso",
          "Partecipanti: Ufficio Programma, Comando Bordo, Capo Sez. "
          "Bacini, NBM, MARISTANAV",
          "Eventuale incontro successivo T-2 per dettagli finali",
          "Verbale di riunione firmato da tutte le parti"],
     "box":("azione", "Presenziare di persona o delegare un Ufficiale con "
                       "potere decisionale e mandato scritto."),
     "tk":None},

    # Slide 14 - Go/No-Go
    {"t":"Il Go / No-Go del Comandante", "k":MOD2, "n":"NAV-G-001 Art. 37",
     "b":["Sbarco di munizioni, esplosivi e combustibili pericolosi",
          "Prosciugamento totale delle sentine",
          "Rizzaggio dei materiali mobili (soprattutto trincarini)",
          "Chiusura della portelleria esterna fino a messa a secco",
          "Pulizia e Gas-Free dei depositi nafta (se previsti lavori a "
          "fuoco)"],
     "box":("stop", "Senza queste operazioni la Sezione Bacini blocca "
                     "l'ingresso. Nessuna eccezione."),
     "tk":"Senza queste operazioni = Sezione Bacini blocca l'ingresso."},

    # Slide 15 - Assetto e segnalazioni
    {"t":"Assetto e Segnalazioni", "k":MOD2, "n":"PSBM001A §5.3",
     "b":["Regolazione assetto longitudinale/trasversale secondo "
          "indicazioni Capo Nucleo NBM",
          "Scopo: ridurre la rotazione della nave sulle taccate in fase "
          "di appoggio",
          "Comunicare a Direzione Arsenale: immersioni, carichi mobili",
          "Segnalare sporgenze sotto chiglia e qualunque anomalia di "
          "scafo"],
     "box":("pretendi", "Verifica scritta dell'assetto concordato prima "
                         "di entrare in bacino."),
     "tk":None},

    # Slide 16 - Check-list PSBM001A.A01
    {"t":"La Check-List PSBM001A.A01", "k":MOD2, "n":"27 step",
     "b":["Documento in 27 step che accompagna l'intera manovra",
          "Include: verifica impianti A.I. e prosciugamento, pulizia "
          "platea, assetto, ruoli",
          "Step 13 (critico): Comando del Comandante",
          "Step 17 (critico): Passaggio comando all'Arsenale"],
     "box":("azione", "DM o Ufficiale di guardia compila e firma gli "
                       "step di competenza, in tempo reale."),
     "tk":None},

    # ------------------- MODULO 3 -------------------
    # Slide 18 - Imboccatura
    {"t":"Imboccatura: Le Regole", "k":MOD3, "n":"NAV-G-001 Art. 38",
     "b":["La nave imbocca la porta del bacino in galleggiamento",
          "DIVIETO ASSOLUTO di fare uso delle eliche",
          "Si procede ESCLUSIVAMENTE con i tonneggi",
          "I rimorchiatori di MARISTANAV NON entrano in bacino"],
     "box":("vietato", "Niente eliche, solo tonneggi. Niente "
                        "rimorchiatori dentro il bacino."),
     "tk":"VIETATO: niente eliche, solo tonneggi."},

    # Slide 19 - Posizionamento
    {"t":"Posizionamento in Bacino", "k":MOD3, "n":"Step 15 PSBM001A.A01",
     "b":["La nave si porta sui riferimenti visivi del Nucleo NBM",
          "Il Comando di Bordo garantisce l'assetto concordato",
          "Verifica visiva con personale palombaro",
          "Comunicazione costante via VHF con Capo Nucleo NBM"],
     "box":("pretendi", "L'assetto e' tuo dovere fino al passaggio "
                         "ufficiale di comando."),
     "tk":None},

    # Slide 20 - Passaggio di comando (overview)
    {"t":"Il Passaggio di Comando: Momento Critico", "k":MOD3,
     "n":"3 Fasi",
     "b":["E' il punto in cui si giocano le responsabilita' penali",
          "Fase A: nave in galleggiamento -> Comando = COMANDANTE",
          "Fase B: installazione paranchi -> preparazione passaggio",
          "Fase C: comunicazione VHF -> Comando = CAPO NUCLEO NBM"],
     "box":("warning", "Questo momento si gioca in tribunale. "
                        "Documenta tutto."),
     "tk":"Questo momento si gioca in tribunale."},

    # Slide 21 - Fase A
    {"t":"Fase A: Comando del Comandante", "k":MOD3, "n":"Galleggiamento",
     "b":["La nave galleggia sui propri cavi di tonneggio",
          "Comando e Controllo = COMANDANTE",
          "Ogni decisione e' del C.te",
          "Le squadre di bordo sono al loro posto di manovra"],
     "box":("azione", "Verifica che il libro di manovra registri data, "
                       "ora e ufficiali in plancia."),
     "tk":None},

    # Slide 22 - Fase B
    {"t":"Fase B: Preparazione", "k":MOD3, "n":"Paranchi",
     "b":["Il personale Nucleo NBM installa i paranchi di centratura",
          "Collega i paranchi agli argani di terra",
          "Il personale di bordo collega l'altro capo alla nave",
          "Nessuno ha ancora il comando esclusivo della nave"],
     "box":("warning", "Zona grigia: vigila con attenzione, ogni gesto "
                        "puo' essere contestato."),
     "tk":None},

    # Slide 23 - Fase C
    {"t":"Fase C: Il Passaggio Ufficiale", "k":MOD3, "n":"VHF",
     "b":["Il Capo Nucleo NBM chiama il Comandante via VHF",
          "Messaggio: \"Nave [Nome], paranchi in tensione, comando e "
          "controllo passa all'Arsenale\"",
          "Da quel secondo: Comando = CAPO NUCLEO NBM",
          "Registrazione VHF e annotazione sul giornale di bordo"],
     "box":("pretendi", "Pretendi che la comunicazione sia chiara, "
                         "registrata e nota a tutti."),
     "tk":"Pretendi: chiarezza, registrazione, conoscenza condivisa."},

    # Slide 24 - Prosciugamento
    {"t":"Prosciugamento e Appoggio", "k":MOD3,
     "n":"Step 15-20 PSBM001A.A01",
     "b":["Prosciugamento del bacino in modo controllato",
          "Stop a circa 2 metri dal piano di posa (Step 17)",
          "Ispezione pagliolo e gargame a cura dei palombari "
          "(Step 15, 19)",
          "Ripresa prosciugamento -> appoggio sulle taccate",
          "Operazioni subacquee di centratura e civatura (Step 19-20)"],
     "box":("rif", "Annotare gli step compiuti con orario e firma del "
                    "responsabile."),
     "tk":None},

    # Slide 25 - Uscita
    {"t":"L'Uscita: Il Comando Torna a Te", "k":MOD3, "n":"Galleggiamento",
     "b":["Durante l'uscita la responsabilita' torna al Comandante nel "
          "momento in cui la nave va in galleggiamento",
          "Non appena la nave \"spinge\" sull'acqua, il comando e' tuo",
          "Comunicazione VHF speculare a quella di ingresso",
          "Annotazione su giornale di bordo con orario preciso"],
     "box":("regola", "Galleggiamento = Comando tuo. Dal secondo zero."),
     "tk":"Galleggiamento = comando tuo."},

    # Slide 26 - Divieto pesi
    {"t":"Divieto Assoluto: Spostare Pesi", "k":MOD3, "n":"NAV-G-001 Art. 40",
     "b":["VIETATO vuotare/riempire caldaie o spostare pesi senza "
          "autorizzazione Direzione Arsenale",
          "Spostare zavorre o carburanti puo' far ribaltare la nave "
          "sulle taccate",
          "In caso di imbarco/sbarco materiali -> accordo preventivo con "
          "Marinarsen",
          "Ogni operazione va registrata per iscritto"],
     "box":("vietato", "Niente pesi in movimento senza permesso scritto "
                        "della Direzione Arsenale."),
     "tk":"VIETATO: niente pesi in movimento senza permesso."},

    # Slide 27 - Messa a terra
    {"t":"Messa a Terra dello Scafo", "k":MOD3, "n":"NAV-G-001 Art. 41",
     "b":["Appena la nave e' a secco, il C.te accerta la messa a terra "
          "dello scafo",
          "Marinarsen fornisce conduttore a contatto metallico con scafo",
          "Terminale collegato al punto di terra del bacino",
          "Verifica collegamento prima di qualsiasi lavoro elettrico"],
     "box":("pretendi", "Foto e firma del collegamento di terra. "
                         "Sempre, prima del primo lavoro a bordo."),
     "tk":None},

    # Slide 28 - Sentinella
    {"t":"Servizio di Sentinella", "k":MOD3, "n":"NAV-G-001 Art. 42",
     "b":["La nave in bacino stabilisce una sentinella sulla porta",
          "Ordine: vigilare che nessuno si soffermi, depositi materiali "
          "o manometta valvole",
          "Se la nave e' senza equipaggio -> vigilanza a cura di "
          "Marinarsen",
          "Cambio guardie e consegne registrati su brogliaccio"],
     "box":("azione", "Consegne scritte e lette ad alta voce al cambio "
                       "guardia. Sempre."),
     "tk":None},

    # Slide 29 - Passerelle
    {"t":"Passerelle: Regole di Transito", "k":MOD3, "n":"NAV-G-001 Art. 43",
     "b":["Scale, androni e plance assicurate con legature di cavo "
          "(bordo + sponda bacino)",
          "VIETATO: affollamento eccessivo, stazionamento, transito a "
          "passo cadenzato",
          "VIETATO: trasporto materiali con peso o ingombro eccessivo",
          "Cartello monitore: numero max persone, distanziate almeno "
          "1 metro"],
     "box":("vietato", "Niente passo cadenzato. Niente affollamento. "
                        "Mai oltre il numero max."),
     "tk":None},

    # Slide 30 - Barriera acque
    {"t":"La Barriera Acque di Processo", "k":MOD3, "n":"PSBM001A §6.1",
     "b":["La barriera separa zona barcaporta (pozzetti) da zona cantiere",
          "Scopo: evitare contaminazione acque mare con acque di processo",
          "Responsabilita' corretto assetto: Comando di Bordo / "
          "Ufficio Programma",
          "Supervisione che le ditte rispettino la barriera"],
     "box":("azione", "Verifica visiva quotidiana dell'integrita' della "
                       "barriera. Annotala."),
     "tk":None},

    # Slide 31 - Rifiuti
    {"t":"Rifiuti in Platea e Banchina", "k":MOD3, "n":"D.Lgs. 152/06",
     "b":["Responsabilita': Comando di Bordo / Ufficio Programma",
          "Supervisione che le ditte rispettino le prescrizioni "
          "contrattuali",
          "DIVIETO assoluto di abbandono rifiuti in platea e banchina",
          "Tracciabilita' smaltimento con FIR / registri"],
     "box":("vietato", "Nessun rifiuto abbandonato. Mai. Foto-evidenza "
                        "e segnalazione al RTL."),
     "tk":None},

    # Slide 32 - Impianti di bacino
    {"t":"Impianti di Bacino: Chi Fa Cosa", "k":MOD3, "n":"Verbale A02",
     "b":["Impianto Antincendio - Funzionamento: Sezione Bacini",
          "Impianto Antincendio - Uso e Regolazione: Comando di Bordo "
          "(SAP)",
          "Prosciugamento platea - Funzionamento e Uso: Sezione Bacini",
          "Alimentazione elettrica - Funzionamento: Sezione Reti"],
     "box":("regola", "Funzionamento = Marinarsen. Uso operativo = Bordo. "
                       "Non confonderli."),
     "tk":None},

    # Slide 33 - Preposti
    {"t":"I Tuoi Rappresentanti (Preposti)", "k":MOD3,
     "n":"Art. 113 SMM-PREVA-1062",
     "b":["Il C.te designa rappresentanti per l'alta vigilanza",
          "Devono avere DPI adeguati ai rischi del cantiere",
          "Seguono gli spostamenti delle ditte all'interno della nave",
          "NON interferiscono con il lavoro della ditta, ma segnalano "
          "al RTL"],
     "box":("azione", "Designa per iscritto. Forma. Dota di DPI. "
                       "Verifica."),
     "tk":None},

    # ------------------- MODULO 4 -------------------
    # Slide 35 - Rubrica
    {"t":"La Rubrica da Tenere in Plancia", "k":MOD4, "n":"H24 / Orario",
     "b":["Guardia Bacini (H24): 22916  -  emergenze, notte e festivi",
          "Capo Nucleo Bacini in Muratura: 23762  -  problemi operativi",
          "Segreteria Bacini in Muratura: 23765",
          "Sezione Reti Marinarsen: 22064  -  guasti alimentazione "
          "elettrica",
          "Sezione Sanitaria: 22841  -  infortuni in orario "
          "(fuori orario -> 118)"],
     "box":("pretendi", "Stampa questa slide e tienila plastificata in "
                         "plancia e in alloggio C.te."),
     "tk":"Stampala. Plastificala. Tienila in plancia."},

    # Slide 36 - Verbale PSBM001A.A02
    {"t":"Il Verbale PSBM001A.A02", "k":MOD4, "n":"PSBM001A.A02",
     "b":["Da firmare entro il giorno successivo all'appoggio, PRIMA "
          "dell'avvio lavori",
          "Passaggio di consegne tra Sezione Bacini e Comando di Bordo",
          "Senza questo verbale firmato = non si entra in cantiere",
          "Prevede anche la cessione in uso dell'impianto antincendio"],
     "box":("stop", "Niente firma = niente lavori. Vale per tutti, "
                     "anche per il RTL."),
     "tk":None},

    # Slide 37 - Alimentazione terra
    {"t":"Alimentazione da Terra", "k":MOD4, "n":"Sezione Reti  22064",
     "b":["Quadro di alimentazione: identificativo ______________",
          "Tensione: ______ V   |   Potenza prelevabile: ______ kW",
          "Verifica OBBLIGATORIA della messa a terra del quadro di "
          "banchina",
          "In caso di guasto -> Sezione Reti 22064 o Guardia Bacini 22916"],
     "box":("azione", "Compila i campi mancanti con valori reali al "
                       "momento della consegna."),
     "tk":None},

    # Slide 38 - Cessione antincendio
    {"t":"Cessione Impianto Antincendio", "k":MOD4, "n":"SAP di bordo",
     "b":["Marinarsen cede in uso l'impianto fisso antincendio",
          "Responsabile SAP di bordo autorizzato all'uso",
          "Obbligo: cuffie antirumore nel locale pompe",
          "Obbligo: annotare giornalmente le ore di funzionamento sul "
          "registro"],
     "box":("pretendi", "Registro ore-pompe firmato giornalmente dal SAP "
                         "di bordo."),
     "tk":None},

    # Slide 39 - 4 Elettropompe
    {"t":"Le 4 Elettropompe", "k":MOD4, "n":"4 x 250 mc/h",
     "b":["Elettropompa 1 da 250 mc/h: SI / NO",
          "Elettropompa 2 da 250 mc/h: SI / NO",
          "Elettropompa 3 da 250 mc/h: SI / NO",
          "Elettropompa 4 da 250 mc/h: SI / NO",
          "Manichette collegate n. ____ presso punti di sbocco ____"],
     "box":("azione", "Verifica fisica dello stato di ciascuna pompa "
                       "prima della firma del verbale."),
     "tk":"Annota ogni giorno le ore di funzionamento."},

    # Slide 40 - Passerelle
    {"t":"Consegna Passerelle", "k":MOD4, "n":"NAV-G-001 Art. 43",
     "b":["Passerelle consegnate n. matricola: ______________",
          "Carico massimo previsto: ______ kg",
          "Passaggio simultaneo massimo: n. _____ persone",
          "Obbligo: distanziate almeno 1 metro l'una dall'altra"],
     "box":("warning", "Cartello monitore obbligatorio all'imbocco di "
                        "ogni passerella."),
     "tk":None},

    # Slide 41 - Allagamento bordo
    {"t":"Allagamento a Bordo", "k":MOD4, "n":"Guardia Bacini 22916",
     "b":["VIETATO scaricare in platea qualsiasi liquido",
          "In caso di allagamento locali di bordo -> chiamare Guardia "
          "Bacini 22916",
          "Marinarsen predispone la raccolta acque di processo",
          "Documentare l'evento con foto e relazione scritta"],
     "box":("vietato", "Mai scaricare in platea. Mai. Anche acqua "
                        "limpida."),
     "tk":"VIETATO: mai scaricare in platea."},

    # Slide 42 - Evacuazione
    {"t":"Evacuazione dall'Unita'", "k":MOD4,
     "n":"Punto di Raccolta n. 11",
     "b":["Punto di Raccolta: n. 11  -  Zona Raccolta FERRATI",
          "Planimetria consegnata con verbale PSBM001A.A02",
          "Trasporto infortunato fino all'ambulanza -> a cura del "
          "Comando di Bordo",
          "Verifica delle vie di esodo prima dell'avvio lavori"],
     "box":("azione", "Affiggi planimetria di evacuazione in plancia e "
                       "ai posti di guardia."),
     "tk":None},

    # Slide 43 - Infortunio
    {"t":"Infortunio: Chi Chiama Chi", "k":MOD4, "n":"22841 / 118",
     "b":["In orario di servizio -> Sezione Sanitaria 22841",
          "Fuori orario -> 118",
          "Trasporto infortunato fuori dall'Unita' fino ad ambulanza -> "
          "a cura dell'Unita'",
          "Avviso immediato a Guardia Bacini 22916 in ogni caso"],
     "box":("warning", "Conserva il numero del soccorritore di bordo "
                        "designato. Verifica idoneita' e formazione."),
     "tk":None},

    # Slide 44 - Guardia e ronda
    {"t":"Guardia e Ronda", "k":MOD4, "n":"PSBM001A.A02",
     "b":["Modalita' definite nel verbale PSBM001A.A02",
          "Con Unita' in bacino, personale di guardia con prodel in "
          "dotazione",
          "Comunicazione continua con Guardia Bacini via VHF su canale "
          "predefinito",
          "Ronda con frequenza definita dal C.te e annotata su "
          "brogliaccio"],
     "box":("azione", "Test radio prodel a ogni cambio guardia. "
                       "Annotalo."),
     "tk":None},

    # Slide 45 - Chi fa cosa
    {"t":"Chi Fa Cosa: Riepilogo", "k":MOD4, "n":"Matrice RACI",
     "b":["Messa a terra scafo e sentinelle porta -> Comando di Bordo",
          "Funzionamento pompe A.I. e prosciugamento -> Sezione Bacini",
          "Uso e regolazione pressione rete A.I. -> Comando di Bordo "
          "(SAP)",
          "Gestione rifiuti e barriera -> C.te / Ufficio Programma"],
     "box":("regola", "Funzionamento impianti = Marinarsen. "
                       "Uso operativo = Bordo."),
     "tk":None},

    # Slide 46 - Formazione impianto
    {"t":"Formazione sull'Impianto Bacino", "k":MOD4, "n":"PSBM001A.A02",
     "b":["Il Nucleo NBM forma e informa il personale di bordo "
          "sull'impianto di bacino",
          "Verbalizzazione della formazione nel modello PSBM001A.A02",
          "Personale SAP autorizzato all'uso dell'impianto antincendio",
          "Aggiornamento periodico in base alla rotazione equipaggio"],
     "box":("azione", "Conserva i verbali di formazione: sono la tua "
                       "prova di diligenza."),
     "tk":None},

    # Slide 47 - Raccomandazioni
    {"t":"Raccomandazioni Specifiche", "k":MOD4, "n":"PSBM001A.A02",
     "b":["Definite nel verbale PSBM001A.A02 caso per caso",
          "Personalizzate in base alle lavorazioni previste",
          "Firmate da tutte le parti: Comando UN, Capo Sez. Bacini, Capo "
          "Nucleo, Ufficio Programma, Legale Rappresentante Ditta",
          "Aggiornate ad ogni variante di configurazione"],
     "box":("pretendi", "Tutte le firme presenti, datate, leggibili. "
                         "Nessuna eccezione."),
     "tk":None},

    # Slide 48 - 5 errori fatali
    {"t":"I 5 Errori Fatali", "k":MOD4, "n":"Da NON fare",
     "b":["1. Dare ordini diretti agli operai privati -> diventi loro "
          "Preposto",
          "2. Firmare il Verbale Consegne senza ispezionare le pompe "
          "A.I. -> nave vulnerabile",
          "3. Permettere accesso a spazi confinati senza Gas-Free "
          "certificato -> reato penale",
          "4. Spostare zavorre o carburanti senza avvisare Capo Nucleo "
          "NBM -> rischio ribaltamento",
          "5. Lasciare la portelleria esterna aperta dopo messa a secco "
          "-> rischio caduta dall'alto"],
     "box":("stop", "Cinque errori che possono mandarti in procura. "
                     "Memorizzali."),
     "tk":"Evita questi 5 errori e dormirai sonni tranquilli."},

    # ------------------- MODULO 5 -------------------
    # Slide 50 - Triangolo normativo
    {"t":"Il Triangolo Normativo: Quale Norma a Bordo?", "k":MOD5,
     "n":"Art. 26 / Titolo IV / D.Lgs. 272",
     "b":["Art. 26 D.Lgs. 81/08 (ditta singola): documento = DUVRI. "
          "Ruolo C.te = coordina e vigila",
          "Titolo IV D.Lgs. 81/08 (cantiere civile complesso): figura = "
          "CSE. Documento = PSC",
          "D.Lgs. 272/1999 (cantiere navale / bacino): figura = RTL. "
          "Documento = Documento di Sicurezza",
          "I tre regimi non si sommano: si applica quello pertinente al "
          "contesto"],
     "box":("regola", "A bordo la norma civile (81/08) viene "
                       "\"navalizzata\" dal D.Lgs. 272/99."),
     "tk":"81/08 navalizzato dal D.Lgs. 272/99."},

    # Slide 51 - Scenario 1: ditta singola
    {"t":"Scenario 1: La Ditta Singola", "k":MOD5,
     "n":"Art. 26 D.Lgs. 81/08",
     "b":["Quando: una sola impresa appaltatrice, nessun rischio "
          "di sovrapposizione",
          "C.te verifica l'Idoneita' Tecnico-Professionale (ITP)",
          "C.te fornisce i rischi specifici del bordo",
          "C.te redige e firma il DUVRI",
          "Non serve CSE, non serve PSC, non serve Impresa Capo Commessa"],
     "box":("regola", "1 Ditta = DUVRI. Responsabilita' operativa al "
                       "Datore di Lavoro della ditta."),
     "tk":None},

    # Slide 52 - Scenario 2: bacino
    {"t":"Scenario 2: Sosta Lavori e Bacino", "k":MOD5,
     "n":"Art. 116 D.Lgs. 272/99",
     "b":["Quando: soste lavori, bacino, piu' ditte che interferiscono",
          "Figura chiave: l'Impresa Capo Commessa e il R.T.L. "
          "(Responsabile Tecnico dei Lavori)",
          "L'R.T.L. sostituisce e adatta il ruolo del CSE civile al "
          "contesto navale",
          "Documento: Documento di Sicurezza (cronoprogramma, rischi, "
          "emergenze)"],
     "box":("regola", "In bacino l'interlocutore unico per le "
                       "interferenze e' l'R.T.L., non l'operaio."),
     "tk":None},

    # Slide 53 - Matrice di scelta
    {"t":"La Matrice di Scelta del Comandante", "k":MOD5, "n":"Matrice",
     "b":["Piccolo intervento (1 ditta) -> Art. 26 D.Lgs. 81/08 -> "
          "DUVRI -> Datore di Lavoro Ditta",
          "Cantiere complesso / sosta lavori / bacino -> D.Lgs. 272/99 "
          "Art. 116 -> Documento di Sicurezza -> R.T.L.",
          "Lavori in area isolata (es. bonifica amianto) -> Art. 118 "
          "SMM-PREVA-1062 -> Verbale Consegna Aree -> RTL + Guardia "
          "Fuochi"],
     "box":("stop", "Se in sosta lavori manca il Documento di Sicurezza, "
                     "il C.te ha il dovere di bloccare l'accesso."),
     "tk":"Manca il Documento di Sicurezza? Si blocca."},

    # Slide 54 - Documento di Sicurezza
    {"t":"Il Documento di Sicurezza", "k":MOD5, "n":"Art. 116 D.Lgs. 272/99",
     "b":["Redatto dall'Impresa Capo Commessa prima di iniziare",
          "Contiene: cronoprogramma, rischi interferenziali, misure di "
          "emergenza",
          "Consegnato in copia controfirmata a tutti i Datori di lavoro "
          "a bordo",
          "Se manca -> STOP AI LAVORI"],
     "box":("stop", "Documento di Sicurezza assente o non firmato = "
                     "lavori sospesi seduta stante."),
     "tk":None},

    # Slide 55 - Riunioni di sicurezza
    {"t":"Le Riunioni di Sicurezza", "k":MOD5, "n":"Art. 116 c.4",
     "b":["RTL convoca riunioni periodiche (giornate della sicurezza)",
          "Comando di bordo invia il suo incaricato",
          "I verbali integrano il Documento di Sicurezza",
          "RTL definisce la cadenza delle riunioni successive gia' "
          "nella prima"],
     "box":("azione", "Conserva tutti i verbali. Sono parte integrante "
                       "del Documento di Sicurezza."),
     "tk":None},

    # Slide 56 - Ambienti confinati
    {"t":"Accesso ad Ambienti Confinati", "k":MOD5,
     "n":"Art. 111 / DPR 177/2011",
     "b":["Sentine, cassoni, serbatoi, doppi fondi",
          "Accesso autorizzato dal COMANDANTE",
          "La ditta deve dimostrare il possesso dei requisiti DPR "
          "177/2011",
          "Pretendi di vedere il Permesso di Lavoro e il monitoraggio "
          "atmosfera (O2, LEL, H2S)"],
     "box":("stop", "Nessun accesso senza autorizzazione del Comandante "
                     "e Gas-Free certificato."),
     "tk":"Nessun accesso senza Gas-Free certificato."},

    # Slide 57 - Cosa fare in emergenza
    {"t":"Cosa Fare in Emergenza", "k":MOD5, "n":"22916 / 22841 / 22064",
     "b":["Incendio platea o bordo -> Guardia Bacini 22916. Attiva SAP "
          "bordo, usa rete idrica bacino",
          "Allagamento locale -> Guardia Bacini 22916 per pompe "
          "raccolta acque di processo",
          "Infortunio ditta -> la ditta chiama 118. Tu metti a "
          "disposizione passerella e chiami Sez. Sanitaria 22841",
          "Blackout elettrico -> Sezione Reti 22064"],
     "box":("warning", "Tieni accanto al telefono di plancia la rubrica "
                        "in formato grande e leggibile."),
     "tk":None},

    # ------------------- MODULO 6 -------------------
    # Slide 58 - Restituzione bacino
    {"t":"Restituzione del Bacino e Uscita", "k":MOD6,
     "n":"Art. 106 SMM-PREVA-1062",
     "b":["Prima dell'allagamento per l'uscita: verifica che tutte le "
          "aree isolate siano state riconsegnate",
          "Rimozione materiali e chiusura verbali di consegna aree",
          "PDCA (lessons learned): cosa ha funzionato, cosa migliorare",
          "Aggiornamento DVR se introdotte varianti di configurazione "
          "(Art. 106)"],
     "box":("azione", "Convoca un debrief con DM, SAP e Ufficio "
                       "Programma entro 7 giorni dall'uscita."),
     "tk":None},

    # Slide 59 - 5 punti chiave
    {"t":"I 5 Punti Chiave e Contatti", "k":MOD6, "n":"Sintesi",
     "b":["1. Alta Vigilanza, non ingerenza",
          "2. Rispetta la timeline: T-60, T-7, T-2",
          "3. Conosci la rubrica: 22916, 23762, 22064, 22841",
          "4. Firma solo dopo aver verificato di persona",
          "5. Il passaggio di comando avviene via VHF: pretendi "
          "chiarezza, registrazione, conoscenza"],
     "box":("regola", "Queste cinque regole sono la tua assicurazione "
                       "professionale e personale."),
     "tk":"Queste sono le tue fonti di salvezza."},
]


def main():
    slides = []

    # ---------- SLIDE 1: COPERTINA ----------
    slides.append(slide_cover(
        title="Gestione del Rischio Cantieristico e Responsabilita' di Bordo",
        subtitle="Guida operativa per l'ingresso, la permanenza e l'uscita "
                 "dal Bacino di Carenaggio",
        footer_txt="Rif: SMM-PREVA-1062 (Titolo VII)  |  PSBM001A Marinarsen "
                   "Taranto  |  NAV-G-001",
        kicker="CONVEGNO OPERATIVO  -  COMANDANTI E UFFICIALI DI BORDO"))

    # Indici delle slide che chiudono i moduli (per inserire i riepiloghi)
    # Stiamo ricostruendo l'ordine: 1 cover, 2..9 mod1 contenuti, 10 riep1,
    # 11..16 mod2, 17 riep2, 18..33 mod3, 34 riep3, 35..48 mod4, 49 riep4,
    # 50..57 mod5, 58..59 mod6, 60 Q&A.

    # ---------- SLIDE 2..9 : MODULO 1 (8 slide contenuto) ----------
    no = 2
    for item in SLIDES[0:8]:
        slides.append(slide_content(
            title=item["t"], kicker=item["k"], norm=item["n"],
            bullets=item["b"], box=item["box"], takeaway=item["tk"],
            footer_left=FOOT, slide_no=no, total=TOTAL))
        no += 1

    # ---------- SLIDE 10: Riepilogo Modulo 1 ----------
    slides.append(slide_riepilogo(
        modulo_num=1,
        quote="Vigila, non ingerire. Il confine e' la delega formale e la "
              "tracciabilita' documentale.",
        next_module="MODULO 2  -  La Timeline Operativa",
        footer_left=FOOT, slide_no=10, total=TOTAL))

    # ---------- SLIDE 11..16: MODULO 2 (6 slide contenuto) ----------
    no = 11
    for item in SLIDES[8:14]:
        slides.append(slide_content(
            title=item["t"], kicker=item["k"], norm=item["n"],
            bullets=item["b"], box=item["box"], takeaway=item["tk"],
            footer_left=FOOT, slide_no=no, total=TOTAL))
        no += 1

    # ---------- SLIDE 17: Riepilogo Modulo 2 ----------
    slides.append(slide_riepilogo(
        modulo_num=2,
        quote="L'ingresso in bacino si prepara due mesi prima. Se la nave "
              "non e' pronta, il bacino si ferma.",
        next_module="MODULO 3  -  Il Giorno dell'Ingresso",
        footer_left=FOOT, slide_no=17, total=TOTAL))

    # ---------- SLIDE 18..33: MODULO 3 (16 slide contenuto) ----------
    no = 18
    for item in SLIDES[14:30]:
        slides.append(slide_content(
            title=item["t"], kicker=item["k"], norm=item["n"],
            bullets=item["b"], box=item["box"], takeaway=item["tk"],
            footer_left=FOOT, slide_no=no, total=TOTAL))
        no += 1

    # ---------- SLIDE 34: Riepilogo Modulo 3 ----------
    slides.append(slide_riepilogo(
        modulo_num=3,
        quote="Il passaggio di comando avviene via VHF. Pretendi che sia "
              "chiaro, registrato e che tutti lo conoscano.",
        next_module="MODULO 4  -  Rubrica, Ruoli e Consegne",
        footer_left=FOOT, slide_no=34, total=TOTAL))

    # ---------- SLIDE 35..48: MODULO 4 (14 slide contenuto) ----------
    no = 35
    for item in SLIDES[30:44]:
        slides.append(slide_content(
            title=item["t"], kicker=item["k"], norm=item["n"],
            bullets=item["b"], box=item["box"], takeaway=item["tk"],
            footer_left=FOOT, slide_no=no, total=TOTAL))
        no += 1

    # ---------- SLIDE 49: Riepilogo Modulo 4 ----------
    slides.append(slide_riepilogo(
        modulo_num=4,
        quote="Firma il verbale di consegne solo dopo aver ispezionato le "
              "pompe e le passerelle. Se manca la pressione, la nave e' "
              "vulnerabile.",
        next_module="MODULO 5  -  Cantiere Vivo, Normative e Interferenze",
        footer_left=FOOT, slide_no=49, total=TOTAL))

    # ---------- SLIDE 50..57: MODULO 5 (8 slide contenuto) ----------
    no = 50
    for item in SLIDES[44:52]:
        slides.append(slide_content(
            title=item["t"], kicker=item["k"], norm=item["n"],
            bullets=item["b"], box=item["box"], takeaway=item["tk"],
            footer_left=FOOT, slide_no=no, total=TOTAL))
        no += 1

    # ---------- SLIDE 58..59: MODULO 6 (2 slide contenuto) ----------
    no = 58
    for item in SLIDES[52:54]:
        slides.append(slide_content(
            title=item["t"], kicker=item["k"], norm=item["n"],
            bullets=item["b"], box=item["box"], takeaway=item["tk"],
            footer_left=FOOT, slide_no=no, total=TOTAL))
        no += 1

    # ---------- SLIDE 60: Q&A finale ----------
    slides.append(slide_qa(
        title="Domande & Risposte",
        subtitle="Spazio per 2-3 domande pratiche dai Comandanti",
        stimolo="Qual e' il dubbio che vi ha tenuto svegli la notte prima "
                "del vostro ultimo bacino?",
        takeaway="Grazie per l'attenzione  -  Buon mare e bacini "
                 "tranquilli.",
        footer_left=FOOT, slide_no=60, total=TOTAL))

    # Verifica conteggio
    assert len(slides) == TOTAL, (
        f"Attese {TOTAL} slide, generate {len(slides)}")

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "Gestione_Rischio_Cantieristico_Bacino.pptx")
    build_pptx(slides, out,
               title="Gestione del Rischio Cantieristico e "
                     "Responsabilita' di Bordo")
    print(f"Creato: {out}")
    print(f"Numero slide: {len(slides)}")


if __name__ == "__main__":
    main()
