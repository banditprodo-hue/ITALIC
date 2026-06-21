# -*- coding: utf-8 -*-
"""
Generatore PPTX "SLIDE RISCHIO MECCANICO"
Corso di formazione lavoratori - D.Lgs. 81/08
Modulo 1: Rischio Meccanico (60 slide)
Modulo 2: Movimentazione Manuale dei Carichi (40 slide)

Costruisce un file .pptx valido (OOXML) usando solo la libreria standard,
perche' python-pptx non e' installabile in questo ambiente (rete bloccata).
Palette "Safety Industrial": Nero / Giallo Segnale / Grigio Antracite.
"""
import os
import zipfile

# ---------------------------------------------------------------------------
# Palette "Safety Industrial"
# ---------------------------------------------------------------------------
BLACK      = "111111"   # nero
ANTHRACITE = "2B2B2B"   # grigio antracite
YELLOW     = "FFD200"   # giallo segnale
WHITE      = "FFFFFF"
LIGHTGRAY  = "F2F2F2"
MIDGRAY    = "8C8C8C"
DARKTEXT   = "262626"
GRAYTEXT   = "595959"

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


def para(runs, algn="l", bullet=False, level=0, bullet_color=YELLOW,
         space_before=600, space_after=0, line=110000):
    """Costruisce un paragrafo. runs = lista di stringhe <a:r>..."""
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
    fill_xml = ''
    if fill is None:
        fill_xml = '<a:noFill/>'
    else:
        fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    line_xml = ''
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


def slide_cover(title, subtitle, footer, kicker):
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=BLACK, name="bg"))
    # blocco antracite
    s.append(shape_rect(0, 1500000, SW, 2750000, fill=ANTHRACITE, name="band"))
    # barra gialla
    s.append(shape_rect(0, 1500000, 380000, 2750000, fill=YELLOW, name="ybar"))
    # kicker
    s.append(shape_text(900000, 1680000, 10000000, 500000,
        [para([run(kicker, 1500, YELLOW, bold=True)], space_before=0)]))
    # titolo
    s.append(shape_text(900000, 2150000, 10400000, 1400000,
        [para([run(title, 4400, WHITE, bold=True)], space_before=0, line=104000)]))
    # sottotitolo
    s.append(shape_text(900000, 3500000, 10400000, 600000,
        [para([run(subtitle, 2000, YELLOW)], space_before=0)]))
    # footer
    s.append(shape_rect(0, SH-520000, SW, 520000, fill=ANTHRACITE, name="ft"))
    s.append(shape_rect(0, SH-520000, SW, 50000, fill=YELLOW, name="ftacc"))
    s.append(shape_text(900000, SH-470000, 10400000, 420000,
        [para([run(footer, 1300, WHITE, bold=True)], space_before=0)], anchor="ctr"))
    return wrap_slide(s)


def slide_divider(modulo, title, points, slide_no, total):
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=ANTHRACITE, name="bg"))
    # blocco numero modulo
    s.append(shape_rect(0, 0, 3600000, SH, fill=BLACK, name="numblock"))
    s.append(shape_rect(3600000, 0, 90000, SH, fill=YELLOW, name="vbar"))
    s.append(shape_text(300000, 1850000, 3100000, 600000,
        [para([run(modulo, 1600, YELLOW, bold=True)], algn="ctr", space_before=0)]))
    # grande numero
    num = modulo.split()[-1] if modulo.split() else ""
    s.append(shape_text(300000, 2350000, 3100000, 1900000,
        [para([run(num, 13000, YELLOW, bold=True)], algn="ctr", space_before=0)],
        anchor="ctr"))
    # titolo modulo
    s.append(shape_text(4050000, 1900000, 7700000, 1500000,
        [para([run(title, 3600, WHITE, bold=True)], space_before=0, line=104000)]))
    s.append(shape_rect(4080000, 3450000, 1500000, 50000, fill=YELLOW, name="acc"))
    # punti
    pp = [para([run(p, 1700, LIGHTGRAY)], bullet=True, space_before=500)
          for p in points]
    s.append(shape_text(4050000, 3650000, 7700000, 2400000, pp))
    # footer numero slide
    s.append(shape_text(SW-1800000, SH-560000, 1500000, 400000,
        [para([run(f"{slide_no} / {total}", 1200, MIDGRAY)], algn="r",
              space_before=0)], anchor="ctr"))
    return wrap_slide(s)


def slide_content(title, norm, bullets, graphic, footer_left, slide_no, total):
    """bullets: lista di (testo, livello). livello 0 o 1."""
    reset_ids()
    s = []
    # sfondo
    s.append(shape_rect(0, 0, SW, SH, fill=WHITE, name="bg"))
    # header antracite
    s.append(shape_rect(0, 0, SW, 1080000, fill=ANTHRACITE, name="header"))
    # barra gialla sotto header
    s.append(shape_rect(0, 1080000, SW, 52000, fill=YELLOW, name="hacc"))
    # tab gialla a sinistra dell'header
    s.append(shape_rect(0, 0, 150000, 1080000, fill=YELLOW, name="htab"))
    # titolo
    s.append(shape_text(420000, 90000, 8900000, 900000,
        [para([run(title, 2600, WHITE, bold=True)], space_before=0, line=102000)],
        anchor="ctr"))
    # riferimento normativo (in alto a destra)
    if norm:
        s.append(shape_rect(9300000, 300000, 2620000, 480000, fill=YELLOW, name="normtag"))
        s.append(shape_text(9320000, 300000, 2580000, 480000,
            [para([run(norm, 1100, BLACK, bold=True)], algn="ctr", space_before=0)],
            anchor="ctr"))
    # corpo: bullet
    bp = []
    for (txt, lvl) in bullets:
        if lvl == 0:
            bp.append(para([run(txt, 1700, DARKTEXT)], bullet=True, level=0,
                           space_before=560))
        else:
            bp.append(para([run(txt, 1450, GRAYTEXT)], bullet=True, level=1,
                           bullet_color=MIDGRAY, space_before=320))
    s.append(shape_text(470000, 1400000, 7100000, 4900000, bp))
    # box suggerimento grafico
    bx, by, bw, bh = 7820000, 1400000, 3950000, 4750000
    s.append(shape_rect(bx, by, bw, bh, fill=LIGHTGRAY, line=ANTHRACITE,
                        line_w=9525, name="gbox"))
    s.append(shape_rect(bx, by, bw, 430000, fill=ANTHRACITE, name="ghdr"))
    s.append(shape_rect(bx, by, 130000, 430000, fill=YELLOW, name="gtab"))
    s.append(shape_text(bx+200000, by, bw-220000, 430000,
        [para([run("SUGGERIMENTO GRAFICO", 1150, YELLOW, bold=True)],
              space_before=0)], anchor="ctr"))
    s.append(shape_text(bx+170000, by+560000, bw-340000, bh-700000,
        [para([run(graphic, 1400, DARKTEXT, italic=True)], space_before=0,
              line=118000)]))
    # footer
    s.append(shape_rect(0, SH-470000, SW, 470000, fill=ANTHRACITE, name="footer"))
    s.append(shape_rect(0, SH-470000, SW, 44000, fill=YELLOW, name="facc"))
    s.append(shape_text(420000, SH-470000, 9000000, 470000,
        [para([run(footer_left, 1050, WHITE)], space_before=0)], anchor="ctr"))
    s.append(shape_text(SW-2200000, SH-470000, 1780000, 470000,
        [para([run(f"{slide_no} / {total}", 1100, YELLOW, bold=True)], algn="r",
              space_before=0)], anchor="ctr"))
    return wrap_slide(s)


def slide_closing(title, points, footer):
    reset_ids()
    s = []
    s.append(shape_rect(0, 0, SW, SH, fill=BLACK, name="bg"))
    s.append(shape_rect(0, 1350000, SW, 90000, fill=YELLOW, name="acc"))
    s.append(shape_text(900000, 600000, 10400000, 800000,
        [para([run(title, 4000, YELLOW, bold=True)], space_before=0)]))
    pp = [para([run(p, 1800, WHITE)], bullet=True, space_before=600)
          for p in points]
    s.append(shape_text(900000, 1700000, 10400000, 3800000, pp))
    s.append(shape_rect(0, SH-520000, SW, 520000, fill=ANTHRACITE, name="ft"))
    s.append(shape_text(900000, SH-520000, 10400000, 520000,
        [para([run(footer, 1300, WHITE, bold=True)], space_before=0)], anchor="ctr"))
    return wrap_slide(s)



# ---------------------------------------------------------------------------
# Scaffolding OOXML
# ---------------------------------------------------------------------------
THEME = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="SafetyIndustrial">
<a:themeElements>
<a:clrScheme name="SafetyIndustrial">
<a:dk1><a:srgbClr val="111111"/></a:dk1>
<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="2B2B2B"/></a:dk2>
<a:lt2><a:srgbClr val="F2F2F2"/></a:lt2>
<a:accent1><a:srgbClr val="FFD200"/></a:accent1>
<a:accent2><a:srgbClr val="2B2B2B"/></a:accent2>
<a:accent3><a:srgbClr val="8C8C8C"/></a:accent3>
<a:accent4><a:srgbClr val="595959"/></a:accent4>
<a:accent5><a:srgbClr val="D9D9D9"/></a:accent5>
<a:accent6><a:srgbClr val="000000"/></a:accent6>
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


def build_pptx(slides_xml, out_path, title="SLIDE RISCHIO MECCANICO"):
    n = len(slides_xml)
    # [Content_Types].xml
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

    # _rels/.rels
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>')

    # presentation.xml
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

    # presentation rels
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

    # master rels
    master_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
        '</Relationships>')

    # layout rels
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
        '<cp:keywords>D.Lgs. 81/08; Rischio Meccanico; MMC; NIOSH</cp:keywords>'
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



# ---------------------------------------------------------------------------
# CONTENUTI - MODULO 1: RISCHIO MECCANICO
# Ogni voce: {"t": titolo, "n": rif. normativo, "b": bullet, "g": grafica}
# Un bullet che inizia con ">" e' di secondo livello.
# ---------------------------------------------------------------------------
FOOT1 = "Modulo 1 - Rischio Meccanico  |  Formazione lavoratori D.Lgs. 81/08"
FOOT2 = "Modulo 2 - Movimentazione Manuale dei Carichi  |  D.Lgs. 81/08"

M1 = [
 {"t":"Obiettivi del modulo","n":"D.Lgs. 81/08",
  "b":["Riconoscere i pericoli di origine meccanica nei luoghi di lavoro",
       "Comprendere le misure di prevenzione e protezione",
       "Conoscere ripari, dispositivi e interblocchi di sicurezza",
       "Applicare procedure sicure in uso e manutenzione",
       "Individuare i DPI adeguati al rischio meccanico"],
  "g":"Infografica di apertura con icone di ingranaggi, riparo e DPI su fascia gialla/antracite, stile industriale."},

 {"t":"Quadro normativo di riferimento","n":"D.Lgs. 81/08",
  "b":["Testo Unico Sicurezza: D.Lgs. 9 aprile 2008, n. 81",
       "Titolo III - Uso delle attrezzature di lavoro e DPI",
       "Direttiva Macchine 2006/42/CE (recepita con D.Lgs. 17/2010)",
       "Norme tecniche armonizzate (serie UNI EN ISO 12100)",
       "Obblighi di datore di lavoro, dirigenti, preposti e lavoratori"],
  "g":"Schema a piramide normativa: Direttive UE in alto, D.Lgs. 81/08 al centro, norme tecniche alla base."},

 {"t":"Cos'e' il rischio meccanico","n":"UNI EN ISO 12100",
  "b":["Insieme dei pericoli generati da elementi di macchine e attrezzature",
       "Deriva da parti in movimento, superfici, spigoli, energia accumulata",
       "Puo' causare lesioni traumatiche immediate",
       "E' presente in produzione, manutenzione, pulizia e regolazione"],
  "g":"Foto tecnica di organi meccanici in movimento con sovrapposti i punti di pericolo evidenziati in giallo."},

 {"t":"Pericolo, rischio e danno","n":"Art. 2 D.Lgs. 81/08",
  "b":["Pericolo: proprieta' intrinseca con potenziale di causare danno",
       "Rischio: probabilita' x gravita' del danno (R = P x D)",
       "Danno: lesione fisica o alterazione della salute",
       "Esposizione: condizione che mette in contatto uomo e pericolo"],
  "g":"Diagramma R = P x D con matrice di rischio a colori (verde/giallo/rosso) in chiave industriale."},

 {"t":"Le macchine: definizione","n":"Dir. 2006/42/CE",
  "b":["Insieme equipaggiato di un sistema di azionamento",
       "Parti collegate per un'applicazione ben determinata",
       "Comprende quasi-macchine e attrezzature intercambiabili",
       "Deve rispettare i Requisiti Essenziali di Sicurezza (RES)"],
  "g":"Disegno tecnico esploso di una macchina utensile con etichette dei principali gruppi funzionali."},

 {"t":"Marcatura CE e documentazione","n":"Dir. 2006/42/CE",
  "b":["Marcatura CE: conformita' ai requisiti essenziali di sicurezza",
       "Dichiarazione CE di conformita' obbligatoria",
       "Fascicolo tecnico e manuale d'uso in lingua italiana",
       "Targa con dati identificativi del fabbricante"],
  "g":"Primo piano di una targa macchina con marcatura CE e dettaglio del manuale d'uso e manutenzione."},

 {"t":"Obblighi del datore di lavoro","n":"Art. 71 D.Lgs. 81/08",
  "b":["Mettere a disposizione attrezzature conformi e idonee",
       "Garantire installazione, uso e manutenzione corretti",
       "Sottoporre le attrezzature a controlli periodici",
       "Assicurare informazione, formazione e addestramento",
       "Conservare il registro dei controlli"],
  "g":"Checklist illustrata con icone di verifica su sfondo antracite e segni di spunta gialli."},

 {"t":"Requisiti delle attrezzature","n":"Art. 70 / All. V",
  "b":["Conformita' alle direttive di prodotto applicabili",
       "Per attrezzature antecedenti: requisiti dell'Allegato V",
       "Organi di comando sicuri, visibili e identificabili",
       "Dispositivi di arresto generale e di emergenza",
       "Protezione contro contatti con elementi mobili"],
  "g":"Tabella comparativa requisiti Allegato V vs Direttiva Macchine, con evidenziazione gialla."},

 {"t":"Obblighi dei lavoratori","n":"Art. 20 D.Lgs. 81/08",
  "b":["Usare correttamente attrezzature e dispositivi di sicurezza",
       "Non rimuovere o modificare i dispositivi di protezione",
       "Segnalare guasti e situazioni di pericolo",
       "Utilizzare i DPI messi a disposizione",
       "Partecipare a formazione e addestramento"],
  "g":"Illustrazione di operatore che segnala un'anomalia e indossa correttamente i DPI."},

 {"t":"Pericoli meccanici: panoramica","n":"UNI EN ISO 12100",
  "b":["Schiacciamento, cesoiamento, taglio e sezionamento",
       "Impigliamento, trascinamento, intrappolamento",
       "Urto, perforazione, puntura",
       "Attrito e abrasione",
       "Eiezione di fluidi e proiezione di materiali"],
  "g":"Tavola sinottica con le 10 icone dei pericoli meccanici fondamentali in stile pittogramma."},

 {"t":"Pericolo di schiacciamento","n":"UNI EN ISO 12100",
  "b":["Parte del corpo compressa tra due elementi",
       "Uno mobile verso uno fisso o due elementi mobili",
       "Tipico di presse, stampi, organi di chiusura",
       "Prevenzione: distanze di sicurezza e ripari"],
  "g":"Disegno tecnico di una pressa con freccia che mostra la zona di schiacciamento evidenziata in rosso."},

 {"t":"Pericolo di cesoiamento","n":"UNI EN ISO 12100",
  "b":["Taglio tra due elementi che scorrono uno sull'altro",
       "Azione tipo forbice tra parte mobile e fissa",
       "Frequente in cesoie, presse piegatrici, nastri",
       "Prevenzione: ripari e dispositivi sensibili"],
  "g":"Illustrazione di lama cesoia con vettori di movimento e indicazione della zona pericolosa."},

 {"t":"Taglio e sezionamento","n":"UNI EN ISO 12100",
  "b":["Contatto con elementi affilati o utensili rotanti",
       "Lame, dischi, frese, seghe circolari",
       "Lesioni anche con macchina in arresto (spigoli)",
       "Prevenzione: cuffie di protezione e coltelli divisori"],
  "g":"Sega circolare con cuffia di protezione e coltello divisore evidenziati, vista laterale tecnica."},

 {"t":"Impigliamento","n":"UNI EN ISO 12100",
  "b":["Cattura di indumenti, capelli o monili da parti mobili",
       "Tipico di alberi rotanti, mandrini, viti senza fine",
       "Aggravato da abiti larghi e parti sporgenti",
       "Prevenzione: ripari, abiti aderenti, no monili"],
  "g":"Disegno di un albero rotante con simbolo di divieto di indumenti larghi e capelli sciolti."},

 {"t":"Trascinamento e intrappolamento","n":"UNI EN ISO 12100",
  "b":["Cattura del corpo in punti di convergenza (nip points)",
       "Tra rulli, ingranaggi, cinghie e pulegge",
       "Evoluzione rapida e difficilmente reversibile",
       "Prevenzione: ripari fissi e arresti di emergenza accessibili"],
  "g":"Schema di due rulli controrotanti con il punto di presa (nip point) cerchiato in giallo."},

 {"t":"Urto","n":"UNI EN ISO 12100",
  "b":["Colpo da elementi in movimento rapido",
       "Bracci robotizzati, parti oscillanti, masse sospese",
       "Rischio in zone di accesso e percorsi",
       "Prevenzione: barriere, distanze, segnalazioni"],
  "g":"Illustrazione di un braccio robotico con area di lavoro delimitata e zona di urto in evidenza."},

 {"t":"Perforazione e puntura","n":"UNI EN ISO 12100",
  "b":["Penetrazione da elementi appuntiti o aghi",
       "Punte, trapani, aghi di macchine tessili",
       "Anche da getti ad alta pressione",
       "Prevenzione: ripari e protezione delle punte"],
  "g":"Dettaglio tecnico di una punta da trapano con cono di protezione e indicazione della zona di rischio."},

 {"t":"Attrito e abrasione","n":"UNI EN ISO 12100",
  "b":["Contatto con superfici ruvide in movimento",
       "Mole, nastri abrasivi, cinghie",
       "Lesioni cutanee e ustioni da frizione",
       "Prevenzione: ripari e poggia-pezzo regolabili"],
  "g":"Mola da banco con poggiapezzo e schermo paraschegge, vista frontale con quote di sicurezza."},

 {"t":"Eiezione di fluidi ad alta pressione","n":"UNI EN ISO 12100",
  "b":["Getti di fluido in pressione che penetrano nei tessuti",
       "Sistemi oleodinamici e pneumatici",
       "Lesioni gravi anche da fori microscopici",
       "Prevenzione: manutenzione tubi, schermi, depressurizzazione"],
  "g":"Sezione di un circuito oleodinamico con punto di perdita e getto ad alta pressione evidenziato."},

 {"t":"Proiezione di elementi","n":"UNI EN ISO 12100",
  "b":["Espulsione di pezzi, trucioli o frammenti di utensile",
       "Rottura di mole, utensili o parti del pezzo",
       "Traiettorie imprevedibili ad alta energia",
       "Prevenzione: schermi paraschegge e protezione occhi"],
  "g":"Tornio con schermo trasparente paraschegge e traiettorie di trucioli rappresentate con frecce."},

 {"t":"Organi di trasmissione del moto","n":"All. V D.Lgs. 81/08",
  "b":["Cinghie, pulegge, catene, ingranaggi, alberi",
       "Elementi non destinati al contatto con l'operatore",
       "Devono essere completamente segregati",
       "Prevenzione: ripari fissi totali sull'intera trasmissione"],
  "g":"Disegno tecnico di un gruppo cinghia-puleggia con carter di protezione integrale."},

 {"t":"Punti di presa (nip points)","n":"UNI EN ISO 12100",
  "b":["Zone di convergenza tra elementi in movimento",
       "Rulli, ingranaggi, cinghia-puleggia, nastro-tamburo",
       "Sufficiente l'aggancio di un lembo per il trascinamento",
       "Prevenzione: ripari e dispositivi di prossimita'"],
  "g":"Mappa di una macchina con tutti i nip points contrassegnati da cerchi gialli numerati."},

 {"t":"Elementi mobili di lavoro","n":"All. V D.Lgs. 81/08",
  "b":["Parti che agiscono direttamente sul pezzo",
       "Utensili, lame, punzoni, matrici",
       "Accesso necessario per la lavorazione",
       "Prevenzione: ripari regolabili e dispositivi di protezione"],
  "g":"Illustrazione di una fresatrice con utensile in lavorazione e riparo regolabile in posizione."},

 {"t":"Zone pericolose e accesso","n":"UNI EN ISO 12100",
  "b":["Zona pericolosa: spazio dentro/attorno alla macchina",
       "Persona esposta: chiunque si trovi nella zona",
       "Accesso volontario o involontario al pericolo",
       "Prevenzione: limitare l'accesso e mantenere distanze"],
  "g":"Pianta di una postazione con perimetro della zona pericolosa tratteggiato e distanze quotate."},

 {"t":"Movimenti rotatori pericolosi","n":"UNI EN ISO 12100",
  "b":["Mandrini, alberi, dischi, ventole",
       "Rischio di impigliamento e trascinamento",
       "Sporgenze (viti, chiavette) aumentano il pericolo",
       "Prevenzione: superfici lisce e ripari avvolgenti"],
  "g":"Vista 3D di un mandrino rotante con frecce di rotazione e segnalazione delle sporgenze."},

 {"t":"Movimenti lineari e alternati","n":"UNI EN ISO 12100",
  "b":["Slitte, tavole, stantuffi, magli",
       "Rischio di schiacciamento e cesoiamento",
       "Inversioni di moto rapide e ripetute",
       "Prevenzione: ripari mobili interbloccati"],
  "g":"Schema di una slitta con corsa indicata e zone di fine corsa pericolose evidenziate."},

 {"t":"Stabilita' delle macchine","n":"All. V D.Lgs. 81/08",
  "b":["Rischio di ribaltamento o spostamento",
       "Ancoraggio al suolo quando necessario",
       "Baricentro e distribuzione dei carichi",
       "Prevenzione: fissaggio, livellamento, limiti di carico"],
  "g":"Disegno di una macchina ancorata al pavimento con frecce del baricentro e punti di fissaggio."},

 {"t":"Rotture durante il funzionamento","n":"UNI EN ISO 12100",
  "b":["Cedimento di utensili, mole, elementi sotto sforzo",
       "Fatica dei materiali e sovraccarichi",
       "Proiezione di frammenti ad alta energia",
       "Prevenzione: limiti d'uso, controlli, schermature"],
  "g":"Sequenza illustrata della rottura di una mola con contenimento da parte del carter."},

 {"t":"Gerarchia delle misure (STOP)","n":"Art. 15 D.Lgs. 81/08",
  "b":["Eliminazione/sostituzione del pericolo alla fonte",
       "Misure tecniche: ripari e dispositivi di protezione",
       "Misure organizzative: procedure e formazione",
       "DPI come ultima barriera di protezione"],
  "g":"Piramide STOP a quattro livelli con colori dal verde (eliminazione) al giallo (DPI)."},

 {"t":"Sicurezza nella progettazione","n":"UNI EN ISO 12100",
  "b":["Prevenzione intrinseca: ridurre il pericolo alla fonte",
       "Riduzione delle parti mobili accessibili",
       "Forme arrotondate, superfici lisce, limiti di energia",
       "Priorita' rispetto a ripari e DPI"],
  "g":"Confronto tra una macchina mal progettata e una progettata in sicurezza (split screen tecnico)."},

 {"t":"Ripari: definizione e funzioni","n":"UNI EN ISO 14120",
  "b":["Elemento che fornisce protezione mediante barriera fisica",
       "Impedisce l'accesso alla zona pericolosa",
       "Trattiene proiezioni di materiali e fluidi",
       "Tipologie: fissi, mobili, regolabili"],
  "g":"Esploso tecnico delle tipologie di riparo con etichette: fisso, mobile, interbloccato, regolabile."},

 {"t":"Ripari fissi","n":"UNI EN ISO 14120",
  "b":["Mantenuti in posizione in modo permanente",
       "Fissati con sistemi che richiedono utensili",
       "Per zone con accesso non necessario in esercizio",
       "Robusti, non facilmente rimovibili"],
  "g":"Carter fisso bullonato su un gruppo di trasmissione, con dettaglio dei fissaggi a vite."},

 {"t":"Ripari mobili interbloccati","n":"UNI EN ISO 14119",
  "b":["Apribili senza utensili per accessi frequenti",
       "Collegati a un dispositivo di interblocco",
       "L'apertura arresta le funzioni pericolose",
       "La macchina non riparte alla sola chiusura"],
  "g":"Sportello interbloccato con sensore e schema del circuito di sicurezza che ferma la macchina."},

 {"t":"Ripari con bloccaggio","n":"UNI EN ISO 14119",
  "b":["Riparo interbloccato con dispositivo di bloccaggio",
       "Resta chiuso e bloccato finche' persiste il rischio",
       "Apertura solo a fermo macchina/arresto inerzia",
       "Per movimenti con lunghi tempi di arresto"],
  "g":"Sezione di un elettroserratura di sicurezza con stato bloccato/sbloccato evidenziato."},

 {"t":"Ripari regolabili e autoregolabili","n":"UNI EN ISO 14120",
  "b":["Adattano l'apertura alle dimensioni del pezzo",
       "Regolabili manualmente o automaticamente",
       "Riducono l'accesso alla zona dell'utensile",
       "Tipici di seghe, frese, trapani"],
  "g":"Sega a nastro con riparo regolabile che segue lo spessore del pezzo in lavorazione."},

 {"t":"Dispositivi di protezione","n":"UNI EN ISO 12100",
  "b":["Non costituiscono una barriera materiale",
       "Rilevano la presenza o limitano l'accesso",
       "Barriere immateriali, comandi a due mani, tappeti",
       "Arrestano il pericolo prima del contatto"],
  "g":"Pannello con i diversi dispositivi non materiali rappresentati da pittogrammi tecnici."},

 {"t":"Barriere immateriali (ESPE)","n":"UNI EN 61496",
  "b":["Fasci di luce che rilevano l'ingresso nella zona",
       "Fotocellule, barriere fotoelettriche, scanner laser",
       "L'interruzione del fascio arresta la macchina",
       "Risoluzione e distanza calcolate sul rischio"],
  "g":"Pressa protetta da barriera fotoelettrica con fasci luminosi rappresentati in giallo."},

 {"t":"Tappeti e bordi sensibili","n":"UNI EN ISO 13856",
  "b":["Rilevano la presenza o il contatto",
       "Tappeti sensibili a pressione al suolo",
       "Bordi e barre sensibili su elementi mobili",
       "Comandano l'arresto della funzione pericolosa"],
  "g":"Area di lavoro con tappeto sensibile a pavimento delimitato e zona di rilevamento evidenziata."},

 {"t":"Comandi a due mani","n":"UNI EN ISO 13851",
  "b":["Richiedono l'uso simultaneo di entrambe le mani",
       "Mantengono le mani fuori dalla zona pericolosa",
       "Azionamento contemporaneo entro 0,5 secondi",
       "Proteggono solo l'operatore che li aziona"],
  "g":"Postazione con due pulsanti distanziati e mani dell'operatore impegnate, vista dall'alto."},

 {"t":"Arresto di emergenza","n":"UNI EN ISO 13850",
  "b":["Funzione che arresta rapidamente la macchina",
       "Comandi a fungo rosso su fondo giallo",
       "Accessibili e ben visibili da ogni postazione",
       "Riattivazione solo con azione volontaria"],
  "g":"Primo piano del pulsante a fungo rosso su targhetta gialla con simbolo di emergenza."},

 {"t":"Sistemi di interblocco","n":"UNI EN ISO 14119",
  "b":["Collegano lo stato del riparo alle funzioni pericolose",
       "Apertura del riparo = arresto della macchina",
       "Sensori meccanici, magnetici, codificati",
       "Resistenti all'elusione (tamper)"],
  "g":"Schema funzionale di un interblocco: sensore, logica di sicurezza, attuatore di arresto."},

 {"t":"Affidabilita' dei sistemi (PL/SIL)","n":"EN ISO 13849-1",
  "b":["Le funzioni di sicurezza hanno un'affidabilita' misurabile",
       "Performance Level (PL) da a a e",
       "Safety Integrity Level (SIL) per sistemi elettronici",
       "Maggiore il rischio, maggiore il PL richiesto"],
  "g":"Tabella PL/SIL con scala di affidabilita' crescente e matrice di selezione del rischio."},

 {"t":"Dispositivi di comando","n":"All. V D.Lgs. 81/08",
  "b":["Chiaramente visibili e identificabili",
       "Protetti da azionamenti accidentali",
       "Posizionati fuori dalle zone pericolose",
       "Avviamento solo con azione volontaria su comando"],
  "g":"Pannello comandi con pulsanti codificati per colore e protezioni anti-azionamento accidentale."},

 {"t":"Riavvio inatteso","n":"UNI EN ISO 14118",
  "b":["Rischio di ripartenza non voluta durante l'intervento",
       "Cause: ripristino energia, reset automatico, guasti",
       "Prevenzione del riavvio automatico dopo arresto",
       "Necessario consenso volontario per la ripartenza"],
  "g":"Sequenza che mostra un riavvio inatteso evitato dal blocco del consenso di ripartenza."},

 {"t":"Separazione dalle energie","n":"UNI EN ISO 14118",
  "b":["Sezionamento di tutte le fonti di energia",
       "Elettrica, pneumatica, idraulica, meccanica",
       "Dissipazione delle energie residue e accumulate",
       "Base per interventi sicuri di manutenzione"],
  "g":"Pannello con sezionatore elettrico, valvola pneumatica e scarico energie residue etichettati."},

 {"t":"Procedura di Lockout/Tagout","n":"UNI EN ISO 14118",
  "b":["Lockout: blocco fisico del sezionatore con lucchetto",
       "Tagout: cartello di segnalazione del divieto di manovra",
       "Ogni operatore applica il proprio lucchetto",
       "Verifica dell'energia zero prima di operare"],
  "g":"Sezionatore con lucchetti multipli LOTO e cartellino di avviso, dettaglio fotografico tecnico."},

 {"t":"Manutenzione in sicurezza","n":"Art. 71 c.4 / All. VI",
  "b":["Interventi a macchina ferma quando possibile",
       "Applicazione delle procedure LOTO",
       "Personale formato e autorizzato",
       "Permesso di lavoro per interventi critici"],
  "g":"Tecnico che applica il lucchetto prima dell'intervento, con cartello permesso di lavoro."},

 {"t":"Accesso ai punti di intervento","n":"All. V D.Lgs. 81/08",
  "b":["Punti di manutenzione accessibili in sicurezza",
       "Passerelle, scale e ripari a norma",
       "Evitare interventi in posizioni pericolose",
       "Illuminazione adeguata dell'area di lavoro"],
  "g":"Disegno di una macchina con passerella e punti di lubrificazione accessibili dall'esterno."},

 {"t":"Pulizia e regolazione sicure","n":"All. V D.Lgs. 81/08",
  "b":["Operazioni frequenti ad alto rischio di contatto",
       "Vietato pulire con organi in movimento",
       "Usare utensili idonei, mai le mani",
       "Modalita' protette per registrazioni e set-up"],
  "g":"Operatore che pulisce una macchina ferma con utensile dedicato, simbolo di divieto mani."},

 {"t":"Modalita' di funzionamento speciali","n":"UNI EN ISO 12100",
  "b":["Modi diversi dal normale esercizio (set-up, diagnosi)",
       "Selettore di modo con chiave",
       "Velocita' ridotta e comando ad azione mantenuta",
       "Sospensione controllata di alcune protezioni"],
  "g":"Selettore di modo a chiave con posizioni AUTO/SET-UP e indicazione della velocita' ridotta."},

 {"t":"Segnaletica sulle macchine","n":"Titolo V / All. XXIV",
  "b":["Segnali di avvertimento dei pericoli residui",
       "Pittogrammi su organi mobili e zone calde",
       "Istruzioni e limiti d'uso visibili",
       "Manutenzione e leggibilita' nel tempo"],
  "g":"Set di pittogrammi di pericolo (organi mobili, schiacciamento, alta tensione) su fondo giallo."},

 {"t":"Informazione e addestramento","n":"Art. 73 D.Lgs. 81/08",
  "b":["Informazioni su condizioni d'uso e situazioni anomale",
       "Formazione specifica sull'uso delle attrezzature",
       "Addestramento per attrezzature che lo richiedono",
       "Abilitazione per attrezzature dell'Accordo Stato-Regioni"],
  "g":"Aula di addestramento pratico su macchina con istruttore e operatore, taglio fotografico tecnico."},

 {"t":"DPI per il rischio meccanico","n":"Titolo III Capo II",
  "b":["Ultima barriera quando il rischio non e' eliminabile",
       "Scelti in base alla valutazione dei rischi",
       "Marcati CE e adeguati alla mansione",
       "Obbligo di uso, cura e manutenzione"],
  "g":"Manichino con set completo di DPI antinfortunistici evidenziati uno a uno con didascalie."},

 {"t":"Protezione di mani e braccia","n":"UNI EN 388",
  "b":["Guanti contro tagli, abrasioni, perforazioni",
       "Marcatura EN 388 con livelli di prestazione",
       "Attenzione: vietati con organi rotanti (impigliamento)",
       "Scelta in base al tipo di lavorazione"],
  "g":"Guanto antitaglio con pittogramma EN 388 e spiegazione dei quattro indici prestazionali."},

 {"t":"Protezione di occhi e viso","n":"UNI EN 166",
  "b":["Occhiali e visiere contro proiezioni di frammenti",
       "Schermi facciali per molatura e tornitura",
       "Resistenza all'impatto secondo EN 166",
       "Compatibilita' con altri DPI"],
  "g":"Operatore con visiera durante molatura, con frammenti deviati rappresentati da frecce."},

 {"t":"Protezione di corpo e piedi","n":"UNI EN ISO 20345",
  "b":["Indumenti aderenti per evitare impigliamenti",
       "Calzature con puntale antischiacciamento",
       "Suola antiperforazione dove necessario",
       "Grembiuli e gambali per lavorazioni specifiche"],
  "g":"Calzatura di sicurezza in sezione con puntale e lamina antiperforazione evidenziati."},

 {"t":"Comportamenti sicuri: limiti dei DPI","n":"Art. 20 / Art. 78",
  "b":["I DPI non eliminano il pericolo, lo mitigano",
       "Non sostituiscono ripari e dispositivi",
       "Efficaci solo se indossati correttamente",
       "Controllo, manutenzione e sostituzione regolari"],
  "g":"Confronto: stesso operatore con DPI corretti vs uso scorretto, evidenziando gli errori."},

 {"t":"Sintesi e checklist","n":"D.Lgs. 81/08",
  "b":["Identificare pericoli e zone pericolose della macchina",
       "Verificare ripari, interblocchi e arresti di emergenza",
       "Applicare LOTO in manutenzione e pulizia",
       "Usare i DPI idonei e segnalare le anomalie",
       "Non manomettere mai i dispositivi di sicurezza"],
  "g":"Checklist finale a spunte gialle su fondo antracite, stile riepilogo operativo del modulo."},
]



# ---------------------------------------------------------------------------
# CONTENUTI - MODULO 2: MOVIMENTAZIONE MANUALE DEI CARICHI
# ---------------------------------------------------------------------------
M2 = [
 {"t":"Obiettivi del modulo","n":"Titolo VI D.Lgs. 81/08",
  "b":["Comprendere cosa si intende per MMC",
       "Conoscere l'anatomia del rachide e le patologie correlate",
       "Applicare il metodo NIOSH di valutazione",
       "Rispettare i limiti di peso per sesso ed eta'",
       "Adottare tecniche corrette di sollevamento"],
  "g":"Apertura modulo con silhouette di lavoratore che solleva un carico e icona della colonna vertebrale."},

 {"t":"Cos'e' la MMC","n":"Art. 167 D.Lgs. 81/08",
  "b":["Operazioni di trasporto o sostegno di un carico",
       "Comprende sollevare, deporre, spingere, tirare, portare",
       "Eseguite da uno o piu' lavoratori",
       "Con rischio dorso-lombare per condizioni sfavorevoli"],
  "g":"Quattro icone delle azioni MMC: sollevare, trasportare, spingere, tirare, in stile pittogramma."},

 {"t":"Quadro normativo MMC","n":"Art. 167-169",
  "b":["Titolo VI del D.Lgs. 81/08 dedicato alla MMC",
       "Art. 167: campo di applicazione e definizioni",
       "Art. 168: obblighi del datore di lavoro",
       "Art. 169: informazione, formazione e addestramento"],
  "g":"Schema degli articoli 167-169 con i rispettivi contenuti su fascia gialla/antracite."},

 {"t":"Allegato XXXIII","n":"All. XXXIII",
  "b":["Elenca gli elementi di riferimento e i fattori di rischio",
       "Caratteristiche del carico e sforzo fisico richiesto",
       "Caratteristiche dell'ambiente di lavoro",
       "Rimando alle norme tecniche ISO 11228"],
  "g":"Riproduzione stilizzata dell'Allegato XXXIII con le quattro categorie di fattori evidenziate."},

 {"t":"Attivita' comprese nella MMC","n":"Art. 167 D.Lgs. 81/08",
  "b":["Sollevamento e deposizione di carichi",
       "Trasporto manuale a breve e media distanza",
       "Spinta e traino di carrelli e oggetti",
       "Movimenti ripetitivi di carichi leggeri (rimando ISO)"],
  "g":"Collage di scenari lavorativi di MMC in magazzino, officina e logistica, stile illustrativo."},

 {"t":"Anatomia della colonna","n":"Riferimento ISO 11228",
  "b":["33-34 vertebre in cinque regioni",
       "Cervicale, dorsale, lombare, sacro, coccige",
       "La regione lombare sostiene i carichi maggiori",
       "Unita' funzionale: due vertebre e il disco"],
  "g":"Tavola anatomica della colonna vertebrale con le cinque regioni etichettate e colorate."},

 {"t":"Le curve fisiologiche","n":"Biomeccanica",
  "b":["Lordosi cervicale e lombare (concavita' posteriore)",
       "Cifosi dorsale e sacrale (convessita' posteriore)",
       "Le curve ammortizzano i carichi assiali",
       "Mantenere le curve riduce la pressione discale"],
  "g":"Profilo laterale della colonna con le curve fisiologiche e frecce di distribuzione del carico."},

 {"t":"I dischi intervertebrali","n":"Biomeccanica",
  "b":["Cuscinetti tra i corpi vertebrali",
       "Nucleo polposo centrale e anello fibroso",
       "Funzione di ammortizzazione e mobilita'",
       "Si deteriorano con sovraccarico e disidratazione"],
  "g":"Sezione di un disco intervertebrale con nucleo polposo e anello fibroso evidenziati."},

 {"t":"Il segmento L5-S1","n":"Biomeccanica",
  "b":["Cerniera tra colonna lombare e sacro",
       "Sopporta le sollecitazioni maggiori",
       "Punto critico nei sollevamenti scorretti",
       "Sede frequente di ernie e lombalgie"],
  "g":"Illustrazione biomeccanica del segmento L5-S1 con la compressione dei dischi sotto carico."},

 {"t":"Muscoli e legamenti","n":"Biomeccanica",
  "b":["Muscoli paravertebrali stabilizzano il rachide",
       "Addominali e core contribuiscono al sostegno",
       "Legamenti limitano i movimenti estremi",
       "Affaticamento muscolare aumenta il rischio"],
  "g":"Schema della muscolatura paravertebrale e del core con vettori di stabilizzazione."},

 {"t":"Biomeccanica del sollevamento","n":"Biomeccanica",
  "b":["Il rachide agisce come una leva sfavorevole",
       "Carico lontano dal corpo = braccio di leva maggiore",
       "Aumenta la forza sui muscoli e sui dischi",
       "Avvicinare il carico riduce la sollecitazione"],
  "g":"Diagramma a leva: carico, fulcro L5-S1 e forza muscolare con bracci di leva quotati."},

 {"t":"Carico sui dischi e leva","n":"Biomeccanica",
  "b":["A schiena flessa la pressione discale si moltiplica",
       "Pochi kg lontani equivalgono a forti compressioni",
       "Le torsioni aumentano lo stress sull'anello fibroso",
       "La postura corretta riduce il carico effettivo"],
  "g":"Confronto della pressione sul disco L5-S1 in posizione eretta, seduta e flessa con valori relativi."},

 {"t":"Patologie da sovraccarico","n":"All. XXXIII",
  "b":["Disturbi muscolo-scheletrici del rachide",
       "Lombalgie, lombosciatalgie, ernie discali",
       "Discopatie e alterazioni degenerative",
       "Spesso a insorgenza progressiva e cronica"],
  "g":"Mappa del corpo con evidenziazione delle aree dorso-lombari colpite dai disturbi MMC."},

 {"t":"Lombalgia acuta e cronica","n":"Sorveglianza sanitaria",
  "b":["Dolore localizzato nella regione lombare",
       "Forma acuta: insorgenza improvvisa (colpo della strega)",
       "Forma cronica: persistente o ricorrente",
       "Tra le principali cause di assenza dal lavoro"],
  "g":"Illustrazione di un lavoratore con dolore lombare e zona dolente evidenziata in rosso."},

 {"t":"Ernia del disco","n":"Biomeccanica",
  "b":["Fuoriuscita del nucleo polposo dall'anello fibroso",
       "Compressione delle radici nervose",
       "Dolore irradiato all'arto inferiore (sciatalgia)",
       "Favorita da sollevamenti scorretti e ripetuti"],
  "g":"Sezione che mostra un disco erniato con compressione della radice nervosa evidenziata."},

 {"t":"Fattori di rischio individuali","n":"All. XXXIII",
  "b":["Eta', sesso e caratteristiche fisiche",
       "Inadeguatezza fisica al compito",
       "Patologie pregresse del rachide",
       "Abbigliamento e calzature inadeguati"],
  "g":"Infografica dei fattori individuali con icone di eta', condizione fisica e abbigliamento."},

 {"t":"Fattori legati al carico","n":"All. XXXIII",
  "b":["Peso eccessivo o ingombrante",
       "Difficile da afferrare o instabile",
       "Contenuto che puo' spostarsi",
       "Spigoli, bordi taglienti o temperatura"],
  "g":"Carico con le caratteristiche di rischio annotate: peso, presa, instabilita', spigoli."},

 {"t":"Fattori legati allo sforzo","n":"All. XXXIII",
  "b":["Sforzo fisico eccessivo o brusco",
       "Torsione o flessione del tronco",
       "Carico distante dal corpo o in alto",
       "Posizione instabile del corpo"],
  "g":"Sequenza di posture scorrette (torsione, flessione, carico alto) con simbolo di divieto."},

 {"t":"Fattori legati all'ambiente","n":"All. XXXIII",
  "b":["Spazio libero insufficiente",
       "Pavimento irregolare o scivoloso",
       "Dislivelli, scale e gradini",
       "Microclima e illuminazione inadeguati"],
  "g":"Postazione di lavoro con criticita' ambientali evidenziate: pavimento, spazio, illuminazione."},

 {"t":"Fattori organizzativi","n":"All. XXXIII",
  "b":["Frequenza e durata delle movimentazioni",
       "Ritmi imposti non modificabili dal lavoratore",
       "Periodi di riposo o recupero insufficienti",
       "Distanze di trasporto eccessive"],
  "g":"Timeline di un turno con cicli di movimentazione e pause, evidenziando i recuperi."},

 {"t":"Metodi di valutazione","n":"ISO 11228",
  "b":["ISO 11228-1: sollevamento e trasporto",
       "ISO 11228-2: spinta e traino",
       "ISO 11228-3: movimentazione ad alta frequenza",
       "NIOSH come riferimento per il sollevamento"],
  "g":"Tabella dei metodi di valutazione con la serie ISO 11228 e ambito di applicazione."},

 {"t":"Il metodo NIOSH","n":"ISO 11228-1",
  "b":["Stima il peso limite raccomandato per un compito",
       "Considera geometria, frequenza e qualita' della presa",
       "Confronta il peso reale con quello raccomandato",
       "Fornisce un indice di rischio sintetico"],
  "g":"Schema a blocchi del metodo NIOSH: input geometrici, fattori e output (RWL e indice)."},

 {"t":"NIOSH: peso limite (RWL)","n":"ISO 11228-1",
  "b":["RWL = peso massimo raccomandato in quelle condizioni",
       "Parte da una costante di peso ideale",
       "Ridotta da sei fattori moltiplicativi",
       "Validi tutti i fattori tra 0 e 1"],
  "g":"Formula RWL = LC x HM x VM x DM x AM x FM x CM rappresentata in modo grafico e chiaro."},

 {"t":"NIOSH: la costante di peso","n":"ISO 11228-1",
  "b":["Costante di carico (LC) in condizioni ideali",
       "Valore di riferimento di 23 kg",
       "Riferito a popolazione lavorativa sana",
       "Punto di partenza del calcolo dell'RWL"],
  "g":"Illustrazione del sollevamento in condizioni ideali con il valore di 23 kg evidenziato."},

 {"t":"NIOSH: i fattori moltiplicativi","n":"ISO 11228-1",
  "b":["HM distanza orizzontale; VM altezza di presa",
       "DM dislocazione verticale; AM angolo di asimmetria",
       "FM frequenza dei sollevamenti",
       "CM qualita' della presa del carico"],
  "g":"Pannello con i sei fattori NIOSH, ciascuno con un'icona e una breve descrizione."},

 {"t":"NIOSH: indice di sollevamento","n":"ISO 11228-1",
  "b":["LI = Peso effettivo / Peso limite (RWL)",
       "Misura sintetica del livello di rischio",
       "Calcolabile per singolo compito o composito",
       "Guida le priorita' di intervento"],
  "g":"Formula LI = Peso / RWL con un indicatore a semaforo del livello di rischio."},

 {"t":"Interpretare l'indice (LI)","n":"ISO 11228-1",
  "b":["LI <= 1: rischio accettabile per la maggioranza",
       "1 < LI <= 3: rischio presente, intervenire",
       "LI > 3: rischio elevato, azione prioritaria",
       "Obiettivo: riportare LI il piu' vicino possibile a 1"],
  "g":"Scala a colori dell'indice LI: verde (<=1), giallo (1-3), rosso (>3) con soglie indicate."},

 {"t":"Massa di riferimento","n":"All. XXXIII / ISO 11228",
  "b":["Valore di peso usato come riferimento di legge",
       "Per la popolazione adulta sana: 25 kg (uomini)",
       "Da intendersi in condizioni ottimali",
       "Si riduce in presenza di fattori sfavorevoli"],
  "g":"Confronto visivo tra masse di riferimento per fasce di popolazione, con valori in kg."},

 {"t":"Limiti per sesso ed eta'","n":"ISO 11228-1",
  "b":["Maschi adulti: riferimento circa 25 kg",
       "Femmine adulte: riferimento circa 20 kg",
       "Giovani e anziani: valori ridotti",
       "Tutela rafforzata per minori e gestanti"],
  "g":"Tabella dei limiti di peso per sesso e fascia d'eta' con barre proporzionali ai valori."},

 {"t":"Spinta e traino","n":"ISO 11228-2",
  "b":["Valutare la forza iniziale e di mantenimento",
       "Preferire la spinta al traino quando possibile",
       "Maniglie all'altezza corretta e buona aderenza",
       "Manutenzione di ruote e pavimentazioni"],
  "g":"Operatore che spinge un carrello con vettori di forza e altezza corretta delle maniglie."},

 {"t":"Tecnica corretta di sollevamento","n":"Buone pratiche",
  "b":["Piedi stabili e leggermente divaricati",
       "Piegare le ginocchia, schiena in posizione neutra",
       "Carico vicino al corpo, presa salda",
       "Sollevare con le gambe, senza torsioni"],
  "g":"Sequenza passo-passo del sollevamento corretto con ginocchia piegate e schiena dritta."},

 {"t":"Errori comuni","n":"Buone pratiche",
  "b":["Sollevare a schiena flessa e gambe tese",
       "Ruotare il tronco con il carico in mano",
       "Sollevare carichi troppo distanti dal corpo",
       "Movimenti bruschi e senza preparazione"],
  "g":"Confronto affiancato tra sollevamento errato e corretto, con gli errori barrati in rosso."},

 {"t":"Trasporto e deposito","n":"Buone pratiche",
  "b":["Mantenere il carico vicino e centrato",
       "Percorsi liberi e ben illuminati",
       "Deposito senza torsioni, accompagnando il carico",
       "Altezze di presa e deposito ergonomiche"],
  "g":"Sequenza di trasporto e deposito con percorso libero e altezze di lavoro ottimali."},

 {"t":"Organizzazione del posto","n":"Ergonomia",
  "b":["Posizionare i carichi tra le ginocchia e le spalle",
       "Ridurre distanze, dislivelli e torsioni",
       "Alternare i compiti e introdurre pause",
       "Adeguare ritmi e frequenze"],
  "g":"Postazione ergonomica con zona di presa ottimale evidenziata tra ginocchia e spalle."},

 {"t":"Ausili meccanici","n":"Art. 168 D.Lgs. 81/08",
  "b":["Carrelli, transpallet e sollevatori",
       "Nastri, rulliere e tavoli regolabili",
       "Manipolatori e bracci servoassistiti",
       "Priorita': eliminare o ridurre la MMC"],
  "g":"Galleria di ausili meccanici: transpallet, sollevatore, manipolatore, con didascalie tecniche."},

 {"t":"Sorveglianza sanitaria","n":"Art. 168 D.Lgs. 81/08",
  "b":["Prevista quando la valutazione evidenzia il rischio",
       "Affidata al medico competente",
       "Visite preventive e periodiche",
       "Giudizio di idoneita' alla mansione"],
  "g":"Ambulatorio del medico competente con cartella sanitaria e schema delle visite periodiche."},

 {"t":"Informazione e formazione","n":"Art. 169 D.Lgs. 81/08",
  "b":["Informazione sui rischi e sul peso dei carichi",
       "Formazione sulle modalita' corrette di MMC",
       "Addestramento pratico alle tecniche",
       "Coinvolgimento di RLS e medico competente"],
  "g":"Sessione di addestramento pratico alla MMC con istruttore e lavoratori, taglio fotografico."},

 {"t":"Sintesi e buone pratiche","n":"Titolo VI D.Lgs. 81/08",
  "b":["Valutare sempre il rischio prima di movimentare",
       "Usare ausili e ridurre i pesi alla fonte",
       "Applicare le tecniche corrette di sollevamento",
       "Rispettare i limiti per sesso ed eta'",
       "Partecipare a sorveglianza sanitaria e formazione"],
  "g":"Riepilogo finale a spunte gialle su fondo antracite con le regole d'oro della MMC."},
]


# ---------------------------------------------------------------------------
# ASSEMBLAGGIO PRESENTAZIONE (100 slide)
# ---------------------------------------------------------------------------
def split_bullets(items):
    out = []
    for s in items:
        if s.startswith(">"):
            out.append((s[1:].strip(), 1))
        else:
            out.append((s, 0))
    return out


def main():
    slides = []
    total = 100

    # Slide 1 - Cover (parte del Modulo 1)
    slides.append(slide_cover(
        title="Rischio Meccanico e Movimentazione Manuale dei Carichi",
        subtitle="Corso di formazione e informazione dei lavoratori",
        footer="D.Lgs. 81/08 - Titolo III, Titolo VI - Direttiva Macchine 2006/42/CE",
        kicker="SICUREZZA SUL LAVORO"))

    # Slide 2 - Divider Modulo 1
    slides.append(slide_divider(
        modulo="MODULO 1",
        title="Rischio Meccanico",
        points=["Pericoli meccanici e zone pericolose",
                "Ripari, dispositivi e interblocchi",
                "Manutenzione in sicurezza e LOTO",
                "DPI per il rischio meccanico"],
        slide_no=2, total=total))

    # Slide 3..60 - Contenuti Modulo 1 (58 slide)
    no = 3
    for item in M1:
        slides.append(slide_content(
            title=item["t"], norm=item.get("n", ""),
            bullets=split_bullets(item["b"]), graphic=item["g"],
            footer_left=FOOT1, slide_no=no, total=total))
        no += 1

    # Slide 61 - Divider Modulo 2
    slides.append(slide_divider(
        modulo="MODULO 2",
        title="Movimentazione Manuale dei Carichi",
        points=["Anatomia del rachide e patologie",
                "Metodo NIOSH e indice di sollevamento",
                "Limiti di peso per sesso ed eta'",
                "Tecniche corrette e ausili meccanici"],
        slide_no=61, total=total))

    # Slide 62..99 - Contenuti Modulo 2 (38 slide)
    no = 62
    for item in M2:
        slides.append(slide_content(
            title=item["t"], norm=item.get("n", ""),
            bullets=split_bullets(item["b"]), graphic=item["g"],
            footer_left=FOOT2, slide_no=no, total=total))
        no += 1

    # Slide 100 - Chiusura
    slides.append(slide_closing(
        title="Messaggi chiave",
        points=["La sicurezza inizia dalla conoscenza dei pericoli",
                "Ripari e dispositivi non vanno mai manomessi",
                "In manutenzione: sempre procedura LOTO",
                "Movimentare i carichi con tecnica corretta e ausili",
                "Segnalare ogni anomalia: la prevenzione e' di tutti"],
        footer="Grazie per l'attenzione - Formazione lavoratori D.Lgs. 81/08"))

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "SLIDE RISCHIO MECCANICO.pptx")
    build_pptx(slides, out)
    print(f"Creato: {out}")
    print(f"Numero slide: {len(slides)}")


if __name__ == "__main__":
    main()
