"""
Genera un logo personalizzato per l'ebook di PROTA DOMENICO.
Tema: blockchain / cripto - stile moderno con catena di blocchi e iniziali "PD".
"""
from PIL import Image, ImageDraw, ImageFont
import math
import os

OUT = "/projects/sandbox/ITALIC/logo_prota_domenico.png"

# Tela quadrata ad alta risoluzione
SIZE = 1024
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Palette: oro Bitcoin + blu blockchain + sfondo scuro
GOLD = (247, 147, 26, 255)        # arancione Bitcoin
GOLD_LIGHT = (255, 200, 80, 255)
BLUE = (28, 35, 90, 255)          # blu profondo
BLUE_LIGHT = (52, 104, 192, 255)
WHITE = (255, 255, 255, 255)
DARK = (15, 20, 45, 255)

# 1) Sfondo: cerchio gradiente (simulato con cerchi concentrici)
cx, cy = SIZE // 2, SIZE // 2
R = SIZE // 2 - 10
for i in range(R, 0, -1):
    t = i / R
    r = int(DARK[0] * t + BLUE[0] * (1 - t))
    g = int(DARK[1] * t + BLUE[1] * (1 - t))
    b = int(DARK[2] * t + BLUE[2] * (1 - t))
    draw.ellipse([cx - i, cy - i, cx + i, cy + i], fill=(r, g, b, 255))

# 2) Anello esterno dorato
ring_w = 18
draw.ellipse([cx - R, cy - R, cx + R, cy + R], outline=GOLD, width=ring_w)

# 3) Catena di blocchi attorno (esagoni piccoli a corona)
N_BLOCKS = 8
ring_radius = R - 70
block_size = 70
for i in range(N_BLOCKS):
    ang = (2 * math.pi * i) / N_BLOCKS - math.pi / 2
    bx = cx + int(ring_radius * math.cos(ang))
    by = cy + int(ring_radius * math.sin(ang))
    # Esagono
    pts = []
    for k in range(6):
        a = math.pi / 6 + k * math.pi / 3
        pts.append((bx + block_size / 2 * math.cos(a),
                    by + block_size / 2 * math.sin(a)))
    color = GOLD if i % 2 == 0 else GOLD_LIGHT
    draw.polygon(pts, fill=color, outline=WHITE)

# Linee che collegano i blocchi (catena)
for i in range(N_BLOCKS):
    ang1 = (2 * math.pi * i) / N_BLOCKS - math.pi / 2
    ang2 = (2 * math.pi * (i + 1)) / N_BLOCKS - math.pi / 2
    x1 = cx + int(ring_radius * math.cos(ang1))
    y1 = cy + int(ring_radius * math.sin(ang1))
    x2 = cx + int(ring_radius * math.cos(ang2))
    y2 = cy + int(ring_radius * math.sin(ang2))
    draw.line([(x1, y1), (x2, y2)], fill=GOLD_LIGHT, width=6)

# 4) Cerchio interno scuro per le iniziali
inner_r = R - 200
draw.ellipse([cx - inner_r, cy - inner_r, cx + inner_r, cy + inner_r],
             fill=DARK, outline=GOLD, width=10)

# 5) Iniziali "PD" stilizzate
def get_font(size):
    candidates = [
        "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]
    for c in candidates:
        if os.path.exists(c):
            return ImageFont.truetype(c, size)
    return ImageFont.load_default()

font_big = get_font(360)
text = "PD"
bbox = draw.textbbox((0, 0), text, font=font_big)
tw = bbox[2] - bbox[0]
th = bbox[3] - bbox[1]
# offset di disegno (tiene conto del bbox iniziale)
tx = cx - tw // 2 - bbox[0]
ty = cy - th // 2 - bbox[1] - 20
# Ombra
draw.text((tx + 6, ty + 6), text, font=font_big, fill=(0, 0, 0, 180))
# Testo principale (oro)
draw.text((tx, ty), text, font=font_big, fill=GOLD)

# 6) Sottotesto "CRYPTO GUIDE"
font_sub = get_font(58)
sub = "CRYPTO GUIDE"
bb2 = draw.textbbox((0, 0), sub, font=font_sub)
sw = bb2[2] - bb2[0]
draw.text((cx - sw // 2 - bb2[0], cy + 150), sub, font=font_sub, fill=WHITE)

# 7) Simbolo Bitcoin (₿) in alto e Ethereum (Ξ) in basso, piccoli
font_sym = get_font(70)
draw.text((cx - 30, cy - 270), "\u20BF", font=font_sym, fill=GOLD)        # ₿
draw.text((cx - 25, cy + 230), "\u039E", font=font_sym, fill=BLUE_LIGHT)  # Ξ (Xi maiuscolo come Ether)

img.save(OUT, "PNG")
print(f"Logo salvato in: {OUT}")
print(f"Dimensione: {os.path.getsize(OUT)} byte")
