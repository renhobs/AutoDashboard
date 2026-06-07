#!/usr/bin/env python3
"""
Blankt das Kennzeichen im Seat Leon Foto weiß aus.
Aufruf: python3 scripts/blank_plate.py
Erwartet: img/seat-leon-mk4.png
Schreibt: img/seat-leon-mk4.png (überschreibt)
"""
import sys
from pathlib import Path

try:
    from PIL import Image, ImageDraw
except ImportError:
    sys.exit("Pillow nicht installiert — bitte: pip3 install Pillow")

SRC = Path(__file__).parent.parent / "img" / "seat-leon-mk4.png"
if not SRC.exists():
    sys.exit(f"Datei nicht gefunden: {SRC}\nBitte zuerst das Bild als img/seat-leon-mk4.png speichern.")

img = Image.open(SRC).convert("RGBA")
w, h = img.size
print(f"Bildgröße: {w}x{h}")

draw = ImageDraw.Draw(img)

# Kennzeichen-Koordinaten (relativ zur Bildgröße 1536x864 bei 4:3 crop)
# Das Kennzeichen liegt bei ca. x: 12%–37%, y: 57%–65% des Bildes
x1 = int(w * 0.115)
y1 = int(h * 0.565)
x2 = int(w * 0.375)
y2 = int(h * 0.650)

# Weiß drüber malen (RGBA: weiß, voll opak)
draw.rectangle([x1, y1, x2, y2], fill=(255, 255, 255, 255))

# Speichern
img.save(SRC, "PNG")
print(f"Kennzeichen geblankiert: [{x1},{y1}] – [{x2},{y2}]")
print(f"Gespeichert: {SRC}")
