"""Draws colourIcon.png, the Colour by content button: a letter A over a bar in
the four standard modelling colours (blue input, black formula, green sheet
link, dark red external link).
Run from this folder with Pillow installed: python make_colourIcon.py"""
from PIL import Image, ImageDraw, ImageFont
S = 256
COLOURS = [(0, 0, 255, 255), (0, 0, 0, 255), (0, 128, 0, 255), (192, 0, 0, 255)]
F = ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', 190)
im = Image.new('RGBA', (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(im)
# The letter, centred above the bar.
b = d.textbbox((0, 0), 'A', font=F, anchor='ls')
d.text((S / 2 - (b[0] + b[2]) / 2, 168), 'A', font=F, fill=(64, 64, 64, 255), anchor='ls')
# The bar: four equal blocks.
left, right, top, bottom = 16, S - 16, 188, 244
w = (right - left) / 4
for i, c in enumerate(COLOURS):
    d.rectangle([left + i * w, top, left + (i + 1) * w, bottom], fill=c)
im.resize((32, 32), Image.LANCZOS).save('colourIcon.png')
# The same picture as one of Digit's own icons ("digit:colour"), which is what
# a quick-button position shows when Colour by content takes its place.
import os
ICONS = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'icons')
small = Image.open('colourIcon.png')
small.save(os.path.join(ICONS, 'colour.png'))
flat = Image.new('RGB', (32, 32), 'white')
flat.paste(small, (0, 0), small)
flat.save(os.path.join(ICONS, 'form', 'colour.bmp'))
