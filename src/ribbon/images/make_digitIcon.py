"""Draws digitIcon.png, the Digit menu button: a dark green tile with # $ % ,
Run from this folder with Pillow installed: python make_digitIcon.py"""
from PIL import Image, ImageDraw, ImageFont
S=256; GREEN=(16,110,60,255)
F=ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',96)
def make(gap):
    im=Image.new('RGBA',(S,S),(0,0,0,0)); d=ImageDraw.Draw(im)
    d.rounded_rectangle([8,8,S-8,S-8], radius=40, fill=GREEN)
    # Two rows sharing a baseline each; two columns centred close to the middle.
    asc,desc=F.getmetrics()
    cap=d.textbbox((0,0),"#",font=F); cap_h=cap[3]-cap[1]
    row_h=cap_h+gap
    top=(S-(2*cap_h+gap))/2            # block of two cap-height rows, centred
    cols=[S/2-gap/2-cap_h*0.42, S/2+gap/2+cap_h*0.42]
    for r,row in enumerate([("#","$"),("%",",")]):
        baseline=top+cap_h+r*row_h     # bottom of the cap height for this row
        for c,t in enumerate(row):
            b=d.textbbox((0,0),t,font=F,anchor="ls")
            w=b[2]-b[0]
            y=baseline
            if t==",":                     # centre the comma in its corner, level with the middle of %
                y=baseline-cap_h/2-((b[1]+b[3])/2)
            d.text((cols[c]-w/2-b[0], y), t, font=F, fill='white', anchor="ls")
    return im
v=make(26)
v.resize((32,32),Image.LANCZOS).save('digitIcon.png')
