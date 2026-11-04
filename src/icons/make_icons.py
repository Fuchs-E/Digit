"""Draws the Digit number icons (32x32 PNG, one per symbol and colour) into this folder.
The build packs every PNG here into the add-in; presets refer to them as "digit:<file name>".
Run with Pillow installed: python make_icons.py"""
from PIL import Image, ImageDraw, ImageFont
import os
S=256
BOLD='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
COLORS={'green':(16,110,60,255),'blue':(31,78,140,255)}
# (id, text lines). Up to two lines; sized to fit.
ICONS=[('hash','#'),('int','0'),('dec1','0.0'),('dec2','0.00'),('thou','#,0'),('thou1','#,0.0'),
       ('k','K'),('m','M'),('neg','(1)'),('dollar','$'),('dollar2','$.00'),
       ('pct','%'),('pct1','0.0%'),('mult','1x'),('sci','1E3'),('text','@'),
       ('plus','+/-'),('date','31'),('month','mmm')]
def fit(d,t,maxw,maxh,start=150):
    sz=start
    while sz>20:
        f=ImageFont.truetype(BOLD,sz); b=d.textbbox((0,0),t,font=f)
        if b[2]-b[0]<=maxw and b[3]-b[1]<=maxh: return f,b
        sz-=4
    return f,b
def draw(t,color):
    im=Image.new('RGBA',(S,S),(0,0,0,0)); d=ImageDraw.Draw(im)
    d.rounded_rectangle([8,8,S-8,S-8],radius=40,fill=color)
    if t=='31':   # little calendar: white page with a band
        d.rounded_rectangle([52,58,204,206],radius=14,fill='white')
        d.rectangle([52,58,204,96],fill=(220,220,220)); 
        f,b=fit(d,'31',120,90,110); d.text((128-(b[2]-b[0])/2-b[0],152-(b[3]-b[1])/2-b[1]),'31',font=f,fill=color); return im
    f,b=fit(d,t,196,150)
    d.text((S/2-(b[2]-b[0])/2-b[0],S/2-(b[3]-b[1])/2-b[1]),t,font=f,fill='white')
    return im
OUT=os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(OUT,'form'),exist_ok=True)
tiles={}
for cname,c in COLORS.items():
    for iid,t in ICONS:
        im=draw(t,c).resize((32,32),Image.LANCZOS); im.save(os.path.join(OUT,f'{iid}_{cname}.png'))
        # Same picture on white as a 24-bit BMP, for the settings window (forms
        # cannot show transparency, and LoadPicture reads BMP but not PNG).
        flat=Image.new('RGB',(32,32),'white'); flat.paste(im,(0,0),im)
        flat.save(os.path.join(OUT,'form',f'{iid}_{cname}.bmp')); tiles[(iid,cname)]=im
