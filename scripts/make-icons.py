import sys, math
# Maakt icons/icon{16,32,48,128}.png (tegel in Teams-paars met wit vinkje). Gebruik: python3 scripts/make-icons.py icons  (vereist Pillow)
from PIL import Image, ImageDraw, ImageFilter, ImageChops
out=sys.argv[1]
N=1024
def squircle_mask(n, inset, k=5.0):
    m=Image.new("L",(n,n),0); d=ImageDraw.Draw(m)
    c=n/2; r=n/2-inset; pts=[]
    for i in range(720):
        t=2*math.pi*i/720; ct,st=math.cos(t),math.sin(t)
        x=c+r*math.copysign(abs(ct)**(2/k),ct); y=c+r*math.copysign(abs(st)**(2/k),st)
        pts.append((x,y))
    d.polygon(pts,fill=255); return m
def vgrad(n, top, bot):
    g=Image.new("RGB",(1,n))
    for y in range(n):
        t=y/(n-1); g.putpixel((0,y),tuple(round(top[i]+(bot[i]-top[i])*t) for i in range(3)))
    return g.resize((n,n))
def render(size, detail=True):
    inset = N*0.06 if size>=48 else N*0.02
    mask=squircle_mask(N, inset)
    img=Image.new("RGBA",(N,N),(0,0,0,0))
    bg=vgrad(N,(124,127,226),(68,71,160)).convert("RGBA")
    img.paste(bg,(0,0),mask)
    if detail:
        # zachte glans bovenaan (glasachtig)
        hl=Image.new("L",(N,N),0); hd=ImageDraw.Draw(hl)
        hd.ellipse((-N*0.25,-N*0.75,N*1.25,N*0.42),fill=70)
        hl=hl.filter(ImageFilter.GaussianBlur(N*0.04))
        hl=ImageChops.multiply(hl,mask)
        white=Image.new("RGBA",(N,N),(255,255,255,0)); white.putalpha(hl)
        img=Image.alpha_composite(img,white)
        # lichtrand binnenin
        edge=ImageChops.subtract(mask, squircle_mask(N, inset+N*0.008))
        e=Image.new("RGBA",(N,N),(255,255,255,0)); e.putalpha(edge.point(lambda v: v*0.35))
        img=Image.alpha_composite(img,e)
    # vinkje
    w = N*(0.115 if size>=48 else 0.16)
    pts=[(N*0.29,N*0.53),(N*0.445,N*0.675),(N*0.72,N*0.355)] if size>=48 else [(N*0.24,N*0.53),(N*0.43,N*0.71),(N*0.77,N*0.32)]
    if detail:
        sh=Image.new("L",(N,N),0); sd=ImageDraw.Draw(sh)
        sd.line([(x,y+N*0.018) for x,y in pts],fill=90,width=int(w),joint="curve")
        for x,y in (pts[0],pts[-1]): sd.ellipse((x-w/2,y+N*0.018-w/2,x+w/2,y+N*0.018+w/2),fill=90)
        sh=sh.filter(ImageFilter.GaussianBlur(N*0.02)); sh=ImageChops.multiply(sh,mask)
        s=Image.new("RGBA",(N,N),(30,30,90,0)); s.putalpha(sh); img=Image.alpha_composite(img,s)
    ck=Image.new("RGBA",(N,N),(0,0,0,0)); cd=ImageDraw.Draw(ck)
    cd.line(pts,fill=(255,255,255,255),width=int(w),joint="curve")
    for x,y in (pts[0],pts[-1]): cd.ellipse((x-w/2,y-w/2,x+w/2,y+w/2),fill=(255,255,255,255))
    img=Image.alpha_composite(img,ck)
    return img.resize((size,size),Image.LANCZOS)
for s in (16,32,48,128):
    render(s, detail=s>=32).save(f"{out}/icon{s}.png", optimize=True)
render(512).save(f"{out}/icon512.png")
print("ok")
