import zipfile,re,json,sys,os,copy,io
from lxml import etree
from PIL import Image
W="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"; WP="{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A="{http://schemas.openxmlformats.org/drawingml/2006/main}"; R="{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG="{http://schemas.openxmlformats.org/package/2006/relationships}"; XML="{http://www.w3.org/XML/1998/namespace}"
src,dst=sys.argv[1],sys.argv[2]; cfg=json.loads(sys.argv[3]) if len(sys.argv)>3 else {}
T="/tmp/claude-0/-home-user-Wealth-managment-software/4853c681-1399-5e51-b2e7-c4edae34dd33/scratchpad/tr/"
FG="/home/user/Wealth-managment-software/secciones/valuacion/graficos/figuras/"
HERE=os.getcwd()
os.chdir(T); sys.path.insert(0,"."); from imgs import componer; os.chdir(HERE)
EN=json.load(open(T+"en_imgs.json"))
def ent(prefix): return next(e for e in EN if e["title"] and e["title"].startswith(prefix))
def sp(t): return re.sub(r"^Figure A?\d+[A-Z]?(?:\.\d)?:\s*","",t)
z=zipfile.ZipFile(src); root=etree.fromstring(z.read("word/document.xml")); body=root.find(W+"body")
rels=etree.fromstring(z.read("word/_rels/document.xml.rels"))
rid2={r.get("Id"):r.get("Target").split("/")[-1] for r in rels}
def txt(e): return "".join(t.text or "" for t in e.iter(W+"t"))
# 1) renumerar referencias y títulos de texto: 25..37 -> +1
MAP={str(n):str(n+1) for n in range(25,38)}
RE=re.compile(r"\b(Figures?)(\s+)((?:\d+[AB]?)(?:(?:\s*,\s*|\s+and\s+|\s*[–-]\s*|\s*&\s*)\d+[AB]?)*)(?![\.\d]*\d)")
NUM=re.compile(r"\d+[AB]?")
def remap(m): return m.group(1)+m.group(2)+NUM.sub(lambda n:MAP.get(n.group(0),n.group(0)),m.group(3))
nref=0
for p in root.iter(W+"p"):
    ts=list(p.iter(W+"t")); full="".join(t.text or "" for t in ts)
    if "Figure" not in full or not RE.search(full): continue
    spans=[];pos=0
    for t in ts: spans.append((pos,pos+len(t.text or ""))); pos+=len(t.text or "")
    out=[t.text or "" for t in ts]
    ms=[m for m in RE.finditer(full) if remap(m)!=m.group(0)]
    # trabajar de atrás hacia adelante sobre el texto concatenado y reasignar
    new=full
    for m in reversed(ms): new=new[:m.start()]+remap(m)+new[m.end():]; nref+=1
    if new==full: continue
    # reparto: todo el texto nuevo en los runs, conservando los límites donde no hay cambio
    # estrategia: si hay cambios, poner por segmentos usando diferencias de longitud por run
    delta=0; ptr=0; res=[]
    marks=sorted([(m.start(),m.end(),remap(m)) for m in ms])
    for (s,e),t in zip(spans,ts):
        piece=full[s:e]; o=0; buf=""
        for a,b,rep in marks:
            if b<=s or a>=e: continue
            if a>=s and b<=e: buf+=full[s+o:a]+rep; o=b-s
            elif a>=s: buf+=full[s+o:a]+rep; o=e-s      # empieza aquí, termina en otro run
            else: o=max(o,b-s) if b<=e else e-s          # continúa de un run anterior
        buf+=piece[o:]; res.append(buf)
    for t,v in zip(ts,res): t.text=v; t.set(XML+"space","preserve")
print("referencias renumeradas:",nref)
# la cita de múltiplos apuntaba (provisoriamente) a la Figura 20: ahora es la 25 nueva
for t in root.iter(W+"t"):
    if t.text and "sits below the price (Figure 20)" in t.text: t.text=t.text.replace("sits below the price (Figure 20)","sits below the price (Figure 25)"); t.set(XML+"space","preserve"); print("cita múltiplos → Figure 25")
# 2) imágenes con número integrado
dims={}
for n in z.namelist():
    m=re.match(r"word/media/image(\d+)\.png",n)
    if m: dims[int(m.group(1))]=Image.open(io.BytesIO(z.read(n))).size
e21A,e21B=ent("Figure 21A"),ent("Figure 21B")
SPEC={25:("26",sp(e21A["title"]),e21A["src"],FG+"en_V18_caprate_spread.png"),
 26:("27",sp(e21B["title"]),e21B["src"],FG+"en_V21_mc_cap_rate.png"),
 27:("28","Mallplaza rose ~4x in five years, ahead of the IPSA; the last leg came with no change in the business",["Weekly closing price, base 100 = 1-Oct-21. Adjusted for the 2024 capital increase. IPSA to 31-Aug-26 (latest available). Milestones: daily closes. Source: Bloomberg; Team analysis."],FG+"en_V19_precio_base100.png"),
 28:("29","Even adding back all deferred taxes, the price is 27% above NAV",["Price / adjusted NAV (EPRA NTA). Adjusted NAV = attributable equity + deferred tax liabilities − goodwill − intangibles. Balance sheet at 30-Jun-26; market cap at 30-Sep-26. Source: Bloomberg; CMF financial statements; Team analysis."],FG+"en_V20_pnav.png"),
 29:("30",sp(ent("Figure 27:")["title"]),ent("Figure 27:")["src"],T+"mEN/word/media/"+ent("Figure 27:")["img"]),
 30:("31",sp(ent("Figure 26:")["title"]),ent("Figure 26:")["src"],T+"mEN/word/media/"+ent("Figure 26:")["img"]),
 35:("36","With historical volatility, only 5% of 52-week price paths end below our target price",["10,000 paths from CLP 3,555. Resampling of weekly returns (Oct-21 to Sep-26), centered at zero; annual volatility ~32%. Source: Bloomberg; Team analysis."],FG+"en_V22_mc_precio_52s.png"),
 36:("37","Bear, base and bull: even with everything in our favor the target price stays 18% below the price (CLP per share)",["Bear: everything that could go wrong, goes wrong; base: everything as we project it; bull: everything in our favor (Appendix 18). Team illustrations. Source: Team analysis."],FG+"en_V10_escenarios.png"),
 37:("38",sp(ent("Figure 34:")["title"]),ent("Figure 34:")["src"],T+"mEN/word/media/"+ent("Figure 34:")["img"])}
os.makedirs("add25",exist_ok=True); nuevas={}
os.chdir(T)
for iid,(n,t,s,ch) in SPEC.items():
    W_,H_=dims[iid]; out=f"{HERE}/add25/image{iid}.png"; componer(W_,H_,f"Figure {n}: {t}",s,ch,out); nuevas[f"word/media/image{iid}.png"]=out
e25=ent("Figure 25:")
componer(500,540,f"Figure 25: {sp(e25['title'])}",e25["src"],FG+"en_V9_pares.png",f"{HERE}/add25/peers.png")
os.chdir(HERE)
nuevas["word/media/image_peers.png"]=f"{HERE}/add25/peers.png"
# 3) relación y anclaje de la figura nueva
rid="rIdPeers25"; rid2[rid]="image_peers.png"
r=etree.SubElement(rels,PKG+"Relationship"); r.set("Id",rid); r.set("Type","http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"); r.set("Target","media/image_peers.png")
def dr_by(n):
    for dr in body.iter(W+"drawing"):
        if rid2[next(dr.iter(A+"blip")).get(R+"embed")]==f"image{n}.png": return dr
molde=dr_by(24); nuevo=copy.deepcopy(molde.getparent())   # <w:r> con el dibujo
d=nuevo.find(W+"drawing"); anc=d.find(WP+"anchor")
next(d.iter(A+"blip")).set(R+"embed",rid)
cx=int(2.47*914400*cfg.get("w",0.9)); cy=int(cx*540/500)
for e in d.iter(WP+"extent",A+"ext"):
    if e.get("cx"): e.set("cx",str(cx)); e.set("cy",str(cy))
dp=anc.find(WP+"docPr"); dp.set("id","9025"); dp.set("name","Figure 25 peers"); 
for e in d.iter("{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr"): e.set("id","9025"); e.set("name","image_peers.png")
ph=anc.find(WP+"positionH"); pv=anc.find(WP+"positionV")
ph.set("relativeFrom","page"); [ph.remove(c) for c in list(ph)]; etree.SubElement(ph,WP+"posOffset").text=str(5179060+2254250-cx)
pv.set("relativeFrom","page"); [pv.remove(c) for c in list(pv)]; etree.SubElement(pv,WP+"posOffset").text=str(int(cfg.get("y",262)*12700))
cita=next(p for p in body.iter(W+"p") if "Every multiple" in txt(p))
cita.append(nuevo)
# 4) columna de la página 7: reducir 24, 21A, 21B y bajar 21A / 21B
def esc(n,f):
    for e in dr_by(n).iter(WP+"extent",A+"ext"):
        if e.get("cx"): e.set("cx",str(round(int(e.get("cx"))*f))); e.set("cy",str(round(int(e.get("cy"))*f)))
def mov(n,pt):
    o=dr_by(n).find(WP+"anchor").find(WP+"positionV").find(WP+"posOffset"); o.text=str(int(o.text)+int(round(pt*12700)))
for n in (24,25,26): esc(n,cfg.get("e",0.9))
for n,pt in cfg.get("mov",{}).items(): mov(int(n),pt)
o=zipfile.ZipFile(dst,"w",zipfile.ZIP_DEFLATED)
for it in z.infolist():
    if it.filename=="word/document.xml": o.writestr(it,etree.tostring(root,xml_declaration=True,encoding="UTF-8",standalone=True))
    elif it.filename=="word/_rels/document.xml.rels": o.writestr(it,etree.tostring(rels,xml_declaration=True,encoding="UTF-8",standalone=True))
    elif it.filename in nuevas: o.writestr(it,open(nuevas[it.filename],"rb").read())
    else: o.writestr(it,z.read(it.filename))
o.writestr("word/media/image_peers.png",open(nuevas["word/media/image_peers.png"],"rb").read())
o.close(); print("ok")
