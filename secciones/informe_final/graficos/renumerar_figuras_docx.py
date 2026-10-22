import zipfile,re,json,sys,os
from lxml import etree
W="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML="{http://www.w3.org/XML/1998/namespace}"
src,dst=sys.argv[1],sys.argv[2]
MAP={"21A":"25","21B":"26","22A":"27","22B":"28","27":"29","26":"30","28":"31","29":"32","30":"33","31":"34","32":"35","33":"36","34":"37"}
RE=re.compile(r"\b(Figures?)(\s+)((?:\d+[AB]?)(?:(?:\s*,\s*|\s+and\s+|\s*[–-]\s*|\s*&\s*)\d+[AB]?)*)(?![\.\d]*\d)")
NUM=re.compile(r"\d+[AB]?")
def remap(m):
    nums=NUM.sub(lambda n:MAP.get(n.group(0),n.group(0)),m.group(3))
    return m.group(1)+m.group(2)+nums
z=zipfile.ZipFile(src); root=etree.fromstring(z.read("word/document.xml"))
cambios=[]
def sustituir(p):
    ts=[t for t in p.iter(W+"t")]
    if not ts: return
    full="".join(t.text or "" for t in ts)
    if "Figure" not in full: return
    new_full=RE.sub(remap,full)
    if new_full==full: return
    # reasignar por segmentos: mapear cada coincidencia a rangos de caracteres
    spans=[]; pos=0
    for t in ts: spans.append((pos,pos+len(t.text or ""))); pos+=len(t.text or "")
    out=[t.text or "" for t in ts]
    off=0
    for m in RE.finditer(full):
        a,b=m.start(),m.end(); rep=remap(m)
        if rep==m.group(0): continue
        cambios.append((m.group(0),rep))
        # primer run afectado recibe el reemplazo completo; los demás pierden su parte
        first=True
        for i,(s,e) in enumerate(spans):
            lo,hi=max(a,s),min(b,e)
            if lo>=hi: continue
            seg=out[i]
            ls,le=lo-s,hi-s
            # posiciones en out[i] pueden haberse desplazado por reemplazos previos: se trabaja con el texto original
            orig=(ts[i].text or "")
            if first: out[i]=out[i].replace(orig[ls:le],rep,1) if orig[ls:le] in out[i] else out[i]; first=False
            else: out[i]=out[i].replace(orig[ls:le],"",1)
    for t,o in zip(ts,out):
        t.text=o; t.set(XML+"space","preserve")
for p in root.iter(W+"p"): sustituir(p)
print(len(cambios),"referencias cambiadas")
for a,b in cambios: print("  ",a,"→",b)
# imágenes con el número dentro
sys.path.insert(0,"../tr")
os.chdir("../tr")
from imgs import componer
os.chdir("../v21")
EN={e["img"]:e for e in json.load(open("../tr/en_imgs.json"))}
FG="/home/user/Wealth-managment-software/secciones/valuacion/graficos/figuras/"
def sinprefijo(t): return re.sub(r"^Figure A?\d+[A-Z]?(?:\.\d)?:\s*","",t)
SPEC={ # id: (n_nuevo, título, fuente, gráfico)
 25:("25",sinprefijo(EN["new_c6.png"]["title"]),EN["new_c6.png"]["src"],FG+"en_V18_caprate_spread.png"),
 26:("26",sinprefijo(EN["new_c3.png"]["title"]),EN["new_c3.png"]["src"],FG+"en_V21_mc_cap_rate.png"),
 27:("27","Mallplaza rose ~4x in five years, ahead of the IPSA; the last leg came with no change in the business",["Weekly closing price, base 100 = 1-Oct-21. Adjusted for the 2024 capital increase. IPSA to 31-Aug-26 (latest available). Milestones: daily closes. Source: Bloomberg; Team analysis."],FG+"en_V19_precio_base100.png"),
 28:("28","Even adding back all deferred taxes, the price is 27% above NAV",["Price / adjusted NAV (EPRA NTA). Adjusted NAV = attributable equity + deferred tax liabilities − goodwill − intangibles. Balance sheet at 30-Jun-26; market cap at 30-Sep-26. Source: Bloomberg; CMF financial statements; Team analysis."],FG+"en_V20_pnav.png"),
 35:("35","With historical volatility, only 5% of 52-week price paths end below our target price",["10,000 paths from CLP 3,555. Resampling of weekly returns (Oct-21 to Sep-26), centered at zero; annual volatility ~32%. Source: Bloomberg; Team analysis."],FG+"en_V22_mc_precio_52s.png"),
 36:("36","Bear, base and bull: even with everything in our favor the target price stays 18% below the price (CLP per share)",["Bear: everything that could go wrong, goes wrong; base: everything as we project it; bull: everything in our favor (Appendix 18). Team illustrations. Source: Team analysis."],FG+"en_V10_escenarios.png"),
 37:("37",sinprefijo(EN["9a4c6a3eeed6c316724a526bd7def097993850a4.png"]["title"]) if False else None,None,None),
}
# Figura 34 original (DCF → precio → consenso): título/fuente del inglés original y su gráfico sin título
e34=EN["b19010530b9f89b139583d8a0b69e20db4e80a4c.png"]
SPEC[37]=("37",sinprefijo(e34["title"]),e34["src"],"../tr/mEN/word/media/b19010530b9f89b139583d8a0b69e20db4e80a4c.png")
from PIL import Image
import io
nuevas={}
for iid,(n,t,s,ch) in SPEC.items():
    W_,H_=Image.open(io.BytesIO(z.read(f"word/media/image{iid}.png"))).size
    os.makedirs("renum",exist_ok=True)
    out=f"renum/image{iid}.png"
    sys.path.insert(0,"../tr"); 
    cwd=os.getcwd(); os.chdir("../tr")
    componer(W_,H_,f"Figure {n}: {t}",s,os.path.join(cwd,ch) if not os.path.isabs(ch) and ch.startswith("renum") else (ch if os.path.isabs(ch) else os.path.join(cwd,ch)),os.path.join(cwd,out))
    os.chdir(cwd); nuevas[f"word/media/image{iid}.png"]=out
o=zipfile.ZipFile(dst,"w",zipfile.ZIP_DEFLATED)
for it in z.infolist():
    if it.filename=="word/document.xml": o.writestr(it,etree.tostring(root,xml_declaration=True,encoding="UTF-8",standalone=True))
    elif it.filename in nuevas: o.writestr(it,open(nuevas[it.filename],"rb").read())
    else: o.writestr(it,z.read(it.filename))
o.close(); print("ok")
