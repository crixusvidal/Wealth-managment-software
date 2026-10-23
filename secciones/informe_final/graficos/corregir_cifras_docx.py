import zipfile,re,json,sys,os,io
from lxml import etree
from PIL import Image
W="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"; XML="{http://www.w3.org/XML/1998/namespace}"
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); root=etree.fromstring(z.read("word/document.xml"))
def reemplazar(p,old,new):
    ts=list(p.iter(W+"t")); full="".join(t.text or "" for t in ts); i=full.find(old)
    if i<0: return False
    pos=0; fin=i+len(old); primero=True
    for t in ts:
        s=t.text or ""; a,b=pos,pos+len(s); pos=b
        if b<=i or a>=fin: continue
        lo=max(i,a)-a; hi=min(fin,b)-a
        t.text=s[:lo]+(new if primero else "")+s[hi:]; t.set(XML+"space","preserve"); primero=False
    return True
R=[("median is ~CLP 2,300, 35% below the price","median is ~CLP 2,270, 36% below the price"),
   ("the price is 27% above NAV (1.28x","the price is 28% above NAV (1.28x"),
   ("the stock rose 39% in the 12 months","the stock rose 37% in the 12 months")]
for old,new in R:
    n=sum(reemplazar(p,old,new) for p in root.iter(W+"p")); print(n,"×",new[:50])
# imágenes con el título dentro: Figura 27 (media 26) y Figura 29 (media 28)
T="/tmp/claude-0/-home-user-Wealth-managment-software/4853c681-1399-5e51-b2e7-c4edae34dd33/scratchpad/tr/"
FG="/home/user/Wealth-managment-software/secciones/valuacion/graficos/figuras/"
HERE=os.getcwd(); os.chdir(T); sys.path.insert(0,"."); from imgs import componer; os.chdir(HERE)
EN=json.load(open(T+"en_imgs.json")); e21B=next(e for e in EN if e["title"] and e["title"].startswith("Figure 21B"))
tit27=re.sub(r"^Figure \d+[A-Z]?:\s*","",e21B["title"]).replace("~CLP 2,300","~CLP 2,270")
print("título Fig 27:",tit27)
dims={k:Image.open(io.BytesIO(z.read(f"word/media/image{k}.png"))).size for k in (26,28)}
os.makedirs("corr",exist_ok=True); os.chdir(T)
componer(*dims[26],f"Figure 27: {tit27}",e21B["src"],FG+"en_V21_mc_cap_rate.png",HERE+"/corr/image26.png")
componer(*dims[28],"Figure 29: Even adding back all deferred taxes, the price is 28% above NAV",["Price / adjusted NAV (EPRA NTA). Adjusted NAV = attributable equity + deferred tax liabilities − goodwill − intangibles. Balance sheet at 30-Jun-26; market cap at 30-Sep-26. Source: Bloomberg; CMF financial statements; Team analysis."],FG+"en_V20_pnav.png",HERE+"/corr/image28.png")
os.chdir(HERE)
o=zipfile.ZipFile(dst,"w",zipfile.ZIP_DEFLATED)
for it in z.infolist():
    if it.filename=="word/document.xml": o.writestr(it,etree.tostring(root,xml_declaration=True,encoding="UTF-8",standalone=True))
    elif it.filename in ("word/media/image26.png","word/media/image28.png"): o.writestr(it,open("corr/"+it.filename.split("/")[-1],"rb").read())
    else: o.writestr(it,z.read(it.filename))
o.close()
