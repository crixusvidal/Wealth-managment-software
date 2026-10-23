import zipfile,re,sys
from lxml import etree
W="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"; WP="{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A="{http://schemas.openxmlformats.org/drawingml/2006/main}"; R="{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"; XML="{http://www.w3.org/XML/1998/namespace}"
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); root=etree.fromstring(z.read("word/document.xml")); body=root.find(W+"body")
rels={r.get("Id"):r.get("Target").split("/")[-1] for r in etree.fromstring(z.read("word/_rels/document.xml.rels"))}
def txt(e): return "".join(t.text or "" for t in e.iter(W+"t"))
def dr_by(n):
    for dr in body.iter(W+"drawing"):
        if rels.get(next(dr.iter(A+"blip")).get(R+"embed"))==f"image{n}.png": return dr
RIGHT=5179060+2254250
def derecha(n,f=1.0):
    dr=dr_by(n); anc=dr.find(WP+"anchor")
    for e in dr.iter(WP+"extent",A+"ext"):
        if e.get("cx") and f!=1.0: e.set("cx",str(round(int(e.get("cx"))*f))); e.set("cy",str(round(int(e.get("cy"))*f)))
    cx=int(anc.find(WP+"extent").get("cx")); ph=anc.find(WP+"positionH"); ph.set("relativeFrom","page")
    for c in list(ph): ph.remove(c)
    etree.SubElement(ph,WP+"posOffset").text=str(RIGHT-cx)
derecha(21,0.88); derecha(25); derecha(26)
# MR1
for t in root.iter(W+"t"):
    if t.text and "coming entirely expected inflation" in t.text:
        t.text=t.text.replace("coming entirely expected inflation","coming entirely from expected inflation"); t.set(XML+"space","preserve"); print("MR1 corregido")
# filas con figura: no partir entre páginas (anexos)
n=0; anexo=False
for ch in body:
    if ch.tag==W+"p" and re.match(r"APPENDICES",txt(ch).strip()): anexo=True
    if anexo and ch.tag==W+"tbl" and ch.find(".//"+W+"drawing") is not None:
        for tr in ch.findall(W+"tr"):
            trpr=tr.find(W+"trPr")
            if trpr is None: trpr=etree.Element(W+"trPr"); tr.insert(0,trpr)
            if trpr.find(W+"cantSplit") is None: etree.SubElement(trpr,W+"cantSplit"); n+=1
print("filas protegidas",n)
o=zipfile.ZipFile(dst,"w",zipfile.ZIP_DEFLATED)
for it in z.infolist(): o.writestr(it,etree.tostring(root,xml_declaration=True,encoding="UTF-8",standalone=True) if it.filename=="word/document.xml" else z.read(it.filename))
o.close()
