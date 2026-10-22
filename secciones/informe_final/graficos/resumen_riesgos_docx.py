"""Ajustes del Word en inglés v2.1 → 2.2: bloque «What would change our view», cabecera de Key Data, MR1 unido, figuras del cuerpo dentro del margen (ver tabla de cambios del turno). Uso: python3 resumen_riesgos_docx.py entrada.docx salida.docx [config_json]"""
import zipfile,re,copy,sys,json
from lxml import etree
W="{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP="{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A="{http://schemas.openxmlformats.org/drawingml/2006/main}"; R="{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
XML="{http://www.w3.org/XML/1998/namespace}"
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); root=etree.fromstring(z.read("word/document.xml")); body=root.find(W+"body")
rels={r.get("Id"):r.get("Target").split("/")[-1] for r in etree.fromstring(z.read("word/_rels/document.xml.rels"))}
def txt(e): return "".join(t.text or "" for t in e.iter(W+"t"))
kids=list(body)
# --- 1) quitar los 8 párrafos vacíos con formato de título (rellenos de la página 9)
i_sum=next(i for i,c in enumerate(kids) if c.tag==W+"p" and txt(c).strip()=="INVESTMENT SUMMARY")
for i in range(i_sum-1,0,-1):
    c=kids[i]
    if c.tag==W+"p" and not txt(c).strip() and c.find(".//"+W+"drawing") is None and c.find(W+"pPr") is not None and c.find(W+"pPr").find(W+"pBdr") is not None:
        body.remove(c)
    elif c.tag==W+"p" and not txt(c).strip() and c.find(".//"+W+"drawing") is None and i>=i_sum-12:
        continue
    else:
        break
kids=list(body)
i_sum=next(i for i,c in enumerate(kids) if c.tag==W+"p" and txt(c).strip()=="INVESTMENT SUMMARY")
h_sum=kids[i_sum]
# --- 2) bloque nuevo de riesgos
h_ref=next(c for c in body.iter(W+"p") if txt(c).strip()=="REGULATORY & ESG RISKS")
b_ref=next(c for c in body.iter(W+"p") if txt(c).startswith("Mitigants: pro-investment"))
def settext(p,texts):
    rs=[r for r in p.iter(W+"r") if r.find(W+"t") is not None]
    for r,t in zip(rs,texts):
        ts=r.findall(W+"t"); ts[0].text=t; ts[0].set(XML+"space","preserve")
        for x in ts[1:]: r.remove(x)
    for r in rs[len(texts):]: r.getparent().remove(r)
nh=copy.deepcopy(h_ref); settext(nh,["WHAT WOULD CHANGE OUR VIEW"])
nb=copy.deepcopy(b_ref)
settext(nb,["Triggers to watch: ","we would revisit the SELL recommendation if long-term rates fell enough to validate the price. These six indicators tell us when that is happening (Figures 26, 27 and 34)."])
key=next(c for c in kids if c.tag==W+"tbl" and txt(c).startswith("RecommendationSELL") and len(c.findall(W+"tr"))==7)
nt=copy.deepcopy(key)
rows=nt.findall(W+"tr")
datos=[("Price-implied nominal WACC","7.60%","Our 10.41% (−281 bp)","Price needs lower long rates"),
 ("10Y UST yield","5.22%","Back toward 4.98% (model Rf)","Higher rates push cap rates up"),
 ("Real cap rate (LTM EBITDA / EV)","6.0% at the price","Our cash flows: 10.1%","Appraisals (Note 11) narrow the gap"),
 ("Cash tax rate","16.8%","27% long run; 23% from 2027","The 23% cut adds +5.2% to the target"),
 ("Real RONIC of expansion","5.3%","Real WACC 7.2%","Growth adds value only above WACC"),
 ("Bull-case target price","CLP 2,915","Price CLP 3,555 (−18.0%)","No case reaches the price")]
cab=("Indicator","Today","Level that would change our view","Why it matters")
anchos=[3100,1700,3200,3500]
for g,w in zip(nt.find(W+"tblGrid"),anchos): g.set(W+"w",str(w))
def poner(tc,texto,color=None,jc=None,negrita=None):
    ps=tc.findall(W+"p"); p=ps[0]
    for x in ps[1:]: tc.remove(x)
    rs=[r for r in p.iter(W+"r")]
    if not rs:
        r=etree.SubElement(p,W+"r"); rpr=etree.SubElement(r,W+"rPr")
        for tag,val in ((W+"sz","17"),(W+"szCs","17")): e=etree.SubElement(rpr,tag); e.set(W+"val",val)
        etree.SubElement(r,W+"t")
        rs=[r]
    r=rs[0]
    for x in rs[1:]: x.getparent().remove(x)
    rpr=r.find(W+"rPr")
    if rpr is None: rpr=etree.Element(W+"rPr"); r.insert(0,rpr)
    for e in list(rpr):
        if e.tag in (W+"sz",W+"szCs",W+"color",W+"b",W+"bCs"): rpr.remove(e)
    if negrita:
        etree.SubElement(rpr,W+"b"); etree.SubElement(rpr,W+"bCs")
    if color: e=etree.SubElement(rpr,W+"color"); e.set(W+"val",color)
    for tag in (W+"sz",W+"szCs"): e=etree.SubElement(rpr,tag); e.set(W+"val","17")
    t=r.find(W+"t")
    if t is None: t=etree.SubElement(r,W+"t")
    t.text=texto; t.set(XML+"space","preserve")
    ppr=p.find(W+"pPr")
    if ppr is None: ppr=etree.Element(W+"pPr"); p.insert(0,ppr)
    for e in ppr.findall(W+"jc"): ppr.remove(e)
    if jc: e=etree.Element(W+"jc"); e.set(W+"val",jc); ppr.insert(len(ppr)-1 if ppr.find(W+"rPr") is not None else len(ppr),e)
    tc_w=int(tc.find(W+"tcPr").find(W+"tcW").get(W+"w"))
for r_,vals in zip(rows[1:],datos):
    for k,(tc,v) in enumerate(zip(r_.findall(W+"tc"),vals)):
        poner(tc,v,color="14284B" if k==0 else None,jc="center" if k==1 else None,negrita=(k==0))
for tc,v,w in zip(rows[0].findall(W+"tc"),cab,anchos):
    poner(tc,v,color="FFFFFF",jc="center" if v=="Today" else None,negrita=True)
for r_ in rows:
    for tc,w in zip(r_.findall(W+"tc"),anchos): tc.find(W+"tcPr").find(W+"tcW").set(W+"w",str(w))
for r_ in rows:
    trpr=r_.find(W+"trPr")
    if trpr is None: trpr=etree.Element(W+"trPr"); r_.insert(0,trpr)
    if trpr.find(W+"cantSplit") is None: etree.SubElement(trpr,W+"cantSplit")
sp=etree.Element(W+"p"); ppr=etree.SubElement(sp,W+"pPr"); s=etree.SubElement(ppr,W+"spacing"); s.set(W+"after","0"); s.set(W+"line","120"); s.set(W+"lineRule","exact")
for e in (nh,nb,nt,sp): h_sum.addprevious(e)
# --- 3) cabecera de Key Data
for c in kids:
    if c is key:
        r0=c.findall(W+"tr")[0]
        for tc,v in zip(r0.findall(W+"tc"),("Key figure","Value","Key figure","Value")):
            poner(tc,v,color="FFFFFF",jc="center" if v=="Value" else None,negrita=True)
# --- 4) MR1: sangría igual que MR2
mr1=next(c for c in body.iter(W+"p") if txt(c).startswith("MR1:"))
mr2=next(c for c in body.iter(W+"p") if txt(c).startswith("MR2:"))
old=mr1.find(W+"pPr"); new=copy.deepcopy(mr2.find(W+"pPr")); 
if old is not None: mr1.replace(old,new)
# --- 4b) MR1 partido en tres párrafos: unirlos
p2=mr1.getnext(); p3=p2.getnext() if p2 is not None else None
if p2 is not None and txt(p2).startswith("1.0%") and p3 is not None and txt(p3).startswith("likelihood"):
    last=[r for r in mr1.iter(W+"r") if r.find(W+"t") is not None][-1]
    lt=last.find(W+"t"); lt.text=(lt.text or "")+" "; lt.set(XML+"space","preserve")
    for r in list(p2.iter(W+"r")): mr1.append(r)
    lastb=[r for r in mr1.iter(W+"r") if r.find(W+"t") is not None][-1]
    lt=lastb.find(W+"t"); lt.text=(lt.text or "")+" "; lt.set(XML+"space","preserve")
    for r in list(p3.iter(W+"r")): mr1.append(r)
    p2.getparent().remove(p2); p3.getparent().remove(p3)
    print("MR1 unido")
# --- 4c) título de la Figura 29: no partir entre páginas
t29=next(p for p in body.iter(W+"p") if txt(p).startswith("Figure 29: Operational risk tornado"))
ppr=t29.find(W+"pPr")
if ppr is None: ppr=etree.Element(W+"pPr"); t29.insert(0,ppr)
for tag in ("keepNext","keepLines"):
    if ppr.find(W+tag) is None: ppr.insert(0,etree.Element(W+tag))
sp=ppr.find(W+"spacing")
if sp is None:
    sp=etree.Element(W+"spacing"); ppr.append(sp)
sp.set(W+"before","600")
# --- 5) figuras
def dr_by_img(n):
    for dr in body.iter(W+"drawing"):
        nm=rels[next(dr.iter(A+"blip")).get(R+"embed")]
        if nm==f"image{n}.png": return dr
def mover(n,dpt):
    anc=dr_by_img(n).find(WP+"anchor"); o=anc.find(WP+"positionV").find(WP+"posOffset"); o.text=str(int(o.text)+int(round(dpt*12700)))
def escalar(n,f):
    dr=dr_by_img(n)
    for e in dr.iter(WP+"extent",A+"ext"):
        if e.get("cx"): e.set("cx",str(round(int(e.get("cx"))*f))); e.set("cy",str(round(int(e.get("cy"))*f)))
cfg=json.loads(sys.argv[3]) if len(sys.argv)>3 else {}
for n,d in cfg.get("mover",{}).items(): mover(int(n),d)
for n,f in cfg.get("escalar",{}).items(): escalar(int(n),f)
# figuras con alineación «derecha del margen»: posición absoluta (igual que las demás de la columna)
for n in cfg.get("absolutas",[]):
    anc=dr_by_img(n).find(WP+"anchor"); ph=anc.find(WP+"positionH"); cx=int(anc.find(WP+"extent").get("cx"))
    ph.set("relativeFrom","page")
    for c in list(ph): ph.remove(c)
    e=etree.SubElement(ph,WP+"posOffset"); e.text=str(5179060+2254250-cx)
# 22A: re-anclar al título de la página 8 y posicionar respecto de la página
if cfg.get("anclar22A"):
    dr=dr_by_img(27); run=dr.getparent()
    destino=next(c for c in body if c.tag==W+"p" and txt(c).strip()=="OUR VALUE VERSUS THE CONSENSUS")
    run.getparent().remove(run); destino.append(run)
    anc=dr.find(WP+"anchor"); ph=anc.find(WP+"positionH"); pv=anc.find(WP+"positionV")
    ph.set("relativeFrom","page"); 
    for c in list(ph): ph.remove(c)
    e=etree.SubElement(ph,WP+"posOffset"); e.text="5179060"
    pv.set("relativeFrom","page")
    for c in list(pv): pv.remove(c)
    e=etree.SubElement(pv,WP+"posOffset"); e.text=str(int(cfg["anclar22A"]*12700))
out=zipfile.ZipFile(dst,"w",zipfile.ZIP_DEFLATED)
for it in z.infolist():
    out.writestr(it,etree.tostring(root,xml_declaration=True,encoding="UTF-8",standalone=True) if it.filename=="word/document.xml" else z.read(it.filename))
out.close(); print("ok")
