"""Cuadra las figuras flotantes (columna derecha) de un Word ya armado: mide en el PDF dónde cae cada una y
corrige tblpY para que no se pisen ni queden dentro del margen inferior. No toca texto ni cifras.

Uso: python3 cuadrar_docx.py entrada.docx salida.docx [--lista] [--escala=0.8]
--escala reduce las figuras de Valuación y pasa a la derecha las de la grilla de cierre.
"""
import re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
SOFFICE = "/root/.claude/skills/synced/508cc6bc-ad76-422f-ada9-ad0fc4279625_0b2a2ef9-bf53-46a5-b8ed-72255cce463a/docx/scripts/office/soffice.py"
ARRIBA, ABAJO, SEP, EXTRA, XCOL = 10.0, 841.89 - 21.0, 5.0, 64.0, 395.0   # pt; EXTRA = título + fuente de la figura


def texto(e):
    return "".join(t.text or "" for t in e.iter(W + "t"))


def flotantes(body):
    """(índice, tabla, etiqueta, alto de imagen en pt) de las tablas flotantes de una columna con figura."""
    out = []
    for i, ch in enumerate(body):
        if ch.tag != W + "tbl":
            continue
        pr = ch.find(W + "tblPr").find(W + "tblpPr")
        imgs = list(ch.iter(WP + "inline"))
        if pr is None or not imgs:
            continue
        cap = next((texto(p) for p in ch.iter(W + "p") if re.match(r"(Figure|Figura) \d+[A-Z]?:", texto(p))), None)
        if cap is None:
            continue
        et = re.match(r"(?:Figure|Figura) (\d+[A-Z]?):", cap).group(1)
        alto = int(imgs[0].find(WP + "extent").get("cy")) / 12700
        out.append((i, ch, et, alto))
    return out


def a_pdf(docx, carpeta):
    subprocess.run([sys.executable, SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", str(carpeta), str(docx)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=240)
    return Path(carpeta) / (Path(docx).stem + ".pdf")


def posiciones(pdf):
    """etiqueta → (página, y) del título de cada figura de la columna derecha."""
    html = subprocess.run(["pdftotext", "-bbox", str(pdf), "-"], capture_output=True, text=True).stdout
    pos = {}
    for p, pag in enumerate(html.split("<page ")[1:]):
        ws = re.findall(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="[\d.]+">([^<]*)</word>', pag)
        for k in range(len(ws) - 1):
            if ws[k][2] in ("Figure", "Figura") and float(ws[k][0]) >= XCOL - 5:
                m = re.fullmatch(r"(\d+[A-Z]?):", ws[k + 1][2])
                if m and m.group(1) not in pos:
                    pos[m.group(1)] = (p, float(ws[k][1]))
    return pos


A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
ANCHO_COL = 3550 * 635                                         # EMU: ancho de la columna derecha


def seccion(body, ini, fin):
    """Índices de body entre el título que calza con ini y el que calza con fin."""
    i0 = next(i for i, e in enumerate(body) if e.tag == W + "p" and re.match(ini, texto(e).strip()))
    i1 = next(i for i, e in enumerate(body) if i > i0 and e.tag == W + "p" and re.match(fin, texto(e).strip()))
    return i0, i1


def escalar(tbl, f):
    for e in tbl.iter(WP + "extent", A_NS + "ext"):
        if e.get("cx"):
            e.set("cx", str(round(int(e.get("cx")) * f)))
            e.set("cy", str(round(int(e.get("cy")) * f)))


def preparar(root, escala, ini, fin):
    """Reduce las figuras de la sección y pasa a la columna derecha las de la grilla de cierre (junto a su cita)."""
    import copy
    body = root.find(W + "body")
    i0, i1 = seccion(body, ini, fin)
    sec = list(body)[i0:i1]
    plantilla = next(t for t in sec if t.tag == W + "tbl" and t.find(W + "tblPr").find(W + "tblpPr") is not None
                     and list(t.iter(WP + "inline")))
    for t in sec:
        if t.tag == W + "tbl" and t.find(W + "tblPr").find(W + "tblpPr") is not None and list(t.iter(WP + "inline")):
            escalar(t, escala)
    for t in sec:
        grid = [int(g.get(W + "w")) for g in t.find(W + "tblGrid")] if t.tag == W + "tbl" else []
        if len(grid) != 3 or t.find(W + "tblPr").find(W + "tblpPr") is not None:
            continue
        for celda in t.iter(W + "tc"):
            if not list(celda.iter(WP + "inline")):
                continue
            ps = [p for p in celda if p.tag == W + "p"]
            cap = next((texto(p) for p in ps if re.match(r"(Figure|Figura) \d+[A-Z]?:", texto(p))), None)
            if not cap:
                continue
            n = re.match(r"(?:Figure|Figura) (\d+[A-Z]?):", cap).group(1)
            nueva = copy.deepcopy(plantilla)
            tc = nueva.find(W + "tr").find(W + "tc")
            for hijo in list(tc):
                if hijo.tag != W + "tcPr":
                    tc.remove(hijo)
            for p in ps:
                tc.append(copy.deepcopy(p))
            cx = max(int(e.get("cx")) for e in nueva.iter(WP + "extent"))
            if cx > ANCHO_COL:
                escalar(nueva, ANCHO_COL / cx)
            nueva.find(W + "tblPr").find(W + "tblpPr").set(W + "tblpY", "0")
            cita = re.compile(r"\((?:Figure|Figura)s? [^)]*\b" + re.escape(n) + r"\b")
            lugar = next((e for e in sec if e.tag == W + "p" and cita.search(texto(e))), None)
            lugar = lugar if lugar is not None else next(e for e in reversed(sec) if e.tag == W + "p" and len(texto(e)) > 200)
            lugar.addnext(nueva)
            for p in ps:
                celda.remove(p)
            celda.append(etree.SubElement(celda, W + "p"))
        if not any(texto(c).strip() or list(c.iter(WP + "inline")) for c in t.iter(W + "tc")):
            t.getparent().remove(t)


def cuadrar(entrada, salida, vueltas=8, lista=False, escala=None):
    tmp = Path(tempfile.mkdtemp())
    work = tmp / "doc.docx"
    shutil.copy(entrada, work)
    if escala:
        z = zipfile.ZipFile(work)
        root = etree.fromstring(z.read("word/document.xml"))
        preparar(root, escala, r"(VALUACIÓN|VALUATION)$", r"(RIESGOS DE INVERSIÓN|INVESTMENT RISKS)")
        out = tmp / "prep.docx"
        with zipfile.ZipFile(work) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                zout.writestr(it, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                              if it.filename == "word/document.xml" else zin.read(it.filename))
        shutil.move(out, work)
    for v in range(vueltas):
        z = zipfile.ZipFile(work)
        root = etree.fromstring(z.read("word/document.xml"))
        fl = flotantes(root.find(W + "body"))
        try:
            pos = posiciones(a_pdf(work, tmp))
        except subprocess.TimeoutExpired:                      # LibreOffice se cuelga con esta disposición: vuelve a la anterior
            print("conversión colgada: se conserva la última disposición convertible", flush=True)
            subprocess.run(["pkill", "-x", "soffice.bin"])
            if (tmp / "previo.docx").exists():
                shutil.copy(tmp / "previo.docx", work)
            break
        shutil.copy(work, tmp / "previo.docx")
        orden = sorted((f for f in fl if f[2] in pos), key=lambda f: pos[f[2]])
        cambios, fondo = 0, {}
        for i, tbl, et, alto in orden:
            p, y = pos[et]
            h = alto + EXTRA
            nuevo_y = max(y, fondo.get(p, ARRIBA) + (SEP if p in fondo else 0))
            pp = p
            if nuevo_y + h > ABAJO:                              # no cabe: pasa arriba de la página siguiente
                pp, nuevo_y = p + 1, max(ARRIBA, fondo.get(p + 1, ARRIBA) + (SEP if p + 1 in fondo else 0))
            fondo[pp] = nuevo_y + h
            delta = (nuevo_y - y) + (pp - p) * (ABAJO - ARRIBA)
            if abs(delta) > 1.5:
                pr = tbl.find(W + "tblPr").find(W + "tblpPr")
                pr.set(W + "tblpY", str(int(pr.get(W + "tblpY")) + round(delta * 20)))
                cambios += 1
        print(f"vuelta {v}: {len(fl)} figuras, {cambios} movidas", flush=True)
        if lista:
            for et, (p, y) in sorted(pos.items(), key=lambda kv: kv[1]):
                print("  ", et, p + 1, round(y))
        if not cambios:
            break
        out = tmp / "nuevo.docx"
        with zipfile.ZipFile(work) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for it in zin.infolist():
                zout.writestr(it, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                              if it.filename == "word/document.xml" else zin.read(it.filename))
        shutil.move(out, work)
    shutil.copy(work, salida)
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    esc = next((float(a.split("=")[1]) for a in sys.argv if a.startswith("--escala=")), None)
    cuadrar(sys.argv[1], sys.argv[2], lista="--lista" in sys.argv, escala=esc)
