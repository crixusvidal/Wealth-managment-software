"""Arma el informe consolidado (ES y EN) en Word y PDF, con las secciones corridas una tras otra
(cada una empieza a mitad de página, justo después de la anterior, como los anexos).

Revisa el PDF: si el título de una sección quedó solo al pie de una página, esa sección pasa a empezar en
la página siguiente (informe_final/documento/saltos_<lang>.json). Con --saltos, también las que quedan con
menos de 4 líneas debajo del título.

Uso (desde secciones/):
    python informe_final/graficos/maquetar.py [es] [en]
"""
import json
import os
import re
import subprocess
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
DOC = RAIZ / "informe_final" / "documento"
MAQ = DOC / "maquetacion"                  # archivos intermedios (plan, posiciones, escalas, ...)
SOFFICE = Path("/root/.claude/skills/synced/508cc6bc-ad76-422f-ada9-ad0fc4279625_0b2a2ef9-bf53-46a5-b8ed-72255cce463a/docx/scripts/office/soffice.py")
TITULOS = {
    "industria": r"(ANÁLISIS DE LA INDUSTRIA|INDUSTRY OVERVIEW)", "esg": r"(AMBIENTAL, SOCIAL|ENVIRONMENTAL, SOCIAL)",
    "fa": r"(ANÁLISIS FINANCIERO|FINANCIAL ANALYSIS)$", "valuacion": r"(VALUACIÓN|VALUATION)$",
    "riesgos": r"(RIESGOS DE INVERSIÓN|INVESTMENT RISKS)", "resumen": r"(RESUMEN DE INVERSIÓN|INVESTMENT SUMMARY)",
    "a_industria": r"(ANEXOS|APPENDICES): (INDUSTRIA|INDUSTRY)", "a_esg": r"(ANEXO|APPENDIX) 8:",
    "a_fa": r"(ANEXOS|APPENDICES): (ANÁLISIS FINANCIERO|FINANCIAL ANALYSIS)", "a_valuacion": r"(ANEXO|APPENDIX) 15:",
    "a_macro": r"(ANEXO|APPENDIX) 20:",
}
PIE = re.compile(r"^(Página|Page|A-)\s*\d+$|^A-\d+$")


def armar(lang, medir=False):
    env = {**os.environ, "NODE_PATH": "/opt/node22/lib/node_modules:/opt/node-tools/node_modules"}
    subprocess.run(["node", str(DOC / "informe_final.js"), lang] + (["--medir"] if medir else []), check=True, cwd=RAIZ, env=env,
                   stdout=subprocess.DEVNULL)
    docx = DOC / f"Mallplaza_Informe_CFA_{lang.upper()}{'_medir' if medir else ''}.docx"
    subprocess.run([sys.executable, str(SOFFICE), "--headless", "--convert-to", "pdf", "--outdir", str(DOC), str(docx)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return docx.with_suffix(".pdf")


def huerfanos(pdf):
    texto = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], check=True, capture_output=True, text=True).stdout
    out = []
    for pagina in texto.split("\f"):
        lineas = [l.strip() for l in pagina.splitlines() if l.strip()]
        lineas = [l for l in lineas if not PIE.match(l) and "CFA Institute Research Challenge" not in l]
        for i, l in enumerate(lineas):
            for clave, rx in TITULOS.items():
                quedan = len(lineas) - i - 1
                if re.match(rx, l) and quedan < 4:
                    out.append((clave, quedan))
    return out


# Posiciones de las figuras flotantes ------------------------------------------------------------
ALTO_PAG, ARRIBA, ABAJO = 841.89, 50.0, 841.89 - 45.0     # A4 en puntos; márgenes 1000 / 900 twips
SEP = 4.0                                                 # separación vertical entre figuras (pt)


def tokens(texto):
    out = []
    for t in re.split(r"\s+", texto.lower()):
        t = re.sub(r"[^\wáéíóúñü]", "", t)
        if t and not t.isdigit() and not re.fullmatch(r"[\d.,]+", t):
            out.append(t)
    return out


def palabras(pdf):
    html = subprocess.run(["pdftotext", "-bbox-layout", str(pdf), "-"], check=True, capture_output=True, text=True).stdout
    out = []
    for p, pagina in enumerate(html.split("<page ")[1:]):
        for m in re.finditer(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="[\d.]+" yMax="([\d.]+)">([^<]*)</word>', pagina):
            for t in tokens(m.group(4).replace("&amp;", "&")):
                out.append((p, float(m.group(2)), t))
    return out


def ubicar(plan, pdf):
    """Página e y (pt) del inicio de cada elemento, buscando sus primeras palabras en el PDF en orden."""
    ws, k, pos = palabras(pdf), 0, []
    seq = [w[2] for w in ws]
    for e in plan["elementos"]:
        clave = tokens(e["inicio"])[:5]
        hallado = None
        for j in range(k, min(len(seq), k + 4000)):
            if seq[j:j + len(clave)] == clave:
                hallado = j
                break
        if hallado is None:
            pos.append(None)
            continue
        k = hallado + len(clave)
        pos.append((ws[hallado][0], ws[hallado][1]))
    for i in range(len(pos)):                     # sin ubicar: hereda la posición del anterior
        if pos[i] is None:
            pos[i] = pos[i - 1] if i else (0, ARRIBA)
    return pos


def posicionar(plan, pdf, adelantar=frozenset()):
    """Ubica cada figura junto al párrafo que la cita. Devuelve (posiciones, espacios): si las figuras de una
    sección no terminan antes del siguiente bloque de ancho completo, ese bloque necesita bajar (espacios)."""
    pos = ubicar(plan, pdf)
    elems = plan["elementos"]
    ultima = max(p for p, _ in pos)
    anchos = [i for i, e in enumerate(elems) if e["tipo"] == "ancho"]
    barreras = [i for i, e in enumerate(elems) if e.get("barrera")]

    def rango(i):
        (p0, y0) = pos[i]
        (p1, y1) = pos[i + 1] if i + 1 < len(pos) else (ultima, ABAJO)
        return (p0, y0), (p1, y1)

    ocupado, salida, espacios = {}, {}, {}
    for fig in sorted(plan["figuras"], key=lambda f: (f["cita"], f["id"])):
        barrera = next((i for i in barreras if i > fig["cita"]), None)
        obst = {}
        for i in anchos:
            if i >= (barrera if barrera is not None else 10**9):
                continue
            (p0, y0), (p1, y1) = rango(i)
            for p in range(p0, p1 + 1):
                obst.setdefault(p, []).append((y0 if p == p0 else ARRIBA, y1 if p == p1 else ABAJO))
        p, y = pos[fig["cita"]]
        if fig.get("clave") in adelantar and p > 0:     # último recurso: puede subir hasta una página antes de su cita
            p, y = p - 1, y
        h = fig["alto"] * 72
        for _ in range(300):
            y = max(y, ARRIBA)
            choques = [r for r in ocupado.get(p, []) + obst.get(p, []) if r[0] < y + h + SEP and y < r[1] + SEP]
            if choques:
                y = max(r[1] for r in choques) + SEP
                continue
            if y + h > ABAJO:
                p, y = p + 1, ARRIBA
                continue
            en_pag = [(pos[i][1], i) for i, e in enumerate(elems) if e["tipo"] == "texto" and pos[i][0] == p
                      and (barrera is None or i < barrera)]
            if not en_pag:       # ningún párrafo de la sección empieza en esta página: el ancla es el último anterior
                previos = [i for i, e in enumerate(elems) if e["tipo"] == "texto" and i <= fig["cita"] or
                           (e["tipo"] == "texto" and (barrera is None or i < barrera) and pos[i][0] < p)]
                ia = max(previos)
                ya = pos[ia][1] - (pos[ia][0] - p) * (ABAJO - ARRIBA)
            else:
                arriba = [x for x in en_pag if x[0] <= y + 1]
                ya, ia = max(arriba) if arriba else min(en_pag)
            salida[fig["id"]] = {"elemento": ia, "offset": round((y - ya) * 20), "pagina": p + 1, "y": round(y, 1)}
            ocupado.setdefault(p, []).append((y, y + h))
            # ¿termina después de que empiece el bloque de ancho completo siguiente? → va a la grilla de cierre
            if barrera is not None and (p, y + h) > pos[barrera]:
                del salida[fig["id"]]
                ocupado[p].pop()
                espacios.setdefault(barrera, []).append(fig["id"])
            break
    return salida, espacios


def main(langs, saltos=False):
    for lang in langs:
        (MAQ / f"posiciones_{lang}.json").unlink(missing_ok=True)
        # Gráficos al costado del texto. Si alguno no cabe antes del siguiente bloque de ancho completo, se achica
        # el alto de los gráficos de ese tramo (hasta 70%) y se vuelve a medir. Lo que igual no quepa (idealmente
        # ninguno, a lo más 2) va en una fila de dos al cierre de la sección.
        archivo_esc = MAQ / f"escalas_{lang}.json"
        escalas = {}
        archivo_esc.write_text("{}")
        for vuelta in range(14):
            pdf_m = armar(lang, medir=True)
            plan = json.loads((MAQ / f"plan_{lang}.json").read_text())
            pos, finales = posicionar(plan, pdf_m)
            fuera = [i for v in finales.values() for i in v]
            print(lang, "vuelta", vuelta, "escalas:", escalas, "abajo:", len(fuera), flush=True)
            if not fuera:
                break
            figs = {f["id"]: f for f in plan["figuras"]}
            secc = {figs[i]["clave"] for i in fuera}
            pag0 = min((v["pagina"] for k, v in pos.items() if figs[int(k)]["clave"] in secc), default=0)
            claves = secc | {figs[int(k)]["clave"] for k, v in pos.items() if v["pagina"] >= pag0}
            nuevas = {c: round(max(0.65, escalas.get(c, 1.0) * 0.93), 3) for c in claves}
            if all(nuevas[c] >= escalas.get(c, 1.0) for c in claves):
                if len(fuera) <= 1:
                    break                               # ya en el mínimo legible
                # aún quedan 2 o más abajo: también se achican los de la página anterior del tramo
                claves = claves | {figs[int(k)]["clave"] for k, v in pos.items() if v["pagina"] >= pag0 - 1}
                nuevas = {c: round(max(0.65, escalas.get(c, 1.0) * 0.93), 3) for c in claves}
                if all(nuevas[c] >= escalas.get(c, 1.0) for c in claves):
                    break
            escalas.update(nuevas)
            archivo_esc.write_text(json.dumps(escalas))
        fuera = [i for v in finales.values() for i in v]
        if fuera:
            figs = {f["id"]: f for f in plan["figuras"]}
            pdf_m = armar(lang, medir=True)
            adel = frozenset(figs[i]["clave"] for i in fuera) | frozenset(escalas)
            pos2, fin2 = posicionar(plan, pdf_m, adel)
            n2 = sum(len(v) for v in fin2.values())
            print(lang, "adelantando", sorted(adel), "abajo:", n2)
            if n2 < len(fuera):
                pos, finales = pos2, fin2
        (MAQ / f"finales_{lang}.json").write_text(json.dumps(finales, indent=1))
        print(lang, "figuras en grilla al cierre de sección:", {k: [plan["figuras"][i]["n"] for i in v] for k, v in finales.items()})
        (MAQ / f"posiciones_{lang}.json").write_text(json.dumps(pos, indent=1))
        for x in (pdf_m, pdf_m.with_suffix(".docx")):
            x.unlink(missing_ok=True)
        sin = [f["n"] for f in plan["figuras"] if str(f["id"]) not in {str(k) for k in pos}]
        if sin:
            print(lang, "figuras sin ubicar:", sin)
        archivo = MAQ / f"saltos_{lang}.json"
        archivo.write_text("[]")
        for vuelta in range(6):
            pdf = armar(lang)
            ya = json.loads(archivo.read_text())
            # título solo (sin ninguna línea debajo): siempre pasa a la página siguiente; con --saltos, también
            # los que quedan con pocas líneas
            nuevos = [k for k, q in huerfanos(pdf) if k.startswith("a_") and k not in ya and (saltos or q < 2)]
            avisos = [k for k, q in huerfanos(pdf) if k not in ya and k not in nuevos]   # el cuerpo ya no usa saltos
            if avisos:
                print(lang, "títulos al pie de página con poco texto debajo (sin salto):", avisos)
            if not nuevos:
                break
            archivo.write_text(json.dumps(json.loads(archivo.read_text()) + nuevos[:1]))   # de a uno: cada salto mueve lo que sigue
        paginas = subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True).stdout
        print(lang, "saltos:", json.loads(archivo.read_text()), re.search(r"Pages:\s+(\d+)", paginas).group(0), pdf)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--saltos"]
    main(args or ["es", "en"], "--saltos" in sys.argv)
