"""Arma el PDF final del informe consolidado con Typst (cuerpo) + anexos (Word → PDF), en español e inglés.

1. exportar.js deja el cuerpo en JSON (mismos textos y numeración que el Word).
2. Typst compila informe.typ: texto corrido y gráficos al costado de su cita. Si algún gráfico no cabe antes del
   Resumen de inversión, se achica el alto de los gráficos de ese tramo (hasta 65%) y, como último recurso,
   pueden subir una página antes de su cita; lo que igual no quepa va en una fila al final.
3. Los anexos salen del Word (informe_final.js --solo=anexos, LibreOffice) y se unen al cuerpo.

Uso (desde secciones/): python informe_final/typst/armar.py [es] [en]
Salida: informe_final/documento/Mallplaza_Informe_CFA_<LANG>.pdf
"""
import json
import os
import subprocess
import sys
from pathlib import Path

import typst
from pypdf import PdfWriter

RAIZ = Path(__file__).resolve().parents[2]
TY = RAIZ / "informe_final" / "typst"
DOC = RAIZ / "informe_final" / "documento"
ENV = {**os.environ, "NODE_PATH": "/opt/node22/lib/node_modules:/opt/node-tools/node_modules"}
SOFFICE = "/root/.claude/skills/synced/508cc6bc-ad76-422f-ada9-ad0fc4279625_0b2a2ef9-bf53-46a5-b8ed-72255cce463a/docx/scripts/office/soffice.py"


def exportar(lang):
    subprocess.run(["node", str(TY / "exportar.js"), lang], check=True, cwd=RAIZ, env=ENV, stdout=subprocess.DEVNULL)
    d = json.loads((TY / "datos" / f"cuerpo_{lang}.json").read_text())
    for s in d["secciones"]:
        for f in s["figuras"]:
            f["archivo"] = "/" + os.path.relpath(f["archivo"], RAIZ)
        for b in s["bloques"]:
            if isinstance(b, dict) and "figancha" in b:
                b["figancha"]["archivo"] = "/" + os.path.relpath(b["figancha"]["archivo"], RAIZ)
    (TY / "datos" / f"cuerpo_{lang}.json").write_text(json.dumps(d, ensure_ascii=False, indent=1))
    return f"/informe_final/typst/datos/cuerpo_{lang}.json"


def compilar(datos, escalas, adelantar, salida=None):
    ins = {"datos": datos, "escalas": json.dumps(escalas), "adelantar": ",".join(sorted(adelantar))}
    if salida:
        typst.compile(str(TY / "informe.typ"), output=str(salida), root=str(RAIZ), sys_inputs=ins)
    r = json.loads(typst.query(str(TY / "informe.typ"), "<resultado>", field="value", root=str(RAIZ), sys_inputs=ins))[0]
    return r


FONDO = 0.55   # escala mínima de los gráficos (legibilidad)


def ajustar(datos):
    escalas, adelantar = {}, set()
    r = compilar(datos, escalas, adelantar)
    for _ in range(14):
        if not r["fuera"]:
            break
        claves = r["claves"]
        tramo = {claves[i] for i in r["ids"]}
        pags = [p for k, (p, _) in r["asig"].items() if claves[int(k)] in tramo]
        pag0 = min(pags) if pags else 0
        # solo se achican la sección que no cupo y las anteriores (liberan su espacio); las posteriores no se tocan
        orden = list(dict.fromkeys(claves))
        tramo |= set(orden[: max(orden.index(c) for c in tramo) + 1])
        # el tramo incluye desde la página 0 si ya no queda nada más por achicar (el mapa de la Figura 1 no se achica)
        if all(escalas.get(c, 1.0) <= FONDO + 1e-6 for c in tramo):
            tramo |= set(claves)
        nuevas = {c: round(max(FONDO, escalas.get(c, 1.0) * 0.93), 3) for c in tramo}
        if all(nuevas[c] >= escalas.get(c, 1.0) for c in tramo):
            break                      # ya no hay cómo achicar: lo que no cupo va al final
        else:
            escalas.update(nuevas)
        r = compilar(datos, escalas, adelantar)
        print("   escalas", escalas, "adelantar", sorted(adelantar), "abajo:", r["fuera"], flush=True)
    return escalas, adelantar, r


def anexos(lang):
    subprocess.run(["node", str(DOC / "informe_final.js"), lang, "--solo=anexos"], check=True, cwd=RAIZ, env=ENV, stdout=subprocess.DEVNULL)
    docx = DOC / f"Mallplaza_Informe_CFA_{lang.upper()}_anexos.docx"
    subprocess.run([sys.executable, SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", str(TY), str(docx)],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    docx.unlink()
    return TY / docx.with_suffix(".pdf").name


def main(langs):
    for lang in langs:
        datos = exportar(lang)
        escalas, adelantar, r = ajustar(datos)
        cuerpo = TY / f"cuerpo_{lang}.pdf"
        compilar(datos, escalas, adelantar, cuerpo)
        pdf_anexos = anexos(lang)
        w = PdfWriter()
        for p in (cuerpo, pdf_anexos):
            w.append(str(p))
        final = DOC / f"Mallplaza_Informe_CFA_{lang.upper()}.pdf"
        with open(final, "wb") as fh:
            w.write(fh)
        for p in (cuerpo, pdf_anexos):
            p.unlink()
        abajo = [r["claves"][i] for i in r["ids"]]
        print(lang, "listo:", final.name, "| gráficos abajo:", r["fuera"], abajo, "| escalas:", escalas, "| adelantar:", sorted(adelantar))


if __name__ == "__main__":
    main(sys.argv[1:] or ["es", "en"])
