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


def armar(lang):
    env = {**os.environ, "NODE_PATH": "/opt/node22/lib/node_modules:/opt/node-tools/node_modules"}
    subprocess.run(["node", str(DOC / "informe_final.js"), lang], check=True, cwd=RAIZ, env=env, stdout=subprocess.DEVNULL)
    docx = DOC / f"Mallplaza_Informe_CFA_{lang.upper()}.docx"
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


def main(langs, saltos=False):
    for lang in langs:
        archivo = DOC / f"saltos_{lang}.json"
        archivo.write_text("[]")
        for vuelta in range(6):
            pdf = armar(lang)
            ya = json.loads(archivo.read_text())
            # título solo (sin ninguna línea debajo): siempre pasa a la página siguiente; con --saltos, también
            # los que quedan con pocas líneas
            nuevos = [k for k, q in huerfanos(pdf) if k not in ya and (saltos or q < 2)]
            avisos = [k for k, q in huerfanos(pdf) if k not in ya and k not in nuevos]
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
