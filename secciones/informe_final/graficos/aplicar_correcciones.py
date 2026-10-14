"""Aplica las correcciones de datos del auditor (JSON) en las fuentes de texto de cada sección.

Solo reemplaza subcadenas exactas (una vez): en los borradores, dentro de la corrida de texto del JSON;
en los constructores .js, dentro del archivo. Lo que no encuentre o encuentre repetido lo informa y no
lo toca. La sección de Riesgos queda fuera por decisión del equipo.

Uso (desde secciones/):
    python informe_final/graficos/aplicar_correcciones.py <correcciones.json> [--aplicar]
"""
import json
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[2]
BORRADORES = {s: RAIZ / s / "documento" / "borrador" / "texto.json" for s in ("negocio", "industria", "esg")}
JS = {"analisis_financiero": RAIZ / "analisis_financiero/documento/analisis_financiero.js",
      "valuacion": RAIZ / "valuacion/documento/valuacion.js",
      "macro": RAIZ / "macro/anexo/anexo_macro.js"}


def corridas(nodo):
    """Todas las listas [texto, negrita, cursiva] y textos sueltos del JSON del borrador."""
    if isinstance(nodo, dict):
        for k, v in nodo.items():
            if k == "runs":
                yield from v
            elif k == "texto" and isinstance(v, str):
                yield nodo
            else:
                yield from corridas(v)
    elif isinstance(nodo, list):
        for x in nodo:
            yield from corridas(x)


def main(ruta, aplicar):
    datos = json.loads(Path(ruta).read_text(encoding="utf-8"))
    hechos, fallidos = [], []
    jsons = {s: json.loads(p.read_text(encoding="utf-8")) for s, p in BORRADORES.items()}
    textos = {s: p.read_text(encoding="utf-8") for s, p in JS.items()}
    for c in datos["correcciones"]:
        s, viejo, nuevo = c["seccion"], c["viejo"], c["nuevo"]
        if s in jsons:
            lugares = []
            for r in corridas(jsons[s]):
                t = r["texto"] if isinstance(r, dict) else r[0]
                if viejo in t:
                    lugares.append(r)
            ocurr = sum((r["texto"] if isinstance(r, dict) else r[0]).count(viejo) for r in lugares)
            if ocurr != 1:
                fallidos.append((s, viejo, f"{ocurr} coincidencias en el borrador (¿cruza negritas?)"))
                continue
            r = lugares[0]
            if isinstance(r, dict):
                r["texto"] = r["texto"].replace(viejo, nuevo)
            else:
                r[0] = r[0].replace(viejo, nuevo)
        elif s in textos:
            n = textos[s].count(viejo)
            if n != 1:
                fallidos.append((s, viejo, f"{n} coincidencias en el .js"))
                continue
            textos[s] = textos[s].replace(viejo, nuevo)
        else:
            fallidos.append((s, viejo, "sección no editable (Riesgos queda para después)"))
            continue
        hechos.append((s, viejo, nuevo, c.get("fuente", "")))
    if aplicar:
        for s, p in BORRADORES.items():
            p.write_text(json.dumps(jsons[s], ensure_ascii=False, indent=1), encoding="utf-8")
        for s, p in JS.items():
            p.write_text(textos[s], encoding="utf-8")
    print(f"{len(hechos)} aplicadas{'' if aplicar else ' (simulación)'}, {len(fallidos)} sin aplicar")
    for h in hechos:
        print("  OK ", h[0], "|", h[1], "→", h[2], "|", h[3])
    for f in fallidos:
        print("  NO ", f[0], "|", f[1], "|", f[2])


if __name__ == "__main__":
    main(sys.argv[1], "--aplicar" in sys.argv)
