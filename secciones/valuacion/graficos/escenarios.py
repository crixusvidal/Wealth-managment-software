"""BOCETO: escenarios bear / base / bull del precio objetivo, con el oso (amarillo) y el toro (negro) atacando.

Eje horizontal = precio objetivo por acción (CLP). El oso embiste desde el caso bear, el toro desde el bull; ambos
van hacia el precio de cierre (línea roja), que ninguno de los dos alcanza. Cifras de datos_anexos.json (réplica del
DCF v2, Anexo 18), que genera valuacion/graficos/anexos_valuacion.py.

Uso (desde secciones/):
    python valuacion/graficos/escenarios.py [carpeta_salida]
Salida: <carpeta>/{es,en}_V10_escenarios.png (sin título ni fuente: van en el documento).
"""
import json
import sys
from pathlib import Path

import matplotlib.pyplot as plt

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional y rutas compartidas
import estilo  # noqa: E402,F401
from estilo import AMARILLO, AZUL, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA, num
from mascotas import dibujar_oso, dibujar_toro  # noqa: E402

ANCHO, ALTO = 3550 / 1440, 2.75     # pulgadas: ancho de la columna de figuras
DPI = 400
P0, P1 = 0, 3900                     # rango de precios del eje (CLP): las barras parten de cero
X0, X1 = 0.10, ANCHO - 0.10          # extremos del eje en pulgadas
GROSOR = 0.085                       # alto de cada barra

T = {
    "es": {"bear": "BEAR", "base": "BASE", "bull": "BULL", "precio": "Precio {p}", "eje": "Precio objetivo por acción (CLP)"},
    "en": {"bear": "BEAR", "base": "BASE", "bull": "BULL", "precio": "Price {p}", "eje": "Target price per share (CLP)"},
}


def x_de(p):
    return X0 + (p - P0) / (P1 - P0) * (X1 - X0)


def carril(ax, lang, caso, y0, color, po, vs, animal=None, borde=None):
    """Barra desde 0 hasta el precio objetivo; el animal se para sobre el extremo, mirando al precio de cierre."""
    t = T[lang]
    x = x_de(po)
    ax.add_patch(plt.Rectangle((X0, y0), x - X0, GROSOR, facecolor=color, edgecolor=borde or "none", linewidth=0.6, zorder=3))
    ax.scatter([x], [y0 + GROSOR / 2], s=24, color=color, edgecolor="white", linewidth=0.8, zorder=5)
    if animal:
        animal(ax, x, y0 + GROSOR - 0.005)
    ax.text(x + 0.1, y0 + 0.40, t[caso], ha="left", va="bottom", fontsize=6.6, fontweight="bold", color=AZUL if caso != "bull" else NEGRO)
    ax.text(x + 0.1, y0 + 0.215, num(po, lang), ha="left", va="bottom", fontsize=9.0, fontweight="bold", color=NEGRO)
    ax.text(x + 0.1, y0 + 0.075, num(vs * 100, lang, 1) + "%", ha="left", va="bottom", fontsize=6.0, color=TINTA_SECUNDARIA)


def figura(lang, esc, precio, salida):
    t = T[lang]
    fig = plt.figure(figsize=(ANCHO, ALTO))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, ANCHO)
    ax.set_ylim(0, ALTO)
    ax.set_aspect("equal")
    ax.axis("off")
    xp = x_de(precio)
    ax.plot([xp, xp], [0.22, ALTO - 0.26], color=ROJO, linewidth=1.3, linestyle=(0, (4, 2)), zorder=2)
    ax.text(xp, ALTO - 0.04, t["precio"].format(p=num(precio, lang)), ha="center", va="top", fontsize=6.8, fontweight="bold", color=ROJO)
    # el animal queda hacia atrás del extremo: su frente (manos / cuernos) apunta al precio
    oso = lambda a, x, y: dibujar_oso(a, x - 0.72, y, 0.78)
    toro = lambda a, x, y: dibujar_toro(a, x - 0.98, y, 0.62)
    carril(ax, lang, "bear", 1.80, AMARILLO, esc["bear"]["po"], esc["bear"]["vs"], oso, borde=AZUL)
    carril(ax, lang, "base", 1.10, AZUL, esc["base"]["po"], esc["base"]["vs"])
    carril(ax, lang, "bull", 0.30, NEGRO, esc["bull"]["po"], esc["bull"]["vs"], toro)
    ax.text(ANCHO / 2, 0.02, t["eje"], ha="center", va="bottom", fontsize=5.6, color=TINTA_SECUNDARIA)
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)
    return salida


if __name__ == "__main__":
    salida = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "figuras"
    d = json.loads((Path(__file__).parent / "figuras" / "datos_anexos.json").read_text())["F"]
    for lang in ("es", "en"):
        print(figura(lang, d["esc"], d["precio"], salida / f"{lang}_V10_escenarios.png"))
