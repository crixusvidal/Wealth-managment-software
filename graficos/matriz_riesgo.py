"""Matriz de riesgo (likelihood x impact) con colores corporativos Mallplaza.

Lee los riesgos de un CSV con columnas id, risk, likelihood, impact (escala 1-5).

Uso:
    python graficos/matriz_riesgo.py [riesgos.csv] [salida.png]
"""
import csv
import sys
from collections import defaultdict
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.colors import to_rgb
from matplotlib.patches import Circle, Rectangle

from sensibilidad_wacc import AMARILLO, CELESTE, NEGRO, ROJO, TINTA_SECUNDARIA

ETIQUETAS = ["Very low", "Low", "Medium", "High", "Very high"]
# Nivel por puntaje likelihood x impact
NIVELES = [(4, "Low", CELESTE), (9, "Medium", AMARILLO), (25, "High", ROJO)]


def tinte(color, fuerza):
    return tuple(1 - fuerza * (1 - c) for c in to_rgb(color))


def nivel(puntaje):
    return next(n for n in NIVELES if puntaje <= n[0])


def leer_riesgos(ruta):
    with open(ruta, newline="", encoding="utf-8") as f:
        return [dict(r, likelihood=int(r["likelihood"]), impact=int(r["impact"]))
                for r in csv.DictReader(f)]


def graficar(ruta_csv, salida):
    riesgos = leer_riesgos(ruta_csv)

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    fig = plt.figure(figsize=(14, 8), dpi=200)
    fig.patch.set_facecolor("white")

    fig.add_artist(Rectangle((0.05, 0.9), 0.007, 0.055, color=ROJO,
                             transform=fig.transFigure))
    fig.text(0.064, 0.91, "RISK MATRIX", fontsize=22, fontweight="bold", color=NEGRO)
    fig.text(0.05, 0.865, "Mallplaza · key risks by likelihood and impact",
             fontsize=11, color=TINTA_SECUNDARIA)

    ax = fig.add_axes([0.1, 0.12, 0.5, 0.7])
    for x in range(1, 6):
        for y in range(1, 6):
            _, _, color = nivel(x * y)
            ax.add_patch(Rectangle((x - 0.5, y - 0.5), 1, 1, facecolor=tinte(color, 0.45),
                                   edgecolor="white", linewidth=3))

    # Riesgos: círculos numerados; varios en la misma celda se reparten
    por_celda = defaultdict(list)
    for r in riesgos:
        por_celda[(r["likelihood"], r["impact"])].append(r)
    desfases = {1: [(0, 0)], 2: [(-0.2, 0), (0.2, 0)],
                3: [(-0.22, 0.15), (0.22, 0.15), (0, -0.2)],
                4: [(-0.2, 0.18), (0.2, 0.18), (-0.2, -0.18), (0.2, -0.18)]}
    for (x, y), grupo in por_celda.items():
        for r, (dx, dy) in zip(grupo, desfases[len(grupo)]):
            _, _, color = nivel(x * y)
            ax.add_patch(Circle((x + dx, y + dy), 0.16, facecolor=color,
                                edgecolor=NEGRO, linewidth=1, zorder=3))
            ax.text(x + dx, y + dy, r["id"], ha="center", va="center", zorder=4,
                    fontsize=10, fontweight="bold",
                    color="white" if color == ROJO else NEGRO)

    ax.set_xlim(0.5, 5.5)
    ax.set_ylim(0.5, 5.5)
    ax.set_aspect("equal")
    ax.set_xticks(range(1, 6), ETIQUETAS)
    ax.set_yticks(range(1, 6), ETIQUETAS)
    ax.set_xlabel("Likelihood", color=TINTA_SECUNDARIA, fontsize=11, labelpad=10)
    ax.set_ylabel("Impact", color=TINTA_SECUNDARIA, fontsize=11, labelpad=10)
    ax.tick_params(length=0, colors=NEGRO, labelsize=10)
    for lado in ax.spines.values():
        lado.set_visible(False)

    # Lista de riesgos ordenada por puntaje
    x0, y = 0.63, 0.8
    fig.text(x0, y, "Risks", fontsize=12, fontweight="bold", color=NEGRO)
    y -= 0.05
    for r in sorted(riesgos, key=lambda r: -r["likelihood"] * r["impact"]):
        _, nombre, color = nivel(r["likelihood"] * r["impact"])
        fig.text(x0 + 0.009, y + 0.007, f"{r['id']:>2}", ha="center", va="center",
                 fontsize=8, fontweight="bold", family="DejaVu Sans Mono",
                 color="white" if color == ROJO else NEGRO,
                 bbox=dict(boxstyle="circle,pad=0.3", facecolor=color,
                           edgecolor=NEGRO, linewidth=0.8))
        fig.text(x0 + 0.025, y, r["risk"], fontsize=10.5, color=NEGRO)
        fig.text(0.96, y, f"L{r['likelihood']} × I{r['impact']} · {nombre}",
                 ha="right", fontsize=9, color=TINTA_SECUNDARIA)
        y -= 0.052

    # Leyenda de niveles
    y -= 0.03
    for i, (_, nombre, color) in enumerate(NIVELES):
        fig.add_artist(Rectangle((x0 + i * 0.1, y), 0.018, 0.022,
                                 facecolor=tinte(color, 0.45),
                                 transform=fig.transFigure))
        fig.text(x0 + i * 0.1 + 0.024, y + 0.004, nombre, fontsize=9.5,
                 color=TINTA_SECUNDARIA)

    fig.text(0.05, 0.03, "Score = likelihood × impact (1–5 scale). "
             "Low ≤ 4 · Medium 5–9 · High ≥ 10.   Source: Team analysis.",
             fontsize=8, color=TINTA_SECUNDARIA)

    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")


if __name__ == "__main__":
    aqui = Path(__file__).parent
    graficar(sys.argv[1] if len(sys.argv) > 1 else aqui / "riesgos.csv",
             sys.argv[2] if len(sys.argv) > 2 else aqui / "matriz_riesgo.png")
