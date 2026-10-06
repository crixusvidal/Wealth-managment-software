"""Mapa de calor de la sensibilidad del valor por acción de Mallplaza (WACC vs g).

Usa la misma grilla que sensibilidad_wacc.py (hoja DCF_Valuation). Rojo = valor
bajo el caso base, celeste = sobre el caso base, gris claro = cerca del caso base.

Uso:
    python graficos/heatmap_sensibilidad_wacc.py <Mallplaza_FA_Exhibits.xlsx> [salida.png]
"""
import sys

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
from matplotlib.patches import Rectangle

from sensibilidad_wacc import (AMARILLO, CELESTE, NEGRO, ROJO, TINTA_SECUNDARIA,
                               leer_grilla, miles, pct)

NEUTRO = "#f2f2f2"


def color_texto(rgba):
    r, g, b = rgba[:3]
    return "white" if 0.299 * r + 0.587 * g + 0.114 * b < 0.55 else NEGRO


def graficar(ruta_xlsx, salida):
    waccs, valores, wacc_base, g_base, precio = leer_grilla(ruta_xlsx)
    gs = list(valores)
    z = np.array([valores[g] for g in gs]).T  # filas = WACC, columnas = g
    i_base = [round(w, 6) for w in waccs].index(round(wacc_base, 6))
    j_base = [round(g, 6) for g in gs].index(round(g_base, 6))
    base = z[i_base, j_base]

    # Filas de mayor a menor WACC: el valor sube hacia abajo a la derecha
    orden = np.argsort(waccs)[::-1]
    z = z[orden]
    waccs = [waccs[i] for i in orden]
    i_base = list(orden).index(i_base)

    cmap = LinearSegmentedColormap.from_list("mallplaza", [ROJO, NEUTRO, CELESTE])
    norm = TwoSlopeNorm(vmin=z.min(), vcenter=base, vmax=z.max())

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    fig = plt.figure(figsize=(8, 9), dpi=200)
    fig.patch.set_facecolor("white")

    # Encabezado
    fig.add_artist(Rectangle((0.07, 0.935), 0.012, 0.04, color=ROJO,
                             transform=fig.transFigure))
    fig.text(0.095, 0.94, "ANÁLISIS DE SENSIBILIDAD", fontsize=20,
             fontweight="bold", color=NEGRO)
    fig.text(0.07, 0.905, "Mallplaza · valor por acción (CLP) según WACC nominal y g",
             fontsize=10.5, color=TINTA_SECUNDARIA)

    # Escenarios
    escenarios = [
        ("PESIMISTA", waccs[0], gs[0], z[0, 0], ROJO),
        ("BASE", wacc_base, g_base, base, AMARILLO),
        ("OPTIMISTA", waccs[-1], gs[-1], z[-1, -1], CELESTE),
    ]
    for k, (nombre, w, g, v, color) in enumerate(escenarios):
        x0 = 0.07 + k * 0.29
        fig.add_artist(Rectangle((x0, 0.865), 0.26, 0.006, color=color,
                                 transform=fig.transFigure))
        fig.text(x0, 0.835, nombre, fontsize=12, fontweight="bold", color=NEGRO)
        fig.text(x0, 0.795, f"CLP {miles(v)}", fontsize=17, fontweight="bold",
                 color=NEGRO)
        fig.text(x0, 0.77, f"WACC {pct(w, 2)} · g {pct(g)}", fontsize=9.5,
                 color=TINTA_SECUNDARIA)

    # Mapa de calor
    ax = fig.add_axes([0.17, 0.2, 0.76, 0.5])
    im = ax.imshow(z, cmap=cmap, norm=norm, aspect="auto")
    for i in range(z.shape[0]):
        for j in range(z.shape[1]):
            ax.text(j, i, miles(z[i, j]), ha="center", va="center", fontsize=13,
                    fontweight="bold" if (i, j) == (i_base, j_base) else "normal",
                    color=color_texto(cmap(norm(z[i, j]))))
    # Separadores blancos entre celdas
    for i in range(z.shape[0] + 1):
        ax.axhline(i - 0.5, color="white", linewidth=3)
    for j in range(z.shape[1] + 1):
        ax.axvline(j - 0.5, color="white", linewidth=3)
    ax.add_patch(Rectangle((j_base - 0.47, i_base - 0.47), 0.94, 0.94, fill=False,
                           edgecolor=AMARILLO, linewidth=4, zorder=3))

    ax.set_xticks(range(len(gs)), [pct(g) for g in gs])
    ax.set_yticks(range(len(waccs)), [pct(w, 2) for w in waccs])
    ax.set_xlabel("g · crecimiento perpetuo nominal", color=TINTA_SECUNDARIA,
                  labelpad=10)
    ax.set_ylabel("WACC nominal CLP", color=TINTA_SECUNDARIA, labelpad=10)
    ax.tick_params(length=0, colors=NEGRO, labelsize=11)
    for lado in ax.spines.values():
        lado.set_visible(False)

    # Barra de color
    cax = fig.add_axes([0.17, 0.09, 0.76, 0.02])
    cb = fig.colorbar(im, cax=cax, orientation="horizontal")
    cb.set_ticks([z.min(), base, z.max()])
    cb.set_ticklabels([miles(z.min()), f"{miles(base)}\ncaso base", miles(z.max())])
    cb.outline.set_visible(False)
    cax.tick_params(length=0, colors=TINTA_SECUNDARIA, labelsize=9)
    cax.set_title("Valor por acción (CLP)", fontsize=9, color=TINTA_SECUNDARIA,
                  loc="left")

    fig.text(0.07, 0.02,
             f"Precio implícito CLP {miles(precio)}: ninguna combinación de la grilla lo alcanza.  "
             "Fuente: Mallplaza_FA_Exhibits, DCF_Valuation.",
             fontsize=8, color=TINTA_SECUNDARIA)

    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    graficar(sys.argv[1],
             sys.argv[2] if len(sys.argv) > 2 else "heatmap_sensibilidad_wacc.png")
