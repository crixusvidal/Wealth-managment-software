"""Mapa de calor de la sensibilidad del valor por acción de Mallplaza (WACC vs g).

Recalcula la grilla con la misma fórmula de la hoja DCF_Valuation (celda D62) sobre
una malla más fina, y verifica que coincide con la grilla original del Excel.
Rojo = valor bajo el caso base, celeste = sobre el caso base.

Uso:
    python graficos/heatmap_sensibilidad_wacc.py <Mallplaza_FA_Exhibits.xlsx> [salida.png]
"""
import sys

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm
from matplotlib.patches import Rectangle

from sensibilidad_wacc import (AMARILLO, CELESTE, NEGRO, ROJO, TINTA_SECUNDARIA,
                               leer_grilla, miles, pct)

NEUTRO = "#f2f2f2"
GS = np.round(np.arange(0.015, 0.04001, 0.0025), 4)  # 1,50% a 4,00%


def leer_insumos(ruta):
    ws = openpyxl.load_workbook(ruta, data_only=True)["DCF_Valuation"]
    columnas = "JKLMNO"
    fcff = np.array([ws[f"{c}28"].value for c in columnas])
    periodos = np.array([ws[f"{c}32"].value for c in columnas])
    ajustes = ws["J45"].value + ws["J46"].value  # deuda neta y minoritarios (negativos)
    return fcff, periodos, ajustes, ws["J48"].value


def valor_por_accion(wacc, g, fcff, periodos, ajustes, acciones):
    vp_flujos = np.sum(fcff / (1 + wacc) ** periodos)
    vp_terminal = fcff[-1] * (1 + g) / (wacc - g) / (1 + wacc) ** 5.5
    return (vp_flujos + vp_terminal + ajustes) / acciones


def color_texto(rgba):
    r, g, b = rgba[:3]
    return "white" if 0.299 * r + 0.587 * g + 0.114 * b < 0.55 else NEGRO


def graficar(ruta_xlsx, salida):
    waccs_xl, valores_xl, wacc_base, g_base, precio = leer_grilla(ruta_xlsx)
    insumos = leer_insumos(ruta_xlsx)

    for g, serie in valores_xl.items():
        for w, v in zip(waccs_xl, serie):
            assert abs(valor_por_accion(w, g, *insumos) - v) < 1e-6, (w, g)

    # 9,00% a 11,00% cada 0,25%, con la WACC base (10,03%) en lugar de 10,00%
    waccs = sorted({round(w, 4) for w in np.arange(0.09, 0.11001, 0.0025)} - {0.10}
                   | {wacc_base}, reverse=True)
    z = np.array([[valor_por_accion(w, g, *insumos) for g in GS] for w in waccs])
    i_base = waccs.index(wacc_base)
    j_base = int(np.argmin(abs(GS - g_base)))
    base = z[i_base, j_base]

    cmap = LinearSegmentedColormap.from_list("mallplaza", [ROJO, NEUTRO, CELESTE])
    norm = TwoSlopeNorm(vmin=z.min(), vcenter=base, vmax=z.max())

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    fig = plt.figure(figsize=(13, 7.5), dpi=200)
    fig.patch.set_facecolor("white")

    fig.add_artist(Rectangle((0.05, 0.905), 0.008, 0.055, color=ROJO,
                             transform=fig.transFigure))
    fig.text(0.065, 0.915, "ANÁLISIS DE SENSIBILIDAD", fontsize=22,
             fontweight="bold", color=NEGRO)
    fig.text(0.05, 0.87, "Mallplaza · valor por acción (CLP) según WACC nominal y g",
             fontsize=11, color=TINTA_SECUNDARIA)

    ax = fig.add_axes([0.1, 0.24, 0.87, 0.58])
    im = ax.imshow(z, cmap=cmap, norm=norm, aspect="auto")
    for i in range(z.shape[0]):
        for j in range(z.shape[1]):
            ax.text(j, i, miles(z[i, j]), ha="center", va="center", fontsize=10,
                    fontweight="bold" if (i, j) == (i_base, j_base) else "normal",
                    color=color_texto(cmap(norm(z[i, j]))))
    for i in range(z.shape[0] + 1):
        ax.axhline(i - 0.5, color="white", linewidth=2)
    for j in range(z.shape[1] + 1):
        ax.axvline(j - 0.5, color="white", linewidth=2)
    ax.add_patch(Rectangle((j_base - 0.46, i_base - 0.44), 0.92, 0.88, fill=False,
                           edgecolor=AMARILLO, linewidth=3, zorder=3))

    ax.set_xticks(range(len(GS)), [pct(g, 2) for g in GS])
    ax.set_yticks(range(len(waccs)), [pct(w, 2) for w in waccs])
    ax.set_xlabel("g · crecimiento perpetuo nominal", color=TINTA_SECUNDARIA,
                  labelpad=10)
    ax.set_ylabel("WACC nominal CLP", color=TINTA_SECUNDARIA, labelpad=10)
    ax.tick_params(length=0, colors=NEGRO, labelsize=10)
    for lado in ax.spines.values():
        lado.set_visible(False)

    cax = fig.add_axes([0.1, 0.09, 0.87, 0.02])
    cb = fig.colorbar(im, cax=cax, orientation="horizontal")
    cb.set_ticks([z.min(), base, z.max()])
    cb.set_ticklabels([miles(z.min()),
                       f"{miles(base)}\ncaso base (WACC {pct(wacc_base, 2)} · g {pct(g_base)})",
                       miles(z.max())])
    cb.outline.set_visible(False)
    cax.tick_params(length=0, colors=TINTA_SECUNDARIA, labelsize=9)
    cax.set_title("Valor por acción (CLP)", fontsize=9, color=TINTA_SECUNDARIA,
                  loc="left")

    fig.text(0.05, 0.015,
             f"Precio implícito CLP {miles(precio)}.  "
             "Fuente: Mallplaza_FA_Exhibits, DCF_Valuation (fórmula de la grilla D62).",
             fontsize=8, color=TINTA_SECUNDARIA)

    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    graficar(sys.argv[1],
             sys.argv[2] if len(sys.argv) > 2 else "heatmap_sensibilidad_wacc.png")
