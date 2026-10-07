"""Football field de Mallplaza: rango de valor por acción por método de múltiplos vs DCF y precio.

Lee la hoja Valor_Implicito (ya recalculada) del libro de comparables.

Uso:
    python graficos/football_field.py <Comparables_..._valoracion.xlsx> [salida.png]
"""
import sys

import matplotlib.pyplot as plt
import openpyxl
from matplotlib.lines import Line2D
from matplotlib.patches import Patch, Rectangle

from sensibilidad_wacc import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA

GRIS_REF = "#c9c9c9"
FONDO = {"facecolor": "white", "edgecolor": "none", "pad": 1.5}
NOMBRES = {"Core": "core peers", "Chile (Parque Arauco, Cencosud Shopping)": "Chilean peers"}


def leer(ruta):
    ws = openpyxl.load_workbook(ruta, data_only=True)["Valor_Implicito"]
    filas = []
    r = 20
    while ws.cell(r, 1).value:
        mult, base, grupo = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value
        base = "Jun-26" if mult == "P/BV" else base
        filas.append((f"{mult} {base} · {NOMBRES[grupo]}", mult in ("P/E", "P/BV"),
                      ws.cell(r, 8).value, ws.cell(r, 9).value, ws.cell(r, 10).value))
        r += 1
    assert all(v is not None for f in filas for v in f[2:]), "recalcular el libro antes de graficar"
    c = {ws.cell(i, 1).value: ws.cell(i, 3).value for i in range(r, r + 8) if ws.cell(i, 1).value}
    precio = ws["C13"].value
    dcf = ws["C14"].value
    objetivo = next(v for k, v in c.items() if k.startswith("Precio objetivo"))
    return filas, precio, dcf, objetivo


def graficar(ruta, salida):
    filas, precio, dcf, objetivo = leer(ruta)
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    fig = plt.figure(figsize=(12, 7), dpi=200)
    fig.patch.set_facecolor("white")
    fig.add_artist(Rectangle((0.05, 0.905), 0.008, 0.055, color=ROJO, transform=fig.transFigure))
    fig.text(0.065, 0.915, "VALUATION SUMMARY", fontsize=22, fontweight="bold", color=NEGRO)
    fig.text(0.05, 0.87, "Mallplaza · value per share (CLP) implied by peer multiples "
             "(25th–75th percentile, median marked) vs DCF and share price", fontsize=11,
             color=TINTA_SECUNDARIA)

    ax = fig.add_axes([0.3, 0.17, 0.66, 0.66])
    for y, (nombre, referencia, lo, med, hi) in enumerate(reversed(filas)):
        color = GRIS_REF if referencia else CELESTE
        ax.barh(y, hi - lo, left=lo, height=0.55, color=color)
        ax.plot([med, med], [y - 0.27, y + 0.27], color=NEGRO, linewidth=2)
        ax.text(lo - 30, y, f"{lo:,.0f}", ha="right", va="center", fontsize=8.5, color=TINTA_SECUNDARIA,
                zorder=4, bbox=FONDO)
        ax.text(hi + 30, y, f"{hi:,.0f}", ha="left", va="center", fontsize=8.5, color=TINTA_SECUNDARIA,
                zorder=4, bbox=FONDO)
        ax.text(-0.02, y, nombre, transform=ax.get_yaxis_transform(), ha="right", va="center",
                fontsize=10, color=NEGRO)
    n = len(filas)
    for x, color, estilo, texto in ((precio, NEGRO, "--", f"Price {precio:,.0f}"),
                                    (dcf, ROJO, "-", f"DCF {dcf:,.0f}"),
                                    (objetivo, AMARILLO, "-", f"Target {objetivo:,.0f}")):
        ax.axvline(x, color=color, linestyle=estilo, linewidth=1.6, zorder=3)
    ax.text(precio, n - 0.35, f"Price {precio:,.0f}", ha="center", va="bottom", fontsize=9,
            fontweight="bold", color=NEGRO, zorder=4, bbox=FONDO)
    ax.text(dcf + 20, n - 0.35, f"DCF {dcf:,.0f}", ha="left", va="bottom", fontsize=9,
            fontweight="bold", color=ROJO, zorder=4, bbox=FONDO)
    ax.text(objetivo - 15, -0.95, f"Target {objetivo:,.0f}", ha="right", va="center", fontsize=9,
            fontweight="bold", color=NEGRO)
    xs = [v for f in filas for v in f[2:]] + [precio, dcf]
    ax.set_xlim(min(xs) - 300, max(xs) + 300)
    ax.set_ylim(-1.2, n - 0.1)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:,.0f}")
    ax.tick_params(axis="x", length=0, colors=TINTA_SECUNDARIA, labelsize=9)
    ax.grid(axis="x", color=GRILLA, linewidth=0.8)
    ax.set_axisbelow(True)
    for lado in ax.spines.values():
        lado.set_visible(False)
    ax.set_xlabel("Value per share (CLP)", color=TINTA_SECUNDARIA, labelpad=8)

    fig.legend(handles=[Patch(color=CELESTE, label="EV multiples (main)"),
                        Patch(color=GRIS_REF, label="Equity multiples (reference only)"),
                        Line2D([], [], color=NEGRO, linewidth=2, label="Median")],
               loc="lower left", bbox_to_anchor=(0.3, 0.035), ncol=3, frameon=False, fontsize=9)
    fig.text(0.05, 0.012, "Source: Bloomberg (prices at 30-Sep-2026, LTM 3Q25–2Q26, BEst consensus); "
             "Team analysis.", fontsize=8, color=TINTA_SECUNDARIA)
    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    graficar(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "football_field.png")
