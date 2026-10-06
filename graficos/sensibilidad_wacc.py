"""Gráfico de sensibilidad del valor por acción de Mallplaza a la WACC y a g.

Lee la grilla "Sensibilidad: valor por acción (CLP) — WACC (filas) vs g (columnas)"
de la hoja DCF_Valuation del archivo de exhibits y guarda un PNG.

Uso:
    python graficos/sensibilidad_wacc.py <Mallplaza_FA_Exhibits.xlsx> [salida.png]
"""
import sys

import matplotlib.pyplot as plt
import openpyxl
from matplotlib.ticker import FuncFormatter

# Colores corporativos Mallplaza
CELESTE = "#64b0e1"
AMARILLO = "#ecb100"
ROJO = "#e9004b"
NEGRO = "#000000"
TINTA_SECUNDARIA = "#5f5f5f"
GRILLA = "#e6e6e6"

TITULO_GRILLA = "Sensibilidad: valor por acción (CLP) — WACC (filas) vs g (columnas)"


def leer_grilla(ruta):
    ws = openpyxl.load_workbook(ruta, data_only=True)["DCF_Valuation"]
    fila_titulo = next(
        c.row for fila in ws.iter_rows() for c in fila if c.value == TITULO_GRILLA
    )
    encabezado = fila_titulo + 1
    gs = [ws.cell(encabezado, c).value for c in range(4, 7)]
    waccs, valores = [], {g: [] for g in gs}
    r = encabezado + 1
    while ws.cell(r, 3).value is not None:
        waccs.append(ws.cell(r, 3).value)
        for i, g in enumerate(gs):
            valores[g].append(ws.cell(r, 4 + i).value)
        r += 1
    wacc_base = ws["J31"].value
    g_base = ws["J37"].value
    precio = ws["J50"].value
    return waccs, valores, wacc_base, g_base, precio


def pct(x, decimales=1):
    return f"{x * 100:.{decimales}f}%".replace(".", ",")


def miles(x):
    return f"{x:,.0f}".replace(",", ".")


def graficar(ruta_xlsx, salida):
    waccs, valores, wacc_base, g_base, precio = leer_grilla(ruta_xlsx)

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    fig, ax = plt.subplots(figsize=(9, 5.5), dpi=200)
    fig.patch.set_facecolor("white")

    estilos = [(CELESTE, "o"), (ROJO, "s"), (AMARILLO, "^")]
    x = [w * 100 for w in waccs]
    for (g, serie), (color, marcador) in zip(valores.items(), estilos):
        es_base = abs(g - g_base) < 1e-9
        ax.plot(
            x, serie, color=color, linewidth=2.5 if es_base else 2,
            marker=marcador, markersize=8, markeredgecolor=NEGRO,
            markeredgewidth=0.8, zorder=3,
            label=f"g = {pct(g)}" + (" (caso base)" if es_base else ""),
        )
        ax.annotate(
            f"g = {pct(g)}", (x[-1], serie[-1]), xytext=(10, 0),
            textcoords="offset points", va="center", fontsize=9, color=NEGRO,
        )

    # Caso base
    base = valores[g_base][[round(w, 6) for w in waccs].index(round(wacc_base, 6))]
    ax.scatter([wacc_base * 100], [base], s=200, facecolor="none",
               edgecolor=NEGRO, linewidth=1.5, zorder=4)
    ax.annotate(
        f"Caso base\nWACC {pct(wacc_base, 2)} · g {pct(g_base)}\nCLP {miles(base)}",
        (wacc_base * 100, base), xytext=(-125, -55), textcoords="offset points",
        fontsize=9, color=NEGRO,
        arrowprops=dict(arrowstyle="-", color=NEGRO, linewidth=0.8),
    )

    # Precio de referencia
    ax.axhline(precio, color=NEGRO, linewidth=1, linestyle=(0, (4, 3)), zorder=2)
    ax.text(x[0], precio, f"Precio implícito CLP {miles(precio)}",
            va="bottom", ha="left", fontsize=9, color=NEGRO)

    ax.set_xticks(x)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: pct(v / 100, 2)))
    ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: miles(v)))
    ax.set_ylim(1400, precio * 1.06)
    ax.set_xlim(x[0] - 0.15, x[-1] + 0.35)
    ax.set_xlabel("WACC nominal CLP", color=TINTA_SECUNDARIA)
    ax.set_ylabel("Valor por acción (CLP)", color=TINTA_SECUNDARIA)

    ax.grid(axis="y", color=GRILLA, linewidth=0.8)
    ax.set_axisbelow(True)
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    for lado in ("left", "bottom"):
        ax.spines[lado].set_color("#bdbdbd")
    ax.tick_params(colors=TINTA_SECUNDARIA)

    fig.text(0.07, 0.955, "Mallplaza — Sensibilidad del valor por acción a la WACC",
             fontsize=14, fontweight="bold", color=NEGRO)
    fig.text(0.07, 0.915,
             "DCF (FCFF), valor por acción en CLP según WACC y crecimiento perpetuo g",
             fontsize=10, color=TINTA_SECUNDARIA)
    ax.legend(loc="upper right", bbox_to_anchor=(1, 0.86), frameon=False, fontsize=9)
    fig.text(0.07, 0.01, "Fuente: Mallplaza_FA_Exhibits, hoja DCF_Valuation.",
             fontsize=8, color=TINTA_SECUNDARIA)

    fig.tight_layout(rect=(0.02, 0.03, 1, 0.9))
    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    graficar(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "sensibilidad_wacc.png")
