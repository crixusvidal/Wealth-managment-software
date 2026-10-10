"""Football field de Mallplaza: valor por acción implícito por múltiplos de cada par vs DCF y precio.

Lee el modelo consolidado (hojas Comparables, Precio_Objetivo, DCF y Supuestos). Cada fila muestra
el rango entre pares del set común, un punto por par (Parque Arauco y Cencosud Shopping con sus
colores corporativos) y el múltiplo ponderado 70% Chile / 30% set común.

Uso:
    python graficos/football_field.py <Valoracion_Consolidada.xlsx> [salida.png]
"""
import sys

import matplotlib.pyplot as plt
import openpyxl
from matplotlib.lines import Line2D
from matplotlib.patches import Patch, Rectangle

from cuota_mercado import MORADO_CENCOSUD, VERDE_PARQUE_ARAUCO
from reajuste_uf import ROJO_CLARO
from sensibilidad_wacc import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA

CELESTE_CLARO = "#c9e3f5"
FONDO = {"facecolor": "white", "edgecolor": "none", "pad": 1.5}
COLOR_PAR = {"Parque Arauco": VERDE_PARQUE_ARAUCO, "Cencosud Shopping": MORADO_CENCOSUD}
# (columna en Comparables, fila en Precio_Objetivo, etiqueta, métrica de Mallplaza, ¿múltiplo de EV?)
MULTIPLOS = [("P", 5, "EV/EBITDA LTM", "K5", True), ("Q", 6, "EV/EBITDA NTM", "O5", True),
             ("R", 7, "EV/EBIT LTM", "L5", True), ("T", 9, "P/BV Jun-26", "N5", False)]


def leer(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    c, po, s, d = wb["Comparables"], wb["Precio_Objetivo"], wb["Supuestos"], wb["DCF"]
    deuda, nci, acciones = s["B15"].value, s["B12"].value, s["B10"].value
    filas = []
    for col, r_po, nombre, celda, es_ev in MULTIPLOS:
        metrica = c[celda].value
        pares = []
        for r in range(6, 13):
            m = c[f"{col}{r}"].value
            if c[f"D{r}"].value != 1 or not isinstance(m, (int, float)):
                continue  # fuera del set común (Fibra Danhos) o NM
            ev = m * metrica
            pares.append((c[f"A{r}"].value, m, (ev - deuda - nci if es_ev else ev) / acciones))
        filas.append((nombre, es_ev, pares, po[f"B{r_po}"].value, po[f"D{r_po}"].value))
    grilla = [d.cell(r, k).value for r in range(61, 66) for k in range(3, 7)]
    dcf = (min(grilla), d["B45"].value, max(grilla))
    pe = (po["B8"].value, po["D8"].value)
    return filas, dcf, pe, po["B15"].value, po["B16"].value, po["B14"].value


def graficar(ruta, salida):
    filas, dcf, pe, objetivo, precio, peso_dcf = leer(ruta)
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    fig = plt.figure(figsize=(12, 7), dpi=200)
    fig.patch.set_facecolor("white")
    fig.add_artist(Rectangle((0.05, 0.905), 0.008, 0.055, color=ROJO, transform=fig.transFigure))
    fig.text(0.065, 0.915, "VALUATION SUMMARY", fontsize=22, fontweight="bold", color=NEGRO)
    fig.text(0.05, 0.87, "Mallplaza · value per share (CLP) implied by each peer's multiple vs DCF "
             "and share price · marker = 70% Chile / 30% peer-set median", fontsize=11,
             color=TINTA_SECUNDARIA)

    ax = fig.add_axes([0.25, 0.2, 0.71, 0.63])
    n = len(filas) + 1
    # DCF: rango de la grilla de sensibilidad (WACC ±1 pp, g 3,0–4,5%)
    y = n - 1
    ax.barh(y, dcf[2] - dcf[0], left=dcf[0], height=0.5, color=ROJO_CLARO)
    ax.plot([dcf[1]] * 2, [y - 0.3, y + 0.3], color=ROJO, linewidth=3, solid_capstyle="butt")
    ax.text(dcf[1], y + 0.36, f"{dcf[1]:,.0f}", ha="center", va="bottom", fontsize=9.5,
            fontweight="bold", color=ROJO, zorder=4, bbox=FONDO)
    for v, lado in ((dcf[0], -1), (dcf[2], 1)):
        ax.text(v + lado * 35, y, f"{v:,.0f}", ha="left" if lado > 0 else "right", va="center",
                fontsize=8.5, color=TINTA_SECUNDARIA, zorder=4, bbox=FONDO)
    ax.text(-0.02, y + 0.12, "DCF (FCFF)", transform=ax.get_yaxis_transform(), ha="right",
            va="center", fontsize=10.5, fontweight="bold", color=NEGRO)
    ax.text(-0.02, y - 0.2, "WACC ±1 pp · g 3.0–4.5%", transform=ax.get_yaxis_transform(),
            ha="right", va="center", fontsize=8.5, color=TINTA_SECUNDARIA)

    for y, (nombre, es_ev, pares, mult, valor) in zip(range(n - 2, -1, -1), filas):
        vals = [p[2] for p in pares]
        ax.barh(y, max(vals) - min(vals), left=min(vals), height=0.5,
                color=CELESTE_CLARO if es_ev else "white", edgecolor=CELESTE,
                linewidth=0 if es_ev else 1, hatch=None if es_ev else "////")
        for par, _, v in sorted(pares, key=lambda p: p[0] in COLOR_PAR):
            chile = par in COLOR_PAR
            ax.scatter(v, y, s=70 if chile else 38, color=COLOR_PAR.get(par, CELESTE),
                       edgecolor="white", linewidth=1, zorder=3 + chile)
        ax.plot([valor] * 2, [y - 0.3, y + 0.3], color=NEGRO, linewidth=3, solid_capstyle="butt",
                zorder=5)
        ax.text(valor, y + 0.36, f"{valor:,.0f}", ha="center", va="bottom", fontsize=9.5,
                fontweight="bold", color=NEGRO, zorder=6, bbox=FONDO)
        for v, lado in ((min(vals), -1), (max(vals), 1)):
            ax.text(v + lado * 70, y, f"{v:,.0f}", ha="left" if lado > 0 else "right",
                    va="center", fontsize=8.5, color=TINTA_SECUNDARIA, zorder=4, bbox=FONDO)
        ax.text(-0.02, y + 0.12, nombre, transform=ax.get_yaxis_transform(), ha="right",
                va="center", fontsize=10.5, fontweight="bold", color=NEGRO)
        ax.text(-0.02, y - 0.2, f"weighted {mult:.1f}x" + ("" if es_ev else " · reference"),
                transform=ax.get_yaxis_transform(), ha="right", va="center", fontsize=8.5,
                color=TINTA_SECUNDARIA)

    ax.axvline(precio, color=NEGRO, linestyle="--", linewidth=1.4, zorder=2)
    ax.axvline(objetivo, color=AMARILLO, linewidth=2, zorder=2)
    ax.text(precio + 25, -0.75, f"Price {precio:,.0f}", ha="left", va="center", fontsize=9,
            fontweight="bold", color=NEGRO)
    ax.text(objetivo - 25, -0.75, f"Target {objetivo:,.0f}", ha="right", va="center", fontsize=9,
            fontweight="bold", color=NEGRO)
    xs = [p[2] for f in filas for p in f[2]] + [dcf[0], dcf[2], precio]
    ax.set_xlim(min(xs) - 450, max(xs) + 450)
    ax.set_ylim(-1.05, n - 0.35)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:,.0f}")
    ax.tick_params(axis="x", length=0, colors=TINTA_SECUNDARIA, labelsize=9)
    ax.grid(axis="x", color=GRILLA, linewidth=0.8)
    ax.set_axisbelow(True)
    for lado in ax.spines.values():
        lado.set_visible(False)
    ax.set_xlabel("Value per share (CLP)", color=TINTA_SECUNDARIA, labelpad=8)

    punto = lambda color, texto: Line2D([], [], marker="o", linestyle="", markersize=8,
                                        markerfacecolor=color, markeredgecolor="white", label=texto)
    fig.legend(handles=[punto(VERDE_PARQUE_ARAUCO, "Parque Arauco"),
                        punto(MORADO_CENCOSUD, "Cencosud Shopping"),
                        punto(CELESTE, "Multiplan, Allos, Iguatemi, IRSA"),
                        Line2D([], [], color=NEGRO, linewidth=3, label="Weighted multiple"),
                        Patch(facecolor=ROJO_CLARO, label="DCF sensitivity range")],
               loc="lower left", bbox_to_anchor=(0.25, 0.055), ncol=5, frameon=False, fontsize=9,
               handletextpad=0.4, columnspacing=1.2)
    fig.text(0.05, 0.035, f"Target = {peso_dcf:.0%} DCF + {1 - peso_dcf:.0%} EV/EBITDA LTM. P/E "
             f"excluded: LTM earnings include property revaluation ({pe[0]:.1f}x → CLP {pe[1]:,.0f}). "
             "Peer set = companies with a WACC beta (Fibra Danhos excluded).", fontsize=8,
             color=TINTA_SECUNDARIA)
    fig.text(0.05, 0.012, "Source: Bloomberg (prices at 30-Sep-2026, LTM 3Q25–2Q26, BEst consensus); "
             "Team analysis.", fontsize=8, color=TINTA_SECUNDARIA)
    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    graficar(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "football_field.png")
