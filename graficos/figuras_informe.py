"""Versiones de los 4 gráficos dimensionadas para la columna de figuras del informe.

Se dibujan al tamaño exacto de impresión (ancho de la columna derecha de la plantilla
Investment Risks) a 400 dpi, sin título (lo pone el pie de figura en Word).

Uso:
    python graficos/figuras_informe.py <FA_Exhibits.xlsx> <Tornado_Sensibilidad.xlsx> [carpeta_salida]
"""
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm, to_rgb
from matplotlib.patches import Rectangle

from heatmap_sensibilidad_wacc import GS, NEUTRO, color_texto, leer_insumos, valor_por_accion
from reajuste_uf import ESCENARIOS, ROJO_CLARO, VARIACION_BASE, impacto
from sensibilidad_wacc import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA, leer_grilla
from tornado_operacional import escenarios, leer_modelo, leer_rangos, proyectar

ANCHO = 3800 / 1440  # pulgadas: columna derecha de informe/investment_risks.js
ALTOS = {"matriz": 2.55, "tornado": 2.55, "heatmap": 2.0, "uf": 1.75}
DPI = 400

NOMBRES_CORTOS = {
    "EBITDA margin": "EBITDA margin",
    "Same-store rent growth 2027–28": "Same-store rent growth",
    "Maintenance + IT capex": "Maintenance capex",
    "Long-term real rent growth 2029–31": "LT real rent growth",
    "Visitor traffic 2027–28 (parking)": "Visitor traffic",
    "Occupancy": "Occupancy",
    "Tenant sales growth 2027–28": "Tenant sales growth",
}


def nueva_figura(clave):
    fig = plt.figure(figsize=(ANCHO, ALTOS[clave]))
    fig.patch.set_facecolor("white")
    return fig


def fuente(fig, texto="Source: Team analysis."):
    fig.text(0.01, 0.012, texto, fontsize=4.5, color=TINTA_SECUNDARIA, va="bottom")


def limpiar(ax):
    for lado in ax.spines.values():
        lado.set_visible(False)


def matriz(salida):
    fig = nueva_figura("matriz")
    ax = fig.add_axes([0.2, 0.2, 0.78, 0.77])
    etiquetas = ["Very low", "Low", "Medium", "High", "Very high"]
    for x in range(1, 6):
        for y in range(1, 6):
            p = x * y
            color = CELESTE if p <= 4 else AMARILLO if p <= 9 else ROJO
            tinte = tuple(1 - 0.5 * (1 - c) for c in to_rgb(color))
            ax.add_patch(Rectangle((x - 0.5, y - 0.5), 1, 1, facecolor=tinte,
                                   edgecolor="white", linewidth=1.5))
    ax.set_xlim(0.5, 5.5)
    ax.set_ylim(0.5, 5.5)
    ax.set_xticks(range(1, 6), etiquetas)
    ax.set_yticks(range(1, 6), etiquetas)
    ax.tick_params(length=0, colors=NEGRO, labelsize=5, pad=2)
    ax.set_xlabel("Likelihood", fontsize=6, color=TINTA_SECUNDARIA, labelpad=3)
    ax.set_ylabel("Impact", fontsize=6, color=TINTA_SECUNDARIA, labelpad=3)
    limpiar(ax)
    fuente(fig)
    fig.savefig(salida, dpi=DPI, facecolor="white")


def tornado(ruta_fa, ruta_tornado, salida):
    m, _ = leer_modelo(ruta_fa)
    base = proyectar(m)[2]
    filas = []
    for nombre, _, k_lo, k_hi in escenarios(m, leer_rangos(ruta_tornado)):
        filas.append((NOMBRES_CORTOS[nombre], proyectar(m, **k_lo)[2], proyectar(m, **k_hi)[2]))
    filas.sort(key=lambda r: (base - r[1], r[2] - r[1]))

    fig = nueva_figura("tornado")
    ax = fig.add_axes([0.37, 0.19, 0.61, 0.74])
    for y, (nombre, lo, hi) in enumerate(filas):
        ax.barh(y, lo - base, left=base, height=0.6, color=ROJO)
        ax.barh(y, hi - base, left=base, height=0.6, color=CELESTE)
        for v, lado in ((lo, -1), (hi, 1)):
            if abs(v - base) >= 0.5:
                ax.text(v + lado * 8, y, f"{v:,.0f}", va="center", fontsize=5,
                        ha="left" if lado > 0 else "right", color=NEGRO)
        ax.text(-0.03, y, nombre, transform=ax.get_yaxis_transform(), ha="right",
                va="center", fontsize=5.5, color=NEGRO)
    ax.axvline(base, color=NEGRO, linewidth=0.8)
    ax.text(base, len(filas) - 0.45, f"Base {base:,.0f}", ha="center", va="bottom",
            fontsize=5.5, fontweight="bold", color=NEGRO, zorder=4,
            bbox={"facecolor": "white", "edgecolor": "none", "pad": 1})
    extremo = max(max(abs(lo - base), abs(hi - base)) for _, lo, hi in filas)
    ax.set_xlim(base - extremo * 1.35, base + extremo * 1.35)
    ax.set_ylim(-0.6, len(filas) - 0.2)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:,.0f}")
    ax.tick_params(axis="x", length=0, colors=TINTA_SECUNDARIA, labelsize=5)
    ax.grid(axis="x", color=GRILLA, linewidth=0.4)
    ax.set_axisbelow(True)
    limpiar(ax)
    ax.set_xlabel("Value per share (CLP)", fontsize=5.5, color=TINTA_SECUNDARIA, labelpad=2)
    for x, color, texto in ((0.37, ROJO, "Historical low"), (0.64, CELESTE, "Historical high")):
        fig.add_artist(Rectangle((x, 0.035), 0.025, 0.03, color=color, transform=fig.transFigure))
        fig.text(x + 0.035, 0.05, texto, fontsize=5, color=NEGRO, va="center")
    fuente(fig)
    fig.savefig(salida, dpi=DPI, facecolor="white")


def heatmap(ruta_fa, salida):
    waccs, _, wacc_base, g_base, _ = leer_grilla(ruta_fa)
    insumos = leer_insumos(ruta_fa)
    waccs = sorted(waccs, reverse=True)
    z = np.array([[valor_por_accion(w, g, *insumos) for g in GS] for w in waccs])
    i_b = waccs.index(wacc_base)
    j_b = int(np.argmin([abs(g - g_base) for g in GS]))
    cmap = LinearSegmentedColormap.from_list("mallplaza", [ROJO, NEUTRO, CELESTE])
    norm = TwoSlopeNorm(vmin=z.min(), vcenter=z[i_b, j_b], vmax=z.max())

    fig = nueva_figura("heatmap")
    ax = fig.add_axes([0.15, 0.24, 0.83, 0.74])
    ax.imshow(z, cmap=cmap, norm=norm, aspect="auto")
    for i in range(z.shape[0]):
        for j in range(z.shape[1]):
            ax.text(j, i, f"{z[i, j]:,.0f}", ha="center", va="center", fontsize=5.5,
                    fontweight="bold" if (i, j) == (i_b, j_b) else "normal",
                    color=color_texto(cmap(norm(z[i, j]))))
    for k in range(z.shape[0] + 1):
        ax.axhline(k - 0.5, color="white", linewidth=1.2)
    for k in range(z.shape[1] + 1):
        ax.axvline(k - 0.5, color="white", linewidth=1.2)
    ax.add_patch(Rectangle((j_b - 0.45, i_b - 0.42), 0.9, 0.84, fill=False,
                           edgecolor=AMARILLO, linewidth=1.2, zorder=3))
    ax.set_xticks(range(len(GS)), [f"{g:.1%}" for g in GS])
    ax.set_yticks(range(len(waccs)), [f"{w:.2%}" for w in waccs])
    ax.tick_params(length=0, colors=NEGRO, labelsize=5, pad=2)
    ax.set_xlabel("Terminal growth (g)", fontsize=5.5, color=TINTA_SECUNDARIA, labelpad=2)
    ax.set_ylabel("WACC", fontsize=5.5, color=TINTA_SECUNDARIA, labelpad=2)
    limpiar(ax)
    fuente(fig)
    fig.savefig(salida, dpi=DPI, facecolor="white")


def uf(salida):
    fig = nueva_figura("uf")
    ax = fig.add_axes([0.04, 0.27, 0.8, 0.66])
    for y, v in zip(range(len(ESCENARIOS))[::-1], ESCENARIOS):
        base = v == VARIACION_BASE
        valor = impacto(v)
        ax.barh(y, valor, height=0.55, color=ROJO if base else ROJO_CLARO,
                edgecolor=ROJO, linewidth=0 if base else 0.5, hatch=None if base else "////")
        ax.text(valor - 600, y, f"{valor:,.0f}".replace("-", "−"), ha="right", va="center",
                fontsize=5.5, fontweight="bold", color=ROJO if base else NEGRO)
        ax.text(700, y + 0.1, f"+{v}%", ha="left", va="center", fontsize=6,
                fontweight="bold", color=NEGRO)
        ax.text(700, y - 0.22, "base case" if base else "linear extrap.", ha="left",
                va="center", fontsize=4.5, color=TINTA_SECUNDARIA)
    ax.axvline(0, color=NEGRO, linewidth=0.8)
    ax.set_xlim(-49_000, 1_000)
    ax.set_ylim(-0.55, len(ESCENARIOS) - 0.45)
    ax.set_yticks([])
    ax.set_xticks([-40_000, -30_000, -20_000, -10_000, 0])
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:,.0f}".replace("-", "−"))
    ax.tick_params(axis="x", length=0, colors=TINTA_SECUNDARIA, labelsize=5)
    ax.grid(axis="x", color=GRILLA, linewidth=0.4)
    ax.set_axisbelow(True)
    limpiar(ax)
    ax.set_xlabel("Impact on earnings (CLP million)", fontsize=5.5,
                  color=TINTA_SECUNDARIA, labelpad=2)
    fuente(fig, "Source: Mallplaza financial statements, 2Q26.")
    fig.savefig(salida, dpi=DPI, facecolor="white")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    salida = Path(sys.argv[3] if len(sys.argv) > 3 else Path(__file__).parent / "informe")
    salida.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    matriz(salida / "fig1_matriz_riesgo.png")
    tornado(sys.argv[1], sys.argv[2], salida / "fig2_tornado.png")
    heatmap(sys.argv[1], salida / "fig3_heatmap.png")
    uf(salida / "fig4_uf.png")
    print(f"Figuras guardadas en {salida}")
