"""Mallplaza: sensibilidad del resultado al aumento de la UF (reajuste de deuda indexada).

Punto base: +3% de UF → −38.691 MM CLP (Nota 26, EEFF jun-26). +1% y +2% son
extrapolaciones lineales.

Uso:
    python graficos/reajuste_uf.py [salida_sin_extension]
"""
import sys

import matplotlib.pyplot as plt

AZUL = "#14284b"
AZUL_SECUNDARIO = "#5a6782"
ROJO = "#b23a48"
ROJO_CLARO = "#e8c3c8"
LINEA = "#d5d9e0"

IMPACTO_BASE = -38_691  # MM CLP para +3% de UF
VARIACION_BASE = 3
ESCENARIOS = [1, 2, 3]


def impacto(variacion_pct):
    return IMPACTO_BASE * variacion_pct / VARIACION_BASE


def clp(x):
    return f"{x:,.0f}".replace("-", "−")


def graficar(salida):
    valores = {v: impacto(v) for v in ESCENARIOS}
    assert [round(v) for v in valores.values()] == [-12_897, -25_794, -38_691]

    plt.rcParams.update({"font.family": "DejaVu Sans", "svg.fonttype": "none"})
    fig = plt.figure(figsize=(16, 9))
    fig.patch.set_facecolor("white")

    fig.text(0.06, 0.885, "Mallplaza: sensitivity to a rise in the UF",
             fontsize=30, fontweight="bold", color=AZUL)

    ax = fig.add_axes([0.06, 0.17, 0.76, 0.6])
    ys = list(range(len(ESCENARIOS)))[::-1]
    for y, v in zip(ys, ESCENARIOS):
        base = v == VARIACION_BASE
        ax.barh(y, valores[v], height=0.5, color=ROJO if base else ROJO_CLARO,
                edgecolor=ROJO, linewidth=0 if base else 1.2,
                hatch=None if base else "///")
        ax.text(valores[v] - 700, y, clp(valores[v]), ha="right", va="center",
                fontsize=19, fontweight="bold", color=ROJO if base else AZUL)
        ax.text(900, y + 0.07, f"+{v}%", ha="left", va="center", fontsize=19,
                fontweight="bold", color=AZUL)
        ax.text(900, y - 0.17, "base case" if base else "linear extrapolation",
                ha="left", va="center", fontsize=11, color=AZUL_SECUNDARIO)
    ax.axvline(0, color=AZUL, linewidth=1.4)
    ax.set_xlim(-45_500, 9_000)
    ax.set_ylim(-0.55, len(ESCENARIOS) - 0.45)
    ax.set_yticks([])
    ax.set_xticks([-40_000, -30_000, -20_000, -10_000, 0])
    ax.xaxis.set_major_formatter(lambda x, _: clp(x))
    ax.tick_params(axis="x", length=0, colors=AZUL_SECUNDARIO, labelsize=13)
    ax.grid(axis="x", color=LINEA, linewidth=0.8)
    ax.set_axisbelow(True)
    for lado in ax.spines.values():
        lado.set_visible(False)
    ax.set_xlabel("Estimated impact on earnings (CLP million)",
                  fontsize=14, color=AZUL_SECUNDARIO, labelpad=10)
    ax.text(900, len(ESCENARIOS) - 0.5, "Increase in UF value", ha="left",
            va="bottom", fontsize=13, color=AZUL_SECUNDARIO)

    fig.text(0.06, 0.04, "Source: Mallplaza financial statements, 2Q26.",
             fontsize=12, color=AZUL_SECUNDARIO)

    fig.savefig(f"{salida}.png", dpi=200, facecolor="white")
    fig.savefig(f"{salida}.svg", facecolor="white")
    print(f"Gráfico guardado en {salida}.png y {salida}.svg")


if __name__ == "__main__":
    graficar(sys.argv[1] if len(sys.argv) > 1 else "reajuste_uf")
