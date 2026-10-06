"""Mallplaza: sensibilidad del resultado al aumento de la UF (reajuste de deuda indexada).

Punto base: +3% de UF → −38.691 MM CLP (atribuido a Nota 26, EEFF jun-26; pendiente
de corroboración). +1% y +2% son extrapolaciones lineales.

Uso:
    python graficos/reajuste_uf.py [salida_sin_extension]
"""
import sys

import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Rectangle

AZUL = "#14284b"
AZUL_SECUNDARIO = "#5a6782"
ROJO = "#b23a48"
ROJO_CLARO = "#e8c3c8"
LINEA = "#d5d9e0"
FONDO_CAJA = "#f4f6f9"

IMPACTO_BASE = -38_691  # MM CLP para +3% de UF
VARIACION_BASE = 3
ESCENARIOS = [1, 2, 3]


def impacto(variacion_pct):
    return IMPACTO_BASE * variacion_pct / VARIACION_BASE


def clp(x):
    return f"{x:,.0f}".replace(",", ".").replace("-", "−")


def graficar(salida):
    valores = {v: impacto(v) for v in ESCENARIOS}
    assert [round(v) for v in valores.values()] == [-12_897, -25_794, -38_691]

    plt.rcParams.update({"font.family": "DejaVu Sans", "svg.fonttype": "none"})
    fig = plt.figure(figsize=(16, 9))
    fig.patch.set_facecolor("white")

    fig.text(0.05, 0.915, "Mallplaza: sensibilidad al aumento de la UF",
             fontsize=28, fontweight="bold", color=AZUL)
    fig.text(0.05, 0.868, "El reajuste de la deuda indexada reduce el resultado contable",
             fontsize=16, color=AZUL_SECUNDARIO)

    # 1. Secuencia de tres pasos
    pasos = [("Sube la UF", AZUL), ("Aumenta el valor de\nla deuda en CLP", AZUL),
             ("Resultado negativo\npor reajuste", ROJO)]
    ancho, alto, y0, x0, sep = 0.235, 0.105, 0.69, 0.05, 0.0925
    for i, (texto, color) in enumerate(pasos):
        x = x0 + i * (ancho + sep)
        fig.add_artist(FancyBboxPatch(
            (x, y0), ancho, alto, boxstyle="round,pad=0,rounding_size=0.012",
            transform=fig.transFigure, facecolor="white", edgecolor=color, linewidth=2))
        fig.text(x + 0.018, y0 + alto - 0.022, str(i + 1), fontsize=12,
                 fontweight="bold", color=color, va="top")
        fig.text(x + ancho / 2, y0 + alto / 2, texto, fontsize=16, fontweight="bold",
                 color=color, ha="center", va="center", linespacing=1.3)
        if i < len(pasos) - 1:
            fig.add_artist(FancyArrowPatch(
                (x + ancho + 0.015, y0 + alto / 2), (x + ancho + sep - 0.015, y0 + alto / 2),
                transform=fig.transFigure, arrowstyle="-|>", mutation_scale=26,
                color=AZUL_SECUNDARIO, linewidth=2))
    fig.text(0.05, 0.645, "La cantidad de UF adeudada permanece constante.",
             fontsize=14, color=AZUL, style="italic")

    # 2. Barras horizontales: impacto negativo desde cero hacia la izquierda
    ax = fig.add_axes([0.13, 0.17, 0.48, 0.40])
    ys = list(range(len(ESCENARIOS)))[::-1]
    for y, v in zip(ys, ESCENARIOS):
        base = v == VARIACION_BASE
        ax.barh(y, valores[v], height=0.56, color=ROJO if base else ROJO_CLARO,
                edgecolor=ROJO, linewidth=0 if base else 1.2,
                hatch=None if base else "///")
        ax.text(valores[v] - 900, y, clp(valores[v]), ha="right", va="center",
                fontsize=16, fontweight="bold", color=ROJO if base else AZUL)
        ax.text(800, y, f"+{v}%", ha="left", va="center", fontsize=16,
                fontweight="bold", color=AZUL)
    ax.axvline(0, color=AZUL, linewidth=1.4)
    ax.set_xlim(-47_500, 5_200)
    ax.set_ylim(-0.6, len(ESCENARIOS) - 0.4)
    ax.set_yticks([])
    ax.set_xticks([-40_000, -30_000, -20_000, -10_000, 0])
    ax.xaxis.set_major_formatter(lambda x, _: clp(x))
    ax.tick_params(axis="x", length=0, colors=AZUL_SECUNDARIO, labelsize=11)
    ax.grid(axis="x", color=LINEA, linewidth=0.8)
    ax.set_axisbelow(True)
    for lado in ax.spines.values():
        lado.set_visible(False)
    ax.set_xlabel("Impacto estimado en resultado (millones de CLP)",
                  fontsize=12, color=AZUL_SECUNDARIO, labelpad=8)
    fig.text(0.611, 0.585, "Aumento del\nvalor de la UF", fontsize=11,
             color=AZUL_SECUNDARIO, ha="left", va="bottom", linespacing=1.2)

    # Leyenda: punto base vs extrapolación
    for i, (texto, relleno, trama) in enumerate(
            [("Punto base del análisis (+3%)", ROJO, None),
             ("Extrapolación lineal (+1% y +2%)", ROJO_CLARO, "///")]):
        x = 0.13 + i * 0.25
        fig.add_artist(Rectangle((x, 0.073), 0.018, 0.026, transform=fig.transFigure,
                                 facecolor=relleno, edgecolor=ROJO,
                                 linewidth=0 if trama is None else 1, hatch=trama))
        fig.text(x + 0.025, 0.086, texto, fontsize=12, color=AZUL, va="center")

    # 3. Recuadro aclaratorio
    fig.add_artist(FancyBboxPatch(
        (0.70, 0.245), 0.25, 0.30, boxstyle="round,pad=0,rounding_size=0.01",
        transform=fig.transFigure, facecolor=FONDO_CAJA, edgecolor="none"))
    fig.add_artist(Rectangle((0.70, 0.245), 0.005, 0.30, transform=fig.transFigure,
                             color=AZUL))
    fig.text(0.722, 0.395,
             "No es una caída del precio\ndel bono ni una salida\ninmediata de caja por el\n"
             "mismo monto: es el aumento\nde la obligación expresada\nen pesos.",
             fontsize=15, color=AZUL, va="center", linespacing=1.5)

    fig.text(0.05, 0.022,
             "Punto base del análisis: +3% de UF → −38.691 millones de CLP, atribuido a la Nota 26 "
             "de los EEFF de junio de 2026. Pendiente de corroboración directa del texto original.\n"
             "Los escenarios de +1% y +2% son extrapolaciones lineales, no cifras adicionales "
             "divulgadas por la compañía.",
             fontsize=10, color=AZUL_SECUNDARIO, linespacing=1.4)

    fig.savefig(f"{salida}.png", dpi=200, facecolor="white")
    fig.savefig(f"{salida}.svg", facecolor="white")
    print(f"Gráfico guardado en {salida}.png y {salida}.svg")


if __name__ == "__main__":
    graficar(sys.argv[1] if len(sys.argv) > 1 else "reajuste_uf")
