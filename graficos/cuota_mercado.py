"""Área apilada: ingresos combinados de Mallplaza, Parque Arauco y Cencosud Shopping y la
participación de cada uno (FY2020–FY2025). FY2022–FY2025 de Bloomberg; FY2020–FY2021 de fuentes
públicas (ver EXTERNOS).

Uso:
    python graficos/cuota_mercado.py <Comparables_..._valoracion.xlsx> [salida.png]
"""
import sys

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.patches import Rectangle

from sensibilidad_wacc import AMARILLO, CELESTE, NEGRO, ROJO, TINTA_SECUNDARIA

EMPRESAS = [("Mallplaza", ROJO, "white"), ("Parque Arauco", CELESTE, NEGRO),
            ("Cencosud Shopping", AMARILLO, NEGRO)]

# Ingresos consolidados (MM CLP) que no están en el libro Bloomberg. Verificar contra los EEFF (CMF).
EXTERNOS = {
    "Mallplaza": {
        2020: 187_329,  # Tabla de EERR IFRS en documento de bonos de Plaza S.A. (2019: 309.598; −39,5%)
        2021: 276_695,  # Mallplaza_FA_Exhibits_v6.xlsx, Inputs rev21 (Bloomberg FA)
    },
    "Parque Arauco": {
        2020: 125_852,  # Prensa (DF) citando reporte a la CMF; −38,8% vs 2019
        2021: 169_859,  # Prensa (DF) citando reporte a la CMF; +35% vs 2020
    },
    "Cencosud Shopping": {
        2020: 146_755,  # Feller Rate y Humphreys (−35,9% vs 2019: 228.990)
        2021: round(286_949.855 / 1.334),  # Estimado: FY2022 (EEFF) / 1,334 (+33,4% informado en prensa)
    },
}
ESTIMADOS = {("Cencosud Shopping", 2021)}


def leer(ruta):
    ws = openpyxl.load_workbook(ruta, data_only=True)["Datos_BBG"]
    ingresos = {}
    for r in range(4, ws.max_row + 1):
        nombre, periodo = ws.cell(r, 1).value, ws.cell(r, 6).value
        if nombre in dict((e[0], 1) for e in EMPRESAS) and str(periodo).startswith("FY"):
            ingresos.setdefault(nombre, {})[int(str(periodo)[2:6])] = ws.cell(r, 7).value
    for nombre, datos in EXTERNOS.items():
        for año, valor in datos.items():
            ingresos[nombre].setdefault(año, valor)
    años = sorted(set.intersection(*(set(v) for v in ingresos.values())))
    return años, np.array([[ingresos[e[0]][a] for a in años] for e in EMPRESAS])


def graficar(ruta, salida):
    años, ing = leer(ruta)
    cuota = ing / ing.sum(axis=0)

    total = ing.sum(axis=0) / 1000  # MM CLP -> miles de millones (CLP bn)
    cagr = (total[-1] / total[0]) ** (1 / (len(años) - 1)) - 1
    cagr_post = (total[-1] / total[años.index(2022)]) ** (1 / (años[-1] - 2022)) - 1

    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10})
    fig = plt.figure(figsize=(12, 7), dpi=200)
    fig.patch.set_facecolor("white")
    fig.add_artist(Rectangle((0.05, 0.905), 0.008, 0.055, color=ROJO, transform=fig.transFigure))
    fig.text(0.065, 0.915, "LISTED CHILEAN MALL OPERATORS: REVENUE AND SHARE", fontsize=20,
             fontweight="bold", color=NEGRO)
    fig.text(0.05, 0.87, f"Mallplaza, Parque Arauco and Cencosud Shopping · combined revenue (CLP bn): "
             f"+{cagr:.0%}/yr since the 2020 COVID trough, +{cagr_post:.0%}/yr since 2022",
             fontsize=11, color=TINTA_SECUNDARIA)

    ax = fig.add_axes([0.08, 0.17, 0.71, 0.63])
    capas = ing / 1000
    ax.stackplot(años, capas, colors=[e[1] for e in EMPRESAS], edgecolor="white", linewidth=1.5)
    base = np.zeros(len(años))
    for (nombre, color, texto), fila, cuota_fila in zip(EMPRESAS, capas, cuota):
        for x, b, v, c in zip(años, base, fila, cuota_fila):
            dx = 0.12 if x == años[0] else -0.12 if x == años[-1] else 0
            marca = "*" if (nombre, x) in ESTIMADOS else ""
            ax.text(x + dx, b + v / 2, f"{c:.0%}{marca}", ha="center", va="center", fontsize=11,
                    fontweight="bold", color=texto)
        ax.text(años[-1] + 0.08, base[-1] + fila[-1] / 2, nombre, ha="left", va="center",
                fontsize=11, fontweight="bold", color=NEGRO)
        base = base + fila
    for i, (x, t) in enumerate(zip(años, total)):
        crec = "" if i == 0 else f"\n+{t / total[i - 1] - 1:.1%} YoY"
        dx = 0.12 if i == 0 else -0.12 if i == len(años) - 1 else 0
        ax.text(x + dx, t + total.max() * 0.03, f"{t:,.0f}{crec}", ha="center", va="bottom",
                fontsize=10.5, fontweight="bold", color=NEGRO, linespacing=1.3)

    ax.set_xlim(años[0], años[-1])
    ax.set_ylim(0, total.max() * 1.22)
    ax.set_xticks(años, [f"FY{a}" for a in años])
    ax.yaxis.set_major_formatter(lambda y, _: f"{y:,.0f}")
    ax.set_ylabel("Revenue (CLP bn, nominal)", color=TINTA_SECUNDARIA, labelpad=8)
    ax.tick_params(length=0, colors=TINTA_SECUNDARIA, labelsize=10)
    ax.tick_params(axis="x", pad=8)
    ax.tick_params(axis="y", pad=6)
    for lado in ax.spines.values():
        lado.set_visible(False)

    fig.text(0.05, 0.09, "Mallplaza FY2025 includes 12 months of the Peruvian assets acquired from "
             "Falabella (1 month in FY2024).", fontsize=8.5, color=TINTA_SECUNDARIA)
    fig.text(0.05, 0.07, "* Cencosud Shopping FY2021 estimated from its reported +33.4% revenue growth "
             "in FY2022.", fontsize=8.5, color=TINTA_SECUNDARIA)
    fig.text(0.05, 0.045, "FY2020–FY2021: company filings via Feller Rate, Humphreys, bond prospectus "
             "and press (DF); to be verified against CMF financial statements.", fontsize=8.5,
             color=TINTA_SECUNDARIA)
    fig.text(0.05, 0.02, "Source: Bloomberg (SALES_REV_TURN, consolidated revenue in CLP); Team "
             "analysis.", fontsize=8.5, color=TINTA_SECUNDARIA)
    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")
    for e, fila in zip(EMPRESAS, cuota):
        print(e[0], " ".join(f"{a}:{v:.1%}" for a, v in zip(años, fila)))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    graficar(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "cuota_mercado.png")
