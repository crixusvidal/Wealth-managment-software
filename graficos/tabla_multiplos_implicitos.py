"""Tabla de múltiplos y cap rates implícitos del DCF (WACC actual y nueva) vs mercado.

Uso:
    python graficos/tabla_multiplos_implicitos.py <Mallplaza_FA_Exhibits.xlsx> [salida.png]
"""
import sys

import matplotlib.pyplot as plt
import openpyxl
from matplotlib.patches import Rectangle

from sensibilidad_wacc import CELESTE, NEGRO, ROJO, TINTA_SECUNDARIA

WACC_ACTUAL, WACC_NUEVA, G = 0.1003, 0.1075, 0.03
EV_MERCADO, PRECIO, EBITDA_NTM = 9_147_859.986, 3555, 617_130.3379  # Bloomberg, 30-sep-2026
FONDO_DESTACADO = "#eaf4fb"
LINEA = "#d9d9d9"


def calcular(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    d, f = wb["DCF_Valuation"], wb["Forecast"]
    fcff = [d[c + "28"].value for c in "JKLMNO"]
    per = [d[c + "32"].value for c in "JKLMNO"]
    dfn, nci, acciones = -d["J45"].value, -d["J46"].value, d["J48"].value
    ltm, e26 = d["H15"].value, f["D104"].value

    def ev(w):
        return (sum(x / (1 + w) ** p for x, p in zip(fcff, per))
                + fcff[-1] * (1 + G) / (w - G) / (1 + w) ** 5.5)

    assert abs((ev(WACC_ACTUAL) - dfn - nci) / acciones - d["J49"].value) < 0.01

    gla_cl = sum(f[c + "41"].value or 0 for c in "DEF")
    gla_pe = sum(f[c + "42"].value or 0 for c in "DEF")
    ebitda_pipe = (gla_cl * f["D60"].value / 1e6 * f["D101"].value
                   + gla_pe * f["D70"].value / 1e6 * f["D15"].value * f["D102"].value)
    capex_pipe = f["D137"].value + f["E137"].value + f["F137"].value
    post = e26 + ebitda_pipe

    cols = []
    for evv, vpa in ((ev(WACC_ACTUAL), None), (ev(WACC_NUEVA), None), (EV_MERCADO, PRECIO)):
        vpa = vpa or (evv - dfn - nci) / acciones
        cols.append([evv, vpa, evv / ltm, evv / e26, evv / EBITDA_NTM, ltm / evv, e26 / evv,
                     post / (evv + capex_pipe)])
    return cols, ebitda_pipe, capex_pipe, gla_cl, gla_pe


def fmt(v, tipo):
    if tipo == "num":
        return f"{v:,.0f}".replace(",", ".")
    if tipo == "x":
        return f"{v:.1f}x".replace(".", ",")
    return f"{v * 100:.2f}%".replace(".", ",")


def graficar(ruta, salida):
    cols, ebitda_pipe, capex_pipe, gla_cl, gla_pe = calcular(ruta)
    filas = [("EV (MM CLP)", "num", False), ("Valor por acción (CLP)", "num", False),
             ("EV/EBITDA LTM", "x", True), ("EV/EBITDA 2026E", "x", False),
             ("EV/EBITDA NTM (consenso)", "x", False),
             ("Cap rate implícito (EBITDA LTM / EV)", "pct", True),
             ("Cap rate 2026E", "pct", False), ("Cap rate post-pipeline", "pct", True)]
    encabezados = [f"WACC actual {WACC_ACTUAL:.2%}".replace(".", ","),
                   f"WACC nueva {WACC_NUEVA:.2%}".replace(".", ","),
                   f"Mercado (precio {PRECIO:,})".replace(",", ".")]

    plt.rcParams.update({"font.family": "DejaVu Sans"})
    fig = plt.figure(figsize=(12, 6.9), dpi=200)
    fig.patch.set_facecolor("white")
    fig.add_artist(Rectangle((0.05, 0.895), 0.008, 0.06, color=ROJO, transform=fig.transFigure))
    fig.text(0.065, 0.905, "MÚLTIPLOS Y CAP RATES IMPLÍCITOS", fontsize=21, fontweight="bold",
             color=NEGRO)
    fig.text(0.05, 0.855, "Mallplaza · DCF con la WACC actual y la nueva (Ke 12%) vs valoración de "
             "mercado · EBITDA como proxy de NOI", fontsize=11, color=TINTA_SECUNDARIA)

    x_lbl, xs = 0.05, [0.53, 0.70, 0.87]
    y0, alto = 0.77, 0.072
    for x, h in zip(xs, encabezados):
        fig.text(x, y0, h, ha="center", va="center", fontsize=11, fontweight="bold",
                 color=ROJO if "nueva" in h else NEGRO)
    fig.add_artist(plt.Line2D([0.05, 0.95], [y0 - 0.035] * 2, color=NEGRO, linewidth=1.2,
                              transform=fig.transFigure))
    for i, (label, tipo, destacada) in enumerate(filas):
        y = y0 - alto * (i + 1)
        if destacada:
            fig.add_artist(Rectangle((0.05, y - alto / 2), 0.9, alto, color=FONDO_DESTACADO,
                                     transform=fig.transFigure, zorder=0))
        fig.text(x_lbl + 0.01, y, label, va="center", fontsize=11.5, color=NEGRO,
                 fontweight="bold" if destacada else "normal")
        for x, col in zip(xs, cols):
            fig.text(x, y, fmt(col[i], tipo), ha="center", va="center", fontsize=12,
                     fontweight="bold" if destacada else "normal",
                     color=TINTA_SECUNDARIA if col is cols[2] and not destacada else NEGRO)
        fig.add_artist(plt.Line2D([0.05, 0.95], [y - alto / 2] * 2, color=LINEA, linewidth=0.6,
                                  transform=fig.transFigure))
    fig.add_artist(Rectangle((0.615, y - alto / 2), 0.17, y0 - 0.035 - (y - alto / 2), fill=False,
                             edgecolor=CELESTE, linewidth=1.5, transform=fig.transFigure))

    n = lambda v: fmt(v, "num")
    yoc = f"{ebitda_pipe / capex_pipe * 100:.1f}".replace(".", ",")
    nota = (f"Cap rate post-pipeline = (EBITDA 2026E + EBITDA del pipeline maduro, {n(ebitda_pipe)} MM CLP) / "
            f"(EV + capex remanente del pipeline, {n(capex_pipe)} MM CLP). Pipeline: {n(gla_cl + gla_pe)} m² "
            f"(Chile {n(gla_cl)}; Perú {n(gla_pe)}); yield on cost {yoc}%.")
    fig.text(0.05, 0.075, nota, fontsize=8.5, color=TINTA_SECUNDARIA, wrap=True)
    fig.text(0.05, 0.035, "Fuente: modelo DCF del equipo (g = 3%); Bloomberg (precio y consenso al "
             "30-sep-2026); release 2T26. Análisis del equipo.", fontsize=8.5, color=TINTA_SECUNDARIA)
    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    graficar(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "tabla_multiplos_implicitos.png")
