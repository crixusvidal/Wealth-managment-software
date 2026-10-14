"""Tabla del costo de la deuda (Kd) de la WACC y cap rates implícitos de los comparables.

Uso:
    python graficos/tabla_kd_caprates.py <Comparables_..._valoracion.xlsx> [salida.png]
"""
import statistics as st
import sys

import matplotlib.pyplot as plt
import openpyxl
from matplotlib.lines import Line2D
from matplotlib.patches import Rectangle

from sensibilidad_wacc import NEGRO, ROJO, TINTA_SECUNDARIA

KD_REAL, INFLACION, TASA = 0.036, 0.03, 0.27   # auditoría DCF_Valuation (FFCF del modelo de valuación)
MARKET_CAP, DEUDA_BRUTA = 7_785_450, 1_659_190  # Bloomberg 30-sep-2026; DFN BBG 1.311.569 + caja 347.621
FONDO_DESTACADO, FONDO_TARGET, LINEA = "#eaf4fb", "#fce4d6", "#d9d9d9"


def pct(v, dec=2):
    return f"{v * 100:.{dec}f}%".replace(".", ",")


def mult(v):
    return f"{v:.1f}x".replace(".", ",")


def leer_peers(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    lt, nt = wb["Valoracion_LTM"], wb["Valoracion_NTM"]
    peers = []
    for r in range(7, 15):
        ev = lt.cell(r, 8).value
        peers.append((lt.cell(r, 1).value, lt.cell(r, 3).value, lt.cell(r, 2).value,
                      ev / lt.cell(r, 10).value, lt.cell(r, 10).value / ev, nt.cell(r, 10).value / ev,
                      lt.cell(r, 6).value / ev))  # LTV = deuda neta / EV de mercado
    assert all(p[3] for p in peers), "recalcular el libro antes de graficar"
    return peers


def tabla(fig, y0, cols_x, encabezados, filas, alto, alineaciones):
    for x, h, al in zip(cols_x, encabezados, alineaciones):
        fig.text(x, y0, h, ha=al, va="center", fontsize=10.5, fontweight="bold", color=NEGRO)
    fig.add_artist(Line2D([0.05, 0.95], [y0 - alto / 2] * 2, color=NEGRO, linewidth=1.2,
                          transform=fig.transFigure))
    for i, (valores, estilo) in enumerate(filas):
        y = y0 - alto * (i + 1)
        if estilo:
            fig.add_artist(Rectangle((0.05, y - alto / 2), 0.9, alto, transform=fig.transFigure,
                                     color=FONDO_TARGET if estilo == "target" else FONDO_DESTACADO,
                                     zorder=0))
        for x, v, al in zip(cols_x, valores, alineaciones):
            fig.text(x, y, v, ha=al, va="center", fontsize=10.5, color=NEGRO,
                     fontweight="bold" if estilo else "normal")
        fig.add_artist(Line2D([0.05, 0.95], [y - alto / 2] * 2, color=LINEA, linewidth=0.6,
                              transform=fig.transFigure))
    return y0 - alto * (len(filas) + 1)


def graficar(ruta, salida):
    peers = leer_peers(ruta)
    kd_nom = (1 + KD_REAL) * (1 + INFLACION) - 1
    kd_neto = kd_nom * (1 - TASA)
    dv = DEUDA_BRUTA / (DEUDA_BRUTA + MARKET_CAP)

    plt.rcParams.update({"font.family": "DejaVu Sans"})
    fig = plt.figure(figsize=(12, 11.5), dpi=200)
    fig.patch.set_facecolor("white")
    fig.add_artist(Rectangle((0.05, 0.935), 0.008, 0.038, color=ROJO, transform=fig.transFigure))
    fig.text(0.065, 0.94, "COSTO DE LA DEUDA Y CAP RATES DE COMPARABLES", fontsize=20,
             fontweight="bold", color=NEGRO)
    fig.text(0.05, 0.912, "Mallplaza · Kd usado en la WACC vs cap rate implícito (EBITDA / EV) "
             "y LTV de los peers · EBITDA como proxy de NOI", fontsize=11, color=TINTA_SECUNDARIA)

    fig.text(0.05, 0.875, "1. Costo de la deuda (Kd) en la WACC", fontsize=12.5, fontweight="bold",
             color=NEGRO)
    filas_kd = [
        (("Kd real en UF (última emisión)", pct(KD_REAL), "Modelo de valuación (FFCF)"), None),
        (("Inflación de largo plazo Chile", pct(INFLACION, 1), "Meta BCCh"), None),
        (("Kd nominal CLP", pct(kd_nom), "(1 + Kd real) × (1 + inflación) − 1"), "dest"),
        (("Tasa de impuesto", pct(TASA, 0), "Tasa legal Chile"), None),
        (("Kd después de impuestos", pct(kd_neto), "Kd nominal × (1 − t)"), "dest"),
        (("Peso de la deuda (D/V)", pct(dv, 1), "Deuda bruta / (deuda bruta + market cap)"), None),
        (("Aporte a la WACC", f"{dv * kd_neto * 100:.2f} pp".replace(".", ","), "D/V × Kd después de impuestos"), None),
    ]
    y = tabla(fig, 0.84, [0.06, 0.47, 0.56], ["Concepto", "Valor", "Cálculo / fuente"], filas_kd,
              0.034, ["left", "right", "left"])
    fig.text(0.05, y - 0.003, "Control: costo financiero neto 3,08% + reajuste UF ~3,7% ≈ 6,8% (modelo "
             "Forecast). El modelo original usaba 3,60% real, sin escudo fiscal.", fontsize=8.5,
             color=TINTA_SECUNDARIA)

    fig.text(0.05, y - 0.03, "2. Cap rate implícito (EBITDA / EV) y LTV de los comparables", fontsize=12.5,
             fontweight="bold", color=NEGRO)
    filas_cr = []
    for nombre, pais, grupo, evx, cl, cn, ltv in peers:
        g = "Objetivo" if grupo == "Target" else grupo
        nota = "*" if nombre == "IRSA" else ""
        filas_cr.append(((nombre, pais, g, mult(evx), pct(cl), pct(cn) + nota, pct(ltv, 1)),
                         "target" if grupo == "Target" else None))
    core = [p for p in peers if p[2] == "Core"]
    chile = [p for p in core if p[1] == "Chile"]
    brasil = [p for p in core if p[1] == "Brasil"]
    for nombre, grupo, f in (("Mediana core", core, st.median), ("Promedio Chile", chile, st.mean),
                             ("Mediana Brasil", brasil, st.median)):
        filas_cr.append(((nombre, "", "", mult(f(p[3] for p in grupo)), pct(f(p[4] for p in grupo)),
                          pct(f(p[5] for p in grupo)), pct(f(p[6] for p in grupo), 1)), "dest"))
    y = tabla(fig, y - 0.07, [0.06, 0.25, 0.355, 0.60, 0.715, 0.83, 0.94],
              ["Empresa", "País", "Grupo", "EV/EBITDA LTM", "Cap rate LTM", "Cap rate NTM", "LTV"],
              filas_cr, 0.034, ["left", "left", "left", "right", "right", "right", "right"])
    fig.text(0.05, y - 0.003, "LTV = deuda neta / EV de mercado (Bloomberg, 30-jun-2026). "
             "* IRSA: consenso de un solo analista; no representativo.", fontsize=8.5,
             color=TINTA_SECUNDARIA)
    fig.text(0.05, 0.02, "Fuente: Bloomberg (precios al 30-sep-2026; LTM 3T25–2T26; consenso BEst); "
             "modelo de valuación del equipo. Análisis del equipo.", fontsize=8.5, color=TINTA_SECUNDARIA)
    fig.savefig(salida, facecolor="white")
    print(f"Gráfico guardado en {salida}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    graficar(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "tabla_kd_caprates.png")
