"""Gráficos de la sección Industry del informe (datos de Excel_DATA_INDUSTRY.xlsx).

Listos para pegar en el informe: título con período, subtítulo, notas y fuente dentro de la
imagen (mismo formato que cuota_mercado.png). Colores: Mallplaza rojo, Parque Arauco verde,
Cencosud morado; el resto con la paleta corporativa de Mallplaza.

Uso:
    python graficos/industria.py <Excel_DATA_INDUSTRY.xlsx> [carpeta_salida]
"""
import math
import sys
import textwrap
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.colors import to_rgb
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from matplotlib.patches import Patch, Rectangle

from cuota_mercado import MORADO_CENCOSUD, VERDE_PARQUE_ARAUCO
from sensibilidad_wacc import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA

COLOR = {"Mallplaza": ROJO, "Parque Arauco": VERDE_PARQUE_ARAUCO, "Cenco Malls": MORADO_CENCOSUD}
TEXTO = {"Mallplaza": "white", "Parque Arauco": NEGRO, "Cenco Malls": "white"}
GRIS_OSCURO, GRIS_CLARO = "#8c8c8c", "#d9d9d9"  # solo grilla y líneas auxiliares
DPI = 300


def tinte(color, fuerza):
    return tuple(1 - fuerza * (1 - c) for c in to_rgb(color))


def limpiar(ax, grilla="y"):
    for lado in ax.spines.values():
        lado.set_visible(False)
    ax.tick_params(length=0, colors=TINTA_SECUNDARIA, labelsize=9)
    if grilla:
        ax.grid(axis=grilla, color=GRILLA, linewidth=0.8)
        ax.set_axisbelow(True)


def guardar(fig, salida):
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)
    print(f"Gráfico guardado en {salida}")


CABECERA, PIE_BASE, PIE_LINEA = 0.82, 0.1, 0.15  # pulgadas
LOGOS = Path(__file__).parent / "logos"
ARCHIVOS_LOGO = {"Mallplaza": "mallplaza.png", "Parque Arauco": "parque_arauco.png",
                 "Cenco Malls": "cencosud_shopping.png"}
ANCHO_LOGO = {"Mallplaza": 0.42, "Parque Arauco": 0.95, "Cenco Malls": 1.0}  # pulgadas


def figura(ancho, alto, titulo, subtitulo, pie):
    """Figura con cabecera (barra roja, título, subtítulo) y pie (notas + fuente).

    `alto` es el alto del área del gráfico; caja() y yc() traducen coordenadas 0-1 de esa área.
    """
    # Notas largas en varias líneas (letra 7 pt: ~19 caracteres por pulgada)
    pie = [l for linea in pie for l in textwrap.wrap(linea, int(ancho * 0.93 * 19))]
    h = alto + CABECERA + PIE_BASE + len(pie) * PIE_LINEA
    fig = plt.figure(figsize=(ancho, h))
    fig.patch.set_facecolor("white")
    fig.area = ((PIE_BASE + len(pie) * PIE_LINEA + 0.12) / h, (alto - 0.12) / h)
    fig.add_artist(Rectangle((0.035, 1 - 0.47 / h), 0.06 / ancho, 0.27 / h, color=ROJO,
                             transform=fig.transFigure))
    fig.text(0.035 + 0.11 / ancho, 1 - 0.45 / h, titulo, fontsize=13.5, fontweight="bold",
             color=NEGRO, va="bottom")
    fig.text(0.035, 1 - 0.7 / h, subtitulo, fontsize=8.5, color=TINTA_SECUNDARIA, va="bottom")
    textos = [fig.text(0.035, (PIE_BASE + i * PIE_LINEA) / h, linea, fontsize=7,
                       color=TINTA_SECUNDARIA, va="bottom") for i, linea in enumerate(reversed(pie))]
    ajustar(fig, fig.texts[0:2] + textos)
    return fig


def ajustar(fig, textos, margen=0.975):
    """Achica la letra de cada texto hasta que quepa en el ancho de la figura."""
    render = fig.canvas.get_renderer()
    for t in textos:
        while t.get_window_extent(render).x1 > margen * fig.bbox.width and t.get_fontsize() > 5:
            t.set_fontsize(t.get_fontsize() - 0.25)


def caja(fig, rect):
    x, y, w, alto = rect
    y0, hc = fig.area
    return [x, y0 + y * hc, w, alto * hc]


def yc(fig, y):
    y0, hc = fig.area
    return y0 + y * hc


def miles(x):
    return f"{x:,.0f}"


# 1. Ocupación consolidada -------------------------------------------------------------------------
def ocupacion(wb, salida):
    ws = wb["Ocupacion_consolidada"]
    periodos = [ws.cell(5, c).value for c in range(2, 7)]
    fig = figura(7, 3.9, f"CONSOLIDATED OCCUPANCY RATE, 2Q{periodos[0][-2:]}–2Q{periodos[-1][-2:]}",
                 "Mallplaza, Parque Arauco and Cenco Malls · occupancy at the end of each 2Q "
                 "(Chile, Peru and Colombia)",
                 ["Each company defines occupancy differently: compare trends rather than levels. "
                  "Cenco 2Q26 includes Plaza Central (Bogotá, acquired Jun-26).",
                  "Source: Mallplaza, Parque Arauco and Cenco Malls 2Q22–2Q26 earnings reports; "
                  "Team analysis."])
    ax = fig.add_axes(caja(fig, [0.08, 0.08, 0.66, 0.9]))
    for r in range(6, 9):
        nombre = ws.cell(r, 1).value
        serie = [ws.cell(r, c).value * 100 for c in range(2, 7)]
        ax.plot(range(5), serie, color=COLOR[nombre], linewidth=2.2, marker="o", markersize=5,
                markeredgecolor="white", markeredgewidth=1, zorder=3)
        ax.text(4.15, serie[-1], f"{nombre} {serie[-1]:.1f}%", va="center", fontsize=9.5,
                fontweight="bold", color=COLOR[nombre])
    ax.axvline(3, color=GRIS_OSCURO, linestyle="--", linewidth=0.9)
    ax.text(3.05, 90.3, "Mallplaza includes\nOpen Plaza from 2Q25", fontsize=7.5,
            color=TINTA_SECUNDARIA, va="bottom")
    ax.set_xticks(range(5), [f"2Q{p[-2:]}" for p in periodos], fontsize=10, color=NEGRO)
    ax.set_ylim(90, 100)
    ax.set_xlim(-0.2, 4.1)
    ax.yaxis.set_major_formatter(lambda y, _: f"{y:.0f}%")
    limpiar(ax)
    guardar(fig, salida)


# 2. Pipeline de GLA ------------------------------------------------------------------------------
def pipeline(wb, salida):
    ws = wb["PIPELINE"]
    datos = [("Mallplaza", ws["D12"].value, ws["E12"].value),
             ("Parque Arauco", ws["N12"].value, ws["O12"].value),
             ("Cenco Malls", ws["I12"].value, 0)]
    fig = figura(7, 3.9, "GLA PIPELINE 2026–2028: ORGANIC VS ACQUISITIONS",
                 "GLA to be added by Mallplaza, Parque Arauco and Cenco Malls (m²) · dated projects only",
                 ["Mallplaza 2S26 inorganic GLA not disclosed (taken as zero).",
                  "Source: Plaza S.A. Release 2Q26 and 2025 Annual Report; Parque Arauco Investor "
                  "Presentation 2Q26; Cenco Malls Corporate Presentation 2Q26; Team analysis."])
    ax = fig.add_axes(caja(fig, [0.11, 0.08, 0.86, 0.9]))
    for x, (nombre, org, inorg) in enumerate(datos):
        ax.bar(x, org, width=0.55, color=COLOR[nombre], linewidth=0)
        ax.text(x, org / 2, miles(org), ha="center", va="center", fontsize=9.5,
                color=TEXTO[nombre])
        if inorg:
            ax.bar(x, inorg, bottom=org, width=0.55, color=tinte(COLOR[nombre], 0.35),
                   edgecolor="white", linewidth=0, hatch="///")
            ax.plot([x - 0.275, x + 0.275], [org, org], color="white", linewidth=1.5)
            ax.text(x, org + inorg / 2, miles(inorg), ha="center", va="center", fontsize=9.5,
                    color=NEGRO, bbox={"facecolor": "white", "edgecolor": "none", "pad": 1.5})
        ax.text(x, org + inorg + 3000, f"{miles(org + inorg)} m²", ha="center", va="bottom",
                fontsize=10.5, fontweight="bold", color=NEGRO)
    ax.set_xticks(range(3), [d[0] for d in datos], fontsize=10, color=NEGRO)
    ax.set_ylim(0, 200_000)
    ax.yaxis.set_major_formatter(lambda y, _: miles(y))
    ax.set_ylabel("GLA (m²)", color=TINTA_SECUNDARIA)
    limpiar(ax)
    ax.legend(handles=[Patch(facecolor=TINTA_SECUNDARIA, label="Organic"),
                       Patch(facecolor=tinte(TINTA_SECUNDARIA, 0.35), hatch="///", edgecolor="white",
                             label="Acquisitions")],
              loc="upper right", frameon=False, fontsize=9)
    guardar(fig, salida)


# 3. Oferta de GLA vs consumo privado -----------------------------------------------------------
def oferta_demanda(wb, salida):
    ws = wb["OFERTA VS DEMANDA ESPERADA"]
    filas = [(ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value) for r in range(2, 6)]
    assert all(isinstance(f[1], float) for f in filas), "abrir y guardar el Excel en Excel (vínculo externo)"
    fig = figura(7, 3.9, "NEW GLA SUPPLY VS PRIVATE CONSUMPTION, CHILE 2026–2028",
                 "Annual growth of the trio's organic GLA in Chile vs expected real private "
                 "consumption growth",
                 ["Organic GLA only (consumption-dependent supply).",
                  "Source: 2Q26 reports of Plaza S.A., Parque Arauco and Cenco Malls; Central Bank "
                  "of Chile, IPoM Sep-2026 (Table II.1); Team analysis."])
    ax = fig.add_axes(caja(fig, [0.08, 0.12, 0.9, 0.86]))
    xs = [0, 1, 2, 3.4]
    for x, (anio, oferta, consumo) in zip(xs, filas):
        for dx, v, color in ((-0.19, oferta, CELESTE), (0.19, consumo, AMARILLO)):
            ax.bar(x + dx, v * 100, width=0.36, color=color)
            ax.text(x + dx, v * 100 + 0.1, f"{v * 100:.1f}%", ha="center", va="bottom", fontsize=9,
                    color=NEGRO)
    ax.axvline(2.7, color=GRIS_CLARO, linewidth=1)
    ax.set_xticks(xs, [str(f[0]).replace("Acumulado", "Cumulative\n").replace(" 2026", "2026")
                       for f in filas], fontsize=10, color=NEGRO)
    ax.set_ylim(0, 7.8)
    ax.yaxis.set_major_formatter(lambda y, _: f"{y:.0f}%")
    limpiar(ax)
    ax.legend(handles=[Patch(color=CELESTE, label="New GLA supply (trio, Chile)"),
                       Patch(color=AMARILLO, label="Real private consumption (IPoM)")],
              loc="upper left", frameon=False, fontsize=9)
    guardar(fig, salida)


# 4. Vencimiento de contratos -------------------------------------------------------------------
def contratos(wb, salida):
    ws = wb["Duración contratos"]
    datos = {}
    for r in range(2, ws.max_row + 1):
        emp, tramo, pct, dur = (ws.cell(r, c).value for c in range(1, 5))
        datos.setdefault(emp, {"tramos": [], "dur": dur})["tramos"].append((tramo, pct))
    empresas = sorted(datos, key=lambda e: datos[e]["dur"])  # mayor duración arriba
    # Vence antes = más riesgo: rojo → amarillo → celeste (paleta Mallplaza)
    rampa = [ROJO, tinte(ROJO, 0.45), AMARILLO, tinte(CELESTE, 0.55), CELESTE]
    fig = figura(7, 3.0, "LEASE MATURITY PROFILE AND AVERAGE REMAINING TERM, 2Q26",
                 "Share of GLA by remaining lease term · Cenco Malls, Mallplaza and Parque Arauco",
                 ["Consolidated figures; the weighting of the average term may differ across companies.",
                  "Source: Cenco Malls 2Q26 Results Presentation (p.14); Mallplaza Release 2Q26 (p.7); "
                  "Parque Arauco 2Q26 Report (p.7) and Investor Presentation 2Q26 (slide 17)."])
    ax = fig.add_axes(caja(fig, [0.17, 0.08, 0.66, 0.74]))
    for y, emp in enumerate(empresas):
        izq = 0
        for i, (tramo, pct) in enumerate(datos[emp]["tramos"]):
            ax.barh(y, pct * 100, left=izq, height=0.6, color=rampa[i], edgecolor="white",
                    linewidth=1.2)
            if pct >= 0.05:
                ax.text(izq + pct * 50, y, f"{pct * 100:.1f}%", ha="center", va="center",
                        fontsize=8, color="white" if i == 0 else NEGRO)
            izq += pct * 100
        ax.text(103, y, f"{datos[emp]['dur']:.1f} yrs", va="center", fontsize=11,
                fontweight="bold", color=NEGRO)
        ax.text(-2, y, emp, ha="right", va="center", fontsize=10, color=NEGRO)
    ax.text(103, len(empresas) - 0.45, "Avg. remaining\nterm", fontsize=7.5, color=TINTA_SECUNDARIA,
            va="bottom")
    ax.set_xlim(0, 100)
    ax.set_yticks([])
    ax.set_xticks([0, 25, 50, 75, 100], ["0%", "25%", "50%", "75%", "100%"])
    limpiar(ax, grilla=None)
    etiquetas = [t.replace("años", "yrs") for t, _ in datos[empresas[0]]["tramos"]]
    fig.legend(handles=[Patch(color=c, label=t) for c, t in zip(rampa, etiquetas)],
               loc="upper left", bbox_to_anchor=(0.16, yc(fig, 1.0)), ncol=5, frameon=False, fontsize=8.5,
               handlelength=1.4, columnspacing=1.2)
    guardar(fig, salida)


# 5. Market share estático por país -------------------------------------------------------------
def market_share_paises(wb, salida):
    ws = wb["MS_estatico_2Q26"]
    otros = tinte(CELESTE, 0.35)
    roles = {"Líder fuera del trío": ("Leader outside the trio", AMARILLO),
             "Otros": ("Others", otros)}
    paises = {}
    for r in range(2, ws.max_row + 1):
        pais, emp, rol = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 4).value
        paises.setdefault(pais, []).append((emp, rol, ws.cell(r, 5).value, ws.cell(r, 7).value))
    for pais, filas in paises.items():  # "Otros" = 1 − resto (fórmula en el Excel)
        resto = 1 - sum(f[2] for f in filas if f[1] != "Otros")
        paises[pais] = [(e, ro, resto if ro == "Otros" else v, n) for e, ro, v, n in filas]

    fig = figura(9, 3.6, "GLA MARKET SHARE BY COUNTRY, 2Q26",
                 "Share of GLA among the main shopping center operators of each country",
                 ["* Pro forma, includes Gran Plaza (4.3% without it). Trio share = Mallplaza + "
                  "Cenco Malls + Parque Arauco.",
                  "Source: Plaza S.A., Institutional Presentation 2Q26 (slide 7); Team analysis."])
    nombres_pais = {"Chile": "Chile", "Perú": "Peru", "Colombia": "Colombia"}
    for k, (pais, filas) in enumerate(paises.items()):
        ax = fig.add_axes(caja(fig, [0.02 + k * 0.33, 0.0, 0.29, 0.86]))
        valores = [f[2] for f in filas]
        colores = [COLOR.get(f[0], roles.get(f[1], (None, otros))[1]) for f in filas]
        cuñas, _ = ax.pie(valores, colors=colores, startangle=90, counterclock=False,
                          wedgeprops={"width": 0.38, "edgecolor": "white", "linewidth": 1.5})
        for cuña, (emp, rol, v, nota) in zip(cuñas, filas):
            ang = math.radians((cuña.theta1 + cuña.theta2) / 2)
            marca = "*" if nota and "pro forma" in nota else ""
            fuera = v < 0.05
            radio = 1.22 if fuera else 0.81
            claro = colores[filas.index((emp, rol, v, nota))] in (otros, VERDE_PARQUE_ARAUCO, AMARILLO)
            ax.text(radio * math.cos(ang), radio * math.sin(ang), f"{v:.0%}{marca}",
                    ha="center", va="center", fontsize=9, fontweight="bold",
                    color=NEGRO if fuera or claro else "white")
        ax.text(0, 0, nombres_pais[pais], ha="center", va="center", fontsize=12,
                fontweight="bold", color=NEGRO)
        trio = sum(f[2] for f in filas if f[0] in COLOR)
        ax.text(0, -1.42, f"Trio share: {trio:.0%}", ha="center", va="center", fontsize=10,
                color=NEGRO)
        ax.set_xlim(-1.3, 1.3)
        ax.set_ylim(-1.55, 1.3)
    fig.legend(handles=[Patch(color=COLOR["Mallplaza"], label="Mallplaza"),
                        Patch(color=COLOR["Cenco Malls"], label="Cenco Malls"),
                        Patch(color=COLOR["Parque Arauco"], label="Parque Arauco"),
                        Patch(color=AMARILLO, label="Leader outside the trio"),
                        Patch(color=otros, label="Others")],
               loc="upper center", bbox_to_anchor=(0.5, yc(fig, 1.02)), ncol=5, frameon=False,
               fontsize=9)
    guardar(fig, salida)


# 6. Cinco fuerzas de Porter --------------------------------------------------------------------
def porter(wb, salida):
    ws = wb["PORTER"]
    # Promedio simple de Chile, Perú y Colombia (columnas C:E), redondeado al entero más cercano
    # con los ,5 hacia arriba (2,4 → 2; 2,5 → 3)
    notas = {ws.cell(r, 1).value: sum(ws.cell(r, c).value for c in (3, 4, 5)) / 3
             for r in range(2, ws.max_row + 1)}
    orden = [("Poder de negociación de los locatarios", "Bargaining power\nof tenants"),
             ("Poder de negociación de los proveedores", "Bargaining power\nof suppliers"),
             ("Amenaza de sustitutos (e-commerce)", "Threat of substitutes\n(e-commerce)"),
             ("Amenaza de nuevos entrantes", "Threat of\nnew entrants"),
             ("Rivalidad entre competidores", "Rivalry among\ncompetitors")]
    valores = [math.floor(notas[k] + 0.5 + 1e-9) for k, _ in orden]
    for (k, _), v in zip(orden, valores):
        print(f"  Porter · {k}: promedio {notas[k]:.2f} → {v}")
    n = len(orden)
    ang = np.linspace(0, 2 * np.pi, n, endpoint=False)
    fig = figura(6.2, 4.7, "PORTER'S FIVE FORCES, 2Q26",
                 "Shopping centers in Chile, Peru and Colombia · 1 = weak force (attractive "
                 "industry), 5 = strong force",
                 ["Score = average of Chile, Peru and Colombia, rounded to the nearest integer (.5 up).",
                  "Source: Team analysis (Porter_industria_2Q26)."])
    ax = fig.add_axes(caja(fig, [0.2, 0.1, 0.6, 0.7]), polar=True)
    ax.set_theta_zero_location("N")
    ax.set_theta_direction(-1)
    for nivel in range(1, 6):  # grilla pentagonal
        ax.plot(np.append(ang, ang[0]), [nivel] * (n + 1), color=GRIS_CLARO, linewidth=0.8)
    for a in ang:
        ax.plot([a, a], [0, 5], color=GRIS_CLARO, linewidth=0.8)
    cerrado = np.append(valores, valores[0])
    ax.fill(np.append(ang, ang[0]), cerrado, color=ROJO, alpha=0.15)
    ax.plot(np.append(ang, ang[0]), cerrado, color=ROJO, linewidth=2.4)
    ax.scatter(ang, valores, s=45, color=ROJO, edgecolor="white", linewidth=1.2, zorder=3)
    for a, v in zip(ang, valores):
        ax.text(a, v + 0.55, f"{v}", ha="center", va="center", fontsize=12, fontweight="bold",
                color=NEGRO)
    for nivel in range(1, 6):
        ax.text(np.radians(36), nivel, str(nivel), fontsize=7.5, color=GRIS_OSCURO,
                ha="center", va="center", bbox={"facecolor": "white", "edgecolor": "none", "pad": 0.5})
    ax.set_xticks(ang, [e for _, e in orden], fontsize=10, color=NEGRO)
    ax.tick_params(axis="x", pad=22)
    ax.set_yticks([])
    ax.set_ylim(0, 5)
    ax.grid(False)
    ax.spines["polar"].set_visible(False)
    guardar(fig, salida)


# 7. Market share evolutivo Chile + crecimiento de la industria (base 100) ----------------------
def market_share_chile(wb, salida):
    """Área apilada de GLA en Chile (mismo formato que cuota_mercado.png) con índice 2Q22 = 100."""
    ws = wb["MS_evolutivo_Chile"]
    periodos = [f"2Q{ws.cell(5, c).value[-2:]}" for c in range(2, 7)]
    filas = {ws.cell(r, 1).value: [ws.cell(r, c).value for c in range(2, 7)] for r in (6, 7, 8)}
    empresas = ["Mallplaza", "Parque Arauco", "Cenco Malls"]  # mismo orden que cuota_mercado.png
    gla = np.array([filas[e] for e in empresas], dtype=float) / 1000  # miles de m²
    total = gla.sum(axis=0)
    indice = total / total[0] * 100
    cuota = gla / total
    x = np.arange(len(periodos))
    cagr = (total[-1] / total[0]) ** (1 / (len(x) - 1)) - 1

    fig = figura(7.4, 4.1, f"CHILEAN GLA OF LISTED MALL OPERATORS, {periodos[0]}–{periodos[-1]}",
                 f"Mallplaza, Parque Arauco and Cenco Malls · GLA in Chile (thousand m²) and share "
                 f"of the trio · index {periodos[0]} = 100 · +{cagr:.1%}/yr",
                 ["Parque Arauco 2Q25 includes the Kennedy Oriente acquisition (70,000 m²). Cenco "
                  "GLA includes Costanera office space.",
                  "Share within the trio, not of the total market.",
                  "Source: Mallplaza, Cenco Malls and Parque Arauco 2Q22–2Q26 earnings reports; "
                  "Team analysis."])
    ax = fig.add_axes(caja(fig, [0.09, 0.08, 0.72, 0.9]))
    ax.stackplot(x, gla, colors=[COLOR[e] for e in empresas], edgecolor="white", linewidth=1.5)
    base = np.zeros(len(x))
    for emp, capa, c in zip(empresas, gla, cuota):
        for i in range(len(x)):
            dx = 0.2 if i == 0 else -0.2 if i == len(x) - 1 else 0
            ax.text(x[i] + dx, base[i] + capa[i] / 2, f"{c[i]:.1%}", ha="center", va="center",
                    fontsize=9, fontweight="bold", color=TEXTO[emp])
        centro = (len(x) - 1, base[-1] + capa[-1] / 2)
        logo = LOGOS / ARCHIVOS_LOGO[emp]
        if logo.exists():
            img = plt.imread(logo)
            caja_logo = OffsetImage(img, zoom=ANCHO_LOGO[emp] * 72 / img.shape[1])
            ax.add_artist(AnnotationBbox(caja_logo, centro, xybox=(12, 0), xycoords="data",
                                         boxcoords="offset points", box_alignment=(0, 0.5),
                                         frameon=False, annotation_clip=False))
        else:
            ax.text(centro[0] + 0.08, centro[1], emp, ha="left", va="center", fontsize=9.5,
                    fontweight="bold", color=NEGRO)
        base = base + capa
    for i, (t, ind) in enumerate(zip(total, indice)):
        dx = 0.2 if i == 0 else -0.2 if i == len(x) - 1 else 0
        detalle = "" if i == 0 else f"\n{t / total[i - 1] - 1:+.1%} YoY"
        ax.text(x[i] + dx, t + total.max() * 0.025, f"{t:,.0f}{detalle}\nIndex {ind:.1f}",
                ha="center", va="bottom", fontsize=8.5, fontweight="bold", color=NEGRO,
                linespacing=1.3)
    ax.set_xlim(x[0], x[-1])
    ax.set_ylim(0, total.max() * 1.25)
    ax.set_xticks(x, periodos, fontsize=9.5)
    ax.yaxis.set_major_formatter(lambda y, _: f"{y:,.0f}")
    ax.set_ylabel("GLA in Chile (thousand m²)", color=TINTA_SECUNDARIA, labelpad=6)
    limpiar(ax, grilla=None)
    guardar(fig, salida)


# 8. Eficiencia operativa: EBITDA por m² -------------------------------------------------------
def eficiencia(wb, salida):
    ws = wb["EF. OPERATIVA (CLPM2)"]
    filas = [(ws.cell(r, 1).value, ws.cell(r, 2).value * 1e6 / ws.cell(r, 3).value * 4)
             for r in range(6, 9)]
    filas.sort(key=lambda f: f[1])
    fig = figura(7, 2.6, "OPERATING EFFICIENCY: EBITDA PER M² OF GLA, 2Q26",
                 "2Q26 EBITDA annualized (×4) / GLA at 30-Jun-2026 · CLP",
                 ["Source: Mallplaza, Cenco Malls and Parque Arauco 2Q26 earnings reports; "
                  "Team analysis."])
    ax = fig.add_axes(caja(fig, [0.17, 0.16, 0.76, 0.82]))
    for y, (emp, v) in enumerate(filas):
        ax.barh(y, v / 1000, height=0.58, color=COLOR[emp])
        ax.text(v / 1000 + 3, y, f"{v / 1000:,.0f}k", va="center", fontsize=10.5,
                fontweight="bold", color=NEGRO)
    ax.set_yticks(range(len(filas)), [f[0] for f in filas], fontsize=10, color=NEGRO)
    ax.set_xlim(0, 270)
    ax.xaxis.set_major_formatter(lambda v, _: f"{v:,.0f}k")
    ax.set_xlabel("EBITDA per m² of GLA, 2Q26 annualized (CLP)", color=TINTA_SECUNDARIA)
    limpiar(ax, grilla="x")
    guardar(fig, salida)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    salida = Path(sys.argv[2] if len(sys.argv) > 2 else Path(__file__).parent / "industria")
    salida.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
    ocupacion(wb, salida / "ind1_ocupacion.png")
    pipeline(wb, salida / "ind2_pipeline_gla.png")
    oferta_demanda(wb, salida / "ind3_oferta_vs_consumo.png")
    contratos(wb, salida / "ind4_vencimiento_contratos.png")
    market_share_paises(wb, salida / "ind5_market_share_paises.png")
    porter(wb, salida / "ind6_porter.png")
    market_share_chile(wb, salida / "ind7_market_share_chile.png")
    eficiencia(wb, salida / "ind8_ebitda_m2.png")
