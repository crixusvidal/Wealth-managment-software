"""Figuras de la sección Valuación y del anexo de la WACC, en español e inglés.

Lee Mallplaza_Figuras_Valuacion.xlsx (cada hoja trae los datos y la especificación del gráfico).
V7 (múltiplos vs tasa local) queda fuera: falta la NTN-B de Brasil (pendiente de Bloomberg);
V8 es el gráfico del equipo. Sin título ni fuente en la imagen: van en el Word (informe/valuacion.js).

Uso:
    python graficos/valuacion.py <Mallplaza_Figuras_Valuacion.xlsx> [carpeta_salida | -] [Valoracion_Consolidada_v2.xlsx]
    (con el modelo v2 también genera V9: valor implícito por múltiplo de cada par)
"""
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.patches import Patch, Rectangle

from pagina_riesgos import num, pct, tinte
from sensibilidad_wacc import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA

AZUL = "#14284B"
GRIS = "#9a9a9a"
CLARO = "#eef1f6"
ANCHO = 3550 / 1440          # columna de figuras del informe (pulgadas)
COMPLETO = 10666 / 1440      # ancho completo (anexo)
DPI = 400
ALTOS = {"V7": 2.25, "V9": 1.95, "V1": 1.85, "V2": 1.7, "V3": 2.05, "V4": 1.8, "V5": 1.6, "V6": 1.75, "W1": 1.9}

T = {
    "es": {"cpa": "CLP por acción", "objetivo": "Precio objetivo", "precio": "Precio actual",
           "v1": {"DCF FCFF (WACC ±50 pb, g 3,0%)": ("DCF FCFF", "±50 pb · g 3,0%"),
                  "DCF FCFF (WACC ±100 pb, g 2,5–4,0%)": ("DCF FCFF", "±100 pb · g 2,5–4,0%"),
                  "APV (chequeo)": ("APV", "chequeo"),
                  "Múltiplos EV/EBITDA (100% set común – 100% Chile)": ("EV/EBITDA LTM", "set común – Chile"),
                  "Múltiplos EV/EBITDA NTM / EV/EBIT LTM": ("NTM · EV/EBIT", "múltiplos alternativos"),
                  "Rango 52 semanas (intradía, 01-10-25 a 30-09-26)": ("Precio 52 semanas", "oct-25 – sep-26")},
           "real": "Tasa real (UF)",
           "v7": {"a": "EV/EBITDA LTM (x) vs bono real local a 10 años", "b": "Spread real: cap rate − bono real local (pb)",
                  "cl": "Chile", "br": "Brasil", "med": "Mediana", "mp": "Mallplaza"},
           "v2": {"Bono en UF 10 años BCCh (real)": "Bono UF 10a (BCCh)",
                  "Bono en pesos 10 años BCCh (nominal)": "Bono en pesos 10a (BCCh)",
                  "Base del modelo: Rf UST + CRP Chile, convertida a CLP": "Rf del modelo (UST + CRP)",
                  "Kd Mallplaza (cupón última emisión UF)": "Kd Mallplaza (emisión UF)",
                  "WACC implícita en el precio actual (DCF inverso)": "WACC implícita en el precio",
                  "WACC real de la tasación (Nota 11, Chile jun-26)": "WACC de la tasación (Nota 11)",
                  "Nuestra WACC terminal": "Nuestra WACC terminal", "Nuestro Ke terminal": "Nuestro Ke terminal"},
           "pb": "pb",
           "v3": {"cols": ["g nominal", "g real", "Reinversión", "Valor DCF", "Δ"],
                  "series": ["RONIC 5,0% (Mallplaza)", "RONIC = WACC (7,2%)", "RONIC 10,0%"],
                  "eje": "g real", "indice": "Valor terminal (g = 0 → 100)"},
           "v4": {"comp": ["Deudores", "CxC\nrelac.", "CxP", "CxP\nrelac.", "Arriendos\nanticip.", "CTN"],
                  "cols": ["Escenario", "CTN/ingresos", "Precio objetivo", "Δ"],
                  "esc": ["Sin capital de trabajo", "Mitad", "Base (dic-25)", "+50%", "Doble"],
                  "mm": "MM CLP, dic-25"},
           "v5": {"Mercado: Mallplaza al precio actual": "Mercado (precio actual)",
                  "Pares chilenos (1 / mediana EV/EBITDA Chile)": "Pares chilenos",
                  "Valor libro (EBITDA LTM / propiedades de inversión)": "Valor libro (PI)",
                  "Compra Gran Plaza, Colombia (NOI / precio) — argumento aislado": "Gran Plaza (hecho posterior)",
                  "Nuestro DCF (EBITDA LTM / EV DCF)": "Nuestro DCF",
                  "Nuestro DCF, valor terminal (EBITDA 2032E / VT)": "Nuestro DCF, valor terminal"},
           "brecha": "brecha", "cap": "Cap rate (EBITDA / EV)",
           "v6": {"dcf": "Valor DCF", "po": "Precio objetivo 80/20", "eje": "WACC terminal", "base": "base",
                  "impl": "implícita"},
           "w1": {"wacc": "WACC nominal", "de": "D/E (eje der.)", "per": ["2S26E", "2027E", "2028E", "2029E",
                                                                           "2030E", "2031E", "Terminal"]}},
    "en": {"cpa": "CLP per share", "objetivo": "Target price", "precio": "Current price",
           "v1": {"DCF FCFF (WACC ±50 pb, g 3,0%)": ("DCF FCFF", "±50 bp · g 3.0%"),
                  "DCF FCFF (WACC ±100 pb, g 2,5–4,0%)": ("DCF FCFF", "±100 bp · g 2.5–4.0%"),
                  "APV (chequeo)": ("APV", "cross-check"),
                  "Múltiplos EV/EBITDA (100% set común – 100% Chile)": ("EV/EBITDA LTM", "peer set – Chile"),
                  "Múltiplos EV/EBITDA NTM / EV/EBIT LTM": ("NTM · EV/EBIT", "alternative multiples"),
                  "Rango 52 semanas (intradía, 01-10-25 a 30-09-26)": ("52-week price", "Oct-25 – Sep-26")},
           "real": "Real rate (UF)",
           "v7": {"a": "EV/EBITDA LTM (x) vs local 10Y real bond", "b": "Real spread: cap rate − local real bond (bp)",
                  "cl": "Chile", "br": "Brazil", "med": "Median", "mp": "Mallplaza"},
           "v2": {"Bono en UF 10 años BCCh (real)": "10Y UF bond (BCCh)",
                  "Bono en pesos 10 años BCCh (nominal)": "10Y peso bond (BCCh)",
                  "Base del modelo: Rf UST + CRP Chile, convertida a CLP": "Model Rf (UST + CRP)",
                  "Kd Mallplaza (cupón última emisión UF)": "Mallplaza Kd (UF issue)",
                  "WACC implícita en el precio actual (DCF inverso)": "WACC implied by the price",
                  "WACC real de la tasación (Nota 11, Chile jun-26)": "Appraisal WACC (Note 11)",
                  "Nuestra WACC terminal": "Our terminal WACC", "Nuestro Ke terminal": "Our terminal Ke"},
           "pb": "bp",
           "v3": {"cols": ["Nominal g", "Real g", "Reinvestment", "DCF value", "Δ"],
                  "series": ["RONIC 5.0% (Mallplaza)", "RONIC = WACC (7.2%)", "RONIC 10.0%"],
                  "eje": "Real g", "indice": "Terminal value (g = 0 → 100)"},
           "v4": {"comp": ["Receiv-\nables", "Rel.-party\nreceiv.", "Payables", "Rel.-party\npayables", "Advance\nrents", "NWC"],
                  "cols": ["Scenario", "NWC/revenue", "Target price", "Δ"],
                  "esc": ["No working capital", "Half", "Base (Dec-25)", "+50%", "Double"],
                  "mm": "CLP MM, Dec-25"},
           "v5": {"Mercado: Mallplaza al precio actual": "Market (current price)",
                  "Pares chilenos (1 / mediana EV/EBITDA Chile)": "Chilean peers",
                  "Valor libro (EBITDA LTM / propiedades de inversión)": "Book value (inv. property)",
                  "Compra Gran Plaza, Colombia (NOI / precio) — argumento aislado": "Gran Plaza (subsequent event)",
                  "Nuestro DCF (EBITDA LTM / EV DCF)": "Our DCF",
                  "Nuestro DCF, valor terminal (EBITDA 2032E / VT)": "Our DCF, terminal value"},
           "brecha": "gap", "cap": "Cap rate (EBITDA / EV)",
           "v6": {"dcf": "DCF value", "po": "80/20 target price", "eje": "Terminal WACC", "base": "base",
                  "impl": "implied"},
           "w1": {"wacc": "Nominal WACC", "de": "D/E (right axis)", "per": ["2H26E", "2027E", "2028E", "2029E",
                                                                             "2030E", "2031E", "Terminal"]}},
}


def figura(clave, ancho=ANCHO):
    fig = plt.figure(figsize=(ancho, ALTOS[clave]))
    fig.patch.set_facecolor("white")
    return fig


def estilo(ax, grilla="x"):
    for lado in ax.spines.values():
        lado.set_visible(False)
    ax.tick_params(length=0, colors=TINTA_SECUNDARIA, labelsize=4.8, pad=1.5)
    if grilla:
        ax.grid(axis=grilla, color=GRILLA, linewidth=0.4)
        ax.set_axisbelow(True)


def guardar(fig, salida):
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


def celdas(ws, rango):
    return [[c.value for c in fila] for fila in ws[rango]]


# V1 · Football field ----------------------------------------------------------------------------
def v1(lang, wb, salida):
    t = T[lang]
    ws = wb["V1_Football"]
    filas = [f for f in celdas(ws, "A5:D10") if f[1] is not None]      # incluye el rango de 52 semanas
    objetivo, precio = ws["B12"].value, ws["B13"].value
    fig = figura("V1")
    ax = fig.add_axes([0.42, 0.17, 0.55, 0.78])
    n = len(filas)
    for k, (metodo, lo, hi, centro) in enumerate(filas):
        y = n - 1 - k
        dcf, apv = metodo.startswith("DCF"), metodo.startswith("APV")
        semanas = metodo.startswith("Rango 52")
        color = (ROJO if "50" in metodo else AZUL) if dcf else (AZUL if apv else NEGRO if semanas else CELESTE)
        if hi - lo > 1:
            ax.barh(y, hi - lo, left=lo, height=0.55, color=color,
                    linewidth=0)
            caja = {"facecolor": "white", "edgecolor": "none", "pad": 0.4}
            ax.text(lo - 25, y, num(lo, lang), ha="right", va="center", fontsize=4.4, color=TINTA_SECUNDARIA,
                    bbox=caja, zorder=5)
            ax.text(hi + 25, y, num(hi, lang), ha="left", va="center", fontsize=4.4, color=TINTA_SECUNDARIA,
                    bbox=caja, zorder=5)
        ax.scatter(centro, y, s=10, color="white", edgecolor=NEGRO, linewidth=0.6, zorder=4)
        if hi - lo <= 1:
            ax.text(centro + 70, y, num(centro, lang), ha="left", va="center", fontsize=4.4, color=TINTA_SECUNDARIA)
        nombre, detalle = t["v1"][metodo]
        ax.text(-0.03, y + 0.12, nombre, transform=ax.get_yaxis_transform(), ha="right", va="center",
                fontsize=4.9, fontweight="bold", color=NEGRO)
        ax.text(-0.03, y - 0.22, detalle, transform=ax.get_yaxis_transform(), ha="right", va="center",
                fontsize=4.1, color=TINTA_SECUNDARIA)
    ax.axvline(objetivo, color=AMARILLO, linewidth=1.1, zorder=2)
    ax.axvline(precio, color=NEGRO, linewidth=0.8, linestyle=(0, (3, 2)), zorder=2)
    ax.text(objetivo - 30, -1.1, f"{t['objetivo']} {num(objetivo, lang)}", ha="right", va="bottom", fontsize=4.6,
            fontweight="bold", color=NEGRO, bbox={"facecolor": "white", "edgecolor": "none", "pad": 0.6})
    ax.text(precio, n - 0.45, f"{t['precio']} {num(precio, lang)}", ha="right", va="bottom", fontsize=4.6,
            fontweight="bold", color=NEGRO, bbox={"facecolor": "white", "edgecolor": "none", "pad": 0.6})
    ax.set_xlim(900, 5100)
    ax.set_xticks([1000, 2000, 3000, 4000, 5000])
    ax.set_ylim(-1.2, n + 0.1)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(lambda x, _: num(x, lang))
    estilo(ax)
    ax.set_xlabel(t["cpa"], fontsize=4.8, color=TINTA_SECUNDARIA, labelpad=2)
    guardar(fig, salida)


# V7 · Múltiplos vs tasa local: la brecha Chile–Brasil es de tasas, no de prima ----------------------
def v7(lang, wb, salida):
    t = T[lang]["v7"]
    ws = wb["V7_Multiplos"]
    pares = [(ws[f"A{r}"].value, ws[f"B{r}"].value, ws[f"C{r}"].value, ws[f"E{r}"].value, ws[f"F{r}"].value)
             for r in range(8, 14)]
    med = {"Chile": (ws["C15"].value, ws["F15"].value), "Brasil": (ws["C16"].value, ws["F16"].value)}
    corto = {"Mallplaza (precio actual)": "Mallplaza", "Cencosud Shopping": "Cencosud S."}
    color = {"Chile": ROJO, "Brasil": AMARILLO}
    fig = figura("V7")
    # Panel A: múltiplo vs bono real local (etiquetas separadas verticalmente dentro de cada país)
    ax = fig.add_axes([0.1, 0.6, 0.86, 0.32])
    for pais in ("Chile", "Brasil"):
        grupo = sorted([p for p in pares if p[1] == pais], key=lambda p: -p[2])
        y_prev = None
        for nombre, _, mult, bono, _ in grupo:
            mp = nombre.startswith("Mallplaza")
            ax.scatter(bono * 100, mult, s=16 if mp else 9, marker="D" if mp else "o", zorder=4,
                       color=AZUL if mp else color[pais], edgecolor="white", linewidth=0.4)
            y_lab = mult if y_prev is None else min(mult, y_prev - 1.25)
            y_prev = y_lab
            ax.text(bono * 100 + 0.3, y_lab, f"{corto.get(nombre, nombre)} {num(mult, lang, 1)}x", fontsize=3.8,
                    va="center", color=AZUL if mp else NEGRO, fontweight="bold" if mp else "normal")
    for pais, (m, _) in med.items():
        x = (0.025 if pais == "Chile" else ws["E6"].value) * 100
        ax.plot([x - 0.6, x + 0.2], [m, m], color=color[pais], linewidth=0.8, linestyle=(0, (2, 1)))
        ax.text(x - 0.7, m, f"{t['med']} {num(m, lang, 1)}x", fontsize=3.6, ha="right", va="center",
                color=TINTA_SECUNDARIA)
    ax.set_xlim(-3.2, 11)
    ax.set_ylim(5.5, 20.5)
    ax.set_xticks([0, 2.5, 5, 7.5, 10])
    ax.xaxis.set_major_formatter(lambda x, _: num(x, lang, 1) + "%")
    ax.set_yticks([8, 12, 16, 20])
    estilo(ax, grilla="y")
    ax.tick_params(labelsize=4.3)
    ax.set_title(t["a"], fontsize=4.6, color=TINTA_SECUNDARIA, loc="left", pad=3)
    # Panel B: spread real (barras horizontales), medianas destacadas
    ax = fig.add_axes([0.3, 0.04, 0.62, 0.43])
    filas = [(corto.get(n, n), pais, sp) for n, pais, _, _, sp in pares if pais == "Chile"]
    filas.append((f"{t['med']} {t['cl']}", "Chile", med["Chile"][1]))
    filas += [(n, pais, sp) for n, pais, _, _, sp in pares if pais == "Brasil"]
    filas.append((f"{t['med']} {t['br']}", "Brasil", med["Brasil"][1]))
    for k, (nombre, pais, sp) in enumerate(filas):
        y = len(filas) - 1 - k
        es_med, mp = nombre.startswith(t["med"]), nombre == "Mallplaza"
        ax.barh(y, sp, height=0.66, color=AZUL if mp else (NEGRO if es_med else color[pais]), zorder=3)
        ax.text(sp + 8, y, num(sp, lang), va="center", fontsize=3.9, fontweight="bold" if es_med else "normal",
                color=NEGRO)
        ax.text(-12, y, nombre, ha="right", va="center", fontsize=3.9, fontweight="bold" if es_med or mp else "normal",
                color=AZUL if mp else NEGRO)
    ax.set_xlim(0, 600)
    ax.set_ylim(-0.6, len(filas) - 0.4)
    ax.set_xticks([])
    ax.set_yticks([])
    estilo(ax, grilla=None)
    fig.text(0.1, 0.5, t["b"], fontsize=4.6, color=TINTA_SECUNDARIA, va="center")
    guardar(fig, salida)


# V2 · Escalera de tasas reales --------------------------------------------------------------------
def v2(lang, wb, salida):
    t = T[lang]
    ws = wb["V2_WACC_Tasas"]
    filas = sorted(celdas(ws, "A7:C14"), key=lambda f: f[2])
    kd, impl = ws["C10"].value, ws["C11"].value
    fig = figura("V2")
    ax = fig.add_axes([0.47, 0.15, 0.5, 0.83])
    for y, (nombre, _, real) in enumerate(filas):
        color = ROJO if "implícita" in nombre else AZUL if nombre.startswith("Nuestra WACC") else \
            NEGRO if nombre.startswith("Nuestro Ke") else CELESTE
        ax.barh(y, real * 100, height=0.6, color=color, linewidth=0)
        ax.text(real * 100 + 0.1, y, pct(real, lang), va="center", fontsize=4.6, fontweight="bold", color=NEGRO)
        ax.text(-0.03, y, t["v2"][nombre], transform=ax.get_yaxis_transform(), ha="right", va="center",
                fontsize=4.6, color=NEGRO, fontweight="bold" if color in (ROJO, AZUL) else "normal")
    # corchete Kd → WACC implícita
    yk = [f[0] for f in filas].index("Kd Mallplaza (cupón última emisión UF)")
    yi = [f[0] for f in filas].index("WACC implícita en el precio actual (DCF inverso)")
    xb = max(kd, impl) * 100 + 0.85
    ax.plot([xb - 0.12, xb, xb, xb - 0.12], [yk, yk, yi, yi], color=ROJO, linewidth=0.6)
    ax.text(xb + 0.1, (yk + yi) / 2, f"+{num(ws['B16'].value, lang)} {t['pb']}", va="center", fontsize=4.6,
            fontweight="bold", color=ROJO)
    ax.set_xlim(0, 10)
    ax.set_ylim(-0.6, len(filas) - 0.4)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:.0f}%")
    estilo(ax)
    ax.set_xlabel(t["real"], fontsize=4.8, color=TINTA_SECUNDARIA, labelpad=2)
    guardar(fig, salida)


# V3 · Value driver: tabla + ilustración ----------------------------------------------------------
def v3(lang, wb, salida):
    t = T[lang]["v3"]
    ws = wb["V3_ValueDriver"]
    tabla = [f for f in celdas(ws, "A14:F17") if f[0] >= 0.03]          # la nota pide mostrar 3,0–4,0%
    fig = figura("V3")
    # Tabla
    axt = fig.add_axes([0.03, 0.6, 0.94, 0.38])
    axt.axis("off")
    xs = [0.0, 0.22, 0.46, 0.7, 0.9]
    alin = ["left", "right", "right", "right", "right"]
    for x, c, a in zip(xs, t["cols"], alin):
        axt.text(x + (0.08 if a == "right" else 0.01), 0.9, c, ha=a, va="center", fontsize=4.8, fontweight="bold",
                 color="white", transform=axt.transAxes)
    axt.add_patch(Rectangle((0, 0.78), 1, 0.24, color=AZUL, transform=axt.transAxes))
    for k, (gn, gr, reinv, _, valor, delta) in enumerate(tabla):
        y = 0.62 - k * 0.22
        base = abs(gn - 0.03) < 1e-9
        if base:
            axt.add_patch(Rectangle((0, y - 0.11), 1, 0.22, color=AMARILLO, transform=axt.transAxes))
        vals = [pct(gn, lang), pct(gr, lang), pct(reinv, lang), num(valor, lang),
                ("±0,0%" if lang == "es" else "±0.0%") if base else pct(delta, lang)]
        for x, v, a in zip(xs, vals, alin):
            axt.text(x + (0.08 if a == "right" else 0.01), y, v, ha=a, va="center", fontsize=4.8,
                     fontweight="bold" if base else "normal", color=NEGRO, transform=axt.transAxes)
    # Línea (panel C)
    ax = fig.add_axes([0.13, 0.12, 0.84, 0.4])
    serie = celdas(ws, "A24:D27")
    gs = [f[0] * 100 for f in serie]
    for j, (color, nombre) in enumerate(zip([ROJO, NEGRO, CELESTE], t["series"])):
        ys = [f[j + 1] for f in serie]
        ax.plot(gs, ys, color=color, linewidth=1.2, marker="o", markersize=2.4, markeredgecolor="white",
                markeredgewidth=0.4)
        ax.text(gs[-1] + 0.05, ys[-1], nombre, fontsize=4.2, color=color,
                va="center", fontweight="bold")
    ax.set_xlim(-0.05, 2.6)
    ax.set_ylim(85, 110)
    ax.set_xticks(gs, [pct(g / 100, lang) for g in gs])
    estilo(ax, grilla="y")
    ax.set_xlabel(t["eje"], fontsize=4.6, color=TINTA_SECUNDARIA, labelpad=1.5)
    ax.set_ylabel(t["indice"], fontsize=4.3, color=TINTA_SECUNDARIA, labelpad=2)
    guardar(fig, salida)


# V4 · Capital de trabajo: cascada + tabla --------------------------------------------------------
def v4(lang, wb, salida):
    t = T[lang]["v4"]
    ws = wb["V4_CapitalTrabajo"]
    comp = [ws[f"D{r}"].value for r in range(5, 10)]
    signos = [1, 1, -1, -1, -1]
    total = ws["D10"].value
    fig = figura("V4")
    ax = fig.add_axes([0.1, 0.52, 0.88, 0.44])
    acum = 0
    for k, (v, s) in enumerate(zip(comp, signos)):
        inicio = acum
        acum += s * v
        ax.bar(k, s * v / 1000, bottom=inicio / 1000, width=0.62, color=CELESTE if s > 0 else ROJO, linewidth=0)
        ax.text(k, max(inicio, acum) / 1000 + 3, ("+" if s > 0 else "−") + num(v / 1000, lang), ha="center",
                va="bottom", fontsize=4.4, color=NEGRO)
        if k < len(comp) - 1:
            ax.plot([k + 0.31, k + 0.69], [acum / 1000] * 2, color=NEGRO, linewidth=0.3)
    assert abs(acum - total) < 1, "la cascada debe cerrar en el CTN"
    ax.bar(len(comp), total / 1000, width=0.62, color=AZUL, linewidth=0)
    ax.text(len(comp), total / 1000 + 3, num(total / 1000, lang), ha="center", va="bottom", fontsize=4.6,
            fontweight="bold", color=AZUL)
    ax.set_xticks(range(len(comp) + 1), t["comp"], fontsize=4.2, linespacing=1.05)
    ax.set_ylim(0, 160)
    ax.yaxis.set_major_formatter(lambda y, _: num(y, lang))
    estilo(ax, grilla="y")
    ax.set_ylabel(t["mm"].replace("MM CLP", "miles de MM CLP") if lang == "es" else t["mm"].replace("CLP MM", "CLP bn"),
                  fontsize=4.2, color=TINTA_SECUNDARIA, labelpad=2)
    # Tabla de sensibilidad
    axt = fig.add_axes([0.03, 0.02, 0.94, 0.33])
    axt.axis("off")
    xs, alin = [0.0, 0.52, 0.78, 0.98], ["left", "right", "right", "right"]
    axt.add_patch(Rectangle((0, 0.83), 1, 0.17, color=AZUL, transform=axt.transAxes))
    for x, c, a in zip(xs, t["cols"], alin):
        axt.text(x + (0.01 if a == "left" else 0), 0.915, c, ha=a, va="center", fontsize=4.5, fontweight="bold",
                 color="white", transform=axt.transAxes)
    for k, r in enumerate(range(15, 20)):
        y = 0.72 - k * 0.165
        base = r == 17
        if base:
            axt.add_patch(Rectangle((0, y - 0.08), 1, 0.16, color=AMARILLO, transform=axt.transAxes))
        d = ws[f"F{r}"].value
        vals = [t["esc"][k], pct(ws[f"C{r}"].value, lang), num(ws[f"E{r}"].value, lang),
                ("±0,0%" if lang == "es" else "±0.0%") if base else ("+" if d > 0 else "") + pct(d, lang)]
        for x, v, a in zip(xs, vals, alin):
            axt.text(x + (0.01 if a == "left" else 0), y, v, ha=a, va="center", fontsize=4.5, color=NEGRO,
                     fontweight="bold" if base else "normal", transform=axt.transAxes)
    guardar(fig, salida)


# V5 · Cap rate implícito vs referencias ----------------------------------------------------------
def v5(lang, wb, salida):
    t = T[lang]
    ws = wb["V5_CapRate"]
    filas = sorted(celdas(ws, "A5:B10"), key=lambda f: f[1])
    fig = figura("V5")
    ax = fig.add_axes([0.42, 0.14, 0.55, 0.84])
    for y, (nombre, v) in enumerate(filas):
        color = ROJO if nombre.startswith("Nuestro") else AMARILLO if "Gran Plaza" in nombre else CELESTE
        ax.barh(y, v * 100, height=0.6, color=color, linewidth=0)
        ax.text(v * 100 + 0.12, y, pct(v, lang), va="center", fontsize=4.6, fontweight="bold", color=NEGRO)
        ax.text(-0.03, y, t["v5"][nombre], transform=ax.get_yaxis_transform(), ha="right", va="center",
                fontsize=4.6, color=NEGRO, fontweight="bold" if color == ROJO else "normal")
    nombres = [f[0] for f in filas]
    ym, yd = nombres.index("Mercado: Mallplaza al precio actual"), nombres.index("Nuestro DCF (EBITDA LTM / EV DCF)")
    xb = 12.6
    ax.plot([xb - 0.15, xb, xb, xb - 0.15], [ym, ym, yd, yd], color=ROJO, linewidth=0.6)
    ax.text(xb + 0.15, (ym + yd) / 2, f"{t['brecha']}\n{num(ws['B12'].value, lang)} {t['pb']}", va="center",
            fontsize=4.5, fontweight="bold", color=ROJO, linespacing=1.1)
    ax.set_xlim(0, 15.5)
    ax.set_ylim(-0.6, len(filas) - 0.4)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:.0f}%")
    estilo(ax)
    ax.set_xlabel(t["cap"], fontsize=4.8, color=TINTA_SECUNDARIA, labelpad=2)
    guardar(fig, salida)


# V6 · DCF inverso -----------------------------------------------------------------------------
def v6(lang, wb, salida):
    t = T[lang]
    tt = t["v6"]
    ws = wb["V6_DCFInverso"]
    filas = celdas(ws, "A5:E15")
    w = [f[1] * 100 for f in filas]
    dcf = [f[2] for f in filas]
    po = [f[3] for f in filas]
    precio = filas[0][4]
    impl, brecha = ws["B17"].value * 100, ws["B18"].value
    base_w = filas[[f[0] for f in filas].index(0)][1] * 100
    fig = figura("V6")
    ax = fig.add_axes([0.13, 0.15, 0.84, 0.8])
    ax.plot(w, dcf, color=ROJO, linewidth=1.3, label=tt["dcf"])
    ax.plot(w, po, color=AZUL, linewidth=1.1, linestyle=(0, (4, 1.5)), label=tt["po"])
    ax.axhline(precio, color=NEGRO, linewidth=0.7, linestyle=(0, (3, 2)))
    ax.text(w[-1], precio + 60, f"{t['precio']} {num(precio, lang)}", ha="right", va="bottom", fontsize=4.5,
            fontweight="bold", color=NEGRO)
    i0 = w.index(base_w)
    ax.scatter([base_w], [dcf[i0]], s=12, color=ROJO, zorder=4, edgecolor="white", linewidth=0.5)
    ax.scatter([impl], [precio], s=12, color=NEGRO, zorder=4, edgecolor="white", linewidth=0.5)
    ax.vlines([impl], 1150, precio, color=NEGRO, linewidth=0.5, linestyle=(0, (1, 1.5)))
    ax.vlines([base_w], 1150, dcf[i0], color=ROJO, linewidth=0.5, linestyle=(0, (1, 1.5)))
    ax.text(impl - 0.08, 1230, f"{tt['impl']} {pct(impl / 100, lang, 2)}", ha="right", va="bottom",
            fontsize=4.4, fontweight="bold", color=NEGRO)
    ax.text(base_w + 0.08, 1230, f"{tt['base']} {pct(base_w / 100, lang, 2)}", ha="left", va="bottom",
            fontsize=4.4, fontweight="bold", color=ROJO)
    ax.annotate("", xy=(impl, 1150), xytext=(base_w, 1150),
                arrowprops={"arrowstyle": "-|>", "color": ROJO, "linewidth": 0.7, "mutation_scale": 5})
    ax.text((impl + base_w) / 2, 1190, f"{num(brecha, lang)} {t['pb']}", ha="center", va="bottom", fontsize=4.6,
            fontweight="bold", color=ROJO)
    ax.set_xlim(min(w) - 0.1, max(w) + 0.1)
    ax.set_ylim(1000, 5300)
    ax.set_yticks(range(1500, 5001, 500))
    ax.xaxis.set_major_formatter(lambda x, _: f"{x:.0f}%")
    ax.yaxis.set_major_formatter(lambda y, _: num(y, lang))
    estilo(ax, grilla="y")
    ax.set_xlabel(tt["eje"], fontsize=4.8, color=TINTA_SECUNDARIA, labelpad=2)
    ax.legend(loc="upper right", frameon=False, fontsize=4.4, handlelength=2)
    guardar(fig, salida)


# V9 · Valor implícito por múltiplo de cada par (football field del equipo, tamaño columna) -------
def v9(lang, ruta_v2, salida):
    from cuota_mercado import MORADO_CENCOSUD, VERDE_PARQUE_ARAUCO
    from football_field import COLOR_PAR, leer
    filas, dcf, _, objetivo, precio, _ = leer(ruta_v2)
    t = T[lang]
    fig = figura("V9")
    ax = fig.add_axes([0.3, 0.2, 0.67, 0.76])
    n = len(filas) + 1
    caja = {"facecolor": "white", "edgecolor": "none", "pad": 0.4}
    y = n - 1
    ax.barh(y, dcf[2] - dcf[0], left=dcf[0], height=0.5, color=ROJO, linewidth=0)
    ax.plot([dcf[1]] * 2, [y - 0.3, y + 0.3], color=NEGRO, linewidth=1.4, solid_capstyle="butt")
    ax.text(dcf[1], y + 0.33, num(dcf[1], lang), ha="center", va="bottom", fontsize=4.4, fontweight="bold", color=ROJO)
    ax.text(-0.03, y + 0.12, "DCF (FCFF)", transform=ax.get_yaxis_transform(), ha="right", va="center",
            fontsize=4.8, fontweight="bold", color=NEGRO)
    ax.text(-0.03, y - 0.22, "±100 pb · g 2,5–4,0%" if lang == "es" else "±100 bp · g 2.5–4.0%",
            transform=ax.get_yaxis_transform(), ha="right", va="center", fontsize=4, color=TINTA_SECUNDARIA)
    for y, (nombre, es_ev, pares, mult, valor) in zip(range(n - 2, -1, -1), filas):
        vals = [p[2] for p in pares]
        ax.barh(y, max(vals) - min(vals), left=min(vals), height=0.5, color=CELESTE if es_ev else "white",
                edgecolor=CELESTE, linewidth=0 if es_ev else 0.4, hatch=None if es_ev else "////////")
        for par, _, v in sorted(pares, key=lambda p: p[0] in COLOR_PAR):
            chile = par in COLOR_PAR
            ax.scatter(v, y, s=9 if chile else 5, color=COLOR_PAR.get(par, AZUL), edgecolor="white",
                       linewidth=0.4, zorder=3 + chile)
        ax.plot([valor] * 2, [y - 0.3, y + 0.3], color=NEGRO, linewidth=1.4, solid_capstyle="butt", zorder=5)
        ax.text(valor, y + 0.33, num(valor, lang), ha="center", va="bottom", fontsize=4.4, fontweight="bold",
                color=NEGRO, zorder=6, bbox=caja)
        ax.text(-0.03, y + 0.12, nombre.replace(" Jun-26", ""), transform=ax.get_yaxis_transform(), ha="right",
                va="center", fontsize=4.8, fontweight="bold", color=NEGRO)
        detalle = (f"ponderado {num(mult, lang, 1)}x" if lang == "es" else f"weighted {num(mult, lang, 1)}x") + \
            ("" if es_ev else " · ref.")
        ax.text(-0.03, y - 0.22, detalle, transform=ax.get_yaxis_transform(), ha="right", va="center", fontsize=4,
                color=TINTA_SECUNDARIA)
    ax.axvline(objetivo, color=AMARILLO, linewidth=1.0, zorder=2)
    ax.axvline(precio, color=NEGRO, linewidth=0.7, linestyle=(0, (3, 2)), zorder=2)
    ax.text(objetivo - 40, -0.75, f"{t['objetivo']} {num(objetivo, lang)}", ha="right", va="center", fontsize=4.4,
            fontweight="bold", color=NEGRO)
    ax.text(precio + 40, -0.75, f"{t['precio']} {num(precio, lang)}", ha="left", va="center", fontsize=4.4,
            fontweight="bold", color=NEGRO)
    xs = [p[2] for f in filas for p in f[2]] + [dcf[0], dcf[2], precio]
    ax.set_xlim(min(xs) - 200, max(xs) + 200)
    ax.set_ylim(-1.05, n - 0.25)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(lambda x, _: num(x, lang))
    estilo(ax)
    ax.set_xlabel(t["cpa"], fontsize=4.6, color=TINTA_SECUNDARIA, labelpad=1.5)
    punto = lambda c, txt: plt.Line2D([], [], marker="o", linestyle="", markersize=3, markerfacecolor=c,
                                      markeredgecolor="white", label=txt)
    fig.legend(handles=[punto(VERDE_PARQUE_ARAUCO, "Parque Arauco"), punto(MORADO_CENCOSUD, "Cencosud"),
                        punto(AZUL, "Multiplan, Allos, Iguatemi, IRSA"),
                        plt.Line2D([], [], color=NEGRO, linewidth=1.4, label="Ponderado" if lang == "es" else "Weighted")],
               loc="lower center", bbox_to_anchor=(0.5, 0.0), ncol=4, frameon=False, fontsize=3.9,
               handletextpad=0.2, columnspacing=0.6)
    guardar(fig, salida)


# W1 · Anexo WACC: WACC nominal y D/E por período -----------------------------------------------
def w1(lang, wb, salida):
    tt = T[lang]["w1"]
    ws = wb["WACC_Construccion"]
    wacc = [ws.cell(65, c).value * 100 for c in range(2, 9)]
    de = [ws.cell(56, c).value * 100 for c in range(2, 9)]
    fig = figura("W1", COMPLETO / 2 - 0.08)
    ax = fig.add_axes([0.1, 0.14, 0.78, 0.8])
    x = np.arange(len(wacc))
    ax2 = ax.twinx()
    ax2.bar(x, de, width=0.55, color=CELESTE, linewidth=0, zorder=1)
    for xi, v in zip(x, de):
        ax2.text(xi, 1.5, pct(v / 100, lang), ha="center", va="bottom", fontsize=5, color=AZUL, fontweight="bold")
    ax.plot(x, wacc, color=ROJO, linewidth=1.5, marker="o", markersize=3.2, markeredgecolor="white",
            markeredgewidth=0.5, zorder=3)
    for xi, v in zip(x, wacc):
        ax.text(xi, v + 0.025, pct(v / 100, lang, 2), ha="center", va="bottom", fontsize=5.2, fontweight="bold",
                color=ROJO)
    ax.set_zorder(ax2.get_zorder() + 1)
    ax.patch.set_visible(False)
    ax.set_ylim(10.0, 10.62)
    ax2.set_ylim(0, 60)
    ax2.set_yticks([0, 10, 20, 30])
    ax.set_xticks(x, tt["per"])
    for a in (ax, ax2):
        for lado in a.spines.values():
            lado.set_visible(False)
        a.tick_params(length=0, colors=TINTA_SECUNDARIA, labelsize=5.2, pad=1.5)
    ax.yaxis.set_major_formatter(lambda y, _: pct(y / 100, lang, 2))
    ax2.yaxis.set_major_formatter(lambda y, _: f"{y:.0f}%")
    ax.legend(handles=[plt.Line2D([], [], color=ROJO, marker="o", markersize=3, label=tt["wacc"]),
                       Patch(color=CELESTE, label=tt["de"])],
              loc="upper right", frameon=False, fontsize=5)
    guardar(fig, salida)


def exportar_wacc(wb, ruta):
    """Datos del anexo de la WACC (hoja WACC_Construccion) para informe/valuacion.js."""
    import json
    ws = wb["WACC_Construccion"]
    v = lambda c: ws[c].value
    datos = {
        "inputs": {k: v(c) for k, c in (("rf", "B5"), ("erp", "B6"), ("crp_cl", "B7"), ("crp_pe", "B8"),
                                        ("crp_co", "B9"), ("pi_cl", "B10"), ("pi_us", "B11"), ("tax", "B12"),
                                        ("kd_real", "B13"), ("spread", "B14"), ("dfn0", "B15"), ("dfn_obj", "B16"),
                                        ("anios", "B17"), ("deuda", "B18"), ("mcap", "B19"), ("ebitda", "B20"),
                                        ("caja", "B21"))},
        "betas": [[ws.cell(r, c).value for c in range(1, 8)] for r in range(32, 39)],
        "beta_u": v("B40"),
        "crp": {"w": [v("B45"), v("B46"), v("B47")], "total": v("B48")},
        "periodos": {str(r): [ws.cell(r, c).value for c in range(2, 9)] for r in range(52, 67)},
        "control": [ws.cell(70, c).value for c in range(2, 9)],
        "rf": {k: v(c) for k, c in (("base", "B73"), ("bono", "B74"), ("dif", "B75"), ("efecto", "B76"),
                                    ("dcf", "B77"), ("po", "B78"))},
    }
    ruta.write_text(json.dumps(datos, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    salida = Path(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[2] != "-" else Path(__file__).parent / "informe" / "valuacion"
    salida.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans", "hatch.linewidth": 0.4})
    wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
    for lang in ("es", "en"):
        for nombre, fn in (("V1_football", v1), ("V2_tasas", v2), ("V3_value_driver", v3),
                           ("V4_capital_trabajo", v4), ("V5_cap_rate", v5), ("V6_dcf_inverso", v6),
                           ("W1_wacc_periodo", w1), ("V7_multiplos", v7)):
            fn(lang, wb, salida / f"{lang}_{nombre}.png")
    exportar_wacc(wb, salida / "datos_wacc.json")
    if len(sys.argv) > 3:   # modelo consolidado v2 para el football field de pares
        for lang in ("es", "en"):
            v9(lang, sys.argv[3], salida / f"{lang}_V9_pares.png")
    print(f"Figuras guardadas en {salida}")
