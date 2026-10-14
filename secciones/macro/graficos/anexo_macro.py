"""Figuras del Anexo macroeconómico (2022–2026), en español e inglés.

Lee data/macro/macro_2022_2026.xlsx (cada cifra con su fuente y estado) y dibuja, sin título ni
fuente dentro de la imagen (van en el Word, informe/anexo_macro.js):
  A0 mapa de transmisión · A1 Brent · A2 tasas soberanas a 10 años · A3 spreads ·
  A4 crecimiento vs inflación mundial · A5 Chile (PIB, IPC, TPM).

Uso:
    python graficos/anexo_macro.py [macro_2022_2026.xlsx] [carpeta_salida]
"""
import sys
from datetime import date
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Patch

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional y rutas compartidas
import estilo  # noqa: E402,F401  (registra las carpetas graficos/ de todas las secciones)
from estilo import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA

RAIZ = Path(__file__).resolve().parent.parent
AZUL = "#14284B"
GRIS = "#9a9a9a"
MEDIO = (10666 - 216) / 2 / 1440   # media columna del anexo (pulgadas)
COMPLETO = 10666 / 1440            # ancho completo
DPI = 400
VALORACION = date(2026, 6, 30)

T = {
    "es": {"meses": "ene feb mar abr may jun jul ago sep oct nov dic".split(),
           "ucrania": "Invasión de\nUcrania", "ormuz": "Conflicto\nen Ormuz", "val": "Valoración",
           "usd": "USD por barril", "pais": ["EE.UU.", "Alemania", "Francia", "Chile"],
           "fr_de": "Francia – Alemania", "cl_us": "Chile – EE.UU.", "pb": "Puntos base",
           "pib": "PIB mundial", "ipc": "Inflación mundial", "pib_cl": "PIB", "ipc_cl": "IPC (dic.)",
           "tpm": "TPM (cierre)", "proy": "proyección",
           "mapa": {"t1": "SHOCKS GLOBALES", "t2": "CANALES EN CHILE", "t3": "EFECTO EN MALLPLAZA",
                    "s": ["Energía\nUcrania 2022 · Ormuz 2026", "Fiscal\ndéficits y deuda → tasas largas",
                          "Geopolítica\nfragmentación y aranceles"],
                    "c": ["Inflación y UF", "Tasas largas locales", "Consumo y PIB", "Tipo de cambio y flujos"],
                    "m": ["Deuda UF: UF +3% = −CLP 38.691 MM; rentas UF compensan",
                          "WACC +100 pb = precio objetivo −12,3%", "Ventas de locatarios y tráfico",
                          "Financiamiento y riesgo país (CRP 1,41% en el Ke)"],
                    "r": ["MR1", "FR1", "MR1 · MR2", "FR1 · MR3"]}},
    "en": {"meses": "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split(),
           "ucrania": "Ukraine\ninvasion", "ormuz": "Hormuz\nconflict", "val": "Valuation",
           "usd": "USD per barrel", "pais": ["US", "Germany", "France", "Chile"],
           "fr_de": "France – Germany", "cl_us": "Chile – US", "pb": "Basis points",
           "pib": "World GDP", "ipc": "World inflation", "pib_cl": "GDP", "ipc_cl": "CPI (Dec.)",
           "tpm": "Policy rate (year-end)", "proy": "projection",
           "mapa": {"t1": "GLOBAL SHOCKS", "t2": "CHANNELS INTO CHILE", "t3": "EFFECT ON MALLPLAZA",
                    "s": ["Energy\nUkraine 2022 · Hormuz 2026", "Fiscal\ndeficits and debt → long rates",
                          "Geopolitics\nfragmentation and tariffs"],
                    "c": ["Inflation and UF", "Local long-term rates", "Consumption and GDP", "FX and capital flows"],
                    "m": ["UF debt: UF +3% = −CLP 38,691 MM; UF rents offset",
                          "WACC +100 bp = target price −12.3%", "Tenant sales and footfall",
                          "Funding and country risk (1.41% CRP in Ke)"],
                    "r": ["MR1", "FR1", "MR1 · MR2", "FR1 · MR3"]}},
}


def num(x, lang, dec=1):
    s = f"{x:,.{dec}f}".replace("-", "−")
    return s.translate(str.maketrans(",.", ".,")) if lang == "es" else s


def leer(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)

    def tabla(nombre, ncol):
        ws = wb[nombre]
        return [[ws.cell(r, c).value for c in range(1, ncol + 1)] for r in range(3, ws.max_row + 1)]

    brent = tabla("Brent", 2)
    tasas = tabla("Tasas10Y", 5)
    fechas = [date(int(m[:4]), int(m[5:7]), 15) for m, _ in brent]
    return {"fechas": fechas, "brent": np.array([v for _, v in brent], float),
            "tasas": np.array([f[1:] for f in tasas], float),
            "mundo": tabla("Mundo", 3), "chile": tabla("Chile", 4)}


def figura(ancho, alto):
    fig = plt.figure(figsize=(ancho, alto))
    fig.patch.set_facecolor("white")
    return fig


def estilo(ax, grilla="y"):
    for lado in ax.spines.values():
        lado.set_visible(False)
    ax.tick_params(length=0, colors=TINTA_SECUNDARIA, labelsize=5.5, pad=2)
    if grilla:
        ax.grid(axis=grilla, color=GRILLA, linewidth=0.4)
        ax.set_axisbelow(True)


def eje_fechas(ax, lang):
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.set_xlim(date(2022, 1, 1), date(2026, 10, 15))


def marcar(ax, fecha, texto, y, ha="left"):
    ax.axvline(fecha, color=TINTA_SECUNDARIA, linewidth=0.6, linestyle=(0, (2, 2)), zorder=1)
    ax.text(fecha, y, " " + texto + " " if ha == "left" else texto + " ", ha=ha, va="top",
            fontsize=4.8, color=TINTA_SECUNDARIA, linespacing=1.15)


def guardar(fig, salida):
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


# A0 · Mapa de transmisión -----------------------------------------------------------------------
def mapa(lang, salida):
    t = T[lang]["mapa"]
    fig = figura(COMPLETO, 1.85)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 40)
    ax.axis("off")
    cols = [(1, 23), (31, 22), (59, 40)]
    for (x, w), titulo in zip(cols, (t["t1"], t["t2"], t["t3"])):
        ax.text(x, 37.6, titulo, fontsize=6, fontweight="bold", color=AZUL, va="center")
        ax.plot([x, x + w], [35.8, 35.8], color=ROJO, linewidth=0.8)

    def caja(x, y, w, h, texto, fondo, letra, peso="normal", tam=5.4):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=1.2",
                                    facecolor=fondo, edgecolor="none"))
        lineas = texto.split("\n")
        if len(lineas) == 2:
            ax.text(x + 1.5, y + h * 0.64, lineas[0], fontsize=tam, fontweight="bold", color=letra, va="center")
            ax.text(x + 1.5, y + h * 0.3, lineas[1], fontsize=tam - 0.8, color=letra, va="center")
        else:
            ax.text(x + 1.5, y + h / 2, texto, fontsize=tam, fontweight=peso, color=letra, va="center")
        return (x, y, w, h)

    ys = [24.5, 13.5, 2.5]
    s_cajas = [caja(cols[0][0], y, cols[0][1], 9, s, AZUL, "white") for y, s in zip(ys, t["s"])]
    yc = [27.5, 19.5, 11.5, 3.5]
    c_cajas = [caja(cols[1][0], y, cols[1][1], 6, c, "#e8f3fb", NEGRO, "bold") for y, c in zip(yc, t["c"])]
    m_cajas = []
    for y, m, r in zip(yc, t["m"], t["r"]):
        mitigante = False
        m_cajas.append(caja(cols[2][0], y, cols[2][1] - 8, 6, m, "#eaf6ea" if mitigante else "#fde7ee", NEGRO))
        ax.add_patch(FancyBboxPatch((cols[2][0] + cols[2][1] - 7, y + 1.3), 7, 3.4,
                                    boxstyle="round,pad=0,rounding_size=1.1",
                                    facecolor=CELESTE if mitigante else ROJO, edgecolor="none"))
        ax.text(cols[2][0] + cols[2][1] - 3.5, y + 3, "UF" if mitigante else r, ha="center", va="center",
                fontsize=4.6, fontweight="bold", color="white")

    def flecha(a, b):
        xa, ya, wa, ha = a
        xb, yb, wb_, hb = b
        ax.add_patch(FancyArrowPatch((xa + wa + 0.6, ya + ha / 2), (xb - 0.6, yb + hb / 2),
                                     arrowstyle="-|>", mutation_scale=5, color=GRIS, linewidth=0.6,
                                     shrinkA=0, shrinkB=0, connectionstyle="arc3,rad=0"))

    # Energía → inflación, consumo; fiscal → tasas; geopolítica → tipo de cambio, consumo
    for i, j in ((0, 0), (0, 2), (1, 1), (2, 3), (2, 2)):
        flecha(s_cajas[i], c_cajas[j])
    for j in range(4):
        flecha(c_cajas[j], m_cajas[j])
    guardar(fig, salida)


# A1 · Brent -------------------------------------------------------------------------------------
def brent(lang, d, salida):
    t = T[lang]
    fig = figura(MEDIO, 2.55)
    ax = fig.add_axes([0.1, 0.1, 0.87, 0.87])
    f, v = d["fechas"], d["brent"]
    ax.fill_between(f, v, color=ROJO, alpha=0.08, linewidth=0)
    ax.plot(f, v, color=ROJO, linewidth=1.3)
    marcar(ax, date(2022, 2, 24), t["ucrania"], 168)
    marcar(ax, date(2026, 2, 28), t["ormuz"], 168, ha="right")
    ax.axvline(VALORACION, color=AZUL, linewidth=0.7, linestyle=(0, (1, 1.5)))
    ax.text(VALORACION, 4, t["val"] + " ", fontsize=4.6, color=AZUL, va="bottom", ha="right")
    for i in (int(np.argmax(v[:12])), int(np.argmax(v[48:56])) + 48, len(v) - 1):
        ax.scatter(f[i], v[i], s=8, color=ROJO, zorder=3, edgecolor="white", linewidth=0.5)
        mes = t["meses"][f[i].month - 1] + "-" + str(f[i].year)[2:]
        ax.text(f[i], v[i] + 5, f"{num(v[i], lang, 0)}\n{mes}", ha="center", va="bottom",
                fontsize=4.8, fontweight="bold", color=NEGRO, linespacing=1.1)
    eje_fechas(ax, lang)
    ax.set_ylim(0, 170)
    ax.set_yticks(range(0, 161, 40))
    ax.set_ylabel(t["usd"], fontsize=5.2, color=TINTA_SECUNDARIA, labelpad=2)
    estilo(ax)
    guardar(fig, salida)


# A2 · Tasas a 10 años ---------------------------------------------------------------------------
COLORES_PAIS = [ROJO, NEGRO, CELESTE, AMARILLO]


def tasas(lang, d, salida):
    t = T[lang]
    fig = figura(MEDIO, 2.55)
    ax = fig.add_axes([0.08, 0.1, 0.72, 0.87])
    f = d["fechas"]
    for k, (pais, color) in enumerate(zip(t["pais"], COLORES_PAIS)):
        serie = d["tasas"][:, k]
        ax.plot(f, serie, color=color, linewidth=1.5 if k == 0 else 1.1, zorder=3 if k == 0 else 2)
        ax.text(date(2026, 10, 25), serie[-1], f"{pais} {num(serie[-1], lang, 1)}%", fontsize=4.9,
                fontweight="bold", color=color if color != AMARILLO else "#b58800", va="center")
    ax.axvline(VALORACION, color=AZUL, linewidth=0.7, linestyle=(0, (1, 1.5)))
    ax.text(VALORACION, -0.85, t["val"] + " ", fontsize=4.6, color=AZUL, va="bottom", ha="right")
    marcar(ax, date(2026, 2, 28), t["ormuz"], 7.6, ha="right")
    eje_fechas(ax, lang)
    ax.set_ylim(-1, 8)
    ax.yaxis.set_major_formatter(lambda y, _: f"{y:.0f}%")
    estilo(ax)
    guardar(fig, salida)


# A3 · Spreads -----------------------------------------------------------------------------------
def spreads(lang, d, salida):
    t = T[lang]
    fig = figura(MEDIO, 2.3)
    ax = fig.add_axes([0.1, 0.11, 0.74, 0.86])
    f, x = d["fechas"], d["tasas"]
    series = [((x[:, 2] - x[:, 1]) * 100, CELESTE, t["fr_de"]), ((x[:, 3] - x[:, 0]) * 100, ROJO, t["cl_us"])]
    for s, color, nombre in series:
        ax.plot(f, s, color=color, linewidth=1.3)
        ax.text(date(2026, 10, 25), s[-1], f"{num(s[-1], lang, 0)} pb" if lang == "es" else f"{num(s[-1], lang, 0)} bp",
                fontsize=4.9, fontweight="bold", color=color, va="center")
        ax.text(f[2], s[2] + 25, nombre, fontsize=5, fontweight="bold", color=color, va="bottom")
    ax.axvline(VALORACION, color=AZUL, linewidth=0.7, linestyle=(0, (1, 1.5)))
    ax.text(VALORACION, 5, t["val"] + " ", fontsize=4.6, color=AZUL, va="bottom", ha="right")
    eje_fechas(ax, lang)
    ax.set_ylim(0, 500)
    ax.set_ylabel(t["pb"], fontsize=5.2, color=TINTA_SECUNDARIA, labelpad=2)
    estilo(ax)
    guardar(fig, salida)


# A4 · Crecimiento vs inflación mundial ---------------------------------------------------------
def mundo(lang, d, salida):
    t = T[lang]
    filas = d["mundo"]
    fig = figura(MEDIO, 2.3)
    ax = fig.add_axes([0.08, 0.11, 0.9, 0.78])
    x = np.arange(len(filas))
    for i, (anio, g, inf) in enumerate(filas):
        proy = anio.endswith("P")
        ax.bar(i, g, width=0.55, color=CELESTE if not proy else "white", edgecolor=CELESTE,
               hatch="////" if proy else None, linewidth=0.8)
        ax.text(i, g / 2, f"{num(g, lang)}%", ha="center", va="center", fontsize=5.2, fontweight="bold",
                color=NEGRO)
    infl = [f[2] for f in filas]
    ax.plot(x, infl, color=ROJO, linewidth=1.4, marker="o", markersize=3, markeredgecolor="white",
            markeredgewidth=0.5, zorder=3)
    for i, v in enumerate(infl):
        ax.text(i, v + 0.45, f"{num(v, lang)}%", ha="center", va="bottom", fontsize=5.2, fontweight="bold",
                color=ROJO)
    ax.set_xticks(x, [f[0].replace("P", "P") for f in filas])
    ax.set_ylim(0, 10.5)
    ax.set_yticks([])
    estilo(ax, grilla=None)
    fig.legend(handles=[Patch(facecolor=CELESTE, label=t["pib"]),
                        plt.Line2D([], [], color=ROJO, marker="o", markersize=3, label=t["ipc"]),
                        Patch(facecolor="white", edgecolor=CELESTE, hatch="////", label=t["proy"])],
               loc="upper center", bbox_to_anchor=(0.53, 1.0), ncol=3, frameon=False, fontsize=5,
               handlelength=1.4, columnspacing=1.2)
    guardar(fig, salida)


# A5 · Chile: PIB (barras) y IPC vs TPM (líneas) en dos paneles -----------------------------------
def chile(lang, d, salida):
    t = T[lang]
    filas = d["chile"]
    anios = [f[0] for f in filas]
    x = np.arange(len(filas))
    fig = figura(MEDIO, 2.2)
    ax1 = fig.add_axes([0.03, 0.12, 0.4, 0.7])
    for i, (anio, g, _, _) in enumerate(filas):
        proy = anio.endswith("P")
        ax1.bar(i, g, width=0.6, color=AZUL if not proy else "white", edgecolor=AZUL,
                hatch="////" if proy else None, linewidth=0.8)
        ax1.text(i, g + 0.1, f"{num(g, lang)}%", ha="center", va="bottom", fontsize=4.8, fontweight="bold",
                 color=AZUL)
    ax1.set_xticks(x, [a[2:] for a in anios])
    ax1.set_ylim(0, 3.2)
    ax1.set_yticks([])
    estilo(ax1, grilla=None)
    ax1.set_title(t["pib_cl"], fontsize=5.5, fontweight="bold", color=AZUL, loc="left", pad=3)

    ax2 = fig.add_axes([0.52, 0.12, 0.36, 0.7])
    for k, color, nombre in ((2, ROJO, t["ipc_cl"]), (3, AMARILLO, t["tpm"])):
        serie = [f[k] for f in filas]
        tinta = color if color != AMARILLO else "#b58800"
        ax2.plot(x, serie, color=color, linewidth=1.3, marker="o", markersize=2.6, markeredgecolor="white",
                 markeredgewidth=0.5, zorder=3)
        dec = 2 if k == 3 else 1
        ax2.text(-0.15, serie[0], f"{num(serie[0], lang, dec)}%", ha="right", va="center", fontsize=4.7,
                 fontweight="bold", color=tinta)
        ax2.text(len(x) - 1 + 0.15, serie[-1] + (0.55 if k == 3 else -0.55),
                 f"{num(serie[-1], lang, dec)}%", ha="left", va="center", fontsize=4.7, fontweight="bold", color=tinta)
    ax2.set_xticks(x, [a[2:] for a in anios])
    ax2.set_xlim(-0.6, len(x) - 0.4)
    ax2.set_ylim(0, 14)
    ax2.set_yticks([])
    estilo(ax2, grilla=None)
    ax2.set_title(f"{t['ipc_cl']} · {t['tpm']}", fontsize=5.5, fontweight="bold", color=AZUL, loc="left", pad=3)
    fig.legend(handles=[Patch(facecolor="white", edgecolor=AZUL, hatch="////", label=t["proy"] + " 2026"),
                        plt.Line2D([], [], color=ROJO, marker="o", markersize=3, label=t["ipc_cl"]),
                        plt.Line2D([], [], color=AMARILLO, marker="o", markersize=3, label=t["tpm"])],
               loc="upper center", bbox_to_anchor=(0.5, 1.0), ncol=3, frameon=False, fontsize=5,
               handlelength=1.4, columnspacing=1.2)
    guardar(fig, salida)


if __name__ == "__main__":
    ruta = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "modelos" / "macro_2022_2026.xlsx"
    salida = Path(sys.argv[2]) if len(sys.argv) > 2 else RAIZ / "graficos" / "figuras"
    salida.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    d = leer(ruta)
    for lang in ("es", "en"):
        mapa(lang, salida / f"{lang}_A0_mapa.png")
        brent(lang, d, salida / f"{lang}_A1_brent.png")
        tasas(lang, d, salida / f"{lang}_A2_tasas.png")
        spreads(lang, d, salida / f"{lang}_A3_spreads.png")
        mundo(lang, d, salida / f"{lang}_A4_mundo.png")
        chile(lang, d, salida / f"{lang}_A5_chile.png")
    print(f"Figuras guardadas en {salida}")
