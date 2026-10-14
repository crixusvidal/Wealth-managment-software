"""Figuras de la sección Financial Analysis, en español e inglés, al ancho de la columna del informe.

Datos del borrador del equipo "Mallplaza_Financial_Analysis_2T26_v2" (EEFF CMF, release 2T26 y
modelo de valoración v2). Sin título ni fuente en la imagen: van en el Word
(informe/analisis_financiero.js).

  FA2  Puente utilidad → FFO 2025 (cascada)
  FA3  Crecimiento 2025 por perímetro y 1S26 descompuesto en real / inflación / moneda
  FA4  Puente EBITDA → caja retenida 2025 (cascada)
  FA5  LTV dic-24 → jun-26 (cascada)
  FA6  Vencimientos de bonos en UF (valor libro dic-25) y tasa efectiva
  FA7  Capex 2023A–2031E por tipo y DFN/EBITDA

Uso:
    python graficos/analisis_financiero.py [carpeta_salida]
"""
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Rectangle

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional y rutas compartidas
import estilo  # noqa: E402,F401  (registra las carpetas graficos/ de todas las secciones)
from estilo import num, tinte
from estilo import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA

AZUL = "#14284B"
GRIS = "#9a9a9a"
ANCHO = 3550 / 1440
DPI = 400
ALTOS = {"FA2": 1.85, "FA3": 1.95, "FA4": 1.75, "FA5": 2.0, "FA6": 2.2, "FA7": 2.95}

T = {
    "es": {
        "fa2": ["Utilidad\n(consol.)", "Revalori-\nzación", "Otros\ngastos op.", "D&A", "Reajuste\nUF deuda", "Impuesto\ndiferido", "FFO\n(consol.)"],
        "fa3_izq": "2025", "fa3_mismo": "Mismo\nperímetro", "fa3_peru": "Compra\nPerú",
        "fa3_seg": ["Chile", "Perú", "Colombia", "Consolidado"], "fa3_der": "1S26 vs 1S25",
        "real": "Real", "infl": "Inflación / UF", "fx": "Moneda",
        "fa4": ["EBITDA", "Impuestos,\nCT y otros", "Flujo\noperativo", "Capex", "Intereses\nnetos", "FCF", "Dividendos", "Caja\nretenida"],
        "fa4_nota": ["Cobertura de dividendos: 1,5x · 2,6x · 1,7x", "Flujo op. / EBITDA: 79% · 84% · 95%", "(2024 · 2025 · 1S26)"],
        "fa5": ["LTV\ndic-24", "Deuda y\nadiciones", "Revalori-\nzación 2025", "LTV\ndic-25", "1S26", "LTV\njun-26"],
        "fa5_nota": ["DFN/EBITDA: 3,78x → 2,49x → 2,36x", "EBITDA/costo fin.: 6,8x → 7,7x → 8,4x"],
        "fa6_eje": "Valor libro dic-25 (CLP bn)", "fa6_s": "Serie S",
        "fa7": ["Real", "Pipeline anunciado", "Mantención + TI", "Expansión post-2028"],
        "fa7_eje": "Capex (CLP bn)", "fa7_dfn": "DFN / EBITDA (x)",
    },
    "en": {
        "fa2": ["Net inc.\n(consol.)", "Fair-value\ngains", "Other op.\nexpenses", "D&A", "UF index.\nof debt", "Deferred\ntax", "FFO\n(consol.)"],
        "fa3_izq": "FY2025", "fa3_mismo": "Same\nperimeter", "fa3_peru": "Peru\nacquisition",
        "fa3_seg": ["Chile", "Peru", "Colombia", "Consolidated"], "fa3_der": "1H26 vs 1H25",
        "real": "Real", "infl": "Inflation / UF", "fx": "Currency",
        "fa4": ["EBITDA", "Tax, WC\n& other", "Op. cash\nflow", "Capex", "Net\ninterest", "FCF", "Divs.\npaid", "Cash\nretained"],
        "fa4_nota": ["Dividend cover: 1.5x · 2.6x · 1.7x", "OCF / EBITDA: 79% · 84% · 95%", "(FY24 · FY25 · 1H26)"],
        "fa5": ["LTV\nDec-24", "Debt &\nadditions", "2025\nrevaluation", "LTV\nDec-25", "1H26", "LTV\nJun-26"],
        "fa5_nota": ["Net debt/EBITDA: 3.78x → 2.49x → 2.36x", "EBITDA/fin. costs: 6.8x → 7.7x → 8.4x"],
        "fa6_eje": "Book value Dec-25 (CLP bn)", "fa6_s": "Series S",
        "fa7": ["Actual", "Announced pipeline", "Maintenance + IT", "Post-2028 expansion"],
        "fa7_eje": "Capex (CLP bn)", "fa7_dfn": "Net debt / EBITDA (x)",
    },
}

# Datos del borrador (CLP bn salvo indicación)
FA2 = [1441.0, -1448.9, 3.5, 3.6, 43.8, 365.1]           # → FFO 408,1
FA3_2025 = (10.4, 21.8)                                   # pp: mismo perímetro, compra Perú (total 32,2%)
FA3_1S26 = {  # real, inflación, moneda (pp aditivos, del split multiplicativo del borrador) → total
    0: (2.8, 3.3, 0.0, 6.1), 1: (4.5, 3.4, 1.0, 8.9), 2: (1.1, 6.2, 8.0, 15.3), 3: (3.0, 3.4, 1.0, 7.4)}
FA4 = [526.2, -85.1, None, -103.6, -57.4, None, -108.5, None]  # None = subtotal
FA5 = [26.9, -3.1, -5.0, None, -1.2, None]
FA6 = [(2027, 40, 3.32), (2028, 59, 4.21), (2029, 161, 1.32), (2030, 6, 4.72), (2031, 198, 4.00),
       (2032, 89, 4.76), (2033, 155, 4.06), (2034, 119, 3.92), (2035, 121, 3.76), (2043, 129, 3.66),
       (2045, 80, 2.09), (2047, 120, 3.44)]
FA7_AÑOS = ["2023A", "2024A", "2025A", "2026E", "2027E", "2028E", "2029E", "2030E", "2031E"]
FA7 = {  # real, pipeline, mantención + TI, expansión post-2028
    "real": [119.1, 76.3, 103.6, 84.2, 0, 0, 0, 0, 0],
    "pipe": [0, 0, 0, 105.6, 198.5, 106.4, 0, 0, 0],
    "mant": [0, 0, 0, 19.4, 42.2, 45.1, 47.7, 49.7, 51.7],
    "exp": [0, 0, 0, 0, 46.3, 95.3, 98.2, 101.1, 104.2]}
FA7_DFN = [3.30, 3.78, 2.49, 2.46, 2.67, 2.88, 3.09, 3.29, 3.50]


def figura(clave):
    fig = plt.figure(figsize=(ANCHO, ALTOS[clave]))
    fig.patch.set_facecolor("white")
    return fig


def limpiar(ax, eje_y=False):
    for k, lado in ax.spines.items():
        lado.set_visible(k == "bottom")
    ax.spines["bottom"].set_color(GRILLA)
    ax.tick_params(length=0, colors=TINTA_SECUNDARIA, labelsize=4.4, pad=1.5)
    if not eje_y:
        ax.set_yticks([])


def guardar(fig, salida):
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


def cascada(ax, pasos, colores, etiquetas, fmt, ancho=0.62):
    """Cascada: pasos con None = subtotal acumulado. Devuelve la lista de niveles."""
    nivel, niveles = 0.0, []
    for x, (v, color) in enumerate(zip(pasos, colores)):
        if v is None or x == 0:
            total = nivel if v is None else v
            ax.bar(x, total, width=ancho, color=color, zorder=3)
            ax.text(x, total, fmt(total), ha="center", va="bottom", fontsize=4.6, fontweight="bold",
                    color=NEGRO)
            nivel = total
        else:
            base = nivel + min(v, 0)
            ax.bar(x, abs(v), bottom=base, width=ancho, color=color, zorder=3)
            ax.text(x, nivel + max(v, 0), fmt(v, signo=True), ha="center", va="bottom",
                    fontsize=4.4, color=NEGRO)
            if x + 1 < len(pasos):
                ax.plot([x + ancho / 2, x + 1 - ancho / 2], [nivel + v] * 2, color=NEGRO,
                        linewidth=0.4, zorder=2)
            nivel += v
        niveles.append(nivel)
    ax.set_xticks(range(len(pasos)), etiquetas)
    return niveles


# FA2 -------------------------------------------------------------------------------------------
def fa2(lang, salida):
    t = T[lang]
    fig = figura("FA2")
    ax = fig.add_axes([0.02, 0.2, 0.96, 0.76])
    pasos = FA2 + [None]
    colores = [AZUL] + [ROJO if v < 0 else CELESTE for v in FA2[1:]] + [AZUL]
    fmt = lambda v, signo=False: ("+" if signo and v > 0 else "") + num(v, lang, 1)
    cascada(ax, pasos, colores, t["fa2"], fmt)
    ax.set_ylim(0, 1650)
    limpiar(ax)
    guardar(fig, salida)


# FA3 -------------------------------------------------------------------------------------------
def fa3(lang, salida):
    t = T[lang]
    fig = figura("FA3")
    # Izquierda: 2025 por perímetro
    ax = fig.add_axes([0.03, 0.2, 0.16, 0.68])
    mismo, peru = FA3_2025
    ax.bar(0, mismo, width=0.7, color=AZUL)
    ax.bar(0, peru, bottom=mismo, width=0.7, color=ROJO)
    ax.text(0, mismo / 2, f"{t['fa3_mismo']}\n+{num(mismo, lang, 1)} pp", ha="center", va="center",
            fontsize=3.7, color="white", linespacing=1.05)
    ax.text(0, mismo + peru / 2, f"{t['fa3_peru']}\n+{num(peru, lang, 1)} pp", ha="center",
            va="center", fontsize=3.7, color="white", linespacing=1.05)
    ax.text(0, mismo + peru + 0.8, num(mismo + peru, lang, 1) + "%", ha="center", va="bottom",
            fontsize=4.8, fontweight="bold", color=NEGRO)
    ax.set_xlim(-0.5, 0.5)
    ax.set_ylim(0, 36)
    ax.set_xticks([0], [t["fa3_izq"]])
    limpiar(ax)
    # Derecha: 1S26 por segmento
    ax = fig.add_axes([0.38, 0.2, 0.58, 0.68])
    for y, k in enumerate(reversed(range(4))):
        real, infl, fx, total = FA3_1S26[k]
        izq = 0
        for v, color in ((real, AZUL), (infl, CELESTE), (fx, AMARILLO)):
            if v:
                ax.barh(y, v, left=izq, height=0.6, color=color)
                izq += v
        if real >= 2:
            ax.text(real / 2, y, num(real, lang, 1) + "%", ha="center", va="center", fontsize=3.9,
                    color="white", fontweight="bold")
        else:  # barra real muy angosta: etiqueta sobre la barra
            ax.text(real / 2, y + 0.42, num(real, lang, 1) + "%", ha="center", va="bottom",
                    fontsize=3.9, color=AZUL, fontweight="bold")
        ax.text(total + 0.3, y, num(total, lang, 1) + "%", ha="left", va="center", fontsize=4.5,
                fontweight="bold", color=NEGRO)
        ax.text(-0.4, y, t["fa3_seg"][k], ha="right", va="center", fontsize=4.5,
                fontweight="bold" if k == 3 else "normal", color=NEGRO)
    ax.set_xlim(0, 18)
    ax.set_ylim(-0.6, 3.6)
    ax.set_xticks([])
    limpiar(ax)
    ax.spines["bottom"].set_visible(False)
    fig.text(0.38, 0.93, t["fa3_der"], fontsize=4.5, color=TINTA_SECUNDARIA, va="center")
    for x, color, texto in ((0.38, AZUL, t["real"]), (0.55, CELESTE, t["infl"]), (0.82, AMARILLO, t["fx"])):
        fig.add_artist(Rectangle((x, 0.06), 0.025, 0.035, color=color, transform=fig.transFigure))
        fig.text(x + 0.035, 0.077, texto, fontsize=4.3, color=NEGRO, va="center")
    guardar(fig, salida)


# FA4 -------------------------------------------------------------------------------------------
def fa4(lang, salida):
    t = T[lang]
    fig = figura("FA4")
    ax = fig.add_axes([0.02, 0.2, 0.96, 0.58])
    colores = [AZUL, ROJO, AZUL, ROJO, ROJO, AZUL, ROJO, CELESTE]
    fmt = lambda v, signo=False: f"({num(-v, lang, 1)})" if v < 0 else num(v, lang, 1)
    cascada(ax, FA4, colores, t["fa4"], fmt)
    ax.set_ylim(0, 600)
    limpiar(ax)
    ax.tick_params(axis="x", labelsize=3.9)
    for i, linea in enumerate(t["fa4_nota"]):
        fig.text(0.98, 0.95 - i * 0.065, linea, ha="right", va="center", fontsize=4.2,
                 fontweight="bold" if i == 0 else "normal",
                 color=NEGRO if i < 2 else TINTA_SECUNDARIA)
    guardar(fig, salida)


# FA5 -------------------------------------------------------------------------------------------
def fa5(lang, salida):
    t = T[lang]
    fig = figura("FA5")
    ax = fig.add_axes([0.02, 0.22, 0.96, 0.55])
    colores = [AZUL, CELESTE, ROJO, AZUL, CELESTE, AZUL]
    fmt = lambda v, signo=False: (num(v, lang, 1) + " pp") if signo else (num(v, lang, 1) + "%")
    cascada(ax, FA5, colores, t["fa5"], fmt)
    ax.set_ylim(0, 31)
    limpiar(ax)
    for i, linea in enumerate(t["fa5_nota"]):
        fig.text(0.98, 0.95 - i * 0.07, linea, ha="right", va="center", fontsize=4.3, color=NEGRO)
    guardar(fig, salida)


# FA6 -------------------------------------------------------------------------------------------
def fa6(lang, salida):
    t = T[lang]
    fig = figura("FA6")
    ax = fig.add_axes([0.02, 0.15, 0.96, 0.8])
    for x, (año, v, tasa) in enumerate(FA6):
        serie_s = año == 2029
        ax.bar(x, v, width=0.66, color=ROJO if serie_s else AZUL, zorder=3)
        ax.text(x, v + 4, num(v, lang), ha="center", va="bottom", fontsize=4.2, fontweight="bold",
                color=NEGRO)
        ax.text(x, v + 19, num(tasa, lang, 2) + "%", ha="center", va="bottom", fontsize=3.4,
                color=ROJO if serie_s else TINTA_SECUNDARIA)
    s = [i for i, f in enumerate(FA6) if f[0] == 2029][0]
    ax.text(s, FA6[s][1] + 36, t["fa6_s"], ha="center", va="bottom", fontsize=3.8,
            fontweight="bold", color=ROJO)
    ax.set_xticks(range(len(FA6)), [str(f[0]) for f in FA6])
    ax.set_ylim(0, 240)
    limpiar(ax)
    ax.tick_params(axis="x", labelsize=3.9)
    guardar(fig, salida)


# FA7 -------------------------------------------------------------------------------------------
def fa7(lang, salida):
    t = T[lang]
    fig = figura("FA7")
    ax = fig.add_axes([0.02, 0.42, 0.96, 0.44])
    x = np.arange(len(FA7_AÑOS))
    base = np.zeros(len(x))
    for clave, color in (("real", AZUL), ("pipe", ROJO), ("mant", CELESTE), ("exp", AMARILLO)):
        v = np.array(FA7[clave])
        ax.bar(x, v, bottom=base, width=0.66, color=color, zorder=3)
        base += v
    for xi, total in zip(x, base):
        ax.text(xi, total + 4, num(total, lang), ha="center", va="bottom", fontsize=4.3,
                fontweight="bold", color=NEGRO)
    ax.set_ylim(0, 330)
    ax.set_xticks([])
    limpiar(ax)
    ax.text(-0.45, 320, t["fa7_eje"], fontsize=4.2, color=TINTA_SECUNDARIA, va="top")
    # DFN / EBITDA
    ax2 = fig.add_axes([0.02, 0.1, 0.96, 0.26])
    ax2.plot(x, FA7_DFN, color=AZUL, linewidth=1, marker="o", markersize=2.2, zorder=3)
    for i, (xi, v) in enumerate(zip(x, FA7_DFN)):
        abajo = i > 0 and v < FA7_DFN[i - 1]  # bajando hacia el punto: etiqueta abajo
        ax2.text(xi, v - 0.22 if abajo else v + 0.18, num(v, lang, 2) + "x", ha="center",
                 va="top" if abajo else "bottom", fontsize=3.9, color=NEGRO)
    ax2.set_ylim(1.7, 4.9)
    ax2.set_xlim(ax.get_xlim())
    ax2.set_xticks(x, FA7_AÑOS)
    limpiar(ax2)
    ax2.tick_params(axis="x", labelsize=4)
    ax2.text(-0.45, 4.88, t["fa7_dfn"], fontsize=4.2, color=TINTA_SECUNDARIA, va="top")
    for i, (color, texto) in enumerate(zip((AZUL, ROJO, CELESTE, AMARILLO), t["fa7"])):
        xx, yy = 0.03 + (i % 2) * 0.5, 0.965 - (i // 2) * 0.04
        fig.add_artist(Rectangle((xx, yy - 0.013), 0.022, 0.026, color=color, transform=fig.transFigure))
        fig.text(xx + 0.03, yy, texto, fontsize=4.1, color=NEGRO, va="center")
    guardar(fig, salida)


def controles():
    assert abs(sum(FA2) - 408.1) < 0.05
    assert all(abs(sum(v[:3]) - v[3]) < 0.05 for v in FA3_1S26.values())
    assert abs(526.2 - 85.1 - 441.1) < 0.05 and abs(441.1 - 103.6 - 57.4 - 280.1) < 0.05
    assert abs(280.1 - 108.5 - 171.6) < 0.05
    assert abs(26.9 - 3.1 - 5.0 - 18.8) < 0.05 and abs(18.8 - 1.2 - 17.6) < 0.05
    assert abs(sum(f[1] for f in FA6) - 1277.7) < 1
    totales = [sum(FA7[k][i] for k in FA7) for i in range(len(FA7_AÑOS))]
    assert [round(v) for v in totales] == [119, 76, 104, 209, 287, 247, 146, 151, 156], totales


if __name__ == "__main__":
    salida = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent / "figuras")
    salida.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    controles()
    for lang in ("es", "en"):
        fa2(lang, salida / f"{lang}_FA2_utilidad_ffo.png")
        fa3(lang, salida / f"{lang}_FA3_crecimiento.png")
        fa4(lang, salida / f"{lang}_FA4_caja.png")
        fa5(lang, salida / f"{lang}_FA5_ltv.png")
        fa6(lang, salida / f"{lang}_FA6_vencimientos.png")
        fa7(lang, salida / f"{lang}_FA7_capex.png")
    print(f"Figuras guardadas en {salida}")
