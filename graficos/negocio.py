"""Figuras de la sección Business Description y sus anexos, en español e inglés.

Lee Excel_bussines_para_armar_graficos.xlsx:
  - N1: dos donas, GLA por país (2T26) e ingresos por fuente (UDM jun-26).
  - NA: activos Tier A por GLA (miles de m²), coloreados por país.
  - NB: tiendas por departamento, % del GLA total, 2020–2T26.
Al ancho de la columna de figuras del informe. Sin título ni fuente en la imagen: van en el Word
(ver negocio/pies_de_figura.md). Además genera una copia con título y fuente dentro de la imagen
(sufijo _titulado), con el formato de la referencia del Excel, para compartir fuera del Word.

Uso:
    python graficos/negocio.py <Excel_bussines_para_armar_graficos.xlsx> [carpeta_salida]
"""
import math
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import openpyxl
from matplotlib.patches import Rectangle

from pagina_riesgos import num, tinte
from sensibilidad_wacc import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA

AZUL = "#14284B"
ANCHO = 3550 / 1440  # columna de figuras del informe (pulgadas)
DPI = 400
ALTOS = {"N1": 3.4, "NA": 2.45, "NB": 1.75}
COLOR_PAIS = {"Chile": ROJO, "Perú": CELESTE, "Colombia": AMARILLO}
COLOR_FUENTE = {"Arriendo": AZUL, "Parking": CELESTE, "Servicios comerciales y otros": AMARILLO,
                "Publicidad": tinte(ROJO, 0.55)}
TEXTO_CLARO = {AZUL, ROJO}  # rellenos donde la etiqueta va en blanco

T = {
    "es": {"gla": "GLA por\npaís", "ing": "Ingresos\npor fuente", "eje_a": "GLA (miles de m²)",
           "nombres": {"Servicios comerciales y otros": "Servicios\ncomerciales y otros"}},
    "en": {"gla": "GLA by\ncountry", "ing": "Revenue\nby source", "eje_a": "GLA (thousand m²)",
           "nombres": {"Perú": "Peru", "Arriendo": "Rent", "Parking": "Parking",
                       "Servicios comerciales y otros": "Commercial\nservices & other",
                       "Publicidad": "Advertising"}},
}


def leer(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    filas = lambda ws, c0, c1: [r[c0:c1] for r in ws.iter_rows(values_only=True)
                                if r[c0] is not None and isinstance(r[c1 - 1], (int, float))]
    donas = filas(wb["Grafico 1, 2 tortas"], 1, 4)
    gla = [(c, v) for d, c, v in donas if d.startswith("A")]
    ingresos = [(c, v) for d, c, v in donas if d.startswith("B")]
    activos = filas(wb["Graf 2. Anexo 1"], 2, 5)
    tiendas = [(str(p), v) for p, v in filas(wb["Graf 3. Anexo 2"], 2, 4)]
    assert abs(sum(v for _, v in gla) - 100) < 0.05 and abs(sum(v for _, v in ingresos) - 100) < 0.05
    return gla, ingresos, activos, tiendas


def figura(clave):
    fig = plt.figure(figsize=(ANCHO, ALTOS[clave]))
    fig.patch.set_facecolor("white")
    return fig


def guardar(fig, salida):
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


def limpiar(ax):
    for lado in ax.spines.values():
        lado.set_visible(False)


def nombre(c, lang):
    return T[lang]["nombres"].get(c, c)


# N1: dos donas ----------------------------------------------------------------------------------
def dona(fig, rect, datos, colores, centro, lang):
    ax = fig.add_axes(rect)
    cuñas, _ = ax.pie([v for _, v in datos], colors=colores, startangle=90, counterclock=False,
                      wedgeprops={"width": 0.46, "edgecolor": "white", "linewidth": 1.2})
    chicas = [k for k, (_, v) in enumerate(datos) if v < 8]
    for k, (cuña, (c, v), color) in enumerate(zip(cuñas, datos, colores)):
        ang = math.radians((cuña.theta1 + cuña.theta2) / 2)
        fuera = k in chicas
        if fuera and len(chicas) > 1:  # cuñas chicas contiguas: abre las etiquetas en abanico
            ang += math.radians(14 * (chicas.index(k) - (len(chicas) - 1) / 2) * -1)
        r = 1.2 if fuera else 0.78
        ax.text(r * math.cos(ang), r * math.sin(ang), num(v, lang, 1) + "%", ha="center",
                va="center", fontsize=4.9, fontweight="bold",
                color="white" if (color in TEXTO_CLARO and not fuera) else NEGRO)
    ax.text(0, 0, centro, ha="center", va="center", fontsize=6.4, fontweight="bold", color=AZUL,
            linespacing=1.15)
    ax.set_xlim(-1.32, 1.32)
    ax.set_ylim(-1.32, 1.32)
    ax.set_aspect("equal")
    return ax


def leyenda(fig, x, y_centro, items, lang):
    paso = 0.045
    y0 = y_centro + paso * (len(items) - 1) / 2
    for k, (c, color) in enumerate(items):
        y = y0 - k * paso
        fig.add_artist(Rectangle((x, y - 0.009), 0.045, 0.018, color=color, transform=fig.transFigure))
        fig.text(x + 0.06, y, nombre(c, lang), fontsize=5.3, color=NEGRO, va="center",
                 linespacing=1.05)


def n1(lang, gla, ingresos, salida):
    t = T[lang]
    fig = figura("N1")
    col_g = [COLOR_PAIS[c] for c, _ in gla]
    col_i = [COLOR_FUENTE[c] for c, _ in ingresos]
    dona(fig, [0.0, 0.52, 0.6, 0.46], gla, col_g, t["gla"], lang)
    dona(fig, [0.0, 0.03, 0.6, 0.46], ingresos, col_i, t["ing"], lang)
    leyenda(fig, 0.6, 0.75, list(zip([c for c, _ in gla], col_g)), lang)
    leyenda(fig, 0.6, 0.26, list(zip([c for c, _ in ingresos], col_i)), lang)
    guardar(fig, salida)


# NA: activos Tier A -----------------------------------------------------------------------------
def na(lang, activos, salida):
    t = T[lang]
    fig = figura("NA")
    ax = fig.add_axes([0.24, 0.17, 0.66, 0.81])
    for y, (activo, pais, v) in enumerate(reversed(activos)):
        ax.barh(y, v, height=0.68, color=COLOR_PAIS[pais])
        ax.text(v + 2.5, y, num(v, lang, 1), va="center", fontsize=5.2, color=NEGRO)
        ax.text(-4, y, activo, va="center", ha="right", fontsize=5.4, color=NEGRO)
    ax.set_yticks([])
    ax.set_xlim(0, 215)
    ax.set_ylim(-0.6, len(activos) - 0.4)
    ax.set_xticks([0, 50, 100, 150, 200])
    ax.xaxis.set_major_formatter(lambda x, _: num(x, lang))
    ax.tick_params(axis="x", length=0, colors=TINTA_SECUNDARIA, labelsize=4.8, pad=1.5)
    ax.grid(axis="x", color=GRILLA, linewidth=0.4)
    ax.set_axisbelow(True)
    limpiar(ax)
    ax.set_xlabel(t["eje_a"], fontsize=5, color=TINTA_SECUNDARIA, labelpad=2)
    for k, pais in enumerate(COLOR_PAIS):
        x = 0.24 + k * 0.24
        fig.add_artist(Rectangle((x, 0.025), 0.035, 0.028, color=COLOR_PAIS[pais],
                                 transform=fig.transFigure))
        fig.text(x + 0.05, 0.039, nombre(pais, lang), fontsize=5.2, color=NEGRO, va="center")
    guardar(fig, salida)


# NB: tiendas por departamento -------------------------------------------------------------------
def nb(lang, tiendas, salida):
    fig = figura("NB")
    ax = fig.add_axes([0.06, 0.15, 0.9, 0.8])
    x = range(len(tiendas))
    v = [p[1] for p in tiendas]
    ax.plot(x, v, color=ROJO, linewidth=1.3, marker="o", markersize=2.6, zorder=3)
    for k, (p, y) in enumerate(tiendas):
        abajo = y < 21  # tras la caída de 2023 la etiqueta va bajo el punto (no cruza la línea)
        ax.text(k, y - 0.9 if abajo else y + 0.9, num(y, lang, 1) + "%", ha="center",
                va="top" if abajo else "bottom", fontsize=5,
                fontweight="bold" if k in (0, len(tiendas) - 1) else "normal", color=NEGRO)
    etiquetas = [p if lang == "en" else p.replace("Q", "T") for p, _ in tiendas]
    ax.set_xticks(list(x), etiquetas)
    ax.set_yticks([])
    ax.set_ylim(16, 29)
    ax.set_xlim(-0.4, len(tiendas) - 0.6)
    ax.tick_params(axis="x", length=0, colors=TINTA_SECUNDARIA, labelsize=5, pad=2)
    ax.axhline(16, color=GRILLA, linewidth=0.6)
    limpiar(ax)
    guardar(fig, salida)


PIES = {
    "es": {"N1": ("Figura 1:", "GLA por país (2T26) e ingresos por fuente (UDM a jun-2026)",
                  "Fuente: Plaza S.A., Release de Resultados 2T26 (p. 19) y Presentación Institucional 2T26 (p. 16)."),
           "NA": ("Anexo A:", "Activos Tier A por GLA, 2T26 (miles de m²)",
                  "Fuente: Plaza S.A., Presentación Institucional 2T26 (p. 12). Tier A = 46,4% del GLA total."),
           "NB": ("Anexo B:", "Tiendas por departamento, % del GLA total",
                  "Fuente: Plaza S.A., Presentación Institucional 2T26 (p. 17).")},
    "en": {"N1": ("Figure 1:", "GLA by country (2Q26) and revenue by source (LTM Jun-2026)",
                  "Source: Plaza S.A., 2Q26 Earnings Release (p. 19) and 2Q26 Institutional Presentation (p. 16)."),
           "NA": ("Appendix A:", "Tier A assets by GLA, 2Q26 (thousand m²)",
                  "Source: Plaza S.A., 2Q26 Institutional Presentation (p. 12). Tier A = 46.4% of total GLA."),
           "NB": ("Appendix B:", "Department stores, % of total GLA",
                  "Source: Plaza S.A., 2Q26 Institutional Presentation (p. 17).")},
}


def titulado(png, clave, lang, salida):
    """Copia de la figura con título (arriba) y fuente (abajo) dentro de la imagen."""
    import textwrap
    n, titulo, fuente = PIES[lang][clave]
    img = plt.imread(png)
    alto_img = ANCHO * img.shape[0] / img.shape[1]
    lineas_t = textwrap.wrap(f"{n} {titulo}", 44)
    lineas_f = textwrap.wrap(fuente, 62)
    cab, pie = 0.12 + 0.13 * len(lineas_t), 0.08 + 0.1 * len(lineas_f)
    h = alto_img + cab + pie
    fig = plt.figure(figsize=(ANCHO, h))
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, pie / h, 1, alto_img / h])
    ax.imshow(img)
    ax.axis("off")
    y = 1 - 0.06 / h
    for k, linea in enumerate(lineas_t):
        if k == 0:
            t = fig.text(0.03, y, n, fontsize=7, fontweight="bold", color=AZUL, va="top")
            ancho_n = t.get_window_extent(fig.canvas.get_renderer()).width / fig.bbox.width
            fig.text(0.03 + ancho_n + 0.012, y, linea[len(n):].strip(), fontsize=7, color=AZUL, va="top")
        else:
            fig.text(0.03, y, linea, fontsize=7, color=AZUL, va="top")
        y -= 0.13 / h
    for k, linea in enumerate(lineas_f):
        fig.text(0.03, (pie - 0.06 - k * 0.1) / h, linea, fontsize=5, style="italic",
                 color=TINTA_SECUNDARIA, va="top")
    guardar(fig, salida)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    salida = Path(sys.argv[2] if len(sys.argv) > 2 else Path(__file__).parent / "informe" / "negocio")
    salida.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    gla, ingresos, activos, tiendas = leer(sys.argv[1])
    for lang in ("es", "en"):
        n1(lang, gla, ingresos, salida / f"{lang}_N1_gla_ingresos.png")
        na(lang, activos, salida / f"{lang}_NA_activos_tier_a.png")
        nb(lang, tiendas, salida / f"{lang}_NB_tiendas_departamento.png")
        for clave, archivo in (("N1", "N1_gla_ingresos"), ("NA", "NA_activos_tier_a"),
                               ("NB", "NB_tiendas_departamento")):
            titulado(salida / f"{lang}_{archivo}.png", clave, lang,
                     salida / f"{lang}_{archivo}_titulado.png")
    print(f"Figuras guardadas en {salida}")
