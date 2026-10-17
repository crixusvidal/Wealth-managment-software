"""Figuras de la sección ESG del informe y de su anexo, en español e inglés.

Datos: Memoria Mallplaza 2025 y Bloomberg (puntaje ESG), según el material del equipo
(Grafiquitos_definitivos.docx y la planilla de sanciones). Sin título ni fuente en la imagen:
van en el Word (informe/esg.js). Paleta Mallplaza: rojo, celeste, amarillo, negro.

Uso:
    python graficos/esg.py [carpeta_salida]
"""
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Circle, FancyBboxPatch, Patch, Rectangle, Wedge

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional y rutas compartidas
import estilo  # noqa: E402,F401  (registra las carpetas graficos/ de todas las secciones)
from estilo import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA

AZUL = "#14284B"
GRIS = "#9a9a9a"
CLARO = "#eef1f6"
COLUMNA = 3550 / 1440                  # columna de figuras del informe (pulgadas)
COMPLETO = 10666 / 1440                # ancho completo (anexo)
DPI = 400

# ---- Datos (Memoria 2025 / Bloomberg) -----------------------------------------------------------
BLOOMBERG = [("ESG", 3.39), ("amb", 2.64), ("soc", 5.62), ("gob", 3.78)]
METAS = [  # (clave, avance %, meta)
    ("res", 94.8), ("agua", 100.9), ("hcfc", 14.9), ("ener", 35.5), ("elec", 89.3)]
ENERGIA = {"renovable": 339_850.6, "no_renovable": 34_009.2}            # MWh 2025
RESIDUOS_VALORIZADOS = 46.2                                                # % de no peligrosos
VALORIZACION = [("biogas", 40.4), ("recic", 36.1), ("compost", 23.5)]     # % de lo valorizado
NPS = [(2022, -8), (2023, 8), (2024, 11), (2025, 13)]
REMUNERACION = [("esg", 33.3), ("otros_nf", 27.7), ("fin", 39.0)]
ACCIONISTAS = [("ctrl", 53.10), ("afp", 23.80), ("may", 11.40), ("fondos", 8.25), ("otros", 3.45)]
RESIDUOS_PAIS = [("Chile", 76), ("Colombia", 16), ("Perú", 7)]
PARES_ESG = [("Cencomalls", [3.74, 3.26, 7.17, 3.41]), ("Parque Arauco", [4.04, 3.86, 5.62, 3.80])]   # general, E, S, G (Bloomberg, beta 4.1)
COLOR_PARES = [AZUL, CELESTE, AMARILLO, NEGRO]

T = {
    "es": {"pilar": {"ESG": "Puntaje ESG", "amb": "Ambiental", "soc": "Social", "gob": "Gobernanza"},
           "escala": "Puntaje Bloomberg (0–10)", "pares": ["General", "E", "S", "G"],
           "ratings": [("MSCI ESG", "BBB"), ("Sustainalytics", "14,4 · riesgo bajo"),
                       ("Dow Jones", "Best-in-Class Chile y MILA"), ("S&P Global", "Sustainability Yearbook"),
                       ("Huella Perú", "2 de 4 estrellas")],
           "metas": {"res": ("Valorización de residuos", "60% al 2025"),
                     "agua": ("Eficiencia hídrica", "−30% m³/GLA vs 2019, al 2025"),
                     "hcfc": ("Eliminación de HCFC", "100% al 2030"),
                     "ener": ("Eficiencia energética", "−22% kWh/GLA vs 2019, al 2035"),
                     "elec": ("Electricidad renovable", "100% al 2026")},
           "avance": "Avance de la meta (2025)",
           "energia": "Energía 2025", "renov": "Renovable", "norenov": "No renovable",
           "residuos": "Valorización de residuos", "valoriz": "Valorizados", "elim": "Eliminación", "reut": "Reutilización", "de_no_pel": "{p}% de los no peligrosos",
           "valor": {"biogas": "Biogás", "recic": "Reciclaje", "compost": "Compostaje"},
           "nps": "NPS de arrendatarios",
           "tiles": [("53%", "mujeres en la fuerza\nlaboral (403)"), ("42%", "mujeres en cargos\ndirectivos (44)"),
                     ("49.325", "personas impactadas\npor programas")],
           "rem": "Remuneración variable\nde ejecutivos",
           "remk": {"esg": "KPIs ESG", "otros_nf": "Otros no financieros", "fin": "Financieros"},
           "dir": "Directorio: 9 miembros", "indep": "independiente", "no_indep": "no independientes",
           "ctrl": "53,10% Falabella\n(controlador)",
           "acc": {"ctrl": "Controlador", "afp": "Fondos de pensiones", "may": "Accionistas mayoritarios",
                   "fondos": "Otros fondos", "otros": "Otros (corredoras)"},
           "pasos": ["Identificar los grupos\nde interés relevantes", "Definir el área\nresponsable",
                     "Fijar objetivos, planes\nde acción y KPIs", "Abrir canales de\ncomunicación y diálogo",
                     "Evaluar la percepción\nde la gestión sostenible"],
           "total_res": "32.425 t\nen 2025"},
    "en": {"pilar": {"ESG": "ESG score", "amb": "Environmental", "soc": "Social", "gob": "Governance"},
           "escala": "Bloomberg score (0–10)", "pares": ["General", "E", "S", "G"],
           "ratings": [("MSCI ESG", "BBB"), ("Sustainalytics", "14.4 · low risk"),
                       ("Dow Jones", "Best-in-Class Chile & MILA"), ("S&P Global", "Sustainability Yearbook"),
                       ("Huella Perú", "2 of 4 stars")],
           "metas": {"res": ("Waste recovery", "60% by 2025"),
                     "agua": ("Water efficiency", "−30% m³/GLA vs 2019, by 2025"),
                     "hcfc": ("HCFC phase-out", "100% by 2030"),
                     "ener": ("Energy efficiency", "−22% kWh/GLA vs 2019, by 2035"),
                     "elec": ("Renewable electricity", "100% by 2026")},
           "avance": "Progress towards target (2025)",
           "energia": "Energy 2025", "renov": "Renewable", "norenov": "Non-renewable",
           "residuos": "Waste recovery", "valoriz": "Recovered", "elim": "Disposal", "reut": "Reuse", "de_no_pel": "{p}% of non-hazardous",
           "valor": {"biogas": "Biogas", "recic": "Recycling", "compost": "Composting"},
           "nps": "Tenant NPS",
           "tiles": [("53%", "women in the\nworkforce (403)"), ("42%", "women in\nmanagement (44)"),
                     ("49,325", "people reached by\ncommunity programs")],
           "rem": "Executive variable\npay",
           "remk": {"esg": "ESG KPIs", "otros_nf": "Other non-financial", "fin": "Financial"},
           "dir": "Board: 9 members", "indep": "independent", "no_indep": "non-independent",
           "ctrl": "53.10% Falabella\n(controller)",
           "acc": {"ctrl": "Controller", "afp": "Pension funds", "may": "Major shareholders",
                   "fondos": "Other funds", "otros": "Others (brokers)"},
           "pasos": ["Identify relevant\nstakeholders", "Assign the\nresponsible area",
                     "Set objectives, action\nplans and KPIs", "Open communication\nand dialogue channels",
                     "Assess perception of\nsustainability management"],
           "total_res": "32,425 t\nin 2025"},
}


def num(x, lang, dec=1):
    s = f"{x:,.{dec}f}".replace("-", "−")
    return s.translate(str.maketrans(",.", ".,")) if lang == "es" else s


def figura(ancho, alto):
    fig = plt.figure(figsize=(ancho, alto))
    fig.patch.set_facecolor("white")
    return fig


def limpio(ax):
    for lado in ax.spines.values():
        lado.set_visible(False)
    ax.set_xticks([])
    ax.set_yticks([])


def guardar(fig, salida):
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


def tarjeta(fig, x, y, w, h, fondo=CLARO, borde=None):
    fig.add_artist(FancyBboxPatch((x, y), w, h, transform=fig.transFigure,
                                  boxstyle="round,pad=0,rounding_size=0.012", facecolor=fondo,
                                  edgecolor=borde or "none", linewidth=0.6))


# E1 · Puntaje Bloomberg y ratings ----------------------------------------------------------------
def scorecard(lang, salida):
    t = T[lang]
    fig = figura(COLUMNA, 2.3)
    ax = fig.add_axes([0.3, 0.7, 0.66, 0.28])
    for y, (k, v) in enumerate(reversed(BLOOMBERG)):
        total = k == "ESG"
        ax.barh(y, 10, height=0.56, color=CLARO)
        ax.barh(y, v, height=0.56, color=ROJO if total else CELESTE)
        ax.text(v + 0.15, y, num(v, lang, 2), va="center", fontsize=5.4, fontweight="bold",
                color=ROJO if total else NEGRO)
        ax.text(-0.25, y, t["pilar"][k], ha="right", va="center", fontsize=5.4,
                fontweight="bold" if total else "normal", color=NEGRO)
    ax.set_xlim(0, 10)
    ax.set_ylim(-0.5, len(BLOOMBERG) - 0.5)
    limpio(ax)
    # Pares: puntaje general y por pilar (cuatro barras chicas por empresa)
    axp = fig.add_axes([0.3, 0.5, 0.66, 0.15])
    axp.set_xlim(0, 4)
    axp.set_ylim(-0.5, len(PARES_ESG) - 0.5)
    for fila, (nombre, vals) in enumerate(PARES_ESG):
        y = len(PARES_ESG) - 1 - fila
        fig.text(0.27, 0.5 + 0.15 * (y + 0.5) / len(PARES_ESG), nombre, ha="right", va="center", fontsize=4.8, color=NEGRO)
        for j, (v, c) in enumerate(zip(vals, COLOR_PARES)):
            axp.add_patch(Rectangle((j + 0.03, y - 0.28), 0.94, 0.56, facecolor=CLARO, linewidth=0))
            axp.add_patch(Rectangle((j + 0.03, y - 0.28), 0.62 * v / 10, 0.56, facecolor=c, linewidth=0))
            axp.text(j + 0.03 + 0.62 * v / 10 + 0.03, y, num(v, lang, 2), va="center", fontsize=3.9, fontweight="bold", color=NEGRO)
    limpio(axp)
    # leyenda de los pares
    for j, (lab, c) in enumerate(zip(t["pares"], COLOR_PARES)):
        x = 0.3 + 0.66 * j / 4
        fig.add_artist(Rectangle((x, 0.455), 0.012, 0.018, transform=fig.transFigure, color=c))
        fig.text(x + 0.02, 0.464, lab, fontsize=4.0, va="center", color=TINTA_SECUNDARIA)
    fig.text(0.3, 0.425, t["escala"], fontsize=4.4, color=TINTA_SECUNDARIA, va="center")
    # Ratings: 2 columnas
    w, h, gx, gy = 0.455, 0.1, 0.03, 0.02
    for i, (agencia, nota) in enumerate(t["ratings"]):
        col, fila = i % 2, i // 2
        x = 0.03 + col * (w + gx)
        y = 0.28 - fila * (h + gy)
        if i == 4:
            x, w2 = 0.03, w * 2 + gx
        else:
            w2 = w
        tarjeta(fig, x, y, w2, h)
        fig.add_artist(Rectangle((x, y), 0.008, h, transform=fig.transFigure, color=ROJO if i == 0 else AZUL))
        fig.text(x + 0.025, y + h * 0.66, agencia, fontsize=4.4, color=TINTA_SECUNDARIA, va="center")
        fig.text(x + 0.025, y + h * 0.3, nota, fontsize=5.2, fontweight="bold", color=AZUL, va="center")
    guardar(fig, salida)


# E2 · Metas ambientales --------------------------------------------------------------------------
def metas(lang, salida):
    t = T[lang]
    fig = figura(COLUMNA, 1.75)
    ax = fig.add_axes([0.03, 0.1, 0.94, 0.88])
    n = len(METAS)
    for i, (k, v) in enumerate(METAS):
        y = n - 1 - i
        nombre, meta = t["metas"][k]
        ax.text(0, y + 0.3, nombre, fontsize=5.3, fontweight="bold", color=AZUL, va="center")
        ax.text(100, y + 0.3, meta, fontsize=4.4, color=TINTA_SECUNDARIA, va="center", ha="right")
        ax.barh(y - 0.08, 100, height=0.3, color=CLARO)
        ax.barh(y - 0.08, min(v, 100), height=0.3, color=ROJO if v >= 90 else CELESTE)
        dentro = v >= 25
        ax.text(min(v, 100) - 1.5 if dentro else v + 1.5, y - 0.08, f"{num(v, lang)}%",
                ha="right" if dentro else "left", va="center", fontsize=4.8, fontweight="bold",
                color="white" if dentro else NEGRO)
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.6, n - 0.35)
    limpio(ax)
    fig.text(0.03, 0.02, t["avance"], fontsize=4.5, color=TINTA_SECUNDARIA)
    guardar(fig, salida)


# E3 · Energía y residuos -------------------------------------------------------------------------
def dona(ax, valores, colores, ancho=0.36):
    inicio = 90
    total = sum(valores)
    for v, c in zip(valores, colores):
        ang = 360 * v / total
        ax.add_patch(Wedge((0, 0), 1, inicio - ang, inicio, width=ancho, facecolor=c, edgecolor="white",
                           linewidth=0.8))
        inicio -= ang
    ax.set_xlim(-1.1, 1.1)
    ax.set_ylim(-1.1, 1.1)
    ax.set_aspect("equal")
    ax.axis("off")


def energia_residuos(lang, salida):
    t = T[lang]
    fig = figura(COLUMNA, 1.8)
    # Energía
    ax1 = fig.add_axes([0.03, 0.3, 0.42, 0.56])
    ren, nor = ENERGIA["renovable"], ENERGIA["no_renovable"]
    dona(ax1, [ren, nor], [ROJO, CELESTE])
    ax1.text(0, 0.1, f"{num(ren / (ren + nor) * 100, lang)}%", ha="center", va="center", fontsize=7.5,
             fontweight="bold", color=ROJO)
    ax1.text(0, -0.25, t["renov"].lower(), ha="center", va="center", fontsize=4.6, color=NEGRO)
    fig.text(0.24, 0.93, t["energia"], ha="center", fontsize=5.4, fontweight="bold", color=AZUL)
    fig.text(0.24, 0.255, f"{num(ren + nor, lang, 0)} MWh", ha="center", fontsize=4.4, color=TINTA_SECUNDARIA)
    # Residuos: desglose de lo valorizado (suma 100% de lo valorizado)
    ax2 = fig.add_axes([0.5, 0.1, 0.5, 0.72])
    colores = [ROJO, NEGRO, CELESTE]
    inicio = 90
    total = sum(p for _, p in VALORIZACION)
    for (k, p), c in zip(VALORIZACION, colores):
        ang = 360 * p / total
        ax2.add_patch(Wedge((0, 0), 1, inicio - ang, inicio, width=0.72, facecolor=c, edgecolor="white", linewidth=0.8))
        mid = np.deg2rad(inicio - ang / 2)
        ax2.text(1.12 * np.cos(mid), 1.12 * np.sin(mid), f"{num(p, lang)}%\n{t['valor'][k]}", ha="right" if np.cos(mid) < 0 else "left",
                 va="center", fontsize=4.4, fontweight="bold", color=NEGRO, linespacing=1.1)
        inicio -= ang
    ax2.text(0, -1.52, f"0% {t['reut']}", ha="center", va="center", fontsize=4.0, color=TINTA_SECUNDARIA)
    ax2.set_xlim(-1.85, 1.85)
    ax2.set_ylim(-1.65, 1.25)
    ax2.set_aspect("equal")
    ax2.axis("off")
    fig.text(0.75, 0.93, t["residuos"], ha="center", fontsize=5.4, fontweight="bold", color=AZUL)
    fig.text(0.75, 0.87, t["de_no_pel"].format(p=num(RESIDUOS_VALORIZADOS, lang)), ha="center", fontsize=4.4, color=TINTA_SECUNDARIA)
    fig.legend(handles=[Patch(color=ROJO, label=t["renov"]), Patch(color=CELESTE, label=t["norenov"])],
               loc="lower center", bbox_to_anchor=(0.24, 0.02), ncol=1, frameon=False, fontsize=4.4,
               handlelength=1, handleheight=0.8, labelspacing=0.25)
    guardar(fig, salida)


# E4 · Social: NPS + indicadores ------------------------------------------------------------------
def social(lang, salida):
    t = T[lang]
    fig = figura(COLUMNA, 1.7)
    ax = fig.add_axes([0.04, 0.12, 0.5, 0.72])
    for i, (anio, v) in enumerate(NPS):
        ax.bar(i, v, width=0.62, color=ROJO if v < 0 else (AZUL if anio == NPS[-1][0] else CELESTE))
        ax.text(i, v + (0.8 if v >= 0 else -0.8), f"{'+' if v > 0 else ''}{num(v, lang, 0)}%",
                ha="center", va="bottom" if v >= 0 else "top", fontsize=5, fontweight="bold", color=NEGRO)
    ax.axhline(0, color=NEGRO, linewidth=0.5)
    ax.set_xticks(range(len(NPS)), [str(a) for a, _ in NPS])
    ax.tick_params(length=0, labelsize=4.8, colors=TINTA_SECUNDARIA, pad=1)
    ax.set_yticks([])
    ax.set_ylim(-13, 17)
    for lado in ax.spines.values():
        lado.set_visible(False)
    fig.text(0.04, 0.92, t["nps"], fontsize=5.4, fontweight="bold", color=AZUL)
    for k, (valor, etiqueta) in enumerate(t["tiles"]):
        y = 0.67 - k * 0.3
        tarjeta(fig, 0.6, y, 0.38, 0.26)
        fig.add_artist(Rectangle((0.6, y), 0.012, 0.26, transform=fig.transFigure, color=ROJO))
        fig.text(0.635, y + 0.175, valor, fontsize=7, fontweight="bold", color=ROJO if k < 2 else AZUL,
                 va="center")
        fig.text(0.635, y + 0.07, etiqueta, fontsize=4.1, color=NEGRO, va="center", linespacing=1.05)
    guardar(fig, salida)


# E5 · Gobierno: remuneración variable y directorio ----------------------------------------------
def gobierno(lang, salida):
    t = T[lang]
    fig = figura(COLUMNA, 1.7)
    ax = fig.add_axes([0.0, 0.2, 0.46, 0.62])
    colores = [ROJO, AMARILLO, CELESTE]
    dona(ax, [v for _, v in REMUNERACION], colores)
    ax.text(0, 0.08, "61%", ha="center", va="center", fontsize=7.5, fontweight="bold", color=ROJO)
    ax.text(0, -0.25, "no fin." if lang == "es" else "non-fin.", ha="center", va="center", fontsize=4.6)
    fig.text(0.23, 0.88, t["rem"], ha="center", fontsize=5.2, fontweight="bold", color=AZUL, linespacing=1.05)
    fig.legend(handles=[Patch(color=c, label=f"{t['remk'][k]} {num(v, lang)}%")
                        for (k, v), c in zip(REMUNERACION, colores)],
               loc="lower center", bbox_to_anchor=(0.23, -0.03), ncol=1, frameon=False, fontsize=4.2,
               handlelength=1, handleheight=0.8, labelspacing=0.2)
    # Directorio: 9 sillas (1 independiente), en ejes con aspecto 1:1 para que sean círculos
    fig.text(0.53, 0.88, t["dir"], fontsize=5.2, fontweight="bold", color=AZUL)
    axd = fig.add_axes([0.53, 0.28, 0.45, 0.52])
    axd.set_xlim(0, 10)
    axd.set_ylim(0, 6.2)
    axd.set_aspect("equal", anchor="W")
    axd.axis("off")
    for i in range(9):
        fila, col = divmod(i, 5)
        axd.add_patch(Circle((0.9 + col * 2, 5.3 - fila * 1.9), 0.75, facecolor=ROJO if i == 0 else AZUL,
                             edgecolor="none"))
    for k, (color, texto) in enumerate(((ROJO, f"1 {t['indep']}"), (AZUL, f"8 {t['no_indep']}"))):
        axd.add_patch(Circle((0.6, 1.6 - k * 1.1), 0.38, facecolor=color, edgecolor="none"))
        axd.text(1.3, 1.6 - k * 1.1, texto, fontsize=4.6, va="center")
    tarjeta(fig, 0.53, 0.04, 0.45, 0.19)
    fig.text(0.55, 0.135, t["ctrl"], fontsize=4.7, fontweight="bold", color=AZUL, va="center", linespacing=1.1)
    guardar(fig, salida)


# Anexo · Estructura de propiedad (barra apilada a lo ancho + leyenda) ---------------------------
def propiedad(lang, salida):
    t = T[lang]
    fig = figura(COMPLETO, 1.0)
    ax = fig.add_axes([0.0, 0.36, 1.0, 0.62])
    colores = [ROJO, CELESTE, AMARILLO, GRIS, NEGRO]
    izq = 0
    for (k, v), c in zip(ACCIONISTAS, colores):
        ax.barh(0, v, left=izq, height=1, color=c, edgecolor="white", linewidth=1.5)
        if v >= 8:
            ax.text(izq + 0.8, 0, f"{num(v, lang, 2)}%", fontsize=10, fontweight="bold",
                    color="white" if c in (ROJO, NEGRO, GRIS) else NEGRO, va="center")
        izq += v
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.5, 0.5)
    limpio(ax)
    for i, ((k, v), c) in enumerate(zip(ACCIONISTAS, colores)):  # 5 columnas iguales, 2 líneas
        x = i * 0.2
        fig.add_artist(Rectangle((x, 0.06), 0.014, 0.2, transform=fig.transFigure, color=c))
        fig.text(x + 0.022, 0.21, t["acc"][k], fontsize=6, va="center", color=NEGRO)
        fig.text(x + 0.022, 0.1, f"{num(v, lang, 2)}%", fontsize=6, fontweight="bold", va="center", color=NEGRO)
    guardar(fig, salida)


# Anexo · Modelo de relacionamiento (5 pasos) -----------------------------------------------------
def relacionamiento(lang, salida):
    t = T[lang]
    fig = figura(COMPLETO, 1.15)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 15.5)
    ax.axis("off")
    for i, texto in enumerate(t["pasos"]):
        cx = 10 + i * 20
        ax.add_patch(Circle((cx, 7.75), 7.2, facecolor=CLARO if i % 2 == 0 else "#fde7ee", edgecolor=CELESTE,
                            linewidth=0.8, alpha=0.95, zorder=1 + i % 2))
        ax.text(cx, 11.3, str(i + 1), ha="center", va="center", fontsize=11, fontweight="bold", color=ROJO, zorder=5)
        ax.text(cx, 7.0, texto, ha="center", va="center", fontsize=5.1, color=NEGRO, linespacing=1.2, zorder=5)
    ax.set_aspect("auto")
    guardar(fig, salida)


# Anexo · Residuos por país -----------------------------------------------------------------------
def residuos_pais(lang, salida):
    t = T[lang]
    fig = figura(COMPLETO / 2 - 0.08, 1.45)
    ax = fig.add_axes([0.02, 0.05, 0.42, 0.9])
    dona(ax, [v for _, v in RESIDUOS_PAIS], [ROJO, AMARILLO, NEGRO])
    ax.text(0, 0, t["total_res"], ha="center", va="center", fontsize=5.6, fontweight="bold", color=AZUL,
            linespacing=1.15)
    for k, ((pais, v), c) in enumerate(zip(RESIDUOS_PAIS, [ROJO, AMARILLO, NEGRO])):
        y = 0.7 - k * 0.22
        fig.add_artist(Rectangle((0.52, y - 0.05), 0.04, 0.1, transform=fig.transFigure, color=c))
        nombre = pais if lang == "es" else {"Perú": "Peru"}.get(pais, pais)
        fig.text(0.59, y, f"{nombre}  {v}%", fontsize=6.5, fontweight="bold", va="center", color=NEGRO)
    guardar(fig, salida)


# Anexo · Directorio con fotos ---------------------------------------------------------------------
FOTOS = Path(__file__).parent / "fotos_directorio"
DIRECTORES = [  # (archivo, nombre, profesión es/en, cargo es/en)
    ("1_cardone", "Sergio Cardone Solari", ("Ingeniero Comercial", "Business engineer"), ("Presidente", "Chairman")),
    ("2_furst_paul", "Paul Fürst Gwinner", ("Ingeniero Comercial", "Business engineer"), ("Director", "Director")),
    ("3_solari", "Carlo Solari Donaggio", ("Ingeniero Civil", "Civil engineer"), ("Director", "Director")),
    ("4_puentes", "Alejandro Puentes Bruno", ("Ingeniero Comercial", "Business engineer"), ("Director", "Director")),
    ("5_eyzaguirre", "Pablo Eyzaguirre Court", ("Ingeniero Civil", "Civil engineer"), ("Director", "Director")),
    ("6_gonzalez", "Dafne González Lizama", ("Abogada", "Lawyer"), ("Directora", "Director")),
    ("7_ostale", "Enrique Ostalé Cambiaso", ("Ingeniero Comercial", "Business engineer"), ("Director", "Director")),
    ("8_furst_ian", "Ian Fürst Garmendia", ("Ingeniero Comercial", "Business engineer"), ("Director", "Director")),
    ("9_ovalle", "Manuel Ovalle Edwards", ("Ingeniero Civil", "Civil engineer"),
     ("Director\nindependiente\nPdte. Comité Dir.", "Independent\ndirector\nChair, Dir. Committee")),
]


def directorio(lang, salida):
    i_lang = 0 if lang == "es" else 1
    fig = figura(COMPLETO, 1.75)
    n_col, ancho_celda = 9, 1 / 9
    for k, (archivo, nombre, prof, cargo) in enumerate(DIRECTORES):
        x0 = k * ancho_celda
        img = plt.imread(FOTOS / f"{archivo}.png")
        lado = ancho_celda * 0.86
        alto_rel = lado * COMPLETO / 1.75            # foto cuadrada en pulgadas
        ax = fig.add_axes([x0 + (ancho_celda - lado) / 2, 0.965 - alto_rel, lado, alto_rel])
        h, w = img.shape[:2]
        yy, xx = np.ogrid[:h, :w]
        mascara = (xx - w / 2) ** 2 + (yy - h / 2) ** 2 > (min(h, w) / 2) ** 2
        img = img.copy()
        img[mascara] = 1.0                            # recorte circular sobre blanco
        ax.imshow(img)
        indep = k == 8
        presi = k == 0
        ax.add_patch(Circle((w / 2 - 0.5, h / 2 - 0.5), min(h, w) / 2 - 2, fill=False,
                            edgecolor=ROJO if indep else (AZUL if presi else "#d5d9e0"),
                            linewidth=2.2 if (indep or presi) else 0.8))
        ax.axis("off")
        cx = x0 + ancho_celda / 2
        y = 0.965 - alto_rel - 0.04
        partes = nombre.split(" ")
        fig.text(cx, y, " ".join(partes[:-2]) + "\n" + " ".join(partes[-2:]), ha="center", va="top",
                 fontsize=5.6, fontweight="bold", color=ROJO if indep else AZUL, linespacing=1.15)
        fig.text(cx, y - 0.17, prof[i_lang], ha="center", va="top", fontsize=4.8, color=TINTA_SECUNDARIA)
        fig.text(cx, y - 0.26, cargo[i_lang], ha="center", va="top", fontsize=4.8,
                 fontweight="bold" if (indep or presi) else "normal", color=NEGRO, linespacing=1.15)
    guardar(fig, salida)


if __name__ == "__main__":
    salida = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "figuras"
    salida.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    for lang in ("es", "en"):
        scorecard(lang, salida / f"{lang}_E1_scorecard.png")
        metas(lang, salida / f"{lang}_E2_metas.png")
        energia_residuos(lang, salida / f"{lang}_E3_energia_residuos.png")
        social(lang, salida / f"{lang}_E4_social.png")
        gobierno(lang, salida / f"{lang}_E5_gobierno.png")
        propiedad(lang, salida / f"{lang}_AE1_propiedad.png")
        relacionamiento(lang, salida / f"{lang}_AE2_relacionamiento.png")
        residuos_pais(lang, salida / f"{lang}_AE3_residuos_pais.png")
        directorio(lang, salida / f"{lang}_AE4_directorio.png")
    print(f"Figuras guardadas en {salida}")
