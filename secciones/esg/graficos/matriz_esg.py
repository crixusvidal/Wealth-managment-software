"""Figura S5 del anexo ESG: matriz de factores ESG y su efecto en el caso de inversión (español e inglés).

Redibuja en la paleta institucional la matriz del borrador del equipo (CFA beta 4.1, imagen S5): nueve factores
(3 ambientales, 3 sociales, 3 de gobierno) con materialidad, visibilidad y horizonte, capacidad de absorber pérdidas,
los canales por los que golpean el caso de inversión y el efecto neto. Materialidad, horizonte, capacidad y efecto neto
son juicios del equipo (marco adaptado del enfoque de S&P Global Ratings para factores ESG de crédito); la evidencia
sale de la Memoria Mallplaza 2025, Bloomberg, MSCI y Sustainalytics. Sin título ni fuente dentro de la imagen.

Uso: python graficos/matriz_esg.py [carpeta_salida]     Salida: {es,en}_AE5_matriz_esg.png
"""
import sys
from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Polygon, Rectangle

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional
import estilo  # noqa: E402,F401
from estilo import AMARILLO, CELESTE, NEGRO, ROJO, TINTA_SECUNDARIA

AZUL = "#14284B"
CLARO = "#eef1f6"
ANCHO, ALTO, DPI = 10666 / 1440, 5.05, 400

# (pilar, clave, materialidad, visibilidad, capacidad, canales{col: 'p'|'s'}, efecto)   efecto: + favorable, = neutro, - desfavorable, ? brecha
FILAS = [
    ("E", "energia", "alta", "medio", "fuerte", {1: "s", 3: "p"}, "+"),
    ("E", "agua", "media", "medio", "moderada", {1: "p", 3: "s"}, "="),
    ("E", "clima", "alta", "largo", "n/d", {0: "p", 1: "s", 3: "s"}, "?"),
    ("S", "arrendatarios", "alta", "corto", "fuerte", {2: "p", 3: "s"}, "+"),
    ("S", "capital", "media", "medio", "moderada", {2: "p"}, "="),
    ("S", "comunidad", "media", "medio", "moderada", {0: "p", 2: "s"}, "="),
    ("G", "directorio", "alta", "largo", "débil", {3: "s", 4: "p"}, "-"),
    ("G", "incentivos", "media", "corto", "fuerte", {4: "p"}, "+"),
    ("G", "transparencia", "media", "corto", "moderada", {3: "s", 4: "p"}, "="),
]

T = {
    "es": {
        "h1": "FACTORES ESG DE RIESGO DE CRÉDITO\n(evidencia de PLAZA)", "h2": "INTEGRACIÓN DE LA EVALUACIÓN",
        "h3": "DÓNDE GOLPEA EL CASO DE INVERSIÓN", "h4": "EFECTO\nNETO",
        "sub": ["Materialidad", "Visibilidad y\nhorizonte", "Capacidad de\nabsorber pérdidas"],
        "canales": ["Riesgo\npaís", "Riesgo\nindustria", "Posición\ncompetitiva", "Flujo de caja\ny deuda", "Gestión y\ngobierno"],
        "grupos": ["Perfil de riesgo del negocio", "Financiero", "Modificador"],
        "chip": {"alta": "Alta", "media": "Media", "corto": "Corto", "medio": "Medio", "largo": "Largo",
                 "fuerte": "Fuerte", "moderada": "Moderada", "débil": "Débil", "n/d": "n/d"},
        "factor": {
            "energia": ("Energía y emisiones GEI", "89,3% de electricidad renovable (meta 100% al 2026); intensidad energética −22% al 2035: avance de 35,5%"),
            "agua": ("Agua, residuos y refrigerantes", "Agua −30% vs 2019 cumplido; valorización de residuos 46,2% (meta 60%); eliminación de HCFC 14,9% (100% al 2030)"),
            "clima": ("Riesgo físico climático", "Exposición sísmica, a inundaciones e incendios en Chile, Perú y Colombia: aún sin cuantificar en el informe"),
            "arrendatarios": ("Arrendatarios y clientes", "NPS de arrendatarios de −8 (2022) a +13 (2025); Cadem: puesto 31 de 385 marcas y ganadora de su categoría"),
            "capital": ("Capital humano", "Mujeres: 53% de la fuerza laboral y 42% de los cargos directivos; brecha salarial regional de 28%; engagement 75% vs meta 80%"),
            "comunidad": ("Comunidad y cadena de suministro", "49.325 personas alcanzadas por programas; criterios de sostenibilidad en 55% de los proveedores regionales y 38% de los extranjeros"),
            "directorio": ("Directorio y accionistas", "1 de 9 directores independiente; 1 mujer en el Directorio; Falabella 53,10% y fondos de pensiones 23,8%"),
            "incentivos": ("Incentivos y controles internos", "33,3% de la remuneración variable ligada a KPIs ESG (61% no financiera); un director independiente preside el Comité de Directores"),
            "transparencia": ("Transparencia y ratings", "Bloomberg ESG 3,39 (0–10); MSCI BBB; Sustainalytics 14,4 (riesgo bajo); DJSI Best-in-Class Chile y MILA"),
        },
        "leyenda": ["Canal principal", "Canal secundario", "Favorable", "Neutral", "Desfavorable", "Brecha: no cuantificada en el informe"],
    },
    "en": {
        "h1": "ESG CREDIT-RISK FACTORS\n(PLAZA evidence)", "h2": "ASSESSMENT INTEGRATION",
        "h3": "WHERE IT HITS THE INVESTMENT CASE", "h4": "NET\nEFFECT",
        "sub": ["Materiality", "Visibility &\ntime horizon", "Ability to\nabsorb losses"],
        "canales": ["Country\nrisk", "Industry\nrisk", "Competitive\nposition", "Cash flow\n/ leverage", "Mgmt &\ngovernance"],
        "grupos": ["Business risk profile", "Financial", "Modifier"],
        "chip": {"alta": "High", "media": "Medium", "corto": "Short", "medio": "Medium", "largo": "Long",
                 "fuerte": "Strong", "moderada": "Moderate", "débil": "Weak", "n/d": "n/a"},
        "factor": {
            "energia": ("Energy & GHG emissions", "89.3% renewable electricity (goal 100% by 2026); energy intensity −22% by 2035: 35.5% achieved"),
            "agua": ("Water, waste & refrigerants", "Water −30% vs 2019 achieved; waste recovery 46.2% (goal 60%); HCFC phase-out 14.9% (100% by 2030)"),
            "clima": ("Climate physical risk", "Seismic, flood and wildfire exposure across Chile, Peru and Colombia: not yet quantified in the report"),
            "arrendatarios": ("Tenants & customers", "Tenant NPS −8 (2022) to +13 (2025); Cadem: rank 31 of 385 brands and category winner"),
            "capital": ("Human capital", "Women: 53% of workforce and 42% of managers; regional pay gap 28%; engagement 75% vs 80% goal"),
            "comunidad": ("Community & supply chain", "49,325 people reached by programs; sustainability criteria applied to 55% of regional and 38% of foreign suppliers"),
            "directorio": ("Board & shareholders", "1 of 9 directors independent; 1 woman on the board; Falabella holds 53.10%, pension funds 23.8%"),
            "incentivos": ("Incentives & internal controls", "33.3% of variable pay tied to ESG KPIs (61% non-financial); an independent director chairs the Directors' Committee"),
            "transparencia": ("Transparency & ratings", "Bloomberg ESG 3.39 (0–10); MSCI BBB; Sustainalytics 14.4 (low risk); DJSI Best-in-Class Chile and MILA"),
        },
        "leyenda": ["Primary channel", "Secondary channel", "Favorable", "Neutral", "Unfavorable", "Gap: not quantified in the report"],
    },
}
PILAR = {"E": CELESTE, "S": AMARILLO, "G": ROJO}
CHIP_MAT = {"alta": (AZUL, "white"), "media": (CELESTE, NEGRO)}
CHIP_CAP = {"fuerte": (CELESTE, NEGRO), "moderada": (AMARILLO, NEGRO), "débil": (ROJO, "white"), "n/d": ("white", TINTA_SECUNDARIA)}
X_TEXTO, X_MAT, X_VIS, X_CAP = 0.27, 2.65, 3.27, 3.89
CANALES_X = [4.88, 5.33, 5.78, 6.23, 6.68]
X_NETO = 7.1
LARGO_TAB = ANCHO


def chip(ax, x, y, w, h, texto, fondo, color, borde=None):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=0.03", facecolor=fondo,
                                edgecolor=borde or "none", linewidth=0.5))
    ax.text(x + w / 2, y + h / 2, texto, ha="center", va="center", fontsize=4.9, fontweight="bold", color=color)


def simbolo(ax, x, y, tipo, r=0.05):
    if tipo == "+":
        ax.add_patch(Polygon([[x - r, y - r * 0.8], [x + r, y - r * 0.8], [x, y + r * 0.9]], closed=True, color=CELESTE))
    elif tipo == "-":
        ax.add_patch(Polygon([[x - r, y + r * 0.8], [x + r, y + r * 0.8], [x, y - r * 0.9]], closed=True, color=ROJO))
    elif tipo == "=":
        ax.add_patch(Rectangle((x - r, y - r * 0.25), 2 * r, r * 0.5, color=NEGRO))
    else:
        ax.text(x, y, "?", ha="center", va="center", fontsize=8, fontweight="bold", color=AMARILLO)


def figura(lang, salida):
    t = T[lang]
    fig = plt.figure(figsize=(ANCHO, ALTO))
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, ANCHO)
    ax.set_ylim(0, ALTO)
    ax.axis("off")
    # Encabezados
    yh, hh = ALTO - 0.46, 0.44
    for x0, x1, txt in ((0.0, 2.55, t["h1"]), (2.6, 4.65, t["h2"]), (4.7, 6.95, t["h3"]), (7.0, ANCHO, t["h4"])):
        ax.add_patch(Rectangle((x0, yh), x1 - x0, hh, facecolor=AZUL, linewidth=0))
        ax.text((x0 + x1) / 2, yh + hh / 2, txt, ha="center", va="center", fontsize=5.0, fontweight="bold", color="white", linespacing=1.15)
    # Subencabezados
    ys = yh - 0.36
    for x, txt in zip((X_MAT + 0.275, X_VIS + 0.275, X_CAP + 0.325), t["sub"]):
        ax.text(x, ys, txt, ha="center", va="center", fontsize=4.5, fontweight="bold", color=AZUL, linespacing=1.1)
    for x, txt in zip(CANALES_X, t["canales"]):
        ax.text(x, ys, txt, ha="center", va="center", fontsize=4.2, fontweight="bold", color=AZUL, linespacing=1.1)
    yg = ys + 0.24
    for (xa, xb, g, col) in ((CANALES_X[0] - 0.17, CANALES_X[2] + 0.17, t["grupos"][0], ROJO), (CANALES_X[3] - 0.17, CANALES_X[3] + 0.17, t["grupos"][1], ROJO),
                             (CANALES_X[4] - 0.17, CANALES_X[4] + 0.17, t["grupos"][2], ROJO)):
        ax.text((xa + xb) / 2, yg, g, ha="center", va="center", fontsize=4.2, fontweight="bold", color=col, style="italic")
        ax.plot([xa, xb], [yg - 0.07, yg - 0.07], color=col, linewidth=0.6)
    # Filas
    alto, gap_grupo, y = 0.40, 0.07, ys - 0.17
    previo = None
    for pilar, clave, mat, vis, cap, canales, efecto in FILAS:
        if previo and pilar != previo:
            y -= gap_grupo
        previo = pilar
        y -= alto
        ax.add_patch(Rectangle((0, y), ANCHO, alto - 0.03, facecolor=CLARO, linewidth=0))
        ax.add_patch(Rectangle((0, y), 0.2, alto - 0.03, facecolor=PILAR[pilar], linewidth=0))
        ax.text(0.1, y + (alto - 0.03) / 2, pilar, ha="center", va="center", fontsize=7, fontweight="bold",
                color="white" if pilar == "G" else NEGRO)
        titulo, evidencia = t["factor"][clave]
        ax.text(X_TEXTO, y + alto - 0.07, titulo, ha="left", va="top", fontsize=5.4, fontweight="bold", color=AZUL)
        # evidencia en dos líneas como máximo
        palabras, lineas, actual = evidencia.split(" "), [], ""
        for pal in palabras:
            if len((actual + " " + pal).strip()) > 60 and actual:
                lineas.append(actual)
                actual = pal
            else:
                actual = (actual + " " + pal).strip()
        lineas.append(actual)
        ax.text(X_TEXTO, y + alto - 0.14, "\n".join(lineas[:3]), ha="left", va="top", fontsize=4.1, color=NEGRO, linespacing=1.15)
        cy = y + (alto - 0.03) / 2 - 0.05
        chip(ax, X_MAT, cy, 0.55, 0.1, t["chip"][mat], *CHIP_MAT[mat])
        chip(ax, X_VIS, cy, 0.55, 0.1, t["chip"][vis], "white", AZUL, borde=AZUL)
        f, c = CHIP_CAP[cap]
        chip(ax, X_CAP, cy, 0.65, 0.1, t["chip"][cap], f, c, borde=AMARILLO if cap == "n/d" else None)
        for col, tipo in canales.items():
            cx, cyy = CANALES_X[col], cy + 0.05
            ax.add_patch(plt.Circle((cx, cyy), 0.055, facecolor=AZUL if tipo == "p" else "white", edgecolor=AZUL, linewidth=0.9, zorder=3))
        simbolo(ax, X_NETO + 0.1, cy + 0.05, efecto)
        if cap == "n/d":                                   # brecha: contorno punteado amarillo
            ax.add_patch(FancyBboxPatch((0.24, y + 0.005), 4.5, alto - 0.04, boxstyle="round,pad=0,rounding_size=0.02", facecolor="none",
                                        edgecolor=AMARILLO, linewidth=0.9, linestyle=(0, (3, 2)), zorder=4))
    # Leyenda
    yl = y - 0.18
    ax.add_patch(plt.Circle((0.25, yl), 0.05, facecolor=AZUL, edgecolor=AZUL, linewidth=0.9))
    ax.text(0.36, yl, t["leyenda"][0], va="center", fontsize=4.4, color=NEGRO)
    ax.add_patch(plt.Circle((1.35, yl), 0.05, facecolor="white", edgecolor=AZUL, linewidth=0.9))
    ax.text(1.46, yl, t["leyenda"][1], va="center", fontsize=4.4, color=NEGRO)
    simbolo(ax, 2.55, yl, "+")
    ax.text(2.67, yl, t["leyenda"][2], va="center", fontsize=4.4, color=NEGRO)
    simbolo(ax, 3.35, yl, "=")
    ax.text(3.47, yl, t["leyenda"][3], va="center", fontsize=4.4, color=NEGRO)
    simbolo(ax, 4.15, yl, "-")
    ax.text(4.27, yl, t["leyenda"][4], va="center", fontsize=4.4, color=NEGRO)
    simbolo(ax, 5.15, yl, "?")
    ax.text(5.27, yl, t["leyenda"][5], va="center", fontsize=4.4, color=NEGRO)
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    salida = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "figuras"
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    for lang in ("es", "en"):
        figura(lang, salida / f"{lang}_AE5_matriz_esg.png")
    print("Matriz ESG guardada en", salida)
