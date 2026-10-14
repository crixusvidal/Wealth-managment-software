"""Figuras de la página Investment Risks, en español e inglés, al tamaño exacto de la columna.

Matriz de riesgos 3x3 (riesgos coloreados por categoría), tornado operacional, sensibilidad del
precio objetivo (WACC x g) y sensibilidad al reajuste de la UF. Sin título ni fuente en la
imagen: van en el texto del Word (informe/pagina_riesgos.js), traducidos.

Uso:
    python graficos/pagina_riesgos.py <FA_Exhibits.xlsx> <Valoracion_Consolidada.xlsx> \\
        <Tornado_Sensibilidad.xlsx> [carpeta_salida]
"""
import copy
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap, TwoSlopeNorm, to_rgb
from matplotlib.patches import FancyBboxPatch, Rectangle

from heatmap_sensibilidad_wacc import NEUTRO, color_texto, grilla
from reajuste_uf import ESCENARIOS, IMPACTO_BASE, VARIACION_BASE, impacto
from sensibilidad_wacc import AMARILLO, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA
from tornado_operacional import escenarios, leer_modelo, leer_rangos, proyectar

ANCHO = 3550 / 1440  # pulgadas: columna de figuras de la plantilla (3550 twips)
ALTOS = {"matriz": 2.05, "tornado": 2.15, "heatmap": 1.7, "uf": 1.4}
DPI = 400
AZUL = "#14284B"  # azul de títulos del informe

# Riesgos: código → (categoría, probabilidad 1-3, impacto 1-3), igual que el texto de la página
RIESGOS = {"OR1": ("op", 2, 2), "OR2": ("op", 1, 1), "FR1": ("fin", 3, 3), "MR1": ("mac", 3, 1),
           "MR2": ("mac", 1, 2), "MR3": ("mac", 2, 2), "RR1": ("reg", 1, 3), "RR2": ("reg", 2, 1)}
CATEGORIAS = {"op": (AZUL, "white"), "fin": (ROJO, "white"), "mac": (AMARILLO, NEGRO),
              "reg": (CELESTE, NEGRO)}

TEXTOS = {
    "en": {"niveles_i": ["Low", "Medium", "High"], "niveles_p": ["Low", "Medium", "High"],
           "impacto": "Impact", "prob": "Likelihood",
           "cat": {"op": "Operational", "fin": "Financial", "mac": "Market & macro",
                   "reg": "Regulatory & ESG"},
           "base": "Base case", "eje_tornado": "DCF value per share (CLP)",
           "bajo": "Historical low", "alto": "Historical high",
           "g": "Terminal growth (g, nominal)", "wacc": "WACC (nominal)",
           "uf_eje": "Change in target price (CLP per share)", "uf_rentas": "UF-indexed rents",
           "uf_deuda": "UF debt", "uf_neto": "Net effect", "uf_util": "earnings"},
    "es": {"niveles_i": ["Bajo", "Medio", "Alto"], "niveles_p": ["Baja", "Media", "Alta"],
           "impacto": "Impacto", "prob": "Probabilidad",
           "cat": {"op": "Operacional", "fin": "Financiero", "mac": "Mercado y macro",
                   "reg": "Regulatorio y ESG"},
           "base": "Caso base", "eje_tornado": "Valor DCF por acción (CLP)",
           "bajo": "Mínimo histórico", "alto": "Máximo histórico",
           "g": "Crecimiento terminal (g, nominal)", "wacc": "WACC (nominal)",
           "uf_eje": "Cambio en el precio objetivo (CLP por acción)", "uf_rentas": "Rentas indexadas a UF",
           "uf_deuda": "Deuda en UF", "uf_neto": "Efecto neto", "uf_util": "utilidad"},
}
# Tornado: nombre del modelo → (etiqueta EN, etiqueta ES, rango ES cuando no es numérico)
NOMBRES_TORNADO = {
    "EBITDA margin": ("EBITDA margin", "Margen EBITDA", None),
    "Same-store rent growth 2027–28": ("Same-store rent 27–28", "Renta mismas tiendas 27–28", None),
    "Maintenance + IT capex": ("Maint. + IT capex", "Capex mant. + TI", None),
    "Long-term real rent growth 2029–31": ("LT real rent 29–31", "Renta real LP 29–31", None),
    "Visitor traffic 2027–28 (parking)": ("Traffic 27–28 (parking)", "Tráfico 27–28 (estac.)", None),
    "Occupancy": ("Occupancy", "Ocupación", None),
    "Tenant sales growth 2027–28": ("Tenant sales 27–28", "Ventas locatarios 27–28", None),
}
RANGOS_TEXTO = {  # rangos no numéricos o con texto
    "Long-term real rent growth 2029–31": ("appraisal range (Note 11)", "rango de tasación (Nota 11)"),
    "Maintenance + IT capex": ("4.8%–6.6% of revenue", "4,8%–6,6% de ingresos"),
}


def num(x, lang, dec=0):
    s = f"{x:,.{dec}f}".replace("-", "−")
    return s.translate(str.maketrans(",.", ".,")) if lang == "es" else s


def pct(x, lang, dec=1):
    return num(x * 100, lang, dec) + "%"


def tinte(color, fuerza):
    return tuple(1 - fuerza * (1 - c) for c in to_rgb(color))


def figura(clave):
    fig = plt.figure(figsize=(ANCHO, ALTOS[clave]))
    fig.patch.set_facecolor("white")
    return fig


def limpiar(ax):
    for lado in ax.spines.values():
        lado.set_visible(False)


def guardar(fig, salida):
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


# Matriz de riesgos ------------------------------------------------------------------------------
def matriz(lang, salida):
    t = TEXTOS[lang]
    fig = figura("matriz")
    ax = fig.add_axes([0.2, 0.29, 0.76, 0.69])
    fondos = {"bajo": tinte(CELESTE, 0.2), "medio": tinte(AMARILLO, 0.22), "alto": tinte(ROJO, 0.16)}
    for i in range(1, 4):          # impacto (eje X)
        for p in range(1, 4):      # probabilidad (eje Y)
            s = i * p
            nivel = "bajo" if s <= 2 else "medio" if s <= 4 else "alto"
            ax.add_patch(Rectangle((i - 0.5, p - 0.5), 1, 1, facecolor=fondos[nivel],
                                   edgecolor="white", linewidth=2))
    por_celda = {}
    for codigo, (cat, p, i) in RIESGOS.items():
        por_celda.setdefault((i, p), []).append((codigo, cat))
    ancho, alto = 0.5, 0.26
    for (i, p), lista in por_celda.items():
        offsets = [0] if len(lista) == 1 else [0.16, -0.16]
        for (codigo, cat), dy in zip(sorted(lista), offsets):
            fondo, letra = CATEGORIAS[cat]
            ax.add_patch(FancyBboxPatch((i - ancho / 2, p + dy - alto / 2), ancho, alto,
                                        boxstyle="round,pad=0,rounding_size=0.11",
                                        facecolor=fondo, edgecolor="white", linewidth=0.8,
                                        zorder=3))
            ax.text(i, p + dy, codigo, ha="center", va="center", fontsize=6, fontweight="bold",
                    color=letra, zorder=4)
    ax.set_xlim(0.5, 3.5)
    ax.set_ylim(0.5, 3.5)
    ax.set_xticks([1, 2, 3], t["niveles_i"])
    ax.set_yticks([1, 2, 3], t["niveles_p"])
    ax.tick_params(length=0, colors=NEGRO, labelsize=5.5, pad=2.5)
    ax.set_xlabel(t["impacto"], fontsize=6, color=TINTA_SECUNDARIA, labelpad=2)
    ax.set_ylabel(t["prob"], fontsize=6, color=TINTA_SECUNDARIA, labelpad=3)
    limpiar(ax)
    # Leyenda de categorías (2 x 2), alineada con la matriz
    for k, cat in enumerate(["op", "fin", "mac", "reg"]):
        x = 0.2 + (k % 2) * 0.38
        y = 0.095 - (k // 2) * 0.055
        fondo, _ = CATEGORIAS[cat]
        fig.add_artist(FancyBboxPatch((x, y - 0.018), 0.05, 0.036, transform=fig.transFigure,
                                      boxstyle="round,pad=0,rounding_size=0.012",
                                      facecolor=fondo, edgecolor="none"))
        fig.text(x + 0.065, y, t["cat"][cat], fontsize=5.5, color=NEGRO, va="center")
    guardar(fig, salida)


# Tornado operacional ---------------------------------------------------------------------------
def tornado(lang, m, rangos, salida):
    t = TEXTOS[lang]
    base = proyectar(m)[2]
    filas = []
    for nombre, rango, k_lo, k_hi in escenarios(m, rangos):
        etiqueta = NOMBRES_TORNADO[nombre][0 if lang == "en" else 1]
        if nombre in RANGOS_TEXTO:
            rango = RANGOS_TEXTO[nombre][0 if lang == "en" else 1]
        else:
            rango = rango.split(" (")[0].replace(" – ", "–")
            if lang == "es":
                rango = rango.replace(".", ",")
        filas.append((etiqueta, rango, proyectar(m, **k_lo)[2], proyectar(m, **k_hi)[2]))
    filas.sort(key=lambda r: (base - r[2], r[3] - r[2]))  # mayor caída arriba (último = arriba)
    fig = figura("tornado")
    ax = fig.add_axes([0.47, 0.19, 0.51, 0.76])
    for y, (etiqueta, rango, lo, hi) in enumerate(filas):
        ax.barh(y, lo - base, left=base, height=0.62, color=ROJO)
        ax.barh(y, hi - base, left=base, height=0.62, color=CELESTE)
        for v, lado in ((lo, -1), (hi, 1)):
            if abs(v - base) >= 0.5:
                ax.text(v + lado * 7, y, num(v, lang), va="center", fontsize=4.8,
                        ha="left" if lado > 0 else "right", color=NEGRO)
        ax.text(-0.03, y + 0.13, etiqueta, transform=ax.get_yaxis_transform(), ha="right",
                va="center", fontsize=4.9, fontweight="bold", color=NEGRO)
        ax.text(-0.03, y - 0.25, rango, transform=ax.get_yaxis_transform(), ha="right",
                va="center", fontsize=4.2, color=TINTA_SECUNDARIA)
    ax.axvline(base, color=NEGRO, linewidth=0.7, zorder=3)
    ax.text(base, len(filas) - 0.42, f"{t['base']} {num(base, lang)}", ha="center", va="bottom",
            fontsize=5.2, fontweight="bold", color=NEGRO, zorder=4,
            bbox={"facecolor": "white", "edgecolor": "none", "pad": 0.8})
    extremo = max(max(abs(lo - base), abs(hi - base)) for _, _, lo, hi in filas)
    ax.set_xlim(base - extremo * 1.45, base + extremo * 1.32)
    ax.set_ylim(-0.6, len(filas) - 0.05)
    ax.set_yticks([])
    paso = 100
    ax.set_xticks(np.arange(np.ceil((base - extremo * 1.3) / paso) * paso, base + extremo * 1.3, paso))
    ax.xaxis.set_major_formatter(lambda x, _: num(x, lang))
    ax.tick_params(axis="x", length=0, colors=TINTA_SECUNDARIA, labelsize=4.6, pad=1.5)
    ax.grid(axis="x", color=GRILLA, linewidth=0.4)
    ax.set_axisbelow(True)
    limpiar(ax)
    ax.set_xlabel(t["eje_tornado"], fontsize=5, color=TINTA_SECUNDARIA, labelpad=2)
    for x, color, texto in ((0.3, ROJO, t["bajo"]), (0.64, CELESTE, t["alto"])):
        fig.add_artist(Rectangle((x, 0.028), 0.03, 0.03, color=color, transform=fig.transFigure))
        fig.text(x + 0.04, 0.043, texto, fontsize=4.8, color=NEGRO, va="center")
    guardar(fig, salida)
    return base, filas


# Sensibilidad del precio objetivo --------------------------------------------------------------
def heatmap(lang, ruta_vf, salida):
    t = TEXTOS[lang]
    waccs, gs, z, i_b, j_b, _ = grilla(ruta_vf)
    cmap = LinearSegmentedColormap.from_list("mallplaza", [ROJO, NEUTRO, CELESTE])
    norm = TwoSlopeNorm(vmin=z.min(), vcenter=z[i_b, j_b], vmax=z.max())
    fig = figura("heatmap")
    ax = fig.add_axes([0.2, 0.2, 0.78, 0.78])
    ax.imshow(z, cmap=cmap, norm=norm, aspect="auto")
    for i in range(z.shape[0]):
        for j in range(z.shape[1]):
            ax.text(j, i, num(z[i, j], lang), ha="center", va="center", fontsize=5.2,
                    fontweight="bold" if (i, j) == (i_b, j_b) else "normal",
                    color=color_texto(cmap(norm(z[i, j]))))
    for k in range(z.shape[0] + 1):
        ax.axhline(k - 0.5, color="white", linewidth=1.1)
    for k in range(z.shape[1] + 1):
        ax.axvline(k - 0.5, color="white", linewidth=1.1)
    ax.add_patch(Rectangle((j_b - 0.45, i_b - 0.41), 0.9, 0.82, fill=False, edgecolor=AMARILLO,
                           linewidth=1.1, zorder=3))
    ax.set_xticks(range(len(gs)), [pct(g, lang) for g in gs])
    ax.set_yticks(range(len(waccs)), [pct(w, lang, 2) for w in waccs])
    ax.tick_params(length=0, colors=NEGRO, labelsize=4.8, pad=2)
    ax.set_xlabel(t["g"], fontsize=5, color=TINTA_SECUNDARIA, labelpad=2)
    ax.set_ylabel(t["wacc"], fontsize=5, color=TINTA_SECUNDARIA, labelpad=2)
    limpiar(ax)
    guardar(fig, salida)
    return waccs, gs, z, i_b, j_b


# Sensibilidad a la UF --------------------------------------------------------------------------
PESO_DCF, MULTIPLOS, OBJETIVO = 0.8, 2951.39, 2075.73  # Precio_Objetivo B14, B13, B15 (modelo v2)
EXPOSICION_UF = -IMPACTO_BASE / (VARIACION_BASE / 100)    # MM CLP de deuda neta en UF (Nota 26)


def sensibilidad_uf(m, s):
    """Shock de nivel de la UF de s (0,01 = 1%) que se materializa en 2027E.

    Rentas: el IPC de Chile de 2027E sube s (rentas, estacionamientos e ingreso por m² de la
    expansión quedan s más altos desde 2027E; el margen fijo escala los costos) y el costo por m²
    de la expansión sube s por la fracción Chile. Deuda: la deuda neta en UF sube s (exposición
    implícita en la Nota 26). WACC y g sin cambio. Devuelve (rentas, deuda) en CLP por acción
    sobre el precio objetivo: rentas pesan 80% (solo DCF); la deuda resta en ambos métodos.
    """
    mm = copy.deepcopy(m)
    mm["ipc"]["cl"][0] += s
    mm["costo_m2"][1:] *= 1 + s * m["mix"]
    rentas = PESO_DCF * (proyectar(mm)[2] - proyectar(m)[2])
    deuda = -s * EXPOSICION_UF / m["acciones"]
    return rentas, deuda


def uf(lang, m, salida):
    t = TEXTOS[lang]
    fig = figura("uf")
    ax = fig.add_axes([0.33, 0.3, 0.64, 0.67])
    filas = []
    for y, v in zip(range(len(ESCENARIOS))[::-1], ESCENARIOS):
        rentas, deuda = sensibilidad_uf(m, v / 100)
        neto = rentas + deuda
        filas.append((v, rentas, deuda, neto))
        ax.barh(y + 0.15, rentas, height=0.28, color=CELESTE)
        ax.barh(y - 0.15, deuda, height=0.28, color=ROJO)
        ax.text(rentas + 0.8, y + 0.15, num(rentas, lang, 1), va="center", fontsize=4.6, color=NEGRO)
        ax.text(deuda - 0.8, y - 0.15, num(deuda, lang, 1), va="center", ha="right", fontsize=4.6,
                color=NEGRO)
        ax.plot(neto, y, marker="D", markersize=3.6, color=NEGRO, zorder=4, linestyle="none")
        ax.text(neto, y + 0.36, f"{'+' if neto > 0 else ''}{num(neto, lang, 1)} "
                f"({'+' if neto > 0 else ''}{pct(neto / OBJETIVO, lang)})", ha="center", va="bottom",
                fontsize=4.6, fontweight="bold", color=NEGRO, zorder=5,
                bbox={"facecolor": "white", "edgecolor": "none", "pad": 0.3})
        ax.text(-0.04, y + 0.12, f"UF +{v}%", transform=ax.get_yaxis_transform(), ha="right",
                va="center", fontsize=5.4, fontweight="bold", color=NEGRO)
        ax.text(-0.04, y - 0.2, f"{t['uf_util']} {num(impacto(v), lang)} MM",
                transform=ax.get_yaxis_transform(), ha="right", va="center", fontsize=4.2,
                color=TINTA_SECUNDARIA)
    ax.axvline(0, color=NEGRO, linewidth=0.7)
    ax.set_xlim(-30, 50)
    ax.set_ylim(-0.5, len(ESCENARIOS) - 0.25)
    ax.set_yticks([])
    ax.set_xticks([-20, 0, 20, 40])
    ax.xaxis.set_major_formatter(lambda x, _: num(x, lang))
    ax.tick_params(axis="x", length=0, colors=TINTA_SECUNDARIA, labelsize=4.6, pad=1.5)
    ax.grid(axis="x", color=GRILLA, linewidth=0.4)
    ax.set_axisbelow(True)
    limpiar(ax)
    ax.set_xlabel(t["uf_eje"], fontsize=5, color=TINTA_SECUNDARIA, labelpad=2)
    leyenda = ((0.05, CELESTE, "s", t["uf_rentas"]), (0.43, ROJO, "s", t["uf_deuda"]),
               (0.72, NEGRO, "D", t["uf_neto"]))
    for x, color, forma, texto in leyenda:
        if forma == "s":
            fig.add_artist(Rectangle((x, 0.03), 0.03, 0.035, color=color, transform=fig.transFigure))
        else:
            fig.add_artist(plt.Line2D([x + 0.015], [0.0475], marker="D", markersize=3.2, color=color,
                                      transform=fig.transFigure))
        fig.text(x + 0.04, 0.0475, texto, fontsize=4.8, color=NEGRO, va="center")
    guardar(fig, salida)
    return filas


if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    ruta_fa, ruta_vf, ruta_tornado = sys.argv[1:4]
    salida = Path(sys.argv[4] if len(sys.argv) > 4 else Path(__file__).parent / "informe" / "riesgos")
    salida.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    m, _ = leer_modelo(ruta_fa, ruta_vf)
    rangos = leer_rangos(ruta_tornado)
    for lang in ("es", "en"):
        matriz(lang, salida / f"{lang}_fig23_matriz.png")
        base, filas = tornado(lang, m, rangos, salida / f"{lang}_fig24_tornado.png")
        heatmap(lang, ruta_vf, salida / f"{lang}_fig25_heatmap.png")
        filas_uf = uf(lang, m, salida / f"{lang}_fig26_uf.png")
    print(f"Base DCF {base:,.2f}")
    for etiqueta, rango, lo, hi in reversed(filas):
        print(f"  {etiqueta:26s} {rango:28s} {lo:7.0f} ({lo / base - 1:+.1%})  {hi:7.0f} ({hi / base - 1:+.1%})")
    for v, rentas, deuda, neto in filas_uf:
        print(f"  UF +{v}%: rentas {rentas:+.1f}  deuda {deuda:+.1f}  neto {neto:+.1f} ({neto / OBJETIVO:+.2%})")
    print(f"Figuras guardadas en {salida}")
