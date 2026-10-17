"""Figuras nuevas del Resumen de inversión (junto a V10, escenarios con el oso y el toro):

  V14  Trayectoria de la acción entre los hitos (oct-25 a sep-26) frente al precio objetivo y al valor DCF.
       Solo hay tres precios confirmados con Bloomberg (mínimo 10-oct-25, máximo 27-abr-26 y cierre 30-sep-26); los demás hitos
       (9-oct-25, 6-nov-25, 25-may-26) son precios IMPLÍCITOS en las notas de prensa (objetivo / (1 + potencial)) y se dibujan
       huecos y con «~»: la línea que une los hitos es solo una guía, no la serie diaria. El 15-dic-25 (elección) no tiene precio.
  V15  Escalera de valores por acción: del DCF propio al precio de mercado y al consenso de 12 meses (Bloomberg, 14 fuentes).

Cifras del modelo v4 (informe_final/graficos/cifras.json). Sin título ni fuente en la imagen: van en el documento.

Uso (desde secciones/): python valuacion/graficos/resumen_figuras.py [carpeta_salida]
Salida: <carpeta>/{es,en}_V14_trayectoria.png y {es,en}_V15_escalera.png
"""
import datetime as dt
import json
import sys
from pathlib import Path

import matplotlib.dates as mdates
import matplotlib.pyplot as plt

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional
import estilo  # noqa: E402,F401
from estilo import AMARILLO, AZUL, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA, num

K = json.loads((Path(__file__).resolve().parents[2] / "informe_final" / "graficos" / "cifras.json").read_text())
FIGS = Path(__file__).parent / "figuras"
ANCHO, DPI = 3550 / 1440, 400
D = dt.date
# (fecha, precio, confirmado con Bloomberg, etiqueta es, etiqueta en, posición del texto)
HITOS = [(D(2025, 10, 9), 2500, False, "", "", (0, 0)),
         (D(2025, 10, 10), 2470, True, "Mín. 52 sem.", "52-wk low", (7, -9)),
         (D(2025, 11, 6), 2800, False, "MSCI", "MSCI", (5, 12)),
         (D(2026, 4, 27), 4700, True, "Máx. 52 sem.", "52-wk high", (-2, 8)),
         (D(2026, 5, 25), 3900, False, "Recorte JPM", "JPM cut", (6, 4)),
         (D(2026, 9, 30), 3555, True, "Cierre", "Close", (-6, -17))]
T = {"es": {"po": "Precio objetivo", "dcf": "Valor DCF", "aprox": "~", "leyenda": ["Confirmado (Bloomberg)", "Implícito en notas de prensa"], "eje": "CLP por acción",
            "eleccion": "Elección 15-dic"},
     "en": {"po": "Target price", "dcf": "DCF value", "aprox": "~", "leyenda": ["Confirmed (Bloomberg)", "Implied in press notes"], "eje": "CLP per share",
            "eleccion": "Election 15-Dec"}}


def trayectoria(lang, salida):
    t = T[lang]
    fig = plt.figure(figsize=(ANCHO, 2.05))
    ax = fig.add_axes([0.14, 0.17, 0.83, 0.68])
    x = [h[0] for h in HITOS]
    y = [h[1] for h in HITOS]
    ax.plot(x, y, color=TINTA_SECUNDARIA, linewidth=0.6, linestyle=(0, (2, 2)), zorder=2)
    for fecha, precio, ok, es, en, (dx, dy) in HITOS:
        ax.scatter([fecha], [precio], s=11, facecolor=AZUL if ok else "white", edgecolor=AZUL, linewidth=0.8, zorder=4)
        rotulo = (es if lang == "es" else en)
        if not rotulo:
            continue
        valor = num(precio, lang) if ok else t["aprox"] + num(precio, lang)
        ax.annotate(f"{rotulo}\n{valor}", (fecha, precio), textcoords="offset points", xytext=(dx, dy), fontsize=3.8, color=NEGRO,
                    ha="left" if dx >= 0 else "right", va="center", linespacing=1.05)
    ax.axhline(K["po"], color=ROJO, linewidth=0.8, linestyle=(0, (4, 2)))
    ax.text(D(2026, 10, 22), K["po"] + 60, f"{t['po']} {num(K['po'], lang)}", fontsize=4.2, color=ROJO, fontweight="bold", va="bottom", ha="right")
    ax.axhline(K["dcf"], color=AZUL, linewidth=0.6, linestyle=(0, (1, 2)))
    ax.text(D(2026, 10, 22), K["dcf"] - 50, f"{t['dcf']} {num(K['dcf'], lang)}", fontsize=4.0, color=AZUL, va="top", ha="right")
    ax.axvline(D(2025, 12, 15), color=GRILLA, linewidth=0.8, zorder=1)
    ax.text(D(2025, 12, 18), 3050, t["eleccion"], fontsize=3.6, color=TINTA_SECUNDARIA, rotation=90, va="bottom")
    ax.set_ylim(1500, 5100)
    ax.set_xlim(D(2025, 9, 25), D(2026, 10, 25))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b-%y"))
    ax.yaxis.set_major_formatter(lambda v, _: num(v, lang))
    ax.tick_params(labelsize=4.0, length=1.5, colors=TINTA_SECUNDARIA)
    ax.grid(axis="y", color=GRILLA, linewidth=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_ylabel(t["eje"], fontsize=4.0, color=TINTA_SECUNDARIA, labelpad=2)
    for i, (txt, lleno) in enumerate(zip(t["leyenda"], (True, False))):
        xx = 0.15 + i * 0.42
        fig.add_artist(plt.Line2D([xx], [0.945], marker="o", markersize=3.2, markerfacecolor=AZUL if lleno else "white", markeredgecolor=AZUL, linestyle="none",
                                  transform=fig.transFigure))
        fig.text(xx + 0.02, 0.945, txt, fontsize=4.1, color=NEGRO, va="center")
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


def escalera(lang, salida):
    t = {"es": ["Nuestro DCF", "Precio objetivo (80/20)", "DCF con la tasa de los tasadores", "Objetivo, supuestos a favor", "Precio de mercado", "Consenso 12 meses"],
         "en": ["Our DCF", "Target price (80/20)", "DCF at the appraisers' rate", "Target, assumptions in our favor", "Market price", "12-month consensus"]}[lang]
    notas = {"es": ["WACC nominal 10,41%", "", f"WACC nominal {K['tasacion_nom'] * 100:.2f}%".replace(".", ","), "impuesto 23%, margen máx., sin CT, tasadores",
                    f"WACC implícita {K['wacc_impl'] * 100:.2f}%".replace(".", ","), f"rango {num(K['consenso']['bajo'], lang)}–{num(K['consenso']['alto'], lang)}"],
             "en": ["nominal WACC 10.41%", "", f"nominal WACC {K['tasacion_nom'] * 100:.2f}%", "23% tax, max. margin, no WC, appraisers",
                    f"implied WACC {K['wacc_impl'] * 100:.2f}%", f"range {num(K['consenso']['bajo'], lang)}–{num(K['consenso']['alto'], lang)}"]}[lang]
    vals = [K["dcf"], K["po"], K["esc"]["tas"]["dcf"], K["supuestos_acum"][-1], K["precio"], K["consenso"]["po_12m"]]
    cols = [AZUL, ROJO, CELESTE, CELESTE, NEGRO, AMARILLO]
    fig = plt.figure(figsize=(ANCHO, 2.05))
    ax = fig.add_axes([0.02, 0.1, 0.96, 0.86])
    n = len(vals)
    for i, (v, c, txt, nota) in enumerate(zip(vals, cols, t, notas)):
        y = n - 1 - i
        ax.barh(y, v, height=0.7, color=c, zorder=3)
        ax.text(v + 45, y + 0.05, num(v, lang), fontsize=5.0, fontweight="bold", color=NEGRO, va="center")
        ax.text(60, y + 0.17, txt, fontsize=4.3, color="white" if c in (AZUL, ROJO, NEGRO) else NEGRO, va="center", fontweight="bold")
        if nota:
            ax.text(60, y - 0.17, nota, fontsize=3.3, color="white" if c in (AZUL, ROJO, NEGRO) else NEGRO, va="center")
    # el consenso descontado a nuestro Ke y su rango
    d = K["consenso"]["po_12m"] / (1 + K["ke_t"])
    ax.plot([K["consenso"]["bajo"], K["consenso"]["alto"]], [-0.46, -0.46], color=NEGRO, linewidth=0.7, zorder=5)
    ax.scatter([d], [-0.46], s=12, facecolor="white", edgecolor=NEGRO, linewidth=0.8, zorder=6)
    ax.text(d - 60, -0.46, ("hoy a nuestro Ke: " if lang == "es" else "today at our Ke: ") + num(d, lang), fontsize=3.5, color=NEGRO, ha="right", va="center")
    ax.axvline(K["precio"], color=ROJO, linewidth=0.6, linestyle=(0, (4, 2)), zorder=2)
    ax.set_xlim(0, 5200)
    ax.set_ylim(-0.75, n - 0.5)
    ax.axis("off")
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    salida = Path(sys.argv[1]) if len(sys.argv) > 1 else FIGS
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    for lang in ("es", "en"):
        trayectoria(lang, salida / f"{lang}_V14_trayectoria.png")
        escalera(lang, salida / f"{lang}_V15_escalera.png")
    print("ok", salida)
