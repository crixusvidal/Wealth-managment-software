"""Figura de supuestos discrecionales del equipo en la valoración: cascada del precio objetivo + tabla de efectos individuales.

Misma forma que la figura de capital de trabajo (V4): arriba una cascada, abajo una tabla. La cascada parte del precio
objetivo (CLP 2.076) y suma, en orden y de forma acumulada, las alternativas que un comprador podría defender
(impuesto de 23%, margen máximo histórico, sin capital de trabajo, g de 4%, WACC de los tasadores). Termina en el
valor «todo a favor» y se compara con el precio de mercado (CLP 3.555). La tabla muestra el efecto de cada supuesto
por separado y suma dos supuestos en contra (tasa efectiva de la deuda y fin del escudo del goodwill tributario).

Todo se calcula con la réplica del DCF v3 de riesgos/graficos/tornado_operacional.py (reproduce CLP 1.856,81).
Precio objetivo = 80% DCF + 20% múltiplos (múltiplos fijos).

Uso (desde secciones/):
    python valuacion/graficos/supuestos_cascada.py <modelo_v3.xlsx> <FA_Exhibits_v6.xlsx> <Tornado.xlsx> [carpeta_salida]
Salida: <carpeta>/{es,en}_V13_supuestos_cascada.png y figuras/datos_supuestos.json (cifras de la figura y del escudo tributario)
"""
import copy
import json
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from matplotlib.patches import Rectangle

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "riesgos" / "graficos"))
import estilo  # noqa: E402,F401
from estilo import AMARILLO, AZUL, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA, num  # noqa: E402
from tornado_operacional import leer_modelo, leer_rangos, proyectar  # noqa: E402

FIGS = Path(__file__).parent / "figuras"
IMPUESTO_LEY = 0.23            # Ley de Reconstrucción (27% → 23%); no está en el modelo
G_ALTO = 0.04                  # crecimiento nominal terminal alternativo
DW_KD = 0.0012                 # +12 pb de WACC con la tasa efectiva de la última emisión (4,33% vs cupón 3,6%)

T = {
    "es": {"pasos": ["Objetivo\nbase", "Impuesto\n23%", "Margen\nmáx.", "Cap. trab.\n= 0", "g 4,0%", "WACC\ntasadores", "Todo a\nfavor"],
           "filas": ["Impuesto de largo plazo 27% → 23% (ley)", "Margen EBITDA al máximo histórico", "Capital de trabajo = 0",
                     "Crecimiento g nominal 3,0% → 4,0%", "WACC de los tasadores (8,72% nominal)", "Tasa efectiva de la deuda (4,33%)",
                     "Sin escudo del goodwill (27% desde ya)"],
           "cols": ["Supuesto (cada uno por separado)", "Δ objetivo", "Δ%"], "mm": "CLP por acción",
           "precio": "Precio {p}", "brecha": "Brecha {b}"},
    "en": {"pasos": ["Base\ntarget", "Tax\n23%", "Max.\nmargin", "Work. cap.\n= 0", "g 4.0%", "Appraisers'\nWACC", "All in\nfavor"],
           "filas": ["Long-run tax 27% → 23% (new law)", "EBITDA margin at historical high", "Working capital = 0",
                     "Nominal terminal growth g 3.0% → 4.0%", "Appraisers' WACC (8.72% nominal)", "Effective cost of debt (4.33%)",
                     "No tax-goodwill shield (27% from now)"],
           "cols": ["Assumption (each on its own)", "Δ target", "Δ%"], "mm": "CLP per share",
           "precio": "Price {p}", "brecha": "Gap {b}"},
}


def calcular(v3, ruta_fa, ruta_tor):
    m, _ = leer_modelo(ruta_fa, v3)
    wb = openpyxl.load_workbook(v3, data_only=True)
    po = wb["Precio_Objetivo"]
    mult, peso, precio = po["B13"].value, po["B14"].value, po["B16"].value
    objetivo = lambda x: peso * x + (1 - peso) * mult
    fb = wb["Forecast_Base"]
    margen26 = fb["C7"].value / fb["C5"].value
    d_margen = leer_rangos(ruta_tor)["Margen EBITDA"][2] - margen26                      # máximo histórico
    wacc_tas = json.loads((FIGS / "datos_anexos.json").read_text())["F"]["param"]["wacc_tas"]
    d_wacc_tas = wacc_tas - m["wacc_t"]
    dcf = wb["DCF"]
    base = proyectar(m)[2]
    assert abs(base - dcf["C49"].value) < 1e-6, "la réplica no reproduce el DCF v3"

    def aplicar(x, kw, a):
        """Aplica una alternativa al modelo copiado (x) o a los argumentos de proyectar (kw)."""
        for k, v in a.items():
            if k == "tax23":
                x["tax"], x["tax_lp"] = np.minimum(m["tax"], IMPUESTO_LEY), IMPUESTO_LEY
            elif k == "tax27":
                x["tax"] = np.full(len(m["tax"]), m["tax_lp"])
            elif k == "ct0":
                x["k_ct"] = 0.0
            else:
                kw[k] = v

    def valor(alts):
        x, kw = copy.deepcopy(m), {}
        for a in alts:
            aplicar(x, kw, a)
        return objetivo(proyectar(x, **kw)[2])

    alternativas = [("tax23", dict(tax23=True)), ("margen", dict(d_margen=d_margen)), ("ct0", dict(ct0=True)),
                    ("g", dict(g=G_ALTO)), ("wacc", dict(d_wacc=d_wacc_tas))]
    acum = [objetivo(base)] + [valor([a for _, a in alternativas[: i + 1]]) for i in range(len(alternativas))]
    indiv = [(c, valor([a]) - objetivo(base)) for c, a in alternativas + [("kd", dict(d_wacc=DW_KD)), ("goodwill", dict(tax27=True))]]
    # escudo tributario del goodwill / pérdidas: (27% − tasa en caja) × EBIT, 2S26E–2031E
    ebit = np.array([dcf[c + "14"].value for c in "CDEFGH"])
    tasa = np.array([dcf[c + "15"].value for c in "CDEFGH"])
    escudo = (m["tax_lp"] - tasa) * ebit
    return {"base": objetivo(base), "acum": acum, "indiv": indiv, "precio": precio, "dcf": base,
            "escudo_mm": escudo.tolist(), "escudo_total_mm": float(escudo.sum()), "escudo_ps": float(escudo.sum() / m["acciones"]),
            "tasa_caja": tasa.tolist(), "tax_lp": m["tax_lp"]}


def figura(lang, r, salida):
    """Cascada del precio objetivo + tabla de efectos individuales con fichas de color (mismo color que el paso de la cascada)."""
    from matplotlib.patches import FancyBboxPatch
    t = T[lang]
    fig = plt.figure(figsize=(3550 / 1440, 2.62), dpi=400)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0.115, 0.505, 0.865, 0.455])
    acum = r["acum"]
    n = len(acum)
    precio = r["precio"]
    ax.set_axisbelow(True)
    ax.axhspan(acum[-1], precio, xmin=0, xmax=1, color=ROJO, alpha=0.06, linewidth=0)          # brecha frente al precio
    ax.bar(0, acum[0], width=0.66, color=AZUL, linewidth=0, zorder=3)
    ax.text(0, acum[0] + 45, num(acum[0], lang), ha="center", va="bottom", fontsize=4.8, fontweight="bold", color=AZUL)
    for k in range(1, n):
        d = acum[k] - acum[k - 1]
        col = CELESTE if d > 0 else ROJO
        ax.bar(k, d, bottom=acum[k - 1], width=0.66, color=col, linewidth=0, zorder=3)
        ax.plot([k - 1 + 0.33, k - 0.33], [acum[k - 1]] * 2, color=TINTA_SECUNDARIA, linewidth=0.35, zorder=2)
        ax.text(k, max(acum[k], acum[k - 1]) + 45, ("+" if d > 0 else "−") + num(abs(d), lang), ha="center", va="bottom",
                fontsize=4.5, fontweight="bold", color=AZUL if d > 0 else ROJO)
    ax.plot([n - 1 + 0.33, n - 0.33], [acum[-1]] * 2, color=TINTA_SECUNDARIA, linewidth=0.35, zorder=2)
    ax.bar(n, acum[-1], width=0.66, color=NEGRO, linewidth=0, zorder=3)
    ax.text(n, acum[-1] + 45, num(acum[-1], lang), ha="center", va="bottom", fontsize=4.8, fontweight="bold", color=NEGRO)
    ax.axhline(precio, color=ROJO, linewidth=0.9, linestyle=(0, (4, 2)), zorder=4)
    ax.text(-0.5, precio + 50, t["precio"].format(p=num(precio, lang)), ha="left", va="bottom", fontsize=4.8, fontweight="bold", color=ROJO)
    # llave de la brecha entre «todo a favor» y el precio
    xb = n + 0.52
    ax.annotate("", xy=(xb, precio), xytext=(xb, acum[-1]), arrowprops=dict(arrowstyle="<->", color=ROJO, linewidth=0.6, shrinkA=0, shrinkB=0), annotation_clip=False)
    ax.text(xb - 0.08, (precio + acum[-1]) / 2, t["brecha"].format(b=num(precio - acum[-1], lang)), ha="right", va="center",
            fontsize=4.6, fontweight="bold", color=ROJO)
    ax.set_xticks(range(n + 1), t["pasos"], fontsize=4.1, linespacing=1.05)
    ax.set_ylim(1500, 3900)
    ax.set_xlim(-0.55, n + 0.62)
    ax.yaxis.set_major_formatter(lambda y, _: num(y, lang))
    ax.set_yticks([1500, 2000, 2500, 3000, 3500])
    ax.tick_params(labelsize=4.1, length=0, colors=TINTA_SECUNDARIA, pad=1.5)
    ax.grid(axis="y", color=GRILLA, linewidth=0.4)
    for s_ in ("top", "right", "left", "bottom"):
        ax.spines[s_].set_visible(False)
    ax.set_ylabel(t["mm"], fontsize=4.1, color=TINTA_SECUNDARIA, labelpad=2)
    # tabla de efectos individuales (fichas)
    axt = fig.add_axes([0.03, 0.01, 0.94, 0.4])
    axt.axis("off")
    axt.set_xlim(0, 1); axt.set_ylim(0, 1)
    axt.add_patch(FancyBboxPatch((0, 0.875), 1, 0.125, boxstyle="round,pad=0,rounding_size=0.02", color=AZUL, linewidth=0))
    cx = [0.045, 0.80, 0.975]
    for x, c, a in zip(cx, t["cols"], ["left", "right", "right"]):
        axt.text(x, 0.938, c, ha=a, va="center", fontsize=4.4, fontweight="bold", color="white")
    nf = len(r["indiv"])
    alto = 0.84 / nf
    for k, (clave, d) in enumerate(r["indiv"]):
        y0 = 0.84 - (k + 1) * alto
        yc = y0 + alto / 2
        en_contra = clave in ("kd", "goodwill")
        if en_contra:
            axt.add_patch(Rectangle((0, y0 + 0.006), 1, alto - 0.012, color=AMARILLO, alpha=0.28, linewidth=0))
        elif k % 2 == 0:
            axt.add_patch(Rectangle((0, y0 + 0.006), 1, alto - 0.012, color=CELESTE, alpha=0.12, linewidth=0))
        col = CELESTE if d > 0 else ROJO
        axt.add_patch(Rectangle((0.012, yc - 0.03), 0.014, 0.06, color=col, linewidth=0))
        axt.text(0.045, yc, t["filas"][k], ha="left", va="center", fontsize=4.2, color=NEGRO)
        sg = "+" if d > 0 else "−"
        axt.text(0.80, yc, sg + num(abs(d), lang), ha="right", va="center", fontsize=4.3, fontweight="bold", color=AZUL if d > 0 else ROJO)
        axt.text(0.975, yc, sg + num(abs(d) / r["base"] * 100, lang, 1) + "%", ha="right", va="center", fontsize=4.3, color=AZUL if d > 0 else ROJO)
    fig.savefig(salida, facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    v3, fa, tor = sys.argv[1:4]
    salida = Path(sys.argv[4]) if len(sys.argv) > 4 else FIGS
    r = calcular(v3, fa, tor)
    (FIGS / "datos_supuestos.json").write_text(json.dumps(r, indent=1, ensure_ascii=False))
    for lang in ("es", "en"):
        figura(lang, r, salida / f"{lang}_V13_supuestos_cascada.png")
    print({k: round(v, 1) for k, v in zip(["base"] + [c for c, _ in r["indiv"]], [r["base"]] + [d for _, d in r["indiv"]])}, [round(a) for a in r["acum"]],
          round(r["escudo_total_mm"]), round(r["escudo_ps"], 1))
