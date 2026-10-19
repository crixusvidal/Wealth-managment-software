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
    t = T[lang]
    fig = plt.figure(figsize=(3550 / 1440, 2.3), dpi=400)
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0.13, 0.48, 0.85, 0.5])
    acum = r["acum"]
    n = len(acum)
    ax.bar(0, acum[0], width=0.62, color=AZUL, linewidth=0)
    ax.text(0, acum[0] + 40, num(acum[0], lang), ha="center", va="bottom", fontsize=4.4, fontweight="bold", color=AZUL)
    for k in range(1, n):
        d = acum[k] - acum[k - 1]
        ax.bar(k, d, bottom=acum[k - 1], width=0.62, color=CELESTE if d > 0 else ROJO, linewidth=0)
        ax.text(k, max(acum[k], acum[k - 1]) + 40, ("+" if d > 0 else "−") + num(abs(d), lang), ha="center", va="bottom", fontsize=4.4, color=NEGRO)
        ax.plot([k - 0.69, k - 0.31], [acum[k - 1]] * 2, color=NEGRO, linewidth=0.3)
    ax.bar(n, acum[-1], width=0.62, color=NEGRO, linewidth=0)
    ax.text(n, acum[-1] + 40, num(acum[-1], lang), ha="center", va="bottom", fontsize=4.6, fontweight="bold", color=NEGRO)
    ax.axhline(r["precio"], color=ROJO, linewidth=0.8, linestyle=(0, (4, 2)))
    ax.text(-0.45, r["precio"] + 40, t["precio"].format(p=num(r["precio"], lang)), ha="left", va="bottom", fontsize=4.6, fontweight="bold", color=ROJO)
    ax.text(n + 0.4, r["precio"] + 40, t["brecha"].format(b=num(r["precio"] - acum[-1], lang)), ha="right", va="bottom", fontsize=4.4, color=ROJO)
    ax.set_xticks(range(n + 1), t["pasos"], fontsize=4.2, linespacing=1.05)
    ax.set_ylim(1500, 3900)
    ax.set_xlim(-0.55, n + 0.55)
    ax.yaxis.set_major_formatter(lambda y, _: num(y, lang))
    ax.tick_params(labelsize=4.2, length=1.5, colors=TINTA_SECUNDARIA)
    ax.grid(axis="y", color=GRILLA, linewidth=0.4)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_ylabel(t["mm"], fontsize=4.2, color=TINTA_SECUNDARIA, labelpad=2)
    # tabla de efectos individuales
    axt = fig.add_axes([0.03, 0.0, 0.94, 0.36])
    axt.axis("off")
    xs, alin = [0.0, 0.87, 0.99], ["left", "right", "right"]
    axt.add_patch(Rectangle((0, 0.87), 1, 0.13, color=AZUL, transform=axt.transAxes))
    for x, c, a in zip(xs, t["cols"], alin):
        axt.text(x + (0.01 if a == "left" else 0), 0.935, c, ha=a, va="center", fontsize=4.6, fontweight="bold", color="white", transform=axt.transAxes)
    for k, (clave, d) in enumerate(r["indiv"]):
        y = 0.79 - k * 0.1175
        if clave in ("kd", "goodwill"):
            axt.add_patch(Rectangle((0, y - 0.055), 1, 0.11, color=AMARILLO, linewidth=0, transform=axt.transAxes))
        vals = [t["filas"][k], ("+" if d > 0 else "−") + num(abs(d), lang),
                ("+" if d > 0 else "−") + num(abs(d) / r["base"] * 100, lang, 1) + "%"]
        for x, v, a in zip(xs, vals, alin):
            axt.text(x + (0.01 if a == "left" else 0), y, v, ha=a, va="center", fontsize=4.6, color=NEGRO, transform=axt.transAxes)
    axt.set_xlim(0, 1)
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
