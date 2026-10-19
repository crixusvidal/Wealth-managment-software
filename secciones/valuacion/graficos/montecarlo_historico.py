"""Monte Carlo con distribuciones tomadas de los datos históricos del libro Mallplaza_Datos_Graficos_Informe.xlsx (Bloomberg y Banco Central).

  V21  Valor por acción implícito en el cap rate = BCU 10 años + spread. El spread se remuestrea de su historia 4T22–3T25 (11 trimestres) y el BCU
       de su historia 4T22–30-sep-26 (15 puntos), ambos con suavizado de núcleo gaussiano (regla de Silverman); 10.000 iteraciones, semilla 42.
       Valor = (EBITDA UDM 2T26 / cap rate − deuda neta − minoritarios) / acciones (hoja Fig_CapRate).
  V22  Precio de la acción a 52 semanas: remuestreo (bootstrap) de los 261 retornos logarítmicos semanales de Mallplaza (oct-21 a sep-26),
       centrados en cero para no extrapolar el rally; 10.000 trayectorias, semilla 42, desde CLP 3.555.
Además genera la copia «_titulado» (con nombre y fuente dentro de la imagen) de las figuras V16–V22.

Uso (desde secciones/): python valuacion/graficos/montecarlo_historico.py
"""
import json
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))
import estilo  # noqa: E402,F401
from estilo import AMARILLO, AZUL, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA, num  # noqa: E402
from titulado import titulado  # noqa: E402
import datos_mercado as dm  # noqa: E402

N, SEMILLA, PRECIO, OBJETIVO, CONSENSO = 10_000, 42, 3555.0, 2077.2, 4111.0
ANCHO, DPI = 3550 / 1440, 400
FIGS = Path(__file__).parent / "figuras"


def hist_base(lang, datos, lineas, xlabel, salida, xlim, anot):
    fig = plt.figure(figsize=(ANCHO, 2.05))
    ax = fig.add_axes([0.10, 0.17, 0.87, 0.74])
    ax.hist(datos, bins=45, weights=np.ones_like(datos) * 100 / len(datos), color=AZUL, edgecolor="white", linewidth=0.3, zorder=3)
    top = ax.get_ylim()[1]
    for v, col, txt, dy, ha, dx in lineas:
        ax.axvline(v, color=col, linewidth=0.9, linestyle=(0, (4, 2)), zorder=4)
        ax.text(v + dx, top * dy, txt, fontsize=4.2, color=NEGRO if col == AMARILLO else col, va="top", ha=ha, fontweight="bold")
    ax.text(0.95, 0.62, anot, transform=ax.transAxes, fontsize=4.4, color=AZUL, fontweight="bold", va="top", ha="right", linespacing=1.3)
    ax.set_xlim(*xlim)
    ax.set_xlabel(xlabel, fontsize=4.2, color=TINTA_SECUNDARIA, labelpad=2)
    ax.set_ylabel("Simulaciones (%)" if lang == "es" else "Simulations (%)", fontsize=4.2, color=TINTA_SECUNDARIA, labelpad=2)
    ax.tick_params(labelsize=4.0, length=1.5, colors=TINTA_SECUNDARIA)
    ax.xaxis.set_major_formatter(lambda v, _: num(v, lang))
    ax.grid(axis="y", color=GRILLA, linewidth=0.4, zorder=0)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


def kde_muestra(obs, n, rng):
    obs = np.asarray(obs, float)
    h = 1.06 * obs.std(ddof=1) * len(obs) ** (-1 / 5)          # regla de Silverman
    return rng.choice(obs, n) + rng.normal(0, h, n)


def correr():
    wb = openpyxl.load_workbook(dm.LIBRO, data_only=True)
    cr = [r for r in wb["Fig_CapRate"].iter_rows(min_row=3, values_only=True) if hasattr(r[1], "year") and r[8] is not None]
    spread = [r[10] for r in cr if r[0] not in ("4T25", "1T26", "2T26", "30-sep-26")]      # 4T22–3T25
    bcu = [r[9] * 100 for r in cr]
    ebitda, neto, acc = cr[-1][6], cr[-1][3] + cr[-1][4], 2190.0
    rng = np.random.default_rng(SEMILLA)
    cap = (kde_muestra(bcu, N, rng) + kde_muestra(spread, N, rng) / 100) / 100
    valor = (ebitda / cap - neto) / acc
    pr = [r for r in wb["Fig_Precio"].iter_rows(min_row=4, values_only=True) if hasattr(r[0], "year") and isinstance(r[1], (int, float))]
    p = np.array([r[1] for r in pr], float)
    ret = np.diff(np.log(p)); ret = ret - ret.mean()
    final = PRECIO * np.exp(rng.choice(ret, (N, 52)).sum(axis=1))
    st = lambda x: {"media": float(x.mean()), "mediana": float(np.median(x)), "p5": float(np.percentile(x, 5)), "p95": float(np.percentile(x, 95))}
    return valor, final, {"valor_implicito": {**st(valor), "p_bajo_precio": float((valor < PRECIO).mean()), "n_spread": len(spread), "n_bcu": len(bcu),
                                              "spread_prom": float(np.mean(spread)), "spread_desv": float(np.std(spread, ddof=1))},
                          "precio_52s": {**st(final), "p_bajo_objetivo": float((final <= OBJETIVO).mean()), "p_bajo_precio": float((final < PRECIO).mean()),
                                         "vol_anual": float(ret.std(ddof=1) * np.sqrt(52)), "n_retornos": len(ret)}}


if __name__ == "__main__":
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    valor, final, r = correr()
    (FIGS / "datos_montecarlo_historico.json").write_text(json.dumps(r, ensure_ascii=False, indent=1))
    v, pr = r["valor_implicito"], r["precio_52s"]
    for lang in ("es", "en"):
        es = lang == "es"
        hist_base(lang, valor, [(OBJETIVO, ROJO, ("Objetivo " if es else "Target ") + num(OBJETIVO, lang), 0.97, "right", -40), (PRECIO, NEGRO, ("Precio " if es else "Price ") + num(PRECIO, lang), 0.97, "left", 40)],
                  "Valor por acción implícito (CLP)" if es else "Implied value per share (CLP)", FIGS / f"{lang}_V21_mc_cap_rate.png", (1000, 4400),
                  f"{num(v['p_bajo_precio'] * 100, lang, 0)}% de las\nsimulaciones\nbajo el precio" if es else f"{num(v['p_bajo_precio'] * 100, lang, 0)}% of\nsimulations\nbelow the price")
        hist_base(lang, final, [(OBJETIVO, ROJO, ("Objetivo\n" if es else "Target\n") + num(OBJETIVO, lang), 0.97, "left", 40), (PRECIO, NEGRO, ("Precio actual " if es else "Current price ") + num(PRECIO, lang), 0.985, "left", 40)],
                  "Precio simulado a 52 semanas (CLP)" if es else "Simulated price at 52 weeks (CLP)", FIGS / f"{lang}_V22_mc_precio_52s.png", (1000, 7500),
                  f"{num(pr['p_bajo_objetivo'] * 100, lang, 0)}% de las\ntrayectorias\ntermina bajo\nel objetivo" if es else f"{num(pr['p_bajo_objetivo'] * 100, lang, 0)}% of paths\nend below\nthe target")
        # copias con nombre y fuente
        T = {
            "V16_montecarlo": ("Distribución del precio objetivo en 10.000 simulaciones" if es else "Target price distribution in 10,000 simulations", "Precio objetivo = 80% DCF + 20% múltiplos; CLP por acción" if es else "Target price = 80% DCF + 20% multiples; CLP per share",
                               ["Distribuciones triangulares con los extremos históricos 4T23–2T26 del tornado de Riesgos; WACC normal (σ 50 pb); semilla 42. Fuente: análisis del equipo." if es else "Triangular distributions with the 4Q23–2Q26 historical extremes of the Risks tornado; normal WACC shift (σ 50 bp); seed 42. Source: Team analysis."]),
            "V17_montecarlo_tornado": ("Qué mueve el precio objetivo simulado" if es else "What moves the simulated target price", "Contribución a la varianza (correlación de rangos, %)" if es else "Contribution to variance (rank correlation, %)",
                                       ["Fuente: simulación de Monte Carlo del equipo (10.000 iteraciones, semilla 42) sobre la réplica del DCF." if es else "Source: Team Monte Carlo simulation (10,000 iterations, seed 42) on the DCF replica."]),
            "V18_caprate_spread": ("El premio sobre la tasa real libre de riesgo se redujo a la mitad: de ~615 pb a 335 pb" if es else "The premium over the real risk-free rate has halved: from ~615 bp to 335 bp", "Cap rate pro forma de Mallplaza y BCU 10 años, % real; spread en pb" if es else "Mallplaza pro forma cap rate and 10-year BCU, % real; spread in bp",
                                  ["Cap rate = EBITDA UDM / EV. 4T24–3T25 pro forma con Perú 12 meses (estimación del equipo). BCU 10 años: promedio mensual del mes de cierre. Fuente: Bloomberg; Banco Central de Chile; análisis del equipo." if es else "Cap rate = LTM EBITDA / EV. 4Q24–3Q25 pro forma with 12 months of Peru (team estimate). 10-year BCU: monthly average of the closing month. Source: Bloomberg; Central Bank of Chile; Team analysis."]),
            "V19_precio_base100": ("Mallplaza subió ~4x en cinco años, más que el IPSA; el último tramo llegó sin cambios en el negocio" if es else "Mallplaza rose ~4x in five years, ahead of the IPSA; the last leg came with no change in the business", "Precio de cierre semanal, base 100 = 1-oct-21" if es else "Weekly closing price, base 100 = 1-Oct-21",
                                   ["Ajustado por el aumento de capital de 2024. IPSA hasta 31-ago-26 (último dato disponible). Hitos: cierres diarios. Fuente: Bloomberg; análisis del equipo." if es else "Adjusted for the 2024 capital increase. IPSA to 31-Aug-26 (latest available). Milestones: daily closes. Source: Bloomberg; Team analysis."]),
            "V20_pnav": ("Aun sumando todos los impuestos diferidos, el precio está 27% sobre el NAV" if es else "Even adding back all deferred taxes, the price is 27% above NAV", "Precio / NAV ajustado (EPRA NTA)" if es else "Price / adjusted NAV (EPRA NTA)",
                         ["NAV ajustado = patrimonio atribuible + pasivos por impuestos diferidos − goodwill − intangibles. Balance al 30-jun-26; capitalización al 30-sep-26. Fuente: Bloomberg; EEFF CMF; análisis del equipo." if es else "Adjusted NAV = attributable equity + deferred tax liabilities − goodwill − intangibles. Balance sheet at 30-Jun-26; market cap at 30-Sep-26. Source: Bloomberg; CMF financial statements; Team analysis."]),
            "V21_mc_cap_rate": (("Si el spread del cap rate volviera a su historia, la acción valdría ~CLP " + num(round(v["mediana"], -2), "es") + " (mediana)") if es else ("If the cap-rate spread returned to its history, the stock would be worth ~CLP " + num(round(v["mediana"], -2), "en") + " (median)"), "Valor por acción implícito en 10.000 simulaciones (CLP)" if es else "Implied value per share in 10,000 simulations (CLP)",
                                ["Cap rate = BCU 10 años + spread; ambos remuestreados de su historia (spread 4T22–3T25, 11 trimestres; BCU 4T22–sep-26, 15 puntos) con suavizado gaussiano; EBITDA UDM 2T26 y deuda neta de jun-26. Fuente: Bloomberg; Banco Central de Chile; análisis del equipo." if es else "Cap rate = 10-year BCU + spread; both resampled from their history (spread 4Q22–3Q25, 11 quarters; BCU 4Q22–Sep-26, 15 points) with Gaussian smoothing; LTM EBITDA 2Q26 and Jun-26 net debt. Source: Bloomberg; Central Bank of Chile; Team analysis."]),
            "V22_mc_precio_52s": ("Distribución del precio a 52 semanas según la volatilidad histórica" if es else "52-week price distribution implied by historical volatility", "10.000 trayectorias desde CLP 3.555 (CLP por acción)" if es else "10,000 paths from CLP 3,555 (CLP per share)",
                                  ["Remuestreo de los retornos semanales de Mallplaza (oct-21 a sep-26), centrados en cero para no extrapolar el rally; volatilidad anual ~" + num(pr['vol_anual'] * 100, 'es', 0) + "%. Fuente: Bloomberg; análisis del equipo." if es else "Bootstrap of Mallplaza weekly returns (Oct-21 to Sep-26), demeaned so as not to extrapolate the rally; annual volatility ~" + num(pr['vol_anual'] * 100, 'en', 0) + "%. Source: Bloomberg; Team analysis."]),
        }
        for k, (t, s, n) in T.items():
            titulado(FIGS / f"{lang}_{k}.png", t, s, n, FIGS / f"{lang}_{k}_titulado.png")
    print({k: {a: round(b, 3) for a, b in d.items()} for k, d in r.items()})
