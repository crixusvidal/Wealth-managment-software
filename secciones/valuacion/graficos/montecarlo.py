"""Simulación de Monte Carlo del precio objetivo sobre la réplica del DCF del equipo (riesgos/graficos/tornado_operacional.py,
que reproduce exactamente el DCF del modelo de valoración). Reemplaza las probabilidades cualitativas de los escenarios.

Qué hace
  1. VALIDACIÓN: con todas las variables en su valor base, el DCF debe dar CLP 1.859 y el precio objetivo CLP 2.077; si no, se detiene.
  2. Simula N = 10.000 iteraciones con semilla fija (SEMILLA = 42) para que el resultado sea reproducible.
  3. Guarda estadísticos, la distribución y el tornado de contribución a la varianza en figuras/datos_montecarlo.json y dibuja
     {es,en}_V16_montecarlo.png (histograma) y {es,en}_V17_montecarlo_tornado.png (contribución a la varianza).

Variables y distribuciones (mínimos y máximos históricos 4T23–2T26 del tornado de Riesgos)
  Triangulares (mínimo, base, máximo): margen EBITDA, ocupación, arriendo mismas tiendas 2027–28 (SSR), ventas de locatarios 2027–28
  (SSS; llegan al ingreso por la renta variable), crecimiento real del arriendo 2029–31 (piso / base / techo de la Nota 11),
  visitas 2027–28, capex de mantención + TI (4,8% / 5,5% / 6,6% de ingresos) y g nominal (2,5% / 3,0% / 4,0%).
  WACC: desplazamiento paralelo normal, media 0, desviación 50 pb, truncado en ±150 pb (todos los períodos).
  Impuesto: caso base (27% de largo plazo); la rebaja a 23% no entra y se reporta aparte.
  Correlación 0,5 entre SSR y SSS (cópula gaussiana, supuesto del equipo); el resto de las variables es independiente.
  Precio objetivo de cada iteración = 80% × DCF + 20% × múltiplos fijos (CLP 2.951).

Uso (desde secciones/): python valuacion/graficos/montecarlo.py
"""
import json
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from scipy.stats import norm, rankdata

RAIZ = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(RAIZ / "riesgos" / "graficos"))
sys.path.insert(0, str(RAIZ / "comun"))
from tornado_operacional import CAPEX_MANT, LARGO, NOTA11, leer_modelo, leer_rangos, proyectar  # noqa: E402
import estilo  # noqa: E402,F401  (paleta institucional)
from estilo import AMARILLO, AZUL, CELESTE, GRILLA, NEGRO, ROJO, TINTA_SECUNDARIA, num  # noqa: E402

N, SEMILLA = 10_000, 42
RHO_SS = 0.5                       # correlación SSR–SSS (cópula gaussiana)
SIGMA_WACC, TOPE_WACC = 0.005, 0.015
G_TRI = (0.025, 0.030, 0.040)
ANCHO, DPI = 3550 / 1440, 400      # mismo ancho de columna y resolución que las demás figuras
FIGS = Path(__file__).parent / "figuras"
V4 = RAIZ / "valuacion/modelos/Mallplaza_Valoracion_Consolidada_v4.xlsx"
FA = RAIZ / "analisis_financiero/modelos/Mallplaza_FA_Exhibits_v7.xlsx"
TOR = RAIZ / "riesgos/modelos/Mallplaza_Tornado_Sensibilidad_Operacional_2026_2028.xlsx"


def tri_inv(u, lo, modo, hi):
    """Inversa de la distribución triangular (mínimo, moda, máximo) para u ~ U(0,1)."""
    fc = (modo - lo) / (hi - lo)
    return np.where(u < fc, lo + np.sqrt(u * (hi - lo) * (modo - lo)), hi - np.sqrt((1 - u) * (hi - lo) * (hi - modo)))


def simular():
    m, control = leer_modelo(str(FA), str(V4))
    wb = openpyxl.load_workbook(V4, data_only=True)["Precio_Objetivo"]
    mult, peso, precio = wb["B13"].value, wb["B14"].value, wb["B16"].value
    consenso = json.loads((RAIZ / "informe_final/graficos/cifras.json").read_text())["consenso"]["po_12m"]
    # 1. validación del caso base
    base = proyectar(m)[2]
    po_base = peso * base + (1 - peso) * mult
    assert abs(base - control["valor"]) < 1e-6 and round(base) == 1859 and round(po_base) == 2077, \
        f"la réplica no reproduce el caso base (DCF {base:.2f}, objetivo {po_base:.2f}): se detiene"
    rg = leer_rangos(str(TOR))
    rng = np.random.default_rng(SEMILLA)
    # 2. números aleatorios: uniformes independientes y el par SSR–SSS correlacionado (cópula gaussiana)
    U = {k: rng.random(N) for k in ("margen", "occ", "visitas", "largo", "capex", "g")}
    z = rng.standard_normal((N, 2))
    z[:, 1] = RHO_SS * z[:, 0] + np.sqrt(1 - RHO_SS ** 2) * z[:, 1]
    u_ssr, u_sss = norm.cdf(z[:, 0]), norm.cdf(z[:, 1])
    wacc = np.clip(rng.normal(0, SIGMA_WACC, N), -TOPE_WACC, TOPE_WACC)   # truncado por recorte en ±150 pb
    # 3. valores de cada variable (el desplazamiento se mide contra el valor base del tornado)
    def tri(nombre, u):
        b, lo, hi = rg[nombre]
        return tri_inv(u, lo, b, hi) - b
    d_margen, d_occ, d_park = tri("Margen EBITDA", U["margen"]), tri("Ocupación", U["occ"]), tri("Visitas — crecimiento orgánico", U["visitas"])
    d_ssr, d_sss = tri("Same-store rent growth (SSR)", u_ssr), tri("Same-store sales growth (SSS)", u_sss)
    mant = tri_inv(U["capex"], CAPEX_MANT[0], m["mant"], CAPEX_MANT[1])
    g = tri_inv(U["g"], *G_TRI)
    base_largo = {p: np.asarray(x["ss"])[LARGO].copy() for p, x in m["paises"].items()}
    modo_largo = {p: float(v.mean()) for p, v in base_largo.items()}
    d_largo = {p: tri_inv(U["largo"], NOTA11[p][0], modo_largo[p], NOTA11[p][1]) - modo_largo[p] for p in NOTA11}
    dcf = np.empty(N)
    for i in range(N):
        sl = {p: base_largo[p] + d_largo[p][i] for p in NOTA11}
        dcf[i] = proyectar(m, d_occ=d_occ[i], d_ss_corto=d_ssr[i] + m["share_var"] * d_sss[i], d_park_corto=d_park[i], d_margen=d_margen[i],
                           mant=mant[i], ss_largo=sl, d_wacc=wacc[i], g=g[i])[2]
    po = peso * dcf + (1 - peso) * mult
    entradas = {"Margen EBITDA": d_margen, "Ocupación": d_occ, "Arriendo mismas tiendas 2027–28": d_ssr, "Ventas de locatarios 2027–28": d_sss,
                "Crecimiento real 2029–31": d_largo["cl"], "Visitas 2027–28": d_park, "Capex de mantención + TI": mant, "WACC": wacc, "g nominal": g}
    rp = rankdata(po)
    rho = {k: float(np.corrcoef(rankdata(v), rp)[0, 1]) for k, v in entradas.items()}
    tot = sum(r ** 2 for r in rho.values())
    contrib = {k: r ** 2 / tot for k, r in rho.items()}
    q = lambda p: float(np.percentile(po, p))
    return {"n": N, "semilla": SEMILLA, "dcf_base": base, "po_base": po_base, "precio": precio, "consenso": consenso, "mult": mult, "peso": peso,
            "stats": {"media": float(po.mean()), "mediana": q(50), "desv": float(po.std(ddof=1)), "p5": q(5), "p25": q(25), "p75": q(75), "p95": q(95),
                      "min": float(po.min()), "max": float(po.max())},
            "dcf": {"media": float(dcf.mean()), "mediana": float(np.median(dcf)), "p5": float(np.percentile(dcf, 5)), "p95": float(np.percentile(dcf, 95))},
            "p_bajo_precio": float((po < precio).mean()), "p_sobre_precio": float((po >= precio).mean()), "p_sobre_consenso": float((po >= consenso).mean()),
            "p_bajo_objetivo": float((po < po_base).mean()), "rho_rank": rho, "contribucion": contrib,
            "supuestos": {"rho_ssr_sss": RHO_SS, "sigma_wacc_pb": SIGMA_WACC * 1e4, "tope_wacc_pb": TOPE_WACC * 1e4, "g": list(G_TRI), "capex": [CAPEX_MANT[0], m["mant"], CAPEX_MANT[1]],
                          "rangos": {k: list(v) for k, v in rg.items()}, "nota11": {p: list(v) for p, v in NOTA11.items()}}}, po


T = {"es": {"bajo": "{p}% de las\nsimulaciones\nbajo el precio", "eje": "Precio objetivo simulado (CLP por acción)", "y": "Simulaciones (%)", "precio": "Precio {v}",
            "obj": "Objetivo {v}", "cons": "Consenso {v}", "var": "Contribución a la varianza del precio objetivo (%)",
            "nom": {"Margen EBITDA": "Margen EBITDA", "Ocupación": "Ocupación", "Arriendo mismas tiendas 2027–28": "Arriendo mismas tiendas 2027–28",
                    "Ventas de locatarios 2027–28": "Ventas de locatarios 2027–28", "Crecimiento real 2029–31": "Crecimiento real 2029–31", "Visitas 2027–28": "Visitas 2027–28",
                    "Capex de mantención + TI": "Capex mantención + TI", "WACC": "WACC", "g nominal": "g nominal"}},
     "en": {"bajo": "{p}% of\nsimulations\nbelow the price", "eje": "Simulated target price (CLP per share)", "y": "Simulations (%)", "precio": "Price {v}",
            "obj": "Target {v}", "cons": "Consensus {v}", "var": "Contribution to target-price variance (%)",
            "nom": {"Margen EBITDA": "EBITDA margin", "Ocupación": "Occupancy", "Arriendo mismas tiendas 2027–28": "Same-store rent 2027–28",
                    "Ventas de locatarios 2027–28": "Tenant sales 2027–28", "Crecimiento real 2029–31": "Real rent growth 2029–31", "Visitas 2027–28": "Visitors 2027–28",
                    "Capex de mantención + TI": "Maintenance + IT capex", "WACC": "WACC", "g nominal": "Nominal g"}}}


def histograma(r, po, lang, salida):
    t = T[lang]
    fig = plt.figure(figsize=(ANCHO, 2.05))
    ax = fig.add_axes([0.10, 0.17, 0.87, 0.74])
    pesos = np.ones_like(po) * 100 / len(po)
    ax.hist(po, bins=40, weights=pesos, color=AZUL, edgecolor="white", linewidth=0.3, zorder=3)
    top = ax.get_ylim()[1]
    ax.set_xlim(1400, 5100)
    for v, c, txt, dy, ha, dx in ((r["po_base"], ROJO, t["obj"], 0.97, "left", 25), (r["precio"], NEGRO, t["precio"], 0.80, "right", -25),
                                  (r["consenso"], AMARILLO, t["cons"], 0.97, "left", 25)):
        ax.axvline(v, color=c, linewidth=0.9, linestyle=(0, (4, 2)), zorder=4)
        ax.text(v + dx, top * dy, txt.format(v=num(v, lang)), fontsize=4.2, color=NEGRO if c == AMARILLO else c, va="top", ha=ha, fontweight="bold")
    ax.text(0.34, 0.62, t["bajo"].format(p=num(round(r["p_bajo_precio"] * 100, 1), lang, 0 if r["p_bajo_precio"] > 0.9995 else 1)), transform=ax.transAxes, fontsize=4.4,
            color=AZUL, fontweight="bold", va="top", ha="left")
    ax.set_xlabel(t["eje"], fontsize=4.2, color=TINTA_SECUNDARIA, labelpad=2)
    ax.set_ylabel(t["y"], fontsize=4.2, color=TINTA_SECUNDARIA, labelpad=2)
    ax.tick_params(labelsize=4.0, length=1.5, colors=TINTA_SECUNDARIA)
    ax.xaxis.set_major_formatter(lambda v, _: num(v, lang))
    ax.grid(axis="y", color=GRILLA, linewidth=0.4, zorder=0)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


def tornado(r, lang, salida):
    t = T[lang]
    items = sorted(r["contribucion"].items(), key=lambda kv: kv[1])
    fig = plt.figure(figsize=(ANCHO, 2.05))
    ax = fig.add_axes([0.36, 0.17, 0.60, 0.76])
    ax.barh([t["nom"][k] for k, _ in items], [v * 100 for _, v in items], color=AZUL, height=0.65, zorder=3)
    for i, (_, v) in enumerate(items):
        ax.text(v * 100 + 0.8, i, num(v * 100, lang, 1), va="center", fontsize=4.2, color=NEGRO)
    ax.set_xlabel(t["var"], fontsize=4.2, color=TINTA_SECUNDARIA, labelpad=2)
    ax.tick_params(labelsize=4.2, length=1.5, colors=TINTA_SECUNDARIA)
    ax.grid(axis="x", color=GRILLA, linewidth=0.4, zorder=0)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)


if __name__ == "__main__":
    plt.rcParams.update({"font.family": "DejaVu Sans"})
    r, po = simular()
    (FIGS / "datos_montecarlo.json").write_text(json.dumps(r, ensure_ascii=False, indent=1))
    for lang in ("es", "en"):
        histograma(r, po, lang, FIGS / f"{lang}_V16_montecarlo.png")
        tornado(r, lang, FIGS / f"{lang}_V17_montecarlo_tornado.png")
    s = r["stats"]
    print({"validación": [round(r["dcf_base"], 1), round(r["po_base"], 1)], "media": round(s["media"]), "mediana": round(s["mediana"]), "p5": round(s["p5"]), "p95": round(s["p95"]),
           "P(<precio)": round(r["p_bajo_precio"], 4), "P(>=consenso)": round(r["p_sobre_consenso"], 4)})
    print({k: round(v, 3) for k, v in sorted(r["contribucion"].items(), key=lambda kv: -kv[1])})
