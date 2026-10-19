"""Cifras del informe, leídas de los modelos vigentes (sin cifras escritas a mano en los textos).

Lee el modelo de valoración v4, el Exhibits v7 (motor operativo), el libro de figuras y los JSON de las figuras de Valuación,
y vuelca informe_final/graficos/cifras.json: escalares (precios, valor DCF, tasas, escenarios, TIR, múltiplos) y series por período
(DCF, capex, deuda, FFO) que usan los constructores de texto (comun/cifras.js) y los anexos de Análisis Financiero y Valuación.
Sustituye a datos_resumen.py como fuente de las cifras del Resumen de inversión y del Abstract.

Uso (desde secciones/):
    python informe_final/graficos/cifras_informe.py <Valoracion_Consolidada_v4.xlsx> <FA_Exhibits_v7.xlsx> <Figuras_Valuacion_recalc.xlsx>
Salida: informe_final/graficos/cifras.json
"""
import json
import sys
from pathlib import Path

import numpy as np
import openpyxl

RAIZ = Path(__file__).resolve().parents[2]
FIG_VAL = RAIZ / "valuacion" / "graficos" / "figuras"
COLS = "CDEFGH"                 # 2S26E … 2031E
COLS_T = "CDEFGHI"              # + terminal


sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "riesgos" / "graficos"))
from tornado_operacional import LARGO, leer_modelo, proyectar  # noqa: E402


def sensibilidades(ruta_fa, ruta_modelo):
    """Δ del valor DCF por acción (CLP) por +1 pp de cada palanca (réplica del DCF): relación precio (arriendo) → ventas → ingresos."""
    m, _ = leer_modelo(ruta_fa, ruta_modelo)
    base = proyectar(m)[2]
    ss_lp = {p: x["ss"][LARGO].copy() + 0.01 for p, x in m["paises"].items()}
    return {"base": base, "ssr_1pp": proyectar(m, d_ss_corto=0.01)[2] - base,
            "sss_1pp": proyectar(m, d_ss_corto=0.01 * m["share_var"])[2] - base, "share_var": m["share_var"],
            "ssr_lp_1pp": proyectar(m, ss_largo=ss_lp)[2] - base, "occ_1pp": proyectar(m, d_occ=0.01)[2] - base,
            "margen_1pp": proyectar(m, d_margen=0.01)[2] - base, "wacc_100pb": proyectar(m, d_wacc=0.01)[2] - base,
            "wacc_m100pb": proyectar(m, d_wacc=-0.01)[2] - base}


def tabla_fa1(fa, k):
    """Proyección 2026E–2031E de la tabla ROE/DuPont de Análisis Financiero (Figura 1) con el caso base de valoración v4.

    Reglas (las mismas del Exhibits, Forecast filas 148–155, con la inflación de 3,0% del modelo): propiedades de inversión al cierre
    = anterior + revalorización (PI × π) + capex total del año; revalorización neta = bruta × (1 − 27% de impuesto diferido) × (1 − %
    minoritarios); utilidad reportada = utilidad distribuible + revalorización neta; patrimonio = anterior + utilidad reportada −
    dividendos ordinarios − distribuciones adicionales; otros activos constantes (jun-26). 2026E: revalorización del Exhibits (1S26 real)."""
    fx, e1 = fa["Forecast"], fa["E1_DuPont"]
    s = k["s"]
    ing = [s["ing_cal"][0]] + s["ing"][1:6]
    pi, t, mino = k["pi"], fx["D149"].value, fx["D113"].value
    otros = fx["D152"].value
    capex = [fx["D142"].value] + [abs(s["capex_base"][i]) + abs(s["capex_exp"][i]) for i in range(1, 6)]
    ffo, dist, div, adic = s["ffo"], s["distribuible_cal"], s["div_ord_cal"], s["dist_adic_cal"]
    activos, patrim = [e1["H16"].value], [e1["H17"].value]          # dic-25 (2025A)
    pic = fx["D151"].value - fx["D148"].value - capex[0]            # PI dic-25
    ni = []
    for i in range(6):
        reval_bruta = fx["D148"].value if i == 0 else pic * pi
        reval_neta = reval_bruta * (1 - t) * (1 - mino)
        pic = pic + reval_bruta + capex[i]
        ni.append(dist[i] + reval_neta)
        activos.append(pic + otros)
        patrim.append(patrim[-1] + ni[i] - div[i] - adic[i])
    prom = lambda a: [(a[i] + a[i + 1]) / 2 for i in range(6)]
    a_p, p_p = prom(activos), prom(patrim)
    return {"ingresos": ing, "margen_ebitda": [s["ebitda_cal"][0] / ing[0]] + [s["ebitda"][i] / ing[i] for i in range(1, 6)],
            "margen_ffo": [ffo[i] / ing[i] for i in range(6)], "utilidad_reportada": ni, "distribuible": dist,
            "margen_neto_norm": [dist[i] / ing[i] for i in range(6)], "rotacion": [ing[i] / a_p[i] for i in range(6)],
            "apalancamiento": [a_p[i] / p_p[i] for i in range(6)], "roe_rep": [ni[i] / p_p[i] for i in range(6)],
            "roe_norm": [dist[i] / p_p[i] for i in range(6)], "roa_norm": [dist[i] / a_p[i] for i in range(6)],
            "activos": activos, "patrimonio": patrim}


def serie(ws, fila, cols=COLS):
    return [ws[f"{c}{fila}"].value for c in cols]


def main(ruta_modelo, ruta_fa, ruta_fig):
    wb = openpyxl.load_workbook(ruta_modelo, data_only=True)
    fa = openpyxl.load_workbook(ruta_fa, data_only=True)
    fig = openpyxl.load_workbook(ruta_fig, data_only=True)
    s, d, p, w, ex, fb, de, ct, apv = (wb[n] for n in ("Supuestos", "DCF", "Precio_Objetivo", "WACC", "Expansion", "Forecast_Base", "Deuda",
                                                         "Capital_Trabajo", "APV"))
    fx = fa["Forecast"]
    da = json.loads((FIG_VAL / "datos_anexos.json").read_text())
    tir = json.loads((FIG_VAL / "datos_tir.json").read_text())
    tp = json.loads((FIG_VAL / "datos_tir_proyecto.json").read_text())
    sup = json.loads((FIG_VAL / "datos_supuestos.json").read_text())
    acc, precio = s["B9"].value, s["B8"].value
    dfn, nci = s["B10"].value, s["B11"].value
    ebitda_udm = s["B13"].value
    semanas = next(r for r in fig["V1_Football"].iter_rows(min_row=5, max_row=10, values_only=True) if r[0] and str(r[0]).startswith("Rango 52"))
    div27 = de["D16"].value / acc
    po_v = p["B15"].value
    ev_mercado = precio * acc + dfn + nci
    ev_dcf = d["C44"].value
    ev_po = po_v * acc + dfn + nci
    pi = s["B34"].value
    real = lambda x: (1 + x) / (1 + pi) - 1
    esc = da["F"]["esc"]
    escudo = sup["escudo_total_mm"]
    k = {
        # --- precio y valor ---
        "precio": precio, "acciones_mm": acc, "dfn_mm": dfn, "nci_mm": nci, "ebitda_udm_mm": ebitda_udm, "pi": pi,
        "dcf": d["C49"].value, "mult_valor": p["B13"].value, "mult_x": p["B5"].value, "peso_dcf": p["B14"].value, "po": po_v,
        "po_vs": p["B17"].value, "dividendo_27": div27, "dividend_yield": div27 / precio, "ret_total": (po_v + div27) / precio - 1,
        "rango_52s": [semanas[1], semanas[2]], "capitalizacion_mm": precio * acc,
        "ev_dcf_mm": ev_dcf, "ev_mercado_mm": ev_mercado, "ev_po_mm": ev_po, "ev_mercado_vs_dcf": ev_mercado / ev_dcf - 1,
        "vp_exp_mm": d["C43"].value, "vp_vt_mm": d["C42"].value, "vt_mm": d["C41"].value, "vt_ev": d["C54"].value,
        "ps_exp": d["C43"].value / acc, "ps_vt": d["C42"].value / acc, "ps_deuda": -(d["C45"].value + d["C46"].value) / acc,
        "ev_ebitda_dcf": d["C55"].value, "ev_ebitda_po": ev_po / ebitda_udm, "ev_ebitda_mercado": ev_mercado / ebitda_udm,
        "cap_dcf": d["C56"].value, "cap_vt": d["C57"].value, "cap_mercado": p["B22"].value, "cap_pares": p["B23"].value,
        "cap_libro": 0.075, "cap_po": ebitda_udm / ev_po, "cap_po_real": real(ebitda_udm / ev_po),
        "dcf_vs_precio": d["C51"].value, "precio_vs_dcf": precio / d["C49"].value - 1,
        # --- tasas ---
        "wacc_2s": w["C43"].value, "wacc_t": w["I43"].value, "wacc_2s_real": w["C44"].value, "wacc_t_real": w["I44"].value,
        "ke_t": w["I39"].value, "ke_t_real": real(w["I39"].value), "kd_t_at": w["I42"].value, "kd_real": s["B33"].value,
        "beta_t": w["I37"].value, "beta_2s": w["C37"].value, "beta_u": w["F13"].value, "crp": w["C21"].value, "rf": s["B17"].value,
        "erp": s["B22"].value, "g": s["B88"].value, "g_real": s["B89"].value, "tasacion_real": w["C47"].value,
        "tasacion_nom": da["F"]["param"]["wacc_tas"], "wacc_impl": da["F"]["implicita"]["wacc"], "wacc_impl_pb": da["F"]["implicita"]["brecha"],
        "wacc_impl_real": real(da["F"]["implicita"]["wacc"]), "wacc_impl_po": da["F"]["implicita"]["wacc_po"],
        "ronic": ex["C38"].value, "yield_costo": ex["C37"].value, "mix_chile": s["B75"].value,
        "dfn_ebitda0": s["B41"].value, "dfn_ebitda_obj": s["B42"].value,
        # --- APV, otras pruebas ---
        "apv": apv["C21"].value, "apv_sin_deuda": apv["C22"].value, "apv_escudos": apv["C23"].value, "apv_ku": apv["C5"].value,
        "dcf_vf": d["C61"].value, "wacc_vf": s["B109"].value, "rf_pesos_dcf": da["F"]["rf"]["dcf"], "rf_pesos_po": da["F"]["rf"]["po"],
        "deuda_tabla": da["F"]["deuda"], "esc": esc, "efectos": da["F"]["efectos"],
        "esc_param": da["F"]["param"],
        # --- TIR ---
        "tir": tir, "tir_proyecto": tp,
        # --- escudo tributario y cascada ---
        "escudo_mm": escudo, "escudo_ps": sup["escudo_ps"], "supuestos_acum": sup["acum"], "supuestos_indiv": sup["indiv"],
        "supuestos_base": sup["base"],
        # --- pipeline y expansión ---
        "pipeline_m2": s["B70"].value, "pipeline_usd_mm": s["B69"].value, "pipeline_chile_m2": s["B73"].value, "pipeline_peru_m2": s["B74"].value,
        "otros_m2": fx["D45"].value, "costo_usd_m2": s["B71"].value, "gla_anual": s["B66"].value, "m2_maduro": d["C33"].value,
        "cohortes_m2": [(ex[f"A{r}"].value, ex[f"B{r}"].value) for r in range(12, 17)],
        "rampa": [s["B78"].value, s["B79"].value, s["B80"].value],
        # --- series por período (2S26E … 2031E; terminal al final cuando corresponde) ---
        "s": {
            "ing": serie(d, 7, COLS_T), "ing_base": serie(d, 5), "ing_exp": serie(d, 6), "ebitda": serie(d, 11, COLS_T), "ebitda_base": serie(d, 9),
            "ebitda_exp": serie(d, 10), "margen": serie(d, 12, COLS_T), "da": serie(d, 13, COLS_T), "ebit": serie(d, 14, COLS_T),
            "tasa": serie(d, 15, COLS_T), "imp": serie(d, 16, COLS_T), "nopat": serie(d, 17, COLS_T), "capex_base": serie(d, 19, COLS_T),
            "capex_exp": serie(d, 20, COLS_T), "ct": serie(d, 21, COLS_T), "fcff": serie(d, 22, COLS_T), "wacc": serie(d, 25, COLS_T),
            "fd": serie(d, 28), "vp": serie(d, 29), "crec": serie(d, 8),
            "ing_cal": serie(fb, 5), "mant": serie(fb, 15), "capex_proy": serie(fb, 14), "capex_usd": serie(fb, 12), "usdclp": serie(fb, 13),
            "costo_m2": serie(ex, 7), "ing_cl": [fx[f"{c}91"].value for c in "DEFGHI"], "ing_pe": [fx[f"{c}92"].value for c in "DEFGHI"],
            "ing_co": [fx[f"{c}93"].value for c in "DEFGHI"], "elim": [fx["D94"].value] + [0] * 5,
            "ebitda_cal": [fx[f"{c}104"].value for c in "DEFGHI"], "margen_cl": [fx[f"{c}101"].value for c in "DEFGHI"],
            "margen_pe": [fx[f"{c}102"].value for c in "DEFGHI"], "margen_co": [fx[f"{c}103"].value for c in "DEFGHI"],
            "gla_cl": [fx["D41"].value, fx["E41"].value, fx["F41"].value + fx["D45"].value, 0, 0, 0],
            "gla_pe": [fx["D42"].value, fx["E42"].value, fx["F42"].value, 0, 0, 0],
            "dfn_ap": serie(de, 7), "dfn_cierre": serie(de, 21), "dfn_ebitda": serie(de, 22), "ebitda_deuda": serie(de, 6),
            "cf_neto": serie(de, 11), "reajuste": serie(de, 13), "distribuible": serie(de, 15), "div_ord": serie(de, 16),
            "dist_adic": serie(de, 20), "deuda_bruta": serie(de, 23),
            "ffo_exhibits": [fx[f"{c}114"].value for c in "DEFGHI"], "reajuste_exhibits": [fx[f"{c}118"].value for c in "DEFGHI"],
            "cons_ing": [fx["D97"].value, fx["E97"].value],
        },
        "gla_exp_anual": s["B66"].value,
        "ing_1s26": s["B54"].value, "ebitda_1s26": s["B55"].value, "ing25": s["B56"].value, "k_ct": s["B84"].value * s["B85"].value,
        "ct_ing": s["B85"].value, "mant_pct": s["B58"].value, "da_pct": s["B57"].value, "tax_lp": s["B59"].value,
        "deudores": ct["E6"].value, "anticipos": ct["E12"].value,
        "sens": sensibilidades(ruta_fa, ruta_modelo),
        # football field (V1): [mínimo, máximo] de cada método, CLP por acción
        "football": {"dcf_50": [fig["V1_Football"]["B5"].value, fig["V1_Football"]["C5"].value],
                     "dcf_100": [fig["V1_Football"]["B6"].value, fig["V1_Football"]["C6"].value],
                     "mult_set_cl": [fig["V1_Football"]["B8"].value, fig["V1_Football"]["C8"].value],
                     "ntm_ebit": [fig["V1_Football"]["B9"].value, fig["V1_Football"]["C9"].value]},
        "precio_vs_tasadores": precio / esc["tas"]["dcf"] - 1,
        "precio_vs_dcf_pct": precio / d["C49"].value - 1,
        "mult_pe": 8644, "pb_precio": precio * acc / 4820898, "pb_po": po_v * acc / 4820898,
        "mc": json.loads((FIG_VAL / "datos_montecarlo.json").read_text()),   # simulación de Monte Carlo (valuacion/graficos/montecarlo.py)
        "consenso": {"po_12m": 4111, "bajo": 3377, "alto": 4600, "fuentes": 14, "valor": 2.85},
    }
    # FFO = utilidad distribuible + reajuste UF antes de impuestos (2026E: año completo de Exhibits; después, hoja Deuda)
    k["s"]["ffo"] = [fx["D114"].value] + [de[f"{c}15"].value + de[f"{c}13"].value for c in "DEFGH"]
    k["s"]["distribuible_cal"] = [fx["D119"].value] + [de[f"{c}15"].value for c in "DEFGH"]
    k["s"]["div_ord_cal"] = [s["B49"].value] + [de[f"{c}16"].value for c in "DEFGH"]
    k["s"]["dist_adic_cal"] = [s["B49"].value * 0 + de["C20"].value + 0] + [de[f"{c}20"].value for c in "DEFGH"]
    k["fa1"] = tabla_fa1(fa, k)
    salida = Path(__file__).parent / "cifras.json"
    salida.write_text(json.dumps(k, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
    print(salida, {x: round(k[x], 4) for x in ("dcf", "po", "po_vs", "wacc_impl", "ronic", "vt_ev", "ret_total")})


if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    main(*sys.argv[1:4])
