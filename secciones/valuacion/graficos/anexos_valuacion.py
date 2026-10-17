"""Datos de los anexos [B], [C], [E] y [F] de Valuación (DCF, supuestos, múltiplos, escenarios).

Vuelca figuras/datos_anexos.json para documento/valuacion.js. Todas las cifras salen de los modelos
(openpyxl, data_only=True); los escenarios bear/bull/tasadores, que no están en el modelo, se calculan
con la réplica del DCF (escenarios_dcf.py, sobre riesgos/graficos/tornado_operacional.py; reproduce el valor del modelo exacto).

Uso:
    python graficos/anexos_valuacion.py <Valoracion_Consolidada_v4.xlsx> <Figuras_Valuacion_recalc.xlsx>
                                        <FA_Exhibits_v7.xlsx> <Tornado_Sensibilidad_Operacional.xlsx>
"""
import copy
import json
import sys
from pathlib import Path

import numpy as np
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "riesgos" / "graficos"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from tornado_operacional import leer_modelo, leer_rangos, proyectar  # noqa: E402

COLS = "CDEFGHI"           # 2S26E … 2031E, terminal (hojas DCF, Forecast_Base, Expansion)


def fila(ws, r, cols=COLS):
    return [ws[f"{c}{r}"].value for c in cols]


def anexo_b(wb):
    d, s = wb["DCF"], wb["Sensibilidad"]
    v = lambda c: d[c].value
    acc = v("C48")
    filas = {k: fila(d, r) for k, r in (("ing", 7), ("ebitda", 11), ("margen", 12), ("da", 13), ("ebit", 14),
                                         ("tasa", 15), ("imp", 16), ("nopat", 17), ("capex_base", 19),
                                         ("capex_exp", 20), ("ct", 21), ("fcff", 22), ("wacc", 25), ("fd", 28),
                                         ("vp", 29))}
    return {
        "filas": filas,
        "puente": {"vp_exp": v("C43"), "vp_vt": v("C42"), "ev": v("C44"), "dfn": v("C45"), "nci": v("C46"),
                   "patrimonio": v("C47"), "acciones": acc, "valor": v("C49"),
                   "ps_exp": v("C43") / acc, "ps_vt": v("C42") / acc, "ps_deuda": (v("C45") + v("C46")) / acc},
        "vt": {"wacc_t": v("C38"), "g": v("C39"), "reinv": v("C40"), "vt": v("C41"), "vt_ev": v("C54"),
               "ev_ebitda": v("C55"), "cap_ev": v("C56"), "cap_vt": v("C57"), "wacc_real": v("C58"),
               "m2_maduro": v("C33"), "ebitda_norm": v("C35")},
        "sens": {"g": [s[f"{c}5"].value for c in "BCDE"], "dw": [s[f"A{r}"].value for r in range(6, 11)],
                 "v": [[s[f"{c}{r}"].value for c in "BCDE"] for r in range(6, 11)]},
    }


def anexo_c(wb):
    f, e, d, su, ct = wb["Forecast_Base"], wb["Expansion"], wb["DCF"], wb["Supuestos"], wb["Capital_Trabajo"]
    s = lambda c: su[c].value
    cal = [f[f"{c}5"].value for c in "CDEFGH"]
    return {
        "filas": {"ing_cal": cal, "crec": fila(d, 8, "CDEFGH"), "ing_exp": fila(d, 6, "CDEFGH"),
                  "margen": fila(f, 9, "CDEFGH"), "tasa": fila(f, 11, "CDEFGH"), "capex_usd": fila(f, 12, "CDEFGH"),
                  "usdclp": fila(f, 13, "CDEFGH"), "capex_proy": fila(f, 14, "CDEFGH"), "mant": fila(f, 15, "CDEFGH"),
                  "capex_base": fila(f, 16, "CDEFGH"), "capex_exp": fila(e, 24, "CDEFGH"),
                  "costo_m2": fila(e, 7, "CDEFGH"), "ct": [-x for x in fila(d, 21, "CDEFGH")]},
        "sup": {"ing_1s26": s("B54"), "ebitda_1s26": s("B55"), "ing25": s("B56"), "da": s("B57"), "mant": s("B58"),
                "tax_lp": s("B59"), "gla_anual": s("B66"), "gla_22": s("B63"), "gla_23": s("B64"), "gla_25": s("B65"),
                "primer": s("B67"), "pago_t2": s("B68"), "pipe_usd": s("B69"), "pipe_m2": s("B70"),
                "costo_usd_m2": s("B71"), "mix_cl": s("B75"), "margen_cl": s("B76"), "margen_pe": s("B77"),
                "rampa": [s("B78"), s("B79"), s("B80")], "yield": e["C37"].value, "ronic": e["C38"].value,
                "ct_ing": s("B85"), "g": s("B88"), "dfn_obj": s("B42"), "anio_obj": s("B43"), "spread": s("B45"),
                "payout": s("B48"), "deudores": ct["E6"].value, "anticip": ct["E12"].value, "ctn": ct["E17"].value,
                "ebitda_ltm": s("B13"), "ing_ltm": wb["Comparables"]["J5"].value},
    }


def anexo_e(wb, wf):
    c, po, v7 = wb["Comparables"], wb["Precio_Objetivo"], wf["V7_Multiplos"]
    pares = []
    for r in range(5, 13):
        g = lambda col: c[f"{col}{r}"].value
        pares.append({"nombre": g("A"), "pais": g("C"), "set": g("D"), "chile": g("E"), "mcap": g("F"),
                      "ev": g("I"), "ebitda": g("K"), "m": [g(x) for x in "PQRST"]})
    est = {k: [c[f"{x}{r}"].value for x in "PQRST"] for k, r in (("med_set", 15), ("med_cl", 16), ("pond", 17))}
    metodos = [[po[f"B{r}"].value, po[f"C{r}"].value, po[f"D{r}"].value] for r in range(5, 10)]
    pond = [[v7[f"{x}{r}"].value for x in "ABCDE"] for r in range(26, 31)]
    s = wb["Supuestos"]
    return {"pares": pares, "est": est, "metodos": metodos, "pond": pond,
            "puente": {"mult": po["B5"].value, "ebitda": po["C5"].value, "ev": po["B5"].value * po["C5"].value,
                       "deuda_bbg": s["B14"].value, "nci": c["H5"].value, "acciones": s["B9"].value,
                       "valor": po["D5"].value},
            "peso_cl": s["B93"].value, "peso_dcf": s["B92"].value}


def anexo_f(wb, wf, ruta_fa, ruta_modelo, ruta_tor):
    import escenarios_dcf
    r = escenarios_dcf.calcular(ruta_fa, ruta_modelo, ruta_tor)
    p = r["param"]
    a, w = wb["APV"], wf["WACC_Construccion"]
    t2 = wb["Sensibilidad"]
    return {
        "esc": r["esc"], "ponderado": r["ponderado"], "prob": r["prob"], "efectos": r["efectos"],
        "param": {"dw_bear": p["dw"], "dw_bull": -p["dw"], "g_bull": p["g"], "m_lo": p["m_lo"], "m_hi": p["m_hi"],
                  "margen26": p["margen26"], "tax_bull": p["tax_bull"], "wacc_tas": p["wacc_tas"],
                  "wacc_tas_real": p["wacc_tas_real"], "wacc_t": p["wacc_t"], "g": p["g"], "share_var": p["share_var"],
                  "rangos": p["rangos"], "bajo": p["bajo"], "alto": p["alto"]},
        "apv": {"ku": a["C5"].value, "sin_deuda": a["C22"].value, "escudos": a["C23"].value, "valor": a["C21"].value,
                "dif": wb["DCF"]["C64"].value},
        "rf": {"bono": w["B74"].value, "dcf": w["B77"].value, "po": w["B78"].value, "efecto": w["B76"].value},
        "deuda": {"spread": [t2[f"{c}15"].value for c in "BCD"],   # fila 16 = estructura actual (Supuestos B41)
                  "filas": [[wb["Supuestos"]["B41"].value if r_ == 16 else
                             float(t2[f"A{r_}"].value.rstrip("x").replace(",", "."))]
                            + [t2[f"{c}{r_}"].value for c in "BCD"] for r_ in range(16, 20)]},
        "vf": {"dcf": wb["DCF"]["C61"].value, "wacc": wb["Supuestos"]["B109"].value},
        "implicita": {"wacc": wf["V6_DCFInverso"]["B17"].value, "brecha": wf["V6_DCFInverso"]["B18"].value,
                      "wacc_po": wf["V6_DCFInverso"]["B19"].value},
        "precio": r["precio"], "mult": r["mult"], "peso": r["peso"],
    }


if __name__ == "__main__":
    if len(sys.argv) < 5:
        sys.exit(__doc__)
    ruta_v2, ruta_fig, ruta_fa, ruta_tor = sys.argv[1:5]
    wb = openpyxl.load_workbook(ruta_v2, data_only=True)
    wf = openpyxl.load_workbook(ruta_fig, data_only=True)
    datos = {"B": anexo_b(wb), "C": anexo_c(wb), "E": anexo_e(wb, wf),
             "F": anexo_f(wb, wf, ruta_fa, ruta_v2, ruta_tor)}
    salida = Path(__file__).parent / "figuras" / "datos_anexos.json"
    salida.write_text(json.dumps(datos, ensure_ascii=False, indent=1, default=float))
    print(salida)
