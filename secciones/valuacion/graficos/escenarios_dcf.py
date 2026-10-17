"""Escenarios bear / base / bull del precio objetivo, definidos con las mismas palancas del tornado de Riesgos (Figura 24).

  Bear = todo sale mal: cada palanca operacional en su mínimo histórico a la vez (margen EBITDA, ocupación, arriendo en mismas
         tiendas = precio, ventas de los locatarios = renta variable, crecimiento real de largo plazo en el piso de la tasación,
         visitas, capex de mantención en el máximo) y WACC +100 pb en todos los períodos.
  Base = estimación central del modelo (DCF 80% + múltiplos 20%).
  Bull = todo sale bien: cada palanca en su máximo histórico a la vez, WACC −100 pb en todos los períodos e impuesto de 23%
         (Ley de Reconstrucción) desde que baja de 27%.

El precio (arriendo por m², SSR) y las ventas de los locatarios (SSS) se mueven juntos: la renta variable (5,7% de los
ingresos) transmite las ventas al ingreso del operador. g nominal se mantiene en 3,0% en todos los casos (0% real: con RONIC
bajo la WACC real, crecer sobre la inflación resta valor). Se calculan con la réplica del DCF (riesgos/graficos/tornado_operacional.py).

Uso como módulo:
    from escenarios_dcf import calcular
    r = calcular(ruta_exhibits, ruta_modelo, ruta_tornado)
"""
import copy
import sys
from pathlib import Path

import numpy as np
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "riesgos" / "graficos"))
from tornado_operacional import escenarios as palancas_tornado  # noqa: E402
from tornado_operacional import leer_modelo, leer_rangos, proyectar  # noqa: E402
from matriz_datos import pesos_escenarios  # noqa: E402

IMPUESTO_BULL = 0.23           # Ley de Reconstrucción (9-oct-26): 27% → 23%; no está en el modelo
DW = 0.01                      # ±100 pb de WACC en todos los períodos
PESO_PROB = pesos_escenarios()         # 30 / 60 / 10: se derivan de la matriz de riesgo (FR1 alta, OR1 media; regla en matriz_datos.py)


def combinar(lista):
    """Suma d_ss_corto (precio + ventas) y toma el valor de cada palanca restante."""
    out = {}
    for kw in lista:
        for k, v in kw.items():
            out[k] = out.get(k, 0.0) + v if k == "d_ss_corto" else v
    return out


def calcular(ruta_fa, ruta_modelo, ruta_tor):
    m, control = leer_modelo(ruta_fa, ruta_modelo)
    base = proyectar(m)[2]
    assert abs(base - control["valor"]) < 1e-6, "la réplica no reproduce el DCF del modelo"
    wb = openpyxl.load_workbook(ruta_modelo, data_only=True)
    po = wb["Precio_Objetivo"]
    mult, peso, precio = po["B13"].value, po["B14"].value, po["B16"].value
    objetivo = lambda x: peso * x + (1 - peso) * mult
    rangos = leer_rangos(ruta_tor)
    pal = palancas_tornado(m, rangos)
    bajo, alto = combinar([p[2] for p in pal]), combinar([p[3] for p in pal])
    bull_m = copy.deepcopy(m)
    bull_m["tax"], bull_m["tax_lp"] = np.minimum(m["tax"], IMPUESTO_BULL), IMPUESTO_BULL
    wacc_tas = openpyxl.load_workbook(ruta_modelo.replace("v4", "v4"), data_only=True)["WACC"]["C47"].value
    wacc_tas_nom = (1 + wacc_tas) * (1 + m["inflacion"]) - 1
    corridas = {
        "bear": proyectar(m, d_wacc=DW, **bajo),
        "base": proyectar(m),
        "bull": proyectar(bull_m, d_wacc=-DW, **alto),
        "tas": proyectar(m, d_wacc=wacc_tas_nom - m["wacc_t"]),
    }
    esc = {k: {"dcf": c[2], "po": objetivo(c[2]), "vs": objetivo(c[2]) / precio - 1, "fcff": [float(x) for x in c[1]]}
           for k, c in corridas.items()}
    # efecto individual de cada palanca (mismo orden que el tornado), para la tabla de supuestos del Anexo 18
    efectos = [{"palanca": p[0], "rango": p[1], "bajo": proyectar(m, **p[2])[2] - base, "alto": proyectar(m, **p[3])[2] - base}
               for p in pal]
    efectos += [{"palanca": "WACC ±100 pb", "rango": "±100 pb", "bajo": proyectar(m, d_wacc=DW)[2] - base,
                 "alto": proyectar(m, d_wacc=-DW)[2] - base},
                {"palanca": "Impuesto 23%", "rango": "27% → 23%", "bajo": 0.0, "alto": proyectar(bull_m)[2] - base}]
    prob = sum(PESO_PROB[k] * esc[k]["po"] for k in PESO_PROB)
    ponderado = {"po": prob, "dcf": sum(PESO_PROB[k] * esc[k]["dcf"] for k in PESO_PROB), "vs": prob / precio - 1}
    margen26 = wb["Forecast_Base"]["C7"].value / wb["Forecast_Base"]["C5"].value
    _, m_lo, m_hi = rangos["Margen EBITDA"]
    return {
        "esc": {k: {a: b for a, b in v.items() if a != "fcff"} for k, v in esc.items()},
        "fcff": {k: v["fcff"] for k, v in esc.items()},
        "ponderado": ponderado, "prob": PESO_PROB, "efectos": efectos,
        "param": {"dw": DW, "m_lo": m_lo, "m_hi": m_hi, "margen26": margen26, "tax_bull": IMPUESTO_BULL, "wacc_tas": wacc_tas_nom,
                  "wacc_tas_real": wacc_tas, "wacc_t": m["wacc_t"], "g": m["g"], "share_var": m["share_var"],
                  "rangos": {k: list(v) for k, v in rangos.items()}, "bajo": {k: v for k, v in bajo.items() if k != "ss_largo"},
                  "alto": {k: v for k, v in alto.items() if k != "ss_largo"}},
        "precio": precio, "mult": mult, "peso": peso,
    }


if __name__ == "__main__":
    raiz = Path(__file__).resolve().parents[2]
    r = calcular(str(raiz / "analisis_financiero/modelos/Mallplaza_FA_Exhibits_v7.xlsx"),
                 str(raiz / "valuacion/modelos/Mallplaza_Valoracion_Consolidada_v4.xlsx"),
                 str(raiz / "riesgos/modelos/Mallplaza_Tornado_Sensibilidad_Operacional_2026_2028.xlsx"))
    for k, v in r["esc"].items():
        print(k, {a: round(b, 4 if a == "vs" else 1) for a, b in v.items()})
    print("ponderado", {a: round(b, 3) for a, b in r["ponderado"].items()})
    for e in r["efectos"]:
        print(e["palanca"], e["rango"], round(e["bajo"], 1), round(e["alto"], 1))
