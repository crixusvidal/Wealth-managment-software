"""Matriz de riesgo de la página de Riesgos (probabilidad × impacto, escala 1–3) y las probabilidades de los escenarios que se derivan de ella.

Única fuente: la leen pagina_riesgos.py (figura de la matriz) y valuacion/graficos/escenarios_dcf.py (pesos bear/base/bull).

Regla (supuestos del equipo, a la vista en el informe):
  1. Cada variable del bear y del bull es el EXTREMO histórico de una palanca del tornado (margen, ocupación, arriendo, ventas,
     crecimiento real, capex, visitas) o ±100 pb de WACC. Llegar a un extremo es, por definición, remoto: ocurrió una sola vez en
     las ~11 observaciones trimestrales del informe (≈ 9%).
  2. La probabilidad de que UNA variable llegue a su extremo en 12 meses depende de la probabilidad que la matriz de riesgo asigna
     al riesgo al que pertenece: baja (1) 2%, media (2) 5%, alta (3) 10%. Tope 10% ≈ frecuencia observada.
       Bear (extremo adverso): WACC +100 pb → FR1 (alta: 10%); las palancas operacionales → OR1 (media: 5%).
       Bull (extremo favorable): se invierte la categoría (alta ↔ baja, media ↔ media): WACC −100 pb → baja (2%); operacionales → media (5%).
  3. Las palancas de un mismo riesgo comparten ciclo y se mueven juntas (correlación perfecta: cuentan una vez); los riesgos distintos
     (tasas y operación) son independientes. Probabilidad conjunta = producto por riesgo:
       Bear = 10% × 5% = 0,5%;   Bull = 2% × 5% = 0,1%;   Base = el resto de la distribución (99,4%).
  4. La rebaja del impuesto a 23% (Ley de Reconstrucción, ya presentada) no entra en la probabilidad conjunta del bull.
"""
# código → (categoría, probabilidad 1-3, impacto 1-3), igual que el texto de la página de Riesgos
RIESGOS = {"OR1": ("op", 2, 2), "OR2": ("op", 1, 1), "FR1": ("fin", 3, 3), "MR1": ("mac", 3, 1),
           "MR2": ("mac", 1, 2), "MR3": ("mac", 2, 2), "RR1": ("reg", 1, 3), "RR2": ("reg", 2, 1)}
PROB_EXTREMO = {1: 0.02, 2: 0.05, 3: 0.10}       # baja / media / alta: probabilidad de que una variable llegue a su extremo en 12 meses
# variable → riesgo de la matriz al que pertenece
VARIABLES = {"wacc": "FR1", "margen": "OR1", "ocupacion": "OR1", "arriendo": "OR1", "ventas": "OR1", "crecimiento_real": "OR1",
             "capex": "OR1", "visitas": "OR1"}


def prob_variable(var, caso):
    """Probabilidad individual de que la variable llegue a su extremo adverso (bear) o favorable (bull)."""
    cat = RIESGOS[VARIABLES[var]][1]
    return PROB_EXTREMO[cat if caso == "bear" else 4 - cat]


def pesos_escenarios():
    out = {}
    for caso in ("bear", "bull"):
        p = 1.0
        for r in sorted(set(VARIABLES.values())):                       # una vez por riesgo (las palancas de un riesgo van juntas)
            var = next(v for v, rr in VARIABLES.items() if rr == r)
            p *= prob_variable(var, caso)
        out[caso] = p
    return {"bear": out["bear"], "base": round(1 - out["bear"] - out["bull"], 10), "bull": out["bull"]}


if __name__ == "__main__":
    print(pesos_escenarios(), {v: (prob_variable(v, "bear"), prob_variable(v, "bull")) for v in VARIABLES})
