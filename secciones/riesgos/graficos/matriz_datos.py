"""Matriz de riesgo de la página de Riesgos (probabilidad × impacto, escala 1–3) y los pesos de los escenarios que se derivan de ella.

Única fuente: la leen pagina_riesgos.py (figura de la matriz) y valuacion/graficos/escenarios_dcf.py (pesos bear/base/bull).

Regla de los pesos (todos los supuestos son del equipo y están a la vista):
  1. Probabilidad cualitativa → probabilidad numérica a 12 meses: baja (1) 10%, media (2) 30%, alta (3) 50%.
  2. Cada escenario es la ocurrencia conjunta de los riesgos de la matriz que mueven sus palancas, con correlación perfecta
     («todo sale mal» / «todo sale bien» ocurren juntos); la probabilidad conjunta es entonces la del riesgo menos probable.
       Bear: FR1 (tasas: WACC +100 pb) y OR1 (deterioro operacional: palancas del tornado en su mínimo).
       Bull: los contrarios de FR1 y OR1. Un riesgo de probabilidad alta de empeorar tiene baja de mejorar (3 → 1), uno medio
             tiene media (2 → 2) y uno bajo, alta (1 → 3).
  3. Base = 1 − bear − bull.
Con FR1 alta y OR1 media: bear = mín(50%, 30%) = 30%; bull = mín(10%, 30%) = 10%; base = 60%.
"""
# código → (categoría, probabilidad 1-3, impacto 1-3), igual que el texto de la página de Riesgos
RIESGOS = {"OR1": ("op", 2, 2), "OR2": ("op", 1, 1), "FR1": ("fin", 3, 3), "MR1": ("mac", 3, 1),
           "MR2": ("mac", 1, 2), "MR3": ("mac", 2, 2), "RR1": ("reg", 1, 3), "RR2": ("reg", 2, 1)}
PROB_NUM = {1: 0.10, 2: 0.30, 3: 0.50}            # baja / media / alta, a 12 meses (supuesto del equipo)
RIESGOS_BEAR = ("FR1", "OR1")                    # tasas y deterioro operacional: las palancas del bear y del bull


def pesos_escenarios():
    pb = [PROB_NUM[RIESGOS[c][1]] for c in RIESGOS_BEAR]
    pf = [PROB_NUM[4 - RIESGOS[c][1]] for c in RIESGOS_BEAR]
    bear, bull = min(pb), min(pf)
    return {"bear": bear, "base": round(1 - bear - bull, 10), "bull": bull}


if __name__ == "__main__":
    print(pesos_escenarios())
