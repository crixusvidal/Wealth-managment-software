"""Matriz de riesgo de la página de Riesgos (categoría, probabilidad 1-3 e impacto 1-3; categorías cualitativas).

La leen pagina_riesgos.py (figura de la matriz). Las probabilidades de los escenarios ya no salen de aquí: se estiman con la simulación
de Monte Carlo (valuacion/graficos/montecarlo.py).
"""
RIESGOS = {"OR1": ("op", 2, 2), "OR2": ("op", 1, 1), "FR1": ("fin", 3, 3), "MR1": ("mac", 3, 1),
           "MR2": ("mac", 1, 2), "MR3": ("mac", 2, 2), "RR1": ("reg", 1, 3), "RR2": ("reg", 2, 1)}
