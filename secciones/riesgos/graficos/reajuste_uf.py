"""Mallplaza: sensibilidad del resultado al aumento de la UF (biblioteca).

Punto base: +3% de UF → −38.691 MM CLP (Nota 26, EEFF jun-26). +1% y +2% son
extrapolaciones lineales. La usa pagina_riesgos.py (Figura 26).
"""
import sys
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional y rutas compartidas
import estilo  # noqa: E402,F401  (registra las carpetas graficos/ de todas las secciones)

ROJO_CLARO = "#f8b3c9"  # tinte al 30% del rojo Mallplaza

IMPACTO_BASE = -38_691  # MM CLP para +3% de UF
VARIACION_BASE = 3
ESCENARIOS = [1, 2, 3]


def impacto(variacion_pct):
    return IMPACTO_BASE * variacion_pct / VARIACION_BASE
