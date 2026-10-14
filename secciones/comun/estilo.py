"""Paleta institucional de Mallplaza y utilidades de formato compartidas por todas las secciones.

Colores del informe (los únicos que usan los gráficos):
    ROJO #e9004b · CELESTE #64b0e1 · AMARILLO #ecb100 · NEGRO · AZUL #14284B (títulos)
    Parque Arauco #56d371 · Cencosud #3a0d7f (solo para identificar a esos pares)

Al importarse agrega al path las carpetas graficos/ de todas las secciones, para que un script
pueda reutilizar el motor de otra sección (p. ej. riesgos usa el DCF replicado de valuación).
"""
import sys
from pathlib import Path

from matplotlib.colors import to_rgb

SECCIONES = Path(__file__).resolve().parents[1]
for _carpeta in sorted(SECCIONES.glob("*/graficos")):
    if str(_carpeta) not in sys.path:
        sys.path.append(str(_carpeta))
LOGOS = Path(__file__).parent / "logos"

ROJO = "#e9004b"
CELESTE = "#64b0e1"
AMARILLO = "#ecb100"
NEGRO = "#000000"
AZUL = "#14284B"
VERDE_PARQUE_ARAUCO = "#56d371"
MORADO_CENCOSUD = "#3a0d7f"
TINTA_SECUNDARIA = "#5f5f5f"  # solo texto secundario (ejes, notas)
GRILLA = "#e6e6e6"            # solo líneas de grilla


def num(x, lang, dec=0):
    """Número con separadores del idioma (es: 1.234,5 · en: 1,234.5) y signo menos tipográfico."""
    s = f"{x:,.{dec}f}".replace("-", "−")
    return s.translate(str.maketrans(",.", ".,")) if lang == "es" else s


def pct(x, lang, dec=1):
    return num(x * 100, lang, dec) + "%"


def tinte(color, fuerza):
    """Mezcla el color con blanco (fuerza 1 = color puro). Usar solo para fondos de celdas."""
    return tuple(1 - fuerza * (1 - c) for c in to_rgb(color))
