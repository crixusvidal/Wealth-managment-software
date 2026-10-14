"""Réplica del DCF del modelo v2 para la sensibilidad WACC vs g (biblioteca).

Calcula el precio objetivo (80% DCF + 20% múltiplos) en una grilla 5 x 5 con el caso base al
centro (WACC por período con desplazamiento paralelo de ±1 pp vs g de 2,0% a 4,0%). Replica la
hoja Sensibilidad del Excel; la columna g = 2,0% no existe en el Excel y se calcula con la misma
fórmula. La usa pagina_riesgos.py (Figura 25) y plantilla_informe/graficos/figuras_informe.py.
"""
import sys
from pathlib import Path

import numpy as np
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional y rutas compartidas
import estilo  # noqa: E402,F401  (registra las carpetas graficos/ de todas las secciones)
from estilo import NEGRO

NEUTRO = "#f2f2f2"


PASOS_G = [-0.01, -0.005, 0.0, 0.005, 0.01]   # g alrededor del caso base (5 columnas)
PASOS_W = [-0.01, -0.005, 0.0, 0.005, 0.01]   # desplazamiento paralelo de la WACC (5 filas)


def leer_insumos(ruta):
    """Modelo v2: cada período se descuenta con su propia WACC (hoja DCF, fila 25)."""
    wb = openpyxl.load_workbook(ruta, data_only=True)
    d, s, ex = wb["DCF"], wb["Supuestos"], wb["Expansion"]
    v = lambda c: s[c].value
    return {"fcff": np.array([d[f"{c}22"].value for c in "CDEFGH"]),
            "waccs": np.array([d[f"{c}25"].value for c in "CDEFGH"]), "wacc_t": d["I25"].value,
            "frac": np.array([d[f"{c}26"].value for c in "CDEFGH"]),
            "ing_norm": d["C34"].value, "ebitda_norm": d["C35"].value,
            "tax": v("B59"), "da": v("B57"), "mant": v("B58"), "k_ct": v("B84") * v("B85"),
            "inflacion": v("B34"), "ronic": ex["C38"].value,
            "ajustes": v("B10") + v("B11"), "acciones": v("B9")}


def dcf_por_accion(d_wacc, g, x):
    """Valor DCF por acción con la WACC de todos los períodos desplazada en d_wacc (hoja Sensibilidad)."""
    w = x["waccs"] + d_wacc
    cierre = np.cumprod((1 + w) ** -x["frac"])
    mitad = np.concatenate([[1.0], cierre[:-1]]) * (1 + w) ** -(x["frac"] / 2)
    ing, ebitda = x["ing_norm"] * (1 + g), x["ebitda_norm"] * (1 + g)
    ebit = ebitda - x["da"] * ing
    nopat = ebit * (1 - x["tax"])
    g_real = (1 + g) / (1 + x["inflacion"]) - 1
    fcff_t = (nopat + x["da"] * ing - x["mant"] * ing - x["k_ct"] * (ing - x["ing_norm"])
              - g_real / x["ronic"] * nopat)
    vt = fcff_t / (x["wacc_t"] + d_wacc - g) * cierre[-1]
    return (np.sum(x["fcff"] * mitad) + vt - x["ajustes"]) / x["acciones"]


def grilla(ruta):
    """Precio objetivo (peso DCF × DCF + resto × múltiplos) por WACC (filas, mayor a menor) y g (columnas).

    El DCF de cada celda se recalcula con la lógica de la hoja Sensibilidad y se verifica contra ella;
    el valor por múltiplos queda fijo (no depende de la WACC ni de g).
    """
    wb = openpyxl.load_workbook(ruta, data_only=True)
    x = leer_insumos(ruta)
    po, sens, s = wb["Precio_Objetivo"], wb["Sensibilidad"], wb["Supuestos"]
    g_base, peso, multiplos = s["B88"].value, po["B14"].value, po["B13"].value
    for r in range(6, 11):      # verificación contra la tabla 1 del Excel
        for c in range(2, 6):
            esperado = sens.cell(r, c).value
            calculado = dcf_por_accion(sens.cell(r, 1).value, sens.cell(5, c).value, x)
            assert abs(calculado - esperado) < 1e-6, (r, c, calculado, esperado)
    gs = [round(g_base + p, 6) for p in PASOS_G]
    deltas = sorted(PASOS_W, reverse=True)       # +1 pp arriba, −1 pp abajo
    waccs = [x["wacc_t"] + dw for dw in deltas]
    z = np.array([[peso * dcf_por_accion(dw, g, x) + (1 - peso) * multiplos for g in gs] for dw in deltas])
    i_base, j_base = deltas.index(0.0), gs.index(round(g_base, 6))
    assert abs(z[i_base, j_base] - po["B15"].value) < 1e-6, "el centro debe ser el precio objetivo"
    return waccs, gs, z, i_base, j_base, po["B16"].value


def color_texto(rgba):
    r, g, b = rgba[:3]
    return "white" if 0.299 * r + 0.587 * g + 0.114 * b < 0.55 else NEGRO
