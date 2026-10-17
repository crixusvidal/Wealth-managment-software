"""TIR (IRR) de la sección de Valuación: firma vs WACC y accionista vs Ke, siempre en nominal CLP y real (UF).

Lee las celdas ya calculadas de la hoja TIR_Accionista del modelo v3 (fórmulas vivas, recalculado con LibreOffice),
comprueba cada TIR con scipy (TIR independiente a partir de los flujos de las hojas DCF y Deuda) y vuelca
figuras/datos_tir.json para documento/valuacion.js.

Uso:
    python graficos/tir_valuacion.py <Mallplaza_Valoracion_Consolidada_v3.xlsx>
"""
import json
import sys
from pathlib import Path

import numpy as np
import openpyxl
from scipy.optimize import brentq

SALIDA = Path(__file__).resolve().parent / "figuras" / "datos_tir.json"
FILAS = {"firma": 7, "acc_mercado": 8, "acc_implicita": 9, "acc_objetivo": 10, "acc_dcf": 11}
ESCENARIOS = {"bear": 114, "base": 115, "bull": 116}


def verificar(wb):
    """TIR independientes (scipy) con los flujos del modelo: deben coincidir con la hoja TIR_Accionista."""
    d, de, su, w = wb["DCF"], wb["Deuda"], wb["Supuestos"], wb["WACC"]
    cols = "CDEFGH"
    fr = np.array([d[c + "26"].value for c in cols])
    cierre = np.cumsum(fr)
    medio = cierre - fr / 2
    fcff = np.array([d[c + "22"].value for c in cols])
    ft, g = d["I22"].value, su["B88"].value
    acc, precio = su["B9"].value, su["B8"].value
    ev = precio * acc + su["B10"].value + su["B11"].value
    r_firma = brentq(lambda r: (fcff / (1 + r) ** medio).sum() + ft / (r - g) / (1 + r) ** cierre[-1] - ev, g + 1e-3, 0.3)
    dist = np.array([de[c + "16"].value + de[c + "20"].value for c in cols])
    salida = d["C41"].value - de["H21"].value - su["B11"].value
    tir_acc = lambda p: brentq(lambda r: (dist / (1 + r) ** medio).sum() + salida / (1 + r) ** cierre[-1] - p * acc, -0.5, 0.5)
    tir_firma = lambda p: brentq(lambda r: (fcff / (1 + r) ** medio).sum() + ft / (r - g) / (1 + r) ** cierre[-1] - (p * acc + su["B10"].value + su["B11"].value), g + 1e-3, 0.3)
    return {"firma_objetivo": tir_firma(wb["Precio_Objetivo"]["B15"].value), "firma_dcf": tir_firma(d["C49"].value), "firma": r_firma, "acc_mercado": tir_acc(precio), "acc_objetivo": tir_acc(wb["Precio_Objetivo"]["B15"].value),
            "acc_dcf": tir_acc(d["C49"].value)}


def main(ruta):
    v = openpyxl.load_workbook(ruta, data_only=True)
    ws = v["TIR_Accionista"]
    out = {"pi": ws["C17"].value, "precio": ws["C18"].value, "objetivo": ws["C19"].value, "dcf": ws["C20"].value,
           "ke": ws["C28"].value, "ke_real": ws["D28"].value, "wacc": ws["C27"].value, "wacc_real": ws["D27"].value,
           "g": ws["C26"].value, "salida_acc_ps": ws["C49"].value, "salida_impl_ps": ws["C52"].value}
    for k, r in FILAS.items():
        out[k] = {"nom": ws[f"C{r}"].value, "real": ws[f"D{r}"].value, "costo_nom": ws[f"E{r}"].value,
                  "costo_real": ws[f"F{r}"].value, "brecha_nom": ws[f"G{r}"].value, "brecha_real": ws[f"H{r}"].value}
    out["esc"] = {k: {"nom": ws[f"C{r}"].value, "real": ws[f"D{r}"].value, "wacc": ws[f"E{r}"].value,
                      "wacc_real": ws[f"F{r}"].value, "brecha_nom": ws[f"G{r}"].value, "brecha_real": ws[f"H{r}"].value}
                  for k, r in ESCENARIOS.items()}
    chk = verificar(v)
    for k in ("firma_objetivo", "firma_dcf"):      # TIR del activo al precio objetivo y al valor DCF (no están en el Excel; mismo método que TIR_Implicita)
        nom = chk.pop(k)
        real = (1 + nom) / (1 + out["pi"]) - 1
        out[k] = {"nom": nom, "real": real, "costo_nom": out["wacc"], "costo_real": out["wacc_real"],
                  "brecha_nom": (nom - out["wacc"]) * 1e4, "brecha_real": (real - out["wacc_real"]) * 1e4}
    for k, x in chk.items():
        assert abs(x - out[k]["nom"]) < 1e-6, f"TIR {k}: hoja {out[k]['nom']} vs scipy {x}"
    # WACC implícita de la Figura 34 (V6): desplaza en paralelo toda la curva de WACC; la TIR de arriba es una tasa única
    v6 = json.loads((SALIDA.parent / "datos_anexos.json").read_text())["F"]["implicita"]
    out["v6"] = {"wacc": v6["wacc"], "brecha_pb": v6["brecha"]}
    out["control"] = {"tir_implicita_c4": v["TIR_Implicita"]["C4"].value}
    SALIDA.write_text(json.dumps(out, indent=1, ensure_ascii=False))
    print(SALIDA, {k: round(out[k]["nom"] * 100, 2) for k in FILAS})


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    main(sys.argv[1])
