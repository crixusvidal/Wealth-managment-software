"""Actualiza los datos de las figuras V1–V6 (Mallplaza_Figuras_Valuacion.xlsx) con el modelo de valoración vigente (v4).

El libro de figuras guarda pegadas las cifras que dependen del DCF (football field, grilla de g, capital de trabajo, cap rate del DCF,
DCF inverso). Este script las recalcula desde el modelo v4 (hojas Sensibilidad, APV, DCF, Expansion, WACC) y desde la réplica del DCF
(riesgos/graficos/tornado_operacional.py, que reproduce el valor del modelo) y deja el libro recalculado con LibreOffice
(Mallplaza_Figuras_Valuacion_recalc.xlsx), que es el que lee valuacion.py.

Uso (desde secciones/):
    python valuacion/graficos/actualizar_figuras.py <Mallplaza_Valoracion_Consolidada_v4.xlsx> <FA_Exhibits_v7.xlsx>
"""
import shutil
import subprocess
import sys
from pathlib import Path

import numpy as np
import openpyxl
from scipy.optimize import brentq

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "riesgos" / "graficos"))
from tornado_operacional import leer_modelo, proyectar  # noqa: E402

RAIZ = Path(__file__).resolve().parents[2]
LIBRO = RAIZ / "valuacion/modelos/Mallplaza_Figuras_Valuacion.xlsx"
LIBRO_RECALC = RAIZ / "valuacion/modelos/Mallplaza_Figuras_Valuacion_recalc.xlsx"
RECALC = "/root/.claude/skills/synced/508cc6bc-ad76-422f-ada9-ad0fc4279625_0b2a2ef9-bf53-46a5-b8ed-72255cce463a/xlsx/scripts/recalc.py"
MULT_CT = {15: 0.0, 16: 0.5, 17: 1.0, 18: 1.5, 19: 2.0}          # filas de V4: multiplicador del capital de trabajo
DW_V6 = [-0.04, -0.035, -0.03, -0.025, -0.02, -0.015, -0.01, -0.005, 0.0, 0.005, 0.01]   # filas 5–15 de V6


def main(ruta_modelo, ruta_fa):
    m, control = leer_modelo(ruta_fa, ruta_modelo)
    base = proyectar(m)[2]
    assert abs(base - control["valor"]) < 1e-6
    wm = openpyxl.load_workbook(ruta_modelo, data_only=True)
    sens, d, po, w, ex, apv = wm["Sensibilidad"], wm["DCF"], wm["Precio_Objetivo"], wm["WACC"], wm["Expansion"], wm["APV"]
    precio, mult, peso = po["B16"].value, po["B13"].value, po["B14"].value
    wb = openpyxl.load_workbook(LIBRO)

    # V1: football field (WACC ±50 pb con g 3,0%; ±100 pb con g 2,5–4,0%; APV)
    v1 = wb["V1_Football"]
    col_g3 = [sens[f"C{r}"].value for r in range(6, 11)]          # g 3,0%, ΔWACC −100…+100 pb
    todos = [sens[f"{c}{r}"].value for c in "BCDE" for r in range(6, 11)]
    v1["B5"], v1["C5"], v1["D5"] = col_g3[3], col_g3[1], base          # +50 pb (mínimo) y −50 pb (máximo)
    v1["B6"], v1["C6"], v1["D6"] = min(todos), max(todos), base
    apv_v = apv["C21"].value
    v1["B7"], v1["C7"], v1["D7"] = apv_v, apv_v, apv_v

    # V2: WACC implícita (se completa abajo), nuestra WACC y Ke terminales
    v2 = wb["V2_WACC_Tasas"]
    v2["B13"], v2["B14"] = w["I43"].value, w["I39"].value

    # V3: g nominal 2,5–4,0% (valor y FCFF 2032E de la hoja Sensibilidad)
    v3 = wb["V3_ValueDriver"]
    v3["B4"], v3["B7"] = ex["C37"].value, w["I44"].value
    for r, c in zip(range(14, 18), "BCDE"):
        v3[f"D{r}"], v3[f"E{r}"] = sens[f"{c}25"].value, sens[f"{c}8"].value

    # V4: capital de trabajo escalado (réplica del DCF)
    v4 = wb["V4_CapitalTrabajo"]
    for r, k in MULT_CT.items():
        mm = dict(m)
        mm["k_ct"] = m["k_ct"] * k
        v4[f"D{r}"] = round(proyectar(mm)[2], 1)
        v4[f"E{r}"] = f"={peso}*D{r}+{1 - peso}*{mult}"
    ct_total = base - proyectar({**m, "k_ct": 0.0})[2]
    v4["A22"] = (f"Lectura: todo el capital de trabajo vale ~CLP {abs(ct_total):.0f} por acción ({abs(ct_total) / base * 100:.1f}% del DCF). "
                 "El valor no depende de este supuesto.")

    # V5: cap rates (mercado, pares, DCF y valor terminal)
    v5 = wb["V5_CapRate"]
    v5["B5"], v5["B6"] = po["B22"].value, po["B23"].value
    v5["B9"], v5["B10"] = d["C56"].value, d["C57"].value

    # V6: DCF inverso (réplica) y WACC implícita
    v6 = wb["V6_DCFInverso"]
    for r, dw in zip(range(5, 16), DW_V6):
        v6[f"C{r}"] = round(proyectar(m, d_wacc=dw)[2], 1)
    wacc_t = m["wacc_t"]
    dw_dcf = brentq(lambda x: proyectar(m, d_wacc=x)[2] - precio, -0.06, 0.0)
    dw_po = brentq(lambda x: peso * proyectar(m, d_wacc=x)[2] + (1 - peso) * mult - precio, -0.06, 0.0)
    v6["B17"], v6["B19"] = wacc_t + dw_dcf, wacc_t + dw_po
    v2["B11"] = wacc_t + dw_dcf
    pi = m["inflacion"]
    real = (1 + wacc_t + dw_dcf) / (1 + pi) - 1
    v6["A22"] = (f"Mensaje: para justificar CLP {precio:,.0f} el mercado debe descontar a ~{(wacc_t + dw_dcf) * 100:.2f}% nominal "
                 f"(~{real * 100:.1f}% real), {abs(dw_dcf) * 1e4:.0f} pb bajo nuestra WACC terminal.").replace(",", ".")
    v2["A20"] = (f"Lectura: a CLP {precio:,.0f} la WACC real implícita (~{real * 100:.1f}%) supera el bono UF a 10 años del BCCh "
                 "(2,5% real) y queda bajo la WACC de la tasación (5,55% real).").replace(",", ".")

    # V7: sensibilidad a la ponderación Chile/pares del múltiplo (el DCF va pegado en la fórmula del precio objetivo)
    v7 = wb["V7_Multiplos"]
    for r in range(26, 31):
        v7[f"D{r}"] = f"={peso}*{base}+{1 - peso}*C{r}"
    wb.save(LIBRO)
    shutil.copy(LIBRO, LIBRO_RECALC)
    subprocess.run([sys.executable, RECALC, str(LIBRO_RECALC)], check=True, stdout=subprocess.DEVNULL)
    # F. Rf del bono en pesos: efecto en la WACC (pb, fórmula del libro) → DCF por la réplica con ese desplazamiento
    efecto_pb = openpyxl.load_workbook(LIBRO_RECALC, data_only=True)["WACC_Construccion"]["B76"].value
    dcf_rf = proyectar(m, d_wacc=efecto_pb / 1e4)[2]
    wc = wb["WACC_Construccion"]
    wc["B77"] = round(dcf_rf)
    wc["C77"] = f"Réplica del DCF v4 con WACC {efecto_pb:.0f} pb en todos los períodos"
    wc["A81"] = (f"Es la elección más conservadora: con el bono en pesos el precio objetivo sube a ~CLP {peso * round(dcf_rf) + (1 - peso) * mult:,.0f}, "
                 f"todavía {abs((peso * round(dcf_rf) + (1 - peso) * mult) / precio - 1) * 100:.0f}% bajo el precio.").replace(",", ".")
    wb.save(LIBRO)
    shutil.copy(LIBRO, LIBRO_RECALC)
    subprocess.run([sys.executable, RECALC, str(LIBRO_RECALC)], check=True, stdout=subprocess.DEVNULL)
    print({"dcf": round(base, 2), "implicita_dcf": round(wacc_t + dw_dcf, 5), "implicita_po": round(wacc_t + dw_po, 5),
           "brecha_pb": round(dw_dcf * 1e4, 1), "ct_total": round(ct_total, 1)})


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    main(*sys.argv[1:3])
