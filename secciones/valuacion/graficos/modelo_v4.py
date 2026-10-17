"""Construye el modelo de valoración v4 con el pipeline completo con fecha (170.200 m²).

Hasta v3 el forecast incluía solo 123.700 m² del pipeline (Chile 67.900 m², Perú 55.800 m²) con USD 445 MM de capex,
mientras el costo por m² y el resto del informe (Industria, Anexos) usan el pipeline completo de la compañía:
170.200 m² y USD 613 MM. La diferencia son los «Otros proyectos Chile» (46.500 m², fechados en 2028 por la Memoria 2025,
p. 45; USD 134 MM). El motor del equipo ya trae el interruptor (hoja Forecast, celdas D43:D45).

Pasos:
  1. Exhibits v7 = Exhibits v6 con el interruptor «Otros proyectos Chile» = 1 y apertura en 2028 (recalculado con LibreOffice).
  2. Modelo v4 = modelo v3 con Forecast_Base (ingresos, EBITDA y capex USD 2026E–2031E) tomado del Exhibits v7 y el pipeline
     de Chile en Supuestos!B73 = 114.400 m² (67.900 + 46.500): el mix Chile/Perú de la expansión continua pasa de 54,9% a 67,2%.
  Todo lo demás (WACC, impuestos, capital de trabajo, TIR, escenarios) se recalcula por fórmula.

Uso (desde secciones/): python valuacion/graficos/modelo_v4.py
Entradas: analisis_financiero/modelos/Mallplaza_FA_Exhibits_v6.xlsx, valuacion/modelos/Mallplaza_Valoracion_Consolidada_v3.xlsx
Salidas: .../Mallplaza_FA_Exhibits_v7.xlsx y valuacion/modelos/Mallplaza_Valoracion_Consolidada_v4.xlsx
"""
import shutil
import subprocess
import sys
from pathlib import Path

import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill

RAIZ = Path(__file__).resolve().parents[2]
RECALC = "/root/.claude/skills/synced/508cc6bc-ad76-422f-ada9-ad0fc4279625_0b2a2ef9-bf53-46a5-b8ed-72255cce463a/xlsx/scripts/recalc.py"
V6 = RAIZ / "analisis_financiero/modelos/Mallplaza_FA_Exhibits_v6.xlsx"
V7 = RAIZ / "analisis_financiero/modelos/Mallplaza_FA_Exhibits_v7.xlsx"
M3 = RAIZ / "valuacion/modelos/Mallplaza_Valoracion_Consolidada_v3.xlsx"
M4 = RAIZ / "valuacion/modelos/Mallplaza_Valoracion_Consolidada_v4.xlsx"
M3_ESG = RAIZ / "valuacion/modelos/Mallplaza_Valoracion_Consolidada_v3_ESG.xlsx"      # v3 + hoja ESG_Ahorros
M4_ESG = RAIZ / "valuacion/modelos/Mallplaza_Valoracion_Consolidada_v4_ESG.xlsx"
TOR = RAIZ / "riesgos/modelos/Mallplaza_Tornado_Sensibilidad_Operacional_2026_2028.xlsx"
OTROS_M2 = 46500
AÑO_APERTURA = 2028
AMARILLO = PatternFill("solid", fgColor="FFFF00")


def recalcular(ruta):
    subprocess.run([sys.executable, RECALC, str(ruta)], check=True, stdout=subprocess.DEVNULL)


def exhibits_v7():
    shutil.copy(V6, V7)
    wb = openpyxl.load_workbook(V7)
    f = wb["Forecast"]
    assert f["D45"].value == OTROS_M2, "los «Otros proyectos Chile» deben ser 46.500 m²"
    f["D43"], f["D44"] = 1, AÑO_APERTURA
    for c in ("D43", "D44"):
        f[c].fill = AMARILLO
    f["J43"] = "v7: activado — pipeline completo con fecha (170.200 m²); Memoria 2025, p. 45: apertura en 2028"
    f["J44"] = "v7: 2028 (en v6 estaba 2029 y el interruptor en 0)"
    wb.save(V7)
    recalcular(V7)


def modelo_v4(origen=M3, destino=M4):
    fx = openpyxl.load_workbook(V7, data_only=True)["Forecast"]
    wb = openpyxl.load_workbook(origen)
    fb, su = wb["Forecast_Base"], wb["Supuestos"]
    for k, c in enumerate("CDEFGH"):            # 2026E–2031E: Forecast_Base C:H ↔ Exhibits D:I
        e = "DEFGHI"[k]
        fb[f"{c}5"] = fx[f"{e}95"].value        # ingresos ordinarios consolidados
        fb[f"{c}7"] = fx[f"{e}104"].value       # EBITDA (definición compañía)
        fb[f"{c}12"] = round(fx[f"{e}136"].value, 6)   # capex de proyectos (MMUSD)
        for r in (5, 7, 12):
            fb[f"{c}{r}"].fill = AMARILLO
    fb["A2"] = ("Cifras del forecast del equipo en azul (Exhibits v7), con el pipeline completo con fecha: 170.200 m² y USD 613 MM "
                "(incluye «Otros proyectos Chile», 46.500 m² que abren en 2028). Sin la expansión post-2028 (hoja Expansion).")
    fb["I5"] = "Exhibits v7, Forecast fila «Ingresos ordinarios consolidados» (con Otros proyectos Chile)."
    fb["I7"] = "Exhibits v7, Forecast bloque E (con Otros proyectos Chile)."
    su["B73"] = 67900 + OTROS_M2
    su["B73"].fill = AMARILLO
    su["A73"] = "m² del pipeline en Chile (67.900 + 46.500 de «Otros proyectos Chile», v4)"
    su["A1"] = "Supuestos — modelo consolidado de valoración Mallplaza (Plaza S.A.) · v4"
    ca = wb["Cambios"]
    ca["A1"] = "Mallplaza (Plaza S.A.) — Valorización consolidada v4"
    ca["A33"] = ("NUEVO · v4: pipeline completo con fecha (170.200 m², USD 613 MM): se activan los «Otros proyectos Chile» "
                 "(46.500 m², apertura 2028, USD 134 MM) en el forecast (Exhibits v7) y Supuestos!B73 = 114.400 m²; "
                 "cambian Forecast_Base (filas 5, 7 y 12), el mix de la expansión y, por fórmula, DCF, Deuda, TIR y Sensibilidad.")
    ca["A33"].fill = AMARILLO
    wb.save(destino)
    recalcular(destino)


def corrida(ruta, celdas, leer):
    """Copia el modelo, cambia celdas de Supuestos, recalcula con LibreOffice y devuelve leer(wb_valores)."""
    tmp = Path(ruta).with_name("_tmp_corrida.xlsx")
    wb = openpyxl.load_workbook(ruta)
    for c, v in celdas.items():
        wb["Supuestos"][c] = v
    wb.save(tmp)
    recalcular(tmp)
    try:
        return leer(openpyxl.load_workbook(tmp, data_only=True))
    finally:
        tmp.unlink(missing_ok=True)


def valores_estaticos(ruta):
    """Valores que el modelo guarda pegados (no son fórmulas): se recalculan cambiando Supuestos y la réplica del DCF."""
    from scipy.optimize import brentq
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import escenarios_dcf
    wb = openpyxl.load_workbook(ruta, data_only=True)
    su, d = wb["Supuestos"], wb["DCF"]
    dcf = lambda w: w["DCF"]["C49"].value
    init = su["B41"].value
    out = {"tabla2": {"actual": corrida(ruta, {"B42": init, "B45": 0.0}, dcf)}}
    for obj in (3.0, 3.5, 4.0):
        for sp in (0.0, 0.0025, 0.005):
            out["tabla2"][f"{obj}|{sp}"] = corrida(ruta, {"B42": obj, "B45": sp}, dcf)
    out["vf"] = corrida(ruta, {"B40": 0}, lambda w: {"dcf": w["DCF"]["C49"].value, "wacc": w["DCF"]["I25"].value})
    out["vf"]["po"] = su["B92"].value * out["vf"]["dcf"] + (1 - su["B92"].value) * wb["Precio_Objetivo"]["B13"].value
    # tasa única que iguala el DCF al precio (hoja TIR_Implicita, celda C4: era «Buscar objetivo»)
    fr = np.array([d[c + "26"].value for c in "CDEFGH"])
    cierre = np.cumsum(fr)
    medio = cierre - fr / 2
    fcff = np.array([d[c + "22"].value for c in "CDEFGH"])
    ft, g = d["I22"].value, su["B88"].value
    ev = su["B8"].value * su["B9"].value + su["B10"].value + su["B11"].value
    out["tir_implicita"] = brentq(lambda r: (fcff / (1 + r) ** medio).sum() + ft / (r - g) / (1 + r) ** cierre[-1] - ev, g + 1e-3, 0.3)
    out["esc"] = escenarios_dcf.calcular(str(V7), str(ruta), str(TOR))
    return out


def fijar_estaticos(ruta, v):
    wb = openpyxl.load_workbook(ruta)
    sens, su = wb["Sensibilidad"], wb["Supuestos"]
    t2 = v["tabla2"]
    for k, r in zip(("actual", 3.0, 3.5, 4.0), (16, 17, 18, 19)):
        for sp, c in zip((0.0, 0.0025, 0.005), "BCD"):
            sens[f"{c}{r}"] = round(t2["actual"] if k == "actual" else t2[f"{k}|{sp}"], 1)
    su["B108"], su["B109"], su["B110"] = v["vf"]["dcf"], v["vf"]["wacc"], v["vf"]["po"]
    ti = wb["TIR_Implicita"]
    ti["C4"] = round(v["tir_implicita"], 10)
    ti["A33"] = ("Chequeo: con C4 = WACC del modelo por período el valor da el valor DCF de la hoja DCF (CLP 1.859); "
                 "aquí se usa una tasa única para todos los períodos y el terminal.")
    ta, e = wb["TIR_Accionista"], v["esc"]
    fc = e["fcff"]
    for fila, k, shift, g in ((120, "bear", 100, e["param"]["g"]), (140, "bull", -100, e["param"]["g"])):
        for c, x in zip("CDEFGHI", fc[k]):
            ta[f"{c}{fila}"] = x
        ta[f"C{fila + 1}"], ta[f"C{fila + 2}"] = g, shift
    ta["I114"], ta["I115"], ta["I116"] = e["esc"]["bear"]["dcf"], e["esc"]["base"]["dcf"], e["esc"]["bull"]["dcf"]
    ta["A115"], ta["A129"] = "Base (modelo v4)", "Escenario · Base: modelo v4"
    ta["A111"] = ("Bear/bull: FCFF pegados desde la réplica del DCF (escenarios_dcf.py; azul; no son fórmulas, no hay hoja de escenarios "
                  "en el modelo). Base: vínculo vivo a DCF. Cada escenario se compara con su propia WACC terminal.")
    ta["A119"] = ("Escenario · Bear (todo sale mal): WACC +100 pb en todos los períodos y cada palanca operacional en su mínimo histórico "
                  "(margen 75,8%, ocupación, arriendo y ventas en mismas tiendas, visitas, capex de mantención 6,6%)")
    ta["A139"] = ("Escenario · Bull (todo sale bien): WACC −100 pb, cada palanca operacional en su máximo histórico (margen 81,3%, "
                  "ocupación, arriendo y ventas en mismas tiendas, visitas, capex de mantención 4,8%) e impuesto de 23%")
    if "ESG_Ahorros" in wb.sheetnames:
        wb["ESG_Ahorros"]["A61"] = "Mismo cuadro como % del precio objetivo (CLP 2.077)"
    wb.save(ruta)
    recalcular(ruta)


if __name__ == "__main__":
    exhibits_v7()
    modelo_v4(M3, M4)
    modelo_v4(M3_ESG, M4_ESG)
    v = valores_estaticos(M4)
    fijar_estaticos(M4, v)
    fijar_estaticos(M4_ESG, v)
    wb = openpyxl.load_workbook(M4, data_only=True)
    d, po = wb["DCF"], wb["Precio_Objetivo"]
    print("DCF", round(d["C49"].value, 2), "PO", round(po["B15"].value, 2), "upside", round(po["B17"].value, 4),
          "VT/EV", round(d["C54"].value, 4), "RONIC", round(wb["Expansion"]["C38"].value, 4))
    print("tabla2", {k: round(x, 1) for k, x in v["tabla2"].items()}, "VF", v["vf"], "TIR implícita", v["tir_implicita"])
