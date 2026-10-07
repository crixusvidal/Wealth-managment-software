"""Agrega la hoja Valor_Implicito al libro de comparables (Bloomberg) de Mallplaza.

Aplica los múltiplos de los peers (P25 / mediana / P75) a las métricas de Mallplaza, con
fórmulas vinculadas a las hojas existentes, y combina el resultado con el DCF (80/20).
Las hojas originales no se modifican.

Uso:
    python valoracion/valoracion_multiplos.py <Comparables_Mallplaza_CFA_BBG.xlsx> <salida.xlsx>
"""
import sys

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

NAVY, SALMON, MEDIAN_FILL, YELLOW = "1F3864", "FCE4D6", "D9E1F2", "FFFF00"
BLUE, GREEN, BLACK = "0000FF", "008000", "000000"
NUM = r"#,##0;\(#,##0\);\-"
MULT = r'0.0\x;\(0.0"x)";\-'
PCT = r"0.0%;\(0.0%\);\-"
thin = Side(style="thin", color="BFBFBF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)

DCF_POR_ACCION = 2104.837  # Mallplaza_FA_Exhibits_v6.xlsx, DCF_Valuation!J49 (WACC 10,03%, g 3,0%)

# (método, base, grupo, hoja, columna del múltiplo, filas de peers, celda de la métrica, tipo)
METODOS = [
    ("EV/EBITDA", "LTM", "Core", "Valoracion_LTM", "O", "8:12", "C5", "EV"),
    ("EV/EBITDA", "NTM", "Core", "Valoracion_NTM", "O", "8:12", "C6", "EV"),
    ("EV/EBIT", "LTM", "Core", "Valoracion_LTM", "P", "8:12", "C7", "EV"),
    ("EV/EBITDA", "LTM", "Chile (Parque Arauco, Cencosud Shopping)", "Valoracion_LTM", "O", "8:9", "C5", "EV"),
    ("EV/EBITDA", "NTM", "Chile (Parque Arauco, Cencosud Shopping)", "Valoracion_NTM", "O", "8:9", "C6", "EV"),
    ("P/E", "NTM", "Core", "Valoracion_NTM", "N", "8:12", "C8", "Equity"),
    ("P/BV", "30-jun-2026", "Core", "Valoracion_LTM", "R", "8:12", "C9", "Equity"),
    ("P/BV", "30-jun-2026", "Chile (Parque Arauco, Cencosud Shopping)", "Valoracion_LTM", "R", "8:9", "C9", "Equity"),
]


def estilo(c, bold=False, color=BLACK, fill=None, fmt=None, size=10, wrap=False, border=True):
    c.font = Font(name="Arial", size=size, bold=bold, color=color)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        c.number_format = fmt
    if border:
        c.border = BOX
    c.alignment = Alignment(vertical="center", wrap_text=wrap)


def encabezado(ws, fila, textos):
    for col, t in enumerate(textos, start=1):
        c = ws.cell(fila, col, t)
        estilo(c, bold=True, color="FFFFFF", fill=NAVY, wrap=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def construir(ruta_in, ruta_out):
    wb = openpyxl.load_workbook(ruta_in)
    ws = wb.create_sheet("Valor_Implicito", index=0)
    ws.sheet_view.showGridLines = False

    estilo(ws["A1"], bold=True, size=13, border=False)
    ws["A1"] = "Valor implícito de Mallplaza por múltiplos comparables (precio al 30-sep-2026)"
    estilo(ws["A2"], size=8, border=False)
    ws["A2"] = ("Aplica los múltiplos de los peers (hojas Valoracion_LTM y Valoracion_NTM) a las métricas de "
                "Mallplaza. Montos en MM CLP; valor por acción en CLP. Mallplaza excluido de los estadísticos.")

    # 1. Datos de Mallplaza
    estilo(ws["A4"], bold=True, size=11, border=False)
    ws["A4"] = "1. Datos de Mallplaza"
    datos = [
        ("EBITDA LTM", "=Valoracion_LTM!J7", NUM, GREEN, "Bloomberg, LTM 3T25–2T26"),
        ("EBITDA NTM (consenso)", "=Valoracion_NTM!J7", NUM, GREEN, "Bloomberg BEst, 13 analistas"),
        ("EBIT LTM", "=Valoracion_LTM!K7", NUM, GREEN, "EBITDA − D&A"),
        ("Utilidad neta NTM (consenso)", "=Valoracion_NTM!L7", NUM, GREEN, "Bloomberg BEst; puede incluir revalorización"),
        ("Patrimonio controlador (book value)", "=Valoracion_LTM!M7", NUM, GREEN, "30-jun-2026"),
        ("Deuda neta", "=Valoracion_LTM!F7", NUM, GREEN, "Bloomberg NET_DEBT, 30-jun-2026 (misma definición que los peers)"),
        ("Minoritarios", "=Valoracion_LTM!G7", NUM, GREEN, "30-jun-2026"),
        ("Acciones (MM)", "=Datos_BBG!O8/Datos_BBG!P8", "#,##0.0", BLACK, "Market cap / precio"),
        ("Precio al 30-sep-2026 (CLP)", "=Datos_BBG!P8", NUM, GREEN, "Bloomberg PX_LAST"),
        ("Valor por acción DCF (CLP)", DCF_POR_ACCION, NUM, BLUE, "Modelo DCF del equipo"),
        ("Peso DCF en el precio objetivo", 0.8, PCT, BLUE, "Supuesto del equipo"),
        ("Peso múltiplos en el precio objetivo", "=1-C15", PCT, BLACK, ""),
    ]
    for i, (label, val, fmt, color, nota) in enumerate(datos):
        r = 5 + i
        a = ws.cell(r, 1, label)
        estilo(a)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        c = ws.cell(r, 3, val)
        estilo(c, color=color, fmt=fmt, fill=YELLOW if r in (14, 15) else None)
        n = ws.cell(r, 4, nota)
        estilo(n, size=8, border=False)
    ws["C14"].comment = Comment(
        "Mallplaza_FA_Exhibits_v6.xlsx, hoja DCF_Valuation, celda J49: valor por acción con "
        "WACC 10,03% y g 3,0% (fecha de valorización 30-06-2026). Actualizar si cambia el DCF.", "Team")
    ws["C15"].comment = Comment("Ponderación definida por el equipo: 80% DCF / 20% múltiplos.", "Team")

    # 2. Valor implícito por acción
    estilo(ws["A18"], bold=True, size=11, border=False)
    ws["A18"] = "2. Valor implícito por acción"
    encabezado(ws, 19, ["Múltiplo", "Base", "Peers", "Múltiplo P25", "Múltiplo mediana", "Múltiplo P75",
                        "Métrica Mallplaza (MM CLP)", "Valor/acción P25 (CLP)", "Valor/acción mediana (CLP)",
                        "Valor/acción P75 (CLP)", "Mediana vs precio"])
    for i, (mult, base, grupo, hoja, col, filas, metrica, tipo) in enumerate(METODOS):
        r = 20 + i
        a, b = filas.split(":")
        rango = f"{hoja}!{col}{a}:{col}{b}"
        valores = [mult, base, grupo,
                   f"=QUARTILE({rango},1)", f"=MEDIAN({rango})", f"=QUARTILE({rango},3)", f"={metrica}"]
        for j, v in enumerate(valores, start=1):
            c = ws.cell(r, j, v)
            estilo(c, fmt=MULT if j in (4, 5, 6) else NUM if j == 7 else None,
                   color=GREEN if j in (4, 5, 6) else BLACK, wrap=j == 3,
                   fill=MEDIAN_FILL if j == 5 else None)
        for j, m in zip((8, 9, 10), "DEF"):
            if tipo == "EV":
                f = f"=({m}{r}*$G{r}-$C$10-$C$11)/$C$12"
            else:
                f = f"={m}{r}*$G{r}/$C$12"
            estilo(ws.cell(r, j, f), fmt=NUM, bold=j == 9, fill=MEDIAN_FILL if j == 9 else None)
        estilo(ws.cell(r, 11, f"=I{r}/$C$13-1"), fmt=PCT)

    # 3. Precio objetivo combinado
    fila = 20 + len(METODOS) + 1
    estilo(ws.cell(fila, 1), bold=True, size=11, border=False)
    ws.cell(fila, 1, "3. Precio objetivo combinado (80% DCF / 20% múltiplos)")
    combinado = [
        ("Valor por múltiplos: EV/EBITDA LTM, mediana Core (CLP)", "=I20", SALMON),
        ("Valor DCF (CLP)", "=C14", None),
        ("Precio objetivo ponderado (CLP)", f"=C15*C{fila + 2}+C16*C{fila + 1}", SALMON),
        ("Precio al 30-sep-2026 (CLP)", "=C13", None),
        ("Upside / (downside)", f"=C{fila + 3}/C{fila + 4}-1", SALMON),
    ]
    for i, (label, f, fill) in enumerate(combinado):
        r = fila + 1 + i
        estilo(ws.cell(r, 1, label), bold=fill is not None, fill=fill)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        estilo(ws.cell(r, 3, f), bold=fill is not None, fmt=PCT if "Upside" in label else NUM, fill=fill)
    ws.cell(fila + 1, 4, "Método principal: EV/EBITDA (neutral a la contabilidad de valor razonable vs costo). "
                         "Cambiar la referencia de esta celda para usar otro método de la tabla 2.")
    estilo(ws.cell(fila + 1, 4), size=8, border=False)

    notas_ini = fila + 8
    estilo(ws.cell(notas_ini, 1), bold=True, size=11, border=False)
    ws.cell(notas_ini, 1, "Notas")
    notas = [
        "EV implícito = múltiplo × métrica de Mallplaza; patrimonio = EV − deuda neta − minoritarios; valor por acción = patrimonio / acciones.",
        "P/E y P/BV se aplican directo sobre patrimonio. P/BV mezcla peers a valor razonable (Chile) con peers a costo (Brasil): usar solo como referencia.",
        "P/E NTM: la utilidad de consenso de Mallplaza puede incluir revalorización de propiedades, lo que infla el valor implícito; usar solo como referencia.",
        "Peers Core = Parque Arauco, Cencosud Shopping, Multiplan, Allos, Iguatemi. IRSA y Fibra Danhos (Sensibilidad) quedan fuera de los estadísticos.",
        "Los peers brasileños transan a múltiplos menores por tasas de interés y riesgo país más altos; por eso se muestra también el grupo Chile.",
        "P25 y P75 con QUARTILE sobre los peers del grupo (con dos peers, entre el mínimo y el máximo).",
    ]
    for i, t in enumerate(notas):
        c = ws.cell(notas_ini + 1 + i, 1, f"• {t}")
        estilo(c, size=9, border=False)

    widths = {"A": 16, "B": 22, "C": 24, "D": 12, "E": 12, "F": 12, "G": 15, "H": 14, "I": 15, "J": 14, "K": 11}
    for k, w in widths.items():
        ws.column_dimensions[k].width = w
    for r in range(20, 20 + len(METODOS)):
        ws.row_dimensions[r].height = 26
    ws.row_dimensions[19].height = 40
    wb.save(ruta_out)
    print(f"Guardado {ruta_out}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    construir(sys.argv[1], sys.argv[2])
