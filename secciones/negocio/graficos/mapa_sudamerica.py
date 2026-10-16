"""BOCETO de la Figura 1 (Business Description): mapa de Sudamérica con los datos de cada país.

Chile, Perú y Colombia en rojo institucional; el resto del continente solo con contorno celeste. Cada país lleva
su ficha con: participación en el GLA, m² de GLA, participación en los ingresos y posición de mercado.
Debajo, los ingresos por fuente (UDM jun-26) como barra apilada (no son datos por país).

Datos (todos de los modelos del informe, nada escrito a mano):
  - GLA por país (%) e ingresos por fuente (%):  negocio/modelos/Excel_bussines_para_armar_graficos.xlsx
  - GLA total (m²) y posición/cuota de mercado:   industria/modelos/Excel_DATA_INDUSTRY.xlsx
  - Ingresos por país (%):                        valuacion/modelos/Mallplaza_Valoracion_Consolidada_v2.xlsx (WACC)
  - m² por país = % del GLA × GLA total (derivado)
Los límites de los países salen del dataset naturalearth_lowres que trae geopandas (<1.0).

Uso (desde secciones/):
    python negocio/graficos/mapa_sudamerica.py [carpeta_salida]
Salida: <carpeta>/{es,en}_N1_mapa_boceto.png
"""
import sys
from pathlib import Path

import geopandas as gpd
import matplotlib.pyplot as plt
import openpyxl
from matplotlib.patches import FancyBboxPatch, Rectangle

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "comun"))  # paleta institucional y rutas compartidas
import estilo  # noqa: E402,F401  (registra las carpetas graficos/ de todas las secciones)
from estilo import AMARILLO, CELESTE, NEGRO, ROJO, num
import negocio  # noqa: E402  (lector del Excel de Negocio y colores de las fuentes de ingreso)

SECCIONES = Path(__file__).resolve().parents[2]
AZUL = "#14284B"
DPI = 400
ANCHO, ALTO = 3.5, 6.4          # pulgadas de diseño; luego se escala al ancho de la columna (2,47 in)
EXT = (-82.5, -33.5, -56.5, 14.5)   # lon mín, lon máx, lat mín, lat máx
PAISES = ("Chile", "Perú", "Colombia")
NATURAL = {"Chile": "Chile", "Perú": "Peru", "Colombia": "Colombia"}
ANCLA = {"Chile": (-70.9, -33.4), "Perú": (-75.2, -9.5), "Colombia": (-73.6, 3.8)}   # punto de cada país
FICHA = {"Chile": (-48.5, -31.0), "Perú": (-48.5, -8.5), "Colombia": (-48.5, 6.5)}   # centro de cada ficha (lon, lat)

T = {
    "es": {"nombres": {"Chile": "CHILE", "Perú": "PERÚ", "Colombia": "COLOMBIA"},
           "gla": "GLA", "ing": "Ingresos", "pos": "Posición", "cuota": "cuota de GLA",
           "mm2": "M m²", "total": "37 centros", "total2": "{} M m² de GLA", "barra": "Ingresos por fuente (UDM jun-26)",
           "fuentes": {"Arriendo": "Arriendo", "Parking": "Parking",
                       "Servicios comerciales y otros": "Servicios comerciales y otros", "Publicidad": "Publicidad"},
           "nota": "* Colombia pro forma con Gran Plaza"},
    "en": {"nombres": {"Chile": "CHILE", "Perú": "PERU", "Colombia": "COLOMBIA"},
           "gla": "GLA", "ing": "Revenue", "pos": "Rank", "cuota": "GLA share",
           "mm2": "M m²", "total": "37 centers", "total2": "{} M m² of GLA", "barra": "Revenue by source (LTM Jun-26)",
           "fuentes": {"Arriendo": "Rent", "Parking": "Parking",
                       "Servicios comerciales y otros": "Commercial services & other", "Publicidad": "Advertising"},
           "nota": "* Colombia pro forma incl. Gran Plaza"},
}


def datos(ruta_negocio, ruta_industria, ruta_v2):
    gla, ingresos, _, _ = negocio.leer(ruta_negocio)
    wi = openpyxl.load_workbook(ruta_industria, data_only=True)
    total = next(r[2] for r in wi["EF. OPERATIVA (CLPM2)"].iter_rows(values_only=True) if r[0] == "Mallplaza")  # GLA m²
    cuota = {}
    for r in wi["MS_estatico_2Q26"].iter_rows(min_row=2, values_only=True):
        if r[1] == "Mallplaza":
            cuota[r[0]] = (r[5], r[4] * 100)            # país → ("#1", 33.0)
    ws = openpyxl.load_workbook(ruta_v2, data_only=True)["WACC"]
    ing_pais = {}
    for r in ws.iter_rows(values_only=True):
        for pais, clave in (("Chile", "Ponderación Chile"), ("Perú", "Ponderación Perú"), ("Colombia", "Ponderación Colombia")):
            if isinstance(r[0], str) and r[0].startswith(clave) and isinstance(r[2], float):
                ing_pais[pais] = r[2] * 100
    assert set(ing_pais) == set(PAISES) and set(cuota) == set(PAISES), (ing_pais, cuota)
    gla = dict(gla)
    return {"gla": gla, "gla_m2": {p: gla[p] / 100 * total for p in PAISES}, "total": total, "ing": ing_pais,
            "cuota": cuota, "fuentes": ingresos}


def ficha(ax, lang, pais, d, centro):
    """Recuadro con los datos del país, centrado en (lon, lat)."""
    t = T[lang]
    w, h = 27.0, 12.5                                    # en grados (el eje es lon/lat)
    x0, y0 = centro[0] - w / 2, centro[1] - h / 2
    ax.add_patch(FancyBboxPatch((x0, y0), w, h, boxstyle="round,pad=0,rounding_size=1.2", facecolor="white",
                                edgecolor=ROJO, linewidth=1.3, zorder=5))
    ax.text(x0 + 1.4, y0 + h - 1.4, t["nombres"][pais], fontsize=8.6, fontweight="bold", color=ROJO, va="top", zorder=6)
    pos, cuota = d["cuota"][pais]
    nota = "*" if pais == "Colombia" else ""
    filas = [(t["gla"], f"{num(d['gla'][pais], lang, 1)}%  ·  {num(d['gla_m2'][pais] / 1e6, lang, 2)} {t['mm2']}"),
             (t["ing"], f"{num(d['ing'][pais], lang, 1)}%"),
             (t["pos"], f"{pos}  ·  {num(cuota, lang, 0)}% {t['cuota']}{nota}")]
    for k, (a, b) in enumerate(filas):
        y = y0 + h - 4.4 - k * 2.75
        ax.text(x0 + 1.4, y, a, fontsize=6.9, color=AZUL, fontweight="bold", va="center", zorder=6)
        ax.text(x0 + 8.6, y, b, fontsize=6.9, color=NEGRO, va="center", zorder=6)
    # línea guía hacia el país
    ax.plot([ANCLA[pais][0], x0], [ANCLA[pais][1], centro[1]], color=AZUL, linewidth=0.8, zorder=4)
    ax.scatter([ANCLA[pais][0]], [ANCLA[pais][1]], s=16, facecolor="white", edgecolor=AZUL, linewidth=1.0, zorder=7)


def barra(fig, rect, lang, fuentes):
    t = T[lang]
    ax = fig.add_axes(rect)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.text(0, 1.55, t["barra"], fontsize=7.4, fontweight="bold", color=AZUL, va="bottom")
    x = 0
    for c, v in fuentes:
        color = negocio.COLOR_FUENTE[c]
        ax.add_patch(Rectangle((x, 0.1), v, 0.8, facecolor=color, edgecolor="white", linewidth=0.8))
        if v > 8:
            ax.text(x + v / 2, 0.5, num(v, lang, 1) + "%", ha="center", va="center", fontsize=7.2, fontweight="bold",
                    color="white" if color in (AZUL, ROJO) else NEGRO)
        x += v
    # leyenda en dos columnas con el valor de cada fuente
    for k, (c, v) in enumerate(fuentes):
        col, fila = k % 2, k // 2
        xl, yl = col * 60, -0.55 - fila * 0.95
        ax.add_patch(Rectangle((xl, yl - 0.2), 3.2, 0.4, facecolor=negocio.COLOR_FUENTE[c], clip_on=False))
        ax.text(xl + 5.0, yl, f"{t['fuentes'][c]}  {num(v, lang, 1)}%", fontsize=6.6, color=NEGRO, va="center", clip_on=False)


def figura(lang, d, salida):
    t = T[lang]
    sa = gpd.read_file(gpd.datasets.get_path("naturalearth_lowres"))
    sa = sa[(sa.continent == "South America") & (sa.name != "Falkland Is.")]
    fig = plt.figure(figsize=(ANCHO, ALTO))
    fig.patch.set_facecolor("white")
    ax = fig.add_axes([0.0, 0.205, 1.0, 0.795])
    ax.set_xlim(EXT[0], EXT[1])
    ax.set_ylim(EXT[2], EXT[3])
    ax.set_aspect(1 / 0.93)                       # grados de longitud algo más cortos que los de latitud
    ax.axis("off")
    for _, f in sa.iterrows():
        operador = f["name"] in NATURAL.values()
        gpd.GeoSeries([f.geometry]).plot(ax=ax, facecolor=ROJO if operador else "white",
                                         edgecolor="white" if operador else CELESTE, linewidth=0.7 if operador else 0.8, zorder=2)
    for pais in PAISES:
        ficha(ax, lang, pais, d, FICHA[pais])
    # total consolidado en el Atlántico sur
    ax.text(-45.5, -48.5, t["total"], fontsize=9.0, fontweight="bold", color=AZUL, ha="center", va="center")
    ax.text(-45.5, -52.4, t["total2"].format(num(d["total"] / 1e6, lang, 2)), fontsize=7.4, color=AZUL, ha="center", va="center")
    ax.text(EXT[1] - 0.4, EXT[2] + 0.2, t["nota"], fontsize=4.9, color="#5f5f5f", ha="right", va="bottom", style="italic")
    barra(fig, [0.04, 0.105, 0.92, 0.038], lang, d["fuentes"])
    fig.savefig(salida, dpi=DPI, facecolor="white")
    plt.close(fig)
    return salida


if __name__ == "__main__":
    salida = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent / "figuras"
    d = datos(SECCIONES / "negocio/modelos/Excel_bussines_para_armar_graficos.xlsx",
              SECCIONES / "industria/modelos/Excel_DATA_INDUSTRY.xlsx",
              SECCIONES / "valuacion/modelos/Mallplaza_Valoracion_Consolidada_v2.xlsx")
    print({k: d[k] for k in ("gla", "gla_m2", "total", "ing", "cuota")})
    for lang in ("es", "en"):
        print(figura(lang, d, salida / f"{lang}_N1_mapa_boceto.png"))
