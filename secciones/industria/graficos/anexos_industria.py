"""Textos y tablas de los Anexos 3 a 7 de Industria (español e inglés) a partir de la especificación del equipo.

Lee modelos/Anexos_Industry_especificacion.xlsx (una hoja por anexo: textos y datos) y vuelca figuras/datos_anexos.json,
que lee documento/industria.js para armar el anexo con el formato del informe (ver Industria_Mallplaza_v7_equipo.docx).
Los textos y cifras en español salen tal cual de la especificación; el inglés es traducción de esos mismos textos y las
tablas se formatean en cada idioma (punto de miles y coma decimal en español; al revés en inglés). Los gráficos de los
Anexos 3, 4 y 5 son los de graficos/industria.py (ind2, ind1 e ind7); los datos de esos gráficos se verifican contra la
especificación.

Uso (desde secciones/): python industria/graficos/anexos_industria.py [Anexos_Industry_especificacion.xlsx]
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

AQUI = Path(__file__).resolve().parent
SPEC = AQUI.parent / "modelos" / "Anexos_Industry_especificacion.xlsx"
DATOS = AQUI.parent / "modelos" / "Excel_DATA_INDUSTRY.xlsx"
SALIDA = AQUI / "figuras" / "datos_anexos.json"


def en_num(s):
    """Convierte números escritos en español (170.200; 95,6%) a formato inglés (170,200; 95.6%)."""
    return s.translate(str.maketrans(".,", ",."))


def millar(x, lang):
    t = f"{x:,.0f}"
    return t.replace(",", ".") if lang == "es" else t


def pct(x, lang, d=1):
    t = f"{x * 100:.{d}f}%"
    return t.replace(".", ",") if lang == "es" else t


EN = {
    "titulo_general": "APPENDICES: INDUSTRY AND COMPETITIVE POSITIONING",
    "a3_titulo": "APPENDIX 3: NEW GLA SUPPLY, 2026-2028",
    "a3_intro": "The trio will add 432,030 m² between 2026 and 2028, of which 334,030 m² are organic (6.4% of their current total GLA; in Chile, 219,900 m², equivalent to 6.7% of their GLA in that country). Acquisitions (98,000 m², all by Parque Arauco) change owner but add no supply to the industry; the supply-versus-demand analysis for Chile (Figure 3) therefore uses organic GLA only.",
    "a3_subA": "A. PIPELINE BY COMPANY (m²)", "a3_subB": "B. GLA BY OPENING YEAR (m²)",
    "a3_notaB": "Org./GLA = organic GLA 2026-2028 over consolidated GLA at 2Q26. Mallplaza: 2028 includes 46,500 m² of “Other projects Chile”, dated 2028 by the 2025 Annual Report (p.45). Parque Arauco: 2026 = 2H26 (acquisitions). Cenco: deliveries given as a range are placed at the end of the range.",
    "a3_fuente": "Source: Plaza S.A., 2Q26 Release (p.13) and 2025 Annual Report (p.45); Parque Arauco, 2Q26 Investor Presentation (slide 23); Cenco Malls, 2Q26 Corporate Presentation (slides 15-17). Dated projects only.",
    "a3_cols": ["Company", "2026", "2027", "2028", "Organic", "Acquis.", "Total", "Org./GLA"],
    "a4_titulo": "APPENDIX 4: OCCUPANCY",
    "a4_intro": "The consolidated occupancy of the three operators has stayed between 93% and 99% since 2Q22, even through the inflation and high-rate cycle. Each company defines occupancy differently, so trends are compared rather than levels.",
    "a4_sub": "A. CONSOLIDATED OCCUPANCY RATE, 2Q22-2Q26",
    "a4_fuente": "Source: 2Q22-2Q26 earnings reports of the three companies. Mallplaza includes Open Plaza since 2Q25.",
    "a5_titulo": "APPENDIX 5: MARKET SHARE IN CHILE",
    "a5_intro": "The trio's GLA in Chile grew 7.3% since 2Q22. Mallplaza gave up 1.2 pp of share because it grew only through expansions, while Parque Arauco added Kennedy Oriente (70,000 m²) by purchase.",
    "a5_sub": "A. SHARE WITHIN THE TRIO, 2Q22-2Q26 (% OF GLA)",
    "a5_fuente": "Source: 2Q22-2Q26 earnings reports of the three companies; Team analysis.",
    "a6_titulo": "APPENDIX 6: PORTER'S FIVE FORCES",
    "a6_intro": "Attractive industry: with whole-number scores by country weighted by Mallplaza's GLA, the only intense force is rivalry (4), which plays out in experience and assets rather than price, since in 2Q26 same-store rents rose in Chile at all three operators (table B).",
    "a6_subA": "A. SCORES BY COUNTRY AND MAIN EVIDENCE (1 = WEAK FORCE; 5 = INTENSE FORCE)",
    "a6_fuenteA": "Source: Team analysis. Final score = average weighted by Mallplaza's GLA at 2Q26 (Chile 61.9%, Peru 26.2%, Colombia 11.9%), rounded to the nearest integer. For new entrants the weighted score sits at the threshold (1.5) and the team sets it at 2 because of lower entry barriers in Peru and Colombia.",
    "a6_subB": "B. PRICE INDICATORS BY OPERATOR, 2Q26",
    "a6_fuenteB": "Source: Plaza S.A., 2Q26 Release (p.18); Cenco Malls, 2Q26 Earnings Presentation (p.17); Parque Arauco, 2Q26 Earnings Report (p.5, p.6 and p.21). Each indicator follows each company's definition: Mallplaza's SSR considers fixed rent in Chile and the change in the UF; Cenco's includes variable rent (Cenco 1Q26: “SSR Chile reflects the effect of lower sales”); Parque Arauco's SAR measures rent collected in areas leased in both periods. *Includes common expenses and promotion fund: its level is not comparable, its change is. Mallplaza's revenue per m² change in Colombia (+8.0%) is a team calculation: 84,790 / 78,545 − 1. Everything else is published data.",
    "a6_colsA": ["Force", "Chile", "Peru", "Colombia", "Final", "Main evidence"],
    "a6_fuerzas": ["Rivalry", "New entrants", "Tenant bargaining power", "Supplier bargaining power", "Substitutes"],
    "a6_evidencia": [
        "CR3 76% / 37% / 13%; reinvestment race in experience; asset purchases (Minka, Plaza Central; Quilín and Chicureo announced); supply above consumption in 2027-28. Stable shares and margins of 71-90%. No price war: positive SSR/SAR in Chile (table B).",
        "Capital intensive (USD 613 million for 170,200 m² at Mallplaza); incumbents' land banks; Peru and Colombia less concentrated; Colombia with most malls under horizontal ownership.",
        "Chile occupancy 96-99%; contracts averaging 5.0-8.7 years; top 5 = 18% of Mallplaza's revenue; Chile SSR above CPI (6.9% vs 4.8%, 1Q23-2Q26); department stores 19.0% of GLA.",
        "Common expenses passed through to tenants (2025 Annual Report, note 2.20); cost of sales 6.2% of Mallplaza's revenue.",
        "93% of Mallplaza's rental income is fixed; 2Q26 visits +4.4% Chile, +1.3% Peru, +2.5% Colombia; shift to experience. No official e-commerce data by country."],
    "a6_colsB": ["Indicator", "Country", "Mallplaza", "Cenco Malls", "Parque Arauco"],
    "a6_indic": ["Same-store or same-area rent (YoY change)", "Monthly revenue per m² (YoY change)", "Occupancy cost (2Q25)"],
    "a7_titulo": "APPENDIX 7: MAIN ASSUMPTIONS",
    "a7_intro": "Criteria used so that figures are comparable across operators and consistent with the official 2Q26 reports.",
    "a7_cols": ["Topic", "Assumption"],
    "a7_temas": ["Market share", "Pipeline", "Supply vs demand", "EBITDA per m²", "Contracts", "Real sales", "Real Plaza (Peru)", "Dates"],
    "a7_textos": [
        "Measured by GLA (not by sales) over each country's main operators (slide 7, 2Q26 Institutional Presentation). Partial HHI: considers only the operators identified on slide 7, so it is a floor (the real HHI is equal or higher); classified with the 2023 DOJ/FTC Merger Guidelines thresholds (above 1,800 = highly concentrated). Colombia pro forma: includes Gran Plaza (announced acquisition of 179,667 m²; 2Q26 Release); without it, Mallplaza's share would be 4.3%.",
        "Only dated projects in official reports at 2Q26 and the 2025 Annual Report; dates given as a range are placed at the end of the range; acquisitions are not new supply for the industry.",
        "Chile only and only m² that depend on consumption (excludes offices and multifamily); base GLA excludes Cenco's offices. Private consumption from the September 2026 IPoM is used as a demand indicator, not as a sales projection.",
        "EBITDA per each company's definition (none includes the IAS 40 revaluation); 2Q26 EBITDA not annualized; GLA consolidated at 100% of controlled assets. Parque Arauco: multifamily NOI (CLP 503 million) is subtracted, since its size is reported in units and not in m²; effect −0.7%, does not change the ranking.",
        "Profile and duration weighted by GLA. Cenco states it explicitly; Parque Arauco presents its maturity profile as % of GLA but does not state the basis of its average duration; Mallplaza does not publish the basis. In both cases GLA is assumed.",
        "Tenant sales in Chile 2Q26 vs 2Q25 deflated with the 12-month CPI of June 2026 (4.3%; INE, release of 08-07-2026). Cenco 2Q25 calculated from its published change (−1.9%).",
        "InRetail 2Q26 adjusted EBITDA, Shopping Malls segment (S/ 131 million; excludes investment property revaluation) ÷ total GLA of 888,000 m² (InRetail, 2Q26 Corporate Presentation, p.26 and p.29). Mallplaza Peru: published monthly EBITDA per m² (S/ 53) × 3 (2Q26 Release, p.18). Margins are not compared: InRetail calculates its own on net rent.",
        "Company data as of 30-Jun-2026; macro projections from the September 2026 IPoM as the most recent available."],
}
PAIS_EN = {"Perú": "Peru"}


def en_celda(s):
    s = en_num(s)
    return s.replace("No publica", "Not published")


def leer(spec=SPEC):
    wb = openpyxl.load_workbook(spec, data_only=True)
    a3, a4, a5, a6, a7 = (wb[n] for n in ("Anexo_3_Pipeline", "Anexo_4_Ocupacion", "Anexo_5_MS_Chile", "Anexo_6_Porter", "Anexo_7_Supuestos"))
    es = {"titulo_general": wb["LEEME_estilo"]["C21"].value,
          "a3_titulo": a3["B6"].value, "a3_intro": a3["B7"].value, "a3_subA": a3["B8"].value, "a3_subB": a3["B9"].value,
          "a3_notaB": a3["B10"].value, "a3_fuente": a3["B11"].value,
          "a3_cols": [a3.cell(22, c).value for c in range(1, 9)].copy(),
          "a4_titulo": a4["B6"].value, "a4_intro": a4["B7"].value, "a4_sub": a4["B8"].value, "a4_fuente": a4["B9"].value,
          "a5_titulo": a5["B6"].value, "a5_intro": a5["B7"].value, "a5_sub": a5["B8"].value, "a5_fuente": a5["B9"].value,
          "a6_titulo": a6["B6"].value, "a6_intro": a6["B7"].value, "a6_subA": a6["B8"].value, "a6_fuenteA": a6["B9"].value,
          "a6_subB": a6["B10"].value, "a6_fuenteB": a6["B11"].value + " " + a6["B32"].value,
          "a6_colsA": [a6.cell(14, c).value for c in range(1, 7)], "a6_colsB": [a6.cell(22, c).value for c in range(1, 6)],
          "a6_fuerzas": [a6.cell(r, 1).value for r in range(15, 20)], "a6_evidencia": [a6.cell(r, 6).value for r in range(15, 20)],
          "a6_indic": [a6["A23"].value, a6["A26"].value, a6["A29"].value],
          "a7_titulo": a7["B6"].value, "a7_intro": a7["B7"].value, "a7_cols": [a7["A10"].value, a7["B10"].value],
          "a7_temas": [a7.cell(r, 1).value for r in range(11, 19)], "a7_textos": [a7.cell(r, 2).value for r in range(11, 19)]}
    t3 = [[a3.cell(r, c).value for c in range(1, 9)] for r in range(23, 27)]
    t6a = [[a6.cell(r, c).value for c in range(1, 6)] for r in range(15, 20)]
    t6b = [[a6.cell(r, c).value for c in range(1, 6)] for r in range(23, 32)]
    return es, t3, t6a, t6b, wb


def verificar_graficos(wb):
    """Los gráficos de los Anexos 3-5 salen de Excel_DATA_INDUSTRY.xlsx: sus datos deben coincidir con la especificación."""
    d = openpyxl.load_workbook(DATOS, data_only=True)
    p = d["PIPELINE"]
    a3 = wb["Anexo_3_Pipeline"]
    assert (p["D12"].value, p["I12"].value, p["N12"].value, p["O12"].value) == (a3["B16"].value, a3["B18"].value, a3["B17"].value, a3["C17"].value)
    o, a4 = d["Ocupacion_consolidada"], wb["Anexo_4_Ocupacion"]
    for fila_spec, fila_dato in ((14, 5), (15, 6), (16, 7)):
        for c in range(2, 7):
            assert abs(a4.cell(fila_spec, c).value - o.cell(fila_dato + 1, c).value) < 5e-4, (fila_spec, c)
    m, a5 = d["MS_evolutivo_Chile"], wb["Anexo_5_MS_Chile"]
    for fila_spec, fila_dato in ((14, 13), (15, 14), (16, 15)):
        for c in range(2, 7):
            assert abs(a5.cell(fila_spec, c).value - m.cell(fila_dato, c).value) < 5e-4, (fila_spec, c)
    return True


def tablas(es, t3, t6a, t6b, lang):
    cols = es["a3_cols"] if lang == "es" else EN["a3_cols"]
    filas3 = []
    for f in t3:
        filas3.append([f[0]] + [("–" if (v == 0) else millar(v, lang)) for v in f[1:7]] + [pct(f[7], lang)])
    a3 = {"cols": cols, "filas": filas3, "destacar": [0], "negritaFilas": [3]}
    esl = lang == "es"
    fz = es["a6_fuerzas"] if esl else EN["a6_fuerzas"]
    ev = es["a6_evidencia"] if esl else EN["a6_evidencia"]
    a6a = {"cols": es["a6_colsA"] if esl else EN["a6_colsA"], "destacar": [0],
           "filas": [[fz[i], *f[1:5], ev[i]] for i, f in enumerate(t6a)]}
    ind = es["a6_indic"] if esl else EN["a6_indic"]
    filas6b, grupo = [], -1
    for f in t6b:
        if f[0]:
            grupo += 1
        pais = f[1] if esl else PAIS_EN.get(f[1], f[1])
        vals = [str(v) for v in f[2:5]]
        filas6b.append([ind[grupo] if f[0] else "", pais] + (vals if esl else [en_celda(v) for v in vals]))
    a6b = {"cols": es["a6_colsB"] if esl else EN["a6_colsB"], "destacar": [0], "filas": filas6b}
    return a3, a6a, a6b


def ajustar_gla_industria(es):
    """El GLA de la industria en Chile (no el del trío) es la base de la oferta nueva: ≈ GLA del trío sin oficinas ÷ 76% (v8 del texto)."""
    es["a3_intro"] = es["a3_intro"].replace("equivalentes al 6,7% de su GLA en ese país", "equivalentes al 5,1% del GLA de la industria en Chile y al 6,7% del GLA del propio trío")
    es["a3_notaB"] += " En Chile, la oferta orgánica del trío (219.900 m²) equivale al 5,1% del GLA estimado de la industria (≈4,3 millones de m²: 3,29 millones del trío sin oficinas ÷ 76%, su participación en los principales operadores)."
    es["a7_textos"] = [t + (" GLA de la industria en Chile ≈ GLA del trío sin oficinas ÷ 76% (participación del trío en los principales operadores, slide 7); la oferta del trío es una cota inferior de la oferta de la industria."
                            if es["a7_temas"][i] == "Oferta vs demanda" else "") for i, t in enumerate(es["a7_textos"])]
    es["a6_evidencia"] = [t.replace("oferta sobre consumo en 2027-28", "oferta que alcanza al consumo en 2027") for t in es["a6_evidencia"]]
    EN["a3_intro"] = EN["a3_intro"].replace("equivalent to 6.7% of their GLA in that country", "equivalent to 5.1% of the industry's GLA in Chile and 6.7% of the trio's own GLA")
    EN["a3_notaB"] += " In Chile, the trio's organic supply (219,900 m²) equals 5.1% of the estimated industry GLA (≈4.3 million m²: 3.29 million of the trio excluding offices ÷ 76%, its share of the main operators)."
    EN["a6_evidencia"] = [t.replace("supply above consumption in 2027-28", "supply catching up with consumption in 2027") for t in EN["a6_evidencia"]]
    if "a7_textos" in EN:
        EN["a7_textos"] = [t + (" Industry GLA in Chile ≈ trio GLA excluding offices ÷ 76% (the trio's share of the main operators, slide 7); the trio's supply is a lower bound of the industry's supply."
                                if i == 2 else "") for i, t in enumerate(EN["a7_textos"])]
    return es


def main(spec=SPEC):
    es, t3, t6a, t6b, wb = leer(spec)
    es = ajustar_gla_industria(es)
    verificar_graficos(wb)
    # el total del trío debe cuadrar con la suma de las tres empresas
    assert sum(f[6] for f in t3[:3]) == t3[3][6] and sum(f[4] for f in t3[:3]) == t3[3][4]
    salida = {}
    for lang in ("es", "en"):
        base = es if lang == "es" else EN
        a3, a6a, a6b = tablas(es, t3, t6a, t6b, lang)
        txt = {k: v for k, v in base.items() if not k.startswith("a3_cols")}
        # el tema y el texto de cada supuesto del Anexo 7 forman la tabla
        txt["a7_tabla"] = {"cols": base["a7_cols"], "filas": [[t, x] for t, x in zip(base["a7_temas"], base["a7_textos"])]}
        salida[lang] = {**txt, "a3_tabla": a3, "a6_tablaA": a6a, "a6_tablaB": a6b}
    SALIDA.write_text(json.dumps(salida, ensure_ascii=False, indent=1))
    print(SALIDA, "· tablas:", len(t3), "filas A3,", len(t6a), "A6-A,", len(t6b), "A6-B, 8 supuestos")


if __name__ == "__main__":
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else SPEC)
