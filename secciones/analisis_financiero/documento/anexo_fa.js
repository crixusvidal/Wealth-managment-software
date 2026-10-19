// Anexos de Análisis Financiero (9 a 14 en el informe consolidado; 1 a 6 en la numeración local), versión v2 del equipo
// («Contenido nuevo v2»): supuestos del forecast, estados financieros con balance vertical y fuentes y usos de caja (2D), ratios con
// ROIC − WACC real, puente DFN/EBITDA (3B) y cifras del texto (3C), EBITDA ajustado, proyección del caso base con el capex separado y
// comparación con los pares en base comparable. Las cifras históricas vienen de los EEFF y del release (tal como las entregó el
// equipo); la proyección (Anexo 5) sale del modelo de valoración v4 vía comun/cifras.js.
const cifras = require("../../comun/cifras");

module.exports = (H) => ({
  anexos(lang) {
    const { K, f } = cifras(lang);
    const es = lang === "es", s = K.s;
    const h = (x) => (es ? x[0] : x[1]);                    // [español, inglés]
    const num = (x, d = 1) => (x < 0 ? "(" + f.n(Math.abs(x), d) + ")" : f.n(x, d));
    const mm = (x, d = 1) => num(x / 1000, d);              // CLP MM → CLP mil MM
    const dash = "—";
    const hist = (txt) => (es ? txt : txt.replace(/,/g, "§").replace(/\./g, ",").replace(/§/g, "."));   // cifras escritas en formato español → inglés
    const T = (rows) => rows.map((r) => r.map((c) => (c === "n.s." ? (es ? c : "n.m.") : typeof c === "string" && /^[−+(]?\d[\d.,]*[x%)]*(?: pp| \/ \d[\d.,]*)?$/.test(c) ? hist(c) : c)));
    const ingTot = [s.ing_cal[0], ...s.ing.slice(1, 6)];
    const ebitdaTot = [s.ebitda_cal[0], ...s.ebitda.slice(1, 6)];
    const da = [5558, ...s.da.slice(1, 6)];                 // 2026E = 1S26 real (2.993) + 2S26E (2.565)
    const ebit = ebitdaTot.map((x, i) => x - da[i]);
    const elim = ingTot.map((x, i) => x - s.ing_exp[i] * (i ? 1 : 0) - (s.ing_cl[i] + s.ing_pe[i] + s.ing_co[i]));
    const capexTot = [84241 + abs(s.capex_base[0]) + abs(s.capex_exp[0]), ...[1, 2, 3, 4, 5].map((i) => abs(s.capex_base[i]) + abs(s.capex_exp[i]))];
    function abs(x) { return Math.abs(x); }
    const fila = (nombre, arr, d = 1, fn = mm) => [nombre, ...arr.map((x) => (x === null ? dash : fn(x, d)))];
    const pct = (arr) => arr.map((x) => f.pc(x, 1));

    // ---------------------------------------------------------------- Anexo 9 (1): supuestos
    const a1 = [h(["Anexo 1:", "Appendix 1:"]), h(["Supuestos del forecast (CLP nominales, 2026E–2031E)", "Forecast assumptions (nominal CLP, 2026E–2031E)"]),
      h([["Variable", "Supuesto", "Base", "Tag"], ["Variable", "Assumption", "Basis", "Tag"]]),
      h([[
        ["Ancla 2026E", "1S26 real + 2S26E que repite el crecimiento en moneda local de 1S26 (Chile 6,1%, Perú 7,9%, Colombia 7,3%)", "EEFF intermedios jun-26", "R/S"],
        ["Inflación", "IPC de consenso 2026–28; metas de los bancos centrales desde 2029 (CL 3%, PE 2%, CO 3%)", "Arriendos indexados a UF/IPC", "P/S"],
        ["Crecimiento real mismas tiendas (sin parking)", "Chile 1,5% → 0,6%; Perú 1,5% → 0,4%; Colombia 2,5% → 2,1% (2027E → 2031E)", "Puntos medios de la tasación (Nota 11)", "S"],
        ["Ocupación", "Chile 96,1% y Colombia 97,5% constantes; Perú 93,8% → 94,6% en 2028E", "2T26 real; Perú vuelve a 2025", "R/S"],
        ["Parking", "12,1% / 6,4% / 9,0% de los ingresos 2026; crecimiento real 9% → 2%", "Release 2T26 (UDM +23,7%)", "R/S"],
        ["GLA del pipeline", `8.900 m² (2026), 72.950 m² (2027), 88.350 m² (2028, con 46.500 m² de otros proyectos en Chile); rampa 35/90/100%; ${f.n(K.pipeline_m2)} m² en total`, "Release 2T26; Memoria 2025 (p. 45)", "R/S"],
        ["Expansión continua", `${f.n(K.gla_anual)} m²/año que abren desde 2029E (${f.pc(K.mix_chile, 0)} Chile / ${f.pc(1 - K.mix_chile, 0)} Perú); USD ${f.n(K.costo_usd_m2)}/m² a IPC; capex 50% en t−2 y 50% en t−1 (desde 2027E); margen 80,7%; yield on cost ${f.pc(K.yield_costo, 1)}`, "GLA orgánico prom. 2022, 2023, 2025; costo/m² del pipeline", "C/S"],
        ["Tipo de cambio", "Consenso de fin de año 2026–28; paridad de poder de compra después", "Bloomberg ECFC, 5-oct-26", "P/S"],
        ["Margen EBITDA", "Chile 80%, Perú 82%, Colombia 74% (≈ 79,7% consolidado)", "2024–UDM 2T26; sin el máximo 2025", "S"],
        ["D&A", "0,704% de los ingresos desde 2S26E; 2026E = 1S26 real (2.993) + 2S26E (2.565)", "UDM 2T26 (4.766 / 677.333)", "C"],
        ["Impuesto en caja", "16,8% (2026E) → 19% → 21,5% → 24% → 27% (2030E–31E), tasa de Chile sobre todo el EBIT", "Impuesto corriente 1S26; Perú/Colombia en Anexo 18", "C/S"],
        ["Capex de mantención + TI", "5,5% de los ingresos (2S26E: medio año)", "2024 (4,8%) – UDM 2T26 (~6,6%)", "S"],
        ["Capex incremental: pipeline", `USD ${f.n(s.capex_usd.slice(0, 3).reduce((a, b) => a + b, 0), 0)} MM restantes a jun-26 (USD 445 MM de proyectos con fecha + USD 134 MM de otros proyectos en Chile), repartidos hasta la apertura (2S26E–28E)`, "Release 2T26, pipeline; análisis del equipo", "R/S"],
        ["Capex incremental: expansión", "Expansión continua (arriba); capex incremental terminal = 0 (g real 0%)", `RONIC ${f.pc(K.ronic, 1)} < WACC real ${f.pc(K.wacc_t_real, 1)}`, "C/S"],
        ["Capital de trabajo", "Capital de trabajo neto operativo en 15,6% de los ingresos (dic-25); 2026E cubre solo 2S26E", "EEFF dic-25, primer cierre con Perú completo", "C/S"],
        ["Apalancamiento", "DFN/EBITDA 2,36x (jun-26) → 3,5x a dic-31E, lineal; la caja sobre el plan se devuelve con recompras o dividendos extraordinarios", "Supuesto del equipo; pares 2,1x–5,7x", "S"],
        ["Costo de la deuda", "Costo financiero neto 3,08% de la DFN inicial + 25 pb reales por vuelta sobre 2,36x; reajuste UF sobre 79% de la deuda", "1S26 anualizado; supuesto del equipo", "C/S"],
        ["Excluido: greenfield (Mallplaza Cali)", "USD 0,7 MM UDM jun-26 vs USD 55,2 MM en 2023", "Inmaterial", "S"],
        ["Excluido: Gran Plaza", "Hecho posterior: contrato del 31-jul-26 sujeto a la SIC (Nota 30); USD 376 MM a cap rate 9,4% (NOI UDM); ~−CLP 11/acción a nuestro cap rate (Anexo 18)", "Posterior a la valoración (30-jun-26)", "S"],
        ["Payout", "40% de la utilidad distribuible con un año de rezago, más las distribuciones adicionales", "Política del directorio (Nota 20e)", "R/S"]],
      [
        ["2026E anchor", "1H26 actual + 2H26E repeating the 1H26 local-currency growth (Chile 6.1%, Peru 7.9%, Colombia 7.3%)", "Interim financials Jun-26", "R/S"],
        ["Inflation", "Consensus CPI 2026–28; central-bank targets from 2029 (CL 3%, PE 2%, CO 3%)", "Rents indexed to UF/CPI", "P/S"],
        ["Real same-store growth (ex parking)", "Chile 1.5% → 0.6%; Peru 1.5% → 0.4%; Colombia 2.5% → 2.1% (2027E → 2031E)", "Appraisal midpoints (Note 11)", "S"],
        ["Occupancy", "Chile 96.1% and Colombia 97.5% constant; Peru 93.8% → 94.6% by 2028E", "2Q26 actual; Peru returns to 2025", "R/S"],
        ["Parking", "12.1% / 6.4% / 9.0% of 2026 revenue; real growth 9% → 2%", "2Q26 release (LTM +23.7%)", "R/S"],
        ["Pipeline GLA", `8,900 m² (2026), 72,950 m² (2027), 88,350 m² (2028, incl. 46,500 m² of other projects in Chile); 35/90/100% ramp; ${f.n(K.pipeline_m2)} m² in total`, "2Q26 release; 2025 Annual Report (p. 45)", "R/S"],
        ["Continuous expansion", `${f.n(K.gla_anual)} m²/yr opening from 2029E (${f.pc(K.mix_chile, 0)} Chile / ${f.pc(1 - K.mix_chile, 0)} Peru); USD ${f.n(K.costo_usd_m2)}/m² indexed to CPI; capex 50% in t−2 and 50% in t−1 (from 2027E); 80.7% margin; yield on cost ${f.pc(K.yield_costo, 1)}`, "Organic GLA avg 2022, 2023, 2025; pipeline cost/m²", "C/S"],
        ["FX", "Year-end consensus 2026–28; purchasing-power parity afterwards", "Bloomberg ECFC, 5-Oct-26", "P/S"],
        ["EBITDA margin", "Chile 80%, Peru 82%, Colombia 74% (≈ 79.7% consolidated)", "2024–LTM 2Q26; 2025 peak not extrapolated", "S"],
        ["D&A", "0.704% of revenue from 2H26E; 2026E = 1H26 actual (2,993) + 2H26E (2,565)", "LTM 2Q26 (4,766 / 677,333)", "C"],
        ["Cash tax", "16.8% (2026E) → 19% → 21.5% → 24% → 27% (2030E–31E), Chile's rate on all EBIT", "1H26 current tax; Peru/Colombia in Appendix 18", "C/S"],
        ["Maintenance + IT capex", "5.5% of revenue (2H26E: half year)", "2024 (4.8%) – LTM 2Q26 (~6.6%)", "S"],
        ["Incremental capex: pipeline", `USD ${f.n(s.capex_usd.slice(0, 3).reduce((a, b) => a + b, 0), 0)}m remaining at Jun-26 (USD 445m dated projects + USD 134m other projects in Chile), spread until opening (2H26E–28E)`, "2Q26 release, pipeline; Team analysis", "R/S"],
        ["Incremental capex: expansion", "Continuous expansion (above); terminal incremental capex = 0 (real g 0%)", `RONIC ${f.pc(K.ronic, 1)} < real WACC ${f.pc(K.wacc_t_real, 1)}`, "C/S"],
        ["Working capital", "Net operating working capital at 15.6% of revenue (Dec-25); 2026E covers 2H26E only", "Dec-25 financials, first close with full Peru", "C/S"],
        ["Leverage", "Net debt/EBITDA 2.36x (Jun-26) → 3.5x by Dec-31E, linear; cash above plan returned via buybacks or special dividends", "Team assumption; peers 2.1x–5.7x", "S"],
        ["Cost of debt", "Net financial cost 3.08% of opening net debt + 25bp real per turn above 2.36x; UF indexation on 79% of debt", "1H26 annualized; team assumption", "C/S"],
        ["Excluded: greenfield (Mallplaza Cali)", "USD 0.7m LTM Jun-26 vs USD 55.2m in 2023", "Immaterial", "S"],
        ["Excluded: Gran Plaza", "Subsequent event: 31-Jul-26 contract subject to the SIC (Note 30); USD 376m at a 9.4% cap rate (LTM NOI); ~−CLP 11/share at our cap rate (Appendix 18)", "After the valuation date (30-Jun-26)", "S"],
        ["Payout", "40% of distributable profit with a one-year lag, plus the additional distributions", "Board policy (Note 20e)", "R/S"]]])];
    // la 3ª posición contiene [esp, ing]: h() ya eligió; ajustar nota
    const nota1 = es ? "Tags: R reportado · C cálculo del equipo · S supuesto del equipo · P consenso de mercado. Fuente: EEFF Plaza S.A. (CMF); release 2T26; Bloomberg ECFC; PwC Tax Summaries (jul-26); análisis del equipo."
      : "Tags: R reported · C team calculation · S team assumption · P market consensus. Source: Plaza S.A. financial statements (CMF); 2Q26 release; Bloomberg ECFC; PwC Tax Summaries (Jul-26); Team analysis.";

    // ---------------------------------------------------------------- Anexo 10 (2): estados financieros
    const IS = T([["370,9", "419,9", "494,6", "653,9", "677,3"], ["282,7", "322,2", "380,6", "526,2", "544,9"], ["(2,2)", "(2,2)", "(2,1)", "(3,6)", "(4,8)"],
      ["83,1", "193,7", "157,3", "1.445,4", "1.190,4"], ["363,6", "513,8", "535,9", "1.968,1", "1.730,6"], [dash, "(23,8)", "(30,9)", "(54,0)", "(44,9)"],
      [dash, "(55,3)", "(53,2)", "(43,8)", "(50,4)"], ["157,7", "349,8", "335,9", "1.431,0", "1.258,1"], ["218,9", "257,2", "303,2", "405,3", "419,9"]]);
    const BS = T([[dash, "405,0", "472,1", "507,1", "6,4%"], ["202,0", "179,2", "268,3", "347,6", "4,4%"], [dash, "5.510,1", "7.111,5", "7.478,6", "93,6%"],
      ["4.166,5", "5.356,5", "6.961,9", "7.309,8", "91,5%"], ["4.796,3", "5.915,1", "7.583,6", "7.985,7", "100%"], [dash, "344,5", "307,1", "203,2", "2,5%"],
      [dash, "2.379,5", "2.752,0", "2.910,8", "36,5%"], ["1.264,0", "1.617,6", "1.579,3", "1.633,4", "20,5%"], ["656,5", "892,9", "1.248,8", "1.313,7", "16,5%"],
      [dash, "3.191,0", "4.524,6", "4.871,7", "61,0%"], ["2.601,1", "3.157,9", "4.482,9", "4.820,9", "60,4%"]]);
    const CF = T([["295,1", "300,8", "441,1", "258,2"], ["(119,1)", "(76,3)", "(103,6)", "(84,2)"], ["(164,6)", "(400,4)", "(0,3)", dash], [dash, "308,4", dash, dash],
      ["(124,0)", "1,3", "(86,5)", "5,5"], ["(55,6)", "(54,4)", "(68,2)", "(30,1)"], ["(0,0)", "(125,8)", "(108,5)", "(87,6)"]]);
    const a2 = {
      eIS: h([["Estado de resultados", "2022A", "2023A", "2024A", "2025A", "UDM 2T26"], ["Income statement", "2022A", "2023A", "2024A", "2025A", "LTM 2Q26"]]),
      rIS: h([["Ingresos", "EBITDA (definición compañía)", "D&A", "Revalorización y otras partidas op., neto", "Resultado operacional", "Costo financiero neto", "Reajuste UF de la deuda", "Utilidad atribuible", "FFO atribuible (compañía)"],
        ["Revenue", "EBITDA (company definition)", "D&A", "Revaluation and other op. items, net", "Operating profit", "Net financial cost", "UF indexation of debt", "Attributable profit", "Attributable FFO (company)"]]),
      eBS: h([["Balance", "dic-23", "dic-24", "dic-25", "jun-26", "% act."], ["Balance sheet", "Dec-23", "Dec-24", "Dec-25", "Jun-26", "% assets"]]),
      rBS: h([["Activos corrientes", "Caja y equivalentes", "Activos no corrientes", "Propiedades de inversión", "Activos totales", "Pasivos corrientes", "Pasivos no corrientes", "Deuda financiera bruta (CP + LP)", "Pasivos por impuesto diferido", "Patrimonio total", "Patrimonio atribuible"],
        ["Current assets", "Cash and equivalents", "Non-current assets", "Investment property", "Total assets", "Current liabilities", "Non-current liabilities", "Gross financial debt (ST + LT)", "Deferred tax liabilities", "Total equity", "Attributable equity"]]),
      eCF: h([["Flujo de caja", "2023A", "2024A", "2025A", "1S26"], ["Cash flow", "2023A", "2024A", "2025A", "1H26"]]),
      rCF: h([["Flujo operativo", "Capex", "Adquisiciones (incl. minoritarios)", "Emisión de acciones", "Endeudamiento neto / (pago)", "Intereses pagados", "Dividendos pagados"],
        ["Operating cash flow", "Capex", "Acquisitions (incl. minorities)", "Share issuance", "Net borrowing / (repayment)", "Interest paid", "Dividends paid"]]),
      nota: es ? "UDM = 2025 − 1S25 + 1S26. 2024 reexpresado en los EEFF dic-25; 2022A de los EEFF dic-23 reexpresados. Análisis vertical: % de los activos totales a jun-26; el balance es 92% propiedades a valor razonable y los impuestos diferidos (CLP 1.314 mil MM) equivalen al 80% de la deuda bruta. Fuente: EEFF Plaza S.A. (CMF) dic-24, dic-25, jun-26."
        : "LTM = 2025 − 1H25 + 1H26. 2024 restated in the Dec-25 financials; 2022A from the restated Dec-23 financials. Vertical analysis: % of total assets at Jun-26; the balance sheet is 92% investment property at fair value and deferred taxes (CLP 1,314bn) equal 80% of gross debt. Source: Plaza S.A. financial statements (CMF) Dec-24, Dec-25, Jun-26." };
    // 2D: fuentes y usos
    const SU = T([["441,1", "258,2"], ["245,7", "57,9"], ["17,6", "12,5"], ["704,5", "328,7"], ["103,6", "84,2"], ["332,2", "52,4"], ["68,2", "30,1"], ["108,5", "87,6"], ["2,9", "1,4"], ["89,0", "72,9"], ["704,5", "328,7"],
      ["2,6x", "1,7x"], ["1,8x", "1,8x"], ["18,1", "78,0"], ["0,29x", "0,26x"]]);
    const rSU = h([["Flujo operativo", "Deuda nueva (bonos + préstamos)", "Intereses recibidos y otros", "Total fuentes", "Capex (PI + intangibles + PP&E)", "Amortización de deuda", "Intereses pagados", "Dividendos pagados", "Minoritarios y otros", "Aumento de caja (antes de FX)", "Total usos", "Cobertura del dividendo: FCF antes de deuda", "Cobertura tras amortización neta de deuda", "Compromisos contractuales de capex (Nota 11)", "Covenant DFN / patrimonio (≤ 2,0x; Nota 25)"],
      ["Operating cash flow", "New debt (bonds + loans)", "Interest received and other", "Total sources", "Capex (IP + intangibles + PP&E)", "Debt repayment", "Interest paid", "Dividends paid", "Minorities and other", "Cash increase (before FX)", "Total uses", "Dividend cover: FCF before debt", "Cover after net debt repayment", "Contractual capex commitments (Note 11)", "Covenant net debt / equity (≤ 2.0x; Note 25)"]]);
    const DO = T([["402,1", "458,3", "+56,1", "78%"], ["(92,3)", "(84,4)", "+7,9", "11%"], ["(25,9)", "(33,5)", "−7,6", "−11%"], ["(49,6)", "(45,1)", "+4,6", "6%"], ["(48,5)", "(37,5)", "+11,0", "15%"], ["0,5", "0,4", "−0,1", "0%"], ["186,3", "258,2", "+71,9", "100%"]]);
    const rDO = h([["Cobros por ventas y servicios", "Pagos a proveedores", "Pagos a empleados", "Otros pagos operativos", "Impuestos pagados", "Otras entradas", "Flujo operativo"],
      ["Collections from sales and services", "Payments to suppliers", "Payments to employees", "Other operating payments", "Taxes paid", "Other inflows", "Operating cash flow"]]);
    // ---------------------------------------------------------------- Anexo 11 (3): ratios, puente, cifras del texto
    const RT = T([["8,1%", "8,3%", "9,7%", dash], ["4,2%", "4,5%", "5,5%", dash], ["6,5%", "6,6%", "7,3%", dash], ["−0,7", "−0,6", "+0,1", dash], ["1,78", "1,18", "1,54", "2,50"], ["0,83", "0,52", "0,87", "1,71"],
      ["0,82", "0,85", "0,68", "0,64"], ["25,5%", "26,9%", "18,8%", "17,6%"], ["3,30", "3,78", "2,49", "2,36"], ["6,9", "6,8", "7,7", "8,4"], ["0,090", "0,092", "0,097", dash], ["12,8%", "16,3%", "15,6%", dash], ["65", "n.s.", "66", "50"]]);
    const rRT = h([["ROE, normalizado (utilidad distribuible)", "ROA, normalizado", "ROIC (NOPAT al 27%, capital a valor razonable)", "ROIC − WACC real (7,2%), pp", "Razón corriente (x)", "Caja / pasivos corrientes (x)", "Pasivos totales / patrimonio (x)", "LTV (DFN / propiedades de inversión)", "DFN / EBITDA (x)", "EBITDA / costos financieros (x)", "Rotación de activos (x)", "Capital de trabajo neto operativo / ingresos", "Días de cobro"],
      ["ROE, normalized (distributable profit)", "ROA, normalized", "ROIC (NOPAT at 27%, capital at fair value)", "ROIC − real WACC (7.2%), pp", "Current ratio (x)", "Cash / current liabilities (x)", "Total liabilities / equity (x)", "LTV (net debt / investment property)", "Net debt / EBITDA (x)", "EBITDA / financial costs (x)", "Asset turnover (x)", "Net operating working capital / revenue", "Collection days"]]);
    const PM = T([["23,8%", "23,5%", "20,1%", "21,2%"], ["(0,5%)", "(0,4%)", "(0,6%)", "(0,9%)"], ["23,3%", "23,0%", "19,5%", "20,3%"], ["24.006 / 19.241", "", "", ""]]);
    const rPM = h([["Costos operativos (costo de ventas + GAV) / ingresos", "(−) D&A incluida en esos costos / ingresos", "= 1 − margen EBITDA", "Ingreso / EBITDA mensual por m² (CLP, UDM jun-26)"],
      ["Operating costs (cost of sales + G&A) / revenue", "(−) D&A included in those costs / revenue", "= 1 − EBITDA margin", "Monthly revenue / EBITDA per m² (CLP, LTM Jun-26)"]]);
    const PU = T([["1.438,4", "380,6", "3,78x", dash], ["1.438,4", "475,5", "3,03x", "−0,75x"], ["1.438,4", "544,9", "2,64x", "−0,39x"], ["1.285,7", "544,9", "2,36x", "−0,28x"]]);
    const rPU = h([["dic-24 reportado (Perú consolida un mes)", "Perímetro: Perú 12 meses en vez de 1 (UDM jun-26: 104,3 − 9,5)", "Crecimiento del EBITDA del resto (a jun-26)", "Menor DFN = jun-26 reportado"],
      ["Dec-24 reported (Peru consolidates one month)", "Perimeter: Peru 12 months instead of 1 (LTM Jun-26: 104.3 − 9.5)", "EBITDA growth of the rest (to Jun-26)", "Lower net debt = Jun-26 reported"]]);
    const CT = T([["29,0%", ""], ["19,1%", ""], ["66%", ""], ["21,8 pp", ""], ["10,7%", ""], ["+0,3%", ""], ["−1,4%", ""], ["10,0%", ""]]);
    const fuentesCT = h([["EEFF dic-25, Nota\u00a025", "Nota 11: puntos medios dic-24 / dic-25", "Cálculo", "EEFF dic-25, Nota\u00a025", "EEFF dic-25, Nota\u00a025", "Bloomberg EE/ANR, 17-oct-26", "Bloomberg EE, 17-oct-26", "Anexo 14"],
      ["Dec-25 financials, Note\u00a025", "Note 11: Dec-24 / Dec-25 midpoints", "Calculation", "Dec-25 financials, Note\u00a025", "Dec-25 financials, Note\u00a025", "Bloomberg EE/ANR, 17-Oct-26", "Bloomberg EE, 17-Oct-26", "Appendix 14"]]);
    const rCT = h([["Alza de la PI de Chile por valor razonable (1.141,2 / PI Chile dic-24 3.933,0)", "Efecto de la tasa: (6,62% − 0,5%) / (5,64% − 0,5%) − 1, perpetuidad", "Participación de la tasa en el alza («dos tercios»)", "Crecimiento 2025 por perímetro: (121,3 − 13,6) / 494,6", "Crecimiento 2025 sin Perú: 532,6 / 481,0 − 1", "Ingresos 2026E (704,0) vs consenso BEst (701,7; 13 analistas)", "EBITDA 2026E (561,3) vs consenso BEst (569,1)", "ROE 2025 con el mismo método que los pares"],
      ["Rise in Chile investment property from fair value (1,141.2 / Chile IP Dec-24 3,933.0)", "Rate effect: (6.62% − 0.5%) / (5.64% − 0.5%) − 1, perpetuity", "Rate share of the rise (“two-thirds”)", "2025 growth from perimeter: (121.3 − 13.6) / 494.6", "2025 growth ex-Peru: 532.6 / 481.0 − 1", "2026E revenue (704.0) vs BEst consensus (701.7; 13 analysts)", "2026E EBITDA (561.3) vs BEst consensus (569.1)", "2025 ROE with the same method as peers"]]);
    // ---------------------------------------------------------------- Anexo 12 (4): EBITDA
    const EB = T([["365,8", "515,9", "538,0", "1.971,7", "1.735,4"], ["(85,4)", "(195,6)", "(162,5)", "(1.448,9)", "(1.194,5)"], ["2,3", "1,8", "5,2", "3,5", "4,1"], ["282,7", "322,2", "380,6", "526,2", "544,9"], ["0,0", "0,0", "0,0", "0,0", "0,0"]]);
    const rEB = h([["Resultado operacional + D&A", "(−) Otros ingresos op. (principalmente valor razonable)", "(+) Otros gastos op. (no recurrentes)", "EBITDA ajustado", "Control vs EBITDA de la compañía"],
      ["Operating profit + D&A", "(−) Other op. income (mainly fair value)", "(+) Other op. expenses (non-recurring)", "Adjusted EBITDA", "Check vs company EBITDA"]]);
    // ---------------------------------------------------------------- Anexo 13 (5): proyección del caso base (modelo v4)
    const a5col = [
      fila(h(["Ingresos — Chile", "Revenue — Chile"]), s.ing_cl), fila(h(["Ingresos — Perú", "Revenue — Peru"]), s.ing_pe), fila(h(["Ingresos — Colombia", "Revenue — Colombia"]), s.ing_co),
      fila(h(["Eliminaciones", "Eliminations"]), elim.map((x) => (Math.abs(x) < 50 ? null : x))),
      fila(h(["Ingresos — expansión continua", "Revenue — continuous expansion"]), [0, 0, 0, ...s.ing_exp.slice(3, 6)].map((x) => (x ? x : null))),
      fila(h(["Ingresos", "Revenue"]), ingTot), [h(["Crecimiento", "Growth"]), ...s.crec.map((x) => f.pc(x, 1))],
      fila("EBITDA", ebitdaTot), [h(["Margen EBITDA", "EBITDA margin"]), ...ebitdaTot.map((x, i) => f.pc(x / ingTot[i], 1))], fila("D&A", da.map((x) => -x)), fila("EBIT", ebit),
      [h(["Tasa de impuesto en caja", "Cash tax rate"]), ...s.tasa.slice(0, 6).map((x) => f.pc(x, 1))],
      [h(["Capex 1S26 real (sin apertura)", "Capex 1H26 actual (no breakdown)"]), num(-84.2), dash, dash, dash, dash, dash],
      fila(h(["Capex — mantención + TI", "Capex — maintenance + IT"]), s.mant.map((x) => -x)),
      fila(h(["Capex incremental — pipeline", "Incremental capex — pipeline"]), s.capex_proy.map((x) => (x ? -x : null))),
      fila(h(["Capex incremental — expansión continua", "Incremental capex — continuous expansion"]), s.capex_exp.slice(0, 6).map((x) => (x ? x : null))),
      fila(h(["Capex total", "Total capex"]), capexTot.map((x) => -x)), fila(h(["Δ capital de trabajo", "Δ working capital"]), s.ct.slice(0, 6)),
      fila("FFO", s.ffo), fila(h(["Utilidad distribuible", "Distributable profit"]), s.distribuible_cal),
      fila(h(["Dividendos ordinarios (40%)", "Ordinary dividends (40%)"]), s.div_ord_cal.map((x) => -x)), fila(h(["Distribuciones adicionales", "Additional distributions"]), s.dist_adic_cal.map((x) => -x)),
      [h(["DFN / EBITDA (x)", "Net debt / EBITDA (x)"]), ...s.dfn_ebitda.map((x) => f.n(x, 2))]];
    const a5 = { enc: [h(["CLP mil MM", "CLP bn"]), "2026E", "2027E", "2028E", "2029E", "2030E", "2031E"], filas: a5col,
      nota: es ? "2026E = 1S26 real + 2S26E: D&A 3,0 + 2,6; el capex de 1S26 (84,2) no tiene apertura por tipo en el release; Δ capital de trabajo y distribuciones adicionales cubren solo 2S26E. FFO = (EBITDA − costo financiero neto) × (1 − impuesto en caja), antes de minoritarios (~1%); utilidad distribuible = FFO − reajuste UF (aproximación del modelo para proyectar; 2025A es la cifra auditada de CLP 369 mil MM, Nota 20e). Fuente: análisis del equipo."
        : "2026E = 1H26 actual + 2H26E: D&A 3.0 + 2.6; 1H26 capex (84.2) has no breakdown by type in the release; Δ working capital and additional distributions cover 2H26E only. FFO = (EBITDA − net financial cost) × (1 − cash tax), before minorities (~1%); distributable profit = FFO − UF indexation (model proxy for projections; 2025A is the audited CLP 369bn, Note 20e). Source: Team analysis." };
    // ---------------------------------------------------------------- Anexo 14 (6): pares
    const PR = T([["27,0%", "2,0%", "2,3%"], ["7,6%", "7,5%", "7,7%"], ["37,5%", "9,8%", "10,5%"], ["10,0%", "7,8%", "8,0%"], ["9,7%", dash, dash], ["18,8%", "42,9%", "16,3%"], ["18,8%", "40,6%", "16,3%"], ["2,49", "5,68", "2,11"], ["1,69", "2,69", "1,51"]]);
    const rPR = h([["Revalorización / propiedades iniciales", "EBITDA / propiedades de inversión (consolidado)", "ROE, reportado", "ROE normalizado — mismo método para los tres", "ROE normalizado — Mallplaza con utilidad distribuible", "LTV (deuda neta / PI)", "LTV con inversiones en JV/asociadas en el denominador", "DFN / EBITDA (x)", "Activos / patrimonio atribuible, cierre (x)"],
      ["Revaluation / opening property", "EBITDA / investment property (consolidated)", "ROE, reported", "ROE normalized — same method for all three", "ROE normalized — Mallplaza on distributable profit", "LTV (net debt / IP)", "LTV with JV/associate investments in the denominator", "Net debt / EBITDA (x)", "Assets / attributable equity, closing (x)"]]);

    // ---------------------------------------------------------------- composición
    const B = (n, t) => H.caption(n, t, 100);
    const tb = (enc, an, filas, o) => H.tabla(enc, an, filas, o);
    const A = (x, i) => x[i];
    const w5 = [2325, 620, 620, 620, 620, 420], wc = [2125, 775, 775, 775, 775];
    const f1 = a1[3];
    const bloque1 = [B(a1[0], a1[1]), tb(a1[2], [1050, 2350, 1375, 450], f1, { izq: 4, tam: 11 }), H.fuente(nota1)];
    const bloque2 = [B(h(["Anexo 2:", "Appendix 2:"]), h(["Estados financieros resumidos (CLP mil MM)", "Summary financial statements (CLP bn)"])),
      tb(a2.eIS, [2125, 620, 620, 620, 620, 620], a2.rIS.map((r, i) => [r, ...IS[i]]), { destacar: [7], sub: [1] }),
      tb(a2.eBS, [2125, 615, 615, 615, 615, 540], a2.rBS.map((r, i) => [r, ...BS[i]]), { destacar: [4], sub: [0, 2, 5, 6] }),
      tb(a2.eCF, [2125, 775, 775, 775, 775], a2.rCF.map((r, i) => [r, ...CF[i]]), {}),
      H.fuente(a2.nota),
      B(h(["Anexo 2D:", "Appendix 2D:"]), h(["Fuentes y usos de caja y variación del flujo operativo (CLP mil MM)", "Sources and uses of cash and change in operating cash flow (CLP bn)"])),
      tb(h([["", "2025", "1S26"], ["", "2025", "1H26"]]), [3725, 750, 750], rSU.map((r, i) => [r, ...SU[i]]), { destacar: [3, 10], sub: [] }),
      tb(h([["Δ flujo operativo", "1S25", "1S26", "Δ", "%Δ"], ["Δ operating cash flow", "1H25", "1H26", "Δ", "%Δ"]]), [2125, 620, 620, 620, 620], rDO.map((r, i) => [r, ...DO[i]]), { destacar: [6] }),
      H.fuente(es ? "El alza de 1S26 viene de cobros (en parte saldos de diciembre: días de cobro de 66 a 50) e impuestos pagados menores; no se extrapola. FCF = flujo operativo − capex − intereses netos (2025: 68,2 pagados − 10,8 recibidos = 57,4, estos últimos dentro de «Intereses recibidos y otros»; 1S26: 30,1 − 8,5 = 21,6). Fuente: EEFF dic-25 y jun-26, estado de flujos directo; Notas 11 y 25."
        : "The 1H26 rise comes from collections (partly December balances: receivable days from 66 to 50) and lower taxes paid; it is not extrapolated. FCF = operating cash flow − capex − net interest (2025: 68.2 paid − 10.8 received = 57.4, the latter within “Interest received and other”; 1H26: 30.1 − 8.5 = 21.6). Source: Dec-25 and Jun-26 financials, direct cash-flow statement; Notes 11 and 25.")];
    const bloque3 = [B(h(["Anexo 3:", "Appendix 3:"]), h(["Tablero de ratios", "Ratio dashboard"])),
      tb(["", ...h([["2023", "2024", "2025", "jun-26"], ["2023", "2024", "2025", "Jun-26"]])], wc, rRT.map((r, i) => [r, ...RT[i]]), { destacar: [3] }),
      tb([h(["Puente de margen (punto 2)", "Margin bridge (point 2)"]), "2023", "2024", "2025", h(["1S26", "1H26"])], wc, rPM.map((r, i) => [r, ...PM[i]]), { destacar: [2] }),
      H.fuente(es ? "El ROIC es un rendimiento en caja sobre capital a valor razonable (sin la revalorización de las propiedades): se compara con la WACC real, no con la nominal. Los costos operativos incluyen la D&A; sin ella son exactamente 1 − margen EBITDA. Apalancamiento jun-26 sobre EBITDA UDM; días de cobro dic-24 n.s. (Perú consolida un mes). Fuente: EEFF (CMF); Análisis Razonado 1S26; Release 2Q26 (p. 8; m²); análisis del equipo."
        : "ROIC is a cash return on capital at fair value (excluding property revaluation): it is compared with the real WACC, not the nominal one. Operating costs include D&A; without it they are exactly 1 − EBITDA margin. Jun-26 leverage on LTM EBITDA; Dec-24 collection days n.m. (Peru consolidates one month). Source: Financial statements (CMF); MD&A 1H26; 2Q26 Release (p. 8; m²); Team analysis."),
      B(h(["Anexo 3B:", "Appendix 3B:"]), h(["Puente DFN/EBITDA dic-24 → jun-26 (CLP mil MM)", "Net debt/EBITDA bridge Dec-24 → Jun-26 (CLP bn)"])),
      tb(["", es ? "DFN" : "Net debt", "EBITDA", "x", "Δ"], [2325, 700, 700, 700, 700], rPU.map((r, i) => [r, ...PU[i]]), { destacar: [3] }),
      H.fuente(es ? "Perú = ingresos − costo de ventas − GAV del segmento (EEFF dic-24 Nota 31; dic-25 Nota 25; jun-26 Nota 24). Usar el UDM jun-26 como proxy del año completo hace de 3,03x una cota inferior del pro forma de 2024."
        : "Peru = revenue − cost of sales − segment G&A (Dec-24 Note 31; Dec-25 Note 25; Jun-26 Note 24). Using LTM Jun-26 as a proxy for the full year makes 3.03x a lower bound of the 2024 pro forma."),
      B(h(["Anexo 3C:", "Appendix 3C:"]), h(["Cifras del texto: cálculo y fuente", "Figures in the text: calculation and source"])),
      tb(h([["Cifra", "Valor", "Fuente"], ["Figure", "Value", "Source"]]), [3125, 650, 1450], rCT.map((r, i) => [r, CT[i][0], fuentesCT[i]]), { izq: 1 })];
    const bloque4 = [B(h(["Anexo 4:", "Appendix 4:"]), h(["El EBITDA que usamos es el de la compañía: la revalorización se elimina una sola vez (CLP mil MM)", "The EBITDA we use is the company's: revaluation is removed once (CLP bn)"])),
      tb(["", "2022A", "2023A", "2024A", "2025A", h(["UDM 2T26", "LTM 2Q26"])], [2125, 620, 620, 620, 620, 620], rEB.map((r, i) => [r, ...EB[i]]), { destacar: [3] }),
      H.fuente(es ? "El campo EBITDA de Bloomberg incluye la revalorización y no se usa. Fuente: EEFF (CMF), Notas 23.1–23.3; release 2T26; análisis del equipo." : "Bloomberg's EBITDA field includes revaluation and is not used. Source: Financial statements (CMF), Notes 23.1–23.3; 2Q26 release; Team analysis.")];
    const bloque5 = [B(h(["Anexo 5:", "Appendix 5:"]), h(["Proyección del caso base (CLP mil MM)", "Base-case projection (CLP bn)"])),
      tb(a5.enc, [1815, 565, 565, 565, 565, 565, 565], a5.filas, { destacar: [5, 15], sub: [7] }), H.fuente(a5.nota)];
    const bloque6 = [B(h(["Anexo 6:", "Appendix 6:"]), h(["Con el mismo ajuste, Mallplaza tiene el ROE normalizado más alto de sus pares locales (2025)", "With the same adjustment, Mallplaza has the highest normalized ROE among its local peers (2025)"])),
      tb(["", "Mallplaza", "Parque Arauco", "Cenco Malls"], [2325, 933, 933, 1034], rPR.map((r, i) => [r, ...PR[i]]), { destacar: [3] }),
      H.fuente(es ? "Normalizado = utilidad atribuible − revalorización × (1 − 27%) × % atribuible a la controladora; revalorización = resultado operacional − (EBITDA − D&A) en Mallplaza y Parque Arauco, y −otros ingresos op. netos en Cenco Malls. Límites: 27% uniforme (Perú 29,5%, Colombia 35%); EBITDA de Cenco Malls = ingresos × margen ajustado de Bloomberg; los tres operan en Chile, Perú y Colombia, con pesos distintos. Inversión de Parque Arauco en asociadas: CLP 209,2 mil MM (dic-25). Fuente: EEFF (CMF); Bloomberg FA (PARAUCO CI, CENCOMAL CI; 28-sep-26); análisis del equipo."
        : "Normalized = attributable profit − revaluation × (1 − 27%) × % attributable to the parent; revaluation = operating profit − (EBITDA − D&A) at Mallplaza and Parque Arauco, and −net other operating income at Cenco Malls. Limits: uniform 27% (Peru 29.5%, Colombia 35%); Cenco Malls EBITDA = revenue × Bloomberg adjusted margin; all three operate in Chile, Peru and Colombia, with different weights. Parque Arauco's investment in associates: CLP 209.2bn (Dec-25). Source: Financial statements (CMF); Bloomberg FA (PARAUCO CI, CENCOMAL CI; 28-Sep-26); Team analysis.")];
    return { bloque1, bloque2, bloque3, bloque4, bloque5, bloque6 };
  },
});
