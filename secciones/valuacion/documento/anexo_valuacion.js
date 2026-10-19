// Anexos de Valuación 15 a 18 (en el código: [B] DCF, [C] supuestos, [E] múltiplos, [F] escenarios), versión v2 del equipo con las
// secciones nuevas: 15D desglose del DCF y 15E cuadratura, 16D calendario de deuda, 18D sensibilidades de la revisión. Todas las
// cifras salen de graficos/figuras/datos_anexos.json (modelo v4 y réplica del DCF; python graficos/anexos_valuacion.py).
const cifras = require("../../comun/cifras");

module.exports = (H) => ({
  construir(lang) {
    const { titulo, lead, subtitulo, tabla, fuente, dosColumnas, fmt, DA, NB, caption, imagen, run } = H;
    const { K } = cifras(lang);
    const es = lang === "es", f = fmt(lang);
    const h = (e, i) => (es ? e : i);
    const pc = (x, d = 1) => f(x * 100, d) + "%";
    const mm = (x) => f(x, 0);
    const neg = (x, d = 0) => (x < 0 ? "(" + f(Math.abs(x), d) + ")" : f(x, d));
    const xm = (x) => (x === null || x === undefined ? NB : f(x, 1) + "x");
    const pb = (x) => `${x > 0 ? "+" : "−"}${f(Math.abs(x), 0)}`;
    const per = es ? ["2S26E", "2027E", "2028E", "2029E", "2030E", "2031E", "Terminal"] : ["2H26E", "2027E", "2028E", "2029E", "2030E", "2031E", "Terminal"];
    const der = (n, desde = 1) => Array.from({ length: n - desde }, (_, i) => i + desde);
    const aPer = [3166, 1070, 1070, 1070, 1070, 1070, 1070, 1080];
    const out = [];
    const B = DA.B, P = B.puente, V = B.vt, C = DA.C, S = C.sup, N = DA.N, Dg = N.desglose, F = DA.F, Q = F.param, ES = F.esc;
    const row = (nombre, arr, fn = mm, extra = []) => [nombre, ...arr.map((x) => (x === null || x === undefined ? NB : fn(x))), ...extra];

    // ================================================================= Anexo 15 [B]: DCF
    out.push(titulo(h("ANEXO [B]: FLUJO DE CAJA DESCONTADO", "APPENDIX [B]: DISCOUNTED CASH FLOW"), 0, true));
    out.push(lead(h(`FCFF en CLP nominales desde el 30-jun-26, con convención de mitad de período y la WACC de cada período (Anexo [Z]). El capex se separa en mantención y TI (5,5% de los ingresos) y capex incremental (pipeline en curso y expansión continua). El valor terminal usa el FCFF 2032E con g de ${pc(V.g)} nominal (0% real: el capex incremental terminal es cero porque el RONIC queda bajo la WACC real) y la WACC terminal de ${pc(V.wacc_t, 2)}, e incluye ${mm(V.m2_maduro)} m² de GLA que maduran después de 2031. Explica el ${pc(V.vt_ev, 0)} del EV. Las secciones D y E desglosan cada línea y prueban que cuadra con el DCF.`,
      `FCFF in nominal CLP from 30-Jun-26, mid-period convention and each period's WACC (Appendix [Z]). Capex is split into maintenance and IT (5.5% of revenue) and incremental capex (pipeline under way and continuous expansion). Terminal value uses 2032E FCFF with nominal g of ${pc(V.g)} (0% real: terminal incremental capex is zero because RONIC is below the real WACC) and a terminal WACC of ${pc(V.wacc_t, 2)}, and includes ${mm(V.m2_maduro)} m² of GLA maturing after 2031. It explains ${pc(V.vt_ev, 0)} of EV. Sections D and E break down each line and prove it ties to the DCF.`)));
    out.push(subtitulo(h("A. FLUJO DE CAJA LIBRE DE LA FIRMA POR PERÍODO (CLP MM)", "A. FIRM FREE CASH FLOW BY PERIOD (CLP MM)"), 80));
    const fl = B.filas, T7 = (k, fn = mm) => fl[k].map((x) => (x === null || x === undefined ? NB : fn(x)));
    const filasA = [
      row(h("Ingresos totales", "Total revenue"), fl.ing), row(h("EBITDA", "EBITDA"), fl.ebitda), [h("Margen EBITDA", "EBITDA margin"), ...T7("margen", pc)],
      [h("Tasa de impuesto en caja", "Cash tax rate"), ...T7("tasa", pc)], row("NOPAT", fl.nopat), row(h("(+) D&A", "(+) D&A"), fl.da),
      row(h("Capex de mantención + TI", "Maintenance + IT capex"), [...Dg.capex_mant.map((x) => -x), Dg.mant_term], neg),
      row(h("Capex incremental: pipeline en curso", "Incremental capex: pipeline"), [...Dg.capex_pipeline.map((x) => -x), 0], neg),
      row(h("Capex incremental: expansión continua", "Incremental capex: continuous expansion"), [...Dg.capex_exp, 0], neg),
      row(h("Inversión en capital de trabajo", "Working capital investment"), fl.ct, neg), row("FCFF", fl.fcff),
      [h("WACC nominal del período", "Nominal WACC of the period"), ...T7("wacc", (x) => pc(x, 2))],
      [h("Factor de descuento (mitad del período)", "Discount factor (mid-period)"), ...fl.fd.slice(0, 6).map((x) => f(x, 4)), NB],
      row(h("VP del FCFF / del valor terminal", "PV of FCFF / terminal value"), [...fl.vp.slice(0, 6), P.vp_vt])];
    // el capex mantención del terminal es 5,5% de los ingresos del terminal (negativo en el DCF)
    out.push(tabla(["", ...per], aPer, filasA, { derecha: der(8), destacar: [10] }));
    const filasP = [[P.vp_exp, P.ps_exp], [P.vp_vt, P.ps_vt], [P.ev, P.ev / P.acciones], [P.dfn, P.dfn / P.acciones], [P.nci, P.nci / P.acciones], [P.patrimonio, P.valor]]
      .map(([a, b], i) => [h(["VP del FCFF explícito (2S26E–2031E)", "VP del valor terminal", "Valor de empresa (EV)", "(−) Deuda financiera neta 30-jun-26", "(−) Minoritarios (libro)", "Valor del patrimonio / por acción"],
        ["PV of explicit FCFF (2H26E–2031E)", "PV of terminal value", "Enterprise value (EV)", "(−) Net financial debt 30-Jun-26", "(−) Minorities (book)", "Equity value / per share"])[i], neg(a), neg(b)]);
    const sens = B.sens.v.map((r, i) => [`${pb(B.sens.dw[i] * 1e4)} ${es ? "pb" : "bp"}`, ...r.map((x) => mm(x))]);
    out.push(dosColumnas(
      [subtitulo(h("B. DEL VALOR DE EMPRESA AL VALOR POR ACCIÓN", "B. FROM ENTERPRISE VALUE TO VALUE PER SHARE"), 100),
        tabla(h(["Concepto", "CLP MM", "CLP/acción"], ["Item", "CLP MM", "CLP/share"]), [3025, 1200, 1000], filasP, { derecha: [1, 2], destacar: [5] })],
      [subtitulo(h("C. SENSIBILIDAD DEL VALOR DCF (CLP/ACCIÓN)", "C. DCF VALUE SENSITIVITY (CLP/SHARE)"), 100),
        tabla([h("Δ WACC \\ g", "Δ WACC \\ g"), ...B.sens.g.map((g) => "g " + pc(g))], [1225, 1000, 1000, 1000, 1000], sens, { derecha: [1, 2, 3, 4], destacar: [2] })]));
    out.push(fuente(h(`Valor terminal al 31-dic-31: CLP ${mm(V.vt)} MM (cap rate implícito del VT ${pc(V.cap_vt)}; EV/EBITDA UDM implícito del EV ${f(V.ev_ebitda, 1)}x). La deuda neta de Bloomberg (CLP 1.311.569 MM, usada en múltiplos) difiere de la DFN en los pasivos por arrendamiento NIIF 16 (CLP 25.828 MM a jun-26). Fuente: análisis del equipo; ${mm(P.acciones)} MM de acciones.`,
      `Terminal value at 31-Dec-31: CLP ${mm(V.vt)} MM (implied terminal cap rate ${pc(V.cap_vt)}; implied LTM EV/EBITDA of EV ${f(V.ev_ebitda, 1)}x). Bloomberg net debt (CLP 1,311,569 MM, used in multiples) differs from net financial debt by IFRS 16 lease liabilities (CLP 25,828 MM at Jun-26). Source: Team analysis; ${mm(P.acciones)} MM shares.`)));
    // 15D desglose
    out.push(subtitulo(h("D. DESGLOSE DE LAS LÍNEAS DEL DCF (CLP MM SALVO INDICACIÓN)", "D. BREAKDOWN OF THE DCF LINES (CLP MM UNLESS STATED)"), 100));
    const ps = ["cl", "pe", "co"];
    const triple = (i) => ps.map((p) => f(Dg.ing_pais[p][i] / 1000, 1)).join(" / ");
    const trip2 = (i) => ps.map((p) => f(Dg.ebitda_pais[p][i] / 1000, 1)).join(" / ");
    const calFilas = (fn) => [NB, ...[0, 1, 2, 3, 4, 5].map(fn), NB];
    const gla = (x) => mm(x);
    const filasD = [
      [h("Ingresos: base + expansión (D1)", "Revenue: base + expansion (D1)"), ...Array(7).fill(NB)],
      row(h("Ingresos base (motor operativo)", "Base revenue (operating engine)"), [...Dg.ing_base, null]),
      row(h("(+) Expansión continua (cohortes desde 2029)", "(+) Continuous expansion (cohorts from 2029)"), [...Dg.ing_exp, null]),
      [h("Chile / Perú / Colombia, año calendario (CLP mil MM)", "Chile / Peru / Colombia, calendar year (CLP bn)"), ...[0, 1, 2, 3, 4, 5].map(triple), NB],
      [h(`Puente al terminal: + maduración ${mm(Dg.madura_m2)} m²; × (1 + ${pc(V.g)})`, `Bridge to terminal: + maturation of ${mm(Dg.madura_m2)} m²; × (1 + ${pc(V.g)})`), NB, NB, NB, NB, NB, mm(Dg.madura_ing), mm(Dg.indexa_ing)],
      [h("EBITDA y D&A (D2–D3)", "EBITDA and D&A (D2–D3)"), ...Array(7).fill(NB)],
      row(h("EBITDA base", "Base EBITDA"), [...Dg.ebitda_base, null]),
      row(h(`(+) EBITDA de la expansión (margen ${pc(Dg.margen_exp)})`, `(+) Expansion EBITDA (margin ${pc(Dg.margen_exp)})`), [...Dg.ebitda_exp, null]),
      [h("EBITDA Chile / Perú / Colombia, año calendario (80/82/74%)", "EBITDA Chile / Peru / Colombia, calendar year (80/82/74%)"), ...[0, 1, 2, 3, 4, 5].map(trip2), NB],
      row(h(`D&A = ${pc(Dg.da_pct, 3)} × ingresos totales`, `D&A = ${pc(Dg.da_pct, 3)} × total revenue`), Dg.da),
      [h("Impuesto, capex y capital de trabajo (D4–D6)", "Tax, capex and working capital (D4–D6)"), ...Array(7).fill(NB)],
      row("EBIT", Dg.ebit), row(h("(−) Impuesto = EBIT × tasa", "(−) Tax = EBIT × rate"), Dg.imp),
      [h("Capex incremental (a): pipeline, USD MM × USDCLP", "Incremental capex (a): pipeline, USD MM × USDCLP"), ...[0, 1, 2].map((i) => `${f(Dg.capex_usd[i], 1)} × ${f(Dg.usdclp[i], 0)}`), NB, NB, NB, NB],
      [h("GLA nuevo que abre en el año (m²; pipeline y expansión)", "New GLA opening in the year (m²; pipeline and expansion)"), ...Dg.gla_nuevo.map(gla), NB],
      [h("Capex total / ingresos", "Total capex / revenue"), ...Dg.capex_ing.map((x) => pc(x))],
      [h("Capex incremental / ingresos", "Incremental capex / revenue"), ...Dg.capex_inc_ing.map((x) => pc(x))],
      [h(`Δ ingresos para capital de trabajo (× ${pc(Dg.ct_pct)})`, `Δ revenue for working capital (× ${pc(Dg.ct_pct)})`), ...Dg.dIng_ct.map((x) => mm(x))]];
    out.push(tabla(["", ...per], aPer, filasD, { derecha: der(8), destacar: [], sub: [0, 5, 10] }));
    out.push(fuente(h(`Ingresos por país del Anexo 13 (año calendario; precisión 0,1 mil MM). Capex 1S26 sin apertura por tipo: el release solo da la apertura UDM (USD 132,5 MM: mantención 40,0; TI y otros 10,2; brownfield 52,5; reconversión 29,1; greenfield 0,7). Yield on cost ${pc(K.yield_costo)}; RONIC real ${pc(K.ronic)} vs WACC real ${pc(K.wacc_t_real)}.`,
      `Country revenue from Appendix 13 (calendar year; 0.1 bn precision). 1H26 capex has no breakdown by type: the release only gives the LTM breakdown (USD 132.5m: maintenance 40.0; IT and other 10.2; brownfield 52.5; reconversion 29.1; greenfield 0.7). Yield on cost ${pc(K.yield_costo)}; real RONIC ${pc(K.ronic)} vs real WACC ${pc(K.wacc_t_real)}.`)));
    // 15E cuadratura
    const cu = N.cuadratura;
    const nombresCu = h(["Ingresos totales", "EBITDA total", "D&A", "Impuesto operacional", "NOPAT", "Capex total (mantención + incremental)", "Inversión en capital de trabajo", "FCFF", "VP del FCFF", "Valor terminal / VP del VT / EV", "Valor por acción (CLP)"],
      ["Total revenue", "Total EBITDA", "D&A", "Operating tax", "NOPAT", "Total capex (maintenance + incremental)", "Working capital investment", "FCFF", "PV of FCFF", "Terminal value / PV of TV / EV", "Value per share (CLP)"]);
    const clavesCu = ["Ingresos totales", "EBITDA total", "D&A", "Impuesto operacional", "NOPAT", "Capex total (mantención + incremental)", "Inversión en capital de trabajo", "FCFF", "VP del FCFF", "Valor terminal / VP del VT / EV", "Valor por acción (CLP)"];
    const maxCu = Math.max(...clavesCu.map((k) => cu[k]));
    out.push(dosColumnas(
      [subtitulo(h("E. CONTROL DE CUADRATURA (CLP MM)", "E. TIE-OUT CHECK (CLP MM)"), 100),
        tabla(h(["Línea: desglose − DCF", "Máx. |dif.| 2S26E–Terminal"], ["Line: breakdown − DCF", "Max. |diff.| 2H26E–Terminal"]), [3925, 1300], nombresCu.map((n, i) => [n, cu[clavesCu[i]] < 1e-6 ? "0" : f(cu[clavesCu[i]], 3)]), { derecha: [1] })],
      [subtitulo(h("LECTURA", "READING"), 100),
        lead(h(`Máxima diferencia absoluta en todas las líneas y períodos: ${maxCu < 1e-6 ? "0" : f(maxCu, 3)}. El desglose solo reordena el DCF: EV CLP ${mm(P.ev)} MM, valor DCF CLP ${mm(P.valor)} y precio objetivo CLP ${mm(F.esc.base.po)}. El capex incremental (pipeline + expansión) pesa ${pc(Dg.capex_inc_ing[1], 0)}–${pc(Dg.capex_inc_ing[0], 0)} de los ingresos en 2S26E–2027E y ~${pc(Dg.capex_inc_ing[4], 0)} desde 2029E; la mantención, 5,5%. El terminal solo financia mantención: crecer con un RONIC real de ${pc(K.ronic)} bajo una WACC real de ${pc(K.wacc_t_real)} restaría valor (Figura 30).`,
          `Maximum absolute difference across all lines and periods: ${maxCu < 1e-6 ? "0" : f(maxCu, 3)}. The breakdown only reorders the DCF: EV CLP ${mm(P.ev)} MM, DCF value CLP ${mm(P.valor)} and target price CLP ${mm(F.esc.base.po)}. Incremental capex (pipeline + expansion) weighs ${pc(Dg.capex_inc_ing[1], 0)}–${pc(Dg.capex_inc_ing[0], 0)} of revenue in 2H26E–2027E and ~${pc(Dg.capex_inc_ing[4], 0)} from 2029E; maintenance, 5.5%. The terminal only funds maintenance: growing at a real RONIC of ${pc(K.ronic)} under a real WACC of ${pc(K.wacc_t_real)} would subtract value (Figure 30).`)),
        lead(h(`El valor terminal parte de los ingresos 2031E (${mm(DA.B.filas.ing[5])}), suma la maduración de la expansión ya financiada (+${mm(Dg.madura_ing - Dg.ing_exp[5])}; el ingreso maduro bruto, ${mm(Dg.madura_ing)}, incluye ${mm(Dg.ing_exp[5])} ya presentes en 2031E) e indexa por inflación (+${mm(Dg.indexa_ing)}) hasta ${mm(Dg.ing_term)}; el FCFF 2032E = NOPAT ${mm(B.filas.nopat[6])} + D&A ${mm(B.filas.da[6])} − mantención ${mm(-Dg.mant_term)} − capital de trabajo ${mm(-B.filas.ct[6])} = ${mm(B.filas.fcff[6])}.`,
          `Terminal value starts from 2031E revenue (${mm(DA.B.filas.ing[5])}), adds the maturation of the expansion already funded (+${mm(Dg.madura_ing - Dg.ing_exp[5])}; the gross mature revenue, ${mm(Dg.madura_ing)}, includes ${mm(Dg.ing_exp[5])} already in 2031E) and indexes for inflation (+${mm(Dg.indexa_ing)}) to ${mm(Dg.ing_term)}; 2032E FCFF = NOPAT ${mm(B.filas.nopat[6])} + D&A ${mm(B.filas.da[6])} − maintenance ${mm(-Dg.mant_term)} − working capital ${mm(-B.filas.ct[6])} = ${mm(B.filas.fcff[6])}.`))]));

    // ================================================================= Anexo 16 [C]: supuestos de la proyección
    out.push(titulo(h("ANEXO [C]: SUPUESTOS DE LA PROYECCIÓN", "APPENDIX [C]: PROJECTION ASSUMPTIONS"), 200));
    out.push(lead(h(`El forecast base sigue el motor operativo del equipo con el pipeline completo con fecha (${mm(S.pipe_m2)} m², pipeline anunciado hasta 2028, incluidos 46.500 m² de otros proyectos en Chile); desde ${S.primer} se suma una expansión continua al ritmo histórico de apertura. Los ingresos 2026E son año calendario; el período 2S26E descuenta el 1S26 real (CLP ${mm(S.ing_1s26)} MM de ingresos).`,
      `The base forecast follows the team's operating engine with the full dated pipeline (${mm(S.pipe_m2)} m², announced pipeline through 2028, including 46,500 m² of other projects in Chile); from ${S.primer} a continuous expansion at the historical opening pace is added. 2026E revenue is calendar-year; the 2H26E period nets out the 1H26 actual (CLP ${mm(S.ing_1s26)} MM of revenue).`)));
    out.push(subtitulo(h("A. PROYECCIÓN DE INGRESOS, CAPEX Y CAPITAL DE TRABAJO (CLP MM)", "A. REVENUE, CAPEX AND WORKING CAPITAL PROJECTION (CLP MM)"), 80));
    const cf = C.filas, per6 = per.slice(0, 6);
    const capexIncr = cf.capex_proy.map((x, i) => Math.round(x) + Math.round(cf.capex_exp[i]));   // capex incremental total = pipeline en curso + expansión continua (suma de las cifras mostradas)
    const filasC = [
      row(h("Ingresos base (año calendario; la columna 2S26E muestra 2026E completo)", "Base revenue (calendar year; the 2H26E column shows full-year 2026E)"), cf.ing_cal), [h("Crecimiento de ingresos totales (año calendario)", "Total revenue growth (calendar year)"), ...cf.crec.map((x) => pc(x))],
      row(h("Ingresos de la expansión continua", "Continuous expansion revenue"), cf.ing_exp), [h("Margen EBITDA base", "Base EBITDA margin"), ...cf.margen.map((x) => pc(x))],
      [h("Tasa de impuesto en caja", "Cash tax rate"), ...cf.tasa.map((x) => pc(x))], row(h("Capex de mantención + TI (5,5% de ingresos)", "Maintenance + IT capex (5.5% of revenue)"), cf.mant),
      [h("Capex incremental: proyectos en curso (USD MM)", "Incremental capex: projects under way (USD MM)"), ...cf.capex_usd.map((x) => f(x, 1))], [h("USDCLP promedio", "Average USDCLP"), ...cf.usdclp.map((x) => f(x, 0))],
      row(h("Capex incremental: proyectos en curso", "Incremental capex: projects under way"), cf.capex_proy), row(h("Capex incremental: expansión continua", "Incremental capex: continuous expansion"), cf.capex_exp),
      row(h("Capex incremental total", "Total incremental capex"), capexIncr), row(h("Capex total", "Total capex"), cf.capex_base.map((x, i) => x + cf.capex_exp[i])),
      [h("GLA nuevo que abre en el año (m²)", "New GLA opening in the year (m²)"), ...Dg.gla_nuevo.map(gla)], [h("Costo por m² de expansión (CLP MM, IPC)", "Expansion cost per m² (CLP MM, CPI)"), ...cf.costo_m2.map((x) => f(x, 2))],
      row(h("Inversión en capital de trabajo", "Working capital investment"), cf.ct)];
    out.push(tabla(["", ...per6], [3466, 1200, 1200, 1200, 1200, 1200, 1200], filasC, { derecha: der(7), destacar: [11] }));
    const rampa = S.rampa.map((x) => pc(x, 0)).join(" / ");
    const filasExp = [[h(`GLA nuevo por año desde ${S.primer}`, `New GLA per year from ${S.primer}`), `${mm(S.gla_anual)} m²`], [h("Base: promedio orgánico 2022, 2023 y 2025", "Base: organic average 2022, 2023 and 2025"), `${mm(S.gla_22)} / ${mm(S.gla_23)} / ${mm(S.gla_25)} m²`],
      [h("Pipeline en curso (Release 2Q26; Memoria 2025)", "Pipeline under way (2Q26 Release; 2025 Annual Report)"), `USD ${f(S.pipe_usd, 1)} MM / ${mm(S.pipe_m2)} m²`], [h("Costo por m² (= pipeline)", "Cost per m² (= pipeline)"), `USD ${mm(S.costo_usd_m2)}`],
      [h("Desembolso del capex (desde 2027E)", "Capex disbursement (from 2027E)"), `${pc(S.pago_t2, 0)} ${es ? "en" : "in"} t−2 / ${pc(1 - S.pago_t2, 0)} ${es ? "en" : "in"} t−1`], [h("Rampa de ingresos (año 1 / 2 / 3+)", "Revenue ramp (year 1 / 2 / 3+)"), rampa],
      [h("Mix Chile / Perú; margen EBITDA", "Chile / Peru mix; EBITDA margin"), `${pc(S.mix_cl)} / ${pc(1 - S.mix_cl)}; ${pc(S.margen_cl, 0)} / ${pc(S.margen_pe, 0)}`],
      [h("Yield on cost / RONIC real después de impuestos", "Yield on cost / real after-tax RONIC"), `${pc(S.yield)} / ${pc(S.ronic)}`]];
    const filasOp = [[h("Margen EBITDA UDM 2T26", "LTM 2Q26 EBITDA margin"), pc(S.ebitda_ltm / S.ing_ltm)], [h("Capex de mantención + TI", "Maintenance + IT capex"), h(`${pc(S.mant)} de los ingresos`, `${pc(S.mant)} of revenue`)],
      [h("Capex incremental terminal", "Terminal incremental capex"), h("0 (g real 0%)", "0 (real g 0%)")], [h("D&A / ingresos (UDM 2T26)", "D&A / revenue (LTM 2Q26)"), pc(S.da, 3)],
      [h("Impuesto en caja", "Cash tax"), `${pc(cf.tasa[0])} → ${pc(S.tax_lp, 0)} (2030)`], [h("Capital de trabajo neto / ingresos (dic-25)", "Net working capital / revenue (Dec-25)"), pc(S.ct_ing)],
      [h("Deudores / arriendos anticipados (dic-25, CLP MM)", "Receivables / prepaid rents (Dec-25, CLP MM)"), `${mm(S.deudores)} / ${mm(S.anticip)}`], [h("g terminal nominal (real)", "Terminal nominal g (real)"), `${pc(S.g)} (0,0%)`.replace(",0%)", es ? ",0%)" : ".0%)")],
      [h("DFN/EBITDA objetivo; Kd real por vuelta", "Target net debt/EBITDA; real Kd per turn"), `${f(S.dfn_obj, 1)}x ${h("en", "in")} ${S.anio_obj}; +${f(S.spread * 1e4, 0)} ${es ? "pb" : "bp"}`], [h("Payout sobre utilidad distribuible", "Payout on distributable profit"), pc(S.payout, 0)]];
    out.push(dosColumnas(
      [subtitulo(h("B. EXPANSIÓN CONTINUA", "B. CONTINUOUS EXPANSION"), 100), tabla(h(["Supuesto", "Valor"], ["Assumption", "Value"]), [2925, 2300], filasExp, { derecha: [1] })],
      [subtitulo(h("C. OPERACIÓN, CAPITAL Y VALOR TERMINAL", "C. OPERATIONS, CAPITAL AND TERMINAL VALUE"), 100), tabla(h(["Supuesto", "Valor"], ["Assumption", "Value"]), [2925, 2300], filasOp, { derecha: [1] })]));
    // 16D calendario de deuda
    const Dd = N.deuda;
    const filasDd = [
      row(h("DFN de apertura", "Opening net debt"), Dd.apertura), row(h("(−) FCFF", "(−) FCFF"), Dd.fcff.map((x) => -x), neg), row(h("(+) Costo financiero neto, después de impuestos", "(+) Net financial cost, after tax"), Dd.costo_at),
      row(h("(+) Reajuste UF, después de impuestos", "(+) UF indexation, after tax"), Dd.reajuste_at), row(h("(+) Dividendos ordinarios", "(+) Ordinary dividends"), Dd.div),
      row(h("(+) Distribución adicional (recompra / dividendo extra)", "(+) Additional distribution (buyback / special dividend)"), Dd.adic), row(h("DFN de cierre", "Closing net debt"), Dd.cierre),
      [h("DFN / EBITDA de cierre (x)", "Closing net debt / EBITDA (x)"), ...Dd.ratio.map((x) => f(x, 2) + "x")], row(h(`Deuda bruta de cierre (caja constante ${mm(Dd.caja)})`, `Closing gross debt (constant cash ${mm(Dd.caja)})`), Dd.bruta)];
    out.push(subtitulo(h("D. CALENDARIO DE LA DEUDA (CLP MM)", "D. DEBT SCHEDULE (CLP MM)"), 100));
    out.push(tabla(["", ...per6], [3466, 1200, 1200, 1200, 1200, 1200, 1200], filasDd, { derecha: der(7), destacar: [6] }));
    out.push(fuente(h(`DFN(t) = DFN(t−1) − FCFF + costo financiero y reajuste UF después de impuestos + dividendos ordinarios + distribución adicional, que lleva la DFN a la trayectoria objetivo (${f(Dd.ratio0, 2)}x → ${f(Dd.ratio_obj, 1)}x). Costo financiero: ${pc(Dd.costo0, 2)} de la DFN + ${f(Dd.spread * 1e4, 0)} pb por vuelta sobre ${f(Dd.ratio0, 2)}x; reajuste UF sobre ${pc(Dd.uf, 0)} de la deuda a ${pc(S.g)}. Vencimientos de bonos: Figura 6. Fuente: Release 2T26; EEFF dic-25; análisis del equipo.`,
      `Net debt(t) = net debt(t−1) − FCFF + financial cost and UF indexation after tax + ordinary dividends + additional distribution, which takes net debt to the target path (${f(Dd.ratio0, 2)}x → ${f(Dd.ratio_obj, 1)}x). Financial cost: ${pc(Dd.costo0, 2)} of net debt + ${f(Dd.spread * 1e4, 0)} bp per turn above ${f(Dd.ratio0, 2)}x; UF indexation on ${pc(Dd.uf, 0)} of debt at ${pc(S.g)}. Bond maturities: Figure 6. Source: 2Q26 Release; Dec-25 financials; Team analysis.`)));

    // ================================================================= Anexo 17 [E]: múltiplos
    const E = DA.E, PE = E.puente;
    out.push(titulo(h("ANEXO [E]: MÚLTIPLOS COMPARABLES", "APPENDIX [E]: COMPARABLE MULTIPLES"), 0, true));
    out.push(lead(h(`Usamos el mismo set de pares que la beta de la WACC. EV/EBITDA UDM es el método principal (neutral al apalancamiento y a la revalorización NIC 40); el múltiplo pondera ${pc(E.peso_cl, 0)} la mediana chilena y ${pc(1 - E.peso_cl, 0)} la del set común. P/E y P/BV mezclan pares a valor razonable y a costo: solo referencia. Precio al 30-sep-26; deuda neta Bloomberg y tipo de cambio al 30-jun-26.`,
      `We use the same peer set as the WACC beta. LTM EV/EBITDA is the main method (neutral to leverage and to IAS 40 revaluation); the multiple weights the Chilean median ${pc(E.peso_cl, 0)} and the common-set median ${pc(1 - E.peso_cl, 0)}. P/E and P/BV mix fair-value and cost-based peers: reference only. Price at 30-Sep-26; Bloomberg net debt and FX at 30-Jun-26.`)));
    out.push(subtitulo(h("A. PARES (CLP MM; MÚLTIPLOS EN VECES)", "A. PEERS (CLP MM; MULTIPLES IN TIMES)"), 80));
    const paisT = es ? { Chile: "Chile", Brasil: "Brasil", Argentina: "Argentina", "México": "México" } : { Chile: "Chile", Brasil: "Brazil", Argentina: "Argentina", "México": "Mexico" };
    const filasE = E.pares.map((p) => [(p.nombre === "Cencosud Shopping" ? "Cenco Malls" : p.nombre) + (p.set === 0 ? h(" (excluida)", " (excluded)") : ""), paisT[p.pais] || p.pais, mm(p.mcap), mm(p.ev), mm(p.ebitda), ...p.m.map(xm)]);
    h(["Mediana set común (6)", "Mediana Chile (2)", "Ponderado 70/30"], ["Common-set median (6)", "Chile median (2)", "Weighted 70/30"]).forEach((t, i) => filasE.push([t, NB, NB, NB, NB, ...E.est[["med_set", "med_cl", "pond"][i]].map(xm)]));
    out.push(tabla(h(["Empresa", "País", "Market cap", "EV", "EBITDA UDM", "EV/EBITDA UDM", "EV/EBITDA NTM", "EV/EBIT UDM", "P/E UDM", "P/BV"],
      ["Company", "Country", "Market cap", "EV", "LTM EBITDA", "LTM EV/EBITDA", "NTM EV/EBITDA", "LTM EV/EBIT", "LTM P/E", "P/BV"]), [2066, 900, 1150, 1150, 1100, 900, 900, 850, 800, 850], filasE,
      { derecha: der(10, 2), destacar: [0, filasE.length - 1] }));
    const nombresM = h(["EV/EBITDA UDM (principal)", "EV/EBITDA NTM", "EV/EBIT UDM", "P/E UDM (ref.)", "P/BV (ref.)"], ["LTM EV/EBITDA (main)", "NTM EV/EBITDA", "LTM EV/EBIT", "LTM P/E (ref.)", "P/BV (ref.)"]);
    const filasM = E.metodos.map(([mult, met, val], i) => [nombresM[i], xm(mult), mm(met), mm(val)]);
    const base = E.pond.find((r) => Math.abs(r[0] - E.peso_cl) < 1e-9);
    const filasW = E.pond.map(([w, mult, val, po, d]) => [pc(w, 0), xm(mult), mm(val), mm(po), (d > 0 ? "+" : "") + pc(d)]);
    out.push(dosColumnas(
      [subtitulo(h("B. VALOR POR MÚLTIPLO", "B. VALUE BY MULTIPLE"), 100), tabla(h(["Múltiplo", "Ponderado", "Métrica CLP MM", "CLP/acción"], ["Multiple", "Weighted", "Metric CLP MM", "CLP/share"]), [1925, 900, 1300, 1100], filasM, { derecha: [1, 2, 3], destacar: [0] })],
      [subtitulo(h("C. SENSIBILIDAD A LA PONDERACIÓN", "C. SENSITIVITY TO THE WEIGHTING"), 100), tabla(h(["Peso Chile", "Múltiplo", "Valor múlt.", "Precio obj.", "Δ"], ["Chile weight", "Multiple", "Mult. value", "Target price", "Δ"]), [1025, 900, 1100, 1100, 1100], filasW, { derecha: [0, 1, 2, 3, 4], destacar: [E.pond.indexOf(base)] })]));
    out.push(fuente(h(`Puente: ${f(PE.mult, 2)}x × EBITDA UDM CLP ${mm(PE.ebitda)} MM = EV ${mm(PE.ev)} − deuda neta Bloomberg ${mm(PE.deuda_bbg)} − minoritarios ${mm(PE.nci)} = CLP ${mm(PE.valor)} por acción (${mm(PE.acciones)} MM de acciones). Fibra Danhos queda fuera por su régimen FIBRA y por no tener beta en la WACC; IRSA reporta bajo NIC 29. Fuente: Bloomberg; análisis del equipo.`,
      `Bridge: ${f(PE.mult, 2)}x × LTM EBITDA CLP ${mm(PE.ebitda)} MM = EV ${mm(PE.ev)} − Bloomberg net debt ${mm(PE.deuda_bbg)} − minorities ${mm(PE.nci)} = CLP ${mm(PE.valor)} per share (${mm(PE.acciones)} MM shares). Fibra Danhos is excluded because of its FIBRA regime and because it has no beta in the WACC; IRSA reports under IAS 29. Source: Bloomberg; Team analysis.`)));

    // ================================================================= Anexo 18 [F]: escenarios y robustez
    const obj = (x) => F.peso * x + (1 - F.peso) * F.mult;
    const vs = (x) => pc(x / F.precio - 1);
    out.push(titulo(h("ANEXO [F]: ESCENARIOS Y ROBUSTEZ", "APPENDIX [F]: SCENARIOS AND ROBUSTNESS"), 200));
    out.push(lead(h(`Ningún escenario ni método alternativo alcanza el precio de CLP ${mm(F.precio)}. Los escenarios se calculan con la réplica del DCF (reproduce exactamente el valor base) y el precio objetivo mantiene el ${pc(F.peso, 0)} DCF y el valor por múltiplos de CLP ${mm(F.mult)}. Bear es todo lo que sale mal a la vez; bull, todo lo que sale bien; ambos con las mismas palancas del tornado de Riesgos.`,
      `No scenario or alternative method reaches the CLP ${mm(F.precio)} price. Scenarios are computed with the DCF replica (it reproduces the base value exactly) and the target price keeps the ${pc(F.peso, 0)} DCF weight and the multiples value of CLP ${mm(F.mult)}. Bear is everything going wrong at once; bull, everything going right; both use the same levers as the Risks tornado.`)));
    out.push(subtitulo(h("A. ESCENARIOS", "A. SCENARIOS"), 80));
    const Pp = Q, R = Pp.rangos, pr = F.prob;
    const wacc0 = DA.B.filas.wacc[0];
    const filasF = [
      [h("Bear", "Bear"), h(`Todo sale mal: WACC nominal +${f(Q.dw_bear * 1e4, 0)} pb en todos los períodos; margen EBITDA ${pc(Q.m_lo)}, ocupación ${pc(R["Ocupación"][1])}, arriendo mismas tiendas 2027–28 ${pc(R["Same-store rent growth (SSR)"][1])}, ventas de locatarios ${pc(R["Same-store sales growth (SSS)"][1])}, crecimiento real 2029–31 en el piso de la Nota 11, visitas ${pc(R["Visitas — crecimiento orgánico"][1])}, capex de mantención 6,6%`,
        `Everything goes wrong: nominal WACC +${f(Q.dw_bear * 1e4, 0)} bp in all periods; EBITDA margin ${pc(Q.m_lo)}, occupancy ${pc(R["Ocupación"][1])}, same-store rent 2027–28 ${pc(R["Same-store rent growth (SSR)"][1])}, tenant sales ${pc(R["Same-store sales growth (SSS)"][1])}, 2029–31 real growth at the Note 11 floor, visitors ${pc(R["Visitas — crecimiento orgánico"][1])}, maintenance capex 6.6%`),
        mm(ES.bear.dcf), mm(ES.bear.po), vs(ES.bear.po)],
      [h("Base", "Base"), h(`Central: modelo v4; WACC nominal ${pc(wacc0, 2)} → ${pc(Q.wacc_t, 2)}; g ${pc(DA.B.vt.g)}; margen ${pc(Q.margen26)}`, `Central: model v4; nominal WACC ${pc(wacc0, 2)} → ${pc(Q.wacc_t, 2)}; g ${pc(DA.B.vt.g)}; margin ${pc(Q.margen26)}`), mm(ES.base.dcf), mm(ES.base.po), vs(ES.base.po)],
      [h("Bull", "Bull"), h(`Todo sale bien: WACC nominal −${f(Q.dw_bear * 1e4, 0)} pb; margen ${pc(Q.m_hi)}, ocupación ${pc(R["Ocupación"][2])}, arriendo mismas tiendas ${pc(R["Same-store rent growth (SSR)"][2])}, ventas de locatarios ${pc(R["Same-store sales growth (SSS)"][2])}, crecimiento real en el techo de la Nota 11, visitas ${pc(R["Visitas — crecimiento orgánico"][2])}, capex de mantención 4,8%; impuesto en caja con techo de ${pc(Q.tax_bull, 0)}`,
        `Everything goes right: nominal WACC −${f(Q.dw_bear * 1e4, 0)} bp; margin ${pc(Q.m_hi)}, occupancy ${pc(R["Ocupación"][2])}, same-store rent ${pc(R["Same-store rent growth (SSR)"][2])}, tenant sales ${pc(R["Same-store sales growth (SSS)"][2])}, real growth at the Note 11 ceiling, visitors ${pc(R["Visitas — crecimiento orgánico"][2])}, maintenance capex 4.8%; cash tax capped at ${pc(Q.tax_bull, 0)}`),
        mm(ES.bull.dcf), mm(ES.bull.po), vs(ES.bull.po)],
      [h("Tasadores", "Appraisers"), h(`WACC nominal ${pc(Q.wacc_tas, 2)} (${pc(Q.wacc_tas_real, 2)} real, Nota 11) en todos los períodos`, `Nominal WACC ${pc(Q.wacc_tas, 2)} (${pc(Q.wacc_tas_real, 2)} real, Note 11) in all periods`), mm(ES.tas.dcf), mm(ES.tas.po), vs(ES.tas.po)]];
    out.push(tabla(h(["Escenario", "Supuestos (todos a la vez)", "DCF (CLP)", "Precio objetivo", "vs precio"], ["Scenario", "Assumptions (all at once)", "DCF (CLP)", "Target price", "vs price"]),
      [1100, 6266, 1100, 1200, 1000], filasF, { derecha: [2, 3, 4], destacar: [1] }));
    out.push(fuente(h("Bear y bull son escenarios extremos ilustrativos, sin probabilidad asignada: la probabilidad de los resultados se estima con la simulación de la sección siguiente (Figura M1). Fuente: análisis del equipo.",
      "Bear and bull are illustrative extreme scenarios with no probability assigned: the probability of outcomes is estimated with the simulation in the next section (Figure M1). Source: Team analysis.")));
    // ---- simulación de Monte Carlo (valuacion/graficos/montecarlo.py; datos_montecarlo.json vía cifras.json)
    const MC = K.mc, st = MC.stats, sp = MC.supuestos, rgo = sp.rangos;
    const mp = (x) => pc(x, 2), mc1 = (x) => pc(x, 1);
    const filasMC = [
      [h("Media", "Mean"), mm(st.media)], [h("Mediana", "Median"), mm(st.mediana)], [h("Desviación estándar", "Standard deviation"), mm(st.desv)],
      ["P5 / P25", `${mm(st.p5)} / ${mm(st.p25)}`], ["P75 / P95", `${mm(st.p75)} / ${mm(st.p95)}`],
      [h(`Prob. objetivo ≥ CLP ${mm(MC.precio)} (precio)`, `Prob. target ≥ CLP ${mm(MC.precio)} (price)`), mc1(MC.p_sobre_precio)],
      [h(`Prob. objetivo ≥ CLP ${mm(MC.consenso)} (consenso)`, `Prob. target ≥ CLP ${mm(MC.consenso)} (consensus)`), mc1(MC.p_sobre_consenso)],
      [h("Prob. objetivo < CLP " + mm(MC.po_base) + " (caso base)", "Prob. target < CLP " + mm(MC.po_base) + " (base case)"), mc1(MC.p_bajo_objetivo)]];
    const ctr = Object.entries(MC.contribucion).sort((a, b) => b[1] - a[1]).slice(0, 4);
    const nomC = { "WACC": "WACC", "Arriendo mismas tiendas 2027–28": h("Arriendo mismas tiendas", "Same-store rent"), "Margen EBITDA": h("Margen EBITDA", "EBITDA margin"), "g nominal": h("g nominal", "Nominal g") };
    const bw = (MEDIA) => MEDIA / 1440 / (3550 / 1440);
    out.push(subtitulo(h("MONTE CARLO: DISTRIBUCIÓN DEL PRECIO OBJETIVO", "MONTE CARLO SIMULATION: TARGET PRICE DISTRIBUTION"), 120));
    out.push(dosColumnas(
      [tabla(h(["Precio objetivo simulado (CLP)", "Valor"], ["Simulated target price (CLP)", "Value"]), [3400, 1500], filasMC, { derecha: [1] }),
        new H.Paragraph({ spacing: { before: 40, after: 20 }, children: [run(h(`Contribución a la varianza (correlación de rangos): ${ctr.map(([k, v]) => `${nomC[k] || k} ${mc1(v)}`).join("; ")}.`, `Contribution to variance (rank correlation): ${ctr.map(([k, v]) => `${nomC[k] || k} ${mc1(v)}`).join("; ")}.`), { size: 13 })] })],
      [caption(h("Figura M1:", "Figure M1:"), h(`Distribución del precio objetivo en ${f(MC.n, 0)} simulaciones`, `Target price distribution in ${f(MC.n, 0)} simulations`), 20),
        imagen(`${lang}_V16_montecarlo.png`, 3550 / 1440, 2.05), fuente(h(`${pc(MC.p_bajo_precio, 0)} de las simulaciones queda bajo el precio de CLP ${mm(MC.precio)}. Fuente: análisis del equipo.`, `${pc(MC.p_bajo_precio, 0)} of simulations stay below the CLP ${mm(MC.precio)} price. Source: Team analysis.`))]));
    const tr = (k, d = 1) => `${pc(rgo[k][1], d)}/${pc(rgo[k][0], d)}/${pc(rgo[k][2], d)}`;
    out.push(fuente(h(`Supuestos: ${f(MC.n, 0)} iteraciones, semilla ${MC.semilla}; distribuciones triangulares (mín/base/máx, extremos 4T23–2T26 del tornado de Riesgos): margen EBITDA ${tr("Margen EBITDA")}, ocupación ${tr("Ocupación")}, arriendo mismas tiendas 2027–28 ${tr("Same-store rent growth (SSR)")}, ventas de locatarios ${tr("Same-store sales growth (SSS)")} (correlación ${f(sp.rho_ssr_sss, 1)} con el arriendo, cópula gaussiana; el resto independiente), visitas ${tr("Visitas — crecimiento orgánico")}, capex de mantención + TI ${pc(sp.capex[0])}/${pc(sp.capex[1])}/${pc(sp.capex[2])}, g nominal ${pc(sp.g[0])}/${pc(sp.g[1])}/${pc(sp.g[2])}, crecimiento real 2029–31 entre piso y techo de la Nota 11. WACC: desplazamiento normal, media 0, desviación ${f(sp.sigma_wacc_pb, 0)} pb, truncado en ±${f(sp.tope_wacc_pb, 0)} pb. Impuesto 27% (sin la rebaja a 23%). Objetivo = 80% DCF + 20% múltiplos fijos (CLP ${mm(MC.mult)}); con todo en su valor base reproduce DCF CLP ${mm(MC.dcf_base)} y objetivo CLP ${mm(MC.po_base)}. Fuente: análisis del equipo.`,
      `Assumptions: ${f(MC.n, 0)} iterations, seed ${MC.semilla}; triangular distributions (min/base/max, 4Q23–2Q26 extremes of the Risks tornado): EBITDA margin ${tr("Margen EBITDA")}, occupancy ${tr("Ocupación")}, same-store rent 2027–28 ${tr("Same-store rent growth (SSR)")}, tenant sales ${tr("Same-store sales growth (SSS)")} (${f(sp.rho_ssr_sss, 1)} correlation with rent, Gaussian copula; the rest independent), visitors ${tr("Visitas — crecimiento orgánico")}, maintenance + IT capex ${pc(sp.capex[0])}/${pc(sp.capex[1])}/${pc(sp.capex[2])}, nominal g ${pc(sp.g[0])}/${pc(sp.g[1])}/${pc(sp.g[2])}, real growth 2029–31 between the Note 11 floor and ceiling. WACC: normal shift, mean 0, standard deviation ${f(sp.sigma_wacc_pb, 0)} bp, truncated at ±${f(sp.tope_wacc_pb, 0)} bp. Tax 27% (no cut to 23%). Target = 80% DCF + 20% fixed multiples (CLP ${mm(MC.mult)}); with everything at base it reproduces the CLP ${mm(MC.dcf_base)} DCF and CLP ${mm(MC.po_base)} target. Source: Team analysis.`)));
    // efecto individual de cada palanca (tornado) para que se vea de dónde sale cada escenario
    const nombresEf = { "EBITDA margin": h("Margen EBITDA", "EBITDA margin"), "Occupancy": h("Ocupación", "Occupancy"), "Same-store rent growth 2027–28": h("Arriendo en mismas tiendas 2027–28 (precio)", "Same-store rent 2027–28 (price)"),
      "Long-term real rent growth 2029–31": h("Crecimiento real del arriendo 2029–31", "Real rent growth 2029–31"), "Maintenance + IT capex": h("Capex de mantención + TI", "Maintenance + IT capex"),
      "Visitor traffic 2027–28 (parking)": h("Visitas 2027–28 (parking)", "Visitor traffic 2027–28 (parking)"), "Tenant sales growth 2027–28": h("Ventas de locatarios 2027–28 (renta variable)", "Tenant sales 2027–28 (variable rent)"),
      "WACC ±100 pb": h("WACC nominal ±100 pb", "Nominal WACC ±100 bp"), "Impuesto 23%": h("Impuesto 23% (solo bull)", "23% tax (bull only)") };
    const filasEf = F.efectos.map((e) => [nombresEf[e.palanca] || e.palanca, (e.bajo < 0 ? "−" : e.bajo > 0 ? "+" : "") + (e.bajo === 0 ? "0" : mm(Math.abs(e.bajo))), "+" + mm(e.alto)]);
    const A = F.apv, Rr = F.rf;
    const nr = h([`APV (Ku ${pc(A.ku, 2)}): valor sin deuda`, "(+) escudos fiscales de la deuda", "= APV por acción", `Rf bono en pesos ${pc(Rr.bono, 2)} (WACC ${pb(Rr.efecto)} pb)`, `DCF VF: WACC constante ${pc(F.vf.wacc, 2)}`, `WACC implícita en el precio: ${pc(F.implicita.wacc, 2)}`],
      [`APV (Ku ${pc(A.ku, 2)}): unlevered value`, "(+) tax shields of debt", "= APV per share", `Peso-bond Rf ${pc(Rr.bono, 2)} (WACC ${pb(Rr.efecto)} bp)`, `DCF VF: constant WACC ${pc(F.vf.wacc, 2)}`, `WACC implied by the price: ${pc(F.implicita.wacc, 2)}`]);
    const filasR = [[nr[0], mm(A.sin_deuda), NB, NB], [nr[1], mm(A.escudos), NB, NB], [nr[2], mm(A.valor), mm(obj(A.valor)), vs(obj(A.valor))], [nr[3], mm(Rr.dcf), mm(Rr.po), vs(Rr.po)],
      [nr[4], mm(F.vf.dcf), mm(obj(F.vf.dcf)), vs(obj(F.vf.dcf))], [nr[5], mm(F.precio), NB, NB]];
    const filasDeu = F.deuda.filas.map(([t, ...v], i) => [f(t, i ? 1 : 2) + "x" + (i ? "" : h(" (actual)", " (current)")), ...v.map((x) => mm(x))]);
    const iBase = F.deuda.filas.findIndex((r) => Math.abs(r[0] - S.dfn_obj) < 1e-9);
    out.push(dosColumnas(
      [subtitulo(h("B. MÉTODOS Y SUPUESTOS ALTERNATIVOS", "B. ALTERNATIVE METHODS AND ASSUMPTIONS"), 100),
        tabla(h(["Prueba", "DCF (CLP)", "Precio obj.", "vs precio"], ["Test", "DCF (CLP)", "Target price", "vs price"]), [2625, 850, 900, 850], filasR, { derecha: [1, 2, 3], destacar: [2] }),
        subtitulo(h("C. OBJETIVO DE DEUDA VS SPREAD (DCF, CLP)", "C. DEBT TARGET VS SPREAD (DCF, CLP)"), 100),
        tabla([h("DFN/EBITDA \\ spread", "Net debt/EBITDA \\ spread"), ...F.deuda.spread.map((sp) => `${f(sp * 1e4, 0)} ${es ? "pb" : "bp"}`)], [1825, 1100, 1150, 1150], filasDeu, { derecha: [1, 2, 3], destacar: [iBase] })],
      [subtitulo(h("EFECTO DE CADA PALANCA SOLA EN EL DCF (CLP/ACCIÓN)", "EFFECT OF EACH LEVER ALONE ON THE DCF (CLP/SHARE)"), 100),
        tabla(h(["Palanca (extremos del tornado)", "Bear", "Bull"], ["Lever (tornado extremes)", "Bear", "Bull"]), [3425, 900, 900], filasEf, { derecha: [1, 2] }),
        subtitulo(h("D. SENSIBILIDADES ADICIONALES (FUERA DEL CASO BASE)", "D. ADDITIONAL SENSITIVITIES (OUTSIDE THE BASE CASE)"), 100),
        tabla(h(["Sensibilidad", "DCF", "Δ DCF", "Precio obj.", "vs precio"], ["Sensitivity", "DCF", "Δ DCF", "Target price", "vs price"]), [2325, 600, 600, 700, 700],
          N.sens_revision.map((r) => [h({ tax: "Impuesto por jurisdicción: Perú 29,5%, Colombia 35% (tasa ponderada 28,3% vs 27%)", da: "D&A al run-rate de 1S26 (0,88% vs 0,704% de los ingresos)", gp: "Gran Plaza incluida: USD 376 MM a cap rate 9,4% (orden de magnitud)" }[r.clave],
            { tax: "Tax by jurisdiction: Peru 29.5%, Colombia 35% (weighted 28.3% vs 27%)", da: "D&A at the 1H26 run-rate (0.88% vs 0.704% of revenue)", gp: "Gran Plaza included: USD 376m at a 9.4% cap rate (order of magnitude)" }[r.clave]),
            mm(r.dcf), pb(r.delta), mm(r.po), vs(r.po)]), { derecha: [1, 2, 3, 4] })]));
    out.push(fuente(h("Bear y bull: cada palanca en el mínimo/máximo histórico del tornado de Riesgos (Earnings Releases 4Q23 y 3Q25 para el margen; Nota 11 para el crecimiento real de largo plazo) en todos los años; el bull no recalcula el escudo fiscal de la WACC. Sección D: tasa de cada año escalada por la mezcla de EBITDA por país (Chile ~70%, Perú ~20%, Colombia ~10%; PwC Tax Summaries, Perú revisado 13-jul-26, Colombia 21-jul-26); Gran Plaza a nuestro cap rate implícito (10,1%) − precio pagado, con un cap rate sobre EBITDA aplicado a un NOI. Ninguna cambia la recomendación. Fuente: análisis del equipo.",
      "Bear and bull: every lever at the historical minimum/maximum of the Risks tornado (4Q23 and 3Q25 earnings releases for margin; Note 11 for long-run real growth) in all years; the bull does not recompute the WACC tax shield. Section D: each year's rate scaled by the EBITDA mix by country (Chile ~70%, Peru ~20%, Colombia ~10%; PwC Tax Summaries, Peru reviewed 13-Jul-26, Colombia 21-Jul-26); Gran Plaza at our implied cap rate (10.1%) − price paid, applying an EBITDA cap rate to an NOI. None changes the recommendation. Source: Team analysis.")));
    return out;
  },
});
