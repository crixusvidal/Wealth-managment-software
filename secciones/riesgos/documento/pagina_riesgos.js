// Página "Investment Risks" en español e inglés (A4, mismo formato que el informe).
// Figuras: graficos/informe/riesgos/{es,en}_fig2x_*.png (python graficos/pagina_riesgos.py ...).
// Uso: node informe/pagina_riesgos.js
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, ImageRun, Packer, PageNumber, Paragraph,
  ShadingType, Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType, LineRuleType,
} = require("docx");

const AZUL = "14284B", ROJO = "E9004B", TINTA = "1A1A1A", GRIS = "5F5F5F", FONDO = "EEF1F6";
const FUENTE = "Arial";
const COL_TEXTO = 6900, COL_SEP = 216, COL_FIG = 3550;
const FIG_DIR = path.join(__dirname, "..", "graficos", "figuras");
const PULGADA_PX = 96; // docx-js mide las imágenes en píxeles a 96 dpi
const ALTOS = { matriz: 2.05, tornado: 2.15, heatmap: 1.7, uf: 1.4, tasas: 1.66 }; // = graficos/pagina_riesgos.py (tasas: valuacion/graficos/sensibilidad_tasas.py)
const DAX = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "..", "valuacion", "graficos", "figuras", "datos_anexos.json"), "utf8"));
const SENS = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "..", "valuacion", "graficos", "figuras", "datos_sens_tasas.json"), "utf8")).pt;
const pc1 = (x, c) => (Math.abs(x) * 100).toFixed(1).replace(".", c);   // porcentaje sin signo

const ANCHO_FIG = COL_FIG / 1440;

const LANG = {
  es: {
    archivo: "Mallplaza_Investment_Risks_ES.docx", lang: "es-CL", pagina: "Página ",
    titulo: "RIESGOS DE INVERSIÓN",
    intro: "El perfil de riesgo de Mallplaza es moderado: su operación es resiliente y el valor depende sobre todo de la tasa de descuento (FR1, probabilidad alta). Un alza de las tasas largas baja el valor y refuerza la recomendación; lo que rompería la tesis de VENDER es lo contrario, una compresión adicional de cap rates o una prima estructural por los activos Tier A. Aun en el escenario bull, el precio objetivo no alcanza la cotización.",
    mitigantes: "Mitigantes: ",
    secciones: [
      { titulo: "EVALUACIÓN DE RIESGOS", parrafos: [
        "La Figura 23 clasifica cada riesgo según su probabilidad y su impacto sobre el valor por acción (Figuras 24, 25 y 26). Los riesgos que no se pudieron cuantificar se clasifican según el análisis cualitativo del equipo sobre la empresa."] },
      { titulo: "RIESGOS OPERACIONALES", riesgos: [
        { cod: "OR1", nombre: "Deterioro operacional.", texto: "Caídas de margen EBITDA, rentas, tráfico u ocupación hacia sus mínimos históricos. El peor caso, un margen de 75,8%, reduce el valor DCF a CLP 1.671 (−10,0%), y el piso histórico del crecimiento de rentas en mismas tiendas lo lleva a CLP 1.782 (−4,0%) (Figura 24).", tag: "No sistemático · Prob. media · Impacto medio",
          mit: "métricas estables: margen EBITDA anual de 76,2% (2022) a 80,5% (2025) y ocupación entre 95,1% y 96,6% en 4T23–2T26. Tras el COVID, los ingresos del 1T22 ya superaban en 18% a los del 1T19." },
        { cod: "OR2", nombre: "Tiendas ancla.", texto: "Las anclas, en su mayoría del grupo Falabella, atraen el flujo de visitas. Su reducción o salida bajaría las visitas y las rentas de las tiendas menores. El impacto va de −1,3% a −4,0% (Figura 24).", tag: "No sistemático · Prob. baja · Impacto bajo",
          mit: "contratos de largo plazo (duración remanente promedio de 7,1 años) y capacidad probada de reconvertir espacios ancla en tiendas menores con mayor renta por m²." }] },
      { titulo: "RIESGOS FINANCIEROS", riesgos: [
        { cod: "FR1", nombre: "Tasas de interés y costo de capital.", texto: "Un alza de tasas largas elevaría la WACC y el costo de refinanciar la deuda y el pipeline. La probabilidad es alta: el UST 10Y pasó de 4,98% (Rf del modelo) a 5,22% al 8-oct-26, máximo en 24 años (desde 2002); la Fed subió su tasa en septiembre y las tasas largas locales también suben. Con WACC +50 pb y +100 pb, el precio objetivo cae 6,5% y 12,3% (Figura 25).", tag: "Sistemático · Prob. alta · Impacto alto",
          mit: "79% de la deuda en UF a tasa real fija, duración de 6,6 años, DFN/EBITDA de 2,36x y calificación AAA local acotan el traspaso al costo de la deuda; los arriendos en UF compensan alzas inflacionarias. El alza de la tasa de descuento del patrimonio no es cubrible." }] },
      { titulo: "RIESGOS DE MERCADO Y MACROECONÓMICOS", riesgos: [
        { cod: "MR1", nombre: "Inflación y ciclo macro.", texto: `El 79% de la deuda está en UF: una UF 3% más alta resta CLP 38.691 MM a la utilidad contable (Nota 26). Sobre el precio objetivo, en cambio, el efecto es pequeño, porque las rentas indexadas suben con la inflación: una alza de 100 pb de la tasa larga nominal que venga toda de inflación esperada resta solo ${pc1(SENS.infl_100.pct, ",")}%; si viniera toda de tasa real, ${pc1(SENS.real_100.pct, ",")}% (Figura 26). El riesgo está en que un menor consumo impida traspasar la inflación a las rentas: si solo se reajustara el 70% del flujo, el efecto de la inflación subiría a ${pc1(SENS.infl_100_theta70.pct, ",")}%. Un menor consumo también frena las nuevas líneas de negocio (Anexo [X]).`, tag: "Sistemático · Prob. alta · Impacto bajo",
          mit: "rentas indexadas a la UF, modelo basado en arriendo más que en ventas, deuda baja aunque creciente y capacidad activa de rotar el portafolio." },
        { cod: "MR2", nombre: "Comercio electrónico.", texto: "Si el consumo migra masivamente al canal online, el mall pierde tráfico y atractivo. El e-commerce en Chile superó US$ 10.000 MM en 2025 (12,6% de las ventas del comercio). El peor caso de tráfico solo resta 1,3% al valor (Figura 24).", tag: "Sistemático · Prob. baja · Impacto medio",
          mit: "el canal online está madurando sin efecto visible en el modelo (visitas +16% en 2025). Los malls ofrecen entretención y gastronomía, difíciles de replicar online." },
        { cod: "MR3", nombre: "Tipo de cambio.", texto: "Cerca del 28% de los ingresos viene de Perú y Colombia, y el modelo supone un sol (PEN) que sube de CLP 268 a 290 entre 2026 y 2031. Una depreciación de 10% del PEN y del COP desde 2027E resta 4,3% al valor DCF, antes del efecto de las coberturas y de la deuda local (cálculo del equipo).", tag: "Sistemático · Prob. media · Impacto medio",
          mit: "Mallplaza radica su deuda en las monedas de cada país y, si no es posible, usa derivados de cobertura (forwards, swaps y cross currency swaps) (Análisis Razonado 2T26, p. 14)." } ] },
      { titulo: "RIESGOS REGULATORIOS Y ESG", riesgos: [
        { cod: "RR1", nombre: "Interlocking y conflictos de interés.", texto: "Falabella controla el 53,1% y arrienda espacios (Falabella, Sodimac, Tottus), lo que genera cuentas por cobrar con relacionadas. Fernando de Peña, gerente de Mallplaza por décadas, preside Falabella desde marzo de 2026. Contratos fuera de mercado afectarían directamente las rentas y el valor.", tag: "No sistemático · Prob. baja · Impacto alto",
          mit: "comité de directores de tres miembros y revisión de operaciones con partes relacionadas (Ley 18.046). Aun así, solo 1 de 9 directores es independiente." },
        { cod: "RR2", nombre: "Regulatorio y político.", texto: "Cambios tributarios, de permisos y de política económica en los tres países. En Chile, el oficialismo no tiene una mayoría holgada en la Cámara para el resto de su agenda (El Dínamo, 11-mar-26). En Perú, Fuerza Popular no tiene mayoría en el nuevo Congreso bicameral y el MEF estima en hasta S/ 50.000 MM el costo de las leyes del Congreso anterior (El Comercio). En Colombia, Hacienda estima un déficit de 7,8% del PIB y propone recortar el gasto en al menos 1% del PIB (El Tiempo, jul-26), y el Banco de la República subió su tasa a 12,25% (30-sep-26), lo que frena el consumo. Un retraso en permisos posterga el pipeline (USD 613 MM, 170.200 m² hasta 2028).", tag: "Sistemático · Prob. media · Impacto bajo",
          mit: "agendas pro inversión. En Chile, Kast promulgó la Ley de Reconstrucción (9-oct-26), que baja el impuesto corporativo de 27% a 23% en forma gradual desde 2027 (Emol, 9-oct-26), S&P reafirmó la calificación soberana A estable y la Ley 21.770 (sep-25) acorta los permisos. En Perú, el alza del salario mínimo a S/ 1.300 y la duplicación de Pensión 65 apoyan el consumo (Infobae, 28-jul-26). Colombia aporta solo ~11% de los ingresos. El modelo mantiene 27%: la rebaja en Chile sumaría ~3,9% al precio objetivo." }] },
    ],
    figuras: [
      { n: "Figura 23:", t: "Matriz de riesgos", img: "es_fig23_matriz.png", alto: ALTOS.matriz,
        fuente: "Color de fondo = severidad (probabilidad × impacto). Fuente: análisis del equipo." },
      { n: "Figura 24:", t: "Tornado de riesgo operacional (CLP por acción)", img: "es_fig24_tornado.png", alto: ALTOS.tornado,
        fuente: "Límites = mínimo y máximo histórico de cada variable (releases de Mallplaza 4T23–2T26; EEFF para capex de mantención y tasación, Nota 11). WACC 10,41% y g 3,0% constantes. Fuente: Mallplaza; equipo." },
      { n: "Figura 25:", t: "Sensibilidad del precio objetivo: WACC vs g (CLP)", img: "es_fig25_heatmap.png", alto: ALTOS.heatmap,
        fuente: "Precio objetivo = 80% DCF + 20% múltiplos EV/EBITDA (CLP 2.951). La columna g = 2,0% usa la fórmula de la hoja Sensibilidad. Fuente: equipo." },
      { n: "Figura 26:", t: "Precio objetivo según la tasa larga y la parte de su alza que viene de inflación (CLP)", img: "../../../valuacion/graficos/figuras/es_V12_mix_tasas_inflacion.png", alto: ALTOS.tasas,
        fuente: "BCP a 10 años (base 6,25%); alza repartida entre tasa real e inflación; 85% del flujo indexado (supuesto del equipo). Fuente: BCCh; equipo." },
    ],
  },
  en: {
    archivo: "Mallplaza_Investment_Risks_EN.docx", lang: "en-US", pagina: "Page ",
    titulo: "INVESTMENT RISKS",
    intro: "Mallplaza's risk profile is moderate: its operations are resilient and value depends mainly on the discount rate (FR1, high likelihood). A rise in long-term rates lowers value and reinforces the recommendation; what would break the SELL thesis is the opposite, a further compression of cap rates or a structural premium for Tier A assets. Even in the bull scenario, the target price does not reach the market price.",
    mitigantes: "Mitigants: ",
    secciones: [
      { titulo: "RISK ASSESSMENT", parrafos: [
        "Figure 23 classifies each risk by likelihood and by its impact on value per share (Figures 24, 25 and 26). Risks that could not be quantified are classified according to the team's qualitative analysis of the company."] },
      { titulo: "OPERATIONAL RISKS", riesgos: [
        { cod: "OR1", nombre: "Operating deterioration.", texto: "EBITDA margin, rents, footfall or occupancy falling to their historical lows. The worst case, a 75.8% margin, cuts DCF value to CLP 1,671 (−10.0%), and the historical floor of same-store rent growth takes it to CLP 1,782 (−4.0%) (Figure 24).", tag: "Non‑systematic · Medium likelihood · Medium impact",
          mit: "stable metrics: annual EBITDA margin rose from 76.2% (2022) to 80.5% (2025) and occupancy held between 95.1% and 96.6% in 4Q23–2Q26. After COVID, 1Q22 revenue was already 18% above 1Q19." },
        { cod: "OR2", nombre: "Anchor stores.", texto: "Anchors, mostly from the Falabella group, drive footfall. Shrinking or losing them would lower visits and rents from smaller stores. The impact ranges from −1.3% to −4.0% (Figure 24).", tag: "Non‑systematic · Low likelihood · Low impact",
          mit: "long-term leases (average remaining term of 7.1 years) and a proven ability to convert anchor space into smaller stores with higher rent per m²." }] },
      { titulo: "FINANCIAL RISKS", riesgos: [
        { cod: "FR1", nombre: "Interest rates and cost of capital.", texto: "Higher long-term rates would raise the WACC and the cost of refinancing debt and the pipeline. Likelihood is high: the 10Y UST moved from 4.98% (model Rf) to 5.22% on 8-Oct-26, a 24-year high (since 2002); the Fed hiked in September and the Central Bank of Chile reports rising local long-term rates. With WACC +50 bp and +100 bp, the target price falls 6.5% and 12.3% (Figure 25).", tag: "Systematic · High likelihood · High impact",
          mit: "79% of debt is fixed real-rate in UF, with 6.6 years of duration, 2.36x net debt/EBITDA and a local AAA rating, which limit the pass-through to the cost of debt; UF-linked rents offset inflation. A higher cost of equity cannot be hedged." }] },
      { titulo: "MARKET & MACROECONOMIC RISKS", riesgos: [
        { cod: "MR1", nombre: "Inflation and macro cycle.", texto: `79% of debt is in UF: a 3% higher UF takes CLP 38,691 million off accounting earnings (Note 26). On the target price, however, the effect is small, because indexed rents rise with inflation: a 100 bp rise in the long nominal rate coming entirely from expected inflation takes off only ${pc1(SENS.infl_100.pct, ".")}%; if it came entirely from the real rate, ${pc1(SENS.real_100.pct, ".")}% (Figure 26). The risk is that weaker consumption prevents passing inflation through to rents: if only 70% of the cash flow were indexed, the effect of inflation would rise to ${pc1(SENS.infl_100_theta70.pct, ".")}%. Weaker consumption also slows the new business lines (Appendix [X]).`, tag: "Systematic · High likelihood · Low impact",
          mit: "UF-indexed rents, a model based on rent rather than sales, low (if rising) leverage and an active ability to rotate the portfolio." },
        { cod: "MR2", nombre: "E‑commerce.", texto: "If consumption migrates massively online, malls lose footfall and appeal. E‑commerce in Chile exceeded US$10 billion in 2025, around 12.6% of retail sales. The worst traffic case only takes 1.3% off value (Figure 24).", tag: "Systematic · Low likelihood · Medium impact",
          mit: "the online channel is maturing with no visible effect on the model (visits +16% in 2025). Malls offer entertainment and dining, hard to replicate online." },
        { cod: "MR3", nombre: "Foreign exchange.", texto: "Around 28% of revenue comes from Peru and Colombia, and the model assumes the sol (PEN) rises from CLP 268 to 290 between 2026 and 2031. A 10% depreciation of the PEN and COP from 2027E takes 4.3% off DCF value, before hedges and the effect on local-currency debt (team estimate).", tag: "Systematic · Medium likelihood · Medium impact",
          mit: "Mallplaza raises debt in each country's currency and, where not possible, uses hedging derivatives (forwards, swaps and cross currency swaps) (Management Discussion 2Q26, p. 14)." } ] },
      { titulo: "REGULATORY & ESG RISKS", riesgos: [
        { cod: "RR1", nombre: "Interlocking and conflicts of interest.", texto: "Falabella controls 53.1% and leases space through Falabella, Sodimac and Tottus, which creates receivables with related parties. Fernando de Peña, Mallplaza's CEO for decades, has chaired Falabella since March 2026. Off-market contracts would directly affect rents and value.", tag: "Non‑systematic · Low likelihood · High impact",
          mit: "a three-member directors' committee and review of related-party transactions (Law 18,046). Even so, safeguards are limited: only 1 of 9 directors is independent." },
        { cod: "RR2", nombre: "Regulatory and political.", texto: "Changes in tax, permitting and economic policy in all three countries. In Chile, the ruling coalition lacks a comfortable majority in the lower house for the rest of its agenda (El Dínamo, 11-Mar-26). In Peru, Fuerza Popular lacks a majority in the new bicameral Congress and the MEF puts the cost of laws passed by the previous Congress at up to S/ 50,000 million (El Comercio). In Colombia, the finance ministry puts the deficit at 7.8% of GDP and proposes spending cuts of at least 1% of GDP (El Tiempo, Jul-26), and the central bank raised its rate to 12.25% (30-Sep-26), which slows consumption. Permit delays push back the pipeline (USD 613 million, 170,200 m² to 2028).", tag: "Systematic · Medium likelihood · Low impact",
          mit: "pro-investment agendas. In Chile, Kast signed the Reconstruction Law (9-Oct-26), which gradually cuts corporate tax from 27% to 23% from 2027 (Emol, 9-Oct-26), S&P affirmed Chile's A stable rating and Law 21,770 (Sep-25) shortens permitting. In Peru, the minimum-wage rise to S/ 1,300 and the doubling of Pensión 65 support consumption (Infobae, 28-Jul-26). Colombia contributes only ~11% of revenue. The model keeps 27%: the Chilean cut would add ~3.9% to the target price." }] },
    ],
    figuras: [
      { n: "Figure 23:", t: "Risk matrix", img: "en_fig23_matriz.png", alto: ALTOS.matriz,
        fuente: "Background colour = severity (likelihood × impact). Source: Team analysis." },
      { n: "Figure 24:", t: "Operational risk tornado (CLP per share)", img: "en_fig24_tornado.png", alto: ALTOS.tornado,
        fuente: "Limits = historical minimum and maximum of each variable (Mallplaza releases 4Q23–2Q26; financial statements for maintenance capex and appraisal, Note 11). WACC 10.41% and g 3.0% constant. Source: Mallplaza; Team analysis." },
      { n: "Figure 25:", t: "Target price sensitivity: WACC vs g (CLP)", img: "en_fig25_heatmap.png", alto: ALTOS.heatmap,
        fuente: "Target price = 80% DCF + 20% EV/EBITDA multiples (CLP 2,951). The g = 2.0% column uses the same formula as the Sensibilidad sheet. Source: Team analysis." },
      { n: "Figure 26:", t: "Target price by long nominal rate and the share of its rise that comes from inflation (CLP)", img: "../../../valuacion/graficos/figuras/en_V12_mix_tasas_inflacion.png", alto: ALTOS.tasas,
        fuente: "10-year BCP (base 6.25%, 30-Sep-26); the rise is split between real rate and inflation; 85% of cash flow is indexed to inflation (team assumption). Source: Central Bank of Chile; Team analysis." },
    ],
  },
};

// Cifras del modelo v4 en los textos de riesgos: los números salen de comun/cifras.js (tornado, sensibilidad WACC × g, escenarios) ------
for (const lang of ["es", "en"]) {
  const { K, f } = require("../../comun/cifras")(lang);
  const es = lang === "es";
  const ef = Object.fromEntries(K.efectos.map((e) => [e.palanca, e]));
  const dcf = K.dcf, pct = (x) => f.pc(x / dcf, 1);
  const SENS_G3 = DAX.B.sens.v.map((r) => r[1]);                       // valor DCF con g = 3,0%: ΔWACC −100, −50, 0, +50, +100 pb
  const po = (v) => K.peso_dcf * v + (1 - K.peso_dcf) * K.mult_valor;
  const caida = (v) => f.pc(Math.abs(po(v) / K.po - 1), 1);
  const L = LANG[lang];
  const buscar = (cod) => L.secciones.flatMap((s) => s.riesgos || []).find((r) => r.cod === cod);
  const mg = ef["EBITDA margin"].bajo, ssr = ef["Same-store rent growth 2027–28"].bajo, vis = ef["Visitor traffic 2027–28 (parking)"].bajo;
  buscar("OR1").texto = es
    ? `Caídas de margen EBITDA, rentas, tráfico u ocupación hacia sus mínimos históricos. El peor caso, un margen trimestral de ${f.pc(K.esc_param.m_lo, 1)}, reduce el valor DCF a CLP ${f.n(dcf + mg)} (${pct(mg)}), y el piso histórico del crecimiento de rentas en mismas tiendas lo lleva a CLP ${f.n(dcf + ssr)} (${pct(ssr)}) (Figura 24). Todas las palancas a la vez, con la WACC +100 pb, definen el escenario bear (precio objetivo CLP ${f.n(K.esc.bear.po)}, ${f.pc(K.esc.bear.vs, 1)} frente al precio).`
    : `EBITDA margin, rents, footfall or occupancy falling to their historical lows. The worst case, a ${f.pc(K.esc_param.m_lo, 1)} quarterly margin, cuts DCF value to CLP ${f.n(dcf + mg)} (${pct(mg)}), and the historical floor of same-store rent growth takes it to CLP ${f.n(dcf + ssr)} (${pct(ssr)}) (Figure 24). All levers together, with WACC +100 bp, define the bear scenario (target price CLP ${f.n(K.esc.bear.po)}, ${f.pc(K.esc.bear.vs, 1)} against the share price).`;
  buscar("OR2").texto = es
    ? `Las anclas, en su mayoría del grupo Falabella, atraen el flujo de visitas. Su reducción o salida bajaría las visitas y las rentas de las tiendas menores. El impacto va de ${pct(vis)} a ${pct(ssr)} (Figura 24).`
    : `Anchors, mostly from the Falabella group, drive footfall. Shrinking or losing them would lower visits and rents from smaller stores. The impact ranges from ${pct(vis)} to ${pct(ssr)} (Figure 24).`;
  buscar("FR1").texto = es
    ? `Las tasas largas dominan el valor, en ambos sentidos. Un alza eleva la WACC y el costo de refinanciar la deuda y el pipeline, y es lo más probable: el UST 10Y pasó de 4,98% (Rf del modelo) a 5,22% al 8-oct-26, máximo desde 2002 (por verificar); la Fed subió su tasa en septiembre y las tasas largas locales también suben. Con WACC nominal +50 pb y +100 pb, el precio objetivo cae ${caida(SENS_G3[3])} y ${caida(SENS_G3[4])} (Figura 25), lo que refuerza la recomendación. Lo que rompería la tesis es lo contrario: con WACC −100 pb sube ${caida(SENS_G3[0])}, a CLP ${f.n(po(SENS_G3[0]))}, y aun así queda ${f.pc(1 - po(SENS_G3[0]) / K.precio, 0)} bajo el precio.`
    : `Long-term rates dominate value, in both directions. A rise lifts the WACC and the cost of refinancing debt and the pipeline, and is the more likely outcome: the 10Y UST moved from 4.98% (model Rf) to 5.22% on 8-Oct-26, the highest since 2002 (to be verified); the Fed hiked in September and the Central Bank of Chile reports rising local long-term rates. With nominal WACC +50 bp and +100 bp, the target price falls ${caida(SENS_G3[3])} and ${caida(SENS_G3[4])} (Figure 25), which reinforces the recommendation. What would break the thesis is the opposite: with WACC −100 bp it rises ${caida(SENS_G3[0])}, to CLP ${f.n(po(SENS_G3[0]))}, and still sits ${f.pc(1 - po(SENS_G3[0]) / K.precio, 0)} below the price.`;
  const rr1 = buscar("RR1"), rr2 = buscar("RR2");
  rr1.texto += es
    ? ` Hay también un riesgo de asignación de capital: la expansión rinde menos que su costo (RONIC real de ${f.pc(K.ronic, 1)} frente a una WACC real de ${f.pc(K.wacc_t_real, 1)}), de modo que seguir creciendo beneficia a la utilidad contable y a un controlador con otros negocios en esos malls más que al minoritario (por verificar; ver Valuación).`
    : ` There is also a capital-allocation risk: expansion earns less than its cost (real RONIC of ${f.pc(K.ronic, 1)} against a real WACC of ${f.pc(K.wacc_t_real, 1)}), so continued growth benefits accounting profit and a controller with other businesses in those malls more than the minority shareholder (to be verified; see Valuation).`;
  const tax = f.pc(K.supuestos_indiv[0][1] / K.supuestos_base, 1);
  rr2.mit = rr2.mit.replace(/~3,9%|~3\.9%/, "~" + tax);
}

const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || 16, bold: o.bold,
  italics: o.italics, color: o.color || TINTA });
const sinBordes = { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } };
const cuerpo = { line: 228, lineRule: LineRuleType.AUTO, before: 0 };

function subtitulo(texto) {
  return new Paragraph({ spacing: { before: 110, after: 50 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ROJO, space: 1 } },
    children: [run(texto, { bold: true, color: AZUL })] });
}

function columnaTexto(L) {
  const out = [];
  for (const s of L.secciones) {
    out.push(subtitulo(s.titulo));
    for (const p of s.parrafos || [])
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { ...cuerpo, after: 40 },
        children: [run(p)] }));
    for (const r of s.riesgos || []) {
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { ...cuerpo, after: 14 },
        children: [run(`${r.cod}: `, { bold: true, color: ROJO }), run(r.nombre + " ", { bold: true }),
          run(r.texto + " "), run(`(${r.tag})`, { italics: true, color: GRIS })] }));
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { ...cuerpo, after: 60 },
        children: [run(L.mitigantes, { bold: true, color: AZUL }), run(r.mit)] }));
    }
  }
  return out;
}

function columnaFiguras(L) {
  const out = [];
  L.figuras.forEach((f, k) => {
    out.push(new Paragraph({ spacing: { before: k === 0 ? 40 : 100, after: 30, line: 228, lineRule: LineRuleType.AUTO },
      children: [run(f.n + " ", { bold: true, color: AZUL, size: 14 }), run(f.t, { color: AZUL, size: 14 })] }));
    out.push(new Paragraph({ spacing: { before: 0, after: 20, line: 240, lineRule: LineRuleType.AUTO }, children: [
      new ImageRun({ type: "png", data: fs.readFileSync(path.join(FIG_DIR, f.img)),
        transformation: { width: Math.round(ANCHO_FIG * PULGADA_PX), height: Math.round(f.alto * PULGADA_PX) } })] }));
    out.push(new Paragraph({ spacing: { before: 0, after: 0, line: 216, lineRule: LineRuleType.AUTO },
      children: [run(f.fuente, { size: 11, color: GRIS })] }));
  });
  return out;
}

function documento(L) {
  const cel = (w, children, extra = {}) => new TableCell({ width: { size: w, type: WidthType.DXA },
    borders: sinBordes, margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP,
    children, ...extra });
  const dosColumnas = () => new Table({ width: { size: COL_TEXTO + COL_SEP + COL_FIG, type: WidthType.DXA },
    columnWidths: [COL_TEXTO, COL_SEP, COL_FIG],
    borders: { ...sinBordes, insideHorizontal: sinBordes.top, insideVertical: sinBordes.top },
    rows: [new TableRow({ children: [
      cel(COL_TEXTO, columnaTexto(L)),
      cel(COL_SEP, [new Paragraph({ children: [] })]),
      cel(COL_FIG, columnaFiguras(L)),
    ] })] });
  return ({
    styles: { default: { document: { run: { font: FUENTE, size: 16 }, paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 },
        margin: { top: 1000, right: 620, bottom: 900, left: 620, header: 450, footer: 400 },
        pageNumbers: { start: 9 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0", space: 2 } },
        children: [run("Mallplaza (Plaza S.A.)  |  CFA Institute Research Challenge", { size: 14, color: GRIS })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [run(L.pagina, { size: 14, color: GRIS }),
          new TextRun({ children: [PageNumber.CURRENT], font: FUENTE, size: 14, color: GRIS })] })] }) },
      children: [
        new Paragraph({ spacing: { after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 2 } },
          children: [run(L.titulo, { bold: true, color: AZUL, size: 30 })] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { ...cuerpo, after: 80 },
          shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO },
          children: [run(L.intro)] }),
        dosColumnas(),
      ],
    }],
  });
}

module.exports = { documento, LANG };

if (require.main === module) (async () => {
  for (const L of Object.values(LANG)) {
    const salida = path.join(__dirname, L.archivo);
    fs.writeFileSync(salida, await Packer.toBuffer(new Document(documento(L))));
    console.log(salida);
  }
})();
