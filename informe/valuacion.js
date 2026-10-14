// Sección Valuación (texto del borrador del equipo, cifras sincronizadas con el modelo v2, citas a figuras)
// + Anexo de construcción de la WACC, en español e inglés. Figuras pendientes = recuadro punteado.
// Figuras y datos: graficos/informe/valuacion/ (python graficos/valuacion.py <Mallplaza_Figuras_Valuacion.xlsx>).
// Uso: node informe/valuacion.js
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, ImageRun, LineRuleType, Packer, PageNumber,
  Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

const AZUL = "14284B", ROJO = "E9004B", TINTA = "1A1A1A", GRIS = "5F5F5F", FONDO = "EEF1F6";
const FUENTE = "Arial", ANCHO = 10666, COL_TEXTO = 6900, COL_SEP = 216, COL_FIG = 3550, MEDIA = 5225;
const FIG = path.join(__dirname, "..", "graficos", "informe", "valuacion");
const D = JSON.parse(fs.readFileSync(path.join(FIG, "datos_wacc.json"), "utf8"));
const PX = 96;
const ALTO = { V9: 2.15, V1: 1.85, V2: 1.9, V3: 2.35, V4: 2.3, V5: 1.75, V6: 1.95, W1: 1.9 }; // = graficos/valuacion.py

const fmt = (lang) => (x, dec = 1) => {
  const s = x.toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec }).replace("-", "−");
  return lang === "es" ? s.replace(/[.,]/g, (c) => (c === "." ? "," : ".")) : s;
};

const L = {
  es: {
    archivo: "Mallplaza_Valuacion_ES.docx", pag: "Página ", pagA: "A-", titulo: "VALUACIÓN",
    lead: "Recomendamos VENDER Mallplaza con un precio objetivo de CLP 2.076 por acción (80% DCF y 20% múltiplos), 41,6% bajo el cierre de CLP 3.555 del 30-sep-26. Ningún método justifica el precio actual (Figura 27).",
    paginas: [
      { texto: [
          { h1: "SÍNTESIS Y PONDERACIÓN" },
          { p: "Asignamos 80% al DCF porque los flujos de Mallplaza son muy predecibles: el 81,4% de los ingresos UDM son arriendos y el 93% de ellos es renta fija. Los múltiplos (20%) contrastan el modelo con lo que paga hoy el mercado; limitamos su peso porque en Chile solo hay dos pares directos (Parque Arauco y Cencosud Shopping) y porque no capturan la expansión de GLA ni la trayectoria de deuda de Mallplaza, que sí están en el DCF." },
          { h1: "DCF: ESTRUCTURA Y TASA DE DESCUENTO" },
          { p: "Valoramos el FCFF en dos etapas, en CLP nominales, al 30-jun-26 y con convención de mitad de período. La etapa explícita (2S26E–2031E) incluye el capex y la maduración de los proyectos que abren hasta 2032 (108.301 m² de GLA madura en el año terminal). Desde 2032, una perpetuidad de Gordon explica el 76% del EV, sin GLA neta nueva y con capex solo de mantención (Anexo [B])." },
          { p: ["WACC. ", "Cambia por período a medida que la DFN/EBITDA converge de 2,36x a 3,5x en 2031 (D/E de 20,0% a 28,5%): 10,54% en 2S26E y 10,41% en el terminal (7,32% y 7,20% reales). El Ke se calcula en USD y se pasa a CLP con el diferencial de inflación: Rf 4,98% + beta 1,14 × ERP 4,14% + CRP 1,41% = 11,10% (11,93% en CLP, terminal). La beta re-apalanca la mediana desapalancada de seis pares (0,94) y el CRP pondera el riesgo soberano por ingresos (71,8% Chile, 17,7% Perú y 10,5% Colombia). El Kd parte del cupón de la última emisión en UF (3,6% real) más 25 pb por vuelta de apalancamiento sobre 2,36x: 4,92% después de impuestos en 2S26E y 5,11% en el terminal. A CLP 3.555 el patrimonio rinde casi lo mismo que esa deuda (Figura 28; detalle en el Anexo [Z])."] },
          { p: ["Estructura de capital. ", "La compañía usa su capacidad de deuda para crecer (llegó a 3,78x en 2024 tras comprar Open Plaza Perú). Como el FCFF ya cubre el capex, la deuda adicional se devuelve a los accionistas como distribución (CLP 1,79 billones acumulados a 2031): no entra en el EV, pero hace que la deuda supuesta en la WACC exista en el plan (Anexo [D])."] },
          { h1: "DCF: PROYECCIÓN OPERATIVA" },
          { p: "El EBITDA reportado (CLP 544.935 MM UDM, margen de 80,5%) excluye la revalorización de propiedades (NIC 40); proyectamos un margen de 79,8% hasta 2031E. La tasa de impuesto operacional sube de 16,8% en 2026E a 27% en 2030E, a medida que se diluyen los beneficios vigentes. El capex base (pipeline más mantención de 5,5% de los ingresos) llega a CLP 240.660 MM en 2027 y baja a ~CLP 50.000 MM desde 2029. Desde 2029 abren además 27.075 m² al año (promedio histórico), con capex de CLP 95.000–104.000 MM anuales, una rampa de 35%/90%/100% y un yield on cost de 6,9%, al costo del pipeline identificado (USD 3.603/m²: 170.200 m² y USD 613,2 MM). El capital de trabajo (15,6% × Δ ingresos) exige CLP 7.100–9.900 MM al año y mueve el precio solo ±1,2% (Anexo [C])." },
          { h1: "DCF: VALOR TERMINAL Y CAP RATE IMPLÍCITO" },
          { p: "Usamos g = 3,0% nominal (0% real). El RONIC real de la expansión (5,0%) es menor que la WACC real (7,2%): crecer más allá del horizonte destruye valor y un g mayor lo reduce (Figura 29). El DCF implica un cap rate de 10,1% sobre el EV y de 11,3% sobre el valor terminal, frente a 6,0% del mercado, 6,4% de los pares chilenos, 7,5% a valor libro y 5,55% de los tasadores (Nota 11) (Figura 30). Esa brecha no es una inconsistencia: es el núcleo de la tesis. Para justificar el precio actual, la WACC tendría que ser 285 pb menor (Figura 31)." },
        ],
        figuras: [
          ["Figura 27:", "Ningún método justifica el precio actual (CLP por acción)", "V1_football", "Rango 52 semanas pendiente (Bloomberg). Fuente: análisis del equipo; Bloomberg."],
          ["Figura 28:", "A CLP 3.555 el patrimonio rinde casi lo mismo que la deuda", "V2_tasas", "Tasas reales en UF. Fuente: BCCh (30-06-26); Mallplaza; análisis del equipo."],
          ["Figura 29:", "Crecer bajo la WACC destruye valor: g real = 0%", "V3_value_driver", "Arriba: valor DCF según g. Abajo: valor terminal relativo según el RONIC. Fuente: análisis del equipo."]] },
      { texto: [
          { h1: "VALUACIÓN RELATIVA: MÚLTIPLOS (20%)" },
          { p: "El set es el mismo de la beta: Parque Arauco y Cencosud Shopping (Chile), Allos, Multiplan e Iguatemi (Brasil) e IRSA (Argentina, con cautela por NIC 29 y un solo analista en el NTM). Excluimos Fibra Danhos, que como FIBRA distribuye el 95% de su resultado fiscal (Anexo [E]). Brasil transa a múltiplos más bajos por sus tasas y primas de riesgo (mediana de 9,0x EV/EBITDA frente a 15,7x en Chile; Figura 32). Como el 71,8% de los ingresos viene de Chile, el múltiplo pondera 70% la mediana chilena y 30% la del set (11,1x): 14,4x." },
          { p: "Usamos EV/EBITDA porque es neutral al apalancamiento y a la revalorización (NIC 40). A 14,4x, el valor es CLP 2.951 por acción. Mallplaza transa a 16,8x LTM: 7% sobre la mediana chilena, bajo Parque Arauco (18,4x) y 17% sobre el múltiplo ponderado (Figura 33). Excluimos el P/E (CLP 8.644, distorsionado por la revalorización). Los múltiplos usan la deuda neta de Bloomberg (CLP 1.311.569 MM), igual que los pares; el DCF, la DFN de la compañía (CLP 1.285.740 MM)." },
          { h1: "SENSIBILIDAD Y RIESGOS DE LA VALUACIÓN" },
          { p: ["Sensibilidad. ", "Con la WACC ±100 pb y g entre 2,5% y 4,0%, el DCF va de CLP 1.379 a 2.329 (Figura 25; Anexo [F]), y el peor caso operacional histórico le resta 10,0% (Figura 24). Todos los múltiplos quedan bajo el precio: EV/EBITDA LTM CLP 2.951 (−17%), NTM CLP 2.914 (−18%), EV/EBIT CLP 3.098 (−13%), P/BV CLP 3.171 (−11%) y la mediana chilena sola CLP 3.294 (−7%) (Figura 27)."] },
          { p: ["Robustez. ", "El APV da CLP 1.894 (2% sobre el DCF); con el bono en pesos como Rf, el precio objetivo sube a ~CLP 2.439 (−31%; Anexo [Z]); y con un objetivo de DFN/EBITDA entre 2,36x y 4,0x, el DCF va de CLP 1.819 a 1.894."] },
          { p: ["Precio objetivo. ", "80% × CLP 1.857 + 20% × CLP 2.951 = CLP 2.076, un retorno de −41,6% frente a CLP 3.555 (30-sep-26). Llevar el DCF del 30-jun al 30-sep-26 lo subiría a ~CLP 1.910 y el precio objetivo a ~CLP 2.118 (−40,4%): la recomendación no cambia."] },
          { p: ["Riesgos al alza. ", "Una compresión sostenida de los cap rates por mayor demanda institucional por activos defensivos, o una prima estructural para los activos Tier A de Mallplaza (ver Riesgos de inversión)."] },
        ],
        figuras: [
          ["Figura 30:", "El mercado paga 6,0%; nuestros flujos valen 10,1%", "V5_cap_rate", "Gran Plaza: hecho posterior, no entra al modelo. Fuente: Mallplaza; Bloomberg; DF; análisis del equipo."],
          ["Figura 31:", "El precio exige una WACC 285 pb menor", "V6_dcf_inverso", "Fuente: análisis del equipo."],
          ["Figura 32:", "Múltiplos de los pares vs tasa local", null, "Gráfico pendiente (falta la NTN-B de Brasil, Bloomberg)."],
          ["Figura 33:", "Valor implícito por múltiplo de cada par vs DCF y precio (CLP por acción)", "V9_pares", "Precio objetivo = 80% DCF + 20% EV/EBITDA LTM. P/E excluido (incluye revalorización). Fuente: Bloomberg (30-09-26); análisis del equipo."]] },
    ],
    anexo: {
      titulo: "ANEXO [Z]: CONSTRUCCIÓN DE LA WACC",
      lead: "La WACC se calcula por período: el D/E converge linealmente desde la estructura actual (20,0%) hasta la objetivo (28,5%, DFN/EBITDA 3,5x) en 2031. El Ke se construye en dólares (Rf del Tesoro de EE.UU. + beta re-apalancada × ERP + CRP ponderado por ingresos) y se convierte a pesos con el diferencial de inflación.",
      formulas: ["βL = βU × [1 + (1 − t) × D/E]  (Hamada)", "Ke USD = Rf + βL × ERP + CRP ponderado", "Ke CLP = (1 + Ke USD) × (1 + π Chile) / (1 + π EE.UU.) − 1", "Kd CLP = (1 + Kd real + spread) × (1 + π Chile) − 1", "WACC = E/V × Ke CLP + D/V × Kd CLP × (1 − t)"],
      sA: "A. SUPUESTOS DE MERCADO Y ESTRUCTURA",
      inputs: [["Rf: bono del Tesoro EE.UU. 10 años", "rf", "%", "Tasa_WACC (FFCF)"], ["Prima por riesgo de mercado (ERP)", "erp", "%", "Damodaran"],
        ["Inflación de largo plazo Chile", "pi_cl", "%", "Meta BCCh"], ["Inflación de largo plazo EE.UU. (breakeven 10Y)", "pi_us", "%", "Bloomberg, jun-26"],
        ["Tasa de impuesto Chile", "tax", "%", "Ley de la Renta"], ["Kd real: cupón última emisión en UF", "kd_real", "%", "Mallplaza"],
        ["Aumento del Kd real por vuelta de DFN/EBITDA", "spread", "%", "Supuesto del equipo"], ["DFN/EBITDA inicial (30-06-26) → objetivo", "dfn", "x", "Release 2Q26; supuesto del equipo"],
        ["Deuda bruta / market cap (estructura actual, MM CLP)", "estructura", "", "Tasa_WACC (FFCF)"]],
      colsA: ["Supuesto", "Valor", "Fuente"],
      sB: "B. BETA: PARES DESAPALANCADOS",
      colsB: ["Empresa", "β apalancada", "Deuda (MM CLP)", "Market cap (MM CLP)", "Impuesto", "β desapalancada"],
      medianaB: "Mediana (6 pares incluidos)", notaB: "Fuente: Bloomberg (betas), Tasa_WACC. Costa Verde se excluye porque no está en el set de comparables.",
      sC: "C. RIESGO PAÍS PONDERADO POR INGRESOS",
      colsC: ["País", "Ponderación (ingresos LTM)", "CRP (Damodaran)", "Aporte"], paises: ["Chile", "Perú", "Colombia"], totalC: "CRP ponderado",
      sD: "D. WACC POR PERÍODO",
      filasD: [["53", "DFN/EBITDA", "x"], ["56", "D/E", "%"], ["57", "D/V", "%"], ["59", "Beta re-apalancada", "b"], ["60", "Ke USD", "%2"], ["61", "Ke nominal CLP", "%2"],
        ["62", "Kd real (UF) con spread", "%2"], ["63", "Kd nominal CLP", "%2"], ["64", "Kd después de impuestos", "%2"], ["65", "WACC nominal CLP", "%2"], ["66", "WACC real", "%2"]],
      per: ["2S26E", "2027E", "2028E", "2029E", "2030E", "2031E", "Terminal"],
      fW: ["Figura Z1:", "WACC nominal y D/E por período", "Control: la réplica coincide con la hoja WACC del modelo v2 (diferencia 0 pb en todos los períodos). Fuente: Damodaran; Bloomberg; BCCh; análisis del equipo."],
      sF: "E. ELECCIÓN DE LA TASA LIBRE DE RIESGO",
      colsF: ["Alternativa", "Valor"],
      filasF: [["Base del modelo: Rf EE.UU. + CRP Chile, en CLP", "base", "%2"], ["Bono en pesos 10 años BCCh (30-06-26)", "bono", "%2"], ["Diferencia", "dif", "pb"],
        ["Efecto en la WACC terminal (diferencia × E/V)", "efecto", "pb"], ["Valor DCF con el bono en pesos (CLP)", "dcf", "clp"], ["Precio objetivo 80/20 en ese caso (CLP)", "po", "clp"]],
      notasF: ["Por qué UST + CRP y no el bono en pesos: el bono local ya incluye el riesgo Chile, y sumarle el CRP lo contaría dos veces; además, un 28% de los ingresos viene de Perú y Colombia, cuyo riesgo se agrega con el CRP ponderado.",
        "Es la elección más conservadora: con el bono en pesos el precio objetivo sube a ~CLP 2.439, todavía 31% bajo el precio de CLP 3.555. La recomendación no cambia.",
        "Fuente: BCCh, Informativo Diario 2-jul-26; réplica del DCF v2 con la WACC desplazada en todos los períodos."],
    },
  },
  en: {
    archivo: "Mallplaza_Valuacion_EN.docx", pag: "Page ", pagA: "A-", titulo: "VALUATION",
    lead: "We recommend SELL on Mallplaza with a target price of CLP 2,076 per share (80% DCF and 20% multiples), 41.6% below the CLP 3,555 close of 30-Sep-26. No method justifies the current price (Figure 27).",
    paginas: [
      { texto: [
          { h1: "SUMMARY AND WEIGHTING" },
          { p: "We give the DCF an 80% weight because Mallplaza's cash flows are highly predictable: 81.4% of LTM revenue is rent and 93% of that is fixed. Multiples (20%) test the model against what the market pays today; we cap their weight because Chile has only two direct peers (Parque Arauco and Cencosud Shopping) and because multiples do not capture Mallplaza's GLA expansion or debt path, which the DCF does." },
          { h1: "DCF: STRUCTURE AND DISCOUNT RATE" },
          { p: "We value FCFF in two stages, in nominal CLP, at 30-Jun-26 with the mid-year convention. The explicit stage (2H26E–2031E) includes the capex and ramp-up of projects opening through 2032 (108,301 m² of mature GLA in the terminal year). From 2032, a Gordon perpetuity accounts for 76% of EV, with no new net GLA and maintenance capex only (Appendix [B])." },
          { p: ["WACC. ", "It changes by period as net debt/EBITDA converges from 2.36x to 3.5x by 2031 (D/E from 20.0% to 28.5%): 10.54% in 2H26E and 10.41% in the terminal period (7.32% and 7.20% real). Ke is computed in USD and converted to CLP with the inflation differential: Rf 4.98% + beta 1.14 × ERP 4.14% + CRP 1.41% = 11.10% (11.93% in CLP, terminal). The beta relevers the median unlevered beta of six peers (0.94), and the CRP weights sovereign risk by revenue (71.8% Chile, 17.7% Peru and 10.5% Colombia). Kd starts from the latest UF bond coupon (3.6% real) plus 25 bp per turn of leverage above 2.36x: 4.92% after tax in 2H26E and 5.11% in the terminal period. At CLP 3,555, equity yields almost the same as that debt (Figure 28; details in Appendix [Z])."] },
          { p: ["Capital structure. ", "The company uses its debt capacity to grow (it reached 3.78x in 2024 after buying Open Plaza Peru). Since FCFF already covers capex, the extra debt is returned to shareholders as a distribution (CLP 1.79 trillion cumulative to 2031): it does not enter the EV, but it makes the debt assumed in the WACC exist in the plan (Appendix [D])."] },
          { h1: "DCF: OPERATING FORECAST" },
          { p: "Reported EBITDA (CLP 544,935 million LTM, 80.5% margin) excludes property revaluation (IAS 40); we forecast a 79.8% margin through 2031E. The operating tax rate rises from 16.8% in 2026E to 27% in 2030E as existing benefits fade. Base capex (pipeline plus maintenance at 5.5% of revenue) peaks at CLP 240,660 million in 2027 and falls to ~CLP 50,000 million from 2029. From 2029, 27,075 m² also open each year (historical average), with capex of CLP 95,000–104,000 million a year, a 35%/90%/100% ramp and a 6.9% yield on cost, at the cost of the identified pipeline (USD 3,603/m²: 170,200 m² and USD 613.2 million). Working capital (15.6% × Δ revenue) requires CLP 7,100–9,900 million a year and moves the price by only ±1.2% (Appendix [C])." },
          { h1: "DCF: TERMINAL VALUE AND IMPLIED CAP RATE" },
          { p: "We use g = 3.0% nominal (0% real). The real RONIC on expansion (5.0%) is below the real WACC (7.2%): growing beyond the horizon destroys value and a higher g lowers it (Figure 29). The DCF implies a cap rate of 10.1% on EV and 11.3% on terminal value, versus 6.0% for the market, 6.4% for Chilean peers, 7.5% at book value and 5.55% for the appraisers (Note 11) (Figure 30). That gap is not an inconsistency: it is the core of the thesis. To justify the current price, the WACC would have to be 285 bp lower (Figure 31)." },
        ],
        figuras: [
          ["Figure 27:", "No method justifies the current price (CLP per share)", "V1_football", "52-week range pending (Bloomberg). Source: Team analysis; Bloomberg."],
          ["Figure 28:", "At CLP 3,555 equity yields almost the same as debt", "V2_tasas", "Real rates in UF. Source: Central Bank of Chile (30-Jun-26); Mallplaza; Team analysis."],
          ["Figure 29:", "Growing below the WACC destroys value: real g = 0%", "V3_value_driver", "Top: DCF value by g. Bottom: relative terminal value by RONIC. Source: Team analysis."]] },
      { texto: [
          { h1: "RELATIVE VALUATION: MULTIPLES (20%)" },
          { p: "The set is the same as for the beta: Parque Arauco and Cencosud Shopping (Chile), Allos, Multiplan and Iguatemi (Brazil) and IRSA (Argentina, used with caution because of IAS 29 and a single NTM analyst). We exclude Fibra Danhos, which as a FIBRA distributes 95% of its taxable income (Appendix [E]). Brazil trades at lower multiples because of its rates and risk premia (median 9.0x EV/EBITDA versus 15.7x in Chile; Figure 32). Since 71.8% of revenue comes from Chile, the multiple weights the Chilean median at 70% and the full-set median (11.1x) at 30%: 14.4x." },
          { p: "We use EV/EBITDA because it is neutral to leverage and to revaluation (IAS 40). At 14.4x, value is CLP 2,951 per share. Mallplaza trades at 16.8x LTM: 7% above the Chilean median, below Parque Arauco (18.4x) and 17% above the weighted multiple (Figure 33). We exclude P/E (CLP 8,644, distorted by revaluation). Multiples use Bloomberg net debt (CLP 1,311,569 million), as for peers; the DCF uses the company's net financial debt (CLP 1,285,740 million)." },
          { h1: "VALUATION SENSITIVITY AND RISKS" },
          { p: ["Sensitivity. ", "With WACC ±100 bp and g between 2.5% and 4.0%, the DCF ranges from CLP 1,379 to 2,329 (Figure 25; Appendix [F]), and the historical worst operating case takes 10.0% off (Figure 24). Every multiple sits below the price: EV/EBITDA LTM CLP 2,951 (−17%), NTM CLP 2,914 (−18%), EV/EBIT CLP 3,098 (−13%), P/BV CLP 3,171 (−11%) and the Chilean median alone CLP 3,294 (−7%) (Figure 27)."] },
          { p: ["Robustness. ", "APV gives CLP 1,894 (2% above the DCF); with the peso bond as Rf, the target price rises to ~CLP 2,439 (−31%; Appendix [Z]); and with a net debt/EBITDA target between 2.36x and 4.0x, the DCF ranges from CLP 1,819 to 1,894."] },
          { p: ["Target price. ", "80% × CLP 1,857 + 20% × CLP 2,951 = CLP 2,076, a −41.6% return against CLP 3,555 (30-Sep-26). Rolling the DCF from 30-Jun to 30-Sep-26 would lift it to ~CLP 1,910 and the target to ~CLP 2,118 (−40.4%): the recommendation does not change."] },
          { p: ["Upside risks. ", "A sustained compression of cap rates on stronger institutional demand for defensive assets, or a structural premium for Mallplaza's Tier A assets (see Investment Risks)."] },
        ],
        figuras: [
          ["Figure 30:", "The market pays 6.0%; our cash flows are worth 10.1%", "V5_cap_rate", "Gran Plaza: subsequent event, not in the model. Source: Mallplaza; Bloomberg; DF; Team analysis."],
          ["Figure 31:", "The price requires a WACC 285 bp lower", "V6_dcf_inverso", "Source: Team analysis."],
          ["Figure 32:", "Peer multiples vs local rates", null, "Chart pending (Brazil NTN-B yield missing, Bloomberg)."],
          ["Figure 33:", "Value implied by each peer's multiple vs DCF and price (CLP per share)", "V9_pares", "Target = 80% DCF + 20% EV/EBITDA LTM. P/E excluded (includes revaluation). Source: Bloomberg (30-Sep-26); Team analysis."]] },
    ],
    anexo: {
      titulo: "APPENDIX [Z]: WACC CONSTRUCTION",
      lead: "The WACC is computed per period: D/E converges linearly from the current structure (20.0%) to the target (28.5%, net debt/EBITDA 3.5x) in 2031. The cost of equity is built in dollars (US Treasury Rf + re-levered beta × ERP + revenue-weighted CRP) and converted to pesos with the inflation differential.",
      formulas: ["βL = βU × [1 + (1 − t) × D/E]  (Hamada)", "Ke USD = Rf + βL × ERP + weighted CRP", "Ke CLP = (1 + Ke USD) × (1 + π Chile) / (1 + π US) − 1", "Kd CLP = (1 + real Kd + spread) × (1 + π Chile) − 1", "WACC = E/V × Ke CLP + D/V × Kd CLP × (1 − t)"],
      sA: "A. MARKET AND CAPITAL STRUCTURE ASSUMPTIONS",
      inputs: [["Rf: 10Y US Treasury", "rf", "%", "Tasa_WACC (FFCF)"], ["Equity risk premium (ERP)", "erp", "%", "Damodaran"],
        ["Long-term inflation, Chile", "pi_cl", "%", "Central Bank target"], ["Long-term inflation, US (10Y breakeven)", "pi_us", "%", "Bloomberg, Jun-26"],
        ["Chilean tax rate", "tax", "%", "Income Tax Law"], ["Real Kd: latest UF bond coupon", "kd_real", "%", "Mallplaza"],
        ["Real Kd increase per turn of net debt/EBITDA", "spread", "%", "Team assumption"], ["Net debt/EBITDA initial (30-Jun-26) → target", "dfn", "x", "2Q26 release; team assumption"],
        ["Gross debt / market cap (current structure, CLP MM)", "estructura", "", "Tasa_WACC (FFCF)"]],
      colsA: ["Assumption", "Value", "Source"],
      sB: "B. BETA: UNLEVERED PEERS",
      colsB: ["Company", "Levered β", "Debt (CLP MM)", "Market cap (CLP MM)", "Tax rate", "Unlevered β"],
      medianaB: "Median (6 peers included)", notaB: "Source: Bloomberg (betas), Tasa_WACC. Costa Verde is excluded because it is not in the comparables set.",
      sC: "C. REVENUE-WEIGHTED COUNTRY RISK",
      colsC: ["Country", "Weight (LTM revenue)", "CRP (Damodaran)", "Contribution"], paises: ["Chile", "Peru", "Colombia"], totalC: "Weighted CRP",
      sD: "D. WACC BY PERIOD",
      filasD: [["53", "Net debt/EBITDA", "x"], ["56", "D/E", "%"], ["57", "D/V", "%"], ["59", "Re-levered beta", "b"], ["60", "Ke USD", "%2"], ["61", "Nominal Ke CLP", "%2"],
        ["62", "Real Kd (UF) incl. spread", "%2"], ["63", "Nominal Kd CLP", "%2"], ["64", "After-tax Kd", "%2"], ["65", "Nominal WACC CLP", "%2"], ["66", "Real WACC", "%2"]],
      per: ["2H26E", "2027E", "2028E", "2029E", "2030E", "2031E", "Terminal"],
      fW: ["Figure Z1:", "Nominal WACC and D/E by period", "Control: the replica matches the WACC sheet of model v2 (0 bp difference in every period). Source: Damodaran; Bloomberg; Central Bank of Chile; Team analysis."],
      sF: "E. CHOICE OF THE RISK-FREE RATE",
      colsF: ["Alternative", "Value"],
      filasF: [["Model base: US Rf + Chile CRP, in CLP", "base", "%2"], ["10Y peso bond, Central Bank (30-Jun-26)", "bono", "%2"], ["Difference", "dif", "pb"],
        ["Effect on terminal WACC (difference × E/V)", "efecto", "pb"], ["DCF value using the peso bond (CLP)", "dcf", "clp"], ["80/20 target price in that case (CLP)", "po", "clp"]],
      notasF: ["Why UST + CRP and not the peso bond: the local bond already embeds Chile risk, so adding the CRP would double-count it; also, 28% of revenue comes from Peru and Colombia, whose risk enters through the weighted CRP.",
        "It is the more conservative choice: with the peso bond the target price rises to ~CLP 2,439, still 31% below the CLP 3,555 share price. The recommendation does not change.",
        "Source: Central Bank of Chile, Informativo Diario 2-Jul-26; DCF v2 replica with the WACC shifted in every period."],
    },
  },
};

const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || 16, bold: o.bold, italics: o.italics, color: o.color || TINTA });
const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: nada, bottom: nada, left: nada, right: nada };
const linea = (l = 228) => ({ line: l, lineRule: LineRuleType.AUTO });
const vacia = () => new Paragraph({ children: [] });
const titulo = (t) => new Paragraph({ spacing: { after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 2 } },
  children: [run(t, { bold: true, color: AZUL, size: 30 })] });
const subtitulo = (t, antes = 110) => new Paragraph({ spacing: { before: antes, after: 50 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ROJO, space: 1 } }, children: [run(t, { bold: true, color: AZUL })] });
const caption = (n, t, antes = 100) => new Paragraph({ spacing: { before: antes, after: 30, ...linea() },
  children: [run(n + " ", { bold: true, color: AZUL, size: 14 }), run(t, { color: AZUL, size: 14 })] });
const fuente = (t) => new Paragraph({ spacing: { before: 10, after: 0, ...linea(216) }, children: [run(t, { size: 11, color: GRIS })] });
const imagen = (archivo, anchoIn, altoIn) => new Paragraph({ spacing: { before: 0, after: 20 }, children: [
  new ImageRun({ type: "png", data: fs.readFileSync(path.join(FIG, archivo)), transformation: { width: Math.round(anchoIn * PX), height: Math.round(altoIn * PX) } })] });
const celda = (w, children, extra = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: sinBordes,
  margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP, children, ...extra });

// Página de la sección: texto a la izquierda, figuras (o recuadro de gráfico pendiente) a la derecha
const parrafo = (x) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 60, ...linea() },
  children: Array.isArray(x) ? [run(x[0], { bold: true, color: AZUL }), run(x[1])] : [run(x)] });
const pendiente = (t) => new Table({ width: { size: COL_FIG, type: WidthType.DXA }, columnWidths: [COL_FIG],
  rows: [new TableRow({ height: { value: 1700, rule: "atLeast" }, children: [new TableCell({ width: { size: COL_FIG, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, color: "auto", fill: "F7F8FB" },
    borders: Object.fromEntries(["top", "bottom", "left", "right"].map((k) => [k, { style: BorderStyle.DASHED, size: 6, color: "9AA3B5" }])),
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(t, { size: 14, italics: true, color: GRIS })] })] })] })] });

function pagina(lang, p) {
  const izq = [];
  p.texto.forEach((b, i) => izq.push(b.h1 ? subtitulo(b.h1, i === 0 ? 20 : 90) : parrafo(b.p)));
  const der = [];
  p.figuras.forEach(([n, t, archivo, f], k) => {
    der.push(caption(n, t, k === 0 ? 20 : 120));
    if (archivo) {
      der.push(imagen(`${lang}_${archivo}.png`, COL_FIG / 1440, ALTO[archivo.split("_")[0]]));
      der.push(fuente(f));
    } else {
      der.push(pendiente(f));
    }
  });
  return new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [COL_TEXTO, COL_SEP, COL_FIG],
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
    rows: [new TableRow({ cantSplit: true, children: [celda(COL_TEXTO, izq), celda(COL_SEP, [vacia()]), celda(COL_FIG, der)] })] });
}

function tabla(cols, anchos, filas, { derecha = [], destacar = [] } = {}) {
  const borde = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  const enc = new TableRow({ tableHeader: true, children: cols.map((c, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: AZUL }, borders: sinBordes, margins: { top: 35, bottom: 35, left: 80, right: 80 },
    children: [new Paragraph({ alignment: derecha.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [run(c, { bold: true, size: 13, color: "FFFFFF" })] })] })) });
  const cuerpo = filas.map((f, r) => new TableRow({ children: f.map((v, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: destacar.includes(r) ? "FDF1CC" : r % 2 ? "FFFFFF" : "F7F8FB" },
    borders: { ...sinBordes, bottom: borde }, margins: { top: 25, bottom: 25, left: 80, right: 80 }, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing: linea(216), alignment: derecha.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [run(String(v), { size: 13, bold: destacar.includes(r), color: TINTA })] })] })) }));
  return new Table({ width: { size: anchos.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: anchos,
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows: [enc, ...cuerpo] });
}

function anexo(lang) {
  const A = L[lang].anexo, f = fmt(lang), I = D.inputs;
  const pc = (x, dec = 2) => f(x * 100, dec) + "%";
  const val = (k, u) => k === "dfn" ? `${f(I.dfn0, 2)}x → ${f(I.dfn_obj, 1)}x` : k === "estructura" ? `${f(I.deuda / 1000, 0)} / ${f(I.mcap / 1000, 0)} mil` .replace(" mil", lang === "es" ? " (miles)" : " (thousands)") : pc(I[k]);
  const filasA = A.inputs.map(([n, k, u, s]) => [n, val(k, u), s]);
  const incl = D.betas.filter((b) => b[6] === 1);
  const filasB = incl.map((b) => [b[0], f(b[1], 3), f(b[2], 0), f(b[3], 0), pc(b[4], 0), f(b[5], 3)]);
  filasB.push([A.medianaB, "", "", "", "", f(D.beta_u, 4)]);
  const crps = [I.crp_cl, I.crp_pe, I.crp_co];
  const filasC = A.paises.map((p, i) => [p, pc(D.crp.w[i], 1), pc(crps[i]), pc(D.crp.w[i] * crps[i])]);
  filasC.push([A.totalC, "100,0%".replace(",", lang === "es" ? "," : "."), "", pc(D.crp.total)]);
  const filasD = A.filasD.map(([r, n, u]) => [n, ...D.periodos[r].map((x) => u === "x" ? f(x, 2) + "x" : u === "b" ? f(x, 3) : pc(x, u === "%" ? 1 : 2))]);
  const filasF = A.filasF.map(([n, k, u]) => [n, u === "pb" ? `${f(D.rf[k], 0)} ${lang === "es" ? "pb" : "bp"}` : u === "clp" ? f(D.rf[k], 0) : pc(D.rf[k])]);
  return [
    titulo(A.titulo),
    new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, ...linea() }, shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO }, children: [run(A.lead)] }),
    ...A.formulas.map((t) => new Paragraph({ spacing: { after: 10 }, indent: { left: 240 }, children: [run(t, { size: 14, color: AZUL, bold: true })] })),
    subtitulo(A.sA, 140), tabla(A.colsA, [4800, 2400, 3466], filasA, { derecha: [1] }),
    subtitulo(A.sB, 140), tabla(A.colsB, [2766, 1400, 1900, 2100, 1100, 1400], filasB, { derecha: [1, 2, 3, 4, 5], destacar: [filasB.length - 1] }), fuente(A.notaB),
    subtitulo(A.sC, 140), tabla(A.colsC, [3166, 2700, 2400, 2400], filasC, { derecha: [1, 2, 3], destacar: [filasC.length - 1] }),
    subtitulo(A.sD, 140), tabla(["", ...A.per], [2866, 1100, 1100, 1100, 1100, 1100, 1100, 1200], filasD, { derecha: [1, 2, 3, 4, 5, 6, 7], destacar: [9] }),
    new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [MEDIA, COL_SEP, MEDIA], borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
      rows: [new TableRow({ cantSplit: true, children: [
        celda(MEDIA, [caption(A.fW[0], A.fW[1], 140), imagen(`${lang}_W1_wacc_periodo.png`, MEDIA / 1440, ALTO.W1), fuente(A.fW[2])]),
        celda(COL_SEP, [vacia()]),
        celda(MEDIA, [subtitulo(A.sF, 140), tabla(A.colsF, [3725, 1500], filasF, { derecha: [1], destacar: [5] }),
          ...A.notasF.map((t, i) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: i ? 30 : 80, after: 0, ...linea(216) },
            children: [run(t, i === A.notasF.length - 1 ? { size: 11, color: GRIS } : { size: 14 })] }))]),
      ] })] }),
  ];
}

const encabezado = () => new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0", space: 2 } },
  children: [run("Mallplaza (Plaza S.A.)  |  CFA Institute Research Challenge", { size: 14, color: GRIS })] })] });
const pie = (p) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  run(p, { size: 14, color: GRIS }), new TextRun({ children: [PageNumber.CURRENT], font: FUENTE, size: 14, color: GRIS })] })] });

function documento(lang) {
  const X = L[lang];
  const pag = { size: { width: 11906, height: 16838 }, margin: { top: 1000, right: 620, bottom: 900, left: 620, header: 450, footer: 400 } };
  return new Document({
    styles: { default: { document: { run: { font: FUENTE, size: 16 } } } },
    sections: [
      { properties: { page: pag }, headers: { default: encabezado() }, footers: { default: pie(X.pag) },
        children: [titulo(X.titulo),
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, ...linea() }, shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO }, children: [run(X.lead)] }),
          ...X.paginas.flatMap((p, k) => k ? [new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: [] }), pagina(lang, p)] : [pagina(lang, p)])] },
      { properties: { page: { ...pag, pageNumbers: { start: 1 } } }, headers: { default: encabezado() }, footers: { default: pie(X.pagA) }, children: anexo(lang) },
    ],
  });
}

(async () => {
  for (const lang of ["es", "en"]) {
    const salida = path.join(__dirname, L[lang].archivo);
    fs.writeFileSync(salida, await Packer.toBuffer(documento(lang)));
    console.log(salida);
  }
})();
