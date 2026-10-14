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
          { p: "Asignamos 80% al DCF porque Mallplaza opera infraestructura comercial con flujos de caja muy predecibles: el 81,4% de los ingresos UDM son arriendos y, de ellos, el 93% es renta fija (~76% de los ingresos totales). El DCF es el método que mejor captura el valor intrínseco del negocio y su plan de expansión." },
          { p: "Los múltiplos (20%) contrastan el modelo con lo que el mercado paga hoy por la industria. Limitamos su peso por dos razones: (i) en Chile solo hay dos pares directos y líquidos (Parque Arauco y Cencosud Shopping), lo que da poca profundidad estadística; y (ii) los múltiplos reflejan el sentimiento actual, pero no la expansión de GLA ni la trayectoria de financiamiento de Mallplaza, que sí están en el DCF." },
          { h1: "DCF: ESTRUCTURA DEL MODELO" },
          { p: "Valoramos un flujo de caja libre para la firma (FCFF) en dos etapas, en CLP nominales, con fecha de valoración 30-jun-26 y convención de mitad de período. La etapa explícita (2S26E–2031E) recoge el plan de inversión de corto plazo y la expansión de GLA; la etapa terminal (desde 2032) es una perpetuidad de Gordon que explica el 76% del enterprise value, una proporción coherente con activos de larga duración (Anexo [B])." },
          { p: "El horizonte explícito incorpora el capex y la maduración de ingresos de los proyectos que abren hasta 2032, que suman 108.301 m² de GLA madura en el año terminal. Ese crecimiento se financia con la caja operativa y con más deuda: la DFN/EBITDA sube de 2,36x a 3,5x en 2031. Desde 2032 el modelo no agrega GLA neta y el capex terminal cubre solo la mantención (5,5% de los ingresos), porque expandirse con un retorno menor a la WACC real destruiría valor (ver Valor terminal)." },
          { h1: "DCF: TASA DE DESCUENTO (WACC)" },
          { p: "La WACC nominal en CLP cambia por período a medida que la estructura de capital converge a su objetivo: 10,54% en 2S26E y 10,41% en el terminal (7,32% y 7,20% en términos reales). Se construye con un CAPM ajustado por riesgo país (enfoque de Damodaran); el detalle está en el Anexo [Z]." },
          { p: ["Costo del patrimonio (Ke). ", "Como Rf, ERP y CRP están en dólares, el Ke se calcula primero en USD: Rf 4,98% (Tesoro de EE.UU. a 10 años) + beta 1,14 × ERP 4,14% (prima implícita de Damodaran) + CRP 1,41% = 11,10% en el terminal. La beta resulta de desapalancar con Hamada seis pares regionales (mediana 0,94) y re-apalancarla con la estructura de Mallplaza: 1,09 en 2S26E y 1,14 en el terminal. El CRP pondera el riesgo soberano por ingresos (71,8% Chile, 17,7% Perú y 10,5% Colombia). El Ke en USD se pasa a CLP con el diferencial de inflación esperada: 11,70% en 2S26E y 11,93% en el terminal."] },
          { p: ["Costo de la deuda (Kd). ", "Parte del cupón de la última emisión de bonos en UF (3,6% real), llevado a nominal con la meta de inflación del BCCh (3,0%), más 25 pb por cada vuelta de DFN/EBITDA sobre 2,36x (supuesto del equipo). Después de impuestos (27%), sube de 4,92% en 2S26E a 5,11% en el terminal. A CLP 3.555 el patrimonio rinde casi lo mismo que esta deuda (Figura 28)."] },
          { p: ["Estructura de capital. ", "El objetivo de largo plazo es una DFN/EBITDA de 3,5x, alcanzada en forma lineal desde 2,36x hacia 2031: la compañía usa su capacidad de deuda para crecer, como en 2024 (3,78x tras la compra de Open Plaza Perú). A valor de mercado equivale a 77,8% patrimonio y 22,2% deuda (D/E 28,5%); por eso la WACC cambia en el tiempo. La caja que no absorben las inversiones se devuelve como distribución adicional a los accionistas (CLP 1,79 billones acumulados a 2031): no entra en el EV, pero hace que la deuda supuesta en la WACC exista en el plan (Anexo [D])."] },
        ],
        figuras: [
          ["Figura 27:", "Ningún método justifica el precio actual (CLP por acción)", "V1_football", "Rango 52 semanas pendiente (Bloomberg). Fuente: análisis del equipo; Bloomberg."],
          ["Figura 28:", "A CLP 3.555 el patrimonio rinde casi lo mismo que la deuda", "V2_tasas", "Tasas reales en UF. Fuente: BCCh (30-06-26); Mallplaza; análisis del equipo."]] },
      { texto: [
          { h1: "DCF: PROYECCIÓN OPERATIVA Y FCFF" },
          { p: ["EBITDA sin revalorización. ", "El EBITDA reportado (CLP 544.935 MM UDM, margen de 80,5%) excluye las ganancias no realizadas por revalorización de propiedades de inversión (NIC 40). Proyectamos un margen de 79,8% en 2026E, ~70 pb bajo el UDM, que se mantiene en ese nivel hasta 2031E (Anexo [C])."] },
          { p: ["Impuesto operacional. ", "La tasa efectiva sobre el EBIT parte en 16,8% en 2026E (impuesto corriente sobre EBITDA menos costo financiero, que incorpora beneficios tributarios vigentes) y converge a la tasa corporativa de 27% en 2030E: 16,8% → 19,0% → 21,5% → 24,0% → 27,0%. Suponemos que esos beneficios se diluyen para no sobreestimar el valor."] },
          { p: ["Capex. ", "El capex base (pipeline publicado más mantención y TI, de 5,5% de los ingresos) llega a un máximo de CLP 240.660 MM en 2027 y se estabiliza en CLP 48.000–52.000 MM anuales desde 2029. El capex de expansión suma CLP 46.268 MM en 2027 y CLP 95.000–104.000 MM anuales en 2028–2031: desde 2029 abren 27.075 m² al año (promedio del GLA orgánico agregado en 2022, 2023 y 2025), con el capex en los dos años previos, una rampa de ingresos de 35%/90%/100% y un yield on cost de 6,9%. El costo por m² (USD 3.603) sale del pipeline identificado: 170.200 m² y USD 613,2 MM (Release 2T26, p. 13)."] },
          { p: ["Capital de trabajo. ", "Mallplaza mantiene un capital de trabajo neto positivo (CLP 102.039 MM, 15,6% de los ingresos a dic-25), porque las cuentas por cobrar a locatarios (CLP 122.370 MM) superan los arriendos cobrados por anticipado (CLP 20.247 MM). La inversión incremental (15,6% × Δ ingresos) es baja y predecible: CLP 7.100–9.900 MM al año, entre 0,7% y 1,3% de los ingresos. Cambiar este supuesto mueve el precio en ±1,2% (Figura 29)."] },
          { h1: "DCF: VALOR TERMINAL Y CAP RATE IMPLÍCITO" },
          { p: ["Crecimiento terminal. ", "Usamos g = 3,0% nominal (0% real), igual a la meta de inflación del BCCh. Lo justifica el value driver: el RONIC real de la expansión (5,0%) es menor que la WACC real (7,2%), de modo que seguir expandiéndose después del horizonte destruiría valor y un g mayor reduce el valor (Figura 30). Por eso la reinversión terminal es cero y los arriendos crecen solo con la inflación."] },
          { p: ["Cap rate implícito. ", "El DCF implica un cap rate real de 10,1% sobre el enterprise value y de 11,3% sobre el valor terminal, muy por sobre el del mercado (6,0%), el de los pares chilenos (6,4%), el cap rate a valor libro (7,5%) y la tasa real que usan los tasadores independientes en la Nota 11 de los EEFF (5,55%, punto medio del rango) (Figura 31). Esta brecha no es una inconsistencia que resolver: es el núcleo de la tesis. Nuestro modelo exige a los flujos mucho más que el mercado y los tasadores; para justificar el precio actual, la WACC tendría que ser 285 pb menor (Figura 32)."] },
        ],
        figuras: [
          ["Figura 29:", "Capital de trabajo: un supuesto que no mueve el precio (±1,2%)", "V4_capital_trabajo", "Fuente: EEFF Mallplaza; análisis del equipo."],
          ["Figura 30:", "Crecer bajo la WACC destruye valor: g real = 0%", "V3_value_driver", "Arriba: valor DCF según g. Abajo: valor terminal relativo según el RONIC. Fuente: análisis del equipo."],
          ["Figura 31:", "El mercado paga 6,0%; nuestros flujos valen 10,1%", "V5_cap_rate", "Gran Plaza: hecho posterior, no entra al modelo. Fuente: Mallplaza; Bloomberg; DF; análisis del equipo."]] },
      { texto: [
          { h1: "VALUACIÓN RELATIVA: MÚLTIPLOS (20%)" },
          { p: ["Peer group. ", "Es el mismo set que define la beta: dos pares chilenos (Parque Arauco y Cencosud Shopping), tres brasileños (Allos, Multiplan e Iguatemi) e IRSA (Argentina), elegidos por modelo de negocio, escala y calidad de los activos de arriendo. IRSA se mantiene con cautela, porque reporta bajo NIC 29 (hiperinflación) y su consenso NTM tiene un solo analista. Excluimos Fibra Danhos (México): como FIBRA debe distribuir el 95% de su resultado fiscal, lo que altera sus múltiplos, y no tiene beta en la WACC (Anexo [E])."] },
          { p: "Los pares brasileños transan a múltiplos estructuralmente más bajos por las mayores tasas y primas de riesgo de Brasil (mediana de 9,0x EV/EBITDA frente a 15,7x en Chile; Figura 33). Aplicar el set completo sin ajuste castigaría a Mallplaza por un riesgo ajeno, cuando el 71,8% de sus ingresos viene de Chile. Por eso el múltiplo pondera 70% la mediana chilena y 30% la del set completo." },
          { p: ["Múltiplo principal. ", "Usamos EV/EBITDA porque es neutral a la estructura de deuda y deja fuera la revalorización de propiedades (NIC 40): mide solo la caja operativa del GLA arrendado. El múltiplo objetivo es 14,4x (70% × 15,7x + 30% × 11,1x) y lleva a CLP 2.951 por acción. Mallplaza transa a 16,8x LTM: 7% sobre la mediana chilena, bajo Parque Arauco (18,4x) y 17% sobre el múltiplo ponderado (Figura 34). Los múltiplos confirman la dirección (la acción está cara); la magnitud del downside viene del DCF."] },
          { p: "Excluimos el P/E porque su valor implícito (CLP 8.644) está distorsionado por la revalorización en la utilidad. Para ser consistentes con los pares, los múltiplos usan la deuda neta de Bloomberg (CLP 1.311.569 MM), levemente distinta de la DFN de la compañía que usa el DCF (CLP 1.285.740 MM)." },
          { h1: "SENSIBILIDAD Y RIESGOS DE LA VALUACIÓN" },
          { p: ["Sensibilidad. ", "Con la WACC ±100 pb y un g entre 2,5% y 4,0%, el valor DCF va de CLP 1.379 a 2.329, con centro en CLP 1.857 (Figura 25; Anexo [F]). Incluso el extremo más favorable queda muy bajo el precio de CLP 3.555. En la operación, el peor caso histórico (margen EBITDA de 75,8%) resta 10,0% al DCF (Figura 24)."] },
          { p: ["Dispersión por múltiplos. ", "Los múltiplos son el escenario más favorable para la acción y aun así todos quedan bajo el precio: EV/EBITDA LTM CLP 2.951 (−17%), EV/EBITDA NTM CLP 2.914 (−18%), EV/EBIT CLP 3.098 (−13%) y P/BV CLP 3.171 (−11%), al igual que la mediana chilena por sí sola (CLP 3.294, −7%) (Figura 27). Que la acción esté cara no depende solo del DCF."] },
          { p: ["Chequeos de robustez. ", "(i) El APV da CLP 1.894 por acción, 2% sobre el DCF; (ii) con el bono en pesos como tasa libre de riesgo, el precio objetivo sube a ~CLP 2.439, todavía 31% bajo el precio (Anexo [Z]); y (iii) con un objetivo de DFN/EBITDA entre 2,36x y 4,0x y un spread de 0 a 50 pb por vuelta, el DCF se mueve entre CLP 1.819 y 1.894."] },
          { p: ["Precio objetivo. ", "80% × CLP 1.857 + 20% × CLP 2.951 = CLP 2.076, un retorno esperado de −41,6% frente al cierre de CLP 3.555 del 30-sep-26. El DCF está a fecha 30-jun-26; llevarlo a la fecha del precio (capitalizando al Ke por 0,25 años) lo subiría a ~CLP 1.910 y el precio objetivo a ~CLP 2.118 (−40,4%), sin cambiar la recomendación."] },
          { p: ["Riesgos al alza. ", "Una compresión adicional y sostenida de los cap rates de mercado, impulsada por mayor apetito institucional por activos inmobiliarios defensivos en la región, o una prima estructural para los activos Tier A de Mallplaza si su calidad y poder de fijación de precios resultan más duraderos de lo que asume el modelo (ver Riesgos de inversión)."] },
        ],
        figuras: [
          ["Figura 32:", "El precio exige una WACC 285 pb menor", "V6_dcf_inverso", "Fuente: análisis del equipo."],
          ["Figura 33:", "Múltiplos de los pares vs tasa local", null, "Gráfico pendiente (falta la NTN-B de Brasil, Bloomberg)."],
          ["Figura 34:", "Valor implícito por múltiplo de cada par vs DCF y precio (CLP por acción)", "V9_pares", "Precio objetivo = 80% DCF + 20% EV/EBITDA LTM. P/E excluido (incluye revalorización de propiedades). Set de pares = empresas con beta en la WACC. Fuente: Bloomberg (precios al 30-09-26); análisis del equipo."]] },
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
          { p: "We give the DCF an 80% weight because Mallplaza runs retail infrastructure with highly predictable cash flows: 81.4% of LTM revenue is rent, and 93% of that is fixed rent (~76% of total revenue). The DCF best captures the intrinsic value of the business and its expansion plan." },
          { p: "Multiples (20%) test the model against what the market pays for the industry today. We cap their weight for two reasons: (i) Chile has only two direct, liquid peers (Parque Arauco and Cencosud Shopping), which gives little statistical depth; and (ii) multiples reflect current sentiment but not Mallplaza's GLA expansion or financing path, which the DCF does capture." },
          { h1: "DCF: MODEL STRUCTURE" },
          { p: "We value a two-stage free cash flow to the firm (FCFF) in nominal CLP, with a 30-Jun-26 valuation date and the mid-year convention. The explicit stage (2H26E–2031E) captures the near-term investment plan and the GLA expansion; the terminal stage (from 2032) is a Gordon perpetuity that accounts for 76% of enterprise value, consistent with long-lived assets (Appendix [B])." },
          { p: "The explicit horizon includes the capex and revenue ramp-up of projects opening through 2032, adding 108,301 m² of mature GLA in the terminal year. That growth is funded with operating cash flow and more debt: net debt/EBITDA rises from 2.36x to 3.5x by 2031. From 2032 the model adds no net GLA and terminal capex covers maintenance only (5.5% of revenue), because expanding at a return below the real WACC would destroy value (see Terminal value)." },
          { h1: "DCF: DISCOUNT RATE (WACC)" },
          { p: "The nominal CLP WACC changes by period as the capital structure converges to its target: 10.54% in 2H26E and 10.41% in the terminal period (7.32% and 7.20% in real terms). It follows a country-risk-adjusted CAPM (Damodaran approach); details are in Appendix [Z]." },
          { p: ["Cost of equity (Ke). ", "Since Rf, ERP and CRP are in dollars, Ke is first computed in USD: Rf 4.98% (10-year US Treasury) + beta 1.14 × ERP 4.14% (Damodaran implied premium) + CRP 1.41% = 11.10% in the terminal period. The beta comes from unlevering six regional peers with Hamada (median 0.94) and relevering it with Mallplaza's structure: 1.09 in 2H26E and 1.14 in the terminal period. The CRP weights sovereign risk by revenue (71.8% Chile, 17.7% Peru and 10.5% Colombia). USD Ke is converted to CLP with the expected inflation differential: 11.70% in 2H26E and 11.93% in the terminal period."] },
          { p: ["Cost of debt (Kd). ", "It starts from the coupon of the latest UF bond issue (3.6% real), converted to nominal with the Central Bank's 3.0% inflation target, plus 25 bp for each turn of net debt/EBITDA above 2.36x (team assumption). After tax (27%), it rises from 4.92% in 2H26E to 5.11% in the terminal period. At CLP 3,555, equity yields almost the same as this debt (Figure 28)."] },
          { p: ["Capital structure. ", "The long-term target is 3.5x net debt/EBITDA, reached linearly from 2.36x by 2031: the company uses its debt capacity to grow, as in 2024 (3.78x after the Open Plaza Peru acquisition). At market value this means 77.8% equity and 22.2% debt (D/E 28.5%), which is why the WACC moves over time. Cash not absorbed by investment is returned to shareholders as an additional distribution (CLP 1.79 trillion cumulative to 2031): it does not enter the EV, but it makes the debt assumed in the WACC exist in the plan (Appendix [D])."] },
        ],
        figuras: [
          ["Figure 27:", "No method justifies the current price (CLP per share)", "V1_football", "52-week range pending (Bloomberg). Source: Team analysis; Bloomberg."],
          ["Figure 28:", "At CLP 3,555 equity yields almost the same as debt", "V2_tasas", "Real rates in UF. Source: Central Bank of Chile (30-Jun-26); Mallplaza; Team analysis."]] },
      { texto: [
          { h1: "DCF: OPERATING FORECAST AND FCFF" },
          { p: ["EBITDA excluding revaluation. ", "Reported EBITDA (CLP 544,935 million LTM, 80.5% margin) excludes unrealised gains from investment property revaluation (IAS 40). We forecast a 79.8% margin in 2026E, ~70 bp below LTM, held at that level through 2031E (Appendix [C])."] },
          { p: ["Operating tax. ", "The effective rate on EBIT starts at 16.8% in 2026E (current tax over EBITDA less financial cost, which includes existing tax benefits) and converges to the 27% corporate rate in 2030E: 16.8% → 19.0% → 21.5% → 24.0% → 27.0%. We assume those benefits fade so as not to overstate value."] },
          { p: ["Capex. ", "Base capex (announced pipeline plus maintenance and IT at 5.5% of revenue) peaks at CLP 240,660 million in 2027 and settles at CLP 48,000–52,000 million a year from 2029. Expansion capex totals CLP 46,268 million in 2027 and CLP 95,000–104,000 million a year in 2028–2031: from 2029, 27,075 m² open each year (average organic GLA added in 2022, 2023 and 2025), with capex in the two prior years, a 35%/90%/100% revenue ramp and a 6.9% yield on cost. The cost per m² (USD 3,603) comes from the identified pipeline: 170,200 m² and USD 613.2 million (2Q26 Release, p. 13)."] },
          { p: ["Working capital. ", "Mallplaza has positive net working capital (CLP 102,039 million, 15.6% of revenue at Dec-25), because tenant receivables (CLP 122,370 million) exceed rents collected in advance (CLP 20,247 million). Incremental investment (15.6% × Δ revenue) is small and predictable: CLP 7,100–9,900 million a year, 0.7%–1.3% of revenue. Changing this assumption moves the price by ±1.2% (Figure 29)."] },
          { h1: "DCF: TERMINAL VALUE AND IMPLIED CAP RATE" },
          { p: ["Terminal growth. ", "We use g = 3.0% nominal (0% real), equal to the Central Bank's inflation target. The value driver supports it: the real RONIC on expansion (5.0%) is below the real WACC (7.2%), so expanding beyond the horizon would destroy value and a higher g lowers value (Figure 30). Terminal reinvestment is therefore zero and rents grow only with inflation."] },
          { p: ["Implied cap rate. ", "The DCF implies a real cap rate of 10.1% on enterprise value and 11.3% on terminal value, well above the market's (6.0%), Chilean peers' (6.4%), the book cap rate (7.5%) and the real rate used by independent appraisers in Note 11 of the financial statements (5.55%, mid-point of the range) (Figure 31). This gap is not an inconsistency to resolve: it is the core of the thesis. Our model asks far more of the cash flows than the market and the appraisers do; to justify the current price, the WACC would have to be 285 bp lower (Figure 32)."] },
        ],
        figuras: [
          ["Figure 29:", "Working capital: an assumption that does not move the price (±1.2%)", "V4_capital_trabajo", "Source: Mallplaza financial statements; Team analysis."],
          ["Figure 30:", "Growing below the WACC destroys value: real g = 0%", "V3_value_driver", "Top: DCF value by g. Bottom: relative terminal value by RONIC. Source: Team analysis."],
          ["Figure 31:", "The market pays 6.0%; our cash flows are worth 10.1%", "V5_cap_rate", "Gran Plaza: subsequent event, not in the model. Source: Mallplaza; Bloomberg; DF; Team analysis."]] },
      { texto: [
          { h1: "RELATIVE VALUATION: MULTIPLES (20%)" },
          { p: ["Peer group. ", "It is the same set that defines the beta: two Chilean peers (Parque Arauco and Cencosud Shopping), three Brazilian ones (Allos, Multiplan and Iguatemi) and IRSA (Argentina), chosen for business model, scale and quality of rental assets. IRSA is kept with caution, as it reports under IAS 29 (hyperinflation) and its NTM consensus has a single analyst. We exclude Fibra Danhos (Mexico): as a FIBRA it must distribute 95% of its taxable income, which distorts its multiples, and it has no beta in the WACC (Appendix [E])."] },
          { p: "Brazilian peers trade at structurally lower multiples because of Brazil's higher rates and risk premia (median 9.0x EV/EBITDA versus 15.7x in Chile; Figure 33). Applying the full set unadjusted would penalise Mallplaza for a risk it does not carry, when 71.8% of its revenue comes from Chile. The multiple therefore weights the Chilean median at 70% and the full-set median at 30%." },
          { p: ["Main multiple. ", "We use EV/EBITDA because it is neutral to leverage and leaves out property revaluation (IAS 40): it measures only the operating cash of leased GLA. The target multiple is 14.4x (70% × 15.7x + 30% × 11.1x), which gives CLP 2,951 per share. Mallplaza trades at 16.8x LTM: 7% above the Chilean median, below Parque Arauco (18.4x) and 17% above the weighted multiple (Figure 34). Multiples confirm the direction (the stock is expensive); the size of the downside comes from the DCF."] },
          { p: "We exclude P/E because its implied value (CLP 8,644) is distorted by revaluation gains in net income. To stay consistent with peers, multiples use Bloomberg net debt (CLP 1,311,569 million), slightly different from the company's net financial debt used in the DCF (CLP 1,285,740 million)." },
          { h1: "VALUATION SENSITIVITY AND RISKS" },
          { p: ["Sensitivity. ", "With WACC ±100 bp and g between 2.5% and 4.0%, DCF value ranges from CLP 1,379 to 2,329, centred on CLP 1,857 (Figure 25; Appendix [F]). Even the most favourable end is far below the CLP 3,555 price. On operations, the historical worst case (75.8% EBITDA margin) takes 10.0% off the DCF (Figure 24)."] },
          { p: ["Dispersion across multiples. ", "Multiples are the most favourable scenario for the stock, and even so all of them sit below the price: EV/EBITDA LTM CLP 2,951 (−17%), EV/EBITDA NTM CLP 2,914 (−18%), EV/EBIT CLP 3,098 (−13%) and P/BV CLP 3,171 (−11%), as does the Chilean median on its own (CLP 3,294, −7%) (Figure 27). The stock being expensive does not depend on the DCF alone."] },
          { p: ["Robustness checks. ", "(i) APV gives CLP 1,894 per share, 2% above the DCF; (ii) using the peso bond as the risk-free rate, the target price rises to ~CLP 2,439, still 31% below the price (Appendix [Z]); and (iii) with a net debt/EBITDA target between 2.36x and 4.0x and a spread of 0 to 50 bp per turn, the DCF moves between CLP 1,819 and 1,894."] },
          { p: ["Target price. ", "80% × CLP 1,857 + 20% × CLP 2,951 = CLP 2,076, an expected return of −41.6% against the CLP 3,555 close of 30-Sep-26. The DCF is dated 30-Jun-26; rolling it to the price date (compounding at Ke for 0.25 years) would lift it to ~CLP 1,910 and the target price to ~CLP 2,118 (−40.4%), without changing the recommendation."] },
          { p: ["Upside risks. ", "A further, sustained compression of market cap rates, driven by stronger institutional appetite for defensive real estate in the region, or a structural premium for Mallplaza's Tier A assets if their quality and pricing power prove more durable than the model assumes (see Investment Risks)."] },
        ],
        figuras: [
          ["Figure 32:", "The price requires a WACC 285 bp lower", "V6_dcf_inverso", "Source: Team analysis."],
          ["Figure 33:", "Peer multiples vs local rates", null, "Chart pending (Brazil NTN-B yield missing, Bloomberg)."],
          ["Figure 34:", "Value implied by each peer's multiple vs DCF and price (CLP per share)", "V9_pares", "Target = 80% DCF + 20% EV/EBITDA LTM. P/E excluded (includes property revaluation). Peer set = companies with a WACC beta. Source: Bloomberg (prices at 30-Sep-26); Team analysis."]] },
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
