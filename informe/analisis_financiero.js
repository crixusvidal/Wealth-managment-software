// Sección Financial Analysis (2 páginas) + Anexos 1–6 (2 páginas), en español e inglés.
// Fuente: borrador del equipo "Mallplaza_Financial_Analysis_2T26_v2" (cuadrado con el modelo de valoración v2).
// Figuras: graficos/informe/financiero/ (python graficos/analisis_financiero.py).
// Uso: node informe/analisis_financiero.js
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, ImageRun, LineRuleType, Packer, PageNumber,
  Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

const AZUL = "14284B", ROJO = "E9004B", TINTA = "1A1A1A", GRIS = "5F5F5F", FONDO = "EEF1F6";
const FUENTE = "Arial", ANCHO = 10666, COL_TEXTO = 6900, COL_SEP = 216, COL_FIG = 3550, MEDIA = 5225;
const FIG = path.join(__dirname, "..", "graficos", "informe", "financiero");
const PX = 96;
const ALTO = { FA2: 1.85, FA3: 1.95, FA4: 1.75, FA5: 1.6, FA6: 1.7, FA7: 2.2 }; // = graficos/analisis_financiero.py

// Cifras en formato inglés; en español se intercambian punto y coma
const es = (x) => String(x).replace(/[.,]/g, (c) => (c === "." ? "," : "."));
const loc = (lang) => (x) => (lang === "es" ? es(x) : String(x));

// ---------------------------------------------------------------------------------------------
// Tablas numéricas (comunes a ambos idiomas)
const FA1_AÑOS = ["’22 A", "’23 A", "’24 A", "’25 A", "’26 E", "’27 E", "’28 E", "’29 E", "’30 E", "’31 E"];
const FA1 = [
  ["371", "420", "495", "654", "704", "767", "821", "870", "916", "961"],
  ["76.2", "76.7", "77.0", "80.5", "79.7", "79.7", "79.7", "79.8", "79.8", "79.8"],
  ["59.0", "61.3", "61.3", "62.0", "60.9", "60.0", "57.7", "55.3", "52.7", "52.2"],
  ["158", "350", "336", "1,431", "639", "592", "610", "624", "627", "649"],
  ["96", "198", "238", "369", "380", "428", "434", "437", "432", "445"],
  ["25.8", "47.0", "48.2", "56.5", "54.0", "55.7", "52.9", "50.2", "47.1", "46.2"],
  [".086", ".090", ".092", ".097", ".089", ".091", ".092", ".093", ".093", ".094"],
  ["1.94", "1.91", "1.86", "1.77", "1.69", "1.71", "1.74", "1.79", "1.84", "1.89"],
  ["7.1", "14.4", "11.7", "37.5", "13.7", "12.0", "11.9", "11.8", "11.8", "12.0"],
  ["4.3", "8.1", "8.3", "9.7", "8.1", "8.6", "8.4", "8.3", "8.1", "8.2"],
  ["2.2", "4.2", "4.5", "5.5", "4.8", "5.1", "4.8", "4.6", "4.4", "4.3"],
];
const A2_IS = [["370.9", "419.9", "494.6", "653.9", "677.3"], ["282.7", "322.2", "380.6", "526.2", "544.9"], ["(2.2)", "(2.2)", "(2.1)", "(3.6)", "(4.8)"],
  ["83.1", "193.7", "157.3", "1,445.4", "1,190.4"], ["363.6", "513.8", "535.9", "1,968.1", "1,730.6"], ["—", "(23.8)", "(30.9)", "(54.0)", "(44.9)"],
  ["—", "(55.3)", "(53.2)", "(43.8)", "(50.4)"], ["157.7", "349.8", "335.9", "1,431.0", "1,258.1"], ["218.9", "257.2", "303.2", "405.3", "419.9"]];
const A2_BS = [["202.0", "179.2", "268.3", "347.6"], ["4,166.5", "5,356.5", "6,961.9", "7,309.8"], ["4,796.3", "5,915.1", "7,583.6", "7,985.7"],
  ["1,264.0", "1,617.6", "1,579.3", "1,633.4"], ["656.5", "892.9", "1,248.8", "1,313.7"], ["2,601.1", "3,157.9", "4,482.9", "4,820.9"]];
const A2_CF = [["295.1", "300.8", "441.1", "258.2"], ["(119.1)", "(76.3)", "(103.6)", "(84.2)"], ["(164.6)", "(400.4)", "(0.3)", "—"], ["—", "308.4", "—", "—"],
  ["(124.0)", "1.3", "(86.5)", "5.5"], ["(55.6)", "(54.4)", "(68.2)", "(30.1)"], ["(0.0)", "(125.8)", "(108.5)", "(87.6)"]];
const A3 = [["8.1%", "8.3%", "9.7%", "—"], ["4.2%", "4.5%", "5.5%", "—"], ["6.5%", "6.6%", "7.3%", "—"], ["1.78", "1.18", "1.54", "2.50"], ["0.83", "0.52", "0.87", "1.71"],
  ["0.82", "0.85", "0.68", "0.64"], ["25.5%", "26.9%", "18.8%", "17.6%"], ["3.30", "3.78", "2.49", "2.36"], ["6.9", "6.8", "7.7", "8.4"], ["0.090", "0.092", "0.097", "—"],
  ["23.8%", "23.5%", "20.1%", "21.2%"], ["65", "n.m.", "66", "50"]];
const A4 = [["365.8", "515.9", "538.0", "1,971.7", "1,735.4"], ["(85.4)", "(195.6)", "(162.5)", "(1,448.9)", "(1,194.5)"], ["2.3", "1.8", "5.2", "3.5", "4.1"],
  ["282.7", "322.2", "380.6", "526.2", "544.9"], ["0.0", "0.0", "0.0", "0.0", "0.0"]];
const A5 = [["498.4", "533.9", "571.2", "603.1", "628.7", "652.8"], ["131.3", "145.2", "159.8", "171.6", "178.6", "184.9"], ["75.2", "88.1", "89.4", "92.2", "96.8", "101.8"],
  ["(1.0)", "—", "—", "—", "—", "—"], ["—", "—", "—", "3.1", "11.7", "21.8"], ["704.0", "767.2", "820.5", "870.1", "915.9", "961.3"], ["7.7%", "9.0%", "7.0%", "6.0%", "5.3%", "5.0%"],
  ["561.3", "611.4", "654.2", "694.0", "730.6", "766.9"], ["79.7%", "79.7%", "79.7%", "79.8%", "79.8%", "79.8%"], ["(5.0)", "(5.4)", "(5.8)", "(6.1)", "(6.4)", "(6.8)"],
  ["556.4", "606.0", "648.5", "687.9", "724.1", "760.1"], ["16.8%", "19.0%", "21.5%", "24.0%", "27.0%", "27.0%"], ["(209.2)", "(240.7)", "(151.5)", "(47.7)", "(49.7)", "(51.7)"],
  ["—", "(46.3)", "(95.3)", "(98.2)", "(101.1)", "(104.2)"], ["(3.9)", "(9.9)", "(8.3)", "(7.7)", "(7.1)", "(7.1)"], ["428.5", "460.4", "473.1", "481.5", "482.3", "501.6"],
  ["380.0", "427.7", "434.4", "436.9", "431.6", "444.6"], ["(147.7)", "(152.0)", "(171.1)", "(173.8)", "(174.8)", "(172.6)"], ["(121.0)", "(236.1)", "(268.0)", "(380.1)", "(378.6)", "(404.8)"],
  ["2.46", "2.67", "2.88", "3.09", "3.29", "3.50"]];
const A6 = [["27.0%", "2.0%", "2.3%"], ["7.3% (Chile)", "7.6%", "7.8%"], ["37.5%", "9.8%", "10.5%"], ["9.7%", "7.8%", "8.0%"], ["18.8%", "42.9%", "16.3%"], ["2.49", "5.68", "2.11"], ["1.69", "2.69", "1.51"]];

// ---------------------------------------------------------------------------------------------
const L = {
  es: {
    archivo: "Mallplaza_Analisis_Financiero_ES.docx", pag: "Página ", pagA: "A-",
    titulo: "ANÁLISIS FINANCIERO", tituloA: "ANEXOS: ANÁLISIS FINANCIERO",
    lead: "El motor de arriendos de Mallplaza crece cerca de 3% real, con un margen EBITDA de 80% y un apalancamiento sobre flujos bajo 2,5x; la utilidad reportada, el crecimiento y el apalancamiento contable sobrestiman ese motor y deben normalizarse antes de proyectar.",
    fa1: ["Figura 1:", "El ROE normalizado de 8–10% es el retorno real; el ROE reportado depende de la tasación",
      ["CLP mil MM", "Ingresos", "Margen EBITDA %", "Margen FFO %", "Utilidad neta, reportada", "Utilidad distribuible", "Margen neto, normalizado %", "Rotación de activos (x)", "Activos / patrimonio (x)", "ROE, reportado %", "ROE, normalizado %", "ROA, normalizado %"],
      "Normalizado = utilidad distribuible (2022A: proxy del equipo; el reajuste UF de 2022 refleja ~13% de inflación). Saldos promedio; la serie parte en 2022 (NIC 40 a valor razonable desde el 1-ene-22). 2026E–2031E: caso base de valoración; balance proyectado desde dic-25 (patrimonio + utilidad reportada − distribuciones ordinarias y adicionales; propiedades revalorizadas al 3% de inflación, impuesto diferido 27%; DFN hasta 3,5x EBITDA en 2031E). Fuente: EEFF Plaza S.A. (CMF), Nota 20e; release 2T26 (FFO); análisis del equipo."],
    paginas: [
      { cortes: [[0, [0]], [3, [1]]], texto: [
          { h1: "La utilidad reportada sigue al tasador; el FFO sigue al negocio" },
          { p: "Bajo NIC 40, la revalorización de las propiedades de inversión pasa por resultados. En 2025, una ganancia por valor razonable de CLP 1.447 mil MM (27% del valor inicial de las propiedades, frente a cerca de 2% al año en Parque Arauco y Cencosud Shopping) elevó la utilidad atribuible a CLP 1.431 mil MM (2,2 veces los ingresos) y el ROE reportado a 37,5% (Figura 1). Cerca de dos tercios del alza de 29% en Chile vino de una menor tasa de descuento en la tasación, no de un mayor NOI." },
          { p: "Sin las ganancias por valor razonable netas de impuesto diferido, la utilidad distribuible auditada fue CLP 369 mil MM y el ROE 9,7%, el más alto de los tres pares locales en base comparable (Anexo 6). El FFO, que reconstruimos desde los estados financieros con 0,02% de diferencia frente a los CLP 405 mil MM reportados (Figura 2), mantuvo un margen de 59–62% en 2022–2025, mientras la utilidad neta varió más de cuatro veces. El FFO excluye el reajuste UF de la deuda (CLP 44 mil MM, 11% del FFO); la utilidad distribuible lo incluye y es nuestra ancla conservadora de utilidades." },
          { h1: "Dos tercios del crecimiento de 2025 fueron comprados; el orgánico real es cerca de 3%" },
          { p: "Los ingresos subieron 76% entre 2022 y 2025, pero 21,8 pp del 32,2% de 2025 vinieron del perímetro: el portafolio peruano comprado a Falabella, consolidado desde el 4 de diciembre de 2024. Sin Perú, los ingresos crecieron 10,7% en 2025 y 7,1% en 1S26." },
          { p: "Sin moneda ni inflación, el crecimiento reportado de 7,4% en 1S26 baja a cerca de 3,0% real (Figura 3). Chile creció 2,8% tras un reajuste UF de 3,1% (5,1% en 2025); Perú, 4,5% pese a que la ocupación bajó a 93,8%; y Colombia solo cerca de 1% una vez descontadas una apreciación del COP de 7,5% y una inflación de 6,1%." },
          { p: "El crecimiento es contractual, no de ventas. La renta variable es 5,1% de los ingresos y cayó 1,1% en 1S26 mientras las ventas de los locatarios subían 6,7%, porque las rentas mínimas (renta en mismas tiendas +5,5%) crecen más que las ventas (ventas en mismas tiendas +3,7%). Los ingresos se comportan como un bono, pero su crecimiento real converge al 0,4–0,9% que la propia compañía supone para Chile en su tasación." },
          { h1: "El margen subió por mix y recuperación de gastos; el apalancamiento operativo se agotó" },
          { p: "El margen EBITDA subió de 76,2% en 2022 a 80,5% en 2025 por dos razones. Perú consolida con un margen de segmento de 83%, frente a 80% en Chile y 73% en Colombia. Y el margen de Chile ganó 2,4 pp por una mayor recuperación de gastos comunes, que se presentan netos de costo (Nota 2.20): es real, pero depende de la ocupación y de la negociación con locatarios, así que no es estructural." },
          { p: "Con el costo de ventas en 6% de los ingresos y costos operativos totales cercanos a 20%, queda poco costo fijo que diluir; con la D&A en 0,7% de los ingresos, el EBIT (CLP 523 mil MM en 2025) sigue al EBITDA. En 1S26 el EBITDA creció 7,4%, en línea con los ingresos, y el margen se mantuvo en 79,7%. Los gastos de administración subieron 18,1% por remuneraciones (+18,9%: transición del gerente general y un año completo de Perú), un gasto en incobrables que se recupera desde una base 2025 favorecida por recuperaciones (0,64% de los ingresos, bajo el 1,22% de 2024) y un impuesto al patrimonio único en Colombia de CLP 1,7 mil MM. Sin estos dos últimos, el margen de 1S26 sería 80,4%, 1,1 pp más." },
        ],
        figuras: [
          ["Figura 2:", "Sin revalorizaciones ni partidas no monetarias, CLP 1.441 mil MM de utilidad se vuelven CLP 408 mil MM de FFO (CLP mil MM)", "FA2_utilidad_ffo", "Celeste suma, rojo resta, azul = totales. Tras minoritarios (CLP 2,8 mil MM), el FFO atribuible es CLP 405,2 mil MM vs 405,3 reportado. Fuente: EEFF Plaza S.A. (CMF) 2025, Notas 13.1 y 23; análisis del equipo."],
          ["Figura 3:", "Dos tercios del crecimiento 2025 fueron comprados; en 1S26 la moneda y la inflación explican más de la mitad", "FA3_crecimiento", "Segmentos separados en forma multiplicativa (real × inflación × moneda) desde el crecimiento en moneda local (Chile 6,1%, Perú 7,9%, Colombia 7,3%); consolidado ponderado por ingresos 2025. Fuente: EEFF intermedios (CMF), Nota 24; UF (SII/BCCh); Bloomberg; INEI; DANE; análisis del equipo."]] },
      { cortes: [[0, [0]], [4, [1, 2]], [8, [3]]], texto: [
          { h1: "La caja cubre el dividendo 2,6x; el techo es la política, no la liquidez" },
          { p: "El flujo operativo subió de CLP 301 mil MM en 2024 a CLP 441 mil MM en 2025, convirtiendo 98,7% del EBITDA antes de impuestos. El impuesto pagado fue solo 15% del resultado operacional, protegido por goodwill tributario de reorganizaciones en Chile y pérdidas tributarias (Nota 13), y subirá hacia la tasa legal de 27%. Después de capex e intereses netos, el flujo de caja libre antes del servicio de la deuda fue CLP 280 mil MM y cubrió los dividendos 2,6x (Figura 4)." },
          { p: "El salto de 38,6% del flujo operativo en 1S26 no es capacidad de generación: el 59% corresponde a cobranzas de diciembre que se revierten vía capital de trabajo. Los días de cobro bajaron de 66 a 50, sin señales de peor recaudación." },
          { p: "El capex está subiendo. Los compromisos contractuales de inversión aumentaron a CLP 78 mil MM en junio de 2026 desde CLP 18 mil MM en diciembre de 2025, y el capex en propiedades de 1S26 fue CLP 81 mil MM (+176%). El caso base financia los USD 445 MM restantes del pipeline en 2026–2028 y luego sigue agregando cerca de 27.000 m² de GLA al año, el ritmo orgánico de 2022–2025, por lo que el capex se estabiliza cerca de 16% de los ingresos y no en el piso de mantención de 5,5% (Figura 7)." },
          { h1: "Los indicadores de flujo justifican el AAA; los ratios contables lo exageran" },
          { p: "El LTV bajó de 26,9% a 18,8% en 2025, pero 5,0 pp de la caída de 8,0 pp vinieron de revalorizar el denominador; con las propiedades al valor de 2024 más las adiciones, el LTV habría sido 23,8% (Figura 5). Los indicadores de flujo son el lente honesto y respaldan el alza de Feller Rate a AAA (25 de septiembre de 2026): la DFN/EBITDA bajó de 3,78x a 2,36x con el EBITDA UDM de junio de 2026 (la cifra de 2024 incluye solo un mes de Perú), y el EBITDA cubre 8,4x los costos financieros." },
          { p: "La deuda bruta de CLP 1.633 mil MM está 79% en UF, 14% en PEN y 7% en COP, en línea con las monedas de los ingresos; el 95% es de largo plazo, con una duración de 6,6 años. Los bonos en UF pagan un cupón real de 3,36% (3,48% efectivo), pero con el reajuste UF el costo nominal total fue cercano a 7% en 2025. Los vencimientos se concentran en 2029, 2031 y 2033 (Figura 6); refinanciar la serie S (1,32% efectivo) a la tasa de la última emisión (4,33%) sumaría cerca de CLP 4,9 mil MM al año. La liquidez es holgada: a junio de 2026 la razón corriente era 2,50x y la caja cubría 1,71x los pasivos corrientes." },
          { p: "Nuestro caso base usa esa holgura. La DFN/EBITDA converge linealmente de 2,36x a 3,5x a fines de 2031E, entre Cencosud Shopping (2,1x) y Parque Arauco (5,7x), y cada vuelta adicional suma 25 pb al costo real de la deuda. La caja por sobre el payout de 40% y el plan de capex se devuelve con recompras o dividendos extraordinarios: CLP 1,8 billones acumulados en 2S26E–2031E (Anexo 5). Con 3,5x, el nuevo AAA probablemente quedaría bajo presión." },
          { h1: "El forecast proyecta lo que respalda la historia, incluido el ritmo de expansión" },
          { p: "Los ingresos crecen 7,7% en 2026E y 5,0% en 2031E (CAGR 2026E–2031E de 6,4%), a medida que el crecimiento real en mismas tiendas converge a los puntos medios de la tasación, maduran los 123.700 m² del pipeline con fecha y, desde 2029E, abren cerca de 27.000 m² al año al ritmo histórico (Figura 7); los ingresos 2026E y 2027E quedan a menos de 0,6% del consenso. El margen EBITDA se mantiene en 79,7–79,8%, bajo el máximo de 2025. El impuesto en caja se normaliza de 16,8% a 27% hacia 2030E y los costos financieros suben con el apalancamiento, por lo que el margen FFO baja de 60,9% a 52,2%. El capital de trabajo se mantiene en 15,6% de los ingresos, el nivel de diciembre de 2025 (Anexo 1). El GLA nuevo rinde 6,9% sobre el costo, un retorno real de 5,0% después de impuestos frente a una WACC real de 7,2%: la expansión suma ingresos, pero no valor." },
        ],
        figuras: [
          ["Figura 4:", "El flujo de caja libre cubre los dividendos 1,5–2,6x aun con más capex (2025, CLP mil MM)", "FA4_caja", "Estado de flujos (método directo; intereses pagados en financiamiento). FCF antes de amortizaciones. Capex = propiedades de inversión + intangibles + PP&E (15,8% de los ingresos en 2025; 24,8% en 1S26). Fuente: EEFF Plaza S.A. (CMF); análisis del equipo."],
          ["Figura 5:", "Dos tercios de la caída del LTV en 2025 son revalorización, no desapalancamiento", "FA5_ltv", "LTV = DFN / propiedades de inversión. Deuda y adiciones = LTV con propiedades al valor 2024 más adiciones 2025 (23,8%). Ratios jun-26 sobre EBITDA UDM. Fuente: EEFF (CMF), Notas 11 y 26; Análisis Razonado 2025 y 1S26; análisis del equipo."],
          ["Figura 6:", "El riesgo de refinanciamiento está en 2029, 2031 y 2033, no en el corto plazo", "FA6_vencimientos", "Bonos en UF, valor libro dic-25 (total CLP 1.277,7 mil MM); etiqueta = tasa real efectiva (años con dos series: promedio ponderado). Excluye CLP 301,6 mil MM de deuda bancaria en PEN/COP. Fuente: EEFF dic-25, Nota 14; análisis del equipo."],
          ["Figura 7:", "Tras el pipeline, la expansión mantiene el capex cerca de CLP 150 mil MM al año mientras el apalancamiento sube a 3,5x", "FA7_capex", "2026E incluye CLP 84,2 mil MM gastados en 1S26; el apalancamiento 2024A incluye un mes de EBITDA de Perú. Fuente: EEFF (CMF); release 2T26 (pipeline); caso base de valoración (Anexos 1 y 5)."]] },
    ],
    anexos: {
      a1: ["Anexo 1:", "Supuestos del forecast (CLP nominales, 2026E–2031E)", ["Variable", "Supuesto", "Base", "Tag"],
        [["Ancla 2026E", "1S26 real + 2S26E que repite el crecimiento en moneda local de 1S26 (Chile 6,1%, Perú 7,9%, Colombia 7,3%)", "EEFF intermedios jun-26", "R/S"],
          ["Inflación", "IPC de consenso 2026–28; metas de los bancos centrales desde 2029 (CL 3%, PE 2%, CO 3%)", "Arriendos indexados a UF/IPC", "P/S"],
          ["Crecimiento real mismas tiendas (sin parking)", "Chile 1,5% → 0,6%; Perú 1,5% → 0,4%; Colombia 2,5% → 2,1% (2027E → 2031E)", "Puntos medios de la tasación (Nota 11)", "S"],
          ["Ocupación", "Chile 96,1% y Colombia 97,5% constantes; Perú 93,8% → 94,6% en 2028E", "2T26 real; Perú vuelve a 2025", "R/S"],
          ["Parking", "12,1% / 6,4% / 9,0% de los ingresos 2026; crecimiento real 9% → 2%", "Release 2T26 (UDM +23,7%)", "R/S"],
          ["GLA del pipeline", "8.900 m² (2026), 72.950 m² (2027), 41.850 m² (2028); rampa 35/90/100%", "Release 2T26, proyectos", "R/S"],
          ["Expansión post-2028", "27.075 m² al año desde 2029E (55% Chile / 45% Perú); USD 3.603/m² indexado a IPC, 50% en t−2 y 50% en t−1; margen EBITDA 80,9%; yield on cost 6,9%", "GLA orgánico promedio 2022, 2023, 2025; costo por m² del pipeline", "C/S"],
          ["Tipo de cambio", "Consenso de fin de año 2026–28; paridad de poder de compra después", "Bloomberg ECFC, 5-oct-26", "P/S"],
          ["Margen EBITDA", "Chile 80%, Perú 82%, Colombia 74% (≈ 79,7% consolidado)", "2024–UDM 2T26; no se extrapola el máximo 2025", "S"],
          ["D&A", "0,70% de los ingresos", "UDM 2T26 (4,8 / 677,3)", "C"],
          ["Impuesto en caja", "16,8% (2026E) → 19% → 21,5% → 24% → 27% (2030E–31E)", "Impuesto corriente 1S26; escudos se agotan", "C/S"],
          ["Capex de proyectos", "USD 445 MM restantes a jun-26, repartidos hasta la apertura (2026E–28E)", "Release 2T26, pipeline", "R/S"],
          ["Mantención + TI", "5,5% de los ingresos", "2024 (4,8%) – UDM 2T26 (~6,6%)", "S"],
          ["Capital de trabajo", "Capital de trabajo neto operativo en 15,6% de los ingresos (dic-25); 2026E cubre solo 2S26E", "EEFF dic-25, primer cierre con Perú completo", "C/S"],
          ["Apalancamiento", "DFN/EBITDA 2,36x (jun-26) → 3,5x a dic-31E, lineal; la caja sobre el plan se devuelve con recompras o dividendos extraordinarios", "Supuesto del equipo; pares 2,1x–5,7x", "S"],
          ["Costo de la deuda", "Costo financiero neto 3,08% de la DFN inicial + 25 pb reales por vuelta sobre 2,36x; reajuste UF sobre 79% de la deuda", "1S26 anualizado; supuesto del equipo", "C/S"],
          ["Excluido", "Greenfield (USD 0,7 MM UDM vs 55,2 MM en 2023), Gran Plaza; proyectos sin fecha en Chile reemplazados por la expansión post-2028", "Sin fecha o inmaterial", "S"],
          ["Payout", "40% de la utilidad distribuible con un año de rezago, más las distribuciones adicionales", "Política del directorio (Nota 20e)", "R/S"]],
        "Tags: R reportado · C cálculo del equipo · S supuesto del equipo · P consenso de mercado. Fuente: EEFF Plaza S.A. (CMF); release 2T26; Bloomberg ECFC; análisis del equipo."],
      a2: ["Anexo 2:", "Estados financieros resumidos (CLP mil MM)",
        ["Estado de resultados", "2022A", "2023A", "2024A", "2025A", "UDM 2T26"],
        ["Ingresos", "EBITDA (definición compañía)", "D&A", "Revalorización y otras partidas op., neto", "Resultado operacional", "Costo financiero neto", "Reajuste UF de la deuda", "Utilidad atribuible", "FFO atribuible (compañía)"],
        ["Balance", "dic-23", "dic-24", "dic-25", "jun-26"],
        ["Caja y equivalentes", "Propiedades de inversión", "Activos totales", "Deuda financiera bruta", "Pasivos por impuesto diferido", "Patrimonio atribuible"],
        ["Flujo de caja", "2023A", "2024A", "2025A", "1S26"],
        ["Flujo operativo", "Capex", "Adquisiciones (incl. minoritarios)", "Emisión de acciones", "Endeudamiento neto / (pago)", "Intereses pagados", "Dividendos pagados"],
        "UDM = 2025 − 1S25 + 1S26. 2024 reexpresado en los EEFF dic-25; 2022A de los EEFF dic-23 reexpresados; las líneas de financiamiento 2022 no están en los archivos usados. Fuente: EEFF Plaza S.A. (CMF) dic-24, dic-25, jun-26."],
      a3: ["Anexo 3:", "Tablero de ratios", ["", "2023", "2024", "2025", "jun-26"],
        ["ROE, normalizado", "ROA, normalizado", "ROIC (NOPAT al 27%, capital a valor razonable)", "Razón corriente (x)", "Caja / pasivos corrientes (x)", "Pasivos totales / patrimonio (x)", "LTV (DFN / propiedades de inversión)", "DFN / EBITDA (x)", "EBITDA / costos financieros (x)", "Rotación de activos (x)", "Costos operativos / ingresos", "Días de cobro"],
        "Apalancamiento jun-26 sobre EBITDA UDM; costos operativos = costo de ventas + administración (1S26). Días de cobro dic-24 n.s. (Perú consolida un mes). Los retornos a jun-26 requieren el balance de jun-25. Fuente: EEFF (CMF); Análisis Razonado 1S26; análisis del equipo."],
      a4: ["Anexo 4:", "El EBITDA que usamos es el de la compañía: la revalorización se elimina una sola vez (CLP mil MM)", ["", "2022A", "2023A", "2024A", "2025A", "UDM 2T26"],
        ["Resultado operacional + D&A", "(−) Otros ingresos op. (principalmente valor razonable)", "(+) Otros gastos op. (no recurrentes)", "EBITDA ajustado", "Control vs EBITDA de la compañía"],
        "El campo EBITDA de Bloomberg incluye la revalorización y no se usa. Fuente: EEFF (CMF), Notas 23.1–23.3; release 2T26; análisis del equipo."],
      a5: ["Anexo 5:", "Proyección del caso base (CLP mil MM)", ["CLP mil MM", "2026E", "2027E", "2028E", "2029E", "2030E", "2031E"],
        ["Ingresos — Chile", "Ingresos — Perú", "Ingresos — Colombia", "Eliminaciones", "Ingresos — expansión post-2028", "Ingresos", "Crecimiento", "EBITDA", "Margen EBITDA", "D&A", "EBIT", "Tasa de impuesto en caja", "Capex — pipeline + mantención", "Capex — expansión post-2028", "Δ capital de trabajo", "FFO", "Utilidad distribuible", "Dividendos ordinarios (40%)", "Distribuciones adicionales", "DFN / EBITDA (x)"],
        "2026E = 1S26 real + 2S26E; el capex incluye CLP 84,2 mil MM gastados en 1S26; Δ capital de trabajo y distribuciones adicionales cubren solo 2S26E. FFO = (EBITDA − costo financiero neto) × (1 − impuesto en caja), antes de minoritarios (~1%); utilidad distribuible = FFO − reajuste UF. Fuente: modelo de valoración v2 (hojas DCF, Deuda y Expansion); análisis del equipo."],
      a6: ["Anexo 6:", "En base comparable, Mallplaza tiene el ROE normalizado más alto de sus pares locales (2025)", ["", "Mallplaza", "Parque Arauco", "Cencosud Shopping"],
        ["Revalorización / propiedades iniciales", "EBITDA / propiedades de inversión", "ROE, reportado", "ROE, normalizado", "LTV", "DFN / EBITDA (x)", "Activos / patrimonio atribuible, cierre (x)"],
        "Mallplaza normalizado con la utilidad distribuible auditada; pares con utilidad neta − revalorización × (1 − 27%). Las propiedades de Parque Arauco excluyen JV, lo que sobrestima su LTV. Fuente: EEFF (CMF); Bloomberg FA (PARAUCO CI, CENCOMAL CI; 28-sep-26); análisis del equipo."],
    },
  },
  en: {
    archivo: "Mallplaza_Analisis_Financiero_EN.docx", pag: "Page ", pagA: "A-",
    titulo: "FINANCIAL ANALYSIS", tituloA: "APPENDIX: FINANCIAL ANALYSIS",
    lead: "Mallplaza's rental engine grows at roughly 3% in real terms with an 80% EBITDA margin and cash-flow leverage below 2.5x; headline profit, growth and leverage all overstate that engine and must be normalized before forecasting.",
    fa1: ["Figure 1:", "Normalized ROE of 8–10% is the real return; reported ROE is appraisal-driven",
      ["CLP bn", "Revenue", "EBITDA margin %", "FFO margin %", "Net income, reported", "Distributable profit", "Net margin, normalized %", "Asset turnover (x)", "Assets / equity (x)", "ROE, reported %", "ROE, normalized %", "ROA, normalized %"],
      "Normalized = distributable profit (2022A team proxy; 2022 UF indexation inflated by ~13% inflation). Average balances; series starts in 2022 (IAS 40 fair value applied retroactively from 1 Jan 2022). 2026E–2031E: valuation base case; balance sheet rolled forward from Dec-25 (equity + reported profit − ordinary and additional distributions; property revalued at 3% inflation, 27% deferred tax; net debt rising to 3.5x EBITDA by 2031E). Source: Plaza S.A. financial statements (CMF), Note 20e; 2Q26 results release (FFO); Team analysis."],
    paginas: [
      { cortes: [[0, [0]], [3, [1]]], texto: [
          { h1: "Reported earnings track the appraiser; FFO tracks the business" },
          { p: "Under IAS 40, revaluations of investment property run through the income statement. In 2025 a CLP 1,447bn fair-value gain, 27% of opening property value against about 2% a year at Parque Arauco and Cencosud Shopping, lifted attributable net income to CLP 1,431bn (2.2x revenue) and reported ROE to 37.5% (Figure 1). About two-thirds of Chile's 29% uplift came from a lower appraisal discount rate, not from higher NOI." },
          { p: "Excluding fair-value gains net of deferred tax, audited distributable profit was CLP 369bn and ROE 9.7%, the highest of the three local peers on a comparable basis (Appendix 6). FFO, which we rebuild from the statements to within 0.02% of the reported CLP 405bn (Figure 2), held a 59–62% margin over 2022–2025 while net income swung more than fourfold. FFO excludes the UF indexation of debt (CLP 44bn, 11% of FFO); distributable profit includes it and is our conservative earnings anchor." },
          { h1: "Two-thirds of 2025 growth was bought; organic real growth is about 3%" },
          { p: "Revenue rose 76% between 2022 and 2025, but 21.8pp of 2025's 32.2% came from perimeter: the Peruvian portfolio bought from Falabella, consolidated from 4 December 2024. Ex-Peru, revenue grew 10.7% in 2025 and 7.1% in 1H26." },
          { p: "Once currency and inflation are removed, 1H26 reported growth of 7.4% falls to about 3.0% real (Figure 3). Chile grew 2.8% after 3.1% UF indexation (5.1% in 2025); Peru 4.5% despite occupancy slipping to 93.8%; Colombia only about 1% once a 7.5% COP appreciation and 6.1% inflation are stripped out." },
          { p: "Growth is contractual, not sales-driven. Variable rent is 5.1% of revenue and fell 1.1% in 1H26 while tenant sales rose 6.7%, because minimum rents (same-store rent +5.5%) are outpacing tenant sales (same-store sales +3.7%). Revenue is therefore bond-like, but its real growth is converging on the 0.4–0.9% the company assumes for Chile in its own appraisal." },
          { h1: "Margin gains came from mix and cost recovery; operating leverage is spent" },
          { p: "EBITDA margin rose from 76.2% in 2022 to 80.5% in 2025 for two reasons. Peru consolidates at an 83% segment margin, against 80% in Chile and 73% in Colombia. Chile's margin gained 2.4pp on a higher recovery of common charges, which are presented net of cost (Note 2.20): real, but dependent on occupancy and tenant negotiations, so not structural." },
          { p: "With cost of sales at 6% of revenue and total operating costs near 20%, little fixed cost is left to dilute; with D&A at 0.7% of revenue, EBIT (CLP 523bn in 2025) tracks EBITDA. In 1H26 EBITDA grew 7.4%, in line with revenue, and margin held at 79.7%. Administrative expense rose 18.1% on payroll (+18.9%: CEO transition and a full year of Peru), a bad-debt charge rebounding from a 2025 base flattered by recoveries (0.64% of revenue, below 1.22% in 2024), and a one-off Colombian wealth tax of CLP 1.7bn. Excluding the last two, 1H26 margin would be 80.4%, up 1.1pp." },
        ],
        figuras: [
          ["Figure 2:", "Removing fair-value gains and non-cash items turns CLP 1,441bn of profit into CLP 408bn of FFO (CLP bn)", "FA2_utilidad_ffo", "Light blue adds back, red removes, navy = totals. After minorities (CLP 2.8bn), FFO attributable is CLP 405.2bn vs CLP 405.3bn reported. Source: Plaza S.A. financial statements (CMF) FY2025, Notes 13.1 and 23; Team analysis."],
          ["Figure 3:", "Two-thirds of 2025 growth was bought; in 1H26 currency and inflation explain more than half", "FA3_crecimiento", "Segments split multiplicatively (real × inflation × FX) from local-currency growth (Chile 6.1%, Peru 7.9%, Colombia 7.3%); consolidated weighted by FY2025 revenue. Source: Plaza S.A. interim financial statements (CMF), Note 24; UF (SII/BCCh); Bloomberg FX; INEI; DANE; Team analysis."]] },
      { cortes: [[0, [0]], [4, [1, 2]], [8, [3]]], texto: [
          { h1: "Cash covers the dividend 2.6x; the ceiling is policy, not liquidity" },
          { p: "Operating cash flow rose from CLP 301bn in 2024 to CLP 441bn in 2025, converting 98.7% of EBITDA before tax. Cash tax was only 15% of operating profit, sheltered by tax goodwill from Chilean reorganizations and loss carry-forwards (Note 13), and will rise toward the 27% statutory rate. After capex and net interest, free cash flow before debt service was CLP 280bn, covering dividends 2.6x (Figure 4)." },
          { p: "The 38.6% jump in 1H26 operating cash flow is not earnings power: 59% of it was December collections unwinding through working capital. Receivable days fell from 66 to 50, with no sign of weaker collection." },
          { p: "Capex is stepping up. Contractual investment commitments rose to CLP 78bn at June 2026 from CLP 18bn in December 2025, and 1H26 property capex was CLP 81bn (+176%). The base case funds USD 445m of remaining pipeline over 2026–2028 and then keeps adding about 27,000 m² of GLA a year, the 2022–2025 organic pace, so capex settles near 16% of revenue rather than the 5.5% maintenance floor (Figure 7)." },
          { h1: "Cash-flow metrics earn the AAA; equity-based ratios flatter it" },
          { p: "LTV fell from 26.9% to 18.8% in 2025, but 5.0pp of the 8.0pp drop came from revaluing the denominator; with property at 2024 value plus additions, LTV would have been 23.8% (Figure 5). Cash-flow metrics are the honest lens and support Feller Rate's upgrade to AAA (25 September 2026): net debt/EBITDA fell from 3.78x to 2.36x on June 2026 LTM EBITDA (the 2024 figure includes only one month of Peru), and EBITDA covers financial costs 8.4x." },
          { p: "Gross debt of CLP 1,633bn is 79% UF, 14% PEN and 7% COP, matching revenue currencies; 95% is long term, with a 6.6-year duration. The UF bonds carry a 3.36% real coupon (3.48% effective), but with UF indexation the all-in nominal cost was about 7% in 2025. Maturities cluster in 2029, 2031 and 2033 (Figure 6); refinancing series S (1.32% effective) at the latest issue's 4.33% would add about CLP 4.9bn a year. Liquidity is ample: at June 2026 the current ratio was 2.50x and cash covered current liabilities 1.71x." },
          { p: "Our base case spends that headroom. Net debt/EBITDA converges linearly from 2.36x to 3.5x by end-2031E, between Cencosud Shopping (2.1x) and Parque Arauco (5.7x), and each extra turn adds 25bp to the real cost of debt. Cash beyond the 40% payout and the capex plan is returned through buybacks or special dividends, CLP 1.8tn cumulatively over 2H26E–2031E (Appendix 5). At 3.5x the new AAA would likely come under pressure." },
          { h1: "The forecast carries forward what history supports, including the pace of expansion" },
          { p: "Revenue grows 7.7% in 2026E and 5.0% in 2031E (2026E–2031E CAGR 6.4%), as real same-store growth converges to the appraisal midpoints, 123,700 m² of dated pipeline matures and, from 2029E, about 27,000 m² a year opens at the historical pace (Figure 7); 2026E and 2027E revenue sit within 0.6% of consensus. EBITDA margin stays at 79.7–79.8%, below the 2025 peak. Cash tax normalizes from 16.8% to 27% by 2030E and financial costs rise with leverage, so FFO margin slides from 60.9% to 52.2%. Working capital is held at 15.6% of revenue, the December 2025 level (Appendix 1). New GLA yields 6.9% on cost, a 5.0% real return after tax against a 7.2% real WACC: expansion adds revenue but not value." },
        ],
        figuras: [
          ["Figure 4:", "Free cash flow covers dividends 1.5–2.6x even as capex rises (FY2025, CLP bn)", "FA4_caja", "Cash flow statement (direct method; interest paid in financing). FCF before debt repayment. Capex = investment property + intangibles + PP&E (15.8% of revenue in 2025; 24.8% in 1H26). Source: Plaza S.A. financial statements (CMF); Team analysis."],
          ["Figure 5:", "Two-thirds of the 2025 fall in LTV is revaluation, not deleveraging", "FA5_ltv", "LTV = net financial debt / investment property. Debt & additions = LTV with property at 2024 value plus 2025 additions (23.8%). Jun-26 ratios on LTM EBITDA. Source: Plaza S.A. financial statements (CMF), Notes 11 and 26; MD&A 2025 and 1H26; Team analysis."],
          ["Figure 6:", "Refinancing risk sits in 2029, 2031 and 2033, not in the near term", "FA6_vencimientos", "UF bonds, book value Dec-25 (total CLP 1,277.7bn); label = effective real rate (two-series years: book-value-weighted). Excludes CLP 301.6bn of PEN/COP bank debt. Source: Plaza S.A. financial statements Dec-25, Note 14; Team analysis."],
          ["Figure 7:", "After the pipeline, expansion keeps capex near CLP 150bn a year while leverage rebuilds to 3.5x", "FA7_capex", "2026E includes CLP 84.2bn spent in 1H26; 2024A leverage includes one month of Peru EBITDA. Source: Plaza S.A. financial statements (CMF); 2Q26 results release (pipeline); valuation base case (Appendices 1 and 5)."]] },
    ],
    anexos: {
      a1: ["Appendix 1:", "Forecast assumptions (nominal CLP, 2026E–2031E)", ["Driver", "Assumption", "Basis", "Tag"],
        [["2026E anchor", "1H26 actual + 2H26E repeating 1H26 local-currency growth (Chile 6.1%, Peru 7.9%, Colombia 7.3%)", "Interim FS Jun-26", "R/S"],
          ["Inflation", "Consensus CPI 2026–28; central-bank targets from 2029 (CL 3%, PE 2%, CO 3%)", "Leases indexed to UF/CPI", "P/S"],
          ["Real same-store growth (ex-parking)", "Chile 1.5% → 0.6%; Peru 1.5% → 0.4%; Colombia 2.5% → 2.1% (2027E → 2031E)", "Midpoints of company appraisal ranges (Note 11)", "S"],
          ["Occupancy", "Chile 96.1%, Colombia 97.5% flat; Peru 93.8% → 94.6% by 2028E", "2Q26 actual; Peru back to 2025", "R/S"],
          ["Parking", "12.1% / 6.4% / 9.0% of 2026 revenue; real growth 9% → 2%", "2Q26 release (LTM +23.7%)", "R/S"],
          ["Pipeline GLA", "8,900 m² (2026), 72,950 m² (2027), 41,850 m² (2028); ramp 35/90/100%", "2Q26 release, projects", "R/S"],
          ["Post-2028 expansion", "27,075 m² a year opening from 2029E (55% Chile / 45% Peru); USD 3,603/m² indexed to CPI, spent 50% in t−2 and 50% in t−1; 80.9% EBITDA margin; 6.9% yield on cost", "Average organic GLA 2022, 2023, 2025; pipeline cost per m²", "C/S"],
          ["FX", "Consensus year-end 2026–28; purchasing-power parity thereafter", "Bloomberg ECFC, 5 Oct 2026", "P/S"],
          ["EBITDA margin", "Chile 80%, Peru 82%, Colombia 74% (≈ 79.7% consolidated)", "2024–LTM 2Q26; 2025 peak not extrapolated", "S"],
          ["D&A", "0.70% of revenue", "LTM 2Q26 (4.8 / 677.3)", "C"],
          ["Cash tax rate", "16.8% (2026E) → 19% → 21.5% → 24% → 27% (2030E–31E)", "1H26 current tax; shields run off", "C/S"],
          ["Project capex", "USD 445m remaining at Jun-26, spread to final opening (2026E–28E)", "2Q26 release, pipeline", "R/S"],
          ["Maintenance + IT capex", "5.5% of revenue", "2024 (4.8%) – LTM 2Q26 (~6.6%)", "S"],
          ["Working capital", "Net operating working capital at 15.6% of revenue (Dec-25); 2026E covers 2H26E only", "FS Dec-25, first year-end with full Peru", "C/S"],
          ["Leverage", "Net debt / EBITDA 2.36x (Jun-26) → 3.5x by Dec-31E, linear; cash above the plan returned via buybacks or special dividends", "Team assumption; peers 2.1x–5.7x", "S"],
          ["Cost of debt", "Net financial cost 3.08% of opening net debt + 25bp real per turn of leverage above 2.36x; UF indexation on 79% of debt", "1H26 annualized; team assumption", "C/S"],
          ["Excluded", "Greenfield (USD 0.7m LTM vs 55.2m in 2023), Gran Plaza; undated Chile projects replaced by the post-2028 expansion", "No date or immaterial", "S"],
          ["Payout", "40% of distributable profit, one-year lag, plus the additional distributions above", "Board policy (Note 20e)", "R/S"]],
        "Tags: R reported · C team calculation · S team assumption · P market consensus. Source: Plaza S.A. financial statements (CMF); 2Q26 results release; Bloomberg ECFC; Team analysis."],
      a2: ["Appendix 2:", "Summary financial statements (CLP bn)",
        ["Income statement", "2022A", "2023A", "2024A", "2025A", "LTM 2Q26"],
        ["Revenue", "EBITDA (company definition)", "D&A", "Fair-value gains & other op. items, net", "Operating income", "Net financial cost", "UF indexation of debt", "Net income attributable", "FFO attributable (company)"],
        ["Balance sheet", "Dec-23", "Dec-24", "Dec-25", "Jun-26"],
        ["Cash and equivalents", "Investment property", "Total assets", "Gross financial debt", "Deferred tax liabilities", "Equity attributable to owners"],
        ["Cash flow", "2023A", "2024A", "2025A", "1H26"],
        ["Operating cash flow", "Capex", "Acquisitions incl. minorities", "Equity issued", "Net borrowing / (repayment)", "Interest paid", "Dividends paid"],
        "LTM = FY2025 − 1H25 + 1H26. 2024 as restated in Dec-25; 2022A from restated Dec-23 statements; 2022 financing lines not in the files used. Source: Plaza S.A. financial statements (CMF) Dec-24, Dec-25, Jun-26."],
      a3: ["Appendix 3:", "Ratio dashboard", ["", "2023", "2024", "2025", "Jun-26"],
        ["ROE, normalized", "ROA, normalized", "ROIC (NOPAT at 27%, IC at fair value)", "Current ratio (x)", "Cash / current liabilities (x)", "Total liabilities / equity (x)", "LTV (net debt / inv. property)", "Net debt / EBITDA (x)", "EBITDA / financial costs (x)", "Asset turnover (x)", "Operating costs / revenue", "Receivable days"],
        "Jun-26 leverage on LTM EBITDA; operating costs = cost of sales + admin (1H26). Dec-24 receivable days n.m. (Peru consolidated one month). Jun-26 returns need the Jun-25 balance sheet. Source: Plaza S.A. financial statements (CMF); MD&A 1H26; Team analysis."],
      a4: ["Appendix 4:", "The EBITDA we use is the company's: fair-value gains are removed once (CLP bn)", ["", "2022A", "2023A", "2024A", "2025A", "LTM 2Q26"],
        ["Operating income + D&A", "(−) Other op. income (mostly fair value)", "(+) Other op. expenses (non-recurring)", "EBITDA, adjusted", "Check vs company EBITDA"],
        "Bloomberg's EBITDA field includes fair-value gains and is not used. Source: Plaza S.A. financial statements (CMF), Notes 23.1–23.3; 2Q26 release; Team analysis."],
      a5: ["Appendix 5:", "Base-case projection (CLP bn)", ["CLP bn", "2026E", "2027E", "2028E", "2029E", "2030E", "2031E"],
        ["Revenue — Chile", "Revenue — Peru", "Revenue — Colombia", "Eliminations", "Revenue — post-2028 expansion", "Revenue", "Growth", "EBITDA", "EBITDA margin", "D&A", "EBIT", "Cash tax rate", "Capex — pipeline + maintenance", "Capex — post-2028 expansion", "Δ working capital", "FFO", "Distributable profit", "Ordinary dividends (40%)", "Additional distributions", "Net debt / EBITDA (x)"],
        "2026E = 1H26 actual + 2H26E; capex includes CLP 84.2bn spent in 1H26; Δ working capital and additional distributions cover 2H26E only. FFO = (EBITDA − net financial cost) × (1 − cash tax), before minorities (~1%); distributable profit = FFO − UF indexation. Source: valuation workbook v2 (DCF, Deuda and Expansion sheets); Team analysis."],
      a6: ["Appendix 6:", "Like-for-like, Mallplaza earns the highest normalized ROE of its local peers (FY2025)", ["", "Mallplaza", "Parque Arauco", "Cencosud Shopping"],
        ["Fair-value gain / opening property", "EBITDA / investment property", "ROE, reported", "ROE, normalized", "LTV", "Net debt / EBITDA (x)", "Assets / equity attributable, YE (x)"],
        "Mallplaza normalized on audited distributable profit; peers on net income − fair-value gains × (1 − 27%). Parque Arauco property excludes JVs, overstating its LTV. Source: Plaza S.A. financial statements (CMF); Bloomberg FA as-reported (PARAUCO CI, CENCOMAL CI; export 28 Sep 2026); Team analysis."],
    },
  },
};

// ---------------------------------------------------------------------------------------------
const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || 16, bold: o.bold, italics: o.italics, color: o.color || TINTA });
const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: nada, bottom: nada, left: nada, right: nada };
const linea = (l = 228) => ({ line: l, lineRule: LineRuleType.AUTO });
const vacia = () => new Paragraph({ children: [] });
const titulo = (t) => new Paragraph({ spacing: { after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 2 } },
  children: [run(t, { bold: true, color: AZUL, size: 30 })] });
const subtitulo = (t, antes = 90) => new Paragraph({ spacing: { before: antes, after: 50, ...linea() }, keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ROJO, space: 1 } }, children: [run(t, { bold: true, color: AZUL })] });
const caption = (n, t, antes = 90) => new Paragraph({ spacing: { before: antes, after: 30, ...linea() },
  children: [run(n + " ", { bold: true, color: AZUL, size: 14 }), run(t, { color: AZUL, size: 14 })] });
const fuente = (t) => new Paragraph({ spacing: { before: 10, after: 0, ...linea(216) }, alignment: AlignmentType.JUSTIFIED, children: [run(t, { size: 11, color: GRIS })] });
const imagen = (archivo, anchoIn, altoIn) => new Paragraph({ spacing: { before: 0, after: 20, line: 240, lineRule: LineRuleType.AUTO }, children: [
  new ImageRun({ type: "png", data: fs.readFileSync(path.join(FIG, archivo)), transformation: { width: Math.round(anchoIn * PX), height: Math.round(altoIn * PX) } })] });
const celda = (w, children, extra = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: sinBordes,
  margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP, children, ...extra });
const parrafo = (x) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 60, ...linea() }, children: [run(x)] });

// Tabla de datos: enc = encabezados; filas = arrays; opciones de formato
function tabla(enc, anchos, filas, { izq = 1, destacar = [], sub = [], tam = 12 } = {}) {
  const borde = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  const c = (v, i, r, cab) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA }, borders: cab ? sinBordes : { ...sinBordes, bottom: borde },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: cab ? AZUL : destacar.includes(r) ? "FDF1CC" : sub.includes(r) ? "E4E8F0" : r % 2 ? "FFFFFF" : "F7F8FB" },
    margins: { top: 18, bottom: 18, left: 60, right: 60 }, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing: linea(204), alignment: i < izq ? AlignmentType.LEFT : AlignmentType.RIGHT,
      children: [run(String(v), { size: tam, bold: cab || destacar.includes(r) || sub.includes(r), color: cab ? "FFFFFF" : TINTA })] })] });
  return new Table({ width: { size: anchos.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: anchos,
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
    rows: [new TableRow({ tableHeader: true, children: enc.map((v, i) => c(v, i, -1, true)) }),
      ...filas.map((f, r) => new TableRow({ cantSplit: true, children: f.map((v, i) => c(v, i, r, false)) }))] });
}

// Una fila por bloque de texto: cada figura va en la fila del texto que la cita (p.cortes = [[desde, [figuras]], ...])
function pagina(lang, p) {
  const cortes = p.cortes || [[0, p.figuras.map((_, k) => k)]];
  const rows = cortes.map(([desde, figs], c) => {
    const hasta = c + 1 < cortes.length ? cortes[c + 1][0] : p.texto.length;
    const out = [];
    p.texto.slice(desde, hasta).forEach((b, k) => {
      const i = desde + k;
    out.push(b.h1 ? subtitulo(b.h1, i === 0 ? 20 : 90) : parrafo(b.p));
    });
    const izq = out;
    const der = [];
    figs.forEach((n_, k) => {
      const [n, t, archivo, f] = p.figuras[n_];
      const out = der;
      out.push(caption(n, t, k === 0 ? 20 : 90));
      out.push(imagen(`${lang}_${archivo}.png`, COL_FIG / 1440, ALTO[archivo.split("_")[0]]));
      out.push(fuente(f));
    });
    if (!der.length) der.push(vacia());
    return new TableRow({ cantSplit: true, children: [celda(COL_TEXTO, izq), celda(COL_SEP, [vacia()]), celda(COL_FIG, der)] });
  });
  return new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [COL_TEXTO, COL_SEP, COL_FIG],
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows });
}

function figura1(lang) {
  const X = L[lang], f = loc(lang), [n, t, filas, nota] = X.fa1;
  const anchos = [2266, ...Array(10).fill(840)];
  return [caption(n, t, 40),
    tabla([filas[0], ...FA1_AÑOS], anchos, FA1.map((r, i) => [filas[i + 1], ...r.map(f)]), { destacar: [9], tam: 12 }),
    fuente(nota), new Paragraph({ spacing: { after: 60 }, children: [] })];
}

function anexos(lang) {
  const A = L[lang].anexos, f = loc(lang);
  const bloque = (n, t) => caption(n, t, 100);
  const [n1, t1, c1, f1, s1] = A.a1;
  const a1 = [bloque(n1, t1), tabla(c1, [1150, 2250, 1375, 450], f1, { izq: 4, tam: 11 }), fuente(s1)];
  const [n2, t2, eIS, rIS, eBS, rBS, eCF, rCF, s2] = A.a2;
  const w5 = [2125, 620, 620, 620, 620, 620], w4 = [2125, 775, 775, 775, 775];
  const a2 = [bloque(n2, t2),
    tabla(eIS, w5, A2_IS.map((r, i) => [rIS[i], ...r.map(f)]), { destacar: [7], sub: [1] }),
    tabla(eBS, w4, A2_BS.map((r, i) => [rBS[i], ...r.map(f)]), { destacar: [5] }),
    tabla(eCF, w4, A2_CF.map((r, i) => [rCF[i], ...r.map(f)]), {}), fuente(s2)];
  const [n3, t3, e3, r3, s3] = A.a3;
  const a3 = [bloque(n3, t3), tabla(e3, w4, A3.map((r, i) => [r3[i], ...r.map(f)]), { destacar: [0] }), fuente(s3)];
  const [n4, t4, e4, r4, s4] = A.a4;
  const a4 = [bloque(n4, t4), tabla(e4, w5, A4.map((r, i) => [r4[i], ...r.map(f)]), { destacar: [3] }), fuente(s4)];
  const [n5, t5, e5, r5, s5] = A.a5;
  const a5 = [bloque(n5, t5), tabla(e5, [1795, 571, 571, 571, 571, 571, 575], A5.map((r, i) => [r5[i], ...r.map(f)]), { destacar: [5, 15], sub: [7] }), fuente(s5)];
  const [n6, t6, e6, r6, s6] = A.a6;
  const a6 = [bloque(n6, t6), tabla(e6, [2125, 1033, 1033, 1034], A6.map((r, i) => [r6[i], ...r.map(f)]), { destacar: [3] }), fuente(s6)];
  const dos = (izq, der) => new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [MEDIA, COL_SEP, MEDIA],
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
    rows: [new TableRow({ children: [celda(MEDIA, izq), celda(COL_SEP, [vacia()]), celda(MEDIA, der)] })] });
  return [titulo(L[lang].tituloA), dos(a1, a2), new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: [] }), dos([...a3, ...a4], [...a5, ...a6])];
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
    styles: { default: { document: { run: { font: FUENTE, size: 16 }, paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } } } } },
    sections: [
      { properties: { page: pag }, headers: { default: encabezado() }, footers: { default: pie(X.pag) },
        children: [titulo(X.titulo),
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 60, ...linea() }, shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO }, children: [run(X.lead, { bold: true, color: AZUL })] }),
          ...figura1(lang), pagina(lang, X.paginas[0]),
          new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: [] }), pagina(lang, X.paginas[1])] },
      { properties: { page: { ...pag, pageNumbers: { start: 1 } } }, headers: { default: encabezado() }, footers: { default: pie(X.pagA) }, children: anexos(lang) },
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
