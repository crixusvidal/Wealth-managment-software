// Anexo macroeconómico (2 páginas A4), en español e inglés, con el formato visual del informe.
// Figuras: graficos/informe/macro/{es,en}_A*.png (python graficos/anexo_macro.py).
// Uso: node informe/anexo_macro.js
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, ImageRun, LineRuleType, Packer, PageBreak,
  PageNumber, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

const AZUL = "14284B", ROJO = "E9004B", TINTA = "1A1A1A", GRIS = "5F5F5F", FONDO = "EEF1F6";
const FUENTE = "Arial";
const ANCHO = 10666, MEDIA = 5225, SEP = 216;
const FIG = path.join(__dirname, "..", "graficos", "figuras");
const PX = 96;
const ALTO = { A0: 1.85, A1: 2.55, A2: 2.55, A3: 2.3, A4: 2.3, A5: 2.2 }; // = graficos/anexo_macro.py

const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || 15, bold: o.bold,
  italics: o.italics, color: o.color || TINTA });
const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: nada, bottom: nada, left: nada, right: nada };
const linea = (l = 228) => ({ line: l, lineRule: LineRuleType.AUTO });

// Texto con **negrita** inline
function rico(texto, o = {}) {
  return texto.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((p) => p.startsWith("**")
    ? run(p.slice(2, -2), { ...o, bold: true, color: o.boldColor || AZUL }) : run(p, o));
}

const L = {
  es: {
    archivo: "Mallplaza_Anexo_Macro_ES.docx", pie: "A-",
    titulo: "ANEXO [X]: ENTORNO MACROECONÓMICO Y GEOPOLÍTICO",
    lead: "Entre 2022 y 2026 la economía mundial absorbió dos shocks de oferta de origen geopolítico y un alza estructural de las tasas largas. El resultado es un entorno de presiones estanflacionarias y mayor costo de capital que llega a Chile por tres canales —inflación, tasas y consumo— y que está detrás de los dos riesgos más severos del informe (FR1 y MR1). En los gráficos, la línea punteada azul marca la fecha de valoración (30-06-2026); los datos posteriores se muestran como contexto. Las figuras cubren la transmisión del shock (Figura A1), el petróleo (Figura A2), las tasas soberanas (Figura A3), los spreads (Figura A4), el crecimiento e inflación mundial (Figura A5) y Chile (Figura A6).",
    kpis: [
      ["USD 114", "Brent, promedio sep-26", "+61% vs feb-26, antes del conflicto"],
      ["5,22%", "UST 10Y, 8-oct-26", "máximo en 24 años (desde 2002)"],
      ["117 pb", "Spread Francia–Alemania", "30-sep-26; ~150 pb a inicios de oct."],
      ["3,0% | 4,7%", "FMI 2026: PIB | inflación", "crecimiento recortado; desinflación detenida"],
      ["0,25–0,75%", "PIB Chile 2026 (BCCh)", "desde 2–3% proyectado en dic-25"],
    ],
    s1: "1. CÓMO LLEGA EL SHOCK A MALLPLAZA",
    a0: ["Figura A1:", "Mapa de transmisión: de los shocks globales a los riesgos de Mallplaza",
      "Fuente: análisis del equipo; cifras de las Figuras 24–26 del informe (Investment Risks)."],
    s2: "2. DOS SHOCKS DE OFERTA Y UN ALZA GLOBAL DE TASAS",
    a1: ["Figura A2:", "Precio del petróleo Brent, 2022–2026 (USD/barril)",
      "**Dos shocks de oferta.** La invasión rusa de Ucrania (feb-22) llevó el Brent a USD 123 en junio de 2022. En 2026, el conflicto en el estrecho de Ormuz —por donde pasa más de un cuarto del petróleo transportado por mar— lo subió de USD 71 en febrero a USD 117 en abril; tras una tregua en junio, los ataques lo devolvieron a USD 114 en septiembre.",
      "Fuente: EIA (Brent spot, promedio mensual; STEO abr-26, sep-26 y oct-26); Statista."],
    a2: ["Figura A3:", "Rendimiento soberano a 10 años, 2022–2026 (%)",
      "**Las tasas largas suben juntas.** El UST 10Y pasó de 1,8% (ene-22) a 5,0% (sep-26) y tocó 5,22% el 8 de octubre, su máximo en 24 años; el Bund pasó de −0,1% a 3,6%. Los déficits fiscales presionan la tasa libre de riesgo y Chile acompaña el movimiento desde mediados de 2026.",
      "Fuente: FRED (GS10), Bundesbank, Banque de France (vía OCDE) y BCCh (BCP 10 años); promedios mensuales."],
    s3: "3. PRIMA FISCAL Y ESTANFLACIÓN",
    a3: ["Figura A4:", "Spreads a 10 años: Francia–Alemania y Chile–EE.UU. (pb)",
      "**El alza es fiscal y global, no de riesgo Chile.** El spread de Francia sobre Alemania subió de ~40 pb en 2022 a 81 pb en promedio en septiembre de 2026 y a 117 pb al cierre del mes (~150 pb a inicios de octubre, máximo desde 2011–12). El de Chile sobre EE.UU. se comprimió: la tasa local depende cada vez más de la tasa base global.",
      "Fuente: elaboración propia con los datos de la Figura A3 (promedios mensuales); cierre de septiembre: Banque de France / Bundesbank."],
    a4: ["Figura A5:", "Crecimiento e inflación mundial, 2022–2026P (%)",
      "**Estanflación.** El FMI recortó el crecimiento mundial de 2026 a 3,0% y subió la inflación a 4,7% (desde 4,1% en 2025): la desinflación que venía desde 2024 se detuvo por el alza de la energía y los alimentos, en una economía que crece menos.",
      "Fuente: FMI, World Economic Outlook (actualización jul-2026). 2026P = proyección."],
    s4: "4. CHILE: MENOS CRECIMIENTO E INFLACIÓN PERSISTENTE",
    a5: ["Figura A6:", "Chile: PIB, inflación y tasa de política monetaria, 2022–2026P (%)",
      "Fuente: BCCh (IPoM dic-25, jun-26 y sep-26) e INE. PIB 2026: punto medio del rango 0,25–0,75%."],
    chile: [
      "**Menos crecimiento.** El Banco Central recortó el PIB 2026 a 0,25–0,75%, desde 2–3% en diciembre de 2025, con un mercado laboral más débil y confianza a la baja.",
      "**Inflación persistente.** El IPC sube a 3,9%, sobre la meta de 3%, y la TPM se mantiene en 4,5%: no hay espacio para estimular el consumo, motor de las ventas de los locatarios (MR1).",
      "**Fragmentación geopolítica.** Alianzas debilitadas y una política comercial de EE.UU. poco predecible encarecen la cobertura de riesgos y reducen los flujos hacia emergentes. Para Chile, Perú y Colombia esto significa monedas más volátiles y financiamiento más caro, que el modelo recoge con un CRP ponderado de 1,41% en el Ke.",
    ],
    s5: "5. QUÉ SIGNIFICA PARA MALLPLAZA",
    tabla: {
      cols: ["Canal", "2022–2026", "Efecto en Mallplaza", "En el modelo", "Riesgo"],
      filas: [
        ["Energía e inflación", "Brent USD 71 → 114 (feb–sep-26); inflación mundial 4,1% → 4,7%", "Reajusta la deuda en UF (79% del total); los arriendos en UF compensan", "UF +3% = −CLP 38.691 MM de utilidad (Fig. 26)", "MR1"],
        ["Tasas largas", "UST 10Y 4,1% → 5,2%; Bund 2,9% → 3,6% (dic-25 → oct-26)", "Mayor WACC y costo de refinanciar el pipeline", "Rf 4,98%; WACC +100 pb = precio objetivo −12,3% (Fig. 25)", "FR1"],
        ["Consumo y PIB", "PIB Chile 2026: 0,25–0,75%; IPC 3,9%", "Menores ventas de locatarios, tráfico y nuevos formatos", "Tornado: tráfico −1,3% / +2,4% del valor (Fig. 24)", "MR1 · MR2"],
        ["Fragmentación", "Aranceles de EE.UU.; menores flujos a emergentes", "Monedas volátiles y financiamiento más caro", "CRP ponderado 1,41% en el Ke; deuda en moneda local y derivados de cobertura (Análisis Razonado 2T26)", "FR1 · MR3"],
      ],
      anchos: [1650, 2700, 2700, 2616, 1000],
    },
    vigilar: { titulo: "6. QUÉ VIGILAR", cols: ["Indicador", "Último dato", "Señal de alerta", "Qué implica para Mallplaza"],
      filas: [
        ["UST 10Y", "5,22% (8-oct-26)", "Sobre 5,5% de forma sostenida", "≈ WACC +50 pb: precio objetivo de CLP 2.077 a ~1.940 (Fig. 25)"],
        ["Brent", "USD 114 (sep-26)", "Sobre USD 120 o cierre prolongado de Ormuz", "Más inflación y UF: mayor reajuste de la deuda (MR1)"],
        ["IPC Chile", "3,9% (IPoM jun-26)", "Sobre 4% por dos trimestres", "TPM sin recortes; consumo y ventas de locatarios débiles"],
        ["PIB Chile", "0,25–0,75% 2026 (IPoM)", "Contracción en dos trimestres", "Menor tráfico y ventas; presión sobre la ocupación (Fig. 24)"],
        ["Spread Francia–Alemania", "117 pb (30-sep-26)", "Sobre 150 pb", "Contagio fiscal: más presión sobre las tasas largas globales (FR1)"],
      ], anchos: [1900, 1900, 2700, 4166] },
    cierre: "**Lectura para el inversionista.** El entorno eleva la probabilidad de los dos riesgos más severos del informe (FR1 y MR1). El modelo usa una Rf de 4,98% y una WACC nominal terminal de 10,41% (10,54% en 2S26E); con el UST 10Y sobre 5,2% en octubre, el sesgo de la tasa de descuento es al alza, lo que refuerza la recomendación.",
  },
  en: {
    archivo: "Mallplaza_Anexo_Macro_EN.docx", pie: "A-",
    titulo: "APPENDIX [X]: MACROECONOMIC AND GEOPOLITICAL BACKDROP",
    lead: "Between 2022 and 2026 the world economy absorbed two geopolitical supply shocks and a structural rise in long-term rates. The result is a backdrop of stagflationary pressure and a higher cost of capital that reaches Chile through three channels — inflation, rates and consumption — and that sits behind the two most severe risks in the report (FR1 and MR1). In the charts, the dotted blue line marks the valuation date (30-Jun-2026); later data are shown for context. The figures cover the shock transmission (Figure A1), oil (Figure A2), sovereign yields (Figure A3), spreads (Figure A4), world growth and inflation (Figure A5) and Chile (Figure A6).",
    kpis: [
      ["USD 114", "Brent, Sep-26 average", "+61% vs Feb-26, before the conflict"],
      ["5.22%", "10Y UST, 8-Oct-26", "24-year high (since 2002)"],
      ["117 bp", "France–Germany spread", "30-Sep-26; ~150 bp in early Oct."],
      ["3.0% | 4.7%", "IMF 2026: GDP | inflation", "growth cut; disinflation stalled"],
      ["0.25–0.75%", "Chile 2026 GDP (BCCh)", "down from 2–3% expected in Dec-25"],
    ],
    s1: "1. HOW THE SHOCK REACHES MALLPLAZA",
    a0: ["Figure A1:", "Transmission map: from global shocks to Mallplaza's risks",
      "Source: Team analysis; figures from Figures 24–26 of the report (Investment Risks)."],
    s2: "2. TWO SUPPLY SHOCKS AND A GLOBAL RISE IN RATES",
    a1: ["Figure A2:", "Brent crude oil price, 2022–2026 (USD/barrel)",
      "**Two supply shocks.** Russia's invasion of Ukraine (Feb-22) pushed Brent to USD 123 in June 2022. In 2026, the conflict in the Strait of Hormuz — which carries more than a quarter of seaborne oil — lifted it from USD 71 in February to USD 117 in April; after a truce in June, renewed attacks took it back to USD 114 in September.",
      "Source: EIA (Brent spot, monthly average; STEO Apr-26, Sep-26 and Oct-26); Statista."],
    a2: ["Figure A3:", "10-year sovereign yields, 2022–2026 (%)",
      "**Long-term rates rise together.** The 10Y UST moved from 1.8% (Jan-22) to 5.0% (Sep-26) and touched 5.22% on 8 October, a 24-year high; the Bund went from −0.1% to 3.6%. Fiscal deficits push up the risk-free rate, and Chile has followed the move since mid-2026.",
      "Source: FRED (GS10), Bundesbank, Banque de France (via OECD) and BCCh (10Y BCP); monthly averages."],
    s3: "3. FISCAL PREMIUM AND STAGFLATION",
    a3: ["Figure A4:", "10-year spreads: France–Germany and Chile–US (bp)",
      "**The rise is fiscal and global, not Chile risk.** France's spread over Germany rose from ~40 bp in 2022 to an 81 bp average in September 2026 and 117 bp at month-end (~150 bp in early October, the widest since 2011–12). Chile's spread over the US compressed: local rates depend more and more on the global base rate.",
      "Source: Team calculation with Figure A3 data (monthly averages); September close: Banque de France / Bundesbank."],
    a4: ["Figure A5:", "World growth and inflation, 2022–2026P (%)",
      "**Stagflation.** The IMF cut 2026 world growth to 3.0% and raised inflation to 4.7% (from 4.1% in 2025): the disinflation under way since 2024 stalled on higher energy and food prices, in an economy that grows less.",
      "Source: IMF, World Economic Outlook (July 2026 update). 2026P = projection."],
    s4: "4. CHILE: SLOWER GROWTH AND STICKY INFLATION",
    a5: ["Figure A6:", "Chile: GDP, inflation and policy rate, 2022–2026P (%)",
      "Source: BCCh (Monetary Policy Reports Dec-25, Jun-26 and Sep-26) and INE. 2026 GDP: midpoint of the 0.25–0.75% range."],
    chile: [
      "**Slower growth.** The Central Bank cut 2026 GDP growth to 0.25–0.75%, from 2–3% in December 2025, with a weaker labour market and falling confidence.",
      "**Sticky inflation.** CPI rises to 3.9%, above the 3% target, and the policy rate stays at 4.5%: there is no room to support consumption, the driver of tenant sales (MR1).",
      "**Geopolitical fragmentation.** Weaker alliances and unpredictable US trade policy raise hedging costs and reduce flows to emerging markets. For Chile, Peru and Colombia this means more volatile currencies and costlier funding, which the model captures with a 1.41% weighted CRP in the cost of equity.",
    ],
    s5: "5. WHAT IT MEANS FOR MALLPLAZA",
    tabla: {
      cols: ["Channel", "2022–2026", "Effect on Mallplaza", "In the model", "Risk"],
      filas: [
        ["Energy and inflation", "Brent USD 71 → 114 (Feb–Sep-26); world inflation 4.1% → 4.7%", "Indexes UF debt (79% of total); UF-linked rents offset it", "UF +3% = −CLP 38,691 MM of earnings (Fig. 26)", "MR1"],
        ["Long-term rates", "10Y UST 4.1% → 5.2%; Bund 2.9% → 3.6% (Dec-25 → Oct-26)", "Higher WACC and cost of refinancing the pipeline", "Rf 4.98%; WACC +100 bp = target price −12.3% (Fig. 25)", "FR1"],
        ["Consumption and GDP", "Chile 2026 GDP: 0.25–0.75%; CPI 3.9%", "Lower tenant sales, footfall and demand for new formats", "Tornado: traffic −1.3% / +2.4% of value (Fig. 24)", "MR1 · MR2"],
        ["Fragmentation", "US tariffs; lower flows to emerging markets", "Volatile currencies and costlier funding", "1.41% weighted CRP in Ke; debt in local currency and hedging derivatives (MD&A 2Q26)", "FR1 · MR3"],
      ],
      anchos: [1650, 2700, 2700, 2616, 1000],
    },
    vigilar: { titulo: "6. WHAT TO WATCH", cols: ["Indicator", "Latest", "Warning sign", "What it means for Mallplaza"],
      filas: [
        ["10Y UST", "5.22% (8-Oct-26)", "Sustained above 5.5%", "≈ WACC +50 bp: target price from CLP 2,077 to ~1,940 (Fig. 25)"],
        ["Brent", "USD 114 (Sep-26)", "Above USD 120 or a prolonged Hormuz closure", "Higher inflation and UF: larger debt indexation (MR1)"],
        ["Chile CPI", "3.9% (Jun-26 MPR)", "Above 4% for two quarters", "No rate cuts; weak consumption and tenant sales"],
        ["Chile GDP", "0.25–0.75% 2026 (MPR)", "Two quarters of contraction", "Lower footfall and sales; pressure on occupancy (Fig. 24)"],
        ["France–Germany spread", "117 bp (30-Sep-26)", "Above 150 bp", "Fiscal contagion: more pressure on global long-term rates (FR1)"],
      ], anchos: [1900, 1900, 2700, 4166] },
    cierre: "**Investor takeaway.** The backdrop raises the likelihood of the report's two most severe risks (FR1 and MR1). The model uses a 4.98% Rf and a 10.41% terminal nominal WACC (10.54% in 2H26E); with the 10Y UST above 5.2% in October, the bias on the discount rate is upward, which reinforces the recommendation.",
  },
};

function seccion(texto, antes = 140) {
  return new Paragraph({ spacing: { before: antes, after: 70 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ROJO, space: 1 } },
    children: [run(texto, { bold: true, color: AZUL, size: 16 })] });
}
const caption = (n, t) => new Paragraph({ spacing: { before: 40, after: 30, ...linea() },
  children: [run(n + " ", { bold: true, color: AZUL, size: 14 }), run(t, { color: AZUL, size: 14 })] });
const fuente = (t) => new Paragraph({ spacing: { before: 10, after: 0, ...linea(216) },
  children: [run(t, { size: 11, color: GRIS })] });
const imagen = (archivo, anchoIn, altoIn) => new Paragraph({ spacing: { before: 0, after: 20, line: 240, lineRule: LineRuleType.AUTO }, children: [
  new ImageRun({ type: "png", data: fs.readFileSync(path.join(FIG, archivo)),
    transformation: { width: Math.round(anchoIn * PX), height: Math.round(altoIn * PX) } })] });
const texto = (t, despues = 50) => new Paragraph({ alignment: AlignmentType.JUSTIFIED,
  spacing: { before: 0, after: despues, ...linea() }, children: rico(t) });

function celda(w, children, extra = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders: sinBordes,
    margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP, children, ...extra });
}
function dosColumnas(izq, der) {
  return new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [MEDIA, SEP, MEDIA],
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
    rows: [new TableRow({ children: [celda(MEDIA, izq), celda(SEP, [new Paragraph({ children: [] })]), celda(MEDIA, der)] })] });
}
function bloqueFigura(lang, cod, f, conTexto = true) {
  const out = [caption(f[0], f[1]), imagen(`${lang}_${cod}.png`, MEDIA / 1440, ALTO[cod.slice(0, 2)])];
  if (conTexto) out.push(texto(f[2], 20));
  out.push(fuente(f[f.length - 1]));
  return out;
}

function kpis(X) {
  const n = X.kpis.length, sepK = 90, w = Math.floor((ANCHO - sepK * (n - 1)) / n);
  const resto = ANCHO - w * n - sepK * (n - 1);
  const cols = [];
  X.kpis.forEach((_, i) => { cols.push(w + (i === n - 1 ? resto : 0)); if (i < n - 1) cols.push(sepK); });
  const celdas = [];
  X.kpis.forEach(([valor, etiqueta, nota], i) => {
    celdas.push(new TableCell({ width: { size: cols[celdas.length], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO },
      borders: { ...sinBordes, top: { style: BorderStyle.SINGLE, size: 18, color: i === 0 ? ROJO : AZUL } },
      margins: { top: 70, bottom: 70, left: 110, right: 90 }, children: [
        new Paragraph({ spacing: { after: 10 }, children: [run(valor, { bold: true, size: 26, color: i === 0 ? ROJO : AZUL })] }),
        new Paragraph({ spacing: { after: 10, ...linea(216) }, children: [run(etiqueta, { bold: true, size: 13, color: AZUL })] }),
        new Paragraph({ spacing: { after: 0, ...linea(216) }, children: [run(nota, { size: 12, color: GRIS })] }),
      ] }));
    if (i < n - 1) celdas.push(new TableCell({ width: { size: sepK, type: WidthType.DXA }, borders: sinBordes,
      children: [new Paragraph({ children: [] })] }));
  });
  return new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: cols,
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows: [new TableRow({ children: celdas })] });
}

function tablaImplicancias(X, t = X.tabla, colRiesgo = 4) {
  const { cols, filas, anchos } = t;
  const linea1 = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  const enc = new TableRow({ tableHeader: true, children: cols.map((c, i) => new TableCell({
    width: { size: anchos[i], type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, color: "auto", fill: AZUL },
    borders: sinBordes, margins: { top: 50, bottom: 50, left: 90, right: 70 },
    children: [new Paragraph({ children: [run(c, { bold: true, size: 13, color: "FFFFFF" })] })] })) });
  const cuerpo = filas.map((f, r) => new TableRow({ children: f.map((v, i) => new TableCell({
    width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: r % 2 ? "FFFFFF" : "F7F8FB" },
    borders: { ...sinBordes, bottom: linea1 }, margins: { top: 50, bottom: 50, left: 90, right: 70 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing: linea(216), alignment: i === colRiesgo ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [run(v, { size: 13, bold: i === 0 || i === colRiesgo, color: i === colRiesgo ? ROJO : i === 0 ? AZUL : TINTA })] })] })) }));
  return new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: anchos,
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows: [enc, ...cuerpo] });
}

function documento(lang) {
  const X = L[lang];
  const cuerpo = [
    new Paragraph({ spacing: { after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 2 } },
      children: [run(X.titulo, { bold: true, color: AZUL, size: 28 })] }),
    new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 110, ...linea() },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO }, children: [run(X.lead, { size: 16 })] }),
    kpis(X),
    seccion(X.s1, 160),
    caption(X.a0[0], X.a0[1]),
    imagen(`${lang}_A0_mapa.png`, ANCHO / 1440, ALTO.A0),
    fuente(X.a0[2]),
    seccion(X.s2),
    dosColumnas(bloqueFigura(lang, "A1_brent", X.a1), bloqueFigura(lang, "A2_tasas", X.a2)),
    new Paragraph({ children: [new PageBreak()] }),
    seccion(X.s3, 0),
    dosColumnas(bloqueFigura(lang, "A3_spreads", X.a3), bloqueFigura(lang, "A4_mundo", X.a4)),
    seccion(X.s4),
    dosColumnas(bloqueFigura(lang, "A5_chile", X.a5, false), [new Paragraph({ spacing: { before: 60 }, children: [] }),
      ...X.chile.map((t) => texto(t, 70))]),
    seccion(X.s5),
    tablaImplicancias(X),
    seccion(X.vigilar.titulo),
    tablaImplicancias(X, X.vigilar, -1),
    new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: 160, after: 0, ...linea() },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: "FDECF2" },
      border: { left: { style: BorderStyle.SINGLE, size: 24, color: ROJO, space: 6 } },
      indent: { left: 120, right: 120 }, children: rico(X.cierre, { size: 16, boldColor: ROJO }) }),
  ];
  return ({
    styles: { default: { document: { run: { font: FUENTE, size: 15 }, paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 },
        margin: { top: 1000, right: 620, bottom: 900, left: 620, header: 450, footer: 400 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0", space: 2 } },
        children: [run("Mallplaza (Plaza S.A.)  |  CFA Institute Research Challenge", { size: 14, color: GRIS })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
        run(X.pie, { size: 14, color: GRIS }), new TextRun({ children: [PageNumber.CURRENT], font: FUENTE, size: 14, color: GRIS })] })] }) },
      children: cuerpo,
    }],
  });
}

module.exports = { documento, L };

if (require.main === module) (async () => {
  for (const lang of ["es", "en"]) {
    const salida = path.join(__dirname, L[lang].archivo);
    fs.writeFileSync(salida, await Packer.toBuffer(new Document(documento(lang))));
    console.log(salida);
  }
})();
