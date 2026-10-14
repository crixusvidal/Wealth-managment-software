// Informe consolidado: une las secciones en el orden de la plantilla del equipo y numera de corrido las
// figuras (1…N) y los anexos (1…M), actualizando cada cita del texto. El texto de cada sección no cambia:
// solo se reemplazan las etiquetas de figura/anexo según el mapa de renumeración de esa sección.
//
// Cuerpo: Abstract → Negocio → Industria → ESG → Análisis Financiero → Valuación → Riesgos → Resumen de inversión
// Anexos: Negocio → Industria → ESG → Análisis Financiero → Valuación → Macro
// Todo corrido: cada sección (y cada anexo) empieza inmediatamente después de la anterior; solo el bloque
// de anexos parte en página nueva (cambia la numeración a A-1).
//
// Uso (desde secciones/): node informe_final/documento/informe_final.js   → Mallplaza_Informe_CFA_{ES,EN}.docx
// Requiere: informe_final/graficos/datos_resumen.json (python informe_final/graficos/datos_resumen.py ...)
const fs = require("fs");
const path = require("path");
const docx = require("docx");

// Renumeración: se intercepta cada TextRun creado mientras se construye una sección -----------------
let MAPA = null;
const AVISOS = new Set();
const TOKEN = /\[[A-Z]\]|[A-Z]?\d+(?:\.\d+)?|[A-Z](?![\wÁÉÍÓÚáéíóúñ])/g;
const CITA = new RegExp(String.raw`\b(Figuras?|Figures?|Fig\.|figuras?|Anexos?|ANEXOS?|Appendix|Appendices|APPENDIX)(\s+)`
  + String.raw`((?:\[[A-Z]\]|[A-Z]?\d+|[A-Z](?![\wÁÉÍÓÚáéíóúñ]))(?:\s*(?:–|-|,|y|and)\s*(?:\[[A-Z]\]|[A-Z]?\d+|[A-Z](?![\wÁÉÍÓÚáéíóúñ])))*)`, "g");
function renumerar(texto) {
  return texto.replace(CITA, (todo, palabra, esp, lista) => {
    const tipo = /^f/i.test(palabra) ? "fig" : "anx";
    return palabra + esp + lista.replace(TOKEN, (tok) => {
      const nuevo = MAPA[tipo][tok];
      if (nuevo === undefined) { AVISOS.add(`${MAPA.nombre}: "${palabra} ${tok}" sin mapa (se deja igual)`); return tok; }
      return String(nuevo);
    });
  });
}
const TextRunOriginal = docx.TextRun;
docx.TextRun = class extends TextRunOriginal {
  constructor(o) {
    if (MAPA && typeof o === "string") o = renumerar(o);
    else if (MAPA && o && typeof o.text === "string") o = { ...o, text: renumerar(o.text) };
    super(o);
  }
};

// Modo continuo: sin saltos de página forzados ni filas que no puedan partirse entre páginas -----
const ParagraphOriginal = docx.Paragraph;
const esTitulo = () => false;   // sin "conservar con el siguiente": en LibreOffice impide partir las filas
docx.Paragraph = class extends ParagraphOriginal {
  constructor(o) {
    if (o && typeof o === "object") o = { ...o, pageBreakBefore: false, ...(esTitulo(o) ? { keepNext: true } : {}) };
    super(o);
  }
};
const TableRowOriginal = docx.TableRow;
docx.TableRow = class extends TableRowOriginal {
  constructor(o) { super(o && o.cantSplit ? { ...o, cantSplit: false } : o); }
};
docx.PageBreak = class extends TextRunOriginal { constructor() { super({ text: "" }); } };

const { Document, Packer, SectionType, Paragraph, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, VerticalAlign, AlignmentType } = docx;
const F = require("../../comun/formato");
F.OPCIONES.filasPorCita = true;
const S = path.join(__dirname, "..", "..");
const negocio = require(path.join(S, "negocio/documento/negocio.js"));
const industria = require(path.join(S, "industria/documento/industria.js"));
const esg = require(path.join(S, "esg/documento/esg.js"));
const fa = require(path.join(S, "analisis_financiero/documento/analisis_financiero.js"));
const valuacion = require(path.join(S, "valuacion/documento/valuacion.js"));
const riesgos = require(path.join(S, "riesgos/documento/pagina_riesgos.js"));
const macro = require(path.join(S, "macro/anexo/anexo_macro.js"));

// Mapas: etiqueta original de cada sección → número final ---------------------------------------
const rango = (claves, desde) => Object.fromEntries(claves.map((k, i) => [k, desde + i]));
const sub = (claves, anexo) => Object.fromEntries(claves.map((k, i) => [k, `${anexo}.${i + 1}`]));
const MAPAS = {
  negocio: { nombre: "Negocio", fig: { 1: 1 }, anx: { A: 1, B: 2 } },
  industria: { nombre: "Industria", fig: rango(["2", "3", "4", "5", "6"], 2), anx: rango(["A", "B", "C", "D", "E"], 3) },
  esg: { nombre: "ESG", fig: { ...rango(["E1", "E2", "E3", "E4", "E5"], 7), ...sub(["S1", "S2", "S3", "S4"], 8) }, anx: { "[Y]": 8 } },
  fa: { nombre: "Análisis Financiero", fig: rango(["1", "2", "3", "4", "5", "6", "7"], 12), anx: rango(["1", "2", "3", "4", "5", "6"], 9) },
  valuacion: { nombre: "Valuación", fig: { ...rango(["27", "28", "29", "30", "31", "32", "33", "34"], 19), ...sub(["Z1"], 19) },
    anx: { "[B]": 15, "[C]": 16, "[E]": 17, "[F]": 18, "[Z]": 19 } },
  riesgos: { nombre: "Riesgos", fig: rango(["23", "24", "25", "26"], 27), anx: { "[X]": 20 } },
  macro: { nombre: "Macro", fig: { ...sub(["A1", "A2", "A3", "A4", "A5", "A6"], 20), ...rango(["24", "25", "26"], 28) }, anx: { "[X]": 20 } },
};
const con = (clave, fn) => { MAPA = MAPAS[clave]; try { return fn(); } finally { MAPA = null; } };

// Resumen de inversión: datos clave del modelo; el texto lo redacta el equipo ----------------------
const RES = {
  es: { titulo: "RESUMEN DE INVERSIÓN", pend: "PENDIENTE DE REDACCIÓN", datos: "DATOS CLAVE", loc: "de-DE",
    filas: ["Recomendación", "Precio objetivo (CLP)", "Precio 30-sep-26 (CLP)", "Potencial", "Valor DCF / múltiplos (CLP)", "Ponderación DCF / múltiplos",
      "Ticker", "Capitalización (CLP mil MM)", "Rango 52 semanas (CLP)", "Dividendo esperado 2027E (CLP)", "DFN / EBITDA UDM 2T26"], rec: "VENDER",
    t1: "Recomendación y precio objetivo · Tesis de inversión · Resumen de la valuación · Catalizadores · Riesgos principales.",
    t2: "Sección por redactar por el equipo; los datos clave de la derecha ya salen de los modelos.",
    fuente: "Fuente: modelo de valoración v2; Bloomberg (rango 52 semanas); análisis del equipo." },
  en: { titulo: "INVESTMENT SUMMARY", pend: "TO BE WRITTEN", datos: "KEY DATA", loc: "en-US",
    filas: ["Recommendation", "Target price (CLP)", "Price 30-Sep-26 (CLP)", "Upside / downside", "DCF / multiples value (CLP)", "DCF / multiples weight",
      "Ticker", "Market cap (CLP bn)", "52-week range (CLP)", "Expected dividend 2027E (CLP)", "Net debt / LTM EBITDA 2Q26"], rec: "SELL",
    t1: "Recommendation and target price · Investment thesis · Valuation summary · Catalysts · Key risks.",
    t2: "Section to be written by the team; the key data on the right already come from the models.",
    fuente: "Source: valuation model v2; Bloomberg (52-week range); Team analysis." },
};
function resumen(lang) {
  const X = RES[lang];
  const R = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "graficos", "datos_resumen.json"), "utf8"));
  const n = (x, d = 0) => x.toLocaleString(X.loc, { minimumFractionDigits: d, maximumFractionDigits: d });
  const pc = (x, d = 1) => n(x * 100, d).replace("-", "−") + "%";
  const valores = [X.rec, n(R.precio_objetivo), n(R.precio), pc(R.potencial), `${n(R.dcf)} / ${n(R.multiplos)}`,
    `${n(R.peso_dcf * 100)}% / ${n((1 - R.peso_dcf) * 100)}%`, "MALLPLAZ CI", n(R.capitalizacion_mm / 1000),
    `${n(R.rango_52s[0])} – ${n(R.rango_52s[1])}`, `${n(R.dividendo)} (${pc(R.dividend_yield)})`, `${n(R.dfn_mm / R.ebitda_udm_mm, 2)}x`];
  const filas = X.filas.map((t, i) => [t, valores[i]]);
  const pendiente = (t) => new Paragraph({ spacing: { after: 60 }, children: [F.run(t, { italics: true, color: F.GRIS })] });
  const izq = [F.subtitulo(X.pend, 20), pendiente(X.t1), pendiente(X.t2)];
  const der = [F.subtitulo(X.datos, 20), F.tabla(["", ""], [1900, 1650], filas, { derecha: [1], negritaCol0: true }),
    new Paragraph({ spacing: { before: 30 }, children: [F.run(X.fuente, { size: 11, color: F.GRIS })] })];
  const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const celda = (w, children) => new TableCell({ width: { size: w, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
    borders: { top: nada, bottom: nada, left: nada, right: nada }, children });
  const sinB = { top: nada, bottom: nada, left: nada, right: nada, insideHorizontal: nada, insideVertical: nada };
  const cuerpo = new Table({ width: { size: F.ANCHO, type: WidthType.DXA }, columnWidths: [F.COL_TEXTO, F.COL_SEP, F.COL_FIG], borders: sinB,
    rows: [new TableRow({ children: [celda(F.COL_TEXTO, izq), celda(F.COL_SEP, [F.vacia()]), celda(F.COL_FIG, der)] })] });
  // Todo el resumen (título incluido) en una fila que no se parte: si no cabe, pasa entero a la página siguiente
  return [new Table({ width: { size: F.ANCHO, type: WidthType.DXA }, columnWidths: [F.ANCHO], borders: sinB,
    rows: [new TableRowOriginal({ cantSplit: true, children: [celda(F.ANCHO, [F.titulo(X.titulo), cuerpo, F.vacia()])] })] })];
}

// Abstract: tesis en un párrafo + franja de datos clave (cifras del modelo v2) --------------------
const ABS = {
  es: { titulo: "ABSTRACT", loc: "de-DE",
    texto: "Mallplaza (Plaza S.A.) es el mayor operador de centros comerciales de la región andina: 37 centros y 2,36 millones de m² de GLA en Chile, Perú y Colombia, con ingresos casi tan predecibles como un bono (81,4% arriendos, 93% de renta fija y reajustada en UF en Chile) y un margen EBITDA de 80,5%. Es un gran negocio, pero a CLP 3.555 el mercado lo paga como si sus flujos estuvieran casi libres de riesgo. Ponderando 80% nuestro DCF y 20% los múltiplos de pares, llegamos a un precio objetivo de CLP 2.076 y recomendamos VENDER. El desacuerdo con el mercado no está en los flujos, sino en la tasa: el precio exige una WACC de 7,56%, 285 pb bajo nuestro 10,41%, y un cap rate de 6,0% frente al 10,1% que implican nuestros flujos. Crecer no rescata el precio, porque cada m² nuevo rinde 5,0% real frente a una WACC real de 7,2%, y ni el escenario bull (CLP 2.627) ni la tasa de los tasadores (CLP 2.719) alcanzan la cotización. El riesgo principal para la tesis es una nueva compresión de las tasas largas.",
    rotulos: ["Recomendación", "Precio objetivo", "Precio (30-sep-26)", "Potencial"], rec: "VENDER" },
  en: { titulo: "ABSTRACT", loc: "en-US",
    texto: "Mallplaza (Plaza S.A.) is the largest shopping center operator in the Andean region: 37 centers and 2.36 million m² of GLA in Chile, Peru and Colombia, with revenue almost as predictable as a bond (81.4% rents, 93% fixed and UF-indexed in Chile) and an 80.5% EBITDA margin. It is a great business, but at CLP 3,555 the market prices it as if its cash flows were nearly risk-free. Weighting our DCF at 80% and peer multiples at 20%, we reach a target price of CLP 2,076 and recommend SELL. Our disagreement with the market is not about cash flows but about the discount rate: the price requires a 7.56% WACC, 285 bp below our 10.41%, and a 6.0% cap rate against the 10.1% implied by our cash flows. Growth does not rescue the price, because each new m² earns a 5.0% real return against a 7.2% real WACC, and neither the bull case (CLP 2,627) nor the appraisers' rate (CLP 2,719) reaches the share price. The main risk to the thesis is a further compression of long-term rates.",
    rotulos: ["Recommendation", "Target price", "Price (30-Sep-26)", "Upside / downside"], rec: "SELL" },
};
function abstract(lang) {
  const X = ABS[lang];
  const R = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "graficos", "datos_resumen.json"), "utf8"));
  const n = (x, d = 0) => x.toLocaleString(X.loc, { minimumFractionDigits: d, maximumFractionDigits: d });
  const valores = [X.rec, `CLP ${n(R.precio_objetivo)}`, `CLP ${n(R.precio)}`, n(R.potencial * 100, 1).replace("-", "−") + "%"];
  const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const ancho = F.ANCHO / 4;
  const caja = (i) => new TableCell({ width: { size: ancho, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, color: "auto", fill: i === 0 ? F.ROJO : F.AZUL },
    borders: { top: nada, bottom: nada, left: { style: BorderStyle.SINGLE, size: 12, color: "FFFFFF" }, right: nada },
    margins: { top: 60, bottom: 60, left: 120, right: 120 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [F.run(X.rotulos[i], { size: 13, color: "FFFFFF" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [F.run(valores[i], { size: 26, bold: true, color: "FFFFFF" })] })] });
  return [F.titulo(X.titulo),
    new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 240, lineRule: "auto" },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: F.FONDO }, children: [F.run(X.texto, { size: 17 })] }),
    new Table({ width: { size: F.ANCHO, type: WidthType.DXA }, columnWidths: [ancho, ancho, ancho, ancho],
      borders: { top: nada, bottom: nada, left: nada, right: nada, insideHorizontal: nada, insideVertical: nada },
      rows: [new TableRow({ children: [0, 1, 2, 3].map(caja) })] }),
    new Paragraph({ spacing: { after: 160 }, children: [] })];
}

function informe(lang = "es") {
  const d = {
    negocio: con("negocio", () => negocio.documento(lang)),
    industria: con("industria", () => industria.documento(lang)),
    esg: con("esg", () => esg.documento(lang)),
    fa: con("fa", () => fa.documento(lang)),
    valuacion: con("valuacion", () => valuacion.documento(lang)),
    riesgos: con("riesgos", () => riesgos.documento(Object.values(riesgos.LANG).find((x) => x.archivo.endsWith(`_${lang.toUpperCase()}.docx`)))),
    macro: con("macro", () => macro.documento(lang)),
  };
  const pagina = (pageNumbers) => ({ ...F.PAGINA, ...(pageNumbers ? { pageNumbers } : {}) });
  const seccion = (s, tipo, pageNumbers, pie) => ({ ...s, properties: { ...s.properties, type: tipo, page: pagina(pageNumbers) },
    headers: { default: F.encabezado() }, footers: { default: F.pie(pie) } });
  const CONT = SectionType.CONTINUOUS, NUEVA = SectionType.NEXT_PAGE, P = lang === "es" ? "Página " : "Page ", A = "A-";
  // Separador fino entre secciones corridas (la sección siguiente parte con su título)
  const conEspacio = (s) => ({ ...s, children: [new Paragraph({ spacing: { before: 160 }, children: [] }), ...s.children] });
  // Secciones que deben partir en página nueva porque su título quedaría solo al pie (lo decide
  // informe_final/graficos/maquetar.py revisando el PDF); el resto corre seguido.
  const archivoSaltos = path.join(__dirname, `saltos_${lang}.json`);
  const saltos = new Set(fs.existsSync(archivoSaltos) ? JSON.parse(fs.readFileSync(archivoSaltos, "utf8")) : []);
  const corrida = (clave, s, pie) => saltos.has(clave) ? seccion(s, NUEVA, null, pie) : seccion(conEspacio(s), CONT, null, pie);
  return new Document({
    styles: F.ESTILOS,
    sections: [
      seccion({ ...d.negocio.sections[0], children: [...abstract(lang), ...d.negocio.sections[0].children] }, NUEVA, { start: 1 }, P),
      corrida("industria", d.industria.sections[0], P),
      corrida("esg", d.esg.sections[0], P),
      corrida("fa", d.fa.sections[0], P),
      corrida("valuacion", d.valuacion.sections[0], P),
      corrida("riesgos", d.riesgos.sections[0], P),
      corrida("resumen", { children: resumen(lang) }, P),
      seccion(d.negocio.sections[1], NUEVA, { start: 1 }, A),
      corrida("a_industria", d.industria.sections[1], A),
      corrida("a_esg", d.esg.sections[1], A),
      corrida("a_fa", d.fa.sections[1], A),
      corrida("a_valuacion", d.valuacion.sections[1], A),
      corrida("a_macro", d.macro.sections[0], A),
    ],
  });
}

module.exports = { informe, MAPAS };

if (require.main === module) (async () => {
  const langs = process.argv.slice(2).length ? process.argv.slice(2) : ["es", "en"];
  for (const lang of langs) {
    const salida = path.join(__dirname, `Mallplaza_Informe_CFA_${lang.toUpperCase()}.docx`);
    fs.writeFileSync(salida, await Packer.toBuffer(informe(lang)));
    console.log(salida);
  }
  if (AVISOS.size) console.log("Avisos de renumeración:\n  " + [...AVISOS].join("\n  "));
})();
