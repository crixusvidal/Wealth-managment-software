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
const FIG = path.join(__dirname, "..", "graficos", "figuras");
const PX = 96;
const ALTO = { FA2: 1.85, FA3: 1.95, FA4: 1.75, FA5: 2.0, FA6: 2.2, FA7: 2.95 }; // = graficos/analisis_financiero.py

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
// 2026E–2031E (columnas 5–10 de cada fila): proyección del caso base de valoración v4 (informe_final/graficos/cifras_informe.py, tabla_fa1)
{
  const F1 = require("../../comun/cifras")("en").K.fa1;
  const p1 = (x) => (x * 100).toFixed(1), mil = (x) => Math.round(x / 1000).toLocaleString("en-US");
  const proy = [F1.ingresos.map(mil), F1.margen_ebitda.map(p1), F1.margen_ffo.map(p1), F1.utilidad_reportada.map(mil), F1.distribuible.map(mil), F1.margen_neto_norm.map(p1),
    F1.rotacion.map((x) => x.toFixed(3).replace(/^0/, "")), F1.apalancamiento.map((x) => x.toFixed(2)), F1.roe_rep.map(p1), F1.roe_norm.map(p1), F1.roa_norm.map(p1)];
  proy.forEach((r, i) => FA1[i].splice(4, 6, ...r));
}

// ---------------------------------------------------------------------------------------------
const L = {
  es: {
    archivo: "Mallplaza_Analisis_Financiero_ES.docx", pag: "Página ", pagA: "A-",
    titulo: "ANÁLISIS FINANCIERO", tituloA: "ANEXOS: ANÁLISIS FINANCIERO",
    lead: "El motor de arriendos de Mallplaza crece cerca de 3% real, con un margen EBITDA de 80% y un apalancamiento sobre flujos bajo 2,5x; la utilidad reportada, el crecimiento y el apalancamiento contable sobrestiman ese motor y deben normalizarse antes de proyectar.",
    fa1: ["Figura 1:", "El ROE normalizado de 8–10% es el retorno real; el ROE reportado depende de la tasación",
      ["CLP mil MM", "Ingresos", "Margen EBITDA %", "Margen FFO %", "Utilidad neta, reportada", "Utilidad distribuible", "Margen neto, normalizado %", "Rotación de activos (x)", "Activos / patrimonio (x)", "ROE, reportado %", "ROE, normalizado %", "ROA, normalizado %"],
      "Normalizado = utilidad distribuible (2022A: proxy del equipo; el reajuste UF de 2022 refleja ~13% de inflación). Saldos promedio; la serie parte en 2022 (NIC 40 a valor razonable desde el 1-ene-22). 2026E–2031E: caso base de valoración; balance proyectado desde dic-25 (patrimonio + utilidad reportada − distribuciones ordinarias y adicionales; propiedades revalorizadas al 3% de inflación, impuesto diferido 27%; DFN hasta 3,5x EBITDA en 2031E). Fuente: EEFF Plaza S.A. (CMF), Nota 20e; release 2T26 (FFO); análisis del equipo."],
  },
  en: {
    archivo: "Mallplaza_Analisis_Financiero_EN.docx", pag: "Page ", pagA: "A-",
    titulo: "FINANCIAL ANALYSIS", tituloA: "APPENDICES: FINANCIAL ANALYSIS",
    lead: "Mallplaza's rental engine grows at roughly 3% in real terms with an 80% EBITDA margin and cash-flow leverage below 2.5x; headline profit, growth and leverage all overstate that engine and must be normalized before forecasting.",
    fa1: ["Figure 1:", "Normalized ROE of 8–10% is the real return; reported ROE is appraisal-driven",
      ["CLP bn", "Revenue", "EBITDA margin %", "FFO margin %", "Net income, reported", "Distributable profit", "Net margin, normalized %", "Asset turnover (x)", "Assets / equity (x)", "ROE, reported %", "ROE, normalized %", "ROA, normalized %"],
      "Normalized = distributable profit (2022A team proxy; 2022 UF indexation inflated by ~13% inflation). Average balances; series starts in 2022 (IAS 40 fair value applied retroactively from 1 Jan 2022). 2026E–2031E: valuation base case; balance sheet rolled forward from Dec-25 (equity + reported profit − ordinary and additional distributions; property revalued at 3% inflation, 27% deferred tax; net debt rising to 3.5x EBITDA by 2031E). Source: Plaza S.A. financial statements (CMF), Note 20e; 2Q26 results release (FFO); Team analysis."],
  },

};

// Texto del cuerpo (v2 del equipo, cifras del modelo v4 vía comun/cifras.js) -----------------------
const TF = require("./texto_fa");
for (const lang of ["es", "en"]) {
  const T = TF.calcular(lang);
  L[lang].lead = T.lead;
  L[lang].fa1[1] = T.fa1;
  L[lang].paginas = [{ texto: T.p1, figuras: T.figs1 }, { texto: T.p2, figuras: T.figs2 }];
}

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
  const X = require("./anexo_fa")({ caption: (n, t, antes) => caption(n, t, antes), tabla, fuente }).anexos(lang);
  const dos = (izq, der) => new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [MEDIA, COL_SEP, MEDIA],
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
    rows: [new TableRow({ children: [celda(MEDIA, izq), celda(COL_SEP, [vacia()]), celda(MEDIA, der)] })] });
  return [titulo(L[lang].tituloA), dos(X.bloque1, X.bloque2), new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: [] }),
    dos([...X.bloque3], [...X.bloque4, ...X.bloque5, ...X.bloque6])];
}

const encabezado = () => new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0", space: 2 } },
  children: [run("Mallplaza (Plaza S.A.)  |  CFA Institute Research Challenge", { size: 14, color: GRIS })] })] });
const pie = (p) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  run(p, { size: 14, color: GRIS }), new TextRun({ children: [PageNumber.CURRENT], font: FUENTE, size: 14, color: GRIS })] })] });

function documento(lang) {
  const X = L[lang];
  const pag = { size: { width: 11906, height: 16838 }, margin: { top: 1000, right: 620, bottom: 900, left: 620, header: 450, footer: 400 } };
  return ({
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

// Documento de la sección → ../documento · anexo → ../anexo (archivos separados, mismo formato)
module.exports = { documento, L, figura1, FIG, FA1, FA1_AÑOS, loc };

if (require.main === module) (async () => {
  for (const lang of ["es", "en"]) {
    const d = documento(lang);
    const salidas = [[0, path.join(__dirname, L[lang].archivo)],
      [1, path.join(__dirname, "..", "anexo", L[lang].archivo.replace(".docx", "_Anexo.docx"))]];
    for (const [k, salida] of salidas) {
      fs.writeFileSync(salida, await Packer.toBuffer(new Document({ ...d, sections: [d.sections[k]] })));
      console.log(salida);
    }
  }
})();
