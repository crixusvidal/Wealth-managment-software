// Formato principal del informe (A4, texto 6900 + separador 216 + figuras 3550 twips) para las secciones
// que nacen de un borrador Word del equipo: lee el JSON de comun/borrador_a_json.py y lo dibuja sin tocar
// el texto. Las imágenes del borrador se reemplazan por los gráficos del repo (mapa `imagenes`).
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Footer, Header, ImageRun, LineRuleType, PageNumber, Paragraph, ShadingType,
  Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

const AZUL = "14284B", ROJO = "E9004B", TINTA = "1A1A1A", GRIS = "5F5F5F", FONDO = "EEF1F6";
const FUENTE = "Arial", ANCHO = 10666, COL_TEXTO = 6900, COL_SEP = 216, COL_FIG = 3550, MEDIA = 5225;
const PAGINA = { size: { width: 11906, height: 16838 }, margin: { top: 1000, right: 620, bottom: 900, left: 620, header: 450, footer: 400 } };
const ESTILOS = { default: { document: { run: { font: FUENTE, size: 16 }, paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } } } } };

const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || 16, bold: o.bold, italics: o.italics, color: o.color || TINTA });
const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: nada, bottom: nada, left: nada, right: nada };
const linea = (l = 228) => ({ line: l, lineRule: LineRuleType.AUTO });
const vacia = () => new Paragraph({ children: [] });
const titulo = (t, salto = false) => new Paragraph({ pageBreakBefore: salto, spacing: { after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 2 } }, children: [run(t, { bold: true, color: AZUL, size: 30 })] });
const subtitulo = (t, antes = 110) => new Paragraph({ spacing: { before: antes, after: 50 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ROJO, space: 1 } }, children: [run(t, { bold: true, color: AZUL })] });
const caption = (n, t, antes = 100) => new Paragraph({ spacing: { before: antes, after: 30, ...linea() },
  children: [run(n + " ", { bold: true, color: AZUL, size: 14 }), run(t, { color: AZUL, size: 14 })] });
const fuente = (runs) => new Paragraph({ spacing: { before: 10, after: 0, ...linea(216) },
  children: runs.map(([t, , i]) => run(t, { size: 11, color: GRIS, italics: i })) });
const parrafo = (runs) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 60, ...linea() },
  children: runs.map(([t, b, i]) => run(t, b ? { bold: true, color: AZUL, italics: i } : { italics: i })) });
const celda = (w, children, extra = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: sinBordes,
  margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP, children, ...extra });

// Alto proporcional de un PNG (lee el encabezado IHDR)
function altoPng(archivo, anchoIn) {
  const b = fs.readFileSync(archivo);
  return anchoIn * b.readUInt32BE(20) / b.readUInt32BE(16);
}
const imagen = (archivo, anchoIn, altoIn) => new Paragraph({ spacing: { before: 0, after: 20, line: 240, lineRule: LineRuleType.AUTO }, children: [
  new ImageRun({ type: "png", data: fs.readFileSync(archivo), transformation: { width: Math.round(anchoIn * 96), height: Math.round((altoIn || altoPng(archivo, anchoIn)) * 96) } })] });

function tabla(cols, anchos, filas, { derecha = [], destacar = [], negritaCol0 = false } = {}) {
  const borde = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  const enc = new TableRow({ tableHeader: true, children: cols.map((c, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: AZUL }, borders: sinBordes, margins: { top: 35, bottom: 35, left: 80, right: 80 },
    children: [new Paragraph({ alignment: derecha.includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT, children: [run(c, { bold: true, size: 13, color: "FFFFFF" })] })] })) });
  const cuerpo = filas.map((f, r) => new TableRow({ cantSplit: true, children: f.map((v, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: destacar.includes(r) ? "FDF1CC" : r % 2 ? "FFFFFF" : "F7F8FB" },
    borders: { ...sinBordes, bottom: borde }, margins: { top: 25, bottom: 25, left: 80, right: 80 }, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing: linea(216), alignment: derecha.includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [run(String(v), { size: 13, bold: destacar.includes(r) || (negritaCol0 && i === 0), color: negritaCol0 && i === 0 ? AZUL : TINTA })] })] })) }));
  return new Table({ width: { size: anchos.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: anchos,
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows: [enc, ...cuerpo] });
}

const encabezado = () => new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0", space: 2 } },
  children: [run("Mallplaza (Plaza S.A.)  |  CFA Institute Research Challenge", { size: 14, color: GRIS })] })] });
const pie = (p) => new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
  run(p, { size: 14, color: GRIS }), new TextRun({ children: [PageNumber.CURRENT], font: FUENTE, size: 14, color: GRIS })] })] });

// Bloques de una columna: textos, subtítulos, figuras (caption + imagen del repo + fuente)
function columna(bloques, anchoTw, img) {
  const out = [];
  let previo = null;
  for (const b of bloques) {
    const primero = !out.length;
    if (b.tipo === "h1") out.push(subtitulo(b.texto, primero ? 20 : 90));
    else if (b.tipo === "caption") out.push(caption(b.n, b.texto, primero ? 20 : 90));
    else if (b.tipo === "imagen") { const [archivo, alto] = img(b.archivo); out.push(imagen(archivo, anchoTw / 1440, alto)); }
    else if (b.tipo === "fuente" || (b.tipo === "p" && previo === "imagen")) out.push(fuente(b.runs));
    else if (b.tipo === "p") out.push(parrafo(b.runs));
    previo = b.tipo;
  }
  return out.length ? out : [vacia()];
}

// Sección completa a partir del JSON del borrador. opciones.imagenes: { "image1.png": ruta | [ruta, altoIn] }
// Devuelve { cuerpo: [...], anexo: [...] } (el anexo empieza en el título "ANEXO..." del borrador)
function seccionBorrador(json, opciones) {
  const img = (nombre) => {
    const v = opciones.imagenes[nombre];
    if (!v) throw new Error(`Sin gráfico del repo para ${nombre} (${json.origen})`);
    return Array.isArray(v) ? v : [v, null];
  };
  const tablas = { cols: null };
  const partes = { cuerpo: [], anexo: [] };
  let destino = partes.cuerpo;
  json.bloques.forEach((b) => {
    if (opciones.soloCuerpo && destino === partes.anexo) return;
    if (b.tipo === "titulo") {
      const esAnexo = /^(ANEXO|APPENDIX)/.test(b.texto);
      if (esAnexo && destino === partes.cuerpo) destino = partes.anexo;
      if (opciones.soloCuerpo && destino === partes.anexo) return;
      const nivel1 = destino.length === 0 || /^(ANEXO|APPENDIX)/.test(b.texto);
      destino.push(nivel1 ? titulo(b.texto, esAnexo && destino.length > 0) : subtitulo(b.texto, 140));
    } else if (b.tipo === "pagina") {
      b.filas.forEach((f) => destino.push(new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [COL_TEXTO, COL_SEP, COL_FIG],
        borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
        rows: [new TableRow({ children: [celda(COL_TEXTO, columna(f.izq, COL_TEXTO, img)), celda(COL_SEP, [vacia()]), celda(COL_FIG, columna(f.der, COL_FIG, img))] })] })));
    } else if (b.tipo === "figuras") {
      b.filas.forEach((f) => destino.push(new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [MEDIA, COL_SEP, MEDIA],
        borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
        rows: [new TableRow({ cantSplit: true, children: [celda(MEDIA, columna(f[0], opciones.anchoAnexo || MEDIA, img)), celda(COL_SEP, [vacia()]),
          celda(MEDIA, columna(f[1], opciones.anchoAnexo || MEDIA, img))] })] })));
    } else if (b.tipo === "caption") destino.push(caption(b.n, b.texto, 140));
    else if (b.tipo === "fuente") destino.push(fuente(b.runs));
    else if (b.tipo === "parrafo") destino.push(parrafo(b.runs));
    else if (b.tipo === "tabla") {
      const total = b.anchos.reduce((a, x) => a + x, 0);
      const anchos = b.anchos.map((x) => Math.round(x * ANCHO / total));
      anchos[anchos.length - 1] += ANCHO - anchos.reduce((a, x) => a + x, 0);
      const derecha = b.filas[0].map((_, i) => i).filter((i) => b.filas.slice(1).every((f) => /^[\d.,%−-]*$/.test(f[i])));
      destino.push(tabla(b.filas[0], anchos, b.filas.slice(1), { derecha, negritaCol0: true }));
    }
  });
  return partes;
}

module.exports = {
  AZUL, ROJO, TINTA, GRIS, FONDO, FUENTE, ANCHO, COL_TEXTO, COL_SEP, COL_FIG, MEDIA, PAGINA, ESTILOS,
  run, titulo, subtitulo, caption, fuente, parrafo, tabla, imagen, altoPng, encabezado, pie, seccionBorrador, vacia,
};
