// Cuerpo del informe consolidado con el diseño de los informes ganadores: el texto corre seguido en su columna
// (6900 twips, sangría derecha) y las figuras van en la columna derecha (3550 twips) como tablas flotantes, a la
// altura del párrafo que las cita. Si no caben ahí, bajan; si no caben en la página, saltan a la siguiente.
//
// Dos pasadas (las coordina informe_final/graficos/maquetar.py):
//   1. medir:  documento solo con el texto y los bloques de ancho completo + plan_<lang>.json (elementos y figuras)
//   2. armar:  con posiciones_<lang>.json (página, párrafo ancla y desplazamiento de cada figura, calculados
//              sobre el PDF de la pasada 1), inserta cada figura flotante antes de su párrafo ancla.
const fs = require("fs");
const path = require("path");
const docx = require("docx");
const F = require("../../comun/formato");

const { Paragraph, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, ShadingType, LineRuleType,
  TableAnchorType, OverlapType } = docx;
const SANGRIA = { right: F.COL_SEP + F.COL_FIG };
const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: nada, bottom: nada, left: nada, right: nada };

// Párrafos de la columna de texto ----------------------------------------------------------------
const pCol = (runs) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, indent: SANGRIA,
  spacing: { before: 0, after: 60, line: 228, lineRule: LineRuleType.AUTO },
  children: runs.map(([t, b, i]) => F.run(t, b ? { bold: true, color: F.AZUL, italics: i } : { italics: i })) });
const h1Col = (t) => new Paragraph({ indent: SANGRIA, spacing: { before: 110, after: 50 }, mantenerConSiguiente: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: F.ROJO, space: 1 } }, children: [F.run(t, { bold: true, color: F.AZUL })] });
const tituloCol = (t) => new Paragraph({ indent: SANGRIA, spacing: { before: 120, after: 80 }, mantenerConSiguiente: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: F.AZUL, space: 2 } }, children: [F.run(t, { bold: true, color: F.AZUL, size: 30 })] });
const leadAncho = (t, destacado) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, indent: SANGRIA, spacing: { after: 80, line: 228, lineRule: LineRuleType.AUTO },
  shading: { type: ShadingType.CLEAR, color: "auto", fill: F.FONDO }, children: [F.run(t, destacado ? { bold: true, color: F.AZUL } : {})] });
const textoRuns = (runs) => runs.map((r) => r[0]).join("");

// Figura: título + gráfico + fuente, en una tabla de la columna derecha
const ALTO_MAX = 1.8;   // pulgadas
// Escala del alto máximo por sección (la ajusta maquetar.py para que todas las figuras quepan al costado del texto)
const ESCALAS = {};
function medidas(fig, anchoTw = F.COL_FIG, altoMax = ALTO_MAX * (ESCALAS[fig.clave] || 1)) {
  const w = anchoTw / 1440, h = F.altoPng(fig.archivo, w);
  return h > altoMax ? [w * altoMax / h, altoMax] : [w, h];
}
function figuraHijos(fig, anchoTw, altoMax) {
  const [w, h] = medidas(fig, anchoTw, altoMax);
  const img = new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 20, line: 240, lineRule: LineRuleType.AUTO }, children: [
    new docx.ImageRun({ type: "png", data: fs.readFileSync(fig.archivo), transformation: { width: Math.round(w * 96), height: Math.round(h * 96) } })] });
  return [F.caption(fig.n, fig.titulo, 0), img,
    new Paragraph({ spacing: { before: 10, after: 0, line: 216, lineRule: LineRuleType.AUTO },
      children: fig.fuente.map(([t, , i]) => F.run(t, { size: 11, color: F.GRIS, italics: i })) })];
}
function figuraFlotante(fig, offsetTw, marcador = false) {
  return new Table({ width: { size: F.COL_FIG, type: WidthType.DXA }, columnWidths: [F.COL_FIG],
    float: { horizontalAnchor: TableAnchorType.MARGIN, verticalAnchor: TableAnchorType.TEXT,
      absoluteHorizontalPosition: F.COL_TEXTO + F.COL_SEP, absoluteVerticalPosition: Math.round(offsetTw),
      overlap: OverlapType.NEVER, leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0 },
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
    rows: [new TableRow({ children: [new TableCell({ width: { size: F.COL_FIG, type: WidthType.DXA }, borders: sinBordes,
      margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: marcador ? [F.vacia()] : figuraHijos(fig) })] })] });
}
// Alto estimado de la figura (pulgadas): título + gráfico + fuente
const altoFigura = (fig) => 0.05 + Math.ceil((fig.n.length + fig.titulo.length + 1) / 52) * 0.135
  + medidas(fig)[1] + 0.04 + Math.ceil(textoRuns(fig.fuente).length / 78) * 0.11 + 0.12;

// Construcción de los elementos -------------------------------------------------------------------
// elemento = { tipo: "texto" | "ancho", hijos: [docx], inicio: "primeras palabras para ubicarlo en el PDF" }
// figura   = { id, n, titulo, archivo, alto, fuente: runs, cita: índice del elemento que la cita }
function nuevoPlan() { return { elementos: [], figuras: [] }; }
const ancho = (plan, hijos, inicio, barrera = false) => plan.elementos.push({ tipo: "ancho", hijos, inicio, barrera });
const texto = (plan, hijo, inicio) => plan.elementos.push({ tipo: "texto", hijos: [hijo], inicio }) - 1;

// Asigna cada figura al primer párrafo de la sección que la cita (o al primero de la sección)
function citar(plan, desde, figs) {
  figs.forEach((fig) => {
    const num = fig.n.replace(/^(Figura|Figure)\s+/, "").replace(/:$/, "");
    const rx = new RegExp(`(Figuras?|Figures?)\\s+(\\S+[–-])?${num.replace(/[.]/g, "\\.")}(?![\\w.])`, "i");
    let i = plan.elementos.findIndex((e, k) => k >= desde && e.tipo === "texto" && rx.test(e.inicioCompleto || ""));
    if (i < 0) i = plan.elementos.findIndex((e, k) => k >= desde && e.tipo === "texto");
    plan.figuras.push({ ...fig, id: plan.figuras.length, cita: i });
  });
}

// Secciones que vienen de un borrador (JSON): títulos y párrafos al texto; figuras de la columna derecha
function seccionBorrador(plan, json, imagenes, clave) {
  const desde = plan.elementos.length, figs = [];
  for (const b of json.bloques) {
    if (b.tipo === "titulo") { if (/^(ANEXO|APPENDI)/.test(b.texto)) break; texto(plan, tituloCol(b.texto), b.texto); continue; }
    if (b.tipo !== "pagina") continue;
    for (const f of b.filas) {
      for (const x of f.izq) {
        if (x.tipo === "h1") texto(plan, h1Col(x.texto), x.texto);
        else if (x.tipo === "p") { const i = texto(plan, pCol(x.runs), textoRuns(x.runs)); plan.elementos[i].inicioCompleto = textoRuns(x.runs); }
      }
      let g = null;
      for (const x of f.der) {
        if (x.tipo === "caption") { g = { n: x.n, titulo: x.texto, fuente: [] }; figs.push(g); }
        else if (x.tipo === "imagen" && g) { const v = imagenes[x.archivo]; [g.archivo, g.alto] = Array.isArray(v) ? v : [v, null]; }
        else if (g && (x.tipo === "fuente" || x.tipo === "p")) g.fuente = x.runs;
      }
    }
  }
  citar(plan, desde, figs.filter((x) => x.archivo).map((x) => ({ ...x, clave })));
}

// Secciones escritas en .js (FA, Valuación): L[lang] con titulo, lead y paginas[{texto, figuras}]
function seccionModulo(plan, X, carpeta, lang, clave, extra = {}) {
  const desde = plan.elementos.length, figs = [];
  texto(plan, tituloCol(X.titulo), X.titulo);
  texto(plan, leadAncho(X.lead, extra.leadDestacado), X.lead);
  if (extra.antes) ancho(plan, extra.antes, extra.inicioAntes);
  for (const p of X.paginas) {
    for (const b of p.texto) {
      if (b.h1) texto(plan, h1Col(b.h1), b.h1);
      else if (b.p) {
        const runs = Array.isArray(b.p) ? [[b.p[0], true, false], [b.p[1], false, false]] : [[b.p, false, false]];
        const i = texto(plan, pCol(runs), textoRuns(runs)); plan.elementos[i].inicioCompleto = textoRuns(runs);
      }
    }
    for (const [n, t, archivo, f] of p.figuras) {
      if (!archivo) continue;
      figs.push({ n, titulo: t, archivo: path.join(carpeta, `${lang}_${archivo}.png`), alto: null, fuente: [[f, false, false]], clave });
    }
  }
  citar(plan, desde, figs);
}

// Documento: medir (sin figuras) o armar (figuras flotantes en sus posiciones) ---------------------
// envolver(fig, fn): construye la figura con la renumeración de su sección (la fija informe_final.js)
// finales: { índice de elemento: [ids] } — figuras que no caben junto a su texto: van en una grilla de dos
// columnas justo antes de ese bloque de ancho completo (cierre de la sección)
function grilla(figs, envolver) {
  const filas = [];
  for (let k = 0; k < figs.length; k += 2) filas.push(figs.slice(k, k + 2));
  const celda = (fig) => new TableCell({ width: { size: F.MEDIA, type: WidthType.DXA }, borders: sinBordes,
    margins: { top: 0, bottom: 120, left: 0, right: 0 }, children: fig ? envolver(fig, () => figuraHijos(fig, F.MEDIA, 2.6)) : [F.vacia()] });
  return [new Paragraph({ spacing: { before: 120 }, children: [] }),
    new Table({ width: { size: F.ANCHO, type: WidthType.DXA }, columnWidths: [F.MEDIA, F.COL_SEP, F.MEDIA],
      borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
      rows: filas.map(([a, b]) => new TableRow({ cantSplit: true, noPartir: true, children: [celda(a), new TableCell({ width: { size: F.COL_SEP, type: WidthType.DXA }, borders: sinBordes, children: [F.vacia()] }), celda(b)] })) })];
}
function hijos(plan, posiciones, envolver = (fig, fn) => fn(), finales = {}) {
  const porElemento = {};
  for (const [i, ids] of Object.entries(finales)) (porElemento[i] ||= []).push(...grilla(ids.map((id) => plan.figuras[id]), envolver));
  if (posiciones) for (const fig of plan.figuras) {
    const pos = posiciones[fig.id];
    if (pos) (porElemento[pos.elemento] ||= []).push(envolver(fig, () => figuraFlotante(fig, pos.offset, pos.marcador)));
  }
  return plan.elementos.flatMap((e, i) => [...(porElemento[i] || []), ...e.hijos]);
}
const resumenPlan = (plan) => ({
  elementos: plan.elementos.map((e, i) => ({ i, tipo: e.tipo, barrera: !!e.barrera, inicio: (e.inicio || "").slice(0, 120) })),
  figuras: plan.figuras.map((f) => ({ id: f.id, n: f.n, cita: f.cita, clave: f.clave, alto: altoFigura(f) })),
});

module.exports = { ESCALAS, nuevoPlan, ancho, texto, seccionBorrador, seccionModulo, hijos, resumenPlan };
