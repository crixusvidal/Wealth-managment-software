// Sección Valuación (texto del borrador del equipo, cifras sincronizadas con el modelo v2, citas a figuras)
// + Anexos [B] DCF, [C] supuestos, [E] múltiplos, [F] escenarios y [Z] construcción de la WACC, en español e inglés. Figuras pendientes = recuadro punteado.
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
const FIG = path.join(__dirname, "..", "graficos", "figuras");
const D = JSON.parse(fs.readFileSync(path.join(FIG, "datos_wacc.json"), "utf8"));
const PX = 96;

// ---- Marcado de lo nuevo o cambiado (propuesta TIR + auditoría nominal/real) ----
// true: resalta en amarillo (#ecb100) cada fragmento nuevo o cambiado y antepone la etiqueta «NUEVO»/«NEW»;
// false: quita todas las marcas de una vez (el texto nuevo queda integrado, sin resaltar).
const MARCAR_NUEVO = false;
const AMARILLO = "ECB100";
const NUEVO = { es: [], en: [] };                         // fragmentos marcados (se registran con N)
const N = (lang, texto) => { NUEVO[lang].push(texto); return texto; };
const SUP = JSON.parse(fs.readFileSync(path.join(FIG, "datos_supuestos.json"), "utf8"));   // graficos/supuestos_cascada.py (cascada de supuestos y escudo tributario)
const TP = JSON.parse(fs.readFileSync(path.join(FIG, "datos_tir_proyecto.json"), "utf8"));   // graficos/tir_proyecto.py (modelo v3, hoja TIR_Proyecto)
const TIR = JSON.parse(fs.readFileSync(path.join(FIG, "datos_tir.json"), "utf8"));   // graficos/tir_valuacion.py (modelo v3, hoja TIR_Accionista)
const ALTO = { V7: 2.25, V9: 1.95, V1: 1.85, V2: 1.7, V3: 1.6, V4: 1.8, V13: 2.3, V5: 1.6, V6: 1.75, W1: 1.9 }; // = graficos/valuacion.py

const fmt = (lang) => (x, dec = 1) => {
  const s = x.toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec }).replace("-", "−");
  return lang === "es" ? s.replace(/[.,]/g, (c) => (c === "." ? "," : ".")) : s;
};

const L = {
  es: {
    archivo: "Mallplaza_Valuacion_ES.docx", pag: "Página ", pagA: "A-", titulo: "VALUACIÓN",
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

// Texto del cuerpo (cifras del modelo v4 vía comun/cifras.js) ---------------------------------------
const TV = require("./texto_valuacion");
for (const lang of ["es", "en"]) {
  const T = TV.calcular(lang);
  const tabla = (b) => (b.t ? { t: b.t === "TIR" ? TV.tablaTIR(lang) : TV.tablaEscenarios(lang) } : b);
  L[lang].lead = T.lead;
  const { K, f } = require("../../comun/cifras")(lang);
  const Z = L[lang].anexo;
  Z.notasF[1] = lang === "es"
    ? `Es la elección más conservadora: con el bono en pesos el precio objetivo sube a ~CLP ${f.n(K.rf_pesos_po)}, todavía ${f.pc(Math.abs(K.rf_pesos_po / K.precio - 1), 0)} bajo el precio de CLP ${f.n(K.precio)}. La recomendación no cambia.`
    : `It is the more conservative choice: with the peso bond the target price rises to ~CLP ${f.n(K.rf_pesos_po)}, still ${f.pc(Math.abs(K.rf_pesos_po / K.precio - 1), 0)} below the CLP ${f.n(K.precio)} share price. The recommendation does not change.`;
  Z.notasF[2] = Z.notasF[2].replace("réplica del DCF v2", "réplica del DCF v4").replace("DCF v2 replica", "DCF v4 replica");
  Z.fW[2] = Z.fW[2].replace("modelo v2", "modelo v4").replace("model v2", "model v4");
  Z.sD = `${Z.sD} (${lang === "es" ? "Figura" : "Figure"} Z1)`;   // la figura del anexo se cita en su propio anexo
  L[lang].paginas = [{ texto: T.p1.map(tabla), figuras: T.figs1 }, { texto: T.p2.map(tabla), figuras: T.figs2 }];
}

const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || 16, bold: o.bold, italics: o.italics, color: o.color || TINTA,
  shading: o.nuevo && MARCAR_NUEVO ? { type: ShadingType.CLEAR, color: "auto", fill: AMARILLO } : undefined });
// Divide un texto en trozos marcados / no marcados según los fragmentos registrados con N()
function trozos(lang, texto) {
  if (!MARCAR_NUEVO || !lang) return [{ t: texto, nuevo: false }];
  const m = new Array(texto.length).fill(false);
  for (const f of NUEVO[lang]) for (let i = texto.indexOf(f); f && i >= 0; i = texto.indexOf(f, i + f.length)) m.fill(true, i, i + f.length);
  const out = [];
  for (let i = 0; i < texto.length;) { let j = i; while (j < texto.length && m[j] === m[i]) j++; out.push({ t: texto.slice(i, j), nuevo: m[i] }); i = j; }
  return out;
}
const etiqueta = (lang, o = {}) => [run(lang === "es" ? "NUEVO" : "NEW", { ...o, bold: true, size: 12, nuevo: true }), run(" ", o)];
function marcado(lang, texto, o = {}) {
  const tr = trozos(lang, texto);
  if (!tr.some((x) => x.nuevo)) return [run(texto, o)];
  return [...etiqueta(lang, o), ...tr.map((x) => run(x.t, { ...o, nuevo: x.nuevo }))];
}
// Cada fragmento marcado debe aparecer exactamente una vez en el texto de su idioma (si no, el resaltado sería ambiguo)
function validarMarcas() {
  for (const lang of ["es", "en"]) {
    const textos = [];
    for (const p of L[lang].paginas) {
      for (const b of p.texto) { if (b.p) textos.push(b.p); if (b.h1) textos.push(b.h1); }
      for (const f of p.figuras) textos.push(f[0] + " " + f[1]);
    }
    const todo = textos.join("\u0000");
    for (const f of new Set(NUEVO[lang])) {
      const n = todo.split(f).length - 1;
      if (n !== 1) throw new Error(`Marca «${f.slice(0, 60)}…» (${lang}) aparece ${n} veces`);
    }
  }
}
const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: nada, bottom: nada, left: nada, right: nada };
const linea = (l = 228) => ({ line: l, lineRule: LineRuleType.AUTO });
const vacia = () => new Paragraph({ children: [] });
const titulo = (t, antes = 0, salto = false) => new Paragraph({ pageBreakBefore: salto, spacing: { before: antes, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 2 } },
  children: [run(t, { bold: true, color: AZUL, size: 30 })] });
const subtitulo = (t, antes = 110) => new Paragraph({ spacing: { before: antes, after: 50 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ROJO, space: 1 } }, children: [run(t, { bold: true, color: AZUL })] });
const caption = (n, t, antes = 100, lang) => new Paragraph({ spacing: { before: antes, after: 30, ...linea() },
  children: [run(n + " ", { bold: true, color: AZUL, size: 14 }), ...marcado(lang, t, { color: AZUL, size: 14 })] });
const fuente = (t) => new Paragraph({ spacing: { before: 10, after: 0, ...linea(216) }, children: [run(t, { size: 11, color: GRIS })] });
const imagen = (archivo, anchoIn, altoIn) => new Paragraph({ spacing: { before: 0, after: 20, line: 240, lineRule: LineRuleType.AUTO }, children: [
  new ImageRun({ type: "png", data: fs.readFileSync(path.join(FIG, archivo)), transformation: { width: Math.round(anchoIn * PX), height: Math.round(altoIn * PX) } })] });
const celda = (w, children, extra = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: sinBordes,
  margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP, children, ...extra });

// Página de la sección: texto a la izquierda, figuras (o recuadro de gráfico pendiente) a la derecha
const parrafo = (x, lang) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: 0, after: 60, ...linea() },
  children: Array.isArray(x) ? [run(x[0], { bold: true, color: AZUL }), run(x[1])] : marcado(lang, x) });
const pendiente = (t) => new Table({ width: { size: COL_FIG, type: WidthType.DXA }, columnWidths: [COL_FIG],
  rows: [new TableRow({ height: { value: 1700, rule: "atLeast" }, children: [new TableCell({ width: { size: COL_FIG, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER, shading: { type: ShadingType.CLEAR, color: "auto", fill: "F7F8FB" },
    borders: Object.fromEntries(["top", "bottom", "left", "right"].map((k) => [k, { style: BorderStyle.DASHED, size: 6, color: "9AA3B5" }])),
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(t, { size: 14, italics: true, color: GRIS })] })] })] })] });

// Una fila por bloque de texto: cada figura va en la fila del texto que la cita (p.cortes = [[desde, [figuras]], ...])
function pagina(lang, p) {
  const cortes = p.cortes || [[0, p.figuras.map((_, k) => k)]];
  const rows = cortes.map(([desde, figs], c) => {
    const hasta = c + 1 < cortes.length ? cortes[c + 1][0] : p.texto.length;
    const out = [];
    p.texto.slice(desde, hasta).forEach((b, k) => {
      const i = desde + k;
    if (b.h1) out.push(subtitulo(b.h1, i === 0 ? 20 : 90));
    else if (b.t) out.push(...bloqueTabla(lang, b.t));
    else out.push(parrafo(b.p, lang));
    });
    const izq = out;
    const der = [];
    figs.forEach((n_, k) => {
      const [n, t, archivo, f] = p.figuras[n_];
      const out = der;
      out.push(caption(n, t, k === 0 ? 20 : 90, lang));
      if (archivo) {
        out.push(imagen(`${lang}_${archivo}.png`, COL_FIG / 1440, ALTO[archivo.split("_")[0]]));
        out.push(fuente(f));
      } else {
        out.push(pendiente(f));
      }
    });
    if (!der.length) der.push(vacia());
    return new TableRow({ cantSplit: true, children: [celda(COL_TEXTO, izq), celda(COL_SEP, [vacia()]), celda(COL_FIG, der)] });
  });
  return new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [COL_TEXTO, COL_SEP, COL_FIG],
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows });
}

function tabla(cols, anchos, filas, { derecha = [], destacar = [], sub = [], nuevo = false } = {}) {
  const borde = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  const enc = new TableRow({ tableHeader: true, children: cols.map((c, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: AZUL }, borders: sinBordes, margins: { top: 35, bottom: 35, left: 80, right: 80 },
    children: [new Paragraph({ alignment: derecha.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [run(c, { bold: true, size: 13, color: "FFFFFF" })] })] })) });
  const cuerpo = filas.map((f, r) => new TableRow({ children: f.map((v, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: destacar.includes(r) ? "FDF1CC" : sub.includes(r) ? "E4E8F0" : r % 2 ? "FFFFFF" : "F7F8FB" },
    borders: { ...sinBordes, bottom: borde }, margins: { top: 25, bottom: 25, left: 80, right: 80 }, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing: linea(216), alignment: derecha.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [run(String(v), { size: 13, bold: destacar.includes(r) || sub.includes(r), color: TINTA, nuevo })] })] })) }));
  return new Table({ width: { size: anchos.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: anchos,
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows: [enc, ...cuerpo] });
}

// Tabla de una página de texto: título, tabla y nota (todo marcado como nuevo si t.nuevo)
function bloqueTabla(lang, t) {
  const o = { nuevo: !!t.nuevo };
  return [
    new Paragraph({ keepNext: true, spacing: { before: 40, after: 30, ...linea() }, children: [...(o.nuevo && MARCAR_NUEVO ? etiqueta(lang, { color: AZUL, size: 14 }) : []),
      run(t.titulo, { bold: true, color: AZUL, size: 14, ...o })] }),
    tabla(t.cols, t.anchos, t.filas, { derecha: t.derecha || [2, 3, 4], destacar: t.destacar, ...o }),
    ...(t.nota ? [new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { before: 20, after: 60, ...linea(216) }, children: [run(t.nota, { size: 11, color: GRIS, ...o })] })] : [new Paragraph({ spacing: { after: 60 }, children: [] })]),
  ];
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
    titulo(A.titulo, 0, true),
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

const DA = JSON.parse(fs.readFileSync(path.join(FIG, "datos_anexos.json"), "utf8"));
const NB = "\u00a0"; // celda vacía con la altura de la fuente de la tabla
const lead = (t) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 40, ...linea(216) },
  shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO }, children: [run(t, { size: 14 })] });
const dosColumnas = (izq, der) => new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [MEDIA, COL_SEP, MEDIA],
  borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
  rows: [new TableRow({ cantSplit: true, children: [celda(MEDIA, izq), celda(COL_SEP, [vacia()]), celda(MEDIA, der)] })] });

const anexosValuacion = (lang) => require("./anexo_valuacion")({ titulo, lead, subtitulo, tabla, fuente, dosColumnas, fmt, DA, NB, caption, vacia, celda, imagen, run, AZUL, GRIS, TINTA }).construir(lang);

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
          new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, ...linea() }, shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO }, children: [run(X.lead)] }),
          ...X.paginas.flatMap((p, k) => k ? [new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: [] }), pagina(lang, p)] : [pagina(lang, p)])] },
      { properties: { page: { ...pag, pageNumbers: { start: 1 } } }, headers: { default: encabezado() }, footers: { default: pie(X.pagA) }, children: [...anexosValuacion(lang), ...anexo(lang)] },
    ],
  });
}

// Documento de la sección → ../documento · anexo → ../anexo (archivos separados, mismo formato)
validarMarcas();
module.exports = { documento, L, FIG };

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
