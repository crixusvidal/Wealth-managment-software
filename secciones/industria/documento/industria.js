// Sección Análisis de la Industria y Posicionamiento Competitivo + anexos, en el formato principal del informe.
// Texto: borrador del equipo (inglés: borrador/texto_en.json, traducción fiel del mismo) (borrador/Industria_Mallplaza_BORRADOR_FINAL.docx → borrador/texto.json con
// comun/borrador_a_json.py), sin cambios. Gráficos: los del repo (graficos/industria.py), en vez de las capturas.
// Uso: node industria/documento/industria.js
const fs = require("fs");
const path = require("path");
const { AlignmentType, BorderStyle, Document, Packer, Paragraph, ShadingType, Table, TableCell, TableRow, VerticalAlign, WidthType } = require("docx");
const F = require("../../comun/formato");

const FIG = path.join(__dirname, "..", "graficos", "figuras");
const ARCHIVO = "Mallplaza_Industria_ES.docx";
// Figura 2–6 del cuerpo y anexos A–C del borrador → gráficos de graficos/industria.py
const IMAGENES = { "image1.png": "ind5b_market_share_paises_vertical", "image2.png": "ind3_oferta_vs_consumo", "image3.png": "ind6_porter",
  "image4.png": "ind8_ebitda_m2", "image5.png": "ind4_vencimiento_contratos", "image6.png": "ind2_pipeline_gla",
  "image7.png": "ind7_market_share_chile", "image8.png": "ind1_ocupacion" };
const ALTOS = fs.existsSync(path.join(FIG, "altos.json")) ? JSON.parse(fs.readFileSync(path.join(FIG, "altos.json"), "utf8")) : {};

const imagenesDe = (lang) => Object.fromEntries(Object.entries(IMAGENES).map(([k, n]) => [k, [path.join(FIG, `${lang}_${n}.png`), ALTOS[n]]]));
const textoDe = (lang) => JSON.parse(fs.readFileSync(path.join(__dirname, "borrador", lang === "es" ? "texto.json" : "texto_en.json"), "utf8"));

// Anexos 3 a 7 con el formato del informe (especificación del equipo: modelos/Anexos_Industry_especificacion.xlsx;
// textos y tablas en graficos/figuras/datos_anexos.json, que genera graficos/anexos_industria.py)
const sinBordes = { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } };
const celda = (w, hijos) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: sinBordes, margins: { top: 0, bottom: 0, left: 0, right: 0 },
  verticalAlign: VerticalAlign.TOP, children: hijos });
const dosColumnas = (izq, der, wIzq = F.MEDIA) => new Table({ width: { size: F.ANCHO, type: WidthType.DXA }, columnWidths: [wIzq, F.COL_SEP, F.ANCHO - wIzq - F.COL_SEP],
  borders: { ...sinBordes, insideHorizontal: sinBordes.top, insideVertical: sinBordes.top },
  rows: [new TableRow({ children: [celda(wIzq, [...izq, F.vacia()]), celda(F.COL_SEP, [F.vacia()]), celda(F.ANCHO - wIzq - F.COL_SEP, [...der, F.vacia()])] })] });
const tituloAnexo = (t, antes = 140, salto = false) => new Paragraph({ pageBreakBefore: salto, spacing: { before: antes, after: 70 }, keepNext: true,
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: F.AZUL, space: 2 } }, children: [F.run(t, { bold: true, color: F.AZUL, size: 24 })] });
const recuadro = (t) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, line: 228, lineRule: "auto" },
  shading: { type: ShadingType.CLEAR, color: "auto", fill: F.FONDO }, children: [F.run(t)] });
const fuenteGris = (t) => F.fuente([[t, false, true]]);
const imagenAnexo = (lang, nombre, anchoTw) => F.imagen(path.join(FIG, `${lang}_${nombre}.png`), anchoTw / 1440);

function anexoNuevo(lang) {
  const A = JSON.parse(fs.readFileSync(path.join(FIG, "datos_anexos.json"), "utf8"))[lang];
  const t3 = A.a3_tabla, t6a = A.a6_tablaA, t6b = A.a6_tablaB, t7 = A.a7_tabla;
  const W3 = 4800, anch3 = [900, 640, 640, 640, 740, 700, 740, 600];   // tabla B: 5.600 de ancho; el gráfico A usa 4.800
  return [
    F.titulo(A.titulo_general),
    // Anexo 3: gráfico (A) y tabla (B) lado a lado
    tituloAnexo(A.a3_titulo, 60), recuadro(A.a3_intro),
    dosColumnas(
      [F.subtitulo(A.a3_subA, 40), imagenAnexo(lang, "ind2_pipeline_gla", W3)],
      [F.subtitulo(A.a3_subB, 40), F.tabla(t3.cols, anch3, t3.filas, { alDerecha: [1, 2, 3, 4, 5, 6, 7], destacar: t3.destacar, negritaFilas: t3.negritaFilas }),
        fuenteGris(A.a3_notaB)], W3),
    fuenteGris(A.a3_fuente),
    // Anexos 4 y 5: mitad y mitad
    dosColumnas(
      [tituloAnexo(A.a4_titulo, 160), recuadro(A.a4_intro), F.subtitulo(A.a4_sub, 40), imagenAnexo(lang, "ind1_ocupacion", F.MEDIA), fuenteGris(A.a4_fuente)],
      [tituloAnexo(A.a5_titulo, 160), recuadro(A.a5_intro), F.subtitulo(A.a5_sub, 40), imagenAnexo(lang, "ind7_market_share_chile", F.MEDIA), fuenteGris(A.a5_fuente)]),
    // Anexo 6: Porter, dos tablas
    tituloAnexo(A.a6_titulo, 0, true), recuadro(A.a6_intro),
    F.subtitulo(A.a6_subA, 40),
    F.tabla(t6a.cols, [1500, 600, 560, 780, 560, F.ANCHO - 4000], t6a.filas, { derecha: [1, 2, 3, 4], destacar: t6a.destacar }),
    fuenteGris(A.a6_fuenteA),
    F.subtitulo(A.a6_subB, 120),
    F.tabla(t6b.cols, [2900, 800, 2300, 2300, F.ANCHO - 8300], t6b.filas, { alDerecha: [2, 3, 4], destacar: t6b.destacar }),
    fuenteGris(A.a6_fuenteB),
    // Anexo 7: supuestos
    tituloAnexo(A.a7_titulo, 200), recuadro(A.a7_intro),
    F.tabla(t7.cols, [1800, F.ANCHO - 1800], t7.filas, { negritaCol0: true }),
  ];
}

function documento(lang = "es") {
  const json = JSON.parse(fs.readFileSync(path.join(__dirname, "borrador", lang === "es" ? "texto.json" : "texto_en.json"), "utf8"));
  const imagenes = imagenesDe(lang);
  const { cuerpo } = F.seccionBorrador(json, { imagenes, anchoAnexo: F.COL_FIG });
  const anexo = anexoNuevo(lang);
  return {
    styles: F.ESTILOS,
    sections: [
      { properties: { page: F.PAGINA }, headers: { default: F.encabezado() }, footers: { default: F.pie(lang === "es" ? "Página " : "Page ") }, children: cuerpo },
      { properties: { page: { ...F.PAGINA, pageNumbers: { start: 1 } } }, headers: { default: F.encabezado() }, footers: { default: F.pie("A-") }, children: anexo },
    ],
  };
}

module.exports = { documento, imagenesDe, textoDe, anexoNuevo };

if (require.main === module) (async () => {
  for (const lang of ["es", "en"]) {
    const d = documento(lang);
    const archivo = ARCHIVO.replace("_ES", lang === "es" ? "_ES" : "_EN");
    const salidas = [[0, path.join(__dirname, archivo)], [1, path.join(__dirname, "..", "anexo", archivo.replace(".docx", "_Anexo.docx"))]];
    for (const [k, salida] of salidas) {
      fs.writeFileSync(salida, await Packer.toBuffer(new Document({ ...d, sections: [d.sections[k]] })));
      console.log(salida);
    }
  }
})();
