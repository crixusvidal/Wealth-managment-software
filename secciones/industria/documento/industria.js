// Sección Análisis de la Industria y Posicionamiento Competitivo + anexos, en el formato principal del informe.
// Texto: borrador del equipo (borrador/Industria_Mallplaza_BORRADOR_FINAL.docx → borrador/texto.json con
// comun/borrador_a_json.py), sin cambios. Gráficos: los del repo (graficos/industria.py), en vez de las capturas.
// Uso: node industria/documento/industria.js
const fs = require("fs");
const path = require("path");
const { Document, Packer } = require("docx");
const F = require("../../comun/formato");

const FIG = path.join(__dirname, "..", "graficos", "figuras");
const ARCHIVO = "Mallplaza_Industria_ES.docx";
// Figura 2–6 del cuerpo y anexos A–C del borrador → gráficos de graficos/industria.py
const IMAGENES = { "image1.png": "ind5b_market_share_paises_vertical", "image2.png": "ind3_oferta_vs_consumo", "image3.png": "ind6_porter",
  "image4.png": "ind8_ebitda_m2", "image5.png": "ind4_vencimiento_contratos", "image6.png": "ind2_pipeline_gla",
  "image7.png": "ind7_market_share_chile", "image8.png": "ind1_ocupacion" };
const ALTOS = fs.existsSync(path.join(FIG, "altos.json")) ? JSON.parse(fs.readFileSync(path.join(FIG, "altos.json"), "utf8")) : {};

function documento(lang = "es") {
  const json = JSON.parse(fs.readFileSync(path.join(__dirname, "borrador", "texto.json"), "utf8"));
  const imagenes = Object.fromEntries(Object.entries(IMAGENES).map(([k, n]) => [k, [path.join(FIG, `${lang}_${n}.png`), ALTOS[n]]]));
  const { cuerpo, anexo } = F.seccionBorrador(json, { imagenes, anchoAnexo: F.COL_FIG });
  return {
    styles: F.ESTILOS,
    sections: [
      { properties: { page: F.PAGINA }, headers: { default: F.encabezado() }, footers: { default: F.pie("Página ") }, children: cuerpo },
      { properties: { page: { ...F.PAGINA, pageNumbers: { start: 1 } } }, headers: { default: F.encabezado() }, footers: { default: F.pie("A-") }, children: anexo },
    ],
  };
}

module.exports = { documento };

if (require.main === module) (async () => {
  const d = documento("es");
  const salidas = [[0, path.join(__dirname, ARCHIVO)], [1, path.join(__dirname, "..", "anexo", ARCHIVO.replace(".docx", "_Anexo.docx"))]];
  for (const [k, salida] of salidas) {
    fs.writeFileSync(salida, await Packer.toBuffer(new Document({ ...d, sections: [d.sections[k]] })));
    console.log(salida);
  }
})();
