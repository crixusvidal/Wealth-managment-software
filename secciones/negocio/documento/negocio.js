// Sección Descripción del Negocio + anexos, en el formato principal del informe.
// Texto: borrador del equipo (inglés: borrador/texto_en.json, traducción fiel del mismo) (borrador/Business_Description_BORRADOR_FINAL.docx → borrador/texto.json con
// comun/borrador_a_json.py), sin cambios. Gráficos: los del repo (graficos/negocio.py), en vez de las capturas.
// Uso: node negocio/documento/negocio.js
const fs = require("fs");
const path = require("path");
const { Document, Packer } = require("docx");
const F = require("../../comun/formato");

const FIG = path.join(__dirname, "..", "graficos", "figuras");
const ARCHIVO = "Mallplaza_Negocio_ES.docx";
const IMAGENES = { "image1.png": "N1_mapa", "image2.png": "NA_activos_tier_a", "image3.png": "NB_tiendas_departamento" };

const imagenesDe = (lang) => Object.fromEntries(Object.entries(IMAGENES).map(([k, n]) => [k, path.join(FIG, `${lang}_${n}.png`)]));
const textoDe = (lang) => JSON.parse(fs.readFileSync(path.join(__dirname, "borrador", lang === "es" ? "texto.json" : "texto_en.json"), "utf8"));

function documento(lang = "es") {
  const json = JSON.parse(fs.readFileSync(path.join(__dirname, "borrador", lang === "es" ? "texto.json" : "texto_en.json"), "utf8"));
  const imagenes = imagenesDe(lang);
  const { cuerpo, anexo } = F.seccionBorrador(json, { imagenes, anchoAnexo: F.COL_FIG });
  return {
    styles: F.ESTILOS,
    sections: [
      { properties: { page: F.PAGINA }, headers: { default: F.encabezado() }, footers: { default: F.pie(lang === "es" ? "Página " : "Page ") }, children: cuerpo },
      { properties: { page: { ...F.PAGINA, pageNumbers: { start: 1 } } }, headers: { default: F.encabezado() }, footers: { default: F.pie("A-") }, children: anexo },
    ],
  };
}

module.exports = { documento, imagenesDe, textoDe };

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
