// Cifras del informe: lee informe_final/graficos/cifras.json (python informe_final/graficos/cifras_informe.py) y da funciones de
// formato por idioma. Los textos del informe usan estas funciones en vez de cifras escritas a mano: así el texto, las figuras
// y los anexos salen de la misma fuente (el modelo de valoración v4).
//   const { K, f } = require("../../comun/cifras")("es");
//   f.n(K.po) → "2.077" · f.pc(K.po_vs, 1) → "−41,6%" · f.pb(K.wacc_impl_pb) → "281" · f.m(K.ev_dcf_mm) → "5.407.032" (CLP MM)
const fs = require("fs");
const path = require("path");

const K = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "informe_final", "graficos", "cifras.json"), "utf8"));

function formatos(lang) {
  const es = lang === "es";
  const sep = (s) => (es ? s.replace(/[.,]/g, (c) => (c === "." ? "," : ".")) : s);
  const n = (x, d = 0) => sep(Math.abs(x).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }));
  const sg = (x, d = 0) => (x < 0 && Math.round(Math.abs(x) * 10 ** d) > 0 ? "−" : "") + n(x, d);
  const f = {
    n: sg,                                              // número con signo (menos tipográfico)
    pc: (x, d = 1) => sg(x * 100, d) + "%",             // 0,1234 → 12,3%
    pcs: (x, d = 1) => (x >= 0 ? "+" : "−") + n(x * 100, d) + "%",     // con signo explícito
    pb: (x) => n(x, 0),                                 // puntos base, valor absoluto
    m: (x, d = 0) => sg(x, d),                          // CLP MM
    mm: (x, d = 1) => sg(x / 1000, d),                  // CLP mil MM (miles de millones)
    bn: (x, d = 2) => sg(x / 1e6, d),                   // billones (10^12) con el valor en CLP MM
    x: (v, d = 1) => sg(v, d) + "x",
  };
  return f;
}

module.exports = (lang) => ({ K, f: formatos(lang), formatos });
