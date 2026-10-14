// Exporta el cuerpo del informe consolidado a JSON para la maquetación en Typst (informe.typ).
// Mismo contenido que el Word (informe_final.js): textos de cada sección con las figuras y anexos ya renumerados.
// Uso (desde secciones/): node informe_final/typst/exportar.js <es|en>  → informe_final/typst/datos/cuerpo_<lang>.json
const fs = require("fs");
const path = require("path");
const I = require("../documento/informe_final.js");

const lang = process.argv[2] || "es";
const R = (clave, t) => I.con(clave, () => I.renumerar(t));
const runsR = (clave, runs) => runs.map(([t, b, i]) => [R(clave, t), !!b, !!i]);
const texto = (runs) => runs.map((r) => r[0]).join("");

// Cada figura se ancla al primer párrafo de su sección que la cita (o al primero de la sección)
function anclar(bloques, figuras) {
  return figuras.map((fig) => {
    const num = fig.n.replace(/^(Figura|Figure)\s+/, "").replace(/:$/, "");
    const rx = new RegExp(`(Figuras?|Figures?)\\s+(\\S+[–-])?${num.replace(/[.]/g, "\\.")}(?![\\w.])`, "i");
    let i = bloques.findIndex((b) => b.p && rx.test(texto(b.p)));
    if (i < 0) i = bloques.findIndex((b) => b.p);
    return { ...fig, ancla: i };
  });
}

function borrador(clave, json, imagenes) {
  const sec = { clave, titulo: null, bloques: [], figuras: [] };
  for (const b of json.bloques) {
    if (b.tipo === "titulo") { if (/^(ANEXO|APPENDI)/.test(b.texto)) break; sec.titulo = R(clave, b.texto); continue; }
    if (b.tipo !== "pagina") continue;
    for (const f of b.filas) {
      for (const x of f.izq) {
        if (x.tipo === "h1") sec.bloques.push({ h1: R(clave, x.texto) });
        else if (x.tipo === "p") sec.bloques.push({ p: runsR(clave, x.runs) });
      }
      let g = null;
      for (const x of f.der) {
        if (x.tipo === "caption") { g = { n: R(clave, x.n), titulo: R(clave, x.texto), fuente: [] }; sec.figuras.push(g); }
        else if (x.tipo === "imagen" && g) { const v = imagenes[x.archivo]; g.archivo = Array.isArray(v) ? v[0] : v; }
        else if (g && (x.tipo === "fuente" || x.tipo === "p")) g.fuente = runsR(clave, x.runs);
      }
    }
  }
  sec.figuras = anclar(sec.bloques, sec.figuras.filter((x) => x.archivo));
  return sec;
}

function modulo(clave, X, carpeta) {
  const sec = { clave, titulo: R(clave, X.titulo), lead: R(clave, X.lead), bloques: [], figuras: [] };
  for (const p of X.paginas) {
    for (const b of p.texto) {
      if (b.h1) sec.bloques.push({ h1: R(clave, b.h1) });
      else if (b.p) sec.bloques.push({ p: Array.isArray(b.p) ? [[R(clave, b.p[0]), true, false], [R(clave, b.p[1]), false, false]] : [[R(clave, b.p), false, false]] });
    }
    for (const [n, t, archivo, f] of p.figuras) {
      if (archivo) sec.figuras.push({ n: R(clave, n), titulo: R(clave, t), archivo: path.join(carpeta, `${lang}_${archivo}.png`), fuente: [[R(clave, f), false, false]] });
    }
  }
  sec.figuras = anclar(sec.bloques, sec.figuras);
  return sec;
}

const datos = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "graficos", "datos_resumen.json"), "utf8"));
const fa = I.fa, X = fa.L[lang], f = fa.loc(lang);
const [n1, t1, filas1, nota1] = X.fa1;
const secFA = modulo("fa", X, fa.FIG);
secFA.leadDestacado = true;
secFA.tabla = { n: R("fa", n1), titulo: R("fa", t1), encabezado: [filas1[0], ...fa.FA1_AÑOS],
  filas: fa.FA1.map((r, i) => [filas1[i + 1], ...r.map(f)]), destacar: [9], nota: R("fa", nota1) };

// Riesgos: texto corrido con sus 4 figuras al costado (runs con color: [texto, negrita, cursiva, color])
function secRiesgos() {
  const L = Object.values(I.riesgos.LANG).find((x) => x.archivo.endsWith(`_${lang.toUpperCase()}.docx`));
  const dir = path.join(__dirname, "..", "..", "riesgos", "graficos", "figuras");
  const sec = { clave: "riesgos", titulo: R("riesgos", L.titulo), lead: R("riesgos", L.intro), bloques: [], figuras: [] };
  for (const s of L.secciones) {
    sec.bloques.push({ h1: R("riesgos", s.titulo) });
    for (const p of s.parrafos || []) sec.bloques.push({ p: [[R("riesgos", p), false, false]] });
    for (const r of s.riesgos || []) {
      sec.bloques.push({ p: [[`${r.cod}: `, true, false, "rojo"], [r.nombre + " ", true, false, "tinta"],
        [R("riesgos", r.texto) + " ", false, false], [`(${r.tag})`, false, true, "gris"]], pegado: true });
      sec.bloques.push({ p: [[L.mitigantes, true, false], [R("riesgos", r.mit), false, false]] });
    }
  }
  sec.figuras = anclar(sec.bloques, L.figuras.map((x) => ({ n: R("riesgos", x.n), titulo: R("riesgos", x.t),
    archivo: path.join(dir, x.img), fuente: [[R("riesgos", x.fuente), false, false]] })));
  return sec;
}

const salida = {
  lang,
  abstract: I.ABS[lang], resumen: I.RES[lang], datos,
  secciones: [
    I.con("negocio", () => borrador("negocio", I.negocio.textoDe(lang), I.negocio.imagenesDe(lang))),
    I.con("industria", () => borrador("industria", I.industria.textoDe(lang), I.industria.imagenesDe(lang))),
    I.con("esg", () => borrador("esg", I.esg.textoDe(lang), I.esg.imagenesDe(lang))),
    secFA,
    modulo("valuacion", I.valuacion.L[lang], I.valuacion.FIG),
    I.con("riesgos", secRiesgos),
  ],
};
const archivo = path.join(__dirname, "datos", `cuerpo_${lang}.json`);
fs.writeFileSync(archivo, JSON.stringify(salida, null, 1));
console.log(archivo, salida.secciones.map((s) => `${s.clave}: ${s.bloques.length} bloques, ${s.figuras.length} figuras`).join(" · "));
