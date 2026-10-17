// Cuerpo del informe consolidado en Typst, con el diseño de los informes ganadores:
// el texto corre seguido en su columna y cada gráfico va al costado, a la altura del párrafo que lo cita.
// Los gráficos se ubican en el primer espacio libre de la columna derecha (sin encimarse ni pisar los bloques
// de ancho completo); los que no caben antes del Resumen de inversión van en una fila al final.
//
// Entradas (sys.inputs): datos = ruta del JSON de exportar.js · escalas = JSON {sección: escala del alto máximo}
//                        adelantar = secciones cuyos gráficos pueden subir hasta una página antes de su cita
// Compilar desde secciones/ (root): ver armar.py

#let D = json(sys.inputs.at("datos", default: "/informe_final/typst/datos/cuerpo_es.json"))
#let ESC = json(bytes(sys.inputs.at("escalas", default: "{}")))
#let ADEL = sys.inputs.at("adelantar", default: "").split(",")
#let ES = D.lang == "es"

// Paleta y medidas (twips → in, igual que el Word) --------------------------------------------------
#let AZUL = rgb("#14284B")
#let ROJO = rgb("#E9004B")
#let AMARILLO = rgb("#ECB100")
#let TINTA = rgb("#1A1A1A")
#let GRIS = rgb("#5F5F5F")
#let FONDO = rgb("#EEF1F6")
#let LINEA = rgb("#D5D9E0")
#let tw(n) = n / 1440 * 1in
#let ML = tw(620)
#let TXT = tw(6900)
#let SEP = tw(216)
#let FIG = tw(3550)
#let TOTAL = TXT + SEP + FIG
#let MT = tw(560)
#let MB = tw(760)
#let ALTO-MAX = 1.8in
#let HUECO = 4pt

#set document(title: if ES [Mallplaza — Informe CFA] else [Mallplaza — CFA Report])
#set text(font: ("Liberation Sans", "Arial"), size: 8pt, fill: TINTA, lang: D.lang, hyphenate: true)
#set par(justify: true, leading: 0.42em, spacing: 0.5em)

// Piezas de texto -----------------------------------------------------------------------------------
#let color-de(c) = if c == "rojo" { ROJO } else if c == "gris" { GRIS } else if c == "tinta" { TINTA } else { none }
#let corridas(rs) = for r in rs {
  let (t, b, i) = (r.at(0), r.at(1), r.at(2))
  let c = if r.len() > 3 { color-de(r.at(3)) } else if b { AZUL } else { TINTA }
  text(weight: if b { "bold" } else { "regular" }, style: if i { "italic" } else { "normal" }, fill: c, t)
}
#let titulo(t) = block(width: TOTAL, above: 10pt, below: 5pt, stroke: (bottom: 2.25pt + AZUL), inset: (bottom: 2pt),
  text(15pt, weight: "bold", fill: AZUL, t))
#let subtitulo(t) = block(width: 100%, above: 7pt, below: 3.5pt, stroke: (bottom: 1pt + ROJO), inset: (bottom: 1.5pt),
  sticky: true, text(weight: "bold", fill: AZUL, t))
#let lead(t, destacado: false) = block(width: TOTAL, fill: FONDO, inset: (x: 3pt, y: 3pt), below: 6pt,
  par(justify: true, text(fill: if destacado { AZUL } else { TINTA }, weight: if destacado { "bold" } else { "regular" }, t)))

// Bloques de ancho completo: marcan su alto para que los gráficos no los pisen
#let ancho(body) = [#metadata("ini") <obs-ini>#body#metadata("fin") <obs-fin>]

// Figura (título + gráfico + fuente) --------------------------------------------------------------
#let FIGS = D.secciones.enumerate().map(((si, s)) => s.figuras.map(f => (..f, clave: s.clave, sec: si))).flatten()
#let figura(f, ancho: FIG) = context {
  // el mapa de Sudamérica (Figura 1) es la figura principal de Negocio: se muestra más alto para que se lean sus fichas
  let maxh = if "N1_mapa" in f.archivo { 3.2in } else { ALTO-MAX * ESC.at(f.clave, default: 1.0) }
  let im = image(f.archivo, width: ancho)
  let h = measure(im).height
  block(width: ancho, breakable: false, {
    set par(justify: false, leading: 0.3em)
    text(7pt, fill: AZUL)[*#f.n* #f.titulo]
    v(2pt)
    align(center, if h > maxh { image(f.archivo, height: maxh) } else { im })
    v(1pt)
    set par(justify: true)
    text(5.5pt, fill: GRIS, corridas(f.fuente.map(r => (r.at(0), false, r.at(2), "gris"))))
  })
}

// Asignación de posiciones (se recalcula en cada página: es determinista) ---------------------------
#let ARRIBA = MT
#let ABAJO = 297mm - MB
#let antes(a, b) = a.at(0) < b.at(0) or (a.at(0) == b.at(0) and a.at(1) <= b.at(1))
#let asignar() = {
  let anclas = query(<ancla>)
  let inicios = query(<sec-ini>).map(l => l.location().position())
  let ini = query(<obs-ini>).map(l => l.location().position())
  let fin = query(<obs-fin>).map(l => l.location().position())
  let obst = (:)
  for (a, b) in ini.zip(fin) {
    for p in range(a.page, b.page + 1) {
      let y0 = if p == a.page { a.y } else { ARRIBA }
      let y1 = if p == b.page { b.y } else { ABAJO }
      obst.insert(str(p), obst.at(str(p), default: ()) + ((y0, y1),))
    }
  }
  let barr = query(<barrera>)
  let barrera = if barr.len() > 0 { let q = barr.first().location().position(); (q.page, q.y) } else { (9999, 0pt) }
  let ocup = (:)
  let asig = (:)
  let fuera = ()
  for a in anclas {
    let id = a.value
    let f = FIGS.at(id)
    let h = measure(figura(f)).height
    // Cada gráfico parte desde el inicio de su sección (no desde su cita): así los gráficos se apilan desde
    // arriba y ocupan el primer espacio libre de la columna derecha, en el orden en que se citan.
    let q = inicios.at(f.sec, default: a.location().position())
    let (p, y) = (q.page, q.y)
    if f.clave in ADEL and p > 1 { p = p - 1 }
    let ok = false
    for _ in range(300) {
      y = calc.max(y, ARRIBA)
      let choques = (ocup.at(str(p), default: ()) + obst.at(str(p), default: ())).filter(r =>
        r.at(0) < y + h + HUECO and y < r.at(1) + HUECO)
      if choques.len() > 0 { y = calc.max(..choques.map(r => r.at(1))) + HUECO; continue }
      if y + h > ABAJO { p += 1; y = ARRIBA; continue }
      if not antes((p, y + h), barrera) { break }
      ok = true
      break
    }
    if ok {
      asig.insert(str(id), (p, y))
      ocup.insert(str(p), ocup.at(str(p), default: ()) + ((y, y + h),))
    } else { fuera.push(id) }
  }
  (asig, fuera)
}

#let colocar() = context {
  let (asig, _) = asignar()
  let pag = here().page()
  for (id, (p, y)) in asig {
    if p == pag { place(top + left, dx: ML + TXT + SEP, dy: y, figura(FIGS.at(int(id)))) }
  }
}

#let gris-pie(t) = text(7pt, fill: GRIS, t)
#set page(paper: "a4", margin: (top: MT, bottom: MB, left: ML, right: ML + SEP + FIG),
  footer: block(width: TOTAL, align(center, gris-pie[#if ES [Página] else [Page] #context counter(page).display()])),
  foreground: colocar())

// Abstract ------------------------------------------------------------------------------------------
#let fmt(x, d: 0) = {
  let s = str(calc.round(calc.abs(x), digits: d))
  let (ent, dec) = if "." in s { s.split(".") } else { (s, "") }
  while dec.len() < d { dec += "0" }
  let grupos = ()
  while ent.len() > 3 { grupos.insert(0, ent.slice(ent.len() - 3)); ent = ent.slice(0, ent.len() - 3) }
  grupos.insert(0, ent)
  let (mil, coma) = if ES { (".", ",") } else { (",", ".") }
  (if x < 0 { "−" } else { "" }) + grupos.join(mil) + (if d > 0 { coma + dec } else { "" })
}
#let R = D.datos
#let A = D.abstract
// Portada con los colores de Mallplaza: logo sobre blanco, banda roja con la recomendación y franja azul con los datos
#let portada() = {
  let blanco(t, tam, ..a) = text(tam, fill: white, ..a, t)
  block(width: TOTAL, clip: true, radius: 2pt, below: 0pt, {
    grid(columns: (1.0in, 1fr), rows: (0.86in,),
      block(fill: white, width: 100%, height: 0.86in, inset: 4pt, stroke: (left: 1pt + ROJO, top: 1pt + ROJO, bottom: 1pt + ROJO),
        align(center + horizon, image("/comun/logos/mallplaza.png", height: 0.76in))),
      block(fill: ROJO, width: 100%, height: 0.86in, inset: (x: 9pt, y: 6pt), {
        set par(justify: false, leading: 0.3em)
        blanco("Mallplaza (Plaza S.A.)  ·  MALLPLAZ CI", 17pt, weight: "bold")
        v(4pt)
        grid(columns: (1fr, 1fr), row-gutter: 6pt,
          blanco(A.rotulos.at(0) + ": " + A.rec, 12pt, weight: "bold"),
          align(right, blanco(A.rotulos.at(1) + ": CLP " + fmt(R.precio_objetivo) + " (" + fmt(R.potencial * 100, d: 1) + "%)", 12pt, weight: "bold")),
          blanco(if ES [Sector: Centros comerciales (Chile, Perú, Colombia)] else [Sector: Shopping centers (Chile, Peru, Colombia)], 8.5pt, style: "italic"),
          align(right, blanco(A.rotulos.at(2) + ": CLP " + fmt(R.precio), 8.5pt, style: "italic")))
      }))
    block(width: 100%, fill: AZUL, inset: (x: 9pt, y: 3.5pt), {
      let datos = if ES {
        (("DCF", fmt(R.dcf)), ("Múltiplos", fmt(R.multiplos)), ("Ponderación", fmt(R.peso_dcf * 100) + "% / " + fmt((1 - R.peso_dcf) * 100) + "%"),
         ("Capitalización", "CLP " + fmt(R.capitalizacion_mm / 1000) + " MM"), ("Rango 52 sem.", fmt(R.rango_52s.at(0)) + " – " + fmt(R.rango_52s.at(1))))
      } else {
        (("DCF", fmt(R.dcf)), ("Multiples", fmt(R.multiplos)), ("Weighting", fmt(R.peso_dcf * 100) + "% / " + fmt((1 - R.peso_dcf) * 100) + "%"),
         ("Market cap", "CLP " + fmt(R.capitalizacion_mm / 1000) + " bn"), ("52-wk range", fmt(R.rango_52s.at(0)) + " – " + fmt(R.rango_52s.at(1))))
      }
      grid(columns: (1fr,) * 5, column-gutter: 4pt, ..datos.map(((a, b)) => align(center, text(5.8pt, fill: AMARILLO, weight: "bold", a) + linebreak() + text(7.6pt, fill: white, weight: "bold", b))))
    })
  })
  v(4pt)
}
#ancho({
  portada()
  titulo(A.titulo)
  block(width: TOTAL, fill: FONDO, inset: 4pt, below: 4pt, par(justify: true, text(8.5pt, A.texto)))
  let valores = (A.rec, "CLP " + fmt(R.precio_objetivo), "CLP " + fmt(R.precio), fmt(R.potencial * 100, d: 1) + "%")
  grid(columns: (1fr,) * 4, column-gutter: 1.5pt, ..range(4).map(i => block(width: 100%, fill: if i == 0 { ROJO } else { AZUL },
    inset: (y: 4pt), align(center, text(6.5pt, fill: white, A.rotulos.at(i)) + linebreak() + text(13pt, weight: "bold", fill: white, valores.at(i))))))
  v(6pt)
})

// Secciones -----------------------------------------------------------------------------------------
#let k0 = 0
#for (si, s) in D.secciones.enumerate() {
  [#metadata(si) <sec-ini>]
  let base = D.secciones.slice(0, si).map(x => x.figuras.len()).sum(default: 0)
  // Análisis Financiero empieza arriba de una página nueva (pedido del equipo)
  if s.clave == "fa" { pagebreak(weak: true) }
  ancho({
    titulo(s.titulo)
    if "lead" in s { lead(s.lead, destacado: s.at("leadDestacado", default: false)) }
    if "tabla" in s {
      let T = s.tabla
      text(7pt, fill: AZUL)[*#T.n* #T.titulo]
      v(-2pt)
      let n = T.encabezado.len()
      // la tabla ocupa todo el ancho de la hoja (hasta el margen derecho de la columna de figuras)
      pad(right: -(SEP + FIG), table(columns: (2.9fr,) + (1fr,) * (n - 1), stroke: (x, y) => (bottom: if y > 0 { 0.5pt + LINEA } else { none }),
        inset: (x: 3pt, y: 1.6pt), align: (x, y) => if x == 0 { left } else { right },
        fill: (x, y) => if y == 0 { AZUL } else if (y - 1) in T.destacar { rgb("#FDF1CC") } else if calc.odd(y) { rgb("#F7F8FB") } else { white },
        ..T.encabezado.map(c => text(6pt, weight: "bold", fill: white, c)),
        ..T.filas.enumerate().map(((r, fila)) => fila.map(c => text(6pt, weight: if r in T.destacar { "bold" } else { "regular" }, c))).flatten()))
      v(-3pt)
      text(5.5pt, fill: GRIS, T.nota)
      v(4pt)
    }
  })
  for (bi, b) in s.bloques.enumerate() {
    let anclas = s.figuras.enumerate().filter(((j, f)) => f.ancla == bi).map(((j, f)) => [#metadata(base + j) <ancla>])
    if "h1" in b { anclas.join(); subtitulo(b.h1) }
    else { par[#anclas.join()#corridas(b.p)] }
  }
}

// Gráficos que no cupieron al costado + Resumen de inversión -----------------------------------------
#context {
  let (asig, fuera) = asignar()
  [#metadata((fuera: fuera.len(), ids: fuera, asig: asig, claves: FIGS.map(f => f.clave))) <resultado>]
  if fuera.len() > 0 {
    ancho(grid(columns: (1fr, 1fr), column-gutter: SEP, row-gutter: 8pt, ..fuera.map(id => figura(FIGS.at(id), ancho: (TOTAL - SEP) / 2))))
  }
}
#[#metadata("barrera") <barrera>]
#let X = D.resumen
#ancho(block(breakable: false, width: TOTAL, {
  titulo(X.titulo)
  let valores = (X.rec, fmt(R.precio_objetivo), fmt(R.precio), fmt(R.potencial * 100, d: 1) + "%",
    fmt(R.dcf) + " / " + fmt(R.multiplos), fmt(R.peso_dcf * 100) + "% / " + fmt((1 - R.peso_dcf) * 100) + "%", "MALLPLAZ CI",
    fmt(R.capitalizacion_mm / 1000), fmt(R.rango_52s.at(0)) + " – " + fmt(R.rango_52s.at(1)),
    fmt(R.dividendo) + " (" + fmt(R.dividend_yield * 100, d: 1) + "%)", fmt(R.dfn_mm / R.ebitda_udm_mm, d: 2) + "x")
  grid(columns: (TXT, FIG), column-gutter: SEP,
    {
      subtitulo(X.pend)
      text(style: "italic", fill: GRIS, X.t1); parbreak(); text(style: "italic", fill: GRIS, X.t2)
    },
    {
      subtitulo(X.datos)
      table(columns: (1.15fr, 1fr), stroke: (bottom: 0.5pt + LINEA), inset: (x: 3pt, y: 1.8pt), align: (x, y) => if x == 0 { left } else { right },
        fill: (x, y) => if calc.even(y) { rgb("#F7F8FB") } else { white },
        ..X.filas.zip(valores).map(((a, b)) => (text(6.5pt, weight: "bold", fill: AZUL, a), text(6.5pt, b))).flatten())
      text(5.5pt, fill: GRIS, X.fuente)
    })
}))
