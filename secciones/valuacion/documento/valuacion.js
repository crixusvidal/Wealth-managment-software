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
const MARCAR_NUEVO = true;
const AMARILLO = "ECB100";
const NUEVO = { es: [], en: [] };                         // fragmentos marcados (se registran con N)
const N = (lang, texto) => { NUEVO[lang].push(texto); return texto; };
const TIR = JSON.parse(fs.readFileSync(path.join(FIG, "datos_tir.json"), "utf8"));   // graficos/tir_valuacion.py (modelo v3, hoja TIR_Accionista)
const ALTO = { V7: 2.25, V9: 1.95, V1: 1.85, V2: 1.7, V3: 2.05, V4: 1.8, V5: 1.6, V6: 1.75, W1: 1.9 }; // = graficos/valuacion.py

const fmt = (lang) => (x, dec = 1) => {
  const s = x.toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec }).replace("-", "−");
  return lang === "es" ? s.replace(/[.,]/g, (c) => (c === "." ? "," : ".")) : s;
};

// Formato de las cifras de la TIR por idioma: pc = porcentaje, pb = puntos base (valor absoluto)
const TT = Object.fromEntries(["es", "en"].map((lang) => { const f = fmt(lang);
  return [lang, { pc: (x, d = 2) => f(x * 100, d) + "%", pb: (x) => f(Math.abs(x), 0) }]; }));

// Tabla chica de la TIR: nominal CLP vs real (UF) y brecha frente al costo de capital (WACC para la firma, Ke para el accionista)
function tablaTIR(lang) {
  const t = TT[lang], T = TIR, es = lang === "es";
  const fila = (nombre, k, costo, d) => [nombre, t.pc(T[k].nom, d), t.pc(T[k].real, d),
    `${costo} ${t.pc(T[k].costo_nom)} / ${t.pc(T[k].costo_real)}`, `${T[k].brecha_nom < 0 ? "−" : "+"}${t.pb(T[k].brecha_nom)} / ${T[k].brecha_real < 0 ? "−" : "+"}${t.pb(T[k].brecha_real)} ${es ? "pb" : "bp"}`];
  return { nuevo: true, anchos: [2300, 740, 640, 1800, 1420], derecha: [1, 2, 3, 4], destacar: [0],
    titulo: es ? "TIR al precio de mercado frente al costo de capital: nominal CLP y real (UF)" : "IRR at the market price vs cost of capital: nominal CLP and real (UF)",
    cols: es ? ["Quien compra a CLP 3.555", "Nominal", "Real", "Costo cap. nom./real", "Brecha nom./real"]
             : ["Buying at CLP 3,555", "Nominal", "Real", "Cost of cap. nom./real", "Gap nom./real"],
    filas: [fila(es ? "Firma (FCFF), tasa única" : "Firm (FCFF), single rate", "firma", "WACC", 2),
      fila(es ? "Accionista, sale a nuestro valor" : "Shareholder, exits at our value", "acc_mercado", "Ke", 1),
      fila(es ? "Accionista, compra al objetivo" : "Shareholder, buys at target", "acc_objetivo", "Ke", 1)],
    nota: es ? `TIR: tasa que iguala el valor de los flujos al precio. Firma: tasa única sobre el FCFF y el valor terminal (la Figura 34 desplaza toda la curva de WACC y da ${t.pc(TIR.v6.wacc)}; ambas dan CLP 3.555). Accionista: distribuciones del modelo (con la adicional que lleva la DFN/EBITDA a 3,5x) y salida al 31-dic-31 a nuestro valor. Real = (1 + nom.) / (1 + ${t.pc(T.pi, 1)}) − 1. Fuente: modelo v3, hoja TIR_Accionista.`
             : `IRR: the rate that equates the value of the cash flows to the price. Firm: a single rate on FCFF and terminal value (Figure 34 shifts the whole WACC curve and gives ${t.pc(TIR.v6.wacc)}; both give CLP 3,555). Shareholder: the model's distributions (with the extra one that takes net debt/EBITDA to 3.5x) and exit at 31-Dec-31 at our value. Real = (1 + nom.) / (1 + ${t.pc(T.pi, 1)}) − 1. Source: model v3, TIR_Accionista sheet.` };
}

const L = {
  es: {
    archivo: "Mallplaza_Valuacion_ES.docx", pag: "Página ", pagA: "A-", titulo: "VALUACIÓN",
    lead: "Recomendamos VENDER Mallplaza con un precio objetivo de CLP 2.076 por acción, 41,6% bajo el cierre de CLP 3.555 del 30-sep-26 (−39,7% de retorno total con dividendos). Nuestro precio objetivo queda incluso bajo el mínimo de las últimas 52 semanas (CLP 2.470).",
    paginas: [
      { texto: [
          { h1: "UN PRECIO QUE NINGÚN MÉTODO ALCANZA" },
          { p: `En doce meses la acción subió de CLP 2.470 a 4.700 y hoy cotiza en CLP 3.555. Ninguno de nuestros métodos llega a ese precio: el DCF va de CLP 1.379 a 2.329 aun moviendo la WACC ±100 pb y g entre 2,5% y 4,0%${N("es", " (nominales)")}, y los múltiplos más favorables tampoco lo alcanzan (Figura 27). Damos 80% del peso al DCF porque Mallplaza es, en el fondo, un cobrador de arriendos: el 81,4% de sus ingresos son arriendos y el 93% de ellos es renta fija. El 20% en múltiplos contrasta el modelo con lo que el mercado paga hoy, pero con cautela: en Chile hay solo dos pares directos y ninguno captura el plan de expansión ni la trayectoria de deuda de la compañía.` },
          { p: "La Figura 27 resume la tesis en una imagen. El APV, que valora la empresa sin deuda y suma aparte los escudos fiscales, da CLP 1.894, apenas 2% sobre el DCF: el resultado no depende de cómo modelamos el financiamiento. Los múltiplos van de CLP 2.152 (mediana del set común) a 3.294 (mediana chilena), y el EV/EBITDA NTM y el EV/EBIT quedan entre CLP 2.914 y 3.098. Todo cae a la izquierda del precio; lo único que lo alcanza es el propio rango de cotización del último año. Incluso la mediana chilena de múltiplos, el escenario más favorable para la acción, queda 7% bajo el precio." },
          { h1: "LA HISTORIA ESTÁ EN LA TASA" },
          { p: "Proyectamos el flujo de caja libre de la firma en CLP nominales desde el 30-jun-26, con horizonte explícito hasta 2031E y un valor terminal que explica el 76% del EV (Anexo [B]). Los CLP 1.857 por acción son CLP 583 de flujos explícitos y CLP 1.884 de valor terminal, menos CLP 610 de deuda y minoritarios: el resultado depende, sobre todo, de la tasa de descuento." },
          { p: `Esa tasa se mueve con la estructura de capital. Mallplaza usa su capacidad de deuda para crecer, como en 2024 tras comprar Open Plaza Perú (3,78x): suponemos que la DFN/EBITDA sube de 2,36x a 3,5x en 2031, con 25 pb más de costo real por cada vuelta, y que la deuda nueva (CLP 1,40 billones) financia todo el capex mientras la caja operativa se distribuye a los accionistas. Así, la WACC nominal pasa de 10,54% en 2S26E a 10,41% en el terminal (7,20% real): un ${N("es", "Ke nominal de 11,93% (8,67% real; Rf 4,98%")}, beta re-apalancada de 1,14 desde una mediana de seis pares de 0,94, ERP 4,14% y CRP de 1,41% ponderado por ingresos) y un ${N("es", "Kd nominal de 5,11% después de impuestos")}, basado en el ${N("es", "cupón real de 3,6% (UF)")} de la última emisión (Anexo [Z]).` },
          { p: `Construimos el Ke en dólares porque sus tres insumos están en esa moneda ${N("es", "y son nominales")}: la Rf del Tesoro de EE.UU., la prima por riesgo de mercado implícita que publica Damodaran y el riesgo país, que ponderamos por ingresos (71,8% Chile, 17,7% Perú y 10,5% Colombia). Luego lo llevamos a pesos con el diferencial entre la inflación esperada de Chile (3,0%, meta del BCCh) y la de EE.UU. (2,24%, breakeven a 10 años). ${N("es", "La beta se desapalanca con Hamada (seis pares regionales) y se re-apalanca cada período (1,09 en 2S26E, 1,14 en el terminal).")}` },
          { p: `${N("es", `La primera señal es la TIR, lo que gana quien compra a CLP 3.555 si se cumplen nuestros flujos: para la firma, ${TT.es.pc(TIR.firma.nom)} nominal (${TT.es.pc(TIR.firma.real)} real en UF), ${TT.es.pb(TIR.firma.brecha_nom)} pb bajo su WACC y apenas ~80 pb reales sobre el costo real de su deuda (Figura 28); para el accionista, ${TT.es.pc(TIR.acc_mercado.nom, 1)} nominal frente a un Ke de ${TT.es.pc(TIR.ke)}.`)} El mercado está pagando por la acción casi como si fuera un bono.` },
          { t: tablaTIR("es") },
          { h1: "CRECER NO RESCATA EL PRECIO" },
          { p: "No es que seamos pesimistas con la operación. Proyectamos un margen EBITDA de 79,7% en 2026E (sin revalorizaciones, bajo el 80,5% UDM), un impuesto que converge de 16,8% a 27%, el pipeline anunciado (USD 613 MM y 170.200 m²) y, desde 2029, 27.075 m² nuevos al año, al ritmo histórico. El consenso de 13 analistas espera un EBITDA a 12 meses apenas ~3% sobre el nuestro: la diferencia con el mercado no está en los flujos." },
          { p: `El capex sigue el plan de la compañía: el pipeline y la mantención (5,5% de los ingresos) llevan el capex base a un máximo de CLP 240.660 MM en 2027, y la expansión continua suma entre CLP 95.000 y 104.000 MM al año desde 2028, al costo del pipeline identificado (USD 3.603 por m²) y con una rampa de ingresos de 35%, 90% y 100%. El capital de trabajo, en 15,6% de los ingresos, exige CLP 7.100–9.900 MM al año y mueve el precio en ±1,2% (Figura 29${N("es", "; Anexo [C]")}).` },
          { p: `Dos supuestos merecen una explicación. La tasa de impuesto en caja parte en 16,8% porque hoy la compañía usa goodwill tributario y pérdidas acumuladas; suponemos que esos escudos se agotan y que la tasa llega a 27% en 2030, para no sobreestimar el valor. ${N("es", "Y el capital de trabajo crece con los ingresos (cuentas por cobrar de CLP 122.370 MM sobre arriendos anticipados de CLP 20.247 MM), no con los m² nuevos.")}` },
          { p: `El problema es el retorno de crecer. Cada m² nuevo rinde ${N("es", "6,9% real antes de impuestos")} sobre su costo, un RONIC real de 5,0% después de impuestos, bajo una WACC real de 7,2%. La expansión suma ingresos, pero no valor: por eso el valor terminal crece solo con la inflación (g de 3,0% nominal, 0% real) y un g mayor reduce el valor (Figura 30). Suponer más crecimiento no cerraría la brecha con el precio; la agrandaría.` },
          { p: `Por eso el horizonte explícito incluye el capex y la maduración de los proyectos que abren hasta 2032 (108.301 m² de GLA madura en el año terminal), pero desde ahí el modelo no agrega centros comerciales: el capex terminal cubre solo la mantención (5,5% de los ingresos). ${N("es", "Con estos retornos, crecer más reduciría su valor.")}` },
        ],
        figuras: [
          ["Figura 27:", "Ningún método justifica el precio actual (CLP por acción)", "V1_football", "Rango de 52 semanas intradía (01-10-25 a 30-09-26; mínimo 10-10-25, máximo 27-04-26). Fuente: análisis del equipo; Bloomberg."],
          ["Figura 28:", `${N("es", "A CLP 3.555 la firma rinde casi lo mismo que su deuda")}`, "V2_tasas", "Tasas reales en UF. Fuente: BCCh (30-06-26); Mallplaza; análisis del equipo."],
          ["Figura 29:", "Capital de trabajo: un supuesto que no mueve el precio (±1,2%)", "V4_capital_trabajo", "Fuente: EEFF Mallplaza; análisis del equipo."],
          ["Figura 30:", "Crecer bajo la WACC destruye valor: g real = 0%", "V3_value_driver", "Arriba: valor DCF según g. Abajo: valor terminal relativo según el RONIC. Fuente: análisis del equipo."]] },
      { texto: [
          { h1: "LOS PARES CONFIRMAN LA DIRECCIÓN, NO LA MAGNITUD" },
          { p: `Mallplaza transa a 16,8x EV/EBITDA, sobre la mediana chilena (15,7x) y muy por encima de los pares brasileños (9,0x). La brecha con Brasil no es de calidad, sino de tasas: con bonos reales locales de 2,5% en Chile y 8,1% en Brasil (NTN-B), el spread del cap rate${N("es", " (comparable a una tasa real, porque las rentas se reajustan)")} sobre la tasa local es parecido en ambos países, 385 pb en Chile y 302 pb en Brasil (Figura 31). Por eso ponderamos 70% la mediana chilena y 30% la del set común (que incluye IRSA; excluimos Fibra Danhos por su régimen FIBRA): 14,4x, que lleva a CLP 2.951 por acción. La ponderación pesa poco: entre 0% y 100% Chile, el precio objetivo solo va de CLP 1.916 a 2.144 (Anexo [E]).` },
          { p: `El contexto macro explica la diferencia. Al 30-jun-26 Brasil tenía ${N("es", "una Selic nominal de 14,25%")}, un bono nominal a 10 años de 14,34% e inflación esperada de 5,3% para 2026, frente a ${N("es", "una TPM nominal de 4,5%")}, un bono nominal de 5,43% y 4,1% de inflación en Chile, con ${N("es", "crecimientos reales del PIB")} parecidos (2,0% y 1,6%). Un inversionista en Brasil exige más a cualquier activo, no solo a los centros comerciales; aplicar ese descuento a Mallplaza, que genera el 71,8% de sus ingresos en Chile, sería castigarla por un riesgo que no corre.` },
          { p: "Elegimos el set con el mismo criterio que la beta. IRSA se mantiene con cautela, porque reporta bajo NIC 29 y su consenso NTM tiene un solo analista; Fibra Danhos queda fuera porque, como FIBRA, debe repartir el 95% de su resultado fiscal y no tiene beta en la WACC. Usamos EV/EBITDA porque es neutral al apalancamiento y a la revalorización (NIC 40), y medimos a todos con la deuda neta de Bloomberg (CLP 1.311.569 MM para Mallplaza, cerca de los CLP 1.285.740 MM de DFN que usa el DCF)." },
          { p: "Tampoco la calidad justifica la prima. Mallplaza transa 29% sobre Cencosud Shopping, pero su EBITDA por m² es solo 1,8% mayor (ver Industria); solo al múltiplo de Parque Arauco (18,4x), que opera con 5,7x de deuda, valdría más que su precio. Todos los múltiplos, del EV/EBITDA NTM (CLP 2.914) al P/BV (CLP 3.171), quedan bajo el precio (Figura 32): confirman que la acción está cara, y la magnitud la da el DCF. Excluimos el P/E (CLP 8.644), distorsionado por la revalorización de propiedades." },
          { p: "Puesto en precios: a la mediana chilena por sí sola la acción valdría CLP 3.294 (−7%); al múltiplo de Cencosud, CLP 2.621; y solo al de Parque Arauco, CLP 3.967. Para defender CLP 3.555 con múltiplos hay que suponer que Mallplaza merece la valoración del par más apalancado del grupo. Con flujos esperados ocurre lo mismo: transa a 14,8x el EBITDA NTM de consenso, frente a 12,5x del múltiplo ponderado de pares." },
          { h1: "LO QUE EL MERCADO TENDRÍA QUE CREER" },
          { p: `El precio implica un EV de CLP 9,12 billones, 69% sobre el nuestro, y un cap rate de 6,0%, frente a 10,1% en nuestro DCF, 6,4% en los pares chilenos y 7,5% a valor libro (Figura 33). El mercado paga 1,61 veces un patrimonio que ya está a valor razonable, tasado con una tasa real de ~5,55% (Nota 11): paga más que los propios tasadores. Para justificar CLP 3.555 la ${N("es", "WACC nominal tendría que ser 285 pb menor (7,56%; ")}Figura 34); incluso con la tasa de los tasadores${N("es", " (5,55% real, 8,72% nominal)")}, el DCF sería CLP 2.660.` },
          { p: `Llevamos la tesis a sus extremos. En un caso bear (${N("es", "WACC nominal +100 pb")} y el margen trimestral mínimo de 75,8%) el precio objetivo cae a CLP 1.713; en un caso bull que suma ${N("es", "WACC nominal −100 pb, g nominal de 2,5%")}, el margen trimestral máximo de 81,3% y la rebaja del impuesto a 23%, sube a CLP 2.627, todavía 26% bajo el precio. El APV (CLP 1.894), una Rf basada en el bono en pesos (~CLP 2.439) y otros objetivos de deuda (DCF de CLP 1.819 a 1.894) llevan a la misma conclusión (Anexo [F]). ${N("es", `La TIR de la firma al precio de mercado casi no cambia entre escenarios (${TT.es.pc(TIR.esc.bear.nom, 1)}, ${TT.es.pc(TIR.esc.base.nom, 1)} y ${TT.es.pc(TIR.esc.bull.nom, 1)} nominal en bear, base y bull; ${TT.es.pc(TIR.esc.bear.real, 1)}, ${TT.es.pc(TIR.esc.base.real, 1)} y ${TT.es.pc(TIR.esc.bull.real, 1)} real) y queda siempre bajo la WACC de cada uno (${TT.es.pb(TIR.esc.bear.brecha_nom)}, ${TT.es.pb(TIR.esc.base.brecha_nom)} y ${TT.es.pb(TIR.esc.bull.brecha_nom)} pb nominales): ni el bull deja margen de seguridad.`)}` },
          { p: `El APV también dice cuánto vale de verdad la deuda. De sus CLP 1.894 por acción, CLP 1.632 corresponden a la empresa sin deuda y CLP 261 a los escudos fiscales de la deuda proyectada: la estrategia de apalancarse hasta 3,5x agrega valor, pero ni sumándola se acerca al precio. Y la elección de la tasa libre de riesgo es la conservadora: usar ${N("es", "el bono nominal en pesos (5,43%)")} y además sumar el riesgo país contaría dos veces el riesgo Chile, y aun así el precio objetivo solo subiría a ~CLP 2.439, 31% bajo el precio.` },
          { p: "Tampoco cambian el resultado los detalles de calendario. Con un dividendo esperado de ~CLP 69 por acción (2,0%), el retorno total esperado es −39,7%. Y como el DCF está al 30-jun-26 y el precio al 30-sep-26, la comparación es levemente favorable al mercado: si el saldo del dividendo 2026 (~CLP 27 por acción) ya se pagó antes de esa fecha, nuestro valor estaría algo sobrestimado, no subestimado." },
          { p: `Así llegamos a CLP 2.076 (80% × 1.857 + 20% × 2.951): 10,8x EBITDA, un cap rate de 9,3% y 0,94x el valor libro. Llevar el DCF al 30-sep-26 lo subiría a ~CLP 2.118, sin cambiar la recomendación. El catalizador es el mismo que sostiene la tesis: con el ${N("es", "UST 10Y nominal")} en 5,22%, máximo desde 2002, las tasas presionan los cap rates, y cada nueva tasación los acerca a nuestro valor. Nos equivocaríamos si los cap rates se comprimen aún más por la demanda institucional por activos defensivos, o si el mercado sostiene una prima estructural por los activos Tier A (ver Riesgos de inversión).` },
          { p: `Lo que vigilaremos: la tasación de la Nota 11 en cada cierre (la tasa real de descuento es el puente entre nuestro valor y el libro), el spread del bono UF de la compañía frente a ${N("es", "su cupón real de 3,6%")} (con la ${N("es", "tasa efectiva real de 4,33%")} de la última emisión, la ${N("es", "WACC nominal subiría 12 pb")} y el precio objetivo bajaría a CLP 2.042), la rebaja del impuesto corporativo a 23% (+5,1% al precio objetivo, aún no incluida) y la clasificación AAA, que probablemente quedaría bajo presión con 3,5x de DFN/EBITDA.` },
        ],
        figuras: [
          ["Figura 31:", "Chile y Brasil cobran la misma prima: los múltiplos difieren por las tasas", "V7_multiplos", "Bono real: UF 10 años (BCCh) y NTN-B 2035 (8,08%) al 30-06-26. Spread = cap rate (1 / EV/EBITDA) − bono real. IRSA se omite (NIC 29). Fuente: Bloomberg; BCCh; análisis del equipo."],
          ["Figura 32:", "Valor implícito por múltiplo de cada par vs DCF y precio (CLP por acción)", "V9_pares", "Precio objetivo = 80% DCF + 20% EV/EBITDA LTM. P/E excluido (incluye revalorización). Fuente: Bloomberg (30-09-26); análisis del equipo."],
          ["Figura 33:", "El mercado paga 6,0%; nuestros flujos valen 10,1%", "V5_cap_rate", "Gran Plaza: hecho posterior, no entra al modelo. Fuente: Mallplaza; Bloomberg; DF; análisis del equipo."],
          ["Figura 34:", `${N("es", "El precio exige una WACC nominal 285 pb menor")}`, "V6_dcf_inverso", "Fuente: análisis del equipo."]] },
    ],
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
    lead: "We recommend SELL on Mallplaza with a target price of CLP 2,076 per share, 41.6% below the CLP 3,555 close of 30-Sep-26 (−39.7% total return including dividends). Our target sits even below the stock's 52-week low (CLP 2,470).",
    paginas: [
      { texto: [
          { h1: "A PRICE NO METHOD REACHES" },
          { p: `In twelve months the stock rose from CLP 2,470 to 4,700 and now trades at CLP 3,555. None of our methods reaches that price: the DCF ranges from CLP 1,379 to 2,329 even with WACC ±100 bp and g between 2.5% and 4.0%${N("en", " (nominal)")}, and the most favourable multiples fall short too (Figure 27). We give the DCF an 80% weight because Mallplaza is, at heart, a rent collector: 81.4% of revenue is rent and 93% of that is fixed. The 20% on multiples tests the model against what the market pays today, but with caution: Chile has only two direct peers, and neither captures the company's expansion plan or debt path.` },
          { p: "Figure 27 sums up the thesis in one picture. APV, which values the firm unlevered and adds the tax shields separately, gives CLP 1,894, just 2% above the DCF: the result does not depend on how we model financing. Multiples range from CLP 2,152 (common-set median) to 3,294 (Chilean median), and EV/EBITDA NTM and EV/EBIT sit between CLP 2,914 and 3,098. Everything falls to the left of the price; the only thing that reaches it is the stock's own trading range over the past year. Even the Chilean median multiple, the most favourable case for the stock, sits 7% below the price." },
          { h1: "THE STORY IS IN THE RATE" },
          { p: "We project free cash flow to the firm in nominal CLP from 30-Jun-26, with an explicit horizon to 2031E and a terminal value that accounts for 76% of EV (Appendix [B]). The CLP 1,857 per share is CLP 583 of explicit cash flows and CLP 1,884 of terminal value, less CLP 610 of debt and minorities: the result depends, above all, on the discount rate." },
          { p: `That rate moves with the capital structure. Mallplaza uses its debt capacity to grow, as in 2024 after buying Open Plaza Peru (3.78x): we assume net debt/EBITDA rises from 2.36x to 3.5x by 2031, with 25 bp more real cost per turn, and that new debt (CLP 1.40 trillion) funds all capex while operating cash is distributed to shareholders. The nominal WACC thus moves from 10.54% in 2H26E to 10.41% in the terminal period (7.20% real): an ${N("en", "11.93% nominal Ke (8.67% real; Rf 4.98%")}, relevered beta of 1.14 from a six-peer median of 0.94, ERP 4.14% and a 1.41% revenue-weighted CRP) and a ${N("en", "5.11% nominal after-tax Kd")} based on the ${N("en", "3.6% real (UF) coupon")} of the latest bond (Appendix [Z]).` },
          { p: `We build Ke in dollars because its three inputs are in that currency${N("en", " and are nominal")}: the US Treasury Rf, the implied equity risk premium published by Damodaran and country risk, which we weight by revenue (71.8% Chile, 17.7% Peru and 10.5% Colombia). We then convert it to pesos with the gap between expected inflation in Chile (3.0%, the central bank's target) and in the US (2.24%, 10-year breakeven). ${N("en", "The beta is unlevered with Hamada (six regional peers) and relevered each period (1.09 in 2H26E, 1.14 in the terminal period).")}` },
          { p: `${N("en", `The first signal is the IRR, what someone buying at CLP 3,555 earns if our cash flows come true: for the firm, ${TT.en.pc(TIR.firma.nom)} nominal (${TT.en.pc(TIR.firma.real)} real in UF), ${TT.en.pb(TIR.firma.brecha_nom)} bp below its WACC and just ~80 bp real above its real cost of debt (Figure 28); for the shareholder, ${TT.en.pc(TIR.acc_mercado.nom, 1)} nominal against a ${TT.en.pc(TIR.ke)} Ke.`)} The market is pricing the stock almost as if it were a bond.` },
          { t: tablaTIR("en") },
          { h1: "GROWTH DOES NOT RESCUE THE PRICE" },
          { p: "This is not pessimism about operations. We forecast a 79.7% EBITDA margin in 2026E (excluding revaluation, below the 80.5% LTM), a tax rate converging from 16.8% to 27%, the announced pipeline (USD 613 million and 170,200 m²) and, from 2029, 27,075 m² of new GLA a year at the historical pace. The 13-analyst consensus expects NTM EBITDA only ~3% above ours: the disagreement with the market is not about cash flows." },
          { p: `Capex follows the company's plan: the pipeline and maintenance (5.5% of revenue) take base capex to a peak of CLP 240,660 million in 2027, and continuous expansion adds CLP 95,000–104,000 million a year from 2028, at the cost of the identified pipeline (USD 3,603 per m²) with a 35%, 90% and 100% revenue ramp. Working capital, at 15.6% of revenue, requires only CLP 7,100–9,900 million a year and moves the price by ±1.2% (Figure 29${N("en", "; Appendix [C]")}).` },
          { p: "Two assumptions deserve explanation. The cash tax rate starts at 16.8% because the company currently uses tax goodwill and loss carry-forwards; we assume those shields run off and the rate reaches 27% by 2030, so as not to overstate value. And working capital is positive because tenant receivables (CLP 122,370 million) exceed rents collected in advance (CLP 20,247 million): it grows with revenue, not with new m²." },
          { p: `The problem is the return on growth. Each new m² yields ${N("en", "6.9% real pre-tax")} on cost, a 5.0% real RONIC after tax, below a 7.2% real WACC. Expansion adds revenue but not value, so terminal value grows only with inflation (g of 3.0% nominal, 0% real) and a higher g lowers value (Figure 30). Assuming more growth would not close the gap with the price; it would widen it.` },
          { p: `That is why the explicit horizon includes the capex and ramp-up of projects opening through 2032 (108,301 m² of mature GLA in the terminal year), but from then on the model adds no new malls: terminal capex covers maintenance only (5.5% of revenue). ${N("en", "At these returns, growing more would reduce its value.")}` },
        ],
        figuras: [
          ["Figure 27:", "No method justifies the current price (CLP per share)", "V1_football", "52-week intraday range (1-Oct-25 to 30-Sep-26; low 10-Oct-25, high 27-Apr-26). Source: Team analysis; Bloomberg."],
          ["Figure 28:", `${N("en", "At CLP 3,555 the firm yields almost the same as its debt")}`, "V2_tasas", "Real rates in UF. Source: Central Bank of Chile (30-Jun-26); Mallplaza; Team analysis."],
          ["Figure 29:", "Working capital: an assumption that does not move the price (±1.2%)", "V4_capital_trabajo", "Source: Mallplaza financial statements; Team analysis."],
          ["Figure 30:", "Growing below the WACC destroys value: real g = 0%", "V3_value_driver", "Top: DCF value by g. Bottom: relative terminal value by RONIC. Source: Team analysis."]] },
      { texto: [
          { h1: "PEERS CONFIRM THE DIRECTION, NOT THE SIZE" },
          { p: `Mallplaza trades at 16.8x EV/EBITDA, above the Chilean median (15.7x) and far above Brazilian peers (9.0x). The gap with Brazil is about rates, not quality: with local real bonds at 2.5% in Chile and 8.1% in Brazil (NTN-B), the cap-rate spread${N("en", " (comparable to a real rate, since rents are indexed)")} over the local rate is similar in both countries, 385 bp in Chile and 302 bp in Brazil (Figure 31). We therefore weight the Chilean median at 70% and the common-set median at 30% (which includes IRSA; we exclude Fibra Danhos for its FIBRA regime): 14.4x, giving CLP 2,951 per share. The weighting matters little: between 0% and 100% Chile, the target price only moves from CLP 1,916 to 2,144 (Appendix [E]).` },
          { p: `Macro explains the difference. At 30-Jun-26 Brazil had ${N("en", "a 14.25% nominal Selic")}, a 14.34% nominal 10-year bond and 5.3% expected 2026 inflation, against ${N("en", "a 4.5% nominal policy rate")}, a 5.43% nominal bond and 4.1% inflation in Chile, with similar ${N("en", "real GDP")} growth (2.0% and 1.6%). An investor in Brazil demands more from every asset, not just shopping malls; applying that discount to Mallplaza, which earns 71.8% of revenue in Chile, would penalise it for a risk it does not carry. Mallplaza's own spread (346 bp) is in fact tighter than the Chilean median (385 bp): investors already pay it a premium within Chile.` },
          { p: "We chose the set on the same basis as the beta. IRSA is kept with caution, since it reports under IAS 29 and its NTM consensus has a single analyst; Fibra Danhos is excluded because, as a FIBRA, it must distribute 95% of its taxable income and has no beta in the WACC. We use EV/EBITDA because it is neutral to leverage and to revaluation (IAS 40), and we measure everyone with Bloomberg net debt (CLP 1,311,569 million for Mallplaza, close to the CLP 1,285,740 million of net financial debt used in the DCF)." },
          { p: "Nor does quality justify the premium. Mallplaza trades 29% above Cencosud Shopping, yet its EBITDA per m² is only 1.8% higher (see Industry); only at Parque Arauco's multiple (18.4x), with 5.7x leverage, would it be worth more than its price. Every multiple, from EV/EBITDA NTM (CLP 2,914) to P/BV (CLP 3,171), sits below the price (Figure 32): multiples confirm the stock is expensive, and the DCF sets the size of the gap. We exclude P/E (CLP 8,644), distorted by property revaluation." },
          { p: "In price terms: at the Chilean median alone the stock would be worth CLP 3,294 (−7%); at Cencosud's multiple, CLP 2,621; and only at Parque Arauco's, CLP 3,967. Defending CLP 3,555 with multiples requires assuming Mallplaza deserves the valuation of the most leveraged peer in the group. The forward view tells the same story: Mallplaza trades at 14.8x consensus NTM EBITDA against a 12.5x weighted peer multiple. We keep Brazil at 30% rather than zero because the Chilean median rests on only two companies." },
          { h1: "WHAT THE MARKET WOULD HAVE TO BELIEVE" },
          { p: `The price implies an EV of CLP 9.12 trillion, 69% above ours, and a 6.0% cap rate, against 10.1% in our DCF, 6.4% for Chilean peers and 7.5% at book value (Figure 33). The market pays 1.61x an equity base already at fair value, appraised at a ~5.55% real rate (Note 11): it pays more than the appraisers themselves. To justify CLP 3,555, the ${N("en", "nominal WACC would have to be 285 bp lower (7.56%; ")}Figure 34); even at the appraisers' rate${N("en", " (5.55% real, 8.72% nominal)")}, the DCF would be CLP 2,660.` },
          { p: `We pushed the thesis to its extremes. In a bear case (${N("en", "nominal WACC +100 bp")} and the 75.8% quarterly-low margin) the target falls to CLP 1,713; in a bull case combining ${N("en", "nominal WACC −100 bp, nominal g of 2.5%")}, the 81.3% quarterly-peak margin and the cut in corporate tax to 23%, it rises to CLP 2,627, still 26% below the price. APV (CLP 1,894), a peso-bond Rf (~CLP 2,439) and other debt targets (DCF of CLP 1,819 to 1,894) lead to the same conclusion (Appendix [F]). ${N("en", `The firm's IRR at the market price barely moves across scenarios (${TT.en.pc(TIR.esc.bear.nom, 1)}, ${TT.en.pc(TIR.esc.base.nom, 1)} and ${TT.en.pc(TIR.esc.bull.nom, 1)} nominal in bear, base and bull; ${TT.en.pc(TIR.esc.bear.real, 1)}, ${TT.en.pc(TIR.esc.base.real, 1)} and ${TT.en.pc(TIR.esc.bull.real, 1)} real) and always stays below each one's WACC (${TT.en.pb(TIR.esc.bear.brecha_nom)}, ${TT.en.pb(TIR.esc.base.brecha_nom)} and ${TT.en.pb(TIR.esc.bull.brecha_nom)} bp nominal): even the bull case leaves no margin of safety.`)}` },
          { p: `APV also shows what the debt is really worth. Of its CLP 1,894 per share, CLP 1,632 is the unlevered firm and CLP 261 the tax shields of projected debt: levering up to 3.5x adds value, but even with it the price is far away. And our risk-free choice is the conservative one: using ${N("en", "the nominal peso bond (5.43%)")} and then adding country risk would count Chile risk twice, and even so the target price would only rise to ~CLP 2,439, 31% below the price. In other words, no reasonable choice of inputs rescues the current valuation; the gap is too wide to be a modelling artefact.` },
          { p: "Calendar details do not change the outcome either. With an expected dividend of ~CLP 69 per share (2.0%), the expected total return is −39.7%. And since the DCF is dated 30-Jun-26 and the price 30-Sep-26, the comparison slightly favours the market: if the balance of the 2026 dividend (~CLP 27 per share) was already paid before that date, our value would be somewhat overstated, not understated." },
          { p: `That brings us to CLP 2,076 (80% × 1,857 + 20% × 2,951): 10.8x EBITDA, a 9.3% cap rate and 0.94x book value. Rolling the DCF to 30-Sep-26 would lift it to ~CLP 2,118 without changing the call. The catalyst is the same force behind the thesis: with the ${N("en", "nominal 10Y UST")} at 5.22%, its highest since 2002, rates pressure cap rates, and each new appraisal moves them toward our value. We would be wrong if cap rates compress further on institutional demand for defensive assets, or if the market sustains a structural premium for Tier A assets (see Investment Risks).` },
          { p: `What we will watch: the Note 11 appraisal at each reporting date (the real discount rate is the bridge between our value and book), the spread of the company's UF bonds over their ${N("en", "3.6% real coupon")} (at the latest issue's ${N("en", "4.33% real effective rate")}, the ${N("en", "nominal WACC would rise 12 bp")} and the target would fall to CLP 2,042), the cut in corporate tax to 23% (+5.1% on the target, not yet included) and the AAA rating, which would likely come under pressure at 3.5x net debt/EBITDA. In short, the operating story is sound; it is the price of that story that no longer adds up.` },
        ],
        figuras: [
          ["Figure 31:", "Chile and Brazil charge the same premium: multiples differ because of rates", "V7_multiplos", "Real bond: 10Y UF (Central Bank of Chile) and NTN-B 2035 (8.08%) at 30-Jun-26. Spread = cap rate (1 / EV/EBITDA) − real bond. IRSA omitted (IAS 29). Source: Bloomberg; Central Bank of Chile; Team analysis."],
          ["Figure 32:", "Value implied by each peer's multiple vs DCF and price (CLP per share)", "V9_pares", "Target = 80% DCF + 20% EV/EBITDA LTM. P/E excluded (includes revaluation). Source: Bloomberg (30-Sep-26); Team analysis."],
          ["Figure 33:", "The market pays 6.0%; our cash flows are worth 10.1%", "V5_cap_rate", "Gran Plaza: subsequent event, not in the model. Source: Mallplaza; Bloomberg; DF; Team analysis."],
          ["Figure 34:", `${N("en", "The price requires a nominal WACC 285 bp lower")}`, "V6_dcf_inverso", "Source: Team analysis."]] },
    ],
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

function tabla(cols, anchos, filas, { derecha = [], destacar = [], nuevo = false } = {}) {
  const borde = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  const enc = new TableRow({ tableHeader: true, children: cols.map((c, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: AZUL }, borders: sinBordes, margins: { top: 35, bottom: 35, left: 80, right: 80 },
    children: [new Paragraph({ alignment: derecha.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [run(c, { bold: true, size: 13, color: "FFFFFF" })] })] })) });
  const cuerpo = filas.map((f, r) => new TableRow({ children: f.map((v, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: destacar.includes(r) ? "FDF1CC" : r % 2 ? "FFFFFF" : "F7F8FB" },
    borders: { ...sinBordes, bottom: borde }, margins: { top: 25, bottom: 25, left: 80, right: 80 }, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing: linea(216), alignment: derecha.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [run(String(v), { size: 13, bold: destacar.includes(r), color: TINTA, nuevo })] })] })) }));
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

// Anexos [B], [C], [E] y [F]: datos de graficos/figuras/datos_anexos.json (graficos/anexos_valuacion.py, modelo v2)
const DA = JSON.parse(fs.readFileSync(path.join(FIG, "datos_anexos.json"), "utf8"));
const AX = {
  es: {
    mm: "CLP MM", cpa: "CLP/acción", per: ["2S26E", "2027E", "2028E", "2029E", "2030E", "2031E", "Terminal"],
    pb: "pb", paises: { Chile: "Chile", Brasil: "Brasil", Argentina: "Argentina", "México": "México" },
    B: {
      titulo: "ANEXO [B]: FLUJO DE CAJA DESCONTADO",
      lead: (v) => `FCFF en CLP nominales desde el 30-jun-26, con convención de mitad de período y la WACC de cada período (Anexo [Z]). El valor terminal usa el FCFF 2032E con g de ${v.g} nominal (0% real, sin reinversión porque el RONIC queda bajo la WACC real) y la WACC terminal de ${v.w}, e incluye ${v.m2} m² de GLA que maduran después de 2031. Explica el ${v.vtev} del EV.`,
      sA: "A. FLUJO DE CAJA LIBRE DE LA FIRMA POR PERÍODO (CLP MM)",
      filas: [["Ingresos totales", "ing", "mm"], ["EBITDA", "ebitda", "mm"], ["Margen EBITDA", "margen", "%"], ["Tasa de impuesto en caja", "tasa", "%"],
        ["NOPAT", "nopat", "mm"], ["(+) D&A", "da", "mm"], ["Capex base (pipeline + mantención)", "capex_base", "mm"],
        ["Capex de expansión / reinversión terminal", "capex_exp", "mm"], ["Inversión en capital de trabajo", "ct", "mm"], ["FCFF", "fcff", "mm"],
        ["WACC nominal del período", "wacc", "%2"], ["Factor de descuento (mitad del período)", "fd", "f"], ["VP del FCFF / del valor terminal", "vp", "mm"]],
      sB: "B. DEL VALOR DE EMPRESA AL VALOR POR ACCIÓN",
      colsB: ["Concepto", "CLP MM", "CLP/acción"],
      puente: ["VP del FCFF explícito (2S26E–2031E)", "VP del valor terminal", "Valor de empresa (EV)", "(−) Deuda financiera neta 30-jun-26", "(−) Minoritarios (libro)", "Valor del patrimonio / por acción"],
      sC: "C. SENSIBILIDAD DEL VALOR DCF (CLP/ACCIÓN)",
      esquina: "Δ WACC \\ g",
      nota: (v) => `Valor terminal al 31-dic-31: CLP ${v.vt} MM (cap rate implícito ${v.cap}; EV/EBITDA UDM implícito ${v.mult}). 2S26E: medio año. Fuente: modelo v2, hojas DCF y Sensibilidad; ${v.acc} MM de acciones.`,
    },
    C: {
      titulo: "ANEXO [C]: SUPUESTOS DE LA PROYECCIÓN",
      lead: (v) => `El forecast base sigue el motor operativo del equipo (pipeline anunciado hasta 2028); desde ${v.primer} se suma una expansión continua al ritmo histórico de apertura. Los ingresos 2026E son año calendario; el período 2S26E descuenta el 1S26 real (CLP ${v.ing1s} MM de ingresos).`,
      sA: "A. PROYECCIÓN DE INGRESOS, CAPEX Y CAPITAL DE TRABAJO (CLP MM)",
      filas: [["Ingresos base (año calendario)", "ing_cal", "mm"], ["Crecimiento de ingresos totales", "crec", "%"], ["Ingresos de la expansión continua", "ing_exp", "mm"],
        ["Margen EBITDA base", "margen", "%"], ["Tasa de impuesto en caja", "tasa", "%"], ["Capex de proyectos en curso (USD MM)", "capex_usd", "u1"],
        ["USDCLP promedio", "usdclp", "u0"], ["Capex de proyectos en curso", "capex_proy", "mm"], ["Capex de mantención + TI (5,5% de ingresos)", "mant", "mm"],
        ["Capex base total", "capex_base", "mm"], ["Capex de expansión continua", "capex_exp", "mm"], ["Costo por m² de expansión (CLP MM, IPC)", "costo_m2", "u2"],
        ["Inversión en capital de trabajo", "ct", "mm"]],
      destacar: [9],
      sB: "B. EXPANSIÓN CONTINUA", sC: "C. OPERACIÓN, CAPITAL Y VALOR TERMINAL", cols: ["Supuesto", "Valor"],
      expansion: (v) => [["GLA nuevo por año desde " + v.primer, `${v.gla} m²`], ["Base: promedio orgánico 2022, 2023 y 2025", `${v.gla22} / ${v.gla23} / ${v.gla25} m²`],
        ["Pipeline en curso (Release 2Q26)", `USD ${v.pipe} MM / ${v.m2} m²`], ["Costo por m² (= pipeline)", `USD ${v.costo}`],
        ["Desembolso del capex", `${v.t2} en t−2 / ${v.t1} en t−1`], ["Rampa de ingresos (año 1 / 2 / 3+)", v.rampa],
        ["Mix Chile / Perú; margen EBITDA", `${v.mix} / ${v.mixpe}; ${v.mcl} / ${v.mpe}`], ["Yield on cost / RONIC real después de impuestos", `${v.yld} / ${v.ronic}`]],
      operacion: (v) => [["Margen EBITDA UDM 2T26", v.mudm], ["Capex de mantención + TI", `${v.mant} de los ingresos`], ["D&A / ingresos", v.da],
        ["Impuesto en caja", `${v.tax0} → ${v.taxlp} (2030)`], ["Capital de trabajo neto / ingresos (dic-25)", v.ct],
        ["Deudores / arriendos anticipados (dic-25, CLP MM)", `${v.deud} / ${v.antic}`], ["g terminal nominal (real)", `${v.g} (0,0%)`],
        ["DFN/EBITDA objetivo; Kd real por vuelta", `${v.dfn}x en ${v.anio}; +${v.spread} pb`], ["Payout sobre utilidad distribuible", v.payout]],
      nota: "Fuente: modelo v2, hojas Forecast_Base, Expansion, Capital_Trabajo y Supuestos (Release 2Q26; EEFF dic-25; Exhibits v6).",
    },
    E: {
      titulo: "ANEXO [E]: MÚLTIPLOS COMPARABLES",
      lead: (v) => `Usamos el mismo set de pares que la beta de la WACC. EV/EBITDA UDM es el método principal (neutral al apalancamiento y a la revalorización NIC 40); el múltiplo pondera ${v.cl} la mediana chilena y ${v.set} la del set común. P/E y P/BV mezclan pares a valor razonable y a costo: solo referencia. Precio al 30-sep-26; deuda neta Bloomberg y tipo de cambio al 30-jun-26.`,
      sA: "A. PARES (CLP MM; MÚLTIPLOS EN VECES)",
      cols: ["Empresa", "País", "Market cap", "EV", "EBITDA UDM", "EV/EBITDA UDM", "EV/EBITDA NTM", "EV/EBIT UDM", "P/E UDM", "P/BV"],
      excluida: " (excluida)", est: ["Mediana set común (6)", "Mediana Chile (2)", "Ponderado 70/30"],
      sB: "B. VALOR POR MÚLTIPLO", colsB: ["Múltiplo", "Ponderado", "Métrica CLP MM", "CLP/acción"],
      metodos: ["EV/EBITDA UDM (principal)", "EV/EBITDA NTM", "EV/EBIT UDM", "P/E UDM (ref.)", "P/BV (ref.)"],
      sC: "C. SENSIBILIDAD A LA PONDERACIÓN", colsC: ["Peso Chile", "Múltiplo", "Valor múlt.", "Precio obj.", "Δ"],
      nota: (v) => `Puente: ${v.mult}x × EBITDA UDM CLP ${v.ebitda} MM = EV ${v.ev} − deuda neta Bloomberg ${v.deuda} − minoritarios ${v.nci} = CLP ${v.valor} por acción (${v.acc} MM de acciones). Fibra Danhos queda fuera por su régimen FIBRA y por no tener beta en la WACC; IRSA reporta bajo NIC 29. Fuente: Bloomberg; modelo v2, hojas Comparables y Precio_Objetivo; Figuras_Valuacion, V7.`,
    },
    F: {
      titulo: "ANEXO [F]: ESCENARIOS Y ROBUSTEZ",
      lead: (v) => `Ningún escenario ni método alternativo alcanza el precio de CLP ${v.precio}. Los escenarios se calculan con la réplica del DCF v2 (reproduce exactamente el valor base) y el precio objetivo mantiene el ${v.peso} DCF y el valor por múltiplos de CLP ${v.mult}.`,
      sA: "A. ESCENARIOS", cols: ["Escenario", "Supuestos", "DCF (CLP)", "Precio objetivo", "vs precio"],
      esc: (v) => [["Bear", `WACC +${v.dw} pb en todos los períodos; margen EBITDA ${v.mlo} (mínimo histórico) en vez de ${v.m26}`],
        ["Base", `Modelo v2: WACC ${v.w0} → ${v.wt}; g ${v.g}; margen ${v.m26}`],
        ["Bull", `WACC −${v.dw} pb; g ${v.gb}; margen ${v.mhi} (máximo histórico); impuesto en caja con techo de ${v.tax}`],
        ["Tasadores", `WACC nominal ${v.wtas} (${v.wtasr} real, Nota 11) en todos los períodos`]],
      sB: "B. MÉTODOS Y SUPUESTOS ALTERNATIVOS", colsB: ["Prueba", "DCF (CLP)", "Precio obj.", "vs precio"],
      rob: (v) => [`APV (Ku ${v.ku}): valor sin deuda`, "(+) escudos fiscales de la deuda", "= APV por acción", `Rf bono en pesos ${v.bono} (WACC ${v.efecto} pb)`,
        `DCF VF: WACC constante ${v.wvf}`, `WACC implícita en el precio: ${v.wimp}`],
      sC: "C. OBJETIVO DE DEUDA VS SPREAD (DCF, CLP)", esquina: "DFN/EBITDA \\ spread", actual: " (actual)",
      nota: "Bear y bull: margen EBITDA del mínimo/máximo histórico (Earnings Releases 4Q23 y 3Q25) en todos los años; el bull no recalcula el escudo fiscal de la WACC. Fuente: réplica del DCF v2 (riesgos/graficos/tornado_operacional.py); modelo v2, hojas APV, Sensibilidad (tabla 2, estática) y DCF; Figuras_Valuacion, WACC_Construccion y V6.",
    },
  },
  en: {
    mm: "CLP million", cpa: "CLP/share", per: ["2H26E", "2027E", "2028E", "2029E", "2030E", "2031E", "Terminal"],
    pb: "bp", paises: { Chile: "Chile", Brasil: "Brazil", Argentina: "Argentina", "México": "Mexico" },
    B: {
      titulo: "APPENDIX [B]: DISCOUNTED CASH FLOW",
      lead: (v) => `FCFF in nominal CLP from 30-Jun-26, with the mid-period convention and each period's WACC (Appendix [Z]). Terminal value uses 2032E FCFF with ${v.g} nominal g (0% real, no reinvestment because RONIC is below the real WACC) and the ${v.w} terminal WACC, and includes ${v.m2} m² of GLA maturing after 2031. It accounts for ${v.vtev} of EV.`,
      sA: "A. FREE CASH FLOW TO THE FIRM BY PERIOD (CLP MILLION)",
      filas: [["Total revenue", "ing", "mm"], ["EBITDA", "ebitda", "mm"], ["EBITDA margin", "margen", "%"], ["Cash tax rate", "tasa", "%"],
        ["NOPAT", "nopat", "mm"], ["(+) D&A", "da", "mm"], ["Base capex (pipeline + maintenance)", "capex_base", "mm"],
        ["Expansion capex / terminal reinvestment", "capex_exp", "mm"], ["Investment in working capital", "ct", "mm"], ["FCFF", "fcff", "mm"],
        ["Nominal WACC of the period", "wacc", "%2"], ["Discount factor (mid-period)", "fd", "f"], ["PV of FCFF / of terminal value", "vp", "mm"]],
      sB: "B. FROM ENTERPRISE VALUE TO VALUE PER SHARE",
      colsB: ["Item", "CLP million", "CLP/share"],
      puente: ["PV of explicit FCFF (2H26E–2031E)", "PV of terminal value", "Enterprise value (EV)", "(−) Net financial debt 30-Jun-26", "(−) Minorities (book)", "Equity value / per share"],
      sC: "C. DCF VALUE SENSITIVITY (CLP/SHARE)",
      esquina: "Δ WACC \\ g",
      nota: (v) => `Terminal value at 31-Dec-31: CLP ${v.vt} million (implied cap rate ${v.cap}; implied LTM EV/EBITDA ${v.mult}). 2H26E: half year. Source: model v2, DCF and Sensibilidad sheets; ${v.acc} million shares.`,
    },
    C: {
      titulo: "APPENDIX [C]: FORECAST ASSUMPTIONS",
      lead: (v) => `The base forecast follows the team's operating engine (announced pipeline to 2028); from ${v.primer} a continuous expansion is added at the historical opening pace. 2026E revenue is calendar-year; the 2H26E period deducts actual 1H26 (CLP ${v.ing1s} million of revenue).`,
      sA: "A. REVENUE, CAPEX AND WORKING CAPITAL FORECAST (CLP MILLION)",
      filas: [["Base revenue (calendar year)", "ing_cal", "mm"], ["Total revenue growth", "crec", "%"], ["Continuous-expansion revenue", "ing_exp", "mm"],
        ["Base EBITDA margin", "margen", "%"], ["Cash tax rate", "tasa", "%"], ["Capex on projects under way (USD million)", "capex_usd", "u1"],
        ["Average USDCLP", "usdclp", "u0"], ["Capex on projects under way", "capex_proy", "mm"], ["Maintenance + IT capex (5.5% of revenue)", "mant", "mm"],
        ["Total base capex", "capex_base", "mm"], ["Continuous-expansion capex", "capex_exp", "mm"], ["Expansion cost per m² (CLP million, CPI)", "costo_m2", "u2"],
        ["Investment in working capital", "ct", "mm"]],
      destacar: [9],
      sB: "B. CONTINUOUS EXPANSION", sC: "C. OPERATIONS, CAPITAL AND TERMINAL VALUE", cols: ["Assumption", "Value"],
      expansion: (v) => [["New GLA per year from " + v.primer, `${v.gla} m²`], ["Basis: organic average 2022, 2023 and 2025", `${v.gla22} / ${v.gla23} / ${v.gla25} m²`],
        ["Pipeline under way (2Q26 release)", `USD ${v.pipe} million / ${v.m2} m²`], ["Cost per m² (= pipeline)", `USD ${v.costo}`],
        ["Capex disbursement", `${v.t2} in t−2 / ${v.t1} in t−1`], ["Revenue ramp (year 1 / 2 / 3+)", v.rampa],
        ["Chile / Peru mix; EBITDA margin", `${v.mix} / ${v.mixpe}; ${v.mcl} / ${v.mpe}`], ["Yield on cost / real after-tax RONIC", `${v.yld} / ${v.ronic}`]],
      operacion: (v) => [["LTM 2Q26 EBITDA margin", v.mudm], ["Maintenance + IT capex", `${v.mant} of revenue`], ["D&A / revenue", v.da],
        ["Cash tax rate", `${v.tax0} → ${v.taxlp} (2030)`], ["Net working capital / revenue (Dec-25)", v.ct],
        ["Receivables / advance rents (Dec-25, CLP million)", `${v.deud} / ${v.antic}`], ["Nominal (real) terminal g", `${v.g} (0.0%)`],
        ["Target net debt/EBITDA; real Kd per turn", `${v.dfn}x in ${v.anio}; +${v.spread} bp`], ["Payout on distributable income", v.payout]],
      nota: "Source: model v2, Forecast_Base, Expansion, Capital_Trabajo and Supuestos sheets (2Q26 release; Dec-25 financial statements; Exhibits v6).",
    },
    E: {
      titulo: "APPENDIX [E]: COMPARABLE MULTIPLES",
      lead: (v) => `We use the same peer set as the WACC beta. LTM EV/EBITDA is the main method (neutral to leverage and to IAS 40 revaluation); the multiple weights the Chilean median at ${v.cl} and the common-set median at ${v.set}. P/E and P/BV mix fair-value and cost-basis peers: reference only. Price at 30-Sep-26; Bloomberg net debt and FX at 30-Jun-26.`,
      sA: "A. PEERS (CLP MILLION; MULTIPLES IN TIMES)",
      cols: ["Company", "Country", "Market cap", "EV", "LTM EBITDA", "EV/EBITDA LTM", "EV/EBITDA NTM", "EV/EBIT LTM", "P/E LTM", "P/BV"],
      excluida: " (excluded)", est: ["Common-set median (6)", "Chile median (2)", "Weighted 70/30"],
      sB: "B. VALUE BY MULTIPLE", colsB: ["Multiple", "Weighted", "Metric (CLP mn)", "CLP/share"],
      metodos: ["EV/EBITDA LTM (main)", "EV/EBITDA NTM", "EV/EBIT LTM", "P/E LTM (ref.)", "P/BV (ref.)"],
      sC: "C. SENSITIVITY TO THE WEIGHTING", colsC: ["Chile weight", "Multiple", "Mult. value", "Target", "Δ"],
      nota: (v) => `Bridge: ${v.mult}x × LTM EBITDA of CLP ${v.ebitda} million = EV ${v.ev} − Bloomberg net debt ${v.deuda} − minorities ${v.nci} = CLP ${v.valor} per share (${v.acc} million shares). Fibra Danhos is excluded for its FIBRA regime and because it has no beta in the WACC; IRSA reports under IAS 29. Source: Bloomberg; model v2, Comparables and Precio_Objetivo sheets; Figuras_Valuacion, V7.`,
    },
    F: {
      titulo: "APPENDIX [F]: SCENARIOS AND ROBUSTNESS",
      lead: (v) => `No scenario or alternative method reaches the CLP ${v.precio} price. Scenarios are computed with the DCF v2 replica (which reproduces the base value exactly), and the target keeps the ${v.peso} DCF weight and the CLP ${v.mult} multiples value.`,
      sA: "A. SCENARIOS", cols: ["Scenario", "Assumptions", "DCF (CLP)", "Target price", "vs price"],
      esc: (v) => [["Bear", `WACC +${v.dw} bp in every period; EBITDA margin ${v.mlo} (historical low) instead of ${v.m26}`],
        ["Base", `Model v2: WACC ${v.w0} → ${v.wt}; g ${v.g}; margin ${v.m26}`],
        ["Bull", `WACC −${v.dw} bp; g ${v.gb}; margin ${v.mhi} (historical high); cash tax capped at ${v.tax}`],
        ["Appraisers", `Nominal WACC ${v.wtas} (${v.wtasr} real, Note 11) in every period`]],
      sB: "B. ALTERNATIVE METHODS AND INPUTS", colsB: ["Test", "DCF (CLP)", "Target", "vs price"],
      rob: (v) => [`APV (Ku ${v.ku}): unlevered value`, "(+) tax shields of debt", "= APV per share", `Rf peso bond ${v.bono} (WACC ${v.efecto} bp)`,
        `DCF VF: constant WACC ${v.wvf}`, `WACC implied by the price: ${v.wimp}`],
      sC: "C. DEBT TARGET VS SPREAD (DCF, CLP)", esquina: "ND/EBITDA \\ spread", actual: " (current)",
      nota: "Bear and bull: EBITDA margin at the historical low/high (4Q23 and 3Q25 earnings releases) in every year; the bull case does not recompute the WACC tax shield. Source: DCF v2 replica (riesgos/graficos/tornado_operacional.py); model v2, APV, Sensibilidad (table 2, static) and DCF sheets; Figuras_Valuacion, WACC_Construccion and V6.",
    },
  },
};

const NB = "\u00a0"; // celda vacía con la altura de la fuente de la tabla
const lead = (t) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 40, ...linea(216) },
  shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO }, children: [run(t, { size: 15 })] });
const dosColumnas = (izq, der) => new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [MEDIA, COL_SEP, MEDIA],
  borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
  rows: [new TableRow({ cantSplit: true, children: [celda(MEDIA, izq), celda(COL_SEP, [vacia()]), celda(MEDIA, der)] })] });

function anexosValuacion(lang) {
  const T = AX[lang], f = fmt(lang);
  const pc = (x, dec = 1) => f(x * 100, dec) + "%";
  const mm = (x) => f(x, 0);
  const pb = (x) => `${x > 0 ? "+" : ""}${f(x, 0)}`;
  const celdaFmt = (x, u) => x === null || x === undefined ? NB : u === "mm" ? mm(x) : u === "%" ? pc(x) : u === "%2" ? pc(x, 2)
    : u === "f" ? f(x, 4) : u === "u0" ? f(x, 0) : u === "u1" ? f(x, 1) : f(x, 2);
  const der = (n, desde = 1) => Array.from({ length: n - desde }, (_, i) => i + desde);
  const anchosPer = [2866, 1100, 1100, 1100, 1100, 1100, 1100, 1200];
  const out = [];

  // [B] DCF
  const B = DA.B, P = B.puente, V = B.vt;
  out.push(titulo(T.B.titulo), lead(T.B.lead({ g: pc(V.g), w: pc(V.wacc_t, 2), m2: f(V.m2_maduro, 0), vtev: pc(V.vt_ev, 0) })));
  out.push(subtitulo(T.B.sA, 80));
  const filasB = T.B.filas.map(([n, k, u]) => {
    const vals = B.filas[k].slice();
    if (k === "vp") vals[6] = P.vp_vt;
    return [n, ...vals.map((x) => celdaFmt(x, u))];
  });
  out.push(tabla(["", ...T.per], anchosPer, filasB, { derecha: der(8), destacar: [9] }));
  const filasP = [[P.vp_exp, P.ps_exp], [P.vp_vt, P.ps_vt], [P.ev, P.ev / P.acciones], [P.dfn, P.dfn / P.acciones], [P.nci, P.nci / P.acciones],
    [P.patrimonio, P.valor]].map(([a, b], i) => [T.B.puente[i], mm(a), mm(b)]);
  const sens = B.sens.v.map((r, i) => [`${pb(B.sens.dw[i] * 1e4)} ${T.pb}`, ...r.map((x) => mm(x))]);
  out.push(dosColumnas(
    [subtitulo(T.B.sB, 100), tabla(T.B.colsB, [3025, 1200, 1000], filasP, { derecha: [1, 2], destacar: [5] })],
    [subtitulo(T.B.sC, 100), tabla([T.B.esquina, ...B.sens.g.map((g) => "g " + pc(g))], [1225, 1000, 1000, 1000, 1000], sens, { derecha: [1, 2, 3, 4], destacar: [2] })]));
  out.push(fuente(T.B.nota({ vt: mm(V.vt), cap: pc(V.cap_ev), mult: f(V.ev_ebitda, 1) + "x", acc: mm(P.acciones) })));

  // [C] Supuestos
  const C = DA.C, S = C.sup;
  out.push(titulo(T.C.titulo, 260), lead(T.C.lead({ primer: S.primer, ing1s: mm(S.ing_1s26) })));
  out.push(subtitulo(T.C.sA, 80));
  const filasC = T.C.filas.map(([n, k, u]) => [n, ...C.filas[k].map((x) => celdaFmt(x, u))]);
  out.push(tabla(["", ...T.per.slice(0, 6)], [3466, 1200, 1200, 1200, 1200, 1200, 1200], filasC, { derecha: der(7), destacar: T.C.destacar }));
  const vC = { primer: S.primer, gla: mm(S.gla_anual), gla22: mm(S.gla_22), gla23: mm(S.gla_23), gla25: mm(S.gla_25), pipe: f(S.pipe_usd, 1), m2: mm(S.pipe_m2),
    costo: mm(S.costo_usd_m2), t2: pc(S.pago_t2, 0), t1: pc(1 - S.pago_t2, 0), rampa: S.rampa.map((x) => pc(x, 0)).join(" / "),
    mix: pc(S.mix_cl), mixpe: pc(1 - S.mix_cl), mcl: pc(S.margen_cl, 0), mpe: pc(S.margen_pe, 0), yld: pc(S.yield), ronic: pc(S.ronic),
    mudm: pc(S.ebitda_ltm / S.ing_ltm), mant: pc(S.mant), da: pc(S.da, 2), tax0: pc(C.filas.tasa[0]), taxlp: pc(S.tax_lp, 0), ct: pc(S.ct_ing),
    deud: mm(S.deudores), antic: mm(S.anticip), g: pc(S.g), dfn: f(S.dfn_obj, 1), anio: S.anio_obj, spread: f(S.spread * 1e4, 0), payout: pc(S.payout, 0) };
  out.push(dosColumnas(
    [subtitulo(T.C.sB, 100), tabla(T.C.cols, [2925, 2300], T.C.expansion(vC), { derecha: [1] })],
    [subtitulo(T.C.sC, 100), tabla(T.C.cols, [2925, 2300], T.C.operacion(vC), { derecha: [1] })]));
  out.push(fuente(T.C.nota));

  // [E] Múltiplos
  const E = DA.E, PE = E.puente;
  const xm = (x) => (x === null || x === undefined ? NB : f(x, 1) + "x");
  out.push(titulo(T.E.titulo, 0, true), lead(T.E.lead({ cl: pc(E.peso_cl, 0), set: pc(1 - E.peso_cl, 0) })));
  out.push(subtitulo(T.E.sA, 80));
  const filasE = E.pares.map((p) => [p.nombre + (p.set === 0 ? T.E.excluida : ""), T.paises[p.pais] || p.pais, mm(p.mcap), mm(p.ev), mm(p.ebitda), ...p.m.map(xm)]);
  ["med_set", "med_cl", "pond"].forEach((k, i) => filasE.push([T.E.est[i], NB, NB, NB, NB, ...E.est[k].map(xm)]));
  out.push(tabla(T.E.cols, [2066, 900, 1150, 1150, 1100, 900, 900, 850, 800, 850], filasE, { derecha: der(10, 2), destacar: [0, filasE.length - 1] }));
  const filasM = E.metodos.map(([mult, met, val], i) => [T.E.metodos[i], xm(mult), mm(met), mm(val)]);
  const base = E.pond.find((r) => Math.abs(r[0] - E.peso_cl) < 1e-9);
  const filasW = E.pond.map(([w, mult, val, po, d]) => [pc(w, 0), xm(mult), mm(val), mm(po), (d > 0 ? "+" : "") + pc(d)]);
  out.push(dosColumnas(
    [subtitulo(T.E.sB, 100), tabla(T.E.colsB, [1925, 900, 1300, 1100], filasM, { derecha: [1, 2, 3], destacar: [0] })],
    [subtitulo(T.E.sC, 100), tabla(T.E.colsC, [1025, 900, 1100, 1100, 1100], filasW, { derecha: [0, 1, 2, 3, 4], destacar: [E.pond.indexOf(base)] })]));
  out.push(fuente(T.E.nota({ mult: f(PE.mult, 2), ebitda: mm(PE.ebitda), ev: mm(PE.ev), deuda: mm(PE.deuda_bbg), nci: mm(PE.nci), valor: mm(PE.valor), acc: mm(PE.acciones) })));

  // [F] Escenarios y robustez
  const F = DA.F, Q = F.param, ES = F.esc;
  const obj = (x) => F.peso * x + (1 - F.peso) * F.mult;
  const vs = (x) => pc(x / F.precio - 1);
  out.push(titulo(T.F.titulo, 260), lead(T.F.lead({ precio: mm(F.precio), peso: pc(F.peso, 0), mult: mm(F.mult) })));
  out.push(subtitulo(T.F.sA, 80));
  const w0 = DA.B.filas.wacc[0];
  const textos = T.F.esc({ dw: f(Q.dw_bear * 1e4, 0), mlo: pc(Q.m_lo), m26: pc(Q.margen26), w0: pc(w0, 2), wt: pc(Q.wacc_t, 2), g: pc(DA.B.vt.g),
    gb: pc(Q.g_bull), mhi: pc(Q.m_hi), tax: pc(Q.tax_bull, 0), wtas: pc(Q.wacc_tas, 2), wtasr: pc(Q.wacc_tas_real, 2) });
  const filasF = ["bear", "base", "bull", "tas"].map((k, i) => [...textos[i], mm(ES[k].dcf), mm(ES[k].po), pc(ES[k].vs)]);
  out.push(tabla(T.F.cols, [1250, 5616, 1200, 1400, 1200], filasF, { derecha: [2, 3, 4], destacar: [1] }));
  const A = F.apv, R = F.rf;
  const nr = T.F.rob({ ku: pc(A.ku, 2), bono: pc(R.bono, 2), efecto: pb(R.efecto), wvf: pc(F.vf.wacc, 2), wimp: pc(F.implicita.wacc, 2), brecha: pb(F.implicita.brecha) });
  const filasR = [[nr[0], mm(A.sin_deuda), NB, NB], [nr[1], mm(A.escudos), NB, NB], [nr[2], mm(A.valor), mm(obj(A.valor)), vs(obj(A.valor))],
    [nr[3], mm(R.dcf), mm(R.po), vs(R.po)], [nr[4], mm(F.vf.dcf), mm(obj(F.vf.dcf)), vs(obj(F.vf.dcf))], [nr[5], mm(F.precio), NB, NB]];
  const filasD = F.deuda.filas.map(([t, ...v], i) => [f(t, i ? 1 : 2) + "x" + (i ? "" : T.F.actual), ...v.map((x) => mm(x))]);
  const iBase = F.deuda.filas.findIndex((r) => Math.abs(r[0] - DA.C.sup.dfn_obj) < 1e-9);
  out.push(dosColumnas(
    [subtitulo(T.F.sB, 100), tabla(T.F.colsB, [2625, 850, 900, 850], filasR, { derecha: [1, 2, 3], destacar: [2] })],
    [subtitulo(T.F.sC, 100), tabla([T.F.esquina, ...F.deuda.spread.map((s) => `${f(s * 1e4, 0)} ${T.pb}`)], [1825, 1100, 1150, 1150], filasD, { derecha: [1, 2, 3], destacar: [iBase] })]));
  out.push(fuente(T.F.nota));
  return out;
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
