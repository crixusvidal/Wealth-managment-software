// Página "Investment Risks" en español e inglés (A4, mismo formato que el informe).
// Figuras: graficos/informe/riesgos/{es,en}_fig2x_*.png (python graficos/pagina_riesgos.py ...).
// Uso: node informe/pagina_riesgos.js
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, ImageRun, Packer, PageNumber, Paragraph,
  ShadingType, Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType, LineRuleType,
} = require("docx");

const AZUL = "14284B", ROJO = "E9004B", TINTA = "1A1A1A", GRIS = "5F5F5F", FONDO = "EEF1F6";
const FUENTE = "Arial";
const COL_TEXTO = 6900, COL_SEP = 216, COL_FIG = 3550;
const FIG_DIR = path.join(__dirname, "..", "graficos", "informe", "riesgos");
const PULGADA_PX = 96; // docx-js mide las imágenes en píxeles a 96 dpi
const ALTOS = { matriz: 2.15, tornado: 2.35, heatmap: 1.75, uf: 1.5 }; // = graficos/pagina_riesgos.py
const ANCHO_FIG = COL_FIG / 1440;

const LANG = {
  es: {
    archivo: "Mallplaza_Investment_Risks_ES.docx", lang: "es-CL", pagina: "Página ",
    titulo: "RIESGOS DE INVERSIÓN",
    intro: "El perfil de riesgo de Mallplaza es moderado: su operación es resiliente y el valor depende sobre todo de variables de mercado. El principal riesgo, de probabilidad alta, es un alza de las tasas de largo plazo, que presiona el rendimiento de los bonos y eleva la WACC; le siguen la inflación y un menor dinamismo macro. Ni siquiera los escenarios favorables justifican el precio de mercado a la fecha del informe.",
    mitigantes: "Mitigantes: ",
    secciones: [
      { titulo: "EVALUACIÓN DE RIESGOS", parrafos: [
        "La Figura 23 ubica cada riesgo en una matriz 3x3 según su probabilidad y su impacto (bajo, medio, alto); el color de cada etiqueta indica su categoría. El impacto se mide por el efecto sobre el valor por acción, usando el tornado operacional (Figura 24), la sensibilidad a WACC y g (Figura 25) y la sensibilidad a la UF (Figura 26). Los riesgos que no se pudieron cuantificar con el modelo (RR1 y RR2) se ubican según el análisis cualitativo del equipo. El riesgo de tasas (FR1) ocupa la celda de mayor severidad (alta/alta) y el de inflación y ciclo macro (MR1) también queda en la zona roja; el deterioro operacional (OR1) y el cambiario (MR3) quedan en el centro."] },
      { titulo: "RIESGOS OPERACIONALES", riesgos: [
        { cod: "OR1", nombre: "Deterioro operacional.", texto: "Caídas de margen EBITDA, rentas, tráfico u ocupación hacia sus mínimos históricos. El peor caso, un margen de 75,8%, reduce el valor DCF a CLP 1.671 (−10,0%), y el piso histórico del crecimiento de rentas en mismas tiendas lo lleva a CLP 1.782 (−4,0%) (Figura 24).", tag: "No sistemático · Prob. media · Impacto medio",
          mit: "métricas estables: margen EBITDA de 76,2% (2022) a 80,5% (2025) y ocupación entre 95,5% y 96,6% en 2024-25. Tras el COVID, los ingresos del 1T22 ya superaban en 18% a los del 1T19." },
        { cod: "OR2", nombre: "Tiendas ancla.", texto: "Las anclas, en su mayoría del grupo Falabella, atraen el flujo de visitas. Su reducción o salida bajaría las visitas y las rentas de las tiendas menores. El impacto va de −1,3% a −4,0% (Figura 24).", tag: "No sistemático · Prob. baja · Impacto bajo",
          mit: "contratos de largo plazo (duración remanente promedio de 7,1 años) y capacidad probada de reconvertir espacios ancla en tiendas menores con mayor renta por m²." }] },
      { titulo: "RIESGOS FINANCIEROS", riesgos: [
        { cod: "FR1", nombre: "Tasas de interés y costo de capital.", texto: "Un alza de tasas largas elevaría la WACC y el costo de refinanciar la deuda y el pipeline. La probabilidad es alta: el UST 10Y pasó de 4,98% (Rf del modelo) a 5,22% al 8-oct-26, máximo en 24 años (desde 2002); la Fed subió su tasa en septiembre y las tasas largas locales también suben. Con WACC +50 pb y +100 pb, el precio objetivo cae 6,5% y 12,3% (Figura 25).", tag: "Sistemático · Prob. alta · Impacto alto",
          mit: "79% de la deuda en UF a tasa fija, duración de ~6,7 años, DFN/EBITDA de 2,36x y rating AA+/BBB acotan el traspaso al costo de la deuda; los arriendos en UF compensan alzas inflacionarias. El alza de la tasa de descuento del patrimonio no es cubrible." }] },
      { titulo: "RIESGOS DE MERCADO Y MACROECONÓMICOS", riesgos: [
        { cod: "MR1", nombre: "Inflación y ciclo macro.", texto: "El 79% de la deuda está en UF: una UF 3% más alta resta CLP 38.691 MM a la utilidad contable (Nota 26). Sobre el precio objetivo, en cambio, el efecto es levemente positivo: las rentas indexadas suman CLP 40,8 por acción y la mayor deuda en UF resta 17,7, con un neto de +23,1 (+1,1%) (Figura 26). El riesgo está en que un menor consumo impida traspasar la UF a las rentas: sin ese traspaso, la UF +3% resta 0,9% al precio objetivo. Un menor consumo también frena las nuevas líneas de negocio (Anexo [X]).", tag: "Sistemático · Prob. alta · Impacto medio",
          mit: "rentas indexadas a la UF, modelo basado en arriendo más que en ventas, deuda baja aunque creciente y capacidad activa de rotar el portafolio." },
        { cod: "MR2", nombre: "Comercio electrónico.", texto: "Si el consumo migra masivamente al canal online, el mall pierde tráfico y atractivo. El e-commerce en Chile superó US$ 10.000 MM en 2025 (12,6% de las ventas del comercio). El peor caso de tráfico solo resta 1,3% al valor (Figura 24).", tag: "Sistemático · Prob. baja · Impacto medio",
          mit: "el canal online está madurando sin efecto visible en el modelo (visitas +16% en 2025). Los malls ofrecen entretención y gastronomía, difíciles de replicar online." },
        { cod: "MR3", nombre: "Tipo de cambio.", texto: "Cerca del 28% de los ingresos viene de Perú y Colombia, y el modelo supone un sol que sube de CLP 268 a 290 entre 2026 y 2031. Una depreciación de 10% del PEN y del COP desde 2027E resta 4,3% al valor DCF, antes del efecto de las coberturas y de la deuda local (cálculo del equipo).", tag: "Sistemático · Prob. media · Impacto medio",
          mit: "Mallplaza radica su deuda en las monedas de cada país y, si no es posible, usa derivados de cobertura (forwards, swaps y cross currency swaps) (Análisis Razonado 2T26, p. 14)." } ] },
      { titulo: "RIESGOS REGULATORIOS Y ESG", riesgos: [
        { cod: "RR1", nombre: "Interlocking y conflictos de interés.", texto: "Falabella controla el 53,05% y arrienda espacios (Falabella, Sodimac, Tottus), lo que genera cuentas por cobrar con relacionadas. Fernando de Peña, gerente de Mallplaza por décadas, preside Falabella desde marzo de 2026. Contratos fuera de mercado afectarían directamente las rentas y el valor.", tag: "No sistemático · Prob. baja · Impacto alto",
          mit: "comité de directores de tres miembros y revisión de operaciones con partes relacionadas (Ley 18.046). Aun así, solo 1 de 9 directores es independiente." },
        { cod: "RR2", nombre: "Regulatorio y político.", texto: "Los tres gobiernos asumieron en 2026 con agendas pro inversión, pero con distinto margen fiscal y legislativo (ver Contexto político-económico): en Chile la rebaja del impuesto corporativo ya es ley, en Perú el gobierno no tiene mayoría propia y en Colombia el ajuste fiscal y las tasas altas pesan sobre el consumo. Un retraso en permisos posterga ingresos del pipeline (USD 613 MM, 170.200 m² hasta 2028), con sensibilidad acotada.", tag: "Sistemático · Prob. media · Impacto bajo",
          mit: "la Ley 21.770 (sep-25) acorta los plazos de permisos y Colombia aporta solo ~11% de los ingresos. El modelo mantiene 27%: aplicar la rebaja a la porción Chile subiría el valor DCF ~5,4% (+3,9% el precio objetivo), aún lejos del precio de mercado." }] },
      { titulo: "ANÁLISIS DE SENSIBILIDAD", pag2: true, parrafos: [
        "El valor es más sensible a la tasa de descuento que a la operación en los escenarios evaluados: el peor caso operacional (margen EBITDA) resta 10,0% al valor DCF (Figura 24), mientras que +100 pb de WACC restan 17,2% (12,3% del precio objetivo, Figura 25). El precio actual implica un cap rate cercano a 6,0%, frente a 10,1% en el DCF: el mercado descuenta una tasa baja y la acción queda más expuesta a un alza de tasas. Un g mayor resta valor: RONIC 5,0% < WACC real 7,2%. La UF casi no mueve el valor (+1,1% con UF +3%), porque las rentas indexadas compensan la deuda en UF. Hasta la mejor celda de la Figura 25 (CLP 2.487) queda bajo el precio (3.555): se mantiene la recomendación."] },
    ],
    politica: {
      titulo: "CONTEXTO POLÍTICO-ECONÓMICO 2026",
      cab: ["País", "Gobierno", "Medidas y señales", "Implicancia para Mallplaza"],
      filas: [
        ["Chile\n~71% ingresos", "José Antonio Kast\n(desde mar-26)",
          "Promulgó la Ley de Reconstrucción (9-oct-26): impuesto corporativo de 27% a 23% gradual desde 2027 e invariabilidad tributaria [1]. Presupuesto 2027 con gasto +2,7% vs la ley 2026 [2]. S&P reafirmó A estable (sep-26) [3]. TPM 4,5% [4]. Senado estable, Cámara frágil [5].",
          "Positivo: menor impuesto (+3,9% al precio objetivo si se aplica, no incluido) y permisos más rápidos (Ley 21.770). Riesgo: aprobar el resto de la agenda depende de una Cámara sin mayoría holgada."],
        ["Perú\n~19% ingresos", "Keiko Fujimori\n(desde 28-jul-26)",
          "Salario mínimo de S/ 1.130 a 1.300 y Pensión 65 duplicada (S/ 350 a 700) [6]. Presupuesto 2027 de S/ 266.506 MM, coherente con déficit de 1,4% del PIB [7]. BCRP mantiene 4,25% [8]. Fuerza Popular sin mayoría en el Congreso bicameral [9].",
          "Positivo para el consumo y las ventas de locatarios. Riesgo: gobernabilidad y costo fiscal de leyes del Congreso anterior (hasta S/ 50.000 MM según el MEF) [7]."],
        ["Colombia\n~11% ingresos", "Abelardo de la Espriella\n(desde 7-ago-26)",
          "Hacienda estima un déficit de 7,8% del PIB y propone recortar el gasto en al menos 1% del PIB [10]; reforma tributaria para bajar la carga a las empresas [11]. El Banco de la República subió su tasa a 12,25% (30-sep-26) [12]. S&P BB- estable [13].",
          "Negativo de corto plazo: ajuste fiscal y tasas altas frenan el consumo y presionan el COP (MR3). Peso acotado en el valor."],
      ],
      fuentes: "[1] Emol, 9-oct-26, “Ley de Reconstrucción promulgada”; CNN Chile, oct-26. [2] Diario Financiero / El Mostrador, 3-oct-26. [3] Ministerio de Hacienda, informe sobre Fitch y S&P, sep-26. [4] Banco Central de Chile, RPM oct-26. [5] El Dínamo, 11-mar-26. [6] Infobae / La República (Perú), 28-jul-26. [7] Infobae, 29-ago-26; El Comercio (MEF). [8] Infobae, 11-sep-26. [9] América TV, 2026. [10] El Tiempo / La Silla Vacía, jul-26. [11] Valora Analitik / El Colombiano, 2026. [12] Banco de la República, comunicado 30-sep-26. [13] IRC – MinHacienda Colombia. Participación en ingresos 2026E: modelo del equipo. Fuente: análisis del equipo.",
    },
    figuras: [
      { n: "Figura 23:", t: "Matriz de riesgos", img: "es_fig23_matriz.png", alto: ALTOS.matriz,
        fuente: "Color de fondo = severidad (probabilidad × impacto). Fuente: análisis del equipo." },
      { n: "Figura 24:", t: "Tornado de riesgo operacional (CLP por acción)", img: "es_fig24_tornado.png", alto: ALTOS.tornado,
        fuente: "Límites = mínimo y máximo histórico de cada variable: releases de resultados de Mallplaza 4T23–2T26 (margen EBITDA, ocupación, rentas y ventas en mismas tiendas, visitas) y EEFF (capex de mantención 2024 y UDM 2T26; rango de crecimiento de la tasación, Nota 11). WACC terminal 10,41% y g 3,0% constantes. Fuente: Mallplaza; análisis del equipo." },
      { n: "Figura 25:", t: "Sensibilidad del precio objetivo: WACC vs g (CLP)", img: "es_fig25_heatmap.png", alto: ALTOS.heatmap,
        fuente: "Precio objetivo = 80% DCF + 20% múltiplos EV/EBITDA (CLP 2.951). La columna g = 2,0% usa la fórmula de la hoja Sensibilidad. Fuente: equipo." },
      { n: "Figura 26:", t: "Sensibilidad del precio objetivo a un alza de la UF (CLP por acción)", img: "es_fig26_uf.png", alto: ALTOS.uf, pag2: true,
        fuente: "Alza de nivel de la UF desde 2027E, con WACC y g constantes. Rentas: IPC de Chile del modelo (80% del precio objetivo vía DCF); deuda: exposición neta en UF implícita en la Nota 26 (CLP 1,29 billones). Utilidad: +3% reportado, +1% y +2% por extrapolación lineal. Fuente: EEFF de Mallplaza a jun-26, Nota 26; análisis del equipo." },
    ],
  },
  en: {
    archivo: "Mallplaza_Investment_Risks_EN.docx", lang: "en-US", pagina: "Page ",
    titulo: "INVESTMENT RISKS",
    intro: "Mallplaza's risk profile is moderate: its operations are resilient and value depends mainly on market variables. The main risk, with high likelihood, is a rise in long-term rates, which pushes up bond yields and the WACC; it is followed by inflation and weaker macro momentum. Not even the favourable scenarios justify the market price at the report date.",
    mitigantes: "Mitigants: ",
    secciones: [
      { titulo: "RISK ASSESSMENT", parrafos: [
        "Figure 23 places each risk in a 3x3 matrix by likelihood and impact (low, medium, high); the colour of each label shows its category. Impact is measured as the effect on value per share, using the operational tornado (Figure 24), the WACC and g sensitivity (Figure 25) and the UF sensitivity (Figure 26). Risks that could not be quantified with the model (RR1 and RR2) are placed according to the team's qualitative analysis. Interest-rate risk (FR1) sits in the most severe cell (high/high), and inflation and the macro cycle (MR1) also fall in the red zone; operating deterioration (OR1) and FX (MR3) sit in the centre."] },
      { titulo: "OPERATIONAL RISKS", riesgos: [
        { cod: "OR1", nombre: "Operating deterioration.", texto: "EBITDA margin, rents, footfall or occupancy falling to their historical lows. The worst case, a 75.8% margin, cuts DCF value to CLP 1,671 (−10.0%), and the historical floor of same-store rent growth takes it to CLP 1,782 (−4.0%) (Figure 24).", tag: "Non‑systematic · Medium likelihood · Medium impact",
          mit: "stable metrics: EBITDA margin rose from 76.2% (2022) to 80.5% (2025) and occupancy held between 95.5% and 96.6% in 2024-25. After COVID, 1Q22 revenue was already 18% above 1Q19." },
        { cod: "OR2", nombre: "Anchor stores.", texto: "Anchors, mostly from the Falabella group, drive footfall. Shrinking or losing them would lower visits and rents from smaller stores. The impact ranges from −1.3% to −4.0% (Figure 24).", tag: "Non‑systematic · Low likelihood · Low impact",
          mit: "long-term leases (average remaining term of 7.1 years) and a proven ability to convert anchor space into smaller stores with higher rent per m²." }] },
      { titulo: "FINANCIAL RISKS", riesgos: [
        { cod: "FR1", nombre: "Interest rates and cost of capital.", texto: "Higher long-term rates would raise the WACC and the cost of refinancing debt and the pipeline. Likelihood is high: the 10Y UST moved from 4.98% (model Rf) to 5.22% on 8-Oct-26, a 24-year high (since 2002); the Fed hiked in September and the Central Bank of Chile reports rising local long-term rates. With WACC +50 bp and +100 bp, the target price falls 6.5% and 12.3% (Figure 25).", tag: "Systematic · High likelihood · High impact",
          mit: "79% of debt is fixed-rate in UF, with ~6.7 years of duration, 2.36x net debt/EBITDA and AA+/BBB ratings, which limit the pass-through to the cost of debt; UF-linked rents offset inflation. A higher cost of equity cannot be hedged." }] },
      { titulo: "MARKET & MACROECONOMIC RISKS", riesgos: [
        { cod: "MR1", nombre: "Inflation and macro cycle.", texto: "79% of debt is in UF: a 3% higher UF takes CLP 38,691 million off accounting earnings (Note 26). On the target price, however, the effect is slightly positive: indexed rents add CLP 40.8 per share and the higher UF debt subtracts 17.7, a net +23.1 (+1.1%) (Figure 26). The risk is that weaker consumption prevents passing the UF through to rents: without that pass-through, UF +3% takes 0.9% off the target price. Weaker consumption also slows the new business lines (Appendix [X]).", tag: "Systematic · High likelihood · Medium impact",
          mit: "UF-indexed rents, a model based on rent rather than sales, low (if rising) leverage and an active ability to rotate the portfolio." },
        { cod: "MR2", nombre: "E‑commerce.", texto: "If consumption migrates massively online, malls lose footfall and appeal. E‑commerce in Chile exceeded US$10 billion in 2025, around 12.6% of retail sales. The worst traffic case only takes 1.3% off value (Figure 24).", tag: "Systematic · Low likelihood · Medium impact",
          mit: "the online channel is maturing with no visible effect on the model (visits +16% in 2025). Malls offer entertainment and dining, hard to replicate online." },
        { cod: "MR3", nombre: "Foreign exchange.", texto: "Around 28% of revenue comes from Peru and Colombia, and the model assumes the sol rises from CLP 268 to 290 between 2026 and 2031. A 10% depreciation of the PEN and COP from 2027E takes 4.3% off DCF value, before hedges and the effect on local-currency debt (team estimate).", tag: "Systematic · Medium likelihood · Medium impact",
          mit: "Mallplaza raises debt in each country's currency and, where not possible, uses hedging derivatives (forwards, swaps and cross currency swaps) (Management Discussion 2Q26, p. 14)." } ] },
      { titulo: "REGULATORY & ESG RISKS", riesgos: [
        { cod: "RR1", nombre: "Interlocking and conflicts of interest.", texto: "Falabella controls 53.05% and leases space through Falabella, Sodimac and Tottus, which creates receivables with related parties. Fernando de Peña, Mallplaza's CEO for decades, has chaired Falabella since March 2026. Off-market contracts would directly affect rents and value.", tag: "Non‑systematic · Low likelihood · High impact",
          mit: "a three-member directors' committee and review of related-party transactions (Law 18,046). Even so, safeguards are limited: only 1 of 9 directors is independent." },
        { cod: "RR2", nombre: "Regulatory and political.", texto: "All three governments took office in 2026 with pro-investment agendas, but with different fiscal and legislative room (see Political and economic context): in Chile the corporate tax cut is already law, in Peru the government lacks its own majority, and in Colombia fiscal tightening and high rates weigh on consumption. Permit delays push back pipeline revenue (USD 613 million, 170,200 m² to 2028), with limited value sensitivity.", tag: "Systematic · Medium likelihood · Low impact",
          mit: "Law 21,770 (Sep-25) shortens permitting times and Colombia contributes only ~11% of revenue. The model keeps 27%: applying the cut to the Chilean share would lift DCF value ~5.4% (+3.9% on the target price), still far from the market price." }] },
      { titulo: "SENSITIVITY ANALYSIS", pag2: true, parrafos: [
        "Value is more sensitive to the discount rate than to operations in the scenarios assessed. The worst operating case, EBITDA margin, takes 10.0% off DCF value (Figure 24), while +100 bp of WACC takes 17.2% off (12.3% of the target price, Figure 25). The current price implies a cap rate close to 6.0%, versus 10.1% in the DCF: the market prices in a low discount rate, which increases the stock's exposure to higher rates. A higher g lowers value because RONIC (5.0%) is below the real WACC (7.2%). The UF barely moves value (+1.1% with UF +3%), as indexed rents offset UF debt. Even the best cell in Figure 25 (CLP 2,487) is below the CLP 3,555 share price, so the recommendation stands."] },
    ],
    politica: {
      titulo: "POLITICAL AND ECONOMIC CONTEXT 2026",
      cab: ["Country", "Government", "Measures and signals", "Implication for Mallplaza"],
      filas: [
        ["Chile\n~71% of revenue", "José Antonio Kast\n(since Mar-26)",
          "Signed the Reconstruction Law (9-Oct-26): corporate tax from 27% to 23% gradually from 2027 and tax invariability [1]. 2027 budget with spending +2.7% vs the 2026 law [2]. S&P affirmed A stable (Sep-26) [3]. Policy rate 4.5% [4]. Stable Senate, fragile lower house [5].",
          "Positive: lower tax (+3.9% on the target price if applied, not included) and faster permits (Law 21,770). Risk: passing the rest of the agenda depends on a lower house without a comfortable majority."],
        ["Peru\n~19% of revenue", "Keiko Fujimori\n(since 28-Jul-26)",
          "Minimum wage from S/ 1,130 to 1,300 and Pensión 65 doubled (S/ 350 to 700) [6]. 2027 budget of S/ 266,506 million, consistent with a 1.4% of GDP deficit [7]. BCRP holds at 4.25% [8]. Fuerza Popular lacks a majority in the bicameral Congress [9].",
          "Positive for consumption and tenant sales. Risk: governability and the fiscal cost of laws passed by the previous Congress (up to S/ 50,000 million per the MEF) [7]."],
        ["Colombia\n~11% of revenue", "Abelardo de la Espriella\n(since 7-Aug-26)",
          "The finance ministry puts the deficit at 7.8% of GDP and proposes spending cuts of at least 1% of GDP [10]; tax reform to lower the corporate burden [11]. The central bank raised its rate to 12.25% (30-Sep-26) [12]. S&P BB- stable [13].",
          "Short-term negative: fiscal tightening and high rates slow consumption and pressure the COP (MR3). Limited weight in value."],
      ],
      fuentes: "[1] Emol, 9-Oct-26, “Ley de Reconstrucción promulgada”; CNN Chile, Oct-26. [2] Diario Financiero / El Mostrador, 3-Oct-26. [3] Chilean Ministry of Finance, note on Fitch and S&P, Sep-26. [4] Central Bank of Chile, Oct-26 meeting. [5] El Dínamo, 11-Mar-26. [6] Infobae / La República (Peru), 28-Jul-26. [7] Infobae, 29-Aug-26; El Comercio (MEF). [8] Infobae, 11-Sep-26. [9] América TV, 2026. [10] El Tiempo / La Silla Vacía, Jul-26. [11] Valora Analitik / El Colombiano, 2026. [12] Banco de la República, statement 30-Sep-26. [13] IRC – Colombian Ministry of Finance. Revenue shares 2026E: team model. Source: Team analysis.",
    },
    figuras: [
      { n: "Figure 23:", t: "Risk matrix", img: "en_fig23_matriz.png", alto: ALTOS.matriz,
        fuente: "Background colour = severity (likelihood × impact). Source: Team analysis." },
      { n: "Figure 24:", t: "Operational risk tornado (CLP per share)", img: "en_fig24_tornado.png", alto: ALTOS.tornado,
        fuente: "Limits = historical low and high of each driver: Mallplaza earnings releases 4Q23–2Q26 (EBITDA margin, occupancy, same-store rent and sales, visitors) and financial statements (maintenance capex 2024 and LTM 2Q26; appraisal growth range, Note 11). Terminal WACC 10.41% and g 3.0% held constant. Source: Mallplaza; Team analysis." },
      { n: "Figure 25:", t: "Target price sensitivity: WACC vs g (CLP)", img: "en_fig25_heatmap.png", alto: ALTOS.heatmap,
        fuente: "Target price = 80% DCF + 20% EV/EBITDA multiples (CLP 2,951). The g = 2.0% column uses the same formula as the Sensibilidad sheet. Source: Team analysis." },
      { n: "Figure 26:", t: "Target price sensitivity to a rise in the UF (CLP per share)", img: "en_fig26_uf.png", alto: ALTOS.uf, pag2: true,
        fuente: "UF level shock from 2027E, with WACC and g constant. Rents: model's Chile CPI (80% of the target price via DCF); debt: net UF exposure implied by Note 26 (CLP 1.29 trillion). Earnings: +3% reported, +1% and +2% linear extrapolation. Source: Mallplaza financial statements Jun-26, Note 26; Team analysis." },
    ],
  },
};

const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || 16, bold: o.bold,
  italics: o.italics, color: o.color || TINTA });
const sinBordes = { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } };
const cuerpo = { line: 228, lineRule: LineRuleType.AUTO, before: 0 };

function subtitulo(texto) {
  return new Paragraph({ spacing: { before: 110, after: 50 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ROJO, space: 1 } },
    children: [run(texto, { bold: true, color: AZUL })] });
}

function columnaTexto(L, pag2) {
  const out = [];
  for (const s of L.secciones.filter((x) => !!x.pag2 === pag2)) {
    out.push(subtitulo(s.titulo));
    for (const p of s.parrafos || [])
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { ...cuerpo, after: 40 },
        children: [run(p)] }));
    for (const r of s.riesgos || []) {
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { ...cuerpo, after: 14 },
        children: [run(`${r.cod}: `, { bold: true, color: ROJO }), run(r.nombre + " ", { bold: true }),
          run(r.texto + " "), run(`(${r.tag})`, { italics: true, color: GRIS })] }));
      out.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { ...cuerpo, after: 60 },
        children: [run(L.mitigantes, { bold: true, color: AZUL }), run(r.mit)] }));
    }
  }
  return out;
}

function columnaFiguras(L, pag2) {
  const out = [];
  L.figuras.filter((x) => !!x.pag2 === pag2).forEach((f, k) => {
    out.push(new Paragraph({ spacing: { before: k === 0 ? 40 : 100, after: 30, line: 228, lineRule: LineRuleType.AUTO },
      children: [run(f.n + " ", { bold: true, color: AZUL, size: 14 }), run(f.t, { color: AZUL, size: 14 })] }));
    out.push(new Paragraph({ spacing: { before: 0, after: 20 }, children: [
      new ImageRun({ type: "png", data: fs.readFileSync(path.join(FIG_DIR, f.img)),
        transformation: { width: Math.round(ANCHO_FIG * PULGADA_PX), height: Math.round(f.alto * PULGADA_PX) } })] }));
    out.push(new Paragraph({ spacing: { before: 0, after: 0, line: 216, lineRule: LineRuleType.AUTO },
      children: [run(f.fuente, { size: 11, color: GRIS })] }));
  });
  return out;
}

function cuadroPolitico(L) {
  const P = L.politica, anchos = [1300, 1700, 4250, 3416];
  const borde = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  const celda = (w, texto, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    borders: { top: borde, bottom: borde, left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
      right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" } },
    margins: { top: 40, bottom: 40, left: 70, right: 70 }, verticalAlign: VerticalAlign.TOP,
    children: texto.split("\n").map((t, k) => new Paragraph({ spacing: { before: 0, after: 0, line: 216, lineRule: LineRuleType.AUTO },
      alignment: o.just ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
      children: [run(t, { size: o.size || (k ? 13 : 14), bold: o.bold && !k, color: o.color || (k ? GRIS : TINTA) })] })) });
  return [
    subtitulo(P.titulo),
    new Table({ width: { size: anchos.reduce((a, b) => a + b), type: WidthType.DXA }, columnWidths: anchos,
      rows: [
        new TableRow({ tableHeader: true, children: P.cab.map((c, k) => celda(anchos[k], c, { bold: true, color: "FFFFFF", fill: AZUL })) }),
        ...P.filas.map((f, r) => new TableRow({ cantSplit: true, children: f.map((c, k) => celda(anchos[k], c,
          { bold: k < 2, just: k >= 2, fill: r % 2 ? FONDO : undefined })) })),
      ] }),
    new Paragraph({ spacing: { before: 50, after: 0, line: 216, lineRule: LineRuleType.AUTO }, alignment: AlignmentType.JUSTIFIED,
      children: [run(P.fuentes, { size: 11, color: GRIS })] }),
  ];
}

function documento(L) {
  const cel = (w, children, extra = {}) => new TableCell({ width: { size: w, type: WidthType.DXA },
    borders: sinBordes, margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP,
    children, ...extra });
  const dosColumnas = (pag2) => new Table({ width: { size: COL_TEXTO + COL_SEP + COL_FIG, type: WidthType.DXA },
    columnWidths: [COL_TEXTO, COL_SEP, COL_FIG],
    borders: { ...sinBordes, insideHorizontal: sinBordes.top, insideVertical: sinBordes.top },
    rows: [new TableRow({ children: [
      cel(COL_TEXTO, columnaTexto(L, pag2)),
      cel(COL_SEP, [new Paragraph({ children: [] })]),
      cel(COL_FIG, columnaFiguras(L, pag2)),
    ] })] });
  return new Document({
    styles: { default: { document: { run: { font: FUENTE, size: 16 } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 },
        margin: { top: 1000, right: 620, bottom: 900, left: 620, header: 450, footer: 400 },
        pageNumbers: { start: 9 } } },
      headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0", space: 2 } },
        children: [run("Mallplaza (Plaza S.A.)  |  CFA Institute Research Challenge", { size: 14, color: GRIS })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [run(L.pagina, { size: 14, color: GRIS }),
          new TextRun({ children: [PageNumber.CURRENT], font: FUENTE, size: 14, color: GRIS })] })] }) },
      children: [
        new Paragraph({ spacing: { after: 80 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 2 } },
          children: [run(L.titulo, { bold: true, color: AZUL, size: 30 })] }),
        new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { ...cuerpo, after: 80 },
          shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO },
          children: [run(L.intro)] }),
        dosColumnas(false),
        new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0 }, children: [] }),
        dosColumnas(true),
        ...cuadroPolitico(L),
      ],
    }],
  });
}

(async () => {
  for (const L of Object.values(LANG)) {
    const salida = path.join(__dirname, L.archivo);
    fs.writeFileSync(salida, await Packer.toBuffer(documento(L)));
    console.log(salida);
  }
})();
