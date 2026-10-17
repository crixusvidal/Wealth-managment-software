// Sección ESG del informe + anexo de gobierno corporativo y sostenibilidad (español e inglés).
// Figuras: graficos/informe/esg/{es,en}_*.png (python graficos/esg.py).
// Uso: node informe/esg.js
const fs = require("fs");
const path = require("path");
const {
  AlignmentType, BorderStyle, Document, Footer, Header, ImageRun, LineRuleType, Packer, PageBreak,
  PageNumber, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, VerticalAlign, WidthType,
} = require("docx");

const AZUL = "14284B", ROJO = "E9004B", TINTA = "1A1A1A", GRIS = "5F5F5F", FONDO = "EEF1F6";
const FUENTE = "Arial", ANCHO = 10666, COL_TEXTO = 6900, COL_SEP = 216, COL_FIG = 3550;
const MEDIA = 5225;
const FIG = path.join(__dirname, "..", "graficos", "figuras");
const PX = 96;
const ALTO = { E1: 2.3, E2: 1.75, E3: 1.8, E4: 1.7, E5: 1.7, E6: 2.5, AE1: 1.0, AE2: 1.15, AE3: 1.45, AE4: 1.75, AE5: 5.05 }; // = graficos/esg.py

const run = (text, o = {}) => new TextRun({ text, font: FUENTE, size: o.size || 16, bold: o.bold,
  italics: o.italics, color: o.color || TINTA });
const nada = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const sinBordes = { top: nada, bottom: nada, left: nada, right: nada };
const linea = (l = 228) => ({ line: l, lineRule: LineRuleType.AUTO });
function rico(texto, o = {}) {
  return texto.split(/(\*\*[^*]+\*\*)/).filter(Boolean).map((p) => p.startsWith("**")
    ? run(p.slice(2, -2), { ...o, bold: true, color: o.boldColor || AZUL }) : run(p, o));
}

const L = {
  es: {
    archivo: "Mallplaza_ESG_ES.docx", pag: "Página ", pagAnexo: "A-",
    titulo: "AMBIENTAL, SOCIAL Y GOBIERNO CORPORATIVO (ESG)",
    intro: "El enfoque ESG de Plaza es transversal a todas sus decisiones: busca desarrollar su modelo de negocios respetando los parámetros ESG y contribuir a un desarrollo sostenible, lo que además le abre beneficios financieros externos y de su propia operación. Bloomberg califica a la compañía con 3,39 puntos (Figura E1), bajo la media de la industria a nivel global (5–6 puntos) y de su competencia directa en Chile; esto contrasta con su perfil “verde” local, donde figura entre las principales empresas inmobiliarias (REM) de la región.",
    secciones: [
      { titulo: null, parrafos: [
        "**Beneficios financieros potenciales:** el sistema de gestión de la energía (SGE) certificado bajo la ISO 50001 mejora el desempeño energético y reduce los costos operativos. Los beneficios externos dependen de su imagen ESG: en Chile, la Ley I+D (CORFO) permite descontar del impuesto de primera categoría el 35% de lo invertido en investigación y desarrollo (hasta 52,55% del total, sumando el gasto); en Perú, el reconocimiento de “Huella Perú” (2 de 4 estrellas) facilita la obtención de permisos y licencias; y en Colombia, la Ley 2099 de 2021 permite descontar de la renta las inversiones en control, conservación y mejoramiento del medio ambiente."] },
      { titulo: "AMBIENTAL", parrafos: [
        "El enfoque de Plaza es una gestión operativa eficiente, sustentable y circular, con tres focos estratégicos: eficiencia energética e hídrica, huella de carbono y gestión de residuos. Su principal objetivo es la carbono neutralidad, con la meta de ser Net Zero en emisiones directas e indirectas al 2035 (65,6% de avance en 2025).",
        "**Índices de sostenibilidad:** la compañía está en el DJSI Chile y el DJSI MILA, y es la única empresa inmobiliaria (REM) de Sudamérica en el MSCI ESG, con clasificación “BBB”. Además, forma parte del S&P Global Sustainability Yearbook y Morningstar Sustainalytics le asigna un riesgo ESG de 14,4 (bajo) (Figura E1).",
        "**Energía:** la meta es reducir 22% la intensidad energética del gasto común al 2035 (35,5% de avance) y abastecerse con 100% de electricidad renovable al 2026 (89,3% de avance) (Figura E2). En 2025 la energía renovable representó el 90,9% del consumo total de 373.860 MWh (Figura E3), con una intensidad de 0,16 MWh/m² de GLA (0,14 renovable y 0,014 no renovable).",
        "**Agua:** la meta de reducir 30% el consumo directo de agua al 2025 se cumplió en 100,9%, apoyada por las certificaciones LEED y EDGE. En 2025 la intensidad fue de 0,77 m³/m² de GLA.",
        "**HCFC:** busca eliminar el 100% de su uso al 2030 (14,9% de avance), clave para la gestión de emisiones de gases de efecto invernadero, ámbito en el que Bloomberg reconoce a Plaza como líder de la industria global.",
        "**Residuos:** la meta de valorizar el 60% de los residuos al 2025 registra un avance de 94,8%; además, busca desviar de los vertederos el 75% de los residuos sólidos al 2027 y ya eliminó el 97% de los plásticos de un solo uso de la operación. En 2025 valorizó el 46,2% de sus residuos no peligrosos (Figura E3).",
        "**Gestión climática:** Plaza adopta las recomendaciones TCFD para integrar los riesgos físicos y de transición, realiza un due diligence ambiental en la adquisición de activos y opera bajo un programa de cumplimiento ambiental alineado con la ISO 14001, monitoreado por el comité de sustentabilidad."] },
      { titulo: "SOCIAL", parrafos: [
        "La compañía aborda el ámbito social con un modelo de relación ética de largo plazo, basado en la creación de valor compartido, que pone a sus grupos de interés en el centro de su estrategia, alineado con la ISO 45001 y las directrices de la OIT. En 2025 incorporó un programa de detección y respuesta ante riesgos de su política de derechos humanos. Gestiona a sus grupos de interés con mecanismos como diálogos comunitarios y evaluaciones de materialidad (Anexo [Y]) y atiende cualquier incidencia social a través de su canal de integridad. Cadem la ubicó en el puesto 31 de 365 “Marcas Ciudadanas” y ganadora de la categoría Centros Comerciales (2026). En 2025 sus programas y proyectos impactaron positivamente a 49.325 personas (Figura E4).",
        "**Colaboradores y proveedores:** capacita a sus trabajadores, ofrece programas para quienes se acercan a la jubilación y evalúa a sus proveedores con criterios de sostenibilidad (55,3% de los proveedores locales y 38,4% de los extranjeros).",
        "**Arrendatarios:** gestiona la relación con sus locatarios mediante un modelo de relacionamiento centrado en la propuesta de valor y la experiencia; el NPS pasó de −8% en 2022 a 13% en 2025 (Figura E4).",
        "**Seguridad y privacidad:** cada centro cuenta con manuales de emergencia adaptados a su contexto geográfico y en 2025 realizó 52 simulacros con la comunidad. Implementó la ISO/IEC 27001 para proteger la confidencialidad, integridad y disponibilidad de la información ante riesgos cibernéticos.",
        "**Diversidad:** las mujeres representan el 53% de la fuerza laboral (403 personas) y el 42% de los cargos directivos (44), con una brecha salarial media de 28% a nivel regional. La meta al 2028 es que el 60% de los centros esté certificado como espacio inclusivo, y Plaza obtuvo el “Sello Mayor”, que certifica su adaptación física y operacional para personas mayores."] },
      { titulo: "GOBIERNO CORPORATIVO", parrafos: [
        "El gobierno corporativo se rige por los principios del G20 y de la OCDE, alineado con la NCG 461 de la CMF, para asegurar una gestión transparente, con control de riesgos y una estructura sólida que dé seguimiento a la estrategia y a los compromisos ESG.",
        "**Directorio y comités:** el Directorio, de 9 miembros, define los objetivos, estrategias y políticas de Plaza, pero solo uno es independiente (Figura E5). El comité de directores, de tres miembros y presidido por el director independiente, revisa las operaciones con partes relacionadas (Ley 18.046). Falabella controla el 53,10% de la propiedad y los fondos de pensiones suman 23,8% (Anexo [Y]). La baja independencia del Directorio es el principal punto débil de gobierno (RR1, Investment Risks).",
        "**Incentivos ESG:** los KPIs ESG se miden anualmente y forman parte de la remuneración variable del Gerente General, gerentes y subgerentes. Los indicadores no financieros pesan 61% (33,3% vinculado directamente a materias ESG y 27,7% a otros temas) y los financieros 39% (Figura E5).",
        "**Ética y cumplimiento:** en 2025 no registró incumplimientos de la legislación ni casos de corrupción. Registró 27 incumplimientos de derechos del consumidor (Ley 19.496), 2 de derechos laborales, 1 procedimiento de la Superintendencia del Medio Ambiente y 8 denuncias de acoso o violencia en el trabajo (Anexo [Y])."] },
    ],
    figuras: [
      ["Figura E1:", "Puntaje ESG de Bloomberg y ratings externos", "E1_scorecard", "Fuente: Bloomberg (puntaje ESG); MSCI, Morningstar Sustainalytics, S&P Global, Huella Perú; Memoria Mallplaza 2025."],
      ["Figura E2:", "Metas ambientales: avance a 2025", "E2_metas", "Fuente: Memoria Mallplaza 2025."],
      ["Figura E3:", "Energía renovable y valorización de residuos, 2025", "E3_energia_residuos", "Valorización: porcentaje sobre los residuos no peligrosos; el desglose suma 100% de lo valorizado. Fuente: Memoria Mallplaza 2025."],
      ["Figura E4:", "Indicadores sociales", "E4_social", "NPS = % de promotores (notas 9–10) menos % de detractores (0–6). Fuente: Memoria Mallplaza 2025."],
      ["Figura E5:", "Incentivos ESG y composición del Directorio", "E5_gobierno", "Fuente: Memoria Mallplaza 2025."],
    ],
    anexo: {
      titulo: "ANEXO [Y]: GOBIERNO CORPORATIVO Y SOSTENIBILIDAD",
      lead: "Información de respaldo de la sección ESG: estructura de propiedad, Directorio, cumplimiento, modelo de relacionamiento con grupos de interés, cultura organizacional y residuos. Fuente principal: Memoria Mallplaza 2025.",
      s1: "1. ESTRUCTURA DE PROPIEDAD (31-12-2025)",
      f1: ["Figura S1:", "Participación por tipo de accionista", "Fuente: Memoria Mallplaza 2025 (NCG 461 2.3.3). No hay ADRs ni certificados emitidos en el extranjero."],
      t1: { cols: ["Diez mayores accionistas", "Acciones", "% de acciones"], anchos: [6266, 2200, 2200], filas: [
        ["Desarrollos Inmobiliarios SpA (Falabella)", "1.161.849.244", "53,10%"], ["Inversiones Avenida Borgoño Limitada", "197.439.955", "9,02%"],
        ["Banchile Corredores de Bolsa S.A.", "59.447.149", "2,71%"], ["BCI Corredor de Bolsa S.A.", "53.553.765", "2,45%"],
        ["AFP Habitat S.A. para Fondo de Pensión C", "50.733.976", "2,32%"], ["Bolsa de Comercio de Santiago, Bolsa de Valores", "41.532.310", "1,90%"],
        ["AFP Provida S.A. para Fondo de Pensión C", "37.746.440", "1,72%"], ["Banco Santander por cuenta de inversión extranjera", "35.594.345", "1,63%"],
        ["AFP Capital S.A. Fondo de Pensión tipo C", "34.539.038", "1,58%"], ["AFP Habitat S.A. Fondo de Pensión tipo B", "30.039.097", "1,37%"]] },
      s2: "2. DIRECTORIO (ELEGIDO EL 22-04-2025)",
      f2: ["Figura S2:", "Miembros del Directorio (rojo: director independiente; azul: presidente)"],
      t2: { cols: ["Director", "Profesión", "Cargo"], anchos: [3600, 2700, 4366], filas: [
        ["Sergio Cardone Solari", "Ingeniero Comercial", "Presidente"], ["Paul Fürst Gwinner", "Ingeniero Comercial", "Director"],
        ["Carlo Solari Donaggio", "Ingeniero Civil", "Director"], ["Alejandro Puentes Bruno", "Ingeniero Comercial", "Director"],
        ["Pablo Eyzaguirre Court", "Ingeniero Civil", "Director"], ["Dafne González Lizama", "Abogada", "Directora"],
        ["Enrique Ostalé Cambiaso", "Ingeniero Comercial", "Director"], ["Ian Fürst Garmendia", "Ingeniero Comercial", "Director"],
        ["Manuel Ovalle Edwards", "Ingeniero Civil", "Director independiente y presidente del Comité de Directores"]] },
      n2: "9 miembros titulares, sin suplentes; un director independiente. Cada director tiene como máximo cuatro cargos en otras sociedades que cotizan en bolsa. Tomás Müller, Manuela Sánchez y Juan Pablo Montero fueron directores hasta el 22-04-2025. Fuente: Memoria Mallplaza 2025.",
      s3: "3. ÉTICA Y CUMPLIMIENTO, 2025",
      t3: { cols: ["Materia", "Casos"], anchos: [8666, 2000], filas: [
        ["Corrupción", "0"],
        ["Incumplimientos de derechos de los clientes (Ley 19.496 de Protección de los Derechos del Consumidor)", "27"],
        ["Incumplimientos de derechos de los trabajadores", "2"], ["Procedimientos de la Superintendencia del Medio Ambiente", "1"],
        ["Denuncias de acoso laboral, sexual y de violencia en el trabajo", "8"]] },
      n3: "Fuente: Memoria Mallplaza 2025.",
      s4: "4. MODELO DE RELACIONAMIENTO CON GRUPOS DE INTERÉS",
      f4: ["Figura S3:", "Proceso de implementación del modelo de relacionamiento", "Fuente: Memoria Mallplaza 2025."],
      s5: "5. CULTURA ORGANIZACIONAL: ENCUESTA DE CULTURA 2025",
      kpis: [["75%", "Engagement", "−2 pp vs 2024 · meta 2025: 80%"], ["37", "eNPS", "−10 vs 2024"], ["88%", "Liderazgo", "+2 pp vs 2024"], ["93%", "Participación", "+3 pp vs 2024"]],
      n5: "El liderazgo (88%) se consolida como el principal activo para apalancar la estrategia. Fuente: Memoria Mallplaza 2025.",
      s6: "6. RESIDUOS, 2025",
      f6: ["Figura S4:", "Generación de residuos por país", "Fuente: Memoria Mallplaza 2025."],
      r6: [
        "**32.424,86 toneladas** de residuos a nivel regional: 76% en Chile, 16% en Colombia y 7% en Perú.",
        "**46,2% valorizado:** 14.987,57 t de residuos no peligrosos se destinaron a valorización energética con biogás (40,4%), reciclaje (36,1%) y compostaje (23,5%); el resto se destinó a eliminación, principalmente rellenos sanitarios.",
        "**Residuos peligrosos:** 9,03 t, una cifra marginal por la naturaleza de la operación.",
      ],
      s7: "7. MATRIZ DE RIESGO ESG",
      f7: ["Figura S5:", "Matriz de factores ESG y su efecto en el caso de inversión", "Materialidad, horizonte, capacidad de absorber pérdidas y efecto neto son juicios del equipo. Marco adaptado del enfoque de S&P Global Ratings para factores ESG de crédito. Fuente: Memoria Mallplaza 2025, Bloomberg, MSCI, Sustainalytics, S&P Global; análisis del equipo."],
      s8: "8. TRACK RECORD DEL EQUIPO EJECUTIVO",
      t8: { cols: ["Nombre", "Cargo en PLAZA", "Antigüedad", "Experiencia previa", "Perfil y responsabilidades"], anchos: [1900, 2300, 1000, 2000, 3466], filas: [
        ["Pablo Pulido Sierra", "Gerente General (desde 2026)", "10 años", "Grupo Casino, Grupo Éxito", "Perfil comercial; experiencia en integraciones regionales."],
        ["Fernando de Peña Iver", "Gerente General (hasta 2025)", "35 años", "Trayectoria en Mallplaza", "Conduce la compañía hasta 2025; amplia trayectoria en el retail latinoamericano."],
        ["Derek Schweitzer Tang", "Gerente de Administración y Finanzas (CFO)", "Desde 2023", "BRMalls", "Experiencia internacional; su área lidera la estructura financiera y el ingreso a índices globales (MSCI)."],
        ["Carlos Honorato C.", "CTO, Gerente de Tecnología", "2 años", "Ripley.com, Neoris, Alto Movup, ProChile", "Responsable de la modernización de las ventas omnicanal y de la ciberseguridad corporativa."],
        ["María Sol de Cabo", "Gerente de Personas y Cultura", "Desde 2025", "Betterfly", "Más de 20 años de trayectoria internacional en gestión de talento."],
        ["Oscar Munizaga Delfín", "Gerente de Desarrollo y Negocios", "34 años", "Trayectoria en Mallplaza", "Responsable del crecimiento mediante M&A y desarrollos greenfield."],
        ["Luis Hernán Silva", "Gerente Legal, Riesgos y Cumplimiento", "27 años", "Trayectoria en Mallplaza", "Responsable legal corporativo y de la gestión de riesgos y cumplimiento."]] },
      n8: "Fuente: información pública de la compañía; Punteo del equipo (track record).",
    },
  },
  en: {
    archivo: "Mallplaza_ESG_EN.docx", pag: "Page ", pagAnexo: "A-",
    titulo: "ENVIRONMENTAL, SOCIAL AND GOVERNANCE (ESG)",
    intro: "Plaza's ESG approach cuts across all of its decisions: it seeks to grow its business model within ESG parameters and contribute to sustainable development, which also opens up external financial benefits and savings in its own operations. Bloomberg rates the company at 3.39 points (Figure E1), below the global industry average (5–6 points) and its direct competitors in Chile; this contrasts with its local “green” profile, where it ranks among the region's leading real estate management (REM) companies.",
    secciones: [
      { titulo: null, parrafos: [
        "**Potential financial benefits:** its ISO 50001-certified energy management system improves energy performance and lowers operating costs. External benefits depend on its ESG standing: in Chile, the R&D Law (CORFO) allows 35% of R&D spending to be credited against first-category income tax (up to 52.55% in total including the expense deduction); in Peru, the “Huella Perú” recognition (2 of 4 stars) eases permits and licences; and in Colombia, Law 2099 of 2021 allows investments in environmental control, conservation and improvement to be deducted from income tax."] },
      { titulo: "ENVIRONMENTAL", parrafos: [
        "Plaza aims for efficient, sustainable and circular operations, with three strategic focus areas: energy and water efficiency, carbon footprint and waste management. Its main goal is carbon neutrality, with a target of Net Zero direct and indirect emissions by 2035 (65.6% progress in 2025).",
        "**Sustainability indices:** the company is included in the DJSI Chile and DJSI MILA, and is the only South American real estate management (REM) company in the MSCI ESG indices, rated “BBB”. It is also part of the S&P Global Sustainability Yearbook, and Morningstar Sustainalytics assigns it an ESG risk score of 14.4 (low) (Figure E1).",
        "**Energy:** the targets are a 22% cut in common-area energy intensity by 2035 (35.5% progress) and 100% renewable electricity by 2026 (89.3% progress) (Figure E2). In 2025 renewables supplied 90.9% of total consumption of 373,860 MWh (Figure E3), an intensity of 0.16 MWh/m² of GLA (0.14 renewable and 0.014 non-renewable).",
        "**Water:** the target of a 30% cut in direct water consumption by 2025 was met at 100.9%, supported by LEED and EDGE certifications. In 2025 intensity was 0.77 m³/m² of GLA.",
        "**HCFC:** it plans to phase out 100% of HCFC use by 2030 (14.9% progress), key to managing greenhouse gas emissions, an area where Bloomberg recognises Plaza as a global industry leader.",
        "**Waste:** the target of recovering 60% of waste by 2025 shows 94.8% progress; it also aims to divert 75% of solid waste from landfill by 2027 and has already eliminated 97% of single-use plastics from its operations. In 2025 it recovered 46.2% of its non-hazardous waste (Figure E3).",
        "**Climate management:** Plaza follows the TCFD recommendations to integrate physical and transition risks, performs environmental due diligence on asset acquisitions and runs an environmental compliance programme aligned with ISO 14001, overseen by the sustainability committee."] },
      { titulo: "SOCIAL", parrafos: [
        "The company approaches social matters through a model of long-term ethical relationships based on shared value creation, placing its stakeholders at the centre of its strategy, in line with ISO 45001 and ILO guidelines. In 2025 it added a programme to detect and respond to risks under its human rights policy. It engages stakeholders through community dialogues and materiality assessments (Appendix [Y]) and handles any social incident through its integrity channel. Cadem ranked it 31st of 365 “Citizen Brands” and winner of the Shopping Centres category (2026). In 2025 its programmes and projects positively reached 49,325 people (Figure E4).",
        "**Employees and suppliers:** it trains its staff, offers programmes for employees nearing retirement and assesses suppliers on sustainability criteria (55.3% of local suppliers and 38.4% of foreign ones).",
        "**Tenants:** it manages tenant relationships through an engagement model focused on value proposition and experience; tenant NPS rose from −8% in 2022 to 13% in 2025 (Figure E4).",
        "**Safety and privacy:** every centre has emergency manuals adapted to its local context, and 52 community drills were held in 2025. It implemented ISO/IEC 27001 to protect the confidentiality, integrity and availability of information against cyber risks.",
        "**Diversity:** women make up 53% of the workforce (403 people) and 42% of management positions (44), with an average gender pay gap of 28% across the region. The 2028 target is for 60% of centres to be certified as inclusive spaces, and Plaza earned the “Sello Mayor” seal, which certifies its physical and operational adaptation for older people."] },
      { titulo: "GOVERNANCE", parrafos: [
        "Corporate governance follows the G20/OECD principles and the CMF's General Rule 461, to ensure transparent management, risk control and a solid structure to monitor the strategy and ESG commitments.",
        "**Board and committees:** the 9-member Board sets Plaza's objectives, strategies and policies, but only one director is independent (Figure E5). The three-member directors' committee, chaired by the independent director, reviews related-party transactions (Law 18,046). Falabella controls 53.10% of the shares and pension funds hold 23.8% (Appendix [Y]). Low board independence is the main governance weakness (RR1, Investment Risks).",
        "**ESG incentives:** ESG KPIs are measured annually and are part of the variable pay of the CEO, managers and deputy managers. Non-financial indicators weigh 61% (33.3% linked directly to ESG matters and 27.7% to other topics) and financial indicators 39% (Figure E5).",
        "**Ethics and compliance:** in 2025 there were no breaches of laws and regulations and no corruption cases. There were 27 breaches of consumer rights (Law 19,496), 2 of labour rights, 1 proceeding by the Environment Superintendency and 8 complaints of workplace harassment or violence (Appendix [Y])."] },
    ],
    figuras: [
      ["Figure E1:", "Bloomberg ESG score and external ratings", "E1_scorecard", "Source: Bloomberg (ESG score); MSCI, Morningstar Sustainalytics, S&P Global, Huella Perú; Mallplaza Annual Report 2025."],
      ["Figure E2:", "Environmental targets: progress to 2025", "E2_metas", "Source: Mallplaza Annual Report 2025."],
      ["Figure E3:", "Renewable energy and waste recovery, 2025", "E3_energia_residuos", "Recovery: share of non-hazardous waste; the breakdown adds up to 100% of recovered waste. Source: Mallplaza Annual Report 2025."],
      ["Figure E4:", "Social indicators", "E4_social", "NPS = % promoters (scores 9–10) minus % detractors (0–6). Source: Mallplaza Annual Report 2025."],
      ["Figure E5:", "ESG incentives and board composition", "E5_gobierno", "Source: Mallplaza Annual Report 2025."],
    ],
    anexo: {
      titulo: "APPENDIX [Y]: CORPORATE GOVERNANCE AND SUSTAINABILITY",
      lead: "Supporting information for the ESG section: ownership structure, Board of Directors, compliance, stakeholder engagement model, organisational culture and waste. Main source: Mallplaza Annual Report 2025.",
      s1: "1. OWNERSHIP STRUCTURE (31-DEC-2025)",
      f1: ["Figure S1:", "Shareholding by type of shareholder", "Source: Mallplaza Annual Report 2025 (NCG 461 2.3.3). No ADRs or certificates issued abroad."],
      t1: { cols: ["Ten largest shareholders", "Shares", "% of shares"], anchos: [6266, 2200, 2200], filas: [
        ["Desarrollos Inmobiliarios SpA (Falabella)", "1,161,849,244", "53.10%"], ["Inversiones Avenida Borgoño Limitada", "197,439,955", "9.02%"],
        ["Banchile Corredores de Bolsa S.A.", "59,447,149", "2.71%"], ["BCI Corredor de Bolsa S.A.", "53,553,765", "2.45%"],
        ["AFP Habitat S.A. Pension Fund C", "50,733,976", "2.32%"], ["Santiago Stock Exchange", "41,532,310", "1.90%"],
        ["AFP Provida S.A. Pension Fund C", "37,746,440", "1.72%"], ["Banco Santander on behalf of foreign investors", "35,594,345", "1.63%"],
        ["AFP Capital S.A. Pension Fund C", "34,539,038", "1.58%"], ["AFP Habitat S.A. Pension Fund B", "30,039,097", "1.37%"]] },
      s2: "2. BOARD OF DIRECTORS (ELECTED 22-APR-2025)",
      f2: ["Figure S2:", "Board members (red: independent director; blue: chairman)"],
      t2: { cols: ["Director", "Profession", "Position"], anchos: [3600, 2700, 4366], filas: [
        ["Sergio Cardone Solari", "Business engineer", "Chairman"], ["Paul Fürst Gwinner", "Business engineer", "Director"],
        ["Carlo Solari Donaggio", "Civil engineer", "Director"], ["Alejandro Puentes Bruno", "Business engineer", "Director"],
        ["Pablo Eyzaguirre Court", "Civil engineer", "Director"], ["Dafne González Lizama", "Lawyer", "Director"],
        ["Enrique Ostalé Cambiaso", "Business engineer", "Director"], ["Ian Fürst Garmendia", "Business engineer", "Director"],
        ["Manuel Ovalle Edwards", "Civil engineer", "Independent director and chair of the Directors' Committee"]] },
      n2: "9 full members, no alternates; one independent director. Each director holds at most four positions in other listed companies. Tomás Müller, Manuela Sánchez and Juan Pablo Montero were directors until 22-Apr-2025. Source: Mallplaza Annual Report 2025.",
      s3: "3. ETHICS AND COMPLIANCE, 2025",
      t3: { cols: ["Matter", "Cases"], anchos: [8666, 2000], filas: [
        ["Corruption", "0"],
        ["Breaches of customer rights (Consumer Protection Law 19,496)", "27"],
        ["Breaches of employee rights", "2"], ["Proceedings by the Environment Superintendency", "1"],
        ["Complaints of workplace, sexual harassment and violence at work", "8"]] },
      n3: "Source: Mallplaza Annual Report 2025.",
      s4: "4. STAKEHOLDER ENGAGEMENT MODEL",
      f4: ["Figure S3:", "Implementation process of the engagement model", "Source: Mallplaza Annual Report 2025."],
      s5: "5. ORGANISATIONAL CULTURE: 2025 CULTURE SURVEY",
      kpis: [["75%", "Engagement", "−2 pp vs 2024 · 2025 target: 80%"], ["37", "eNPS", "−10 vs 2024"], ["88%", "Leadership", "+2 pp vs 2024"], ["93%", "Participation", "+3 pp vs 2024"]],
      n5: "Leadership (88%) is the main asset to leverage the strategy. Source: Mallplaza Annual Report 2025.",
      s6: "6. WASTE, 2025",
      f6: ["Figure S4:", "Waste generation by country", "Source: Mallplaza Annual Report 2025."],
      r6: [
        "**32,424.86 tonnes** of waste across the region: 76% in Chile, 16% in Colombia and 7% in Peru.",
        "**46.2% recovered:** 14,987.57 t of non-hazardous waste went to energy recovery through biogas (40.4%), recycling (36.1%) and composting (23.5%); the rest went to disposal, mainly sanitary landfills.",
        "**Hazardous waste:** 9.03 t, a marginal amount given the nature of the business.",
      ],
      s7: "7. ESG RISK MATRIX",
      f7: ["Figure S5:", "ESG factor matrix and its effect on the investment case", "Materiality, time horizon, ability to absorb losses and net effect are team assessments. Framework adapted from S&P Global Ratings' approach to ESG credit factors. Source: Mallplaza Annual Report 2025, Bloomberg, MSCI, Sustainalytics, S&P Global; Team analysis."],
      s8: "8. EXECUTIVE TEAM TRACK RECORD",
      t8: { cols: ["Name", "Role at PLAZA", "Tenure", "Previous experience", "Profile and responsibilities"], anchos: [1900, 2300, 1000, 2000, 3466], filas: [
        ["Pablo Pulido Sierra", "CEO (since 2026)", "10 years", "Grupo Casino, Grupo Éxito", "Commercial profile; experience in regional integrations."],
        ["Fernando de Peña Iver", "CEO (until 2025)", "35 years", "Career at Mallplaza", "Led the company until 2025; extensive career in Latin American retail."],
        ["Derek Schweitzer Tang", "Chief Administration and Finance Officer (CFO)", "Since 2023", "BRMalls", "International experience; his area leads the financial structure and the entry into global indices (MSCI)."],
        ["Carlos Honorato C.", "CTO, Chief Technology Officer", "2 years", "Ripley.com, Neoris, Alto Movup, ProChile", "Responsible for modernizing omnichannel sales and corporate cybersecurity."],
        ["María Sol de Cabo", "Chief People and Culture Officer", "Since 2025", "Betterfly", "Over 20 years of international experience in talent management."],
        ["Oscar Munizaga Delfín", "Chief Development and Business Officer", "34 years", "Career at Mallplaza", "Responsible for growth through M&A and greenfield developments."],
        ["Luis Hernán Silva", "Chief Legal, Risk and Compliance Officer", "27 years", "Career at Mallplaza", "Responsible for corporate legal affairs and risk and compliance management."]] },
      n8: "Source: company public information; Team notes (track record).",
    },
  },
};

const subtitulo = (t, antes = 110) => new Paragraph({ spacing: { before: antes, after: 50 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ROJO, space: 1 } },
  children: [run(t, { bold: true, color: AZUL })] });
const parrafo = (t, despues = 60) => new Paragraph({ alignment: AlignmentType.JUSTIFIED,
  spacing: { before: 0, after: despues, ...linea(228) }, children: rico(t) });
const caption = (n, t, antes = 100) => new Paragraph({ spacing: { before: antes, after: 30, ...linea() },
  children: [run(n + " ", { bold: true, color: AZUL, size: 14 }), run(t, { color: AZUL, size: 14 })] });
const fuente = (t) => new Paragraph({ spacing: { before: 10, after: 0, ...linea(216) }, children: [run(t, { size: 11, color: GRIS })] });
const imagen = (archivo, anchoIn, altoIn) => new Paragraph({ spacing: { before: 0, after: 20, line: 240, lineRule: LineRuleType.AUTO }, children: [
  new ImageRun({ type: "png", data: fs.readFileSync(path.join(FIG, archivo)),
    transformation: { width: Math.round(anchoIn * PX), height: Math.round(altoIn * PX) } })] });
const celda = (w, children, extra = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: sinBordes,
  margins: { top: 0, bottom: 0, left: 0, right: 0 }, verticalAlign: VerticalAlign.TOP, children, ...extra });
const vacia = () => new Paragraph({ children: [] });
const titulo = (t) => new Paragraph({ spacing: { after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: AZUL, space: 2 } },
  children: [run(t, { bold: true, color: AZUL, size: 30 })] });
const lead = (t) => new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { after: 80, ...linea() },
  shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO }, children: [run(t)] });

function bloque(lang, secciones, figuras, primero) {
  const izq = [];
  secciones.forEach((sec) => {
    if (sec.titulo) izq.push(subtitulo(sec.titulo, izq.length ? 110 : 20));
    sec.parrafos.forEach((p) => izq.push(parrafo(p)));
  });
  const der = [];
  figuras.forEach(([n, t, archivo, f], k) => {
    der.push(caption(n, t, k === 0 ? 20 : 110));
    der.push(imagen(`${lang}_${archivo}.png`, COL_FIG / 1440, ALTO[archivo.split("_")[0]]));
    der.push(fuente(f));
  });
  return new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [COL_TEXTO, COL_SEP, COL_FIG],
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
    rows: [new TableRow({ cantSplit: true, children: [celda(COL_TEXTO, izq), celda(COL_SEP, [vacia()]), celda(COL_FIG, der)] })] });
}

// Página 1: ESG, Ambiental y Social (Figuras E1–E4); página 2: Gobierno corporativo (Figura E5)
function seccionInforme(lang) {
  const X = L[lang];
  return [titulo(X.titulo), lead(X.intro),
    bloque(lang, X.secciones.slice(0, 3), X.figuras.slice(0, 4)),
    new Paragraph({ pageBreakBefore: true, spacing: { after: 0 }, children: [] }),
    bloque(lang, X.secciones.slice(3), X.figuras.slice(4))];
}

function tabla({ cols, anchos, filas }, alinearDer = []) {
  const borde = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0" };
  const enc = new TableRow({ tableHeader: true, children: cols.map((c, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: AZUL }, borders: sinBordes, margins: { top: 40, bottom: 40, left: 90, right: 90 },
    children: [new Paragraph({ alignment: alinearDer.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [run(c, { bold: true, size: 14, color: "FFFFFF" })] })] })) });
  const cuerpo = filas.map((f, r) => new TableRow({ children: f.map((v, i) => new TableCell({ width: { size: anchos[i], type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, color: "auto", fill: r % 2 ? "FFFFFF" : "F7F8FB" }, borders: { ...sinBordes, bottom: borde },
    margins: { top: 35, bottom: 35, left: 90, right: 90 }, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing: linea(216), alignment: alinearDer.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [run(v, { size: 14, bold: i === 0 && r === 0, color: TINTA })] })] })) }));
  return new Table({ width: { size: anchos.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: anchos,
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows: [enc, ...cuerpo] });
}

function kpis(lista) {
  const sep = 120, n = lista.length, w = Math.floor((ANCHO - sep * (n - 1)) / n);
  const cols = [];
  lista.forEach((_, i) => { cols.push(i === n - 1 ? ANCHO - (w + sep) * (n - 1) : w); if (i < n - 1) cols.push(sep); });
  const celdas = [];
  lista.forEach(([v, e, nota], i) => {
    celdas.push(new TableCell({ width: { size: cols[celdas.length], type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, color: "auto", fill: FONDO },
      borders: { ...sinBordes, top: { style: BorderStyle.SINGLE, size: 18, color: i === 0 ? ROJO : AZUL } },
      margins: { top: 70, bottom: 70, left: 120, right: 90 }, children: [
        new Paragraph({ spacing: { after: 10 }, children: [run(v, { bold: true, size: 30, color: i === 0 ? ROJO : AZUL })] }),
        new Paragraph({ spacing: { after: 10 }, children: [run(e, { bold: true, size: 14, color: AZUL })] }),
        new Paragraph({ children: [run(nota, { size: 12, color: GRIS })] })] }));
    if (i < n - 1) celdas.push(new TableCell({ width: { size: sep, type: WidthType.DXA }, borders: sinBordes, children: [vacia()] }));
  });
  return new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: cols,
    borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada }, rows: [new TableRow({ children: celdas })] });
}

function anexo(lang) {
  const A = L[lang].anexo;
  return [
    titulo(A.titulo), lead(A.lead),
    subtitulo(A.s1, 60), caption(A.f1[0], A.f1[1], 20), imagen(`${lang}_AE1_propiedad.png`, ANCHO / 1440, ALTO.AE1), fuente(A.f1[2]),
    new Paragraph({ spacing: { before: 80 }, children: [] }), tabla(A.t1, [1, 2]),
    subtitulo(A.s2, 160), caption(A.f2[0], A.f2[1], 20),
    imagen(`${lang}_AE4_directorio.png`, ANCHO / 1440, ALTO.AE4), fuente(A.n2),
    subtitulo(A.s3, 160), tabla(A.t3, [1]), fuente(A.n3),
    subtitulo(A.s4, 160), caption(A.f4[0], A.f4[1], 20), imagen(`${lang}_AE2_relacionamiento.png`, ANCHO / 1440, ALTO.AE2), fuente(A.f4[2]),
    subtitulo(A.s5, 160), kpis(A.kpis), fuente(A.n5),
    subtitulo(A.s6, 160),
    new Table({ width: { size: ANCHO, type: WidthType.DXA }, columnWidths: [MEDIA, COL_SEP, MEDIA],
      borders: { ...sinBordes, insideHorizontal: nada, insideVertical: nada },
      rows: [new TableRow({ children: [
        celda(MEDIA, [caption(A.f6[0], A.f6[1], 20), imagen(`${lang}_AE3_residuos_pais.png`, MEDIA / 1440, ALTO.AE3), fuente(A.f6[2])]),
        celda(COL_SEP, [vacia()]),
        celda(MEDIA, [new Paragraph({ spacing: { before: 200 }, children: [] }), ...A.r6.map((t) => parrafo(t, 80))]),
      ] })] }),
    subtitulo(A.s7, 160), caption(A.f7[0], A.f7[1], 20), imagen(`${lang}_AE5_matriz_esg.png`, ANCHO / 1440, ALTO.AE5), fuente(A.f7[2]),
    subtitulo(A.s8, 160), tabla(A.t8), fuente(A.n8),
  ];
}

function encabezado() {
  return new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D5D9E0", space: 2 } },
    children: [run("Mallplaza (Plaza S.A.)  |  CFA Institute Research Challenge", { size: 14, color: GRIS })] })] });
}
function pie(prefijo) {
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
    run(prefijo, { size: 14, color: GRIS }), new TextRun({ children: [PageNumber.CURRENT], font: FUENTE, size: 14, color: GRIS })] })] });
}

// Cuerpo en español: texto del borrador del equipo (borrador/CFA_ESG_PLAZA_beta_4_1.docx → borrador/texto.json,
// con comun/borrador_a_json.py), sin cambios, en el formato principal y con los gráficos del repo.
// Inglés: borrador/texto_en.json (traducción fiel del mismo texto).
const FIGS_CUERPO = ["E1_scorecard", "E2_metas", "E3_energia_residuos", "E4_social", "E5_gobierno", "E6_ahorro_esg"];
const imagenesDe = (lang) => Object.fromEntries(FIGS_CUERPO.map((n, i) => [`image${i + 1}.png`, [path.join(FIG, `${lang}_${n}.png`), ALTO[n.split("_")[0]]]]));
const textoDe = (lang) => JSON.parse(fs.readFileSync(path.join(__dirname, "borrador", lang === "es" ? "texto.json" : "texto_en.json"), "utf8"));
function cuerpo(lang) {
  const BORRADOR = path.join(__dirname, "borrador", lang === "es" ? "texto.json" : "texto_en.json");
  if (!fs.existsSync(BORRADOR)) return seccionInforme(lang);
  const { seccionBorrador } = require("../../comun/formato");
  const imagenes = imagenesDe(lang);
  return seccionBorrador(JSON.parse(fs.readFileSync(BORRADOR, "utf8")), { imagenes, soloCuerpo: true }).cuerpo;
}

function documento(lang) {
  const X = L[lang];
  const pagina = { size: { width: 11906, height: 16838 }, margin: { top: 1000, right: 620, bottom: 900, left: 620, header: 450, footer: 400 } };
  return ({
    styles: { default: { document: { run: { font: FUENTE, size: 16 }, paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: "auto" } } } } },
    sections: [
      { properties: { page: pagina }, headers: { default: encabezado() }, footers: { default: pie(X.pag) }, children: cuerpo(lang) },
      { properties: { page: { ...pagina, pageNumbers: { start: 1 } } }, headers: { default: encabezado() }, footers: { default: pie(X.pagAnexo) }, children: anexo(lang) },
    ],
  });
}

// Documento de la sección → ../documento · anexo → ../anexo (archivos separados, mismo formato)
module.exports = { documento, L, imagenesDe, textoDe };

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
