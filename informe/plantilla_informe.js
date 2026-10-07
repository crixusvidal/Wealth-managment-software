// Mallplaza — CFA Institute Research Challenge: report format template (A4).
// 10 body pages (sections in the team's order) + 10 appendix pages. Format only: every
// text area is a Word placeholder and every figure slot is an empty box, except the
// Investment Risks page, which carries its four charts.
// Usage: node informe/plantilla_informe.js
//   → informe/Mallplaza_CFA_Report_template.docx and informe/Investment_Risks_template.docx
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, Header, Footer,
  PageNumber, WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, HeightRule,
  LineRuleType,
} = require("docx");

const NAVY = "14284B";
const RED = "E9004B";
const GREY = "5F5F5F";
const PLACEHOLDER = "808081"; // marker color: post-processed into Word placeholders
const BOX = "F2F4F7";
const RULE = "D5D9E0";
const FONT = "Arial";
const G = path.join(__dirname, "..", "graficos", "informe");

const PAGE = { width: 11906, height: 16838 }; // A4
const MARGIN = { top: 1000, bottom: 900, left: 620, right: 620, header: 450, footer: 400 };
const CONTENT_W = PAGE.width - MARGIN.left - MARGIN.right;   // 10666
const RIGHT = 3550, GUTTER = 216, LEFT = CONTENT_W - RIGHT - GUTTER;  // ~65% / ~33%
const ROW_H = 14100;          // reserves the rest of the page under the page title
const ROW_H_CONT = 14560;     // continuation page: no title, so the columns get its space
const ROW_H_SUMMARY = 13650;  // same, on pages that also carry a summary bar
const BOX_H = 4300;           // empty figure box: three per page
const HALF_H = 6550;          // half-page block (title + columns) on a shared page
const HALF_BOX_H = 5200;      // one figure box in a half-page block
const DXA_PER_PX = 96 / 1440;

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none,
                    insideHorizontal: none, insideVertical: none };

let figureNo = 0;

const run = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 16, bold: o.bold,
                                            color: o.color, italics: o.italics });
const ph = (text, o = {}) => run(text, { ...o, color: PLACEHOLDER, italics: true });

function body(children, o = {}) {
  return new Paragraph({ children, spacing: { before: 0, after: o.after ?? 40, line: 228 },
                         alignment: o.align || AlignmentType.JUSTIFIED });
}

function pageTitle(children, noBreak) {
  return new Paragraph({
    children, pageBreakBefore: !noBreak, spacing: { after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: NAVY, space: 2 } },
  });
}

function sectionHeader(text) {
  return new Paragraph({
    children: [run(text.toUpperCase(), { bold: true, color: NAVY })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: RED, space: 1 } },
    spacing: { before: 100, after: 50 },
  });
}

function summaryBar(hint) {
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "EEF1F6" },
    spacing: { after: 80, line: 228 }, children: [ph(hint)],
  });
}

const cell = (children, width, extra = {}) => new TableCell({
  children, width: { size: width, type: WidthType.DXA }, borders: noBorders,
  margins: { top: 0, bottom: 0, left: 0, right: 0 }, ...extra });

// Text column (~65%) and figure column (~33%); figuresLeft swaps their sides.
function twoColumns(text, figures, height, figuresLeft = false) {
  const t = cell(text, LEFT, { verticalAlign: VerticalAlign.TOP });
  const f = cell(figures, RIGHT, { verticalAlign: VerticalAlign.TOP });
  const gap = cell([new Paragraph("")], GUTTER);
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: figuresLeft ? [RIGHT, GUTTER, LEFT] : [LEFT, GUTTER, RIGHT],
    borders: noBorders,
    rows: [new TableRow({ height: { value: height, rule: HeightRule.ATLEAST },
                          children: figuresLeft ? [f, gap, t] : [t, gap, f] })],
  });
}

function caption(n, children) {
  return new Paragraph({ spacing: { before: 60, after: 20 },
    children: [run(`Figure ${n}: `, { bold: true, color: NAVY, size: 14 }), ...children] });
}

function source(children) {
  return new Paragraph({ spacing: { before: 20, after: 60 }, children: [
    run("Source: ", { color: GREY, size: 12 }), ...children] });
}

function box(width, height, hint) {
  return new Table({
    width: { size: width, type: WidthType.DXA }, columnWidths: [width], borders: noBorders,
    rows: [new TableRow({
      height: { value: height, rule: HeightRule.EXACT },
      children: [cell([new Paragraph({ alignment: AlignmentType.CENTER, children: [ph(hint, { size: 14 })] })],
                      width, { verticalAlign: VerticalAlign.CENTER,
                               shading: { type: ShadingType.CLEAR, color: "auto", fill: BOX } })],
    })],
  });
}

function emptyFigure(hint, height = BOX_H) {
  return [caption(++figureNo, [ph(`[${hint}]`, { size: 14 })]),
          box(RIGHT, height, "[Insert chart or table]"),
          source([ph("[Company filings / Team analysis]", { size: 12 })])];
}

function image(file, widthDxa) {
  const data = fs.readFileSync(path.join(G, file));
  const w = data.readUInt32BE(16), h = data.readUInt32BE(20);
  const wpx = widthDxa * DXA_PER_PX;
  return new ImageRun({ type: "png", data, transformation: { width: wpx, height: wpx * h / w },
                        altText: { title: file, description: file, name: file } });
}

function imageFigure(text, file) {
  return [caption(++figureNo, [run(text, { color: NAVY, size: 14 })]),
          new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.CENTER,
                          children: [image(file, RIGHT)] })];
}

function writeHere(title, hint) {
  return [sectionHeader(title), body([ph(`[Write here: ${hint}]`)])];
}

// Body page block: subtitles with writing areas and a column of figure slots. A half-page
// block (height HALF_H) starts without a page break when newPage is false.
// A continuation page (no title) passes title = null and only breaks the page.
function standardPage(title, subsections, figures, o = {}) {
  const start = title
    ? pageTitle([run(title, { bold: true, color: NAVY, size: 30 })], o.newPage === false)
    : new Paragraph({ pageBreakBefore: true, spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT },
                      children: [] });
  return [
    start,
    twoColumns(subsections.flatMap(([t, h]) => writeHere(t, h)),
               figures.flatMap((f) => emptyFigure(f, o.boxH)),
               o.height || (title ? ROW_H : ROW_H_CONT), o.figuresLeft),
  ];
}

function risk(id, hint, figs) {
  return [
    body([run(`${id}: `, { bold: true, color: RED }),
          ph(`[Risk title, e.g. ${hint}.]`, { bold: true }), run(" "),
          ph(`[Describe the risk: what could happen, how likely it is, and its impact on revenue, EBITDA or value per share; reference Figures ${figs} where relevant.]`)],
         { after: 10 }),
    body([run("Mitigants: ", { bold: true, color: NAVY }),
          ph("[How the company or the thesis offsets this risk.]")], { after: 50 }),
  ];
}

function riskPage(first = false, figuresLeft = false) {
  const n = figureNo + 1;
  const figs = `${n}–${n + 3}`;
  const right = [
    ...imageFigure("Risk matrix", "fig1_matriz_riesgo.png"),
    ...imageFigure("Operational risk tornado (CLP per share)", "fig2_tornado.png"),
    ...imageFigure("Value per share sensitivity: WACC vs g", "fig3_heatmap.png"),
    ...imageFigure("Earnings sensitivity to a rise in the UF", "fig4_uf.png"),
  ];
  const left = [
    sectionHeader("Risk assessment"),
    body([ph(`[Explain how likelihood and impact were scored (1–5) in Figure ${n} and which risks fall in the high-risk zone.]`)]),
    sectionHeader("Operational risks"),
    ...risk("OR1", "EBITDA margin compression", figs),
    ...risk("OR2", "slowdown in same-store rent growth", figs),
    ...risk("OR3", "higher maintenance capex", figs),
    sectionHeader("Financial risks"),
    ...risk("FR1", "revaluation of UF-indexed debt", figs),
    ...risk("FR2", "higher cost of capital", figs),
    sectionHeader("Market & macroeconomic risks"),
    ...risk("MR1", "FX exposure in Peru and Colombia", figs),
    sectionHeader("Regulatory & ESG risks"),
    ...risk("RR1", "related-party transactions", figs),
    sectionHeader("Sensitivity analysis"),
    body([ph(`[Comment on Figures ${n + 1}–${n + 3}: which drivers move the value per share most, the downside versus the base case, and whether the recommendation holds.]`)]),
  ];
  return [
    pageTitle([run("INVESTMENT RISKS", { bold: true, color: NAVY, size: 30 })], first),
    summaryBar("[Summary: overall risk view for Mallplaza and the main risks to the recommendation and target price.]"),
    twoColumns(left, right, ROW_H_SUMMARY, figuresLeft),
  ];
}

function keyDataTable() {
  const LABEL = 1750;
  const rows = ["Recommendation", "Target price (CLP)", "Current price (CLP)", "Upside / downside",
                "Ticker", "Market cap (CLP bn)", "52-week range (CLP)", "Dividend yield"];
  const line = { style: BorderStyle.SINGLE, size: 4, color: RULE };
  return new Table({
    width: { size: RIGHT, type: WidthType.DXA }, columnWidths: [LABEL, RIGHT - LABEL],
    borders: { ...noBorders, insideHorizontal: line, bottom: line },
    rows: rows.map((r) => new TableRow({ children: [
      cell([new Paragraph({ spacing: { before: 40, after: 40 }, children: [run(r, { bold: true, color: NAVY, size: 14 })] })], LABEL),
      cell([new Paragraph({ spacing: { before: 40, after: 40 }, alignment: AlignmentType.RIGHT, children: [ph("[  ]", { size: 14 })] })], RIGHT - LABEL),
    ] })),
  });
}

function investmentSummaryPage() {
  const left = [
    ...writeHere("Recommendation and target price", "BUY / HOLD / SELL, target price, upside and time horizon."),
    ...writeHere("Investment thesis", "the two or three main reasons behind the recommendation."),
    ...writeHere("Valuation summary", "DCF and multiples result and how they support the target price."),
    ...writeHere("Key risks", "the risks that could change the recommendation (see Investment Risks)."),
  ];
  const right = [
    new Paragraph({ spacing: { before: 60, after: 40 }, children: [run("KEY DATA", { bold: true, color: NAVY, size: 14 })] }),
    keyDataTable(),
    ...emptyFigure("Figure title, e.g. Share price vs. IPSA"),
    ...emptyFigure("Figure title, e.g. Valuation summary"),
  ];
  return [
    pageTitle([run("INVESTMENT SUMMARY", { bold: true, color: NAVY, size: 30 })], true),
    twoColumns(left, right, ROW_H),
  ];
}

// Page budget (10 pages): Investment Summary 1 · Business Description 0.5 · Industry 1.5 ·
// ESG 1 · Financial Analysis 2 · Valuation 2 · Investment Risks 1 · Conclusion 1.
// The figure column alternates: right on odd pages, left on even pages.
function bodyPages() {
  figureNo = 0;
  const L = { figuresLeft: true };
  return [
    ...investmentSummaryPage(),                                                     // p1 right
    ...standardPage("BUSINESS DESCRIPTION", [                                        // p2 left (top half)
      ["Who we are and what we do", "company overview, history and listing."],
      ["Business model", "fixed and variable rents, parking and services."],
      ["Geographic footprint and asset mix", "portfolio by country, GLA and asset type."],
      ["Key growth drivers", "pipeline, expansions and acquisitions."],
      ["Corporate structure and governance", "controlling shareholder and corporate structure."],
    ], ["Figure title, e.g. GLA and revenue by country"], { ...L, height: HALF_H, boxH: HALF_BOX_H }),
    ...standardPage("INDUSTRY OVERVIEW AND COMPETITIVE POSITIONING", [                // p2 left (bottom half)
      ["Industry overview (retail and shopping malls)", "structure of the retail and mall industry in Chile, Peru and Colombia."],
      ["Market size and growth", "market size, growth and penetration."],
    ], ["Figure title, e.g. Retail sales growth by country"],
    { ...L, height: HALF_H, boxH: HALF_BOX_H, newPage: false }),
    ...standardPage(null, [                                                          // p3 right (Industry cont.)
      ["Competitive dynamics", "main competitors and recent moves."],
      ["Porter's five forces", "assessment of each force and its intensity."],
      ["Mallplaza's competitive position", "advantages versus peers (location, scale, mix)."],
      ["Barriers to entry and bargaining power", "land, permits, capital, and tenant and supplier power."],
    ], ["Figure title, e.g. Porter's five forces", "Figure title, e.g. Peer comparison: occupancy and rent",
        "Figure title, e.g. Market share by operator"]),
    ...standardPage("ESG", [                                                          // p4 left
      ["Environmental management", "energy, emissions, water and certifications."],
      ["Social impact and community relations", "employees, tenants, visitors and communities."],
      ["Corporate governance", "board, independence, related parties and incentives."],
      ["Regulatory compliance and standards", "regulation and reporting standards that apply."],
      ["ESG risks and opportunities", "material ESG issues and their link to value."],
    ], ["Figure title, e.g. ESG scorecard", "Figure title, e.g. Energy and emissions intensity",
        "Figure title, e.g. Board composition"], L),
    ...standardPage("FINANCIAL ANALYSIS", [                                           // p5 right
      ["Historical results and trends", "revenue, EBITDA and FFO over time and what drove them."],
      ["Margin and profitability analysis", "EBITDA margin by country and profitability trends."],
      ["Cash flow", "operating cash flow, capex and dividends."],
    ], ["Figure title, e.g. Revenue and EBITDA 2021–2025", "Figure title, e.g. EBITDA margin by country",
        "Figure title, e.g. FFO bridge"]),
    ...standardPage(null, [                                                          // p6 left (FA cont.)
      ["Capital structure and liquidity", "net debt, maturities, UF exposure and liquidity."],
      ["Key ratios (ROIC, ROE, leverage)", "ratio analysis and comparison with peers."],
    ], ["Figure title, e.g. Debt maturity profile", "Figure title, e.g. Net debt/EBITDA and LTV",
        "Figure title, e.g. DuPont ROE decomposition"], L),
    ...standardPage("VALUATION", [                                                    // p7 right
      ["Valuation methodology (DCF + comps)", "methods used and their weighting."],
      ["Key assumptions", "growth, margins, capex and terminal value."],
      ["Cash flow projections", "FCFF forecast 2026E–2031E."],
    ], ["Figure title, e.g. Valuation summary (football field)", "Figure title, e.g. Key assumptions",
        "Figure title, e.g. FCFF 2026E–2031E"]),
    ...standardPage(null, [                                                          // p8 left (Valuation cont.)
      ["Cost of capital (WACC, Ke)", "risk-free rate, beta, ERP, country risk and cost of debt."],
      ["Sensitivity and scenario analysis", "bear, base and bull cases and their drivers."],
      ["Value per share and upside", "value per share versus the current price."],
    ], ["Figure title, e.g. WACC build-up", "Figure title, e.g. Scenario analysis",
        "Figure title, e.g. Target price vs. current price"], L),
    ...riskPage(),                                                                    // p9 right
    ...standardPage("CONCLUSION", [                                                   // p10 left
      ["General conclusion", "how the analysis fits together into the investment case."],
      ["Investment recommendation", "final recommendation and the reasoning behind it."],
      ["Catalysts and drivers", "events or trends that could unlock value and their timing."],
      ["Key risks", "the main risks to the recommendation."],
      ["Target price", "final target price and upside versus the current price."],
    ], ["Figure title, e.g. Investment case summary", "Figure title, e.g. Catalyst timeline",
        "Figure title, e.g. Scenario target prices"], L),
  ];
}

const APPENDICES = ["Income statement 2022A–2031E", "Balance sheet", "Cash flow statement",
  "Forecast assumptions", "DCF valuation detail", "WACC calculation", "Comparable companies",
  "DuPont and key ratios", "Sensitivity and scenario analysis", "ESG scorecard and risk detail"];

function appendixPages() {
  return APPENDICES.flatMap((hint, i) => [
    pageTitle([run(`APPENDIX ${i + 1}: `, { bold: true, color: NAVY, size: 30 }),
               ph(`[Title, e.g. ${hint}]`, { bold: true, size: 30 })], i === 0),
    box(CONTENT_W, ROW_H - 700, "[Insert table or exhibit]"),
    source([ph("[Company filings / Team analysis]", { size: 12 })]),
  ]);
}

function header() {
  return new Header({ children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 2 } },
    children: [run("Mallplaza (Plaza S.A.)  |  CFA Institute Research Challenge", { color: GREY, size: 14 })],
  })] });
}

function footer(prefix) {
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
    new TextRun({ children: [prefix, PageNumber.CURRENT], font: FONT, size: 14, color: GREY })] })] });
}

function section(children, prefix) {
  return {
    properties: { page: { size: PAGE, margin: MARGIN, pageNumbers: { start: 1 } } },
    headers: { default: header() }, footers: { default: footer(prefix) }, children,
  };
}

const styles = { default: { document: { run: { font: FONT, size: 16 } } } };

// Wrap each placeholder run in a content control: clicking it selects the hint and
// typing replaces it with normal (non-italic, dark) text.
function toContentControls(xml) {
  let id = 1000;
  return xml.replace(/<w:r>(?:(?!<\/w:r>).)*?w:val="808081"(?:(?!<\/w:r>).)*?<\/w:r>/gs, (r) => {
    const bold = /<w:b\/>|<w:b w:val="true"\/>/.test(r) ? "<w:b/><w:bCs/>" : "";
    const size = (r.match(/<w:sz w:val="(\d+)"\/>/) || [, "16"])[1];
    const rPr = `<w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>${bold}` +
                `<w:color w:val="1A1A1A"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr>`;
    return `<w:sdt><w:sdtPr>${rPr}<w:id w:val="${id++}"/><w:showingPlcHdr/></w:sdtPr>` +
           `<w:sdtContent>${r}</w:sdtContent></w:sdt>`;
  });
}

async function save(doc, file) {
  const zip = await require("jszip").loadAsync(await Packer.toBuffer(doc));
  const f = "word/document.xml";
  zip.file(f, toContentControls(await zip.file(f).async("string")));
  const out = path.join(__dirname, file);
  fs.writeFileSync(out, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  console.log(out);
}

(async () => {
  await save(new Document({ styles, sections: [section(bodyPages(), "Page "),
                                               section(appendixPages(), "A-")] }),
             "Mallplaza_CFA_Report_template.docx");
  figureNo = 0;
  await save(new Document({ styles, sections: [section(riskPage(true), "Page ")] }),
             "Investment_Risks_template.docx");
})();
