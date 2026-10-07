// One-page "Investment Risks" template (CFA Research Challenge style) for Mallplaza.
// Usage: node informe/investment_risks.js  → informe/Investment_Risks_template.docx
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign, LineRuleType,
} = require("docx");

const NAVY = "14284B";
const RED = "E9004B";
const GREY = "5F5F5F";
const PLACEHOLDER = "808081"; // marker color: post-processed into Word placeholders
const FONT = "Arial";
const G = path.join(__dirname, "..", "graficos", "informe");

const PAGE_W = 12240, MARGIN = 620;
const CONTENT = PAGE_W - 2 * MARGIN;   // 11000
const LEFT = 6950, GUTTER = 250, RIGHT = CONTENT - LEFT - GUTTER;
const DXA_PER_PX = 96 / 1440;          // image sizes are in px at 96 dpi

const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: none, bottom: none, left: none, right: none,
                    insideHorizontal: none, insideVertical: none };

const ph = (text, opts = {}) => new TextRun({ text, color: PLACEHOLDER, italics: true,
  font: FONT, size: opts.size || 16, bold: opts.bold });

function body(children, opts = {}) {
  return new Paragraph({ children, spacing: { before: 0, after: opts.after ?? 40, line: 228 },
                         alignment: opts.align || AlignmentType.JUSTIFIED });
}

function sectionHeader(text) {
  return new Paragraph({
    children: [new TextRun({ text: text.toUpperCase(), bold: true, color: NAVY, font: FONT, size: 16 })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: RED, space: 1 } },
    spacing: { before: 100, after: 50 },
  });
}

// Free-writing area under a subtitle: one content control that accepts any number of
// paragraphs, followed by an empty spacer that reserves room on the page.
function zone(hint, reserve) {
  return [
    body([ph(hint)], { after: 0 }),
    new Paragraph({ children: [], spacing: { before: 0, after: 0, line: reserve,
                                              lineRule: LineRuleType.EXACT } }),
  ];
}

function image(file, widthDxa) {
  const data = fs.readFileSync(path.join(G, file));
  const w = data.readUInt32BE(16), h = data.readUInt32BE(20);
  const wpx = widthDxa * DXA_PER_PX;
  return new ImageRun({ type: "png", data, transformation: { width: wpx, height: wpx * h / w },
                        altText: { title: file, description: file, name: file } });
}

function figure(n, caption, file, widthDxa) {
  return [
    new Paragraph({ spacing: { before: 60, after: 20 },
      children: [new TextRun({ text: `Figure ${n}: `, bold: true, color: NAVY, font: FONT, size: 14 }),
                 new TextRun({ text: caption, color: NAVY, font: FONT, size: 14 })] }),
    new Paragraph({ spacing: { after: 0 }, alignment: AlignmentType.CENTER, children: [image(file, widthDxa)] }),
  ];
}

const cell = (children, width, extra = {}) => new TableCell({
  children, width: { size: width, type: WidthType.DXA }, borders: noBorders,
  margins: { top: 0, bottom: 0, left: 0, right: 0 }, ...extra });

// Left column (~65%): subtitles with free-writing areas. Right column (~35%): figures.
const left = [
  sectionHeader("Risk assessment"),
  ...zone("[Write here: how likelihood and impact were scored (1–5) in Figure 1 and which risks fall in the high-risk zone.]", 500),
  sectionHeader("Operational risks"),
  ...zone("[Write here. Suggested format per risk: OR1 – title (e.g. EBITDA margin compression), description, impact on value (Figure 2) and mitigants. Continue with OR2, OR3…]", 2800),
  sectionHeader("Financial risks"),
  ...zone("[Write here. e.g. FR1 – revaluation of UF-indexed debt (Figure 4); FR2 – higher cost of capital (Figure 3). Include description and mitigants.]", 1700),
  sectionHeader("Market & macroeconomic risks"),
  ...zone("[Write here. e.g. MR1 – FX exposure in Peru and Colombia. Include description and mitigants.]", 1100),
  sectionHeader("Regulatory & ESG risks"),
  ...zone("[Write here. e.g. RR1 – related-party transactions. Include description and mitigants.]", 1100),
  sectionHeader("Sensitivity analysis"),
  ...zone("[Write here: which drivers move the value per share most (Figures 2–4), the downside versus the base case, and whether the recommendation holds.]", 1900),
];

const right = [
  ...figure(1, "Risk matrix", "fig1_matriz_riesgo.png", RIGHT),
  ...figure(2, "Operational risk tornado (value per share, CLP)", "fig2_tornado.png", RIGHT),
  ...figure(3, "Value per share sensitivity: WACC vs g", "fig3_heatmap.png", RIGHT),
  ...figure(4, "Earnings sensitivity to a rise in the UF", "fig4_uf.png", RIGHT),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 16 } } } },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 15840 },
                          margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    children: [
      new Paragraph({
        children: [new TextRun({ text: "INVESTMENT RISKS", bold: true, color: NAVY, font: FONT, size: 30 })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: NAVY, space: 2 } },
        spacing: { after: 80 },
      }),
      new Paragraph({
        shading: { type: ShadingType.CLEAR, color: "auto", fill: "EEF1F6" },
        spacing: { after: 80, line: 228 },
        children: [ph("[Summary: overall risk view for Mallplaza and the main risks to the recommendation and target price.]", { size: 16 })],
      }),
      new Table({
        width: { size: CONTENT, type: WidthType.DXA }, columnWidths: [LEFT, GUTTER, RIGHT],
        borders: noBorders,
        rows: [new TableRow({ children: [
          cell(left, LEFT, { verticalAlign: VerticalAlign.TOP }),
          cell([new Paragraph("")], GUTTER),
          cell(right, RIGHT, { verticalAlign: VerticalAlign.TOP }),
        ] })],
      }),
    ],
  }],
});

// Wrap each placeholder paragraph in a block-level content control: clicking selects the
// hint, typing replaces it with normal text, and Enter adds paragraphs inside the area.
function toContentControls(xml) {
  let id = 1000;
  return xml.replace(/<w:p>(?:(?!<\/w:p>).)*?w:val="808081"(?:(?!<\/w:p>).)*?<\/w:p>/gs, (p) => {
    const size = (p.match(/<w:sz w:val="(\d+)"\/>/) || [, "16"])[1];
    const rPr = `<w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>` +
                `<w:color w:val="1A1A1A"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr>`;
    return `<w:sdt><w:sdtPr>${rPr}<w:id w:val="${id++}"/><w:showingPlcHdr/></w:sdtPr>` +
           `<w:sdtContent>${p}</w:sdtContent></w:sdt>`;
  });
}

const out = path.join(__dirname, "Investment_Risks_template.docx");
Packer.toBuffer(doc)
  .then((buf) => require("jszip").loadAsync(buf))
  .then(async (zip) => {
    const f = "word/document.xml";
    zip.file(f, toContentControls(await zip.file(f).async("string")));
    fs.writeFileSync(out, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
    console.log(out);
  });
