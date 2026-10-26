// 私立高校 Web出願チェックシート（Word・A4縦）を生成する
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, ShadingType, BorderStyle, VerticalAlign,
} = require("docx");

const FONT = "Yu Gothic";
const W = 10206; // 本文幅（A4 - 左右余白 850×2）
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "808080" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

const t = (text, opts = {}) => new TextRun({ text, font: { eastAsia: FONT, ascii: FONT, hAnsi: FONT }, size: 19, ...opts });
const p = (children, opts = {}) => new Paragraph({ children: Array.isArray(children) ? children : [children], spacing: { before: 0, after: 0 }, ...opts });

function cell(content, width, opts = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map((c) =>
    typeof c === "string" ? p(t(c, opts.run || {}), { alignment: opts.align }) : c);
  return new TableCell({
    children: paras,
    width: { size: width, type: WidthType.DXA },
    borders: BORDERS,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    shading: opts.fill ? { type: ShadingType.CLEAR, color: "auto", fill: opts.fill } : undefined,
    columnSpan: opts.span,
  });
}

// ── データ（data/checklist.json を Word・Excel・Web で共有） ──
const data = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "data", "checklist.json"), "utf8"));
const sections = data.sections.map((s) => [s.title, s.items]);

// ── 部品 ─────────────────────────────
const COLS = [520, 5780, 2306, 800, 800]; // No / 確認項目 / 補足 / 生徒 / 担任
const header = new TableRow({
  tableHeader: true,
  children: ["No", "確認項目", "補足・メモ", "生徒", "担任"].map((h, i) =>
    cell(h, COLS[i], { fill: "1F4E79", align: AlignmentType.CENTER, run: { bold: true, color: "FFFFFF" } })),
});

const rows = [header];
let n = 1;
for (const [title, items] of sections) {
  rows.push(new TableRow({
    cantSplit: true,
    children: [cell(title, W, { span: 5, fill: "DDEBF7", run: { bold: true, size: 21 } })],
  }));
  for (const [item, note] of items) {
    rows.push(new TableRow({
      cantSplit: true,
      children: [
        cell(String(n++), COLS[0], { align: AlignmentType.CENTER }),
        cell(item, COLS[1]),
        cell(note, COLS[2], { run: { size: 16, color: "555555" } }),
        cell("□", COLS[3], { align: AlignmentType.CENTER, run: { size: 24 } }),
        cell("□", COLS[4], { align: AlignmentType.CENTER, run: { size: 24 } }),
      ],
    }));
  }
}
const checklist = new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: COLS, rows });

// 生徒情報
const infoCols = [1400, 3700, 1400, 3706];
const info = new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: infoCols,
  rows: [new TableRow({ children: [
    cell("年　組　番", infoCols[0], { fill: "F2F2F2", align: AlignmentType.CENTER }),
    cell("　　年　　組　　番", infoCols[1]),
    cell("氏　名", infoCols[2], { fill: "F2F2F2", align: AlignmentType.CENTER }),
    cell("", infoCols[3]),
  ] })],
});

// 受験校の一覧
const schoolCols = [1846, 1400, 1100, 1180, 1180, 1100, 1100, 1300];
const schoolHead = data.schoolColumns;
const schools = new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: schoolCols,
  rows: [
    new TableRow({ children: schoolHead.map((h, i) => cell(h, schoolCols[i], { fill: "F2F2F2", align: AlignmentType.CENTER, run: { bold: true, size: 15 } })) }),
    ...[1, 2, 3].map(() => new TableRow({ height: { value: 520, rule: "atLeast" }, children: schoolCols.map((w) => cell("", w)) })),
  ],
});

// 保護者確認
const signCols = [2000, 3103, 2000, 3103];
const sign = new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: signCols,
  rows: [new TableRow({ height: { value: 600, rule: "atLeast" }, children: [
    cell("保護者確認", signCols[0], { fill: "F2F2F2", align: AlignmentType.CENTER }),
    cell("", signCols[1]),
    cell("担任確認", signCols[2], { fill: "F2F2F2", align: AlignmentType.CENTER }),
    cell("", signCols[3]),
  ] })],
});

const heading = (text) => p(t(text, { bold: true, size: 22, color: "1F4E79" }), { spacing: { before: 200, after: 80 } });

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 19 } } } },
  sections: [{
    properties: { page: { margin: { top: 720, bottom: 720, left: 850, right: 850 } } },
    children: [
      p(t(data.title, { bold: true, size: 32 }), { alignment: AlignmentType.CENTER, spacing: { after: 60 } }),
      p(t(data.subtitle, { size: 17, color: "555555" }), { alignment: AlignmentType.CENTER, spacing: { after: 160 } }),
      info,
      heading("受験する高校"),
      schools,
      heading("チェック項目"),
      checklist,
      heading("最終確認"),
      sign,
    ],
  }],
});

const out = path.join(__dirname, "..", "output", "私立高校_Web出願チェックシート.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log(out); });
