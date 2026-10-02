// Word (.docx) export of one APC calculation (A4 landscape, structured
// tables), rendered from the shared ReportModel with the `docx` package.

import type * as DocxNs from "docx";
import type { ApcResult, MfEntry } from "./apc";
import { buildReportModel, type ReportModel, type ReportTable } from "./reportModel";
import { downloadBlob } from "./download";

type Docx = typeof DocxNs;

// A4 in twentieths of a point (DXA). `docx` expects portrait dimensions and
// swaps them itself when the orientation is LANDSCAPE.
const A4_SHORT = 11906;
const A4_LONG = 16838;
const MARGIN = 720;
const CONTENT_W = A4_LONG - 2 * MARGIN; // landscape content width

const INK = "0B0B0B";
const MUTED = "52514E";
const HEAD_FILL = "1F3864";
const FOOT_FILL = "ECEFF4";
const ALT_FILL = "FAFAF8";
const DANGER = "D03B3B";
const DANGER_FILL = "FDF3F3";
const HERO_FILL = "F5F7FB";

/** Build the Word document (no download). Exposed for tests. */
export function buildApcDocx(d: Docx, model: ReportModel): DocxNs.Document {
  const {
    Document, Paragraph, TextRun, Table, TableRow, TableCell,
    WidthType, AlignmentType, BorderStyle, ShadingType, PageOrientation,
    HeadingLevel, TableLayoutType,
  } = d;

  const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const noBorders = {
    top: noBorder, bottom: noBorder, left: noBorder, right: noBorder,
    insideHorizontal: noBorder, insideVertical: noBorder,
  };
  const cellMargins = { top: 50, bottom: 50, left: 90, right: 90 };

  const text = (
    value: string,
    opts: { bold?: boolean; color?: string; size?: number; italics?: boolean } = {}
  ) => new TextRun({ text: value, font: "Calibri", color: opts.color ?? INK, ...opts });

  const para = (
    value: string,
    opts: {
      bold?: boolean; color?: string; size?: number; italics?: boolean;
      align?: (typeof AlignmentType)[keyof typeof AlignmentType];
      before?: number; after?: number;
    } = {}
  ) =>
    new Paragraph({
      alignment: opts.align,
      spacing: { before: opts.before ?? 0, after: opts.after ?? 0 },
      children: [text(value, opts)],
    });

  // ---- Inputs table (borderless) ----
  const metaW = Math.round(CONTENT_W * 0.62);
  const metaLabelW = 3200;
  const metaTable = new Table({
    width: { size: metaW, type: WidthType.DXA },
    columnWidths: [metaLabelW, metaW - metaLabelW],
    layout: TableLayoutType.FIXED,
    borders: noBorders,
    rows: model.meta.map(
      ([k, v]) =>
        new TableRow({
          children: [
            new TableCell({
              width: { size: metaLabelW, type: WidthType.DXA },
              margins: cellMargins,
              children: [para(k, { bold: true, color: MUTED, size: 18 })],
            }),
            new TableCell({
              width: { size: metaW - metaLabelW, type: WidthType.DXA },
              margins: cellMargins,
              children: [para(v, { size: 18 })],
            }),
          ],
        })
    ),
  });

  // ---- APC-2 highlight box ----
  const heroTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: HEAD_FILL },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: HEAD_FILL },
      left: { style: BorderStyle.SINGLE, size: 6, color: HEAD_FILL },
      right: { style: BorderStyle.SINGLE, size: 6, color: HEAD_FILL },
      insideHorizontal: noBorder, insideVertical: noBorder,
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: CONTENT_W, type: WidthType.DXA },
            shading: { fill: HERO_FILL, type: ShadingType.CLEAR, color: "auto" },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            children: [
              para(model.hero.label, { color: MUTED, size: 18 }),
              para(model.hero.value, { bold: true, color: HEAD_FILL, size: 36 }),
              para(model.hero.perDay, { color: MUTED, size: 18 }),
            ],
          }),
        ],
      }),
    ],
  });

  const children: (DocxNs.Paragraph | DocxNs.Table)[] = [
    new Paragraph({
      heading: HeadingLevel.TITLE,
      spacing: { after: 60 },
      children: [text(model.title, { bold: true, size: 32 })],
    }),
    para(model.subtitle, { color: MUTED, size: 20, after: 160 }),
    metaTable,
    para("", { after: 120 }),
    heroTable,
    para("", { after: 160 }),
  ];

  if (model.alert) {
    children.push(
      new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: [CONTENT_W],
        layout: TableLayoutType.FIXED,
        borders: {
          top: { style: BorderStyle.SINGLE, size: 6, color: DANGER },
          bottom: { style: BorderStyle.SINGLE, size: 6, color: DANGER },
          left: { style: BorderStyle.SINGLE, size: 6, color: DANGER },
          right: { style: BorderStyle.SINGLE, size: 6, color: DANGER },
          insideHorizontal: noBorder, insideVertical: noBorder,
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: CONTENT_W, type: WidthType.DXA },
                shading: { fill: DANGER_FILL, type: ShadingType.CLEAR, color: "auto" },
                margins: { top: 100, bottom: 100, left: 200, right: 200 },
                children: [para(model.alert, { color: DANGER, size: 19 })],
              }),
            ],
          }),
        ],
      }),
      para("", { after: 160 })
    );
  }

  const dataTable = (t: ReportTable, width: number) => {
    const totalW = t.widths?.reduce((s, w) => s + w, 0) ?? t.head.length;
    const colW = t.head.map((_, i) =>
      Math.round(((t.widths ? t.widths[i] : 1) / totalW) * width)
    );
    const cell = (
      value: string,
      i: number,
      kind: "head" | "body" | "foot",
      rowIndex: number
    ) => {
      const numeric = t.numericCols.includes(i);
      const imputed = kind === "body" && value.includes("(avg)");
      const emphasized = kind === "body" && !!t.emphasizeLastRow && rowIndex === t.body.length - 1;
      const fill =
        kind === "head" ? HEAD_FILL
        : kind === "foot" || emphasized ? FOOT_FILL
        : rowIndex % 2 === 1 ? ALT_FILL : undefined;
      return new TableCell({
        width: { size: colW[i], type: WidthType.DXA },
        margins: cellMargins,
        verticalAlign: d.VerticalAlign.CENTER,
        shading: fill ? { fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
        children: [
          para(value, {
            size: kind === "head" ? 15 : 16,
            bold: kind !== "body" || emphasized,
            italics: imputed,
            color: kind === "head" ? "FFFFFF" : imputed ? MUTED : INK,
            align:
              kind === "head" ? AlignmentType.CENTER : numeric ? AlignmentType.RIGHT : AlignmentType.LEFT,
          }),
        ],
      });
    };
    const rows: DocxNs.TableRow[] = [
      new TableRow({
        tableHeader: true,
        children: t.head.map((h, i) => cell(h, i, "head", 0)),
      }),
      ...t.body.map(
        (r, ri) => new TableRow({ children: r.map((v, i) => cell(v, i, "body", ri)) })
      ),
      ...(t.foot ?? []).map(
        (r) => new TableRow({ children: r.map((v, i) => cell(v, i, "foot", 0)) })
      ),
    ];
    return new Table({
      width: { size: width, type: WidthType.DXA },
      columnWidths: colW,
      layout: TableLayoutType.FIXED,
      rows,
    });
  };

  const section = (t: ReportTable, width = CONTENT_W) => {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        keepNext: true,
        spacing: { before: 200, after: 60 },
        children: [text(t.title, { bold: true, size: 22 })],
      })
    );
    if (t.note) children.push(para(t.note, { color: MUTED, size: 17, after: 80 }));
    children.push(dataTable(t, width));
  };

  section(model.table1);
  section(model.table2);
  section(model.table3, Math.round(CONTENT_W * 0.7));

  children.push(
    new Paragraph({
      heading: HeadingLevel.HEADING_2,
      keepNext: true,
      spacing: { before: 240, after: 60 },
      children: [text("Notes", { bold: true, size: 20 })],
    }),
    ...model.notes.map(
      (n) =>
        new Paragraph({
          bullet: { level: 0 },
          spacing: { after: 40 },
          children: [text(n, { color: MUTED, size: 17 })],
        })
    )
  );

  return new Document({
    creator: "APC 2.0",
    title: model.title,
    styles: {
      default: { document: { run: { font: "Calibri", size: 20, color: INK } } },
    },
    sections: [
      {
        properties: {
          page: {
            size: { orientation: PageOrientation.LANDSCAPE, width: A4_SHORT, height: A4_LONG },
            margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN },
          },
        },
        children,
      },
    ],
  });
}

export async function exportApcDocx(
  plaza: { code: number; name: string },
  result: ApcResult,
  uptoMonth: string,
  mfEntries: MfEntry[],
  growth: number
) {
  const docx = await import("docx");
  const model = buildReportModel(plaza, result, uptoMonth, mfEntries, growth);
  const doc = buildApcDocx(docx, model);
  const blob = await docx.Packer.toBlob(doc);
  downloadBlob(blob, `${model.fileBase}.docx`);
}
