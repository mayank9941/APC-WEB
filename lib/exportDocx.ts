// Word (.docx) export of one APC calculation laid out like NHAI's single
// A4-landscape sheet (same structure as the PDF): header bar + plaza/MF
// block, Table-1 with lettered formula row and rounded % column, Table-2
// beside the chart, Table-3 with the formula box, Table-4 (APC-3).

import type * as DocxNs from "docx";
import type { ApcResult, MfEntry, PlazaInfo } from "./apc";
import { buildSheetModel, type SheetModel } from "./reportModel";
import { renderChartPng } from "./chartImage";
import { downloadBlob } from "./download";

type Docx = typeof DocxNs;

// A4 in DXA. `docx` expects portrait dimensions and swaps them for landscape.
const A4_SHORT = 11906;
const A4_LONG = 16838;
const MARGIN = 560;
const CONTENT_W = A4_LONG - 2 * MARGIN;
const MM = CONTENT_W / 281; // DXA per mm of the PDF layout (281 mm content width)

const INK = "000000";
const GREY = "6E6E6E";
const RED = "C80000";
const LINE = "787878";
const PEACH = "FCE4D6";
const PEACH_LINE = "F4B183";
const YELLOW = "FFFF00";
const ORANGE = "FFC000";
const BLUE_FILL = "DDEBF7";

const BASE = 13; // half-points (6.5pt)

export interface DocxAssets {
  /** PNG bytes of the Table-2 chart; omitted when no canvas is available. */
  chartPng?: Uint8Array | null;
}

/** Build the Word document (no download). Exposed for tests. */
export function buildApcDocx(d: Docx, model: SheetModel, assets: DocxAssets = {}): DocxNs.Document {
  const {
    Document, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
    WidthType, AlignmentType, BorderStyle, ShadingType, PageOrientation,
    TableLayoutType, VerticalAlign,
  } = d;

  type Border = { style: (typeof BorderStyle)[keyof typeof BorderStyle]; size: number; color: string };
  type Borders = { top: Border; bottom: Border; left: Border; right: Border };
  const solid: Border = { style: BorderStyle.SINGLE, size: 4, color: LINE };
  const none: Border = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const allSolid: Borders = { top: solid, bottom: solid, left: solid, right: solid };
  const allNone: Borders = { top: none, bottom: none, left: none, right: none };
  const tableNoBorders = { ...allNone, insideHorizontal: none, insideVertical: none };
  const tableSolid = { ...allSolid, insideHorizontal: solid, insideVertical: solid };
  const pad = { top: 25, bottom: 25, left: 50, right: 50 };

  type Align = (typeof AlignmentType)[keyof typeof AlignmentType];
  interface TextOpts { bold?: boolean; color?: string; size?: number; italics?: boolean; align?: Align; before?: number; after?: number }
  const run = (value: string, o: TextOpts = {}) =>
    new TextRun({ text: value, font: "Calibri", color: o.color ?? INK, bold: o.bold, italics: o.italics, size: o.size ?? BASE });
  const para = (value: string | DocxNs.TextRun[], o: TextOpts = {}) =>
    new Paragraph({
      alignment: o.align ?? AlignmentType.CENTER,
      spacing: { before: o.before ?? 0, after: o.after ?? 0 },
      children: typeof value === "string" ? [run(value, o)] : value,
    });

  interface CellOpts { width: number; fill?: string; align?: Align; bold?: boolean; color?: string; size?: number; borders?: Borders; colSpan?: number; children?: (DocxNs.Paragraph | DocxNs.Table)[]; vAlign?: DocxNs.TableVerticalAlign }
  const cell = (value: string, o: CellOpts) =>
    new TableCell({
      width: { size: o.width, type: WidthType.DXA },
      columnSpan: o.colSpan,
      margins: pad,
      verticalAlign: o.vAlign ?? VerticalAlign.CENTER,
      borders: o.borders ?? allSolid,
      shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
      children: o.children ?? [para(value, { bold: o.bold, color: o.color, size: o.size, align: o.align })],
    });
  const grid = (widths: number[], rows: DocxNs.TableRow[], borders = tableSolid) =>
    new Table({
      width: { size: widths.reduce((s, w) => s + w, 0), type: WidthType.DXA },
      columnWidths: widths,
      layout: TableLayoutType.FIXED,
      borders,
      rows,
    });
  const layoutCell = (width: number, children: (DocxNs.Paragraph | DocxNs.Table)[], vAlign: DocxNs.TableVerticalAlign = VerticalAlign.TOP) =>
    new TableCell({
      width: { size: width, type: WidthType.DXA },
      borders: allNone,
      verticalAlign: vAlign,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
      children,
    });
  const spacer = (after = 80) => para("", { after });
  const mm = (v: number) => Math.round(v * MM);

  // ---------- Header: title bar + Calculated on ----------
  const titleW = mm(196);
  const calcW = mm(54);
  const header = grid(
    [titleW, CONTENT_W - titleW - calcW, calcW],
    [
      new TableRow({
        children: [
          cell(model.title, { width: titleW, fill: PEACH, bold: true, size: 20, borders: allNone }),
          layoutCell(CONTENT_W - titleW - calcW, [para("")]),
          layoutCell(calcW, [
            grid([mm(24), mm(30)], [
              new TableRow({ children: [cell("Calculated on", { width: mm(24), bold: true, size: 14 }), cell(model.calculatedOn, { width: mm(30), size: 14 })] }),
            ]),
          ]),
        ],
      }),
    ],
    tableNoBorders
  );

  // ---------- Plaza code / MF block / plaza info ----------
  const mfLabels = model.mf.labels.length ? model.mf.labels : ["MF"];
  const mfVals = (v: string[]) => (v.length ? v : ["-"]);
  const mfLabelW = mm(24);
  const mfColW = mm(19);
  const mfWidths = [mfLabelW, ...mfLabels.map(() => mfColW)];
  const mfRow = (name: string, vals: string[], o: { fill?: string; red?: boolean; bold?: boolean }) =>
    new TableRow({
      children: [
        cell(name, { width: mfLabelW, align: AlignmentType.RIGHT, bold: true, color: o.red ? RED : INK, borders: allNone, size: 14 }),
        ...vals.map((v) => cell(v, { width: mfColW, fill: o.fill, bold: o.bold, size: 14 })),
      ],
    });
  const mfTable = grid(mfWidths, [
    new TableRow({ children: [cell("", { width: mfLabelW, borders: allNone }), ...mfLabels.map((l) => cell(l, { width: mfColW, bold: true, size: 14 }))] }),
    mfRow("Old remittance", mfVals(model.mf.oldRemittance), { fill: YELLOW, red: true }),
    mfRow("New remittance", mfVals(model.mf.newRemittance), { fill: YELLOW, red: true }),
    mfRow("MF", mfVals(model.mf.factor), { bold: true }),
  ], tableNoBorders);

  const infoTable = grid([mm(24), mm(30)], model.plazaInfo.map(([k, v]) =>
    new TableRow({ children: [cell(k, { width: mm(24), align: AlignmentType.RIGHT, bold: true, size: 14 }), cell(v, { width: mm(30), size: 14 })] })
  ));

  const plazaW = mm(64);
  const infoW = mm(54);
  const infoRow = grid(
    [plazaW, CONTENT_W - plazaW - infoW, infoW],
    [
      new TableRow({
        children: [
          layoutCell(plazaW, [
            para([run("Plaza code:  ", { size: 15 }), run(model.plazaCode, { bold: true, color: RED, size: 16 })], { align: AlignmentType.LEFT, before: 60 }),
            para(model.plazaName, { align: AlignmentType.LEFT, size: 14, before: 40 }),
          ]),
          layoutCell(CONTENT_W - plazaW - infoW, [mfTable]),
          layoutCell(infoW, [infoTable]),
        ],
      }),
    ],
    tableNoBorders
  );

  // ---------- Table-1 ----------
  const t1 = model.table1;
  const t1Widths = [40, 27, 20, 18, 18, 26, 16, 28, 22, 23].map(mm);
  const roundW = mm(14);
  const t1All = [...t1Widths, roundW];
  const headRow = new TableRow({
    tableHeader: true,
    children: [...t1.head.map((h, i) => cell(h, { width: t1Widths[i], bold: true })), cell("", { width: roundW, borders: allNone })],
  });
  const letterRow = new TableRow({
    children: [...t1.letters.map((l, i) => cell(l, { width: t1Widths[i] })), cell("", { width: roundW, borders: allNone })],
  });
  const bodyRows = t1.body.map((r, ri) =>
    new TableRow({
      children: [
        ...r.map((v, i) =>
          cell(v, {
            width: t1Widths[i],
            align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
            fill: i === 2 || i === 7 || i === 8 ? BLUE_FILL : i === 3 || i === 4 ? YELLOW : undefined,
          })
        ),
        cell(t1.rounded[ri], { width: roundW, borders: allNone, align: AlignmentType.RIGHT }),
      ],
    })
  );
  const weightedRow = new TableRow({
    children: [
      cell(t1.weightedLabel, { width: t1Widths.slice(0, 6).reduce((s, w) => s + w, 0), colSpan: 6, bold: true, align: AlignmentType.RIGHT }),
      cell(t1.weighted, { width: t1Widths[6], fill: ORANGE, bold: true, size: 15 }),
      cell("", { width: t1Widths.slice(7).reduce((s, w) => s + w, 0), colSpan: 3, borders: allNone }),
      cell(t1.roundedTotal, { width: roundW, fill: YELLOW, bold: true, borders: allNone, align: AlignmentType.RIGHT }),
    ],
  });
  const table1 = grid(t1All, [headRow, letterRow, ...bodyRows, weightedRow], tableNoBorders);

  // ---------- Table-2 + chart ----------
  const t2 = model.table2;
  const t2Widths = [16, 26, 32, 34].map(mm);
  const table2 = grid(t2Widths, [
    new TableRow({
      tableHeader: true,
      children: t2.head.map((h, i) => cell(h, { width: t2Widths[i], bold: true, color: i === 1 ? RED : INK })),
    }),
    ...t2.body.map((r) =>
      new TableRow({
        children: r.map((v, i) => cell(v, { width: t2Widths[i], bold: i === 0, fill: i === 2 ? BLUE_FILL : undefined, color: r[0].endsWith("*") ? GREY : INK })),
      })
    ),
    new TableRow({
      children: t2.average.map((v, i) => cell(v, { width: t2Widths[i], bold: true, fill: i === 3 ? ORANGE : undefined, size: i === 3 ? 15 : BASE })),
    }),
  ]);

  const t2W = t2Widths.reduce((s, w) => s + w, 0);
  const chartW = CONTENT_W - t2W - mm(8);
  const chartPx = Math.round((chartW / 1440) * 96);
  const chartChildren: (DocxNs.Paragraph | DocxNs.Table)[] = assets.chartPng
    ? [new Paragraph({ children: [new ImageRun({ type: "png", data: assets.chartPng, transformation: { width: chartPx, height: Math.round(chartPx / 2) } })] })]
    : [
        grid([chartW], [
          new TableRow({
            children: [
              new TableCell({
                width: { size: chartW, type: WidthType.DXA },
                borders: { top: { ...solid, color: PEACH_LINE, size: 8 }, bottom: { ...solid, color: PEACH_LINE, size: 8 }, left: { ...solid, color: PEACH_LINE, size: 8 }, right: { ...solid, color: PEACH_LINE, size: 8 } },
                margins: { top: 200, bottom: 200, left: 120, right: 120 },
                children: [
                  para(model.chart.title, { bold: true, color: RED, size: 17 }),
                  para(model.chart.points.map((p) => `${p.label}: ${p.value === null ? "-" : Math.round(p.value).toLocaleString("en-IN")}`).join("   "), { size: 12, color: GREY, before: 80 }),
                ],
              }),
            ],
          }),
        ]),
      ];

  const row2 = grid(
    [t2W, mm(8), chartW],
    [
      new TableRow({
        children: [
          layoutCell(t2W, [
            para(t2.title, { bold: true, size: 16, align: AlignmentType.LEFT, after: 40 }),
            table2,
            ...(t2.imputedNote ? [para(t2.imputedNote, { size: 12, color: GREY, align: AlignmentType.LEFT, before: 40 })] : []),
          ]),
          layoutCell(mm(8), [para("")]),
          layoutCell(chartW, [spacer(120), ...chartChildren]),
        ],
      }),
    ],
    tableNoBorders
  );

  // ---------- Table-3 + formula + Table-4 ----------
  const t3 = model.table3;
  const t3Widths = [10, 52, 46].map(mm);
  const table3 = grid(t3Widths, t3.rows.map((r, ri) =>
    new TableRow({
      children: r.map((v, i) =>
        cell(v, {
          width: t3Widths[i],
          align: i === 2 ? AlignmentType.RIGHT : AlignmentType.CENTER,
          bold: ri === t3.rows.length - 1,
          fill: ri === t3.rows.length - 1 && i === 2 ? ORANGE : undefined,
          size: ri === t3.rows.length - 1 && i === 2 ? 15 : BASE,
        })
      ),
    })
  ));
  const t3W = t3Widths.reduce((s, w) => s + w, 0);

  const formulaW = mm(72);
  const formulaBox = grid([formulaW], [
    new TableRow({
      children: [
        new TableCell({
          width: { size: formulaW, type: WidthType.DXA },
          borders: allSolid,
          margins: { top: 100, bottom: 100, left: 120, right: 120 },
          children: [
            para([new TextRun({ text: model.formula.heading, font: "Calibri", bold: true, size: 15, underline: {} })], { align: AlignmentType.LEFT, after: 60 }),
            ...model.formula.lines.map((l) => para(l, { size: 14, align: AlignmentType.LEFT, after: 20 })),
          ],
        }),
      ],
    }),
  ]);

  const t4Widths = [50, 28].map(mm);
  const t4W = t4Widths[0] + t4Widths[1];
  const table4: DocxNs.Table | DocxNs.Paragraph = model.table4
    ? grid(t4Widths, [
        new TableRow({ children: [cell(model.table4.title, { width: t4W, colSpan: 2, bold: true })] }),
        ...model.table4.rows.map((r, ri) =>
          new TableRow({
            children: [
              cell(r[0], { width: t4Widths[0], align: AlignmentType.RIGHT, bold: ri === model.table4!.rows.length - 1 }),
              cell(r[1], { width: t4Widths[1], bold: ri === model.table4!.rows.length - 1, fill: ri === model.table4!.rows.length - 1 ? ORANGE : undefined, size: ri === model.table4!.rows.length - 1 ? 15 : BASE }),
            ],
          })
        ),
      ])
    : para("Table-4: Calculation of APC-3 - enter old/new remittance for an MF revision to compute APC-3.", { size: 13, color: GREY, align: AlignmentType.LEFT });

  const gapW = CONTENT_W - t3W - formulaW - t4W;
  const row3 = grid(
    [t3W, Math.round(gapW / 2), formulaW, Math.round(gapW / 2), t4W],
    [
      new TableRow({
        children: [
          layoutCell(t3W, [
            para(t3.title, { bold: true, size: 16, align: AlignmentType.LEFT, after: 40 }),
            table3,
            para(t3.perDay, { size: 14, align: AlignmentType.RIGHT, before: 40 }),
          ]),
          layoutCell(Math.round(gapW / 2), [para("")]),
          layoutCell(formulaW, [spacer(400), formulaBox]),
          layoutCell(Math.round(gapW / 2), [para("")]),
          layoutCell(t4W, [spacer(300), table4]),
        ],
      }),
    ],
    tableNoBorders
  );

  const children: (DocxNs.Paragraph | DocxNs.Table)[] = [
    header,
    spacer(60),
    infoRow,
    spacer(80),
    para([run(t1.title + " ", { bold: true, size: 16 }), run(t1.period, { bold: true, size: 16, color: RED })], { after: 40 }),
    table1,
    spacer(100),
    row2,
    spacer(100),
    row3,
  ];
  if (model.alert) children.push(spacer(60), para(model.alert, { bold: true, color: RED, size: 13, align: AlignmentType.LEFT }));

  return new Document({
    creator: "APC 2.0",
    title: model.title,
    styles: { default: { document: { run: { font: "Calibri", size: BASE, color: INK } } } },
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
  plaza: PlazaInfo,
  result: ApcResult,
  uptoMonth: string,
  mfEntries: MfEntry[],
  growth: number
) {
  const docx = await import("docx");
  const model = buildSheetModel(plaza, result, uptoMonth, mfEntries, growth);
  const chartPng = renderChartPng(model.chart.title, model.chart.points);
  const doc = buildApcDocx(docx, model, { chartPng });
  const blob = await docx.Packer.toBlob(doc);
  downloadBlob(blob, `${model.fileBase}.docx`);
}
