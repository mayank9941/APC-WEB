// PDF export of one APC calculation as a single A4-landscape page laid out
// like NHAI's "APC of <plaza> Toll Plaza as per new methodology" sheet:
// header bar + plaza/MF block, Table-1 (with lettered formula row and the
// rounded % column outside), Table-2 beside the line chart, Table-3 with the
// admin-expense formula box, and Table-4 (APC-3) at the bottom right.

import type { jsPDF } from "jspdf";
import type { UserOptions } from "jspdf-autotable";
import type { ApcResult, MfEntry, PlazaInfo } from "./apc";
import { fmtIN } from "./apc";
import { buildSheetModel, type SheetChartPoint, type SheetModel } from "./reportModel";
import { downloadBlob } from "./download";

type AutoTableFn = (doc: jsPDF, options: UserOptions) => void;
type RGB = [number, number, number];

const PAGE_W = 297;
const PAGE_H = 210;
const M = 8; // page margin
const RIGHT = PAGE_W - M;

const INK: RGB = [0, 0, 0];
const GREY: RGB = [110, 110, 110];
const RED: RGB = [200, 0, 0];
const LINE: RGB = [120, 120, 120];
const PEACH: RGB = [252, 228, 214];
const PEACH_LINE: RGB = [244, 177, 131];
const YELLOW: RGB = [255, 255, 0];
const ORANGE: RGB = [255, 192, 0];
const BLUE_FILL: RGB = [221, 235, 247];
const SERIES: RGB = [31, 78, 121];

const BASE_FONT = 6;
const PAD = 0.5;
const BOTTOM = 4; // autotable bottom margin so tables use the whole page

function finalY(doc: jsPDF): number {
  const d = doc as unknown as { lastAutoTable?: { finalY?: number } };
  return d.lastAutoTable?.finalY ?? 0;
}

function baseStyles(): Partial<UserOptions["styles"]> {
  return {
    font: "helvetica",
    fontSize: BASE_FONT,
    cellPadding: PAD,
    lineColor: LINE,
    lineWidth: 0.15,
    textColor: INK,
    valign: "middle",
    halign: "center",
    overflow: "linebreak",
  };
}

function label(
  doc: jsPDF, text: string, x: number, y: number, size = 7.5, bold = true, color: RGB = INK,
  align: "left" | "center" | "right" = "left"
) {
  doc.setFont("helvetica", bold ? "bold" : "normal");
  doc.setFontSize(size);
  doc.setTextColor(...color);
  doc.text(text, x, y, { align });
  return doc.getTextWidth(text);
}

/** Build the PDF document (no download). Exposed for tests. */
export function buildApcPdf(jsPDFCtor: typeof jsPDF, autoTable: AutoTableFn, model: SheetModel): jsPDF {
  const doc = new jsPDFCtor({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setProperties({ title: model.title, subject: "APC-2 / APC-3 calculation", creator: "APC 2.0" });

  // ---------- Title bar + "Calculated on" ----------
  let y = M;
  doc.setFillColor(...PEACH);
  doc.rect(M, y, 196, 7, "F");
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(...INK);
  doc.text(model.title, M + 98, y + 4.8, { align: "center" });

  autoTable(doc, {
    startY: y,
    margin: { left: RIGHT - 54, bottom: BOTTOM },
    tableWidth: 54,
    theme: "grid",
    styles: { ...baseStyles(), fontSize: 7 },
    columnStyles: { 0: { cellWidth: 24, fontStyle: "bold" }, 1: { cellWidth: 30 } },
    body: [["Calculated on", model.calculatedOn]],
  });

  // ---------- Plaza code / MF block / plaza info ----------
  y += 10;
  label(doc, "Plaza code:", M, y + 3.5, 7.5, false);
  doc.setFillColor(...BLUE_FILL);
  doc.rect(M + 17, y, 16, 4.6, "F");
  label(doc, model.plazaCode, M + 25, y + 3.5, 8, true, RED, "center");
  label(doc, model.plazaName, M + 8, y + 9, 7, false);

  const mfCount = Math.max(model.mf.labels.length, 1);
  const mfColW = 19;
  const mfLabelW = 24;
  const mfW = mfLabelW + mfColW * mfCount;
  const mfX = 72;
  const mfLabels = model.mf.labels.length ? model.mf.labels : ["MF"];
  const mfRow = (name: string, vals: string[]) => [name, ...(vals.length ? vals : ["-"])];
  const mfColumnStyles: UserOptions["columnStyles"] = { 0: { cellWidth: mfLabelW, halign: "right", fontStyle: "bold" } };
  for (let i = 0; i < mfCount; i++) mfColumnStyles[i + 1] = { cellWidth: mfColW };
  autoTable(doc, {
    startY: y - 2,
    margin: { left: mfX, bottom: BOTTOM },
    tableWidth: mfW,
    theme: "grid",
    styles: { ...baseStyles(), fontSize: 7 },
    columnStyles: mfColumnStyles,
    head: [["", ...mfLabels]],
    headStyles: { fillColor: [255, 255, 255], textColor: INK, fontStyle: "bold", lineWidth: 0.15, lineColor: LINE },
    body: [
      mfRow("Old remittance", model.mf.oldRemittance),
      mfRow("New remittance", model.mf.newRemittance),
      mfRow("MF", model.mf.factor),
    ],
    didParseCell: (d) => {
      if (d.section === "head" && d.column.index === 0) d.cell.styles.lineWidth = 0;
      if (d.section === "body" && d.column.index === 0) {
        d.cell.styles.lineWidth = 0;
        if (d.row.index < 2) d.cell.styles.textColor = RED;
      }
      if (d.section === "body" && d.column.index > 0 && d.row.index < 2) d.cell.styles.fillColor = YELLOW;
      if (d.section === "body" && d.row.index === 2) d.cell.styles.fontStyle = "bold";
    },
  });
  const mfEnd = finalY(doc);

  autoTable(doc, {
    startY: y - 2,
    margin: { left: RIGHT - 54, bottom: BOTTOM },
    tableWidth: 54,
    theme: "grid",
    styles: { ...baseStyles(), fontSize: 7 },
    columnStyles: { 0: { cellWidth: 24, halign: "right", fontStyle: "bold" }, 1: { cellWidth: 30 } },
    body: model.plazaInfo,
  });
  y = Math.max(mfEnd, finalY(doc)) + 3;

  // ---------- Table-1 ----------
  const t1 = model.table1;
  const t1W = 238;
  const w = label(doc, t1.title + " ", M + t1W / 2 - 30, y + 3, 8);
  label(doc, t1.period, M + t1W / 2 - 30 + w, y + 3, 8, true, RED);
  y += 5;

  const t1Widths = [40, 27, 20, 18, 18, 26, 16, 28, 22, 23]; // sums to 238
  const t1Cols: UserOptions["columnStyles"] = {};
  t1Widths.forEach((cw, i) => (t1Cols[i] = { cellWidth: cw }));
  const roundX = M + t1W + 2;
  const roundW = 12;
  autoTable(doc, {
    startY: y,
    margin: { left: M, bottom: BOTTOM },
    tableWidth: t1W,
    theme: "grid",
    styles: baseStyles(),
    columnStyles: t1Cols,
    head: [t1.head, t1.letters],
    headStyles: { fillColor: [255, 255, 255], textColor: INK, fontStyle: "bold", lineWidth: 0.15, lineColor: LINE },
    body: [
      ...t1.body,
      [{ content: t1.weightedLabel, colSpan: 6, styles: { halign: "right", fontStyle: "bold" } }, t1.weighted, "", "", ""],
    ],
    didParseCell: (d) => {
      if (d.section === "head" && d.row.index === 1) d.cell.styles.fontStyle = "normal";
      if (d.section === "body" && d.row.index < t1.body.length) {
        if (d.column.index === 2 || d.column.index === 7 || d.column.index === 8) d.cell.styles.fillColor = BLUE_FILL;
        if (d.column.index === 3 || d.column.index === 4) d.cell.styles.fillColor = YELLOW;
      }
      if (d.section === "body" && d.row.index === t1.body.length) {
        if (d.column.index === 6) {
          d.cell.styles.fillColor = ORANGE;
          d.cell.styles.fontStyle = "bold";
          d.cell.styles.fontSize = 7.5;
        }
        if (d.column.index >= 7) d.cell.styles.lineWidth = 0;
      }
    },
    // Rounded % column written outside the bordered table, aligned to each row.
    didDrawCell: (d) => {
      if (d.section !== "body" || d.column.index !== 9) return;
      const cy = d.cell.y + d.cell.height / 2 + 1;
      if (d.row.index < t1.body.length) {
        label(doc, t1.rounded[d.row.index], roundX + roundW, cy, BASE_FONT, false, INK, "right");
      } else {
        doc.setFillColor(...YELLOW);
        doc.rect(roundX, d.cell.y + 0.3, roundW, d.cell.height - 0.6, "F");
        label(doc, t1.roundedTotal, roundX + roundW - 0.8, cy, BASE_FONT + 0.5, true, INK, "right");
      }
    },
  });
  y = finalY(doc) + 3;

  // ---------- Table-2 (left) + chart (right) ----------
  const t2 = model.table2;
  const t2W = 108;
  label(doc, t2.title, M + 2, y + 3, 8);
  const t2Top = y + 5;
  autoTable(doc, {
    startY: t2Top,
    margin: { left: M, bottom: BOTTOM },
    tableWidth: t2W,
    theme: "grid",
    styles: baseStyles(),
    columnStyles: { 0: { cellWidth: 16, fontStyle: "bold" }, 1: { cellWidth: 26 }, 2: { cellWidth: 32 }, 3: { cellWidth: 34 } },
    head: [t2.head],
    headStyles: { fillColor: [255, 255, 255], textColor: INK, fontStyle: "bold", lineWidth: 0.15, lineColor: LINE },
    body: [...t2.body, t2.average],
    didParseCell: (d) => {
      if (d.section === "head" && d.column.index === 1) d.cell.styles.textColor = RED;
      if (d.section === "body") {
        const isAvg = d.row.index === t2.body.length;
        if (d.column.index === 2 && !isAvg) d.cell.styles.fillColor = BLUE_FILL;
        if (isAvg) {
          d.cell.styles.fontStyle = "bold";
          if (d.column.index === 3) {
            d.cell.styles.fillColor = ORANGE;
            d.cell.styles.fontSize = 7.5;
          }
        }
        if (d.cell.text.join("").endsWith("*")) d.cell.styles.textColor = GREY;
      }
    },
  });
  const t2End = finalY(doc);
  if (t2.imputedNote) label(doc, t2.imputedNote, M, t2End + 2.5, 5.5, false, GREY);

  const chartX = M + t2W + 8;
  const chartW = RIGHT - chartX;
  drawChart(doc, model.chart.title, model.chart.points, chartX, y, chartW, t2End - y);

  // ---------- Table-3 (left), formula box (middle), Table-4 (right) ----------
  y = t2End + (t2.imputedNote ? 5 : 3);
  const t3 = model.table3;
  const t3W = 108;
  label(doc, t3.title, M + 2, y + 3, 8);
  const t3Top = y + 5;
  autoTable(doc, {
    startY: t3Top,
    margin: { left: M, bottom: BOTTOM },
    tableWidth: t3W,
    theme: "grid",
    styles: baseStyles(),
    columnStyles: { 0: { cellWidth: 10 }, 1: { cellWidth: 52 }, 2: { cellWidth: 46, halign: "right" } },
    body: t3.rows,
    didParseCell: (d) => {
      if (d.row.index === t3.rows.length - 1) {
        d.cell.styles.fontStyle = "bold";
        if (d.column.index === 2) {
          d.cell.styles.fillColor = ORANGE;
          d.cell.styles.fontSize = 7.5;
        }
      }
    },
  });
  const t3End = finalY(doc);
  label(doc, t3.perDay, M + t3W, t3End + 3, 7, false, INK, "right");

  // Formula box
  const fx = M + t3W + 10;
  const fw = 72;
  const fTop = t3Top + 6;
  doc.setDrawColor(...LINE);
  doc.setLineWidth(0.3);
  doc.rect(fx, fTop, fw, 24);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(7.5);
  doc.setTextColor(...INK);
  doc.text(model.formula.heading, fx + 3, fTop + 5);
  doc.setLineWidth(0.2);
  doc.line(fx + 3, fTop + 6, fx + 3 + doc.getTextWidth(model.formula.heading), fTop + 6);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(7);
  model.formula.lines.forEach((ln, i) => doc.text(ln, fx + 3, fTop + 11 + i * 4.2));

  // Table-4
  const t4W = 78;
  const t4X = RIGHT - t4W;
  if (model.table4) {
    autoTable(doc, {
      startY: t3Top,
      margin: { left: t4X, bottom: BOTTOM },
      tableWidth: t4W,
      theme: "grid",
      styles: baseStyles(),
      columnStyles: { 0: { cellWidth: 50, halign: "right" }, 1: { cellWidth: 28 } },
      head: [[{ content: model.table4.title, colSpan: 2, styles: { halign: "center", fontStyle: "bold" } }]],
      headStyles: { fillColor: [255, 255, 255], textColor: INK, lineWidth: 0.15, lineColor: LINE },
      body: model.table4.rows,
      didParseCell: (d) => {
        if (d.section === "body" && d.row.index === model.table4!.rows.length - 1) {
          d.cell.styles.fontStyle = "bold";
          if (d.column.index === 1) {
            d.cell.styles.fillColor = ORANGE;
            d.cell.styles.fontSize = 7.5;
          }
        }
      },
    });
  } else {
    label(doc, "Table-4: Calculation of APC-3", t4X, t3Top + 3, 7.5);
    label(doc, "Enter old/new remittance for an MF revision to compute APC-3.", t4X, t3Top + 8, 6.5, false, GREY);
  }

  // ---------- Alert / footer ----------
  const footY = PAGE_H - 4;
  if (model.alert) label(doc, model.alert, M, footY, 6.5, true, RED);
  label(doc, "APC 2.0", RIGHT, footY, 6, false, GREY, "right");

  return doc;
}

/** Line chart drawn with jsPDF primitives, styled like the sheet's Excel chart. */
function drawChart(doc: jsPDF, title: string, points: SheetChartPoint[], x: number, y: number, w: number, h: number) {
  doc.setDrawColor(...PEACH_LINE);
  doc.setLineWidth(0.6);
  doc.rect(x, y, w, h);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(8.5);
  doc.setTextColor(...RED);
  doc.text(title, x + w / 2, y + 6, { align: "center" });

  const values = points.map((p) => p.value).filter((v): v is number => v !== null && Number.isFinite(v));
  const maxV = values.length ? Math.max(...values) : 0;
  const step = niceStep(maxV);
  const top = step > 0 ? Math.ceil(maxV / step) * step : 1;
  const ticks = step > 0 ? Math.round(top / step) : 1;

  const left = x + 20;
  const right = x + w - 6;
  const bottom = y + h - 10;
  const plotTop = y + 11;
  const ph = bottom - plotTop;
  const pw = right - left;

  // Gridlines + y labels
  doc.setFont("helvetica", "normal");
  doc.setFontSize(5.5);
  doc.setTextColor(...GREY);
  doc.setDrawColor(217, 217, 217);
  doc.setLineWidth(0.15);
  for (let i = 0; i <= ticks; i++) {
    const v = (top / ticks) * i;
    const gy = bottom - (ph * i) / ticks;
    doc.line(left, gy, right, gy);
    doc.text(fmtIN(v, 0), left - 1.5, gy + 1, { align: "right" });
  }
  doc.setDrawColor(...GREY);
  doc.setLineWidth(0.25);
  doc.line(left, bottom, right, bottom);

  // x labels
  const n = points.length;
  const xAt = (i: number) => left + (pw * (i + 0.5)) / n;
  points.forEach((p, i) => doc.text(p.label, xAt(i), bottom + 3.5, { align: "center" }));

  // Series
  doc.setDrawColor(...SERIES);
  doc.setLineWidth(0.5);
  let prev: [number, number] | null = null;
  points.forEach((p, i) => {
    if (p.value === null || !Number.isFinite(p.value) || top <= 0) {
      prev = null;
      return;
    }
    const cx = xAt(i);
    const cy = bottom - (ph * p.value) / top;
    if (prev) doc.line(prev[0], prev[1], cx, cy);
    prev = [cx, cy];
  });
  points.forEach((p, i) => {
    if (p.value === null || !Number.isFinite(p.value) || top <= 0) return;
    const cx = xAt(i);
    const cy = bottom - (ph * p.value) / top;
    if (p.imputed) {
      doc.setFillColor(255, 255, 255);
      doc.circle(cx, cy, 0.9, "FD");
    } else {
      doc.setFillColor(...SERIES);
      doc.circle(cx, cy, 0.9, "F");
    }
  });
}

function niceStep(maxV: number): number {
  if (maxV <= 0) return 0;
  const raw = maxV / 6;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const nice = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
  return nice * mag;
}

export async function exportApcPdf(
  plaza: PlazaInfo,
  result: ApcResult,
  uptoMonth: string,
  mfEntries: MfEntry[],
  growth: number
) {
  const [{ jsPDF }, { default: autoTable }] = await Promise.all([
    import("jspdf"),
    import("jspdf-autotable"),
  ]);
  const model = buildSheetModel(plaza, result, uptoMonth, mfEntries, growth);
  const doc = buildApcPdf(jsPDF, autoTable as AutoTableFn, model);
  downloadBlob(doc.output("blob"), `${model.fileBase}.pdf`);
}
