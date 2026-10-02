// PDF export of one APC calculation (A4 landscape, structured tables),
// rendered from the shared ReportModel with jsPDF + jspdf-autotable.

import type { jsPDF } from "jspdf";
import type { UserOptions } from "jspdf-autotable";
import type { ApcResult, MfEntry } from "./apc";
import { buildReportModel, type ReportModel, type ReportTable } from "./reportModel";
import { downloadBlob } from "./download";

type AutoTableFn = (doc: jsPDF, options: UserOptions) => void;

const PAGE_W = 297;
const PAGE_H = 210;
const MARGIN = 12;
const CONTENT_W = PAGE_W - 2 * MARGIN;
const INK: [number, number, number] = [11, 11, 11];
const MUTED: [number, number, number] = [82, 81, 78];
const HEAD_FILL: [number, number, number] = [31, 56, 100];
const FOOT_FILL: [number, number, number] = [236, 239, 244];
const DANGER: [number, number, number] = [208, 59, 59];
const TOTAL_PAGES = "{total_pages}";

function finalY(doc: jsPDF): number {
  const d = doc as unknown as { lastAutoTable?: { finalY?: number } };
  return d.lastAutoTable?.finalY ?? MARGIN;
}

/** Build the PDF document (no download). Exposed for tests. */
export function buildApcPdf(jsPDFCtor: typeof jsPDF, autoTable: AutoTableFn, model: ReportModel): jsPDF {
  const doc = new jsPDFCtor({ orientation: "landscape", unit: "mm", format: "a4" });
  doc.setProperties({ title: model.title, subject: "APC-2 calculation", creator: "APC 2.0" });

  let y = MARGIN;

  // ---- Title block ----
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(...INK);
  doc.text(model.title, MARGIN, y + 5);
  y += 9;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(...MUTED);
  doc.text(model.subtitle, MARGIN, y + 3);
  y += 8;

  // ---- Inputs (left) and APC-2 highlight (right) ----
  const metaW = CONTENT_W * 0.62;
  autoTable(doc, {
    startY: y,
    margin: { left: MARGIN, right: PAGE_W - MARGIN - metaW },
    tableWidth: metaW,
    theme: "plain",
    styles: { fontSize: 9, cellPadding: { top: 1.2, bottom: 1.2, left: 0, right: 3 }, textColor: INK },
    columnStyles: { 0: { fontStyle: "bold", cellWidth: 42, textColor: MUTED } },
    body: model.meta,
  });
  const metaEnd = finalY(doc);

  const heroX = MARGIN + metaW + 8;
  const heroW = CONTENT_W - metaW - 8;
  const heroH = 26;
  doc.setDrawColor(...HEAD_FILL);
  doc.setFillColor(245, 247, 251);
  doc.roundedRect(heroX, y, heroW, heroH, 2, 2, "FD");
  doc.setFontSize(9);
  doc.setTextColor(...MUTED);
  doc.text(model.hero.label, heroX + 5, y + 7);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.setTextColor(...HEAD_FILL);
  doc.text(model.hero.value, heroX + 5, y + 16);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.setTextColor(...MUTED);
  doc.text(model.hero.perDay, heroX + 5, y + 22);

  y = Math.max(metaEnd, y + heroH) + 4;

  // ---- Alert ----
  if (model.alert) {
    doc.setFontSize(9.5);
    doc.setTextColor(...DANGER);
    const lines = doc.splitTextToSize(model.alert, CONTENT_W - 8) as string[];
    const boxH = lines.length * 4.6 + 5;
    doc.setDrawColor(...DANGER);
    doc.setFillColor(253, 243, 243);
    doc.roundedRect(MARGIN, y, CONTENT_W, boxH, 1.5, 1.5, "FD");
    doc.text(lines, MARGIN + 4, y + 5.5);
    y += boxH + 5;
    doc.setTextColor(...INK);
  }

  // ---- Tables ----
  y = drawTable(doc, autoTable, model.table1, y);
  y = drawTable(doc, autoTable, model.table2, y);
  y = drawTable(doc, autoTable, model.table3, y, CONTENT_W * 0.7);

  // ---- Notes ----
  y = ensureRoom(doc, y, 30);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(9.5);
  doc.setTextColor(...INK);
  doc.text("Notes", MARGIN, y + 4);
  y += 7;
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(...MUTED);
  for (const n of model.notes) {
    const lines = doc.splitTextToSize(n, CONTENT_W - 4) as string[];
    y = ensureRoom(doc, y, lines.length * 4 + 2);
    doc.text("-", MARGIN, y + 3);
    doc.text(lines, MARGIN + 4, y + 3);
    y += lines.length * 4 + 1.5;
  }

  // ---- Footer on every page ----
  const pages = doc.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(...MUTED);
    doc.text(model.title, MARGIN, PAGE_H - 6);
    doc.text(`Page ${p} of ${TOTAL_PAGES}`, PAGE_W - MARGIN, PAGE_H - 6, { align: "right" });
  }
  doc.putTotalPages(TOTAL_PAGES);

  return doc;
}

function ensureRoom(doc: jsPDF, y: number, needed: number): number {
  if (y + needed > PAGE_H - MARGIN - 6) {
    doc.addPage();
    return MARGIN;
  }
  return y;
}

function drawTable(
  doc: jsPDF,
  autoTable: AutoTableFn,
  t: ReportTable,
  y: number,
  width: number = CONTENT_W
): number {
  // Keep the whole table on one page when it can fit; otherwise keep the
  // heading with at least the header and a couple of rows.
  const estHeight = (t.body.length + (t.foot?.length ?? 0) + 2) * 8.5 + 20;
  y = ensureRoom(doc, y, Math.min(estHeight, PAGE_H - 2 * MARGIN - 10));
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(...INK);
  doc.text(t.title, MARGIN, y + 4);
  y += 7;
  if (t.note) {
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    doc.setTextColor(...MUTED);
    const lines = doc.splitTextToSize(t.note, width) as string[];
    doc.text(lines, MARGIN, y + 3);
    y += lines.length * 3.8 + 2;
  }

  const columnStyles: UserOptions["columnStyles"] = {};
  const totalW = t.widths?.reduce((s, w) => s + w, 0) ?? 0;
  t.head.forEach((_, i) => {
    const style: Record<string, unknown> = {};
    if (t.numericCols.includes(i)) style.halign = "right";
    if (t.widths && totalW > 0) style.cellWidth = (t.widths[i] / totalW) * width;
    columnStyles[i] = style;
  });

  autoTable(doc, {
    startY: y,
    margin: { left: MARGIN, right: PAGE_W - MARGIN - width, bottom: MARGIN + 4 },
    tableWidth: width,
    theme: "grid",
    styles: {
      font: "helvetica",
      fontSize: 8,
      cellPadding: 1.6,
      lineColor: [195, 194, 183],
      lineWidth: 0.2,
      textColor: INK,
      valign: "middle",
      overflow: "linebreak",
    },
    headStyles: {
      fillColor: HEAD_FILL,
      textColor: [255, 255, 255],
      fontStyle: "bold",
      halign: "center",
      valign: "middle",
      fontSize: 7.5,
    },
    footStyles: {
      fillColor: FOOT_FILL,
      textColor: INK,
      fontStyle: "bold",
    },
    alternateRowStyles: { fillColor: [250, 250, 248] },
    columnStyles,
    head: [t.head],
    body: t.body,
    foot: t.foot,
    showFoot: "lastPage",
    didParseCell: (data) => {
      // Numeric columns in the footer stay right-aligned; label cells left.
      if (data.section === "foot" && t.numericCols.includes(data.column.index)) {
        data.cell.styles.halign = "right";
      }
      if (t.emphasizeLastRow && data.section === "body" && data.row.index === t.body.length - 1) {
        data.cell.styles.fontStyle = "bold";
        data.cell.styles.fillColor = FOOT_FILL;
      }
      if (data.section === "body" && data.cell.text.join("").includes("(avg)")) {
        data.cell.styles.textColor = MUTED;
        data.cell.styles.fontStyle = "italic";
      }
    },
  });

  return finalY(doc) + 8;
}

export async function exportApcPdf(
  plaza: { code: number; name: string },
  result: ApcResult,
  uptoMonth: string,
  mfEntries: MfEntry[],
  growth: number
) {
  const [{ jsPDF }, { default: autoTable }] = await Promise.all([
    import("jspdf"),
    import("jspdf-autotable"),
  ]);
  const model = buildReportModel(plaza, result, uptoMonth, mfEntries, growth);
  const doc = buildApcPdf(jsPDF, autoTable as AutoTableFn, model);
  downloadBlob(doc.output("blob"), `${model.fileBase}.pdf`);
}
