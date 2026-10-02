// Client-side Excel export of one APC calculation (inputs + Tables 1-3),
// laid out like the approved "2. APC" reference sheet.

import type { ApcResult, MfEntry } from "./apc";
import { fmtIN, fmtShort, monthLabel } from "./apc";
import { downloadBlob } from "./download";

const BOLD = { bold: true };
const NUM2 = "#,##0.00";
const NUM4 = "0.0000";
const INT = "#,##0";

export async function exportApcXlsx(
  plaza: { code: number; name: string },
  result: ApcResult,
  uptoMonth: string,
  mfEntries: MfEntry[],
  growth: number
) {
  const ExcelJS = (await import("exceljs")).default;
  const { table1, table2, table3, months } = result;

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("APC", {
    views: [{ showGridLines: true }],
  });
  ws.columns = [
    { width: 42 }, { width: 20 }, { width: 16 }, { width: 14 },
    { width: 16 }, { width: 18 }, { width: 14 }, { width: 18 },
    { width: 18 }, { width: 14 }, { width: 14 },
  ];

  const title = ws.addRow([`APC Report — ${plaza.name} (${plaza.code})`]);
  title.font = { bold: true, size: 14 };

  ws.addRow(["Calculated up to", monthLabel(uptoMonth)]);
  const imputed = table2.rows.filter((r) => r.imputed).map((r) => monthLabel(r.month));
  ws.addRow([
    "12-month window",
    `${monthLabel(months[0])} to ${monthLabel(months[11])} (${table2.monthsUsed} months with data` +
      (imputed.length ? `, ${imputed.length} filled with the average: ${imputed.join(", ")})` : ")"),
  ]);
  ws.addRow([
    "Traffic growth (%)",
    growth,
    table2.negativeTrend
      ? "Alert: ETC collection trend is negative — growth should be taken as 0%"
      : "",
  ]);
  const validMf = mfEntries.filter((e) => e.date && e.factor > 0);
  ws.addRow([
    "MF revisions",
    validMf.length
      ? validMf.map((e) => `${e.date}: ${e.factor}`).join(", ")
      : "none",
  ]);
  const hero = ws.addRow([
    "APC-2",
    `Rs ${table3.apcCr.toFixed(2)} Cr.`,
    `Rs ${Math.round(table3.apcPerDay).toLocaleString("en-IN")} per day`,
  ]);
  hero.font = { bold: true, size: 12 };

  // ---- Table 1 ----
  ws.addRow([]);
  ws.addRow(["Table 1 — Calculation of ETC Penetration (12 Completed Calendar Months)"]).font = BOLD;
  const h1 = ws.addRow([
    "Category of Vehicle", "NPCI Codes",
    "ETC Transactions (A)", "Cash + UPI", "50% of Exempted",
    "Total TMCC (B)", "% ETC Penetration (C = A×100/B)",
    "ETC Collection + AP Compensation (D)", "Derived Total (E = D×100/C)",
    "% Contribution (F = E/ΣE)", "Revenue Contribution Roundoff % (G, ΣG = 100)",
  ]);
  h1.font = BOLD;
  h1.alignment = { wrapText: true, vertical: "top" };
  for (const r of table1.rows) {
    const row = ws.addRow([
      r.category, r.npciCodes,
      r.etcTxn, r.cashUpiTxn, r.exemptHalf, r.totalTxn,
      r.penetration, r.etcCollectionWithAp, r.derivedTotal,
      r.contribution * 100, r.contributionRounded,
    ]);
    row.getCell(3).numFmt = INT;
    row.getCell(4).numFmt = INT;
    for (const c of [5, 6, 8, 9]) row.getCell(c).numFmt = NUM2;
    for (const c of [7, 10]) row.getCell(c).numFmt = "0.00";
    row.getCell(11).numFmt = '0"%"';
  }
  const t1total = ws.addRow([
    "Weighted Average ETC Penetration — Σ(C×D)/Σ(D)",
    "", "", "", "", "",
    table1.weightedPenetration, "", "",
    table1.rows.reduce((s, r) => s + r.contribution, 0) * 100,
    table1.contributionTotal,
  ]);
  t1total.font = BOLD;
  t1total.getCell(7).numFmt = "0.00";
  t1total.getCell(10).numFmt = "0.00";
  t1total.getCell(11).numFmt = '0"%"';

  // ---- Table 2 ----
  ws.addRow([]);
  ws.addRow(["Table 2 — Monthly Average Daily ETC Collection"]).font = BOLD;
  ws.addRow([
    "Annual Pass compensation is shown for information only and is not part of B. " +
      "Months marked (avg) had no data and take the average of the months that do.",
  ]);
  const h2 = ws.addRow([
    "Month", "Avg Daily AP Compensation (info)", "Avg Daily ETC Collection (A)",
    "MF Multiplier", "ETC at revised fee rates (B = A × MF)",
  ]);
  h2.font = BOLD;
  h2.alignment = { wrapText: true, vertical: "top" };
  for (const r of table2.rows) {
    const filled = r.etcDaily !== null;
    const row = ws.addRow([
      monthLabel(r.month) + (r.imputed ? " (avg)" : ""),
      filled ? r.apDaily : "—",
      filled ? r.etcDaily : "—",
      r.mfMultiplier,
      filled ? r.normalizedDaily : "no data",
    ]);
    for (const c of [2, 3, 5]) row.getCell(c).numFmt = NUM2;
    row.getCell(4).numFmt = NUM4;
    if (r.imputed) row.font = { italic: true, color: { argb: "FF52514E" } };
  }
  const t2avg = ws.addRow([
    "Average",
    table2.avgApDaily, table2.avgEtcDaily, "", table2.avgNormalizedDaily,
  ]);
  t2avg.font = BOLD;
  for (const c of [2, 3, 5]) t2avg.getCell(c).numFmt = NUM2;

  // ---- Table 3 ----
  ws.addRow([]);
  ws.addRow(["Table 3 — Calculation of APC-2"]).font = BOLD;
  const h3 = ws.addRow(["#", "Item", "Value (Rs)"]);
  h3.font = BOLD;
  // Rupee values carry their short form inline, e.g. 3,40,14,000 (3.40 Cr.)
  const t3: [number, string, string][] = [
    [1, "Average Daily FASTag Collection", money(table3.avgDailyFastagCollection)],
    [2, "FASTag Penetration (in %)", table3.fastagPenetration.toFixed(2)],
    [3, "Annual Average Daily Collection", money(table3.annualAvgDailyCollection)],
    [4, "Annual Expected Collection", money(table3.annualExpectedCollection)],
    [5, "Traffic Growth (in %)", `${table3.trafficGrowthPct}%`],
    [6, "Net Expected Collection", money(table3.netExpectedCollection)],
    [7, "Less Administrative Charges", money(table3.adminCharges)],
    [8, "Less Contractor Profit @5%", money(table3.contractorProfit)],
    [9, "APC-2", `${money(table3.apcCr * 1e7)} = Rs ${fmtIN(table3.apcPerDay)} per day`],
  ];
  for (const [n, label, value] of t3) {
    const row = ws.addRow([n, label, value]);
    row.getCell(3).alignment = { horizontal: "right" };
    if (n === 9) row.font = BOLD;
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  downloadBlob(blob, `APC_${plaza.code}_${uptoMonth}.xlsx`);
}

function money(n: number): string {
  return Math.abs(n) < 1e3 ? fmtIN(n, 0) : `${fmtIN(n, 0)} (${fmtShort(n)})`;
}
