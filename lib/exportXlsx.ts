// Client-side Excel export of one APC calculation (inputs + Tables 1-3),
// laid out like the approved "2. APC" reference sheet.

import type { ApcResult, MfEntry } from "./apc";
import { monthLabel } from "./apc";

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
  ws.addRow([
    "12-month window",
    `${monthLabel(months[0])} to ${monthLabel(months[11])} (${table2.monthsUsed} months with data)`,
  ]);
  ws.addRow(["Traffic growth (%)", growth]);
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
    "% Contribution (F = E/ΣE)", "Roundoff (G)",
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
    for (const c of [7, 10, 11]) row.getCell(c).numFmt = "0.00";
  }
  const t1total = ws.addRow([
    "Weighted Average ETC Penetration — Σ(C×D)/Σ(D)",
    "", "", "", "", "",
    table1.weightedPenetration, "", "", "",
    table1.contributionTotal,
  ]);
  t1total.font = BOLD;
  t1total.getCell(7).numFmt = "0.00";
  t1total.getCell(11).numFmt = "0.00";

  // ---- Table 2 ----
  ws.addRow([]);
  ws.addRow(["Table 2 — Monthly Average Daily Collection (Seasonal Factor)"]).font = BOLD;
  const h2 = ws.addRow([
    "Month", "Avg Daily AP Compensation", "Avg Daily ETC Collection",
    "Avg Daily ETC + AP (A)", "MF Multiplier", "Normalized (B = A × MF)",
  ]);
  h2.font = BOLD;
  h2.alignment = { wrapText: true, vertical: "top" };
  for (const r of table2.rows) {
    const row = ws.addRow([
      monthLabel(r.month),
      r.hasData ? r.apDaily : "—",
      r.hasData ? r.etcDaily : "—",
      r.hasData ? r.totalDaily : "—",
      r.mfMultiplier,
      r.hasData ? r.normalizedDaily : "no data",
    ]);
    for (const c of [2, 3, 4, 6]) row.getCell(c).numFmt = NUM2;
    row.getCell(5).numFmt = NUM4;
  }
  const t2avg = ws.addRow([
    "Average",
    table2.avgApDaily, "", table2.avgTotalDaily, "", table2.avgNormalizedDaily,
  ]);
  t2avg.font = BOLD;
  for (const c of [2, 4, 6]) t2avg.getCell(c).numFmt = NUM2;

  // ---- Table 3 ----
  ws.addRow([]);
  ws.addRow(["Table 3 — Calculation of APC-2"]).font = BOLD;
  const t3: [number, string, number | string, string][] = [
    [1, "Average Daily FASTag Collection", table3.avgDailyFastagCollection, ""],
    [2, "FASTag Penetration (In decimal)", table3.fastagPenetration, ""],
    [3, "Seasonal Factor (In decimal)", "-", ""],
    [4, "Annual Average Daily Collection", table3.annualAvgDailyCollection, ""],
    [5, "Annual Expected Collection", table3.annualExpectedCollection, cr(table3.annualExpectedCollection)],
    [6, "Traffic Growth (in %)", table3.trafficGrowthPct, ""],
    [7, "Net Expected Collection", table3.netExpectedCollection, cr(table3.netExpectedCollection)],
    [8, "Less Administrative Charges", table3.adminCharges, cr(table3.adminCharges)],
    [9, "Less Contractor Profit @5%", table3.contractorProfit, cr(table3.contractorProfit)],
    [10, "APC-2 (Cr)", table3.apcCr, `${table3.apcCr.toFixed(2)} Cr.`],
  ];
  for (const [n, label, value, inCr] of t3) {
    const row = ws.addRow([n, label, value, inCr]);
    if (typeof value === "number") row.getCell(3).numFmt = NUM2;
    if (n === 10) row.font = BOLD;
  }

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `APC_${plaza.code}_${uptoMonth}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

function cr(n: number): string {
  return `${(n / 1e7).toFixed(2)} Cr.`;
}
