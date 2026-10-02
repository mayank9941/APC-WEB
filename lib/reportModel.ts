// A document-neutral view of one APC calculation: every value already
// formatted as text, laid out as titles, tables and notes. The PDF and Word
// exports render this model so they stay identical in structure and wording.
// Text is kept to characters available in the PDF's standard fonts
// (no rupee sign, Greek letters or typographic dashes).

import type { ApcResult, MfEntry } from "./apc";
import { fmtIN, fmtMoney, monthLabel } from "./apc";

export interface ReportTable {
  title: string;
  note?: string;
  head: string[];
  body: string[][];
  foot?: string[][];
  /** Column indexes that hold numbers (right-aligned). */
  numericCols: number[];
  /** Relative column widths (same length as head); optional. */
  widths?: number[];
  /** Render the last body row as a highlighted total (Table 3's APC-2 row). */
  emphasizeLastRow?: boolean;
}

export interface ReportModel {
  title: string;
  subtitle: string;
  meta: [string, string][];
  hero: { label: string; value: string; perDay: string };
  /** Growth-assumption alert text, or null when the trend is not negative. */
  alert: string | null;
  table1: ReportTable;
  table2: ReportTable;
  table3: ReportTable;
  notes: string[];
  /** File name without extension. */
  fileBase: string;
}

export function buildReportModel(
  plaza: { code: number; name: string },
  result: ApcResult,
  uptoMonth: string,
  mfEntries: MfEntry[],
  growth: number
): ReportModel {
  const { table1, table2, table3, months } = result;
  const first = monthLabel(months[0]);
  const last = monthLabel(months[11]);
  const validMf = mfEntries.filter((e) => e.date && e.factor > 0);
  const imputedMonths = table2.rows.filter((r) => r.imputed).map((r) => monthLabel(r.month));

  const windowText =
    `${first} to ${last} (${table2.monthsUsed} month${table2.monthsUsed === 1 ? "" : "s"} with data` +
    (table2.monthsImputed > 0
      ? `, ${table2.monthsImputed} filled with the average: ${imputedMonths.join(", ")})`
      : ")");

  const alert = table2.negativeTrend
    ? `Traffic growth should be taken as 0%: the average daily ETC collection shows a ` +
      `negative trend over ${first} - ${last} (about Rs ${fmtIN(Math.abs(table2.trendSlope), 0)} ` +
      `per day less each month).` +
      (growth > 0 ? ` This report uses ${growth}%.` : " This report uses 0%.")
    : null;

  const fTotal = table1.rows.reduce((s, r) => s + r.contribution, 0) * 100;

  const t1: ReportTable = {
    title: "Table 1 - Calculation of ETC Penetration (Period: 12 Completed Calendar Months)",
    note: `Summed over ${first} - ${last}`,
    head: [
      "Category of Vehicle",
      "NPCI Codes",
      "Total ETC Transactions (A)",
      "Cash + UPI Transactions",
      "50% of Exempted",
      "Total Transactions as per TMCC (B)",
      "% ETC Penetration (C = A x 100 / B)",
      "ETC Collection + Annual Pass Compensation (D)",
      "Derived Total Collection (E = D x 100 / C)",
      "% Contribution (F = E / Sum of E)",
      "Revenue Contribution Roundoff % (G)",
    ],
    body: table1.rows.map((r) => [
      r.category,
      r.npciCodes,
      fmtIN(r.etcTxn),
      fmtIN(r.cashUpiTxn),
      fmtIN(r.exemptHalf, 1),
      fmtIN(r.totalTxn, 1),
      r.penetration.toFixed(2),
      fmtIN(r.etcCollectionWithAp),
      fmtIN(r.derivedTotal),
      `${(r.contribution * 100).toFixed(2)}%`,
      `${r.contributionRounded}%`,
    ]),
    foot: [[
      "Weighted Average ETC Penetration = Sum(C x D) / Sum(D)",
      "", "", "", "", "",
      table1.weightedPenetration.toFixed(2),
      "", "",
      `${fTotal.toFixed(2)}%`,
      `${table1.contributionTotal}%`,
    ]],
    numericCols: [2, 3, 4, 5, 6, 7, 8, 9, 10],
    widths: [18.5, 9, 8.5, 7.5, 7, 9, 8.5, 10, 10, 8.5, 8.5],
  };

  const t2: ReportTable = {
    title: "Table 2 - Monthly Average Daily ETC Collection",
    note:
      "Annual Pass compensation is shown for information only and is not part of B. " +
      "Months marked (avg) had no data and take the average of the months that do.",
    head: [
      "Month",
      "Monthly Average Daily Annual Pass Compensation (info)",
      "ETC Collection, average daily (A)",
      "MF Multiplier",
      "Monthly Avg ETC Collection as per revised Fee Rates (B = A x MF)",
    ],
    body: table2.rows.map((r) => [
      monthLabel(r.month) + (r.imputed ? " (avg)" : ""),
      r.etcDaily === null ? "-" : fmtIN(r.apDaily, 2),
      r.etcDaily === null ? "-" : fmtIN(r.etcDaily, 2),
      r.mfMultiplier.toFixed(4),
      r.etcDaily === null ? "no data" : fmtIN(r.normalizedDaily, 2),
    ]),
    foot: [[
      "Average",
      fmtIN(table2.avgApDaily, 2),
      fmtIN(table2.avgEtcDaily, 2),
      "",
      fmtIN(table2.avgNormalizedDaily, 2),
    ]],
    numericCols: [1, 2, 3, 4],
    widths: [14, 22, 20, 14, 30],
  };

  const t3: ReportTable = {
    title: "Table 3 - Calculation of APC-2",
    head: ["#", "Item", "Value (Rs)"],
    body: [
      ["1", "Average Daily FASTag Collection", fmtMoney(table3.avgDailyFastagCollection)],
      ["2", "FASTag Penetration (in %)", table3.fastagPenetration.toFixed(2)],
      ["3", "Annual Average Daily Collection", fmtMoney(table3.annualAvgDailyCollection)],
      ["4", "Annual Expected Collection", fmtMoney(table3.annualExpectedCollection)],
      ["5", "Traffic Growth (in %)", `${table3.trafficGrowthPct}%`],
      ["6", "Net Expected Collection", fmtMoney(table3.netExpectedCollection)],
      ["7", "Less Administrative Charges", fmtMoney(table3.adminCharges)],
      ["8", "Less Contractor Profit @5%", fmtMoney(table3.contractorProfit)],
      [
        "9",
        "APC-2",
        `${fmtMoney(table3.apcCr * 1e7)} = Rs ${fmtIN(table3.apcPerDay)} per day`,
      ],
    ],
    numericCols: [2],
    widths: [6, 44, 50],
    emphasizeLastRow: true,
  };

  const notes = [
    "Table 1: C = A x 100 / B; D = ETC collection + Annual Pass compensation; E = D x 100 / C; " +
      "F = E / Sum of E; G = F rounded to a whole percentage so that the column totals 100%.",
    "Table 2: A = monthly ETC collection / days in month; MF = fee-revision multiplier " +
      "(day-weighted in the revision month); B = A x MF. Average of B feeds Table 3 item 1.",
    "Table 3: item 3 = item 1 x 100 / item 2; item 4 = item 3 x 365; item 6 = item 4 x (1 + growth); " +
      "item 7 = 3 Cr x (1 - exp(-item 6 / 6 Cr)); item 8 = (item 6 - item 7) x 5%; " +
      "APC-2 = item 6 - item 7 - item 8.",
  ];

  return {
    title: `APC Report - ${plaza.name} (Plaza code ${plaza.code})`,
    subtitle: `APC-2 for ${monthLabel(uptoMonth)}`,
    meta: [
      ["Calculated up to", monthLabel(uptoMonth)],
      ["12-month window", windowText],
      ["Traffic growth", `${growth}%`],
      [
        "MF (fee-rate revisions)",
        validMf.length ? validMf.map((e) => `${e.date}: ${e.factor}`).join(", ") : "none",
      ],
      ["Generated on", new Date().toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" })],
    ],
    hero: {
      label: "APC-2",
      value: `Rs ${table3.apcCr.toFixed(2)} Cr.`,
      perDay: `Rs ${fmtIN(table3.apcPerDay)} per day`,
    },
    alert,
    table1: t1,
    table2: t2,
    table3: t3,
    notes,
    fileBase: `APC_${plaza.code}_${uptoMonth}`,
  };
}
