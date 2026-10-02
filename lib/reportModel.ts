// A document-neutral view of one APC calculation laid out like NHAI's
// single-page "APC of <plaza> Toll Plaza as per new methodology" sheet.
// Every value is already formatted as text; the PDF and Word exports render
// this model so they stay identical in structure and wording. Text is kept
// to characters available in the PDF's standard fonts (no rupee sign).

import type { ApcResult, MfEntry, PlazaInfo } from "./apc";
import {
  feeRateYearLabel,
  fmtDateDMY,
  fmtIN,
  fmtMoney,
  mfDateLabel,
  mfFactor,
  monthFull,
  monthShort,
} from "./apc";

export interface SheetMf {
  /** Column captions, e.g. "MF1 (01.10.25)". */
  labels: string[];
  oldRemittance: string[];
  newRemittance: string[];
  factor: string[];
}

export interface SheetTable1 {
  title: string;
  period: string; // printed in red after the title
  head: string[]; // 10 columns
  letters: string[]; // formula row under the header
  body: string[][]; // 6 rows x 10 columns
  rounded: string[]; // whole-number % per row (outside the table)
  roundedTotal: string; // "100%"
  weightedLabel: string;
  weighted: string;
}

export interface SheetTable2 {
  title: string;
  head: string[]; // Month | AP | A | B
  body: string[][];
  average: string[];
  imputedNote: string | null;
}

export interface SheetChartPoint {
  label: string;
  value: number | null;
  imputed: boolean;
}

export interface SheetTable3 {
  title: string;
  rows: [string, string, string][]; // #, item, value (short form inline)
  perDay: string; // "[Rs 386027 per day]"
}

export interface SheetTable4 {
  title: string;
  rows: [string, string][];
}

export interface SheetModel {
  title: string;
  calculatedOn: string;
  plazaCode: string;
  plazaName: string;
  mf: SheetMf;
  plazaInfo: [string, string][];
  table1: SheetTable1;
  table2: SheetTable2;
  chart: { title: string; points: SheetChartPoint[] };
  table3: SheetTable3;
  formula: { heading: string; lines: string[] };
  table4: SheetTable4 | null;
  alert: string | null;
  fileBase: string;
}

const dash = (v: unknown) => (v === null || v === undefined || v === "" ? "-" : String(v));

export function buildSheetModel(
  plaza: PlazaInfo,
  result: ApcResult,
  uptoMonth: string,
  mfEntries: MfEntry[],
  growth: number
): SheetModel {
  const { table1, table2, table3, table4, months } = result;
  const first = months[0];
  const last = months[11];

  const validMf = mfEntries
    .filter((e) => e.date && mfFactor(e) > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
  const mf: SheetMf = {
    labels: validMf.map((e, i) => `MF${i + 1} (${mfDateLabel(e.date)})`),
    oldRemittance: validMf.map((e) => (e.oldRemittance ? fmtIN(e.oldRemittance) : "-")),
    newRemittance: validMf.map((e) => (e.newRemittance ? fmtIN(e.newRemittance) : "-")),
    factor: validMf.map((e) => mfFactor(e).toFixed(4)),
  };

  const imputedMonths = table2.rows.filter((r) => r.imputed).map((r) => monthShort(r.month));

  const t1: SheetTable1 = {
    title: "Table-1: Calculation of ETC Penetration",
    period: `(Period: ${monthFull(first)} to ${monthFull(last)})`,
    head: [
      "Category of Vehicle",
      "Mapper Vehicle Description as per NPCI",
      "Total ETC Transactions as per NPCI",
      "Total Cash+UPI @TMCC",
      "Exempted Transaction @TMCC",
      "Total Transactions for APC",
      "% ETC Penetration",
      "ETC Collection as per NPCI + Annual Pass Compensation",
      "Derived Total Collection",
      "% Contribution of Different Categories of vehicles in Collection",
    ],
    letters: ["", "", "A", "B1", "B2", "B = A + B1 + 50% of B2", "C", "D", "E = (D x 100 / C)", "F = E / SUM(E)"],
    body: table1.rows.map((r) => [
      r.category,
      r.npciCodes,
      fmtIN(r.etcTxn),
      fmtIN(r.cashUpiTxn),
      fmtIN(r.exemptTxn),
      fmtIN(r.totalTxn),
      r.penetration.toFixed(2),
      fmtIN(r.etcCollectionWithAp),
      fmtIN(r.derivedTotal),
      `${(r.contribution * 100).toFixed(2)}%`,
    ]),
    rounded: table1.rows.map((r) => `${r.contributionRounded}%`),
    roundedTotal: `${table1.contributionTotal}%`,
    weightedLabel: "Weighted Average ETC Penetration = SUM (C x D) / SUM (D)",
    weighted: table1.weightedPenetration.toFixed(4),
  };

  const fy = feeRateYearLabel(uptoMonth);
  const t2: SheetTable2 = {
    title: "Table-2: Calculation of Annual Average Daily ETC Collection",
    head: [
      "Month",
      "Monthly Average Daily Annual Pass Compensation",
      "Monthly Average Daily ETC Collection as per Actual Fee Rates, [A]",
      `Monthly Avg ETC Collection as per revised ${fy} Fee Rates [B = A x MF]`,
    ],
    body: table2.rows.map((r) => [
      monthShort(r.month) + (r.imputed ? "*" : ""),
      r.etcDaily === null ? "-" : fmtIN(r.apDaily),
      r.etcDaily === null ? "-" : fmtIN(r.etcDaily),
      r.etcDaily === null ? "-" : fmtIN(r.normalizedDaily),
    ]),
    average: ["Average", fmtIN(table2.avgApDaily), fmtIN(table2.avgEtcDaily), fmtIN(table2.avgNormalizedDaily)],
    imputedNote: imputedMonths.length
      ? `* ${imputedMonths.join(", ")}: no data; average of the months with data is used.`
      : null,
  };

  const t3: SheetTable3 = {
    title: "Table-3: Calculation of APC-2",
    rows: [
      ["1", "Average Daily FASTag Collection", fmtMoney(table3.avgDailyFastagCollection)],
      ["2", "FASTag Penetration (In decimal)", table3.fastagPenetration.toFixed(4)],
      ["3", "Annual Average Daily Collection", fmtMoney(table3.annualAvgDailyCollection)],
      ["4", "Annual Expected Collection", fmtMoney(table3.annualExpectedCollection)],
      ["5", "Traffic Growth (in %)", `${table3.trafficGrowthPct}`],
      ["6", "Net expected Collection", fmtMoney(table3.netExpectedCollection)],
      ["7", "Less Administrative Charges", fmtMoney(table3.adminCharges)],
      ["8", "Less Contractor Profit @5%", fmtMoney(table3.contractorProfit)],
      ["9", "APC-2 (Cr)", `${table3.apcCr.toFixed(2)} Cr.`],
    ],
    perDay: `[Rs ${Math.round(table3.apcPerDay)} per day]`,
  };

  const t4: SheetTable4 | null = table4
    ? {
        title: "Table-4: Calculation of APC-3",
        rows: [
          ["Present Remittance (Yearly) =", fmtIN(table4.presentRemittanceYearly)],
          ["Present Annual Pass Compensation (Yearly) =", fmtIN(table4.presentApCompensationYearly)],
          ["Net Remittance =", fmtIN(table4.netRemittance)],
          ["increment =", `${table4.incrementPct.toFixed(2)}%`],
          ["APC-3 (Cr) =", table4.apc3Cr.toFixed(2)],
        ],
      }
    : null;

  const alert = table2.negativeTrend
    ? `Traffic growth should be taken as 0%: the average daily ETC collection shows a negative ` +
      `trend over ${monthShort(first)} - ${monthShort(last)}` +
      (growth > 0 ? `; this sheet uses ${growth}%.` : "; this sheet uses 0%.")
    : null;

  return {
    title: `APC of ${plaza.name} Toll Plaza as per new methodology`,
    calculatedOn: fmtDateDMY(new Date()),
    plazaCode: String(plaza.code),
    plazaName: plaza.name,
    mf,
    plazaInfo: [
      ["No of Lane:", dash(plaza.lanes)],
      ["Type:", dash(plaza.type)],
      ["PIU:", dash(plaza.piu)],
      ["RO:", dash(plaza.ro)],
      ["State:", dash(plaza.state)],
    ],
    table1: t1,
    table2: t2,
    chart: {
      title: `ETC Collection as per revised ${fy} fee rate`,
      points: table2.rows.map((r) => ({
        label: monthShort(r.month),
        value: r.normalizedDaily,
        imputed: r.imputed,
      })),
    },
    table3: t3,
    formula: {
      heading: "Formula to calculate Administrative Expenses:",
      lines: [
        "Y = 3 x (1 - e^(-X/6))",
        "Where, X = Net Expected Collection in Cr/Annum",
        "Y = Administrative Expenses in Cr/Annum",
      ],
    },
    table4: t4,
    alert,
    fileBase: `APC_${plaza.code}_${uptoMonth}`,
  };
}
