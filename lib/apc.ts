// APC calculation, replicating NHAI's sheet "2. APC".
// All money values in rupees unless noted.

export const CATEGORIES = [
  "Car/Jeep/Van/Light Motor Vehicle",
  "LCV/Light Good Vehicle or Mini Bus",
  "Bus/Truck (Two Axles)",
  "Three Axle Commercial Vehicles",
  "HCM or EME or MAV (4 to 6 Axles)",
  "Oversized Vehicles (Seven or More Axles)",
] as const;

export const NPCI_CODES: Record<string, string> = {
  "Car/Jeep/Van/Light Motor Vehicle": "VC4, VC20",
  "LCV/Light Good Vehicle or Mini Bus": "VC5, VC6, VC9",
  "Bus/Truck (Two Axles)": "VC7, VC10",
  "Three Axle Commercial Vehicles": "VC8, VC11",
  "HCM or EME or MAV (4 to 6 Axles)": "VC12, VC13, VC14",
  "Oversized Vehicles (Seven or More Axles)": "VC15, VC16, VC17",
};

export interface MonthlyRow {
  month: string; // "YYYY-MM"
  category: string;
  etc_cnt: number;
  cash_cnt: number;
  upi_cnt: number;
  exempt_cnt: number;
  tmcc_total: number;
  etc_collection: number;
  ap_cnt: number;
  ap_compensation: number;
}

export interface MfEntry {
  date: string; // effective date "YYYY-MM-DD"
  factor: number; // new rate / old rate
}

export interface Table1Row {
  category: string;
  npciCodes: string;
  etcTxn: number; // A
  cashUpiTxn: number;
  exemptHalf: number;
  totalTxn: number; // B = ETC + Cash + UPI + 50% exempt
  penetration: number; // C = A*100/B
  etcCollectionWithAp: number; // D = ETC collection + AP compensation
  derivedTotal: number; // E = D*100/C
  contribution: number; // F = E / sum(E)
  contributionRounded: number; // G
}

export interface Table1 {
  rows: Table1Row[];
  weightedPenetration: number; // sum(C*D)/sum(D)
  contributionTotal: number; // sum(G), should be 1
  roundOffOk: boolean;
}

export interface Table2Row {
  month: string;
  days: number;
  hasData: boolean;
  apDaily: number | null; // B
  etcDaily: number | null; // C
  totalDaily: number | null; // D = B + C
  mfMultiplier: number;
  normalizedDaily: number | null; // E = D * multiplier
}

export interface Table2 {
  rows: Table2Row[];
  avgApDaily: number;
  avgTotalDaily: number; // average of D over months with data
  avgNormalizedDaily: number; // average of E over months with data
  monthsUsed: number;
}

export interface Table3 {
  avgDailyFastagCollection: number; // 1
  fastagPenetration: number; // 2
  annualAvgDailyCollection: number; // 4
  annualExpectedCollection: number; // 5
  trafficGrowthPct: number; // 6
  netExpectedCollection: number; // 7
  adminCharges: number; // 8
  contractorProfit: number; // 9
  apcCr: number; // 10, in crores
  apcPerDay: number; // rupees per day
}

export function daysInMonth(monthKey: string): number {
  const [y, m] = monthKey.split("-").map(Number);
  return new Date(y, m, 0).getDate();
}

/** The 12 calendar months ending at `uptoMonth` (inclusive), oldest first. */
export function windowMonths(uptoMonth: string): string[] {
  const [y, m] = uptoMonth.split("-").map(Number);
  const out: string[] = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(y, m - 1 - i, 1);
    out.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
  }
  return out;
}

/** Latest completed month "YYYY-MM" relative to now. */
export function lastCompletedMonth(now: Date = new Date()): string {
  const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

/**
 * MF multiplier for a month, replicating the Excel's day-weighted blending.
 * Entries: {date, factor = new/old}. For a month fully before a revision the
 * collection is multiplied by that factor (later revisions compound); the
 * revision month gets ((k-1)*f + (n-k+1)*1)/n where k = effective day,
 * n = days in month (Excel: Oct-25 revision on the 20th -> (19*MF1+12)/31).
 */
export function mfMultiplier(monthKey: string, entries: MfEntry[]): number {
  let mult = 1;
  for (const e of entries) {
    if (!e.date || !e.factor || e.factor <= 0) continue;
    const effMonth = e.date.slice(0, 7);
    if (monthKey < effMonth) {
      mult *= e.factor;
    } else if (monthKey === effMonth) {
      const k = Number(e.date.slice(8, 10));
      const n = daysInMonth(monthKey);
      mult *= ((k - 1) * e.factor + (n - k + 1)) / n;
    }
    // months on/after the revision month: factor already reflected -> *1
  }
  return mult;
}

type CatTotals = {
  etc_cnt: number;
  cash_cnt: number;
  upi_cnt: number;
  exempt_cnt: number;
  tmcc_total: number;
  etc_collection: number;
  ap_compensation: number;
};

export function computeTable1(rows: MonthlyRow[], months: string[]): Table1 {
  const monthSet = new Set(months);
  const totals = new Map<string, CatTotals>();
  for (const cat of CATEGORIES) {
    totals.set(cat, {
      etc_cnt: 0, cash_cnt: 0, upi_cnt: 0, exempt_cnt: 0,
      tmcc_total: 0, etc_collection: 0, ap_compensation: 0,
    });
  }
  for (const r of rows) {
    if (!monthSet.has(r.month)) continue;
    const t = totals.get(r.category);
    if (!t) continue;
    t.etc_cnt += r.etc_cnt;
    t.cash_cnt += r.cash_cnt;
    t.upi_cnt += r.upi_cnt;
    t.exempt_cnt += r.exempt_cnt;
    t.tmcc_total += r.tmcc_total;
    t.etc_collection += r.etc_collection;
    t.ap_compensation += r.ap_compensation;
  }

  const base = CATEGORIES.map((cat) => {
    const t = totals.get(cat)!;
    const penetration = t.tmcc_total > 0 ? (t.etc_cnt * 100) / t.tmcc_total : 100;
    const d = t.etc_collection + t.ap_compensation;
    return {
      category: cat,
      npciCodes: NPCI_CODES[cat],
      etcTxn: t.etc_cnt,
      cashUpiTxn: t.cash_cnt + t.upi_cnt,
      exemptHalf: 0.5 * t.exempt_cnt,
      totalTxn: t.tmcc_total,
      penetration,
      etcCollectionWithAp: d,
      derivedTotal: penetration > 0 ? (d * 100) / penetration : 0,
    };
  });

  const sumDerived = base.reduce((s, r) => s + r.derivedTotal, 0);
  const rows1: Table1Row[] = base.map((r) => {
    const contribution = sumDerived > 0 ? r.derivedTotal / sumDerived : 0;
    return {
      ...r,
      contribution,
      contributionRounded: Math.round(contribution * 100) / 100,
    };
  });

  const sumD = base.reduce((s, r) => s + r.etcCollectionWithAp, 0);
  const weightedPenetration =
    sumD > 0
      ? base.reduce((s, r) => s + r.penetration * r.etcCollectionWithAp, 0) / sumD
      : 0;
  const contributionTotal = rows1.reduce((s, r) => s + r.contributionRounded, 0);

  return {
    rows: rows1,
    weightedPenetration,
    contributionTotal,
    roundOffOk: Math.abs(contributionTotal - 1) < 1e-9,
  };
}

export function computeTable2(
  rows: MonthlyRow[],
  months: string[],
  mfEntries: MfEntry[]
): Table2 {
  const byMonth = new Map<string, { etc: number; ap: number; any: boolean }>();
  for (const m of months) byMonth.set(m, { etc: 0, ap: 0, any: false });
  for (const r of rows) {
    const b = byMonth.get(r.month);
    if (!b) continue;
    b.etc += r.etc_collection;
    b.ap += r.ap_compensation;
    b.any = true;
  }

  const out: Table2Row[] = months.map((m) => {
    const b = byMonth.get(m)!;
    const days = daysInMonth(m);
    const mult = mfMultiplier(m, mfEntries);
    // A month with no rows or zero collection is excluded (Excel AVERAGEIF "<>0").
    const hasData = b.any && b.etc + b.ap > 0;
    if (!hasData) {
      return {
        month: m, days, hasData: false,
        apDaily: null, etcDaily: null, totalDaily: null,
        mfMultiplier: mult, normalizedDaily: null,
      };
    }
    const apDaily = b.ap / days;
    const etcDaily = b.etc / days;
    const totalDaily = apDaily + etcDaily;
    return {
      month: m, days, hasData: true,
      apDaily, etcDaily, totalDaily,
      mfMultiplier: mult,
      normalizedDaily: totalDaily * mult,
    };
  });

  const used = out.filter((r) => r.hasData);
  const avg = (f: (r: Table2Row) => number) =>
    used.length > 0 ? used.reduce((s, r) => s + f(r), 0) / used.length : 0;

  return {
    rows: out,
    avgApDaily: avg((r) => r.apDaily!),
    avgTotalDaily: avg((r) => r.totalDaily!),
    avgNormalizedDaily: avg((r) => r.normalizedDaily!),
    monthsUsed: used.length,
  };
}

export function computeTable3(
  avgDailyFastagCollection: number,
  fastagPenetration: number,
  trafficGrowthPct: number
): Table3 {
  const growth = Math.max(0, trafficGrowthPct); // 0% if negative
  const annualAvgDailyCollection =
    fastagPenetration > 0 ? (avgDailyFastagCollection * 100) / fastagPenetration : 0;
  const annualExpectedCollection = annualAvgDailyCollection * 365;
  const netExpectedCollection = annualExpectedCollection * (1 + growth / 100);
  const adminCharges = 3 * (1 - Math.exp(-netExpectedCollection / 60_000_000)) * 1e7;
  const contractorProfit = (netExpectedCollection - adminCharges) * 0.05;
  const apcCr =
    Math.round(((netExpectedCollection - adminCharges - contractorProfit) / 1e7) * 100) / 100;
  return {
    avgDailyFastagCollection,
    fastagPenetration,
    annualAvgDailyCollection,
    annualExpectedCollection,
    trafficGrowthPct: growth,
    netExpectedCollection,
    adminCharges,
    contractorProfit,
    apcCr,
    apcPerDay: Math.round((apcCr * 1e7) / 365),
  };
}

export interface ApcResult {
  months: string[];
  table1: Table1;
  table2: Table2;
  table3: Table3;
}

export function computeApc(
  rows: MonthlyRow[],
  uptoMonth: string,
  mfEntries: MfEntry[],
  trafficGrowthPct: number
): ApcResult {
  const months = windowMonths(uptoMonth);
  const table1 = computeTable1(rows, months);
  const table2 = computeTable2(rows, months, mfEntries);
  const table3 = computeTable3(
    table2.avgNormalizedDaily,
    table1.weightedPenetration,
    trafficGrowthPct
  );
  return { months, table1, table2, table3 };
}

// ---- formatting helpers ----

export function fmtIN(n: number | null | undefined, digits = 0): string {
  if (n === null || n === undefined || Number.isNaN(n)) return "-";
  return n.toLocaleString("en-IN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

export function fmtCr(n: number): string {
  return `${(n / 1e7).toFixed(2)} Cr.`;
}

export function monthLabel(monthKey: string): string {
  const [y, m] = monthKey.split("-").map(Number);
  const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${names[m - 1]}-${y}`;
}
