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
  factor: number; // new rate / old rate (= newRemittance / oldRemittance when both given)
  oldRemittance?: number; // daily remittance before the revision (Rs)
  newRemittance?: number; // daily remittance after the revision (Rs)
}

/** Factor implied by the remittances, or the stored factor when they are absent. */
export function mfFactor(e: MfEntry): number {
  if (e.oldRemittance && e.newRemittance && e.oldRemittance > 0) {
    return e.newRemittance / e.oldRemittance;
  }
  return e.factor;
}

/** Plaza master data; fields other than code/name are optional. */
export interface PlazaInfo {
  code: number;
  name: string;
  lanes?: number | string | null;
  type?: string | null;
  piu?: string | null;
  ro?: string | null;
  state?: string | null;
}

export interface Table1Row {
  category: string;
  npciCodes: string;
  etcTxn: number; // A
  cashUpiTxn: number; // B1
  exemptTxn: number; // B2 (full exempted count)
  exemptHalf: number; // 50% of B2
  totalTxn: number; // B = A + B1 + 50% of B2
  penetration: number; // C = A*100/B
  etcCollectionWithAp: number; // D = ETC collection + AP compensation
  derivedTotal: number; // E = D*100/C
  contribution: number; // F = E / sum(E), as a fraction
  contributionRounded: number; // G = whole-number %, largest-remainder so sum(G) = 100
}

export interface Table1 {
  rows: Table1Row[];
  weightedPenetration: number; // sum(C*D)/sum(D)
  contributionTotal: number; // sum(G) in %, always 100 when there is any collection
  roundOffOk: boolean;
}

export interface Table2Row {
  month: string;
  days: number;
  /** True when the month has actual collection data. */
  hasData: boolean;
  /** True when the month had no data and A / AP were filled with the average of the months that do. */
  imputed: boolean;
  apDaily: number | null; // Annual Pass compensation per day (display only)
  etcDaily: number | null; // A = ETC collection per day at actual fee rates
  mfMultiplier: number;
  normalizedDaily: number | null; // B = A * MF multiplier
}

export interface Table2 {
  rows: Table2Row[];
  avgApDaily: number; // display only, not used downstream
  avgEtcDaily: number; // average of A over all 12 months (imputed months included)
  avgNormalizedDaily: number; // average of B over all 12 months (imputed months included)
  /** Months with actual data. */
  monthsUsed: number;
  /** Months filled with the average of the months that have data. */
  monthsImputed: number;
  /** Least-squares slope of B over the months with actual data (rupees/day per month). */
  trendSlope: number;
  /** True when the overall trend of B across the window is downward. */
  negativeTrend: boolean;
}

export interface Table4 {
  newRemittanceDaily: number; // latest revised daily remittance (Rs)
  presentRemittanceYearly: number; // x 365
  latestApDaily: number; // latest month's average daily AP compensation
  presentApCompensationYearly: number; // x 365
  netRemittance: number; // present remittance - AP compensation
  incrementPct: number;
  apc3Cr: number; // net remittance x (1 + increment), in crores
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
    const f = mfFactor(e);
    if (!e.date || !f || f <= 0) continue;
    const effMonth = e.date.slice(0, 7);
    if (monthKey < effMonth) {
      mult *= f;
    } else if (monthKey === effMonth) {
      const k = Number(e.date.slice(8, 10));
      const n = daysInMonth(monthKey);
      mult *= ((k - 1) * f + (n - k + 1)) / n;
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
      exemptTxn: t.exempt_cnt,
      exemptHalf: 0.5 * t.exempt_cnt,
      totalTxn: t.tmcc_total,
      penetration,
      etcCollectionWithAp: d,
      derivedTotal: penetration > 0 ? (d * 100) / penetration : 0,
    };
  });

  const sumDerived = base.reduce((s, r) => s + r.derivedTotal, 0);
  const contributions = base.map((r) =>
    sumDerived > 0 ? r.derivedTotal / sumDerived : 0
  );
  const roundedPct = roundPercentagesToTotal(contributions.map((c) => c * 100));
  const rows1: Table1Row[] = base.map((r, i) => ({
    ...r,
    contribution: contributions[i],
    contributionRounded: roundedPct[i],
  }));

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
    roundOffOk: sumDerived <= 0 || Math.abs(contributionTotal - 100) < 1e-9,
  };
}

/**
 * Round percentages to whole numbers so that they still add up to exactly
 * 100 (largest-remainder method). Returns all zeros if the inputs sum to 0.
 */
export function roundPercentagesToTotal(pcts: number[], total = 100): number[] {
  const sum = pcts.reduce((s, v) => s + v, 0);
  if (sum <= 0) return pcts.map(() => 0);
  const scaled = pcts.map((v) => (v * total) / sum);
  const floors = scaled.map(Math.floor);
  let remaining = total - floors.reduce((s, v) => s + v, 0);
  const order = scaled
    .map((v, i) => ({ i, frac: v - floors[i] }))
    .sort((a, b) => b.frac - a.frac || a.i - b.i);
  for (const { i } of order) {
    if (remaining <= 0) break;
    floors[i] += 1;
    remaining -= 1;
  }
  return floors;
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

  const actual: Table2Row[] = months.map((m) => {
    const b = byMonth.get(m)!;
    const days = daysInMonth(m);
    const mult = mfMultiplier(m, mfEntries);
    // A month with no rows or zero ETC collection has no usable data.
    const hasData = b.any && b.etc > 0;
    if (!hasData) {
      return {
        month: m, days, hasData: false, imputed: false,
        apDaily: null, etcDaily: null,
        mfMultiplier: mult, normalizedDaily: null,
      };
    }
    const apDaily = b.ap / days;
    const etcDaily = b.etc / days;
    return {
      month: m, days, hasData: true, imputed: false,
      apDaily, etcDaily,
      mfMultiplier: mult,
      // Annual Pass compensation is shown for information only; the
      // normalized figure is ETC collection at revised fee rates.
      normalizedDaily: etcDaily * mult,
    };
  });

  const used = actual.filter((r) => r.hasData);
  const avgOf = (list: Table2Row[], f: (r: Table2Row) => number) =>
    list.length > 0 ? list.reduce((s, r) => s + f(r), 0) / list.length : 0;
  const fillEtc = avgOf(used, (r) => r.etcDaily!);
  const fillAp = avgOf(used, (r) => r.apDaily!);

  // Months without data take the average of the months that have data
  // (A and AP), then get their own MF multiplier applied for B.
  const out: Table2Row[] = actual.map((r) => {
    if (r.hasData || used.length === 0) return r;
    return {
      ...r,
      imputed: true,
      apDaily: fillAp,
      etcDaily: fillEtc,
      normalizedDaily: fillEtc * r.mfMultiplier,
    };
  });

  const filled = out.filter((r) => r.etcDaily !== null);
  // Trend is judged on actual data only; imputed months would flatten it.
  const trendSlope = linearTrendSlope(
    actual.map((r) => r.normalizedDaily)
  );

  return {
    rows: out,
    avgApDaily: avgOf(filled, (r) => r.apDaily!),
    avgEtcDaily: avgOf(filled, (r) => r.etcDaily!),
    avgNormalizedDaily: avgOf(filled, (r) => r.normalizedDaily!),
    monthsUsed: used.length,
    monthsImputed: out.filter((r) => r.imputed).length,
    trendSlope,
    negativeTrend: trendSlope < 0,
  };
}

/**
 * Least-squares slope of a series indexed by position (month number in the
 * window); null entries (months without data) are skipped. Returns 0 when
 * fewer than two points are available.
 */
export function linearTrendSlope(values: (number | null)[]): number {
  const pts: [number, number][] = [];
  values.forEach((v, i) => {
    if (v !== null && v !== undefined && Number.isFinite(v)) pts.push([i, v]);
  });
  if (pts.length < 2) return 0;
  const n = pts.length;
  const meanX = pts.reduce((s, [x]) => s + x, 0) / n;
  const meanY = pts.reduce((s, [, y]) => s + y, 0) / n;
  let sxy = 0;
  let sxx = 0;
  for (const [x, y] of pts) {
    sxy += (x - meanX) * (y - meanY);
    sxx += (x - meanX) ** 2;
  }
  return sxx > 0 ? sxy / sxx : 0;
}

/** Default traffic growth assumption (%). Use 0% when the collection trend is negative. */
export const DEFAULT_TRAFFIC_GROWTH_PCT = 5;

/** Default yearly increment (%) applied to the net remittance for APC-3. */
export const DEFAULT_INCREMENT_PCT = 2.5;

/**
 * Table 4 (APC-3): present remittance from the latest fee revision, less the
 * present Annual Pass compensation, grown by the increment. Returns null when
 * no revision with a new remittance has been entered.
 */
export function computeTable4(
  mfEntries: MfEntry[],
  table2: Table2,
  incrementPct: number = DEFAULT_INCREMENT_PCT
): Table4 | null {
  const withRemittance = mfEntries
    .filter((e) => e.date && e.newRemittance && e.newRemittance > 0)
    .sort((a, b) => a.date.localeCompare(b.date));
  const latest = withRemittance[withRemittance.length - 1];
  if (!latest) return null;
  const actual = table2.rows.filter((r) => r.hasData);
  const latestApDaily = actual.length ? actual[actual.length - 1].apDaily! : 0;
  const newRemittanceDaily = latest.newRemittance!;
  const presentRemittanceYearly = newRemittanceDaily * 365;
  const presentApCompensationYearly = latestApDaily * 365;
  const netRemittance = presentRemittanceYearly - presentApCompensationYearly;
  const apc3Cr = Math.round((netRemittance * (1 + incrementPct / 100)) / 1e7 * 100) / 100;
  return {
    newRemittanceDaily,
    presentRemittanceYearly,
    latestApDaily,
    presentApCompensationYearly,
    netRemittance,
    incrementPct,
    apc3Cr,
  };
}

export function computeTable3(
  avgDailyFastagCollection: number,
  fastagPenetration: number,
  trafficGrowthPct: number = DEFAULT_TRAFFIC_GROWTH_PCT
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
  table4: Table4 | null;
}

export function computeApc(
  rows: MonthlyRow[],
  uptoMonth: string,
  mfEntries: MfEntry[],
  trafficGrowthPct: number = DEFAULT_TRAFFIC_GROWTH_PCT,
  incrementPct: number = DEFAULT_INCREMENT_PCT
): ApcResult {
  const months = windowMonths(uptoMonth);
  const table1 = computeTable1(rows, months);
  const table2 = computeTable2(rows, months, mfEntries);
  const table3 = computeTable3(
    table2.avgNormalizedDaily,
    table1.weightedPenetration,
    trafficGrowthPct
  );
  const table4 = computeTable4(mfEntries, table2, incrementPct);
  return { months, table1, table2, table3, table4 };
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

/** Short Indian form: 3.40 Cr. / 8.85 L / 34.0 k / 950 */
export function fmtShort(n: number): string {
  const a = Math.abs(n);
  if (a >= 1e7) return `${(n / 1e7).toFixed(2)} Cr.`;
  if (a >= 1e5) return `${(n / 1e5).toFixed(2)} L`;
  if (a >= 1e3) return `${(n / 1e3).toFixed(1)} k`;
  return fmtIN(n, 0);
}

/** Rupee amount with its short form inline, e.g. "3,40,14,000 (3.40 Cr.)". */
export function fmtMoney(n: number): string {
  if (Math.abs(n) < 1e3) return fmtIN(n, 0);
  return `${fmtIN(n, 0)} (${fmtShort(n)})`;
}

const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const MONTH_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

export function monthLabel(monthKey: string): string {
  const [y, m] = monthKey.split("-").map(Number);
  return `${MONTH_NAMES[m - 1]}-${y}`;
}

/** "Aug-25" style label used in the NHAI sheet. */
export function monthShort(monthKey: string): string {
  const [y, m] = monthKey.split("-").map(Number);
  return `${MONTH_NAMES[m - 1]}-${String(y).slice(2)}`;
}

/** "August-2025" style label used in the Table-1 period. */
export function monthFull(monthKey: string): string {
  const [y, m] = monthKey.split("-").map(Number);
  return `${MONTH_FULL[m - 1]}-${y}`;
}

/** Indian financial year label ("2026-27") of the month after `monthKey`. */
export function feeRateYearLabel(monthKey: string): string {
  const [y, m] = monthKey.split("-").map(Number);
  const next = new Date(y, m, 1); // month after the window end
  const fyStart = next.getMonth() >= 3 ? next.getFullYear() : next.getFullYear() - 1;
  return `${fyStart}-${String(fyStart + 1).slice(2)}`;
}

/** "DD-MM-YYYY" as printed in the sheet's "Calculated on" cell. */
export function fmtDateDMY(d: Date): string {
  return `${String(d.getDate()).padStart(2, "0")}-${String(d.getMonth() + 1).padStart(2, "0")}-${d.getFullYear()}`;
}

/** "01.10.25" style label for an MF effective date "YYYY-MM-DD". */
export function mfDateLabel(date: string): string {
  if (!date) return "";
  const [y, m, d] = date.split("-");
  return `${d}.${m}.${y.slice(2)}`;
}
