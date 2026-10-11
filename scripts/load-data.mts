// Loads NHAI's monthly ETC / Non-ETC / Annual Pass workbook into Neon Postgres.
//
//   npx tsx scripts/load-data.mts <workbook.xlsx>            # dry run: parse + report only
//   npx tsx scripts/load-data.mts <workbook.xlsx> --apply    # replace DB contents
//   npx tsx scripts/load-data.mts <workbook.xlsx> --apply --add-missing-plazas [--point-based-only]
//
// Two workbook layouts are understood (columns matched by name, case-insensitive):
//  * "ETC & Non-ETC" / "Annual Pass"  (Month, NETC, PLAZA, VEHICLE_CLASS, ETC_COUNT,
//    ETC_COLLECTION, ANNUAL_PASS_COUNT, CASH_COUNT, UPI_COUNT, EXEMPT_COUNT ...
//    / ANNUAL_PASS_QUALIFIED_TXN, ANNUAL_PASS_COMPENSATION)
//  * "ETC & Non-ETC Data" / "Annual Pass Data"  (Month as "May-2025", Plaza, NETC,
//    Ownership, Toll Type, Closed-Loop/Point-Based, Vehicle Class, ETC_Traffic,
//    ETC_Collection, Annual_Pass_Traffic, CASH_Traffic, UPI_Traffic,
//    EXEMPTION_Traffic ... / Total_Qualified_Transactions, Annual_Pass_Compensation)
//
// ETC sheet -> plaza_monthly_data, one row per plaza x month x NHAI category.
// Raw vehicle classes are mapped to the 6 categories (4-6 Axle and HCM/EME
// merge); unmapped classes (Tractors, 2/3 Wheeler, Blank/Others, WAVC, ...)
// are dropped. tmcc_total = etc + cash + upi + 50% of exempted.
// Annual Pass sheet -> annual_pass_monthly.
// --point-based-only keeps rows whose Closed-Loop/Point-Based column says
// Point-Based (when that column exists).
//
// --apply DELETEs all rows from both tables and inserts the new data inside
// one transaction. pb_plazas is left untouched unless --add-missing-plazas
// (required when the workbook has plazas missing from pb_plazas, because
// plaza_monthly_data.plaza_code references pb_plazas).
// DATABASE_URL is read from the environment or .env.local.
// If the local network's DNS cannot resolve *.neon.tech, set
// DNS_SERVERS=8.8.8.8,1.1.1.1 to resolve the database host via those servers.

import { readFileSync, existsSync } from "node:fs";
import path from "node:path";
import ExcelJS from "exceljs";
import { neon, neonConfig } from "@neondatabase/serverless";
import { CATEGORIES } from "../lib/apc.ts";

type Category = (typeof CATEGORIES)[number];

/** Raw VEHICLE_CLASS -> NHAI category. Classes not listed are ignored. */
const CLASS_MAP: Record<string, Category> = {
  "Car/Jeep/Van/LMV": CATEGORIES[0],
  "LCV": CATEGORIES[1],
  "Bus/Truck": CATEGORIES[2],
  "Upto 3 Axle Vehicle": CATEGORIES[3],
  "4 to 6 Axle": CATEGORIES[4],
  "HCM/EME": CATEGORIES[4],
  "7 or more Axle": CATEGORIES[5],
};

const BATCH = 4000;

/** Column aliases (upper-cased) across the known workbook layouts. */
const COLS = {
  month: ["MONTH"],
  netc: ["NETC", "NETC_CODE", "PLAZA_CODE"],
  plaza: ["PLAZA", "PLAZA_NAME", "NAME"],
  vc: ["VEHICLE_CLASS", "VEHICLE CLASS", "VC"],
  etc: ["ETC_COUNT", "ETC_TRAFFIC", "ETC_CNT"],
  etcCol: ["ETC_COLLECTION"],
  ap: ["ANNUAL_PASS_COUNT", "ANNUAL_PASS_TRAFFIC", "AP_CNT"],
  cash: ["CASH_COUNT", "CASH_TRAFFIC", "CASH_CNT"],
  upi: ["UPI_COUNT", "UPI_TRAFFIC", "UPI_CNT"],
  exempt: ["EXEMPT_COUNT", "EXEMPTION_TRAFFIC", "EXEMPT_CNT"],
  loop: ["CLOSED-LOOP/POINT-BASED", "CLOSED_LOOP/POINT_BASED", "PLAZA_TYPE"],
  comp: ["ANNUAL_PASS_COMPENSATION"],
  qual: ["ANNUAL_PASS_QUALIFIED_TXN", "TOTAL_QUALIFIED_TRANSACTIONS", "TOTAL_QUALIFIED_TRANSACTION"],
};

const MONTHS3 = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];

interface MonthlyAgg {
  plaza_code: number;
  month: string; // YYYY-MM-01
  category: Category;
  etc_cnt: number;
  cash_cnt: number;
  upi_cnt: number;
  exempt_cnt: number;
  etc_collection: number;
  ap_cnt: number;
}

interface ApRow {
  plaza_code: number;
  month: string;
  compensation: number;
  qualified_txn: number;
}

// ---------------------------------------------------------------- helpers

function loadEnv() {
  if (process.env.DATABASE_URL) return;
  const p = path.resolve(".env.local");
  if (!existsSync(p)) return;
  for (const line of readFileSync(p, "utf8").split("\n")) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, "");
  }
}

/**
 * Optional: route the Neon HTTP driver's fetch through undici with hostnames
 * resolved by DNS-over-HTTPS (dns.google) and, as fallback, the UDP servers
 * in DNS_SERVERS. For networks whose resolver refuses neon.tech lookups.
 * No effect when DNS_SERVERS is unset.
 */
async function configureDns() {
  const servers = process.env.DNS_SERVERS?.split(",").map((x) => x.trim()).filter(Boolean);
  if (!servers?.length) return;
  const dns = await import("node:dns");
  const resolver = new dns.Resolver();
  resolver.setServers(servers);
  const { Agent, fetch: ufetch } = await import("undici");
  const cache = new Map<string, string[]>();

  const viaDoh = async (host: string): Promise<string[]> => {
    const r = await fetch(`https://dns.google/resolve?name=${encodeURIComponent(host)}&type=A`);
    if (!r.ok) throw new Error(`DoH HTTP ${r.status}`);
    const j = (await r.json()) as { Answer?: { type: number; data: string }[] };
    const a = (j.Answer ?? []).filter((x) => x.type === 1).map((x) => x.data);
    if (!a.length) throw new Error(`DoH: no A record for ${host}`);
    return a;
  };
  const viaUdp = (host: string) =>
    new Promise<string[]>((res, rej) => resolver.resolve4(host, (e, a) => (e || !a.length ? rej(e ?? new Error("no A")) : res(a))));
  const resolve = async (host: string): Promise<string[]> => {
    const hit = cache.get(host);
    if (hit) return hit;
    let addrs: string[];
    try {
      addrs = await viaDoh(host);
    } catch {
      addrs = await viaUdp(host);
    }
    cache.set(host, addrs);
    return addrs;
  };

  type LookupCb = (err: NodeJS.ErrnoException | null, address?: unknown, family?: number) => void;
  const lookup = (host: string, opts: { all?: boolean } | undefined, cb: LookupCb) => {
    resolve(host).then(
      (addrs) => (opts?.all ? cb(null, addrs.map((address) => ({ address, family: 4 }))) : cb(null, addrs[0], 4)),
      (err: NodeJS.ErrnoException) => cb(err)
    );
  };
  const agent = new Agent({ connect: { lookup } as unknown as ConstructorParameters<typeof Agent>[0] extends { connect?: infer C } ? C : never });
  neonConfig.fetchFunction = (url: string, init: Record<string, unknown>) =>
    ufetch(url, { ...init, dispatcher: agent } as Parameters<typeof ufetch>[1]);
  console.log(`DNS: resolving via DNS-over-HTTPS, fallback ${servers.join(", ")}`);
}

function monthKey(v: unknown): string | null {
  if (v instanceof Date) {
    return `${v.getUTCFullYear()}-${String(v.getUTCMonth() + 1).padStart(2, "0")}-01`;
  }
  if (typeof v === "string") {
    const s = v.trim();
    const iso = s.match(/^(\d{4})-(\d{2})/);
    if (iso) return `${iso[1]}-${iso[2]}-01`;
    const mon = s.match(/^([A-Za-z]{3})[a-z]*[-\s'/]*(\d{2}|\d{4})$/); // "May-2025", "May 25", "May-25"
    if (mon) {
      const mi = MONTHS3.indexOf(mon[1].toUpperCase());
      if (mi >= 0) {
        const y = mon[2].length === 2 ? 2000 + Number(mon[2]) : Number(mon[2]);
        return `${y}-${String(mi + 1).padStart(2, "0")}-01`;
      }
    }
  }
  if (typeof v === "number") {
    // Excel serial date
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-01`;
  }
  return null;
}

function num(v: unknown): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  if (typeof v === "object" && v && "result" in v) return num((v as { result: unknown }).result);
  const n = Number(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function text(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "object" && v && "richText" in v) {
    return (v as { richText: { text: string }[] }).richText.map((t) => t.text).join("");
  }
  return String(v).replace(/[\u200b-\u200d\ufeff]/g, "").trim();
}

function headerIndex(ws: ExcelJS.Worksheet): Map<string, number> {
  const idx = new Map<string, number>();
  ws.getRow(1).eachCell((cell, col) => idx.set(text(cell.value).toUpperCase(), col));
  return idx;
}

function find(idx: Map<string, number>, aliases: string[]): number | undefined {
  for (const a of aliases) {
    const c = idx.get(a);
    if (c) return c;
  }
  return undefined;
}

function need(idx: Map<string, number>, aliases: string[], sheet: string): number {
  const c = find(idx, aliases);
  if (!c) throw new Error(`Sheet "${sheet}": none of [${aliases.join(", ")}] found. Columns: ${[...idx.keys()].join(", ")}`);
  return c;
}

/** Find a sheet by title, falling back to one whose header row has the given column. */
function findSheet(wb: ExcelJS.Workbook, re: RegExp, headerAliases: string[], what: string): ExcelJS.Worksheet {
  const byTitle = wb.worksheets.find((w) => re.test(w.title));
  if (byTitle) return byTitle;
  const byHeader = wb.worksheets.find((w) => find(headerIndex(w), headerAliases) !== undefined);
  if (byHeader) return byHeader;
  throw new Error(`${what} sheet not found. Sheets: ${wb.worksheets.map((w) => w.title || "(untitled)").join(", ")}`);
}

// ---------------------------------------------------------------- parse

function parseMonthly(ws: ExcelJS.Worksheet, pointBasedOnly: boolean) {
  const idx = headerIndex(ws);
  const S = ws.title;
  const c = {
    month: need(idx, COLS.month, S),
    netc: need(idx, COLS.netc, S),
    plaza: need(idx, COLS.plaza, S),
    vc: need(idx, COLS.vc, S),
    etc: need(idx, COLS.etc, S),
    etcCol: need(idx, COLS.etcCol, S),
    ap: need(idx, COLS.ap, S),
    cash: need(idx, COLS.cash, S),
    upi: need(idx, COLS.upi, S),
    exempt: need(idx, COLS.exempt, S),
    loop: find(idx, COLS.loop),
  };
  if (pointBasedOnly && !c.loop) throw new Error(`--point-based-only given but sheet "${S}" has no Closed-Loop/Point-Based column`);

  const agg = new Map<string, MonthlyAgg>();
  const plazaNames = new Map<number, string>();
  const plazaLoop = new Map<number, string>();
  const dropped = new Map<string, number>();
  const loopCounts = new Map<string, number>();
  const months = new Set<string>();
  let rows = 0;
  let skipped = 0;
  let filtered = 0;

  ws.eachRow((row, n) => {
    if (n === 1) return;
    const month = monthKey(row.getCell(c.month).value);
    const code = Number(text(row.getCell(c.netc).value));
    const vc = text(row.getCell(c.vc).value);
    if (!month || !Number.isInteger(code) || code <= 0) {
      skipped++;
      return;
    }
    if (c.loop) {
      const loop = text(row.getCell(c.loop).value);
      loopCounts.set(loop || "(blank)", (loopCounts.get(loop || "(blank)") ?? 0) + 1);
      plazaLoop.set(code, loop);
      if (pointBasedOnly && !/point/i.test(loop)) {
        filtered++;
        return;
      }
    }
    rows++;
    months.add(month);
    const name = text(row.getCell(c.plaza).value);
    if (name && !plazaNames.has(code)) plazaNames.set(code, name);

    const cat = CLASS_MAP[vc];
    if (!cat) {
      dropped.set(vc || "(blank)", (dropped.get(vc || "(blank)") ?? 0) + 1);
      return;
    }
    const key = `${code}|${month}|${cat}`;
    let a = agg.get(key);
    if (!a) {
      a = { plaza_code: code, month, category: cat, etc_cnt: 0, cash_cnt: 0, upi_cnt: 0, exempt_cnt: 0, etc_collection: 0, ap_cnt: 0 };
      agg.set(key, a);
    }
    a.etc_cnt += num(row.getCell(c.etc).value);
    a.cash_cnt += num(row.getCell(c.cash).value);
    a.upi_cnt += num(row.getCell(c.upi).value);
    a.exempt_cnt += num(row.getCell(c.exempt).value);
    a.etc_collection += num(row.getCell(c.etcCol).value);
    a.ap_cnt += num(row.getCell(c.ap).value);
  });

  return { rows, skipped, filtered, loopCounts, months: [...months].sort(), agg: [...agg.values()], plazaNames, plazaLoop, dropped };
}

function parseAnnualPass(ws: ExcelJS.Worksheet, pointBasedOnly: boolean) {
  const idx = headerIndex(ws);
  const S = ws.title;
  const c = {
    month: need(idx, COLS.month, S),
    netc: need(idx, COLS.netc, S),
    comp: need(idx, COLS.comp, S),
    qual: find(idx, COLS.qual),
    loop: find(idx, COLS.loop),
  };
  const out = new Map<string, ApRow>();
  let skipped = 0;
  let dupes = 0;
  let filtered = 0;
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const month = monthKey(row.getCell(c.month).value);
    const code = Number(text(row.getCell(c.netc).value));
    if (!month || !Number.isInteger(code) || code <= 0) {
      skipped++;
      return;
    }
    if (pointBasedOnly && c.loop && !/point/i.test(text(row.getCell(c.loop).value))) {
      filtered++;
      return;
    }
    const key = `${code}|${month}`;
    const comp = num(row.getCell(c.comp).value);
    const qual = c.qual ? num(row.getCell(c.qual).value) : 0;
    const existing = out.get(key);
    if (existing) {
      dupes++;
      existing.compensation += comp; // same plaza-month twice: sum
      existing.qualified_txn += qual;
    } else {
      out.set(key, { plaza_code: code, month, compensation: comp, qualified_txn: qual });
    }
  });
  return { rows: [...out.values()], skipped, dupes, filtered };
}

// ---------------------------------------------------------------- main

async function main() {
  const args = process.argv.slice(2);
  const file = args.find((a) => !a.startsWith("--"));
  const apply = args.includes("--apply");
  const addMissing = args.includes("--add-missing-plazas");
  const pointBasedOnly = args.includes("--point-based-only");
  if (!file) {
    console.error("Usage: npx tsx scripts/load-data.mts <workbook.xlsx> [--apply] [--add-missing-plazas] [--point-based-only]");
    process.exit(1);
  }

  console.log(`Reading ${file} ...`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const etcWs = findSheet(wb, /etc.*non.?etc/i, COLS.vc, "ETC & Non-ETC");
  const apWs = findSheet(wb, /^annual.?pass/i, COLS.comp, "Annual Pass");
  console.log(`sheets: "${etcWs.title || "(untitled)"}", "${apWs.title || "(untitled)"}"`);

  const monthly = parseMonthly(etcWs, pointBasedOnly);
  const ap = parseAnnualPass(apWs, pointBasedOnly);

  const byCat = new Map<string, number>();
  for (const r of monthly.agg) byCat.set(r.category, (byCat.get(r.category) ?? 0) + 1);

  console.log("\n== ETC & Non-ETC ==");
  console.log(`source rows: ${monthly.rows} (skipped ${monthly.skipped} without month/NETC` +
    (monthly.filtered ? `, filtered out ${monthly.filtered} non point-based` : "") + ")");
  if (monthly.loopCounts.size) {
    const loopPlazas = new Map<string, number>();
    for (const l of monthly.plazaLoop.values()) loopPlazas.set(l, (loopPlazas.get(l) ?? 0) + 1);
    console.log("plaza types in file:", Object.fromEntries(loopPlazas), pointBasedOnly ? "(loading Point-Based only)" : "(loading all)");
  }
  console.log(`months: ${monthly.months[0]} .. ${monthly.months[monthly.months.length - 1]} (${monthly.months.length})`);
  console.log(`plazas: ${monthly.plazaNames.size}`);
  console.log(`rows to load into plaza_monthly_data: ${monthly.agg.length}`);
  for (const cat of CATEGORIES) console.log(`  ${cat}: ${byCat.get(cat) ?? 0}`);
  console.log("dropped vehicle classes:", Object.fromEntries(monthly.dropped));
  console.log("\n== Annual Pass ==");
  console.log(`rows to load into annual_pass_monthly: ${ap.rows.length} (skipped ${ap.skipped}, merged duplicates ${ap.dupes}` +
    (ap.filtered ? `, filtered out ${ap.filtered} non point-based` : "") + ")");

  loadEnv();
  await configureDns();
  if (!process.env.DATABASE_URL) {
    console.log("\nDATABASE_URL not set: parse-only dry run finished. Add it to .env.local to compare with / write to the DB.");
    return;
  }
  const sql = neon(process.env.DATABASE_URL);

  // Schema check
  const cols = await sql`
    SELECT table_name, column_name, data_type
    FROM information_schema.columns
    WHERE table_schema = 'public' AND table_name IN ('pb_plazas', 'plaza_monthly_data', 'annual_pass_monthly')
    ORDER BY table_name, ordinal_position`;
  console.log("\n== DB schema ==");
  for (const t of ["pb_plazas", "plaza_monthly_data", "annual_pass_monthly"]) {
    console.log(`  ${t}: ${cols.filter((c) => c.table_name === t).map((c) => `${c.column_name} ${c.data_type}`).join(", ")}`);
  }
  const monthlyCols = new Set(cols.filter((c) => c.table_name === "plaza_monthly_data").map((c) => c.column_name as string));
  for (const c of ["plaza_code", "month", "category", "etc_cnt", "cash_cnt", "upi_cnt", "exempt_cnt", "tmcc_total", "etc_collection", "ap_cnt"]) {
    if (!monthlyCols.has(c)) throw new Error(`plaza_monthly_data is missing column ${c}`);
  }
  const hasApComp = monthlyCols.has("ap_compensation");

  const [counts] = await sql`
    SELECT (SELECT count(*) FROM pb_plazas)::int AS plazas,
           (SELECT count(*) FROM plaza_monthly_data)::int AS monthly,
           (SELECT count(*) FROM annual_pass_monthly)::int AS ap,
           (SELECT min(month) FROM plaza_monthly_data) AS min_month,
           (SELECT max(month) FROM plaza_monthly_data) AS max_month`;
  console.log(`\nDB now: pb_plazas=${counts.plazas}, plaza_monthly_data=${counts.monthly} (${counts.min_month} .. ${counts.max_month}), annual_pass_monthly=${counts.ap}`);

  const known = new Set((await sql`SELECT plaza_code FROM pb_plazas`).map((r) => Number(r.plaza_code)));
  const missing = [...monthly.plazaNames].filter(([code]) => !known.has(code));
  console.log(`plazas in workbook not in pb_plazas: ${missing.length}` + (missing.length ? ` e.g. ${missing.slice(0, 5).map(([c, n]) => `${c} ${n}`).join("; ")}` : ""));
  if (missing.length && !addMissing) {
    console.log("  (their monthly rows will load but won't be searchable until added; pass --add-missing-plazas to insert code+name)");
  }

  if (!apply) {
    console.log("\nDry run only. Re-run with --apply to replace the DB contents.");
    return;
  }

  // ---- Apply: delete + insert in one transaction ----
  console.log("\nApplying ...");
  const queries = [];
  queries.push(sql`DELETE FROM annual_pass_monthly`);
  queries.push(sql`DELETE FROM plaza_monthly_data`);

  if (addMissing && missing.length) {
    queries.push(sql`
      INSERT INTO pb_plazas (plaza_code, plaza_name)
      SELECT * FROM unnest(${missing.map(([c]) => c)}::int[], ${missing.map(([, n]) => n)}::text[])
      ON CONFLICT DO NOTHING`);
  }

  for (let i = 0; i < monthly.agg.length; i += BATCH) {
    const b = monthly.agg.slice(i, i + BATCH);
    const codes = b.map((r) => r.plaza_code);
    const months = b.map((r) => r.month);
    const cats = b.map((r) => r.category);
    const etc = b.map((r) => r.etc_cnt);
    const cash = b.map((r) => r.cash_cnt);
    const upi = b.map((r) => r.upi_cnt);
    const ex = b.map((r) => r.exempt_cnt);
    const tmcc = b.map((r) => r.etc_cnt + r.cash_cnt + r.upi_cnt + 0.5 * r.exempt_cnt);
    const col = b.map((r) => r.etc_collection);
    const apc = b.map((r) => r.ap_cnt);
    if (hasApComp) {
      queries.push(sql`
        INSERT INTO plaza_monthly_data
          (plaza_code, month, category, etc_cnt, cash_cnt, upi_cnt, exempt_cnt, tmcc_total, etc_collection, ap_cnt, ap_compensation)
        SELECT *, 0 FROM unnest(
          ${codes}::int[], ${months}::date[], ${cats}::text[], ${etc}::bigint[], ${cash}::bigint[],
          ${upi}::bigint[], ${ex}::bigint[], ${tmcc}::numeric[], ${col}::numeric[], ${apc}::bigint[])`);
    } else {
      queries.push(sql`
        INSERT INTO plaza_monthly_data
          (plaza_code, month, category, etc_cnt, cash_cnt, upi_cnt, exempt_cnt, tmcc_total, etc_collection, ap_cnt)
        SELECT * FROM unnest(
          ${codes}::int[], ${months}::date[], ${cats}::text[], ${etc}::bigint[], ${cash}::bigint[],
          ${upi}::bigint[], ${ex}::bigint[], ${tmcc}::numeric[], ${col}::numeric[], ${apc}::bigint[])`);
    }
  }

  for (let i = 0; i < ap.rows.length; i += BATCH) {
    const b = ap.rows.slice(i, i + BATCH);
    queries.push(sql`
      INSERT INTO annual_pass_monthly (plaza_code, month, compensation, qualified_txn)
      SELECT * FROM unnest(${b.map((r) => r.plaza_code)}::int[], ${b.map((r) => r.month)}::date[],
                           ${b.map((r) => r.compensation)}::numeric[], ${b.map((r) => Math.round(r.qualified_txn))}::bigint[])`);
  }

  console.log(`running ${queries.length} statements in one transaction ...`);
  await sql.transaction(queries);

  const [after] = await sql`
    SELECT (SELECT count(*) FROM pb_plazas)::int AS plazas,
           (SELECT count(*) FROM plaza_monthly_data)::int AS monthly,
           (SELECT count(*) FROM annual_pass_monthly)::int AS ap,
           (SELECT min(month) FROM plaza_monthly_data) AS min_month,
           (SELECT max(month) FROM plaza_monthly_data) AS max_month`;
  console.log(`DB after: pb_plazas=${after.plazas}, plaza_monthly_data=${after.monthly} (${after.min_month} .. ${after.max_month}), annual_pass_monthly=${after.ap}`);
  console.log("Done.");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
