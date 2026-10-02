// Renders a sample PDF + DOCX from synthetic data to /tmp for visual checks.
// Usage: npx tsx scripts/render-sample.mts
import { computeApc, CATEGORIES } from "../lib/apc.ts";
import type { MonthlyRow } from "../lib/apc.ts";
import { buildReportModel } from "../lib/reportModel.ts";
import { buildApcPdf } from "../lib/exportPdf.ts";
import { buildApcDocx } from "../lib/exportDocx.ts";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as docx from "docx";
import { writeFileSync } from "node:fs";

const months: string[] = [];
for (let i = 0; i < 12; i++) { const d = new Date(2025, 8 + i, 1); months.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`); }
const rows: MonthlyRow[] = [];
const shares = [0.4133, 0.1266, 0.2266, 0.0833, 0.1133, 0.0369];
months.forEach((m, mi) => {
  if (m === "2026-01" || m === "2026-05") return; // missing months
  CATEGORIES.forEach((cat, ci) => {
    const etc = (30_000_000 - mi * 600_000) * shares[ci];
    rows.push({ month: m, category: cat, etc_cnt: 1000, cash_cnt: 100, upi_cnt: 50, exempt_cnt: 20,
      tmcc_total: 1160, etc_collection: etc, ap_cnt: 10, ap_compensation: ci === 0 ? 900_000 : 0 });
  });
});
const mf = [{ date: "2025-10-20", factor: 1.05 }];
const r = computeApc(rows, months[11], mf);
const t2 = r.table2;
console.log("monthsUsed", t2.monthsUsed, "imputed", t2.monthsImputed, "negative", t2.negativeTrend);
const jan = t2.rows.find(x => x.month === "2026-01")!;
const actual = t2.rows.filter(x => x.hasData);
const avgEtcActual = actual.reduce((s, x) => s + x.etcDaily!, 0) / actual.length;
console.log("Jan imputed:", jan.imputed, "etcDaily == avg(actual):", Math.abs(jan.etcDaily! - avgEtcActual) < 1e-9, "B = A*MF:", Math.abs(jan.normalizedDaily! - jan.etcDaily! * jan.mfMultiplier) < 1e-9);
console.log("avgNormalized over 12:", t2.avgNormalizedDaily.toFixed(2), "APC:", r.table3.apcCr);

const model = buildReportModel({ code: 330085, name: "Hathitala" }, r, months[11], mf, 5);
console.log("Table 3 rows:"); for (const row of model.table3.body) console.log("  ", row.join(" | "));
console.log("Table 2 sample:", model.table2.body[4].join(" | "));
console.log("alert:", model.alert);

const doc = buildApcPdf(jsPDF, autoTable as unknown as Parameters<typeof buildApcPdf>[1], model);
writeFileSync("/tmp/apc_test.pdf", Buffer.from(doc.output("arraybuffer")));
console.log("pdf pages:", doc.getNumberOfPages());

const wdoc = buildApcDocx(docx, model);
writeFileSync("/tmp/apc_test.docx", await docx.Packer.toBuffer(wdoc));
console.log("docx written");
