// Runs lib/apc.ts against the dev server's API data and prints the result,
// for diffing with verify_apc.py (the independent Python implementation).
// Usage: node --experimental-strip-types scripts/check.mts <code> <uptoMonth> <growth> [mf date:factor ...]

import { computeApc, type MfEntry } from "../lib/apc.ts";

const [code, upto, growthArg, ...mfSpecs] = process.argv.slice(2);
const mfEntries: MfEntry[] = mfSpecs.map((s) => {
  const i = s.lastIndexOf(":");
  return { date: s.slice(0, i), factor: Number(s.slice(i + 1)) };
});

const res = await fetch(`http://localhost:3000/api/plazas/${code}`);
if (!res.ok) throw new Error(`API ${res.status}`);
const data = await res.json();

const result = computeApc(data.rows, upto, mfEntries, Number(growthArg));
console.log(JSON.stringify({ ...result, plaza: data.plaza }, null, 2));
