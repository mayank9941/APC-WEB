"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import {
  MfEntry,
  MonthlyRow,
  computeApc,
  fmtCr,
  fmtIN,
  lastCompletedMonth,
  monthLabel,
} from "@/lib/apc";
import CollectionChart from "./CollectionChart";
import { exportApcXlsx } from "@/lib/exportXlsx";

interface ApiResponse {
  plaza: { code: number; name: string };
  months: string[];
  rows: MonthlyRow[];
}

export default function ApcReport({ code }: { code: string }) {
  const [data, setData] = useState<ApiResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [uptoMonth, setUptoMonth] = useState<string>("");
  const [mfEntries, setMfEntries] = useState<MfEntry[]>([]);
  const [growth, setGrowth] = useState<number>(0);

  useEffect(() => {
    fetch(`/api/plazas/${code}`)
      .then(async (r) => {
        if (!r.ok) {
          const body = await r.json().catch(() => ({}));
          throw new Error(body.error || `HTTP ${r.status}`);
        }
        return r.json();
      })
      .then((d: ApiResponse) => {
        setData(d);
        // Default "up to" month: latest available month before the current
        // month; if none qualifies, the latest available month.
        const completed = lastCompletedMonth();
        const candidates = d.months.filter((m) => m <= completed);
        const def = candidates.length
          ? candidates[candidates.length - 1]
          : d.months[d.months.length - 1];
        if (def) setUptoMonth(def);
      })
      .catch((e) => setError(e.message));
  }, [code]);

  const result = useMemo(() => {
    if (!data || !uptoMonth) return null;
    return computeApc(data.rows, uptoMonth, mfEntries, growth);
  }, [data, uptoMonth, mfEntries, growth]);

  if (error) {
    return (
      <div>
        <Link href="/" className="back-link">← Back to search</Link>
        <div className="error-box">{error}</div>
      </div>
    );
  }
  if (!data || !result) {
    return <div className="loading">Loading plaza data…</div>;
  }
  if (data.months.length === 0) {
    return (
      <div>
        <Link href="/" className="back-link">← Back to search</Link>
        <div className="report-header">
          <h1>{data.plaza.name}</h1>
          <span className="plaza-code">Plaza code: {data.plaza.code}</span>
        </div>
        <div className="error-box">No monthly data available for this plaza.</div>
      </div>
    );
  }

  const { table1, table2, table3, months } = result;

  const setMf = (i: number, patch: Partial<MfEntry>) => {
    setMfEntries((list) =>
      list.map((e, j) => (j === i ? { ...e, ...patch } : e))
    );
  };

  return (
    <div>
      <Link href="/" className="back-link">← Back to search</Link>

      <div className="report-header">
        <h1>{data.plaza.name}</h1>
        <span className="plaza-code">Plaza code: {data.plaza.code}</span>
      </div>
      <p className="report-subtitle">
        APC for {monthLabel(uptoMonth)} — window {monthLabel(months[0])} to{" "}
        {monthLabel(months[11])} ({table2.monthsUsed} month
        {table2.monthsUsed === 1 ? "" : "s"} with data)
      </p>

      <div className="apc-hero top">
        <div className="hero-main">
          <span className="label">APC-2</span>
          <span className="value">₹ {table3.apcCr.toFixed(2)} Cr.</span>
          <span className="per-day">
            [Rs {fmtIN(table3.apcPerDay)} per day]
          </span>
        </div>
        <button
          className="btn export"
          onClick={() =>
            exportApcXlsx(data.plaza, result, uptoMonth, mfEntries, growth)
          }
        >
          ⬇ Export to Excel
        </button>
      </div>

      <div className="controls">
        <div className="control-group">
          <label htmlFor="upto">Calculate up to</label>
          <select
            id="upto"
            value={uptoMonth}
            onChange={(e) => setUptoMonth(e.target.value)}
          >
            {[...data.months].reverse().map((m) => (
              <option key={m} value={m}>
                {monthLabel(m)}
              </option>
            ))}
          </select>
          <span className="hint">12-month window ends at this month</span>
        </div>

        <div className="control-group">
          <label htmlFor="growth">Traffic growth (%)</label>
          <select
            id="growth"
            value={growth}
            onChange={(e) => setGrowth(Number(e.target.value))}
          >
            <option value={0}>0%</option>
            <option value={5}>5%</option>
          </select>
          <span className="hint">Default 0%</span>
        </div>

        <div className="control-group">
          <span className="group-label">MF (fee-rate revision factors)</span>
          {mfEntries.map((e, i) => (
            <div className="mf-row" key={i}>
              <input
                type="date"
                value={e.date}
                onChange={(ev) => setMf(i, { date: ev.target.value })}
                aria-label={`MF ${i + 1} effective date`}
              />
              <input
                type="number"
                step="0.0001"
                min="0"
                value={e.factor}
                onChange={(ev) => setMf(i, { factor: Number(ev.target.value) })}
                aria-label={`MF ${i + 1} factor`}
              />
              <button
                className="btn remove"
                onClick={() =>
                  setMfEntries((list) => list.filter((_, j) => j !== i))
                }
              >
                ✕
              </button>
            </div>
          ))}
          <button
            className="btn"
            onClick={() =>
              setMfEntries((list) => [...list, { date: "", factor: 1 }])
            }
          >
            + Add MF
          </button>
          <span className="hint">
            Factor = new rate ÷ old rate, applied from its effective date
            (day-weighted in the revision month)
          </span>
        </div>
      </div>

      {/* ---------------- Table 1 ---------------- */}
      <section className="card">
        <h2>
          Table 1 — Calculation of ETC Penetration (Period: 12 Completed
          Calendar Months)
        </h2>
        <p className="card-note">
          Summed over {monthLabel(months[0])} – {monthLabel(months[11])}
        </p>
        <div className="table-wrap">
          <table className="apc">
            <thead>
              <tr>
                <th>Category of Vehicle</th>
                <th>
                  Mapper Vehicle Description as per NPCI
                </th>
                <th className="num">
                  Total ETC Transactions as per NPCI
                  <span className="formula">A</span>
                </th>
                <th className="num">
                  Cash + UPI Transactions
                </th>
                <th className="num">
                  50% of Exempted
                </th>
                <th className="num">
                  Total Transactions as per TMCC (ETC+Cash+UPI+50% of exempted)
                  <span className="formula">B</span>
                </th>
                <th className="num">
                  % ETC Penetration
                  <span className="formula">C = A×100/B</span>
                </th>
                <th className="num">
                  ETC Collection as per NPCI + Annual Pass Compensation
                  <span className="formula">D</span>
                </th>
                <th className="num">
                  Derived Total Collection
                  <span className="formula">E = D×100/C</span>
                </th>
                <th className="num">
                  % Contribution
                  <span className="formula">F = E/ΣE</span>
                </th>
                <th className="num">
                  Revenue Contribution (Roundoff)
                  <span className="formula">G</span>
                </th>
              </tr>
            </thead>
            <tbody>
              {table1.rows.map((r) => (
                <tr key={r.category}>
                  <td>{r.category}</td>
                  <td>{r.npciCodes}</td>
                  <td className="num">{fmtIN(r.etcTxn)}</td>
                  <td className="num">{fmtIN(r.cashUpiTxn)}</td>
                  <td className="num">{fmtIN(r.exemptHalf, 1)}</td>
                  <td className="num">{fmtIN(r.totalTxn, 1)}</td>
                  <td className="num">{r.penetration.toFixed(2)}</td>
                  <td className="num">{fmtIN(r.etcCollectionWithAp)}</td>
                  <td className="num">{fmtIN(r.derivedTotal)}</td>
                  <td className="num">{(r.contribution * 100).toFixed(2)}%</td>
                  <td className="num">{r.contributionRounded.toFixed(2)}</td>
                </tr>
              ))}
              <tr className="total">
                <td colSpan={6}>
                  Weighted Average ETC Penetration — Σ(C×D)/Σ(D)
                </td>
                <td className="num">{table1.weightedPenetration.toFixed(2)}</td>
                <td colSpan={3}></td>
                <td className="num">{table1.contributionTotal.toFixed(2)}</td>
              </tr>
            </tbody>
          </table>
        </div>
        {!table1.roundOffOk && (
          <div className="warn">
            Please CHECK Round off Contribution (total ≠ 1.00)
          </div>
        )}
      </section>

      {/* ---------------- Table 2 + chart ---------------- */}
      <section className="card">
        <h2>Table 2 — Monthly Average Daily Collection (Seasonal Factor)</h2>
        <p className="card-note">
          Months without data are excluded from the averages
        </p>
        <div className="table-wrap">
          <table className="apc">
            <thead>
              <tr>
                <th>Month</th>
                <th className="num">
                  Monthly Average Daily Annual Pass Compensation
                </th>
                <th className="num">ETC Collection (avg daily)</th>
                <th className="num">
                  Monthly Average Daily ETC Collection as per Actual Fee Rates +
                  Annual Pass Compensation
                  <span className="formula">A</span>
                </th>
                <th className="num">
                  MF Multiplier
                </th>
                <th className="num">
                  Monthly Avg ETC Collection as per revised Fee Rates
                  <span className="formula">B = A × MF</span>
                </th>
              </tr>
            </thead>
            <tbody>
              {table2.rows.map((r) => (
                <tr key={r.month} className={r.hasData ? "" : "excluded"}>
                  <td>{monthLabel(r.month)}</td>
                  <td className="num">{r.hasData ? fmtIN(r.apDaily, 2) : "—"}</td>
                  <td className="num">{r.hasData ? fmtIN(r.etcDaily, 2) : "—"}</td>
                  <td className="num">{r.hasData ? fmtIN(r.totalDaily, 2) : "—"}</td>
                  <td className="num">{r.mfMultiplier.toFixed(4)}</td>
                  <td className="num">
                    {r.hasData ? fmtIN(r.normalizedDaily, 2) : "no data"}
                  </td>
                </tr>
              ))}
              <tr className="average">
                <td>Average</td>
                <td className="num">{fmtIN(table2.avgApDaily, 2)}</td>
                <td className="num"></td>
                <td className="num">{fmtIN(table2.avgTotalDaily, 2)}</td>
                <td className="num"></td>
                <td className="num">{fmtIN(table2.avgNormalizedDaily, 2)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      <section className="card">
        <h2>Average Daily ETC Collection (normalized to current fee rates)</h2>
        <p className="card-note">₹ per day, by month — column B of Table 2</p>
        <CollectionChart rows={table2.rows} />
        <p className="chart-note">
          Gaps indicate months with no data (excluded from the calculation).
        </p>
      </section>

      {/* ---------------- Table 3 ---------------- */}
      <section className="card">
        <h2>Table 3 — Calculation of APC-2</h2>
        <div className="table-wrap">
          <table className="apc">
            <tbody>
              <tr>
                <td>1</td>
                <td>Average Daily FASTag Collection</td>
                <td className="num">{fmtIN(table3.avgDailyFastagCollection, 2)}</td>
                <td></td>
              </tr>
              <tr>
                <td>2</td>
                <td>FASTag Penetration (In decimal)</td>
                <td className="num">{table3.fastagPenetration.toFixed(2)}</td>
                <td></td>
              </tr>
              <tr>
                <td>3</td>
                <td>Seasonal Factor (In decimal)</td>
                <td className="num">-</td>
                <td></td>
              </tr>
              <tr>
                <td>4</td>
                <td>Annual Average Daily Collection</td>
                <td className="num">{fmtIN(table3.annualAvgDailyCollection, 2)}</td>
                <td></td>
              </tr>
              <tr>
                <td>5</td>
                <td>Annual Expected Collection</td>
                <td className="num">{fmtIN(table3.annualExpectedCollection, 2)}</td>
                <td className="num">{fmtCr(table3.annualExpectedCollection)}</td>
              </tr>
              <tr>
                <td>6</td>
                <td>Traffic Growth (in %)</td>
                <td className="num">{table3.trafficGrowthPct}</td>
                <td></td>
              </tr>
              <tr>
                <td>7</td>
                <td>Net Expected Collection</td>
                <td className="num">{fmtIN(table3.netExpectedCollection, 2)}</td>
                <td className="num">{fmtCr(table3.netExpectedCollection)}</td>
              </tr>
              <tr>
                <td>8</td>
                <td>Less Administrative Charges</td>
                <td className="num">{fmtIN(table3.adminCharges, 2)}</td>
                <td className="num">{fmtCr(table3.adminCharges)}</td>
              </tr>
              <tr>
                <td>9</td>
                <td>Less Contractor Profit @5%</td>
                <td className="num">{fmtIN(table3.contractorProfit, 2)}</td>
                <td className="num">{fmtCr(table3.contractorProfit)}</td>
              </tr>
              <tr className="total">
                <td>10</td>
                <td>APC-2 (Cr)</td>
                <td className="num">{table3.apcCr.toFixed(2)}</td>
                <td className="num">{table3.apcCr.toFixed(2)} Cr.</td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}
