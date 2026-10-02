"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import {
  DEFAULT_INCREMENT_PCT,
  DEFAULT_TRAFFIC_GROWTH_PCT,
  MfEntry,
  MonthlyRow,
  PlazaInfo,
  computeApc,
  fmtIN,
  fmtMoney,
  lastCompletedMonth,
  mfFactor,
  monthLabel,
} from "@/lib/apc";
import CollectionChart from "./CollectionChart";
import { exportApcXlsx } from "@/lib/exportXlsx";
import { exportApcPdf } from "@/lib/exportPdf";
import { exportApcDocx } from "@/lib/exportDocx";

interface ApiResponse {
  plaza: PlazaInfo;
  months: string[];
  rows: MonthlyRow[];
}

export default function ApcReport({ code }: { code: string }) {
  const [data, setData] = useState<ApiResponse | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [uptoMonth, setUptoMonth] = useState<string>("");
  const [mfEntries, setMfEntries] = useState<MfEntry[]>([]);
  const [growth, setGrowth] = useState<number>(DEFAULT_TRAFFIC_GROWTH_PCT);
  const [increment, setIncrement] = useState<number>(DEFAULT_INCREMENT_PCT);

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
    return computeApc(data.rows, uptoMonth, mfEntries, growth, increment);
  }, [data, uptoMonth, mfEntries, growth, increment]);

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

  const { table1, table2, table3, table4, months } = result;
  const plazaDetails = [
    ["Lanes", data.plaza.lanes],
    ["Type", data.plaza.type],
    ["PIU", data.plaza.piu],
    ["RO", data.plaza.ro],
    ["State", data.plaza.state],
  ].filter(([, v]) => v !== null && v !== undefined && v !== "") as [string, string | number][];
  const showGrowthAlert = table2.negativeTrend && growth > 0;

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
      {plazaDetails.length > 0 && (
        <p className="plaza-details">
          {plazaDetails.map(([k, v]) => (
            <span key={k}>
              <b>{k}:</b> {v}
            </span>
          ))}
        </p>
      )}
      <p className="report-subtitle">
        APC for {monthLabel(uptoMonth)} — window {monthLabel(months[0])} to{" "}
        {monthLabel(months[11])} ({table2.monthsUsed} month
        {table2.monthsUsed === 1 ? "" : "s"} with data
        {table2.monthsImputed > 0 &&
          `, ${table2.monthsImputed} filled with the average`}
        )
      </p>

      <div className="apc-hero top">
        <div className="hero-main">
          <span className="label">APC-2</span>
          <span className="value">₹ {table3.apcCr.toFixed(2)} Cr.</span>
          <span className="per-day">
            [Rs {fmtIN(table3.apcPerDay)} per day]
          </span>
        </div>
        <div className="export-group">
          <button
            className="btn export"
            onClick={() =>
              exportApcPdf(data.plaza, result, uptoMonth, mfEntries, growth)
            }
          >
            ⬇ PDF
          </button>
          <button
            className="btn export"
            onClick={() =>
              exportApcDocx(data.plaza, result, uptoMonth, mfEntries, growth)
            }
          >
            ⬇ Word
          </button>
          <button
            className="btn export"
            onClick={() =>
              exportApcXlsx(data.plaza, result, uptoMonth, mfEntries, growth)
            }
          >
            ⬇ Excel
          </button>
        </div>
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
            <option value={5}>5%</option>
            <option value={0}>0%</option>
          </select>
          <span className="hint">
            Default {DEFAULT_TRAFFIC_GROWTH_PCT}%; take 0% when the collection
            trend is negative
          </span>
        </div>

        <div className="control-group">
          <label htmlFor="increment">Increment for APC-3 (%)</label>
          <input
            id="increment"
            type="number"
            step="0.5"
            min="0"
            value={increment}
            onChange={(e) => setIncrement(Number(e.target.value))}
          />
          <span className="hint">Default {DEFAULT_INCREMENT_PCT}% on net remittance</span>
        </div>

        <div className="control-group mf-group">
          <span className="group-label">MF (fee-rate revisions)</span>
          {mfEntries.length > 0 && (
            <div className="mf-row mf-head">
              <span>Effective date</span>
              <span>Old remittance</span>
              <span>New remittance</span>
              <span>MF</span>
              <span></span>
            </div>
          )}
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
                step="1"
                min="0"
                placeholder="Old ₹/day"
                value={e.oldRemittance ?? ""}
                onChange={(ev) =>
                  setMf(i, { oldRemittance: ev.target.value === "" ? undefined : Number(ev.target.value) })
                }
                aria-label={`MF ${i + 1} old remittance`}
              />
              <input
                type="number"
                step="1"
                min="0"
                placeholder="New ₹/day"
                value={e.newRemittance ?? ""}
                onChange={(ev) =>
                  setMf(i, { newRemittance: ev.target.value === "" ? undefined : Number(ev.target.value) })
                }
                aria-label={`MF ${i + 1} new remittance`}
              />
              <span className="mf-factor" aria-label={`MF ${i + 1} factor`}>
                {mfFactor(e) > 0 ? mfFactor(e).toFixed(4) : "—"}
              </span>
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
            MF = new remittance ÷ old remittance, applied from its effective
            date (day-weighted in the revision month). The latest new
            remittance also drives Table 4.
          </span>
        </div>
      </div>

      {showGrowthAlert && (
        <div className="alert" role="alert">
          <strong>Traffic growth should be taken as 0%.</strong> The average
          daily ETC collection shows a negative trend over{" "}
          {monthLabel(months[0])} – {monthLabel(months[11])} (about ₹
          {fmtIN(Math.abs(table2.trendSlope), 0)} per day less each month).
          <button className="btn alert-action" onClick={() => setGrowth(0)}>
            Set growth to 0%
          </button>
        </div>
      )}

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
                  Total Cash + UPI @TMCC
                  <span className="formula">B1</span>
                </th>
                <th className="num">
                  Exempted Transactions @TMCC
                  <span className="formula">B2</span>
                </th>
                <th className="num">
                  Total Transactions for APC
                  <span className="formula">B = A + B1 + 50% of B2</span>
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
                  Revenue Contribution (Roundoff) %
                  <span className="formula">G = round(F), ΣG = 100%</span>
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
                  <td className="num">{fmtIN(r.exemptTxn)}</td>
                  <td className="num">{fmtIN(r.totalTxn, 1)}</td>
                  <td className="num">{r.penetration.toFixed(2)}</td>
                  <td className="num">{fmtIN(r.etcCollectionWithAp)}</td>
                  <td className="num">{fmtIN(r.derivedTotal)}</td>
                  <td className="num">{(r.contribution * 100).toFixed(2)}%</td>
                  <td className="num">{r.contributionRounded}%</td>
                </tr>
              ))}
              <tr className="total">
                <td colSpan={6}>
                  Weighted Average ETC Penetration — Σ(C×D)/Σ(D)
                </td>
                <td className="num">{table1.weightedPenetration.toFixed(2)}</td>
                <td colSpan={2}></td>
                <td className="num">{(table1.rows.reduce((s, r) => s + r.contribution, 0) * 100).toFixed(2)}%</td>
                <td className="num">{table1.contributionTotal}%</td>
              </tr>
            </tbody>
          </table>
        </div>
        {!table1.roundOffOk && (
          <div className="warn">
            Please CHECK Round off Contribution (total ≠ 100%)
          </div>
        )}
      </section>

      {/* ---------------- Table 2 + chart ---------------- */}
      <section className="card">
        <h2>Table 2 — Monthly Average Daily ETC Collection</h2>
        <p className="card-note">
          Annual Pass compensation is shown for information only and is not
          part of B. Months marked <em>avg</em> had no data and take the
          average of the months that do.
        </p>
        <div className="table-wrap">
          <table className="apc">
            <thead>
              <tr>
                <th>Month</th>
                <th className="num">
                  Monthly Average Daily Annual Pass Compensation
                </th>
                <th className="num">
                  ETC Collection (avg daily)
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
              {table2.rows.map((r) => {
                const filled = r.etcDaily !== null;
                return (
                  <tr
                    key={r.month}
                    className={r.imputed ? "imputed" : filled ? "" : "excluded"}
                  >
                    <td>
                      {monthLabel(r.month)}
                      {r.imputed && <span className="badge">avg</span>}
                    </td>
                    <td className="num">{filled ? fmtIN(r.apDaily, 2) : "—"}</td>
                    <td className="num">{filled ? fmtIN(r.etcDaily, 2) : "—"}</td>
                    <td className="num">{r.mfMultiplier.toFixed(4)}</td>
                    <td className="num">
                      {filled ? fmtIN(r.normalizedDaily, 2) : "no data"}
                    </td>
                  </tr>
                );
              })}
              <tr className="average">
                <td>Average</td>
                <td className="num">{fmtIN(table2.avgApDaily, 2)}</td>
                <td className="num">{fmtIN(table2.avgEtcDaily, 2)}</td>
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
          Hollow points are months with no data, filled with the average of the
          months that have data.
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
                <td className="num">{fmtMoney(table3.avgDailyFastagCollection)}</td>
              </tr>
              <tr>
                <td>2</td>
                <td>FASTag Penetration (in %)</td>
                <td className="num">{table3.fastagPenetration.toFixed(2)}</td>
              </tr>
              <tr>
                <td>3</td>
                <td>Annual Average Daily Collection</td>
                <td className="num">{fmtMoney(table3.annualAvgDailyCollection)}</td>
              </tr>
              <tr>
                <td>4</td>
                <td>Annual Expected Collection</td>
                <td className="num">{fmtMoney(table3.annualExpectedCollection)}</td>
              </tr>
              <tr>
                <td>5</td>
                <td>Traffic Growth (in %)</td>
                <td className="num">{table3.trafficGrowthPct}%</td>
              </tr>
              <tr>
                <td>6</td>
                <td>Net Expected Collection</td>
                <td className="num">{fmtMoney(table3.netExpectedCollection)}</td>
              </tr>
              <tr>
                <td>7</td>
                <td>Less Administrative Charges</td>
                <td className="num">{fmtMoney(table3.adminCharges)}</td>
              </tr>
              <tr>
                <td>8</td>
                <td>Less Contractor Profit @5%</td>
                <td className="num">{fmtMoney(table3.contractorProfit)}</td>
              </tr>
              <tr className="total">
                <td>9</td>
                <td>APC-2</td>
                <td className="num">
                  {fmtMoney(table3.apcCr * 1e7)}
                  <span className="sub">₹ {fmtIN(table3.apcPerDay)} per day</span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>

      {/* ---------------- Table 4 ---------------- */}
      <section className="card">
        <h2>Table 4 — Calculation of APC-3</h2>
        {table4 ? (
          <div className="table-wrap">
            <table className="apc">
              <tbody>
                <tr>
                  <td>Present Remittance (Yearly)</td>
                  <td className="num">
                    {fmtMoney(table4.presentRemittanceYearly)}
                    <span className="sub">₹ {fmtIN(table4.newRemittanceDaily)} per day × 365</span>
                  </td>
                </tr>
                <tr>
                  <td>Present Annual Pass Compensation (Yearly)</td>
                  <td className="num">
                    {fmtMoney(table4.presentApCompensationYearly)}
                    <span className="sub">₹ {fmtIN(table4.latestApDaily)} per day × 365</span>
                  </td>
                </tr>
                <tr>
                  <td>Net Remittance</td>
                  <td className="num">{fmtMoney(table4.netRemittance)}</td>
                </tr>
                <tr>
                  <td>Increment</td>
                  <td className="num">{table4.incrementPct.toFixed(2)}%</td>
                </tr>
                <tr className="total">
                  <td>APC-3</td>
                  <td className="num">{table4.apc3Cr.toFixed(2)} Cr.</td>
                </tr>
              </tbody>
            </table>
          </div>
        ) : (
          <p className="card-note">
            Enter the old and new remittance for at least one MF revision to
            compute APC-3 from the present remittance.
          </p>
        )}
      </section>
    </div>
  );
}
