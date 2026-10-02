"use client";

import {
  CartesianGrid,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { Table2Row } from "@/lib/apc";
import { fmtIN, monthLabel } from "@/lib/apc";

interface Props {
  rows: Table2Row[];
}

interface Point {
  month: string;
  value: number | null;
  imputed: boolean;
}

/** Solid dot for actual months, hollow dot for months filled with the average. */
function PointDot(props: {
  cx?: number;
  cy?: number;
  payload?: Point;
  r?: number;
}) {
  const { cx, cy, payload } = props;
  if (cx == null || cy == null || !payload || payload.value == null) return null;
  const r = props.r ?? 4;
  return payload.imputed ? (
    <circle cx={cx} cy={cy} r={r} fill="var(--surface)" stroke="var(--series-1)" strokeWidth={2} strokeDasharray="2 1.5" />
  ) : (
    <circle cx={cx} cy={cy} r={r} fill="var(--series-1)" stroke="var(--surface)" strokeWidth={2} />
  );
}

function ChartTooltip({
  active,
  payload,
  label,
}: {
  active?: boolean;
  payload?: Array<{ value: number; payload?: Point }>;
  label?: string;
}) {
  if (!active || !payload || payload.length === 0 || payload[0].value == null) {
    return null;
  }
  const imputed = payload[0].payload?.imputed;
  return (
    <div
      style={{
        background: "var(--surface)",
        border: "1px solid var(--border)",
        borderRadius: 8,
        padding: "8px 12px",
        boxShadow: "0 4px 12px rgba(0,0,0,0.12)",
        fontSize: 13,
      }}
    >
      <div style={{ fontWeight: 650, fontSize: 14 }}>
        ₹ {fmtIN(payload[0].value, 0)}
      </div>
      <div style={{ color: "var(--muted)" }}>
        {label}
        {imputed ? " · average of months with data" : ""}
      </div>
    </div>
  );
}

export default function CollectionChart({ rows }: Props) {
  const data: Point[] = rows.map((r) => ({
    month: monthLabel(r.month),
    value: r.normalizedDaily,
    imputed: r.imputed,
  }));

  return (
    <div style={{ width: "100%", height: 300 }}>
      <ResponsiveContainer>
        <LineChart data={data} margin={{ top: 12, right: 18, bottom: 4, left: 12 }}>
          <CartesianGrid stroke="var(--grid)" strokeWidth={1} vertical={false} />
          <XAxis
            dataKey="month"
            tick={{ fill: "var(--muted)", fontSize: 12 }}
            tickLine={false}
            axisLine={{ stroke: "var(--baseline)" }}
            interval={0}
            angle={-35}
            textAnchor="end"
            height={52}
          />
          <YAxis
            tick={{ fill: "var(--muted)", fontSize: 12 }}
            tickLine={false}
            axisLine={false}
            tickFormatter={(v: number) => fmtIN(v, 0)}
            width={82}
          />
          <Tooltip
            content={<ChartTooltip />}
            cursor={{ stroke: "var(--baseline)", strokeWidth: 1 }}
          />
          <Line
            type="linear"
            dataKey="value"
            stroke="var(--series-1)"
            strokeWidth={2}
            strokeLinejoin="round"
            strokeLinecap="round"
            connectNulls={false}
            dot={<PointDot r={4} />}
            activeDot={<PointDot r={5} />}
            isAnimationActive={false}
          />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}
