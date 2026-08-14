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
}

function ChartTooltip({
  active,
  payload,
  label,
}: {
  active?: boolean;
  payload?: Array<{ value: number }>;
  label?: string;
}) {
  if (!active || !payload || payload.length === 0 || payload[0].value == null) {
    return null;
  }
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
      <div style={{ color: "var(--muted)" }}>{label}</div>
    </div>
  );
}

export default function CollectionChart({ rows }: Props) {
  const data: Point[] = rows.map((r) => ({
    month: monthLabel(r.month),
    value: r.normalizedDaily,
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
            dot={{
              r: 4,
              fill: "var(--series-1)",
              stroke: "var(--surface)",
              strokeWidth: 2,
            }}
            activeDot={{
              r: 5,
              fill: "var(--series-1)",
              stroke: "var(--surface)",
              strokeWidth: 2,
            }}
            isAnimationActive={false}
          />
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}
