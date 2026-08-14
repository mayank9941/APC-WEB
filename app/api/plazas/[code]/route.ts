import { NextResponse } from "next/server";
import { sql } from "@/lib/db";
import { CATEGORIES } from "@/lib/apc";

const CAR = CATEGORIES[0]; // "Car/Jeep/Van/Light Motor Vehicle"

export async function GET(
  _req: Request,
  { params }: { params: Promise<{ code: string }> }
) {
  const { code } = await params;
  const plazaCode = Number(code);
  if (!Number.isInteger(plazaCode)) {
    return NextResponse.json({ error: "Invalid plaza code" }, { status: 400 });
  }

  const plazaRows = await sql`
    SELECT plaza_code AS code, plaza_name AS name
    FROM pb_plazas WHERE plaza_code = ${plazaCode}
  `;
  if (plazaRows.length === 0) {
    return NextResponse.json({ error: "Plaza not found" }, { status: 404 });
  }

  // AP compensation comes from NHAI's real report (annual_pass_monthly),
  // attributed to the Car category; plaza-months absent from the report are 0.
  // The estimated plaza_monthly_data.ap_compensation column is no longer used.
  const rows = await sql`
    SELECT to_char(d.month, 'YYYY-MM') AS month, d.category,
           d.etc_cnt::float8       AS etc_cnt,
           d.cash_cnt::float8      AS cash_cnt,
           d.upi_cnt::float8       AS upi_cnt,
           d.exempt_cnt::float8    AS exempt_cnt,
           d.tmcc_total::float8    AS tmcc_total,
           d.etc_collection::float8 AS etc_collection,
           d.ap_cnt::float8        AS ap_cnt,
           COALESCE(ap.compensation, 0)::float8 AS ap_compensation
    FROM plaza_monthly_data d
    LEFT JOIN annual_pass_monthly ap
      ON ap.plaza_code = d.plaza_code AND ap.month = d.month
     AND d.category = ${CAR}
    WHERE d.plaza_code = ${plazaCode}
    ORDER BY d.month, d.category
  `;

  const months = [...new Set(rows.map((r) => r.month as string))].sort();
  return NextResponse.json({ plaza: plazaRows[0], months, rows });
}
