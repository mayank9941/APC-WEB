import { NextResponse } from "next/server";
import { sql } from "@/lib/db";

export const revalidate = 3600;

export async function GET() {
  const rows = await sql`
    SELECT p.plaza_code AS code, p.plaza_name AS name,
           EXISTS (SELECT 1 FROM plaza_monthly_data d WHERE d.plaza_code = p.plaza_code) AS has_data
    FROM pb_plazas p
    ORDER BY lower(p.plaza_name), p.plaza_code
  `;
  return NextResponse.json(rows);
}
