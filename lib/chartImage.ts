// Renders the Table-2 line chart to a PNG (browser only, via <canvas>) so it
// can be embedded in the Word export. Returns null when no DOM is available.

import { fmtIN } from "./apc";
import type { SheetChartPoint } from "./reportModel";

export function renderChartPng(
  title: string,
  points: SheetChartPoint[],
  width = 1400,
  height = 700
): Uint8Array | null {
  if (typeof document === "undefined") return null;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = "#f4b183";
  ctx.lineWidth = 4;
  ctx.strokeRect(2, 2, width - 4, height - 4);

  ctx.fillStyle = "#c80000";
  ctx.font = "bold 34px Helvetica, Arial, sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(title, width / 2, 56);

  const values = points.map((p) => p.value).filter((v): v is number => v !== null && Number.isFinite(v));
  const maxV = values.length ? Math.max(...values) : 0;
  const raw = maxV / 5;
  const mag = raw > 0 ? Math.pow(10, Math.floor(Math.log10(raw))) : 1;
  const norm = raw / mag;
  const step = raw > 0 ? (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag : 0;
  const top = step > 0 ? Math.ceil(maxV / step) * step : 1;
  const ticks = step > 0 ? Math.round(top / step) : 1;

  const left = 150;
  const right = width - 50;
  const bottom = height - 80;
  const plotTop = 100;
  const ph = bottom - plotTop;
  const pw = right - left;

  ctx.font = "22px Helvetica, Arial, sans-serif";
  ctx.fillStyle = "#6e6e6e";
  ctx.strokeStyle = "#d9d9d9";
  ctx.lineWidth = 1.5;
  ctx.textAlign = "right";
  for (let i = 0; i <= ticks; i++) {
    const gy = bottom - (ph * i) / ticks;
    ctx.beginPath();
    ctx.moveTo(left, gy);
    ctx.lineTo(right, gy);
    ctx.stroke();
    ctx.fillText(fmtIN((top / ticks) * i, 0), left - 12, gy + 8);
  }
  ctx.strokeStyle = "#6e6e6e";
  ctx.beginPath();
  ctx.moveTo(left, bottom);
  ctx.lineTo(right, bottom);
  ctx.stroke();

  const n = points.length;
  const xAt = (i: number) => left + (pw * (i + 0.5)) / n;
  ctx.textAlign = "center";
  points.forEach((p, i) => ctx.fillText(p.label, xAt(i), bottom + 32));

  ctx.strokeStyle = "#1f4e79";
  ctx.lineWidth = 4;
  ctx.beginPath();
  let started = false;
  points.forEach((p, i) => {
    if (p.value === null || !Number.isFinite(p.value)) {
      started = false;
      return;
    }
    const cx = xAt(i);
    const cy = bottom - (ph * p.value) / top;
    if (started) ctx.lineTo(cx, cy);
    else ctx.moveTo(cx, cy);
    started = true;
  });
  ctx.stroke();
  points.forEach((p, i) => {
    if (p.value === null || !Number.isFinite(p.value)) return;
    const cx = xAt(i);
    const cy = bottom - (ph * p.value) / top;
    ctx.beginPath();
    ctx.arc(cx, cy, 8, 0, Math.PI * 2);
    ctx.fillStyle = p.imputed ? "#ffffff" : "#1f4e79";
    ctx.fill();
    ctx.lineWidth = 3;
    ctx.stroke();
  });

  const dataUrl = canvas.toDataURL("image/png");
  const b64 = dataUrl.slice(dataUrl.indexOf(",") + 1);
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}
