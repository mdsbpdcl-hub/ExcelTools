// Charts for the optional summary sheet. The Excel library used here cannot create native Excel charts,
// so charts are drawn to a picture (PNG) and placed in the workbook. This is stated in the app and run summary.

import type { CellValue, Table } from './types';
import type { ChartSpec } from './workflow';
import { cellToString, compareValues, toNumber } from './util';
import { colIndex } from './ops/helpers';
import { Acc, AGG_LABELS } from './ops/pivot';

export interface ChartData {
  title: string;
  type: ChartSpec['type'];
  labels: string[];
  values: number[];
  points?: { x: number; y: number }[];
  note?: string;
}

export function prepareChart(t: Table, spec: ChartSpec): ChartData | string {
  const ci = colIndex(t, spec.category);
  const vi = spec.agg === 'count' && !spec.value ? -1 : colIndex(t, spec.value);
  if (ci < 0) return `Chart skipped: column "${spec.category}" not found in ${t.sheetName}.`;
  if (vi < 0 && !(spec.agg === 'count' && !spec.value)) return `Chart skipped: column "${spec.value}" not found in ${t.sheetName}.`;
  const title = spec.title || (spec.type === 'scatter' ? `${spec.value} vs ${spec.category}` : `${AGG_LABELS[spec.agg]} of ${spec.value || 'rows'} by ${spec.category}`);
  if (spec.type === 'scatter') {
    const points: { x: number; y: number }[] = [];
    for (const row of t.rows) {
      const x = toNumber(row[ci]), y = toNumber(row[vi]);
      if (x !== null && y !== null) points.push({ x, y });
      if (points.length >= 5000) break;
    }
    if (!points.length) return `Chart skipped: "${spec.category}" and "${spec.value}" have no numeric pairs.`;
    return { title, type: 'scatter', labels: [], values: [], points, note: points.length >= 5000 ? 'First 5,000 points shown.' : undefined };
  }
  const groups = new Map<string, { label: CellValue; acc: Acc }>();
  for (const row of t.rows) {
    const k = cellToString(row[ci]).trim() || '(blank)';
    if (!groups.has(k)) groups.set(k, { label: row[ci], acc: new Acc() });
    groups.get(k)!.acc.add(vi < 0 ? 1 : row[vi], spec.agg === 'median');
  }
  let entries = [...groups.entries()].map(([k, g]) => ({ k, label: g.label, v: Number(g.acc.result(spec.agg === 'count' && vi < 0 ? 'count' : spec.agg) ?? 0) }));
  if (spec.type === 'line' || spec.type === 'waterfall') entries.sort((a, b) => compareValues(a.label, b.label));
  else entries.sort((a, b) => b.v - a.v);
  let note: string | undefined;
  const max = Math.max(2, spec.maxCategories || 12);
  if (entries.length > max) {
    if (spec.type === 'line' || spec.type === 'waterfall') { note = `First ${max} of ${entries.length} categories shown.`; entries = entries.slice(0, max); }
    else {
      const rest = entries.slice(max - 1).reduce((s, e) => s + e.v, 0);
      note = `${entries.length - max + 1} smaller categories grouped as "Others".`;
      entries = [...entries.slice(0, max - 1), { k: 'Others', label: 'Others', v: rest }];
    }
  }
  return { title, type: spec.type, labels: entries.map((e) => e.k), values: entries.map((e) => e.v), note };
}

const PALETTE = ['#6aa882', '#4f8bc9', '#e0a03a', '#c8655a', '#8a7fc2', '#5bb3b0', '#b9a24b', '#9aa3ad', '#d17fa8', '#7c9c4f', '#4d6b8a', '#c98a5b'];
const fmt = (n: number) => (Math.abs(n) >= 1e7 ? `${(n / 1e7).toFixed(1)} Cr` : Math.abs(n) >= 1e5 ? `${(n / 1e5).toFixed(1)} L` : Math.abs(n) >= 1000 ? n.toLocaleString('en-IN', { maximumFractionDigits: 0 }) : String(Math.round(n * 100) / 100));

type Ctx = OffscreenCanvasRenderingContext2D;

/** Draws a chart to PNG. Returns null where the browser has no OffscreenCanvas (the run summary says so). */
export async function renderChartPng(d: ChartData, width = 900, height = 480): Promise<Uint8Array | null> {
  if (typeof OffscreenCanvas === 'undefined') return null;
  const canvas = new OffscreenCanvas(width, height);
  const ctx = canvas.getContext('2d') as Ctx | null;
  if (!ctx) return null;
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);
  ctx.fillStyle = '#1f2937';
  ctx.font = '600 18px Segoe UI, Roboto, sans-serif';
  ctx.fillText(d.title.slice(0, 80), 20, 30);
  if (d.note) { ctx.font = '12px Segoe UI, Roboto, sans-serif'; ctx.fillStyle = '#6b7280'; ctx.fillText(d.note, 20, height - 10); }
  const area = { x: 80, y: 50, w: width - 110, h: height - 130 };
  if (d.type === 'pie') drawPie(ctx, d, width, height);
  else if (d.type === 'scatter') drawScatter(ctx, d, area);
  else drawBars(ctx, d, area);
  const blob = await canvas.convertToBlob({ type: 'image/png' });
  return new Uint8Array(await blob.arrayBuffer());
}

function axes(ctx: Ctx, area: { x: number; y: number; w: number; h: number }, min: number, max: number) {
  ctx.strokeStyle = '#e5e7eb';
  ctx.fillStyle = '#6b7280';
  ctx.font = '11px Segoe UI, Roboto, sans-serif';
  ctx.textAlign = 'right';
  for (let i = 0; i <= 5; i++) {
    const v = min + ((max - min) * i) / 5;
    const y = area.y + area.h - (area.h * i) / 5;
    ctx.beginPath(); ctx.moveTo(area.x, y); ctx.lineTo(area.x + area.w, y); ctx.stroke();
    ctx.fillText(fmt(v), area.x - 6, y + 4);
  }
  ctx.textAlign = 'left';
}

function drawBars(ctx: Ctx, d: ChartData, area: { x: number; y: number; w: number; h: number }) {
  const n = d.values.length || 1;
  let bars: { start: number; end: number; color: string }[];
  if (d.type === 'waterfall') {
    let run = 0;
    bars = d.values.map((v) => { const b = { start: run, end: run + v, color: v >= 0 ? '#6aa882' : '#c8655a' }; run += v; return b; });
    bars.push({ start: 0, end: run, color: '#4f8bc9' });
  } else bars = d.values.map((v) => ({ start: 0, end: v, color: '#6aa882' }));
  const labels = d.type === 'waterfall' ? [...d.labels, 'Total'] : d.labels;
  const all = bars.flatMap((b) => [b.start, b.end]);
  const min = Math.min(0, ...all), max = Math.max(0, ...all) || 1;
  axes(ctx, area, min, max);
  const y = (v: number) => area.y + area.h - ((v - min) / (max - min)) * area.h;
  const slot = area.w / (d.type === 'waterfall' ? n + 1 : n);
  if (d.type === 'line') {
    ctx.strokeStyle = '#4f8bc9'; ctx.lineWidth = 2; ctx.beginPath();
    d.values.forEach((v, i) => { const px = area.x + slot * (i + 0.5); i ? ctx.lineTo(px, y(v)) : ctx.moveTo(px, y(v)); });
    ctx.stroke();
    ctx.fillStyle = '#4f8bc9';
    d.values.forEach((v, i) => { ctx.beginPath(); ctx.arc(area.x + slot * (i + 0.5), y(v), 3, 0, Math.PI * 2); ctx.fill(); });
  } else {
    bars.forEach((b, i) => {
      ctx.fillStyle = b.color;
      const top = y(Math.max(b.start, b.end)), bottom = y(Math.min(b.start, b.end));
      ctx.fillRect(area.x + slot * i + slot * 0.15, top, slot * 0.7, Math.max(1, bottom - top));
    });
  }
  ctx.fillStyle = '#374151';
  ctx.font = '11px Segoe UI, Roboto, sans-serif';
  labels.forEach((l, i) => {
    ctx.save();
    ctx.translate(area.x + slot * (i + 0.5), area.y + area.h + 8);
    ctx.rotate(labels.length > 8 ? -Math.PI / 5 : 0);
    ctx.textAlign = labels.length > 8 ? 'right' : 'center';
    ctx.fillText(l.slice(0, 18), 0, 10);
    ctx.restore();
  });
}

function drawScatter(ctx: Ctx, d: ChartData, area: { x: number; y: number; w: number; h: number }) {
  const pts = d.points ?? [];
  const xs = pts.map((p) => p.x), ys = pts.map((p) => p.y);
  const minX = Math.min(...xs), maxX = Math.max(...xs) || minX + 1, minY = Math.min(0, ...ys), maxY = Math.max(...ys) || 1;
  axes(ctx, area, minY, maxY);
  ctx.fillStyle = 'rgba(79,139,201,0.6)';
  for (const p of pts) {
    const px = area.x + ((p.x - minX) / (maxX - minX || 1)) * area.w;
    const py = area.y + area.h - ((p.y - minY) / (maxY - minY || 1)) * area.h;
    ctx.beginPath(); ctx.arc(px, py, 3, 0, Math.PI * 2); ctx.fill();
  }
  ctx.fillStyle = '#6b7280'; ctx.font = '11px Segoe UI, Roboto, sans-serif';
  ctx.fillText(fmt(minX), area.x, area.y + area.h + 18);
  ctx.textAlign = 'right'; ctx.fillText(fmt(maxX), area.x + area.w, area.y + area.h + 18); ctx.textAlign = 'left';
}

function drawPie(ctx: Ctx, d: ChartData, width: number, height: number) {
  const total = d.values.reduce((s, v) => s + Math.max(0, v), 0) || 1;
  const cx = width * 0.33, cy = height / 2 + 10, r = Math.min(width, height) * 0.33;
  let a = -Math.PI / 2;
  d.values.forEach((v, i) => {
    const slice = (Math.max(0, v) / total) * Math.PI * 2;
    ctx.fillStyle = PALETTE[i % PALETTE.length];
    ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, r, a, a + slice); ctx.closePath(); ctx.fill();
    a += slice;
  });
  ctx.font = '12px Segoe UI, Roboto, sans-serif';
  d.labels.forEach((l, i) => {
    const y = 70 + i * 22;
    ctx.fillStyle = PALETTE[i % PALETTE.length]; ctx.fillRect(width * 0.66, y - 10, 12, 12);
    ctx.fillStyle = '#374151';
    ctx.fillText(`${l.slice(0, 22)} — ${fmt(d.values[i])} (${((Math.max(0, d.values[i]) / total) * 100).toFixed(1)}%)`, width * 0.66 + 18, y);
  });
}
