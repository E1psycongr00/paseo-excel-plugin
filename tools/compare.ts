import { readFileSync } from "node:fs";
import { parseBytes } from "../server/service";
import { cellsInRows } from "../server/sheet-builder";

// Compares parser output with the Excel oracle: displayed text, sizes, hidden rows/cols, page breaks.
const [file, oracleFile] = process.argv.slice(2) as [string, string];
const oracle = JSON.parse(readFileSync(oracleFile, "utf8"));
const wb = parseBytes(new Uint8Array(readFileSync(file)), "r", file);
let total = 0, match = 0;
const mismatches: string[] = [];
for (const [i, truth] of (oracle.sheets as any[]).entries()) {
  const sheet = wb.sheet(i);
  const L = sheet.layout;
  const byKey = new Map(cellsInRows(sheet, 0, Number.MAX_SAFE_INTEGER).map((c) => [`${c.r},${c.c}`, c]));
  for (const t of truth.cells as any[]) {
    total++;
    const got = byKey.get(`${t.r},${t.c}`)?.v ?? "";
    if (got.trim() === String(t.text).trim()) match++;
    else if (mismatches.length < 25) mismatches.push(`${L.name}!${t.r},${t.c}: excel=${JSON.stringify(t.text)} ours=${JSON.stringify(got)} f=${t.formula}`);
  }
  const colAt = (c: number) => { const col = L.cols.find((x) => c >= x.min && c <= x.max); return col ? (col.hidden ? 0 : col.width) : L.defaultColWidth; };
  const rowMap = new Map(L.rows.map((r) => [r.r, r]));
  const rowAt = (r: number) => { const row = rowMap.get(r); return row ? (row.hidden ? 0 : row.height) : L.defaultRowHeight; };
  const colDiff = (truth.colPx as number[]).map((px, c) => px - colAt(c)).filter((d) => d !== 0).length;
  const rowDiff = (truth.rowPx as number[]).map((px, r) => px - rowAt(r)).filter((d) => d !== 0).length;
  const fmt = (b: any[]) => b.map((x) => `${x.at}${x.manual ? "m" : ""}`).join(",");
  console.log(`${L.name}: colPx diff ${colDiff}/${truth.colPx.length}, rowPx diff ${rowDiff}/${truth.rowPx.length}, zoom ours=${L.print.scale} excel=${truth.zoom}`);
  console.log(`   rowBreaks excel=[${fmt(truth.rowBreaks)}] ours=[${fmt(L.print.rowBreaks)}]  colBreaks excel=[${fmt(truth.colBreaks)}] ours=[${fmt(L.print.colBreaks)}]`);
}
console.log(`text match ${match}/${total} (${((match / total) * 100).toFixed(1)}%)`);
for (const m of mismatches) console.log("  ", m);
