import { readFileSync } from "node:fs";
import { parseBytes } from "../server/service";
import { cellsInRows } from "../server/sheet-builder";

// Measures the daemon-side work behind the first screen: open, sheet layout, first cell chunk.
for (const file of process.argv.slice(2)) {
  const bytes = new Uint8Array(readFileSync(file));
  const t0 = performance.now();
  const wb = parseBytes(bytes, "bench", file);
  const meta = JSON.stringify(wb.meta);
  const t1 = performance.now();
  const sheet = wb.sheet(wb.meta.activeSheet);
  const layout = JSON.stringify(sheet.layout);
  const t2 = performance.now();
  const chunk = JSON.stringify(cellsInRows(sheet, 0, sheet.layout.chunkRows));
  const t3 = performance.now();
  const mid = JSON.stringify(cellsInRows(sheet, 25000, 25000 + sheet.layout.chunkRows));
  const t4 = performance.now();
  const mem = process.memoryUsage().heapUsed / 1024 / 1024;
  console.log(`${file}: ${(bytes.length / 1024 / 1024).toFixed(1)}MB cells=${sheet.layout.cellCount}`);
  console.log(`  open ${(t1 - t0).toFixed(0)}ms (meta ${meta.length}B) | sheet ${(t2 - t1).toFixed(0)}ms (layout ${layout.length}B) | chunk ${(t3 - t2).toFixed(1)}ms (${(chunk.length / 1024).toFixed(0)}KB) | mid chunk ${(t4 - t3).toFixed(1)}ms | first screen ${(t3 - t0).toFixed(0)}ms | heap ${mem.toFixed(0)}MB`);
}
