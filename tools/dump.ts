import { readFileSync } from "node:fs";
import { parseXlsx } from "../server/xlsx/index";

const file = process.argv[2]!;
const t0 = performance.now();
const wb = parseXlsx(new Uint8Array(readFileSync(file)), "r1");
const t1 = performance.now();
console.log("open ms", (t1 - t0).toFixed(1), "styles", wb.meta.styles.length, "sheets", wb.meta.sheets.map((s) => `${s.name}:${s.state}:${s.tabColor}`).join(", "));
const which = process.argv[3] ? [Number(process.argv[3])] : wb.meta.sheets.map((s) => s.index);
for (const i of which) {
  const t2 = performance.now();
  const s = wb.sheet(i);
  const L = s.layout;
  console.log(`\n== ${L.name} (${(performance.now() - t2).toFixed(1)}ms) cells=${L.cellCount} last=${L.lastRow},${L.lastCol} extent=${L.rowCount}x${L.colCount} defCol=${L.defaultColWidth} defRow=${L.defaultRowHeight}`);
  console.log("view", JSON.stringify(L.view));
  console.log("print", JSON.stringify(L.print));
  console.log("cols", JSON.stringify(L.cols));
  console.log("merges", L.merges.length, "drawings", JSON.stringify(L.drawings.map((d) => ({ k: d.kind, a: d.anchor, img: d.image }))));
  console.log("comments", JSON.stringify(L.comments));
  console.log("sample cells", JSON.stringify(s.cells.filter((c) => c.k !== "z").slice(0, 12)));
}
