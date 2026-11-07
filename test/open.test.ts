import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { initialSheet, type OpenResult } from "../shared/model";
import { openWorkbook, parseBytes, sheetBundle } from "../server/service";

const fixture = resolve(__dirname, "../fixtures/features.xlsx");

describe("open bundle", () => {
  it("carries the first sheet's layout and first-screen cells", async () => {
    const result = JSON.parse((await openWorkbook(fixture)).json) as OpenResult;
    expect(result.initial?.layout.index).toBe(initialSheet(result.meta, undefined));
    expect(result.initial?.chunks.map((chunk) => chunk.rowStart)).toEqual([0]);
    expect(result.initial?.chunks[0]!.cells.length).toBeGreaterThan(0);
  });

  it("bundles the requested sheet", async () => {
    const result = JSON.parse((await openWorkbook(fixture, 1)).json) as OpenResult;
    expect(result.initial?.layout.index).toBe(1);
  });

  it("covers frozen rows and the saved top rows on the app's chunk grid", () => {
    const parsed = parseBytes(new Uint8Array(readFileSync(fixture)), "t", fixture).sheet(0);
    const { layout } = parsed;
    const moved = { ...parsed, layout: { ...layout, lastRow: 1000, rowCount: 1100, view: { ...layout.view, topRow: 450, frozenRows: 1 } } };
    const size = layout.chunkRows;
    expect(sheetBundle(moved).chunks.map((chunk) => chunk.rowStart)).toEqual([0, Math.floor(450 / size) * size]);
    const boundary = { ...moved, layout: { ...moved.layout, view: { ...moved.layout.view, topRow: size - 10, frozenRows: 0 } } };
    expect(sheetBundle(boundary).chunks.map((chunk) => chunk.rowStart)).toEqual([0, size]);
  });

  it("reports unparsed sheet size for prefetch, and 0 once parsed", () => {
    const workbook = parseBytes(new Uint8Array(readFileSync(fixture)), "t", fixture);
    expect(workbook.sheetBytes(0)).toBeGreaterThan(0);
    workbook.sheet(0);
    expect(workbook.sheetBytes(0)).toBe(0);
  });
});
