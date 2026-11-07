import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { parseBytes } from "../server/service";
import { cellsInRows } from "../server/sheet-builder";

// Ground truth comes from Excel itself (tools/excel-oracle.ps1). Each workbook is checked in both
// its original .xlsx form and the .xls copy Excel saved, so both parsers meet the same bar.
const root = join(__dirname, "..", "fixtures");
const cases = [
  { name: "features", files: ["features.xlsx", "oracle/features.xls"] },
  { name: "sample-career", files: ["sample-career.xlsx", "oracle/sample-career.xls"] },
];

interface OracleSheet {
  name: string;
  visible: number;
  colPx: number[];
  rowPx: number[];
  hiddenRows: number[];
  hiddenCols: number[];
  rowBreaks: { at: number; manual: boolean }[];
  colBreaks: { at: number; manual: boolean }[];
  cells: { r: number; c: number; text: string }[];
}

for (const testCase of cases) {
  const oraclePath = join(root, "oracle", `${testCase.name}.oracle.json`);
  for (const file of testCase.files) {
    const path = join(root, file);
    const available = existsSync(path) && existsSync(oraclePath);
    describe.skipIf(!available)(`${file} matches Excel`, () => {
      const oracle = available ? (JSON.parse(readFileSync(oraclePath, "utf8")) as { sheets: OracleSheet[] }) : { sheets: [] };
      const workbook = available ? parseBytes(new Uint8Array(readFileSync(path)), "test", path) : null;

      it("lists sheets with their visibility", () => {
        const states = oracle.sheets.map((sheet) => (sheet.visible === -1 ? "visible" : sheet.visible === 0 ? "hidden" : "veryHidden"));
        expect(workbook!.meta.sheets.map((sheet) => sheet.name)).toEqual(oracle.sheets.map((sheet) => sheet.name));
        expect(workbook!.meta.sheets.map((sheet) => sheet.state)).toEqual(states);
      });

      oracle.sheets.forEach((truth, index) => {
        it(`${truth.name}: displayed text`, () => {
          const sheet = workbook!.sheet(index);
          const byKey = new Map(cellsInRows(sheet, 0, Number.MAX_SAFE_INTEGER).map((cell) => [`${cell.r},${cell.c}`, cell.v]));
          const wrong = truth.cells.filter((cell) => (byKey.get(`${cell.r},${cell.c}`) ?? "").trim() !== String(cell.text).trim());
          expect(wrong.slice(0, 5)).toEqual([]);
        });

        it(`${truth.name}: column widths, hidden rows and columns`, () => {
          const { layout } = workbook!.sheet(index);
          const colAt = (c: number) => {
            const col = layout.cols.find((entry) => c >= entry.min && c <= entry.max);
            return col ? (col.hidden ? 0 : col.width) : layout.defaultColWidth;
          };
          expect(truth.colPx.map((_, c) => colAt(c))).toEqual(truth.colPx);
          const hiddenRows = layout.rows.filter((row) => row.hidden).map((row) => row.r).filter((r) => r < truth.rowPx.length);
          expect(hiddenRows).toEqual(truth.hiddenRows);
          const hiddenCols = layout.cols.filter((col) => col.hidden).flatMap((col) => Array.from({ length: col.max - col.min + 1 }, (_, i) => col.min + i)).filter((c) => c < truth.colPx.length);
          expect(hiddenCols).toEqual(truth.hiddenCols);
        });

        it(`${truth.name}: page breaks within one row or column`, () => {
          const { print } = workbook!.sheet(index).layout;
          const close = (ours: { at: number; manual: boolean }[], excel: { at: number; manual: boolean }[]) => {
            expect(ours.length).toBe(excel.length);
            ours.forEach((brk, i) => {
              expect(brk.manual).toBe(excel[i]!.manual);
              expect(Math.abs(brk.at - excel[i]!.at)).toBeLessThanOrEqual(brk.manual ? 0 : 1);
            });
          };
          if (truth.visible !== -1) return; // Excel reports no breaks for hidden sheets.
          close(print.rowBreaks, truth.rowBreaks);
          close(print.colBreaks, truth.colBreaks);
        });
      });
    });
  }
}
