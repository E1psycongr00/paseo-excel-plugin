import { describe, expect, it } from "vitest";
import { buildAxes, hiddenRuns, openedMask } from "../client/geometry";
import type { SheetLayout } from "../shared/model";

function layout(partial: Partial<SheetLayout>): SheetLayout {
  return {
    index: 0,
    name: "Sheet1",
    view: { mode: "normal", showGrid: true, showHeaders: true, zoom: 100, frozenRows: 0, frozenCols: 0, topRow: 0, leftCol: 0, rtl: false },
    defaultColWidth: 64,
    defaultRowHeight: 20,
    rowCount: 10,
    colCount: 6,
    lastRow: 9,
    lastCol: 5,
    cols: [],
    rows: [],
    merges: [],
    drawings: [],
    comments: [],
    links: [],
    print: { area: null, pageRange: null, rowBreaks: [], colBreaks: [], order: "down", scale: 100, estimated: false },
    cellCount: 0,
    chunkRows: 400,
    ...partial,
  };
}

describe("hidden runs", () => {
  const sheet = layout({
    cols: [{ min: 1, max: 2, width: 50, hidden: true, style: null }, { min: 4, max: 4, width: 0, hidden: true, style: null }],
    rows: [3, 4, 5].map((r) => ({ r, height: 30, hidden: true, style: null })),
  });

  it("groups consecutive hidden entries by their first index", () => {
    const axes = buildAxes(sheet, 1, false);
    expect([...hiddenRuns(axes.cols.hidden)]).toEqual([[1, 2], [4, 4]]);
    expect([...hiddenRuns(axes.rows.hidden)]).toEqual([[3, 5]]);
  });

  it("lays out only the opened runs", () => {
    const base = buildAxes(sheet, 1, false);
    const colMask = openedMask(base.cols.count, hiddenRuns(base.cols.hidden), new Set([1]));
    const axes = buildAxes(sheet, 1, { row: () => false, col: (c) => colMask[c] === 1 });
    expect([0, 1, 2, 3, 4, 5].map((c) => axes.cols.size(c))).toEqual([64, 50, 50, 64, 0, 64]);
    expect(axes.rows.size(4)).toBe(0);
    // Still flagged hidden, so the opened run is shaded.
    expect(axes.cols.hidden[1]).toBe(1);
  });

  it("gives every row one height in quick review, hidden rows still folded", () => {
    const axes = buildAxes(layout({ rows: [{ r: 0, height: 60, hidden: false, style: null }, { r: 1, height: 30, hidden: true, style: null }] }), 1, false, 22);
    expect([axes.rows.size(0), axes.rows.size(1), axes.rows.size(2)]).toEqual([22, 0, 22]);
  });
});
