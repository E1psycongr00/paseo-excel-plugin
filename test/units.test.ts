import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { Axis, effectiveAlign, Merges, overflowSpan, visibleWindow } from "../client/geometry";
import { applyTint, resolveColor, DEFAULT_PALETTE, DEFAULT_THEME } from "../server/colors";
import { formatColor, formatValue } from "../server/format";
import { computePages } from "../server/pages";
import { parseAreaList, parseCellRef, parseRangeRef } from "../server/refs";
import { columnWidthPx, defaultColumnPx } from "../server/units";
import { rk } from "../server/xls/reader";
import { scanXml } from "../server/xml";
import { unzipLimited } from "../server/xlsx/zip";
import type { Cell } from "../shared/model";

describe("references", () => {
  it("parses cells, ranges and defined-name areas", () => {
    expect(parseCellRef("AB12")).toEqual({ r: 11, c: 27 });
    expect(parseRangeRef("$A$1:$K$40")).toEqual([0, 0, 39, 10]);
    expect(parseRangeRef("$1:$3")).toEqual([0, 0, 2, 16383]);
    expect(parseAreaList("'등급 및 실적'!$A$1:$K$40")).toEqual([{ sheet: "등급 및 실적", range: [0, 0, 39, 10] }]);
  });
});

describe("sizes", () => {
  it("converts stored widths like Excel", () => {
    expect(columnWidthPx(6.88671875, 9)).toBe(62);
    expect(columnWidthPx(8.88671875, 9)).toBe(80);
    expect(defaultColumnPx(undefined, undefined, 7)).toBe(64);
  });
});

describe("colors and formats", () => {
  it("resolves theme tints and indexed colors", () => {
    const context = { palette: DEFAULT_PALETTE, theme: DEFAULT_THEME };
    expect(resolveColor({ theme: "0" }, context, null)).toBe("#FFFFFF");
    expect(resolveColor({ indexed: "10" }, context, null)).toBe("#FF0000");
    expect(applyTint("000000", 0.5).toUpperCase()).toBe("808080");
  });

  it("formats numbers, dates and section colors", () => {
    expect(formatValue("#,##0", 1234567.8, false)).toBe("1,234,568");
    expect(formatValue("yyyy-mm-dd", 45000, false)).toBe("2023-03-15");
    expect(formatValue("[DBNum4]General", 123, false)).toBe("일이삼");
    expect(formatColor("#,##0;[Red]-#,##0", -5)).toBe("#FF0000");
    expect(formatColor("#,##0;[Red]-#,##0", 5)).toBeUndefined();
  });

  it("decodes RK numbers", () => {
    expect(rk((100 << 2) | 0x02)).toBe(100);
    expect(rk((12345 << 2) | 0x03)).toBe(123.45);
  });
});

describe("xml scanner", () => {
  it("reads attributes with entities and strips main-namespace prefixes", () => {
    const seen: string[] = [];
    scanXml(`<x:c r="A1" t="s"><x:v>1&amp;2</x:v></x:c><x14:cfRule a='&lt;'/>`, {
      open: (name, attrs) => seen.push(`${name}:${JSON.stringify(attrs)}`),
      text: (text) => seen.push(text),
    });
    expect(seen).toEqual([`c:{"r":"A1","t":"s"}`, `v:{}`, "1&2", `x14:cfRule:{"a":"<"}`]);
  });
});

describe("page breaks", () => {
  it("breaks columns that overflow the printable width", () => {
    const print = computePages({
      colPx: () => 100,
      rowPx: () => 20,
      usedRange: [0, 0, 9, 19],
      setup: { paperSize: 9, manualRowBreaks: [5], manualColBreaks: [], printArea: null, titleRows: null, titleCols: null },
    });
    expect(print.rowBreaks[0]).toEqual({ at: 5, manual: true });
    expect(print.colBreaks.length).toBeGreaterThan(1);
  });
});

describe("grid geometry", () => {
  const axis = new Axis(10, (i) => (i === 3 ? 0 : 10), (i) => i === 3, 1);
  it("skips hidden entries", () => {
    expect(axis.start(4)).toBe(30);
    expect(axis.indexAt(35)).toBe(4);
    expect(axis.total).toBe(90);
  });

  it("windows in blocks", () => {
    const win = visibleWindow(axis, axis, 0, 0, 30, 30);
    expect(win.r0).toBe(0);
    expect(win.r1).toBeGreaterThanOrEqual(3);
  });

  it("lets left-aligned text spill into empty neighbours", () => {
    const cell: Cell = { r: 0, c: 0, s: 0, v: "long text", k: "s" };
    const row = new Map<number, Cell>([[0, cell], [3, { r: 0, c: 3, s: 0, v: "x", k: "s" }]]);
    expect(overflowSpan(cell, undefined, row, new Merges([]), { c0: 0, c1: 9 })).toEqual([0, 2]);
    expect(effectiveAlign(undefined, { ...cell, k: "n" })).toBe("right");
  });
});

describe("unzip limits", () => {
  const zip = zipSync({ "a.xml": new Uint8Array(4000), "b.xml": new Uint8Array(4000), "xl/media/big.png": new Uint8Array(9000) });
  const all = () => true;

  it("keeps entries within the limits and skips filtered ones", () => {
    const files = unzipLimited(zip, (name) => !name.startsWith("xl/media/"), { entry: 5000, total: 8000 });
    expect(Object.keys(files).sort()).toEqual(["a.xml", "b.xml"]);
    expect(files["a.xml"]!.length).toBe(4000);
  });

  it("rejects a single oversized entry and an oversized total", () => {
    expect(() => unzipLimited(zip, all, { entry: 3000, total: 1e6 })).toThrow(/a\.xml/);
    expect(() => unzipLimited(zip, (name) => name.endsWith(".xml"), { entry: 5000, total: 7000 })).toThrow(/전체 크기/);
  });

  it("rejects a header that claims a huge size before allocating it", () => {
    const lying = zipSync({ "sheet.xml": strToU8("<x/>") });
    const view = new DataView(lying.buffer, lying.byteOffset, lying.byteLength);
    for (let i = lying.length - 22; i >= 0; i--) {
      if (view.getUint32(i, true) === 0x02014b50) {
        view.setUint32(i + 24, 0xf0000000, true); // central directory: uncompressed size
        break;
      }
    }
    expect(() => unzipLimited(lying, all)).toThrow(/너무 큰 항목/);
  });
});
