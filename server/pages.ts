import type { PageBreak, PrintInfo, Range } from "../shared/model";

export interface PageSetup {
  paperSize?: number;
  orientation?: "portrait" | "landscape";
  /** Percent. */
  scale?: number;
  fitToPage?: boolean;
  fitToWidth?: number;
  fitToHeight?: number;
  /** Inches. */
  margins?: { left: number; right: number; top: number; bottom: number };
  order?: "down" | "over";
  manualRowBreaks: number[];
  manualColBreaks: number[];
  printArea: Range | null;
  titleRows: [number, number] | null;
  titleCols: [number, number] | null;
}

// Paper sizes in inches, by OOXML/BIFF paper code.
const PAPER: Record<number, [number, number]> = {
  1: [8.5, 11], 2: [8.5, 11], 3: [11, 17], 4: [17, 11], 5: [8.5, 14], 6: [5.5, 8.5], 7: [7.25, 10.5],
  8: [11.69, 16.54], 9: [8.27, 11.69], 10: [8.27, 11.69], 11: [5.83, 8.27], 12: [10.12, 14.33],
  13: [7.17, 10.12], 14: [8.5, 13], 15: [8.46, 10.83], 16: [10, 14], 17: [11, 17], 18: [8.5, 11],
  19: [3.88, 8.88], 20: [4.13, 9.5], 21: [4.5, 10.38], 22: [4.75, 11], 23: [5, 11.5],
  24: [17, 22], 25: [22, 34], 26: [34, 44], 27: [4.33, 8.66], 28: [6.38, 9.02], 29: [12.76, 18.03],
  30: [9.02, 12.76], 31: [4.49, 6.38], 32: [4.49, 9.02], 33: [9.84, 13.9], 34: [6.93, 9.84],
  35: [6.93, 3.94], 36: [4.33, 9.06], 37: [3.88, 7.5], 38: [3.63, 6.5], 39: [14.88, 11],
  40: [8.5, 12], 41: [8.5, 13], 66: [16.54, 23.39], 70: [4.13, 5.83], 256: [8.27, 11.69],
};

// Calibrated against Excel 16: printed columns come out narrower, printed rows slightly taller.
const PRINT_COL_FACTOR = 0.96;
const PRINT_ROW_FACTOR = 1.009;

const DEFAULT_MARGINS = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 };

export interface PageInput {
  /** Pixel sizes of every laid-out column and row (0 when hidden). */
  colPx: (c: number) => number;
  rowPx: (r: number) => number;
  usedRange: Range | null;
  setup: PageSetup;
}

export function computePages({ colPx, rowPx, usedRange, setup }: PageInput): PrintInfo {
  const range = setup.printArea ?? usedRange;
  const order = setup.order ?? "down";
  if (!range) {
    return { area: setup.printArea, pageRange: null, rowBreaks: [], colBreaks: [], order, scale: setup.scale ?? 100, estimated: false };
  }
  const [r1, c1, r2, c2] = range;
  let [paperW, paperH] = PAPER[setup.paperSize ?? 9] ?? PAPER[9]!;
  if (setup.orientation === "landscape") [paperW, paperH] = [paperH, paperW];
  const margins = setup.margins ?? DEFAULT_MARGINS;
  const printableW = Math.max(36, (paperW - margins.left - margins.right) * 72);
  const printableH = Math.max(36, (paperH - margins.top - margins.bottom) * 72);
  // Printed cells come out slightly narrower than on screen (printer font metrics). Calibrated
  // against Excel 16 page breaks; see tools/compare.ts.
  const colPt = (px: number) => px * 0.75 * PRINT_COL_FACTOR;
  const rowPt = (px: number) => px * 0.75 * PRINT_ROW_FACTOR;

  let totalW = 0;
  for (let c = c1; c <= c2; c++) totalW += colPt(colPx(c));
  let totalH = 0;
  for (let r = r1; r <= r2; r++) totalH += rowPt(rowPx(r));

  let scale = (setup.scale ?? 100) / 100;
  if (setup.fitToPage) {
    const fitW = setup.fitToWidth ?? 1;
    const fitH = setup.fitToHeight ?? 1;
    const byW = fitW > 0 && totalW > 0 ? (printableW * fitW) / totalW : Infinity;
    const byH = fitH > 0 && totalH > 0 ? (printableH * fitH) / totalH : Infinity;
    scale = Math.min(1, byW, byH);
    scale = Math.max(0.1, Math.floor(scale * 100) / 100);
  }

  const titleH = setup.titleRows ? sumRange(setup.titleRows, (r) => rowPt(rowPx(r))) * scale : 0;
  const titleW = setup.titleCols ? sumRange(setup.titleCols, (c) => colPt(colPx(c))) * scale : 0;
  const manualRows = new Set(setup.manualRowBreaks);
  const manualCols = new Set(setup.manualColBreaks);

  const colBreaks = breaks(c1, c2, (c) => colPt(colPx(c)) * scale, printableW, titleW, setup.titleCols, manualCols);
  const rowBreaks = breaks(r1, r2, (r) => rowPt(rowPx(r)) * scale, printableH, titleH, setup.titleRows, manualRows);
  return {
    area: setup.printArea,
    pageRange: range,
    rowBreaks,
    colBreaks,
    order,
    scale: Math.round(scale * 100),
    estimated: rowBreaks.some((brk) => !brk.manual) || colBreaks.some((brk) => !brk.manual),
  };
}

function sumRange([a, b]: [number, number], size: (i: number) => number): number {
  let total = 0;
  for (let i = a; i <= b; i++) total += size(i);
  return total;
}

function breaks(
  start: number,
  end: number,
  size: (i: number) => number,
  page: number,
  titles: number,
  titleRange: [number, number] | null,
  manual: Set<number>,
): PageBreak[] {
  const out: PageBreak[] = [];
  let used = 0;
  let first = true;
  for (let i = start; i <= end; i++) {
    const length = size(i);
    if (i > start && manual.has(i)) {
      out.push({ at: i, manual: true });
      used = 0;
      first = false;
    }
    // Titles repeat on every page that does not already contain them.
    const available = first || !titleRange || (i >= titleRange[0] && i <= titleRange[1]) ? page : page - titles;
    if (used > 0 && used + length > available + 0.01) {
      out.push({ at: i, manual: false });
      used = 0;
      first = false;
    }
    used += length;
  }
  return out;
}
