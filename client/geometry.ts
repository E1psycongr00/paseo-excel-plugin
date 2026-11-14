import type { Cell, CellStyle, ColRange, Range, RowInfo, SheetLayout } from "../shared/model";

/** Positions along one axis. Hidden entries collapse to zero unless shown. */
export class Axis {
  readonly count: number;
  private readonly starts: Float64Array;
  readonly hidden: Uint8Array;

  constructor(count: number, size: (index: number) => number, isHidden: (index: number) => boolean, scale: number) {
    this.count = count;
    this.starts = new Float64Array(count + 1);
    this.hidden = new Uint8Array(count);
    let at = 0;
    for (let i = 0; i < count; i++) {
      this.starts[i] = at;
      if (isHidden(i)) this.hidden[i] = 1;
      at += size(i) * scale;
    }
    this.starts[count] = at;
  }

  start(index: number): number {
    return this.starts[Math.max(0, Math.min(this.count, index))]!;
  }

  end(index: number): number {
    return this.start(index + 1);
  }

  size(index: number): number {
    return this.start(index + 1) - this.start(index);
  }

  get total(): number {
    return this.starts[this.count]!;
  }

  /** The index whose span contains `offset` (the last one at or past the end). */
  indexAt(offset: number): number {
    let lo = 0;
    let hi = this.count - 1;
    if (offset <= 0) return 0;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.starts[mid]! <= offset) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}

/** Which hidden rows and columns to lay out (shaded) instead of collapsing. */
export interface Reveal {
  row(r: number): boolean;
  col(c: number): boolean;
}

const REVEAL_ALL: Reveal = { row: () => true, col: () => true };
const REVEAL_NONE: Reveal = { row: () => false, col: () => false };

/** `rowHeight` lays every row out at that height (hidden rows still collapse) instead of the sheet's own. */
export function buildAxes(layout: SheetLayout, scale: number, reveal: boolean | Reveal, rowHeight?: number) {
  const show = reveal === true ? REVEAL_ALL : reveal === false ? REVEAL_NONE : reveal;
  const colRanges = layout.cols;
  const findCol = (c: number): ColRange | undefined => {
    let lo = 0;
    let hi = colRanges.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const col = colRanges[mid]!;
      if (c < col.min) hi = mid - 1;
      else if (c > col.max) lo = mid + 1;
      else return col;
    }
    return undefined;
  };
  const rowMap = new Map<number, RowInfo>(layout.rows.map((row) => [row.r, row]));
  const zeroRows = layout.defaultRowHeight === 0;
  const fallbackRow = zeroRows ? 20 : layout.defaultRowHeight;
  const cols = new Axis(
    layout.colCount,
    (c) => {
      const col = findCol(c);
      if (!col) return layout.defaultColWidth;
      if (col.hidden) return show.col(c) ? col.width || layout.defaultColWidth : 0;
      return col.width;
    },
    (c) => findCol(c)?.hidden ?? false,
    scale,
  );
  const rows = new Axis(
    layout.rowCount,
    (r) => {
      const row = rowMap.get(r);
      if (!row) return zeroRows ? (show.row(r) ? rowHeight ?? fallbackRow : 0) : rowHeight ?? layout.defaultRowHeight;
      if (row.hidden) return show.row(r) ? rowHeight ?? (row.height || fallbackRow) : 0;
      return rowHeight ?? row.height;
    },
    (r) => {
      const row = rowMap.get(r);
      return row ? row.hidden : zeroRows;
    },
    scale,
  );
  return { cols, rows, findCol, rowMap };
}

/**
 * Runs of consecutive hidden entries, keyed by their first index, with the run's last index. A run
 * is what one tap on its marker opens or closes, like a block hidden together in Excel.
 */
export function hiddenRuns(hidden: Uint8Array): Map<number, number> {
  const runs = new Map<number, number>();
  for (let i = 0; i < hidden.length; i++) {
    if (!hidden[i]) continue;
    const start = i;
    while (i + 1 < hidden.length && hidden[i + 1]) i++;
    runs.set(start, i);
  }
  return runs;
}

/** Marks every index of the opened runs. */
export function openedMask(count: number, runs: Map<number, number>, opened: ReadonlySet<number>): Uint8Array {
  const mask = new Uint8Array(count);
  for (const start of opened) {
    const end = runs.get(start);
    if (end !== undefined) mask.fill(1, start, end + 1);
  }
  return mask;
}

export interface Window {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

export const ROW_BLOCK = 16;
export const COL_BLOCK = 6;
// Cells kept mounted beyond the viewport, so a fling reaches drawn tiles before the JS thread catches up.
const OVERSCAN_ROWS = 12;
const OVERSCAN_COLS = 3;

/** Visible cells plus overscan, snapped to blocks so scrolling re-renders rarely. */
export function visibleWindow(rows: Axis, cols: Axis, x: number, y: number, width: number, height: number): Window {
  const r0 = Math.floor(Math.max(0, rows.indexAt(y) - OVERSCAN_ROWS) / ROW_BLOCK) * ROW_BLOCK;
  const r1 = Math.min(rows.count - 1, (Math.floor((rows.indexAt(y + height) + OVERSCAN_ROWS) / ROW_BLOCK) + 1) * ROW_BLOCK - 1);
  const c0 = Math.floor(Math.max(0, cols.indexAt(x) - OVERSCAN_COLS) / COL_BLOCK) * COL_BLOCK;
  const c1 = Math.min(cols.count - 1, (Math.floor((cols.indexAt(x + width) + OVERSCAN_COLS) / COL_BLOCK) + 1) * COL_BLOCK - 1);
  return { r0, r1, c0, c1 };
}

export const sameWindow = (a: Window, b: Window) => a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1;

/** A block of the window that draws and caches on its own. */
export interface Tile extends Window {
  /** First block row or column of the window: draws merges anchored above or left of it. */
  top: boolean;
  left: boolean;
}

/** Splits a window into grid-aligned blocks; a scroll then only draws the blocks that came into view. */
export function windowTiles(win: Window): Tile[] {
  const out: Tile[] = [];
  if (win.r0 > win.r1 || win.c0 > win.c1) return out;
  for (let r = win.r0; r <= win.r1; r = (Math.floor(r / ROW_BLOCK) + 1) * ROW_BLOCK) {
    const r1 = Math.min(win.r1, (Math.floor(r / ROW_BLOCK) + 1) * ROW_BLOCK - 1);
    for (let c = win.c0; c <= win.c1; c = (Math.floor(c / COL_BLOCK) + 1) * COL_BLOCK) {
      const c1 = Math.min(win.c1, (Math.floor(c / COL_BLOCK) + 1) * COL_BLOCK - 1);
      out.push({ r0: r, r1, c0: c, c1, top: r === win.r0, left: c === win.c0 });
    }
  }
  return out;
}

/**
 * Exactly one tile of a window draws each merge: the one holding its top-left cell, or, when that
 * cell lies above or left of the window, the edge tile where the merge enters the window.
 */
export function ownsMerge(merge: Range, tile: Tile): boolean {
  if (merge[2] < tile.r0 || merge[3] < tile.c0) return false;
  const r = merge[0] < tile.r0 && tile.top ? tile.r0 : merge[0];
  const c = merge[1] < tile.c0 && tile.left ? tile.c0 : merge[1];
  return r >= tile.r0 && r <= tile.r1 && c >= tile.c0 && c <= tile.c1;
}

/** Merge lookup: anchor cell and covered cells. */
export class Merges {
  private readonly byCell = new Map<number, Range>();
  // Huge merges (whole rows or columns) are checked directly instead of indexed per cell.
  private readonly large: Range[] = [];
  readonly list: Range[];
  constructor(merges: Range[]) {
    this.list = merges;
    for (const merge of merges) {
      const area = (merge[2] - merge[0] + 1) * (merge[3] - merge[1] + 1);
      if (area > 4096) {
        this.large.push(merge);
        continue;
      }
      for (let r = merge[0]; r <= merge[2]; r++) {
        for (let c = merge[1]; c <= merge[3]; c++) this.byCell.set(r * 16384 + c, merge);
      }
    }
  }

  at(r: number, c: number): Range | undefined {
    const hit = this.byCell.get(r * 16384 + c);
    if (hit || this.large.length === 0) return hit;
    return this.large.find((m) => r >= m[0] && r <= m[2] && c >= m[1] && c <= m[3]);
  }

  intersecting(win: Window): Range[] {
    return this.list.filter((m) => m[2] >= win.r0 && m[0] <= win.r1 && m[3] >= win.c0 && m[1] <= win.c1);
  }

  /** Merges drawn by this tile; see ownsMerge. */
  ownedBy(tile: Tile): Range[] {
    return this.list.filter((m) => m[0] <= tile.r1 && m[1] <= tile.c1 && ownsMerge(m, tile));
  }
}

/** Cells grouped by row for quick access while rendering. */
export type RowCells = Map<number, Map<number, Cell>>;

export function indexCells(cells: Cell[], into: RowCells = new Map()): RowCells {
  for (const cell of cells) {
    let row = into.get(cell.r);
    if (!row) {
      row = new Map();
      into.set(cell.r, row);
    }
    row.set(cell.c, cell);
  }
  return into;
}

/** Effective horizontal alignment once "general" resolves by value kind. */
export function effectiveAlign(style: CellStyle | undefined, cell: Cell | undefined): "left" | "center" | "right" {
  const h = style?.h ?? "general";
  if (h === "left" || h === "fill" || h === "justify") return "left";
  if (h === "center" || h === "centerContinuous" || h === "distributed") return "center";
  if (h === "right") return "right";
  if (!cell) return "left";
  if (cell.k === "n" || cell.k === "d") return "right";
  if (cell.k === "b" || cell.k === "e") return "center";
  return "left";
}

/**
 * Columns a non-wrapped text cell may spill into: Excel lets text run over empty neighbours.
 * Returns [first, last] columns, inclusive.
 */
export function overflowSpan(
  cell: Cell,
  style: CellStyle | undefined,
  row: Map<number, Cell> | undefined,
  merges: Merges,
  limit: { c0: number; c1: number },
  /** Width the text needs and the column axis: the span then stops where the text ends, as in Excel. */
  need?: { width: number; cols: Axis },
): [number, number] {
  if (cell.k !== "s" || !cell.v || style?.wrap || style?.shrink || (style?.rotation ?? 0) !== 0) return [cell.c, cell.c];
  if (merges.at(cell.r, cell.c)) return [cell.c, cell.c];
  const align = effectiveAlign(style, cell);
  const blank = (c: number) => {
    const other = row?.get(c);
    return (!other || (other.v === "" && !other.f)) && !merges.at(cell.r, c);
  };
  const c = cell.c;
  // Centered text grows equally to both sides; left and right aligned text grows one way.
  const extra = need ? need.width - need.cols.size(c) : Infinity;
  const side = align === "center" ? extra / 2 : extra;
  const rightShort = (last: number) => !need || need.cols.end(last) - need.cols.end(c) < side;
  const leftShort = (first: number) => !need || need.cols.start(c) - need.cols.start(first) < side;
  let first = c;
  let last = c;
  if (extra <= 0) return [c, c];
  if (align === "left" || align === "center") {
    while (last + 1 <= limit.c1 && rightShort(last) && blank(last + 1)) last++;
  }
  if (align === "right" || align === "center") {
    while (first - 1 >= limit.c0 && leftShort(first) && blank(first - 1)) first--;
  }
  return [first, last];
}

/** Advance widths in em for ASCII; between Calibri and the phone system fonts, a little on the wide side. */
function asciiEm(code: number): number {
  if (code >= 48 && code <= 57) return 0.56; // digits
  if (code === 32) return 0.3;
  // , . : ; ' ! | i l j  (narrow punctuation and letters; thousands separators were counted as digits)
  if (code === 44 || code === 46 || code === 58 || code === 59 || code === 39 || code === 33 || code === 124 || code === 105 || code === 108 || code === 106) return 0.28;
  if (code === 45 || code === 40 || code === 41 || code === 47) return 0.36; // - ( ) /
  if (code === 37) return 0.78; // %
  if (code === 77 || code === 87 || code === 109 || code === 119) return 0.85; // M W m w
  if (code >= 65 && code <= 90) return 0.62; // other capitals
  if (code >= 97 && code <= 122) return 0.52; // other lowercase
  return 0.56;
}

/** Rough text width for "####", shrink-to-fit and spill decisions. */
export function estimateTextWidth(text: string, fontPx: number): number {
  let em = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    em += code >= 0x1100 ? 1 : code < 128 ? asciiEm(code) : 0.6;
  }
  return em * fontPx;
}
