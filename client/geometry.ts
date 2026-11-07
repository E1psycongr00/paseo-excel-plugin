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

export function buildAxes(layout: SheetLayout, scale: number, showHidden: boolean) {
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
      if (col.hidden) return showHidden ? col.width || layout.defaultColWidth : 0;
      return col.width;
    },
    (c) => findCol(c)?.hidden ?? false,
    scale,
  );
  const rows = new Axis(
    layout.rowCount,
    (r) => {
      const row = rowMap.get(r);
      if (!row) return zeroRows ? (showHidden ? fallbackRow : 0) : layout.defaultRowHeight;
      if (row.hidden) return showHidden ? row.height || fallbackRow : 0;
      return row.height;
    },
    (r) => {
      const row = rowMap.get(r);
      return row ? row.hidden : zeroRows;
    },
    scale,
  );
  return { cols, rows, findCol, rowMap };
}

export interface Window {
  r0: number;
  r1: number;
  c0: number;
  c1: number;
}

const ROW_BLOCK = 16;
const COL_BLOCK = 6;

/** Visible cells plus overscan, snapped to blocks so scrolling re-renders rarely. */
export function visibleWindow(rows: Axis, cols: Axis, x: number, y: number, width: number, height: number): Window {
  const r0 = Math.floor(Math.max(0, rows.indexAt(y) - 4) / ROW_BLOCK) * ROW_BLOCK;
  const r1 = Math.min(rows.count - 1, (Math.floor((rows.indexAt(y + height) + 4) / ROW_BLOCK) + 1) * ROW_BLOCK - 1);
  const c0 = Math.floor(Math.max(0, cols.indexAt(x) - 2) / COL_BLOCK) * COL_BLOCK;
  const c1 = Math.min(cols.count - 1, (Math.floor((cols.indexAt(x + width) + 2) / COL_BLOCK) + 1) * COL_BLOCK - 1);
  return { r0, r1, c0, c1 };
}

export const sameWindow = (a: Window, b: Window) => a.r0 === b.r0 && a.r1 === b.r1 && a.c0 === b.c0 && a.c1 === b.c1;

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
): [number, number] {
  if (cell.k !== "s" || !cell.v || style?.wrap || style?.shrink || (style?.rotation ?? 0) !== 0) return [cell.c, cell.c];
  if (merges.at(cell.r, cell.c)) return [cell.c, cell.c];
  const align = effectiveAlign(style, cell);
  const blank = (c: number) => {
    const other = row?.get(c);
    return (!other || (other.v === "" && !other.f)) && !merges.at(cell.r, c);
  };
  let first = cell.c;
  let last = cell.c;
  if (align === "left" || align === "center") {
    while (last + 1 <= limit.c1 && blank(last + 1)) last++;
  }
  if (align === "right" || align === "center") {
    while (first - 1 >= limit.c0 && blank(first - 1)) first--;
  }
  return [first, last];
}

/** Rough text width for "####" and shrink-to-fit decisions. */
export function estimateTextWidth(text: string, fontPx: number): number {
  let width = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    width += code >= 0x1100 ? fontPx : code === 32 ? fontPx * 0.3 : code >= 48 && code <= 57 ? fontPx * 0.56 : fontPx * 0.55;
  }
  return width;
}
