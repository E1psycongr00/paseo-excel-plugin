import type { Cell, CellComment, CellLink, ColRange, DrawingItem, Range, RowInfo, SheetLayout, SheetView } from "../shared/model";
import { computePages, type PageSetup } from "./pages";
import { columnWidthPx, defaultColumnPx, defaultRowHeightPt, maxDigitWidth, ptToPx } from "./units";

export interface RawSheet {
  index: number;
  name: string;
  defaultFont: { name: string; size: number };
  baseColWidth?: number;
  defaultColWidth?: number;
  /** Points. */
  defaultRowHeight?: number;
  /** Rows are hidden unless listed (sheetFormatPr zeroHeight). */
  zeroHeight?: boolean;
  /** Widths in characters, as stored. */
  cols: { min: number; max: number; width?: number; hidden: boolean; style: number | null }[];
  /** Heights in points. */
  rows: Map<number, { ht?: number; hidden: boolean; style: number | null }>;
  cells: Cell[];
  merges: Range[];
  view: SheetView;
  setup: PageSetup;
  drawings: DrawingItem[];
  comments: CellComment[];
  links: CellLink[];
}

export interface ParsedSheet {
  layout: SheetLayout;
  /** Sorted by row, then column. Numbers may still be raw; read through cellsInRows. */
  cells: Cell[];
  finalize: (cell: Cell) => void;
}

const CHUNK_ROWS = 400;
const PAD_ROWS = 40;
const PAD_COLS = 8;
const MIN_ROWS = 60;
const MIN_COLS = 20;
const MAX_LAYOUT_ROWS = 1048576;
const MAX_LAYOUT_COLS = 16384;

export function buildSheet(raw: RawSheet, finalize: (cell: Cell) => void = () => {}): ParsedSheet {
  const mdw = maxDigitWidth(raw.defaultFont.name, raw.defaultFont.size);
  const defaultColWidth = defaultColumnPx(raw.baseColWidth, raw.defaultColWidth, mdw);
  const defaultRowHeight = Math.round(ptToPx(raw.defaultRowHeight ?? defaultRowHeightPt(raw.defaultFont.name, raw.defaultFont.size)));

  const cells = raw.cells.sort((a, b) => a.r - b.r || a.c - b.c);
  let lastRow = -1;
  let lastCol = -1;
  for (const cell of cells) {
    if (cell.r > lastRow) lastRow = cell.r;
    if (cell.c > lastCol) lastCol = cell.c;
  }
  for (const [, , r2, c2] of raw.merges) {
    lastRow = Math.max(lastRow, r2);
    lastCol = Math.max(lastCol, c2);
  }
  const usedRange: Range | null = lastRow >= 0 && lastCol >= 0 ? [0, 0, lastRow, lastCol] : null;

  // Blank cells inside a merge keep their style for the merge's border, but never show text.
  if (raw.merges.length > 0) {
    const mergesByRow = new Map<number, Range[]>();
    for (const merge of raw.merges) {
      for (let r = merge[0]; r <= merge[2]; r++) {
        const list = mergesByRow.get(r);
        if (list) list.push(merge);
        else mergesByRow.set(r, [merge]);
      }
    }
    for (const cell of cells) {
      const list = mergesByRow.get(cell.r);
      if (!list) continue;
      for (const [r1, c1, , c2] of list) {
        if (cell.c >= c1 && cell.c <= c2 && (cell.r !== r1 || cell.c !== c1)) {
          cell.v = "";
          cell.k = "z";
          delete cell.n;
          delete cell.rich;
          delete cell.f;
          delete cell.raw;
          break;
        }
      }
    }
  }

  const cols: ColRange[] = raw.cols
    .filter((col) => col.max >= col.min)
    .map((col) => ({
      min: col.min,
      max: Math.min(col.max, MAX_LAYOUT_COLS - 1),
      width: col.width === undefined ? defaultColWidth : columnWidthPx(col.width, mdw),
      hidden: col.hidden || col.width === 0,
      style: col.style,
    }))
    .sort((a, b) => a.min - b.min);

  const rows: RowInfo[] = [];
  for (const [r, row] of raw.rows) {
    const height = row.ht === undefined ? defaultRowHeight : Math.round(ptToPx(row.ht));
    const hidden = row.hidden || height === 0;
    if (height === defaultRowHeight && !hidden && row.style === null && !raw.zeroHeight) continue;
    rows.push({ r, height, hidden, style: row.style });
  }
  rows.sort((a, b) => a.r - b.r);

  for (const drawing of raw.drawings) {
    const end = drawing.anchor.to ?? drawing.anchor.from;
    lastRow = Math.max(lastRow, end.r);
    lastCol = Math.max(lastCol, end.c);
  }
  const area = raw.setup.printArea;
  const extentRow = Math.max(lastRow, area ? area[2] : -1);
  const extentCol = Math.max(lastCol, area ? area[3] : -1);
  const rowCount = Math.min(MAX_LAYOUT_ROWS, Math.max(MIN_ROWS, extentRow + 1 + PAD_ROWS, raw.view.topRow + MIN_ROWS));
  const colCount = Math.min(MAX_LAYOUT_COLS, Math.max(MIN_COLS, extentCol + 1 + PAD_COLS, raw.view.leftCol + MIN_COLS));

  const rowMap = new Map(rows.map((row) => [row.r, row]));
  const colPx = (c: number) => {
    const col = findCol(cols, c);
    if (!col) return defaultColWidth;
    return col.hidden ? 0 : col.width;
  };
  const rowPx = (r: number) => {
    const row = rowMap.get(r);
    if (!row) return raw.zeroHeight ? 0 : defaultRowHeight;
    return row.hidden ? 0 : row.height;
  };
  const print = computePages({ colPx, rowPx, usedRange, setup: raw.setup });

  return {
    cells,
    finalize,
    layout: {
      index: raw.index,
      name: raw.name,
      view: raw.view,
      defaultColWidth,
      defaultRowHeight: raw.zeroHeight ? 0 : defaultRowHeight,
      rowCount,
      colCount,
      lastRow,
      lastCol,
      cols,
      rows,
      merges: raw.merges,
      drawings: raw.drawings,
      comments: raw.comments,
      links: raw.links,
      print,
      cellCount: cells.length,
      chunkRows: CHUNK_ROWS,
    },
  };
}

function findCol(cols: ColRange[], c: number): ColRange | undefined {
  let lo = 0;
  let hi = cols.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const col = cols[mid]!;
    if (c < col.min) hi = mid - 1;
    else if (c > col.max) lo = mid + 1;
    else return col;
  }
  return undefined;
}

/** Cells whose row lies in [rowStart, rowEnd). */
export function cellsInRows(sheet: ParsedSheet, rowStart: number, rowEnd: number): Cell[] {
  const { cells } = sheet;
  let lo = 0;
  let hi = cells.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (cells[mid]!.r < rowStart) lo = mid + 1;
    else hi = mid;
  }
  const out: Cell[] = [];
  for (let i = lo; i < cells.length && cells[i]!.r < rowEnd; i++) {
    const cell = cells[i]!;
    if (cell.n !== undefined) sheet.finalize(cell);
    out.push(cell);
  }
  return out;
}

export function defaultView(): SheetView {
  return { mode: "normal", showGrid: true, showHeaders: true, zoom: 100, frozenRows: 0, frozenCols: 0, topRow: 0, leftCol: 0, rtl: false };
}
