// Render model shared by the daemon parsers and the app renderer. Everything is plain JSON.
// Sizes are CSS pixels at 100% zoom; colors are "#RRGGBB" (or "#RRGGBBAA" when translucent).

export type SheetState = "visible" | "hidden" | "veryHidden";

export interface SheetSummary {
  index: number;
  name: string;
  state: SheetState;
  tabColor: string | null;
}

export interface WorkbookMeta {
  format: "xlsx" | "xls";
  revision: string;
  sheets: SheetSummary[];
  /** Sheet selected when the file was saved. */
  activeSheet: number;
  styles: CellStyle[];
  warnings: string[];
}

export type BorderLine = "solid" | "dashed" | "dotted" | "double";

export interface BorderSide {
  /** Line width in px at 100% zoom. */
  w: number;
  line: BorderLine;
  color: string;
}

export interface CellStyle {
  font: {
    name: string;
    /** Points. */
    size: number;
    bold?: boolean;
    italic?: boolean;
    underline?: "single" | "double";
    strike?: boolean;
    color: string;
    vert?: "superscript" | "subscript";
  };
  fill: string | null;
  border: {
    l?: BorderSide;
    r?: BorderSide;
    t?: BorderSide;
    b?: BorderSide;
    /** Diagonal from bottom-left to top-right. */
    up?: BorderSide;
    /** Diagonal from top-left to bottom-right. */
    down?: BorderSide;
  };
  h: HAlign;
  v: VAlign;
  wrap?: boolean;
  shrink?: boolean;
  indent?: number;
  /** 0-180 as in OOXML, or 255 for vertical stacked text. */
  rotation?: number;
}

export type HAlign = "general" | "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
export type VAlign = "top" | "center" | "bottom" | "justify" | "distributed";

export interface RichRun {
  text: string;
  font?: Partial<CellStyle["font"]>;
}

export type CellKind = "n" | "s" | "b" | "e" | "d" | "z";

/** A cell, addressed by 0-based row and column. */
export interface Cell {
  r: number;
  c: number;
  /** Style index into WorkbookMeta.styles. */
  s: number;
  /** Displayed text after number formatting. Empty for blank styled cells. */
  v: string;
  /** Value kind: number, string, boolean, error, date, blank. */
  k: CellKind;
  /** Formula text without the leading "=". */
  f?: string;
  /** Raw value when it differs from the displayed text. */
  raw?: string;
  rich?: RichRun[];
  /** Color forced by the number format, e.g. [Red]. */
  fc?: string;
  /** Daemon-internal: a number not yet formatted. Never sent to the app. */
  n?: number;
}

export interface ColRange {
  /** First and last 0-based column, inclusive. */
  min: number;
  max: number;
  width: number;
  hidden: boolean;
  style: number | null;
}

export interface RowInfo {
  r: number;
  height: number;
  hidden: boolean;
  style: number | null;
}

/** [r1, c1, r2, c2] inclusive, 0-based. */
export type Range = [number, number, number, number];

export interface Anchor {
  /** Cell position with px offsets inside the cell. */
  from: { r: number; c: number; dx: number; dy: number };
  /** Either an end cell or an extent in px. */
  to?: { r: number; c: number; dx: number; dy: number };
  ext?: { w: number; h: number };
  /** Absolute position in px for absoluteAnchor. */
  pos?: { x: number; y: number };
}

export interface DrawingItem {
  id: string;
  kind: "image" | "shape" | "chart";
  anchor: Anchor;
  /** Image key for image.chunk. */
  image?: string;
  text?: string;
  fill?: string | null;
  line?: string | null;
  name?: string;
}

export interface CellComment {
  r: number;
  c: number;
  author: string;
  text: string;
}

export interface CellLink {
  r: number;
  c: number;
  target: string;
  /** In-workbook location such as Sheet2!A1. */
  internal?: boolean;
}

export interface PageBreak {
  /** Index of the first row (or column) of the next page. */
  at: number;
  manual: boolean;
}

export interface PrintInfo {
  /** Explicit print area, or null when Excel prints the used range. */
  area: Range | null;
  /** Range the page grid covers. */
  pageRange: Range | null;
  rowBreaks: PageBreak[];
  colBreaks: PageBreak[];
  /** "down" = down, then over (Excel default). */
  order: "down" | "over";
  /** Effective print scale in percent. */
  scale: number;
  /** True when automatic breaks come from our own page computation. */
  estimated: boolean;
}

export interface SheetView {
  mode: "normal" | "pageBreakPreview" | "pageLayout";
  showGrid: boolean;
  showHeaders: boolean;
  /** Percent. */
  zoom: number;
  frozenRows: number;
  frozenCols: number;
  topRow: number;
  leftCol: number;
  rtl: boolean;
}

export interface SheetLayout {
  index: number;
  name: string;
  view: SheetView;
  defaultColWidth: number;
  defaultRowHeight: number;
  /** Extent the renderer lays out, including blank padding. */
  rowCount: number;
  colCount: number;
  /** Last used row/column (0-based), -1 when empty. */
  lastRow: number;
  lastCol: number;
  cols: ColRange[];
  /** Rows that differ from the default. */
  rows: RowInfo[];
  merges: Range[];
  drawings: DrawingItem[];
  comments: CellComment[];
  links: CellLink[];
  print: PrintInfo;
  cellCount: number;
  /** Rows per chunk for sheet.cells. */
  chunkRows: number;
}

export interface CellChunk {
  sheet: number;
  rowStart: number;
  rowEnd: number;
  cells: Cell[];
}

/** A sheet's layout with the cell chunks its first screen needs. */
export interface SheetBundle {
  layout: SheetLayout;
  chunks: CellChunk[];
}

export interface OpenResult {
  meta: WorkbookMeta;
  /** Bundle for `initialSheet(meta, requested)`; absent when that sheet failed to parse. */
  initial: SheetBundle | null;
}

/** Sheet a tab opens on: the saved one, else the active sheet if visible, else the first visible. */
export function initialSheet(meta: WorkbookMeta, saved: number | undefined): number {
  if (saved !== undefined && saved < meta.sheets.length) return saved;
  if (meta.sheets[meta.activeSheet]?.state === "visible") return meta.activeSheet;
  return meta.sheets.find((sheet) => sheet.state === "visible")?.index ?? 0;
}
