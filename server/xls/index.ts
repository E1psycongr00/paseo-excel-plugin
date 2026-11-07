import CFB from "cfb";
import type { Cell, CellComment, CellStyle, DrawingItem, HAlign, Range, RichRun, SheetSummary, VAlign, WorkbookMeta } from "../../shared/model";
import { type ColorContext, DEFAULT_PALETTE, DEFAULT_THEME, mix, PATTERN_DENSITY } from "../colors";
import { builtinFormat, finalizeNumber, formatValue, isDateFormatCached } from "../format";
import type { PageSetup } from "../pages";
import { buildSheet, defaultView, type ParsedSheet, type RawSheet } from "../sheet-builder";
import { columnWidthPx, defaultColumnPx, defaultRowHeightPt, maxDigitWidth, ptToPx } from "../units";
import { type ImageData, type ParsedWorkbook, sniffMime } from "../workbook";
import { borderSide, type FontDef } from "../xlsx/styles";
import { type BiffRecord, f64, i16, readRecords, rk, Segments, shortString, u16, u32, xlString } from "./reader";

const R = {
  BOF: 0x0809, EOF: 0x000a, FILEPASS: 0x002f, DATEMODE: 0x0022, WINDOW1: 0x003d, FONT: 0x0031, FORMAT: 0x041e,
  XF: 0x00e0, PALETTE: 0x0092, BOUNDSHEET: 0x0085, SST: 0x00fc, NAME: 0x0018, MSODRAWINGGROUP: 0x00eb,
  DEFCOLWIDTH: 0x0055, STANDARDWIDTH: 0x0099, DEFAULTROWHEIGHT: 0x0225, COLINFO: 0x007d, ROW: 0x0208,
  BLANK: 0x0201, MULBLANK: 0x00be, NUMBER: 0x0203, RK: 0x027e, MULRK: 0x00bd, LABELSST: 0x00fd, LABEL: 0x0204,
  RSTRING: 0x00d6, BOOLERR: 0x0205, FORMULA: 0x0006, STRING: 0x0207, MERGEDCELLS: 0x00e5, WINDOW2: 0x023e,
  SCL: 0x00a0, PANE: 0x0041, HPAGEBREAKS: 0x001b, VPAGEBREAKS: 0x001a, SETUP: 0x00a1, LEFTMARGIN: 0x0026,
  RIGHTMARGIN: 0x0027, TOPMARGIN: 0x0028, BOTTOMMARGIN: 0x0029, WSBOOL: 0x0081, MSODRAWING: 0x00ec,
  OBJ: 0x005d, TXO: 0x01b6, NOTE: 0x001c,
};

const ERRORS: Record<number, string> = { 0x00: "#NULL!", 0x07: "#DIV/0!", 0x0f: "#VALUE!", 0x17: "#REF!", 0x1d: "#NAME?", 0x24: "#NUM!", 0x2a: "#N/A" };
const H: HAlign[] = ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"];
const V: VAlign[] = ["top", "center", "bottom", "justify", "distributed"];
const BORDER = ["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"];
const PATTERNS = ["none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"];

interface Globals {
  date1904: boolean;
  activeTab: number;
  fonts: FontDef[];
  formats: Map<number, string>;
  xfs: Uint8Array[];
  palette: string[];
  sheets: { name: string; state: SheetSummary["state"]; offset: number; type: number }[];
  sst: { text: string; rich?: RichRun[] }[];
  printAreas: Map<number, Range>;
  titles: Map<number, { rows: [number, number] | null; cols: [number, number] | null }>;
  blips: (ImageData | null)[];
}

export function parseXls(bytes: Uint8Array, revision: string): ParsedWorkbook {
  const container = CFB.read(bytes, { type: "array" });
  if (CFB.find(container, "EncryptionInfo")) throw new Error("암호로 보호된 파일은 열 수 없습니다.");
  const entry = CFB.find(container, "Workbook") ?? CFB.find(container, "Book");
  if (!entry || !entry.content) throw new Error("Excel 통합 문서 스트림을 찾을 수 없습니다.");
  const stream = entry.content instanceof Uint8Array ? entry.content : new Uint8Array(entry.content as ArrayLike<number>);
  const bof = readRecords(stream, 0)[0];
  if (!bof || bof.type !== R.BOF || u16(bof.data, 0) !== 0x0600) {
    throw new Error("Excel 5.0/95 이전 형식(BIFF5)은 지원하지 않습니다. Excel 97-2003 이상 형식으로 저장해 주세요.");
  }
  const globals = parseGlobals(readRecords(stream, 0));
  const colors: ColorContext = { palette: globals.palette, theme: DEFAULT_THEME };
  const defaultFont = globals.fonts[0] ?? { name: "Arial", size: 10, color: "#000000" };
  const styles = globals.xfs.map((xf) => xfStyle(xf, globals, colors, defaultFont));
  const formats = globals.xfs.map((xf) => {
    const id = u16(xf, 2);
    return globals.formats.get(id) ?? builtinFormat(id);
  });

  const worksheets = globals.sheets.filter((sheet) => sheet.type === 0 || sheet.type === 1);
  const meta: WorkbookMeta = {
    format: "xls",
    revision,
    sheets: worksheets.map((sheet, index) => ({ index, name: sheet.name, state: sheet.state, tabColor: null })),
    activeSheet: Math.min(Math.max(0, globals.activeTab), Math.max(0, worksheets.length - 1)),
    styles,
    warnings: [],
  };
  const parsed = new Map<number, ParsedSheet>();
  const imageKeys = new Map<string, number>();

  return {
    meta,
    sheet(index) {
      const cached = parsed.get(index);
      if (cached) return cached;
      const info = worksheets[index];
      if (!info) throw new Error(`시트 ${index}가 없습니다.`);
      const boundIndex = globals.sheets.indexOf(info);
      const raw = parseWorksheet(readRecords(stream, info.offset), {
        index,
        name: info.name,
        globals,
        formats,
        defaultFont,
        printArea: globals.printAreas.get(boundIndex) ?? null,
        titles: globals.titles.get(boundIndex) ?? { rows: null, cols: null },
        imageKeys,
      });
      const sheet = buildSheet(raw, (cell) => finalizeNumber(cell, formats[cell.s] ?? "General", globals.date1904));
      parsed.set(index, sheet);
      return sheet;
    },
    image(key) {
      const blip = imageKeys.get(key);
      return blip === undefined ? null : globals.blips[blip] ?? null;
    },
  };
}

function parseGlobals(records: BiffRecord[]): Globals {
  const globals: Globals = {
    date1904: false,
    activeTab: 0,
    fonts: [],
    formats: new Map(),
    xfs: [],
    palette: [...DEFAULT_PALETTE],
    sheets: [],
    sst: [],
    printAreas: new Map(),
    titles: new Map(),
    blips: [],
  };
  const fontRecords: Uint8Array[] = [];
  for (const record of records) {
    const d = record.data;
    switch (record.type) {
      case R.FILEPASS:
        throw new Error("암호로 보호된 파일은 열 수 없습니다.");
      case R.DATEMODE:
        globals.date1904 = u16(d, 0) === 1;
        break;
      case R.WINDOW1:
        globals.activeTab = u16(d, 10);
        break;
      case R.FONT:
        fontRecords.push(d);
        break;
      case R.FORMAT: {
        const id = u16(d, 0);
        globals.formats.set(id, xlString(d, 2).text);
        break;
      }
      case R.XF:
        globals.xfs.push(d);
        break;
      case R.PALETTE: {
        const count = u16(d, 0);
        for (let i = 0; i < count; i++) {
          const o = 2 + i * 4;
          globals.palette[8 + i] = [d[o]!, d[o + 1]!, d[o + 2]!].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
        }
        break;
      }
      case R.BOUNDSHEET: {
        const state = d[4]! & 0x03;
        globals.sheets.push({
          offset: u32(d, 0),
          state: state === 1 ? "hidden" : state === 2 ? "veryHidden" : "visible",
          type: d[5]!,
          name: shortString(d, 6).text,
        });
        break;
      }
      case R.SST:
        globals.sst = parseSst(record);
        break;
      case R.NAME:
        parseName(d, globals);
        break;
      case R.MSODRAWINGGROUP:
        globals.blips = parseBlipStore([d, ...record.continues]);
        break;
    }
  }
  // Colors resolve after PALETTE, which can follow the FONT records.
  const colors: ColorContext = { palette: globals.palette, theme: DEFAULT_THEME };
  globals.fonts = fontRecords.map((d) => parseFontRecord(d, colors));
  // Font index 4 does not exist in BIFF; indexes above it shift down by one.
  if (globals.fonts.length > 4) globals.fonts.splice(4, 0, globals.fonts[0]!);
  return globals;
}

function paletteColor(index: number, colors: ColorContext, fallback: string): string {
  if (index >= 64) return fallback;
  const hex = colors.palette[index];
  return hex ? `#${hex}` : fallback;
}

function parseFontRecord(d: Uint8Array, colors: ColorContext): FontDef {
  const height = u16(d, 0) / 20;
  const flags = u16(d, 2);
  const color = u16(d, 4);
  const weight = u16(d, 6);
  const script = u16(d, 8);
  const underline = d[10]!;
  const font: FontDef = { name: shortString(d, 14).text, size: height, color: paletteColor(color, colors, "#000000") };
  if (weight >= 700) font.bold = true;
  if (flags & 0x02) font.italic = true;
  if (flags & 0x08) font.strike = true;
  if (underline === 1 || underline === 0x21) font.underline = "single";
  else if (underline === 2 || underline === 0x22) font.underline = "double";
  if (script === 1) font.vert = "superscript";
  else if (script === 2) font.vert = "subscript";
  return font;
}

function xfStyle(xf: Uint8Array, globals: Globals, colors: ColorContext, defaultFont: FontDef): CellStyle {
  const font = globals.fonts[u16(xf, 0)] ?? defaultFont;
  const align = xf[6]!;
  const rotation = xf[7]!;
  const indentShrink = xf[8]!;
  const border1 = u32(xf, 10);
  const border2 = u32(xf, 14);
  const fillColors = u16(xf, 18);
  const side = (style: number, color: number) => borderSide(BORDER[style], paletteColor(color, colors, "#000000"));
  const pattern = PATTERNS[(border2 >>> 26) & 0x3f] ?? "none";
  const fore = paletteColor(fillColors & 0x7f, colors, "#000000");
  const back = paletteColor((fillColors >>> 7) & 0x7f, colors, "#FFFFFF");
  const style: CellStyle = {
    font: { ...font },
    fill: pattern === "none" ? null : pattern === "solid" ? fore : mix(fore, back, PATTERN_DENSITY[pattern] ?? 0.5),
    border: {},
    h: H[align & 0x07] ?? "general",
    v: V[(align >>> 4) & 0x07] ?? "bottom",
  };
  const border = style.border;
  const l = side(border1 & 0x0f, (border1 >>> 16) & 0x7f);
  const r = side((border1 >>> 4) & 0x0f, (border1 >>> 23) & 0x7f);
  const t = side((border1 >>> 8) & 0x0f, border2 & 0x7f);
  const b = side((border1 >>> 12) & 0x0f, (border2 >>> 7) & 0x7f);
  if (l) border.l = l;
  if (r) border.r = r;
  if (t) border.t = t;
  if (b) border.b = b;
  const diagonal = side((border2 >>> 21) & 0x0f, (border2 >>> 14) & 0x7f);
  const diagonalKind = (border1 >>> 30) & 0x03;
  if (diagonal) {
    if (diagonalKind & 0x01) border.down = diagonal;
    if (diagonalKind & 0x02) border.up = diagonal;
  }
  if (align & 0x08) style.wrap = true;
  if (indentShrink & 0x10) style.shrink = true;
  if (indentShrink & 0x0f) style.indent = indentShrink & 0x0f;
  if (rotation) style.rotation = rotation;
  return style;
}

function parseSst(record: BiffRecord): { text: string; rich?: RichRun[] }[] {
  const reader = new Segments([record.data, ...record.continues]);
  reader.skip(4);
  const unique = reader.u32();
  const out: { text: string; rich?: RichRun[] }[] = [];
  for (let i = 0; i < unique && !reader.exhausted; i++) {
    const count = reader.u16();
    const flags = reader.u8();
    const runs = flags & 0x08 ? reader.u16() : 0;
    const ext = flags & 0x04 ? reader.u32() : 0;
    const text = reader.chars(count, (flags & 0x01) === 1);
    const marks: [number, number][] = [];
    for (let r = 0; r < runs; r++) marks.push([reader.u16(), reader.u16()]);
    reader.skip(ext);
    out.push(marks.length > 0 ? { text, rich: richRunsFromMarks(text, marks) } : { text });
  }
  return out;
}

// Rich runs carry font indexes; resolve them to fonts when the cell is built.
function richRunsFromMarks(text: string, marks: [number, number][]): RichRun[] {
  const runs: RichRun[] = [];
  let start = 0;
  let font: number | null = null;
  for (const [at, fontIndex] of marks) {
    if (at > start) runs.push({ text: text.slice(start, at), font: font === null ? undefined : ({ name: `#font${font}` } as RichRun["font"]) });
    start = at;
    font = fontIndex;
  }
  runs.push({ text: text.slice(start), font: font === null ? undefined : ({ name: `#font${font}` } as RichRun["font"]) });
  return runs.filter((run) => run.text.length > 0);
}

function resolveRich(runs: RichRun[] | undefined, fonts: FontDef[]): RichRun[] | undefined {
  if (!runs) return undefined;
  return runs.map((run) => {
    const marker = run.font?.name;
    if (!marker?.startsWith("#font")) return run;
    const font = fonts[Number(marker.slice(5))];
    return font ? { text: run.text, font: { ...font } } : { text: run.text };
  });
}

function parseName(d: Uint8Array, globals: Globals) {
  const flags = u16(d, 0);
  const nameLength = d[3]!;
  const formulaLength = u16(d, 4);
  const sheet = u16(d, 8);
  if (!(flags & 0x20) || sheet === 0 || nameLength === 0) return;
  const wide = (d[14]! & 1) === 1;
  const code = wide ? u16(d, 15) : d[15]!;
  const formulaStart = 15 + nameLength * (wide ? 2 : 1);
  const areas = parseAreaPtgs(d.subarray(formulaStart, formulaStart + formulaLength));
  if (code === 0x06 && areas[0]) {
    globals.printAreas.set(sheet - 1, areas[0]);
  } else if (code === 0x07) {
    const titles: { rows: [number, number] | null; cols: [number, number] | null } = { rows: null, cols: null };
    for (const area of areas) {
      if (area[1] === 0 && area[3] >= 255) titles.rows = [area[0], area[2]];
      else if (area[0] === 0 && area[2] >= 65535) titles.cols = [area[1], area[3]];
    }
    globals.titles.set(sheet - 1, titles);
  }
}

function parseAreaPtgs(rgce: Uint8Array): Range[] {
  const out: Range[] = [];
  let i = 0;
  while (i < rgce.length) {
    const ptg = rgce[i]!;
    if (ptg === 0x29 || ptg === 0x49 || ptg === 0x69) i += 3; // PtgMemFunc + cce
    else if (ptg === 0x10) i += 1; // PtgUnion
    else if (ptg === 0x3b || ptg === 0x5b || ptg === 0x7b) {
      out.push([u16(rgce, i + 3), u16(rgce, i + 7) & 0x3fff, u16(rgce, i + 5), u16(rgce, i + 9) & 0x3fff]);
      i += 11;
    } else if (ptg === 0x3a || ptg === 0x5a || ptg === 0x7a) {
      const row = u16(rgce, i + 3);
      const col = u16(rgce, i + 5) & 0x3fff;
      out.push([row, col, row, col]);
      i += 7;
    } else break;
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Office drawing (Escher) records

interface Escher {
  ver: number;
  instance: number;
  type: number;
  data: Uint8Array;
}

function escherRecords(bytes: Uint8Array, start = 0, end = bytes.length): Escher[] {
  const out: Escher[] = [];
  let o = start;
  while (o + 8 <= end) {
    const verInstance = u16(bytes, o);
    const type = u16(bytes, o + 2);
    const length = u32(bytes, o + 4);
    out.push({ ver: verInstance & 0x0f, instance: verInstance >>> 4, type, data: bytes.subarray(o + 8, Math.min(end, o + 8 + length)) });
    o += 8 + length;
  }
  return out;
}

function concat(parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const part of parts) {
    out.set(part, o);
    o += part.length;
  }
  return out;
}

function parseBlipStore(parts: Uint8Array[]): (ImageData | null)[] {
  const blips: (ImageData | null)[] = [];
  const walk = (records: Escher[]) => {
    for (const record of records) {
      if (record.type === 0xf007) {
        blips.push(readFbse(record.data));
      } else if (record.ver === 0x0f) {
        walk(escherRecords(record.data));
      }
    }
  };
  try {
    walk(escherRecords(concat(parts)));
  } catch {
    // A damaged drawing group only costs the pictures.
  }
  return blips;
}

function readFbse(d: Uint8Array): ImageData | null {
  const nameLength = d[33]!;
  const start = 36 + nameLength;
  if (start + 8 > d.length) return null;
  const [blip] = escherRecords(d, start);
  if (!blip) return null;
  const twoUids = (blip.instance & 0x01) === 1;
  switch (blip.type) {
    case 0xf01d: // JPEG
    case 0xf02a: // JPEG (CMYK)
    case 0xf01e: // PNG
    case 0xf029: { // TIFF
      const body = blip.data.subarray(twoUids ? 33 : 17);
      return { bytes: body, mimeType: sniffMime(body) };
    }
    case 0xf01f: { // DIB: prepend a BITMAPFILEHEADER
      const dib = blip.data.subarray(twoUids ? 33 : 17);
      const header = new Uint8Array(14);
      const size = dib.length + 14;
      const headerSize = u32(dib, 0);
      const bitCount = u16(dib, 14);
      const colorsUsed = u32(dib, 32);
      const paletteSize = bitCount <= 8 ? (colorsUsed || 1 << bitCount) * 4 : 0;
      header[0] = 0x42;
      header[1] = 0x4d;
      new DataView(header.buffer).setUint32(2, size, true);
      new DataView(header.buffer).setUint32(10, 14 + headerSize + paletteSize, true);
      return { bytes: concat([header, dib]), mimeType: "image/bmp" };
    }
    default:
      return null; // EMF/WMF are not displayable without a metafile renderer.
  }
}

interface ShapeAnchor {
  colL: number; dxL: number; rwT: number; dyT: number;
  colR: number; dxR: number; rwB: number; dyB: number;
}

interface SheetShape {
  anchor: ShapeAnchor;
  blip: number | null;
  shapeType: number;
  text?: string;
  fill?: string | null;
}

function parseSheetShapes(bytes: Uint8Array, palette: string[]): SheetShape[] {
  const shapes: SheetShape[] = [];
  const walk = (records: Escher[]) => {
    for (const record of records) {
      if (record.type === 0xf004) {
        const inner = escherRecords(record.data);
        let anchor: ShapeAnchor | null = null;
        let blip: number | null = null;
        let shapeType = 0;
        let fill: string | null | undefined;
        for (const item of inner) {
          if (item.type === 0xf00a) shapeType = item.instance;
          else if (item.type === 0xf010 && item.data.length >= 18) {
            const a = item.data;
            anchor = { colL: u16(a, 2), dxL: i16(a, 4), rwT: u16(a, 6), dyT: i16(a, 8), colR: u16(a, 10), dxR: i16(a, 12), rwB: u16(a, 14), dyB: i16(a, 16) };
          } else if (item.type === 0xf00b) {
            const count = item.instance;
            for (let i = 0; i < count; i++) {
              const id = u16(item.data, i * 6) & 0x3fff;
              const value = u32(item.data, i * 6 + 2);
              if (id === 0x0104) blip = value;
              else if (id === 0x0181) fill = bgrColor(value, palette);
            }
          } else if (item.ver === 0x0f) {
            walk([item]);
          }
        }
        if (anchor) shapes.push({ anchor, blip, shapeType, fill });
      } else if (record.ver === 0x0f) {
        walk(escherRecords(record.data));
      }
    }
  };
  try {
    walk(escherRecords(bytes));
  } catch {
    // Ignore damaged drawings.
  }
  return shapes;
}

function bgrColor(value: number, palette: string[]): string | null {
  if (value & 0x08000000) {
    const hex = palette[value & 0xff];
    return hex ? `#${hex}` : null;
  }
  const r = value & 0xff;
  const g = (value >>> 8) & 0xff;
  const b = (value >>> 16) & 0xff;
  return `#${[r, g, b].map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

// ---------------------------------------------------------------------------------------------
// Worksheets

interface SheetInput {
  index: number;
  name: string;
  globals: Globals;
  formats: string[];
  defaultFont: FontDef;
  printArea: Range | null;
  titles: { rows: [number, number] | null; cols: [number, number] | null };
  imageKeys: Map<string, number>;
}

function parseWorksheet(records: BiffRecord[], input: SheetInput): RawSheet {
  const { globals, formats } = input;
  const view = defaultView();
  const setup: PageSetup = { manualRowBreaks: [], manualColBreaks: [], printArea: input.printArea, titleRows: input.titles.rows, titleCols: input.titles.cols };
  const raw: RawSheet = {
    index: input.index,
    name: input.name,
    defaultFont: { name: input.defaultFont.name, size: input.defaultFont.size },
    cols: [],
    rows: new Map(),
    cells: [],
    merges: [],
    view,
    setup,
    drawings: [],
    comments: [],
    links: [],
  };
  const drawingParts: Uint8Array[] = [];
  const noteTexts = new Map<number, string>();
  const notes: { r: number; c: number; id: number; author: string }[] = [];
  let lastObjId: number | null = null;
  let pendingFormula: Cell | null = null;
  let zoomFromScl: number | null = null;
  let slvZoom = 0;
  let normalZoom = 0;
  let frozen = false;

  const push = (cell: Cell) => raw.cells.push(cell);
  const numberCell = (r: number, c: number, s: number, value: number): Cell => {
    const format = formats[s] ?? "General";
    return { r, c, s, v: "", n: value, k: isDateFormatCached(format) ? "d" : "n" };
  };
  const stringCell = (r: number, c: number, s: number, text: string, rich?: RichRun[]): Cell => {
    const format = formats[s] ?? "General";
    const cell: Cell = { r, c, s, v: format.includes("@") && format !== "@" ? formatValue(format, text, globals.date1904) : text, k: "s" };
    const resolved = resolveRich(rich, globals.fonts);
    if (resolved && resolved.some((run) => run.font)) cell.rich = resolved;
    return cell;
  };

  for (let i = 1; i < records.length; i++) {
    const record = records[i]!;
    const d = record.data;
    if (record.type === R.BOF) {
      // Embedded chart substream: skip to its EOF.
      let depth = 1;
      while (depth > 0 && ++i < records.length) {
        if (records[i]!.type === R.BOF) depth++;
        else if (records[i]!.type === R.EOF) depth--;
      }
      continue;
    }
    if (record.type === R.EOF) break;
    try {
      switch (record.type) {
        case R.DEFCOLWIDTH:
          raw.baseColWidth = u16(d, 0);
          break;
        case R.STANDARDWIDTH:
          raw.defaultColWidth = u16(d, 0) / 256;
          break;
        case R.DEFAULTROWHEIGHT:
          if (u16(d, 0) & 0x02) raw.zeroHeight = true;
          raw.defaultRowHeight = u16(d, 2) / 20;
          break;
        case R.COLINFO: {
          const xf = u16(d, 6);
          raw.cols.push({ min: u16(d, 0), max: Math.min(u16(d, 2), 16383), width: u16(d, 4) / 256, hidden: (u16(d, 8) & 0x01) === 1, style: xf > 15 ? xf : null });
          break;
        }
        case R.ROW: {
          const r = u16(d, 0);
          const height = u16(d, 6) & 0x7fff;
          const flags = u16(d, 12);
          const hidden = (flags & 0x20) !== 0;
          const custom = (flags & 0x40) !== 0;
          const formatted = (flags & 0x80) !== 0;
          const style = formatted ? u16(d, 14) & 0x0fff : null;
          raw.rows.set(r, { ht: custom || height ? height / 20 : undefined, hidden, style });
          break;
        }
        case R.BLANK:
          push({ r: u16(d, 0), c: u16(d, 2), s: u16(d, 4), v: "", k: "z" });
          break;
        case R.MULBLANK: {
          const r = u16(d, 0);
          const first = u16(d, 2);
          const count = (d.length - 6) / 2;
          for (let k = 0; k < count; k++) push({ r, c: first + k, s: u16(d, 4 + k * 2), v: "", k: "z" });
          break;
        }
        case R.NUMBER:
          push(numberCell(u16(d, 0), u16(d, 2), u16(d, 4), f64(d, 6)));
          break;
        case R.RK:
          push(numberCell(u16(d, 0), u16(d, 2), u16(d, 4), rk(u32(d, 6))));
          break;
        case R.MULRK: {
          const r = u16(d, 0);
          const first = u16(d, 2);
          const count = (d.length - 6) / 6;
          for (let k = 0; k < count; k++) push(numberCell(r, first + k, u16(d, 4 + k * 6), rk(u32(d, 6 + k * 6))));
          break;
        }
        case R.LABELSST: {
          const entry = globals.sst[u32(d, 6)];
          push(stringCell(u16(d, 0), u16(d, 2), u16(d, 4), entry?.text ?? "", entry?.rich));
          break;
        }
        case R.LABEL:
        case R.RSTRING:
          push(stringCell(u16(d, 0), u16(d, 2), u16(d, 4), xlString(d, 6).text));
          break;
        case R.BOOLERR: {
          const value = d[6]!;
          const isError = d[7] === 1;
          push({ r: u16(d, 0), c: u16(d, 2), s: u16(d, 4), v: isError ? ERRORS[value] ?? "#ERR" : value ? "TRUE" : "FALSE", k: isError ? "e" : "b" });
          break;
        }
        case R.FORMULA: {
          const r = u16(d, 0);
          const c = u16(d, 2);
          const s = u16(d, 4);
          let cell: Cell;
          if (d[12] === 0xff && d[13] === 0xff) {
            const kind = d[6]!;
            if (kind === 0) {
              cell = { r, c, s, v: "", k: "s" };
              pendingFormula = cell;
            } else if (kind === 1) cell = { r, c, s, v: d[8] ? "TRUE" : "FALSE", k: "b" };
            else if (kind === 2) cell = { r, c, s, v: ERRORS[d[8]!] ?? "#ERR", k: "e" };
            else cell = { r, c, s, v: "", k: "s" };
          } else {
            cell = numberCell(r, c, s, f64(d, 6));
          }
          cell.f = "…";
          push(cell);
          break;
        }
        case R.STRING:
          if (pendingFormula) {
            const reader = new Segments([d, ...record.continues]);
            const count = reader.u16();
            const flags = reader.u8();
            const text = reader.chars(count, (flags & 1) === 1);
            const format = formats[pendingFormula.s] ?? "General";
            pendingFormula.v = format.includes("@") && format !== "@" ? formatValue(format, text, globals.date1904) : text;
            pendingFormula = null;
          }
          break;
        case R.MERGEDCELLS: {
          const count = u16(d, 0);
          for (let k = 0; k < count; k++) {
            const o = 2 + k * 8;
            const range: Range = [u16(d, o), u16(d, o + 4), u16(d, o + 2), u16(d, o + 6)];
            if (range[0] !== range[2] || range[1] !== range[3]) raw.merges.push(range);
          }
          break;
        }
        case R.WINDOW2: {
          const flags = u16(d, 0);
          if (!(flags & 0x02)) view.showGrid = false;
          if (!(flags & 0x04)) view.showHeaders = false;
          if (flags & 0x40) view.rtl = true;
          if (flags & 0x0800) view.mode = "pageBreakPreview";
          view.topRow = u16(d, 2);
          view.leftCol = u16(d, 4);
          if (d.length >= 14) {
            slvZoom = u16(d, 10);
            normalZoom = u16(d, 12);
          }
          frozen = (flags & 0x08) !== 0;
          break;
        }
        case R.SCL:
          if (u16(d, 2)) zoomFromScl = Math.round((u16(d, 0) / u16(d, 2)) * 100);
          break;
        case R.PANE:
          // Split panes without freezing scroll together in this viewer.
          if (frozen) {
            view.frozenCols = u16(d, 0);
            view.frozenRows = u16(d, 2);
          }
          break;
        case R.HPAGEBREAKS: {
          const count = u16(d, 0);
          for (let k = 0; k < count; k++) setup.manualRowBreaks.push(u16(d, 2 + k * 6));
          break;
        }
        case R.VPAGEBREAKS: {
          const count = u16(d, 0);
          for (let k = 0; k < count; k++) setup.manualColBreaks.push(u16(d, 2 + k * 6));
          break;
        }
        case R.SETUP: {
          const flags = u16(d, 10);
          if (!(flags & 0x04)) {
            setup.paperSize = u16(d, 0) || 9;
            setup.scale = u16(d, 2) || 100;
            setup.orientation = flags & 0x02 ? "portrait" : "landscape";
          }
          setup.fitToWidth = u16(d, 6);
          setup.fitToHeight = u16(d, 8);
          if (flags & 0x01) setup.order = "over";
          break;
        }
        case R.LEFTMARGIN:
        case R.RIGHTMARGIN:
        case R.TOPMARGIN:
        case R.BOTTOMMARGIN: {
          setup.margins ??= { left: 0.75, right: 0.75, top: 1, bottom: 1 };
          const value = f64(d, 0);
          if (record.type === R.LEFTMARGIN) setup.margins.left = value;
          else if (record.type === R.RIGHTMARGIN) setup.margins.right = value;
          else if (record.type === R.TOPMARGIN) setup.margins.top = value;
          else setup.margins.bottom = value;
          break;
        }
        case R.WSBOOL:
          if (u16(d, 0) & 0x0100) setup.fitToPage = true;
          break;
        case R.MSODRAWING:
          drawingParts.push(d, ...record.continues);
          break;
        case R.OBJ:
          // ftCmo: ft(2) cb(2) ot(2) id(2)
          if (u16(d, 0) === 0x15) lastObjId = u16(d, 4) === 0x19 ? u16(d, 6) : null;
          break;
        case R.TXO:
          if (lastObjId !== null && record.continues.length > 0) {
            const count = u16(d, 10);
            const reader = new Segments(record.continues);
            const flags = reader.u8();
            noteTexts.set(lastObjId, reader.chars(count, (flags & 1) === 1));
          }
          break;
        case R.NOTE:
          notes.push({ r: u16(d, 0), c: u16(d, 2), id: u16(d, 6), author: xlString(d, 8).text });
          break;
      }
    } catch {
      // A truncated record should not drop the rest of the sheet.
    }
  }
  if (view.mode === "pageBreakPreview") view.zoom = zoomFromScl ?? (slvZoom || 60);
  else view.zoom = zoomFromScl ?? (normalZoom || 100);
  setup.paperSize ??= 9;

  raw.comments = notes.map<CellComment>((note) => ({ r: note.r, c: note.c, author: note.author, text: (noteTexts.get(note.id) ?? "").replace(/\r\n?/g, "\n").trim() }));
  raw.drawings = drawingItems(parseSheetShapes(concat(drawingParts), globals.palette), raw, input);
  return raw;
}

// Converts BIFF anchors (1/1024 of a column, 1/256 of a row) to px offsets.
function drawingItems(shapes: SheetShape[], raw: RawSheet, input: SheetInput): DrawingItem[] {
  const mdw = maxDigitWidth(raw.defaultFont.name, raw.defaultFont.size);
  const defaultCol = defaultColumnPx(raw.baseColWidth, raw.defaultColWidth, mdw);
  const defaultRow = ptToPx(raw.defaultRowHeight ?? defaultRowHeightPt(raw.defaultFont.name, raw.defaultFont.size));
  const colPx = (c: number) => {
    const col = raw.cols.find((entry) => c >= entry.min && c <= entry.max);
    if (!col) return defaultCol;
    return col.hidden ? 0 : col.width === undefined ? defaultCol : columnWidthPx(col.width, mdw);
  };
  const rowPx = (r: number) => {
    const row = raw.rows.get(r);
    if (!row) return defaultRow;
    return row.hidden ? 0 : row.ht === undefined ? defaultRow : ptToPx(row.ht);
  };
  const items: DrawingItem[] = [];
  shapes.forEach((shape, index) => {
    const a = shape.anchor;
    const anchor: DrawingItem["anchor"] = {
      from: { c: a.colL, dx: (a.dxL / 1024) * colPx(a.colL), r: a.rwT, dy: (a.dyT / 256) * rowPx(a.rwT) },
      to: { c: a.colR, dx: (a.dxR / 1024) * colPx(a.colR), r: a.rwB, dy: (a.dyB / 256) * rowPx(a.rwB) },
    };
    const id = `xls:${input.index}:${index}`;
    // Comment boxes (shape type 202) are drawn by the comment marker instead.
    if (shape.shapeType === 202) return;
    if (shape.blip !== null && shape.blip > 0 && input.globals.blips[shape.blip - 1]) {
      const key = `blip:${shape.blip - 1}`;
      input.imageKeys.set(key, shape.blip - 1);
      items.push({ id, kind: "image", anchor, image: key });
    } else if (shape.shapeType !== 0 && shape.shapeType !== 201) {
      items.push({ id, kind: "shape", anchor, fill: shape.fill ?? null, line: "#000000" });
    } else if (shape.shapeType === 201) {
      // Form controls and charts share type 201 in BIFF drawings.
    }
  });
  return items;
}
