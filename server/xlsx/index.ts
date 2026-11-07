import { strFromU8 } from "fflate";
import type { Cell, CellComment, CellLink, CellStyle, DrawingItem, Range, RichRun, SheetSummary, SheetView, WorkbookMeta } from "../../shared/model";
import { applyTint } from "../colors";
import { finalizeNumber, formatValue, isDateFormatCached } from "../format";
import type { PageSetup } from "../pages";
import { parseAreaList, parseCellRef, parseRangeRef } from "../refs";
import { buildSheet, defaultView, type ParsedSheet, type RawSheet } from "../sheet-builder";
import { emuToPx } from "../units";
import { mimeFromName, type ImageData, type ParsedWorkbook, summaryState } from "../workbook";
import { type Attrs, child, children, decodeOoxmlEscapes, descendants, parseTree, scanXml, textOf, type XmlNode } from "../xml";
import { type FontDef, type ParsedStyles, parseStyles, runFont } from "./styles";
import { IMAGE_LIMITS, unzipLimited } from "./zip";

interface Rel {
  type: string;
  target: string;
  external: boolean;
}

const REL_SUFFIX = {
  sharedStrings: "/sharedStrings",
  styles: "/styles",
  theme: "/theme",
  drawing: "/drawing",
  comments: "/comments",
  hyperlink: "/hyperlink",
  image: "/image",
  chart: "/chart",
};

function resolvePath(base: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = base.split("/");
  parts.pop();
  for (const segment of target.split("/")) {
    if (segment === "..") parts.pop();
    else if (segment !== "." && segment !== "") parts.push(segment);
  }
  return parts.join("/");
}

function relsPath(part: string): string {
  const slash = part.lastIndexOf("/");
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

interface SharedString {
  text: string;
  rich?: RichRun[];
}

export function parseXlsx(bytes: Uint8Array, revision: string): ParsedWorkbook {
  const files = unzipLimited(bytes, (name) => !/^xl\/media\//i.test(name) && !/\.bin$/i.test(name));
  const read = (path: string): string | undefined => {
    const entry = files[path] ?? files[Object.keys(files).find((name) => name.toLowerCase() === path.toLowerCase()) ?? ""];
    return entry ? strFromU8(entry) : undefined;
  };
  const relsCache = new Map<string, Map<string, Rel>>();
  const rels = (part: string): Map<string, Rel> => {
    const cached = relsCache.get(part);
    if (cached) return cached;
    const map = new Map<string, Rel>();
    const xml = read(relsPath(part));
    if (xml) {
      scanXml(xml, {
        open(name, attrs) {
          if (name !== "Relationship" || !attrs.Id) return;
          const external = attrs.TargetMode === "External";
          map.set(attrs.Id, { type: attrs.Type ?? "", target: external ? attrs.Target ?? "" : resolvePath(part, attrs.Target ?? ""), external });
        },
      });
    }
    relsCache.set(part, map);
    return map;
  };
  const relOfType = (part: string, suffix: string) => [...rels(part).values()].find((rel) => rel.type.endsWith(suffix));

  const rootRels = rels("");
  const officeDoc = [...rootRels.values()].find((rel) => rel.type.endsWith("/officeDocument"));
  const workbookPath = officeDoc?.target ?? "xl/workbook.xml";
  const workbookXml = read(workbookPath);
  if (!workbookXml) throw new Error("통합 문서(workbook.xml)를 찾을 수 없습니다.");
  const workbook = parseTree(workbookXml);
  const book = child(workbook, "workbook");
  const date1904 = ["1", "true"].includes(child(book, "workbookPr")?.attrs.date1904 ?? "");
  const activeSheet = Number(child(child(book, "bookViews"), "workbookView")?.attrs.activeTab ?? 0);

  const stylesRel = relOfType(workbookPath, REL_SUFFIX.styles);
  const themeRel = relOfType(workbookPath, REL_SUFFIX.theme);
  const parsedStyles = parseStyles(stylesRel ? read(stylesRel.target) : undefined, themeRel ? read(themeRel.target) : undefined);
  const sharedRel = relOfType(workbookPath, REL_SUFFIX.sharedStrings);
  const shared = sharedRel ? parseSharedStrings(read(sharedRel.target) ?? "", parsedStyles) : [];

  const sheetNodes = children(child(book, "sheets"), "sheet");
  const sheets: (SheetSummary & { part: string | null })[] = sheetNodes.map((node, index) => {
    const rel = rels(workbookPath).get(node.attrs.id ?? "");
    return { index, name: node.attrs.name ?? `Sheet${index + 1}`, state: summaryState(node.attrs.state), tabColor: null, part: rel?.target ?? null };
  });

  const printAreas = new Map<number, Range>();
  const printTitles = new Map<number, { rows: [number, number] | null; cols: [number, number] | null }>();
  for (const name of children(child(book, "definedNames"), "definedName")) {
    const local = name.attrs.localSheetId;
    if (local === undefined) continue;
    const index = Number(local);
    const formula = textOf(name);
    if (name.attrs.name === "_xlnm.Print_Area") {
      const first = parseAreaList(formula)[0];
      if (first) printAreas.set(index, first.range);
    } else if (name.attrs.name === "_xlnm.Print_Titles") {
      const titles: { rows: [number, number] | null; cols: [number, number] | null } = { rows: null, cols: null };
      for (const { range } of parseAreaList(formula)) {
        if (range[1] === 0 && range[3] === 16383) titles.rows = [range[0], range[2]];
        else if (range[0] === 0 && range[2] === 1048575) titles.cols = [range[1], range[3]];
      }
      printTitles.set(index, titles);
    }
  }

  const imageParts = new Map<string, string>();
  const parsedSheets = new Map<number, ParsedSheet>();
  const warnings: string[] = [];

  // Tab colors live in each sheet's sheetPr; read them cheaply up front.
  for (const sheet of sheets) {
    if (!sheet.part) continue;
    const xml = read(sheet.part);
    if (!xml) continue;
    const head = xml.slice(0, Math.min(xml.length, 4096));
    const match = /<(?:\w+:)?tabColor\s([^>]*)\/?>/.exec(head);
    if (match) {
      const attrs: Attrs = {};
      for (const pair of match[1]!.matchAll(/(\w+)="([^"]*)"/g)) attrs[pair[1]!] = pair[2]!;
      sheet.tabColor = colorFromAttrs(attrs, parsedStyles);
    }
  }

  const meta: WorkbookMeta = {
    format: "xlsx",
    revision,
    sheets: sheets.map(({ part: _part, ...summary }) => summary),
    activeSheet: Math.min(Math.max(0, activeSheet), Math.max(0, sheets.length - 1)),
    styles: parsedStyles.styles,
    warnings,
  };

  return {
    meta,
    sheet(index) {
      const cached = parsedSheets.get(index);
      if (cached) return cached;
      const summary = sheets[index];
      if (!summary) throw new Error(`시트 ${index}가 없습니다.`);
      const xml = summary.part ? read(summary.part) : undefined;
      const raw = parseSheetXml({
        index,
        name: summary.name,
        xml: xml ?? "",
        styles: parsedStyles,
        shared,
        date1904,
        printArea: printAreas.get(index) ?? null,
        titles: printTitles.get(index) ?? { rows: null, cols: null },
      });
      if (summary.part) {
        const sheetRels = rels(summary.part);
        raw.sheet.links = raw.links.map((link) => {
          const rel = link.relId ? sheetRels.get(link.relId) : undefined;
          return { r: link.r, c: link.c, target: rel?.target ?? link.location ?? "", internal: !rel && !!link.location } satisfies CellLink;
        }).filter((link) => link.target);
        if (raw.drawingRel) {
          const rel = sheetRels.get(raw.drawingRel);
          if (rel) raw.sheet.drawings = parseDrawing(rel.target, read(rel.target), rels(rel.target), parsedStyles, imageParts);
        }
        const commentsRel = [...sheetRels.values()].find((rel) => rel.type.endsWith(REL_SUFFIX.comments));
        if (commentsRel) raw.sheet.comments = parseComments(read(commentsRel.target));
      }
      const parsed = buildSheet(raw.sheet, (cell) => finalizeNumber(cell, parsedStyles.formats[cell.s] ?? "General", date1904));
      parsedSheets.set(index, parsed);
      return parsed;
    },
    image(key) {
      const part = imageParts.get(key);
      if (!part) return null;
      const entry = unzipLimited(bytes, (name) => name === part, IMAGE_LIMITS)[part];
      return entry ? ({ bytes: entry, mimeType: mimeFromName(part) } satisfies ImageData) : null;
    },
  };
}

function colorFromAttrs(attrs: Attrs, styles: ParsedStyles): string | null {
  if (attrs.rgb) return `#${attrs.rgb.slice(-6).toUpperCase()}`;
  if (attrs.theme !== undefined) {
    const base = styles.colors.theme[Number(attrs.theme)];
    if (!base) return null;
    return `#${(attrs.tint ? applyTint(base, Number(attrs.tint)) : base).toUpperCase()}`;
  }
  if (attrs.indexed !== undefined) {
    const hex = styles.colors.palette[Number(attrs.indexed)];
    return hex ? `#${hex}` : null;
  }
  return null;
}

function parseSharedStrings(xml: string, styles: ParsedStyles): SharedString[] {
  const out: SharedString[] = [];
  let runs: RichRun[] | null = null;
  let text = "";
  let runText = "";
  let inT = false;
  let phonetic = 0;
  let inRun = false;
  let rPr: XmlNode | null = null;
  let runFontSpec: RichRun["font"] | undefined;
  scanXml(xml, {
    open(name, attrs, selfClosing) {
      if (phonetic > 0) {
        if (!selfClosing) phonetic++;
        return;
      }
      switch (name) {
        case "si":
          runs = null;
          text = "";
          break;
        case "rPh":
          if (!selfClosing) phonetic = 1;
          break;
        case "r":
          inRun = true;
          runText = "";
          runFontSpec = undefined;
          break;
        case "rPr":
          rPr = { name, attrs, children: [], text: "" };
          break;
        case "t":
          if (!selfClosing) inT = true;
          break;
        default:
          if (rPr) rPr.children.push({ name, attrs, children: [], text: "" });
      }
    },
    close(name) {
      if (phonetic > 0) {
        phonetic--;
        return;
      }
      switch (name) {
        case "t":
          inT = false;
          break;
        case "rPr":
          if (rPr) runFontSpec = runFont(rPr, styles.colors);
          rPr = null;
          break;
        case "r":
          inRun = false;
          if (!runs) runs = [];
          runs.push(runFontSpec ? { text: runText, font: runFontSpec } : { text: runText });
          break;
        case "si": {
          const value = decodeOoxmlEscapes(runs ? (runs as RichRun[]).map((run) => run.text).join("") : text);
          const entry: SharedString = { text: value };
          if (runs && (runs as RichRun[]).some((run) => run.font)) {
            entry.rich = (runs as RichRun[]).map((run) => ({ ...run, text: decodeOoxmlEscapes(run.text) }));
          }
          out.push(entry);
          break;
        }
      }
    },
    text(value) {
      if (!inT || phonetic > 0) return;
      if (inRun) runText += value;
      else text += value;
    },
  });
  return out;
}

interface SheetParseInput {
  index: number;
  name: string;
  xml: string;
  styles: ParsedStyles;
  shared: SharedString[];
  date1904: boolean;
  printArea: Range | null;
  titles: { rows: [number, number] | null; cols: [number, number] | null };
}

interface RawLink {
  r: number;
  c: number;
  relId?: string;
  location?: string;
}

const isTrue = (value: string | undefined) => value === "1" || value === "true";

function parseSheetXml(input: SheetParseInput): { sheet: RawSheet; links: RawLink[]; drawingRel: string | null } {
  const { styles, shared, date1904 } = input;
  const formats = styles.formats;
  const view: SheetView = defaultView();
  const setup: PageSetup = { manualRowBreaks: [], manualColBreaks: [], printArea: input.printArea, titleRows: input.titles.rows, titleCols: input.titles.cols };
  const raw: RawSheet = {
    index: input.index,
    name: input.name,
    defaultFont: { name: styles.defaultFont.name, size: styles.defaultFont.size },
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
  const links: RawLink[] = [];
  let drawingRel: string | null = null;
  let sawView = false;
  let breaks: "row" | "col" | null = null;

  // Current cell state.
  let rowIndex = -1;
  let colCursor = 0;
  let cell: { r: number; c: number; s: number; t: string } | null = null;
  let value = "";
  let formula = "";
  let inV = false;
  let inF = false;
  let inIs = false;
  let inT = false;
  let isText = "";
  let phonetic = 0;
  let extDepth = 0;
  let colStyled: boolean | null = null;

  const finishCell = () => {
    if (!cell) return;
    const { r, c, s, t } = cell;
    const format = formats[s] ?? "General";
    const out: Cell = { r, c, s, v: "", k: "z" };
    const hasValue = value !== "" || t === "inlineStr" || t === "str";
    if (t === "s") {
      const entry = shared[Number(value)];
      if (entry) {
        out.v = entry.text;
        out.k = "s";
        if (entry.rich) out.rich = entry.rich;
        if (format.includes("@") && format !== "@") out.v = formatValue(format, entry.text, date1904);
      }
    } else if (t === "inlineStr") {
      out.v = decodeOoxmlEscapes(isText);
      out.k = "s";
    } else if (t === "str") {
      out.v = decodeOoxmlEscapes(value);
      out.k = "s";
    } else if (t === "b") {
      out.v = value === "1" || value === "true" ? "TRUE" : "FALSE";
      out.k = "b";
    } else if (t === "e") {
      out.v = value;
      out.k = "e";
    } else if (hasValue) {
      const number = t === "d" ? isoToSerial(value, date1904) : Number(value);
      if (Number.isFinite(number)) {
        // Formatting waits for the chunk request: most cells of a big sheet are never displayed.
        out.n = number;
        out.k = isDateFormatCached(format) ? "d" : "n";
      } else {
        out.v = value;
        out.k = "s";
      }
    }
    if (formula) out.f = formula;
    // A blank cell with style 0 still matters when it overrides a row or column style.
    if (colStyled === null) colStyled = raw.cols.some((col) => col.style !== null);
    if (out.k !== "z" || out.f || s !== 0 || colStyled || (raw.rows.get(r)?.style ?? null) !== null) raw.cells.push(out);
    cell = null;
  };

  scanXml(input.xml, {
    open(name, attrs, selfClosing) {
      if (extDepth > 0) {
        if (!selfClosing) extDepth++;
        return;
      }
      if (phonetic > 0) {
        if (!selfClosing) phonetic++;
        return;
      }
      switch (name) {
        case "extLst":
          if (!selfClosing) extDepth = 1;
          return;
        case "pageSetUpPr":
          if (isTrue(attrs.fitToPage)) setup.fitToPage = true;
          return;
        case "sheetView":
          if (sawView) return;
          sawView = true;
          if (attrs.view === "pageBreakPreview" || attrs.view === "pageLayout") view.mode = attrs.view;
          if (attrs.showGridLines === "0" || attrs.showGridLines === "false") view.showGrid = false;
          if (attrs.showRowColHeaders === "0" || attrs.showRowColHeaders === "false") view.showHeaders = false;
          if (isTrue(attrs.rightToLeft)) view.rtl = true;
          if (attrs.zoomScale) view.zoom = Number(attrs.zoomScale) || 100;
          if (attrs.topLeftCell) {
            const ref = parseCellRef(attrs.topLeftCell);
            if (ref && ref.r >= 0 && ref.c >= 0) {
              view.topRow = ref.r;
              view.leftCol = ref.c;
            }
          }
          return;
        case "pane":
          if (attrs.state === "frozen" || attrs.state === "frozenSplit") {
            view.frozenCols = Number(attrs.xSplit ?? 0) | 0;
            view.frozenRows = Number(attrs.ySplit ?? 0) | 0;
          }
          return;
        case "sheetFormatPr":
          if (attrs.defaultRowHeight) raw.defaultRowHeight = Number(attrs.defaultRowHeight);
          if (attrs.defaultColWidth) raw.defaultColWidth = Number(attrs.defaultColWidth);
          if (attrs.baseColWidth) raw.baseColWidth = Number(attrs.baseColWidth);
          if (isTrue(attrs.zeroHeight)) raw.zeroHeight = true;
          return;
        case "col":
          raw.cols.push({
            min: Number(attrs.min ?? 1) - 1,
            max: Number(attrs.max ?? attrs.min ?? 1) - 1,
            width: attrs.width !== undefined ? Number(attrs.width) : undefined,
            hidden: isTrue(attrs.hidden),
            style: attrs.style !== undefined && attrs.style !== "0" ? Number(attrs.style) : null,
          });
          return;
        case "row": {
          rowIndex = attrs.r ? Number(attrs.r) - 1 : rowIndex + 1;
          colCursor = 0;
          const style = isTrue(attrs.customFormat) && attrs.s !== undefined ? Number(attrs.s) : null;
          const hidden = isTrue(attrs.hidden);
          const ht = attrs.ht !== undefined ? Number(attrs.ht) : undefined;
          if (ht !== undefined || hidden || style !== null || raw.zeroHeight) raw.rows.set(rowIndex, { ht, hidden, style });
          return;
        }
        case "c": {
          let r = rowIndex;
          let c = colCursor;
          if (attrs.r) {
            const ref = parseCellRef(attrs.r);
            if (ref) {
              if (ref.r >= 0) r = ref.r;
              c = ref.c;
            }
          }
          colCursor = c + 1;
          cell = { r, c, s: attrs.s ? Number(attrs.s) : 0, t: attrs.t ?? "n" };
          value = "";
          formula = "";
          isText = "";
          if (selfClosing) finishCell();
          return;
        }
        case "v":
          if (!selfClosing) inV = true;
          return;
        case "f":
          if (!selfClosing) inF = true;
          return;
        case "is":
          inIs = true;
          return;
        case "rPh":
          if (!selfClosing) phonetic = 1;
          return;
        case "t":
          if (inIs && !selfClosing) inT = true;
          return;
        case "mergeCell": {
          const range = attrs.ref ? parseRangeRef(attrs.ref) : null;
          if (range && (range[0] !== range[2] || range[1] !== range[3])) raw.merges.push(range);
          return;
        }
        case "hyperlink": {
          const range = attrs.ref ? parseRangeRef(attrs.ref) : null;
          if (range) links.push({ r: range[0], c: range[1], relId: attrs.id, location: attrs.location });
          return;
        }
        case "pageMargins":
          setup.margins = {
            left: Number(attrs.left ?? 0.7),
            right: Number(attrs.right ?? 0.7),
            top: Number(attrs.top ?? 0.75),
            bottom: Number(attrs.bottom ?? 0.75),
          };
          return;
        case "pageSetup":
          if (attrs.paperSize) setup.paperSize = Number(attrs.paperSize);
          if (attrs.scale) setup.scale = Number(attrs.scale);
          if (attrs.fitToWidth !== undefined) setup.fitToWidth = Number(attrs.fitToWidth);
          if (attrs.fitToHeight !== undefined) setup.fitToHeight = Number(attrs.fitToHeight);
          if (attrs.orientation === "landscape") setup.orientation = "landscape";
          if (attrs.pageOrder === "overThenDown") setup.order = "over";
          return;
        case "rowBreaks":
          breaks = "row";
          return;
        case "colBreaks":
          breaks = "col";
          return;
        case "brk":
          if (breaks && attrs.id !== undefined && (attrs.man === undefined || isTrue(attrs.man))) {
            (breaks === "row" ? setup.manualRowBreaks : setup.manualColBreaks).push(Number(attrs.id));
          }
          return;
        case "drawing":
          drawingRel = attrs.id ?? null;
          return;
      }
    },
    close(name) {
      if (extDepth > 0) {
        extDepth--;
        return;
      }
      if (phonetic > 0) {
        phonetic--;
        return;
      }
      switch (name) {
        case "v":
          inV = false;
          return;
        case "f":
          inF = false;
          return;
        case "t":
          inT = false;
          return;
        case "is":
          inIs = false;
          return;
        case "c":
          finishCell();
          return;
        case "rowBreaks":
        case "colBreaks":
          breaks = null;
          return;
      }
    },
    text(text) {
      if (phonetic > 0 || extDepth > 0) return;
      if (inV) value += text;
      else if (inF) formula += text;
      else if (inT) isText += text;
    },
  });
  return { sheet: raw, links, drawingRel };
}

function isoToSerial(value: string, date1904: boolean): number {
  const time = Date.parse(value.endsWith("Z") || /[+-]\d\d:\d\d$/.test(value) ? value : `${value}Z`);
  if (!Number.isFinite(time)) return Number.NaN;
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  return (time - epoch) / 86400000;
}

// ---------------------------------------------------------------------------------------------
// Drawings

const SCHEME_INDEX: Record<string, number> = {
  bg1: 0, lt1: 0, tx1: 1, dk1: 1, bg2: 2, lt2: 2, tx2: 3, dk2: 3,
  accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11,
};

const PRESET_COLORS: Record<string, string> = {
  black: "000000", white: "FFFFFF", red: "FF0000", green: "008000", blue: "0000FF", yellow: "FFFF00",
  gray: "808080", grey: "808080", darkGray: "A9A9A9", lightGray: "D3D3D3", orange: "FFA500",
};

function drawingColor(node: XmlNode | undefined, styles: ParsedStyles): string | null {
  if (!node) return null;
  const color = node.children[0];
  if (!color) return null;
  let hex: string | undefined;
  if (color.name === "srgbClr") hex = color.attrs.val;
  else if (color.name === "schemeClr") hex = styles.colors.theme[SCHEME_INDEX[color.attrs.val ?? ""] ?? 1];
  else if (color.name === "sysClr") hex = color.attrs.lastClr ?? "000000";
  else if (color.name === "prstClr") hex = PRESET_COLORS[color.attrs.val ?? ""];
  if (!hex) return null;
  const lumMod = child(color, "lumMod")?.attrs.val;
  const lumOff = child(color, "lumOff")?.attrs.val;
  const shade = child(color, "shade")?.attrs.val;
  const tint = child(color, "tint")?.attrs.val;
  if (lumMod || lumOff) {
    const mod = lumMod ? Number(lumMod) / 100000 : 1;
    const off = lumOff ? Number(lumOff) / 100000 : 0;
    // lumMod/lumOff scale and shift luminance; approximate with the tint helper.
    hex = applyTint(hex, off > 0 ? off : mod - 1);
  } else if (shade) {
    hex = applyTint(hex, Number(shade) / 100000 - 1);
  } else if (tint) {
    hex = applyTint(hex, 1 - Number(tint) / 100000);
  }
  return `#${hex.toUpperCase()}`;
}

function anchorPoint(node: XmlNode | undefined) {
  return {
    c: Number(textOf(child(node, "col")) || 0),
    dx: emuToPx(Number(textOf(child(node, "colOff")) || 0)),
    r: Number(textOf(child(node, "row")) || 0),
    dy: emuToPx(Number(textOf(child(node, "rowOff")) || 0)),
  };
}

function collectAnchors(node: XmlNode, out: XmlNode[]) {
  for (const entry of node.children) {
    if (entry.name === "mc:Fallback") continue;
    if (entry.name === "twoCellAnchor" || entry.name === "oneCellAnchor" || entry.name === "absoluteAnchor") out.push(entry);
    else collectAnchors(entry, out);
  }
}

function parseDrawing(part: string, xml: string | undefined, rels: Map<string, Rel>, styles: ParsedStyles, images: Map<string, string>): DrawingItem[] {
  if (!xml) return [];
  const tree = parseTree(xml);
  const anchors: XmlNode[] = [];
  collectAnchors(tree, anchors);
  const items: DrawingItem[] = [];
  anchors.forEach((anchorNode, index) => {
    const anchor: DrawingItem["anchor"] = { from: anchorPoint(child(anchorNode, "from")) };
    if (anchorNode.name === "twoCellAnchor") anchor.to = anchorPoint(child(anchorNode, "to"));
    const ext = child(anchorNode, "ext");
    if (ext) anchor.ext = { w: emuToPx(Number(ext.attrs.cx ?? 0)), h: emuToPx(Number(ext.attrs.cy ?? 0)) };
    const pos = child(anchorNode, "pos");
    if (pos) anchor.pos = { x: emuToPx(Number(pos.attrs.x ?? 0)), y: emuToPx(Number(pos.attrs.y ?? 0)) };
    const id = `${part}#${index}`;

    const pic = descendants(anchorNode, "pic")[0];
    if (pic) {
      const blip = descendants(pic, "blip")[0];
      const rel = blip?.attrs.embed ? rels.get(blip.attrs.embed) : undefined;
      if (rel && !rel.external) {
        images.set(rel.target, rel.target);
        items.push({ id, kind: "image", anchor, image: rel.target, name: descendants(pic, "cNvPr")[0]?.attrs.name });
      }
      return;
    }
    const frame = descendants(anchorNode, "graphicFrame")[0];
    if (frame) {
      items.push({ id, kind: "chart", anchor, name: descendants(frame, "cNvPr")[0]?.attrs.name ?? "차트" });
      return;
    }
    const shape = descendants(anchorNode, "sp")[0];
    if (shape) {
      const spPr = child(shape, "spPr");
      const style = child(shape, "style");
      const fill = child(spPr, "noFill") ? null : drawingColor(child(spPr, "solidFill"), styles) ?? drawingColor(child(style, "fillRef"), styles);
      const ln = child(spPr, "ln");
      const line = child(ln, "noFill") ? null : drawingColor(child(ln, "solidFill"), styles) ?? drawingColor(child(style, "lnRef"), styles);
      const paragraphs = descendants(child(shape, "txBody"), "p").map((p) => descendants(p, "t").map(textOf).join(""));
      items.push({ id, kind: "shape", anchor, fill, line, text: paragraphs.join("\n").trim() || undefined });
    }
  });
  return items;
}

function parseComments(xml: string | undefined): CellComment[] {
  if (!xml) return [];
  const tree = parseTree(xml);
  const root = child(tree, "comments");
  const authors = children(child(root, "authors"), "author").map(textOf);
  return children(child(root, "commentList"), "comment").flatMap((node) => {
    const ref = node.attrs.ref ? parseCellRef(node.attrs.ref) : null;
    if (!ref || ref.r < 0) return [];
    const textNode = child(node, "text");
    const parts = textNode ? textNode.children.filter((entry) => entry.name !== "rPh").map((entry) => (entry.name === "t" ? textOf(entry) : descendants(entry, "t").map(textOf).join(""))) : [];
    return [{ r: ref.r, c: ref.c, author: authors[Number(node.attrs.authorId ?? 0)] ?? "", text: decodeOoxmlEscapes(parts.join("")).trim() }];
  });
}

export type { FontDef, CellStyle };
