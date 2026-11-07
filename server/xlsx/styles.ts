import type { BorderLine, BorderSide, CellStyle, HAlign, RichRun, VAlign } from "../../shared/model";
import { type ColorContext, type ColorSpec, DEFAULT_PALETTE, DEFAULT_THEME, mix, PATTERN_DENSITY, resolveColor } from "../colors";
import { builtinFormat } from "../format";
import { type Attrs, child, children, descendants, parseTree, type XmlNode } from "../xml";

export interface FontDef {
  name: string;
  size: number;
  bold?: boolean;
  italic?: boolean;
  underline?: "single" | "double";
  strike?: boolean;
  color: string;
  vert?: "superscript" | "subscript";
}

export interface ParsedStyles {
  styles: CellStyle[];
  /** Number format code per cellXfs index. */
  formats: string[];
  defaultFont: FontDef;
  colors: ColorContext;
}

export function parseTheme(xml: string | undefined): string[] {
  if (!xml) return [...DEFAULT_THEME];
  const tree = parseTree(xml);
  const scheme = descendants(tree, "clrScheme")[0];
  if (!scheme) return [...DEFAULT_THEME];
  const pick = (name: string, fallback: string) => {
    const node = child(scheme, name);
    const srgb = child(node, "srgbClr");
    if (srgb?.attrs.val) return srgb.attrs.val;
    const sys = child(node, "sysClr");
    return sys?.attrs.lastClr ?? fallback;
  };
  // Excel's theme index order swaps the light and dark pairs.
  return [
    pick("lt1", "FFFFFF"), pick("dk1", "000000"), pick("lt2", "EEECE1"), pick("dk2", "1F497D"),
    pick("accent1", "4F81BD"), pick("accent2", "C0504D"), pick("accent3", "9BBB59"), pick("accent4", "8064A2"),
    pick("accent5", "4BACC6"), pick("accent6", "F79646"), pick("hlink", "0000FF"), pick("folHlink", "800080"),
  ];
}

function colorSpec(node: XmlNode | undefined): ColorSpec | undefined {
  if (!node) return undefined;
  return node.attrs as Attrs as ColorSpec;
}

const flag = (node: XmlNode | undefined) => node !== undefined && node.attrs.val !== "0" && node.attrs.val !== "false";

export function parseFont(node: XmlNode | undefined, colors: ColorContext, base?: FontDef): FontDef {
  const font: FontDef = base ? { ...base } : { name: "Calibri", size: 11, color: "#000000" };
  if (!node) return font;
  const name = child(node, "name") ?? child(node, "rFont");
  if (name?.attrs.val) font.name = name.attrs.val;
  const size = child(node, "sz");
  if (size?.attrs.val) font.size = Number(size.attrs.val) || font.size;
  if (child(node, "b")) font.bold = flag(child(node, "b"));
  if (child(node, "i")) font.italic = flag(child(node, "i"));
  if (child(node, "strike")) font.strike = flag(child(node, "strike"));
  const underline = child(node, "u");
  if (underline) {
    const value = underline.attrs.val ?? "single";
    font.underline = value === "none" ? undefined : value.startsWith("double") ? "double" : "single";
  }
  const vert = child(node, "vertAlign")?.attrs.val;
  if (vert === "superscript" || vert === "subscript") font.vert = vert;
  const color = child(node, "color");
  if (color) font.color = resolveColor(colorSpec(color), colors, "#000000") ?? "#000000";
  return font;
}

function parseFill(node: XmlNode | undefined, colors: ColorContext): string | null {
  if (!node) return null;
  const pattern = child(node, "patternFill");
  if (pattern) {
    const type = pattern.attrs.patternType ?? (child(pattern, "fgColor") ? "solid" : "none");
    if (type === "none") return null;
    const fg = resolveColor(colorSpec(child(pattern, "fgColor")), colors, "#000000") ?? "#000000";
    if (type === "solid") return fg;
    const bg = resolveColor(colorSpec(child(pattern, "bgColor")), colors, "#FFFFFF") ?? "#FFFFFF";
    return mix(fg, bg, PATTERN_DENSITY[type] ?? 0.5);
  }
  const gradient = child(node, "gradientFill");
  if (gradient) {
    const stops = children(gradient, "stop").map((stop) => resolveColor(colorSpec(child(stop, "color")), colors, null)).filter(Boolean) as string[];
    if (stops.length === 0) return null;
    return stops.length === 1 ? stops[0]! : mix(stops[0]!, stops[stops.length - 1]!, 0.5);
  }
  return null;
}

const BORDER_STYLES: Record<string, { w: number; line: BorderLine }> = {
  thin: { w: 1, line: "solid" },
  hair: { w: 1, line: "dotted" },
  dotted: { w: 1, line: "dotted" },
  dashed: { w: 1, line: "dashed" },
  dashDot: { w: 1, line: "dashed" },
  dashDotDot: { w: 1, line: "dashed" },
  medium: { w: 2, line: "solid" },
  mediumDashed: { w: 2, line: "dashed" },
  mediumDashDot: { w: 2, line: "dashed" },
  mediumDashDotDot: { w: 2, line: "dashed" },
  slantDashDot: { w: 2, line: "dashed" },
  thick: { w: 3, line: "solid" },
  double: { w: 3, line: "double" },
};

export function borderSide(style: string | undefined, color: string | null): BorderSide | undefined {
  if (!style || style === "none") return undefined;
  const spec = BORDER_STYLES[style] ?? BORDER_STYLES.thin!;
  return { w: spec.w, line: spec.line, color: color ?? "#000000" };
}

function parseBorder(node: XmlNode | undefined, colors: ColorContext): CellStyle["border"] {
  const border: CellStyle["border"] = {};
  if (!node) return border;
  const side = (name: string) => {
    const element = child(node, name);
    if (!element) return undefined;
    return borderSide(element.attrs.style, resolveColor(colorSpec(child(element, "color")), colors, "#000000"));
  };
  border.l = side("left") ?? side("start");
  border.r = side("right") ?? side("end");
  border.t = side("top");
  border.b = side("bottom");
  const diagonal = side("diagonal");
  if (diagonal) {
    if (node.attrs.diagonalUp === "1" || node.attrs.diagonalUp === "true") border.up = diagonal;
    if (node.attrs.diagonalDown === "1" || node.attrs.diagonalDown === "true") border.down = diagonal;
  }
  for (const key of Object.keys(border) as (keyof CellStyle["border"])[]) if (!border[key]) delete border[key];
  return border;
}

const H_ALIGN = new Set<HAlign>(["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"]);
const V_ALIGN = new Set<VAlign>(["top", "center", "bottom", "justify", "distributed"]);

export function parseStyles(xml: string | undefined, themeXml: string | undefined): ParsedStyles {
  const theme = parseTheme(themeXml);
  const tree = xml ? parseTree(xml) : undefined;
  const sheet = child(tree, "styleSheet");
  const customPalette = descendants(child(sheet, "colors"), "rgbColor").map((node) => (node.attrs.rgb ?? "FF000000").slice(-6));
  const palette = customPalette.length > 0
    ? [...DEFAULT_PALETTE.slice(0, 8), ...customPalette, ...DEFAULT_PALETTE.slice(8 + customPalette.length)]
    : DEFAULT_PALETTE;
  const colors: ColorContext = { palette, theme };

  const numFmts = new Map<number, string>();
  for (const node of children(child(sheet, "numFmts"), "numFmt")) numFmts.set(Number(node.attrs.numFmtId), node.attrs.formatCode ?? "General");

  const fontNodes = children(child(sheet, "fonts"), "font");
  const fonts = fontNodes.map((node) => parseFont(node, colors));
  const defaultFont = fonts[0] ?? { name: "Calibri", size: 11, color: "#000000" };
  const fills = children(child(sheet, "fills"), "fill").map((node) => parseFill(node, colors));
  const borders = children(child(sheet, "borders"), "border").map((node) => parseBorder(node, colors));

  const styles: CellStyle[] = [];
  const formats: string[] = [];
  const xfs = children(child(sheet, "cellXfs"), "xf");
  for (const xf of xfs.length > 0 ? xfs : [{ name: "xf", attrs: {}, children: [], text: "" } as XmlNode]) {
    const font = fonts[Number(xf.attrs.fontId ?? 0)] ?? defaultFont;
    const alignment = child(xf, "alignment")?.attrs ?? {};
    const h = (alignment.horizontal ?? "general") as HAlign;
    const v = (alignment.vertical ?? "bottom") as VAlign;
    const style: CellStyle = {
      font: { ...font },
      fill: fills[Number(xf.attrs.fillId ?? 0)] ?? null,
      border: borders[Number(xf.attrs.borderId ?? 0)] ?? {},
      h: H_ALIGN.has(h) ? h : "general",
      v: V_ALIGN.has(v) ? v : "bottom",
    };
    if (alignment.wrapText === "1" || alignment.wrapText === "true") style.wrap = true;
    if (alignment.shrinkToFit === "1" || alignment.shrinkToFit === "true") style.shrink = true;
    if (alignment.indent) style.indent = Number(alignment.indent);
    if (alignment.textRotation) style.rotation = Number(alignment.textRotation);
    styles.push(style);
    const id = Number(xf.attrs.numFmtId ?? 0);
    formats.push(numFmts.get(id) ?? builtinFormat(id));
  }
  return { styles, formats, defaultFont, colors };
}

/** Converts a run's font to the fields that differ from the cell's base font. */
export function runFont(rPr: XmlNode | undefined, colors: ColorContext): RichRun["font"] | undefined {
  if (!rPr) return undefined;
  const font: RichRun["font"] = {};
  const name = child(rPr, "rFont") ?? child(rPr, "name");
  if (name?.attrs.val) font.name = name.attrs.val;
  const size = child(rPr, "sz");
  if (size?.attrs.val) font.size = Number(size.attrs.val);
  if (child(rPr, "b")) font.bold = flag(child(rPr, "b"));
  if (child(rPr, "i")) font.italic = flag(child(rPr, "i"));
  if (child(rPr, "strike")) font.strike = flag(child(rPr, "strike"));
  const underline = child(rPr, "u");
  if (underline) font.underline = (underline.attrs.val ?? "single") === "none" ? undefined : (underline.attrs.val ?? "").startsWith("double") ? "double" : "single";
  const vert = child(rPr, "vertAlign")?.attrs.val;
  if (vert === "superscript" || vert === "subscript") font.vert = vert;
  const color = child(rPr, "color");
  if (color) {
    const resolved = resolveColor(colorSpec(color), colors, null);
    if (resolved) font.color = resolved;
  }
  return font;
}
