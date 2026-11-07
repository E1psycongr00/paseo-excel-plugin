import { Platform, type TextStyle } from "react-native";
import type { CellStyle, RichRun } from "../shared/model";

// Korean and Latin fallbacks so web shows the workbook's fonts when the machine has them.
const WEB_FALLBACK = "\"Malgun Gothic\", \"맑은 고딕\", \"Apple SD Gothic Neo\", \"Noto Sans KR\", sans-serif";

export function fontFamily(name: string | undefined): string | undefined {
  if (Platform.OS !== "web" || !name) return undefined;
  return `"${name.replace(/"/g, "")}", ${WEB_FALLBACK}`;
}

export const ptToPx = (pt: number) => (pt * 96) / 72;

export function textStyle(font: Partial<CellStyle["font"]>, zoom: number, colorOverride?: string): TextStyle {
  const size = ptToPx(font.size ?? 11) * zoom * (font.vert ? 0.7 : 1);
  const decorations: string[] = [];
  if (font.underline) decorations.push("underline");
  if (font.strike) decorations.push("line-through");
  const style: TextStyle = {
    fontSize: size,
    lineHeight: Math.ceil(size * 1.22),
    color: colorOverride ?? font.color ?? "#000000",
  };
  const family = fontFamily(font.name);
  if (family) style.fontFamily = family;
  if (font.bold) style.fontWeight = "700";
  if (font.italic) style.fontStyle = "italic";
  if (decorations.length > 0) style.textDecorationLine = decorations.join(" ") as TextStyle["textDecorationLine"];
  if (font.underline === "double") style.textDecorationStyle = "double";
  return style;
}

/** Rich run styles relative to the cell font. */
export function runStyle(run: RichRun, base: CellStyle["font"], zoom: number): TextStyle {
  return textStyle({ ...base, ...run.font }, zoom);
}

export class StyleCache {
  private readonly cache = new Map<string, TextStyle>();
  constructor(private readonly styles: CellStyle[], private readonly zoom: number) {}

  text(id: number, color?: string): TextStyle {
    const key = color ? `${id}|${color}` : `${id}`;
    let style = this.cache.get(key);
    if (!style) {
      const cell = this.styles[id] ?? this.styles[0];
      style = textStyle(cell?.font ?? { name: "Calibri", size: 11, color: "#000000" }, this.zoom, color);
      this.cache.set(key, style);
    }
    return style;
  }

  get(id: number): CellStyle | undefined {
    return this.styles[id];
  }
}
