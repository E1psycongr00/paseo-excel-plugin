// Excel color resolution: indexed palette, theme colors and tint.

export const DEFAULT_PALETTE: readonly string[] = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

/** Office 2007-2013 default theme, used when a file has no theme part. */
export const DEFAULT_THEME: readonly string[] = [
  "FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646", "0000FF", "800080",
];

export interface ColorContext {
  palette: readonly string[];
  theme: readonly string[];
}

export interface ColorSpec {
  rgb?: string;
  theme?: string;
  indexed?: string;
  tint?: string;
  auto?: string;
}

/**
 * Resolves a color element. `fallback` covers `auto` and missing colors (black text, no fill).
 * Returns "#RRGGBB".
 */
export function resolveColor(spec: ColorSpec | undefined, context: ColorContext, fallback: string | null): string | null {
  if (!spec) return fallback;
  let hex: string | null = null;
  if (spec.rgb) {
    const rgb = spec.rgb.replace(/^#/, "");
    hex = rgb.length === 8 ? rgb.slice(2) : rgb.length === 6 ? rgb : null;
  } else if (spec.theme !== undefined) {
    hex = context.theme[Number(spec.theme)] ?? null;
  } else if (spec.indexed !== undefined) {
    const index = Number(spec.indexed);
    // 64 is the system foreground, 65 the system background.
    if (index === 64) return fallback ?? "#000000";
    if (index === 65) return fallback ?? "#FFFFFF";
    hex = context.palette[index] ?? null;
  } else if (spec.auto) {
    return fallback;
  }
  if (!hex) return fallback;
  const tint = spec.tint ? Number(spec.tint) : 0;
  return `#${(tint ? applyTint(hex, tint) : hex).toUpperCase()}`;
}

export function applyTint(hex: string, tint: number): string {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  let [h, s, l] = rgbToHsl(r, g, b);
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const [nr, ng, nb] = hslToRgb(h, s, Math.min(1, Math.max(0, l)));
  return [nr, ng, nb].map((value) => Math.round(value * 255).toString(16).padStart(2, "0")).join("");
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  if (s === 0) return [l, l, l];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [hue(h + 1 / 3), hue(h), hue(h - 1 / 3)];
}

/** Mixes two "#RRGGBB" colors; `amount` is the share of `a`. */
export function mix(a: string, b: string, amount: number): string {
  const channel = (hex: string, at: number) => parseInt(hex.slice(at, at + 2), 16);
  const out = [1, 3, 5].map((at) => Math.round(channel(a, at) * amount + channel(b, at) * (1 - amount)));
  return `#${out.map((value) => value.toString(16).padStart(2, "0")).join("").toUpperCase()}`;
}

/** Pattern fills approximated by the share of foreground pixels. */
export const PATTERN_DENSITY: Record<string, number> = {
  solid: 1,
  darkGray: 0.75,
  mediumGray: 0.5,
  lightGray: 0.25,
  gray125: 0.125,
  gray0625: 0.0625,
  darkHorizontal: 0.5,
  darkVertical: 0.5,
  darkDown: 0.5,
  darkUp: 0.5,
  darkGrid: 0.5,
  darkTrellis: 0.75,
  lightHorizontal: 0.25,
  lightVertical: 0.25,
  lightDown: 0.25,
  lightUp: 0.25,
  lightGrid: 0.4,
  lightTrellis: 0.375,
};
