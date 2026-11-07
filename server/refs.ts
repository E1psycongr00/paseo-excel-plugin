import type { Range } from "../shared/model";

/** "AB12" → { r: 11, c: 27 }. Absolute markers are ignored. */
export function parseCellRef(ref: string): { r: number; c: number } | null {
  let c = 0;
  let i = 0;
  const text = ref.replace(/\$/g, "");
  while (i < text.length) {
    const code = text.charCodeAt(i);
    const upper = code >= 97 && code <= 122 ? code - 32 : code;
    if (upper < 65 || upper > 90) break;
    c = c * 26 + (upper - 64);
    i++;
  }
  const row = Number(text.slice(i));
  if (i === 0 && !Number.isInteger(row)) return null;
  return { r: i === text.length ? -1 : row - 1, c: c - 1 };
}

/** "A1:C3", "A1", "A:C" or "1:3" → inclusive 0-based range. */
export function parseRangeRef(ref: string, maxRow = 1048575, maxCol = 16383): Range | null {
  const [startText, endText = startText] = ref.replace(/\$/g, "").split(":");
  if (!startText) return null;
  const point = (text: string, end: boolean): { r: number; c: number } | null => {
    if (/^\d+$/.test(text)) return { r: Number(text) - 1, c: end ? maxCol : 0 };
    if (/^[A-Za-z]+$/.test(text)) return { r: end ? maxRow : 0, c: (parseCellRef(`${text}1`)?.c ?? 0) };
    return parseCellRef(text);
  };
  const a = point(startText, false);
  const b = point(endText!, true);
  if (!a || !b || a.r < 0 || b.r < 0 || a.c < 0 || b.c < 0) return null;
  return [Math.min(a.r, b.r), Math.min(a.c, b.c), Math.max(a.r, b.r), Math.max(a.c, b.c)];
}

/** Splits a defined-name formula such as "'My Sheet'!$A$1:$K$40,'My Sheet'!$M$1:$N$2". */
export function parseAreaList(formula: string): { sheet: string | null; range: Range }[] {
  const out: { sheet: string | null; range: Range }[] = [];
  const pattern = /(?:(?:'((?:[^']|'')+)'|([^'!,]+))!)?(\$?[A-Za-z]{0,3}\$?\d*(?::\$?[A-Za-z]{0,3}\$?\d*)?)/g;
  for (const part of splitTopLevel(formula)) {
    pattern.lastIndex = 0;
    const match = pattern.exec(part.trim());
    if (!match || !match[3]) continue;
    const range = parseRangeRef(match[3]);
    if (!range) continue;
    const sheet = match[1] !== undefined ? match[1].replace(/''/g, "'") : match[2] ?? null;
    out.push({ sheet, range });
  }
  return out;
}

function splitTopLevel(formula: string): string[] {
  const parts: string[] = [];
  let current = "";
  let quoted = false;
  for (const char of formula) {
    if (char === "'") quoted = !quoted;
    if (char === "," && !quoted) {
      parts.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  parts.push(current);
  return parts;
}

export function columnName(c: number): string {
  let name = "";
  let n = c + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

export const cellName = (r: number, c: number) => `${columnName(c)}${r + 1}`;
