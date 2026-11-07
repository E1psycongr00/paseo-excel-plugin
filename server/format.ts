import SSF from "ssf";

// Built-in formats that depend on the Excel locale. The viewer follows Korean Excel defaults, which
// also read naturally elsewhere (ISO-like dates).
const LOCALE_BUILTINS: Record<number, string> = {
  5: "\"₩\"#,##0;\"₩\"\-#,##0",
  6: "\"₩\"#,##0;[Red]\"₩\"\-#,##0",
  7: "\"₩\"#,##0.00;\"₩\"\-#,##0.00",
  8: "\"₩\"#,##0.00;[Red]\"₩\"\-#,##0.00",
  14: "yyyy-mm-dd",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  22: "yyyy-mm-dd h:mm",
  27: "yyyy\"년\" m\"월\" d\"일\"",
  28: "m\"월\" d\"일\"",
  29: "m\"월\" d\"일\"",
  30: "mm-dd-yy",
  31: "yyyy\"년\" m\"월\" d\"일\"",
  32: "h\"시\" mm\"분\"",
  33: "h\"시\" mm\"분\" ss\"초\"",
  34: "yyyy-mm-dd",
  35: "yyyy-mm-dd",
  36: "yyyy\"년\" m\"월\" d\"일\"",
  37: "#,##0;-#,##0",
  38: "#,##0;[Red]-#,##0",
  39: "#,##0.00;-#,##0.00",
  40: "#,##0.00;[Red]-#,##0.00",
  41: "_-* #,##0_-;\-* #,##0_-;_-* \"-\"_-;_-@_-",
  42: "_-\"₩\"* #,##0_-;\-\"₩\"* #,##0_-;_-\"₩\"* \"-\"_-;_-@_-",
  43: "_-* #,##0.00_-;\-* #,##0.00_-;_-* \"-\"??_-;_-@_-",
  44: "_-\"₩\"* #,##0.00_-;\-\"₩\"* #,##0.00_-;_-\"₩\"* \"-\"??_-;_-@_-",
  50: "yyyy\"년\" m\"월\" d\"일\"",
  51: "m\"월\" d\"일\"",
  52: "yyyy-mm-dd",
  53: "yyyy-mm-dd",
  54: "m\"월\" d\"일\"",
  55: "yyyy-mm-dd",
  56: "yyyy-mm-dd",
  57: "yyyy\"년\" m\"월\" d\"일\"",
  58: "m\"월\" d\"일\"",
};

const table = SSF.get_table() as Record<number, string>;

export function builtinFormat(id: number): string {
  return LOCALE_BUILTINS[id] ?? table[id] ?? "General";
}

const NAMED_COLORS: Record<string, string> = {
  black: "#000000",
  white: "#FFFFFF",
  red: "#FF0000",
  green: "#00FF00",
  blue: "#0000FF",
  yellow: "#FFFF00",
  magenta: "#FF00FF",
  cyan: "#00FFFF",
};

// Split format sections on ";" outside quotes, brackets and escapes.
function sections(format: string): string[] {
  const out: string[] = [];
  let current = "";
  let quoted = false;
  let bracket = false;
  for (let i = 0; i < format.length; i++) {
    const char = format[i]!;
    if (char === "\\" && !quoted) {
      current += char + (format[i + 1] ?? "");
      i++;
      continue;
    }
    if (char === "\"") quoted = !quoted;
    else if (!quoted && char === "[") bracket = true;
    else if (!quoted && char === "]") bracket = false;
    if (char === ";" && !quoted && !bracket) {
      out.push(current);
      current = "";
      continue;
    }
    current += char;
  }
  out.push(current);
  return out;
}

/** The color a format section forces, e.g. [Red]. */
export function formatColor(format: string, value: number | string): string | undefined {
  if (format.indexOf("[") < 0) return undefined;
  const parts = sections(format);
  let section: string;
  if (typeof value === "string") section = parts[3] ?? (parts.length === 1 ? parts[0]! : "");
  else if (parts.some((part) => /\[[<>=]/.test(part))) section = conditionalSection(parts, value);
  else if (value > 0 || parts.length === 1) section = parts[0]!;
  else if (value < 0) section = parts[1] ?? parts[0]!;
  else section = parts[2] ?? parts[0]!;
  const match = /\[(Black|White|Red|Green|Blue|Yellow|Magenta|Cyan|Color\s*(\d+))\]/i.exec(section);
  if (!match) return undefined;
  if (match[2]) return undefined;
  return NAMED_COLORS[match[1]!.toLowerCase()];
}

function conditionalSection(parts: string[], value: number): string {
  for (const part of parts) {
    const condition = /\[(<=|>=|<>|<|>|=)(-?[\d.]+)\]/.exec(part);
    if (!condition) return part;
    const limit = Number(condition[2]);
    const op = condition[1];
    const pass = op === "<" ? value < limit : op === ">" ? value > limit : op === "<=" ? value <= limit
      : op === ">=" ? value >= limit : op === "=" ? value === limit : value !== limit;
    if (pass) return part;
  }
  return parts[parts.length - 1]!;
}

/** Formats a cell whose number was kept raw during parsing. */
export function finalizeNumber(cell: { v: string; n?: number; raw?: string; fc?: string }, format: string, date1904: boolean) {
  const value = cell.n;
  if (value === undefined) return;
  delete cell.n;
  cell.v = formatValue(format, value, date1904);
  const raw = String(value);
  if (cell.v !== raw) cell.raw = raw;
  const color = formatColor(format, value);
  if (color) cell.fc = color;
}

const dateCache = new Map<string, boolean>();

export function isDateFormatCached(format: string): boolean {
  let hit = dateCache.get(format);
  if (hit === undefined) {
    hit = isDateFormat(format);
    dateCache.set(format, hit);
  }
  return hit;
}

export function isDateFormat(format: string): boolean {
  try {
    return SSF.is_date(format);
  } catch {
    return false;
  }
}

// [DBNum1] and friends select Korean or CJK numerals, which SSF does not implement.
const KOREAN_DIGITS = ["영", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"];
const HANJA_DIGITS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];

function dbnum(text: string, kind: string): string {
  // Korean Excel: DBNum1 writes 一二三, DBNum4 writes 일이삼. Others keep Arabic digits.
  const digits = kind === "1" ? HANJA_DIGITS : kind === "4" ? KOREAN_DIGITS : null;
  if (!digits) return text;
  return text.replace(/\d/g, (digit) => digits[Number(digit)]!);
}

export function formatValue(format: string, value: number | string | boolean, date1904: boolean): string {
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  let fmt = format || "General";
  const db = /\[DBNum(\d)\]/i.exec(fmt);
  if (db) fmt = fmt.replace(/\[DBNum\d\]/gi, "");
  try {
    const out = SSF.format(fmt, value, { date1904 }) as string;
    return db ? dbnum(out, db[1]!) : out;
  } catch {
    try {
      return typeof value === "number" ? (SSF.format("General", value) as string) : String(value);
    } catch {
      return String(value);
    }
  }
}
