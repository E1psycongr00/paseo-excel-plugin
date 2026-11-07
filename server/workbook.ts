import type { SheetSummary, WorkbookMeta } from "../shared/model";
import type { ParsedSheet } from "./sheet-builder";

export interface ImageData {
  bytes: Uint8Array;
  mimeType: string;
}

/** Format-independent parsed workbook. Sheets parse on first use. */
export interface ParsedWorkbook {
  meta: WorkbookMeta;
  sheet(index: number): ParsedSheet;
  /** Bytes a sheet's parse would read, or 0 once parsed. Bounds background prefetch. */
  sheetBytes(index: number): number;
  image(key: string): ImageData | null;
}

export function summaryState(value: string | undefined): SheetSummary["state"] {
  return value === "hidden" ? "hidden" : value === "veryHidden" ? "veryHidden" : "visible";
}

export function mimeFromName(name: string): string {
  const ext = name.slice(name.lastIndexOf(".") + 1).toLowerCase();
  switch (ext) {
    case "png": return "image/png";
    case "jpg":
    case "jpeg": return "image/jpeg";
    case "gif": return "image/gif";
    case "bmp": return "image/bmp";
    case "webp": return "image/webp";
    case "svg": return "image/svg+xml";
    case "emf": return "image/emf";
    case "wmf": return "image/wmf";
    case "tif":
    case "tiff": return "image/tiff";
    default: return "application/octet-stream";
  }
}

/** Sniffs image bytes; BIFF blips carry no file name. */
export function sniffMime(bytes: Uint8Array): string {
  if (bytes[0] === 0x89 && bytes[1] === 0x50) return "image/png";
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return "image/jpeg";
  if (bytes[0] === 0x47 && bytes[1] === 0x49) return "image/gif";
  if (bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
  return "application/octet-stream";
}
