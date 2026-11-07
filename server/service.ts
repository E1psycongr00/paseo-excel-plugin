import { readFile, realpath, stat } from "node:fs/promises";
import { extname, isAbsolute, relative, sep } from "node:path";
import { IMAGE_CHUNK_BYTES, STALE_REVISION } from "../shared/rpc";
import { type CellChunk, initialSheet, type OpenResult, type SheetBundle } from "../shared/model";
import { cellsInRows, type ParsedSheet } from "./sheet-builder";
import type { ParsedWorkbook } from "./workbook";
import { parseXls } from "./xls/index";
import { parseXlsx } from "./xlsx/index";

const MAX_FILE_BYTES = 300 * 1024 * 1024;
const CACHE_SIZE = 6;
/** Rows below the saved top row that a bundle covers: a tall phone screen at small zoom. */
const FIRST_SCREEN_ROWS = 60;
// Parsing blocks the daemon's only thread, so a background prefetch must not stall the sheet on
// screen: about 150ms of parsing for either format.
const PREFETCH_MAX_BYTES = 4 * 1024 * 1024;

interface Entry {
  revision: string;
  workbook: ParsedWorkbook;
  imageB64: Map<string, { base64: string; mimeType: string }>;
}

const cache = new Map<string, Entry>();
const pending = new Map<string, Promise<Entry>>();

function remember(path: string, entry: Entry) {
  cache.delete(path);
  cache.set(path, entry);
  while (cache.size > CACHE_SIZE) cache.delete(cache.keys().next().value!);
}

function isCfb(bytes: Uint8Array) {
  return bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
}

function isZip(bytes: Uint8Array) {
  return bytes[0] === 0x50 && bytes[1] === 0x4b;
}

export function parseBytes(bytes: Uint8Array, revision: string, name = ""): ParsedWorkbook {
  if (isZip(bytes)) {
    if (/\.xlsb$/i.test(name)) throw new Error("바이너리 통합 문서(.xlsb)는 아직 지원하지 않습니다.");
    return parseXlsx(bytes, revision);
  }
  if (isCfb(bytes)) return parseXls(bytes, revision);
  throw new Error("Excel 파일 형식이 아닙니다.");
}

export async function loadWorkbook(path: string): Promise<Entry> {
  const info = await stat(path);
  if (!info.isFile()) throw new Error("파일이 아닙니다.");
  if (info.size > MAX_FILE_BYTES) throw new Error(`파일이 너무 큽니다 (${Math.round(info.size / 1024 / 1024)}MB).`);
  const revision = `${Math.trunc(info.mtimeMs)}-${info.size}`;
  const cached = cache.get(path);
  if (cached && cached.revision === revision) {
    remember(path, cached);
    return cached;
  }
  const key = `${path}\0${revision}`;
  const inflight = pending.get(key);
  if (inflight) return inflight;
  const promise = (async () => {
    const bytes = new Uint8Array(await readFile(path));
    const started = performance.now();
    const workbook = parseBytes(bytes, revision, path);
    console.log(`[excel-viewer] parsed ${extname(path)} ${bytes.length}B in ${(performance.now() - started).toFixed(0)}ms`);
    const entry: Entry = { revision, workbook, imageB64: new Map() };
    remember(path, entry);
    return entry;
  })().finally(() => pending.delete(key));
  pending.set(key, promise);
  return promise;
}

async function current(path: string, revision: string): Promise<Entry> {
  const entry = await loadWorkbook(path);
  if (entry.revision !== revision) throw new Error(STALE_REVISION);
  return entry;
}

function chunk(parsed: ParsedSheet, rowStart: number, rowEnd: number): CellChunk {
  return { sheet: parsed.layout.index, rowStart, rowEnd, cells: cellsInRows(parsed, rowStart, rowEnd) };
}

/** Layout plus the chunks (on the app's chunk grid) under the saved top rows and frozen rows. */
export function sheetBundle(parsed: ParsedSheet): SheetBundle {
  const { layout } = parsed;
  const size = layout.chunkRows;
  const lastRow = Math.min(layout.lastRow, layout.rowCount - 1);
  const indexes = new Set<number>();
  if (lastRow >= 0) {
    if (layout.view.frozenRows > 0) indexes.add(0);
    const top = Math.min(layout.view.topRow, lastRow);
    const bottom = Math.min(lastRow, layout.view.topRow + FIRST_SCREEN_ROWS);
    for (let i = Math.floor(top / size); i <= Math.floor(bottom / size); i++) indexes.add(i);
  }
  const chunks = [...indexes].sort((a, b) => a - b).map((i) => chunk(parsed, i * size, (i + 1) * size));
  return { layout, chunks };
}

export async function openWorkbook(path: string, requested?: number) {
  const entry = await loadWorkbook(path);
  const { meta } = entry.workbook;
  let initial: SheetBundle | null = null;
  if (meta.sheets.length > 0) {
    // A sheet that fails to parse still opens the workbook; its own request then shows the error.
    try {
      initial = sheetBundle(entry.workbook.sheet(initialSheet(meta, requested)));
    } catch (error) {
      console.error(`[excel-viewer] first sheet failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  const result: OpenResult = { meta, initial };
  return { json: JSON.stringify(result) };
}

export async function sheetOpen(path: string, revision: string, sheet: number, prefetch: boolean) {
  const entry = await current(path, revision);
  if (prefetch && entry.workbook.sheetBytes(sheet) > PREFETCH_MAX_BYTES) return { json: "null" };
  return { json: JSON.stringify(sheetBundle(entry.workbook.sheet(sheet))) };
}

export async function sheetCells(path: string, revision: string, sheet: number, rowStart: number, rowEnd: number) {
  const entry = await current(path, revision);
  return { json: JSON.stringify(chunk(entry.workbook.sheet(sheet), rowStart, rowEnd)) };
}

export async function imageChunk(path: string, revision: string, key: string, index: number) {
  const entry = await current(path, revision);
  let encoded = entry.imageB64.get(key);
  if (!encoded) {
    const image = entry.workbook.image(key);
    if (!image) throw new Error("이미지를 찾을 수 없습니다.");
    encoded = { base64: Buffer.from(image.bytes).toString("base64"), mimeType: image.mimeType };
    entry.imageB64.set(key, encoded);
  }
  const chunkChars = (IMAGE_CHUNK_BYTES / 3) * 4;
  const chunks = Math.max(1, Math.ceil(encoded.base64.length / chunkChars));
  return { base64: encoded.base64.slice(index * chunkChars, (index + 1) * chunkChars), mimeType: encoded.mimeType, chunks };
}

function contains(root: string, target: string): boolean {
  const path = relative(root, target);
  return path === "" || (path !== ".." && !path.startsWith(`..${sep}`) && !isAbsolute(path));
}

interface WorkspaceLister {
  workspaces: {
    list(input: { filter?: { query?: string }; page?: { limit: number; cursor?: string } }): Promise<{
      entries: { id: string; workspaceDirectory?: string | null; projectRootPath?: string | null }[];
      pageInfo?: { hasMore?: boolean; nextCursor?: string | null } | null;
    }>;
  };
}

/** Resolves a file panel path and requires it to sit inside the workspace. */
export async function resolveWorkspaceFile(paseo: WorkspaceLister, workspaceId: string, path: string): Promise<string> {
  if (!isAbsolute(path) || path.includes("\0")) throw new Error("잘못된 파일 경로입니다.");
  let cursor: string | undefined;
  let root: string | null = null;
  for (let page = 0; page < 50 && !root; page++) {
    const result = await paseo.workspaces.list({ filter: { query: workspaceId }, page: { limit: 200, ...(cursor ? { cursor } : {}) } });
    const workspace = result.entries.find((entry) => entry.id === workspaceId);
    if (workspace) {
      root = workspace.workspaceDirectory || workspace.projectRootPath || null;
      if (!root) throw new Error("워크스페이스 폴더를 찾을 수 없습니다.");
      break;
    }
    const next = result.pageInfo?.nextCursor ?? undefined;
    if (!result.pageInfo?.hasMore || !next || next === cursor) break;
    cursor = next;
  }
  if (!root) throw new Error("워크스페이스를 찾을 수 없습니다.");
  const [realRoot, target] = await Promise.all([realpath(root), realpath(path)]);
  if (!contains(realRoot, target)) throw new Error("워크스페이스 밖의 파일은 열 수 없습니다.");
  return target;
}
