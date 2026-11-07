import { useRpc } from "@getpaseo/plugin/client";
import { type QueryClient, useQueries, useQuery, useQueryClient } from "@tanstack/react-query";
import { useEffect, useMemo, useRef } from "react";
import type { CellChunk, OpenResult, SheetBundle, SheetLayout, WorkbookMeta } from "../shared/model";
import { imageChunkRpc, sheetCellsRpc, sheetOpenRpc, workbookOpenRpc } from "../shared/rpc";
import { indexCells, type RowCells } from "./geometry";

export interface FileKey {
  workspaceId: string;
  path: string;
}

const layoutKey = (file: FileKey, revision: string | undefined, sheet: number) => ["excel.layout", file.workspaceId, file.path, revision, sheet];
const cellsKey = (file: FileKey, revision: string | undefined, sheet: number | undefined, chunk: number) =>
  ["excel.cells", file.workspaceId, file.path, revision, sheet, chunk];

/** Puts a bundle's layout and cell chunks where the sheet's own queries look, so they skip the network. */
function seedSheet(client: QueryClient, file: FileKey, revision: string, bundle: SheetBundle) {
  const { layout } = bundle;
  if (client.getQueryData(layoutKey(file, revision, layout.index)) === undefined) client.setQueryData(layoutKey(file, revision, layout.index), layout);
  for (const chunk of bundle.chunks) {
    const key = cellsKey(file, revision, layout.index, Math.floor(chunk.rowStart / layout.chunkRows));
    if (client.getQueryData(key) === undefined) client.setQueryData(key, chunk);
  }
}

/** Opens the workbook; the response also carries the first sheet's layout and first-screen cells. */
export function useWorkbook(file: FileKey, version: number, savedSheet: number | undefined) {
  const open = useRpc(workbookOpenRpc);
  const client = useQueryClient();
  return useQuery({
    // The saved sheet only picks what comes bundled, so it stays out of the key.
    queryKey: ["excel.open", file.workspaceId, file.path, version],
    queryFn: async () => {
      const result = JSON.parse((await open({ ...file, ...(savedSheet !== undefined ? { sheet: savedSheet } : {}) })).json) as OpenResult;
      if (result.initial) seedSheet(client, file, result.meta.revision, result.initial);
      return result.meta;
    },
    staleTime: Infinity,
    retry: 1,
  });
}

export function useSheetLayout(file: FileKey, revision: string | undefined, sheet: number) {
  const open = useRpc(sheetOpenRpc);
  const client = useQueryClient();
  return useQuery({
    queryKey: layoutKey(file, revision, sheet),
    queryFn: async () => {
      const bundle = JSON.parse((await open({ ...file, revision: revision!, sheet, prefetch: false })).json) as SheetBundle;
      seedSheet(client, file, revision!, bundle);
      return bundle.layout;
    },
    enabled: revision !== undefined,
    staleTime: Infinity,
    retry: 1,
  });
}

/** Wait after the current sheet is up before fetching others, so its own scroll-ahead goes first. */
const PREFETCH_DELAY_MS = 1000;

/**
 * Fetches the other visible sheets' first screens one at a time, nearest tab first, so tapping a tab
 * shows it at once. The daemon declines sheets too large to parse in the background.
 */
export function usePrefetchSheets(file: FileKey, meta: WorkbookMeta, current: number, ready: boolean) {
  const open = useRpc(sheetOpenRpc);
  const openRef = useRef(open);
  openRef.current = open;
  const client = useQueryClient();
  useEffect(() => {
    if (!ready) return;
    let cancelled = false;
    const order = meta.sheets
      .filter((sheet) => sheet.state === "visible" && sheet.index !== current)
      .map((sheet) => sheet.index)
      .sort((a, b) => Math.abs(a - current) - Math.abs(b - current) || a - b);
    const timer = setTimeout(async () => {
      for (const sheet of order) {
        if (cancelled) return;
        if (client.getQueryData(layoutKey(file, meta.revision, sheet)) !== undefined) continue;
        try {
          const bundle = JSON.parse((await openRef.current({ ...file, revision: meta.revision, sheet, prefetch: true })).json) as SheetBundle | null;
          if (!cancelled && bundle) seedSheet(client, file, meta.revision, bundle);
        } catch {
          // Stale file or lost connection: stop quietly; the sheet's own request reports errors.
          return;
        }
      }
    }, PREFETCH_DELAY_MS);
    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [ready, file, meta, current, client]);
}

// One index per received chunk, so a row's Map keeps its identity while other chunks come and go;
// drawn tiles compare row Maps to know whether they can be reused.
const chunkIndex = new WeakMap<CellChunk, RowCells>();

function indexChunk(chunk: CellChunk): RowCells {
  let index = chunkIndex.get(chunk);
  if (!index) {
    index = indexCells(chunk.cells);
    chunkIndex.set(chunk, index);
  }
  return index;
}

/**
 * Loads the row chunks covering [rowStart, rowEnd] and indexes their cells. `ahead` rows on both
 * sides load too, so scrolling reaches rows that are already on the device.
 */
export function useCellRows(file: FileKey, revision: string | undefined, layout: SheetLayout | undefined, rowStart: number, rowEnd: number, ahead = 0) {
  const cells = useRpc(sheetCellsRpc);
  const chunkRows = layout?.chunkRows ?? 400;
  const lastRow = layout ? Math.min(layout.lastRow, layout.rowCount - 1) : -1;
  const chunks: number[] = [];
  if (layout && lastRow >= 0 && rowEnd >= rowStart) {
    const first = Math.floor(Math.max(0, rowStart - ahead) / chunkRows);
    const last = Math.floor(Math.min(lastRow, rowEnd + ahead) / chunkRows);
    for (let i = first; i <= last; i++) chunks.push(i);
  }
  const results = useQueries({
    queries: chunks.map((index) => ({
      queryKey: cellsKey(file, revision, layout?.index, index),
      queryFn: async () => {
        const rowStartChunk = index * chunkRows;
        const response = await cells({ ...file, revision: revision!, sheet: layout!.index, rowStart: rowStartChunk, rowEnd: rowStartChunk + chunkRows });
        return JSON.parse(response.json) as CellChunk;
      },
      enabled: revision !== undefined && layout !== undefined,
      staleTime: Infinity,
      gcTime: 5 * 60_000,
    })),
  });
  const stamp = results.map((result) => result.dataUpdatedAt).join(",");
  const rows = useMemo<RowCells>(() => {
    const index: RowCells = new Map();
    for (const result of results) {
      if (!result.data) continue;
      for (const [r, row] of indexChunk(result.data)) index.set(r, row);
    }
    return index;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [stamp]);
  const loading = results.some((result) => result.isPending);
  const error = results.find((result) => result.error)?.error ?? null;
  return { rows, loading, error };
}

export function useImage(file: FileKey, revision: string, key: string | undefined) {
  const chunk = useRpc(imageChunkRpc);
  return useQuery({
    queryKey: ["excel.image", file.workspaceId, file.path, revision, key],
    queryFn: async () => {
      const first = await chunk({ ...file, revision, key: key!, index: 0 });
      let base64 = first.base64;
      for (let index = 1; index < first.chunks; index++) base64 += (await chunk({ ...file, revision, key: key!, index })).base64;
      return `data:${first.mimeType};base64,${base64}`;
    },
    enabled: key !== undefined,
    staleTime: Infinity,
    gcTime: 10 * 60_000,
  });
}
