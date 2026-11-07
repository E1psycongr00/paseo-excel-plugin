import { useRpc } from "@getpaseo/plugin/client";
import { useQueries, useQuery } from "@tanstack/react-query";
import { useMemo } from "react";
import type { CellChunk, SheetLayout, WorkbookMeta } from "../shared/model";
import { imageChunkRpc, sheetCellsRpc, sheetLayoutRpc, workbookOpenRpc } from "../shared/rpc";
import { indexCells, type RowCells } from "./geometry";

export interface FileKey {
  workspaceId: string;
  path: string;
}

export function useWorkbook(file: FileKey, version: number) {
  const open = useRpc(workbookOpenRpc);
  return useQuery({
    queryKey: ["excel.open", file.workspaceId, file.path, version],
    queryFn: async () => JSON.parse((await open(file)).json) as WorkbookMeta,
    staleTime: Infinity,
    retry: 1,
  });
}

export function useSheetLayout(file: FileKey, revision: string | undefined, sheet: number) {
  const layout = useRpc(sheetLayoutRpc);
  return useQuery({
    queryKey: ["excel.layout", file.workspaceId, file.path, revision, sheet],
    queryFn: async () => JSON.parse((await layout({ ...file, revision: revision!, sheet })).json) as SheetLayout,
    enabled: revision !== undefined,
    staleTime: Infinity,
    retry: 1,
  });
}

/** Loads the row chunks covering [rowStart, rowEnd] and indexes their cells. */
export function useCellRows(file: FileKey, revision: string | undefined, layout: SheetLayout | undefined, rowStart: number, rowEnd: number) {
  const cells = useRpc(sheetCellsRpc);
  const chunkRows = layout?.chunkRows ?? 400;
  const lastRow = layout ? Math.min(layout.lastRow, layout.rowCount - 1) : -1;
  const chunks: number[] = [];
  if (layout && lastRow >= 0) {
    const first = Math.floor(Math.max(0, rowStart) / chunkRows);
    const last = Math.floor(Math.min(lastRow, Math.max(rowStart, rowEnd)) / chunkRows);
    for (let i = first; i <= last; i++) chunks.push(i);
  }
  const results = useQueries({
    queries: chunks.map((index) => ({
      queryKey: ["excel.cells", file.workspaceId, file.path, revision, layout?.index, index],
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
    for (const result of results) if (result.data) indexCells(result.data.cells, index);
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
