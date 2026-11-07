import type { PluginHandlerContext, PluginServerContext } from "@getpaseo/plugin/server";
import { imageChunk, openWorkbook, resolveWorkspaceFile, sheetCells, sheetOpen } from "./server/service";
import { imageChunkRpc, sheetCellsRpc, sheetOpenRpc, workbookOpenRpc } from "./shared/rpc";

// Re-checking the workspace lists workspaces over IPC; doing it every 30s stalled cell requests while
// scrolling. Concurrent requests for the same file share one check.
const RESOLVE_TTL_MS = 5 * 60_000;
const resolved = new Map<string, { path: Promise<string>; at: number }>();

function resolve(paseo: PluginHandlerContext["paseo"], workspaceId: string, path: string): Promise<string> {
  const key = `${workspaceId}\0${path}`;
  const hit = resolved.get(key);
  if (hit && Date.now() - hit.at < RESOLVE_TTL_MS) return hit.path;
  const target = resolveWorkspaceFile(paseo as unknown as Parameters<typeof resolveWorkspaceFile>[0], workspaceId, path);
  resolved.set(key, { path: target, at: Date.now() });
  // A failed check is not cached, so the next request tries again.
  target.catch(() => {
    if (resolved.get(key)?.path === target) resolved.delete(key);
  });
  if (resolved.size > 200) resolved.delete(resolved.keys().next().value!);
  return target;
}

export default function contribute(server: PluginServerContext) {
  server.handle(workbookOpenRpc, async ({ workspaceId, path, sheet }, { paseo }) => openWorkbook(await resolve(paseo, workspaceId, path), sheet));
  server.handle(sheetOpenRpc, async ({ workspaceId, path, revision, sheet, prefetch }, { paseo }) =>
    sheetOpen(await resolve(paseo, workspaceId, path), revision, sheet, prefetch));
  server.handle(sheetCellsRpc, async ({ workspaceId, path, revision, sheet, rowStart, rowEnd }, { paseo }) =>
    sheetCells(await resolve(paseo, workspaceId, path), revision, sheet, rowStart, rowEnd));
  server.handle(imageChunkRpc, async ({ workspaceId, path, revision, key, index }, { paseo }) =>
    imageChunk(await resolve(paseo, workspaceId, path), revision, key, index));
  return () => {
    resolved.clear();
  };
}
