import type { PluginHandlerContext, PluginServerContext } from "@getpaseo/plugin/server";
import { imageChunk, openWorkbook, resolveWorkspaceFile, sheetCells, sheetLayout } from "./server/service";
import { imageChunkRpc, sheetCellsRpc, sheetLayoutRpc, workbookOpenRpc } from "./shared/rpc";

const RESOLVE_TTL_MS = 30_000;
const resolved = new Map<string, { path: string; at: number }>();

async function resolve(paseo: PluginHandlerContext["paseo"], workspaceId: string, path: string): Promise<string> {
  const key = `${workspaceId}\0${path}`;
  const hit = resolved.get(key);
  if (hit && Date.now() - hit.at < RESOLVE_TTL_MS) return hit.path;
  const target = await resolveWorkspaceFile(paseo as unknown as Parameters<typeof resolveWorkspaceFile>[0], workspaceId, path);
  resolved.set(key, { path: target, at: Date.now() });
  if (resolved.size > 200) resolved.delete(resolved.keys().next().value!);
  return target;
}

export default function contribute(server: PluginServerContext) {
  server.handle(workbookOpenRpc, async ({ workspaceId, path }, { paseo }) => openWorkbook(await resolve(paseo, workspaceId, path)));
  server.handle(sheetLayoutRpc, async ({ workspaceId, path, revision, sheet }, { paseo }) =>
    sheetLayout(await resolve(paseo, workspaceId, path), revision, sheet));
  server.handle(sheetCellsRpc, async ({ workspaceId, path, revision, sheet, rowStart, rowEnd }, { paseo }) =>
    sheetCells(await resolve(paseo, workspaceId, path), revision, sheet, rowStart, rowEnd));
  server.handle(imageChunkRpc, async ({ workspaceId, path, revision, key, index }, { paseo }) =>
    imageChunk(await resolve(paseo, workspaceId, path), revision, key, index));
  return () => {
    resolved.clear();
  };
}
