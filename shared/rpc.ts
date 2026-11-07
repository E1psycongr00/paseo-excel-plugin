import { defineRpc } from "@getpaseo/plugin";
import { z } from "zod";

// Payloads travel as JSON strings: a deep Zod schema would validate every cell twice per request.
const fileRef = {
  workspaceId: z.string(),
  /** Absolute path from the file panel's handle. */
  path: z.string(),
};

export const workbookOpenRpc = defineRpc({
  name: "excel.workbook.open",
  input: z.object(fileRef),
  output: z.object({ json: z.string() }),
});

export const sheetLayoutRpc = defineRpc({
  name: "excel.sheet.layout",
  input: z.object({ ...fileRef, revision: z.string(), sheet: z.number().int().min(0) }),
  output: z.object({ json: z.string() }),
});

export const sheetCellsRpc = defineRpc({
  name: "excel.sheet.cells",
  input: z.object({
    ...fileRef,
    revision: z.string(),
    sheet: z.number().int().min(0),
    rowStart: z.number().int().min(0),
    rowEnd: z.number().int().min(0),
  }),
  output: z.object({ json: z.string() }),
});

/** Image bytes as base64, in chunks that stay well under the IPC frame size. */
export const imageChunkRpc = defineRpc({
  name: "excel.image.chunk",
  input: z.object({ ...fileRef, revision: z.string(), key: z.string(), index: z.number().int().min(0) }),
  output: z.object({ base64: z.string(), mimeType: z.string(), chunks: z.number().int() }),
});

export const IMAGE_CHUNK_BYTES = 3 * 64 * 1024; // multiple of 3 so base64 chunks concatenate

export const STALE_REVISION = "stale-revision";
