import { unzipSync, type Unzipped } from "fflate";

export interface UnzipLimits {
  /** Largest single entry after decompression. */
  entry: number;
  /** Sum of every kept entry after decompression. */
  total: number;
}

const MiB = 1024 * 1024;

// A real 50k-row workbook inflates about 7x (4 MB to 29 MB). A V8 string tops out near 512 MiB,
// so a larger sheet XML could not be read anyway.
export const WORKBOOK_LIMITS: UnzipLimits = { entry: 512 * MiB, total: 1024 * MiB };
export const IMAGE_LIMITS: UnzipLimits = { entry: 64 * MiB, total: 64 * MiB };

/**
 * Unzips the entries `keep` accepts, refusing archives whose declared sizes exceed `limits`.
 * fflate allocates each entry's output from its declared size and never grows it, so checking
 * that size before inflating also bounds memory for archives that lie about it.
 */
export function unzipLimited(bytes: Uint8Array, keep: (name: string) => boolean, limits: UnzipLimits = WORKBOOK_LIMITS): Unzipped {
  let total = 0;
  return unzipSync(bytes, {
    filter(file) {
      if (!keep(file.name)) return false;
      if (file.originalSize > limits.entry) throw new Error(`압축을 푼 크기가 너무 큰 항목이 있습니다 (${file.name}, ${Math.round(file.originalSize / MiB)}MB).`);
      total += file.originalSize;
      if (total > limits.total) throw new Error(`압축을 푼 전체 크기가 너무 큽니다 (${Math.round(limits.total / MiB)}MB 초과).`);
      return true;
    },
  });
}
