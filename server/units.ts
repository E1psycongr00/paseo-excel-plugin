// Excel sizing. Column widths are stored in "characters" of the Normal font's maximum digit width
// (MDW); Excel converts them to pixels at 96 DPI. MDW values were measured with Excel 16 on Windows.

const MDW: Record<string, Record<number, number>> = {
  "calibri": { 9: 6, 10: 7, 11: 7, 12: 8 },
  "맑은 고딕": { 9: 7, 10: 7, 11: 8, 12: 9 },
  "malgun gothic": { 9: 7, 10: 7, 11: 8, 12: 9 },
  "돋움": { 9: 6, 10: 6, 11: 9, 12: 8 },
  "dotum": { 9: 6, 10: 6, 11: 9, 12: 8 },
  "굴림": { 9: 6, 10: 7, 11: 8, 12: 9 },
  "gulim": { 9: 6, 10: 7, 11: 8, 12: 9 },
  "바탕": { 9: 6, 10: 7, 11: 8, 12: 9 },
  "batang": { 9: 6, 10: 7, 11: 8, 12: 9 },
  "굴림체": { 9: 6, 10: 7, 11: 8, 12: 8 },
  "돋움체": { 9: 6, 10: 7, 11: 8, 12: 8 },
  "바탕체": { 9: 6, 10: 7, 11: 8, 12: 8 },
  "궁서": { 9: 6, 10: 6, 11: 7, 12: 9 },
  "arial": { 9: 7, 10: 7, 11: 8, 12: 9 },
  "times new roman": { 9: 6, 10: 6, 11: 7, 12: 8 },
  "나눔고딕": { 9: 7, 10: 7, 11: 8, 12: 9 },
};

// Default row heights (points) for the Normal font, measured the same way.
const ROW_HEIGHT: Record<string, Record<number, number>> = {
  "calibri": { 9: 12, 10: 12.75, 11: 15, 12: 15.75 },
  "맑은 고딕": { 9: 12, 10: 13.5, 11: 16.5, 12: 17.25 },
  "malgun gothic": { 9: 12, 10: 13.5, 11: 16.5, 12: 17.25 },
  "나눔고딕": { 9: 12, 10: 13.5, 11: 16.5, 12: 17.25 },
  "arial": { 9: 12, 10: 12.75, 11: 14.25, 12: 15 },
  "times new roman": { 9: 12, 10: 12.75, 11: 15, 12: 15.75 },
};

export function maxDigitWidth(font: string, size: number): number {
  const known = MDW[font.toLowerCase()]?.[Math.round(size)];
  if (known) return known;
  return Math.max(5, Math.round(size * 0.7));
}

export function defaultRowHeightPt(font: string, size: number): number {
  const known = ROW_HEIGHT[font.toLowerCase()]?.[Math.round(size)];
  if (known) return known;
  // Korean system fonts sit at about 1.23x the size, Latin fonts a little higher.
  return Math.round((size * 1.25) * 4) / 4 + 0.75;
}

/** Stored column width (characters, padding included) to pixels. */
export function columnWidthPx(width: number, mdw: number): number {
  if (width <= 0) return 0;
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** Width of `defaultColWidth`/`baseColWidth` when no explicit width is stored. */
export function defaultColumnPx(baseColWidth: number | undefined, defaultColWidth: number | undefined, mdw: number): number {
  if (defaultColWidth !== undefined) return columnWidthPx(defaultColWidth, mdw);
  const base = baseColWidth ?? 8;
  // Excel pads the base width with 5 px of margin and rounds up to a multiple of 8 px.
  const px = Math.trunc(base * mdw + 5);
  return Math.ceil(px / 8) * 8;
}

export const ptToPx = (pt: number) => (pt * 96) / 72;

/** EMU (drawing units) to px. */
export const emuToPx = (emu: number) => emu / 9525;
