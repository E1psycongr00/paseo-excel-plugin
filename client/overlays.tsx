import { memo, type ReactElement } from "react";
import { Image, Text, View } from "react-native";
import type { DrawingItem, SheetLayout } from "../shared/model";
import { type FileKey, useImage } from "./data";
import type { Axis, Window } from "./geometry";

const PAGE_BLUE = "#2F5BEA";

export function drawingRect(item: DrawingItem, rows: Axis, cols: Axis, zoom: number) {
  const { anchor } = item;
  if (anchor.pos && anchor.ext) return { x: anchor.pos.x * zoom, y: anchor.pos.y * zoom, w: anchor.ext.w * zoom, h: anchor.ext.h * zoom };
  const x = cols.start(anchor.from.c) + Math.min(anchor.from.dx * zoom, cols.size(anchor.from.c));
  const y = rows.start(anchor.from.r) + Math.min(anchor.from.dy * zoom, rows.size(anchor.from.r));
  if (anchor.to) {
    const x2 = cols.start(anchor.to.c) + Math.min(anchor.to.dx * zoom, cols.size(anchor.to.c));
    const y2 = rows.start(anchor.to.r) + Math.min(anchor.to.dy * zoom, rows.size(anchor.to.r));
    return { x, y, w: Math.max(0, x2 - x), h: Math.max(0, y2 - y) };
  }
  return { x, y, w: (anchor.ext?.w ?? 0) * zoom, h: (anchor.ext?.h ?? 0) * zoom };
}

function Picture({ file, revision, item, rect }: { file: FileKey; revision: string; item: DrawingItem; rect: { x: number; y: number; w: number; h: number } }) {
  const image = useImage(file, revision, item.image);
  const box = { position: "absolute" as const, left: rect.x, top: rect.y, width: rect.w, height: rect.h };
  if (!image.data) {
    return <View pointerEvents="none" style={[box, { backgroundColor: image.error ? "rgba(200,0,0,0.06)" : "rgba(0,0,0,0.04)" }]} />;
  }
  return (
    <View pointerEvents="none" style={box}>
      <Image source={{ uri: image.data }} style={{ width: rect.w, height: rect.h }} resizeMode="stretch" accessibilityLabel={item.name ?? "그림"} />
    </View>
  );
}

interface DrawingsProps {
  file: FileKey;
  revision: string;
  drawings: DrawingItem[];
  rows: Axis;
  cols: Axis;
  zoom: number;
  view: { x: number; y: number; w: number; h: number };
}

function DrawingsImpl({ file, revision, drawings, rows, cols, zoom, view }: DrawingsProps) {
  const out: ReactElement[] = [];
  for (const item of drawings) {
    const rect = drawingRect(item, rows, cols, zoom);
    if (rect.w <= 0 || rect.h <= 0) continue;
    // Only mount drawings near the viewport; images decode lazily.
    if (rect.x > view.x + view.w + 400 || rect.x + rect.w < view.x - 400 || rect.y > view.y + view.h + 400 || rect.y + rect.h < view.y - 400) continue;
    if (item.kind === "image") {
      out.push(<Picture key={item.id} file={file} revision={revision} item={item} rect={rect} />);
    } else if (item.kind === "chart") {
      out.push(
        <View key={item.id} pointerEvents="none" style={{ position: "absolute", left: rect.x, top: rect.y, width: rect.w, height: rect.h, borderWidth: 1, borderStyle: "dashed", borderColor: "#888888", backgroundColor: "rgba(255,255,255,0.85)", alignItems: "center", justifyContent: "center" }}>
          <Text style={{ color: "#666666", fontSize: 12 * zoom }}>{`차트: ${item.name ?? ""} (표시 안 함)`}</Text>
        </View>,
      );
    } else {
      out.push(
        <View key={item.id} pointerEvents="none" style={{ position: "absolute", left: rect.x, top: rect.y, width: rect.w, height: rect.h, borderWidth: item.line ? 1 : 0, borderColor: item.line ?? "transparent", backgroundColor: item.fill ?? "transparent", alignItems: "center", justifyContent: "center", overflow: "hidden" }}>
          {item.text ? <Text style={{ color: "#000000", fontSize: 11 * (96 / 72) * zoom, textAlign: "center" }}>{item.text}</Text> : null}
        </View>,
      );
    }
  }
  return <>{out}</>;
}

export const Drawings = memo(DrawingsImpl);

interface PagesProps {
  layout: SheetLayout;
  rows: Axis;
  cols: Axis;
  zoom: number;
  win: Window;
  preview: boolean;
}

/** Page break preview: gray outside the print range, blue page edges, page numbers. */
function PagesImpl({ layout, rows, cols, zoom, win, preview }: PagesProps) {
  const range = layout.print.pageRange;
  if (!range) return null;
  const [r1, c1, r2, c2] = range;
  const rowEdges = [r1, ...layout.print.rowBreaks.map((b) => b.at), r2 + 1];
  const colEdges = [c1, ...layout.print.colBreaks.map((b) => b.at), c2 + 1];
  const manualRow = new Set(layout.print.rowBreaks.filter((b) => b.manual).map((b) => b.at));
  const manualCol = new Set(layout.print.colBreaks.filter((b) => b.manual).map((b) => b.at));
  const x1 = cols.start(c1);
  const x2 = cols.end(c2);
  const y1 = rows.start(r1);
  const y2 = rows.end(r2);
  const totalW = cols.end(win.c1);
  const totalH = rows.end(win.r1);
  const out: ReactElement[] = [];
  const thick = preview ? Math.max(2, Math.round(3 * Math.min(1, zoom))) : 1;
  const line = (id: string, x: number, y: number, w: number, h: number, dashed: boolean) => {
    if (dashed) {
      const half = thick / 2;
      out.push(<View key={id} pointerEvents="none" style={{ position: "absolute", left: x - (w === 0 ? half : 0), top: y - (h === 0 ? half : 0), width: w, height: h, borderWidth: half, borderStyle: "dashed", borderColor: preview ? PAGE_BLUE : "#555555" }} />);
    } else {
      out.push(<View key={id} pointerEvents="none" style={{ position: "absolute", left: x - (w === 0 ? thick / 2 : 0), top: y - (h === 0 ? thick / 2 : 0), width: w || thick, height: h || thick, backgroundColor: PAGE_BLUE }} />);
    }
  };

  if (preview) {
    const shade = "rgba(110,110,110,0.32)";
    const top = rows.start(win.r0);
    const left = cols.start(win.c0);
    if (y1 > top) out.push(<View key="g-top" pointerEvents="none" style={{ position: "absolute", left, top, width: totalW - left, height: y1 - top, backgroundColor: shade }} />);
    if (totalH > y2) out.push(<View key="g-bottom" pointerEvents="none" style={{ position: "absolute", left, top: y2, width: totalW - left, height: totalH - y2, backgroundColor: shade }} />);
    if (x1 > left) out.push(<View key="g-left" pointerEvents="none" style={{ position: "absolute", left, top: y1, width: x1 - left, height: y2 - y1, backgroundColor: shade }} />);
    if (totalW > x2) out.push(<View key="g-right" pointerEvents="none" style={{ position: "absolute", left: x2, top: y1, width: totalW - x2, height: y2 - y1, backgroundColor: shade }} />);
    line("edge-t", x1, y1, x2 - x1, 0, false);
    line("edge-b", x1, y2, x2 - x1, 0, false);
    line("edge-l", x1, y1, 0, y2 - y1, false);
    line("edge-r", x2, y1, 0, y2 - y1, false);
  }
  // Preview: manual breaks solid, automatic dashed (as Excel). Normal view: thin dashed lines only.
  for (let i = 1; i < rowEdges.length - 1; i++) {
    line(`rb${i}`, x1, rows.start(rowEdges[i]!), x2 - x1, 0, !preview || !manualRow.has(rowEdges[i]!));
  }
  for (let j = 1; j < colEdges.length - 1; j++) {
    line(`cb${j}`, cols.start(colEdges[j]!), y1, 0, y2 - y1, !preview || !manualCol.has(colEdges[j]!));
  }
  if (preview) {
    const pagesDown = rowEdges.length - 1;
    const pagesAcross = colEdges.length - 1;
    for (let i = 0; i < pagesDown; i++) {
      for (let j = 0; j < pagesAcross; j++) {
        const px1 = cols.start(colEdges[j]!);
        const px2 = cols.start(colEdges[j + 1]!);
        const py1 = rows.start(rowEdges[i]!);
        const py2 = rows.start(rowEdges[i + 1]!);
        if (px2 < cols.start(win.c0) || px1 > totalW || py2 < rows.start(win.r0) || py1 > totalH) continue;
        const number = layout.print.order === "down" ? j * pagesDown + i + 1 : i * pagesAcross + j + 1;
        const size = Math.max(14, Math.min(96, Math.min(px2 - px1, py2 - py1) / 5));
        out.push(
          <View key={`p${i}_${j}`} pointerEvents="none" style={{ position: "absolute", left: px1, top: py1, width: px2 - px1, height: Math.min(py2 - py1, 4000), alignItems: "center", justifyContent: "center" }}>
            <Text style={{ fontSize: size, fontWeight: "700", color: "rgba(130,130,130,0.32)" }}>{`${number} 페이지`}</Text>
          </View>,
        );
      }
    }
  }
  return <>{out}</>;
}

export const Pages = memo(PagesImpl);
