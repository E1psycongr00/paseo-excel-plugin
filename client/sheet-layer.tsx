import { memo, type ReactElement } from "react";
import { Text, View, type TextStyle } from "react-native";
import type { BorderSide, Cell, CellComment, CellStyle, Range, SheetLayout } from "../shared/model";
import { type StyleCache, runStyle } from "./cell-style";
import { type Axis, effectiveAlign, estimateTextWidth, type Merges, overflowSpan, type RowCells, type Window } from "./geometry";

export const GRID_COLOR = "#D9D9D9";
export const SELECT_COLOR = "#217346";
const BIG = 4000;

export interface LayerProps {
  layout: SheetLayout;
  rows: Axis;
  cols: Axis;
  win: Window;
  /** Subtracted from content coordinates (bands draw with their own origin). */
  ox: number;
  oy: number;
  cells: RowCells;
  merges: Merges;
  styles: StyleCache;
  zoom: number;
  showGrid: boolean;
  showHidden: boolean;
  rowStyle: (r: number) => number | null;
  colStyle: (c: number) => number | null;
  comments: Map<number, CellComment>;
  selected: Range | null;
}

const key = (r: number, c: number) => r * 16384 + c;

function lineThickness(side: BorderSide, zoom: number): number {
  return side.w <= 1 ? 1 : Math.max(1, Math.round(side.w * Math.min(1.5, zoom)));
}

// A border side as absolutely positioned views; dashed and dotted use a uniform-border box, which
// renders on every platform (single-side dashed borders do not on Android).
function borderLine(id: string, side: BorderSide, x: number, y: number, length: number, horizontal: boolean, zoom: number, out: ReactElement[]) {
  const t = lineThickness(side, zoom);
  if (side.line === "double") {
    const gap = 1;
    const a = horizontal ? { left: x, top: y - gap - 1, width: length, height: 1 } : { left: x - gap - 1, top: y, width: 1, height: length };
    const b = horizontal ? { left: x, top: y + gap, width: length, height: 1 } : { left: x + gap, top: y, width: 1, height: length };
    out.push(<View key={`${id}a`} pointerEvents="none" style={{ position: "absolute", backgroundColor: side.color, ...a }} />);
    out.push(<View key={`${id}b`} pointerEvents="none" style={{ position: "absolute", backgroundColor: side.color, ...b }} />);
    return;
  }
  const offset = Math.floor(t / 2);
  const rect = horizontal
    ? { left: x, top: y - offset, width: length, height: t }
    : { left: x - offset, top: y, width: t, height: length };
  if (side.line === "solid") {
    out.push(<View key={id} pointerEvents="none" style={{ position: "absolute", backgroundColor: side.color, ...rect }} />);
  } else {
    const half = Math.max(0.5, t / 2);
    out.push(
      <View
        key={id}
        pointerEvents="none"
        style={{
          position: "absolute",
          ...(horizontal ? { left: x, top: y - half, width: length, height: 0 } : { left: x - half, top: y, width: 0, height: length }),
          borderWidth: half,
          borderStyle: side.line,
          borderColor: side.color,
        }}
      />,
    );
  }
}

function diagonal(id: string, side: BorderSide, x: number, y: number, w: number, h: number, up: boolean, zoom: number, out: ReactElement[]) {
  const length = Math.hypot(w, h);
  const angle = (Math.atan2(h, w) * 180) / Math.PI * (up ? -1 : 1);
  out.push(
    <View
      key={id}
      pointerEvents="none"
      style={{
        position: "absolute",
        left: x + w / 2 - length / 2,
        top: y + h / 2 - lineThickness(side, zoom) / 2,
        width: length,
        height: lineThickness(side, zoom),
        backgroundColor: side.color,
        transform: [{ rotate: `${angle}deg` }],
      }}
    />,
  );
}

interface TextBox {
  x: number;
  y: number;
  w: number;
  h: number;
  /** Width of the cell itself (spill excluded) for "####" decisions. */
  cellW: number;
}

function cellText(id: string, cell: Cell, style: CellStyle | undefined, box: TextBox, props: LayerProps, out: ReactElement[]) {
  const { zoom, styles } = props;
  const base = styles.text(cell.s, cell.fc);
  const align = effectiveAlign(style, cell);
  const pad = Math.max(1, 2 * zoom);
  const indent = (style?.indent ?? 0) * 9 * zoom;
  const fontPx = (base.fontSize as number) ?? 14;
  const rotation = style?.rotation ?? 0;
  const vertical = style?.v === "top" ? "flex-start" : style?.v === "center" || style?.v === "justify" || style?.v === "distributed" ? "center" : "flex-end";
  const textAlign: TextStyle["textAlign"] = align;
  let text = cell.v;
  let fontScale = 1;
  const wrap = !!style?.wrap;

  if (!wrap && rotation === 0) {
    if (text.indexOf("\n") >= 0) text = text.replace(/\n/g, " ");
    const available = box.cellW - pad * 2 - indent;
    if (style?.shrink) {
      const width = estimateTextWidth(text, fontPx);
      if (width > available && width > 0) fontScale = Math.max(0.4, available / width);
    } else if ((cell.k === "n" || cell.k === "d") && estimateTextWidth(text, fontPx) > available + 1) {
      text = "#".repeat(Math.max(1, Math.floor(available / (fontPx * 0.56))));
    }
  }
  const textStyle: TextStyle[] = [base, { textAlign }];
  if (fontScale !== 1) textStyle.push({ fontSize: fontPx * fontScale, lineHeight: Math.ceil(fontPx * fontScale * 1.22) });
  const content = cell.rich && text === cell.v
    ? cell.rich.map((run, index) => (
      <Text key={index} style={runStyle(run, (style ?? styles.get(0))!.font, zoom * fontScale)}>{wrap ? run.text : run.text.replace(/\n/g, " ")}</Text>
    ))
    : text;

  if (rotation === 255) {
    out.push(
      <View key={id} pointerEvents="none" style={{ position: "absolute", left: box.x, top: box.y, width: box.w, height: box.h, overflow: "hidden", justifyContent: vertical, alignItems: "center" }}>
        <Text style={[base, { textAlign: "center" }]}>{cell.v.split("").join("\n")}</Text>
      </View>,
    );
    return;
  }
  if (rotation !== 0) {
    // 1-90 rotate counter-clockwise, 91-180 clockwise.
    const degrees = rotation <= 90 ? -rotation : rotation - 90;
    const turned = Math.abs(degrees) > 45;
    const w = turned ? box.h : box.w;
    const h = turned ? box.w : box.h;
    out.push(
      <View key={id} pointerEvents="none" style={{ position: "absolute", left: box.x, top: box.y, width: box.w, height: box.h, overflow: "hidden" }}>
        <View style={{ position: "absolute", left: (box.w - w) / 2, top: (box.h - h) / 2, width: w, height: h, justifyContent: "center", transform: [{ rotate: `${degrees}deg` }] }}>
          <Text style={[base, { textAlign: "center" }]} numberOfLines={wrap ? undefined : 1}>{content}</Text>
        </View>
      </View>,
    );
    return;
  }

  const inner = box.w - pad * 2 - indent;
  let marginLeft: number;
  let width: number;
  if (wrap) {
    width = Math.max(1, inner);
    marginLeft = pad + (align === "left" ? indent : 0);
  } else {
    width = BIG;
    marginLeft = align === "left" ? pad + indent : align === "right" ? box.w - BIG - pad - indent : (box.w - BIG) / 2;
  }
  out.push(
    <View key={id} pointerEvents="none" style={{ position: "absolute", left: box.x, top: box.y, width: box.w, height: box.h, overflow: "hidden", justifyContent: vertical }}>
      <Text style={[...textStyle, { width, marginLeft }]}>{content}</Text>
    </View>,
  );
}

function SheetLayerImpl(props: LayerProps) {
  const { layout, rows, cols, win, ox, oy, cells, merges, styles, zoom, showGrid, showHidden, comments, selected } = props;
  const grid: ReactElement[] = [];
  const blanks: ReactElement[] = [];
  const fills: ReactElement[] = [];
  const texts: ReactElement[] = [];
  const borders: ReactElement[] = [];
  const marks: ReactElement[] = [];

  const left = cols.start(win.c0) - ox;
  const top = rows.start(win.r0) - oy;
  const right = cols.end(win.c1) - ox;
  const bottom = rows.end(win.r1) - oy;

  if (showGrid) {
    for (let c = win.c0; c <= win.c1 + 1; c++) {
      if (c <= win.c1 && cols.size(c) === 0) continue;
      grid.push(<View key={`gc${c}`} style={{ position: "absolute", left: cols.start(c) - ox, top, width: 1, height: bottom - top, backgroundColor: GRID_COLOR }} />);
    }
    for (let r = win.r0; r <= win.r1 + 1; r++) {
      if (r <= win.r1 && rows.size(r) === 0) continue;
      grid.push(<View key={`gr${r}`} style={{ position: "absolute", left, top: rows.start(r) - oy, width: right - left, height: 1, backgroundColor: GRID_COLOR }} />);
    }
  }

  const styleAt = (r: number, c: number, cell: Cell | undefined): number => cell ? cell.s : props.rowStyle(r) ?? props.colStyle(c) ?? 0;

  const drawBorders = (r: number, c: number, style: CellStyle, merge: Range | undefined) => {
    const b = style.border;
    if (!b.l && !b.r && !b.t && !b.b && !b.up && !b.down) return;
    const x = cols.start(c) - ox;
    const y = rows.start(r) - oy;
    const w = cols.size(c);
    const h = rows.size(r);
    if (w === 0 || h === 0) return;
    const id = `b${r}_${c}`;
    // Inside a merge only the perimeter sides draw.
    if (b.l && (!merge || c === merge[1])) borderLine(`${id}l`, b.l, x, y, h, false, zoom, borders);
    if (b.r && (!merge || c === merge[3])) borderLine(`${id}r`, b.r, x + w, y, h, false, zoom, borders);
    if (b.t && (!merge || r === merge[0])) borderLine(`${id}t`, b.t, x, y, w, true, zoom, borders);
    if (b.b && (!merge || r === merge[2])) borderLine(`${id}b`, b.b, x, y + h, w, true, zoom, borders);
    if (!merge) {
      if (b.up) diagonal(`${id}u`, b.up, x, y, w, h, true, zoom, borders);
      if (b.down) diagonal(`${id}d`, b.down, x, y, w, h, false, zoom, borders);
    }
  };

  for (let r = win.r0; r <= win.r1; r++) {
    const h = rows.size(r);
    if (h === 0) continue;
    const row = cells.get(r);
    for (let c = win.c0; c <= win.c1; c++) {
      const w = cols.size(c);
      if (w === 0) continue;
      const cell = row?.get(c);
      const merge = merges.at(r, c);
      const id = styleAt(r, c, cell);
      if (id === 0 && !cell) continue;
      const style = styles.get(id);
      if (style) drawBorders(r, c, style, merge);
      if (merge) continue;
      const x = cols.start(c) - ox;
      const y = rows.start(r) - oy;
      if (style?.fill) fills.push(<View key={`f${r}_${c}`} style={{ position: "absolute", left: x, top: y, width: w, height: h, backgroundColor: style.fill }} />);
      if (cell && cell.v) {
        const [first, last] = overflowSpan(cell, style, row, merges, { c0: Math.max(0, win.c0 - 20), c1: win.c1 + 20 });
        const sx = cols.start(first) - ox;
        const sw = cols.end(last) - cols.start(first);
        if ((first !== c || last !== c) && showGrid) {
          // Spilled text hides the gridlines it covers.
          blanks.push(<View key={`s${r}_${c}`} style={{ position: "absolute", left: sx + 1, top: y + 1, width: sw - 1, height: h - 1, backgroundColor: "#FFFFFF" }} />);
        }
        cellText(`t${r}_${c}`, cell, style, { x: sx, y, w: sw, h, cellW: w }, props, texts);
      }
      const comment = comments.get(key(r, c));
      if (comment) marks.push(commentMark(`m${r}_${c}`, x + w, y, zoom));
    }
  }

  for (const merge of merges.intersecting(win)) {
    const [r1, c1, r2, c2] = merge;
    const x = cols.start(c1) - ox;
    const y = rows.start(r1) - oy;
    const w = cols.end(c2) - cols.start(c1);
    const h = rows.end(r2) - rows.start(r1);
    if (w <= 0 || h <= 0) continue;
    const cell = cells.get(r1)?.get(c1);
    const style = styles.get(styleAt(r1, c1, cell));
    fills.push(<View key={`mf${r1}_${c1}`} style={{ position: "absolute", left: x + (showGrid ? 1 : 0), top: y + (showGrid ? 1 : 0), width: w - (showGrid ? 1 : 0), height: h - (showGrid ? 1 : 0), backgroundColor: style?.fill ?? "#FFFFFF" }} />);
    if (cell && cell.v) cellText(`mt${r1}_${c1}`, cell, style, { x, y, w, h, cellW: w }, props, texts);
    const comment = comments.get(key(r1, c1));
    if (comment) marks.push(commentMark(`mm${r1}_${c1}`, x + w, y, zoom));
  }

  if (showHidden) {
    for (let r = win.r0; r <= win.r1; r++) {
      if (rows.hidden[r]) marks.push(<View key={`hr${r}`} pointerEvents="none" style={{ position: "absolute", left, top: rows.start(r) - oy, width: right - left, height: rows.size(r), backgroundColor: "rgba(120,120,120,0.18)" }} />);
    }
    for (let c = win.c0; c <= win.c1; c++) {
      if (cols.hidden[c]) marks.push(<View key={`hc${c}`} pointerEvents="none" style={{ position: "absolute", left: cols.start(c) - ox, top, width: cols.size(c), height: bottom - top, backgroundColor: "rgba(120,120,120,0.18)" }} />);
    }
  }

  if (selected && selected[2] >= win.r0 && selected[0] <= win.r1 && selected[3] >= win.c0 && selected[1] <= win.c1) {
    const x = cols.start(selected[1]) - ox;
    const y = rows.start(selected[0]) - oy;
    marks.push(
      <View
        key="sel"
        pointerEvents="none"
        style={{ position: "absolute", left: x - 1, top: y - 1, width: cols.end(selected[3]) - cols.start(selected[1]) + 2, height: rows.end(selected[2]) - rows.start(selected[0]) + 2, borderWidth: 2, borderColor: SELECT_COLOR }}
      />,
    );
  }
  void layout;
  return (
    <>
      {grid}
      {blanks}
      {fills}
      {texts}
      {borders}
      {marks}
    </>
  );
}

function commentMark(id: string, x: number, y: number, zoom: number): ReactElement {
  const size = Math.max(4, 6 * zoom);
  return (
    <View
      key={id}
      pointerEvents="none"
      style={{ position: "absolute", left: x - size, top: y, width: 0, height: 0, borderTopWidth: size, borderLeftWidth: size, borderTopColor: "#E81123", borderLeftColor: "transparent" }}
    />
  );
}

export const SheetLayer = memo(SheetLayerImpl);
