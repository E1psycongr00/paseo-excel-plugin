import { useCallback, useEffect, useMemo, useRef, useState, type ReactElement } from "react";
import {
  Animated,
  type GestureResponderEvent,
  type LayoutChangeEvent,
  type NativeScrollEvent,
  type NativeSyntheticEvent,
  Platform,
  Pressable,
  ScrollView,
  Text,
  View,
} from "react-native";
import type { CellComment, Range, SheetLayout, WorkbookMeta } from "../shared/model";
import { StyleCache } from "./cell-style";
import { type FileKey, useCellRows } from "./data";
import { buildAxes, Merges, sameWindow, visibleWindow, type Window } from "./geometry";
import { Drawings, Pages } from "./overlays";
import { GRID_COLOR, SELECT_COLOR, SheetLayer } from "./sheet-layer";
import { attachWebScroll, isWeb, webScrollTo } from "./web";

export interface GridOptions {
  zoom: number;
  showHidden: boolean;
  preview: boolean;
  showGrid: boolean;
  showHeaders: boolean;
}

interface GridProps {
  file: FileKey;
  meta: WorkbookMeta;
  layout: SheetLayout;
  options: GridOptions;
  selected: Range | null;
  onSelect(r: number, c: number): void;
  onZoomWheel?(delta: number): void;
  /** Extra scroll room below the last row, for content the host draws system bars over. */
  bottomClearance?: number;
}

const HEADER_BG = "#F3F3F3";
const HEADER_TEXT = "#444444";
const HEADER_LINE = "#D0D0D0";
const HEADER_ACTIVE = "#D3E7DA";

export function columnLabel(c: number): string {
  let name = "";
  let n = c + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

export function Grid({ file, meta, layout, options, selected, onSelect, onZoomWheel, bottomClearance = 0 }: GridProps) {
  const { zoom, showHidden, preview, showGrid, showHeaders } = options;
  const { rows, cols, findCol, rowMap } = useMemo(() => buildAxes(layout, zoom, showHidden), [layout, zoom, showHidden]);
  const merges = useMemo(() => new Merges(layout.merges), [layout.merges]);
  const styles = useMemo(() => new StyleCache(meta.styles, zoom), [meta.styles, zoom]);
  const comments = useMemo(() => new Map<number, CellComment>(layout.comments.map((note) => [note.r * 16384 + note.c, note])), [layout.comments]);
  const rowStyle = useCallback((r: number) => rowMap.get(r)?.style ?? null, [rowMap]);
  const colStyle = useCallback((c: number) => findCol(c)?.style ?? null, [findCol]);

  const frozenRows = Math.min(layout.view.frozenRows, rows.count);
  const frozenCols = Math.min(layout.view.frozenCols, cols.count);
  const frozenH = rows.start(frozenRows);
  const frozenW = cols.start(frozenCols);
  const headerFont = Math.max(9, Math.round(11 * Math.min(1.2, Math.max(0.7, zoom))));
  const headerH = showHeaders ? Math.round(headerFont * 1.9) : 0;
  const headerW = showHeaders ? Math.round(headerFont * 0.62 * Math.max(3, String(layout.rowCount).length) + 12) : 0;

  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  const scroll = useRef({ x: 0, y: 0 });
  const scrollX = useRef(new Animated.Value(0)).current;
  const scrollY = useRef(new Animated.Value(0)).current;
  const negX = useMemo(() => Animated.multiply(scrollX, -1), [scrollX]);
  const negY = useMemo(() => Animated.multiply(scrollY, -1), [scrollY]);
  const [win, setWin] = useState<Window>({ r0: 0, r1: 40, c0: 0, c1: 20 });
  const [viewRect, setViewRect] = useState({ x: 0, y: 0, w: 0, h: 0 });

  const update = useCallback((x: number, y: number) => {
    scroll.current = { x, y };
    scrollX.setValue(x);
    scrollY.setValue(y);
    const next = visibleWindow(rows, cols, x, y, Math.max(1, viewport.width), Math.max(1, viewport.height));
    setWin((current) => (sameWindow(current, next) ? current : next));
    setViewRect((current) => (Math.abs(current.x - x) > 200 || Math.abs(current.y - y) > 200 || current.w !== viewport.width ? { x, y, w: viewport.width, h: viewport.height } : current));
  }, [rows, cols, viewport.width, viewport.height, scrollX, scrollY]);

  useEffect(() => {
    update(scroll.current.x, scroll.current.y);
  }, [update]);

  const { rows: cells } = useCellRows(file, meta.revision, layout, win.r0, win.r1);
  const { rows: frozenCells } = useCellRows(file, meta.revision, layout, 0, frozenRows > 0 ? frozenRows - 1 : -1);
  const allCells = useMemo(() => {
    if (frozenRows === 0) return cells;
    const merged = new Map(cells);
    for (const [r, row] of frozenCells) merged.set(r, row);
    return merged;
  }, [cells, frozenCells, frozenRows]);

  // Scroll targets: web uses a DOM element, iOS a 2D ScrollView, Android nested ScrollViews.
  const webNode = useRef<View>(null);
  const outer = useRef<ScrollView>(null);
  const inner = useRef<ScrollView>(null);
  const scrollTo = useCallback((x: number, y: number) => {
    if (webScrollTo(webNode.current, x, y)) return;
    if (Platform.OS === "android") {
      outer.current?.scrollTo({ y, animated: false });
      inner.current?.scrollTo({ x, animated: false });
    } else {
      outer.current?.scrollTo({ x, y, animated: false });
    }
    update(x, y);
  }, [update]);

  useEffect(() => {
    if (!isWeb) return;
    return attachWebScroll(webNode.current, (x, y) => update(x, y), onZoomWheel) ?? undefined;
  }, [update, onZoomWheel]);

  // Restore the saved top-left cell once, and keep the same cell in view across zoom changes.
  const anchor = useRef<{ r: number; c: number } | null>(null);
  const initialized = useRef(false);
  useEffect(() => {
    if (viewport.width === 0) return;
    let target = anchor.current;
    if (!initialized.current) {
      initialized.current = true;
      target = { r: layout.view.topRow, c: layout.view.leftCol };
    }
    if (!target) return;
    const x = Math.max(0, cols.start(target.c) - frozenW);
    const y = Math.max(0, rows.start(target.r) - frozenH);
    requestAnimationFrame(() => scrollTo(x, y));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [zoom, viewport.width > 0]);
  useEffect(() => {
    anchor.current = { r: rows.indexAt(scroll.current.y + frozenH), c: cols.indexAt(scroll.current.x + frozenW) };
  });

  const onScroll = (axis: "x" | "y" | "xy") => (event: NativeSyntheticEvent<NativeScrollEvent>) => {
    const { x, y } = event.nativeEvent.contentOffset;
    update(axis === "y" ? scroll.current.x : x, axis === "x" ? scroll.current.y : y);
  };

  const onLayout = (event: LayoutChangeEvent) => {
    const { width, height } = event.nativeEvent.layout;
    setViewport({ width: Math.max(0, width - headerW), height: Math.max(0, height - headerH) });
  };

  const area = useRef<View>(null);
  const press = (event: GestureResponderEvent) => {
    const { pageX, pageY } = event.nativeEvent;
    area.current?.measureInWindow((ax, ay) => {
      const lx = pageX - ax - headerW;
      const ly = pageY - ay - headerH;
      if (lx < 0 || ly < 0) return;
      const x = lx < frozenW ? lx : lx + scroll.current.x;
      const y = ly < frozenH ? ly : ly + scroll.current.y;
      const r = rows.indexAt(y);
      const c = cols.indexAt(x);
      const merge = merges.at(r, c);
      onSelect(merge ? merge[0] : r, merge ? merge[1] : c);
    });
  };

  const totalW = cols.total;
  const totalH = rows.total;
  const layerBase = { layout, rows, cols, cells: allCells, merges, styles, zoom, showGrid, showHidden, rowStyle, colStyle, comments, selected };
  const mainWin: Window = { r0: Math.max(win.r0, frozenRows), r1: win.r1, c0: Math.max(win.c0, frozenCols), c1: win.c1 };

  const content = (
    <View style={{ width: totalW, height: totalH, backgroundColor: "#FFFFFF" }}>
      {mainWin.r0 <= mainWin.r1 && mainWin.c0 <= mainWin.c1 ? <SheetLayer {...layerBase} win={mainWin} ox={0} oy={0} /> : null}
      <Drawings file={file} revision={meta.revision} drawings={layout.drawings} rows={rows} cols={cols} zoom={zoom} view={viewRect} />
      <Pages layout={layout} rows={rows} cols={cols} zoom={zoom} win={win} preview={preview} />
    </View>
  );

  let scroller: ReactElement;
  if (isWeb) {
    scroller = (
      <View ref={webNode} style={{ position: "absolute", left: headerW, top: headerH, right: 0, bottom: 0, overflow: "scroll" }}>
        {content}
      </View>
    );
  } else if (Platform.OS === "android") {
    scroller = (
      <ScrollView ref={outer} style={{ position: "absolute", left: headerW, top: headerH, right: 0, bottom: 0 }} nestedScrollEnabled onScroll={onScroll("y")} scrollEventThrottle={16}>
        <ScrollView ref={inner} horizontal nestedScrollEnabled onScroll={onScroll("x")} scrollEventThrottle={16}>
          {content}
        </ScrollView>
        {bottomClearance > 0 ? <View style={{ height: bottomClearance }} /> : null}
      </ScrollView>
    );
  } else {
    scroller = (
      <ScrollView
        ref={outer}
        style={{ position: "absolute", left: headerW, top: headerH, right: 0, bottom: 0 }}
        contentContainerStyle={{ width: totalW, height: totalH + bottomClearance }}
        directionalLockEnabled={false}
        onScroll={onScroll("xy")}
        scrollEventThrottle={16}
      >
        {content}
      </ScrollView>
    );
  }

  const selectedRows = selected ? [selected[0], selected[2]] : null;
  const selectedCols = selected ? [selected[1], selected[3]] : null;

  const colHeaders = (from: number, to: number, ox: number): ReactElement[] => {
    const out: ReactElement[] = [];
    for (let c = from; c <= to; c++) {
      const w = cols.size(c);
      if (w === 0) continue;
      const active = selectedCols !== null && c >= selectedCols[0]! && c <= selectedCols[1]!;
      out.push(
        <View key={c} style={{ position: "absolute", left: cols.start(c) - ox, top: 0, width: w, height: headerH, borderRightWidth: 1, borderBottomWidth: 1, borderColor: HEADER_LINE, backgroundColor: active ? HEADER_ACTIVE : cols.hidden[c] ? "#E4E4E4" : HEADER_BG, alignItems: "center", justifyContent: "center", overflow: "hidden" }}>
          <Text style={{ fontSize: headerFont, color: active ? SELECT_COLOR : HEADER_TEXT }} numberOfLines={1}>{columnLabel(c)}</Text>
        </View>,
      );
    }
    return out;
  };
  const rowHeaders = (from: number, to: number, oy: number): ReactElement[] => {
    const out: ReactElement[] = [];
    for (let r = from; r <= to; r++) {
      const h = rows.size(r);
      if (h === 0) continue;
      const active = selectedRows !== null && r >= selectedRows[0]! && r <= selectedRows[1]!;
      out.push(
        <View key={r} style={{ position: "absolute", left: 0, top: rows.start(r) - oy, width: headerW, height: h, borderRightWidth: 1, borderBottomWidth: 1, borderColor: HEADER_LINE, backgroundColor: active ? HEADER_ACTIVE : rows.hidden[r] ? "#E4E4E4" : HEADER_BG, alignItems: "center", justifyContent: "center", overflow: "hidden" }}>
          {h >= headerFont ? <Text style={{ fontSize: headerFont, color: active ? SELECT_COLOR : HEADER_TEXT }} numberOfLines={1}>{r + 1}</Text> : null}
        </View>,
      );
    }
    return out;
  };

  const bandLine = frozenRows > 0 || frozenCols > 0 ? "#9E9E9E" : GRID_COLOR;
  return (
    <View ref={area} style={{ flex: 1, backgroundColor: "#FFFFFF", overflow: "hidden" }} onLayout={onLayout}>
      <Pressable style={{ flex: 1 }} onPress={press} accessibilityLabel={`${layout.name} 시트`}>
        {scroller}

        {/* Frozen rows across the scrolling columns */}
        {frozenRows > 0 ? (
          <View pointerEvents="none" style={{ position: "absolute", top: headerH, left: headerW + frozenW, right: 0, height: frozenH, overflow: "hidden", borderBottomWidth: 1, borderColor: bandLine, backgroundColor: "#FFFFFF" }}>
            <Animated.View style={{ position: "absolute", left: 0, top: 0, width: totalW, height: frozenH, transform: [{ translateX: negX }] }}>
              <SheetLayer {...layerBase} win={{ r0: 0, r1: frozenRows - 1, c0: Math.max(win.c0, frozenCols), c1: win.c1 }} ox={frozenW} oy={0} />
            </Animated.View>
          </View>
        ) : null}
        {/* Frozen columns down the scrolling rows */}
        {frozenCols > 0 ? (
          <View pointerEvents="none" style={{ position: "absolute", top: headerH + frozenH, left: headerW, bottom: 0, width: frozenW, overflow: "hidden", borderRightWidth: 1, borderColor: bandLine, backgroundColor: "#FFFFFF" }}>
            <Animated.View style={{ position: "absolute", left: 0, top: 0, width: frozenW, height: totalH, transform: [{ translateY: negY }] }}>
              <SheetLayer {...layerBase} win={{ r0: Math.max(win.r0, frozenRows), r1: win.r1, c0: 0, c1: frozenCols - 1 }} ox={0} oy={frozenH} />
            </Animated.View>
          </View>
        ) : null}
        {frozenRows > 0 && frozenCols > 0 ? (
          <View pointerEvents="none" style={{ position: "absolute", top: headerH, left: headerW, width: frozenW, height: frozenH, overflow: "hidden", borderRightWidth: 1, borderBottomWidth: 1, borderColor: bandLine, backgroundColor: "#FFFFFF" }}>
            <SheetLayer {...layerBase} win={{ r0: 0, r1: frozenRows - 1, c0: 0, c1: frozenCols - 1 }} ox={0} oy={0} />
          </View>
        ) : null}

        {showHeaders ? (
          <>
            <View pointerEvents="none" style={{ position: "absolute", top: 0, left: headerW, right: 0, height: headerH, overflow: "hidden", backgroundColor: HEADER_BG }}>
              {frozenCols > 0 ? colHeaders(0, frozenCols - 1, 0) : null}
              <View style={{ position: "absolute", left: frozenW, top: 0, right: 0, height: headerH, overflow: "hidden" }}>
                <Animated.View style={{ position: "absolute", left: 0, top: 0, width: totalW, height: headerH, transform: [{ translateX: negX }] }}>
                  {colHeaders(Math.max(win.c0, frozenCols), win.c1, frozenW)}
                </Animated.View>
              </View>
            </View>
            <View pointerEvents="none" style={{ position: "absolute", top: headerH, left: 0, bottom: 0, width: headerW, overflow: "hidden", backgroundColor: HEADER_BG }}>
              {frozenRows > 0 ? rowHeaders(0, frozenRows - 1, 0) : null}
              <View style={{ position: "absolute", top: frozenH, left: 0, bottom: 0, width: headerW, overflow: "hidden" }}>
                <Animated.View style={{ position: "absolute", left: 0, top: 0, width: headerW, height: totalH, transform: [{ translateY: negY }] }}>
                  {rowHeaders(Math.max(win.r0, frozenRows), win.r1, frozenH)}
                </Animated.View>
              </View>
            </View>
            <View pointerEvents="none" style={{ position: "absolute", top: 0, left: 0, width: headerW, height: headerH, backgroundColor: HEADER_BG, borderRightWidth: 1, borderBottomWidth: 1, borderColor: HEADER_LINE }} />
          </>
        ) : null}
      </Pressable>
    </View>
  );
}
