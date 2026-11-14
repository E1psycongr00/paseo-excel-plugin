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
import { type Axis, buildAxes, hiddenRuns, Merges, openedMask, type Reveal, sameWindow, visibleWindow, type Window } from "./geometry";
import { Drawings, Pages } from "./overlays";
import { FOLD_COLOR, folded, GRID_COLOR, LITE_ROW_PX, SELECT_COLOR, SheetLayer } from "./sheet-layer";
import { attachWebScroll, isWeb, webScrollTo } from "./web";

export interface GridOptions {
  zoom: number;
  showHidden: boolean;
  preview: boolean;
  showGrid: boolean;
  showHeaders: boolean;
  /** Temporary: single-view text for cells that clearly fit, to compare with the wrapped drawing. */
  slimText: boolean;
  /** Quick review: plain values on uniform rows, without formatting, pictures or page breaks. */
  lite: boolean;
}

interface GridProps {
  file: FileKey;
  meta: WorkbookMeta;
  layout: SheetLayout;
  options: GridOptions;
  selected: Range | null;
  onSelect(r: number, c: number): void;
  onZoomWheel?(delta: number): void;
  /** Zoom chosen with a two-finger pinch on phones. */
  onZoomTo?(zoom: number): void;
  /** Extra scroll room below the last row, for content the host draws system bars over. */
  bottomClearance?: number;
}

/** Rows loaded beyond the window on each side, so the next chunk is on the device before it scrolls in. */
const PREFETCH_ROWS = 150;
// After scrolling stops, the window grows in steps (half a screen of rows and a quarter screen of
// columns per side each step), so the next scroll starts over tiles that are already drawn.
const IDLE_DELAY_MS = 200;
const IDLE_STEP_MS = 150;
const IDLE_STEPS = 2;

export const MIN_ZOOM = 0.25;
export const MAX_ZOOM = 3;
const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value));

/** Minimum time between drawing more cells for what a pinch brings into view. */
const PINCH_GROW_MS = 80;
const TOP_LEFT = "left top";

const unionWindow = (a: Window, b: Window): Window => ({ r0: Math.min(a.r0, b.r0), r1: Math.max(a.r1, b.r1), c0: Math.min(a.c0, b.c0), c1: Math.max(a.c1, b.c1) });

/** Finger distance and midpoint, in page px. */
interface Touch2 {
  dist: number;
  mx: number;
  my: number;
}

interface Pinch {
  /** Fingers when the zoom on screen was taken; the preview transform is measured from here. */
  base: Touch2;
  last: Touch2;
  /** Fingers of a zoom step that is not on screen yet. */
  inFlight: Touch2 | null;
  ended: boolean;
}

function twoTouches(event: GestureResponderEvent): Touch2 | null {
  const touches = event.nativeEvent.touches;
  if (touches.length < 2) return null;
  const a = touches[0]!;
  const b = touches[1]!;
  return { dist: Math.max(1, Math.hypot(a.pageX - b.pageX, a.pageY - b.pageY)), mx: (a.pageX + b.pageX) / 2, my: (a.pageY + b.pageY) / 2 };
}

const contains = (a: Window, b: Window) => a.r0 <= b.r0 && a.r1 >= b.r1 && a.c0 <= b.c0 && a.c1 >= b.c1;

const HEADER_BG = "#F3F3F3";
const HEADER_TEXT = "#444444";
const HEADER_LINE = "#D0D0D0";
const HEADER_ACTIVE = "#D3E7DA";
const HEADER_HIDDEN = "#E4E4E4";
/** How far from a fold marker a header tap still hits it, in px. */
const FOLD_HIT_PX = 12;

interface Opened {
  rows: ReadonlySet<number>;
  cols: ReadonlySet<number>;
}

const NOTHING_OPENED: Opened = { rows: new Set(), cols: new Set() };

/**
 * What a header tap at `offset` toggles: the folded run whose marker is nearest (within reach), else
 * the opened hidden run under the tap. Returns the run's first index.
 */
function foldTarget(axis: Axis, runs: Map<number, number>, opened: ReadonlySet<number>, offset: number, reach: number): number | null {
  let best: number | null = null;
  let bestDistance = reach;
  for (const start of runs.keys()) {
    if (opened.has(start)) continue;
    const distance = Math.abs(axis.start(start) - offset);
    if (distance <= bestDistance) {
      best = start;
      bestDistance = distance;
    }
  }
  if (best !== null) return best;
  const index = axis.indexAt(offset);
  if (!axis.hidden[index] || axis.size(index) === 0) return null;
  for (const [start, end] of runs) if (index >= start && index <= end && opened.has(start)) return start;
  return null;
}

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

export function Grid({ file, meta, layout, options, selected, onSelect, onZoomWheel, onZoomTo, bottomClearance = 0 }: GridProps) {
  const { zoom, showHidden, preview, showGrid, showHeaders, slimText, lite } = options;
  // Hidden rows and columns open run by run from their header markers; "숨김" opens them all.
  const hiddenAxes = useMemo(() => buildAxes(layout, 1, false), [layout]);
  const rowRuns = useMemo(() => hiddenRuns(hiddenAxes.rows.hidden), [hiddenAxes]);
  const colRuns = useMemo(() => hiddenRuns(hiddenAxes.cols.hidden), [hiddenAxes]);
  const [opened, setOpened] = useState<Opened>(NOTHING_OPENED);
  const reveal = useMemo<boolean | Reveal>(() => {
    if (showHidden) return true;
    if (opened.rows.size === 0 && opened.cols.size === 0) return false;
    const rowMask = openedMask(hiddenAxes.rows.count, rowRuns, opened.rows);
    const colMask = openedMask(hiddenAxes.cols.count, colRuns, opened.cols);
    return { row: (r) => rowMask[r] === 1, col: (c) => colMask[c] === 1 };
  }, [showHidden, opened, hiddenAxes, rowRuns, colRuns]);
  const toggleFold = useCallback((axis: "rows" | "cols", start: number) => {
    setOpened((current) => {
      const next = new Set(current[axis]);
      if (!next.delete(start)) next.add(start);
      return { ...current, [axis]: next };
    });
  }, []);
  const { rows, cols, findCol, rowMap } = useMemo(() => buildAxes(layout, zoom, reveal, lite ? LITE_ROW_PX : undefined), [layout, zoom, reveal, lite]);
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

  const idleTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const windowAxes = useRef<unknown>(null);
  const windowCols = useRef<unknown>(null);
  const update = useCallback((x: number, y: number) => {
    scroll.current = { x, y };
    // Native scroll views drive these values on the UI thread (see nativeScroll); setting them from
    // JS there would pull the headers back to a stale offset.
    if (isWeb) {
      scrollX.setValue(x);
      scrollY.setValue(y);
    }
    const width = Math.max(1, viewport.width);
    const height = Math.max(1, viewport.height);
    const windowAt = (step: number) => {
      const mx = step * width * 0.25;
      const my = step * height * 0.5;
      return visibleWindow(rows, cols, Math.max(0, x - mx), Math.max(0, y - my), width + mx * 2, height + my * 2);
    };
    // A window drawn for other axes (zoom, hidden toggle, resize) is replaced; otherwise a larger
    // window that still covers the screen is kept, so pre-drawn tiles survive small scrolls.
    const sameAxes = windowAxes.current === rows && windowCols.current === cols;
    windowAxes.current = rows;
    windowCols.current = cols;
    const next = windowAt(0);
    setWin((current) => (sameWindow(current, next) || (sameAxes && contains(current, next)) ? current : next));
    setViewRect((current) => (Math.abs(current.x - x) > 200 || Math.abs(current.y - y) > 200 || current.w !== viewport.width ? { x, y, w: viewport.width, h: viewport.height } : current));

    if (idleTimer.current) clearTimeout(idleTimer.current);
    let step = 0;
    const grow = () => {
      step++;
      const wider = windowAt(step);
      setWin((current) => (contains(current, wider) ? current : wider));
      idleTimer.current = step < IDLE_STEPS ? setTimeout(grow, IDLE_STEP_MS) : null;
    };
    idleTimer.current = setTimeout(grow, IDLE_DELAY_MS);
  }, [rows, cols, viewport.width, viewport.height, scrollX, scrollY]);

  useEffect(() => {
    update(scroll.current.x, scroll.current.y);
  }, [update]);
  useEffect(() => () => {
    if (idleTimer.current) clearTimeout(idleTimer.current);
  }, []);

  const { rows: cells } = useCellRows(file, meta.revision, layout, win.r0, win.r1, PREFETCH_ROWS);
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
    const pending = pendingScroll.current;
    if (pending) {
      // A pinch step chose this zoom: keep the pinched spot under the fingers.
      pendingScroll.current = null;
      initialized.current = true;
      requestAnimationFrame(() => {
        scrollTo(pending.x, pending.y);
        pinchLanded();
      });
      return;
    }
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

  // Frozen panes and headers follow the scroll on the UI thread; the JS listener only moves the window.
  const updateRef = useRef(update);
  updateRef.current = update;
  const nativeScroll = useMemo(() => {
    const listener = (axis: "x" | "y" | "xy") => (event: NativeSyntheticEvent<NativeScrollEvent>) => {
      const { x, y } = event.nativeEvent.contentOffset;
      updateRef.current(axis === "y" ? scroll.current.x : x, axis === "x" ? scroll.current.y : y);
    };
    return {
      x: Animated.event([{ nativeEvent: { contentOffset: { x: scrollX } } }], { useNativeDriver: true, listener: listener("x") }),
      y: Animated.event([{ nativeEvent: { contentOffset: { y: scrollY } } }], { useNativeDriver: true, listener: listener("y") }),
      xy: Animated.event([{ nativeEvent: { contentOffset: { x: scrollX, y: scrollY } } }], { useNativeDriver: true, listener: listener("xy") }),
    };
  }, [scrollX, scrollY]);

  const area = useRef<View>(null);
  const areaBox = useRef({ x: 0, y: 0, width: 0, height: 0 });
  const measureArea = () => area.current?.measureInWindow((x, y, width, height) => {
    areaBox.current = { x, y, width, height };
  });
  const onLayout = (event: LayoutChangeEvent) => {
    const { width, height } = event.nativeEvent.layout;
    setViewport({ width: Math.max(0, width - headerW), height: Math.max(0, height - headerH) });
    measureArea();
  };

  // Two-finger pinch zooms and pans on phones. While the fingers move, the sheet, the frozen bands and
  // the headers are each scaled inside their own fixed frames (scrolling axes about the fingers, frozen
  // axes about their edge), and cells the gesture brings into view are drawn at the current zoom, so
  // nothing already drawn moves or redraws. On release the zoom is applied once and the scroll keeps
  // the pinched spot under the fingers. Applying the zoom mid-gesture made the screen shake, because
  // the redrawn sheet and the reset preview do not reach the screen in the same frame.
  const pinchEnabled = !isWeb && onZoomTo !== undefined;
  const pinch = useRef<Pinch | null>(null);
  const pendingScroll = useRef<{ x: number; y: number } | null>(null);
  const [pinching, setPinching] = useState(false);
  const pinchScale = useRef(new Animated.Value(1)).current;
  const mainX = useRef(new Animated.Value(0)).current;
  const mainY = useRef(new Animated.Value(0)).current;
  const bandX = useRef(new Animated.Value(0)).current;
  const bandY = useRef(new Animated.Value(0)).current;
  // The bands and headers share transforms with the natively driven scroll offsets, so their pinch
  // values run on the native driver. The sheet's ran on JS and reached the screen a frame apart from
  // them, which made the body and headers jitter against each other; a no-op native animation moves
  // every pinch value to the native driver so one setValue batch updates them all in the same frame.
  useEffect(() => {
    if (!pinchEnabled) return;
    for (const value of [pinchScale, mainX, mainY, bandX, bandY]) {
      Animated.timing(value, { toValue: value === pinchScale ? 1 : 0, duration: 0, useNativeDriver: true }).start();
    }
  }, [pinchEnabled, pinchScale, mainX, mainY, bandX, bandY]);
  const grownAt = useRef(0);
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;
  const resetPinch = useCallback(() => {
    pinchScale.setValue(1);
    for (const value of [mainX, mainY, bandX, bandY]) value.setValue(0);
  }, [pinchScale, mainX, mainY, bandX, bandY]);
  /** Scale of `touch` against `base` at the zoom on screen, within the zoom limits. */
  const pinchScaleOf = (base: Touch2, touch: Touch2) => clamp((zoomRef.current * touch.dist) / base.dist, MIN_ZOOM, MAX_ZOOM) / zoomRef.current;
  const showPinch = (base: Touch2, touch: Touch2) => {
    const scale = pinchScaleOf(base, touch);
    const box = areaBox.current;
    // Base midpoint in the grid area, and how far the fingers moved since.
    const fx = base.mx - box.x;
    const fy = base.my - box.y;
    const dx = touch.mx - base.mx;
    const dy = touch.my - base.my;
    const { x: sx, y: sy } = scroll.current;
    // A view whose top-left sits at o and scales about it maps p to f + d + s(p - f) when shifted by
    // (f - o)(1 - s) + d. The sheet sits at header - scroll; scrolling bands start a frozen size later.
    const tx = (fx - (headerW - sx)) * (1 - scale) + dx;
    const ty = (fy - (headerH - sy)) * (1 - scale) + dy;
    pinchScale.setValue(scale);
    mainX.setValue(tx);
    mainY.setValue(ty);
    bandX.setValue(tx - frozenW * (1 - scale));
    bandY.setValue(ty - frozenH * (1 - scale));

    // Draw the cells now in view at the current zoom; what is drawn stays, so nothing on screen moves.
    const now = Date.now();
    if (now - grownAt.current < PINCH_GROW_MS) return;
    grownAt.current = now;
    const x0 = fx + (headerW - fx - dx) / scale - headerW + sx;
    const y0 = fy + (headerH - fy - dy) / scale - headerH + sy;
    const next = visibleWindow(rows, cols, Math.max(0, x0), Math.max(0, y0), viewport.width / scale, viewport.height / scale);
    setWin((current) => (contains(current, next) ? current : unionWindow(current, next)));
  };
  /** Applies the gesture so far as real zoom and scroll; the preview transform restarts from here. */
  const pinchStep = (current: Pinch) => {
    const { base, last } = current;
    const z = zoomRef.current;
    const next = clamp(Math.round(z * pinchScaleOf(base, last) * 100) / 100, MIN_ZOOM, MAX_ZOOM);
    const k = next / z;
    // The content point under the base midpoint lands under the latest midpoint.
    const box = areaBox.current;
    const fx = base.mx - box.x - headerW;
    const fy = base.my - box.y - headerH;
    const dx = last.mx - base.mx;
    const dy = last.my - base.my;
    const x = clamp((scroll.current.x + fx) * k - (fx + dx), 0, Math.max(0, cols.total * k - viewport.width));
    const y = clamp((scroll.current.y + fy) * k - (fy + dy), 0, Math.max(0, rows.total * k - viewport.height + bottomClearance));
    if (next === z) {
      scrollTo(x, y);
      current.base = last;
      if (current.ended) {
        pinch.current = null;
        resetPinch();
      } else showPinch(current.base, current.last);
      return;
    }
    current.inFlight = last;
    pendingScroll.current = { x, y };
    onZoomTo?.(next);
  };
  /** A zoom step is on screen: measure the preview from where that step was taken. */
  const pinchLanded = () => {
    const current = pinch.current;
    if (!current || !current.inFlight) {
      resetPinch();
      return;
    }
    current.base = current.inFlight;
    current.inFlight = null;
    if (current.ended) pinchStep(current);
    else showPinch(current.base, current.last);
  };
  const wantsPinch = (event: GestureResponderEvent) => event.nativeEvent.touches.length >= 2;
  const startPinch = (event: GestureResponderEvent) => {
    const touch = twoTouches(event);
    if (!touch || pinch.current) return;
    measureArea();
    // The window grows from the pinch now; a pending idle step would replace it.
    if (idleTimer.current) {
      clearTimeout(idleTimer.current);
      idleTimer.current = null;
    }
    pinch.current = { base: touch, last: touch, inFlight: null, ended: false };
    setPinching(true);
  };
  const movePinch = (event: GestureResponderEvent) => {
    const touch = twoTouches(event);
    if (!touch) return;
    const current = pinch.current;
    if (!current || current.ended) {
      if (!current) startPinch(event);
      return;
    }
    current.last = touch;
    showPinch(current.base, touch);
  };
  const endPinch = () => {
    const current = pinch.current;
    setPinching(false);
    if (!current) {
      resetPinch();
      return;
    }
    current.ended = true;
    // With a step still on its way, the last one runs when it lands.
    if (!current.inFlight) pinchStep(current);
  };
  const pinchHandlers = pinchEnabled
    ? {
      onStartShouldSetResponderCapture: wantsPinch,
      onMoveShouldSetResponderCapture: wantsPinch,
      // Returning true keeps the native scroll views from taking the gesture back.
      onResponderGrant: (event: GestureResponderEvent) => {
        startPinch(event);
        return true;
      },
      onResponderMove: movePinch,
      onResponderRelease: endPinch,
      onResponderTerminate: endPinch,
      onResponderTerminationRequest: () => false,
    }
    : {};
  const press = (event: GestureResponderEvent) => {
    const { pageX, pageY } = event.nativeEvent;
    area.current?.measureInWindow((ax, ay) => {
      const lx = pageX - ax - headerW;
      const ly = pageY - ay - headerH;
      if (lx < 0 && ly < 0) return;
      const x = lx < frozenW ? lx : lx + scroll.current.x;
      const y = ly < frozenH ? ly : ly + scroll.current.y;
      // A header tap opens or closes the hidden run at its marker.
      if (ly < 0 || lx < 0) {
        if (showHidden) return;
        const axis = ly < 0 ? "cols" : "rows";
        const start = ly < 0 ? foldTarget(cols, colRuns, opened.cols, x, FOLD_HIT_PX) : foldTarget(rows, rowRuns, opened.rows, y, FOLD_HIT_PX);
        if (start !== null) toggleFold(axis, start);
        return;
      }
      const r = rows.indexAt(y);
      const c = cols.indexAt(x);
      const merge = merges.at(r, c);
      onSelect(merge ? merge[0] : r, merge ? merge[1] : c);
    });
  };

  const totalW = cols.total;
  const totalH = rows.total;
  const layerBase = { layout, rows, cols, cells: allCells, merges, styles, zoom, showGrid, showHidden, rowStyle, colStyle, comments, selected, slimText, lite };
  const mainWin = useMemo<Window>(() => ({ r0: Math.max(win.r0, frozenRows), r1: win.r1, c0: Math.max(win.c0, frozenCols), c1: win.c1 }), [win, frozenRows, frozenCols]);
  const topWin = useMemo<Window>(() => ({ r0: 0, r1: frozenRows - 1, c0: Math.max(win.c0, frozenCols), c1: win.c1 }), [win.c0, win.c1, frozenRows, frozenCols]);
  const sideWin = useMemo<Window>(() => ({ r0: Math.max(win.r0, frozenRows), r1: win.r1, c0: 0, c1: frozenCols - 1 }), [win.r0, win.r1, frozenRows, frozenCols]);
  const cornerWin = useMemo<Window>(() => ({ r0: 0, r1: frozenRows - 1, c0: 0, c1: frozenCols - 1 }), [frozenRows, frozenCols]);

  const content = (
    <Animated.View style={{ width: totalW, height: totalH, backgroundColor: "#FFFFFF", transformOrigin: TOP_LEFT, transform: [{ translateX: mainX }, { translateY: mainY }, { scale: pinchScale }] }}>
      {mainWin.r0 <= mainWin.r1 && mainWin.c0 <= mainWin.c1 ? <SheetLayer {...layerBase} win={mainWin} ox={0} oy={0} /> : null}
      {lite ? null : <Drawings file={file} revision={meta.revision} drawings={layout.drawings} rows={rows} cols={cols} zoom={zoom} view={viewRect} />}
      {lite ? null : <Pages layout={layout} rows={rows} cols={cols} zoom={zoom} win={win} preview={preview} />}
    </Animated.View>
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
      <Animated.ScrollView ref={outer} style={{ position: "absolute", left: headerW, top: headerH, right: 0, bottom: 0 }} scrollEnabled={!pinching} nestedScrollEnabled onScroll={nativeScroll.y} scrollEventThrottle={16}>
        <Animated.ScrollView ref={inner} horizontal scrollEnabled={!pinching} nestedScrollEnabled onScroll={nativeScroll.x} scrollEventThrottle={16}>
          {content}
        </Animated.ScrollView>
        {bottomClearance > 0 ? <View style={{ height: bottomClearance }} /> : null}
      </Animated.ScrollView>
    );
  } else {
    scroller = (
      <Animated.ScrollView
        ref={outer}
        style={{ position: "absolute", left: headerW, top: headerH, right: 0, bottom: 0 }}
        contentContainerStyle={{ width: totalW, height: totalH + bottomClearance }}
        directionalLockEnabled={false}
        scrollEnabled={!pinching}
        onScroll={nativeScroll.xy}
        scrollEventThrottle={16}
      >
        {content}
      </Animated.ScrollView>
    );
  }

  const selectedRows = selected ? [selected[0], selected[2]] : null;
  const selectedCols = selected ? [selected[1], selected[3]] : null;

  const colHeaders = (from: number, to: number, ox: number): ReactElement[] => {
    const out: ReactElement[] = [];
    for (let c = from; c <= to; c++) {
      const w = cols.size(c);
      if (folded(cols, c) && !folded(cols, c - 1)) {
        out.push(<View key={`z${c}`} style={{ position: "absolute", left: cols.start(c) - ox - 2, top: 0, width: 4, height: headerH, backgroundColor: FOLD_COLOR, borderRadius: 2, zIndex: 1 }} />);
      }
      if (w === 0) continue;
      const active = selectedCols !== null && c >= selectedCols[0]! && c <= selectedCols[1]!;
      out.push(
        <View key={c} style={{ position: "absolute", left: cols.start(c) - ox, top: 0, width: w, height: headerH, borderRightWidth: 1, borderBottomWidth: 1, borderColor: HEADER_LINE, backgroundColor: active ? HEADER_ACTIVE : cols.hidden[c] ? HEADER_HIDDEN : HEADER_BG, alignItems: "center", justifyContent: "center", overflow: "hidden" }}>
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
      if (folded(rows, r) && !folded(rows, r - 1)) {
        out.push(<View key={`z${r}`} style={{ position: "absolute", left: 0, top: rows.start(r) - oy - 2, width: headerW, height: 4, backgroundColor: FOLD_COLOR, borderRadius: 2, zIndex: 1 }} />);
      }
      if (h === 0) continue;
      const active = selectedRows !== null && r >= selectedRows[0]! && r <= selectedRows[1]!;
      out.push(
        <View key={r} style={{ position: "absolute", left: 0, top: rows.start(r) - oy, width: headerW, height: h, borderRightWidth: 1, borderBottomWidth: 1, borderColor: HEADER_LINE, backgroundColor: active ? HEADER_ACTIVE : rows.hidden[r] ? HEADER_HIDDEN : HEADER_BG, alignItems: "center", justifyContent: "center", overflow: "hidden" }}>
          {h >= headerFont ? <Text style={{ fontSize: headerFont, color: active ? SELECT_COLOR : HEADER_TEXT }} numberOfLines={1}>{r + 1}</Text> : null}
        </View>,
      );
    }
    return out;
  };

  const bandLine = frozenRows > 0 || frozenCols > 0 ? "#9E9E9E" : GRID_COLOR;
  return (
    <View ref={area} style={{ flex: 1, backgroundColor: "#FFFFFF", overflow: "hidden" }} onLayout={onLayout} {...pinchHandlers}>
      <Pressable style={{ flex: 1 }} onPress={press} accessibilityLabel={`${layout.name} 시트`}>
        {scroller}

        {/* Frozen rows across the scrolling columns */}
        {frozenRows > 0 ? (
          <View pointerEvents="none" style={{ position: "absolute", top: headerH, left: headerW + frozenW, right: 0, height: frozenH, overflow: "hidden", borderBottomWidth: 1, borderColor: bandLine, backgroundColor: "#FFFFFF" }}>
            <Animated.View style={{ position: "absolute", left: 0, top: 0, width: totalW, height: frozenH, transformOrigin: TOP_LEFT, transform: [{ translateX: negX }, { translateX: bandX }, { scale: pinchScale }] }}>
              <SheetLayer {...layerBase} win={topWin} ox={frozenW} oy={0} />
            </Animated.View>
          </View>
        ) : null}
        {/* Frozen columns down the scrolling rows */}
        {frozenCols > 0 ? (
          <View pointerEvents="none" style={{ position: "absolute", top: headerH + frozenH, left: headerW, bottom: 0, width: frozenW, overflow: "hidden", borderRightWidth: 1, borderColor: bandLine, backgroundColor: "#FFFFFF" }}>
            <Animated.View style={{ position: "absolute", left: 0, top: 0, width: frozenW, height: totalH, transformOrigin: TOP_LEFT, transform: [{ translateY: negY }, { translateY: bandY }, { scale: pinchScale }] }}>
              <SheetLayer {...layerBase} win={sideWin} ox={0} oy={frozenH} />
            </Animated.View>
          </View>
        ) : null}
        {frozenRows > 0 && frozenCols > 0 ? (
          <View pointerEvents="none" style={{ position: "absolute", top: headerH, left: headerW, width: frozenW, height: frozenH, overflow: "hidden", borderRightWidth: 1, borderBottomWidth: 1, borderColor: bandLine, backgroundColor: "#FFFFFF" }}>
            <Animated.View style={{ position: "absolute", left: 0, top: 0, width: frozenW, height: frozenH, transformOrigin: TOP_LEFT, transform: [{ scale: pinchScale }] }}>
              <SheetLayer {...layerBase} win={cornerWin} ox={0} oy={0} />
            </Animated.View>
          </View>
        ) : null}

        {showHeaders ? (
          <>
            <View pointerEvents="none" style={{ position: "absolute", top: 0, left: headerW, right: 0, height: headerH, overflow: "hidden", backgroundColor: HEADER_BG }}>
              {frozenCols > 0 ? (
                <Animated.View style={{ position: "absolute", left: 0, top: 0, width: frozenW, height: headerH, transformOrigin: TOP_LEFT, transform: [{ scale: pinchScale }] }}>
                  {colHeaders(0, frozenCols - 1, 0)}
                </Animated.View>
              ) : null}
              <View style={{ position: "absolute", left: frozenW, top: 0, right: 0, height: headerH, overflow: "hidden" }}>
                <Animated.View style={{ position: "absolute", left: 0, top: 0, width: totalW, height: headerH, transformOrigin: TOP_LEFT, transform: [{ translateX: negX }, { translateX: bandX }, { scale: pinchScale }] }}>
                  {colHeaders(Math.max(win.c0, frozenCols), win.c1, frozenW)}
                </Animated.View>
              </View>
            </View>
            <View pointerEvents="none" style={{ position: "absolute", top: headerH, left: 0, bottom: 0, width: headerW, overflow: "hidden", backgroundColor: HEADER_BG }}>
              {frozenRows > 0 ? (
                <Animated.View style={{ position: "absolute", left: 0, top: 0, width: headerW, height: frozenH, transformOrigin: TOP_LEFT, transform: [{ scale: pinchScale }] }}>
                  {rowHeaders(0, frozenRows - 1, 0)}
                </Animated.View>
              ) : null}
              <View style={{ position: "absolute", top: frozenH, left: 0, bottom: 0, width: headerW, overflow: "hidden" }}>
                <Animated.View style={{ position: "absolute", left: 0, top: 0, width: headerW, height: totalH, transformOrigin: TOP_LEFT, transform: [{ translateY: negY }, { translateY: bandY }, { scale: pinchScale }] }}>
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
