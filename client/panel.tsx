import type { PluginTheme } from "@getpaseo/plugin";
import type { PluginFilePanelProps } from "@getpaseo/plugin/client";
import { useQueryClient } from "@tanstack/react-query";
import { useCallback, useEffect, useMemo, useState } from "react";
import { ActivityIndicator, Platform, Pressable, ScrollView, Text, View } from "react-native";
import { initialSheet, type Range, type SheetLayout, type WorkbookMeta } from "../shared/model";
import { STALE_REVISION } from "../shared/rpc";
import { useCellRows, usePrefetchSheets, useSheetLayout, useWorkbook } from "./data";
import { columnLabel, Grid, type GridOptions } from "./grid";
import { openExternal } from "./web";

// Plugin panels get no bottom safe-area inset, and Android draws edge to edge, so the bottom of the
// panel sits under the navigation bar on phones. Compact layouts keep controls at the top and leave
// room below the last row; 48 is Android's three-button bar height and covers the iOS home indicator.
const NATIVE_BOTTOM_CLEARANCE = 48;

const ZOOM_STEPS = [0.25, 0.33, 0.4, 0.5, 0.6, 0.75, 0.85, 1, 1.15, 1.3, 1.5, 1.75, 2, 2.5, 3];

interface TabState {
  sheet?: number;
  /** Quick review (true) or the formatted view (false). */
  lite?: boolean;
}

function readTabState(state: unknown): TabState {
  if (typeof state !== "object" || state === null) return {};
  const { sheet, lite } = state as { sheet?: unknown; lite?: unknown };
  const out: TabState = {};
  if (typeof sheet === "number" && Number.isInteger(sheet) && sheet >= 0) out.sheet = sheet;
  if (typeof lite === "boolean") out.lite = lite;
  return out;
}

// The mode last picked in any tab; new tabs open in it. Quick review until the user picks otherwise.
let lastLite = true;

function errorText(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  return message || "파일을 열지 못했습니다.";
}

export function ExcelFilePanel({ file, workspaceId, theme, layout, tab }: PluginFilePanelProps) {
  const fileKey = useMemo(() => ({ workspaceId, path: file.path }), [workspaceId, file.path]);
  const [version, setVersion] = useState(0);
  const queryClient = useQueryClient();
  useEffect(() => file.watch((event) => {
    if (event.type === "changed") setVersion((value) => value + 1);
  }), [file]);
  const workbook = useWorkbook(fileKey, version, readTabState(tab.state).sheet);

  // A stale revision means the file changed between requests: reopen it.
  const retryStale = useCallback((error: unknown) => {
    if (errorText(error).includes(STALE_REVISION)) {
      setVersion((value) => value + 1);
      void queryClient.invalidateQueries({ queryKey: ["excel.open", workspaceId, file.path] });
      return true;
    }
    return false;
  }, [queryClient, workspaceId, file.path]);

  const styles = useMemo(() => chrome(theme, layout.compact), [theme, layout.compact]);
  if (workbook.error) {
    return (
      <View style={styles.center}>
        <Text style={styles.title}>Excel 파일을 열 수 없습니다</Text>
        <Text style={styles.muted}>{errorText(workbook.error)}</Text>
        <Pressable accessibilityRole="button" onPress={() => void workbook.refetch()} style={styles.button}>
          <Text style={styles.buttonText}>다시 시도</Text>
        </Pressable>
      </View>
    );
  }
  if (!workbook.data) {
    return (
      <View style={styles.center}>
        <ActivityIndicator color={theme.colors.accent} />
        <Text style={styles.muted}>통합 문서를 여는 중…</Text>
      </View>
    );
  }
  return <WorkbookView key={workbook.data.revision} meta={workbook.data} fileKey={fileKey} theme={theme} compact={layout.compact} tab={tab} onStale={retryStale} />;
}

interface WorkbookViewProps {
  meta: WorkbookMeta;
  fileKey: { workspaceId: string; path: string };
  theme: PluginTheme;
  compact: boolean;
  tab: PluginFilePanelProps["tab"];
  onStale(error: unknown): boolean;
}

function WorkbookView({ meta, fileKey, theme, compact, tab, onStale }: WorkbookViewProps) {
  const styles = useMemo(() => chrome(theme, compact), [theme, compact]);
  const saved = readTabState(tab.state);
  // Same choice as the daemon's, so the sheet that came bundled with the open is the one shown.
  const [sheet, setSheet] = useState(() => initialSheet(meta, saved.sheet));
  // Both modes draw from the same parsed cells, so switching only redraws the screen.
  const [lite, setLite] = useState(() => saved.lite ?? lastLite);
  const [showHidden, setShowHidden] = useState(false);
  const selectSheet = (index: number) => {
    setSheet(index);
    tab.setState({ sheet: index, lite });
  };
  const toggleLite = () => {
    const next = !lite;
    setLite(next);
    lastLite = next;
    tab.setState({ sheet, lite: next });
  };
  const layoutQuery = useSheetLayout(fileKey, meta.revision, sheet);
  useEffect(() => {
    if (layoutQuery.error) onStale(layoutQuery.error);
  }, [layoutQuery.error, onStale]);
  usePrefetchSheets(fileKey, meta, sheet, layoutQuery.data !== undefined);

  const hiddenSheets = meta.sheets.filter((entry) => entry.state !== "visible").length;
  const tabBar = (
    <View style={styles.tabs}>
      <ScrollView horizontal showsHorizontalScrollIndicator={false} contentContainerStyle={{ alignItems: "stretch" }}>
        {meta.sheets.map((entry) => {
          if (entry.state !== "visible" && !showHidden) return null;
          const active = entry.index === sheet;
          return (
            <Pressable
              key={entry.index}
              accessibilityRole="tab"
              accessibilityState={{ selected: active }}
              accessibilityLabel={`${entry.name} 시트${entry.state !== "visible" ? ", 숨김" : ""}`}
              onPress={() => selectSheet(entry.index)}
              style={[styles.tab, active && styles.tabActive]}
            >
              {entry.tabColor ? <View style={{ position: "absolute", left: 0, right: 0, bottom: 0, height: 3, backgroundColor: entry.tabColor }} /> : null}
              <Text style={[styles.tabText, active && styles.tabTextActive, entry.state !== "visible" && styles.tabHidden]} numberOfLines={1}>
                {entry.state === "veryHidden" ? `${entry.name} (매우 숨김)` : entry.state === "hidden" ? `${entry.name} (숨김)` : entry.name}
              </Text>
            </Pressable>
          );
        })}
      </ScrollView>
      {hiddenSheets > 0 && !showHidden ? <Text style={styles.tabNote}>{`숨긴 시트 ${hiddenSheets}`}</Text> : null}
    </View>
  );
  return (
    <View style={styles.root}>
      {tabBar}
      {layoutQuery.data ? (
        <SheetView key={`${sheet}`} lite={lite} onToggleLite={toggleLite} meta={meta} layout={layoutQuery.data} fileKey={fileKey} theme={theme} compact={compact} showHidden={showHidden} setShowHidden={setShowHidden} />
      ) : layoutQuery.error ? (
        <View style={styles.center}>
          <Text style={styles.muted}>{errorText(layoutQuery.error)}</Text>
          <Pressable accessibilityRole="button" onPress={() => void layoutQuery.refetch()} style={styles.button}>
            <Text style={styles.buttonText}>다시 시도</Text>
          </Pressable>
        </View>
      ) : (
        <View style={styles.center}>
          <ActivityIndicator color={theme.colors.accent} />
        </View>
      )}
    </View>
  );
}

interface SheetViewProps {
  lite: boolean;
  onToggleLite(): void;
  meta: WorkbookMeta;
  layout: SheetLayout;
  fileKey: { workspaceId: string; path: string };
  theme: PluginTheme;
  compact: boolean;
  showHidden: boolean;
  setShowHidden(value: boolean): void;
}

function fitZoom(layout: SheetLayout, width: number): number {
  const range = layout.print.area ?? layout.print.pageRange;
  const lastCol = range ? range[3] : Math.max(0, layout.lastCol);
  let total = 0;
  for (let c = range ? range[1] : 0; c <= lastCol; c++) {
    const col = layout.cols.find((entry) => c >= entry.min && c <= entry.max);
    total += col ? (col.hidden ? 0 : col.width) : layout.defaultColWidth;
  }
  if (total <= 0 || width <= 0) return 1;
  return Math.max(0.25, Math.min(3, (width - 40) / total));
}

function SheetView({ lite, onToggleLite, meta, layout, fileKey, theme, compact, showHidden, setShowHidden }: SheetViewProps) {
  const styles = useMemo(() => chrome(theme, compact), [theme, compact]);
  const [width, setWidth] = useState(0);
  const [zoom, setZoom] = useState<number | null>(null);
  const [preview, setPreview] = useState(!lite && layout.view.mode === "pageBreakPreview");
  const [selected, setSelected] = useState<Range | null>(null);
  // Temporary comparison toggle for the single-view text drawing; remove once checked against Excel.
  const [slimText, setSlimText] = useState(true);
  const savedZoom = Math.max(0.1, Math.min(4, layout.view.zoom / 100));
  // Phones open at the saved zoom like desktop; fitting every column made text tiny and put more cells
  // on screen. The zoom label toggles fit-to-width.
  const effectiveZoom = zoom ?? savedZoom;

  const stepZoom = (direction: 1 | -1) => {
    const current = effectiveZoom;
    const next = direction > 0 ? ZOOM_STEPS.find((step) => step > current + 0.001) : [...ZOOM_STEPS].reverse().find((step) => step < current - 0.001);
    setZoom(next ?? current);
  };
  const onZoomWheel = useCallback((delta: number) => {
    setZoom((current) => {
      const base = current ?? effectiveZoom;
      return Math.max(0.25, Math.min(3, Math.round(base * (delta > 0 ? 0.9 : 1.1) * 100) / 100));
    });
  }, [effectiveZoom]);
  const onZoomTo = useCallback((value: number) => setZoom(Math.max(0.25, Math.min(3, value))), []);

  const options: GridOptions = {
    zoom: effectiveZoom,
    showHidden,
    preview,
    showGrid: layout.view.showGrid,
    showHeaders: layout.view.showHeaders && !(compact && width < 360),
    slimText,
    lite,
  };
  const select = useCallback((r: number, c: number) => {
    const merge = layout.merges.find((m) => m[0] === r && m[1] === c);
    setSelected((current) => (current && current[0] === r && current[1] === c ? null : merge ?? [r, c, r, c]));
  }, [layout.merges]);

  const details = selected ? <CellDetails meta={meta} layout={layout} fileKey={fileKey} selected={selected} theme={theme} compact={compact} onClose={() => setSelected(null)} /> : null;
  const hiddenCount = layout.rows.filter((row) => row.hidden).length + layout.cols.filter((col) => col.hidden).reduce((sum, col) => sum + (col.max - col.min + 1), 0);
  return (
    <View style={styles.sheet} onLayout={(event) => setWidth(event.nativeEvent.layout.width)}>
      <View style={styles.toolbar}>
        <ToolButton theme={theme} styles={styles} label={lite ? "빠른 검토" : "서식 보기"} onPress={onToggleLite} hint={lite ? "서식 보기로 전환" : "빠른 검토로 전환"} />
        {lite ? null : <ToolButton theme={theme} styles={styles} label={preview ? "페이지 미리 보기" : "기본 보기"} active={preview} onPress={() => setPreview(!preview)} hint="페이지 나누기 미리 보기 전환" />}
        <ToolButton theme={theme} styles={styles} label={showHidden ? "숨김 표시 중" : hiddenCount > 0 ? `숨김 ${hiddenCount}` : "숨김"} active={showHidden} onPress={() => setShowHidden(!showHidden)} hint="숨긴 행, 열, 시트 표시" />
        {lite ? null : <ToolButton theme={theme} styles={styles} label={slimText ? "새 글자" : "기존 글자"} active={slimText} onPress={() => setSlimText(!slimText)} hint="셀 글자 그리기 방식 전환 (비교용)" />}
        <View style={styles.spacer} />
        <ToolButton theme={theme} styles={styles} label="−" onPress={() => stepZoom(-1)} hint="축소" />
        <Pressable accessibilityRole="button" accessibilityLabel="화면에 맞추기" onPress={() => setZoom(zoom === null ? fitZoom(layout, width) : null)} style={styles.zoomLabel}>
          <Text style={styles.toolText}>{`${Math.round(effectiveZoom * 100)}%`}</Text>
        </Pressable>
        <ToolButton theme={theme} styles={styles} label="+" onPress={() => stepZoom(1)} hint="확대" />
      </View>
      {layout.print.estimated && preview ? (
        <Text style={styles.notice} numberOfLines={1}>자동 페이지 나누기(점선)는 추정값이라 실제 인쇄와 한두 줄 다를 수 있습니다.</Text>
      ) : null}
      {details}
      <View style={{ flex: 1 }}>
        <Grid file={fileKey} meta={meta} layout={layout} options={options} selected={selected} onSelect={select} onZoomWheel={compact ? undefined : onZoomWheel} onZoomTo={onZoomTo} bottomClearance={compact && Platform.OS !== "web" ? NATIVE_BOTTOM_CLEARANCE : 0} />
      </View>
    </View>
  );
}

function ToolButton({ label, active, onPress, hint, styles }: { theme: PluginTheme; styles: ReturnType<typeof chrome>; label: string; active?: boolean; onPress(): void; hint: string }) {
  return (
    <Pressable accessibilityRole="button" accessibilityLabel={hint} accessibilityState={{ selected: !!active }} onPress={onPress} style={[styles.tool, active && styles.toolActive]}>
      <Text style={[styles.toolText, active && styles.toolTextActive]}>{label}</Text>
    </Pressable>
  );
}

function CellDetails({ meta, layout, fileKey, selected, theme, compact, onClose }: { meta: WorkbookMeta; layout: SheetLayout; fileKey: { workspaceId: string; path: string }; selected: Range; theme: PluginTheme; compact: boolean; onClose(): void }) {
  const styles = useMemo(() => chrome(theme, compact), [theme, compact]);
  const [r, c] = selected;
  const cells = useCellAt(fileKey, meta.revision, layout, r);
  const cell = cells?.get(c);
  const comment = layout.comments.find((note) => note.r === r && note.c === c);
  const link = layout.links.find((entry) => entry.r === r && entry.c === c);
  const address = `${columnLabel(c)}${r + 1}${selected[2] !== r || selected[3] !== c ? `:${columnLabel(selected[3])}${selected[2] + 1}` : ""}`;
  return (
    <View style={styles.details}>
      <View style={styles.detailsRow}>
        <Text style={styles.address}>{address}</Text>
        <Text style={styles.value} selectable numberOfLines={compact ? 3 : 2}>{cell?.v || "(빈 셀)"}</Text>
        <Pressable accessibilityRole="button" accessibilityLabel="셀 정보 닫기" onPress={onClose} style={styles.close}>
          <Text style={styles.muted}>✕</Text>
        </Pressable>
      </View>
      {cell?.f && cell.f !== "…" ? <Text style={styles.formula} selectable numberOfLines={2}>{`= ${cell.f}`}</Text> : cell?.f ? <Text style={styles.formula}>수식 셀 (계산된 값 표시)</Text> : null}
      {cell?.raw && cell.raw !== cell.v ? <Text style={styles.mutedSmall} selectable>{`원래 값: ${cell.raw}`}</Text> : null}
      {comment ? <Text style={styles.comment} selectable>{`메모${comment.author ? ` (${comment.author})` : ""}: ${comment.text}`}</Text> : null}
      {link ? (
        <Pressable accessibilityRole="link" onPress={() => { if (!link.internal) void openExternal(link.target); }}>
          <Text style={styles.link} numberOfLines={1}>{link.internal ? `문서 내 링크: ${link.target}` : link.target}</Text>
        </Pressable>
      ) : null}
    </View>
  );
}

function useCellAt(fileKey: { workspaceId: string; path: string }, revision: string, layout: SheetLayout, r: number) {
  const { rows } = useCellRows(fileKey, revision, layout, r, r);
  return rows.get(r);
}

function chrome(theme: PluginTheme, compact: boolean) {
  const c = theme.colors;
  const pad = compact ? 6 : 8;
  return {
    root: { flex: 1, backgroundColor: c.surface0 },
    sheet: { flex: 1, backgroundColor: c.surface0 },
    center: { flex: 1, alignItems: "center" as const, justifyContent: "center" as const, gap: 10, padding: 24, backgroundColor: c.surface0 },
    title: { color: c.foreground, fontSize: compact ? 16 : 18, fontWeight: "600" as const },
    muted: { color: c.foregroundMuted, fontSize: 13, textAlign: "center" as const },
    mutedSmall: { color: c.foregroundMuted, fontSize: 12 },
    button: { paddingHorizontal: 16, paddingVertical: 10, borderRadius: 8, backgroundColor: c.accent },
    buttonText: { color: c.accentForeground, fontSize: 14 },
    toolbar: { flexDirection: "row" as const, alignItems: "center" as const, gap: 6, paddingHorizontal: pad, paddingVertical: compact ? 4 : 6, borderBottomWidth: 1, borderColor: c.border, backgroundColor: c.surface1 },
    spacer: { flex: 1 },
    tool: { paddingHorizontal: compact ? 10 : 10, paddingVertical: compact ? 8 : 5, minWidth: compact ? 36 : 28, alignItems: "center" as const, borderRadius: 6, borderWidth: 1, borderColor: c.border, backgroundColor: c.surface0 },
    toolActive: { backgroundColor: c.accent, borderColor: c.accent },
    toolText: { color: c.foreground, fontSize: 13 },
    toolTextActive: { color: c.accentForeground },
    zoomLabel: { paddingHorizontal: 4, paddingVertical: compact ? 8 : 5, minWidth: 48, alignItems: "center" as const },
    notice: { color: c.foregroundMuted, fontSize: 11, paddingHorizontal: pad, paddingVertical: 3, backgroundColor: c.surface1 },
    tabs: { flexDirection: "row" as const, alignItems: "center" as const, minHeight: compact ? 44 : 34, borderBottomWidth: 1, borderColor: c.border, backgroundColor: c.surface1 },
    tab: { paddingHorizontal: compact ? 14 : 12, justifyContent: "center" as const, minHeight: compact ? 44 : 34, borderRightWidth: 1, borderColor: c.border, maxWidth: 220 },
    tabActive: { backgroundColor: c.surface0 },
    tabText: { color: c.foregroundMuted, fontSize: 13 },
    tabTextActive: { color: c.foreground, fontWeight: "600" as const },
    tabHidden: { fontStyle: "italic" as const },
    tabNote: { color: c.foregroundMuted, fontSize: 11, paddingHorizontal: 8 },
    details: { paddingHorizontal: pad + 2, paddingVertical: 6, gap: 3, borderBottomWidth: 1, borderColor: c.border, backgroundColor: c.surface1 },
    detailsRow: { flexDirection: "row" as const, alignItems: "flex-start" as const, gap: 10 },
    address: { color: c.foregroundMuted, fontSize: 13, fontWeight: "600" as const, minWidth: 48 },
    value: { flex: 1, color: c.foreground, fontSize: 13 },
    close: { paddingHorizontal: 6, paddingVertical: 2 },
    formula: { color: c.foregroundMuted, fontSize: 12, fontFamily: undefined },
    comment: { color: c.foreground, fontSize: 12 },
    link: { color: c.accent, fontSize: 12, textDecorationLine: "underline" as const },
  };
}
