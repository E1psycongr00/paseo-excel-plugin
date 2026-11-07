import { Linking, Platform } from "react-native";

// This plugin typechecks without the DOM library. Declare only what this module uses.
declare const window: { open(url: string, target: string, features: string): unknown };

interface ScrollNode {
  scrollLeft: number;
  scrollTop: number;
  clientWidth: number;
  clientHeight: number;
  addEventListener(type: string, listener: (event: { ctrlKey?: boolean; deltaY?: number; preventDefault?(): void }) => void, options?: { passive?: boolean }): void;
  removeEventListener(type: string, listener: (event: never) => void): void;
}

const isScrollNode = (node: unknown): node is ScrollNode =>
  typeof node === "object" && node !== null && "addEventListener" in node && "scrollTop" in node;

export const isWeb = Platform.OS === "web";

export async function openExternal(url: string): Promise<void> {
  if (Platform.OS === "web") {
    window.open(url, "_blank", "noopener,noreferrer");
    return;
  }
  await Linking.openURL(url);
}

/**
 * On web a single element scrolls both axes with real scrollbars. React Native Web's View ref is
 * the DOM element. Returns an unsubscribe function, or null off web.
 */
export function attachWebScroll(
  node: unknown,
  onScroll: (x: number, y: number) => void,
  onZoomWheel?: (deltaY: number) => void,
): (() => void) | null {
  if (!isWeb || !isScrollNode(node)) return null;
  let frame = 0;
  const scroll = () => {
    if (frame) return;
    frame = requestAnimationFrame(() => {
      frame = 0;
      onScroll(node.scrollLeft, node.scrollTop);
    });
  };
  const wheel = (event: { ctrlKey?: boolean; deltaY?: number; preventDefault?(): void }) => {
    if (!event.ctrlKey || !onZoomWheel) return;
    event.preventDefault?.();
    onZoomWheel(event.deltaY ?? 0);
  };
  node.addEventListener("scroll", scroll, { passive: true });
  node.addEventListener("wheel", wheel, { passive: false });
  return () => {
    if (frame) cancelAnimationFrame(frame);
    node.removeEventListener("scroll", scroll as never);
    node.removeEventListener("wheel", wheel as never);
  };
}

export function webScrollTo(node: unknown, x: number, y: number): boolean {
  if (!isWeb || !isScrollNode(node)) return false;
  node.scrollLeft = x;
  node.scrollTop = y;
  return true;
}
