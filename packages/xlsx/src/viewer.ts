import {
  XlsxWorkbook,
  acquireXlsxWorksheet,
  acquireXlsxWorksheetPreview,
  retainXlsxWorksheetReference,
  loadXlsxSheetSource,
  prepareXlsxViewerRowHeights,
  releaseXlsxViewerProjection,
  retainXlsxViewerFonts,
} from './workbook.js';
import type { LoadOptions } from './workbook.js';
import type { Cell, Hyperlink, Row, ViewportRange, Worksheet, XlsxChromeColors, XlsxComment } from './types.js';
import type { FindHighlightColors, HyperlinkTarget, FindMatch, FindMatchesOptions, OoxmlResourceMetrics, ViewerContextMenuEvent, ZoomableViewer } from '@silurus/ooxml-core';
import { nextVisibleIndex, resolveVisibleIndex, countVisible, zoomStepScale, anchoredZoomOffset, openExternalHyperlink, nextZoomStep, prevZoomStep, fitScale } from '@silurus/ooxml-core';
import {
  CallerCanvasMount,
  resolveCanvasViewerMode,
  type CanvasViewerRenderMode,
} from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import {
  HEADER_W,
  HEADER_H,
  pxToColWidth,
  pxToRowHeight,
  invalidateAutoRowHeights,
  invalidateSheetRenderCache,
  getGridGeometryForWorksheet,
  rtlMirrorX,
} from './renderer.js';
import { formatA1, parseA1 } from './a1.js';
import { inheritWorksheetPreviewBounds } from './internal/worksheet-content-bounds.js';
import { viewportPreviewBlocker, type ViewportPreviewBlocker } from './internal/worksheet-preview-eligibility.js';
import { resolveXlsxInternalHyperlink } from './internal-hyperlink.js';
import type {
  CellAddress,
  XlsxSelectionArea,
  XlsxSelectionContext,
  XlsxSelectionContextCell,
  XlsxSelectionContextOptions,
  XlsxElementContext,
  XlsxSelectionInput,
  XlsxSelectionState,
} from './selection.js';
import {
  hitTestXlsxElementContext,
  limitXlsxElementContext,
  projectXlsxElementContext,
  type XlsxElementHitViewport,
} from './element-context.js';
import {
  MAX_SELECTION_CONTEXT_CELLS,
  MAX_SELECTION_CONTEXT_TEXT_CHARACTERS,
  areaContainsCell,
  normalizeSelectionState,
  selectionCoordinateCountUpperBound,
  selectionStateFromReference,
  selectionStatesEqual,
} from './selection.js';
export type { CellAddress } from './selection.js';
import { XlsxFindController, type FindCell, type XlsxMatchLocation } from './find.js';
import type { XlsxCommentsOptions } from './comment-card.js';
import { withViewerRenderContext } from './worker-protocol.js';
import { SheetViewEdits } from './internal/viewer/sheet-view-edits.js';
import { OutlineGutter } from './internal/viewer/outline-gutter.js';
import { SheetTabBar } from './internal/viewer/sheet-tab-bar.js';
import { ZoomControl } from './internal/viewer/zoom-control.js';
import { ValidationPanel } from './internal/viewer/validation-panel.js';
import {
  COMMENT_POPUP_MAX_H,
  COMMENT_POPUP_MAX_W,
  CommentPopup,
  createCommentMap,
} from './internal/viewer/comment-popup.js';
import type { OutlineAxis } from './outline.js';
import {
  GridGeometry,
  MAX_WORKSHEET_COL,
  MAX_WORKSHEET_ROW,
} from './internal/grid-geometry.js';
import {
  SheetAcquisition,
  SheetRenderDispatcher,
  SelectionController,
  ViewportState,
  createSheetViewModel,
  type SheetSelectionMode,
} from './internal/sheet-viewer-runtime.js';
import { CanvasSurface, SheetOverlayHost } from './internal/sheet-surface.js';
import { withXlsxRenderCommitGuard } from './render-orchestrator.js';
import { selectionAutoScrollVelocity } from './selection-auto-scroll.js';
import { worksheetContentBounds } from './internal/worksheet-content-bounds.js';
import type { XlsxSheetLoadOptions } from './delimited-text.js';

export type { XlsxSheetLoadOptions } from './delimited-text.js';

const borrowedWorkbookOption = Symbol('XlsxViewer.borrowedWorkbook');
/** @internal Shared source-loading hook for the two XLSX viewer facades. */
const loadXlsxViewerSource = Symbol('XlsxViewer.loadSource');
// Re-exported for the existing xlsx zoom tests (resize-zoom.test.ts imports it
// from this module) and any consumer that referenced it here before it moved to
// @silurus/ooxml-core. The single source of truth is core (design §5.2).
export { zoomStepScale } from '@silurus/ooxml-core';

/** Max width of the list-validation dropdown panel (CSS px). */
const VALIDATION_PANEL_MAX_W = 240;
/** Max height before the value list scrolls (CSS px). */
const VALIDATION_PANEL_MAX_H = 200;

let nextViewerProjectionId = 1;

/** How {@link XlsxViewer} presents hidden sheets (`<sheet state>`, §18.2.19). */
export type HiddenSheetMode = 'show' | 'skip' | 'dim';

/** Marker attribute on the single injected viewer stylesheet, so the module-
 *  level injector is idempotent and destroy() can leave it in place. */
const VIEWER_STYLE_ATTR = 'data-xlsx-viewer-styles';

/** Class-constant CSS shared by every XlsxViewer: it styles pseudo-elements
 *  (scrollbar, slider track/thumb) that inline `element.style` cannot reach, so
 *  it must live in a stylesheet rather than on the elements. */
const VIEWER_STYLE_CSS =
  `.xlsx-tab-strip::-webkit-scrollbar{display:none}` +
  // The viewport remains focusable so copy shortcuts belong to the active
  // Viewer. Focus stays quiet by default; consumers can opt into a keyboard
  // focus ring without conflating it with the selected-cell border.
  `[data-xlsx-viewport-input]:focus{outline:none}` +
  `[data-xlsx-viewport-input]:focus-visible{outline:2px solid var(--ooxml-xlsx-focus-ring,transparent);outline-offset:-2px}` +
  `.xlsx-tab-nav{background:transparent;transition:background 0.1s;}` +
  `.xlsx-tab-nav:hover{background:color-mix(in srgb,var(--ooxml-xlsx-chrome-text,#444) 8%,transparent);}` +
  // Excel-status-bar zoom slider: a thin uniform gray track (no colored
  // fill on either side of the thumb) with a small round gray handle.
  `.xlsx-zoom-slider{-webkit-appearance:none;appearance:none;background:transparent;height:15px;margin:0;}` +
  `.xlsx-zoom-slider::-webkit-slider-runnable-track{height:4px;background:var(--ooxml-xlsx-chrome-border,#c4c4c4);border-radius:2px;}` +
  `.xlsx-zoom-slider::-webkit-slider-thumb{-webkit-appearance:none;appearance:none;width:12px;height:12px;margin-top:-4px;border-radius:50%;background:var(--ooxml-xlsx-chrome-text-muted,#808080);cursor:pointer;}` +
  `.xlsx-zoom-slider:hover::-webkit-slider-thumb{background:var(--ooxml-xlsx-chrome-text,#5f5f5f);}` +
  `.xlsx-zoom-slider::-moz-range-track{height:4px;background:var(--ooxml-xlsx-chrome-border,#c4c4c4);border-radius:2px;}` +
  `.xlsx-zoom-slider::-moz-range-thumb{width:12px;height:12px;border:none;border-radius:50%;background:var(--ooxml-xlsx-chrome-text-muted,#808080);cursor:pointer;}`;

/**
 * Inject the shared viewer stylesheet into one owning document exactly once,
 * keyed by the {@link VIEWER_STYLE_ATTR} marker. Earlier this ran
 * per-instance, so every mount/unmount cycle leaked another `<style>` into the
 * head (unbounded growth). It is deliberately NEVER removed on destroy: the CSS
 * is a class constant that any still-live viewer may depend on, and a single
 * leftover `<style>` after the last teardown is harmless (a fixed, bounded cost,
 * not a per-instance leak).
 */
function ensureViewerStyleInjected(ownerDocument: Document): void {
  if (!ownerDocument.head) return;
  if (ownerDocument.head.querySelector(`style[${VIEWER_STYLE_ATTR}]`)) return;
  const style = ownerDocument.createElement('style');
  style.setAttribute(VIEWER_STYLE_ATTR, '');
  style.textContent = VIEWER_STYLE_CSS;
  ownerDocument.head.appendChild(style);
}

const XLSX_CHROME_COLOR_PROPERTIES = {
  background: '--ooxml-xlsx-chrome-background',
  surface: '--ooxml-xlsx-chrome-surface',
  mutedSurface: '--ooxml-xlsx-chrome-surface-muted',
  text: '--ooxml-xlsx-chrome-text',
  mutedText: '--ooxml-xlsx-chrome-text-muted',
  border: '--ooxml-xlsx-chrome-border',
  selectedSurface: '--ooxml-xlsx-chrome-selection-background',
  accent: '--ooxml-xlsx-chrome-accent',
} as const satisfies Record<keyof XlsxChromeColors, string>;

function sameChromeColors(left: XlsxChromeColors, right: XlsxChromeColors): boolean {
  return Object.keys(XLSX_CHROME_COLOR_PROPERTIES).every((key) =>
    left[key as keyof XlsxChromeColors] === right[key as keyof XlsxChromeColors]);
}

export interface XlsxSheetViewerOptions extends LoadOptions {
  /** Adaptive decoded-raster memory policy for visible worksheet paints. */
  imageResources?: import('@silurus/ooxml-core').ImageResourceOptions;
  /** Scale factor for cell/header dimensions (default 1). 0.5 = half size. */
  cellScale?: number;
  /**
   * Enable drag-to-resize of column widths / row heights by dragging header
   * borders. Resizing only changes the on-screen view — it never modifies the
   * loaded file. Default: true.
   */
  resizable?: boolean;
  /**
   * Show native horizontal and vertical scrollbars for the worksheet viewport.
   * Default: true. Wheel/trackpad panning remains available when explicitly
   * disabled.
   */
  showScrollbars?: boolean;
  /** Lower/upper bounds for the zoom slider as scale factors. Default 0.1–4
   *  (10%–400%, matching Excel's zoom range). Also the clamp range for the IX9
   *  {@link ZoomableViewer} zoom contract ({@link XlsxViewer.setScale} etc.). */
  zoomMin?: number;
  zoomMax?: number;
  /**
   * IX9 — fires whenever the zoom factor actually changes (`1` = 100%), whatever
   * the source: {@link XlsxViewer.setScale}, {@link XlsxViewer.zoomIn} /
   * {@link XlsxViewer.zoomOut}, {@link XlsxViewer.fitWidth} /
   * {@link XlsxViewer.fitPage}, the built-in zoom slider, the +/- buttons, or a
   * Ctrl/⌘+wheel gesture. Named `onScaleChange` to match the docx/pptx viewers so
   * all five share one notification shape. Not fired when a call resolves to the
   * same (clamped/snapped) scale.
   */
  onScaleChange?: (scale: number) => void;
  onReady?: (sheetNames: string[]) => void;
  /**
   * Called when the active sheet changes, with the new sheet's zero-based
   * `index` and the `total` number of sheets in the workbook. This mirrors the
   * docx `onPageChange` and pptx `onSlideChange` contracts so all three viewers
   * share one callback shape. To get the sheet *name*, look it up by index from
   * `viewer.sheetNames[index]` (or the `sheetNames` array delivered to
   * `onReady`).
   */
  onSheetChange?: (index: number, total: number) => void;
  /**
   * Receives asynchronous Viewer-managed failures that cannot be observed by
   * awaiting the method that started them. Failures from `load()`, including
   * its initial render, always reject that Promise and are not also delivered
   * here. Later event-driven render failures invoke this callback, or fall back
   * to `console.error` when omitted.
   *
   * Stable cases can be narrowed with `OoxmlError`,
   * `OoxmlResourceLimitError`, or `OoxmlDecodedImageLimitError` re-exported by
   * this package. Other failures remain `Error` values; do not parse message
   * text as an API. A `code` of `parser-crashed` identifies a recognized WASM
   * trap, not a reliably classified OOM.
   */
  onError?: (err: Error) => void;
  /** Called with the canonical selection state whenever it actually changes. */
  onSelectionStateChange?: (selection: XlsxSelectionState | null) => void;
  /**
   * Called with a bounded, detached read-only context after selection changes.
   * Rapid changes are coalesced to one notification per animation frame. Use
   * `onSelectionStateChange` instead when canonical UI geometry is required.
   */
  onSelectionContextChange?: (context: XlsxSelectionContext | null) => void;
  /**
   * Called synchronously for a browser `contextmenu` event. The original event
   * can suppress the native menu; `getContext()` resolves the range or element
   * context established at the event target.
   */
  onContextMenu?: (event: ViewerContextMenuEvent<XlsxSelectionContext>) => void;
  /**
   * Enable read-only selection of rendered charts, pictures, and shapes. The
   * selected object exposes element context and receives a non-editable outline.
   * Default false; hit-testing runs only for pointer clicks when enabled.
   */
  enableElementSelection?: boolean;
  /**
   * IX1 (design decision — NOT user-confirmed, integrator may veto). Fires when a
   * cell carrying a hyperlink (ECMA-376 §18.3.1.47) is clicked. Default when
   * omitted: external → {@link openExternalHyperlink} (new tab, sanitised,
   * noopener); internal (`location`) → navigate to the referenced sheet/cell
   * when resolvable. When supplied, this callback fully owns the behaviour and
   * receives the raw {@link HyperlinkTarget} verbatim (URL sanitisation is the
   * default handler's job, so a blocked scheme still reaches a custom callback).
   */
  onHyperlinkClick?: (target: HyperlinkTarget) => void;
  /** IX1 — master switch for hyperlink interactivity. Default `true`. When
   *  `false`, the cell hit-test reports no hyperlink under any cell, so hyperlink
   *  interactivity is disabled entirely: no pointer cursor over a link, no default
   *  navigation (external new-tab / internal sheet jump), and `onHyperlinkClick`
   *  is never called. Hyperlinked cells still render exactly as authored but are
   *  inert. */
  enableHyperlinks?: boolean;
  /**
   * Color of the cell-selection highlight. A single CSS color drives both the
   * selection rectangle's border (drawn in this color) and its fill (the same
   * color made translucent — see {@link selectionOverlayStyle}), so callers pick
   * one accent color instead of a separate border + background. Any CSS color
   * string works (`#1a73e8`, `rgb(...)`, `tomato`, …). Default `#1a73e8`
   * (Google blue), matching the historical look. Can also be changed at runtime
   * via {@link XlsxViewer.setSelectionColor}.
   */
  selectionColor?: string;
  /** CSS backgrounds for ordinary and active in-document search matches. */
  findHighlightColors?: FindHighlightColors;
  /**
   * Show authored cell notes and threaded comments. Pass options to configure
   * resolved-thread visibility. Default true.
   */
  comments?: boolean | XlsxCommentsOptions;
  /**
   * `'main'` (default): parse in a worker, render on the main thread. `'worker'`:
   * parse AND render entirely inside the worker and paint the returned
   * ImageBitmap onto the viewer's canvas, so document rendering never blocks the
   * UI thread. All interaction (scroll, sheet tabs, frozen panes, zoom, cell
   * selection) is unchanged. Requires `Worker` + `OffscreenCanvas`. Built-in
   * math and chart renderers are reconstructed inside the worker from their
   * serializable stable identities.
   */
  /**
   * How hidden / veryHidden sheets (`<sheet state>`, ECMA-376 §18.2.19) are
   * presented:
   * - `'show'` (default): every sheet gets a tab — current behavior.
   * - `'skip'`: hidden/veryHidden sheets get no tab and are jumped over by
   *   `nextSheet`/`prevSheet` and initial load; absolute indices are unchanged,
   *   and an explicit `goToSheet(i)` to a hidden sheet is still honored.
   * - `'dim'`: hidden/veryHidden tabs are shown greyed but stay selectable.
   *
   * Named to match the {@link XlsxViewer.hiddenSheetMode} getter and
   * {@link XlsxViewer.setHiddenSheetMode} setter. Mirrors pptx `hiddenSlideMode`.
   */
  hiddenSheetMode?: HiddenSheetMode;
  /** Called after viewport movement with logical CSS-pixel offsets. */
  onViewportChange?: (offset: XlsxViewportOffset) => void;
}

export interface XlsxViewerOptions extends XlsxSheetViewerOptions {
  /** Show the Excel-style zoom slider at the right end of the sheet-tab bar.
   *  Default `true`. Set `false` to hide it (e.g. when the host supplies its
   *  own zoom control). */
  showZoomSlider?: boolean;
}

type InternalXlsxViewerOptions = (XlsxViewerOptions | XlsxSheetViewerOptions) & {
  [borrowedWorkbookOption]?: XlsxWorkbook;
};

export interface XlsxViewportOffset {
  /** Horizontal CSS-pixel offset from the logical start edge (column A side). */
  readonly x: number;
  /** Vertical CSS-pixel offset from the top of the sheet. */
  readonly y: number;
}

/** Cell bounds in CSS pixels relative to the worksheet viewport's top-left.
 * Values may extend outside the visible viewport for an off-screen cell. */
export interface XlsxCellViewportRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface XlsxScrollToCellOptions {
  readonly align?: 'nearest' | 'start' | 'center' | 'end';
}

export type XlsxCopyResult =
  | Readonly<{ status: 'copied'; cellCount: number; utf16CodeUnits: number }>
  | Readonly<{ status: 'empty-selection' }>
  | Readonly<{ status: 'unsupported-multiple-areas' }>
  | Readonly<{ status: 'too-large'; limit: 'cells' | 'text' }>
  | Readonly<{ status: 'clipboard-unavailable' }>
  | Readonly<{ status: 'clipboard-denied' }>;

type SelectionInterval = Readonly<{ first: number; last: number }>;

function mergeSelectionIntervals(intervals: readonly SelectionInterval[]): SelectionInterval[] {
  const sorted = [...intervals].sort((a, b) => a.first - b.first || a.last - b.last);
  const merged: SelectionInterval[] = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous || interval.first > previous.last + 1) {
      merged.push({ ...interval });
    } else if (interval.last > previous.last) {
      merged[merged.length - 1] = { first: previous.first, last: interval.last };
    }
  }
  return merged;
}

function intervalContains(intervals: readonly SelectionInterval[], value: number): boolean {
  let low = 0;
  let high = intervals.length - 1;
  while (low <= high) {
    const middle = (low + high) >>> 1;
    const interval = intervals[middle];
    if (value < interval.first) high = middle - 1;
    else if (value > interval.last) low = middle + 1;
    else return true;
  }
  return false;
}

function lowerBoundBy<T>(items: readonly T[], value: number, key: (item: T) => number): number {
  let low = 0;
  let high = items.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (key(items[middle]) < value) low = middle + 1;
    else high = middle;
  }
  return low;
}

function orderedBy<T>(items: readonly T[], key: (item: T) => number): readonly T[] {
  for (let index = 1; index < items.length; index++) {
    if (key(items[index - 1]) > key(items[index])) {
      return [...items].sort((left, right) => key(left) - key(right));
    }
  }
  return items;
}

/** Default cell-selection accent (Google blue), used when no `selectionColor`
 *  option is supplied. */
const DEFAULT_SELECTION_COLOR = '#1a73e8';

/** Half-width (CSS px) of the grab zone around a header border for
 *  drag-to-resize (issue #567), and the minimum size a column/row can be
 *  dragged to (logical px) so a collapsed band keeps a grabbable border. */
const RESIZE_GRAB_PX = 4;
const RESIZE_MIN_PX = 5;
// Keep clipboard materialization within the same hard cell-count envelope as a
// worksheet. A sparse range can span billions of coordinates even when only a
// handful of cells are populated, so its rectangular TSV must never be built.
const MAX_CLIPBOARD_CELLS = 250_000;
// Bound the retained TSV and its final joined copy. This is a resource-safety
// contract, not a worksheet semantic limit; callers can handle `too-large`
// without the viewer attempting an unbounded JavaScript string allocation.
const MAX_CLIPBOARD_UTF16_CODE_UNITS = 8 * 1_024 * 1_024;
const DEFAULT_SELECTION_CONTEXT_TEXT_CHARACTERS = 1 * 1_024 * 1_024;
const DEFAULT_SELECTION_CONTEXT_NOTIFICATION_TEXT_CHARACTERS = 65_536;
const MAX_SELECTION_CONTEXT_FIELD_CHARACTERS = 65_536;
const MAX_REENTRANT_SELECTION_NOTIFICATIONS = 100;

function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, Math.max(0, maxCodeUnits));
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

function encodeTsvFieldWithin(value: string, remaining: number): string | null {
  let quoteCount = 0;
  let needsQuotes = false;
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    if (code === 34) { quoteCount++; needsQuotes = true; }
    else if (code === 9 || code === 10 || code === 13) needsQuotes = true;
  }
  const length = value.length + (needsQuotes ? quoteCount + 2 : 0);
  if (length > remaining) return null;
  return needsQuotes ? `"${value.replace(/"/g, '""')}"` : value;
}

/**
 * Pure hit predicate for drag-to-resize (issue #567): given a pointer
 * coordinate `pt` (in the header-strip's CSS-px axis — already RTL-un-mirrored
 * by the caller) and the candidate band trailing edges `edges`, return the band
 * index whose edge is within `grabPx` of `pt`, or `null` if none qualifies.
 *
 * `edges` is the candidate list the caller builds — for the band the pointer is
 * over (`hit`) Excel lets you resize the band whose *trailing* border you grab,
 * so the caller passes both `hit - 1` and `hit` (the neighbour-to-the-far-side
 * and the band itself); the first edge within the grab zone wins, in the order
 * given. An edge that sits at or under the header strip (`edge <= headerExtent`,
 * i.e. scrolled behind the frozen corner) is rejected — you can't grab a border
 * hidden under the header. Kept pure (no DOM, no `this`) so the off-by-one
 * geometry — exact-on-edge, within-grab, just-outside, `[hit-1, hit]` neighbour
 * selection, header rejection — is unit-testable. {@link XlsxViewer.getResizeTarget}
 * does the DOM/geometry and calls this.
 */
export function resizeHitIndex(
  pt: number,
  edges: { index: number; edge: number }[],
  grabPx: number,
  headerExtent: number,
): number | null {
  for (const { index, edge } of edges) {
    if (edge <= headerExtent) continue; // scrolled behind the header strip
    if (Math.abs(pt - edge) <= grabPx) return index;
  }
  return null;
}

/**
 * Derive the selection rectangle's `border` and `background` CSS from a single
 * accent color: the border is the color verbatim and the fill is the same color
 * at 8% opacity via `color-mix`, so any CSS color string (`#rgb`, `rgb(...)`,
 * named) yields a matching translucent fill without the caller computing an
 * rgba. For the default `#1a73e8` this reproduces the historical
 * `rgba(26,115,232,0.08)` fill.
 */
export function selectionOverlayStyle(color: string): { border: string; background: string } {
  return {
    border: `2px solid ${color}`,
    background: `color-mix(in srgb, ${color} 8%, transparent)`,
  };
}

interface SelectionOverlayRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly top: boolean;
  readonly right: boolean;
  readonly bottom: boolean;
  readonly left: boolean;
}

interface SelectionBoundarySegment {
  readonly axis: 'h' | 'v';
  readonly fixed: number;
  readonly start: number;
  readonly end: number;
}

/**
 * Build the single-Area outline from its visible frozen-pane fragments.
 * Splitting collinear edges at every endpoint emits coincident fragment edges
 * only once. Work is bounded by the visible fragment count, not sheet size.
 */
function selectionBoundaryPath(rects: readonly SelectionOverlayRect[]): string {
  const raw: SelectionBoundarySegment[] = [];
  for (const rect of rects) {
    const x2 = rect.x + rect.width;
    const y2 = rect.y + rect.height;
    if (rect.top) raw.push({ axis: 'h', fixed: rect.y, start: rect.x, end: x2 });
    if (rect.right) raw.push({ axis: 'v', fixed: x2, start: rect.y, end: y2 });
    if (rect.bottom) raw.push({ axis: 'h', fixed: y2, start: rect.x, end: x2 });
    if (rect.left) raw.push({ axis: 'v', fixed: rect.x, start: rect.y, end: y2 });
  }

  const groups = new Map<string, SelectionBoundarySegment[]>();
  for (const segment of raw) {
    const key = `${segment.axis}:${segment.fixed}`;
    const group = groups.get(key);
    if (group) group.push(segment);
    else groups.set(key, [segment]);
  }

  const commands: string[] = [];
  for (const segments of groups.values()) {
    const points = [...new Set(segments.flatMap(({ start, end }) => [start, end]))]
      .sort((a, b) => a - b);
    let runStart: number | null = null;
    let runEnd = 0;
    const flush = () => {
      if (runStart === null || runEnd <= runStart) return;
      const { axis, fixed } = segments[0];
      commands.push(axis === 'h'
        ? `M${runStart} ${fixed}H${runEnd}`
        : `M${fixed} ${runStart}V${runEnd}`);
      runStart = null;
    };
    for (let index = 0; index + 1 < points.length; index++) {
      const start = points[index];
      const end = points[index + 1];
      const covered = segments.some((segment) => segment.start < end && segment.end > start);
      if (covered && runStart !== null && start === runEnd) {
        runEnd = end;
      } else {
        flush();
        if (covered) {
          runStart = start;
          runEnd = end;
        }
      }
    }
    flush();
  }
  return commands.join('');
}

let selectionMaskSequence = 0;

const DEFAULT_FIND_HIGHLIGHT = 'color-mix(in srgb, #ffb300 8%, transparent)';
const DEFAULT_FIND_ACTIVE_HIGHLIGHT = 'color-mix(in srgb, #fb8c00 8%, transparent)';

/** Resolve an XLSX find box without altering a caller-provided CSS background. */
export function findHighlightOverlayStyle(
  active: boolean,
  colors: FindHighlightColors = {},
): { border: string; background: string } {
  const accent = active ? '#fb8c00' : '#ffb300';
  const custom = active ? colors.active : colors.match;
  const background = custom ?? (active ? DEFAULT_FIND_ACTIVE_HIGHLIGHT : DEFAULT_FIND_HIGHLIGHT);
  return { border: `2px solid ${custom ?? accent}`, background };
}

type XlsxViewerMount =
  | { readonly kind: 'composite' }
  | {
      readonly kind: 'sheet';
      readonly canvas: HTMLCanvasElement;
      /** Resolved before the caller-owned canvas is reparented. */
      readonly mode: CanvasViewerRenderMode;
    };

class XlsxViewerEngine implements ZoomableViewer {
  private readonly container: HTMLElement;
  /** DOM realm of the mount target. Sheet canvases may belong to a same-origin
   * popup rather than the Window that created the viewer instance. */
  private readonly hostDocument: Document;
  private readonly hostWindow: Window & typeof globalThis;
  private readonly acquisition = new SheetAcquisition();
  private readonly viewport: ViewportState;
  private readonly renderDispatcher: SheetRenderDispatcher;
  /** The single subtree root the constructor appended to the caller's
   *  container. destroy() removes it to return the container to its original
   *  (empty) state. */
  private wrapper!: HTMLDivElement;
  private canvas: HTMLCanvasElement;
  /** Region holding the outline gutters (top/left) and the inset {@link canvasArea}.
   *  When the active sheet has no outlining the gutters collapse to 0 px and this
   *  is a transparent pass-through, so an outline-free sheet lays out identically. */
  private gridRegion!: HTMLDivElement;
  /** Row/column grouping gutters (XL4) beside the grid. */
  private readonly outlineGutter: OutlineGutter;
  /** View-only outline/resize edits, replayed onto every sheet projection. */
  private readonly viewEdits = new SheetViewEdits();
  private readonly projectionId = nextViewerProjectionId++;
  private canvasArea: HTMLDivElement;
  private scrollHost: HTMLDivElement;
  private spacer: HTMLDivElement;
  private readonly surface: CanvasSurface;
  private readonly overlayHost: SheetOverlayHost;
  /** Composite-viewer footer chrome; `null` for sheet mounts, which create no
   *  footer DOM. */
  private readonly sheetTabs: SheetTabBar | null = null;
  private readonly zoomControl: ZoomControl | null = null;
  private currentSheet = 0;
  /** Atomically commits an asynchronously acquired worksheet with its index.
   * Incremented by every navigation and teardown so late acquisitions are no-ops. */
  private sheetRequestGeneration = 0;
  private fontBindingGeneration = 0;
  private fontBinding: Readonly<{ workbook: XlsxWorkbook; release: () => void }> | null = null;
  private _hiddenSheetMode: HiddenSheetMode;
  /** During navigation the outgoing graph stays live for interaction while
   * its lease is released. It can briefly coexist with the incoming graph,
   * so viewer memory can peak at two worksheet models until the swap. */
  private currentWorksheet: Worksheet | null = null;
  private previewCompletion: Promise<Worksheet> | null = null;
  private firstPreviewRender = false;
  /** Counts frames that actually reached the canvas. Completing a pull can
   * supersede the first render before it commits, even when the preview flag
   * has already been cleared by the completion callback. */
  private committedFrameCount = 0;
  private previewPreparedViewport: { width: number; height: number; scale: number } | null = null;
  private previewFallbackReason: ViewportPreviewBlocker | null = null;
  private releaseCurrentWorksheet: (() => void) | null = null;
  /** Authored comments for the selected sheet. Presentation filtering must not
   * erase the application-owned data and selection-context contracts. */
  private currentSourceComments: readonly XlsxComment[] = [];
  /** Latest application-owned comment-list navigation. `scrollToCell()` awaits
   * a render, so an older click must not restore its selection after a newer
   * click or after the current sheet has changed. */
  private commentNavigationGeneration = 0;
  private sourceCommentMap = new Map<string, XlsxComment>();
  /** Viewer-owned projections of workbook-cached worksheets. Only view-mutable
   * size/outline state is copied; immutable cell/content graphs stay shared. */
  private sheetViews = new Map<number, Worksheet>();
  private opts: XlsxViewerOptions;
  private readonly _mountKind: XlsxViewerMount['kind'];
  /** Whether this mount delegates viewport movement to a native scroll host. */
  private readonly _nativeScrollbars: boolean;
  /** 'main' renders on this thread; 'worker' paints worker-produced bitmaps. */
  private readonly _mode: 'main' | 'worker';
  private _borrowed = false;
  /** Workbook for which viewer-local state has been initialized. A borrowed
   * sheet mount defers this work until the caller's first goToSheet(), so it
   * never materializes an unrelated first sheet as a constructor side effect. */
  private preparedWorkbook: XlsxWorkbook | null = null;
  /** Set by {@link destroy} (first line). Guards {@link _reportRenderError} so a
   *  render rejection that lands AFTER teardown is swallowed rather than surfaced
   *  to an `onError` / `console.error` on a dead viewer — parity with the scroll
   *  viewers' `_destroyed` flag. */
  private _destroyed = false;
  private resizeObserver: ResizeObserver | null = null;
  private chromeColors: XlsxChromeColors = {};
  private chromeStyleObserver: MutationObserver | null = null;
  private chromeSchemeMedia: MediaQueryList | null = null;
  private chromeSchemeListener: (() => void) | null = null;
  /** Last offset delivered to onViewportChange. Keeping this in the shared
   *  engine prevents a programmatic scroll followed by the browser's native
   *  scroll event from producing duplicate notifications. */
  private _lastViewportNotification: XlsxViewportOffset | null = null;
  private get anchorCell(): CellAddress | null {
    return this.selectionController.anchor;
  }

  private get activeCell(): CellAddress | null {
    return this.selectionController.active;
  }

  private get selectionMode(): SheetSelectionMode {
    return this.selectionController.mode;
  }

  private get isSelecting(): boolean {
    return this.selectionController.dragging;
  }

  private get selectionPointerId(): number | null {
    return this.selectionController.draggingPointerId;
  }

  /** Claim drag-selection ownership and discard deferred gestures from any
   * other pointer that began before this drag. */
  private beginSelectionDrag(pointerId: number): void {
    if (this.pendingTap?.pointerId !== pointerId) this.pendingTap = null;
    if (this.pendingClick?.pointerId !== pointerId) this.pendingClick = null;
    this.selectionController.beginDrag(pointerId);
  }

  /** Gesture-only pointer anchor for the NEXT `setScale`, in canvasArea-viewport
   *  px (`{ x, y }` from the wheel event, relative to the grid's top-left). Set by
   *  the Ctrl/⌘+wheel handler right before it calls `setScale` so the zoom pivots
   *  on the cursor ("zoom toward the pointer") in BOTH axes, past the fixed
   *  header + frozen-pane lead-in; consumed and cleared by `setScale`. `null` for
   *  every non-gesture source (the public `setScale`, the +/- steppers, the zoom
   *  slider, `fitWidth`/`fitPage`), which keep the historical START-anchored
   *  (top-left) preservation so their behaviour is unchanged. */
  private _pendingZoomAnchor: { x: number; y: number } | null = null;

  // Selection state
  private readonly selectionController = new SelectionController();
  private lastNotifiedSelectionState: XlsxSelectionState | null = null;
  private emittingSelectionChange = false;
  private pendingSelectionChange = false;
  private selectionNotificationScheduled = false;
  private selectionNotificationCount = 0;
  private selectionContextNotificationFrame: number | null = null;
  private selectionContextNotificationMicrotask = false;
  // SpreadsheetML permits explicit row/cell references to appear out of
  // coordinate order. Cache a canonical view once per immutable parsed model
  // so range extraction can use binary search without silently skipping such
  // cells on every subsequent context read.
  private readonly selectionContextRows = new WeakMap<Worksheet, readonly Row[]>();
  private readonly selectionContextCells = new WeakMap<Row, readonly Cell[]>();
  private elementContext: XlsxElementContext | null = null;
  private selectionOverlay: HTMLDivElement;
  /** IX2 — find-highlight overlay (matched-cell boxes). */
  private findOverlay!: HTMLDivElement;
  /** IX2 — find state (matches + active cursor). */
  private _find!: XlsxFindController;
  private keydownHandler: ((e: KeyboardEvent) => void) | null = null;
  // Deferred selection press: committed on pointerup only if the pointer
  // neither moved beyond the tap threshold nor caused a scroll. Used for
  // touch/pen (swipe-to-scroll must not change the cell) and for mouse
  // presses inside the overlay-scrollbar band (a thumb drag must not select
  // the cell underneath).
  private pendingTap:
    | { x: number; y: number; shiftKey: boolean; additiveKey: boolean; pointerId: number }
    | null = null;
  // IX1 — mouse press bookkeeping for hyperlink activation: the down position and
  // the cell under it. On pointerup, if the pointer did not move beyond the tap
  // slop (a genuine click, not a drag-select), a hyperlink on that cell is
  // dispatched. Touch/pen activate through the pendingTap path instead.
  private pendingClick: { x: number; y: number; pointerId: number; cell: CellAddress } | null = null;
  private pendingElementClick:
    | { x: number; y: number; pointerId: number; context: XlsxElementContext }
    | null = null;
  // In-flight column/row resize drag (issue #567). `originScaled` is the fixed
  // LTR edge the resized band grows from (left edge for a column, top for a row)
  // in canvasArea CSS px; `mdw` is captured once so the live px→model-unit
  // conversion is stable across the drag. A resize is a *view-only* adjustment:
  // it mutates the in-memory worksheet's colWidths/rowHeights, never the file.
  private resizeDrag:
    | { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number; pointerId: number }
    | null = null;
  /** Last captured drag-selection pointer, retained while edge scrolling runs. */
  private selectionAutoScrollPointer:
    | { clientX: number; clientY: number; pointerId: number }
    | null = null;
  private selectionAutoScrollFrame: number | null = null;
  private selectionAutoScrollLastTime: number | null = null;

  /** Excel-style hover note for the displayed sheet's comments. */
  private readonly comments: CommentPopup;
  /** IX1 — `"row:col"` → hyperlink for the current sheet, rebuilt on every
   *  showSheet. Keys mirror the renderer's `hyperlinkMap` (1-based row/col, the
   *  first cell of a hyperlink `ref` range per the parser), so a `getCellAt`
   *  {row,col} looks up directly. */
  private hyperlinkMap = new Map<string, Hyperlink>();

  /** List data-validation dropdown arrow and display-only value panel. */
  private readonly validation: ValidationPanel;

  constructor(
    container: HTMLElement,
    opts: XlsxViewerOptions | XlsxSheetViewerOptions = {},
    mount: XlsxViewerMount,
  ) {
    this.container = container;
    this.hostDocument =
      (mount.kind === 'sheet' ? mount.canvas.ownerDocument : container.ownerDocument) ?? document;
    const hostWindow = this.hostDocument.defaultView;
    if (!hostWindow) throw new Error('XlsxViewer requires a document with an active Window');
    this.hostWindow = hostWindow;
    this.opts = opts;
    this._mountKind = mount.kind;
    this._nativeScrollbars = opts.showScrollbars ?? true;
    const borrowedWorkbook = (opts as InternalXlsxViewerOptions)[borrowedWorkbookOption];
    this._borrowed = borrowedWorkbook !== undefined;
    this._mode = mount.kind === 'sheet'
      ? mount.mode
      : resolveCanvasViewerMode('XlsxViewer', opts.mode, borrowedWorkbook);
    this._hiddenSheetMode = opts.hiddenSheetMode ?? 'show';
    this.viewport = new ViewportState(opts.cellScale ?? 1);

    this.wrapper = this.hostDocument.createElement('div');
    this.wrapper.style.cssText =
      `position:relative;width:100%;height:100%;` +
      `background:${mount.kind === 'composite' ? 'var(--ooxml-xlsx-chrome-surface,#fff)' : 'transparent'};` +
      `box-sizing:border-box;font-family:sans-serif;display:flex;flex-direction:column;`;

    // The grid region fills the space above the tab bar. The outline gutters
    // (XL4) sit at its top / left edges and {@link canvasArea} is inset by the
    // gutter extents. With no outlining both extents are 0, so canvasArea covers
    // the whole region exactly as before (byte-identical layout).
    this.gridRegion = this.hostDocument.createElement('div');
    this.gridRegion.style.cssText = `position:relative;flex:1;min-height:0;overflow:hidden;`;

    this.canvasArea = this.hostDocument.createElement('div');
    this.canvasArea.style.cssText = `position:absolute;inset:0;overflow:hidden;`;

    this.canvas = mount.kind === 'sheet' ? mount.canvas : this.hostDocument.createElement('canvas');
    this.canvas.style.cssText = `position:absolute;top:0;left:0;z-index:0;display:block;`;
    this.renderDispatcher = new SheetRenderDispatcher(
      this.canvas,
      this._mode === 'worker',
      this.hostWindow,
    );

    this.scrollHost = this.hostDocument.createElement('div');
    this.scrollHost.setAttribute('data-xlsx-viewport-input', mount.kind);
    this.scrollHost.setAttribute('role', 'region');
    this.scrollHost.setAttribute(
      'aria-label',
      'Spreadsheet viewport. Use Arrow keys to move the selected cell. Press Enter to show its comment.',
    );
    this.scrollHost.tabIndex = 0;
    this.scrollHost.style.cssText =
      `position:absolute;inset:0;` +
      `overflow:${this._nativeScrollbars ? 'auto' : 'clip'};` +
      `z-index:2;background:transparent;` +
      `scrollbar-color:var(--ooxml-xlsx-chrome-scrollbar-color,auto);`;
    this.spacer = this.hostDocument.createElement('div');
    this.spacer.style.cssText = `position:absolute;top:0;left:0;pointer-events:none;`;
    if (this._nativeScrollbars) this.scrollHost.appendChild(this.spacer);
    this.surface = new CanvasSurface(this.canvas, this.canvasArea, this.scrollHost);
    this.overlayHost = new SheetOverlayHost(this.canvasArea, this.canvas, this.scrollHost, {
      commentMaxWidth: COMMENT_POPUP_MAX_W,
      commentMaxHeight: COMMENT_POPUP_MAX_H,
      validationMaxWidth: VALIDATION_PANEL_MAX_W,
      validationMaxHeight: VALIDATION_PANEL_MAX_H,
    });
    this.selectionOverlay = this.overlayHost.selection;
    this.findOverlay = this.overlayHost.find;
    this.comments = new CommentPopup({
      ownerDocument: this.hostDocument,
      canvasArea: this.canvasArea,
      overlayHost: this.overlayHost,
      currentSheet: () => this.currentSheet,
      isRtl: () => this.isRtl,
      isDestroyed: () => this._destroyed,
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
      reportError: (error) => this._reportRenderError(error),
    });
    this.validation = new ValidationPanel({
      ownerDocument: this.hostDocument,
      canvasArea: this.canvasArea,
      surface: this.surface,
      overlayHost: this.overlayHost,
      worksheet: () => this.currentWorksheet,
      workbook: () => this.wb,
      currentSheet: () => this.currentSheet,
      activeCell: () => this.activeCell,
      selectionMode: () => this.selectionMode,
      scale: () => this.viewport.scale,
      isRtl: () => this.isRtl,
      isDestroyed: () => this._destroyed,
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
    });
    this.outlineGutter = new OutlineGutter({
      gridRegion: this.gridRegion,
      canvasArea: this.canvasArea,
      surface: this.surface,
      worksheet: () => this.currentWorksheet,
      scale: () => this.viewport.scale,
      chromeColors: () => this.chromeColors,
      cellRect: (row, col) => this._cellRect(row, col),
      screenX: (x, w) => this.screenX(x, w),
      setBandHidden: (axis, index, hidden) => this.setBandHidden(axis, index, hidden),
      setBandCollapsed: (axis, index, collapsed) => this.setBandCollapsed(axis, index, collapsed),
      afterOutlineMutation: (ws, anchor) => this.afterOutlineMutation(ws, anchor),
    });
    // Inject the shared viewer stylesheet once per module (idempotent). Both
    // mounts use it; the composite footer also hides its tab-strip scrollbar.
    ensureViewerStyleInjected(this.hostDocument);

    if (mount.kind === 'composite') {
      this.sheetTabs = new SheetTabBar(this.hostDocument, {
        hiddenSheetMode: () => this._hiddenSheetMode,
        isHidden: (index) => Boolean(this.wb?.isHidden(index)),
        selectSheet: (index) => {
          void this.goToSheet(index).catch((error) => this._reportRenderError(error));
        },
      });
      if (this.opts.showZoomSlider !== false) {
        this.zoomControl = new ZoomControl(this.hostDocument, {
          setScale: (scale) => this.setScale(scale),
          zoomIn: () => this.zoomIn(),
          zoomOut: () => this.zoomOut(),
        }, this.viewport.scale, this.opts.zoomMin ?? 0.1, this.opts.zoomMax ?? 4);
        this.sheetTabs.append(this.zoomControl.element);
      }
    }

    // canvasArea only — the gutter canvases are attached lazily by
    // layoutGutters when (and only when) the shown sheet actually has an
    // outline, and detached again otherwise. Keeping them OUT of the DOM for
    // outline-free sheets preserves exact element parity with the pre-outline
    // viewer: consumers that count or index `<canvas>` elements (the layouts
    // smoke does `page.locator('canvas').count()`, which includes
    // `display:none` nodes) must see no difference.
    this.gridRegion.appendChild(this.canvasArea);
    this.wrapper.appendChild(this.gridRegion);
    if (this.sheetTabs) this.wrapper.appendChild(this.sheetTabs.tabBar);
    container.appendChild(this.wrapper);
    this.installChromeThemeRefresh();

    // Gutter click handling (XL4): +/- toggles and the numbered level banks
    // (each in its own gutter's header strip; the corner is inert background).
    // Registered once; no-op when a sheet has no gutter (extents 0 ⇒ hidden).
    this.outlineGutter.installListeners();

    if (this._nativeScrollbars) this.surface.on('scroll', () => {
      // Any scroll cancels a deferred tap: the press that started it was a
      // scrollbar-thumb drag (overlay scrollbars) or a touch swipe, not a
      // cell click.
      this.pendingTap = null;
      this.pendingElementClick = null;
      // A comment popup is anchored to a cell's on-screen rect, which moves
      // under the cursor while scrolling — hide it (Excel does the same).
      this.hideCommentPopup();
      // The validation panel is anchored to the cell too; Excel closes its
      // dropdown on scroll, so do the same.
      this.hideValidationPanel();
      // Track the start-anchored position, but only while the host is laid
      // out: a hidden host reports clientWidth 0 and fires bogus scroll
      // events when the browser clamps scrollLeft, which must not overwrite
      // the last real position.
      if (this.scrollHost.clientWidth > 0) {
        const raw = this.scrollHost.scrollLeft;
        const logicalX = this.isRtl ? this.maxScrollLeft - raw : raw;
        this.viewport.setViewportSize(this.scrollHost.clientWidth, this.scrollHost.clientHeight);
        this.viewport.setOffset(logicalX, this.scrollHost.scrollTop);
      }
      this.emitViewportChange();
      // Coalesce into the next frame: a scroll gesture fires many events per
      // frame, and the previous synchronous redraw ran the full render on each
      // one. The overlay update is cheap DOM geometry (no canvas paint) and must
      // track the scroll immediately, so it stays synchronous.
      this.scheduleRender();
      this.updateSelectionOverlay();
      this.updateFindOverlay();
    });

    // Re-render whenever the canvas area changes size. Re-anchor first: a
    // size change shifts maxScrollLeft, and for RTL sheets the native
    // scrollLeft must be re-derived from the start-anchored position or the
    // view drifts (or, after a hidden mount, stays stranded at the far end).
    const resizeObserver = new this.hostWindow.ResizeObserver(() => {
      const offset = { x: this.viewport.x, y: this.viewport.y };
      this.viewport.setViewportSize(this.scrollHost.clientWidth, this.scrollHost.clientHeight);
      this.setViewportLeft(offset.x);
      this.viewportTop = offset.y;
      this.reanchorHorizontalScroll();
      // Re-place the outline gutter strips for the new region size (XL4). This
      // only rewrites styles (no canvasArea size change) so it can't feed back
      // into the observer.
      this.layoutGutters();
      // Container resizes can burst (a live window/pane drag); coalesce the
      // canvas paint into one frame. The re-anchor, overlay and nav updates are
      // cheap and must reflect the new size at once, so they stay synchronous.
      this.scheduleRender();
      this.updateSelectionOverlay();
      this.updateFindOverlay();
      this.sheetTabs?.updateNavButtons();
    });
    resizeObserver.observe(this.gridRegion);
    this.resizeObserver = resizeObserver;

    this.setupSelectionEvents();

    this._find = new XlsxFindController(
      () => this.sheetCount,
      (sheet) => this.wb?.sheetNames[sheet] ?? '',
      (sheet) => this._collectSheetCells(sheet),
    );

    if (borrowedWorkbook) {
      this.acquisition.install(borrowedWorkbook, false);
      if (this._mountKind === 'composite') {
        this.activateWorkbook(borrowedWorkbook).catch((error) => this._reportRenderError(error));
      }
    }

  }

  /**
   * Re-read the CSS custom properties that affect Canvas-painted Viewer chrome.
   * DOM chrome follows inherited CSS variables without help; row/column headers
   * and outline gutters need an explicit repaint because their colors are baked
   * into pixels.
   */
  private refreshChromeTheme(): void {
    if (this._destroyed) return;
    const getComputedStyle = this.hostWindow.getComputedStyle?.bind(this.hostWindow);
    if (!getComputedStyle) return;
    const computed = getComputedStyle(this.wrapper);
    const next: Record<string, string> = {};
    for (const [key, property] of Object.entries(XLSX_CHROME_COLOR_PROPERTIES)) {
      const value = computed.getPropertyValue(property).trim();
      if (value) next[key] = value;
    }
    const nextColors = next as XlsxChromeColors;
    if (sameChromeColors(this.chromeColors, nextColors)) return;
    this.chromeColors = nextColors;
    this.renderGutters();
    this.scheduleRender();
  }

  /** Observe the ordinary ways an application changes theme state. */
  private installChromeThemeRefresh(): void {
    this.refreshChromeTheme();

    const MutationObserverClass = this.hostWindow.MutationObserver ?? globalThis.MutationObserver;
    if (MutationObserverClass) {
      this.chromeStyleObserver = new MutationObserverClass(() => this.refreshChromeTheme());
      for (let target: HTMLElement | null = this.container; target; target = target.parentElement) {
        this.chromeStyleObserver.observe(target, {
          attributes: true,
          attributeFilter: ['class', 'style', 'data-theme'],
        });
      }
    }

    const media = this.hostWindow.matchMedia?.('(prefers-color-scheme: dark)') ?? null;
    if (media) {
      const listener = () => this.refreshChromeTheme();
      media.addEventListener?.('change', listener);
      this.chromeSchemeMedia = media;
      this.chromeSchemeListener = listener;
    }
  }

  /** Every non-empty cell of a sheet with its rendered display text (IX2 find
   *  source). Reads the parsed worksheet model directly — no render — so search
   *  covers the whole sheet, not just the on-screen viewport. */
  private async _collectSheetCells(sheet: number): Promise<FindCell[]> {
    const wb = this.wb;
    if (!wb) return [];
    const lease = await acquireXlsxWorksheet(wb, sheet);
    try {
      const ws = lease.worksheet;
      const cells: FindCell[] = [];
      for (const row of ws.rows) {
        for (const cell of row.cells) {
          const text = wb.cellText(ws, cell);
          if (text !== '') cells.push({ row: cell.row, col: cell.col, text });
        }
      }
      return cells;
    } finally {
      lease.release();
    }
  }

  /**
   * Load an XLSX from URL or ArrayBuffer and render the first sheet.
   *
   * Parse, load, and initial-render failures always reject this Promise.
   * `onError` is reserved for later Viewer-managed work that has no directly
   * awaitable method result, so one failure is never delivered twice.
   */
  async [loadXlsxViewerSource](
    source: string | ArrayBuffer,
    sourceOptions: XlsxSheetLoadOptions = {},
  ): Promise<void> {
    this.assertOpen();
    if (this._borrowed) {
      throw new Error(
        `${this._mountKind === 'sheet' ? 'XlsxSheetViewer' : 'XlsxViewer'}.load() is unsupported ` +
          'on a Viewer created by fromWorkbook(); the borrowed workbook is already loaded.',
      );
    }
    // SC20 atomic swap: retain the previous workbook locally and only tear it down
    // AFTER the new one loads successfully. A re-load thus never orphans the old
    // workbook's worker + pinned WASM allocation (the leak this guards), yet a
    // FAILED re-load keeps the current workbook + its rendered sheet intact rather
    // than dropping to an empty viewer. The 2× memory window is bounded to the
    // load itself (the old workbook is freed the moment the new model arrives).
    try {
      const wb = await this.acquisition.replace(() => XlsxWorkbook[loadXlsxSheetSource](source, {
          password: this.opts.password,
          useGoogleFonts: this.opts.useGoogleFonts,
          cjkFallback: this.opts.cjkFallback,
          maxZipEntryBytes: this.opts.maxZipEntryBytes,
          resourceLimits: this.opts.resourceLimits,
          debug: this.opts.debug,
          onResourceMetrics: this.opts.onResourceMetrics,
          workerTimeoutMs: this.opts.workerTimeoutMs,
          wasmUrl: this.opts.wasmUrl,
          math: this.opts.math,
          threeD: this.opts.threeD,
          regionMap: this.opts.regionMap,
          chartEx: this.opts.chartEx,
          tiff: this.opts.tiff,
          mode: this._mode,
          ...(this.opts.modelSources === undefined ? undefined : { modelSources: this.opts.modelSources }),
        }, sourceOptions), () => {
          // Claim every async-operation generation before closing the old
          // workbook. Rejections caused by its worker termination are stale
          // completion, not errors belonging to the new workbook.
          this.sheetRequestGeneration++;
          this.renderDispatcher.begin();
          this._find.invalidate();
          this.hideValidationPanel();
          this.releaseHostFonts();
        });
      if (!wb) return;
      if (this._destroyed) throw this.destroyedError();
      await this.activateWorkbook(wb);
    } catch (err) {
      if (this._destroyed) throw this.destroyedError();
      throw err instanceof Error ? err : new Error(String(err));
    }
  }

  /** Bind the current acquisition to its independent viewer state. Parsing,
   *  worksheet materialization, archive access, and caches remain workbook-owned. */
  private async activateWorkbook(workbook: XlsxWorkbook, sheetIndex?: number): Promise<void> {
    if (!this.prepareWorkbook(workbook)) return;
    await this.showSheet(sheetIndex ?? this._initialSheet());
  }

  private async ensureHostFonts(workbook: XlsxWorkbook): Promise<boolean> {
    if (this.fontBinding?.workbook === workbook) return true;
    const retain = workbook[retainXlsxViewerFonts];
    // Structural test doubles and pre-feature adapters have no font hook. A
    // real XlsxWorkbook always does; absence means there is nothing to retain.
    if (typeof retain !== 'function') return true;
    const generation = ++this.fontBindingGeneration;
    const release = await retain.call(workbook, this.hostDocument);
    if (
      this._destroyed ||
      generation !== this.fontBindingGeneration ||
      this.wb !== workbook
    ) {
      release();
      return false;
    }
    this.fontBinding?.release();
    this.fontBinding = { workbook, release };
    return true;
  }

  private releaseHostFonts(): void {
    this.fontBindingGeneration++;
    this.fontBinding?.release();
    this.fontBinding = null;
  }

  /** Initialize the viewer-local projection state without choosing a sheet.
   * This split lets a borrowed sheet viewer make goToSheet(index) its first
   * worksheet materialization, while the composite viewer can still open its
   * normal initial sheet automatically. */
  private prepareWorkbook(workbook: XlsxWorkbook): boolean {
    if (this._destroyed || this.wb !== workbook) return false;
    if (this.preparedWorkbook === workbook) return true;
    this._find.invalidate();
    this.viewEdits.clear();
    this.sheetViews.clear();
    this.buildTabs();
    this.preparedWorkbook = workbook;
    this.opts.onReady?.(workbook.sheetNames);
    return true;
  }

  /** The loaded workbook, or throws if {@link load} has not completed. */
  private get workbook(): XlsxWorkbook {
    const workbook = this.acquisition.current;
    if (!workbook) throw new Error('Workbook not loaded');
    return workbook;
  }

  private get wb(): XlsxWorkbook | null {
    return this.acquisition.current;
  }

  /** Internal assignment seam retained for focused viewer-mechanics tests. All
   *  ownership still flows through SheetAcquisition. */
  private set wb(workbook: XlsxWorkbook | null) {
    if (workbook) this.acquisition.install(workbook);
    else this.acquisition.destroy();
  }

  private async showSheet(index: number): Promise<void> {
    const generation = ++this.sheetRequestGeneration;
    this.previewFallbackReason = null;
    const workbook = this.workbook;
    let worksheet: Worksheet;
    let sourceWorksheet: Worksheet;
    let releaseNewWorksheet: (() => void) | undefined;
    let previewCompletion: Promise<Worksheet> | null = null;
    let previewPreparedViewport: { width: number; height: number; scale: number } | null = null;
    try {
      if (!await this.ensureHostFonts(workbook)) return;
      if (!this.isCurrentSheetRequest(generation, workbook)) return;
      if (index !== this.currentSheet && this.currentWorksheet) {
        // Permit cache eviction, but keep the displayed worksheet and every
        // interaction map intact until a replacement is ready to commit.
        this.releaseCurrentWorksheet?.();
        this.releaseCurrentWorksheet = null;
      }
      const lease = await acquireXlsxWorksheetPreview(workbook, index);
      sourceWorksheet = lease.worksheet;
      releaseNewWorksheet = lease.release;
      let eligiblePreview = lease.partial;
      const cachedView = this.sheetViews.get(index);
      const prepareView = (model: Worksheet): Worksheet => {
        const view = cachedView ?? this.createVisibleSheetView(model);
        if (cachedView || lease.partial) view.rows = model.rows;
        if (!cachedView) this.viewEdits.restoreSheetViewState(index, view);
        return view;
      };
      const prepareHeights = (view: Worksheet, refresh: boolean): void => {
        if (refresh) invalidateAutoRowHeights(view);
        const prepareRowHeights = workbook[prepareXlsxViewerRowHeights];
        if (typeof prepareRowHeights === 'function') {
          const measureCanvas = this.hostDocument.createElement('canvas');
          const measureCtx = measureCanvas.getContext('2d');
          if (measureCtx) prepareRowHeights.call(workbook, view, measureCtx);
        }
        this.viewEdits.syncAutomaticRowOverrides(index, view);
      };
      worksheet = prepareView(sourceWorksheet);
      if (lease.partial) {
        const preparedViewport = {
          width: this.canvasArea.clientWidth,
          height: this.canvasArea.clientHeight,
          scale: this.viewport.scale,
        };
        previewPreparedViewport = preparedViewport;
        const visibleRange = () => getGridGeometryForWorksheet(worksheet).visibleRange({
          width: preparedViewport.width,
          height: preparedViewport.height,
          scale: preparedViewport.scale,
          scrollX: 0, scrollY: 0,
          headerWidth: HEADER_W, headerHeight: HEADER_H, buffer: 2,
        });
        let visible = visibleRange();
        let coveringRow = 0;
        for (;;) {
          const needed = Math.max(visible.range.row + visible.range.rows - 1, worksheet.freezeRows ?? 0);
          if (needed > coveringRow) {
            await (lease.waitForRows?.(needed) ?? Promise.resolve());
            coveringRow = needed;
          }
          // Rows that arrived while waiting can change display-derived height
          // and therefore bring additional rows into the first viewport.
          prepareHeights(worksheet, true);
          visible = visibleRange();
          if (Math.max(visible.range.row + visible.range.rows - 1, worksheet.freezeRows ?? 0) <= coveringRow) break;
        }
        // Paint includes the frozen corner, frozen row/column strips, and the
        // scrollable quadrant. Use their union for both row coverage and every
        // dependency check; the renderer can also spill text horizontally from
        // cells outside the visible column band in these rows.
        const painted = {
          row: (worksheet.freezeRows ?? 0) > 0 ? 1 : visible.range.row,
          col: (worksheet.freezeCols ?? 0) > 0 ? 1 : visible.range.col,
          rows: visible.range.row + visible.range.rows - ((worksheet.freezeRows ?? 0) > 0 ? 1 : visible.range.row),
          cols: visible.range.col + visible.range.cols - ((worksheet.freezeCols ?? 0) > 0 ? 1 : visible.range.col),
        };
        this.previewFallbackReason = viewportPreviewBlocker(worksheet, painted, coveringRow);
        if (this.previewFallbackReason) {
          sourceWorksheet = await lease.completion;
          eligiblePreview = false;
          previewPreparedViewport = null;
          worksheet = prepareView(sourceWorksheet);
          prepareHeights(worksheet, true);
        }
      } else prepareHeights(worksheet, false);
      previewCompletion = eligiblePreview ? lease.completion : null;
    } catch (error) {
      releaseNewWorksheet?.();
      if (!this.isCurrentSheetRequest(generation, workbook)) return;
      await this.restoreDisplayedWorksheetLease(workbook, generation);
      throw error;
    }
    if (!this.isCurrentSheetRequest(generation, workbook)) {
      releaseNewWorksheet?.();
      return;
    }

    this.releaseCurrentWorksheet?.();
    this.releaseCurrentWorksheet = releaseNewWorksheet ?? null;
    // Viewer projections share the full cell graph. Keeping inactive entries
    // would defeat workbook eviction even after its cache drops the model.
    this.sheetViews.clear();
    this.sheetViews.set(index, worksheet);
    this.currentSheet = index;
    this.currentWorksheet = worksheet;
    this.previewCompletion = previewCompletion;
    this.firstPreviewRender = previewCompletion !== null;
    this.previewPreparedViewport = previewPreparedViewport;
    if (previewCompletion) {
      void previewCompletion.then((completed) => {
        if (!this.isCurrentSheetRequest(generation, workbook) || this.currentWorksheet !== worksheet) return;
        this.previewCompletion = null;
        this.firstPreviewRender = false;
        this.previewPreparedViewport = null;
        // Chart and sparkline references can resolve more completely once all
        // rows exist. Rebind the viewer to the committed model, preserving only
        // viewer-owned size and outline edits made during the pull.
        const finalized = this.createVisibleSheetView(completed);
        this.viewEdits.restoreSheetViewState(index, finalized);
        this.currentWorksheet = finalized;
        this.sheetViews.set(index, finalized);
        if (completed.parseError) {
          // A later row may make the cursor produce the normal degraded-sheet
          // placeholder. Replace the provisional graph before the next frame.
          this.currentSourceComments = [];
          this.sourceCommentMap.clear();
          this.selectionController.reset();
          this.emitSelectionChange();
          this.updateSelectionOverlay();
          this.buildCommentMap(finalized);
          this.buildHyperlinkMap(finalized);
          this.buildOutline(finalized);
          this.layoutGutters();
          this.updateSpacerSize(finalized);
          this.scheduleRender();
          return;
        }
        invalidateSheetRenderCache(worksheet);
        invalidateAutoRowHeights(worksheet);
        const measureCtx = this.hostDocument.createElement('canvas').getContext('2d');
        if (measureCtx) workbook[prepareXlsxViewerRowHeights](finalized, measureCtx);
        this.viewEdits.syncAutomaticRowOverrides(index, finalized);
        this.currentSourceComments = completed.comments ?? [];
        this.sourceCommentMap = createCommentMap(this.currentSourceComments);
        this.buildCommentMap(finalized);
        this.buildHyperlinkMap(finalized);
        this.buildOutline(finalized);
        this.layoutGutters();
        this.updateSpacerSize(finalized);
        this.scheduleRender();
        this.scheduleSelectionContextNotification();
      }).catch((error: unknown) => {
        if (!this.isCurrentSheetRequest(generation, workbook) || this.currentWorksheet !== worksheet) return;
        this.previewCompletion = null;
        this.firstPreviewRender = false;
        this.previewPreparedViewport = null;
        this.currentWorksheet = null;
        this.releaseCurrentWorksheet?.();
        this.releaseCurrentWorksheet = null;
        this.renderDispatcher.begin();
        if (this._mode === 'worker') {
          // A bitmaprenderer canvas has no 2D context, and resizing it can
          // retain the last transferred frame. Replace it with an empty bitmap.
          const surface = new OffscreenCanvas(1, 1);
          surface.getContext('2d');
          const blank = surface.transferToImageBitmap();
          this.canvas.getContext('bitmaprenderer')?.transferFromImageBitmap(blank);
          blank.close();
        } else {
          this.canvas.getContext('2d')?.clearRect?.(0, 0, this.canvas.width, this.canvas.height);
        }
        this._reportRenderError(error);
      });
    }
    this.currentSourceComments = sourceWorksheet.comments ?? [];
    if (this.opts.comments !== false && this.currentSourceComments.length > 0) {
      void this.comments.loadUi().catch((error) => this._reportRenderError(error));
    }
    this.sourceCommentMap = createCommentMap(this.currentSourceComments);
    this.setElementContext(null);
    this.pendingElementClick = null;
    this.updateFooterDirection();
    this.viewportTop = 0;
    this.selectionController.reset();
    this.emitSelectionChange();
    this.hideCommentPopup();
    this.hideValidationPanel();
    this.updateSelectionOverlay();
    this.updateTabActive(index);
    this.buildCommentMap(this.currentWorksheet);
    this.buildHyperlinkMap(this.currentWorksheet);
    // XL4: build the outline layout for this sheet and size the gutters. Must run
    // before `updateSpacerSize` / render so the inset canvasArea has its final
    // size when the grid geometry is computed.
    this.buildOutline(this.currentWorksheet);
    this.layoutGutters();
    this.updateSpacerSize(this.currentWorksheet);
    // Reset the horizontal scroll origin to the natural START of the sheet.
    // For RTL sheets the start column (col A) lives at the RIGHT, which means
    // the native scrollbar thumb must sit at its right end (max scrollLeft);
    // for LTR sheets the start is scrollLeft=0. updateSpacerSize must run first
    // so scrollWidth reflects the new sheet before we read the max offset.
    this.resetHorizontalScroll();
    const frameBefore = this.committedFrameCount;
    await this.renderCurrentSheet();
    const paintedEarly = this.committedFrameCount > frameBefore;
    if (previewCompletion && !paintedEarly && this.isCurrentSheetRequest(generation, workbook)) {
      await previewCompletion;
      if (!this.isCurrentSheetRequest(generation, workbook)) return;
      await this.renderCurrentSheet();
    }
    if (!this.isCurrentSheetRequest(generation, workbook)) return;
    // Redraw find highlights for the newly shown sheet (the find state survives
    // a sheet switch; only the visible sheet's boxes are drawn).
    this.updateFindOverlay();
    this.emitViewportChange();
    this.opts.onSheetChange?.(index, this.workbook.sheetNames.length);
  }

  private isCurrentSheetRequest(generation: number, workbook: XlsxWorkbook): boolean {
    return !this._destroyed && generation === this.sheetRequestGeneration && this.wb === workbook;
  }

  private async restoreDisplayedWorksheetLease(workbook: XlsxWorkbook, generation: number): Promise<void> {
    if (!this.currentWorksheet || this.releaseCurrentWorksheet) return;
    const release = await retainXlsxWorksheetReference(workbook, this.currentSheet);
    if (this.isCurrentSheetRequest(generation, workbook) && this.currentWorksheet && !this.releaseCurrentWorksheet) {
      this.releaseCurrentWorksheet = release;
    } else {
      release();
    }
  }

  // ─── Outline gutter (XL4: row/column grouping) ────────────────────────────

  /** Recompute the per-axis outline layout for `ws` and bind the sheet's
   *  view-edit stashes. An outline-free sheet collapses both gutters to 0. */
  private buildOutline(ws: Worksheet): void {
    this.viewEdits.bindSheet(this.currentSheet);
    this.outlineGutter.rebuild(ws);
  }

  /** Place the gutters and inset canvasArea by their extents. */
  private layoutGutters(): void {
    this.outlineGutter.layout();
  }

  /** Repaint the gutters for the current scroll offset (after every frame). */
  private renderGutters(): void {
    this.outlineGutter.render();
  }

  /** Align an outline summary band to the scrollable viewport's start without
   * disturbing the perpendicular axis. */
  private scrollOutlineSummaryToStart(axis: OutlineAxis, summary: number): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const cs = this.viewport.scale;
    const offset = getGridGeometryForWorksheet(ws).scrollOffsetForCell(
      axis === 'row' ? summary : 1,
      axis === 'col' ? summary : 1,
      {
        scale: cs,
        viewportWidth: this.canvasArea.clientWidth,
        viewportHeight: this.canvasArea.clientHeight,
        currentX: this.effectiveScrollLeft,
        currentY: this.viewportTop,
        headerWidth: HEADER_W,
        headerHeight: HEADER_H,
        align: 'start',
      },
    );
    if (axis === 'row') this.viewportTop = offset.y;
    else this.setViewportLeft(offset.x);
  }

  private setBandHidden(axis: OutlineAxis, index: number, hidden: boolean): void {
    const ws = this.currentWorksheet;
    if (ws) this.viewEdits.setBandHidden(ws, this.currentSheet, axis, index, hidden);
  }

  private recordSizeOverride(axis: OutlineAxis, index: number): void {
    const ws = this.currentWorksheet;
    if (ws) this.viewEdits.recordSizeOverride(ws, this.currentSheet, axis, index);
  }

  private wireSizeOverrides(): ReturnType<SheetViewEdits['wireSizeOverrides']> {
    return this.viewEdits.wireSizeOverrides(this.currentSheet);
  }

  private setBandCollapsed(axis: OutlineAxis, index: number, collapsed: boolean): void {
    const ws = this.currentWorksheet;
    if (ws) this.viewEdits.setBandCollapsed(ws, this.currentSheet, axis, index, collapsed);
  }

  /** Shared tail of a gutter interaction: invalidate the axis cache, rebuild the
   *  outline (collapsed flags changed), refresh dependent geometry, re-render. */
  private afterOutlineMutation(
    ws: Worksheet,
    anchor?: { axis: OutlineAxis; summary: number },
  ): void {
    GridGeometry.invalidate(ws);
    this.outlineGutter.rebuild(ws);
    this.updateSpacerSize(ws);
    if (anchor) this.scrollOutlineSummaryToStart(anchor.axis, anchor.summary);
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.scheduleRender();
    if (anchor) this.emitViewportChange();
  }

  /** True when the current sheet's grid is laid out right-to-left. */
  private get isRtl(): boolean {
    return this.currentWorksheet?.rightToLeft === true;
  }

  /** Mirror the workbook footer for an RTL sheet (composite mounts only). */
  private updateFooterDirection(): void {
    this.sheetTabs?.setDirection(this.isRtl);
  }

  /** Maximum horizontal logical viewport offset (≥ 0). */
  private get maxScrollLeft(): number {
    this.syncNativeViewportExtent();
    return this.viewport.maxX;
  }

  private get maxScrollTop(): number {
    this.syncNativeViewportExtent();
    return this.viewport.maxY;
  }

  private syncNativeViewportExtent(): void {
    if (!this._nativeScrollbars) return;
    this.viewport.setViewportSize(this.scrollHost.clientWidth, this.scrollHost.clientHeight);
    this.viewport.ensureExtent(this.scrollHost.scrollWidth, this.scrollHost.scrollHeight);
  }

  private get viewportTop(): number {
    if (this._nativeScrollbars) {
      this.syncNativeViewportExtent();
      this.viewport.adoptNativeOffset(this.viewport.x, this.scrollHost.scrollTop);
    }
    return this.viewport.y;
  }

  private set viewportTop(value: number) {
    this.viewport.setOffset(this.viewport.x, value);
    if (this._nativeScrollbars) this.scrollHost.scrollTop = this.viewport.y;
  }

  /**
   * The logical horizontal scroll position used to find the start-of-sheet
   * (col A) edge, in *scaled* CSS pixels — the same unit as
   * `scrollHost.scrollLeft`. The renderer always lays the grid out LTR and then
   * mirrors it (ECMA-376 §18.3.1.87), so the viewer must hand it a position
   * where 0 = the START of the sheet (col A) and increasing values reveal later
   * columns.
   *
   * For LTR that is exactly the native `scrollLeft`. For RTL the sheet starts at
   * the RIGHT, so the native scrollbar runs the opposite way: thumb fully right
   * (`scrollLeft = maxScrollLeft`) is the start, thumb left is the far columns.
   * Inverting here makes wheel/trackpad follow the finger and aligns the
   * thumb↔page mapping with Excel, without depending on browser-specific RTL
   * `scrollLeft` sign conventions.
   */
  private get effectiveScrollLeft(): number {
    if (this._nativeScrollbars) {
      this.syncNativeViewportExtent();
      const raw = this.scrollHost.scrollLeft;
      this.viewport.adoptNativeOffset(this.isRtl ? this.maxScrollLeft - raw : raw, this.viewport.y);
    }
    return this.viewport.x;
  }

  private setViewportLeft(value: number): void {
    this.viewport.setOffset(value, this.viewport.y);
    if (this._nativeScrollbars) {
      this.scrollHost.scrollLeft = this.isRtl
        ? Math.max(0, this.maxScrollLeft - this.viewport.x)
        : this.viewport.x;
    }
  }

  /**
   * Map between the logical-LTR x used by all the cell-geometry math and the
   * on-screen (canvasArea CSS-pixel) x, applying the RTL mirror (ECMA-376
   * §18.3.1.87) via the same {@link rtlMirrorX} the renderer uses. For LTR this
   * is the identity. The mirror is an involution, so this one method serves
   * both cell→px (overlay draw, `w` = cell width) and px→cell (pointer
   * hit-testing, `w` = 0 for a point) — guaranteeing the overlay sits exactly
   * where the cell is drawn and a click resolves to that same cell at every
   * scroll offset. `canvasArea.clientWidth` equals the renderer's `canvasW`.
   */
  private screenX(logicalX: number, w: number): number {
    return this.isRtl ? rtlMirrorX(logicalX, w, this.canvasArea.clientWidth) : logicalX;
  }

  /** Park the scrollbar at the sheet's natural start: scrollLeft=0 for LTR,
   *  the right end for RTL (so col A shows first). */
  private resetHorizontalScroll(): void {
    this.viewport.setOffset(0, this.viewport.y);
    if (this._nativeScrollbars) {
      this.scrollHost.scrollLeft = this.isRtl ? this.maxScrollLeft : 0;
    }
  }

  /** Re-derive the native scrollLeft from the tracked start-anchored
   *  position after the scroll host's size changes. Only RTL needs this:
   *  for LTR the native scrollLeft *is* start-anchored and the browser
   *  already clamps it sensibly on resize. */
  private reanchorHorizontalScroll(): void {
    if (!this._nativeScrollbars) return;
    if (!this.isRtl || this.scrollHost.clientWidth === 0) return;
    const want = Math.max(0, this.maxScrollLeft - this.viewport.x);
    if (Math.abs(this.scrollHost.scrollLeft - want) > 1) {
      this.scrollHost.scrollLeft = want;
    }
  }

  /** 0-based index of the currently displayed sheet. */
  get sheetIndex(): number {
    return this.currentSheet;
  }

  /** Total number of sheets in the loaded workbook. */
  get sheetCount(): number {
    return this.wb?.sheetCount ?? 0;
  }

  /**
   * Navigate to a sheet by index, clamped to range. Canonical navigation verb
   * matching {@link PptxViewer.goToSlide} / {@link DocxViewer.goToPage}.
   */
  async goToSheet(index: number): Promise<void> {
    if (this.sheetCount === 0) return;
    const workbook = this.workbook;
    if (!this.prepareWorkbook(workbook)) return;
    await this.showSheet(Math.max(0, Math.min(index, this.sheetCount - 1)));
  }

  async nextSheet(): Promise<void> {
    await this.goToSheet(this._stepSheet(1));
  }

  async prevSheet(): Promise<void> {
    await this.goToSheet(this._stepSheet(-1));
  }

  /** Logical start-anchored viewport offset in CSS pixels at the current scale. */
  getViewportOffset(): XlsxViewportOffset {
    return {
      x: Math.max(0, this.effectiveScrollLeft),
      y: Math.max(0, this.viewportTop),
    };
  }

  private emitViewportChange(): void {
    const callback = this.opts.onViewportChange;
    if (!callback) return;
    const offset = this.getViewportOffset();
    const previous = this._lastViewportNotification;
    if (previous && previous.x === offset.x && previous.y === offset.y) return;
    this._lastViewportNotification = offset;
    callback(offset);
  }

  /** Move the active sheet viewport without exposing browser RTL scroll rules. */
  async setViewportOffset(offset: XlsxViewportOffset): Promise<void> {
    if (!Number.isFinite(offset.x) || !Number.isFinite(offset.y)) {
      throw new TypeError('XLSX viewport offsets must be finite numbers');
    }
    const x = Math.min(this.maxScrollLeft, Math.max(0, offset.x));
    const y = Math.min(this.maxScrollTop, Math.max(0, offset.y));
    this.setViewportLeft(x);
    this.viewportTop = y;
    await this.renderCurrentSheet();
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.emitViewportChange();
  }

  /** Re-read the mount's CSS box and repaint the current viewport. */
  async relayout(): Promise<void> {
    this.reanchorHorizontalScroll();
    this.layoutGutters();
    if (this.currentWorksheet) this.updateSpacerSize(this.currentWorksheet);
    await this.renderCurrentSheet();
    this.updateSelectionOverlay();
    this.updateFindOverlay();
  }

  async scrollToCell(
    ref: string,
    options: XlsxScrollToCellOptions = {},
  ): Promise<void> {
    const cell = parseA1(ref);
    if (!cell || !this.currentWorksheet) return;
    if (this.previewCompletion) await this.previewCompletion;
    this._scrollCellIntoView(cell.row, cell.col, options.align ?? 'nearest');
    await this.renderCurrentSheet();
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.emitViewportChange();
  }

  /** Next sheet index for sequential nav: skip mode jumps over hidden sheets. */
  private _stepSheet(dir: 1 | -1): number {
    if (this._hiddenSheetMode === 'skip' && this.wb) {
      return nextVisibleIndex(this.currentSheet, dir, (i) => this.wb!.isHidden(i), this.sheetCount);
    }
    return this.currentSheet + dir;
  }

  /** Initial sheet for load() / entering skip mode: land on a visible sheet. */
  private _initialSheet(): number {
    if (this._hiddenSheetMode === 'skip' && this.wb) {
      return resolveVisibleIndex(0, (i) => this.wb!.isHidden(i), this.sheetCount);
    }
    return 0;
  }

  /** Returns the cell at canvas-client coordinates, or null if outside the cell grid. */
  getCellAt(clientX: number, clientY: number): CellAddress | null {
    if (this._destroyed) return null;
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.viewport.scale;

    const rect = this.canvasArea.getBoundingClientRect();
    // Un-mirror the screen x into the logical-LTR layout the geometry below
    // assumes (header on the left). screenX is an involution, so applying it to
    // a screen point recovers the logical point; w = 0 for a point. Done in
    // scaled CSS px (canvasArea space) before converting to logical px.
    const lx = this.screenX(clientX - rect.left, 0);
    const ly = clientY - rect.top;

    const scaledHeaderW = Math.round(HEADER_W * cs);
    const scaledHeaderH = Math.round(HEADER_H * cs);
    if (lx < scaledHeaderW || ly < scaledHeaderH) return null;

    const innerX = lx - scaledHeaderW;
    const innerY = ly - scaledHeaderH;

    return getGridGeometryForWorksheet(ws).cellAt(innerX, innerY, {
      scrollX: this.effectiveScrollLeft,
      scrollY: this.viewportTop,
      scale: cs,
    });
  }

  /** Click-only DrawingML hit test. It walks just the sheet's anchored object
   * arrays and never scans worksheet cells or runs during render/scroll. */
  private elementContextViewport(): XlsxElementHitViewport | null {
    const worksheet = this.currentWorksheet;
    if (!worksheet) return null;
    const width = this.canvasArea.clientWidth;
    const height = this.canvasArea.clientHeight;
    if (width <= 0 || height <= 0) return null;
    const scale = this.viewport.scale;
    const geometry = getGridGeometryForWorksheet(worksheet);
    const visible = geometry.visibleRange({
      width,
      height,
      scale,
      scrollX: this.effectiveScrollLeft,
      scrollY: this.viewportTop,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
      buffer: 2,
    });
    return {
      width,
      height,
      cellScale: scale,
      viewport: visible.range,
      scrollOffsetX: visible.offsetX,
      scrollOffsetY: visible.offsetY,
      freezeRows: worksheet.freezeRows ?? 0,
      freezeCols: worksheet.freezeCols ?? 0,
    };
  }

  private elementContextAt(clientX: number, clientY: number): XlsxElementContext | null {
    if (!this.opts.enableElementSelection || this._destroyed) return null;
    const worksheet = this.currentWorksheet;
    const viewport = this.elementContextViewport();
    if (!worksheet || !viewport) return null;
    const rect = this.canvasArea.getBoundingClientRect();
    return hitTestXlsxElementContext(
      worksheet,
      this.currentSheet,
      { x: clientX - rect.left, y: clientY - rect.top },
      viewport,
    );
  }

  /** Returns the CSS-pixel rect of a cell within canvasArea, or null if not
   *  computable. Mirrors the renderer's per-cell rounding (Math.round(px * cs))
   *  so the selection overlay sits exactly on the canvas's drawn cell borders;
   *  multiplying logical accumulators by `cs` once at the end (the previous
   *  approach) drifted by up to 1 px per cell at non-integer scales.
   */
  private _cellRect(row: number, col: number): { x: number; y: number; w: number; h: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.viewport.scale;
    return getGridGeometryForWorksheet(ws).cellRect(row, col, {
      scale: cs,
      scrollX: this.effectiveScrollLeft,
      scrollY: this.viewportTop,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
    });
  }

  /** Return one cell's viewport-relative CSS-pixel bounds. This is the forward
   * geometry primitive for application-owned comment or annotation overlays. */
  getCellViewportRect(cell: CellAddress | string): XlsxCellViewportRect | null {
    if (this._destroyed) return null;
    const address = typeof cell === 'string' ? parseA1(cell) : cell;
    if (!address || address.row < 1 || address.col < 1) return null;
    const rect = this._cellRect(address.row, address.col);
    return rect
      ? Object.freeze({
          x: this.screenX(rect.x, rect.w),
          y: rect.y,
          width: rect.w,
          height: rect.h,
        })
      : null;
  }

  /** Detached comments for the current sheet, in authored order. */
  getComments(): readonly Readonly<XlsxComment>[] {
    this.assertOpen();
    return structuredClone(this.currentSourceComments);
  }

  /**
   * Reveal and select the cell that owns a comment on an explicit sheet. This deliberately owns
   * no list UI: applications render detached records from `getComments()` and
   * call this navigation primitive from their own rows.
   *
   * Returns `false` when the sheet index is invalid or `cellRef` does not
   * identify a comment on that sheet.
   */
  async goToComment(
    sheetIndex: number,
    cellRef: string,
    options?: XlsxScrollToCellOptions,
  ): Promise<boolean> {
    const target = parseA1(cellRef);
    const workbook = this.wb;
    if (
      !target || !workbook || !Number.isInteger(sheetIndex) ||
      sheetIndex < 0 || sheetIndex >= workbook.sheetCount
    ) {
      return false;
    }
    const generation = ++this.commentNavigationGeneration;
    const comments = sheetIndex === this.currentSheet && this.currentWorksheet !== null
      ? this.currentSourceComments
      : await workbook.getComments(sheetIndex);
    if (this._destroyed) throw this.destroyedError();
    if (generation !== this.commentNavigationGeneration || workbook !== this.wb) return false;
    if (!comments.some((comment) => {
      const cell = parseA1(comment.cellRef);
      return cell?.row === target.row && cell.col === target.col;
    })) return false;

    if (sheetIndex !== this.currentSheet || this.currentWorksheet === null) {
      await this.goToSheet(sheetIndex);
      if (this._destroyed) throw this.destroyedError();
      if (
        generation !== this.commentNavigationGeneration || workbook !== this.wb ||
        sheetIndex !== this.currentSheet
      ) {
        return false;
      }
    }
    const sheetGeneration = this.sheetRequestGeneration;
    const sheet = this.currentSheet;
    const worksheet = this.currentWorksheet;
    await this.scrollToCell(cellRef, options);
    if (this._destroyed) throw this.destroyedError();
    if (
      generation !== this.commentNavigationGeneration ||
      workbook !== this.wb ||
      sheetGeneration !== this.sheetRequestGeneration ||
      sheet !== this.currentSheet ||
      worksheet !== this.currentWorksheet
    ) return false;
    this.setSelection(cellRef);
    return true;
  }

  /** Returns the full selection model, detached from viewer-owned state. */
  get selectionState(): XlsxSelectionState | null {
    return this.selectionController.snapshot();
  }

  /**
   * Set an A1 area (`B2:D5`, `2:4`, `B:D`), a complete canonical state, or
   * `null`. A string describes selection geometry only; its normalized
   * upper-left cell becomes ActiveCell and the Shift-extension anchor.
   */
  setSelection(input: XlsxSelectionInput): void {
    if (this._destroyed) throw new Error('XlsxViewer has been destroyed');
    let next: XlsxSelectionState | null;
    if (typeof input === 'string') {
      next = selectionStateFromReference(input);
      if (!next) throw new SyntaxError(`Invalid XLSX selection reference: ${input}`);
    } else {
      next = input ? normalizeSelectionState(input) : null;
    }
    this.commitSelection(next);
  }

  /**
   * Return a serializable, bounded snapshot of the current selection and the
   * populated cells it covers. Intended for read-only AI/MCP context handoff;
   * it exposes no mutable workbook objects and does not touch the Clipboard API.
   */
  getSelectionContext(options: XlsxSelectionContextOptions = {}): XlsxSelectionContext | null {
    this.assertOpen();
    // This synchronous data API cannot await an unloaded selection. Geometry
    // remains selectable; a fresh context notification follows completion.
    if (this.previewCompletion) return null;
    if (this.elementContext) {
      return limitXlsxElementContext(this.elementContext, options.maxTextCharacters);
    }
    const worksheet = this.currentWorksheet;
    const selection = this.selectionState;
    if (!worksheet || !selection) return null;
    const requestedMax = options.maxCells ?? 1_000;
    if (!Number.isFinite(requestedMax) || requestedMax < 0) {
      throw new RangeError('maxCells must be a finite non-negative number.');
    }
    const maxCells = Math.min(MAX_SELECTION_CONTEXT_CELLS, Math.floor(requestedMax));
    const requestedTextMax = options.maxTextCharacters ?? DEFAULT_SELECTION_CONTEXT_TEXT_CHARACTERS;
    if (!Number.isFinite(requestedTextMax) || requestedTextMax < 0) {
      throw new RangeError('maxTextCharacters must be a finite non-negative number.');
    }
    const maxTextCharacters = Math.min(
      MAX_SELECTION_CONTEXT_TEXT_CHARACTERS,
      Math.floor(requestedTextMax),
    );
    let textCharacters = 0;
    let textTruncated = false;
    const boundedField = (input: string | readonly Readonly<{ text: string }>[]): string => {
      const parts: readonly (string | Readonly<{ text: string }>)[] =
        typeof input === 'string' ? [input] : input;
      const chunks: string[] = [];
      let fieldCharacters = 0;
      for (let index = 0; index < parts.length; index++) {
        const sourcePart = parts[index];
        const part = typeof sourcePart === 'string' ? sourcePart : sourcePart.text;
        const allowed = Math.max(0, Math.min(
          MAX_SELECTION_CONTEXT_FIELD_CHARACTERS - fieldCharacters,
          maxTextCharacters - textCharacters,
        ));
        const chunk = safeUtf16Prefix(part, allowed);
        chunks.push(chunk);
        fieldCharacters += chunk.length;
        textCharacters += chunk.length;
        if (chunk.length < part.length || index + 1 < parts.length && allowed === 0) {
          textTruncated = true;
          break;
        }
      }
      return chunks.join('');
    };
    const sheetSelected = selection.areas.some((area) => area.kind === 'sheet');
    const rowIntervals = mergeSelectionIntervals(selection.areas.flatMap((area) =>
      area.kind === 'rows' ? [{ first: area.firstRow, last: area.lastRow }] : []));
    const columnIntervals = mergeSelectionIntervals(selection.areas.flatMap((area) =>
      area.kind === 'columns'
        ? [{ first: area.firstColumn, last: area.lastColumn }]
        : []));
    const rectangles = selection.areas.flatMap((area) => area.kind === 'cells' ? [area] : []);
    const events = rectangles.flatMap((area, index) => [
      { row: area.top, index, active: true },
      { row: area.bottom + 1, index, active: false },
    ]).sort((a, b) => a.row - b.row || Number(a.active) - Number(b.active));
    const activeRectangles = new Set<number>();
    let eventIndex = 0;
    let activeColumnIntervals: SelectionInterval[] = [];
    const cells: XlsxSelectionContextCell[] = [];
    let cellsTruncated = false;
    const selectedRowIntervals = sheetSelected || columnIntervals.length > 0
      ? [{ first: 1, last: MAX_WORKSHEET_ROW }]
      : mergeSelectionIntervals([
          ...rowIntervals,
          ...rectangles.map((area) => ({ first: area.top, last: area.bottom })),
        ]);
    let rows = this.selectionContextRows.get(worksheet);
    if (!rows) {
      rows = orderedBy(worksheet.rows, (row) => row.index);
      this.selectionContextRows.set(worksheet, rows);
    }

    cellScan: for (const selectedRows of selectedRowIntervals) {
      let rowIndex = lowerBoundBy(rows, selectedRows.first, (row) => row.index);
      while (rowIndex < rows.length) {
        const row = rows[rowIndex++];
        if (row.index > selectedRows.last) break;
        let changed = false;
        while (eventIndex < events.length && events[eventIndex].row <= row.index) {
          const event = events[eventIndex++];
          if (event.active) activeRectangles.add(event.index);
          else activeRectangles.delete(event.index);
          changed = true;
        }
        if (changed) {
          activeColumnIntervals = mergeSelectionIntervals([...activeRectangles].map((index) => ({
            first: rectangles[index].left,
            last: rectangles[index].right,
          })));
        }
        const wholeRow = sheetSelected || intervalContains(rowIntervals, row.index);
        const selectedColumns = wholeRow
          ? [{ first: 1, last: MAX_WORKSHEET_COL }]
          : mergeSelectionIntervals([...columnIntervals, ...activeColumnIntervals]);
        for (const selectedColumnsInterval of selectedColumns) {
          let rowCells = this.selectionContextCells.get(row);
          if (!rowCells) {
            rowCells = orderedBy(row.cells, (cell) => cell.col);
            this.selectionContextCells.set(row, rowCells);
          }
          let cellIndex = lowerBoundBy(rowCells, selectedColumnsInterval.first, (cell) => cell.col);
          while (cellIndex < rowCells.length) {
            const cell = rowCells[cellIndex++];
            if (cell.col > selectedColumnsInterval.last) break;
            const raw = cell.value;
            const sourceComment = this.sourceCommentMap.get(`${cell.row}:${cell.col}`);
            if (raw.type === 'empty' && cell.formula === undefined && !sourceComment) continue;
            if (cells.length >= maxCells) { cellsTruncated = true; break cellScan; }
            const displayText = boundedField(this.wb?.cellText(worksheet, cell) ?? '');
            const value = raw.type === 'text'
              ? boundedField(raw.runs ?? raw.text)
              : raw.type === 'number'
                ? raw.number
                : raw.type === 'bool'
                  ? raw.bool
                  : raw.type === 'error'
                    ? boundedField(raw.error)
                  : null;
            const comment = sourceComment ? {
              root: {
                id: sourceComment.id,
                author: sourceComment.author,
                date: sourceComment.date,
                text: boundedField(sourceComment.rootText ?? sourceComment.text),
                status: sourceComment.resolved ? 'resolved' as const : 'active' as const,
              },
              replies: (sourceComment.replies ?? []).map((reply) => ({
                id: reply.id,
                author: reply.author,
                date: reply.date,
                text: boundedField(reply.text),
                status: reply.resolved ? 'resolved' as const : 'active' as const,
              })),
            } : undefined;
            cells.push({
              address: { row: cell.row, col: cell.col },
              displayText,
              valueType: raw.type,
              value,
              ...(cell.formula === undefined ? {} : { formula: boundedField(cell.formula) }),
              ...(comment === undefined ? {} : { comment }),
            });
            if (textTruncated) break cellScan;
          }
        }
      }
    }
    const truncationReasons: Array<'cells' | 'text'> = [];
    if (cellsTruncated) truncationReasons.push('cells');
    if (textTruncated) truncationReasons.push('text');
    return {
      format: 'xlsx',
      kind: 'range',
      sheetIndex: this.currentSheet,
      sheetName: worksheet.name,
      selection,
      coordinateCountUpperBound: selectionCoordinateCountUpperBound(selection),
      cells,
      truncated: truncationReasons.length > 0,
      truncationReasons,
      maxCells,
      textCharacters,
      maxTextCharacters,
    };
  }

  private commitSelection(next: XlsxSelectionState | null): void {
    this.setElementContext(null);
    const current = this.selectionState;
    if (selectionStatesEqual(current, next)) return;
    this.hideValidationPanel();
    this.selectionController.setState(next);
    this.updateSelectionOverlay();
    if (this.wb) this.scheduleRender();
    this.emitSelectionChange();
  }

  private setElementContext(context: XlsxElementContext | null): boolean {
    if (JSON.stringify(this.elementContext) === JSON.stringify(context)) return false;
    this.elementContext = context ? structuredClone(context) : null;
    this.updateSelectionOverlay();
    this.scheduleSelectionContextNotification();
    return true;
  }

  private scheduleSelectionContextNotification(): void {
    if (!this.opts.onSelectionContextChange || this._destroyed ||
        this.selectionContextNotificationFrame !== null ||
        this.selectionContextNotificationMicrotask) return;
    const notify = () => {
      this.selectionContextNotificationFrame = null;
      this.selectionContextNotificationMicrotask = false;
      if (this._destroyed) return;
      const context = this.getSelectionContext({
        maxTextCharacters: DEFAULT_SELECTION_CONTEXT_NOTIFICATION_TEXT_CHARACTERS,
      });
      this.opts.onSelectionContextChange?.(context ? structuredClone(context) : null);
    };
    if (typeof this.hostWindow.requestAnimationFrame === 'function') {
      this.selectionContextNotificationFrame = this.hostWindow.requestAnimationFrame(notify);
    } else {
      this.selectionContextNotificationMicrotask = true;
      queueMicrotask(notify);
    }
  }

  private emitSelectionChange(): void {
    const state = this.selectionState;
    if (!selectionStatesEqual(state, this.lastNotifiedSelectionState)) {
      this.scheduleSelectionContextNotification();
    }
    if (this.emittingSelectionChange) {
      this.pendingSelectionChange = true;
      this.scheduleSelectionNotification();
      return;
    }
    this.pendingSelectionChange = false;
    if (selectionStatesEqual(state, this.lastNotifiedSelectionState)) {
      this.finishSelectionNotificationChain();
      return;
    }

    if (this.selectionNotificationCount >= MAX_REENTRANT_SELECTION_NOTIFICATIONS) {
      // A callback feedback cycle must not monopolize the main thread. The
      // canonical state remains authoritative; only notifications beyond the
      // documented per-chain safety limit are suppressed.
      this.lastNotifiedSelectionState = state ? structuredClone(state) : null;
      this.finishSelectionNotificationChain();
      return;
    }
    this.selectionNotificationCount++;
    this.lastNotifiedSelectionState = state ? structuredClone(state) : null;
    this.emittingSelectionChange = true;
    try {
      this.opts.onSelectionStateChange?.(state ? structuredClone(state) : null);
    } finally {
      this.emittingSelectionChange = false;
      if (this.pendingSelectionChange ||
          !selectionStatesEqual(this.selectionState, this.lastNotifiedSelectionState)) {
        this.scheduleSelectionNotification();
      } else {
        this.finishSelectionNotificationChain();
      }
    }
  }

  private scheduleSelectionNotification(): void {
    if (this.selectionNotificationScheduled || this._destroyed) return;
    this.selectionNotificationScheduled = true;
    queueMicrotask(() => {
      this.selectionNotificationScheduled = false;
      if (!this._destroyed) this.emitSelectionChange();
    });
  }

  private finishSelectionNotificationChain(): void {
    this.pendingSelectionChange = false;
    this.selectionNotificationCount = 0;
  }

  /**
   * Returns what the header area contains at the given client coordinates.
   * Returns null when the point is in the cell grid (not a header).
   */
  private getHeaderHit(
    clientX: number,
    clientY: number,
  ): { kind: 'corner' } | { kind: 'row'; row: number } | { kind: 'col'; col: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.viewport.scale;
    const rect = this.canvasArea.getBoundingClientRect();
    // Same RTL un-mirror as getCellAt: map the screen x back to the logical-LTR
    // layout (row header on the left) before the header math below.
    const lx = this.screenX(clientX - rect.left, 0);
    const ly = clientY - rect.top;

    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const inRowHeader = lx < headerW;
    const inColHeader = ly < headerH;
    if (!inRowHeader && !inColHeader) return null;
    if (inRowHeader && inColHeader) return { kind: 'corner' };

    const geometry = getGridGeometryForWorksheet(ws);

    if (inRowHeader) {
      // Determine which row was clicked
      const innerY = ly - headerH;
      if (innerY < 0) return { kind: 'corner' };
      const r = geometry.rowAt(innerY, this.viewportTop, cs);
      return r === null ? null : { kind: 'row', row: r };
    }

    // inColHeader
    const innerX = lx - headerW;
    if (innerX < 0) return { kind: 'corner' };
    const c = geometry.colAt(innerX, this.effectiveScrollLeft, cs);
    return c === null ? null : { kind: 'col', col: c };
  }

  /**
   * If the pointer sits on a column/row-header border (within {@link
   * RESIZE_GRAB_PX}), return the resize target: which index to resize and the
   * fixed LTR edge it grows from (in canvasArea CSS px). Excel resizes the band
   * whose *trailing* border you grab — the column to the left of a vertical
   * border, the row above a horizontal one — so both that band and its
   * neighbour-to-the-far-side are checked. Geometry comes straight from {@link
   * getCellRect}, so the grab line always coincides with the drawn border at any
   * scroll offset / zoom / RTL. Returns null off the header borders.
   */
  private getResizeTarget(
    clientX: number,
    clientY: number,
  ): { kind: 'col' | 'row'; index: number; originScaled: number; mdw: number } | null {
    const ws = this.currentWorksheet;
    if (!ws) return null;
    const cs = this.viewport.scale;
    const rect = this.canvasArea.getBoundingClientRect();
    // Un-mirror the screen x to the logical-LTR space getCellRect draws in (the
    // same transform getHeaderHit uses), so the comparison holds for RTL sheets.
    const ptX = this.screenX(clientX - rect.left, 0);
    const ptY = clientY - rect.top;
    const headerW = Math.round(HEADER_W * cs);
    const headerH = Math.round(HEADER_H * cs);
    const mdw = getGridGeometryForWorksheet(ws).maximumDigitWidth;

    // Column borders live in the column-header strip, right of the corner.
    if (ptY <= headerH && ptX > headerW) {
      const hit = this.getHeaderHit(clientX, clientY);
      if (hit?.kind !== 'col') return null;
      const origins = new Map<number, number>(); // index -> fixed LTR origin edge
      const edges: { index: number; edge: number }[] = [];
      for (const c of [hit.col - 1, hit.col]) {
        if (c < 1) continue;
        const r = this._cellRect(1, c); // x is independent of the row
        if (!r) continue;
        origins.set(c, r.x);
        edges.push({ index: c, edge: r.x + r.w }); // trailing (right) border
      }
      const index = resizeHitIndex(ptX, edges, RESIZE_GRAB_PX, headerW);
      if (index === null) return null;
      return { kind: 'col', index, originScaled: origins.get(index) as number, mdw };
    }

    // Row borders live in the row-header strip, below the corner.
    if (ptX <= headerW && ptY > headerH) {
      const hit = this.getHeaderHit(clientX, clientY);
      if (hit?.kind !== 'row') return null;
      const origins = new Map<number, number>(); // index -> fixed LTR origin edge
      const edges: { index: number; edge: number }[] = [];
      for (const rIdx of [hit.row - 1, hit.row]) {
        if (rIdx < 1) continue;
        const r = this._cellRect(rIdx, 1); // y is independent of the column
        if (!r) continue;
        origins.set(rIdx, r.y);
        edges.push({ index: rIdx, edge: r.y + r.h }); // trailing (bottom) border
      }
      const index = resizeHitIndex(ptY, edges, RESIZE_GRAB_PX, headerH);
      if (index === null) return null;
      return { kind: 'row', index, originScaled: origins.get(index) as number, mdw };
    }

    return null;
  }

  /**
   * Apply a live resize drag: size the band from its fixed origin edge to the
   * current pointer, clamp to {@link RESIZE_MIN_PX}, and write the result back
   * into the in-memory worksheet model in its native unit (Excel column widths /
   * points). This is a *view-only* mutation — the file is never written. The
   * memoized axis cache for this sheet is invalidated so every geometry read
   * (spacer, hit-test, overlay, renderer) sees the new size on the next frame.
   */
  private applyResize(clientX: number, clientY: number): void {
    const drag = this.resizeDrag;
    const ws = this.currentWorksheet;
    if (!drag || !ws) return;
    const cs = this.viewport.scale;
    const rect = this.canvasArea.getBoundingClientRect();

    if (drag.kind === 'col') {
      const ptX = this.screenX(clientX - rect.left, 0);
      const sizePx = Math.max(RESIZE_MIN_PX, Math.round((ptX - drag.originScaled) / cs));
      ws.colWidths[drag.index] = pxToColWidth(sizePx, drag.mdw);
      this.recordSizeOverride('col', drag.index);
    } else {
      const ptY = clientY - rect.top;
      const sizePx = Math.max(RESIZE_MIN_PX, Math.round((ptY - drag.originScaled) / cs));
      ws.rowHeights[drag.index] = pxToRowHeight(sizePx);
      this.recordSizeOverride('row', drag.index);
    }

    GridGeometry.invalidate(ws); // sizes changed → rebuild the cumulative-offset axes
    this.updateSpacerSize(ws);
    this.updateSelectionOverlay();
    // Live resize drag fires per pointermove; coalesce the canvas repaint into
    // one frame. The spacer (scrollbar extent) and overlay updates are cheap DOM
    // writes that must track the drag immediately, so they stay synchronous.
    this.scheduleRender();
  }

  /** Refit automatic rows once after a column-resize gesture. Doing this on
   * every pointermove would turn a drag into O(sheet cells × pointer events),
   * while Excel's observable result only needs to be committed at release. */
  private refitAutoRowsAfterColumnResize(): void {
    const ws = this.currentWorksheet;
    const workbook = this.preparedWorkbook;
    if (!ws || !workbook) return;
    const manualRows = this.viewEdits.manualRows(this.currentSheet);
    invalidateAutoRowHeights(ws, manualRows);
    const prepareRowHeights = workbook[prepareXlsxViewerRowHeights];
    if (typeof prepareRowHeights !== 'function') return;
    const measureCanvas = this.hostDocument.createElement('canvas');
    const measureCtx = measureCanvas.getContext('2d');
    if (!measureCtx) return;
    prepareRowHeights.call(workbook, ws, measureCtx);
    this.viewEdits.syncAutomaticRowOverrides(this.currentSheet, ws);
    this.updateSpacerSize(ws);
    this.updateSelectionOverlay();
    this.scheduleRender();
  }

  /**
   * Change the cell-selection highlight color at runtime (see {@link
   * XlsxViewerOptions.selectionColor}). The border takes the color as-is and the
   * fill becomes a translucent shade of it; the current selection repaints
   * immediately.
   */
  setSelectionColor(color: string): void {
    this.opts.selectionColor = color;
    this.updateSelectionOverlay();
  }

  /**
   * Switch the hidden-sheet mode at runtime: restyle the tabs and re-render.
   * Entering `'skip'` while on a hidden sheet advances to the nearest visible.
   */
  async setHiddenSheetMode(mode: HiddenSheetMode): Promise<void> {
    this._hiddenSheetMode = mode;
    this.buildTabs();
    if (mode === 'skip' && this.wb && this.wb.isHidden(this.currentSheet)) {
      await this.showSheet(
        resolveVisibleIndex(this.currentSheet, (i) => this.wb!.isHidden(i), this.sheetCount),
      );
    } else {
      this.updateTabActive(this.currentSheet);
    }
  }

  /** The current hidden-sheet mode. */
  get hiddenSheetMode(): HiddenSheetMode { return this._hiddenSheetMode; }

  /** Number of non-hidden sheets (absolute `sheetCount` is unchanged). */
  get visibleSheetCount(): number {
    if (!this.wb) return 0;
    const wb = this.wb;
    return countVisible((i) => wb.isHidden(i), this.sheetCount);
  }

  /**
   * Copy the selected area as bounded TSV. The same limits apply regardless of
   * whether pointer, keyboard, or API created the selection.
   */
  async copySelection(): Promise<XlsxCopyResult> {
    this.assertOpen();
    if (this.previewCompletion) await this.previewCompletion;
    const ws = this.currentWorksheet;
    const state = this.selectionState;
    if (!ws || !state) return { status: 'empty-selection' };
    if (state.areas.length !== 1) return { status: 'unsupported-multiple-areas' };
    const area = state.areas[0];

    // Whole-row/column/sheet selections are unbounded Excel concepts. Copying
    // narrows them to used cells without changing the logical selection.
    let maxRow = 1, maxCol = 1;
    for (const row of ws.rows) {
      if (row.index > maxRow) maxRow = row.index;
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col;
      }
    }

    const { r1, r2, c1, c2 } = area.kind === 'sheet'
      ? { r1: 1, r2: maxRow, c1: 1, c2: maxCol }
      : area.kind === 'rows'
        ? { r1: area.firstRow, r2: area.lastRow, c1: 1, c2: maxCol }
        : area.kind === 'columns'
          ? { r1: 1, r2: maxRow, c1: area.firstColumn, c2: area.lastColumn }
          : { r1: area.top, r2: area.bottom, c1: area.left, c2: area.right };

    const rowCount = r2 - r1 + 1;
    const colCount = c2 - c1 + 1;
    if (rowCount > Math.floor(MAX_CLIPBOARD_CELLS / colCount)) {
      return { status: 'too-large', limit: 'cells' };
    }
    const cellCount = rowCount * colCount;

    let utf16CodeUnits = Math.max(0, rowCount - 1) + rowCount * Math.max(0, colCount - 1);
    if (utf16CodeUnits > MAX_CLIPBOARD_UTF16_CODE_UNITS) {
      return { status: 'too-large', limit: 'text' };
    }
    const cellMap = new Map<number, Map<number, string>>();
    for (const row of ws.rows) {
      if (row.index < r1 || row.index > r2) continue;
      for (const cell of row.cells) {
        if (cell.col < c1 || cell.col > c2) continue;
        const v = cell.value;
        let text = this.wb?.cellText(ws, cell) ?? '';
        if (!this.wb) {
          if (v.type === 'text') text = v.runs ? v.runs.map((r) => r.text).join('') : v.text;
          else if (v.type === 'number') text = String(v.number);
          else if (v.type === 'bool') text = v.bool ? 'TRUE' : 'FALSE';
          else if (v.type === 'error') text = v.error;
        }
        if (text) {
          const encoded = encodeTsvFieldWithin(
            text,
            MAX_CLIPBOARD_UTF16_CODE_UNITS - utf16CodeUnits,
          );
          if (encoded === null) return { status: 'too-large', limit: 'text' };
          utf16CodeUnits += encoded.length;
          let values = cellMap.get(row.index);
          if (!values) { values = new Map(); cellMap.set(row.index, values); }
          values.set(cell.col, encoded);
        }
      }
    }

    const lines: string[] = [];
    for (let r = r1; r <= r2; r++) {
      const cols: string[] = [];
      const values = cellMap.get(r);
      for (let c = c1; c <= c2; c++) {
        const value = values?.get(c) ?? '';
        cols.push(value);
      }
      lines.push(cols.join('\t'));
    }
    const clipboard = this.hostWindow.navigator.clipboard;
    if (!clipboard) return { status: 'clipboard-unavailable' };
    try {
      await clipboard.writeText(lines.join('\n'));
      return { status: 'copied', cellCount, utf16CodeUnits };
    } catch {
      return { status: 'clipboard-denied' };
    }
  }

  private updateSelectionOverlay(): void {
    this.overlayHost.clearSelection();
    if (this.elementContext) {
      this.drawElementContextOverlay();
      return;
    }
    const state = this.selectionState;
    if (!state) return;
    const cs = this.viewport.scale;
    const ws = this.currentWorksheet;
    if (!ws) return;
    const sp = (px: number) => Math.round(px * cs);
    const headerW = sp(HEADER_W);
    const headerH = sp(HEADER_H);
    const width = this.canvasArea.clientWidth;
    const height = this.canvasArea.clientHeight;
    const geometry = getGridGeometryForWorksheet(ws);
    // Match renderViewport's physical freeze materialization. A legal freeze
    // count may cover the full sheet; it must never create million-row overlay
    // geometry when only a handful of bands can reach this viewport.
    const effective = geometry.effectiveFrozenBands({
      scale: cs, width, height, headerWidth: HEADER_W, headerHeight: HEADER_H,
      rows: ws.freezeRows ?? 0, cols: ws.freezeCols ?? 0,
    });
    const axes = geometry.axesAtScale(cs);
    const frozenW = axes.col.offsetOf(effective.cols + 1);
    const frozenH = axes.row.offsetOf(effective.rows + 1);
    const xPanes = effective.cols > 0
      ? [
          { first: 1, last: effective.cols, start: headerW, end: Math.min(width, headerW + frozenW) },
          { first: effective.cols + 1, last: MAX_WORKSHEET_COL, start: Math.min(width, headerW + frozenW), end: width },
        ]
      : [{ first: 1, last: MAX_WORKSHEET_COL, start: headerW, end: width }];
    const yPanes = effective.rows > 0
      ? [
          { first: 1, last: effective.rows, start: headerH, end: Math.min(height, headerH + frozenH) },
          { first: effective.rows + 1, last: MAX_WORKSHEET_ROW, start: Math.min(height, headerH + frozenH), end: height },
        ]
      : [{ first: 1, last: MAX_WORKSHEET_ROW, start: headerH, end: height }];
    const selectionColor = this.opts.selectionColor ?? DEFAULT_SELECTION_COLOR;
    const { background } = selectionOverlayStyle(selectionColor);
    const seenFragments = new Set<string>();
    const fillSubpaths: string[] = [];
    const overlayRects: SelectionOverlayRect[] = [];

    for (const area of state.areas) {
      const bounds = area.kind === 'cells'
        ? { top: area.top, bottom: area.bottom, left: area.left, right: area.right,
            topEdge: true, bottomEdge: true, leftEdge: true, rightEdge: true }
        : area.kind === 'rows'
          ? { top: area.firstRow, bottom: area.lastRow, left: 1, right: MAX_WORKSHEET_COL,
              topEdge: true, bottomEdge: true, leftEdge: false, rightEdge: false }
          : area.kind === 'columns'
            ? { top: 1, bottom: MAX_WORKSHEET_ROW, left: area.firstColumn, right: area.lastColumn,
                topEdge: false, bottomEdge: false, leftEdge: true, rightEdge: true }
            : { top: 1, bottom: MAX_WORKSHEET_ROW, left: 1, right: MAX_WORKSHEET_COL,
                topEdge: false, bottomEdge: false, leftEdge: false, rightEdge: false };

      for (const yp of yPanes) for (const xp of xPanes) {
        if (xp.end <= xp.start || yp.end <= yp.start) continue;
        const top = Math.max(bounds.top, yp.first);
        const bottom = Math.min(bounds.bottom, yp.last);
        const left = Math.max(bounds.left, xp.first);
        const right = Math.min(bounds.right, xp.last);
        if (top > bottom || left > right) continue;
        const tl = this._cellRect(top, left);
        const br = this._cellRect(bottom, right);
        if (!tl || !br) continue;
        const rawLeft = tl.x;
        const rawTop = tl.y;
        const rawRight = br.x + br.w;
        const rawBottom = br.y + br.h;
        const x = Math.max(rawLeft, xp.start);
        const y = Math.max(rawTop, yp.start);
        const x2 = Math.min(rawRight, xp.end);
        const y2 = Math.min(rawBottom, yp.end);
        const fragmentW = x2 - x;
        const fragmentH = y2 - y;
        if (fragmentW <= 0 || fragmentH <= 0) continue;

        // Only paint a border where the logical selection itself ends. Pane and
        // viewport clips are not selection edges and must not create fake lines.
        const topBorder = bounds.topEdge && top === bounds.top && rawTop >= yp.start;
        const bottomBorder = bounds.bottomEdge && bottom === bounds.bottom && rawBottom <= yp.end;
        const leftBorder = bounds.leftEdge && left === bounds.left && rawLeft >= xp.start;
        const rightBorder = bounds.rightEdge && right === bounds.right && rawRight <= xp.end;
        const screenLeft = this.screenX(x, fragmentW);
        const physicalLeftBorder = this.isRtl ? rightBorder : leftBorder;
        const physicalRightBorder = this.isRtl ? leftBorder : rightBorder;
        const fragmentKey = [
          screenLeft, y, fragmentW, fragmentH,
          topBorder, physicalRightBorder, bottomBorder, physicalLeftBorder,
        ].join('|');
        if (seenFragments.has(fragmentKey)) continue;
        seenFragments.add(fragmentKey);
        // Paint every fragment as a subpath in one SVG fill operation. With a
        // single non-zero fill, overlapping selection areas form a visual union
        // instead of stacking translucent backgrounds and becoming darker.
        fillSubpaths.push(
          `M${screenLeft} ${y}h${fragmentW}v${fragmentH}h${-fragmentW}Z`,
        );
        overlayRects.push({
          x: screenLeft,
          y,
          width: fragmentW,
          height: fragmentH,
          top: topBorder,
          right: physicalRightBorder,
          bottom: bottomBorder,
          left: physicalLeftBorder,
        });
      }
    }

    if (fillSubpaths.length > 0) {
      const svgNamespace = 'http://www.w3.org/2000/svg';
      const svg = this.hostDocument.createElementNS(svgNamespace, 'svg');
      svg.setAttribute('data-xlsx-selection-fill', '');
      svg.style.cssText =
        'position:absolute;inset:0;width:100%;height:100%;overflow:hidden;pointer-events:none;';
      const isMultipleAreaSelection = state.areas.length > 1;
      const activeRect = this._cellRect(state.activeCell.row, state.activeCell.col);
      const maskId = `xlsx-selection-mask-${++selectionMaskSequence}`;
      const defs = this.hostDocument.createElementNS(svgNamespace, 'defs');
      const mask = this.hostDocument.createElementNS(svgNamespace, 'mask');
      mask.setAttribute('id', maskId);
      mask.setAttribute('maskUnits', 'userSpaceOnUse');
      mask.setAttribute('x', '0');
      mask.setAttribute('y', '0');
      mask.setAttribute('width', String(width));
      mask.setAttribute('height', String(height));
      const selectedPath = this.hostDocument.createElementNS(svgNamespace, 'path');
      selectedPath.setAttribute('d', fillSubpaths.join(''));
      selectedPath.setAttribute('fill', '#fff');
      mask.appendChild(selectedPath);

      // Excel leaves ActiveCell unshaded so it remains distinct from the
      // selected cells. ActiveCell stays at the drag origin; only the Area's
      // opposite corner changes during extension.
      if (activeRect) {
        for (const yp of yPanes) for (const xp of xPanes) {
          const clippedX = Math.max(activeRect.x, xp.start);
          const clippedY = Math.max(activeRect.y, yp.start);
          const clippedX2 = Math.min(activeRect.x + activeRect.w, xp.end);
          const clippedY2 = Math.min(activeRect.y + activeRect.h, yp.end);
          if (clippedX2 <= clippedX || clippedY2 <= clippedY) continue;
          const cutout = this.hostDocument.createElementNS(svgNamespace, 'rect');
          cutout.setAttribute('data-xlsx-active-cell-cutout', '');
          cutout.setAttribute('x', String(this.screenX(clippedX, clippedX2 - clippedX)));
          cutout.setAttribute('y', String(clippedY));
          cutout.setAttribute('width', String(clippedX2 - clippedX));
          cutout.setAttribute('height', String(clippedY2 - clippedY));
          cutout.setAttribute('fill', '#000');
          mask.appendChild(cutout);
        }
      }
      defs.appendChild(mask);
      svg.appendChild(defs);

      const fill = this.hostDocument.createElementNS(svgNamespace, 'rect');
      fill.setAttribute('x', '0');
      fill.setAttribute('y', '0');
      fill.setAttribute('width', String(width));
      fill.setAttribute('height', String(height));
      fill.setAttribute('fill', background);
      fill.setAttribute('mask', `url(#${maskId})`);
      svg.appendChild(fill);

      const boundaryPath = isMultipleAreaSelection ? '' : selectionBoundaryPath(overlayRects);
      if (boundaryPath) {
        const boundary = this.hostDocument.createElementNS(svgNamespace, 'path');
        boundary.setAttribute('data-xlsx-selection-border', '');
        boundary.setAttribute('d', boundaryPath);
        boundary.setAttribute('fill', 'none');
        boundary.setAttribute('stroke', selectionColor);
        boundary.setAttribute('stroke-width', '2');
        boundary.setAttribute('stroke-linecap', 'square');
        boundary.setAttribute('stroke-linejoin', 'miter');
        svg.appendChild(boundary);
      }
      if (activeRect && isMultipleAreaSelection) {
        for (const yp of yPanes) for (const xp of xPanes) {
          const clippedX = Math.max(activeRect.x, xp.start);
          const clippedY = Math.max(activeRect.y, yp.start);
          const clippedX2 = Math.min(activeRect.x + activeRect.w, xp.end);
          const clippedY2 = Math.min(activeRect.y + activeRect.h, yp.end);
          if (clippedX2 <= clippedX || clippedY2 <= clippedY) continue;
          const focus = this.hostDocument.createElementNS(svgNamespace, 'rect');
          focus.setAttribute('data-xlsx-active-cell-border', '');
          focus.setAttribute('x', String(this.screenX(clippedX, clippedX2 - clippedX)));
          focus.setAttribute('y', String(clippedY));
          focus.setAttribute('width', String(clippedX2 - clippedX));
          focus.setAttribute('height', String(clippedY2 - clippedY));
          focus.setAttribute('fill', 'none');
          focus.setAttribute('stroke', selectionColor);
          focus.setAttribute('stroke-width', '1');
          svg.appendChild(focus);
        }
      }
      this.overlayHost.appendSelection(svg as unknown as HTMLElement);
    }

    // List data-validation dropdown arrow (ECMA-376 §18.3.1.33). Excel shows an
    // in-cell dropdown button only while the cell is *selected* and only for
    // `list`-type rules — so it is drawn here (selection overlay) rather than in
    // the canvas renderer. The button itself is non-interactive
    // (pointer-events:none); clicks are hit-tested against its rect in the
    // pointerdown handler, which opens a panel listing the allowed values
    // (display only — picking a value never changes the cell).
    this.validation.drawDropdown();
  }

  private drawElementContextOverlay(): void {
    const context = this.elementContext;
    const worksheet = this.currentWorksheet;
    const viewport = this.elementContextViewport();
    if (!context || !worksheet || !viewport || context.sheetIndex !== this.currentSheet) return;
    const projection = projectXlsxElementContext(worksheet, context, viewport);
    if (!projection) return;
    const clip = this.hostDocument.createElement('div');
    clip.setAttribute('data-xlsx-element-context-clip', '');
    clip.style.cssText =
      `position:absolute;left:${projection.clip.x}px;top:${projection.clip.y}px;` +
      `width:${projection.clip.width}px;height:${projection.clip.height}px;` +
      'overflow:hidden;pointer-events:none;';
    const frame = this.hostDocument.createElement('div');
    frame.setAttribute('data-xlsx-element-context-outline', context.elementType);
    const color = this.opts.selectionColor ?? DEFAULT_SELECTION_COLOR;
    frame.style.cssText =
      `position:absolute;left:${projection.rect.x - projection.clip.x}px;` +
      `top:${projection.rect.y - projection.clip.y}px;` +
      `width:${projection.rect.width}px;height:${projection.rect.height}px;` +
      `box-sizing:border-box;border:2px solid ${color};` +
      `background:color-mix(in srgb, ${color} 6%, transparent);` +
      `transform:rotate(${projection.rotation}deg);transform-origin:center;pointer-events:none;`;
    clip.appendChild(frame);
    this.overlayHost.appendSelection(clip);
  }

  // ─── IX2 find-highlight overlay ──────────────────────────────────────────

  /**
   * Redraw the find-highlight overlay: one translucent box per matched cell on
   * the current sheet, the active match in a stronger colour. Uses the SAME
   * `getCellRect` + `screenX` + header/frozen clamp the selection overlay uses,
   * so a box lands exactly on the drawn cell at any scroll offset / zoom / RTL.
   * Rebuilt on every render and scroll (cheap DOM geometry, no canvas paint).
   */
  private updateFindOverlay(): void {
    this.overlayHost.clearFind();
    const ws = this.currentWorksheet;
    if (!ws) return;
    const cs = this.viewport.scale;
    const sp = (px: number) => Math.round(px * cs);
    const headerW = sp(HEADER_W);
    const headerH = sp(HEADER_H);
    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;
    const frozen = getGridGeometryForWorksheet(ws).roundedFrozenExtent(cs);
    const frozenBoundX = headerW + frozen.width;
    const frozenBoundY = headerH + frozen.height;

    // A match accent: same single-color → border + translucent fill derivation
    // the selection overlay uses. The active match uses a warm accent so it is
    // distinguishable from other hits and from the (blue) selection box.
    const other = findHighlightOverlayStyle(false, this.opts.findHighlightColors);
    const active = findHighlightOverlayStyle(true, this.opts.findHighlightColors);

    for (const hl of this._find.sheetHighlights(this.currentSheet)) {
      const rect = this._cellRect(hl.row, hl.col);
      if (!rect) continue;
      let { x, y, w, h } = rect;
      // Clamp against headers + the frozen-pane boundary (scrollable cells that
      // scrolled behind the frozen area are clipped there), mirroring the
      // selection overlay so a highlight never spills over fixed regions.
      if (x < headerW) { w -= headerW - x; x = headerW; }
      if (y < headerH) { h -= headerH - y; y = headerH; }
      if (hl.col > freezeCols && x < frozenBoundX) { w -= frozenBoundX - x; x = frozenBoundX; }
      if (hl.row > freezeRows && y < frozenBoundY) { h -= frozenBoundY - y; y = frozenBoundY; }
      if (w <= 0 || h <= 0) continue;
      const screenLeft = this.screenX(x, w);
      const { border, background } = hl.active ? active : other;
      const box = this.hostDocument.createElement('div');
      box.style.cssText =
        `position:absolute;` +
        `left:${screenLeft}px;top:${y}px;width:${w}px;height:${h}px;` +
        `box-sizing:border-box;border:${border};background:${background};pointer-events:none;`;
      this.overlayHost.appendFind(box);
    }
  }

  /**
   * IX2 — find every occurrence of `query` across every sheet and highlight the
   * matched cells. Returns every match in document order (sheet ascending, then
   * row-major within a sheet), each tagged with its
   * `{ sheet, sheetName, ref, row, col }`. A cell is the search unit: search
   * runs over each cell's *rendered* display text (number formats, dates, rich
   * text flattened), so a query matches what the grid shows. Case-insensitive by
   * default; pass `{ caseSensitive: true }` for an exact match. An empty query
   * clears the find.
   */
  async findText(
    query: string,
    opts: FindMatchesOptions = {},
  ): Promise<FindMatch<XlsxMatchLocation>[]> {
    if (!this.wb) return [];
    const matches = await this._find.find(query, opts);
    this.updateFindOverlay();
    return matches;
  }

  /**
   * IX2 — move to the next match (wrap-around), switching sheets and scrolling
   * the matched cell into view as needed, and highlight it as the active match.
   * Returns the now-active match, or `null` when there are none. Call
   * {@link findText} first.
   */
  async findNext(): Promise<FindMatch<XlsxMatchLocation> | null> {
    return this._activateMatch(this._find.next());
  }

  /** IX2 — move to the previous match (wrap-around). */
  async findPrev(): Promise<FindMatch<XlsxMatchLocation> | null> {
    return this._activateMatch(this._find.prev());
  }

  /** IX2 — clear all highlights and reset the find state. */
  clearFind(): void {
    this._find.invalidate();
    this.updateFindOverlay();
  }

  private async _activateMatch(
    match: FindMatch<XlsxMatchLocation> | null,
  ): Promise<FindMatch<XlsxMatchLocation> | null> {
    if (!match) {
      this.updateFindOverlay();
      return null;
    }
    const { sheet, row, col } = match.location;
    if (sheet !== this.currentSheet) {
      // showSheet resets scroll/selection and re-renders; the find state (and so
      // the highlights) survive because they live on the controller, not the
      // sheet. updateFindOverlay runs after the sheet switch below.
      await this.goToSheet(sheet);
    }
    this._scrollCellIntoView(row, col);
    // Scrolling schedules a coalesced render; draw the highlights now so the
    // active box is visible immediately without waiting a frame.
    this.updateFindOverlay();
    return match;
  }

  /**
   * Scroll the grid so cell (row, col) is comfortably in view. Computes the
   * cell's absolute logical offset from the axis metrics (the same the renderer
   * uses) and nudges the vertical / start-anchored horizontal viewport
   * only when the cell is outside the scrollable viewport — an in-view cell is
   * left where it is (Excel's find behaviour). Frozen cells are always visible,
   * so they need no scroll.
   */
  private _scrollCellIntoView(
    row: number,
    col: number,
    align: NonNullable<XlsxScrollToCellOptions['align']> = 'nearest',
  ): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const cs = this.viewport.scale;
    const offset = getGridGeometryForWorksheet(ws).scrollOffsetForCell(
      row,
      col,
      {
        scale: cs,
        viewportWidth: this.canvasArea.clientWidth,
        viewportHeight: this.canvasArea.clientHeight,
        currentX: this.effectiveScrollLeft,
        currentY: this.viewportTop,
        headerWidth: HEADER_W,
        headerHeight: HEADER_H,
        align,
      },
    );
    this.viewportTop = offset.y;
    this.setViewportLeft(offset.x);
  }

  /** Close the list-validation panel and cancel a pending resolution. */
  private hideValidationPanel(): void {
    this.validation.hide();
  }

  // ─── Comment hover popup ──────────────────────────────────────────────────

  /** Index the displayed sheet's comments for the hover popup. */
  private buildCommentMap(ws: Worksheet): void {
    this.comments.setComments(ws.comments ?? []);
  }

  private createVisibleSheetView(source: Worksheet): Worksheet {
    const worksheet = createSheetViewModel(source);
    if (this.opts.comments === false) {
      const hidden = { ...worksheet, commentRefs: [], comments: [] };
      inheritWorksheetPreviewBounds(worksheet, hidden);
      return hidden;
    }
    // Keep the pre-customization behavior: XLSX historically exposed resolved
    // threaded comments. Consumers may explicitly hide them.
    const commentOptions = typeof this.opts.comments === 'object'
      ? this.opts.comments
      : undefined;
    if (commentOptions?.includeResolved !== false) return worksheet;
    const resolved = new Set(
      (worksheet.comments ?? [])
        .filter((comment) => comment.resolved === true)
        .map((comment) => comment.cellRef),
    );
    if (resolved.size === 0) return worksheet;
    const unresolved = {
      ...worksheet,
      commentRefs: worksheet.commentRefs?.filter((ref) => !resolved.has(ref)),
      comments: worksheet.comments?.filter((comment) => !resolved.has(comment.cellRef)),
    };
    inheritWorksheetPreviewBounds(worksheet, unresolved);
    return unresolved;
  }

  /** IX1 — index the current sheet's hyperlinks by `"row:col"` (1-based, first
   *  cell of the `ref` range) so a clicked/hovered cell resolves in O(1). Keys
   *  match the renderer's `hyperlinkMap` exactly (`${hl.row}:${hl.col}`). */
  private buildHyperlinkMap(ws: Worksheet): void {
    this.hyperlinkMap = new Map();
    for (const hl of ws.hyperlinks ?? []) {
      this.hyperlinkMap.set(`${hl.row}:${hl.col}`, hl);
    }
  }

  /** IX1 — the hyperlink at a cell, or null. `getCellAt` returns 1-based
   *  {row,col}, matching the parser/renderer keying.
   *
   *  Returns null unconditionally when `enableHyperlinks` is `false`: this is the
   *  single gate that disables hyperlink interactivity. Both consumers — the
   *  pointermove pointer-cursor affordance and the click dispatch
   *  ({@link dispatchHyperlink}) — funnel through this hit-test, so a null result
   *  means no cursor change, no default navigation, and no `onHyperlinkClick`. */
  private hyperlinkAtCell(cell: CellAddress): Hyperlink | null {
    if (this.opts.enableHyperlinks === false) return null;
    return this.hyperlinkMap.get(`${cell.row}:${cell.col}`) ?? null;
  }

  /**
   * IX1 — dispatch a click on a hyperlinked cell. Builds a
   * {@link HyperlinkTarget} from the parsed hyperlink (external `url` wins over
   * internal `location`, matching Excel: a `<hyperlink>` carrying both navigates
   * to the external target) and routes it to the caller's `onHyperlinkClick`
   * (which fully owns behaviour) or the built-in default. Returns true when a
   * hyperlink was found and dispatched.
   */
  private dispatchHyperlink(cell: CellAddress): boolean {
    const hl = this.hyperlinkAtCell(cell);
    if (!hl) return false;
    let target: HyperlinkTarget;
    if (hl.url) {
      target = { kind: 'external', url: hl.url };
    } else if (hl.location) {
      target = { kind: 'internal', ref: hl.location };
    } else {
      return false; // parser only emits a hyperlink with url or location
    }
    const custom = this.opts.onHyperlinkClick;
    if (custom) {
      custom(target);
      return true;
    }
    // Built-in default. External: open in a new tab, sanitised against the safe
    // scheme allowlist (a blocked scheme like `javascript:` is a no-op, not a
    // navigation). Internal: best-effort sheet navigation, below.
    if (target.kind === 'external') {
      openExternalHyperlink(target.url, undefined, this.hostWindow);
    } else {
      void this.navigateInternalHyperlink(target.ref).catch(
        (error) => this._reportRenderError(error),
      );
    }
    return true;
  }

  /**
   * IX1 default handler for an internal `location` target (§18.3.1.47): resolve
   * a direct cell/range or an in-scope defined name (§18.2.5), switch sheets when
   * needed, then scroll the first referenced cell into view.
   */
  private async navigateInternalHyperlink(location: string): Promise<void> {
    const target = resolveXlsxInternalHyperlink(
      location,
      this.currentSheet,
      this.sheetNames,
      this.currentWorksheet?.definedNames ?? [],
    );
    if (!target) return;
    if (target.sheetIndex !== this.currentSheet) {
      await this.goToSheet(target.sheetIndex);
    }
    await this.scrollToCell(target.cellRef);
  }

  /** Hide the comment popup and cancel any pending show. */
  private hideCommentPopup(): void {
    this.comments.hide();
  }

  private applyPointerSelection(
    clientX: number,
    clientY: number,
    shiftKey: boolean,
    additiveKey: boolean,
    pointerId: number,
    allowDrag: boolean,
  ): void {
    const headerHit = this.getHeaderHit(clientX, clientY);

    if (headerHit) {
      if (headerHit.kind === 'corner') {
        // Select all — no drag extension needed
        this.selectionController.select({ row: 1, col: 1 }, 'all');
        this.selectionController.endDrag();
      } else if (headerHit.kind === 'row') {
        if (shiftKey && this.anchorCell && this.selectionMode === 'rows') {
          this.selectionController.extend({ row: headerHit.row, col: 1 });
        } else {
          const selected = additiveKey
            ? this.selectionController.add({ row: headerHit.row, col: 1 }, 'rows')
            : (this.selectionController.select({ row: headerHit.row, col: 1 }, 'rows'), true);
          if (allowDrag && selected) {
            this.beginSelectionDrag(pointerId);
            this.scrollHost.setPointerCapture(pointerId);
          }
        }
      } else {
        if (shiftKey && this.anchorCell && this.selectionMode === 'cols') {
          this.selectionController.extend({ row: 1, col: headerHit.col });
        } else {
          const selected = additiveKey
            ? this.selectionController.add({ row: 1, col: headerHit.col }, 'cols')
            : (this.selectionController.select({ row: 1, col: headerHit.col }, 'cols'), true);
          if (allowDrag && selected) {
            this.beginSelectionDrag(pointerId);
            this.scrollHost.setPointerCapture(pointerId);
          }
        }
      }
      this.updateSelectionOverlay();
      void this.renderCurrentSheet().catch((error) => this._reportRenderError(error));
      this.emitSelectionChange();
      return;
    }

    const cell = this.getCellAt(clientX, clientY);
    if (!cell) return;

    let selected = true;
    if (shiftKey && this.anchorCell && this.selectionMode === 'cells') {
      this.selectionController.extend(cell);
    } else {
      selected = additiveKey
        ? this.selectionController.add(cell, 'cells')
        : (this.selectionController.select(cell, 'cells'), true);
    }
    if (allowDrag && selected) {
      this.beginSelectionDrag(pointerId);
      this.scrollHost.setPointerCapture(pointerId);
    }
    this.updateSelectionOverlay();
    if (this.wb) {
      this.renderCurrentSheet().catch((error) => this._reportRenderError(error));
    }
    this.emitSelectionChange();
  }

  /** Browser-visible input box, excluding classic native scrollbar gutters. */
  private viewportInputBounds(): { left: number; top: number; width: number; height: number } {
    const rect = this.canvasArea.getBoundingClientRect();
    const left = rect.left + this.scrollHost.clientLeft;
    const top = rect.top + this.scrollHost.clientTop;
    const availableWidth = Math.max(0, rect.width - this.scrollHost.clientLeft);
    const availableHeight = Math.max(0, rect.height - this.scrollHost.clientTop);
    return {
      left,
      top,
      width: Math.min(availableWidth, this.scrollHost.clientWidth || availableWidth),
      height: Math.min(availableHeight, this.scrollHost.clientHeight || availableHeight),
    };
  }

  /** Extend the active drag selection to the pointer's cell. Captured pointers
   * outside the canvas and auto-scroll ticks clamp to the visible data edge so
   * selection never jumps ahead of the viewport. */
  private extendDragSelection(
    clientX: number,
    clientY: number,
    clampToViewport: boolean,
  ): boolean {
    let pointerX = clientX;
    let pointerY = clientY;
    const bounds = this.viewportInputBounds();
    const outsideViewport = clientX < bounds.left || clientX >= bounds.left + bounds.width ||
      clientY < bounds.top || clientY >= bounds.top + bounds.height;
    if (clampToViewport || outsideViewport) {
      const cs = this.viewport.scale;
      const headerW = Math.round(HEADER_W * cs);
      const headerH = Math.round(HEADER_H * cs);
      const dataLeft = bounds.left + (this.isRtl ? 0 : headerW);
      const dataRight = bounds.left + bounds.width - (this.isRtl ? headerW : 0);
      pointerX = Math.min(dataRight - 1, Math.max(dataLeft + 1, pointerX));
      pointerY = Math.min(
        bounds.top + bounds.height - 1,
        Math.max(bounds.top + headerH + 1, pointerY),
      );
    }

    if (this.selectionMode === 'rows') {
      const hit = clampToViewport ? null : this.getHeaderHit(pointerX, pointerY);
      const row = hit?.kind === 'row'
        ? hit.row
        : this.getCellAt(pointerX, pointerY)?.row;
      if (!row || row === this.activeCell?.row) return false;
      this.selectionController.extend({ row, col: 1 });
      return true;
    }

    if (this.selectionMode === 'cols') {
      const hit = clampToViewport ? null : this.getHeaderHit(pointerX, pointerY);
      const col = hit?.kind === 'col'
        ? hit.col
        : this.getCellAt(pointerX, pointerY)?.col;
      if (!col || col === this.activeCell?.col) return false;
      this.selectionController.extend({ row: 1, col });
      return true;
    }

    const cell = this.getCellAt(pointerX, pointerY);
    if (!cell || (cell.row === this.activeCell?.row && cell.col === this.activeCell?.col)) {
      return false;
    }
    this.selectionController.extend(cell);
    return true;
  }

  private selectionAutoScrollSpeed(): { x: number; y: number } {
    const pointer = this.selectionAutoScrollPointer;
    if (!pointer) return { x: 0, y: 0 };
    const bounds = this.viewportInputBounds();
    return selectionAutoScrollVelocity(
      { x: pointer.clientX - bounds.left, y: pointer.clientY - bounds.top },
      { width: bounds.width, height: bounds.height },
      this.isRtl,
      this.selectionMode,
    );
  }

  private trackSelectionAutoScroll(e: PointerEvent): void {
    if (e.pointerId !== this.selectionPointerId) return;
    this.selectionAutoScrollPointer = {
      clientX: e.clientX,
      clientY: e.clientY,
      pointerId: e.pointerId,
    };
    const speed = this.selectionAutoScrollSpeed();
    if (speed.x === 0 && speed.y === 0) {
      this.stopSelectionAutoScroll();
      return;
    }
    if (this.selectionAutoScrollFrame !== null) return;
    this.selectionAutoScrollLastTime = null;
    this.selectionAutoScrollFrame = this.hostWindow.requestAnimationFrame(
      (time) => this.runSelectionAutoScroll(time),
    );
  }

  private runSelectionAutoScroll(time: number): void {
    this.selectionAutoScrollFrame = null;
    const pointer = this.selectionAutoScrollPointer;
    if (
      !pointer ||
      pointer.pointerId !== this.selectionPointerId ||
      !this.isSelecting ||
      this._destroyed
    ) {
      this.stopSelectionAutoScroll();
      return;
    }

    const speed = this.selectionAutoScrollSpeed();
    if (speed.x === 0 && speed.y === 0) {
      this.stopSelectionAutoScroll();
      return;
    }

    const previousTime = this.selectionAutoScrollLastTime;
    const elapsedSeconds = previousTime === null
      ? 1 / 60
      : Math.min(0.05, Math.max(0, time - previousTime) / 1000);
    this.selectionAutoScrollLastTime = time;

    const beforeX = this.effectiveScrollLeft;
    const beforeY = this.viewportTop;
    this.setViewportLeft(beforeX + speed.x * elapsedSeconds);
    this.viewportTop = beforeY + speed.y * elapsedSeconds;
    const moved = this.effectiveScrollLeft !== beforeX || this.viewportTop !== beforeY;
    const extended = moved && this.extendDragSelection(pointer.clientX, pointer.clientY, true);

    if (moved) {
      this.updateSelectionOverlay();
      this.updateFindOverlay();
      this.scheduleRender();
      this.emitViewportChange();
      if (extended) this.emitSelectionChange();
    }

    if (!moved) {
      this.stopSelectionAutoScroll();
      return;
    }
    this.selectionAutoScrollFrame = this.hostWindow.requestAnimationFrame(
      (nextTime) => this.runSelectionAutoScroll(nextTime),
    );
  }

  private stopSelectionAutoScroll(): void {
    if (this.selectionAutoScrollFrame !== null) {
      this.hostWindow.cancelAnimationFrame(this.selectionAutoScrollFrame);
      this.selectionAutoScrollFrame = null;
    }
    this.selectionAutoScrollPointer = null;
    this.selectionAutoScrollLastTime = null;
  }

  private contextMenuTargetIsSelected(clientX: number, clientY: number): boolean {
    const selection = this.selectionState;
    if (!selection) return false;
    const header = this.getHeaderHit(clientX, clientY);
    if (header?.kind === 'corner') {
      return selection.areas.some((area) => area.kind === 'sheet');
    }
    if (header?.kind === 'row') {
      return selection.areas.some((area) => area.kind === 'sheet' ||
        (area.kind === 'rows' && header.row >= area.firstRow && header.row <= area.lastRow));
    }
    if (header?.kind === 'col') {
      return selection.areas.some((area) => area.kind === 'sheet' ||
        (area.kind === 'columns' &&
          header.col >= area.firstColumn && header.col <= area.lastColumn));
    }
    const cell = this.getCellAt(clientX, clientY);
    return cell !== null && selection.areas.some((area) => areaContainsCell(area, cell));
  }

  private resolveContextMenuContext(event: MouseEvent): Promise<XlsxSelectionContext | null> {
    if (this._destroyed) return Promise.resolve(null);
    const element = this.elementContextAt(event.clientX, event.clientY);
    if (element) {
      this.setElementContext(element);
    } else {
      this.setElementContext(null);
      if (!this.contextMenuTargetIsSelected(event.clientX, event.clientY)) {
        this.applyPointerSelection(event.clientX, event.clientY, false, false, -1, false);
      }
    }
    const context = this.getSelectionContext();
    return Promise.resolve(context ? structuredClone(context) : null);
  }

  private setupSelectionEvents(): void {
    // Distance (CSS px) beyond which a touch/pen pointerdown→pointerup is treated as a swipe (scroll), not a tap.
    const TAP_SLOP = 8;

    if (this.opts.onContextMenu) {
      this.surface.on('contextmenu', (event: MouseEvent) => {
        let context: Promise<XlsxSelectionContext | null> | undefined;
        this.opts.onContextMenu?.({
          originalEvent: event,
          getContext: () => context ??= this.resolveContextMenuContext(event),
        });
      });
    }

    this.surface.on('pointerdown', (e: PointerEvent) => {
      this.scrollHost.focus?.({ preventScroll: true });
      if (e.button !== 0) return;
      if (this.isSelecting && e.pointerId !== this.selectionPointerId) return;

      // Drag-to-resize a column/row from its header border (issue #567). Checked
      // before selection so grabbing the border never moves the cell selection.
      // Gated by the `resizable` option (default true); when off, a header-border
      // press falls through to normal selection behavior.
      const resize = (this.opts.resizable ?? true)
        ? this.getResizeTarget(e.clientX, e.clientY)
        : null;
      if (resize) {
        e.preventDefault();
        this.resizeDrag = { ...resize, pointerId: e.pointerId };
        this.scrollHost.setPointerCapture(e.pointerId);
        this.hideCommentPopup();
        return;
      }

      // List-validation dropdown arrow: if the press lands on the (display-only)
      // arrow button drawn on the active cell, toggle the value panel instead of
      // re-selecting the cell. The arrow's rect is in canvasArea space, so map
      // the client point through canvasArea's box.
      if (this.validation.hitsArrow(e.clientX, e.clientY)) {
        e.preventDefault();
        this.validation.toggle();
        return;
      }

      // A pointerdown on the native scrollbar must not move the cell
      // selection — dragging the thumb would otherwise select whatever cell
      // sits underneath it. Two scrollbar styles need different handling:
      // classic scrollbars reserve layout space, so the press lands in the
      // band between the content box (clientWidth/Height) and the border-box
      // edge and can be rejected exactly; OS overlay scrollbars (macOS
      // "show when scrolling") float over the content without affecting
      // client sizes, so a press near a scrollable edge is geometrically
      // indistinguishable from a cell click. For that case we defer the
      // selection to pointerup via the pendingTap path and cancel it when a
      // scroll event arrives first (the press was a thumb drag). A plain
      // click in the band still selects the cell on release.
      const hostRect = this.scrollHost.getBoundingClientRect();
      const localX = e.clientX - hostRect.left - this.scrollHost.clientLeft;
      const localY = e.clientY - hostRect.top - this.scrollHost.clientTop;
      if (localX >= this.scrollHost.clientWidth || localY >= this.scrollHost.clientHeight) {
        return; // classic scrollbar gutter
      }
      // Overlay scrollbar hit band (~15 CSS px on macOS / Windows 11).
      const OVERLAY_SCROLLBAR_BAND = 16;
      const inOverlayBand = this._nativeScrollbars && (
        (this.scrollHost.scrollWidth > this.scrollHost.clientWidth &&
          this.scrollHost.clientHeight - localY <= OVERLAY_SCROLLBAR_BAND) ||
        (this.scrollHost.scrollHeight > this.scrollHost.clientHeight &&
          this.scrollHost.clientWidth - localX <= OVERLAY_SCROLLBAR_BAND));

      const elementContext = this.elementContextAt(e.clientX, e.clientY);
      if (elementContext) {
        this.pendingTap = null;
        this.pendingClick = null;
        this.pendingElementClick = {
          x: e.clientX,
          y: e.clientY,
          pointerId: e.pointerId,
          context: elementContext,
        };
        return;
      }
      // A cell/header/empty-space press leaves object focus and returns the
      // authoritative context to the existing cell-selection state.
      this.setElementContext(null);

      // Touch / pen: defer selection until pointerup so swipe-to-scroll doesn't change the cell.
      // Mouse: select immediately to preserve drag-to-extend behavior.
      if (e.pointerType !== 'mouse' || inOverlayBand) {
        this.pendingTap = {
          x: e.clientX,
          y: e.clientY,
          shiftKey: e.shiftKey,
          additiveKey: e.ctrlKey || e.metaKey,
          pointerId: e.pointerId,
        };
        return;
      }

      // IX1 — remember the cell under a mouse press so a click (no drag) can
      // activate its hyperlink on release. Recorded before selection so a
      // shift-click extend still tracks the destination cell.
      const downCell = this.getCellAt(e.clientX, e.clientY);
      this.pendingClick = downCell
        ? { x: e.clientX, y: e.clientY, pointerId: e.pointerId, cell: downCell }
        : null;

      this.applyPointerSelection(
        e.clientX,
        e.clientY,
        e.shiftKey,
        e.ctrlKey || e.metaKey,
        e.pointerId,
        true,
      );
    });

    this.surface.on('pointermove', (e: PointerEvent) => {
      // Live column/row resize takes priority over every other pointer behavior.
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        e.preventDefault();
        this.applyResize(e.clientX, e.clientY);
        return;
      }

      // Resize-handle affordance: show the col/row-resize cursor when hovering a
      // header border (mouse only — touch/pen have no hover). Skipped mid-select
      // and when the `resizable` option (default true) is off, so no resize
      // cursor is shown when drag-resize is disabled.
      if (e.pointerType === 'mouse' && !this.isSelecting && (this.opts.resizable ?? true)) {
        const rt = this.getResizeTarget(e.clientX, e.clientY);
        this.scrollHost.style.cursor = rt ? (rt.kind === 'col' ? 'col-resize' : 'row-resize') : '';
        if (rt) {
          this.hideCommentPopup();
          return;
        }
      }

      // Cancel a pending tap once the pointer moves beyond the slop — the user is scrolling.
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) {
          this.pendingTap = null;
        }
      }

      // IX1 — a mouse press that turns into a drag (beyond the slop) is a
      // selection, not a hyperlink click: drop the pending activation.
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingClick.x;
        const dy = e.clientY - this.pendingClick.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) {
          this.pendingClick = null;
        }
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingElementClick.x;
        const dy = e.clientY - this.pendingElementClick.y;
        if (dx * dx + dy * dy > TAP_SLOP * TAP_SLOP) this.pendingElementClick = null;
      }

      // Comment hover popup (mouse only — touch/pen have no hover, so they get
      // the popup on selection instead, below). Suppressed while drag-selecting
      // so the popup doesn't fight the selection rect. A header hover hides it.
      if (e.pointerType === 'mouse' && !this.isSelecting) {
        const hovered = this.getCellAt(e.clientX, e.clientY);
        if (hovered) this.comments.scheduleForCell(hovered);
        else this.hideCommentPopup();
        // IX1 — pointer cursor over a hyperlinked cell. Reached only when the
        // pointer is NOT over a resize border (that path returns above), so the
        // resize cursor is never clobbered. Otherwise clear back to default.
        this.scrollHost.style.cursor =
          hovered && this.hyperlinkAtCell(hovered) ? 'pointer' : '';
      }

      if (!this.isSelecting || e.pointerId !== this.selectionPointerId) return;

      this.trackSelectionAutoScroll(e);
      if (!this.extendDragSelection(e.clientX, e.clientY, false)) return;

      this.updateSelectionOverlay();
      // Drag-select fires per pointermove; coalesce the canvas repaint (the
      // header-highlight bands the renderer draws) into one frame. The overlay
      // rect and the selection-change callback stay synchronous.
      this.scheduleRender();
      this.emitSelectionChange();
    });

    this.surface.on('pointerup', (e: PointerEvent) => {
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        if (this.resizeDrag.kind === 'col') this.refitAutoRowsAfterColumnResize();
        this.scrollHost.releasePointerCapture(e.pointerId);
        this.resizeDrag = null;
        return;
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        const pending = this.pendingElementClick;
        this.pendingElementClick = null;
        const dx = e.clientX - pending.x;
        const dy = e.clientY - pending.y;
        const current = dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP
          ? this.elementContextAt(e.clientX, e.clientY)
          : null;
        if (
          current &&
          current.sheetIndex === pending.context.sheetIndex &&
          current.elementType === pending.context.elementType &&
          current.elementIndex === pending.context.elementIndex &&
          current.shapeIndex === pending.context.shapeIndex
        ) this.setElementContext(current);
        return;
      }
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingTap.x;
        const dy = e.clientY - this.pendingTap.y;
        if (dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP) {
          this.applyPointerSelection(
            e.clientX,
            e.clientY,
            this.pendingTap.shiftKey,
            this.pendingTap.additiveKey,
            e.pointerId,
            false,
          );
          // Touch / pen have no hover, so surface the comment popup on a tap
          // (the active cell after the selection commit). Mouse uses hover.
          if (e.pointerType !== 'mouse' && this.activeCell) {
            const comment = this.comments.commentAt(this.activeCell);
            if (comment) {
              this.hideCommentPopup();
              void this.comments.show(this.activeCell, comment)
                .catch((error) => this._reportRenderError(error));
            } else {
              this.hideCommentPopup();
            }
          }
          // IX1 — a touch/pen tap on a hyperlinked cell activates it.
          if (this.activeCell) this.dispatchHyperlink(this.activeCell);
        }
        this.pendingTap = null;
      }
      const endsSelectionDrag = e.pointerId === this.selectionPointerId;
      if (endsSelectionDrag) this.stopSelectionAutoScroll();
      // IX1 — a mouse click (press+release without a drag) on a hyperlinked cell
      // activates it. The release must still land on the same cell the press did.
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        const dx = e.clientX - this.pendingClick.x;
        const dy = e.clientY - this.pendingClick.y;
        const upCell = this.getCellAt(e.clientX, e.clientY);
        if (
          dx * dx + dy * dy <= TAP_SLOP * TAP_SLOP &&
          upCell &&
          upCell.row === this.pendingClick.cell.row &&
          upCell.col === this.pendingClick.cell.col
        ) {
          this.dispatchHyperlink(this.pendingClick.cell);
        }
        this.pendingClick = null;
      }
      if (endsSelectionDrag) this.selectionController.endDrag(e.pointerId);
    });

    this.surface.on('pointercancel', (e: PointerEvent) => {
      if (this.resizeDrag && this.resizeDrag.pointerId === e.pointerId) {
        if (this.resizeDrag.kind === 'col') this.refitAutoRowsAfterColumnResize();
        this.resizeDrag = null;
      }
      if (this.pendingTap && this.pendingTap.pointerId === e.pointerId) {
        this.pendingTap = null;
      }
      if (this.pendingClick && this.pendingClick.pointerId === e.pointerId) {
        this.pendingClick = null;
      }
      if (this.pendingElementClick?.pointerId === e.pointerId) {
        this.pendingElementClick = null;
      }
      if (e.pointerId === this.selectionPointerId) {
        this.stopSelectionAutoScroll();
        this.selectionController.endDrag(e.pointerId);
      }
    });

    // Ctrl/⌘ + mouse wheel (and trackpad pinch, which the browser reports as a
    // ctrl-wheel) zooms the grid, matching Excel. preventDefault stops the
    // browser's own page zoom. A plain wheel still scrolls the grid natively.
    // The step is exponential in mode-normalized wheel distance (see
    // zoomStepScale), so a trackpad pinch — a high-frequency stream of
    // small-deltaY events — does not zoom away; the total zoom tracks the gesture
    // distance, not the event count, while a mouse wheel remains a gentle 10%.
    this.surface.on(
      'wheel',
      (e: WheelEvent) => {
        if (!(e.ctrlKey || e.metaKey)) {
          if (!this._nativeScrollbars) {
            e.preventDefault();
            const unit = e.deltaMode === WheelEvent.DOM_DELTA_LINE
              ? 16
              : e.deltaMode === WheelEvent.DOM_DELTA_PAGE
                ? Math.max(1, this.scrollHost.clientHeight)
                : 1;
            const horizontal = (e.shiftKey ? e.deltaY : e.deltaX) * unit;
            const vertical = (e.shiftKey ? 0 : e.deltaY) * unit;
            this.setViewportLeft(this.effectiveScrollLeft + horizontal);
            this.viewportTop += vertical;
            this.scheduleRender();
            this.updateSelectionOverlay();
            this.updateFindOverlay();
            this.emitViewportChange();
          }
          return;
        }
        e.preventDefault();
        if (e.deltaY === 0) return;
        // Pointer-anchored zoom: pivot on the cursor, not the top-left corner.
        // Record the pointer relative to the grid's top-left (canvasArea rect,
        // which the scrollHost overlays with inset:0) so `setScale` keeps the
        // cell under the cursor fixed. `scrollHost` and `canvasArea` share a rect.
        // A malformed event (no clientX/Y) yields a non-finite anchor; drop it so
        // `setScale` falls back to the historical START-anchored preservation.
        const { x: ax, y: ay } = this.surface.localPoint(e.clientX, e.clientY);
        this._pendingZoomAnchor =
          Number.isFinite(ax) && Number.isFinite(ay) ? { x: ax, y: ay } : null;
        this.setScale(zoomStepScale(this.viewport.scale, e.deltaY, e.deltaMode));
      },
      { passive: false },
    );

    this.surface.on('pointerleave', (event: PointerEvent) => {
      const next = event.relatedTarget as Node | null;
      if (next && this.comments.contains(next)) return;
      this.hideCommentPopup();
    });

    // A canvas-backed sheet has no native focused cell. Establish the ordinary
    // A1 selection when its viewport receives keyboard focus, then reuse the
    // public selection contract for Arrow-key movement below.
    this.surface.on('focus', () => {
      if (this.currentWorksheet && !this.activeCell) this.setSelection('A1');
    });

    this.keydownHandler = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
        if (e.defaultPrevented || e.isComposing) return;
        const target = e.target as HTMLElement | null;
        const tag = target?.tagName;
        if (target?.isContentEditable || tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;
        e.preventDefault();
        void this.copySelection();
      } else if (
        !e.defaultPrevented && !e.isComposing &&
        !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey &&
        (e.key === 'ArrowUp' || e.key === 'ArrowDown' ||
          e.key === 'ArrowLeft' || e.key === 'ArrowRight')
      ) {
        const current = this.activeCell;
        const rowDelta = e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0;
        const colDelta = e.key === 'ArrowLeft'
          ? (this.isRtl ? 1 : -1)
          : e.key === 'ArrowRight'
            ? (this.isRtl ? -1 : 1)
            : 0;
        const next = current ? {
          row: Math.max(1, Math.min(MAX_WORKSHEET_ROW, current.row + rowDelta)),
          col: Math.max(1, Math.min(MAX_WORKSHEET_COL, current.col + colDelta)),
        } : { row: 1, col: 1 };
        e.preventDefault();
        this.hideCommentPopup();
        const ref = formatA1(next.row, next.col);
        this.setSelection(ref);
        // Selection already schedules the paint. Reuse the ordinary viewport
        // geometry without starting a second immediate render for every key.
        this._scrollCellIntoView(next.row, next.col);
        this.updateSelectionOverlay();
        this.updateFindOverlay();
        this.emitViewportChange();
      } else if (e.key === 'Escape' && this.validation.isOpen()) {
        this.hideValidationPanel();
      } else if (e.key === 'Escape' && this.comments.isOpen()) {
        this.hideCommentPopup();
      } else if (
        e.key === 'Enter' && this.activeCell &&
        !e.defaultPrevented && !e.isComposing &&
        !e.ctrlKey && !e.metaKey && !e.altKey
      ) {
        const comment = this.comments.commentAt(this.activeCell);
        if (comment) {
          e.preventDefault();
          this.hideCommentPopup();
          void this.comments.show(this.activeCell, comment)
            .catch((error) => this._reportRenderError(error));
        }
      }
    };
    this.surface.on('keydown', this.keydownHandler);
  }

  private buildTabs(): void {
    if (!this.sheetTabs) return;
    this.sheetTabs.build(this.workbook.sheetNames, this.workbook.tabColors);
  }

  private updateTabActive(index: number): void {
    this.sheetTabs?.setActive(index);
  }

  /**
   * IX9 {@link ZoomableViewer} — set the cell/header scale (`1` = 100%; the
   * viewer's `cellScale`) and re-lay-out the current sheet. Clamped to the zoom
   * bounds and snapped to whole percent; keeps the slider thumb, percentage label
   * in sync, and fires `onScaleChange` when the resolved scale actually changes.
   */
  setScale(scale: number): void {
    const zoomMin = this.opts.zoomMin ?? 0.1;
    const zoomMax = this.opts.zoomMax ?? 4;
    // Snap to whole percent so the label and cellScale stay tidy.
    const pct = Math.min(
      Math.round(zoomMax * 100),
      Math.max(Math.round(zoomMin * 100), Math.round(scale * 100)),
    );
    const next = pct / 100;
    const prevScale = this.viewport.scale;
    // Consume the gesture-only pointer anchor (Ctrl/⌘+wheel set it just above)
    // FIRST — before the no-op early return — so a gesture whose setScale ends
    // up a NO-OP (pinned at zoomMin/zoomMax, or a small deltaY swallowed by the
    // whole-percent snap) can never leak a stale anchor into a later non-gesture
    // setScale (slider, steppers, fitWidth/fitPage, public API), which must keep
    // the historical START-anchored (top-left) preservation. `null` for every
    // non-gesture source.
    const gestureAnchor = this._pendingZoomAnchor;
    this._pendingZoomAnchor = null;
    if (next === prevScale) return;
    this.viewport.setScale(next);

    this.zoomControl?.sync(next, pct, zoomMin, zoomMax);

    if (this.currentWorksheet) {
      // Preserve the START-anchored effective scroll position across the zoom.
      // The spacer (scrollWidth) is re-sized below, which changes maxScrollLeft;
      // for RTL the native scrollLeft is the inverse of the effective position,
      // so we must re-derive scrollLeft from the preserved effective value or
      // the view would jump toward the start on every zoom step.
      const prevEffective = this.effectiveScrollLeft;
      const prevScrollTop = this.viewportTop;
      // Gutter extents scale with cellScale (XL4); re-lay them out before the
      // spacer/scroll math reads canvasArea's new inset size.
      this.layoutGutters();
      this.updateSpacerSize(this.currentWorksheet);

      if (gestureAnchor) {
        // POINTER-ANCHORED zoom (both axes). The header + frozen band are drawn
        // at a FIXED screen position and do NOT scroll (see getCellAt), but their
        // on-screen size is the UNSCALED extent K × cs — a SCALING lead-in. From
        // getCellAt, the logical row under screen-y `py` is
        //   (py + scrollTop)/cs − K            (K = HEADER_H + frozenH)
        // and requiring that to be invariant across cs makes the K·cs terms
        // cancel exactly:
        //   scrollTop' = ratio·(scrollTop + py) − py
        // — i.e. the RAW pointer is the anchor and the clamp is the native
        // [0, maxScroll] (see anchoredZoomOffset's LEAD-INS note; routing through
        // a lead-in-shifted virtual scroll would distort the low clamp and floor
        // scrollTop at K·cs near the sheet start).

        // Vertical: native scrollTop is start-anchored in both LTR/RTL.
        this.viewportTop = anchoredZoomOffset(prevScrollTop, gestureAnchor.y, prevScale, next, {
          maxScroll: this.maxScrollTop,
        });

        // Horizontal: anchor in the logical-LTR space the grid math uses (the
        // same cancellation holds for K = HEADER_W + frozenW), so RTL is handled
        // by translating the pointer through screenX (an involution) and
        // re-deriving the native scrollLeft from the effective (start-anchored)
        // position, exactly as the START-anchored branch does.
        const anchorLogicalX = this.screenX(gestureAnchor.x, 0);
        const maxLeftV = this.maxScrollLeft;
        const newEffective = anchoredZoomOffset(prevEffective, anchorLogicalX, prevScale, next, {
          maxScroll: maxLeftV,
        });
        this.setViewportLeft(newEffective);
      } else {
        this.setViewportLeft(prevEffective);
      }
    }
    void this.renderCurrentSheet().catch((error) => this._reportRenderError(error));
    this.updateSelectionOverlay();
    this.updateFindOverlay();
    this.sheetTabs?.updateNavButtons();
    // IX9 change notification (fired last, after the view is consistent). Only
    // reached when `next` differs from the prior scale (early-returned above).
    this.opts.onScaleChange?.(next);
  }

  /** IX9 {@link ZoomableViewer} — the current zoom factor (`1` = 100%). This is
   *  the viewer's `cellScale`; `1` before anything is set. */
  getScale(): number {
    return this.viewport.scale;
  }

  /** IX9 {@link ZoomableViewer} — step up to the next rung of the shared zoom
   *  ladder (clamped to `zoomMax` by {@link setScale}). */
  zoomIn(): void {
    this.setScale(nextZoomStep(this.getScale()));
  }

  /** IX9 {@link ZoomableViewer} — step down to the next lower ladder rung. */
  zoomOut(): void {
    this.setScale(prevZoomStep(this.getScale()));
  }

  /**
   * IX9 {@link ZoomableViewer} — fit the used data range's WIDTH to the canvas
   * area. The "content" is the natural (100%) width of the row header plus the
   * used columns; the container is `canvasArea.clientWidth`. A no-op (defers) when
   * nothing is loaded or the container is unlaid-out. Routes through
   * {@link setScale}, so the result is clamped/snapped and fires `onScaleChange`.
   */
  fitWidth(): void {
    this._fit('width');
  }

  /**
   * IX9 {@link ZoomableViewer} — fit the used data range's WIDTH AND HEIGHT inside
   * the canvas area (header + used columns/rows), so the whole used range is
   * visible without scrolling. Takes the tighter of the width- and height-fit
   * factors. Defers when unloaded / unlaid-out; routes through {@link setScale}.
   */
  fitPage(): void {
    this._fit('page');
  }

  /** Shared fit implementation for {@link fitWidth} / {@link fitPage}: derive the
   *  natural (cs=1) content extent of the used data range, ask core's pure
   *  {@link fitScale} for the factor, and apply it via {@link setScale}. */
  private _fit(mode: 'width' | 'page'): void {
    const ws = this.currentWorksheet;
    if (!ws) return;
    const { width, height } = this._naturalContentExtent(ws);
    const scale = fitScale(
      {
        contentWidth: width,
        contentHeight: height,
        containerWidth: this.canvasArea.clientWidth,
        containerHeight: this.canvasArea.clientHeight,
      },
      mode,
    );
    if (scale <= 0) return; // unlaid-out / empty — defer (fitScale's 0 sentinel)
    this.setScale(scale);
  }

  /** Natural (unscaled, cs=1) CSS-px extent of a worksheet's used data range:
   *  the row/column header plus every used column width / row height. Mirrors
   *  {@link updateSpacerSize} at cs=1 (same used-range detection) so the fit
   *  targets exactly the region the spacer/scroll extent covers. */
  private _naturalContentExtent(ws: Worksheet): { width: number; height: number } {
    const { maxRow, maxCol } = worksheetContentBounds(ws);
    return getGridGeometryForWorksheet(ws).logicalContentExtent(
      maxRow,
      maxCol,
      HEADER_W,
      HEADER_H,
    );
  }

  private updateSpacerSize(ws: Worksheet): void {
    const cs = this.viewport.scale;
    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // Find actual scrollable data extent
    let { maxRow, maxCol } = worksheetContentBounds(ws);
    maxRow += 30;
    maxCol += 10;

    // Spacer = rounded header + cumulative per-band-rounded geometry.
    const extent = getGridGeometryForWorksheet(ws).roundedContentExtent(
      maxRow,
      maxCol,
      cs,
      HEADER_W,
      HEADER_H,
    );
    const totalW = extent.width;
    const totalH = extent.height;

    this.spacer.style.width = `${totalW}px`;
    this.spacer.style.height = `${totalH}px`;
    this.viewport.setViewportSize(this.scrollHost.clientWidth, this.scrollHost.clientHeight);
    this.viewport.setExtent(totalW, totalH);
    this.setViewportLeft(this.viewport.x);
    this.viewportTop = this.viewport.y;
  }

  /**
   * Coalesce a re-render into the next animation frame. Called from the
   * high-frequency event-driven paths (scroll, live column/row resize, drag-
   * selection, container resize); a burst of these within one frame schedules a
   * single {@link renderCurrentSheet}, avoiding the previous behavior where every
   * scroll event forced its own synchronous full redraw. Already-scheduled frames
   * are not re-scheduled — the one pending render reads the live scroll/scale
   * state when it runs, so the most recent position always wins without threading
   * a coordinate through. Falls back to a synchronous render when
   * `requestAnimationFrame` is unavailable (e.g. a non-DOM host), preserving the
   * old semantics there.
   */
  private scheduleRender(): void {
    this.renderDispatcher.schedule(() =>
      this.renderCurrentSheet().catch((error) => this._reportRenderError(error)));
  }

  private async renderCurrentSheet(): Promise<void> {
    const generation = this.renderDispatcher.begin();
    try {
      await this._renderCurrentSheet(generation);
    } catch (err) {
      if (!this.renderDispatcher.isCurrent(generation)) return;
      throw err;
    }
  }

  /** Route a render failure to `onError`, or `console.error` when none is given
   *  (never fully silent), and never after teardown. Mirrors the scroll viewers'
   *  `_reportRenderError`. */
  private _reportRenderError(err: unknown): void {
    if (this._destroyed) return;
    const e = err instanceof Error ? err : new Error(String(err));
    if (this.opts.onError) this.opts.onError(e);
    else console.error('[ooxml] XlsxViewer render failed:', e);
  }

  private async _renderCurrentSheet(seq: number): Promise<void> {
    if (!this.currentWorksheet) return;
    if (this.previewCompletion) {
      if (this.firstPreviewRender) {
        const prepared = this.previewPreparedViewport;
        if (!prepared || this.canvasArea.clientWidth !== prepared.width ||
            this.canvasArea.clientHeight !== prepared.height ||
            this.viewport.scale !== prepared.scale ||
            this.viewportTop !== 0 || this.effectiveScrollLeft !== 0) return;
      } else await this.previewCompletion;
      if (!this.renderDispatcher.isCurrent(seq) || this._destroyed) return;
    }
    const ws = this.currentWorksheet;
    const w = this.canvasArea.clientWidth;
    const h = this.canvasArea.clientHeight;
    if (w <= 0 || h <= 0) return;

    // Claim a render generation up front so a later render started while this one
    // awaits the worker can mark this frame stale (worker mode only; see below).
    const cs = this.viewport.scale;
    const dpr = this.surface.dpr;

    const freezeRows = ws.freezeRows ?? 0;
    const freezeCols = ws.freezeCols ?? 0;

    // DOM scrollLeft/scrollTop are in scaled (physical) CSS pixels.
    // Convert to logical pixels for cell-finding by dividing by cs. For RTL
    // sheets effectiveScrollLeft inverts the native scrollLeft so that 0 = col A
    // at the (mirrored) right edge — see the getter for the rationale.
    const visible = getGridGeometryForWorksheet(ws).visibleRange({
      width: w,
      height: h,
      scale: cs,
      scrollX: this.effectiveScrollLeft,
      scrollY: this.viewportTop,
      headerWidth: HEADER_W,
      headerHeight: HEADER_H,
      buffer: 2,
    });
    const viewport: ViewportRange = visible.range;
    const { offsetX, offsetY } = visible;

    const { selectedRowRange, selectedColRange } = this.computeHeaderHighlight();

    const renderOpts = {
      width: w,
      height: h,
      dpr,
      imageResources: this.opts.imageResources,
      cellScale: cs,
      scrollOffsetX: offsetX,
      scrollOffsetY: offsetY,
      freezeRows,
      freezeCols,
      selectedRowRange,
      selectedColRange,
      chromeColors: this.chromeColors,
    };

    const sizeProjection = this.wireSizeOverrides();
    const viewerRenderOpts = withViewerRenderContext(
      sizeProjection ? { ...renderOpts, sizeOverrides: sizeProjection.overrides } : renderOpts,
      getGridGeometryForWorksheet(ws).maximumDigitWidth,
      {
        worksheet: ws,
        projection: sizeProjection
          ? { id: this.projectionId, revision: sizeProjection.revision, autoRowHeightsPrepared: true }
          : undefined,
      },
    );

    if (this._mode === 'worker') {
      // Render the viewport off the main thread and paint the returned bitmap.
      // The selection overlay (geometry-based, from getCellRect) is unaffected.
      // Attach the cumulative view-only size overrides (outline collapse/
      // expand, drag resize) so the worker re-lays the mutated bands — its
      // local sheet cache never sees main-thread model writes on its own.
      const bmp = await this.workbook.renderViewportToBitmap(
        this.currentSheet,
        viewport,
        viewerRenderOpts,
      );
      if (!this.renderDispatcher.commitBitmap(seq, bmp, w, h)) return;
    } else {
      await this.workbook.renderViewport(
        this.canvas,
        this.currentSheet,
        viewport,
        withXlsxRenderCommitGuard(viewerRenderOpts, () =>
          !this._destroyed && this.renderDispatcher.isCurrent(seq),
        ),
      );
      if (!this.renderDispatcher.isCurrent(seq) || this._destroyed) return;
    }
    // XL4: repaint the outline gutters over the fresh grid frame, aligned to the
    // same scroll offset. No-op when the sheet has no outlining.
    this.renderGutters();
    this.firstPreviewRender = false;
    this.committedFrameCount++;
  }

  private computeHeaderHighlight(): {
    selectedRowRange: { start: number; end: number; strong: boolean } | null;
    selectedColRange: { start: number; end: number; strong: boolean } | null;
  } {
    return this.selectionController.headerHighlight();
  }

  get sheetNames(): string[] {
    return this.wb?.sheetNames ?? [];
  }

  /** The underlying <canvas> element the grid is drawn on. */
  get canvasElement(): HTMLCanvasElement {
    return this.canvas;
  }

  /** Latest content-free resource metrics for the loaded workbook. */
  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    if (!this.wb) throw new Error('Workbook not loaded');
    return await this.wb.getResourceMetrics();
  }

  /**
   * Tear down the viewer and release resources.
   *
   * The caller's container is returned to the state it had before construction
   * (empty): the entire wrapper subtree the constructor appended is removed.
   * All document-level listeners are detached — the keydown handler here, and
   * the validation-panel outside-click handler via {@link hideValidationPanel}.
   * Listeners on elements inside the wrapper (scrollHost, tabs, …) need no
   * explicit removal: removing the subtree makes them unreachable and eligible
   * for GC. Safe to call more than once.
   *
   * NOTE: the shared `<style>` in the owning document is intentionally NOT removed —
   * it is a class constant that any still-live viewer may depend on, and one
   * leftover sheet is a bounded, harmless cost (see {@link ensureViewerStyleInjected}).
   */
  destroy(): void {
    if (this._destroyed) return;
    // First line: block any render rejection racing in from surfacing on a dead
    // viewer (checked at the top of _reportRenderError). The acquisition owner
    // invalidates any load still in flight below.
    this._destroyed = true;
    if (this.selectionContextNotificationFrame !== null) {
      this.hostWindow.cancelAnimationFrame(this.selectionContextNotificationFrame);
      this.selectionContextNotificationFrame = null;
    }
    this.selectionContextNotificationMicrotask = false;
    this.stopSelectionAutoScroll();
    this.sheetRequestGeneration++;
    this.resizeObserver?.disconnect();
    this.chromeStyleObserver?.disconnect();
    this.chromeStyleObserver = null;
    if (this.chromeSchemeMedia && this.chromeSchemeListener) {
      this.chromeSchemeMedia.removeEventListener?.('change', this.chromeSchemeListener);
    }
    this.chromeSchemeMedia = null;
    this.chromeSchemeListener = null;
    this.renderDispatcher.destroy();
    this.surface.destroy();
    this.sheetTabs?.destroy();
    this.zoomControl?.destroy();
    this.comments.destroy();
    this.validation.destroy();
    // IX2 — drop the find state (matches + cursor) so a stale
    // findNext()/findPrev() after teardown returns null instead of a match
    // pointing into a dead viewer (same fix as DocxViewer/PptxViewer.destroy).
    this._find.invalidate();
    this.releaseHostFonts();
    const releaseProjection = this.wb?.[releaseXlsxViewerProjection];
    if (typeof releaseProjection === 'function') {
      releaseProjection.call(this.wb, this.projectionId);
    }
    this.currentWorksheet = null;
    this.releaseCurrentWorksheet?.();
    this.releaseCurrentWorksheet = null;
    this.sheetViews.clear();
    this.viewEdits.destroy();
    this.currentSourceComments = [];
    this.sourceCommentMap.clear();
    this.hyperlinkMap.clear();
    this.preparedWorkbook = null;
    this.outlineGutter.destroy();
    this.elementContext = null;
    this.pendingElementClick = null;
    this.selectionController.reset();
    this.lastNotifiedSelectionState = null;
    this.finishSelectionNotificationChain();
    this.acquisition.destroy();
    // Remove the whole UI subtree so the container is empty again. This also
    // detaches every listener bound to elements within it (scrollHost pointer/
    // wheel handlers, tab clicks, zoom slider) without per-element cleanup.
    this.wrapper.remove();
  }

  private assertOpen(): void {
    if (this._destroyed) throw this.destroyedError();
  }

  private destroyedError(): Error {
    return new Error(this._mountKind === 'sheet'
      ? 'XlsxSheetViewer is destroyed'
      : 'XlsxViewer is destroyed');
  }
}

/** Workbook viewer mounted into a container with scrollable grid, sheet tabs,
 * outline gutters, and optional zoom chrome. */
export class XlsxViewer extends XlsxViewerEngine {
  /**
   * Create a workbook Viewer that borrows an already-loaded workbook.
   * Destroying the Viewer leaves the caller-owned workbook open.
   */
  static fromWorkbook(
    container: HTMLElement,
    workbook: XlsxWorkbook,
    opts: Omit<XlsxViewerOptions, keyof LoadOptions> = {},
  ): Omit<XlsxViewer, 'load'> {
    return new XlsxViewer(container, {
      ...opts,
      [borrowedWorkbookOption]: workbook,
    } as InternalXlsxViewerOptions);
  }

  constructor(container: HTMLElement, opts: XlsxViewerOptions = {}) {
    super(container, opts, { kind: 'composite' });
  }

  /** Load an OOXML workbook from a URL or ArrayBuffer. */
  async load(source: string | ArrayBuffer): Promise<void> {
    await this[loadXlsxViewerSource](source);
  }
}

type XlsxSheetViewerSnapshot = Readonly<{
  sheetIndex: number;
  sheetCount: number;
  sheetNames: string[];
  viewport: XlsxViewportOffset;
  selectionState: XlsxSelectionState | null;
  scale: number;
  hiddenSheetMode: HiddenSheetMode;
  visibleSheetCount: number;
}>;

/**
 * Canvas-mounted active-sheet viewer. It instantiates the same workbook,
 * acquisition, geometry, selection, overlay, and render-dispatch engine as
 * {@link XlsxViewer}, but mounts no workbook footer or sheet-tab chrome.
 */
export class XlsxSheetViewer implements ZoomableViewer {
  private readonly engine: XlsxViewerEngine;
  private readonly canvasMount: CallerCanvasMount;
  private destroyed = false;
  private snapshot: XlsxSheetViewerSnapshot;
  private lastMetrics: OoxmlResourceMetrics | undefined;

  /**
   * Create a sheet Viewer that borrows an already-loaded workbook.
   * Destroying the Viewer leaves the caller-owned workbook open.
   */
  static fromWorkbook(
    canvasElement: HTMLCanvasElement,
    workbook: XlsxWorkbook,
    options: Omit<XlsxSheetViewerOptions, keyof LoadOptions> = {},
  ): Omit<XlsxSheetViewer, 'load'> {
    return new XlsxSheetViewer(canvasElement, {
      ...options,
      [borrowedWorkbookOption]: workbook,
    } as InternalXlsxViewerOptions);
  }

  constructor(
    readonly canvasElement: HTMLCanvasElement,
    options: XlsxSheetViewerOptions = {},
  ) {
    const borrowedWorkbook = (options as InternalXlsxViewerOptions)[borrowedWorkbookOption];
    const mode = resolveCanvasViewerMode('XlsxSheetViewer', options.mode, borrowedWorkbook);
    const rect = canvasElement.getBoundingClientRect();
    this.canvasMount = new CallerCanvasMount(canvasElement, {
      wrapperCssText:
        `position:relative;display:inline-block;vertical-align:top;overflow:hidden;` +
        `width:${canvasElement.style.width || `${rect.width || canvasElement.width}px`};` +
        `height:${canvasElement.style.height || `${rect.height || canvasElement.height}px`};`,
      restoreMode: 'style-and-bitmap',
    });
    this.engine = new XlsxViewerEngine(this.canvasMount.wrapper, {
      ...options,
      onResourceMetrics: (metrics) => {
        this.lastMetrics = metrics;
        options.onResourceMetrics?.(metrics);
      },
    }, {
      kind: 'sheet',
      canvas: canvasElement,
      mode,
    });
    this.snapshot = {
      sheetIndex: 0,
      sheetCount: 0,
      sheetNames: [],
      viewport: { x: 0, y: 0 },
      selectionState: null,
      scale: this.engine.getScale(),
      hiddenSheetMode: this.engine.hiddenSheetMode,
      visibleSheetCount: 0,
    };
  }

  /**
   * Load an XLSX worksheet, or reuse the XLSX sheet renderer for one explicitly
   * selected delimited-text source. URL/ArrayBuffer input, reload replacement,
   * callbacks, and destroy ownership match XLSX loading.
   */
  async load(
    source: string | ArrayBuffer,
    options: XlsxSheetLoadOptions = {},
  ): Promise<void> {
    this.assertOpen();
    try {
      await this.engine[loadXlsxViewerSource](source, options);
    } finally {
      if (!this.destroyed) this.captureSnapshot();
    }
    this.assertOpen();
  }

  get sheetIndex(): number { return this.destroyed ? this.snapshot.sheetIndex : this.engine.sheetIndex; }
  get sheetCount(): number { return this.destroyed ? this.snapshot.sheetCount : this.engine.sheetCount; }
  get sheetNames(): string[] {
    return this.destroyed ? [...this.snapshot.sheetNames] : [...this.engine.sheetNames];
  }

  async goToSheet(index: number): Promise<void> {
    this.assertOpen();
    await this.engine.goToSheet(index);
    this.assertOpen();
    this.captureSnapshot();
  }

  async nextSheet(): Promise<void> {
    this.assertOpen();
    await this.engine.nextSheet();
    this.assertOpen();
    this.captureSnapshot();
  }

  async prevSheet(): Promise<void> {
    this.assertOpen();
    await this.engine.prevSheet();
    this.assertOpen();
    this.captureSnapshot();
  }

  getViewportOffset(): XlsxViewportOffset {
    return this.destroyed ? { ...this.snapshot.viewport } : this.engine.getViewportOffset();
  }

  async setViewportOffset(offset: XlsxViewportOffset): Promise<void> {
    this.assertOpen();
    await this.engine.setViewportOffset(offset);
    this.assertOpen();
    this.captureSnapshot();
  }

  async scrollToCell(ref: string, options?: XlsxScrollToCellOptions): Promise<void> {
    this.assertOpen();
    await this.engine.scrollToCell(ref, options);
    this.assertOpen();
    this.captureSnapshot();
  }

  async relayout(): Promise<void> {
    this.assertOpen();
    // The caller canvas remains the sizing authority. A caller may update its
    // inline/CSS box and then call relayout(); promote that box to the mount so
    // the shared engine measures the new viewport rather than its old wrapper.
    const rect = this.canvasElement.getBoundingClientRect();
    if (rect.width > 0) this.canvasMount.wrapper.style.width = `${rect.width}px`;
    if (rect.height > 0) this.canvasMount.wrapper.style.height = `${rect.height}px`;
    await this.engine.relayout();
    this.assertOpen();
    this.captureSnapshot();
  }

  getScale(): number { return this.destroyed ? this.snapshot.scale : this.engine.getScale(); }

  setScale(scale: number): void {
    this.assertOpen();
    this.engine.setScale(scale);
    this.captureSnapshot();
  }

  zoomIn(): void { this.assertOpen(); this.engine.zoomIn(); this.captureSnapshot(); }
  zoomOut(): void { this.assertOpen(); this.engine.zoomOut(); this.captureSnapshot(); }
  fitWidth(): void { this.assertOpen(); this.engine.fitWidth(); this.captureSnapshot(); }
  fitPage(): void { this.assertOpen(); this.engine.fitPage(); this.captureSnapshot(); }

  getCellAt(clientX: number, clientY: number): CellAddress | null {
    return this.destroyed ? null : this.engine.getCellAt(clientX, clientY);
  }

  getCellViewportRect(cell: CellAddress | string): XlsxCellViewportRect | null {
    return this.destroyed ? null : this.engine.getCellViewportRect(cell);
  }

  /** Detached comments for the current sheet, in authored order. */
  getComments(): readonly Readonly<XlsxComment>[] {
    this.assertOpen();
    return this.engine.getComments();
  }

  async goToComment(
    sheetIndex: number,
    cellRef: string,
    options?: XlsxScrollToCellOptions,
  ): Promise<boolean> {
    this.assertOpen();
    const found = await this.engine.goToComment(sheetIndex, cellRef, options);
    this.assertOpen();
    this.captureSnapshot();
    return found;
  }

  get selectionState(): XlsxSelectionState | null {
    const value = this.destroyed ? this.snapshot.selectionState : this.engine.selectionState;
    return value ? structuredClone(value) : null;
  }

  setSelection(selection: XlsxSelectionInput): void {
    this.assertOpen();
    this.engine.setSelection(selection);
    this.captureSnapshot();
  }

  getSelectionContext(options?: XlsxSelectionContextOptions): XlsxSelectionContext | null {
    this.assertOpen();
    return this.engine.getSelectionContext(options);
  }

  async copySelection(): Promise<XlsxCopyResult> {
    this.assertOpen();
    return await this.engine.copySelection();
  }

  setSelectionColor(color: string): void {
    this.assertOpen();
    this.engine.setSelectionColor(color);
  }

  async setHiddenSheetMode(mode: HiddenSheetMode): Promise<void> {
    this.assertOpen();
    await this.engine.setHiddenSheetMode(mode);
    this.assertOpen();
    this.captureSnapshot();
  }

  get hiddenSheetMode(): HiddenSheetMode {
    return this.destroyed ? this.snapshot.hiddenSheetMode : this.engine.hiddenSheetMode;
  }

  get visibleSheetCount(): number {
    return this.destroyed ? this.snapshot.visibleSheetCount : this.engine.visibleSheetCount;
  }

  async findText(
    query: string,
    options?: FindMatchesOptions,
  ): Promise<FindMatch<XlsxMatchLocation>[]> {
    this.assertOpen();
    const matches = await this.engine.findText(query, options);
    this.assertOpen();
    return matches;
  }

  async findNext(): Promise<FindMatch<XlsxMatchLocation> | null> {
    this.assertOpen();
    const match = await this.engine.findNext();
    this.assertOpen();
    this.captureSnapshot();
    return match;
  }

  async findPrev(): Promise<FindMatch<XlsxMatchLocation> | null> {
    this.assertOpen();
    const match = await this.engine.findPrev();
    this.assertOpen();
    this.captureSnapshot();
    return match;
  }

  clearFind(): void { this.assertOpen(); this.engine.clearFind(); }

  async getResourceMetrics(): Promise<OoxmlResourceMetrics> {
    if (this.destroyed) {
      if (this.lastMetrics) return this.lastMetrics;
      throw this.destroyedError();
    }
    this.lastMetrics = await this.engine.getResourceMetrics();
    return this.lastMetrics;
  }

  destroy(): void {
    if (this.destroyed) return;
    this.captureSnapshot();
    this.destroyed = true;
    this.engine.destroy();

    this.canvasMount.restore();
  }

  private captureSnapshot(): void {
    const selectionState = this.engine.selectionState;
    this.snapshot = {
      sheetIndex: this.engine.sheetIndex,
      sheetCount: this.engine.sheetCount,
      sheetNames: [...this.engine.sheetNames],
      viewport: { ...this.engine.getViewportOffset() },
      selectionState: selectionState ? structuredClone(selectionState) : null,
      scale: this.engine.getScale(),
      hiddenSheetMode: this.engine.hiddenSheetMode,
      visibleSheetCount: this.engine.visibleSheetCount,
    };
  }

  private assertOpen(): void {
    if (this.destroyed) throw this.destroyedError();
  }

  private destroyedError(): Error {
    return new Error('XlsxSheetViewer is destroyed');
  }
}
