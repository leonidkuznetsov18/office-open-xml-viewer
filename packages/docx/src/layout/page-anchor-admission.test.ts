import { describe, expect, it, vi } from 'vitest';

// Page-owned anchor admission (issue #1659): page/margin-positioned floating
// tables (§17.4.57) and page/margin-relative drawings (§20.4.3.4-5) whose
// registration couples with body flow. Pagination here is independent of
// glyph metrics: each paragraph is one short glyph on an exact 24 pt line
// (§17.3.1.33), so the 468 pt column (US Letter, 1 in margins) holds 27 lines.

const counters = vi.hoisted(() => ({ passes: 0, failures: 0, limit: Number.POSITIVE_INFINITY }));
vi.mock('./runtime-state.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./runtime-state.js')>();
  return {
    ...actual,
    bodyLayoutKernelOf: (services: Parameters<typeof actual.bodyLayoutKernelOf>[0]) => {
      const kernel = actual.bodyLayoutKernelOf(services);
      return kernel && {
        openBodyLayoutSession: (...args: Parameters<typeof kernel.openBodyLayoutSession>) => {
          counters.passes += 1;
          return kernel.openBodyLayoutSession(...args);
        },
      };
    },
  };
});
// Counts convergence runs that end in the defensive fallback, and lets a test
// force it by lowering the operational pass guard.
vi.mock('./convergence.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./convergence.js')>();
  return {
    ...actual,
    convergeExactStateSteps: function* (
      options: Parameters<typeof actual.convergeExactStateSteps>[0],
    ) {
      try {
        return yield* actual.convergeExactStateSteps({
          ...options,
          limit: Math.max(2, Math.min(options.limit, counters.limit)),
        });
      } catch (error) {
        counters.failures += 1;
        throw error;
      }
    },
  };
});

import { createLayoutServices } from '../layout-runtime.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import type { BodyElement, DocxDocumentModel } from '../types.js';
import {
  paginateBody,
  resolvePageOwnedAnchors,
  type BodyPaginationPassResult,
} from './body-paginator.js';
import type { PageAnchorInputEvent } from './anchor-line-deferral.js';
import { sourceKey } from './source-key.js';
import { layoutFingerprint } from './invariants.js';
import { normalizeLayoutOptions } from './options.js';
import type { DocumentLayout, LayoutRect } from './types.js';

const BODY_TOP = 72;

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
  } as unknown as CanvasRenderingContext2D;
}

function line(extraRuns: readonly unknown[] = []): BodyElement {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0,
    lineSpacing: { value: 24, rule: 'exact', explicit: true },
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: 'x', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'NotInMetrics',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }, ...extraRuns],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as BodyElement;
}

/** A paragraph anchoring an image positioned against the page or margin
 * (page-owned, §20.4.3.4-5) or against its paragraph (host-owned). */
function anchoredImageLine(spec: Readonly<{
  widthPt: number; heightPt: number; yPt: number;
  relativeFrom: 'page' | 'margin' | 'paragraph'; wrap: 'topAndBottom' | 'square';
}>, paragraph: Partial<Record<'keepNext' | 'keepLines' | 'widowControl', boolean>> = {}): BodyElement {
  return {
    ...line([{
      type: 'image', imagePath: 'word/media/anchor.png', mimeType: 'image/png',
      widthPt: spec.widthPt, heightPt: spec.heightPt, anchor: true,
      anchorXPt: 0, anchorYPt: spec.yPt, anchorXFromMargin: true,
      anchorYFromPara: spec.relativeFrom === 'paragraph',
      wrapMode: spec.wrap, wrapSide: 'bothSides',
      anchorXRelativeFrom: 'margin', anchorYRelativeFrom: spec.relativeFrom,
    }]),
    ...paragraph,
  } as BodyElement;
}

/** A paragraph of `count` text lines (each glyph is 5 pt wide at 10 pt). */
function textParagraph(
  count: number,
  widthPt: number,
  paragraph: Partial<Record<'keepNext' | 'keepLines' | 'widowControl', boolean>> = {},
): BodyElement {
  const base = line() as BodyElement & { runs: Array<Record<string, unknown>> };
  const glyphs = Math.max(1, Math.floor(widthPt / 5) - 1) * count;
  return {
    ...base,
    runs: [{ ...base.runs[0]!, text: 'x'.repeat(Math.max(1, glyphs - 1)) }],
    ...paragraph,
  } as BodyElement;
}

type Columns = Readonly<{ widthsPt: readonly number[]; spacePt: number }> | null;

function columnsSpec(columns: Columns) {
  if (!columns) return null;
  return {
    count: columns.widthsPt.length, spacePt: columns.spacePt,
    equalWidth: false, sep: false,
    cols: columns.widthsPt.map((widthPt) => ({ widthPt, spacePt: columns.spacePt })),
  };
}

/** A section break ending a section of the given page size and columns. */
function sectionBreak(
  kind: 'nextPage' | 'continuous',
  pageWidth: number,
  pageHeight: number,
  columns: Columns,
): BodyElement {
  return {
    type: 'sectionBreak', kind, columns: columnsSpec(columns),
    geom: {
      pageWidth, pageHeight, marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36,
    },
  } as unknown as BodyElement;
}

/** A page/margin-positioned floating table (§17.4.57) of exact-height rows. */
function floatingTable(spec: Readonly<{
  widthPt: number; heightPt: number | readonly number[]; tblpY: number;
  vertAnchor: 'page' | 'margin'; distancePt: number;
}>): BodyElement {
  const noBorders = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
  return {
    type: 'table',
    colWidths: [spec.widthPt],
    rows: (typeof spec.heightPt === 'number' ? [spec.heightPt] : spec.heightPt).map((rowHeight) => ({
      cells: [{
        content: [], colSpan: 1, vMerge: null, borders: noBorders,
        background: null, vAlign: 'top', widthPt: null,
      }],
      rowHeight, rowHeightRule: 'exact', isHeader: false,
    })),
    borders: noBorders,
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    tblpPr: {
      leftFromText: spec.distancePt, rightFromText: spec.distancePt,
      topFromText: spec.distancePt, bottomFromText: spec.distancePt,
      horzAnchor: 'margin', horzSpecified: true, vertAnchor: spec.vertAnchor,
      tblpX: 0, tblpY: spec.tblpY,
    },
  } as unknown as BodyElement;
}

function documentModel(
  body: BodyElement[],
  final: Readonly<{ pageWidth: number; pageHeight: number; columns: Columns }> = {
    pageWidth: 612, pageHeight: 792, columns: null,
  },
): DocxDocumentModel {
  return {
    section: {
      pageWidth: final.pageWidth, pageHeight: final.pageHeight,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage', columns: columnsSpec(final.columns),
    },
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

function lines(count: number): BodyElement[] {
  return Array.from({ length: count }, () => line());
}

function layout(model: DocxDocumentModel) {
  const before = { passes: counters.passes, failures: counters.failures };
  const source = layoutSourceStore(model);
  const result = paginateBody(
    source.bodyLayoutInput,
    createLayoutServices(source, { measureContext: measureContext() }),
    normalizeLayoutOptions(undefined, 1_700_000_000_000),
  );
  return {
    layout: result,
    passes: counters.passes - before.passes,
    fallbacks: counters.failures - before.failures,
  };
}

const overlaps = (a: LayoutRect, b: LayoutRect) => (
  a.xPt < b.xPt + b.widthPt - 0.01 && b.xPt < a.xPt + a.widthPt - 0.01
  && a.yPt < b.yPt + b.heightPt - 0.01 && b.yPt < a.yPt + a.heightPt - 0.01
);

/**
 * No text line overlaps a body table or a page-owned drawing, tables do not
 * overlap each other, and every vertical gap from the body top to the last
 * text line of a page is explained by a table or drawing on that page (no
 * phantom wrap left by a stale registration).
 */
function expectWellFormed(result: DocumentLayout): void {
  for (const page of result.pages) {
    const tables = page.layers.body.filter((node) => node.kind === 'table')
      .map((node) => node.flowBounds);
    // Text is checked against page-owned root fragments. A row continuing in
    // a later column or page is host-flow ("fresh-text" frame); its geometry
    // is outside page-owned admission.
    const pageOwnedTables = page.layers.body.filter((node) => (
      node.kind === 'table' && node.sectionFlowOwnership === 'page'
    )).map((node) => node.flowBounds);
    const drawings = page.layers.body.flatMap((node) => node.kind === 'paragraph'
      ? node.drawings.filter((drawing) => drawing.anchorLayer?.verticalOwnership === 'page')
        .map((drawing) => drawing.flowBounds)
      : []);
    // Host-relative drawings may explain wrap gaps; their own overlap with
    // earlier lines is outside page-owned admission.
    const hostDrawings = page.layers.body.flatMap((node) => node.kind === 'paragraph'
      ? node.drawings.filter((drawing) => drawing.anchorLayer?.verticalOwnership === 'host')
        .map((drawing) => drawing.flowBounds)
      : []);
    const text = page.layers.body.flatMap((node) => node.kind === 'paragraph' ? node.lines : [])
      .map((entry) => entry.bounds);
    tables.forEach((table, index) => {
      for (const other of tables.slice(index + 1)) {
        expect(overlaps(table, other), `page ${page.pageIndex} tables`).toBe(false);
      }
    });
    for (const box of text) {
      for (const obstacle of [...pageOwnedTables, ...drawings]) {
        expect(overlaps(box, obstacle), `page ${page.pageIndex} text over float`).toBe(false);
      }
    }
    if (text.length === 0) continue;
    // A float explains a gap within its wrap distance (at most 18 pt here),
    // plus one 24 pt line that could not fit above it.
    const covered = [
      ...text.map((box) => [box.yPt, box.yPt + box.heightPt] as const),
      ...[...tables, ...drawings, ...hostDrawings].map((box) => [
        box.yPt - 18 - 24 - 0.01, box.yPt + box.heightPt + 18.01,
      ] as const),
    ].sort((left, right) => left[0] - right[0]);
    const lastText = Math.max(...text.map((box) => box.yPt + box.heightPt));
    let reach = BODY_TOP;
    for (const [top, bottom] of covered) {
      if (top > reach + 0.01 && top < lastText) {
        expect.fail(`page ${page.pageIndex}: unexplained gap ${reach}..${top}`);
      }
      reach = Math.max(reach, bottom);
    }
  }
}

type Pick = <T>(values: readonly T[]) => T;

const PAGE_SIZES = [[612, 792], [792, 612], [595, 842]] as const;

function randomColumns(pick: Pick, pageWidth: number): Columns {
  const count = pick([1, 1, 2, 3]);
  if (count === 1) return null;
  const spacePt = pick([12, 18, 36]);
  const available = pageWidth - 144 - spacePt * (count - 1);
  const shares = Array.from({ length: count }, () => pick([1, 2, 3]));
  const total = shares.reduce((sum, share) => sum + share, 0);
  const widthsPt = shares.map((share) => Math.floor((available * share) / total));
  return { widthsPt, spacePt };
}

/**
 * A random document of paragraphs (single- and multi-line, keepNext chains,
 * keepLines, widow control), page/margin floating tables (one or several
 * exact rows), page/margin and paragraph-relative drawings, in up to three
 * sections with different page sizes and 1-3 unequal columns. Returns the
 * page-owned anchor count K and C, the keepNext paragraphs plus the anchoring
 * paragraphs that keep lines or control widows, for the bound 1 + 4(2K + C).
 */
function randomDocument(pick: Pick) {
  const body: BodyElement[] = [];
  let anchors = 0;
  let kept = 0;
  const sections = pick([1, 1, 2, 3]);
  let final = { pageWidth: 612, pageHeight: 792, columns: null as Columns };
  for (let section = 0; section < sections; section += 1) {
    const [pageWidth, pageHeight] = pick(PAGE_SIZES);
    const columns = randomColumns(pick, pageWidth);
    const columnWidth = columns ? Math.min(...columns.widthsPt) : pageWidth - 144;
    const count = pick([1, 2, 3, 4]);
    for (let index = 0; index < count; index += 1) {
      body.push(...lines(pick([0, 3, 12, 20, 24, 25, 26, 30])));
      const chain = pick([0, 0, 1, 2]);
      for (let member = 0; member < chain; member += 1) {
        const lineCount = pick([1, 3]);
        const keepLines = pick([false, true]);
        const widowControl = pick([false, true]);
        body.push(textParagraph(lineCount, columnWidth, { keepNext: true, keepLines, widowControl }));
        kept += 1;
      }
      const kind = pick(['table', 'page-drawing', 'page-drawing', 'host-drawing'] as const);
      if (kind === 'table') {
        anchors += 1;
        body.push(floatingTable({
          widthPt: pick([150, 300, 468, 524]),
          heightPt: pick([24, 120, 360, 600, 648, [120, 48, 48, 48, 240], [200, 200]]),
          tblpY: pick([0, 36, 72, 144, 400]),
          vertAnchor: pick(['page', 'margin'] as const),
          distancePt: pick([0, 9, 18]),
        }));
      } else {
        if (kind === 'page-drawing') anchors += 1;
        const keepLines = pick([false, false, true]);
        const widowControl = pick([false, true]);
        kept += kind === 'page-drawing' && (keepLines || widowControl) ? 1 : 0;
        body.push(anchoredImageLine({
          widthPt: pick([150, 468]),
          heightPt: pick([48, 120, 200, 360, 500, 624]),
          yPt: pick([0, 36, 72, 144]),
          relativeFrom: kind === 'host-drawing'
            ? 'paragraph' : pick(['page', 'margin'] as const),
          wrap: pick(['topAndBottom', 'square'] as const),
        }, { keepLines, widowControl }));
      }
    }
    body.push(...lines(pick([0, 1, 2])));
    if (section < sections - 1) {
      body.push(sectionBreak(pick(['nextPage', 'continuous'] as const), pageWidth, pageHeight, columns));
    } else {
      final = { pageWidth, pageHeight, columns };
    }
  }
  return { model: documentModel(body, final), anchors, kept };
}

describe('page-owned anchor admission', () => {
  it('defers a drawing anchor line past a page-anchored table without a phantom wrap', () => {
    // Review repro: registering the 200 pt image on page 0 pushes its anchor
    // line to page 1, after which the table is reached on page 0. The old
    // proof-scoped deferral rejected that page context and cycled; the
    // fallback then kept the image registered on page 0 (text from 272 pt).
    const result = layout(documentModel([
      ...lines(25),
      floatingTable({ widthPt: 300, heightPt: 120, tblpY: 0, vertAnchor: 'page', distancePt: 9 }),
      anchoredImageLine({
        widthPt: 468, heightPt: 200, yPt: 0, relativeFrom: 'margin', wrap: 'topAndBottom',
      }),
      ...lines(2),
    ]));
    expect(result.fallbacks).toBe(0);
    expect(result.layout.pages).toHaveLength(2);
    const first = result.layout.pages[0]!.layers.body[0]!;
    expect(first.kind === 'paragraph' && first.lines[0]!.bounds.yPt).toBe(BODY_TOP);
    expect(result.layout.pages[1]!.layers.body[0]!.source.path[0]).toBe(26);
    expect(result.passes).toBeLessThanOrEqual(1 + 8 * 2);
    expectWellFormed(result.layout);
  });

  it('rejects again when a keepNext predecessor follows the deferred anchor line', () => {
    // Review repro: the 624 pt image and its anchor line fill a page; the
    // keepNext predecessor moves with the deferred line, so on the next page
    // the line does not lead and the image cannot stay with it. That is a
    // second rejection, not an exemption that would overlap the predecessor.
    const result = layout(documentModel([
      ...lines(25),
      { ...line(), keepNext: true } as BodyElement,
      anchoredImageLine({
        widthPt: 468, heightPt: 624, yPt: 0, relativeFrom: 'margin', wrap: 'topAndBottom',
      }),
      line(),
    ]));
    expect(result.fallbacks).toBe(0);
    const pageOf = (index: number) => result.layout.pages.find((page) => (
      page.layers.body.some((node) => node.source.path[0] === index)
    ))!.pageIndex;
    expect([pageOf(25), pageOf(26)]).toEqual([1, 2]);
    expect(result.passes).toBeLessThanOrEqual(1 + 4 * (2 * 1 + 1));
    expectWellFormed(result.layout);
  });

  it('registers a page-anchored table that moves to a later column', () => {
    // Review repro: the table's trial moves its source from column 0 to
    // column 1. Its destination is registered where the section region opens
    // (a column transition is never prescanned), so the trial is read and
    // settles instead of reselecting forever.
    const result = layout(documentModel([
      anchoredImageLine({
        widthPt: 468, heightPt: 120, yPt: 72, relativeFrom: 'page', wrap: 'square',
      }),
      line(),
      floatingTable({
        widthPt: 300, heightPt: [120, 48, 48, 48, 240], tblpY: 72, vertAnchor: 'page', distancePt: 0,
      }),
    ], { pageWidth: 612, pageHeight: 792, columns: { widthsPt: [225, 225], spacePt: 18 } }));
    expect(result.fallbacks).toBe(0);
    expect(result.passes).toBeLessThanOrEqual(1 + 4 * 2 * 2);
    expectWellFormed(result.layout);
  });

  it('admits coupled page-anchored tables in source order', () => {
    // Review repro: T1/T2/T3 destination plans alternated (0, 1, 2) and
    // (absent, 4, 5). T1's own exclusion pushes its source to page 1 (Word
    // source-boundary control C06); T2 and T3 then follow on fresh pages.
    const table = (widthPt: number, heightPt: number) => floatingTable({
      widthPt, heightPt, tblpY: 72, vertAnchor: 'page', distancePt: 0,
    });
    const result = layout(documentModel([
      ...lines(13), table(468, 360), ...lines(1), table(468, 600),
      ...lines(17), table(300, 600), ...lines(2),
    ]));
    expect(result.fallbacks).toBe(0);
    const placed = result.layout.pages.flatMap((page) => page.layers.body
      .filter((node) => node.kind === 'table').map(() => page.pageIndex));
    expect(placed).toEqual([1, 2, 3]);
    expect(result.passes).toBeLessThanOrEqual(1 + 8 * 3);
    expectWellFormed(result.layout);
  });

  it('settles random table, drawing, column, section and keep mixes within the bound', () => {
    // Property (see randomDocument): every sample settles within
    // 1 + 4(2K + C) passes, never reaches the defensive fallback, is well
    // formed and is deterministic.
    let seed = 0x1659;
    const pick = <T>(values: readonly T[]): T => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return values[(seed >>> 8) % values.length]!;
    };
    for (let sample = 0; sample < 500; sample += 1) {
      const { model, anchors, kept } = randomDocument(pick);
      const first = layout(model);
      expect(first.fallbacks, `sample ${sample}`).toBe(0);
      expect(first.passes, `sample ${sample}`).toBeLessThanOrEqual(1 + 4 * (2 * anchors + kept));
      expectWellFormed(first.layout);
      if (sample % 10 === 0) {
        expect(layoutFingerprint(layout(model).layout)).toBe(layoutFingerprint(first.layout));
      }
    }
  }, 600_000);

  it('resolves a stable plan with work linear in its page-owned anchors', () => {
    // Each read of the plan, the observation and the pass's reads is counted;
    // a stable plan of T tables on T pages must cost O(T), not O(T^2).
    const work = (count: number) => {
      let reads = 0;
      class CountingMap<K, V> extends Map<K, V> {
        override get(key: K) { reads += 1; return super.get(key); }
        override *entries(): MapIterator<[K, V]> {
          for (const entry of super.entries()) { reads += 1; yield entry; }
        }
        override [Symbol.iterator]() { return this.entries(); }
        override *keys(): MapIterator<K> { for (const [key] of this.entries()) yield key; }
        override *values(): MapIterator<V> { for (const [, value] of this.entries()) yield value; }
      }
      const plan = new CountingMap<string, Parameters<typeof resolvePageOwnedAnchors>[2] extends
        ReadonlyMap<string, infer V> ? V : never>();
      const events: PageAnchorInputEvent[] = [];
      for (let pageIndex = 0; pageIndex < count; pageIndex += 1) {
        const tableSource = { story: 'body', storyInstance: 'body', path: [pageIndex * 2 + 1] } as const;
        const flowDomainId = `page:${pageIndex}`;
        const bounds = Object.freeze({ xPt: 72, yPt: 72, widthPt: 468, heightPt: 120 });
        const occurrenceId = `table-occurrence:${pageIndex}`;
        plan.set(`table:${sourceKey(tableSource)}`, Object.freeze({
          kind: 'floating-table', occurrenceId, tableSource, bounds, pageIndex, flowDomainId,
        }));
        events.push(Object.freeze({
          kind: 'prescan', pageIndex, flowDomainId,
          anchors: Object.freeze([Object.freeze({
            kind: 'floating-table' as const, occurrenceId, tableSource, bounds,
          })]),
        }));
      }
      const counted = new Proxy(events, {
        get(target, property, receiver) {
          if (typeof property === 'string' && /^\d+$/.test(property)) reads += 1;
          return Reflect.get(target, property, receiver);
        },
      });
      const order = new Map([...plan.keys()].map((key, index) => [key, index] as const));
      reads = 0;
      const result = resolvePageOwnedAnchors(
        order,
        { anchorInputs: counted, layout: { pages: [], diagnostics: [] } } as unknown as BodyPaginationPassResult,
        plan, plan, null,
        { floors: new Map(), exempt: new Set(), retried: new Set() },
      );
      expect(result.frontier).toBeNull();
      return reads;
    };
    const small = work(1_000);
    const large = work(4_000);
    expect(large).toBeLessThanOrEqual(small * 4 * 1.05);
    expect(large).toBeLessThanOrEqual(8 * 4_000);
  });

  it('renders a complete, well-formed fallback when the pass guard is forced', () => {
    // The fallback is defensive: force it by lowering the guard to two
    // passes. It keeps only reads before the changing page, drops later
    // registrations and gives each later anchor a fresh page.
    let seed = 0xfa11;
    const pick = <T>(values: readonly T[]): T => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return values[(seed >>> 8) % values.length]!;
    };
    counters.limit = 2;
    try {
      let forced = 0;
      const models = [
        documentModel([
          ...lines(25),
          floatingTable({ widthPt: 300, heightPt: 120, tblpY: 0, vertAnchor: 'page', distancePt: 9 }),
          anchoredImageLine({
            widthPt: 468, heightPt: 200, yPt: 0, relativeFrom: 'margin', wrap: 'topAndBottom',
          }),
          ...lines(2),
        ]),
        ...Array.from({ length: 60 }, () => randomDocument(pick).model),
      ];
      for (const model of models) {
        const result = layout(model);
        forced += result.fallbacks > 0 ? 1 : 0;
        const sources = new Set(result.layout.pages.flatMap((page) => page.layers.body)
          .map((node) => node.source.path[0]));
        expect(sources.size).toBe(model.body.filter((element) => (
          element.type === 'paragraph' || element.type === 'table'
        )).length);
        expectWellFormed(result.layout);
      }
      expect(forced).toBeGreaterThan(10);
    } finally {
      counters.limit = Number.POSITIVE_INFINITY;
    }
  }, 600_000);
});
