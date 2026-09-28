import { beforeAll, describe, expect, it, vi } from 'vitest';

// Counts body pagination passes: each pass opens exactly one kernel session.
const passes = vi.hoisted(() => ({ opened: 0 }));
vi.mock('./runtime-state.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./runtime-state.js')>();
  return {
    ...actual,
    bodyLayoutKernelOf: (services: Parameters<typeof actual.bodyLayoutKernelOf>[0]) => {
      const kernel = actual.bodyLayoutKernelOf(services);
      return kernel && {
        openBodyLayoutSession: (...args: Parameters<typeof kernel.openBodyLayoutSession>) => {
          passes.opened += 1;
          return kernel.openBodyLayoutSession(...args);
        },
      };
    },
  };
});

import { createLayoutServices } from '../layout-runtime.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import { installStubCanvas, syntheticDocxModel } from '../testing/synthetic-document.js';
import type { DocxDocumentModel } from '../types.js';
import { paginateBody } from './body-paginator.js';
import { layoutFingerprint } from './invariants.js';
import { normalizeLayoutOptions } from './options.js';
import { PaginationAbortError } from './pagination-scheduler.js';
import { layoutDocumentProgressively, type ProgressiveLayoutPreview } from './progressive.js';
import type { DocumentLayout, LayoutPage } from './types.js';

// Page-owned anchors (§20.4.3.5 positionV relativeFrom="margin") are
// resolved by exact-state convergence: the first pass prescans them in source
// order, later passes apply the destinations the previous pass observed. A
// page published from a pass that a later pass supersedes would be revised.
// These cases pin that anchored documents still publish progressively and that
// every published page is already the page of the final layout.

beforeAll(() => {
  installStubCanvas();
});

/** A plain document with a margin-relative square-wrapped image in each listed
 * paragraph. The first pass wraps text around every image from the top of the
 * document, so later passes move pages. */
function anchoredModel(paragraphs: number, anchored: readonly number[]): DocxDocumentModel {
  const model = syntheticDocxModel('plain', { paragraphs });
  for (const index of anchored) {
    const paragraph = model.body[index] as unknown as { runs: unknown[] };
    paragraph.runs = [...paragraph.runs, {
      type: 'image', imagePath: 'word/media/anchor.png', mimeType: 'image/png',
      widthPt: 150, heightPt: 120, anchor: true, anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: true, anchorYFromPara: false,
      wrapMode: 'square', wrapSide: 'bothSides',
      anchorXRelativeFrom: 'margin', anchorYRelativeFrom: 'margin',
    }];
  }
  return model;
}

function open(model: DocxDocumentModel) {
  const source = layoutSourceStore(model);
  return {
    input: source.bodyLayoutInput,
    services: createLayoutServices(source),
    options: normalizeLayoutOptions(undefined, 1_700_000_000_000),
  };
}

function pageFingerprint(page: LayoutPage): string {
  return layoutFingerprint({ pages: [page], diagnostics: [] } as DocumentLayout);
}

function blockingLayout(model: DocxDocumentModel): Readonly<{ layout: DocumentLayout; passes: number }> {
  const before = passes.opened;
  const blocking = open(model);
  const layout = paginateBody(blocking.input, blocking.services, blocking.options);
  return { layout, passes: passes.opened - before };
}

function expectPublishedPagesFinal(
  previews: readonly ProgressiveLayoutPreview[],
  final: DocumentLayout,
): void {
  const counts = previews.map((preview) => preview.layout.pages.length);
  expect(counts).toEqual([...counts].sort((left, right) => left - right));
  for (const preview of previews) {
    expect(preview.exact).toBe(false);
    expect(preview.layout.pages.length).toBeLessThanOrEqual(final.pages.length);
    preview.layout.pages.forEach((page, index) => {
      expect(pageFingerprint(page as LayoutPage))
        .toBe(pageFingerprint(final.pages[index] as LayoutPage));
    });
  }
}

describe('progressive layout with page-owned anchors', () => {
  it('publishes only pages every later convergence pass reproduces', async () => {
    const model = anchoredModel(300, [40, 120]);
    const blocking = blockingLayout(model);
    // Discovery, correction and confirmation: more than one pass must run.
    expect(blocking.passes).toBeGreaterThanOrEqual(3);

    const previews: ProgressiveLayoutPreview[] = [];
    const progressive = open(model);
    const final = await layoutDocumentProgressively(
      progressive.input,
      progressive.services,
      progressive.options,
      { onPreview: (preview) => { previews.push(preview); } },
    );

    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    // The anchored document now publishes before layout completes.
    expect(previews.length).toBeGreaterThan(1);
    expect(previews[0]!.layout.pages.length).toBeLessThan(final.pages.length);
    expectPublishedPagesFinal(previews, final);
  }, 300_000);

  it('emits no stale page when cancelled during convergence', async () => {
    const model = anchoredModel(300, [40, 120]);
    const { layout: final } = blockingLayout(model);
    const everyStep = (onYield: () => void) => ({
      now: () => Number.MAX_SAFE_INTEGER,
      sliceMs: 0,
      yieldToHost: () => { onYield(); return Promise.resolve(); },
    });

    let totalYields = 0;
    const complete = open(model);
    await layoutDocumentProgressively(complete.input, complete.services, complete.options, {
      onPreview: () => {},
      scheduler: everyStep(() => { totalYields += 1; }),
    });

    // Abort points inside the second and later passes, after publications
    // have started, as a destroyed viewer or a replaced load would.
    for (const fraction of [0.4, 0.7, 0.9]) {
      const abortAt = Math.floor(totalYields * fraction);
      const controller = new AbortController();
      const previews: ProgressiveLayoutPreview[] = [];
      let yields = 0;
      let publishedAfterAbort = false;
      const cancelled = open(model);
      await expect(layoutDocumentProgressively(
        cancelled.input,
        cancelled.services,
        cancelled.options,
        {
          onPreview: (preview) => {
            if (controller.signal.aborted) publishedAfterAbort = true;
            previews.push(preview);
          },
          scheduler: {
            ...everyStep(() => {
              yields += 1;
              if (yields === abortAt) controller.abort();
            }),
            signal: controller.signal,
          },
        },
      )).rejects.toBeInstanceOf(PaginationAbortError);
      expect(previews.length).toBeGreaterThan(0);
      expect(publishedAfterAbort).toBe(false);
      expectPublishedPagesFinal(previews, final);
    }
  }, 300_000);
});
