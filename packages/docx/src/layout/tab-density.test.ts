/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { documentBytes, measureContext } from '../test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { runLineBreakerPass } from '../line-breaker/pass-driver.js';
import type { LayoutSeg } from '../line-breaker/model.js';
import { layoutBidiTabStops } from '../line-breaker/tabs.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function tabDocument(count: number, rtl: boolean, float: boolean) {
  const archive = new DocxArchive(documentBytes('', rtl, 0,
    { alignment: 'left', count, text: 'word', noFloat: !float }));
  try {
    return normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
}

function timedLayout(model: ReturnType<typeof tabDocument>) {
  const services = createLayoutServices(model, { measureContext: measureContext() });
  const start = process.cpuUsage();
  const layout = layoutDocument(model, services, { currentDateMs: 0 });
  const ms = elapsedCpuMs(start);
  const text = layout.pages.flatMap((page) => page.layers.body)
    .filter((node) => node.kind === 'paragraph')
    .flatMap((node) => node.lines.flatMap((line) => line.placements))
    .filter((node) => node.kind === 'text').map((node) => node.text).join('');
  expect(text).toBe('word');
  return ms;
}

// Empty cells beyond the margin can collapse onto one RTL line. Repeatedly
// projecting its completed prefix used to take tens of seconds for 32k tabs.
// Compare with the production linear tab walk, allowing ample layout overhead
// and an absolute noise floor for slower/shared CI machines. Per-process CPU
// time excludes scheduling pauses caused by unrelated parallel test workers.
it.each([false, true].flatMap((rtl) => [false, true].map((float) => ({ rtl, float }))))(
  'lays out 32,000 tabs within a linear budget: rtl=$rtl float=$float', ({ rtl, float }) => {
    const count = 32_000;
    const run = () => timedPass(count, rtl, float);
    const items = Array.from({ length: count }, () => ({ isTab: true, width: 0 }));
    const start = process.cpuUsage();
    layoutBidiTabStops(items, [], 0, 468, 36);
    const linearMs = elapsedCpuMs(start);
    expect(run()).toBeLessThan(Math.max(2_000, linearMs * 100));
  }, 120_000,
);

it.each([false, true].flatMap((rtl) => [false, true].map((float) => ({ rtl, float }))))(
  'scales tab fitting from n to 4n: rtl=$rtl float=$float', ({ rtl, float }) => {
    timedPass(2_000, rtl, float);
    timedPass(8_000, rtl, float);
    const sample = (count: number) =>
      Math.min(...Array.from({ length: 3 }, () => timedPass(count, rtl, float)));
    const smallMs = sample(2_000);
    const largeMs = sample(8_000);
    expect(largeMs).toBeLessThan(smallMs * 10 + 100);
  }, 120_000,
);

it('lays out a parser-backed 32,000-tab RTL document within the fitting budget', () => {
  expect(timedLayout(tabDocument(32_000, true, false))).toBeLessThan(2_000);
}, 120_000);

// Time the queue-owning production pass separately from document pagination:
// LTR unreachable tabs intentionally produce many lines, and page-level float
// retries have different costs from resolving a single line's completed cells.
function timedPass(count: number, rtl: boolean, float: boolean) {
  const segs: LayoutSeg[] = Array.from({ length: count }, () =>
    ({ isTab: true, fontSize: 12, measuredWidth: 0 }));
  segs.push({ text: 'word', bold: false, italic: false, underline: false,
    strikethrough: false, fontSize: 12, color: null, fontFamily: null,
    vertAlign: null, measuredWidth: 0 });
  const start = process.cpuUsage();
  const lines = runLineBreakerPass({
    ctx: measureContext(), segs, maxWidth: 468, firstIndent: 0, scale: 1,
    tabStops: [], fontFamilyClasses: {}, tabOriginPx: 0,
    kinsoku: DEFAULT_KINSOKU_RULES, defaultTabPt: 36, marginRightPx: 468,
    baseRtl: rtl, isJustified: false, stretchLastLine: false,
    widthPolicy: 'bounded', overflowPunct: false,
    passContext: { probeHeights: float ? Array(count + 1).fill(12) : null },
    wrapCtx: float ? {
      startPageY: 0, paraX: 0, columnXPt: 0, columnWidthPt: 468,
      floats: [{ kind: 'shape', mode: 'square', imageKey: '',
        imageX: 0, imageY: 0, imageW: 200, imageH: 100,
        xLeft: 0, xRight: 200, yTop: 0, yBottom: 100, side: 'bothSides',
        distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 }],
      lineBoxH: () => 12, pageH: Number.POSITIVE_INFINITY,
    } : undefined,
  });
  const ms = elapsedCpuMs(start);
  expect(lines.flatMap((line) => line.segments).filter((seg) => 'isTab' in seg)).toHaveLength(count);
  expect(lines.flatMap((line) => line.segments).filter((seg) => 'text' in seg)
    .map((seg) => seg.text).join('')).toBe('word');
  return ms;
}

function elapsedCpuMs(start: ReturnType<typeof process.cpuUsage>) {
  const elapsed = process.cpuUsage(start);
  return (elapsed.user + elapsed.system) / 1_000;
}
