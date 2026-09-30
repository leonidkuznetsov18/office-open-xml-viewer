import { afterEach, describe, expect, it, vi } from 'vitest';
import type { HyperlinkTarget } from '@silurus/ooxml-core';
import { layoutParagraph, renderTextBody, shapeTextRotation, type PptxTextRunInfo } from './renderer.js';
import { buildPptxTextLayer } from './text-layer.js';
import type { Paragraph, TextBody, TextRunData } from './types.js';

// ECMA-376 §20.1.10.83 wordArtVert / wordArtVertRtl through the PowerPoint
// renderer: segments become stacked glyphs (core layoutStackedText) and each
// glyph class paints as the issue #1626 controls show.
const SCALE = 1 / 12700; // 1 pt → 1 px

interface Call { text: string; x: number; y: number; rot: number; tx: number; ty: number; align: string }

function mockCtx(): { ctx: CanvasRenderingContext2D; calls: Call[] } {
  let font = '24px serif';
  let textAlign: CanvasTextAlign = 'left';
  let textBaseline: CanvasTextBaseline = 'alphabetic';
  let rot = 0;
  let tx = 0;
  let ty = 0;
  const stack: [number, number, number][] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '24');
  const calls: Call[] = [];
  const ctx = {
    canvas: { style: {} },
    get font() { return font; }, set font(v: string) { font = v; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px', direction: 'ltr',
    get textAlign() { return textAlign; }, set textAlign(v: CanvasTextAlign) { textAlign = v; },
    get textBaseline() { return textBaseline; }, set textBaseline(v: CanvasTextBaseline) { textBaseline = v; },
    measureText: (s: string) => ({
      width: [...s].length * px() * 0.6,
      fontBoundingBoxAscent: px() * 0.9, fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.7, actualBoundingBoxDescent: 0,
    }) as TextMetrics,
    fillText: (text: string, x: number, y: number) => calls.push({ text, x, y, rot, tx, ty, align: textAlign }),
    save: () => { stack.push([rot, tx, ty]); },
    restore: () => { const s = stack.pop(); if (s) [rot, tx, ty] = s; },
    translate: (x: number, y: number) => { tx += x; ty += y; },
    rotate: (a: number) => { rot += a; },
    scale: () => {}, beginPath: () => {}, moveTo: () => {}, lineTo: () => {}, stroke: () => {},
    fill: () => {}, clip: () => {}, rect: () => {}, fillRect: () => {}, setLineDash: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

function body(text: string, vert: string, fontFamilyEa?: string, hyperlink?: string): TextBody {
  const run = {
    type: 'text', text, bold: null, italic: null, underline: false, strikethrough: false,
    fontSize: 24, color: '000000', fontFamily: 'Arial', fontFamilyEa, hyperlink,
  } as TextRunData;
  const para = {
    alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null, spaceLine: null,
    lvl: 0, bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: [], eaLnBrk: true, runs: [run],
  } as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: 24, defaultBold: null, defaultItalic: null,
    lIns: 91440, rIns: 91440, tIns: 45720, bIns: 45720, wrap: 'square', vert, autoFit: 'none',
  };
}

interface FakeEl {
  textContent: string; innerHTML: string; title: string; dataset: Record<string, string>;
  style: Record<string, string>; children: FakeEl[]; listeners: ((e: { preventDefault(): void }) => void)[];
  appendChild(c: FakeEl): void; setAttribute(): void; addEventListener(t: string, f: (e: { preventDefault(): void }) => void): void; click(): void;
}
function fakeEl(): FakeEl {
  const el: FakeEl = {
    textContent: '', innerHTML: '', title: '', dataset: {}, style: {}, children: [], listeners: [],
    appendChild(c) { el.children.push(c); },
    setAttribute() {},
    addEventListener(t, f) { if (t === 'click') el.listeners.push(f); },
    click() { for (const f of el.listeners) f({ preventDefault() {} }); },
  };
  return el;
}
afterEach(() => vi.unstubAllGlobals());

// Arial 24 pt: cell 7/6 × 1.1172 em × 24 = 31.28, descent 0.2119 em.
const CELL = (7 / 6) * (2288 / 2048) * 24;

describe('pptx stacked vertical text (wordArtVert / wordArtVertRtl)', () => {
  it('stacks upright Latin glyphs on the column axis, one cell apart', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body('AB', 'wordArtVert'), 0, 0, 60, 470, SCALE);
    const [a, b] = calls;
    expect(a).toMatchObject({ text: 'A', align: 'center', rot: 0 });
    expect(a.x).toBeCloseTo(7.2 + CELL / 2, 6);
    expect(a.y).toBeCloseTo(3.6 + CELL - (434 / 2048) * 24, 6);
    expect(b.y - a.y).toBeCloseTo(CELL, 6);
    // wordArtVertRtl starts at the right edge.
    const r = mockCtx();
    renderTextBody(r.ctx, body('AB', 'wordArtVertRtl'), 0, 0, 60, 470, SCALE);
    expect(r.calls[0].x).toBeCloseTo(60 - 7.2 - CELL / 2, 6);
  });

  it('turns ASCII brackets 90° and draws CJK brackets with the face vertical glyph', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body('(A「', 'wordArtVert', 'Yu Gothic'), 0, 0, 60, 470, SCALE);
    const paren = calls.find((c) => c.text === '(');
    expect(paren?.rot).toBeCloseTo(Math.PI / 2, 6);
    // 「 is drawn upright as its vertical presentation form ﹁ (U+FE41).
    const bracket = calls.find((c) => c.text === '﹁');
    expect(bracket?.rot).toBe(0);
    expect(calls.some((c) => c.text === '「')).toBe(false);
  });

  it('stacks a grapheme cluster in one cell', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body('e\u0301B', 'wordArtVert'), 0, 0, 60, 470, SCALE);
    expect(calls.map((c) => c.text)).toEqual(['e\u0301', 'B']);
    expect(calls[1].y - calls[0].y).toBeCloseTo(CELL, 6);
  });

  it('hands each glyph its run hyperlink so the text layer installs link handlers', () => {
    const { ctx } = mockCtx();
    const runs: PptxTextRunInfo[] = [];
    renderTextBody(ctx, body('AB', 'wordArtVert', undefined, 'https://example.com/'), 0, 0, 60, 470, SCALE,
      null, 0, false, false, '#000000', undefined, undefined, (r) => runs.push(r));
    const target: HyperlinkTarget = { kind: 'external', url: 'https://example.com/' };
    expect(runs.map((r) => r.hyperlink)).toEqual([target, target]);
    vi.stubGlobal('document', { createElement: () => fakeEl() });
    const layer = fakeEl();
    const onClick = vi.fn<(t: HyperlinkTarget) => void>();
    buildPptxTextLayer(layer as unknown as HTMLDivElement, runs, 960, 540, onClick);
    const spans = layer.children.flatMap((group) => group.children);
    expect(spans).toHaveLength(2);
    spans[0].click();
    expect(onClick).toHaveBeenCalledWith(target);
  });

  it('keeps a stacked body upright under flipH and turns it 180° under flipV', () => {
    expect(shapeTextRotation('wordArtVert', 0, true, false)).toBe(0);
    expect(shapeTextRotation('wordArtVertRtl', 0, false, true)).toBe(180);
    expect(shapeTextRotation('horz', 0, false, true)).toBe(0);
  });
});

// A grapheme extender (variation selector, combining mark) follows its base's
// font slot, so no cluster straddles two segments. Arial latin, Yu Gothic ea.
describe('font slots keep grapheme clusters whole', () => {
  const IVS = '\u{E0100}'; // ideographic variation selector (Latin slot by itself)
  it('stacked: a CJK base and its selector share one Yu Gothic cell', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body(`葛${IVS}B`, 'wordArtVert', 'Yu Gothic'), 0, 0, 60, 470, SCALE);
    expect(calls.map((c) => c.text)).toEqual([`葛${IVS}`, 'B']);
  });

  it('stacked: a bracket with a selector keeps its glyphs (no lone-character substitution)', () => {
    const { ctx, calls } = mockCtx();
    renderTextBody(ctx, body('「\uFE0F', 'wordArtVert', 'Yu Gothic'), 0, 0, 60, 470, SCALE);
    expect(calls.map((c) => c.text)).toEqual(['「\uFE0F']);
  });

  it('horizontal: the cluster stays in the base segment and font', () => {
    const { ctx } = mockCtx();
    const para = body(`葛${IVS}B e\u0301`, 'horz', 'Yu Gothic').paragraphs[0];
    const [line] = layoutParagraph(ctx, para, 10_000, 24, '#000', SCALE, 0);
    const segs = line.segments.filter((g) => g.text);
    expect(segs.map((g) => g.text)).toEqual([`葛${IVS}`, 'B e\u0301']);
    expect(segs[0].font).toContain('Yu Gothic');
    expect(segs[1].font).not.toContain('Yu Gothic');
  });
});

// A cluster whose extender opens the next authored run: clusters are
// segmented over the paragraph text, and the carried extender takes the base
// run's formatting (the second run here is bold).
describe('grapheme clusters across run seams', () => {
  const SEAMS: [string, string][] = [['葛', '\u{E0100}'], ['「', '\uFE0F'], ['가', '\u11A8']];
  const split = (vert: string, base: string, ext: string): TextBody => {
    const b = body(`${base}${ext}B`, vert, 'Yu Gothic');
    const r = b.paragraphs[0].runs[0] as TextRunData;
    b.paragraphs[0].runs = [{ ...r, text: base }, { ...r, text: `${ext}B`, bold: true }];
    return b;
  };
  for (const [base, ext] of SEAMS) {
    it(`stacked: ${JSON.stringify(base + ext)} split across runs stays one cell`, () => {
      for (const vert of ['wordArtVert', 'wordArtVertRtl']) {
        const { ctx, calls } = mockCtx();
        renderTextBody(ctx, split(vert, base, ext), 0, 0, 100, 470, SCALE);
        expect(calls.map((c) => c.text)).toEqual([base + ext, 'B']);
      }
    });
    it(`horizontal: ${JSON.stringify(base + ext)} split across runs stays in the base segment`, () => {
      const { ctx } = mockCtx();
      const [line] = layoutParagraph(ctx, split('horz', base, ext).paragraphs[0], 10_000, 24, '#000', SCALE, 0);
      const segs = line.segments.filter((g) => g.text);
      expect(segs.map((g) => g.text)).toEqual([base + ext, 'B']);
      expect(segs[0].font).toContain('Yu Gothic');
      expect(segs[0].font).not.toContain('bold');
      expect(segs[1].font).toContain('bold');
    });
  }
});
