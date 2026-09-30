import { describe, expect, it } from 'vitest';
import { renderTextBody, shapeTextRotation } from './renderer.js';
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

function body(text: string, vert: string, fontFamilyEa?: string): TextBody {
  const run = {
    type: 'text', text, bold: null, italic: null, underline: false, strikethrough: false,
    fontSize: 24, color: '000000', fontFamily: 'Arial', fontFamilyEa,
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

  it('keeps a stacked body upright under flipH and turns it 180° under flipV', () => {
    expect(shapeTextRotation('wordArtVert', 0, true, false)).toBe(0);
    expect(shapeTextRotation('wordArtVertRtl', 0, false, true)).toBe(180);
    expect(shapeTextRotation('horz', 0, false, true)).toBe(0);
  });
});
