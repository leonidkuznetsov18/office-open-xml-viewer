import { describe, expect, it } from 'vitest';
import { layoutLines, type LayoutTextSeg, type WrapLayoutCtx } from '../line-layout.js';
import type { FloatRect } from './float-wrap.js';

function context(): CanvasRenderingContext2D {
  return { measureText: (text: string) => ({ width: text.length * 5,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
}
function token(text: string): LayoutTextSeg {
  return { text, fontSize: 10, fontFamily: 'Arial', bold: false, italic: false,
    underline: false, strikethrough: false, color: null, vertAlign: null, measuredWidth: 0 };
}
function obstacle(left: number, right: number, authoredWrap: 'square' | 'tight' | 'through'): FloatRect {
  return { kind: 'shape', mode: 'square', authoredWrap, imageKey: 'test',
    imageX: left, imageY: 0, imageW: right - left, imageH: 60,
    xLeft: left, xRight: right, yTop: 0, yBottom: 60, side: 'bothSides',
    distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0,
    ...(authoredWrap === 'square' ? {} : { wrapPolygon: [
      { xPt: left, yPt: 0 }, { xPt: right, yPt: 0 },
      { xPt: right, yPt: 60 }, { xPt: left, yPt: 60 },
    ] }) };
}
function wrap(floats: FloatRect[], rtl = false): WrapLayoutCtx {
  return { floats, paraX: 0, startPageY: 0, columnXPt: 0, columnWidthPt: 200,
    readingDirection: rtl ? 'rtl' : 'ltr', pageH: 800, lineBoxH: () => 10 };
}
describe('Word measured gap continuation (#1670)', () => {
  it.each(['square', 'tight', 'through'] as const)('fills three %s gaps on one baseline before advancing', (mode) => {
    const lines = layoutLines(context(), ['AAAA ', 'BBBB ', 'CCCC ', 'DDDD'].map(token),
      200, 0, 1, [], wrap([obstacle(40, 70, mode), obstacle(110, 160, mode)]));
    expect(lines.map(l => [l.topY, l.xOffset, l.availWidth])).toEqual([
      [0, 0, 40], [0, 70, 40], [0, 160, 40], [10, 0, 40],
    ]);
    expect(lines.flatMap(l => l.segments.map(s => 'text' in s ? s.text : '')).join('')).toBe('AAAA BBBB CCCC DDDD');
  });
  it('uses right-to-left gap order and restarts at the right on the next baseline', () => {
    const lines = layoutLines(context(), ['AAAA ', 'BBBB ', 'CCCC'].map(token),
      200, 0, 1, [], wrap([obstacle(40, 160, 'square')], true), {}, 0,
      undefined, undefined, undefined, undefined, true);
    expect(lines.map(l => [l.topY, l.xOffset])).toEqual([[0, 160], [0, 0], [10, 160]]);
  });
  it('skips gaps that cannot hold the next atom and keeps an unbroken word whole below the float', () => {
    const lines = layoutLines(context(), [token('UNBROKENWORDUNBROKENWORD')], 200, 0, 1, [],
      wrap([obstacle(100, 160, 'square')]));
    expect(lines).toHaveLength(1);
    expect(lines[0].topY).toBe(60);
    expect(lines[0].segments[0]).toMatchObject({ text: 'UNBROKENWORDUNBROKENWORD' });
  });
  it('admits a word across formatting seams as one atom', () => {
    const lines = layoutLines(context(), [token('UNBROKEN'), { ...token('WORD'), joinPrev: true }],
      200, 0, 1, [], wrap([obstacle(40, 160, 'square')]));
    expect(lines.map(line => line.topY)).toEqual([60]);
    expect(lines[0].segments.map(segment => 'text' in segment ? segment.text : '').join('')).toBe('UNBROKENWORD');
  });
  it('shares the tallest gap metrics and advances one whole physical line', () => {
    const ctx = context();
    ctx.measureText = (text: string) => ({ width: text.length * 5,
      fontBoundingBoxAscent: text.startsWith('AAAA') ? 16 : 8,
      fontBoundingBoxDescent: text.startsWith('AAAA') ? 4 : 2 }) as TextMetrics;
    const w = wrap([obstacle(40, 160, 'square')]);
    w.lineBoxH = (ascent, descent) => ascent + descent;
    const lines = layoutLines(ctx, ['AAAA ', 'BBBB ', 'CCCC'].map(token), 200, 0, 1, [], w);
    expect(lines.map(l => [l.topY, l.ascent, l.descent])).toEqual([[0, 16, 4], [0, 16, 4], [20, 8, 2]]);
  });
  it('manual breaks advance vertically instead of continuing in the next gap', () => {
    const lines = layoutLines(context(), [token('AAAA'), { lineBreak: true, fontSize: 10, measuredWidth: 0 }, token('BBBB')],
      200, 0, 1, [], wrap([obstacle(40, 160, 'square')]));
    expect(lines.map(l => [l.topY, l.xOffset])).toEqual([[0, 0], [10, 0]]);
  });
});
