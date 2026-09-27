import { describe, expect, it } from 'vitest';
import { layoutLines, type LayoutTextSeg } from './line-layout.js';

function context(): CanvasRenderingContext2D {
  let font = '10px Arial';
  return {
    get font() { return font; }, set font(value: string) { font = value; },
    letterSpacing: '0px', fontKerning: 'auto',
    measureText(text: string) {
      const symbol = /Segoe UI Symbol|Apple Color Emoji|__ooxml_registered_emoji/u.test(font);
      return {
        width: [...text].length * 5,
        fontBoundingBoxAscent: symbol ? 15 : 8,
        fontBoundingBoxDescent: symbol ? 3 : 2,
        actualBoundingBoxAscent: symbol ? 15 : 8,
        actualBoundingBoxDescent: symbol ? 3 : 2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

function segment(text: string, fontFamily: string): LayoutTextSeg {
  return {
    text, fontFamily, fontSize: 10, position: 3,
    bold: false, italic: false, underline: false, strikethrough: false,
    color: null, vertAlign: null, measuredWidth: 0,
  };
}

function metrics(parts: LayoutTextSeg[], width = 100) {
  const lines = layoutLines(context(), parts, width, 0, 1);
  return lines.map((line) => ({
    box: line.ascent + line.descent,
    advance: line.segments.reduce((sum, part) => sum + part.measuredWidth, 0),
    text: line.segments.map((part) => 'text' in part ? part.text : '').join(''),
  }));
}

describe('Word mixed leading symbol line box', () => {
  it('keeps glyph advance and paint text while ordinary following text owns height', () => {
    const ordinary = metrics([segment('Word', 'Arial')])[0]!;
    for (const symbol of [
      segment('❑', 'Segoe UI Symbol'),
      segment('◼︎', 'Apple Color Emoji'),
    ]) {
      const mixed = metrics([symbol, segment('Word', 'Arial')])[0]!;
      expect(mixed.box).toBe(ordinary.box);
      expect(mixed.advance).toBeGreaterThan(ordinary.advance);
      expect(mixed.text).toBe(`${symbol.text}Word`);
    }
  });

  it('retains native metrics when the symbol is isolated, trailing, or color-presented', () => {
    const ordinary = metrics([segment('Word', 'Arial')])[0]!;
    expect(metrics([segment('❑', 'Segoe UI Symbol')])[0]!.box).toBeGreaterThan(ordinary.box);
    expect(metrics([segment('Word', 'Arial'), segment('❑', 'Segoe UI Symbol')])[0]!.box)
      .toBeGreaterThan(ordinary.box);
    expect(metrics([segment('◼️', 'Apple Color Emoji'), segment('Word', 'Arial')])[0]!.box)
      .toBeGreaterThan(ordinary.box);
    expect(metrics([segment('❑', 'Segoe UI Symbol'), segment('Word', 'Arial')], 5)[0]!.box)
      .toBeGreaterThan(ordinary.box);
  });

  it('uses the authored face after canvas resource registration', () => {
    const symbol = segment('◼︎', '__ooxml_registered_emoji');
    symbol.authoredFontFamily = 'Apple Color Emoji';
    const ordinary = segment('Word', '__ooxml_registered_arial');
    ordinary.authoredFontFamily = 'Arial';
    expect(metrics([symbol])[0]!.box).toBeGreaterThan(metrics([ordinary])[0]!.box);
    expect(metrics([symbol, ordinary])[0]!.box).toBe(metrics([ordinary])[0]!.box);
  });
});
