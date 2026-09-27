import { expect, it } from 'vitest';
import { layoutLines, type LayoutTextSeg } from './line-layout.js';
import { frenchHyphenationOffsets } from './hyphenation.js';

const context = {
  font: '10px serif', letterSpacing: '0px',
  measureText(text: string) {
    return {
      width: text.length * 5,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    } as TextMetrics;
  },
} as unknown as CanvasRenderingContext2D;

function segment(text: string, index: number, language?: string, zone?: number): LayoutTextSeg {
  return {
    text, src: { segIndex: index, charOffset: 0 },
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'serif', vertAlign: null,
    measuredWidth: 0, hyphenationLanguage: language, hyphenationZonePt: zone,
  };
}

it('uses a French dictionary boundary only when the authored zone permits it', () => {
  expect(frenchHyphenationOffsets('organisation', 'fr-CA')).toContain(4);
  expect(frenchHyphenationOffsets('organisation', 'en-US')).toEqual([]);
  const enabled = layoutLines(context, [segment('abcdefg ', 0), segment('organisation', 1, 'fr-CA', 10)], 65, 0, 1);
  expect(enabled[0]!.segments.map((part) => 'text' in part ? part.text : '').join('')).toBe('abcdefg orga-');
  expect(enabled[1]!.segments.map((part) => 'text' in part ? part.text : '').join('')).toBe('nisation');
  const blocked = layoutLines(context, [segment('abcdefg ', 0), segment('organisation', 1, 'fr-CA', 30)], 65, 0, 1);
  expect(blocked[0]!.segments.map((part) => 'text' in part ? part.text : '').join('')).toBe('abcdefg ');
  expect(blocked[1]!.segments.map((part) => 'text' in part ? part.text : '').join('')).toBe('organisation');
});

it('hyphenates a word at a source-run seam without changing the suffix text', () => {
  const parts = [
    segment('abcdefg ', 0), segment('orga', 1, 'fr-CA', 10),
    segment('nisation', 2, 'fr-CA', 10),
  ];
  const lines = layoutLines(context, parts, 65, 0, 1);
  expect(lines[0]!.segments.map((part) => 'text' in part ? part.text : '').join('')).toBe('abcdefg orga-');
  expect(lines[1]!.segments.map((part) => 'text' in part ? part.text : '').join('')).toBe('nisation');
});

it('keeps a Word run-split word in the current line when its dictionary seam fits', () => {
  const parts = [
    segment('abcdefg ', 0), segment('télév', 1, 'fr-CA', 10),
    { ...segment('i', 2, 'fr-CA', 10), joinPrev: true },
    { ...segment('sion, ', 3, 'fr-CA', 10), joinPrev: true },
  ];
  const lines = layoutLines(context, parts, 80, 0, 1);
  expect(lines[0]!.segments.map((part) => 'text' in part ? part.text : '').join('')).toBe('abcdefg télévi-');
  expect(lines[1]!.segments.map((part) => 'text' in part ? part.text : '').join('')).toBe('sion, ');
});
