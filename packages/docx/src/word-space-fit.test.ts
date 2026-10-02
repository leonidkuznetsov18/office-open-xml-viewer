import { describe, expect, it } from 'vitest';
import {
  BAND_DEFICIT_PT, FONTS, VARIANTS, advancePt, layoutStubParagraph, type Variant,
} from './test-support/word-space-fit.test-support.js';

// WORD_COMPRESSED_SPACE_LINE_FIT. Every case is a synthetic Word control
// (issue #1660: 576 coarse, 595 fine and 602 round-3 cells) exported by Word
// for Mac 16.113.3 with Save As PDF; `wordWraps` is read from Word's own glyph
// origins. The stub measures each character with the hmtx advance of the face
// Word embedded (fonts[*].advances) and exposes that face's OS/2 xAvgCharWidth,
// so no Word value is restated as a renderer constant.
//
// Table geometry is not part of this rule. The controls' half-point bordered
// fixed cells end Word's line 1pt inside the authored content width (measured
// by the no-space anchors: wrap at C=85.95, fit at 86.00 for an 85pt natural
// line) and borderless cells end at the content width; the stub line width is
// that measured band.

function sourceChunks(variant: Variant): string[] {
  if (variant.sourceRuns === 'single') return [variant.text];
  return variant.text.match(/[　-鿿＀-￯]+|[^　-鿿＀-￯]+/gu) ?? [];
}

function kerning(variant: Variant): number | undefined {
  if (variant.kern === 'on') return 0;
  if (variant.kern === 'off') return variant.sizePt + 0.5;
  return undefined;
}

function lineCount(variant: Variant, widthTwips: number) {
  return layoutStubParagraph({
    runs: sourceChunks(variant).map((text) => ({
      text, ascii: variant.ascii, eastAsia: variant.eastAsia, sizePt: variant.sizePt,
      bold: variant.bold, kerning: kerning(variant),
    })),
    environment: {
      compatibilityMode: variant.compatibilityMode,
      ...(variant.characterSpacingControl
        ? { characterSpacingControl: variant.characterSpacingControl } : {}),
      enableOpenTypeFeatures: variant.enableOpenTypeFeatures,
      lineWrapLikeWord6: false,
    },
    bandPt: widthTwips / 20 - BAND_DEFICIT_PT[variant.borderEighths]!,
    justification: variant.justification,
  });
}

const inScope = VARIANTS.filter((variant) => variant.outOfScope === null);
const BIZ = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
const MODE_14 = { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' } as const;

describe('WORD_COMPRESSED_SPACE_LINE_FIT controls', () => {
  it('covers every measured cell exactly once', () => {
    expect(VARIANTS.reduce((sum, variant) => sum + variant.widthsTwips.length, 0)).toBe(1773);
    expect(inScope.reduce((sum, variant) => sum + variant.widthsTwips.length, 0)).toBe(1443);
    // Every excluded class names a measured cause outside this rule.
    expect(new Set(VARIANTS.flatMap((variant) => variant.outOfScope
      ? [variant.outOfScope.split(':')[0]] : []))).toEqual(new Set([
      'autospace', 'kerning', 'punctuation', 'geometry', 'latin-only', 'unexplained',
    ]));
  });

  it.each(inScope.map((variant) => [
    `${variant.stage} ${variant.document} ${variant.variant}`, variant,
  ] as const))('%s matches every Word wrap outcome', (_label, variant) => {
    const actual = variant.widthsTwips
      .map((width) => (lineCount(variant, width).length > 1 ? '1' : '0'))
      .join('');
    expect(actual).toBe(variant.wordWraps);
  });

  it('fits the issue #1660 cell on one line with 3.5pt spaces', () => {
    // sample-36 A12 equivalent: 107.5pt band, OpenType features on; Word
    // paints four 3.5pt spaces.
    const [line, ...rest] = layoutStubParagraph({
      runs: [{ ...BIZ, text: '甲甲甲甲 + 乙乙 + 丙丙丙丙' }],
      environment: { ...MODE_14, enableOpenTypeFeatures: true },
      bandPt: 107.5,
      justification: 'left',
    });
    expect(rest).toHaveLength(0);
    const spaces = line!.filter((segment) => segment.text.endsWith(' '))
      .map((segment) => segment.width - advancePt(FONTS['BIZ UDGothic|700']!,
        segment.text.trimEnd(), 8.5));
    expect(spaces).toHaveLength(4);
    for (const space of spaces) expect(space).toBeCloseTo(3.5, 6);
  });

  it('fits identically when a source run starts inside the terminal cluster', () => {
    // Review probe: round-3 kana + closing-parenthesis class at its first fit
    // (102pt band). Splitting the run before the parenthesis, or before the
    // kana, must not change the partition or the retained space advances.
    const text = '甲甲甲甲  + 乙乙 +  丙丙ｱ）';
    const layout = (chunks: readonly string[]) => layoutStubParagraph({
      runs: chunks.map((chunk) => ({ ...BIZ, text: chunk })),
      environment: MODE_14,
      bandPt: 102,
      justification: 'left',
    });
    const joined = layout([text]);
    expect(joined).toHaveLength(1);
    const visible = (lines: ReturnType<typeof layout>) => lines.map((line) =>
      line.map((segment) => segment.text).join(''));
    const total = (lines: ReturnType<typeof layout>) => lines.map((line) =>
      line.reduce((sum, segment) => sum + segment.width, 0));
    for (const split of [[text.slice(0, -1), text.slice(-1)], [text.slice(0, -2), text.slice(-2)]]) {
      const lines = layout(split);
      expect(visible(lines)).toEqual(visible(joined));
      expect(total(lines)[0]).toBeCloseTo(total(joined)[0]!, 9);
    }
    // One pt narrower Word wraps; joined and split agree there too.
    const narrow = (chunks: readonly string[]) => layoutStubParagraph({
      runs: chunks.map((chunk) => ({ ...BIZ, text: chunk })),
      environment: MODE_14, bandPt: 101, justification: 'left',
    }).map((line) => line.map((segment) => segment.text).join(''));
    expect(narrow([text.slice(0, -1), text.slice(-1)])).toEqual(narrow([text]));
  });
});

// Deterministic PRNG (mulberry32) so property cases are reproducible.
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
  };
}

/** Split points inside visible text only: a U+0020 at a source-run boundary
 * has its own registered seam rule (WORD_SOURCE_RUN_SPACE_SEQUENCE). */
function randomChunks(text: string, next: () => number): string[] {
  const characters = [...text];
  const legal = characters.flatMap((character, index) => index > 0
    && character !== ' ' && characters[index - 1] !== ' ' ? [index] : []);
  const cuts = [...new Set(Array.from({ length: 1 + Math.floor(next() * 3) },
    () => legal[Math.floor(next() * legal.length)]!))].sort((a, b) => a - b);
  return [0, ...cuts].map((start, index) =>
    characters.slice(start, cuts[index] ?? characters.length).join(''));
}

describe('WORD_COMPRESSED_SPACE_LINE_FIT properties', () => {
  const mixed = inScope.filter((variant) => variant.sourceRuns === 'single'
    && /[　-鿿＀-￯]/u.test(variant.text));
  const paragraph = (variant: Variant, chunks: readonly string[], widthTwips: number) =>
    layoutStubParagraph({
      runs: chunks.map((text) => ({
        text, ascii: variant.ascii, eastAsia: variant.eastAsia, sizePt: variant.sizePt,
        bold: variant.bold, kerning: kerning(variant),
      })),
      environment: {
        compatibilityMode: variant.compatibilityMode,
        ...(variant.characterSpacingControl
          ? { characterSpacingControl: variant.characterSpacingControl } : {}),
        enableOpenTypeFeatures: variant.enableOpenTypeFeatures,
      },
      bandPt: widthTwips / 20 - BAND_DEFICIT_PT[variant.borderEighths]!,
      justification: variant.justification,
    });
  const summary = (lines: ReturnType<typeof paragraph>) => lines.map((line) => ({
    text: line.map((segment) => segment.text).join(''),
    // Per-segment widths are rounded to 1e-6pt by the stub; compare sums at 1e-4.
    width: Number(line.reduce((sum, segment) => sum + segment.width, 0).toFixed(4)),
    compression: Number(line.reduce((sum, segment) => sum + segment.compression, 0).toFixed(4)),
  }));

  it('is invariant under source-run seams inside visible text', () => {
    const next = random(1660);
    let compared = 0;
    for (const variant of mixed) {
      for (const width of variant.widthsTwips) {
        const joined = summary(paragraph(variant, [variant.text], width));
        for (let trial = 0; trial < 3; trial += 1) {
          const chunks = randomChunks(variant.text, next);
          expect(summary(paragraph(variant, chunks, width)), `${variant.variant} ${width} ${chunks.join('|')}`)
            .toEqual(joined);
          compared += 1;
        }
      }
    }
    expect(compared).toBeGreaterThan(3000);
  });

  it('keeps the measured Latin-terminal control seam-invariant', () => {
    // Review round 2: `ABCD` split as `A` / `BCD` at the 106.25pt band.
    const variant = mixed.find((item) => item.variant === 'cap-latin-word')!;
    const text = variant.text;
    const split = [text.slice(0, -3), text.slice(-3)];
    for (const width of variant.widthsTwips) {
      expect(summary(paragraph(variant, split, width))).toEqual(
        summary(paragraph(variant, [text], width)));
    }
    // Word's first fit (authored 107.25pt, the measured 106.25pt band).
    expect(paragraph(variant, split, 2145)).toHaveLength(1);
  });

  it('shrinks exactly the overflow, equally per space, never below the floor', () => {
    let compressedLines = 0;
    for (const variant of mixed) {
      for (const width of variant.widthsTwips) {
        const band = width / 20 - BAND_DEFICIT_PT[variant.borderEighths]!;
        for (const line of paragraph(variant, [variant.text], width)) {
          const placed = line.reduce((sum, segment) => sum + segment.width, 0);
          const reduction = line.reduce((sum, segment) => sum + segment.compression, 0);
          if (reduction === 0) {
            expect(placed).toBeLessThanOrEqual(band + 1e-9);
            continue;
          }
          compressedLines += 1;
          // Placed advances end exactly at the band: no excess contraction.
          expect(placed).toBeCloseTo(band, 5);
          const perSpace = line.filter((segment) => segment.compression > 0)
            .map((segment) => segment.compression / (segment.text.length - segment.text.trimEnd().length));
          for (const value of perSpace) expect(value).toBeCloseTo(perSpace[0]!, 5);
        }
      }
    }
    expect(compressedLines).toBeGreaterThan(300);
  });

  it('accounts spaces committed after the line was first shrunk', () => {
    // Review round 2: a gap added after an admission must not restore a
    // reduction that was never applied to it. 甲 admits by shrinking, then
    // `+ ` (a new gap) and 乙 follow on the same run.
    const run = { ascii: 'BIZ UDGothic', eastAsia: 'BIZ UDGothic', sizePt: 8.5, bold: true } as const;
    const text = '甲甲甲甲 + 乙乙 + 丙丙丙丙';
    for (const chunks of [
      [text], ['甲甲甲甲 + 乙乙 ', '+ 丙丙丙丙'], ['甲甲甲甲 +', ' 乙乙 + 丙丙丙丙'],
      // Space-only runs committed after the admitting ideograph.
      [text, ' ', ' '], [`${text} `, '丙'],
    ]) {
      const lines = layoutStubParagraph({
        runs: chunks.map((chunk) => ({ ...run, text: chunk })),
        environment: { compatibilityMode: 14, characterSpacingControl: 'compressPunctuation' },
        bandPt: 106.25, justification: 'left',
      });
      // Line one holds the whole text; its visible advance ends exactly at the
      // band, a line-end space keeps its natural advance, and every inner
      // space shrinks by the same 1.0625pt (4.25pt over four spaces).
      const first = lines[0]!;
      expect(first.map((segment) => segment.text).join('').trimEnd()).toBe(text);
      const end = first.at(-1)!;
      const lineEndSpace = end.compression === 0
        ? 4.25 * (end.text.length - end.text.trimEnd().length) : 0;
      const placed = first.reduce((sum, segment) => sum + segment.width, 0) - lineEndSpace;
      expect(placed).toBeCloseTo(106.25, 5);
      const inner = first.filter((segment) => segment.compression > 0);
      expect(inner).toHaveLength(4);
      for (const segment of inner) expect(segment.compression).toBeCloseTo(1.0625, 9);
    }
  });
});
