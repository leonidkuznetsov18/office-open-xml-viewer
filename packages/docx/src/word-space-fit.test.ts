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
