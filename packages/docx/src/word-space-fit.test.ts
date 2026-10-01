import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService, type ResolvedFontMetric } from './layout/text.js';
import { buildSegments, layoutLines, type LayoutTextSeg } from './line-layout.js';
import type { DocRun, DocxDocumentModel } from './types.js';
import { readFileSync } from 'node:fs';

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

interface Variant {
  readonly stage: string;
  readonly document: string;
  readonly variant: string;
  readonly compatibilityMode: number;
  readonly characterSpacingControl: string | null;
  readonly enableOpenTypeFeatures: boolean;
  readonly ascii: string;
  readonly eastAsia: string;
  readonly sizePt: number;
  readonly bold: boolean;
  readonly text: string;
  readonly justification: 'left' | 'both' | 'distribute';
  readonly sourceRuns: string;
  readonly kern: string;
  readonly borderEighths: number;
  readonly outOfScope: string | null;
  readonly widthsTwips: readonly number[];
  readonly wordWraps: string;
}
interface Face {
  readonly unitsPerEm: number;
  readonly xAvgCharWidth: number;
  readonly advances: Readonly<Record<string, number>>;
  readonly ink: Readonly<Record<string, readonly number[]>>;
}
const controls = JSON.parse(readFileSync(new URL('./word-space-fit-controls.json', import.meta.url), 'utf8')) as {
  readonly fonts: Readonly<Record<string, Face>>;
  readonly variants: readonly Variant[];
};
const FONTS = controls.fonts;
const VARIANTS = controls.variants;
const BAND_DEFICIT_PT: Readonly<Record<number, number>> = { 0: 0, 4: 1 };

function faceOf(familyList: string, weight: number): Face {
  const family = /"([^"]+)"/.exec(familyList)?.[1] ?? familyList.split(',')[0]!.trim();
  const face = FONTS[`${family}|${weight}`];
  if (!face) throw new Error(`no face ${family} ${weight}`);
  return face;
}

function advancePt(face: Face, text: string, sizePt: number): number {
  let units = 0;
  for (const character of text) {
    const value = face.advances[character];
    if (value === undefined) throw new Error(`missing advance for ${JSON.stringify(character)}`);
    units += value;
  }
  return (units * sizePt) / face.unitsPerEm;
}

let canvasFont = '';
const canvas = {
  get font() { return canvasFont; },
  set font(value: string) { canvasFont = value; },
  letterSpacing: '0px',
  fontKerning: 'auto',
  measureText(text: string) {
    const px = Number(/([\d.]+)px/.exec(canvasFont)?.[1] ?? 10);
    const weight = /\bbold\b|\b700\b/.test(canvasFont) ? 700 : 400;
    const family = canvasFont.slice(canvasFont.indexOf('px') + 2).trim();
    const width = advancePt(faceOf(family, weight), text, px);
    return {
      width, actualBoundingBoxAscent: px * 0.8, actualBoundingBoxDescent: px * 0.2,
      fontBoundingBoxAscent: px * 0.88, fontBoundingBoxDescent: px * 0.12,
    } as TextMetrics;
  },
} as unknown as CanvasRenderingContext2D;

function services(families: readonly string[]) {
  const empty = { default: null, first: null, even: null };
  const base = createLayoutServices({
    section: {
      pageWidth: 612, pageHeight: 792, marginTop: 72, marginRight: 72,
      marginBottom: 72, marginLeft: 72, headerDistance: 36, footerDistance: 36,
      titlePage: false, evenAndOddHeaders: false,
    },
    body: [], headers: empty, footers: empty,
  } as unknown as DocxDocumentModel, { measureContext: canvas });
  const routes = families.flatMap((family) => [400, 700].map((weight) => ({
    requestedFamily: family, resolvedFamily: family, source: 'local' as const,
    resourceIdentity: `office-local:local("${family}")`, weight,
  })));
  const fontMetrics: Record<string, ResolvedFontMetric> = {};
  for (const [key, face] of Object.entries(FONTS)) {
    const [family, weight] = key.split('|') as [string, string];
    if (!families.includes(family)) continue;
    fontMetrics[key] = {
      family, requestedFamily: family, weight: Number(weight),
      sourceIdentity: `office-local:local("${family}")`,
      averageCharWidthRatio: face.xAvgCharWidth / face.unitsPerEm,
      unicodeRanges: [...new Set(Object.keys(face.advances))]
        .map((character) => character.codePointAt(0)!)
        .map((code) => [code, code] as const),
    };
  }
  const text = createTextLayoutService({
    fonts: createFontResolver(routes),
    measurer: {
      fingerprint: 'word-space-fit-controls',
      measure: (request) => {
        const face = faceOf(request.fontRoute.familyList, request.weight);
        const characters = [...request.text];
        const visible = characters.filter((character) => character !== ' ');
        const advance = advancePt(face, request.text, request.fontSizePt);
        const scalePt = request.fontSizePt / face.unitsPerEm;
        // Glyph outline bounds of the embedded face (glyf), for ink-based
        // punctuation compression; spaces carry no ink.
        let last = characters.length - 1;
        while (last >= 0 && characters[last] === ' ') last -= 1;
        const inkBounds = visible.length === 0 ? undefined : {
          xMinPt: face.ink[characters.find((character) => character !== ' ')!]![0] * scalePt,
          xMaxPt: advancePt(face, characters.slice(0, last).join(''), request.fontSizePt)
            + face.ink[characters[last]!]![1] * scalePt,
          ascentPt: request.fontSizePt * 0.88, descentPt: request.fontSizePt * 0.12,
        };
        return {
          advancePt: advance, ascentPt: request.fontSizePt * 0.88,
          descentPt: request.fontSizePt * 0.12,
          ...(inkBounds ? { inkBounds, horizontalInkBoundsAreTight: true } : {}),
        };
      },
    },
    fontMetrics,
  });
  return { ...base, text };
}

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
  const layoutServices = services([variant.ascii, variant.eastAsia]);
  const runs = sourceChunks(variant).map((text) => ({
    type: 'text', text, fontFamily: variant.ascii, fontFamilyHighAnsi: variant.ascii,
    fontFamilyEastAsia: variant.eastAsia, fontSize: variant.sizePt, bold: variant.bold,
    italic: false, underline: false, strikethrough: false, kerning: kerning(variant),
    lang: 'en-US', langEastAsia: 'ja-JP',
  })) as unknown as DocRun[];
  const segments = buildSegments(runs, {
    pageIndex: 0, totalPages: 1, layoutServices,
    compatibilityMode: variant.compatibilityMode,
    ...(variant.characterSpacingControl
      ? { characterSpacingControl: variant.characterSpacingControl } : {}),
    enableOpenTypeFeatures: variant.enableOpenTypeFeatures,
    lineWrapLikeWord6: false,
  });
  const band = widthTwips / 20 - BAND_DEFICIT_PT[variant.borderEighths]!;
  const justified = variant.justification !== 'left';
  return layoutLines(canvas, segments, band, 0, 1, [], undefined, {}, 0, undefined, undefined,
    36, 0, false, justified, variant.justification === 'distribute', undefined, 'bounded',
    undefined, false);
}

const inScope = VARIANTS.filter((variant) => variant.outOfScope === null);

describe('WORD_COMPRESSED_SPACE_LINE_FIT controls', () => {
  it('covers every measured cell exactly once', () => {
    expect(VARIANTS.reduce((sum, variant) => sum + variant.widthsTwips.length, 0)).toBe(1773);
    expect(inScope.reduce((sum, variant) => sum + variant.widthsTwips.length, 0)).toBe(1471);
    // Every excluded class names a measured cause outside this rule.
    expect(new Set(VARIANTS.flatMap((variant) => variant.outOfScope
      ? [variant.outOfScope.split(':')[0]] : []))).toEqual(new Set([
      'autospace', 'kerning', 'punctuation', 'geometry', 'unexplained',
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
    // sample-36 A12 equivalent: 107.5pt band, Word paints four 3.5pt spaces.
    const variant = VARIANTS.find((item) => item.stage === 'coarse' && item.variant === 'baseline'
      && item.compatibilityMode === 14)!;
    const lines = layoutLines(canvas, buildSegments([{
      type: 'text', text: variant.text, fontFamily: variant.ascii, fontFamilyEastAsia: variant.eastAsia,
      fontSize: 8.5, bold: true, italic: false, underline: false, strikethrough: false,
    } as unknown as DocRun], {
      pageIndex: 0, totalPages: 1, layoutServices: services([variant.ascii]), compatibilityMode: 14,
      characterSpacingControl: 'compressPunctuation', enableOpenTypeFeatures: true,
    }), 107.5, 0, 1);
    expect(lines).toHaveLength(1);
    const spaces = lines[0]!.segments
      .filter((segment) => (segment as LayoutTextSeg).text.endsWith(' '))
      .map((segment) => segment.measuredWidth - advancePt(FONTS['BIZ UDGothic|700']!,
        (segment as LayoutTextSeg).text.trimEnd(), 8.5));
    expect(spaces).toHaveLength(4);
    for (const space of spaces) expect(space).toBeCloseTo(3.5, 6);
  });
});
