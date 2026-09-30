import { describe, it, expect } from 'vitest';
import {
  prepareFloatWrap,
  resolveLineFloatWindow,
  computePreparedLineFloatWindow,
  computePreparedLineFloatWindowWithDiagnostics,
  skipPastTopAndBottom,
  wordMinLineStartPx,
  WORD_MIN_LINE_START_PT,
  LINE_START_GAP_EPS_PT,
  normalizeWrapSide,
  type FloatRect,
} from './float-layout.js';
import {
  layoutLines,
  type LayoutSeg,
  type LayoutTextSeg,
  type WrapLayoutCtx,
} from './line-layout.js';

// ─────────────────────────────────────────────────────────────────────────────
// Word's measured minimum line-start rule beside a float (issue #676).
//
// GROUND TRUTH (fixtures private/sample-19/20/22, Word-exported PDF, pdftotext
// bbox): Word starts a CONTENT line beside a float ONLY when the free horizontal
// gap is ≥ 72pt (= 1 inch = 1440 twips), and always when it is. For a content
// line the threshold is:
//   - text-independent (a short-token line and a long-word line switch at the
//     same width),
//   - font-size-independent (8/12/24pt switch at 72pt),
//   - line-spacing-independent (single/1.5/double switch at 72pt),
// i.e. an ABSOLUTE width, not an em- or line-height-proportional quantity. At a
// gap of 70pt the line flows below the band; at 72pt it sits beside. A first
// word that overruns the ≥1-inch gap is force-broken there (Word's "AFTE"/"R-10"
// wrap), not refused.
//
// SCOPE — this 1-inch rule is the CONTENT-line threshold. A literally-empty /
// anchor-only paragraph's pilcrow uses the NARROWER pilcrow-em threshold
// (paragraphMarkEmPx via resolveEmptyMarkTop / flowMarkLine): Word keeps such a
// mark beside a float down to a sub-inch gap and drops it below only for a
// full-width band (sample-9 p.4 + sample-12 p.2; the #676 change wrongly
// applied 1 inch to empty marks). The (c) case below exercises the layoutLines
// EMPTY-CONTENT-line path (a paragraph carrying an empty text segment), which is
// a content line and keeps the 1-inch rule.
//
// The boundary is INCLUSIVE at 1 inch, but a frame authored so the gap is
// exactly 1 inch computes as content-width − frame-width slightly under 72
// (sample-22 p.7 → 71.963716pt in this renderer). The callers therefore pass
// `wordMinLineStartPx(scale)` = (72 − LINE_START_GAP_EPS_PT) × scale, a half-twip
// rounding tolerance, so the effective threshold is 71.95pt at scale 1: 70/71.9pt
// stay below, 71.95pt and up (incl. the 71.96pt computed for a 72.0pt frame) go
// beside. See issue #676.
//
// This file pins the pure geometry gate (resolveLineFloatWindow) and the
// layoutLines integration that consumes it. It replaced the first-atomic-token-
// width probe (the former requiredLineWidth) for content lines with this single
// grounded 1-inch rule. The literally-empty-paragraph mark path
// (resolveEmptyMarkTop / flowMarkLine) keeps its own pilcrow-em threshold — see
// the SCOPE note above.
// ─────────────────────────────────────────────────────────────────────────────

/** A LEFT-anchored square float band occupying [0, floatRightPx] horizontally on
 *  rows [0, floatBottomPx). paraX is 0, so with a column of `colWpx` the free
 *  RIGHT gap is `colWpx - floatRightPx`. */
function leftBand(floatRightPx: number, floatBottomPx: number): FloatRect {
  return {
    kind: 'shape', mode: 'square', imageKey: 'x',
    imageX: 0, imageY: 0, imageW: floatRightPx, imageH: floatBottomPx,
    xLeft: 0, xRight: floatRightPx, yTop: 0, yBottom: floatBottomPx,
    side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
    paraId: 1,
  } as FloatRect;
}

function polygonFloat(
  authoredWrap: 'tight' | 'through',
  points: readonly Readonly<{ xPt: number; yPt: number }>[],
  overrides: Partial<Extract<FloatRect, { kind: 'shape' }>> = {},
): FloatRect {
  const xs = points.map((point) => point.xPt);
  const ys = points.map((point) => point.yPt);
  const xLeft = Math.min(...xs);
  const xRight = Math.max(...xs);
  const yTop = Math.min(...ys);
  const yBottom = Math.max(...ys);
  return {
    ...leftBand(xRight, yBottom),
    imageKey: authoredWrap,
    authoredWrap,
    wrapPolygon: points,
    xLeft, xRight, yTop, yBottom,
    imageX: xLeft, imageY: yTop,
    imageW: xRight - xLeft, imageH: yBottom - yTop,
    ...overrides,
  };
}

/** Query resolveLineFloatWindow with a given free-gap width (px) at a given
 *  scale, passing EXACTLY what the docx renderer passes for a line-start probe
 *  (`wordMinLineStartPx(scale)`). Returns whether the line was placed BESIDE the
 *  band (topY 0 with a non-zero xOffset) or FLOWED BELOW it (topY advanced past
 *  the band bottom). */
function placeLine(gapPx: number, scale: number): { beside: boolean; topY: number; xOffset: number } {
  const colW = 1000 * scale;
  const floatBottom = 50 * scale;
  const floatRight = colW - gapPx; // leave exactly `gapPx` of free gap on the right
  const win = resolveLineFloatWindow(
    0, wordMinLineStartPx(scale), 10 * scale, 0, colW, [leftBand(floatRight, floatBottom)],
  );
  const beside = win.topY === 0 && win.xOffset > 0;
  return { beside, topY: win.topY, xOffset: win.xOffset };
}

const resolveWithReference = resolveLineFloatWindow as unknown as (
  topY: number,
  requiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  floats: FloatRect[],
  columnXLeftPt?: number,
  columnXRightPt?: number,
  reference?: Readonly<{
    xLeftPt: number;
    xRightPt: number;
    readingDirection: 'ltr' | 'rtl';
  }>,
) => { topY: number; xOffset: number; maxWidth: number };

describe('resolveLineFloatWindow — Word 1-inch line-start gate (issue #676)', () => {
  const fullBand = (id: string, mode: FloatRect['mode'], yTop: number, yBottom: number): FloatRect => ({
    ...leftBand(100, yBottom),
    imageKey: id, mode, yTop, yBottom,
  });
  it('normalizes unknown legacy wrap sides to bothSides', () => {
    expect(normalizeWrapSide('left')).toBe('left');
    expect(normalizeWrapSide('legacy-unknown')).toBe('bothSides');
    expect(normalizeWrapSide(null)).toBe('bothSides');
  });

  it('projects a triangular tight polygon at the current line Y', () => {
    const triangle = polygonFloat('tight', [
      { xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 },
    ]);

    expect(resolveLineFloatWindow(0, 1, 10, 0, 100, [triangle]))
      .toEqual({ topY: 0, xOffset: 0, maxWidth: 20 });
    expect(resolveLineFloatWindow(80, 1, 10, 0, 100, [triangle]))
      .toEqual({ topY: 80, xOffset: 0, maxWidth: 44 });
  });

  it('advances a tight line in whole line-height steps (word-tight-wrap-line-step-advance)', () => {
    const triangle = polygonFloat('tight', [
      { xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 },
    ]);

    // The free side gap first reaches 40 pt at the contour root 200/3; the
    // line retries one line height (1 pt) lower each time, so it lands at 67.
    const result = resolveLineFloatWindow(0, 40, 1, 0, 100, [triangle]);

    expect(result.topY).toBe(67);
    expect(result.maxWidth).toBeGreaterThanOrEqual(40);
  });

  it('matches a per-line retry on sloped and self-intersecting tight polygons', () => {
    const star = Array.from({ length: 9 }, (_, index) => {
      const angle = (index * 4 * Math.PI) / 9;
      return { xPt: 50 + 45 * Math.cos(angle), yPt: 60 + 55 * Math.sin(angle) };
    });
    const shapes = [
      [{ xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 }],
      [{ xPt: 0, yPt: 10 }, { xPt: 100, yPt: 10 }, { xPt: 100, yPt: 30 },
        { xPt: 45, yPt: 60 }, { xPt: 100, yPt: 90 }, { xPt: 0, yPt: 90 }],
      star,
    ];
    for (const points of shapes) {
      const float = polygonFloat('tight', points);
      const prepared = prepareFloatWrap([float]);
      for (const [required, lineHeight] of [[40, 1], [30, 3.3], [55, 7], [10, 0.7]] as const) {
        const start = 0;
        const actual = computePreparedLineFloatWindow(start, required, lineHeight, 0, 100, prepared);
        // Reference: test every line step while the band still meets the polygon.
        let expected: number | null = null;
        for (let y = start; ; y += lineHeight) {
          const window = computePreparedLineFloatWindow(y, required, lineHeight, 0, 100, prepared);
          if (window.topY === y) { expected = y; break; }
          if (y > float.yBottom) break;
        }
        if (expected !== null) expect(actual.topY).toBeCloseTo(expected, 9);
      }
    }
  });

  /** First grid Y (start + k * lineHeight, k >= 0) at which the line fits,
   * testing every step while the band can still meet a polygon. */
  const perStepReference = (
    floats: FloatRect[],
    start: number,
    required: number,
    lineHeight: number,
    left: number,
    width: number,
  ) => {
    const prepared = prepareFloatWrap(floats);
    const bottom = Math.max(...floats.map((float) => float.yBottom));
    for (let k = 0; ; k += 1) {
      const y = start + k * lineHeight;
      const window = computePreparedLineFloatWindow(y, required, lineHeight, left, width, prepared);
      if (window.topY === y) return window;
      if (y > bottom) return null;
    }
  };

  it('keeps usable steps where a contour crosses the paragraph bounds', () => {
    const floats = [
      polygonFloat('tight', [{ xPt: 120, yPt: 30 }, { xPt: 70, yPt: 100 }, { xPt: 0, yPt: 60 }]),
      polygonFloat('tight', [{ xPt: 60, yPt: 0 }, { xPt: 0, yPt: 100 }, { xPt: 20, yPt: 70 }], { side: 'right' }),
    ];
    const actual = computePreparedLineFloatWindow(20, 58, 1, 0, 100, prepareFloatWrap(floats));
    expect(actual.topY).toBe(32);
    expect(actual.maxWidth).toBeCloseTo(58.2857, 3);
    expect(actual).toEqual(perStepReference(floats, 20, 58, 1, 0, 100));
  });

  it('matches per-step retries on random tight polygons that cross the paragraph bounds', () => {
    let seed = 0x1623;
    const random = () => {
      seed = (seed * 1103515245 + 12345) >>> 0;
      return seed / 2 ** 32;
    };
    const sides = ['bothSides', 'left', 'right', 'largest'] as const;
    let compared = 0;
    for (let trial = 0; trial < 200; trial += 1) {
      const floats = Array.from({ length: 1 + Math.floor(random() * 3) }, () => polygonFloat(
        'tight',
        Array.from({ length: 3 + Math.floor(random() * 3) }, () => ({
          xPt: Math.round(-40 + random() * 180),
          yPt: Math.round(random() * 120),
        })),
        { side: sides[Math.floor(random() * sides.length)]! },
      )).filter((float) => float.yBottom > float.yTop && float.xRight > float.xLeft);
      if (floats.length === 0) continue;
      const start = Math.round(random() * 30);
      const required = 5 + Math.round(random() * 70);
      const lineHeight = [0.7, 1, 2.5, 5][Math.floor(random() * 4)]!;
      const expected = perStepReference(floats, start, required, lineHeight, 0, 100);
      if (expected === null) continue;
      const actual = computePreparedLineFloatWindow(
        start, required, lineHeight, 0, 100, prepareFloatWrap(floats),
      );
      expect(actual, `trial ${trial}`).toEqual(expected);
      compared += 1;
    }
    expect(compared).toBeGreaterThan(100);
  }, 120_000);

  it('bounds the tight line-step search by geometry, not by the number of steps', () => {
    const tall = polygonFloat('tight', [
      { xPt: 0, yPt: 0 }, { xPt: 100, yPt: 0 }, { xPt: 100, yPt: 10_000 }, { xPt: 0, yPt: 10_000 },
    ]);
    const { window, diagnostics } = computePreparedLineFloatWindowWithDiagnostics(
      0, 10, 0.05, 0, 100, prepareFloatWrap([tall]),
    );
    expect(window.topY).toBeGreaterThan(10_000);
    expect(window.topY).toBeLessThanOrEqual(10_000 + 0.05 + 1e-6);
    expect(diagnostics.evaluatedYCount).toBeLessThanOrEqual(8);
  });

  it('keeps sweeping to the earliest contour root for through wrap', () => {
    const triangle = polygonFloat('through', [
      { xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 },
    ]);

    const result = resolveLineFloatWindow(0, 40, 1, 0, 100, [triangle]);

    expect(result.topY).toBeCloseTo(200 / 3, 10);
    expect(result.maxWidth).toBeCloseTo(40, 10);
  });

  it('keeps a concave polygon interior gap for through but not tight', () => {
    const concave = [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 80 }, { xPt: 80, yPt: 80 },
      { xPt: 80, yPt: 20 }, { xPt: 20, yPt: 20 },
      { xPt: 20, yPt: 80 }, { xPt: 10, yPt: 80 },
    ];

    expect(resolveLineFloatWindow(30, 1, 10, 0, 100, [polygonFloat('tight', concave)]))
      .toEqual({ topY: 30, xOffset: 0, maxWidth: 10 });
    expect(resolveLineFloatWindow(30, 1, 10, 0, 100, [polygonFloat('through', concave)]))
      .toEqual({ topY: 30, xOffset: 20, maxWidth: 60 });
  });

  it('keeps a through-notch opening whose area begins at the line top', () => {
    const notch = [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 100 }, { xPt: 70, yPt: 100 },
      { xPt: 70, yPt: 40 }, { xPt: 30, yPt: 40 },
      { xPt: 30, yPt: 100 }, { xPt: 10, yPt: 100 },
    ];

    expect(resolveLineFloatWindow(40, 1, 10, 0, 100, [polygonFloat('through', notch)]))
      .toEqual({ topY: 40, xOffset: 30, maxWidth: 40 });
  });

  it('does not apply square compatibility to a polygon gap the square does not constrain', () => {
    const notch = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 100 }, { xPt: 70, yPt: 100 },
      { xPt: 70, yPt: 40 }, { xPt: 30, yPt: 40 },
      { xPt: 30, yPt: 100 }, { xPt: 10, yPt: 100 },
    ]);
    const containedSquare = {
      ...leftBand(20, 100),
      imageKey: 'contained-square',
      xLeft: 15,
      imageX: 15,
      imageW: 5,
    };

    expect(computePreparedLineFloatWindow(
      50, 1, 10, 0, 100,
      prepareFloatWrap([notch, containedSquare]),
      0, 100,
      { xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr' },
      72,
    )).toEqual({ topY: 50, xOffset: 30, maxWidth: 40 });
  });

  it('applies square compatibility when the square constrains the selected polygon gap', () => {
    const notch = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 100 }, { xPt: 70, yPt: 100 },
      { xPt: 70, yPt: 40 }, { xPt: 30, yPt: 40 },
      { xPt: 30, yPt: 100 }, { xPt: 10, yPt: 100 },
    ]);
    const gapBoundingSquare = {
      ...leftBand(35, 100),
      imageKey: 'gap-bounding-square',
      xLeft: 25,
      imageX: 25,
      imageW: 10,
    };

    expect(computePreparedLineFloatWindow(
      50, 1, 10, 0, 100,
      prepareFloatWrap([notch, gapBoundingSquare]),
      0, 100,
      { xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr' },
      72,
    )).toEqual({ topY: 100, xOffset: 0, maxWidth: 100 });
  });

  it('advances to the exact root where an eligible through gap becomes geometrically widest', () => {
    const opening = polygonFloat('through', [
      { xPt: 130, yPt: 0 }, { xPt: 200, yPt: 0 },
      { xPt: 200, yPt: 100 }, { xPt: 195, yPt: 100 },
      { xPt: 160, yPt: 20 }, { xPt: 140, yPt: 20 },
      { xPt: 140, yPt: 100 }, { xPt: 130, yPt: 100 },
    ]);
    const square = { ...leftBand(80, 100), imageKey: 'square-boundary' };

    const result = computePreparedLineFloatWindow(
      50, 1, 0.5, 0, 200,
      prepareFloatWrap([square, opening]),
      0, 200,
      { xLeftPt: 0, xRightPt: 200, readingDirection: 'ltr' },
      72,
    );

    expect(result.topY).toBeCloseTo(620 / 7, 10);
    expect(result).toMatchObject({ xOffset: 140, maxWidth: 50 });
  });

  it('supports an inferred-closure bow-tie whose signed shoelace area is zero', () => {
    const bowTie = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 100 },
      { xPt: 10, yPt: 100 }, { xPt: 90, yPt: 0 },
    ]);

    expect(resolveLineFloatWindow(20, 1, 10, 0, 100, [bowTie]))
      .toEqual({ topY: 20, xOffset: 0, maxWidth: 26 });
  });

  it('applies wrap distances to polygon line-band and horizontal projection', () => {
    const padded = polygonFloat('tight', [
      { xPt: 20, yPt: 20 }, { xPt: 40, yPt: 20 }, { xPt: 30, yPt: 40 },
    ], { xLeft: 15, xRight: 47, yTop: 10, yBottom: 46 });

    expect(resolveLineFloatWindow(45, 1, 1, 0, 100, [padded]))
      .toEqual({ topY: 45, xOffset: 37.5, maxWidth: 62.5 });
  });

  it('applies authored left, right, and largest sides after polygon projection', () => {
    const triangle = [
      { xPt: 20, yPt: 0 }, { xPt: 80, yPt: 0 }, { xPt: 50, yPt: 100 },
    ];
    const resolve = (side: FloatRect['side']) => resolveLineFloatWindow(
      80, 1, 10, 0, 100, [polygonFloat('tight', triangle, { side })],
    );

    expect(resolve('left')).toEqual({ topY: 80, xOffset: 0, maxWidth: 44 });
    expect(resolve('right')).toEqual({ topY: 80, xOffset: 56, maxWidth: 44 });
    expect(resolve('largest')).toEqual({ topY: 80, xOffset: 0, maxWidth: 44 });
  });

  it('selects a polygon largest side from its maximum extent for every line band', () => {
    const asymmetric = prepareFloatWrap([polygonFloat('tight', [
      { xPt: 60, yPt: 0 }, { xPt: 70, yPt: 0 },
      { xPt: 20, yPt: 100 }, { xPt: 10, yPt: 100 },
    ], { side: 'largest' })]);
    const resolve = (topY: number) => computePreparedLineFloatWindow(
      topY, 1, 10, 0, 100, asymmetric,
      0, 100,
      { xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr' },
    );

    // §20.4.3.7 selects the object's right side from its full [10, 70]
    // horizontal extent. Its top contour locally lies on the opposite half of
    // the page, but that line-band projection cannot reverse the selected side.
    expect(resolve(0)).toEqual({ topY: 0, xOffset: 70, maxWidth: 30 });
    expect(resolve(90)).toEqual({ topY: 90, xOffset: 25, maxWidth: 75 });
  });

  it('resolves a centered largest object by the first intersecting line direction', () => {
    const centered = { ...leftBand(60, 100), xLeft: 40, imageX: 40, imageW: 20, side: 'largest' };

    expect(resolveWithReference(0, 1, 10, 0, 100, [centered], 0, 100, {
      xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr',
    })).toEqual({ topY: 0, xOffset: 0, maxWidth: 40 });
    expect(resolveWithReference(0, 1, 10, 0, 100, [centered], 0, 100, {
      xLeftPt: 0, xRightPt: 100, readingDirection: 'rtl',
    })).toEqual({ topY: 0, xOffset: 60, maxWidth: 40 });
  });

  it('intersects the independently selected sides of multiple largest objects', () => {
    const left = { ...leftBand(40, 100), xLeft: 20, imageX: 20, imageW: 20, side: 'largest', imageKey: 'left' };
    const right = { ...leftBand(80, 100), xLeft: 60, imageX: 60, imageW: 20, side: 'largest', imageKey: 'right' };

    expect(resolveWithReference(0, 1, 10, 0, 100, [left, right], 0, 100, {
      xLeftPt: 0, xRightPt: 100, readingDirection: 'ltr',
    })).toEqual({ topY: 0, xOffset: 40, maxWidth: 20 });
  });

  it('rejects tight and through floats without a finite nonzero polygon', () => {
    const missing = { ...leftBand(80, 20), authoredWrap: 'tight' as const };
    const nonfinite = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 }, { xPt: 50, yPt: 20 },
    ], { wrapPolygon: [{ xPt: Number.NaN, yPt: 0 }, { xPt: 90, yPt: 0 }, { xPt: 50, yPt: 20 }] });

    expect(() => resolveLineFloatWindow(0, 1, 10, 0, 100, [missing]))
      .toThrow(/invalid tight wrapPolygon/i);
    expect(() => resolveLineFloatWindow(0, 1, 10, 0, 100, [nonfinite]))
      .toThrow(/invalid through wrapPolygon/i);
  });

  it('crosses more than sixteen chained topAndBottom bottoms', () => {
    const floats = Array.from({ length: 20 }, (_, index) =>
      fullBand(`top-${index}`, 'topAndBottom', index, index + 1));

    expect(skipPastTopAndBottom(0, floats, 0, 100)).toBe(20);
    expect(resolveLineFloatWindow(0, 1, 0.5, 0, 100, floats).topY).toBe(20);
  });

  it('crosses more than sixty-four chained square bottoms', () => {
    const floats = Array.from({ length: 70 }, (_, index) =>
      fullBand(`square-${index}`, 'square', index, index + 1));

    expect(resolveLineFloatWindow(0, 1, 0.5, 0, 100, floats).topY).toBe(70);
  });

  it('rechecks topAndBottom after a square push', () => {
    const floats = [
      fullBand('square', 'square', 0, 10),
      fullBand('top', 'topAndBottom', 10, 20),
    ];

    expect(resolveLineFloatWindow(0, 1, 0.5, 0, 100, floats).topY).toBe(20);
  });

  it('the grounded constant is exactly 1 inch (72pt) with a one-twip tolerance', () => {
    expect(WORD_MIN_LINE_START_PT).toBe(72);
    expect(LINE_START_GAP_EPS_PT).toBe(0.05); // one twip (1/20 pt)
    expect(wordMinLineStartPx(1)).toBeCloseTo(71.95, 10);
    expect(wordMinLineStartPx(2)).toBeCloseTo(143.9, 10);
  });

  it('(a) a 71.9pt gap flows the line BELOW the band (clear of the tolerance band)', () => {
    // 71.9 < 71.95 effective threshold → below. 71.9pt is the largest "below"
    // probe that stays outside the one-twip tolerance (a genuinely sub-inch gap).
    const r = placeLine(71.9, 1);
    expect(r.beside).toBe(false);
    expect(r.topY).toBe(50); // pushed to the band bottom
  });

  it('(b) a 72.0pt gap places the line BESIDE the band (exactly 1 inch)', () => {
    const r = placeLine(72.0, 1);
    expect(r.beside).toBe(true);
    expect(r.topY).toBe(0);
    expect(r.xOffset).toBeGreaterThan(0);
  });

  it('(b) sample-22 p.7: a gap computed at 71.9637pt (a 72.0pt frame) is BESIDE', () => {
    // The exact value this renderer computes for the gap=72.0pt frame — the
    // tolerance exists precisely so this lands beside, matching Word's PDF.
    expect(placeLine(71.963716, 1).beside).toBe(true);
  });

  it('a 70pt gap is below, a 74pt gap is beside (the sample-22 bracket)', () => {
    expect(placeLine(70, 1).beside).toBe(false);
    expect(placeLine(74, 1).beside).toBe(true);
  });

  it('(e) the 70/72pt boundary is identical in PT space at scale 0.75', () => {
    const s = 0.75;
    // 70pt and 72pt gaps expressed in px at this scale must still switch across
    // the 1-inch boundary (requiredWidth is wordMinLineStartPx(scale), so the
    // decision is taken in pt space and is scale-invariant).
    expect(placeLine(70 * s, s).beside).toBe(false);
    expect(placeLine(72 * s, s).beside).toBe(true);
  });

  it('(e) the boundary is identical across scales (absolute pt width)', () => {
    for (const s of [1, 2, 0.5, 1.5, 0.75, 3]) {
      expect(placeLine(70 * s, s).beside).toBe(false); // 70pt < 1 inch → below
      expect(placeLine(72 * s, s).beside).toBe(true);  // 72pt = 1 inch → beside
    }
  });

  it('is a pure width gate: no content input, so font size / empty-vs-filled cannot matter', () => {
    // resolveLineFloatWindow takes only a numeric requiredWidth — there is no
    // content input at all. The gate therefore cannot depend on font size or the
    // line being empty vs. filled; every caller resolves to wordMinLineStartPx.
    // (c)+(d) parity is enforced structurally by the single call site.
    expect(placeLine(71.9, 1).beside).toBe(false);
    expect(placeLine(72.0, 1).beside).toBe(true);
  });

  it('ignores square wrap rectangles wholly outside either side of the paragraph column', () => {
    const outsideRanges = [
      { xLeft: 20, xRight: 100 },
      { xLeft: 190, xRight: 270 },
    ];

    for (const side of ['bothSides', 'left', 'right', 'largest']) {
      for (const range of outsideRanges) {
        const outsideColumn = {
          ...leftBand(80, 120),
          ...range,
          imageX: range.xLeft,
          imageW: range.xRight - range.xLeft,
          side,
        };
        const win = resolveLineFloatWindow(
          20,
          wordMinLineStartPx(1),
          10,
          110,
          70,
          [outsideColumn],
        );

        expect(win).toEqual({ topY: 20, xOffset: 0, maxWidth: 70 });
      }
    }
  });
});

// ── layoutLines integration ──────────────────────────────────────────────────
// Linear mock canvas: glyph advance = perPx · px · chars; ascent/descent 0.8/0.2
// em. Perfectly scale-linear so the wrap ALGORITHM is isolated from font hinting.
function makeLinearCtx(perPx = 0.5): CanvasRenderingContext2D {
  let font = '10px serif';
  const pxOf = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = pxOf();
      const per = p * perPx;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function textSeg(text: string, fontSize = 10, extra: Partial<LayoutTextSeg> = {}): LayoutSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'Times New Roman', vertAlign: null,
    measuredWidth: 0, ...extra,
  } as LayoutSeg;
}

function wrapCtx(floats: FloatRect[]): WrapLayoutCtx {
  return {
    startPageY: 0,
    paraX: 0,
    floats,
    lineBoxH: (asc: number, desc: number) => asc + desc,
    pageH: 100000,
  } as WrapLayoutCtx;
}

/** Was the FIRST line placed beside the band (topY 0, xOffset > 0) or flowed
 *  below it (topY past the band bottom)? */
function firstLinePlacement(lines: ReturnType<typeof layoutLines>): 'beside' | 'below' {
  const l = lines[0];
  return l.topY === 0 && l.xOffset > 0 ? 'beside' : 'below';
}

describe('layoutLines — 1-inch line-start rule end to end (issue #676)', () => {
  const scale = 1;
  const colW = 1000;
  const floatBottom = 50;

  // A gap just under 1 inch (70px) and just over (72px) at scale 1.
  const bandFor = (gapPx: number) => [leftBand(colW - gapPx, floatBottom)];

  it('(c) an empty-content CONTENT line flows below a sub-inch gap and beside a ≥1-inch gap', () => {
    // A content paragraph whose sole segment is empty text — the layoutLines
    // content-line path, which keeps the 1-inch rule. (A literally-empty
    // paragraph with NO runs is placed by resolveEmptyMarkTop against the
    // narrower pilcrow-em threshold instead — see SCOPE note.)
    const emptyBelow = layoutLines(makeLinearCtx(), [textSeg('', 10)], colW, 0, scale, [], wrapCtx(bandFor(70)), {}, 0);
    const emptyBeside = layoutLines(makeLinearCtx(), [textSeg('', 10)], colW, 0, scale, [], wrapCtx(bandFor(72)), {}, 0);
    expect(firstLinePlacement(emptyBelow)).toBe('below');
    expect(firstLinePlacement(emptyBeside)).toBe('beside');
  });

  it('keeps an anchor-host metric-only line on the paragraph-mark threshold', () => {
    const markWrap = {
      ...wrapCtx(bandFor(62)),
      paragraphMarkLineStartWidth: 10,
    };
    const lines = layoutLines(
      makeLinearCtx(),
      [textSeg('', 10, { metricOnly: true })],
      colW,
      0,
      scale,
      [],
      markWrap,
      {},
      0,
    );

    // A zero-advance anchor-character placeholder preserves the run's line
    // metrics, but it does not turn the pilcrow into inline content. The 62pt
    // side gap holds the 10pt mark even though it is below the 1-inch content
    // threshold, so the host line stays beside the float.
    expect(firstLinePlacement(lines)).toBe('beside');
    expect(lines[0].ascent + lines[0].descent).toBe(10);
  });

  it('(c) a text line makes the SAME below/beside decision as the empty line', () => {
    const textBelow = layoutLines(makeLinearCtx(), [textSeg('hi', 10)], colW, 0, scale, [], wrapCtx(bandFor(70)), {}, 0);
    const textBeside = layoutLines(makeLinearCtx(), [textSeg('hi', 10)], colW, 0, scale, [], wrapCtx(bandFor(72)), {}, 0);
    expect(firstLinePlacement(textBelow)).toBe('below');
    expect(firstLinePlacement(textBeside)).toBe('beside');
  });

  it('(d) the below/beside decision is font-size-independent (8pt vs 24pt agree)', () => {
    for (const fs of [8, 24]) {
      const below = layoutLines(makeLinearCtx(), [textSeg('X', fs)], colW, 0, scale, [], wrapCtx(bandFor(70)), {}, 0);
      const beside = layoutLines(makeLinearCtx(), [textSeg('X', fs)], colW, 0, scale, [], wrapCtx(bandFor(72)), {}, 0);
      expect(firstLinePlacement(below)).toBe('below');
      expect(firstLinePlacement(beside)).toBe('beside');
    }
  });

  it('(d) a SHORT token no longer wedges into a sub-inch gap it would have fit', () => {
    // "X" at 10pt is 5px wide — under the old 1-em (10px) probe it might have
    // been rejected, but a longer prior-behaviour concern was a short token
    // fitting a sub-inch sliver. With the 1-inch rule, a 5px-wide token in a
    // 30px gap (well under 1 inch) is sent below, matching Word.
    const lines = layoutLines(makeLinearCtx(), [textSeg('X', 10)], colW, 0, scale, [], wrapCtx(bandFor(30)), {}, 0);
    expect(firstLinePlacement(lines)).toBe('below');
  });

  it('force-wrap: a word wider than a ≥1-inch gap is CHAR-BROKEN in the gap (Word "AFTE"/"R-10")', () => {
    // Gap = 72px (exactly 1 inch). Word "AFTERTENAFTERTEN" = 16 chars × 5px = 80px,
    // wider than the 72px gap. The line IS started beside the band (gap ≥ 1 inch)
    // and the word is force-broken to fit — it is NOT sent below.
    const lines = layoutLines(
      makeLinearCtx(), [textSeg('AFTERTENAFTERTEN', 10)], colW, 0, scale, [], wrapCtx(bandFor(72)), {}, 0,
    );
    // First line beside the band, holding as many chars as fit the 72px gap.
    expect(firstLinePlacement(lines)).toBe('beside');
    expect(lines.length).toBeGreaterThan(1); // the word was split across lines
    const firstText = (lines[0].segments[0] as LayoutTextSeg).text;
    const secondText = (lines[1].segments[0] as LayoutTextSeg).text;
    // 72px gap / 5px per char = 14 chars fit; the split preserves the whole word.
    expect(firstText.length).toBeGreaterThan(0);
    expect(firstText.length).toBeLessThan('AFTERTENAFTERTEN'.length);
    expect(firstText + secondText).toBe('AFTERTENAFTERTEN');
    // The line sat in the gap (xOffset at the band's right edge), not full width.
    expect(lines[0].xOffset).toBeGreaterThan(0);
    expect(lines[0].availWidth).toBeLessThanOrEqual(72 + 1e-6);
  });

  it('a word narrower than the ≥1-inch gap sits beside the band without splitting', () => {
    // Gap = 200px. "AFTER" = 5 chars × 5px = 25px < 200px → sits beside, no split.
    const lines = layoutLines(makeLinearCtx(), [textSeg('AFTER', 10)], colW, 0, scale, [], wrapCtx(bandFor(200)), {}, 0);
    expect(firstLinePlacement(lines)).toBe('beside');
    expect((lines[0].segments[0] as LayoutTextSeg).text).toBe('AFTER');
  });

  it('remeasures a tall line against its actual polygon band', () => {
    const inverted = polygonFloat('tight', [
      { xPt: 50, yPt: 0 }, { xPt: 100, yPt: 100 }, { xPt: 0, yPt: 100 },
    ]);

    const lines = layoutLines(
      makeLinearCtx(), [textSeg('X', 60)], 100, 0, 1, [], wrapCtx([inverted]), {}, 0,
    );

    expect(lines[0].ascent + lines[0].descent).toBe(60);
    expect(lines[0].availWidth).toBe(20);
  });

  it('diagnoses an exact non-adjacent line-window cycle without a pass cap', () => {
    let calls = 0;
    const cycling: WrapLayoutCtx = {
      ...wrapCtx([]),
      lineWindow: (input) => {
        calls += 1;
        return {
          topYPt: input.topYPt,
          xOffsetPt: 0,
          maximumWidthPt: calls % 2 === 1 ? 50 : 100,
        };
      },
    };

    expect(() => layoutLines(
      makeLinearCtx(), [textSeg('AAAAA '), textSeg('BBBBB')], 100, 0, 1, [], cycling, {}, 0,
    )).toThrow(/measure\/resolve cycle did not converge/i);
  });

  it('uses a spec-permitted through opening without the square-only one-inch policy', () => {
    const notch = polygonFloat('through', [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 100 }, { xPt: 70, yPt: 100 },
      { xPt: 70, yPt: 40 }, { xPt: 30, yPt: 40 },
      { xPt: 30, yPt: 100 }, { xPt: 10, yPt: 100 },
    ]);
    const context = { ...wrapCtx([notch]), startPageY: 40 };

    const lines = layoutLines(
      makeLinearCtx(), [textSeg('word', 10)], 100, 0, 1, [], context, {}, 0,
    );

    expect(lines[0]).toMatchObject({ topY: 40, xOffset: 30, availWidth: 40 });
  });
});
