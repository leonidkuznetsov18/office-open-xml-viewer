import { describe, expect, it } from 'vitest';
import { PT_TO_PX, type OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import { bindXlsxOfficeFontRoutes, drawShapeText } from './renderer.js';
import { xlsxWorksheetOfficeFontRequests } from './google-fonts.js';
import { shapeRunLineRatios } from './shape-office-line.js';
import type { ShapeParagraph, ShapeText, ShapeTextRun, Worksheet } from './types.js';

// Expectations below come from Excel for Mac 16.113.2 PDF exports of the
// issue #1604 controls (baselines relative to the shape top, top anchor,
// tIns 3.6 pt). Arial: ascent 0.938 em (usWinAscent + lineGap), descent
// 0.212 em. Meiryo: 1.95 em box, ascent 1.285 em.

type TextRun = Extract<ShapeTextRun, { type: 'text' }>;

function run(text: string, fontFace: string, size: number, bold = false): TextRun {
  return { type: 'text', text, fontFace, fontFaceEa: fontFace, bold, italic: false, size };
}

function para(runs: TextRun[], extra: Partial<ShapeParagraph> = {}): ShapeParagraph {
  return { align: 'l', runs, ...extra };
}

function body(paragraphs: ShapeParagraph[], anchor: 't' | 'b' = 't'): ShapeText {
  return {
    anchor, wrap: 'none', autoFit: 'none',
    lIns: 91440, rIns: 91440, tIns: 45720, bIns: 45720, paragraphs,
  };
}

function context(): { ctx: CanvasRenderingContext2D; draws: Array<{ text: string; y: number; font: string; baseline: string }> } {
  let font = '11px sans-serif';
  const draws: Array<{ text: string; y: number; font: string; baseline: string }> = [];
  const ctx = {
    get font() { return font; }, set font(value: string) { font = value; },
    measureText() { return { width: 40, actualBoundingBoxAscent: 20 }; },
    fillText(text: string, _x: number, y: number) {
      draws.push({ text, y, font, baseline: (ctx as { textBaseline: string }).textBaseline });
    },
    fillStyle: '#000', textBaseline: 'alphabetic',
  } as unknown as CanvasRenderingContext2D;
  return { ctx, draws };
}

function route(family: string, weight: 400 | 700 = 400): OfficeFontFallbackRoute {
  const alias = `__exact_${family.toLowerCase().replace(/\s+/gu, '_')}_${weight}`;
  return {
    requestedFamily: family, family: alias, source: 'local',
    resourceIdentity: `office-local:local("${family}")`,
    weight, style: 'normal', metric: { family: alias, synthesized: false },
  };
}

const ROUTES = {
  arial: route('Arial'),
  meiryo: route('Meiryo'),
  'ms gothic': route('MS Gothic'),
  'yu gothic': route('Yu Gothic'),
  'times new roman': route('Times New Roman'),
};

/** Office PDF positions carry up to 0.16 pt of export rounding in these controls. */
function expectOffice(actual: number, office: number): void {
  expect(Math.abs(actual - office)).toBeLessThanOrEqual(0.16);
}

/** Paint and return each line's baseline in pt below the shape top. */
function baselinesPt(text: ShapeText, routes: Record<string, OfficeFontFallbackRoute> | null = ROUTES): number[] {
  const { ctx, draws } = context();
  if (routes) bindXlsxOfficeFontRoutes(ctx, {} as Worksheet, routes);
  drawShapeText(ctx, text, 400, 900, 1);
  return [...new Set(draws.map((draw) => draw.y))].map((y) => y / PT_TO_PX);
}

describe('Excel shape-text line box from font metrics (#1604)', () => {
  it('projects the Office line box from usWin metrics and the Far East class', () => {
    const sum = (r: ReturnType<typeof shapeRunLineRatios>) => (r ? r.ascentRatio + r.descentRatio : NaN);
    // Excel: Arial pitch 1.150 em (usWin box + hhea lineGap 67).
    expect(sum(shapeRunLineRatios(run('Hg', 'Arial', 24), ROUTES.arial))).toBeCloseTo(1.1499, 4);
    // Excel: Yu Gothic pitch 1.673 em = 1.3 × usWin box (the hhea box would give 1.433).
    expect(sum(shapeRunLineRatios(run('日g', 'Yu Gothic', 24), ROUTES['yu gothic']))).toBeCloseTo(1.6732, 3);
    // Earlier single-line controls: Meiryo UI Bold measured 1.646–1.649 em.
    expect(sum(shapeRunLineRatios(run('予算', 'Meiryo UI', 25, true), route('Meiryo UI', 700)))).toBeCloseTo(1.651, 3);
    // Office and macOS ship Times New Roman with different hhea lineGap: ambiguous.
    expect(shapeRunLineRatios(run('Hg', 'Times New Roman', 24), ROUTES['times new roman'])).toBeUndefined();
    // Unverified resources and distinct East Asian faces stay undefined.
    expect(shapeRunLineRatios(run('Hg', 'Arial', 24), { ...ROUTES.arial, resourceIdentity: 'injected:x' })).toBeUndefined();
    expect(shapeRunLineRatios({ ...run('Hg', 'Arial', 24), fontFaceEa: 'Meiryo' }, ROUTES.arial)).toBeUndefined();
  });

  it('steps same-size lines by the font line box and places the first baseline at the ascent', () => {
    // Excel S-Arial-24: 26.18, 53.78, 81.38 pt. S-Meiryo-24: 34.46, 81.26, 128.09 pt.
    const arial = baselinesPt(body([para([run('A1', 'Arial', 24)]), para([run('A2', 'Arial', 24)]),
      para([run('A3', 'Arial', 24)])]));
    expect(arial.map((v) => +v.toFixed(1))).toEqual([26.1, 53.7, 81.3]);
    const meiryo = baselinesPt(body([para([run('M1', 'Meiryo', 24)]), para([run('M2', 'Meiryo', 24)])]));
    expect(meiryo[0]).toBeCloseTo(34.44, 1);
    expect(meiryo[1] - meiryo[0]).toBeCloseTo(46.8, 1);
  });

  it('unions run ascents and descents on one shared baseline', () => {
    // Excel X-Arial-Meiryo: pitches 35.91 and 38.52 pt around a mixed line.
    const { ctx, draws } = context();
    bindXlsxOfficeFontRoutes(ctx, {} as Worksheet, ROUTES);
    drawShapeText(ctx, body([
      para([run('A1', 'Arial', 24)]),
      para([run('Hx', 'Arial', 24), run('日g', 'Meiryo', 24)]),
      para([run('A3', 'Arial', 24)]),
    ]), 400, 900, 1);
    const y = (text: string) => draws.find((draw) => draw.text === text)!.y / PT_TO_PX;
    expect(y('Hx')).toBe(y('日g'));
    expect(y('Hx') - y('A1')).toBeCloseTo(35.91, 1);
    expect(y('A3') - y('Hx')).toBeCloseTo(38.52, 1);
    // Excel M-Arial-seq: 11→40→11 pt lines pitch 39.84 then 18.86 pt.
    const seq = baselinesPt(body([11, 40, 11].map((size, i) => para([run(`S${i}`, 'Arial', size)]))));
    expectOffice(seq[1] - seq[0], 39.84);
    expectOffice(seq[2] - seq[1], 18.86);
  });

  it('re-divides the line box for spcPct and spcPts', () => {
    const three = (font: string, size: number, spaceLine: ShapeParagraph['spaceLine']) =>
      body([1, 2, 3].map((n) => para([run(`L${n}`, font, size)], { spaceLine })));
    // Excel P-Arial-40-150: first 55.36, pitch 69.0. P-Arial-40-80: first 31.93, pitch 36.84.
    const up = baselinesPt(three('Arial', 40, { type: 'pct', val: 150000 }));
    expect(up[0]).toBeCloseTo(55.35, 1);
    expect(up[1] - up[0]).toBeCloseTo(69.0, 1);
    const down = baselinesPt(three('Arial', 40, { type: 'pct', val: 80000 }));
    expect(down[0]).toBeCloseTo(31.92, 1);
    // Excel P-Meiryo-40-150 (ascent below 75 % of the box): first 84.25 pt.
    expect(baselinesPt(three('Meiryo', 40, { type: 'pct', val: 150000 }))[0]).toBeCloseTo(84.25, 1);
    // Excel T-Arial-24-18 and T-Meiryo-24-18: first 17.14 and 12.82, pitch 18.
    const arialPts = baselinesPt(three('Arial', 24, { type: 'pts', val: 18 }));
    expect(arialPts[0]).toBeCloseTo(17.1, 1);
    expect(arialPts[1] - arialPts[0]).toBeCloseTo(18, 5);
    expect(baselinesPt(three('Meiryo', 24, { type: 'pts', val: 18 }))[0]).toBeCloseTo(12.84, 1);
  });

  it('adds spcAft and spcBef between paragraphs but not before the first', () => {
    const lines = (p1: Partial<ShapeParagraph>, p2: Partial<ShapeParagraph>) => baselinesPt(body([
      para([run('P1', 'Arial', 24)], p1), para([run('P2', 'Arial', 24)], p2),
    ]));
    // Excel B-Arial-both: spcAft 12 + spcBef 18 pt → pitch 57.6 (sum).
    const both = lines({ spaceAfter: { type: 'pts', val: 12 } }, { spaceBefore: { type: 'pts', val: 18 } });
    expect(both[1] - both[0]).toBeCloseTo(57.6, 1);
    // Excel B-Arial-bef-pct50: 50 % of the 27.6 pt natural line → pitch 41.42.
    const pct = lines({}, { spaceBefore: { type: 'pct', val: 50000 } });
    expect(pct[1] - pct[0]).toBeCloseTo(41.4, 1);
    // Excel B-Arial-first-bef: spcBef on the first paragraph leaves 26.23 pt.
    expect(lines({ spaceBefore: { type: 'pts', val: 24 } }, {})[0]).toBeCloseTo(26.1, 1);
  });

  it('anchors the metric line box at the bottom of the text rectangle', () => {
    // Earlier Excel control: one 25 pt Meiryo UI Bold line in a 98 px box with
    // zero insets; top and bottom anchors exposed a 1.646–1.649 em line box.
    const text = { ...body([para([run('予算', 'Meiryo UI', 25, true)])], 'b'),
      lIns: 0, rIns: 0, tIns: 0, bIns: 0 };
    const { ctx, draws } = context();
    bindXlsxOfficeFontRoutes(ctx, {} as Worksheet, { 'meiryo ui:700:normal': route('Meiryo UI', 700) });
    drawShapeText(ctx, text, 230, 98, 1);
    const ratios = shapeRunLineRatios(run('予算', 'Meiryo UI', 25, true), route('Meiryo UI', 700))!;
    expect(draws[0].y).toBeCloseTo(98 - 25 * PT_TO_PX * ratios.descentRatio, 5);
    expect(draws[0].font).toContain('__exact_meiryo_ui_700');
  });

  it('keeps the ordinary 1.2 em box when any run lacks a verified face', () => {
    const plain = baselinesPt(body([para([run('A1', 'Arial', 20)]), para([run('A2', 'Arial', 20)])]), null);
    expect(plain[1] - plain[0]).toBeCloseTo(24, 5);
    const mixed = baselinesPt(body([para([run('A1', 'Arial', 20)]), para([run('T2', 'Times New Roman', 20)])]));
    expect(mixed[1] - mixed[0]).toBeCloseTo(24, 5);
    // Paragraph spacing still applies to the ordinary box.
    const spaced = baselinesPt(body([
      para([run('A1', 'Arial', 20)], { spaceAfter: { type: 'pts', val: 10 } }), para([run('A2', 'Arial', 20)]),
    ]), null);
    expect(spaced[1] - spaced[0]).toBeCloseTo(34, 5);
  });

  it('puts runs of another size on the largest run baseline in the ordinary box', () => {
    const { ctx, draws } = context();
    drawShapeText(ctx, body([para([run('ab', 'Arial', 11), run('CD', 'Arial', 40)])]), 400, 900, 1);
    expect(draws).toHaveLength(2);
    expect(draws[0].y).toBe(draws[1].y);
    expect(draws.every((draw) => draw.baseline === 'alphabetic')).toBe(true);
  });

  it('discovers every catalogued shape run face separately from cell fonts', () => {
    const ws = { rows: [{ cells: [{ value: { type: 'text', text: 'x', runs: [{
      text: 'x', font: { name: 'Calibri', bold: true, italic: false },
    }] } }] }], shapeGroups: [{ shapes: [
      { text: body([para([run('x', 'Meiryo UI', 25, true)]), para([run('y', 'Arial', 11)])]) },
      { text: body([para([{ ...run('z', 'Arial', 11), fontFaceEa: 'Meiryo' }])]) },
    ] }] } as unknown as Worksheet;
    expect(xlsxWorksheetOfficeFontRequests(ws)).toEqual([
      { family: 'Calibri', weight: 700, style: 'normal' },
      { family: 'Meiryo UI', weight: 700, style: 'normal' },
      { family: 'Arial', weight: 400, style: 'normal' },
    ]);
  });
});
