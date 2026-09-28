import { describe, expect, it } from 'vitest';
import { excelDrawingMlLineRatios } from './office-auto-line.js';

describe('excelDrawingMlLineRatios', () => {
  it('uses the usWin box, not hhea, for the Far East 1.3× line (Yu Gothic)', () => {
    const r = excelDrawingMlLineRatios({
      unitsPerEm: 2048, winAscent: 2017, winDescent: 619,
      hheaAscent: 1802, hheaDescent: -455, hheaLineGap: 1024, farEastCodePage: true,
    })!;
    const box = (2017 + 619) / 2048;
    expect(r.ascentRatio).toBeCloseTo(2017 / 2048 + 0.15 * box, 10);
    expect(r.descentRatio).toBeCloseTo(619 / 2048 + 0.15 * box, 10);
  });

  it('adds only the TEXTMETRIC external leading above other faces', () => {
    // Arial: equal boxes, so the whole hhea lineGap is external leading.
    const arial = excelDrawingMlLineRatios({
      unitsPerEm: 2048, winAscent: 1854, winDescent: 434,
      hheaAscent: 1854, hheaDescent: -434, hheaLineGap: 67, farEastCodePage: false,
    })!;
    expect(arial.ascentRatio + arial.descentRatio).toBeCloseTo(2355 / 2048, 10);
    // A usWin box taller than hhea absorbs the lineGap first.
    const tall = excelDrawingMlLineRatios({
      unitsPerEm: 1000, winAscent: 950, winDescent: 250,
      hheaAscent: 900, hheaDescent: -200, hheaLineGap: 150, farEastCodePage: false,
    })!;
    expect(tall.ascentRatio).toBeCloseTo(0.95 + 0.05, 10);
    expect(tall.descentRatio).toBeCloseTo(0.25, 10);
  });
});
