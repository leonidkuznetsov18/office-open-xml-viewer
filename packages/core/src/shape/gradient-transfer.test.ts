import { describe, expect, it } from 'vitest';
import { officeGradientStops, officeTwoStopWeight, usesOfficeTwoStopTransfer } from './gradient-transfer';

const channel = (color: string, offset: number) => Number.parseInt(color.slice(offset, offset + 2), 16);
function sample(stops: readonly { position: number; color: string }[], t: number, offset: number) {
  const index = stops.findIndex(stop => stop.position >= t);
  if (index <= 0) return channel(stops[0].color, offset);
  const low = stops[index - 1]; const high = stops[index];
  const a = (t - low.position) / (high.position - low.position);
  return channel(low.color, offset) + a * (channel(high.color, offset) - channel(low.color, offset));
}

describe('Office two-stop gradient transfer', () => {
  it('applies only to exactly two stops at 0% and 100%', () => {
    expect(usesOfficeTwoStopTransfer([{ position: 1, color: 'FFFFFF' }, { position: 0, color: '000000' }])).toBe(true);
    expect(usesOfficeTwoStopTransfer([{ position: .2, color: '000000' }, { position: .8, color: 'FFFFFF' }])).toBe(false);
    const three = [{ position: 0, color: '000000' }, { position: .5, color: '808080' }, { position: 1, color: 'FFFFFF' }];
    expect(officeGradientStops(three)).toBe(three);
  });

  it('uses the normalized Φ((t-½)/¼) sigma weight', () => {
    // Φ(-1) = .158655, Φ(-2) = .022750, Φ(2) = .977250.
    expect(officeTwoStopWeight(.25)).toBeCloseTo((.158655 - .02275) / .9545, 4);
    expect(officeTwoStopWeight(.5)).toBeCloseTo(.5, 6);
    expect(officeTwoStopWeight(0)).toBe(0);
    expect(officeTwoStopWeight(1)).toBeCloseTo(1, 12);
  });

  it('reproduces the PowerPoint-exported black/white and red/green functions', () => {
    // Values of PowerPoint's 256-entry PDF function at t = k/8 (issue #1599).
    const office = [0, 62, 104, 147, 186, 217, 238, 250, 255];
    const redGreen = [[255, 0], [250, 62], [238, 104], [217, 147], [187, 186], [148, 217], [106, 238], [63, 250], [0, 255]];
    const bw = officeGradientStops([{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }]);
    const rg = officeGradientStops([{ position: 0, color: 'FF0000' }, { position: 1, color: '00FF00' }]);
    expect(bw.length).toBeLessThan(48);
    office.forEach((value, k) => expect(Math.abs(sample(bw, k / 8, 0) - value)).toBeLessThanOrEqual(1.5));
    redGreen.forEach(([red, green], k) => {
      expect(Math.abs(sample(rg, k / 8, 0) - red)).toBeLessThanOrEqual(1.5);
      expect(Math.abs(sample(rg, k / 8, 2) - green)).toBeLessThanOrEqual(1.5);
    });
  });

  it('interpolates alpha with the transferred weight, without gamma', () => {
    const stops = officeGradientStops([{ position: 0, color: '00000000' }, { position: 1, color: '000000FF' }]);
    expect(Math.abs(sample(stops, .5, 6) - 127.5)).toBeLessThanOrEqual(1);
    expect(Math.abs(sample(stops, .25, 6) - 255 * officeTwoStopWeight(.25))).toBeLessThanOrEqual(1);
  });
});
