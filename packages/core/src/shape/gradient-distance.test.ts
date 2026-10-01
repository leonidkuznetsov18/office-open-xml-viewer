import { expect, it } from 'vitest';
import { interiorDistances } from './gradient-distance';

it('measures diagonal concave boundaries and horizontal/vertical clearance independently', () => {
  const width = 9; const height = 7;
  const rgba = new Uint8ClampedArray(width * height * 4);
  // An L with a concave corner at (4,4), surrounded by the required empty
  // border. The diagonal probe rejects Manhattan/chamfer distance and catches
  // a row/column indexing swap on this non-square raster.
  for (let y = 1; y < height - 1; y++) {
    for (let x = 1; x < width - 1; x++) {
      if (x < 4 || y < 4) rgba[(y * width + x) * 4 + 3] = 255;
    }
  }
  const distance = interiorDistances(rgba, width, height);
  expect(distance[3 * width + 3]).toBeCloseTo(Math.SQRT2 - .5);
  expect(distance[2 * width + 2]).toBe(1.5);
  expect(distance[2 * width + 7]).toBe(.5);
  expect(distance[5 * width + 2]).toBe(.5);
  expect(distance[4 * width + 4]).toBe(0);
});
