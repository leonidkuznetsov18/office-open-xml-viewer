import { expect, it } from 'vitest';
import { fontResourceCoversCluster } from './font-cluster-coverage.js';

it('requires all marks or a supported canonical equivalent from the same resource', () => {
  const covered = (points: number[]) => (cp: number) => points.includes(cp);
  expect(fontResourceCoversCluster('A\u0301', covered([0x41]))).toBe(false);
  expect(fontResourceCoversCluster('A\u0301', covered([0x41, 0x301]))).toBe(true);
  expect(fontResourceCoversCluster('A\u0301', covered([0xc1]))).toBe(true);
  expect(fontResourceCoversCluster('\u1100\u1161\u11a8', covered([0xac01]))).toBe(true);
  expect(fontResourceCoversCluster('A\u0301', (cp) => cp === 0x41 ? true : undefined)).toBeUndefined();
});

it('keeps sequence-only ownership unknown even if every scalar has a cmap entry', () => {
  for (const cluster of ['§\ufe0e', '§\ufe0f', '👩\u200d💻']) {
    expect(fontResourceCoversCluster(cluster, () => true)).toBeUndefined();
    expect(fontResourceCoversCluster(cluster, () => false)).toBeUndefined();
  }
});
