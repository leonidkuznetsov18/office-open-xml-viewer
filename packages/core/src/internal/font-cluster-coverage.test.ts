import { expect, it } from 'vitest';
import { fontResourceCoversCluster } from './font-cluster-coverage.js';
import { canonicalClusterPairs } from '../test-fixtures/canonical-clusters.js';

it('requires all marks or a supported canonical equivalent from the same resource', () => {
  const covered = (points: number[]) => (cp: number) => points.includes(cp);
  expect(fontResourceCoversCluster('A\u0301', covered([0x41]))).toBe(false);
  expect(fontResourceCoversCluster('A\u0301', covered([0x41, 0x301]))).toBe(true);
  expect(fontResourceCoversCluster('A\u0301', covered([0xc1]))).toBe(true);
  expect(fontResourceCoversCluster('\u1100\u1161\u11a8', covered([0xac01]))).toBe(true);
  expect(fontResourceCoversCluster('A\u0301', (cp) => cp === 0x41 ? true : undefined)).toBeUndefined();
});

it('certifies the same resource for generated canonical equivalents with remaining marks', () => {
  for (const [a, b] of canonicalClusterPairs()) {
    const points = new Set([...a.normalize('NFC')].map((ch) => ch.codePointAt(0)));
    const covers = (cp: number) => points.has(cp);
    expect([fontResourceCoversCluster(a, covers), fontResourceCoversCluster(b, covers)],
      [...a].map((ch) => ch.codePointAt(0)?.toString(16)).join(' ')).toEqual([true, true]);
  }
});

it('handles reordered marks, canonical singletons and decomposed-only resources consistently', () => {
  for (const [a, b] of [['A\u0301\u0323', 'A\u0323\u0301'], ['\u212b', 'Å'], ['\u2126', 'Ω']]) {
    for (const form of ['NFC', 'NFD'] as const) {
      const points = new Set([...a.normalize(form)].map((ch) => ch.codePointAt(0)));
      expect([fontResourceCoversCluster(a, (cp) => points.has(cp)),
        fontResourceCoversCluster(b, (cp) => points.has(cp))]).toEqual([true, true]);
    }
  }
  // Compatibility decomposition does not confer canonical ownership.
  expect(fontResourceCoversCluster('ﬀ', (cp) => cp === 0x66)).toBe(false);
});

it('accepts resource-supported intermediate compositions and respects canonical blocking', () => {
  const covered = new Set([0xc5, 0x301, 0x307]); // Å + acute + dot; neither complete NFC nor NFD.
  for (const text of ['A\u030a\u0301\u0307', 'Å\u0301\u0307', 'Ǻ\u0307']) {
    expect(fontResourceCoversCluster(text, (cp) => covered.has(cp))).toBe(true);
  }
  // Acute and ring have the same class; acute blocks A+ring composition.
  expect(fontResourceCoversCluster('A\u0301\u030a', (cp) => covered.has(cp))).toBe(false);
  // A lower-class retained mark allows composition across it.
  expect(fontResourceCoversCluster('A\u0323\u030a', (cp) => cp === 0xc5 || cp === 0x323)).toBe(true);
  // Bengali class-zero spacing marks compose with one another, not the base.
  expect(fontResourceCoversCluster('ক\u09c7\u09be', (cp) => cp === 0x995 || cp === 0x9cb)).toBe(true);
});

it('keeps sequence-only ownership unknown even if every scalar has a cmap entry', () => {
  for (const cluster of ['§\ufe0e', '§\ufe0f', '👩\u200d💻']) {
    expect(fontResourceCoversCluster(cluster, () => true)).toBeUndefined();
    expect(fontResourceCoversCluster(cluster, () => false)).toBeUndefined();
  }
});
