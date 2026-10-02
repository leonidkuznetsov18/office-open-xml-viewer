/** Derive canonical pairs from the runtime's Unicode normalization data, the
 * same Unicode source available to the renderer. No copied decomposition table
 * or font-specific list. Appending another combining mark exercises multi-scalar
 * NFC as well as singleton decompositions and algorithmic Hangul composition. */
export function canonicalClusterPairs(): [string, string][] {
  const pairs: [string, string][] = [];
  for (let cp = 0; cp <= 0x10ffff; cp++) {
    if (cp >= 0xd800 && cp <= 0xdfff) continue;
    const scalar = String.fromCodePoint(cp);
    const decomposed = scalar.normalize('NFD');
    if (scalar !== decomposed) pairs.push([scalar + '\u0307', decomposed + '\u0307']);
  }
  return pairs;
}
