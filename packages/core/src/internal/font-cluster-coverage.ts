import { canonicalCombiningClass } from './canonical-combining-class.js';

/** Canonical Canvas spelling for both ownership and the shared text adapter.
 * Logical OOXML runs stay unchanged. Normalizing the display units prevents a
 * singleton-only cmap (e.g. U+212B versus U+00C5) from painting a different
 * resource than the canonical one whose metrics we certified. This is library
 * display policy; CSS Fonts §5.4 does not itself require normalized input. */
export function canonicalFontClusterText(cluster: string): string {
  return cluster.normalize('NFC');
}

/** Certify one resource's ownership of a complete grapheme, rather than lending
 * the base's metrics to a font that cannot paint its marks (CSS Fonts §5.3).
 * The caller supplies resource-specific, three-valued cmap coverage. Evaluate
 * the whole canonical form (UAX #15), never the authored spelling first:
 * equivalent inputs must have identical ownership, including canonical
 * singletons, reordered marks, and composition leaving additional marks.
 * CSS Fonts §5.3 specifies complete-sequence/single-composition matching;
 * Chrome also matches multi-scalar NFC, decomposed-only and intermediate
 * compositions (e.g. a supported composed base plus a remaining mark). Verified
 * with Latin base/marks, canonical singletons and Hangul syllables/jamo plus
 * remaining marks, against single-resource advance and ink-bound oracles.
 * This is browser resource attribution, not an inferred Office line-box rule.
 * Start from NFC, then canonically decompose and recompose only where this
 * resource has the composed glyph, respecting canonical blocking. Do not union
 * different resources' glyphs or enumerate exponentially many equivalent forms.
 * Unknown coverage is an ownership barrier in the fallback chain.
 *
 * Default-ignorable modifiers (variation selectors, ZWJ, etc.) need sequence
 * facts from cmap format 14 / shaping tables, which our scalar catalogues do
 * not retain. Their scalar presence or absence cannot establish cluster
 * ownership. Leave them unknown, including when the base is covered: browsers
 * may select a different resource or use system fallback for that sequence.
 * No font-name rule, glyph-result cache, or retained font bytes are involved.
 */
export function fontResourceCoversCluster(
  cluster: string,
  covers: (codePoint: number) => boolean | undefined,
): boolean | undefined {
  if (!cluster || /\p{Default_Ignorable_Code_Point}/u.test(cluster)) return undefined;
  const coverageOf = (text: string): boolean | undefined => {
    let result: boolean | undefined = true;
    for (const ch of text) {
      const coverage = covers(ch.codePointAt(0) as number);
      if (coverage === false) return false;
      if (coverage === undefined) result = undefined;
    }
    return result;
  };
  const canonical = canonicalFontClusterText(cluster);
  const composed = coverageOf(canonical);
  if (composed === true) return true;
  const decomposed = canonical.normalize('NFD');
  if (decomposed === canonical) return composed;
  // Unicode Standard §3.11 / UAX #15 §1.3 canonical composition, additionally
  // requiring glyph availability. A successful composition is not a blocker;
  // class-zero characters (including spacing marks and Hangul jamo) start the
  // next sequence when they could not compose. Pair normalization supplies
  // Unicode's composition exclusions and algorithmic Hangul, without retaining
  // a second composition table. Work and transient storage are linear in NFD.
  const shaped: string[] = [];
  let starter = -1;
  let lastClass = 0;
  let unknownComposition = false;
  for (const ch of decomposed) {
    const cls = canonicalCombiningClass(ch.codePointAt(0) as number);
    if (starter >= 0 && (lastClass === 0 || lastClass < cls)) {
      const candidate = canonicalFontClusterText(shaped[starter] + ch);
      const cp = candidate.codePointAt(0) as number;
      if (candidate === String.fromCodePoint(cp)) {
        const coverage = covers(cp);
        if (coverage === true) {
          shaped[starter] = candidate;
          continue;
        }
        if (coverage === undefined) unknownComposition = true;
      }
    }
    if (cls === 0) starter = shaped.length;
    shaped.push(ch);
    lastClass = cls;
  }
  const supported = coverageOf(shaped.join(''));
  if (supported === true) return true;
  return composed === undefined || unknownComposition ? undefined : supported;
}
