/** Certify one resource's ownership of a complete grapheme, rather than lending
 * the base's metrics to a font that cannot paint its marks (CSS Fonts §5.3).
 * The caller supplies resource-specific, three-valued cmap coverage. Canonical
 * single-character equivalents are eligible too; absence needs both forms to
 * be absent. Unknown coverage is an ownership barrier in the fallback chain.
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
  let direct: boolean | undefined = true;
  for (const ch of cluster) {
    const coverage = covers(ch.codePointAt(0) as number);
    if (coverage === false) {
      direct = false;
      break;
    }
    if (coverage === undefined) direct = undefined;
  }
  if (direct === true) return true;
  const canonical = cluster.normalize('NFC');
  const cp = canonical.codePointAt(0) as number;
  if (canonical !== cluster && canonical === String.fromCodePoint(cp)) {
    const composed = covers(cp);
    if (composed === true) return true;
    if (composed === undefined) return undefined;
  }
  return direct;
}
