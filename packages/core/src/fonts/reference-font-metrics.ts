import { retainFontSupportFacts, type FontSupportFacts } from '../internal/font-support-facts.js';
import referenceData from './reference-font-metrics-data.json';
import { OPEN_FONT_REFERENCE_PROFILES } from './reference-font-metrics-open.js';

export type ReferenceFontSource = 'office-mac' | 'macos-system' | 'macos-supplemental' | 'published-open-font';
export type ReferenceFontStyle = 'normal' | 'italic';
export interface ReferenceFontMetricProfile {
  readonly source: ReferenceFontSource;
  readonly family: string;
  readonly aliases: readonly string[];
  readonly weight: number;
  readonly style: ReferenceFontStyle;
  readonly unitsPerEm: number;
  /** Signed OS/2 xAvgCharWidth design units, when declared. */
  readonly xAvgCharWidth?: number | null;
  readonly hhea: readonly [ascender: number, descender: number, lineGap: number];
  /** OS/2 [usWinAscent, usWinDescent] design units. Null means the face has no
   * OS/2 table; undefined means this source did not record the field. */
  readonly win?: readonly [ascent: number, descent: number] | null;
  /** OS/2 [sTypoAscender, sTypoDescender, sTypoLineGap], present only when the
   * face sets fsSelection USE_TYPO_METRICS, including older tables that set
   * the bit in practice (same rule as the resource parser). */
  readonly typoMetrics?: readonly [ascender: number, descender: number, lineGap: number];
  /** Derived OS/2 code-page class. Null means this source did not provide the
   * code-page field needed to classify Word's auto-line allocation. */
  readonly farEastCodePage: boolean | null;
  /** OS/2 PANOSE [bFamilyType, bSerifStyle]. Null means the face has no OS/2
   * table; undefined means this source did not record the field. */
  readonly panose?: readonly [familyType: number, serifStyle: number] | null;
  /** True when the face's Unicode cmap maps a CJK Unified Ideograph
   * (U+4E00–U+9FFF). Null includes unreadable/disagreeing maps; undefined means
   * this source did not record coverage. */
  readonly cjkUnifiedIdeographs?: boolean | null;
  /** Unicode cmap presence within the catalogue's bounded symbol domain.
   * Sorted inclusive endpoint pairs, interned across faces. Empty means no
   * mapped symbols; null/undefined means unknown. Not installed-face detection. */
  readonly symbolCoverage?: readonly number[] | null;
  /** Index of this concrete cut's immutable CJK cmap bitmap in the generated
   * catalogue. Family-wide Han classification cannot prove glyph coverage. */
  readonly cjkCoverage?: number | null;
}

export interface FindReferenceFontMetricsOptions {
  readonly source?: ReferenceFontSource;
  readonly weight?: number;
  readonly style?: ReferenceFontStyle;
}

type ParsedReferenceCoverage = Readonly<{ symbol: readonly number[] | null; cjk: number | null }>;
const parsedCoverage = new WeakMap<ReferenceFontMetricProfile, ParsedReferenceCoverage>();
function freezeProfile(profile: ReferenceFontMetricProfile): ReferenceFontMetricProfile {
  Object.freeze(profile.aliases);
  Object.freeze(profile.hhea);
  if (profile.win) Object.freeze(profile.win);
  if (profile.typoMetrics) Object.freeze(profile.typoMetrics);
  if (profile.panose) Object.freeze(profile.panose);
  if (profile.symbolCoverage) Object.freeze(profile.symbolCoverage);
  // Private certificate association survives projection/freezing without
  // widening the public catalogue metric contract. Missing facts stay unknown.
  const support = (profile as ReferenceFontMetricProfile & { supportFacts?: FontSupportFacts }).supportFacts;
  if (support) {
    if (support.erasureSafeRanges) { support.erasureSafeRanges.forEach(Object.freeze); Object.freeze(support.erasureSafeRanges); }
    retainFontSupportFacts(profile, Object.freeze(support));
  }
  const scalar = profile as ReferenceFontMetricProfile & { symbolPossibleCoverage?: readonly number[] | null; cjkPossibleCoverage?: number | null };
  if ('symbolPossibleCoverage' in scalar) {
    if (scalar.symbolPossibleCoverage) Object.freeze(scalar.symbolPossibleCoverage);
    parsedCoverage.set(profile, Object.freeze({ symbol: scalar.symbolPossibleCoverage ?? null, cjk: scalar.cjkPossibleCoverage ?? null }));
  }
  return Object.freeze(profile);
}

// The generator includes every supported static face from the Office and macOS
// catalogs; independently verified, pinned open-font profiles follow them.
let profiles: readonly ReferenceFontMetricProfile[] | undefined;

function getProfiles(): readonly ReferenceFontMetricProfile[] {
  // The JSON payload is statically imported and parsed with the module. Only
  // profile freezing and the alias index are deferred until the first lookup.
  return profiles ??= Object.freeze(
    [...referenceData.profiles.map((profile) => ({
      ...profile,
      symbolPossibleCoverage: profile.symbolPossibleCoverage === null ? null
        : referenceData.symbolCoverages[profile.symbolPossibleCoverage],
      symbolCoverage: profile.symbolCoverage === null ? null
        : referenceData.symbolCoverages[profile.symbolCoverage],
    })) as unknown as ReferenceFontMetricProfile[],
      ...OPEN_FONT_REFERENCE_PROFILES].map(freezeProfile),
  );
}

type OptionBuckets = ReadonlyMap<string, readonly ReferenceFontMetricProfile[]>;
let aliasIndex: ReadonlyMap<string, OptionBuckets> | undefined;
const EMPTY_RESULTS = Object.freeze([]) as readonly ReferenceFontMetricProfile[];

function normalizeFamilyName(value: string): string {
  return value.normalize('NFKC').trim().replace(/\s+/gu, ' ').toLocaleLowerCase('en-US');
}

function optionKey(options: FindReferenceFontMetricsOptions): string {
  return `${options.source ?? '*'}|${options.weight ?? '*'}|${options.style ?? '*'}`;
}

function getAliasIndex(): ReadonlyMap<string, OptionBuckets> {
  if (aliasIndex !== undefined) return aliasIndex;

  const mutable = new Map<string, Map<string, ReferenceFontMetricProfile[]>>();
  for (const profile of getProfiles()) {
    const keys = new Set(profile.aliases.map(normalizeFamilyName));
    const optionKeys = new Set<string>();
    for (const source of [undefined, profile.source] as const) {
      for (const weight of [undefined, profile.weight] as const) {
        for (const style of [undefined, profile.style] as const) {
          optionKeys.add(optionKey({ source, weight, style }));
        }
      }
    }
    for (const alias of keys) {
      let buckets = mutable.get(alias);
      if (buckets === undefined) {
        buckets = new Map();
        mutable.set(alias, buckets);
      }
      for (const key of optionKeys) {
        const matches = buckets.get(key);
        if (matches === undefined) buckets.set(key, [profile]);
        else matches.push(profile);
      }
    }
  }
  for (const buckets of mutable.values()) {
    for (const matches of buckets.values()) Object.freeze(matches);
  }
  // This index is bounded solely by aliases in the immutable generated catalog.
  // Caller-provided queries are normalized once and are never memoized.
  aliasIndex = mutable;
  return aliasIndex;
}

/**
 * Finds metadata-only reference profiles without claiming which font Canvas or
 * the host selected. Multiple results intentionally preserve same-name metric
 * conflicts across and within source catalogs.
 */
export function findReferenceFontMetrics(
  familyOrAlias: string,
  options: FindReferenceFontMetricsOptions = {},
): readonly ReferenceFontMetricProfile[] {
  const query = normalizeFamilyName(familyOrAlias);
  if (!query) return EMPTY_RESULTS;
  return getAliasIndex().get(query)?.get(optionKey(options)) ?? EMPTY_RESULTS;
}

/** A cmap fact, not proof of the drawing resource. Outside the recorded
 * #1653/#1689 symbol sweep domain, or for a source without coverage, return
 * undefined rather than confuse an unrecorded scalar with a missing glyph. */
export function referenceFontCoversSymbol(
  profile: ReferenceFontMetricProfile,
  codePoint: number,
): boolean | undefined {
  if (!isReferenceSymbolCodePoint(codePoint) || profile.symbolCoverage == null) return undefined;
  if (coverageIncludes(profile.symbolCoverage, codePoint)) return true;
  const parsed = parsedCoverage.get(profile);
  return !parsed ? false : parsed.symbol === null || coverageIncludes(parsed.symbol, codePoint) ? undefined : false;
}

/** Per-cut CJK cmap coverage; outside the generated domain stays unknown. */
export function referenceFontCoversCjk(
  profile: ReferenceFontMetricProfile,
  codePoint: number,
): boolean | undefined {
  if (!Number.isInteger(codePoint) || profile.cjkCoverage == null
    || !referenceData.cjkCoverageRanges.some(([lo, hi]) => codePoint >= lo && codePoint <= hi)) return undefined;
  const bits = cjkBitmap(profile.cjkCoverage);
  if (!bits) return undefined;
  let offset = 0;
  for (const [lo, hi] of referenceData.cjkCoverageRanges) {
    if (codePoint >= lo && codePoint <= hi) {
      const index = offset + codePoint - lo;
      if ((bits[index >>> 3] & (1 << (index & 7))) !== 0) return true;
      const parsed = parsedCoverage.get(profile);
      if (!parsed) return false;
      const possible = parsed.cjk === null ? undefined : cjkBitmap(parsed.cjk);
      return !possible || (possible[index >>> 3] & (1 << (index & 7))) !== 0 ? undefined : false;
    }
    offset += hi - lo + 1;
  }
  return undefined;
}

// Decode only queried repertoires. The index is confined to the immutable
// catalogue (one bitmap per generated repertoire); family/glyph queries never
// create cache keys. Each decoded bitmap has the fixed scalar-domain size,
// including its implicit trailing zeros. Neither bytes nor bitmaps escape.
const cjkBitmaps = new Map<number, Uint8Array>();
const CJK_BITMAP_BYTES = Math.ceil(referenceData.cjkCoverageRanges
  .reduce((sum, [lo, hi]) => sum + hi - lo + 1, 0) / 8);
function cjkBitmap(index: number): Uint8Array | undefined {
  if (!Number.isInteger(index) || index < 0 || index >= referenceData.cjkCoverages.length) return undefined;
  const cached = cjkBitmaps.get(index);
  if (cached) return cached;
  const packed = atob(referenceData.cjkCoverages[index]);
  const bits = new Uint8Array(CJK_BITMAP_BYTES);
  let out = 0;
  // PackBits: 0..127 copies n+1 literal bytes; 129..255 repeats the next
  // byte 257-n times; 128 is a no-op. Encoding limits, not font heuristics.
  for (let at = 0; at < packed.length;) {
    const control = packed.charCodeAt(at++);
    if (control === 128) continue;
    const count = control < 128 ? control + 1 : 257 - control;
    if (out + count > bits.length) return undefined;
    if (control < 128) {
      if (at + count > packed.length) return undefined;
      for (let end = at + count; at < end;) bits[out++] = packed.charCodeAt(at++);
    } else {
      if (at >= packed.length) return undefined;
      bits.fill(packed.charCodeAt(at++), out, out + count);
      out += count;
    }
  }
  cjkBitmaps.set(index, bits);
  return bits;
}

function coverageIncludes(ranges: readonly number[], codePoint: number): boolean {
  let lo = 0;
  let hi = ranges.length / 2 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (codePoint < ranges[mid * 2]) hi = mid - 1;
    else if (codePoint > ranges[mid * 2 + 1]) lo = mid + 1;
    else return true;
  }
  return false;
}

/** Whether a scalar lies in the cmap domain recorded by the generator. */
export function isReferenceSymbolCodePoint(codePoint: number): boolean {
  return Number.isInteger(codePoint)
    && referenceData.symbolCoverageRanges.some(([lo, hi]) => codePoint >= lo && codePoint <= hi);
}
