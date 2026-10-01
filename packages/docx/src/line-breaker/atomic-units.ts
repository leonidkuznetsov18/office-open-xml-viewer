import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { LayoutSeg, LayoutTextSeg } from './model.js';
import { hardJoinPrefixEnd } from './advance.js';
import { hasCJKBreakOpportunity } from './text-runs.js';

export interface AtomicTextMeasurement {
  readonly baseRtl: boolean;
  readonly segAdvance: (segment: LayoutTextSeg) => number;
  readonly strAdvance: (segment: LayoutTextSeg, text: string) => number;
}

/** §17.3.2.14 links contiguous fitText runs into one specified width.
 * Internal separators and formatting seams cannot shorten that allocation.
 * Admission and placement consume this same complete resolved region. */
export function measureFitTextUnit(
  first: LayoutTextSeg, following: readonly LayoutSeg[], measurement: AtomicTextMeasurement, startIndex = 0,
): number {
  let width = measurement.segAdvance(first);
  for (let index = startIndex; index < following.length; index += 1) {
    const member = following[index];
    if (!('text' in member) || member.fitTextRegionIndex !== first.fitTextRegionIndex) break;
    width += measurement.segAdvance(member);
  }
  return width;
}

/** The source-seam group used by both gap admission and ordinary placement.
 * A hard no-break seam consumes its protected prefix; external links retain
 * their registered syntax opportunities; CJK followers glue only non-starters.
 * UAX #14 LB7 retains a space suffix across non-textual formatting seams.
 * A following fitText region is admitted separately when its head is reached;
 * at an empty line head, a preceding word can finish before that region moves
 * as its own cell. Mid-line lookahead still avoids splitting joined words.
 * No field/link/run direction grants extra atomic ownership by itself. */
export function measureJoinedTextUnit(
  s: LayoutTextSeg, following: readonly LayoutSeg[], measurement: AtomicTextMeasurement,
  w = measurement.segAdvance(s), trailingSpaceW = 0, startIndex = 0, atLineStart = false,
): Readonly<{ width: number; trailingSpace: number; next: LayoutSeg | undefined }> {
  const { segAdvance, strAdvance } = measurement;
  let groupW = w;
  let groupTrail = s.fitTextRegionIndex === undefined ? trailingSpaceW : 0;
  let groupEnd = startIndex;
  for (; groupEnd < following.length; groupEnd += 1) {
    const f = following[groupEnd];
    if (!('text' in f)) break;
    if (!f.joinPrev || atLineStart && f.fitTextRegionStart) break;
    const fixedCell = f.ruby !== undefined || f.tateChuYoko === true;
    const hardEnd = hardJoinPrefixEnd(f);
    const externalEnd = f.externalLinkBreakOffsets?.[0];
    let end = fixedCell ? f.text.length : hardEnd ?? externalEnd ?? f.text.length;
    if (!fixedCell && hardEnd === undefined && externalEnd === undefined && hasCJKBreakOpportunity(f.text)) {
      // A CJK follower glues its leading non-starters, then resumes ordinary
      // inter-character breaks. The Latin leader never owns the whole CJK run.
      end = 0;
      for (const character of f.text) {
        if (!DEFAULT_KINSOKU_RULES.lineStartForbidden.has(character.codePointAt(0)!)) break;
        end += character.length;
      }
    }
    const prefix = f.text.slice(0, end);
    const fw = end === f.text.length ? segAdvance(f) : strAdvance(f, prefix);
    groupW += fw;
    const trimmed = prefix.replace(/ +$/, '');
    const trailing = !fixedCell && !measurement.baseRtl && prefix.endsWith(' ')
      ? fw - strAdvance(f, trimmed) : 0;
    groupTrail = trimmed.length === 0 && groupTrail > 0 ? groupTrail + trailing : trailing;
    if (end < f.text.length || !fixedCell && externalEnd !== undefined && hardEnd === undefined) break;
  }
  return { width: groupW, trailingSpace: groupTrail, next: following[groupEnd] };
}
