/**
 * Line-breaker projection of WORD_COMPRESSED_SPACE_LINE_FIT (issue #1660).
 *
 * The rule applies only to mixed East Asian / Latin lines in compatibility
 * modes below 15 with a compressing characterSpacingControl. Latin-only lines
 * keep WORD_LATIN_INTERWORD_XAVG_FLOOR unchanged: a line enters this
 * projection only once it (or its candidate) holds East Asian text, which that
 * Latin projection never compresses.
 *
 * Every decision is a pure function of the candidate line's content: while a
 * line is built `currentWidth` stays the natural advance of its committed
 * items, and the shrinkable spaces (count, natural advance and floor) are
 * re-derived from those items on every query. The only state is whether the
 * line was admitted by shrinking. The reduction is materialized once, when the
 * line is finalized, onto each space-bearing segment.
 */
import { calcEffectiveFontPx, EAST_ASIAN_RE } from '../layout/text.js';
import {
  wordCompressedSpaceEastAsianOverflowLimit,
  wordCompressedSpaceFloor,
} from '../layout/line-compatibility.js';
import { COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION } from './text-runs.js';
import type { LayoutSeg, LayoutTextSeg } from './model.js';
import type { PassOperationState } from './pass-operations.js';

export interface MixedSpaceState {
  /** The current line was admitted by shrinking its U+0020; its natural
   * advance may exceed the band until the line is finalized. */
  compressed: boolean;
}

export function createMixedSpaceState(): MixedSpaceState {
  return { compressed: false };
}

/** One candidate unit: consecutive pieces of text that must share a line
 * (a segment, a prefix of one, or a joined group across source runs). */
export interface MixedSpaceCandidate {
  readonly pieces: readonly Readonly<{ segment: LayoutTextSeg; text: string }>[];
  /** Fit width of the whole unit, as used by the ordinary natural fit test. */
  readonly fitWidth: number;
}

interface LineSpaces {
  readonly valid: boolean;
  readonly eastAsian: boolean;
  readonly face: LayoutTextSeg | undefined;
  /** Shrinkable advance of one space (natural minus floor). */
  readonly perSpaceCapacity: number;
  readonly gaps: readonly Readonly<{ segment: LayoutTextSeg; count: number }>[];
  readonly count: number;
}

function sameSpaceFace(candidate: LayoutTextSeg, reference: LayoutTextSeg): boolean {
  return (
    candidate.mixedSpaceAverageWidthRatio === reference.mixedSpaceAverageWidthRatio &&
    candidate.fontRoute?.fingerprint === reference.fontRoute?.fingerprint &&
    candidate.fontFamily === reference.fontFamily &&
    candidate.fontSize === reference.fontSize &&
    candidate.bold === reference.bold &&
    candidate.italic === reference.italic
  );
}

/** Text without U+0020 that keeps its natural advance on a mixed line. */
function neutral(segment: LayoutTextSeg, text: string): boolean {
  return (
    !text.includes(' ') &&
    !segment.verticalRun &&
    !segment.tateChuYoko &&
    !segment.rtl &&
    segment.fitTextRegionIndex === undefined &&
    segment.widthBalanceGridDeltaFactor === undefined
  );
}

function inkless(segment: LayoutSeg): boolean {
  return ('text' in segment && segment.metricOnly === true)
    || ('imagePath' in segment && Boolean(segment.anchor));
}

function trailingSpaceCount(text: string): number {
  return text.length - text.replace(/ +$/u, '').length;
}

function perSpaceCapacity(segment: LayoutTextSeg, count: number, scale: number): number {
  return Math.max(
    0,
    segment.mixedNaturalTrailingSpacePx! / count -
      wordCompressedSpaceFloor(calcEffectiveFontPx(segment, scale), segment.mixedSpaceAverageWidthRatio!),
  );
}

/** Re-derive the shrinkable spaces of the committed line items. */
function scanLine(line: readonly LayoutSeg[], scale: number): LineSpaces {
  let eastAsian = false;
  let face: LayoutTextSeg | undefined;
  let capacity = 0;
  let count = 0;
  const gaps: { segment: LayoutTextSeg; count: number }[] = [];
  const invalid = (): LineSpaces => ({
    valid: false, eastAsian, face, perSpaceCapacity: 0, gaps: [], count: 0,
  });
  for (const item of line) {
    if (inkless(item)) continue;
    if (!('text' in item)) return invalid();
    if (EAST_ASIAN_RE.test(item.text)) eastAsian = true;
    const spaces = trailingSpaceCount(item.text);
    if (
      item.mixedNaturalTrailingSpacePx !== undefined &&
      spaces > 0 &&
      spaces === item.mixedNaturalTrailingSpaceCount &&
      !item.text.slice(0, item.text.length - spaces).includes(' ')
    ) {
      const itemCapacity = perSpaceCapacity(item, spaces, scale);
      if (face && (!sameSpaceFace(item, face) || Math.abs(itemCapacity - capacity) > 1e-6)) {
        return invalid();
      }
      face ??= item;
      capacity = itemCapacity;
      gaps.push({ segment: item, count: spaces });
      count += spaces;
    } else if (!neutral(item, item.text)) {
      return invalid();
    }
  }
  return { valid: true, eastAsian, face, perSpaceCapacity: capacity, gaps, count };
}

/** Total U+0020 reduction the line needs to append `candidate`, or undefined
 * when the candidate belongs on the next line. Pure. */
export function performMixedSpaceRequirement(
  operationState: PassOperationState,
  candidate: MixedSpaceCandidate,
): number | undefined {
  const {
    breakerState,
    availW,
    fitsMeasuredWidth,
    characterGrid,
    baseRtl,
    widthPolicy,
    strNaturalAdvance,
    scale,
  } = operationState;
  const { pieces } = candidate;
  if (
    pieces.length === 0 ||
    baseRtl ||
    widthPolicy !== 'bounded' ||
    characterGrid?.type === 'snapToChars' ||
    characterGrid?.type === 'linesAndChars' ||
    breakerState.currentLine.length === 0
  )
    return undefined;
  const line = scanLine(breakerState.currentLine, scale);
  if (!line.valid || line.count === 0 || !line.face) return undefined;
  // Only mixed East Asian / Latin lines; Latin-only lines keep their own rule.
  if (!line.eastAsian && !pieces.some((piece) => EAST_ASIAN_RE.test(piece.text))) return undefined;
  for (const [index, piece] of pieces.entries()) {
    const spaces = trailingSpaceCount(piece.text);
    const visible = piece.text.slice(0, piece.text.length - spaces);
    if (!piece.segment.fontRoute || !neutral(piece.segment, visible)) return undefined;
    // Trailing spaces become gaps of the line; they must share its face.
    if (
      spaces > 0 &&
      (index !== pieces.length - 1 ||
        piece.segment.mixedSpaceAverageWidthRatio === undefined ||
        !sameSpaceFace(piece.segment, line.face))
    ) {
      return undefined;
    }
  }
  const capacity = line.perSpaceCapacity * line.count;
  if (!(capacity > 0)) return undefined;
  const naturalWidth = breakerState.currentWidth;
  const required = Math.max(0, naturalWidth + candidate.fitWidth - availW());
  if (required > capacity || !fitsMeasuredWidth(naturalWidth + candidate.fitWidth - required, availW())) {
    return undefined;
  }
  // An East Asian final character may overflow by at most half its font size,
  // measured naturally and without trailing closing punctuation, which keeps
  // its own line-end compression.
  const cores = pieces.map((piece) => piece.text.replace(/ +$/u, ''));
  let last = cores.length - 1;
  while (last >= 0) {
    let core = cores[last]!;
    while (core.length > 0 && COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(core.at(-1)!)) {
      core = core.slice(0, -1);
    }
    cores[last] = core;
    if (core.length > 0) break;
    last -= 1;
  }
  const lastCharacter = last >= 0
    ? cores[last]!.at(-1)
    : lastVisibleLineCharacter(breakerState.currentLine);
  const limitSegment = last >= 0 ? pieces[last]!.segment : pieces[0]!.segment;
  const limit = wordCompressedSpaceEastAsianOverflowLimit(
    lastCharacter !== undefined && EAST_ASIAN_RE.test(lastCharacter),
    calcEffectiveFontPx(limitSegment, scale),
  );
  if (limit !== undefined) {
    let coreWidth = 0;
    for (let index = 0; index <= last; index += 1) {
      if (cores[index]!.length > 0) coreWidth += strNaturalAdvance(pieces[index]!.segment, cores[index]!);
    }
    if (naturalWidth + coreWidth - availW() > limit + 1e-9) return undefined;
  }
  return required;
}

function lastVisibleLineCharacter(line: readonly LayoutSeg[]): string | undefined {
  for (let index = line.length - 1; index >= 0; index -= 1) {
    const item = line[index]!;
    if (inkless(item)) continue;
    if (!('text' in item)) return undefined;
    const visible = item.text.replace(/ +$/u, '');
    if (visible.length > 0) return visible.at(-1);
  }
  return undefined;
}

/** Record that the line was admitted by shrinking its spaces. */
export function performMarkMixedSpacesCompressed(operationState: PassOperationState): void {
  operationState.breakerState.mixedSpace.compressed = true;
}

/** Finalize the line: shrink its spaces by exactly what its committed natural
 * advance still exceeds the band (a kinsoku retraction may have shortened it),
 * write each space-bearing segment once, and reset for the next line. */
export function performSettleMixedSpaces(operationState: PassOperationState): void {
  const { breakerState, availW, scale } = operationState;
  if (breakerState.mixedSpace.compressed) {
    const line = scanLine(breakerState.currentLine, scale);
    // Spaces ending the line are not gaps: like every fit decision, the line
    // end is judged without them (they keep their natural advance).
    const gaps = [...line.gaps];
    let lineEndSpacePx = 0;
    for (let index = breakerState.currentLine.length - 1; index >= 0; index -= 1) {
      const item = breakerState.currentLine[index]!;
      if (inkless(item)) continue;
      const gap = gaps.at(-1);
      if (gap?.segment !== item) break;
      gaps.pop();
      lineEndSpacePx += gap.segment.mixedNaturalTrailingSpacePx!;
      if (gap.segment.text.trim().length > 0) break;
    }
    const count = gaps.reduce((sum, gap) => sum + gap.count, 0);
    const needed = Math.max(0, breakerState.currentWidth - lineEndSpacePx - availW());
    const reduction = line.valid && count > 0 ? Math.min(needed, line.perSpaceCapacity * count) : 0;
    if (reduction > 0) {
      const perSpace = reduction / count;
      for (const gap of gaps) {
        const gapReduction = perSpace * gap.count;
        gap.segment.measuredWidth -= gapReduction;
        gap.segment.latinSpaceCompressionPx = gapReduction;
      }
      breakerState.currentWidth -= reduction;
    }
  }
  breakerState.mixedSpace = createMixedSpaceState();
}
