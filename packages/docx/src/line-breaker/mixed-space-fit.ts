/**
 * Line-breaker projection of WORD_COMPRESSED_SPACE_LINE_FIT (issue #1660).
 *
 * The rule applies only to mixed East Asian / Latin lines in compatibility
 * modes below 15 with a compressing characterSpacingControl. It keeps its own
 * per-line state so Latin-only lines retain WORD_LATIN_INTERWORD_XAVG_FLOOR
 * unchanged: a line enters this projection only once it (or its candidate)
 * holds East Asian text, which that Latin projection never compresses.
 *
 * Every U+0020 of the line shrinks by one amount down to
 * min(xAvg / 2, size / 4); text without U+0020 keeps its natural advance in
 * any face or script and is neutral. The reduction is decided per candidate,
 * re-settled once when the line is finalized, and written to each space-bearing
 * segment exactly once.
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
  /** Segments carrying the line's shrinkable U+0020, in line order. */
  gaps: LayoutTextSeg[];
  /** Total U+0020 across `gaps`. */
  spaceCount: number;
  /** Shrinkable advance of one space (natural minus floor), shared by all gaps. */
  perSpaceCapacity: number | undefined;
  /** First gap; every later gap must share its face, size and floor. */
  face: LayoutTextSeg | undefined;
  /** False once the line holds content the projection was not observed with. */
  valid: boolean;
  /** The line holds East Asian text. */
  eastAsian: boolean;
  /** Reduction of each U+0020 currently subtracted from the line width. */
  appliedPerSpace: number;
}

export function createMixedSpaceState(): MixedSpaceState {
  return {
    gaps: [],
    spaceCount: 0,
    perSpaceCapacity: undefined,
    face: undefined,
    valid: true,
    eastAsian: false,
    appliedPerSpace: 0,
  };
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
function neutral(segment: LayoutSeg): boolean {
  return (
    'text' in segment &&
    !segment.text.includes(' ') &&
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

/** Record a segment just committed to the current line. */
export function performTrackMixedSpaces(
  operationState: PassOperationState,
  segment: LayoutSeg,
): void {
  const { breakerState, scale } = operationState;
  const state = breakerState.mixedSpace;
  if (!state.valid || inkless(segment)) return;
  if ('text' in segment && EAST_ASIAN_RE.test(segment.text)) state.eastAsian = true;
  if ('text' in segment && segment.mixedNaturalTrailingSpacePx !== undefined) {
    const ratio = segment.mixedSpaceAverageWidthRatio!;
    const count = segment.mixedNaturalTrailingSpaceCount ?? 1;
    const capacity = Math.max(
      0,
      segment.mixedNaturalTrailingSpacePx / count -
        wordCompressedSpaceFloor(calcEffectiveFontPx(segment, scale), ratio),
    );
    if (
      (state.face && !sameSpaceFace(segment, state.face)) ||
      (state.perSpaceCapacity !== undefined && Math.abs(capacity - state.perSpaceCapacity) > 1e-6)
    ) {
      state.valid = false;
      return;
    }
    state.face ??= segment;
    state.perSpaceCapacity ??= capacity;
    state.gaps.push(segment);
    state.spaceCount += count;
    return;
  }
  if (!neutral(segment)) state.valid = false;
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

/** Total U+0020 reduction needed to append `next` (fit width `nextFitWidth`),
 * or undefined when the candidate belongs on the next line. Pure. */
export function performMixedSpaceRequirement(
  operationState: PassOperationState,
  next: LayoutTextSeg,
  nextFitWidth: number,
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
  const state = breakerState.mixedSpace;
  if (
    !state.valid ||
    state.spaceCount === 0 ||
    !state.face ||
    baseRtl ||
    widthPolicy !== 'bounded' ||
    characterGrid?.type === 'snapToChars' ||
    characterGrid?.type === 'linesAndChars' ||
    !next.fontRoute
  )
    return undefined;
  // Only mixed East Asian / Latin lines; Latin-only lines keep their own rule.
  if (!state.eastAsian && !EAST_ASIAN_RE.test(next.text)) return undefined;
  if (next.text.includes(' ')) {
    if (next.mixedSpaceAverageWidthRatio === undefined || !sameSpaceFace(next, state.face)) {
      return undefined;
    }
    if (next.text.replace(/ +$/u, '').includes(' ')) return undefined;
  } else if (!neutral(next)) {
    return undefined;
  }
  const capacity = (state.perSpaceCapacity ?? 0) * state.spaceCount;
  if (!(capacity > 0)) return undefined;
  const naturalWidth = breakerState.currentWidth + state.appliedPerSpace * state.spaceCount;
  const required = Math.max(0, naturalWidth + nextFitWidth - availW());
  if (required > capacity || !fitsMeasuredWidth(naturalWidth + nextFitWidth - required, availW())) {
    return undefined;
  }
  // An East Asian final character may overflow by at most half its font size,
  // measured naturally and without trailing closing punctuation, which keeps
  // its own line-end compression.
  let core = next.text.replace(/ +$/u, '');
  while (core.length > 0 && COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION.has(core.at(-1)!)) {
    core = core.slice(0, -1);
  }
  const last = core.at(-1) ?? lastVisibleLineCharacter(breakerState.currentLine);
  const limit = wordCompressedSpaceEastAsianOverflowLimit(
    last !== undefined && EAST_ASIAN_RE.test(last),
    calcEffectiveFontPx(next, scale),
  );
  if (limit !== undefined) {
    const coreWidth = core.length > 0 ? strNaturalAdvance(next, core) : 0;
    if (naturalWidth + coreWidth - availW() > limit + 1e-9) return undefined;
  }
  return required;
}

/** Commit a requirement from performMixedSpaceRequirement. */
export function performApplyMixedSpaces(operationState: PassOperationState, required: number): void {
  const { breakerState } = operationState;
  const state = breakerState.mixedSpace;
  breakerState.currentWidth += state.appliedPerSpace * state.spaceCount - required;
  state.appliedPerSpace = required / state.spaceCount;
}

/** Finalize the line: a later kinsoku retraction may need less reduction than
 * the last admission applied, so the retained line is settled to exactly what
 * it still needs, then each space-bearing segment is written once. */
export function performSettleMixedSpaces(operationState: PassOperationState): void {
  const { breakerState, availW } = operationState;
  const state = breakerState.mixedSpace;
  if (state.appliedPerSpace > 0 && state.spaceCount > 0) {
    const naturalWidth = breakerState.currentWidth + state.appliedPerSpace * state.spaceCount;
    let perSpace = state.appliedPerSpace;
    if (state.valid) {
      const needed = Math.max(0, naturalWidth - availW());
      if (needed <= (state.perSpaceCapacity ?? 0) * state.spaceCount) {
        perSpace = needed / state.spaceCount;
      }
    }
    for (const gap of state.gaps) {
      const reduction = perSpace * (gap.mixedNaturalTrailingSpaceCount ?? 1);
      gap.measuredWidth -= reduction;
      gap.latinSpaceCompressionPx = reduction;
    }
    breakerState.currentWidth = naturalWidth - perSpace * state.spaceCount;
  }
  breakerState.mixedSpace = createMixedSpaceState();
}
