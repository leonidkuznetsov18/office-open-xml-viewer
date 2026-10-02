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
 * items, and the line facts (shrinkable space count and floor, East Asian
 * presence, unsupported content, visible-text tail) come from a reversible
 * running summary that mirrors `currentLine` item for item. The summary is
 * updated O(1) per commit through the same helpers that mutate the line, and
 * tests assert it equals a full recomputation. The reduction is materialized
 * once, when the line is finalized, onto each space-bearing segment.
 */
import { calcEffectiveFontPx, EAST_ASIAN_RE } from '../layout/text.js';
import {
  wordCompressedSpaceEastAsianOverflowLimit,
  wordCompressedSpaceFloor,
} from '../layout/line-compatibility.js';
import { COMPRESSIBLE_TRAILING_FULL_WIDTH_PUNCTUATION } from './text-runs.js';
import type { LayoutSeg, LayoutTextSeg } from './model.js';
import type { PassOperationState } from './pass-operations.js';

/** One committed line item's contribution to the running summary. */
interface Contribution {
  readonly item: LayoutSeg;
  readonly inkless: boolean;
  /** Last non-space character of a visible text item. */
  readonly lastVisible: string | undefined;
  /** The item is text holding a non-whitespace character. */
  readonly visibleText: boolean;
  readonly eastAsian: boolean;
  /** Content this projection was not observed with (ends it for the line). */
  readonly invalid: boolean;
  /** Shrinkable trailing U+0020 carried by the item (0 when none). */
  readonly gapCount: number;
  readonly capacity: number;
  /** The gap differs in face or floor from the line's first gap. */
  readonly mismatched: boolean;
}

/**
 * Running summary of the committed line, O(1) per commit. It is maintained
 * only through commitMixedLineItem / popMixedLineItem / replaceLastMixedLineItem,
 * which also perform the matching `currentLine` mutation, and is reset with
 * the line. `contributions` mirrors `currentLine` item for item; the totals
 * equal a full recomputation (scanLine) by construction, which tests assert.
 */
export interface MixedSpaceState {
  /** The current line was admitted by shrinking its U+0020; its natural
   * advance may exceed the band until the line is finalized. */
  compressed: boolean;
  contributions: Contribution[];
  eastAsianItems: number;
  invalidItems: number;
  mismatchedGaps: number;
  spaceCount: number;
  /** Index into `contributions` of the line's first gap, or -1. */
  firstGap: number;
  /** Indices of contributions with visible text, in line order. */
  visibleIndices: number[];
  /** Indices of text contributions with a character other than U+0020. */
  nonSpaceIndices: number[];
  /** Indices of non-inkless, non-text contributions. */
  lastNonText: number[];
}

export function createMixedSpaceState(): MixedSpaceState {
  return {
    compressed: false,
    contributions: [],
    eastAsianItems: 0,
    invalidItems: 0,
    mismatchedGaps: 0,
    spaceCount: 0,
    firstGap: -1,
    visibleIndices: [],
    nonSpaceIndices: [],
    lastNonText: [],
  };
}

let summaryAssertions = false;
let summaryWork = 0;

/** Test hook: compare the running summary with a full recomputation on every
 * query and at finalization, throwing on any difference. */
export function setMixedSpaceSummaryAssertions(enabled: boolean): void {
  summaryAssertions = enabled;
}

/** Test hook: summary items read since the last reset (work, not time). */
export function mixedSpaceSummaryWork(reset = false): number {
  const value = summaryWork;
  if (reset) summaryWork = 0;
  return value;
}

function contributionOf(state: MixedSpaceState, item: LayoutSeg, scale: number): Contribution {
  summaryWork += 1;
  if (inkless(item)) {
    return {
      item, inkless: true, lastVisible: undefined,
      // The visible-content predicate counts every text item, as before.
      visibleText: 'text' in item && /\S/u.test(item.text), eastAsian: false,
      invalid: false, gapCount: 0, capacity: 0, mismatched: false,
    };
  }
  if (!('text' in item)) {
    return {
      item, inkless: false, lastVisible: undefined, visibleText: false, eastAsian: false,
      invalid: true, gapCount: 0, capacity: 0, mismatched: false,
    };
  }
  const eastAsian = EAST_ASIAN_RE.test(item.text);
  const visible = item.text.replace(/ +$/u, '');
  const lastVisible = visible.length > 0 ? visible.at(-1) : undefined;
  const visibleText = /\S/u.test(item.text);
  const spaces = trailingSpaceCount(item.text);
  if (
    item.mixedNaturalTrailingSpacePx !== undefined &&
    spaces > 0 &&
    spaces === item.mixedNaturalTrailingSpaceCount &&
    !item.text.slice(0, item.text.length - spaces).includes(' ')
  ) {
    const capacity = perSpaceCapacity(item, spaces, scale);
    const first = state.firstGap >= 0 ? state.contributions[state.firstGap]! : undefined;
    const mismatched = first !== undefined && (
      !sameSpaceFace(item, first.item as LayoutTextSeg) || Math.abs(capacity - first.capacity) > 1e-6
    );
    return {
      item, inkless: false, lastVisible, visibleText, eastAsian, invalid: false,
      gapCount: spaces, capacity, mismatched,
    };
  }
  return {
    item, inkless: false, lastVisible, visibleText, eastAsian, invalid: !neutral(item, item.text),
    gapCount: 0, capacity: 0, mismatched: false,
  };
}

function addContribution(state: MixedSpaceState, contribution: Contribution): void {
  if (contribution.gapCount > 0 && state.firstGap < 0) state.firstGap = state.contributions.length;
  const index = state.contributions.length;
  state.contributions.push(contribution);
  if (contribution.visibleText) state.visibleIndices.push(index);
  if (contribution.lastVisible !== undefined) state.nonSpaceIndices.push(index);
  if (!contribution.inkless && !('text' in contribution.item)) state.lastNonText.push(index);
  if (contribution.eastAsian) state.eastAsianItems += 1;
  if (contribution.invalid) state.invalidItems += 1;
  if (contribution.mismatched) state.mismatchedGaps += 1;
  state.spaceCount += contribution.gapCount;
}

function removeContribution(state: MixedSpaceState): void {
  const contribution = state.contributions.pop();
  if (!contribution) return;
  const index = state.contributions.length;
  if (state.visibleIndices.at(-1) === index) state.visibleIndices.pop();
  if (state.nonSpaceIndices.at(-1) === index) state.nonSpaceIndices.pop();
  if (state.lastNonText.at(-1) === index) state.lastNonText.pop();
  if (contribution.eastAsian) state.eastAsianItems -= 1;
  if (contribution.invalid) state.invalidItems -= 1;
  if (contribution.mismatched) state.mismatchedGaps -= 1;
  state.spaceCount -= contribution.gapCount;
  if (state.firstGap === state.contributions.length) state.firstGap = -1;
}

type BreakerState = PassOperationState['breakerState'];

/** Append a committed item to the line and its summary. */
type LineItem = BreakerState['currentLine'][number];

export function commitMixedLineItem(breakerState: BreakerState, item: LineItem, scale: number): void {
  breakerState.currentLine.push(item);
  addContribution(breakerState.mixedSpace, contributionOf(breakerState.mixedSpace, item, scale));
}

/** Remove the last committed item from the line and its summary. */
export function popMixedLineItem(breakerState: BreakerState): void {
  breakerState.currentLine.pop();
  removeContribution(breakerState.mixedSpace);
}

/** Replace the last committed item (a kinsoku-retracted head) in both. */
export function replaceLastMixedLineItem(
  breakerState: BreakerState,
  item: LineItem,
  scale: number,
): void {
  popMixedLineItem(breakerState);
  commitMixedLineItem(breakerState, item, scale);
}

/** The committed line holds visible (non-whitespace) text; O(1). */
export function mixedLineHasVisibleText(breakerState: BreakerState): boolean {
  summaryWork += 1;
  return breakerState.mixedSpace.visibleIndices.length > 0;
}

/** Line facts from the running summary; O(1). */
function summaryLine(state: MixedSpaceState): Omit<LineSpaces, 'gaps'> {
  const first = state.firstGap >= 0 ? state.contributions[state.firstGap]! : undefined;
  const valid = state.invalidItems === 0 && state.mismatchedGaps === 0;
  return {
    valid,
    eastAsian: state.eastAsianItems > 0,
    face: valid ? first?.item as LayoutTextSeg | undefined : undefined,
    perSpaceCapacity: valid ? first?.capacity ?? 0 : 0,
    count: valid ? state.spaceCount : 0,
  };
}

function assertSummary(breakerState: BreakerState, scale: number): void {
  if (!summaryAssertions) return;
  const state = breakerState.mixedSpace;
  const full = scanLine(breakerState.currentLine, scale);
  const summary = summaryLine(state);
  const line = breakerState.currentLine;
  const indices = (predicate: (item: LayoutSeg) => boolean) =>
    line.flatMap((item, index) => (predicate(item) ? [index] : []));
  const sameIndices = (left: readonly number[], right: readonly number[]) =>
    left.length === right.length && left.every((value, index) => value === right[index]);
  const same = state.contributions.length === breakerState.currentLine.length
    && sameIndices(state.visibleIndices, indices((item) => 'text' in item && /\S/u.test(item.text)))
    && sameIndices(state.nonSpaceIndices, indices((item) => !inkless(item) && 'text' in item
      && item.text.replace(/ +$/u, '').length > 0))
    && sameIndices(state.lastNonText, indices((item) => !inkless(item) && !('text' in item)))
    && state.contributions.every((contribution, index) => contribution.item === breakerState.currentLine[index])
    && full.valid === summary.valid
    && (!full.valid || (
      full.eastAsian === summary.eastAsian
      && full.face === summary.face
      && full.count === summary.count
      && Math.abs(full.perSpaceCapacity - summary.perSpaceCapacity) <= 1e-9
    ));
  if (!same) {
    throw new Error('WORD_COMPRESSED_SPACE_LINE_FIT line summary drifted from its committed items');
  }
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
    summaryWork += 1;
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
  assertSummary(breakerState, scale);
  const line = summaryLine(breakerState.mixedSpace);
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
    : lastVisibleLineCharacter(breakerState.mixedSpace);
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

function lastVisibleLineCharacter(state: MixedSpaceState): string | undefined {
  // The last visible character of the committed line, unless a non-text item
  // (tab, object) follows it. Text items holding only spaces are skipped.
  summaryWork += 1;
  const visible = state.nonSpaceIndices.at(-1);
  const nonText = state.lastNonText.at(-1);
  if (visible === undefined || (nonText !== undefined && nonText > visible)) return undefined;
  return state.contributions[visible]!.lastVisible;
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
    assertSummary(breakerState, scale);
    const line = summaryLine(breakerState.mixedSpace);
    // Finalization walks the line once; every gap is written exactly once.
    const gaps = breakerState.mixedSpace.contributions.flatMap((contribution) => {
      summaryWork += 1;
      return contribution.gapCount > 0
        ? [{ segment: contribution.item as LayoutTextSeg, count: contribution.gapCount }]
        : [];
    });
    // Spaces ending the line are not gaps: like every fit decision, the line
    // end is judged without them (they keep their natural advance).
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
