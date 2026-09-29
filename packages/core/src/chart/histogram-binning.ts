import type { ChartexHistogramBinning } from '../types/chart';
import { formatChartValWithCode } from './chart-number-format.js';

/** Matches the shared parser's maximum retained ChartEx cache width. */
export const MAX_HISTOGRAM_INPUT_POINTS = 1_048_576;

/**
 * Histogram output is a Canvas primitive plan, not a lossless data cache.
 * Keep it comfortably below the general 10,000-mark Canvas ceiling so an
 * authored microscopic bin size cannot expand a compact source into a large
 * synchronous paint.
 */
export const MAX_HISTOGRAM_BINS = 512;

export type HistogramBinPlan =
  | { kind: 'bins'; categories: string[]; counts: number[] }
  | { kind: 'tooManyInputPoints' };

function finiteOrNull(value: number | null | undefined): number | null {
  return value != null && Number.isFinite(value) ? value : null;
}

/** Remove binary accumulation noise (3.9000000000000004 -> 3.9). */
function cleanEdge(value: number): number {
  if (!Number.isFinite(value)) return value;
  if (value === 0) return 0;
  return Number(value.toPrecision(12));
}

function makeBoundaryLabel(formatCode: string | null | undefined): (value: number) => string {
  const general = !formatCode || formatCode.trim().toLowerCase() === 'general';
  return (value) => {
    const v = cleanEdge(value);
    // General: shortest round-trip decimal. Otherwise the value dimension's
    // authored number format, as PowerPoint applies it to bin edges.
    return general ? String(v) : formatChartValWithCode(v, formatCode);
  };
}

/** Snap a quotient to the nearest integer when it only differs by float noise. */
function snap(q: number): number {
  const r = Math.round(q);
  return Math.abs(q - r) <= 1e-9 * Math.max(1, Math.abs(q)) ? r : q;
}

function collectFinite(source: readonly (number | null | undefined)[]): number[] {
  const out: number[] = [];
  for (const value of source) {
    if (value != null && Number.isFinite(value)) out.push(value);
  }
  return out;
}

/** Office's automatic width: 3.5 * sample sigma / n^(1/3), 2 significant digits
 *  (measured against PowerPoint 16.113). sigma and n span every finite point;
 *  n = 1 or sigma = 0 gives 5. */
function automaticWidth(values: readonly number[]): number {
  const n = values.length;
  if (n <= 1) return 5;
  let mean = 0;
  let m2 = 0;
  let k = 0;
  for (const v of values) {
    k++;
    const delta = v - mean;
    mean += delta / k;
    m2 += delta * (v - mean);
  }
  const sigma = Math.sqrt(m2 / (n - 1));
  if (!(sigma > 0)) return 5;
  const width = Number(((3.5 * sigma) / Math.cbrt(n)).toPrecision(2));
  return width > 0 ? width : 5;
}

/**
 * Aggregate raw ChartEx histogram observations into a bounded bar plan.
 *
 * MS-ODRAWXML defines authored bin size/count and interval boundaries but not
 * automatic bin selection. Omission follows the rule measured from PowerPoint:
 * width = 3.5 * sample sigma / n^(1/3) rounded to 2 significant digits, first
 * edge = minimum (or underflow), count = ceil(range / width). Authored plans
 * beyond the Canvas bound are coarsened over the same domain instead of
 * allocating an unbounded counts array.
 *
 * Labels use Office's text form: `[a, b]` for the first bin, `(a, b]` after it
 * (r-closed); `[a, b)` with a closed last bin (l-closed); `≤ u` / `> o` for
 * r-closed underflow/overflow. The l-closed forms `< u` / `≥ o` are unmeasured
 * and kept symmetric.
 */
export function planHistogramBins(
  source: readonly (number | null | undefined)[],
  options: ChartexHistogramBinning,
  formatCode?: string | null,
): HistogramBinPlan {
  if (source.length > MAX_HISTOGRAM_INPUT_POINTS) {
    return { kind: 'tooManyInputPoints' };
  }
  const boundaryLabel = makeBoundaryLabel(formatCode);

  const intervalClosed = options.intervalClosed === 'r' ? 'r' : 'l';
  let underflow = finiteOrNull(options.underflow);
  let overflow = finiteOrNull(options.overflow);
  if (underflow != null && overflow != null && underflow >= overflow) {
    underflow = null;
    overflow = null;
  }

  const isUnderflow = (value: number): boolean => underflow != null
    && (intervalClosed === 'r' ? value <= underflow : value < underflow);
  const isOverflow = (value: number): boolean => overflow != null
    && (intervalClosed === 'r' ? value > overflow : value >= overflow);

  let regularMin = Number.POSITIVE_INFINITY;
  let regularMax = Number.NEGATIVE_INFINITY;
  let regularCount = 0;
  let underflowCount = 0;
  let overflowCount = 0;
  for (const sourceValue of source) {
    const value = finiteOrNull(sourceValue);
    if (value == null) continue;
    if (isUnderflow(value)) {
      underflowCount++;
    } else if (isOverflow(value)) {
      overflowCount++;
    } else {
      regularCount++;
      regularMin = Math.min(regularMin, value);
      regularMax = Math.max(regularMax, value);
    }
  }
  if (regularCount + underflowCount + overflowCount === 0) {
    return { kind: 'bins', categories: [], counts: [] };
  }

  const categories: string[] = [];
  const counts: number[] = [];
  if (underflow != null) {
    categories.push(`${intervalClosed === 'r' ? '≤' : '<'} ${boundaryLabel(underflow)}`);
    counts.push(underflowCount);
  }

  if (regularCount > 0) {
    const lower = underflow ?? regularMin;
    const upper = overflow ?? regularMax;
    const range = upper - lower;
    if (!Number.isFinite(range)) {
      categories.push(`[${boundaryLabel(lower)}, ${boundaryLabel(upper)}]`);
      counts.push(regularCount);
    } else {
      const authoredSize = finiteOrNull(options.binSize);
      const authoredCount = options.binCount != null && Number.isFinite(options.binCount) && options.binCount > 0
        ? Math.max(1, Math.floor(options.binCount))
        : null;
      const sizeMode = authoredSize != null && authoredSize > 0;
      // Width-driven planning: authored size, or the automatic width.
      const widthDriven = sizeMode || authoredCount == null;
      let nominalWidth = 0;
      let requestedCount: number;
      if (widthDriven) {
        nominalWidth = sizeMode ? authoredSize : automaticWidth(collectFinite(source));
        requestedCount = range > 0 ? Math.max(1, Math.ceil(snap(range / nominalWidth))) : 1;
        if (!Number.isFinite(requestedCount)) requestedCount = 1;
      } else {
        requestedCount = authoredCount;
      }
      const binCount = range <= 0 ? 1 : Math.min(MAX_HISTOGRAM_BINS, requestedCount);
      const usesWidth = widthDriven && requestedCount <= MAX_HISTOGRAM_BINS
        && Number.isFinite(nominalWidth) && nominalWidth > 0;
      const width = usesWidth ? nominalWidth : range === 0 ? 1 : range / binCount;
      const regularCounts = new Array<number>(binCount).fill(0);
      for (const sourceValue of source) {
        const value = finiteOrNull(sourceValue);
        if (value == null) continue;
        if (isUnderflow(value) || isOverflow(value)) continue;
        // Normalize by the whole range for count-driven/coarsened bins.
        // Dividing a subnormal range by binCount can underflow `width` to zero
        // even though both observations are finite (0..Number.MIN_VALUE).
        const position = snap(usesWidth
          ? (value - lower) / width
          : range === 0
            ? 0
            : ((value - lower) / range) * binCount);
        const rawIndex = intervalClosed === 'r' ? Math.ceil(position) - 1 : Math.floor(position);
        // A value exactly on the top edge stays in the last bin.
        const index = Math.max(0, Math.min(binCount - 1, rawIndex));
        regularCounts[index]++;
      }
      const edge = (k: number): number => {
        const raw = usesWidth ? lower + width * k : lower + (range * k) / binCount;
        return overflow == null ? raw : Math.min(raw, overflow);
      };
      for (let index = 0; index < binCount; index++) {
        const start = boundaryLabel(edge(index));
        const end = boundaryLabel(edge(index + 1));
        const last = index === binCount - 1;
        const leftOpen = intervalClosed === 'r' && (underflow != null || index > 0);
        const rightOpen = intervalClosed === 'l' && !(last && overflow == null);
        categories.push(`${leftOpen ? '(' : '['}${start}, ${end}${rightOpen ? ')' : ']'}`);
        counts.push(regularCounts[index]);
      }
    }
  }

  if (overflow != null) {
    categories.push(`${intervalClosed === 'r' ? '>' : '≥'} ${boundaryLabel(overflow)}`);
    counts.push(overflowCount);
  }
  return { kind: 'bins', categories, counts };
}
