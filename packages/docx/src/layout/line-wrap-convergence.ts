import {
  ExactConvergenceError,
  convergeExactState,
} from './convergence.js';
import { LayoutInvariantError } from './diagnostics.js';

interface LineWrapStateSegment {
  readonly src?: unknown;
  readonly text?: string;
}

export interface LineWrapStateLine {
  readonly physicalLineIndex?: number;
  readonly consumedEnd?: unknown;
  readonly topY?: number;
  readonly xOffset?: number;
  readonly availWidth?: number;
  readonly segments: readonly LineWrapStateSegment[];
}

export class LineWrapNonConvergenceError extends LayoutInvariantError {
  readonly reason: 'cycle' | 'limit';
  readonly states: readonly string[];

  constructor(reason: 'cycle' | 'limit', states: readonly string[]) {
    super(
      'NON_CONVERGENCE',
      reason === 'cycle'
        ? `line wrap measure/resolve cycle did not converge (${states.length} states)`
        : `line wrap measure/resolve pass limit did not converge (${states.length} states)`,
    );
    this.name = 'LineWrapNonConvergenceError';
    this.reason = reason;
    this.states = Object.freeze([...states]);
  }
}

export function cloneSegmentsForLinePass<T extends object>(segments: readonly T[]): T[] {
  return segments.map((segment) => ({ ...segment }));
}

function samePhysicalLine(left: LineWrapStateLine | undefined, right: LineWrapStateLine | undefined): boolean {
  if (!left || !right) return false;
  if (left.physicalLineIndex !== undefined || right.physicalLineIndex !== undefined) {
    return left.physicalLineIndex === right.physicalLineIndex;
  }
  return left.topY !== undefined && left.topY === right.topY;
}

function lineWrapState(
  lines: readonly LineWrapStateLine[],
  probeHeights: readonly number[],
): string {
  let physicalIndex = -1;
  return JSON.stringify(lines.map((line, index) => {
    if (!samePhysicalLine(lines[index - 1], line)) physicalIndex += 1;
    return {
      physicalLineIndex: line.physicalLineIndex,
      end: line.consumedEnd, topY: line.topY, xOffset: line.xOffset,
      availableWidth: line.availWidth, probeHeight: probeHeights[physicalIndex],
      segments: line.segments.map(segment => ({ source: segment.src,
        ...(segment.text === undefined ? {} : { text: segment.text }),
      })),
    };
  }));
}

const MAX_LINE_WRAP_PASSES = 16;

/**
 * An adjacent exact-state repeat is the fixed point; a non-adjacent repeat is
 * a real cycle. The pass budget is a fail-closed resource guard for a
 * deterministic orbit whose geometric state cardinality has no useful small
 * bound; exhaustion never accepts stale line geometry. Probes are indexed by
 * physical lines, never gaps. For fixed metrics and one physical line, all G
 * gaps are discovered in one resolving pass and confirmed in the next (three
 * passes including initial measurement), irrespective of G. Variable-height
 * reflow retains the same 16-state fail-closed budget.
 */
export function convergeLineWrap<TLine extends LineWrapStateLine>(
  measure: (probeHeights: readonly number[] | null) => TLine[],
  lineBoxHeight: (line: TLine) => number,
): TLine[] {
  type Pass = Readonly<{
    lines: TLine[];
    probeHeights: readonly number[];
    state: string;
  }>;
  try {
    return convergeExactState<Pass>({
      step: (previous) => {
        const lines = measure(previous?.probeHeights ?? null);
        const probeHeights = Object.freeze(lines.flatMap((line, index) =>
          samePhysicalLine(line, lines[index + 1])
            ? [] : [lineBoxHeight(line)]));
        return Object.freeze({
          lines,
          probeHeights,
          state: lineWrapState(lines, probeHeights),
        });
      },
      stateOf: (pass) => pass.state,
      limit: MAX_LINE_WRAP_PASSES,
    }).value.lines;
  } catch (error) {
    if (error instanceof ExactConvergenceError) {
      throw new LineWrapNonConvergenceError(error.reason, error.states);
    }
    throw error;
  }
}
