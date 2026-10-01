import type { PageAnchorPrescanInput } from './body-layout-kernel.js';
import type { ParagraphLayout } from './types.js';

/**
 * Anchor-line deferral for page-owned drawings (issue #1615), governed by the
 * compatibility rule `word-page-anchor-line-deferral` (anchor-compatibility.ts).
 *
 * A page-owned drawing (positionH/V relativeFrom page or margin, ECMA-376
 * §20.4.3.4-5) is registered on the page of its anchor, so text laid out
 * earlier on that page wraps around it (§20.4.2.17 wrapSquare, §20.4.2.20
 * wrapTopAndBottom). When that wrap would push the anchor line L off page N,
 * the rule keeps page N laid out with only the anchors accepted before L,
 * ends the page just above L, and places the drawing on the page L reaches.
 *
 * Body pagination decides this test on the admission frontier
 * (`resolvePageOwnedAnchors` in body-paginator.ts): the earliest page whose
 * registrations disagree with the landed anchors, where the pass registered
 * exactly the accepted anchors plus L's. A failed test sets a page floor N+1
 * for L's drawings, which every later pass of the run applies.
 */

type PageStartAnchors = PageAnchorPrescanInput['anchors'];

/** One read of anchor-convergence state by a body pagination pass. Everything
 * else a pass reads is identical in every pass of one convergence run. */
export type PageAnchorInputEvent = Readonly<{
  kind: 'prescan';
  pageIndex: number;
  flowDomainId: string;
  anchors: PageStartAnchors;
}> | Readonly<{
  kind: 'page-owned-table';
  pageIndex: number;
  key: string;
  floor: number | undefined;
}> | Readonly<{
  /** An anchor-line floor took effect: the line holding `keys`
   * starts a page after `pageIndex`. Checks that do not apply change nothing
   * and are not reads. */
  kind: 'anchor-line-deferral';
  pageIndex: number;
  keys: readonly string[];
}>;

export function anchorKeysId(keys: readonly string[]): string {
  return [...keys].sort().join('\n');
}

/** Occurrence keys of the page-owned drawings anchored on each line. */
export function pageOwnedAnchorKeysByLine(layout: ParagraphLayout): readonly (readonly string[])[] {
  if (layout.drawings.length === 0) return layout.lines.map(() => []);
  const drawings = new Map(layout.drawings.map((drawing) => [drawing.id, drawing]));
  return layout.lines.map((line) => line.placements.flatMap((placement) => {
    if (placement.kind !== 'drawing') return [];
    const anchor = drawings.get(placement.drawingId)?.anchorLayer;
    if (!anchor || anchor.horizontalOwnership !== 'page' || anchor.verticalOwnership !== 'page') {
      return [];
    }
    return [anchor.acquisitionOccurrenceId ?? anchor.occurrenceId];
  }));
}
