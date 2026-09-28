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
 * Body pagination resolves page-owned anchors by exact-state convergence over
 * whole passes, so the counterfactual "page N with the accepted anchors plus
 * the anchors on L" is only observed in a pass that registered exactly that
 * set on N. A proof records that pass's carry reads up to the end of page N. A
 * later pass applies the deferral only after checking that the same carry
 * reads lead to page N with the accepted anchors registered, so it reproduces
 * exactly the proving pass's page-N context minus the deferred anchors.
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
  /** A proven anchor-line deferral took effect: the line holding `keys`
   * starts a page after `pageIndex`. Checks that do not apply change nothing
   * and are not reads. */
  kind: 'anchor-line-deferral';
  pageIndex: number;
  keys: readonly string[];
}>;

/** A carry read of the proving pass on page N after A's registration. */
type InPageRead = Readonly<{
  kind: 'prescan';
  flowDomainId: string;
  serialized: string;
}> | Readonly<{
  kind: 'page-owned-table';
  key: string;
  floor: number | undefined;
}> | Readonly<{
  kind: 'anchor-line-deferral';
  keysId: string;
}>;

export type AnchorLineDeferralProof = Readonly<{
  /** Occurrence keys of the page-owned drawings anchored on line L. */
  keys: readonly string[];
  /** Page N, on which registering `keys` pushed L to a later page. */
  pageIndex: number;
  /** Serialized reads of the proving pass (shared by its proofs). */
  prefix: readonly string[];
  /** Index in `prefix` of the prescan that registered `keys` on page N. */
  registrationIndex: number;
  /** That prescan without `keys`: what a deferring pass reads there. */
  expectedRegistration: string;
  /** Later carry reads of the proving pass on page N. */
  inPage: readonly InPageRead[];
  identity: string;
}>;

/** Proofs by `anchorLineDeferralKey(anchor, page)`. A proof for one page
 * never replaces another page's proof of the same anchor. */
export type AnchorLineDeferrals = ReadonlyMap<string, AnchorLineDeferralProof>;

export function anchorLineDeferralKey(occurrenceKey: string, pageIndex: number): string {
  return `${pageIndex}\u0000${occurrenceKey}`;
}

export function anchorKeysId(keys: readonly string[]): string {
  return [...keys].sort().join('\n');
}

export function serializeAnchorInput(
  event: PageAnchorInputEvent,
  anchorsIdentity: (anchors: PageStartAnchors) => string,
): string {
  switch (event.kind) {
    case 'prescan':
      return `P|${event.pageIndex}|${event.flowDomainId}|${anchorsIdentity(event.anchors)}`;
    case 'page-owned-table':
      return `T|${event.pageIndex}|${event.key}|${event.floor ?? ''}`;
    case 'anchor-line-deferral':
      return `D|${event.pageIndex}|${anchorKeysId(event.keys)}`;
  }
}

export function createAnchorLineDeferralProof(
  keys: readonly string[],
  pageIndex: number,
  events: readonly PageAnchorInputEvent[],
  serialized: readonly string[],
  registrationIndex: number,
  anchorsIdentity: (anchors: PageStartAnchors) => string,
): AnchorLineDeferralProof {
  const registration = events[registrationIndex];
  if (registration?.kind !== 'prescan' || registration.pageIndex !== pageIndex) {
    throw new Error('Anchor-line deferral proof must start at its page registration');
  }
  const keySet = new Set(keys);
  const expectedRegistration = serializeAnchorInput(Object.freeze({
    ...registration,
    anchors: Object.freeze(registration.anchors.filter((anchor) => !keySet.has(anchor.occurrenceId))),
  }), anchorsIdentity);
  const inPage: InPageRead[] = [];
  for (let index = registrationIndex + 1; index < events.length; index += 1) {
    const event = events[index]!;
    if (event.pageIndex > pageIndex) break;
    if (event.pageIndex < pageIndex) continue;
    if (event.kind === 'prescan') {
      inPage.push(Object.freeze({
        kind: 'prescan', flowDomainId: event.flowDomainId, serialized: serialized[index]!,
      }));
    } else if (event.kind === 'page-owned-table') {
      inPage.push(Object.freeze({ kind: 'page-owned-table', key: event.key, floor: event.floor }));
    } else {
      inPage.push(Object.freeze({ kind: 'anchor-line-deferral', keysId: anchorKeysId(event.keys) }));
    }
  }
  const sortedKeys = Object.freeze([...keys].sort());
  return Object.freeze({
    keys: sortedKeys,
    pageIndex,
    prefix: serialized,
    registrationIndex,
    expectedRegistration,
    inPage: Object.freeze(inPage),
    // State identity for cycle detection. The prefix can hold one read per
    // page, so it enters as a digest; application still compares it exactly.
    identity: `${sortedKeys.join(',')}@${pageIndex}:${registrationIndex}:${
      digest(serialized, registrationIndex)}\u0002${expectedRegistration}\u0002${
      inPage.map((read) => JSON.stringify(read)).join('\u0001')}`,
  });
}

/** Two independent 32-bit FNV-1a lanes over the first `length` reads. */
function digest(values: readonly string[], length: number): string {
  let left = 0x811c9dc5;
  let right = 0x01000193 ^ 0x5bd1e995;
  for (let index = 0; index < length; index += 1) {
    const value = values[index]!;
    for (let offset = 0; offset < value.length; offset += 1) {
      const code = value.charCodeAt(offset);
      left = Math.imul(left ^ code, 0x01000193);
      right = Math.imul(right ^ code, 0x5bd1e995) ^ (right >>> 15);
    }
    left = Math.imul(left ^ 0xff, 0x01000193);
    right = Math.imul(right ^ 0xff, 0x5bd1e995) ^ (right >>> 15);
  }
  return `${(left >>> 0).toString(16)}${(right >>> 0).toString(16)}`;
}

/** What a deferring pass knows about its own reads when it reaches line L. */
export interface AnchorLineDeferralContext {
  /** Serialized reads of this pass so far. */
  readonly reads: readonly string[];
  /** The serialized prescan this pass's plan performs for a flow domain. */
  plannedPrescan(pageIndex: number, flowDomainId: string): string | null;
  tableFloor(key: string): number | undefined;
  /** Deferrals this pass already applied on `pageIndex`. */
  appliedDeferrals(pageIndex: number): ReadonlySet<string>;
}

/** Whether the proof's counterfactual is this pass's own page-N context plus
 * the deferred anchors: identical reads before N's registration, the same
 * registration without them, and the same carry values for every later read
 * the proving pass made on page N. */
export function anchorLineDeferralApplies(
  proof: AnchorLineDeferralProof,
  keys: readonly string[],
  pageIndex: number,
  context: AnchorLineDeferralContext,
): boolean {
  if (proof.pageIndex !== pageIndex || anchorKeysId(proof.keys) !== anchorKeysId(keys)) return false;
  const { reads } = context;
  if (reads.length <= proof.registrationIndex) return false;
  for (let index = 0; index < proof.registrationIndex; index += 1) {
    if (reads[index] !== proof.prefix[index]) return false;
  }
  if (reads[proof.registrationIndex] !== proof.expectedRegistration) return false;
  const applied = context.appliedDeferrals(pageIndex);
  for (const read of proof.inPage) {
    if (read.kind === 'prescan') {
      if (context.plannedPrescan(pageIndex, read.flowDomainId) !== read.serialized) return false;
    } else if (read.kind === 'page-owned-table') {
      if (context.tableFloor(read.key) !== read.floor) return false;
    } else if (!applied.has(read.keysId)) {
      return false;
    }
  }
  return true;
}

export function anchorLineDeferralsIdentity(deferrals: AnchorLineDeferrals): string {
  return [...new Set(deferrals.values())]
    .map((proof) => proof.identity)
    .sort()
    .join('\u0003');
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

