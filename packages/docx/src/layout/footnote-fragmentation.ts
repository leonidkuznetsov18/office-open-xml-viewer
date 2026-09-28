import { sliceParagraphLayout } from './paragraph.js';
import { translateCompleteParagraphLayout } from './retained-geometry-translation.js';
import type { NoteLayout, PaintNode, StoryLayout } from './types.js';

export interface FootnoteCursor {
  readonly blockIndex: number;
  readonly lineIndex: number;
  readonly inlineExtentPt: number;
}

/**
 * ECMA-376 §17.11.1 illustrates a footnote continued across pages, and
 * §17.11.21 locates its band at the page bottom. Word keeps the reference on
 * its body page and admits as many complete
 * note lines as fit below that page's body content. The following page owns a
 * fresh note separator and resumes at the first unpainted line. A note whose
 * first line cannot fit is left to ordinary reference relocation.
 */
export function partitionFootnote(
  acquired: NoteLayout,
  cursor: FootnoteCursor | null,
  capacityPt: number,
): Readonly<{ fragment: NoteLayout; nextCursor: FootnoteCursor | null }> | null {
  if (!Number.isFinite(capacityPt) || capacityPt < 0) throw new RangeError('Invalid footnote capacity');
  // A cursor counts shaped lines, so continuing it in a different text width
  // could silently skip or repeat source text. Reflow-aware source offsets are
  // required before that case can be supported.
  if (cursor && Math.abs(cursor.inlineExtentPt - acquired.flowBounds.widthPt) > 1e-6) {
    throw new Error('Footnote continuation across different text widths is unsupported');
  }
  if (cursor === null && acquired.advancePt <= capacityPt) {
    return Object.freeze({ fragment: acquired, nextCursor: null });
  }
  // Table rows require their own continuation cursor. Relocate the reference
  // until that capability exists; never retain a clipped or page-stale table.
  if (acquired.story.blocks.some((block) => block.kind !== 'paragraph')) return null;
  const separatorPt = acquired.advancePt - acquired.story.advancePt;
  let usedPt = separatorPt;
  if (usedPt >= capacityPt) return null;
  const blocks: PaintNode[] = [];
  const startBlock = cursor?.blockIndex ?? 0;
  let nextCursor: FootnoteCursor | null = null;
  for (let blockIndex = startBlock; blockIndex < acquired.story.blocks.length; blockIndex += 1) {
    const block = acquired.story.blocks[blockIndex]!;
    if (block.kind !== 'paragraph') {
      // Table note continuation needs row-aware partitioning; never clip it.
      if (usedPt + block.advancePt > capacityPt) {
        nextCursor = Object.freeze({ blockIndex, lineIndex: 0, inlineExtentPt: acquired.flowBounds.widthPt });
        break;
      }
      blocks.push(block);
      usedPt += block.advancePt;
      continue;
    }
    const lineStart = blockIndex === startBlock ? (cursor?.lineIndex ?? 0) : 0;
    if (block.lines.length === 0) {
      if (usedPt + block.advancePt > capacityPt) {
        nextCursor = Object.freeze({ blockIndex, lineIndex: 0, inlineExtentPt: acquired.flowBounds.widthPt });
        break;
      }
      blocks.push(translateCompleteParagraphLayout(block, {
        xPt: 0,
        yPt: acquired.story.flowBounds.yPt + usedPt - separatorPt - block.flowBounds.yPt,
      }));
      usedPt += block.advancePt;
      continue;
    }
    let admitted = 0;
    let admittedBlock = null as ReturnType<typeof sliceParagraphLayout> | null;
    for (let end = lineStart + 1; end <= block.lines.length; end += 1) {
      const candidate = sliceParagraphLayout(block, {
        lineStart,
        lineEnd: end,
        continuesFromPrevious: lineStart > 0,
        continuesOnNext: end < block.lines.length,
      });
      if (usedPt + candidate.advancePt > capacityPt + 1e-6) break;
      admitted = end;
      admittedBlock = candidate;
    }
    if (admittedBlock === null) {
      nextCursor = Object.freeze({ blockIndex, lineIndex: lineStart, inlineExtentPt: acquired.flowBounds.widthPt });
      break;
    }
    blocks.push(translateCompleteParagraphLayout(admittedBlock, {
      xPt: 0,
      yPt: acquired.story.flowBounds.yPt + usedPt - separatorPt - admittedBlock.flowBounds.yPt,
    }));
    usedPt += admittedBlock.advancePt;
    if (admitted < block.lines.length) {
      nextCursor = Object.freeze({ blockIndex, lineIndex: admitted, inlineExtentPt: acquired.flowBounds.widthPt });
      break;
    }
  }
  if (blocks.length === 0) return null;
  const storyBounds = Object.freeze({ ...acquired.story.flowBounds, heightPt: usedPt - separatorPt });
  const story: StoryLayout = Object.freeze({
    ...acquired.story,
    blocks: Object.freeze(blocks),
    flowBounds: storyBounds,
    inkBounds: storyBounds,
    advancePt: usedPt - separatorPt,
  });
  const noteBounds = Object.freeze({ ...acquired.flowBounds, heightPt: usedPt });
  return Object.freeze({
    fragment: Object.freeze({
      ...acquired,
      story,
      flowBounds: noteBounds,
      inkBounds: noteBounds,
      advancePt: usedPt,
    }),
    nextCursor,
  });
}
