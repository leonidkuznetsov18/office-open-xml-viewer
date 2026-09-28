import { describe, expect, it } from 'vitest';
import { partitionFootnote } from './footnote-fragmentation.js';
import type { NoteLayout, ParagraphLayout } from './types.js';

const source = { story: 'footnote' as const, storyInstance: '7', path: [0] };

function note(pageIndex: number): NoteLayout {
  const bounds = { xPt: 0, yPt: 6, widthPt: 180, heightPt: 40 };
  const paragraph = {
    kind: 'paragraph', id: `footnote:7:page:${pageIndex}:paragraph`, source,
    flowDomainId: `notes:page:${pageIndex}:footnote:7`, ordinaryFlow: true,
    flowBounds: bounds, inkBounds: bounds, advancePt: 40,
    spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
    lines: Array.from({ length: 4 }, (_, index) => ({
      range: { start: index, end: index + 1 },
      bounds: { xPt: 0, yPt: 6 + index * 10, widthPt: 20, heightPt: 10 },
      baselinePt: 14 + index * 10, advancePt: 10,
      placements: [{
        kind: 'text' as const, range: { start: index, end: index + 1 },
        origin: { xPt: 0, yPt: 14 + index * 10 },
        bounds: { xPt: 0, yPt: 6 + index * 10, widthPt: 20, heightPt: 10 },
        advancePt: 20, decorations: [],
      }],
    })),
  } as unknown as ParagraphLayout;
  return {
    kind: 'note', id: `footnote:7:page:${pageIndex}`, source,
    flowDomainId: `notes:page:${pageIndex}`, ordinaryFlow: true,
    flowBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt: 46 },
    inkBounds: { xPt: 0, yPt: 0, widthPt: 180, heightPt: 46 },
    advancePt: 46, separator: [],
    story: {
      story: 'footnote', blocks: [paragraph], flowBounds: bounds,
      inkBounds: bounds, advancePt: 40, diagnostics: [],
    },
  };
}

describe('page-bottom footnote continuation', () => {
  it('retains each source line exactly once across a page boundary', () => {
    const first = partitionFootnote(note(0), null, 27)!;
    expect(first.nextCursor).toEqual({ blockIndex: 0, lineIndex: 2, inlineExtentPt: 180 });
    expect(first.fragment.story.blocks[0]?.kind).toBe('paragraph');
    expect(first.fragment.advancePt).toBe(26);
    const second = partitionFootnote(note(1), first.nextCursor, 100)!;
    expect(second.nextCursor).toBeNull();
    const ranges = [first.fragment, second.fragment].flatMap((part) =>
      part.story.blocks.flatMap((block) => block.kind === 'paragraph'
        ? block.lines.map((line) => line.range.start) : []));
    expect(ranges).toEqual([0, 1, 2, 3]);
    expect(second.fragment.flowDomainId).toBe('notes:page:1');
  });

  it('relocates a reference if no note line can fit', () => {
    expect(partitionFootnote(note(0), null, 15)).toBeNull();
  });

  it('rejects a changed continuation width before a line cursor can skip text', () => {
    const first = partitionFootnote(note(0), null, 27)!;
    const second = { ...note(1), flowBounds: { ...note(1).flowBounds, widthPt: 160 } };
    expect(() => partitionFootnote(second, first.nextCursor, 100)).toThrow(/different text widths/);
  });

  it('keeps a fitting note whole', () => {
    const whole = partitionFootnote(note(0), null, 46)!;
    expect(whole.nextCursor).toBeNull();
    expect(whole.fragment.advancePt).toBe(46);
  });

  it('retains a trailing empty paragraph when it continues alone', () => {
    const original = note(0);
    const empty = {
      ...original.story.blocks[0],
      id: 'empty', source: { ...source, path: [1] }, lines: [],
      flowBounds: { xPt: 0, yPt: 46, widthPt: 180, heightPt: 11 },
      inkBounds: { xPt: 0, yPt: 46, widthPt: 180, heightPt: 11 },
      advancePt: 11,
    } as ParagraphLayout;
    const story = {
      ...original.story,
      blocks: [original.story.blocks[0]!, empty],
      advancePt: 51,
      flowBounds: { ...original.story.flowBounds, heightPt: 51 },
    };
    const acquired: NoteLayout = {
      ...original, story, advancePt: 57,
      flowBounds: { ...original.flowBounds, heightPt: 57 },
    };
    const first = partitionFootnote(acquired, null, 46)!;
    expect(first.nextCursor).toEqual({ blockIndex: 1, lineIndex: 0, inlineExtentPt: 180 });
    const second = partitionFootnote(acquired, first.nextCursor, 100)!;
    expect(second.nextCursor).toBeNull();
    expect(second.fragment.story.blocks.map((block) => block.source.path)).toEqual([[1]]);
  });
});
