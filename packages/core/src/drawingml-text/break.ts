import { isCjkBreakChar } from '../text/cjk-ranges.js';
import { DEFAULT_KINSOKU_RULES, kinsokuAdjustedSplit } from '../text/kinsoku/index.js';
import { isUax14NoBreakPair, lineBreakClass } from '../text/line-break.js';
import { containsSeaScript, graphemeClusterOffsets, seaMixedBreakOffsets } from '../text/sea-break.js';
import { resolveDrawingMlTabWidths, type DrawingMlTabStop } from './tab.js';

export type DrawingMlInputRun<T> =
  | { type: 'text'; text: string; style: T }
  | { type: 'break' }
  | { type: 'object'; width: number; style: T; payload?: unknown; display?: boolean };

export type DrawingMlLineSegment<T> =
  | { type: 'text'; text: string; style: T; width: number }
  | { type: 'tab'; style: T; width: number }
  | { type: 'object'; style: T; width: number; payload?: unknown };

export interface DrawingMlBrokenLine<T> {
  segments: DrawingMlLineSegment<T>[];
  width: number;
  /** The line ended at an authored `<a:br>` or LF, not a soft wrap. */
  endsWithBreak?: boolean;
}

export interface DrawingMlBreakOptions<T> {
  /** Inner text width after body insets and paragraph margins, in canvas px. */
  maxWidth: number;
  /** Signed first-line indent; continuation lines use the full `maxWidth`. */
  firstLineIndent?: number;
  /** Resolve font and other run properties in the format adapter. */
  measureText(text: string, style: T): number;
  /** Adjacent authored runs may shape together only when all paint metadata agrees. */
  sameStyle?(left: T, right: T): boolean;
  tabStops?: readonly DrawingMlTabStop[];
  /** `a:pPr@defTabSz` in canvas px; Office's ordinary fallback is 1 inch. */
  defaultTabSize?: number;
  /** Pen position relative to the leading inset; includes paragraph margin. */
  tabStartPen?(lineIndex: number): number;
  /** Use a linear fit for negative tracking or a nonmonotone measurement hook. */
  nonMonotoneMeasure?: boolean;
}

type Atom<T> =
  | { type: 'text'; text: string; style: T; run: number }
  | { type: 'tab'; style: T; run: number }
  | { type: 'object'; style: T; run: number; width: number; payload?: unknown };

const latin = /^\p{Script_Extensions=Latin}$/u;
const isSpace = <T>(atom: Atom<T>): boolean => atom.type === 'text' && atom.text === ' ';
const isCjk = <T>(atom: Atom<T>): boolean =>
  atom.type === 'text' && isCjkBreakChar(atom.text.codePointAt(0) ?? 0);

/**
 * Shared DrawingML soft-wrap phase. ECMA-376 CT_TextBody/CT_TextParagraphProperties
 * defines `wrap`, insets, indent and tab stops, but leaves the choice among
 * Unicode soft opportunities to the host. The Office controls C00–C11 establish
 * the boundaries below at matched geometry in PowerPoint and Excel:
 * authored run seams are not breaks, ordinary spaces separate words, a Latin
 * compound hyphen and CJK/Latin seam are breaks, NBSP is not ordinary space,
 * and an overwide unbreakable word falls back to grapheme boundaries. Both
 * hosts put a tab on its next visual line when its following word will not fit.
 * C08 also shows that a CJK punctuation run seam can remain a break boundary;
 * the kinsoku adjustment therefore stays within an authored run here.
 *
 * The input adapter owns fonts and measurement; this phase never changes a
 * run's style or fabricates a font. The output is logical order. Alignment,
 * bidi, line metrics and paint are later phases.
 */
export function breakDrawingMlText<T>(
  runs: readonly DrawingMlInputRun<T>[],
  options: DrawingMlBreakOptions<T>,
): DrawingMlBrokenLine<T>[] {
  const sameStyle = options.sameStyle ?? ((a: T, b: T) => a === b);
  const regions: Atom<T>[][] = [[]];
  let runIndex = 0;
  for (const run of runs) {
    if (run.type === 'break') {
      regions.push([]);
      runIndex++;
      continue;
    }
    if (run.type === 'object') {
      regions.at(-1)!.push({ type: 'object', style: run.style, run: runIndex, width: run.width, payload: run.payload });
      if (run.display) regions.push([]);
      runIndex++;
      continue;
    }
    const offsets = [0, ...graphemeClusterOffsets(run.text), run.text.length];
    for (let i = 0; i + 1 < offsets.length; i++) {
      const text = run.text.slice(offsets[i], offsets[i + 1]);
      if (text === '\n') regions.push([]);
      else if (text === '\t') regions.at(-1)!.push({ type: 'tab', style: run.style, run: runIndex });
      else if (text !== '\r') regions.at(-1)!.push({ type: 'text', text, style: run.style, run: runIndex });
    }
    runIndex++;
  }

  const lines: DrawingMlBrokenLine<T>[] = [];
  for (let regionIndex = 0; regionIndex < regions.length; regionIndex++) {
    const atoms = regions[regionIndex];
    // Office C05/C06: terminal U+0020 spaces do not create a continuation
    // line, even when the suffix crosses an authored run boundary. NBSP stays.
    while (atoms.length > 0 && isSpace(atoms.at(-1)!)) atoms.pop();

    const end = atoms.length;
    const seaBreaks = new Set<number>();
    if (atoms.some((atom) => atom.type === 'text' && containsSeaScript(atom.text))) {
      let full = '';
      const atomOffsets = [0];
      for (const atom of atoms) {
        full += atom.type === 'text' ? atom.text : '\ufffc';
        atomOffsets.push(full.length);
      }
      const offsets = new Set(seaMixedBreakOffsets(full, { cjk: true }));
      for (let i = 1; i < atomOffsets.length; i++) if (offsets.has(atomOffsets[i])) seaBreaks.add(i);
    }

    const mayBreakAt = (index: number): boolean => {
      if (index <= 0 || index >= end) return false;
      const prev = atoms[index - 1];
      const next = atoms[index];
      // A tab is a break-before opportunity, not break-after: C11 carries
      // the tab to the next line and seats its first glyph at the 1-inch stop.
      if (next.type === 'tab') return true;
      if (prev.type === 'tab') return false;
      if (isSpace(prev) || isSpace(next)) return true;
      if (seaBreaks.has(index)) return true;
      if (prev.type === 'object' || next.type === 'object') return true;
      const prevCp = [...prev.text].at(-1)?.codePointAt(0);
      const nextCp = next.text.codePointAt(0);
      if (prevCp === undefined || nextCp === undefined) return false;
      if (prevCp === 0x200b) return true;
      // Observed PowerPoint table controls T00–T08: NBSP binds the adjacent
      // words. At a narrow width Office moves the whole "and NBSP Partner"
      // group after the preceding ordinary space; once that group fits, it
      // stays on the first line. The matching Arial controls bound the rule
      // independently of Segoe UI's host-only font metrics.
      if (isUax14NoBreakPair(prevCp, nextCp)) return false;
      if (isCjk(prev) || isCjk(next)) return true;
      if (index > 1 && (lineBreakClass(prevCp) === 'HY' || lineBreakClass(prevCp) === 'HH')) {
        const before = atoms[index - 2];
        if (before.type === 'text' && latin.test([...before.text].at(-1) ?? '')
            && latin.test([...next.text][0] ?? '')) return true;
      }
      return false;
    };

    const makeSegments = (start: number, stop: number, lineIndex: number): DrawingMlBrokenLine<T> => {
      const segments: DrawingMlLineSegment<T>[] = [];
      for (let i = start; i < stop; i++) {
        const atom = atoms[i];
        if (atom.type === 'text') {
          const last = segments.at(-1);
          if (last?.type === 'text' && sameStyle(last.style, atom.style)) {
            last.text += atom.text;
            continue;
          }
          segments.push({ type: 'text', text: atom.text, style: atom.style, width: 0 });
        } else if (atom.type === 'tab') {
          segments.push({ type: 'tab', style: atom.style, width: 0 });
        } else {
          segments.push({ type: 'object', style: atom.style, width: atom.width, payload: atom.payload });
        }
      }
      let noStopGap = 0;
      const items = segments.map((seg) => {
        if (seg.type === 'text') seg.width = options.measureText(seg.text, seg.style);
        if (seg.type === 'tab') noStopGap = options.measureText(' ', seg.style);
        return { isTab: seg.type === 'tab', width: seg.width };
      });
      if (items.some((item) => item.isTab)) {
        const widths = resolveDrawingMlTabWidths(
          items,
          options.tabStops ?? [],
          options.tabStartPen?.(lineIndex) ?? (lineIndex === 0 ? options.firstLineIndent ?? 0 : 0),
          Infinity,
          noStopGap,
          options.defaultTabSize ?? 0,
        );
        for (let i = 0; i < segments.length; i++) segments[i].width = widths[i];
      }
      return { segments, width: segments.reduce((sum, seg) => sum + seg.width, 0) };
    };

    let start = 0;
    if (end === 0) {
      lines.push({ segments: [], width: 0, ...(regionIndex + 1 < regions.length ? { endsWithBreak: true } : {}) });
      continue;
    }
    while (start < end) {
      const lineIndex = lines.length;
      const budget = options.maxWidth - (lineIndex === 0 ? options.firstLineIndent ?? 0 : 0);
      const widthAt = (stop: number): number => makeSegments(start, stop, lineIndex).width;
      let fit = end;
      if (Number.isFinite(budget) && widthAt(end) > budget) {
        if (options.nonMonotoneMeasure) {
          fit = start;
          for (let i = start + 1; i <= end; i++) {
            if (widthAt(i) <= budget) fit = i;
          }
        } else {
          let lo = start;
          let hi = end;
          while (lo < hi) {
            const mid = Math.ceil((lo + hi) / 2);
            if (widthAt(mid) <= budget) lo = mid;
            else hi = mid - 1;
          }
          fit = lo;
        }
      }
      if (fit >= end) {
        const line = makeSegments(start, end, lineIndex);
        if (regionIndex + 1 < regions.length) line.endsWithBreak = true;
        lines.push(line);
        break;
      }

      let split = 0;
      for (let i = start + 1; i <= fit; i++) if (mayBreakAt(i)) split = i;
      if (split === 0) split = Math.max(start + 1, fit); // overwide word: grapheme-safe emergency break

      // Kinsoku (§17.15.1.58–.60) adjusts an in-run CJK boundary. The Office
      // C08 control is a counterexample at an authored run seam, so leave that
      // boundary intact instead of inferring a cross-run retraction rule.
      if (split < end && atoms[split - 1].run === atoms[split].run) {
        const left = atoms.slice(start, split).flatMap((atom) => atom.type === 'text' ? [...atom.text] : []);
        const right = atoms.slice(split).flatMap((atom) => atom.type === 'text' ? [...atom.text] : []);
        if (left.length > 1 && right.length > 0) {
          const adjusted = kinsokuAdjustedSplit([...left, ...right], left.length, DEFAULT_KINSOKU_RULES, 1);
          const retract = left.length - adjusted;
          if (retract > 0 && split - retract > start) split -= retract;
        }
      }

      let visibleEnd = split;
      while (visibleEnd > start && isSpace(atoms[visibleEnd - 1])) visibleEnd--;
      if (visibleEnd === start) visibleEnd = split; // authored leading space
      lines.push(makeSegments(start, visibleEnd, lineIndex));
      start = split;
      while (start < end && isSpace(atoms[start])) start++;
    }
  }
  return lines;
}
