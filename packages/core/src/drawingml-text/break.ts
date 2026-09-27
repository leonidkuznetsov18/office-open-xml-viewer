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
  /** Advance at a style seam within one authored run (for `rPr@spc`). */
  boundaryAdvance?(left: T, right: T): number;
  tabStops?: readonly DrawingMlTabStop[];
  /** `a:pPr@defTabSz` in canvas px; Office's ordinary fallback is 1 inch. */
  defaultTabSize?: number;
  /** Pen position relative to the leading inset; includes paragraph margin. */
  tabStartPen?(lineIndex: number): number;
  /**
   * Negative tracking can make a longer prefix narrower. Every prefix is then
   * a fit candidate; widths are accumulated from segment heads and
   * per-grapheme advances measured with a bounded left shaping context.
   */
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
      // ECMA-376 §22.1 m:oMathPara is display math. It occupies its own visual
      // line: break after preceding text and before following text. An authored
      // hard break already supplies the first edge.
      if (run.display && regions.at(-1)!.length > 0) regions.push([]);
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
    // Office C05/C06: paragraph-terminal U+0020 spaces do not create a
    // continuation line, even across authored runs. An authored <a:br> is a
    // line ending inside the paragraph: preserve its preceding spaces because
    // they still advance a centred/right-aligned line (observed in an Office
    // PDF with a space before <a:br>).
    // NBSP stays in either case.
    if (regionIndex === regions.length - 1) {
      while (atoms.length > 0 && isSpace(atoms.at(-1)!)) atoms.pop();
    }

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

    // A continuation of an overwide word can begin at every grapheme. Cache
    // these paragraph-local predicates once; searching the whole remaining
    // word on every continuation made a narrow box quadratic in word length.
    const breakAt = new Uint8Array(end + 1);
    const nextBreak = new Int32Array(end + 1);
    const nonTextOrCjk = new Int32Array(end + 1);
    const styleSeams = new Int32Array(end + 1);
    for (let i = 1; i < end; i++) breakAt[i] = mayBreakAt(i) ? 1 : 0;
    nextBreak[end] = end;
    for (let i = end - 1; i >= 0; i--) nextBreak[i] = breakAt[i + 1] ? i + 1 : nextBreak[i + 1];
    for (let i = 0; i < end; i++) {
      const atom = atoms[i];
      const validWordAtom = atom.type === 'text' && !isCjk(atom);
      nonTextOrCjk[i + 1] = nonTextOrCjk[i] + (validWordAtom ? 0 : 1);
      styleSeams[i + 1] = styleSeams[i] + (
        i > 0 && validWordAtom && atoms[i - 1].type === 'text'
          && !isCjk(atoms[i - 1]) && !sameStyle(atoms[i - 1].style, atom.style) ? 1 : 0
      );
    }

    // Most DrawingML paragraphs use one resolved paint style and no tabs or
    // objects. Build UTF-16 offsets once so every fit probe measures exactly
    // the same substring that makeSegments would coalesce, without rebuilding
    // the segment array for each candidate prefix. This also preserves shaping
    // across authored run seams when their resolved styles agree.
    const plainStyle = atoms[0]?.style;
    const plainText = end > 0 && atoms.every((atom) => atom.type === 'text'
      && sameStyle(plainStyle, atom.style))
      ? atoms.map((atom) => atom.type === 'text' ? atom.text : '').join('') : null;
    const plainOffsets = plainText === null ? null : new Int32Array(end + 1);
    if (plainOffsets) {
      for (let i = 0; i < end; i++) {
        const atom = atoms[i];
        plainOffsets[i + 1] = plainOffsets[i] + (atom.type === 'text' ? atom.text.length : 0);
      }
    }

    const isSingleExplicitTabCell = (index: number): boolean => {
      const tab = atoms[index];
      const first = atoms[index + 1];
      return tab?.type === 'tab' && !!options.tabStops?.length
        && first?.type === 'text'
        && /^[\p{Script_Extensions=Latin}\p{Number}]$/u.test(first.text)
        && atoms.slice(index + 1).every((atom) => atom.type === 'text'
          && !isSpace(atom) && !isCjk(atom));
    };

    const appendAtom = (segments: DrawingMlLineSegment<T>[], atom: Atom<T>): void => {
      if (atom.type === 'text') {
        const last = segments.at(-1);
        if (last?.type === 'text' && sameStyle(last.style, atom.style)) {
          last.text += atom.text;
          return;
        }
        segments.push({ type: 'text', text: atom.text, style: atom.style, width: 0 });
      } else if (atom.type === 'tab') {
        segments.push({ type: 'tab', style: atom.style, width: 0 });
      } else {
        segments.push({ type: 'object', style: atom.style, width: atom.width, payload: atom.payload });
      }
    };

    const measureSegments = (segments: DrawingMlLineSegment<T>[], lineIndex: number): number => {
      let noStopGap = 0;
      const items = segments.map((seg, i) => {
        if (seg.type === 'text') {
          seg.width = options.measureText(seg.text, seg.style);
          const previous = segments[i - 1];
          if (previous?.type === 'text') {
            seg.width += options.boundaryAdvance?.(previous.style, seg.style) ?? 0;
          }
        }
        if (seg.type === 'tab') {
          // The non-monotone scan reuses this prefix array after measuring its
          // previous length. A tab's earlier resolved gap is not an input to
          // the next candidate's tab-stop resolution.
          seg.width = 0;
          noStopGap = options.measureText(' ', seg.style);
        }
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
      return segments.reduce((sum, seg) => sum + seg.width, 0);
    };

    const makeSegments = (start: number, stop: number, lineIndex: number): DrawingMlBrokenLine<T> => {
      if (plainText !== null && plainOffsets && plainStyle !== undefined) {
        const text = plainText.slice(plainOffsets[start], plainOffsets[stop]);
        const width = options.measureText(text, plainStyle);
        return { segments: [{ type: 'text', text, style: plainStyle, width }], width };
      }
      const segments: DrawingMlLineSegment<T>[] = [];
      for (let i = start; i < stop; i++) appendAtom(segments, atoms[i]);
      return { segments, width: measureSegments(segments, lineIndex) };
    };

    // Negative tracking (`rPr@spc < 0`) can make a longer prefix narrower, so
    // the fit inspects every candidate prefix and keeps the greatest fitting
    // one. Measuring each candidate afresh made a narrow box cubic in run
    // length. The paragraph is measured once instead: each text segment's
    // first SHAPING_CONTEXT + 1 graphemes as one string, then each later
    // grapheme's advance after the preceding SHAPING_CONTEXT graphemes of the
    // same segment. Tracking is a per-grapheme addend, so the accumulated
    // width equals the measured string's width whenever shaping context
    // (kerning pairs, joining) stays within that window. Segment seams, tabs
    // and objects contribute exactly as measureSegments does.
    const SHAPING_CONTEXT = 8;
    const BLOCK = 32;
    interface NonMonotoneModel {
      /** Start of the coalesced text segment containing each text atom. */
      segStart: Int32Array;
      segEnd: Int32Array;
      /** Segment text width through atom t while t is within its head. */
      head: Float64Array;
      /** Advance of atom t after its SHAPING_CONTEXT predecessors. */
      tail: Float64Array;
      tailSum: Float64Array;
      tailBlockMin: Float64Array;
      /** Line width of atoms [0, i) with segments coalesced from atom 0. */
      natural: Float64Array;
      naturalBlockMin: Float64Array;
      tabsBefore: Int32Array;
    }
    let model: NonMonotoneModel | null = null;
    const atomText = (from: number, to: number): string => {
      let text = '';
      for (let i = from; i < to; i++) {
        const atom = atoms[i];
        if (atom.type === 'text') text += atom.text;
      }
      return text;
    };
    const blockMins = (values: Float64Array): Float64Array => {
      const mins = new Float64Array(Math.ceil(values.length / BLOCK)).fill(Infinity);
      for (let i = 0; i < values.length; i++) {
        const b = Math.floor(i / BLOCK);
        if (values[i] < mins[b]) mins[b] = values[i];
      }
      return mins;
    };
    const buildModel = (): NonMonotoneModel => {
      const segStart = new Int32Array(end).fill(-1);
      const segEnd = new Int32Array(end).fill(-1);
      const head = new Float64Array(end).fill(NaN);
      const tail = new Float64Array(end);
      const tailSum = new Float64Array(end + 1);
      const natural = new Float64Array(end + 1);
      const tabsBefore = new Int32Array(end + 1);
      let first = -1;
      let style: T | undefined;
      let boundary: number | null = null;
      let base = 0;
      let text = 0;
      const close = (stop: number): void => {
        for (let t = first; first >= 0 && t < stop; t++) segEnd[t] = stop;
      };
      for (let t = 0; t < end; t++) {
        const atom = atoms[t];
        tabsBefore[t + 1] = tabsBefore[t] + (atom.type === 'tab' ? 1 : 0);
        if (!(atom.type === 'text' && first >= 0 && sameStyle(style as T, atom.style))) {
          close(t);
          base = natural[t];
          if (atom.type === 'text') {
            boundary = first >= 0 ? options.boundaryAdvance?.(style as T, atom.style) ?? 0 : null;
            first = t;
            style = atom.style;
          } else {
            first = -1;
            style = undefined;
          }
        }
        if (atom.type === 'text') {
          const segmentStyle = style as T;
          segStart[t] = first;
          if (t - first <= SHAPING_CONTEXT) {
            text = options.measureText(atomText(first, t + 1), segmentStyle);
            head[t] = text;
          } else {
            const context = atomText(t - SHAPING_CONTEXT, t);
            tail[t] = options.measureText(context + atom.text, segmentStyle)
              - options.measureText(context, segmentStyle);
            text += tail[t];
          }
          natural[t + 1] = base + (boundary === null ? text : text + boundary);
        } else {
          natural[t + 1] = base + (atom.type === 'object' ? atom.width : 0);
        }
        tailSum[t + 1] = tailSum[t] + tail[t];
      }
      close(end);
      return {
        segStart, segEnd, head, tail, tailSum,
        tailBlockMin: blockMins(tailSum),
        natural,
        naturalBlockMin: blockMins(natural),
        tabsBefore,
      };
    };

    /** Greatest i in (from, to] with offset + values[i] <= budget, else -1. */
    const lastFitting = (
      values: Float64Array, mins: Float64Array, offset: number, from: number, to: number, budget: number,
    ): number => {
      let i = to;
      while (i > from) {
        const lo = Math.floor(i / BLOCK) * BLOCK;
        if (lo > from && offset + mins[lo / BLOCK] > budget) {
          i = lo - 1;
          continue;
        }
        if (offset + values[i] <= budget) return i;
        i--;
      }
      return -1;
    };

    /** Text width of atoms [lineStart, stop) within one coalesced segment. */
    const lineHeadWidths = new Map<number, number>();
    const firstSegmentWidth = (m: NonMonotoneModel, lineStart: number, stop: number): number => {
      if (lineStart === m.segStart[lineStart]) {
        const last = Math.min(stop, lineStart + SHAPING_CONTEXT + 1) - 1;
        return m.head[last] + (m.tailSum[stop] - m.tailSum[last + 1]);
      }
      const headStop = Math.min(stop, lineStart + SHAPING_CONTEXT + 1);
      let width = lineHeadWidths.get(headStop);
      if (width === undefined) {
        width = options.measureText(atomText(lineStart, headStop), atoms[lineStart].style);
        lineHeadWidths.set(headStop, width);
      }
      return width + (m.tailSum[stop] - m.tailSum[headStop]);
    };

    const nonMonotoneFit = (lineStart: number, lineIndex: number, budget: number): number => {
      const m = model ??= buildModel();
      lineHeadWidths.clear();
      if (m.tabsBefore[end] !== m.tabsBefore[lineStart]) {
        return tabbedNonMonotoneFit(m, lineStart, lineIndex, budget);
      }
      const first = atoms[lineStart];
      const firstEnd = first.type === 'text' ? m.segEnd[lineStart] : lineStart + 1;
      const firstWidth = first.type === 'text' ? firstSegmentWidth(m, lineStart, firstEnd)
        : first.type === 'object' ? first.width : 0;
      if (firstEnd < end) {
        // The next segment's seam advance is relative to this line's first style.
        const next = atoms[firstEnd];
        const seam = first.type === 'text' && next.type === 'text' && options.boundaryAdvance
          ? options.boundaryAdvance(first.style, next.style)
            - options.boundaryAdvance(atoms[m.segStart[lineStart]].style, next.style)
          : 0;
        const found = lastFitting(m.natural, m.naturalBlockMin,
          firstWidth + seam - m.natural[firstEnd], firstEnd, end, budget);
        if (found >= 0) return found;
      }
      if (firstWidth <= budget) return firstEnd;
      if (first.type !== 'text') return lineStart;
      const headStop = lineStart + SHAPING_CONTEXT + 1;
      if (firstEnd - 1 > headStop) {
        const offset = firstSegmentWidth(m, lineStart, headStop) - m.tailSum[headStop];
        const found = lastFitting(m.tailSum, m.tailBlockMin, offset, headStop, firstEnd - 1, budget);
        if (found >= 0) return found;
      }
      for (let i = Math.min(firstEnd - 1, headStop); i > lineStart; i--) {
        if (firstSegmentWidth(m, lineStart, i) <= budget) return i;
      }
      return lineStart;
    };

    // Tabs make each prefix's width depend on its stop resolution. Resolve the
    // cached segment widths per candidate instead of re-measuring text.
    const tabGaps = new Map<number, number>();
    const tabbedNonMonotoneFit = (
      m: NonMonotoneModel, lineStart: number, lineIndex: number, budget: number,
    ): number => {
      const items: { isTab: boolean; width: number }[] = [];
      let fit = lineStart;
      let noStopGap = 0;
      let hasTab = false;
      let closedSum = 0;
      let segFirst = -1;
      let boundary: number | null = null;
      for (let i = lineStart + 1; i <= end; i++) {
        const index = i - 1;
        const atom = atoms[index];
        if (atom.type === 'text' && segFirst >= 0 && m.segStart[index] === m.segStart[index - 1]) {
          const text = firstSegmentWidthFrom(m, segFirst, lineStart, i);
          items[items.length - 1].width = boundary === null ? text : text + boundary;
        } else {
          if (items.length > 0) closedSum += items[items.length - 1].width;
          if (atom.type === 'text') {
            const previous = segFirst >= 0 ? atoms[segFirst] : undefined;
            boundary = previous?.type === 'text' ? options.boundaryAdvance?.(previous.style, atom.style) ?? 0 : null;
            segFirst = index;
            const text = firstSegmentWidthFrom(m, segFirst, lineStart, i);
            items.push({ isTab: false, width: boundary === null ? text : text + boundary });
          } else {
            segFirst = -1;
            if (atom.type === 'tab') {
              hasTab = true;
              let gap = tabGaps.get(index);
              if (gap === undefined) {
                gap = options.measureText(' ', atom.style);
                tabGaps.set(index, gap);
              }
              noStopGap = gap;
              items.push({ isTab: true, width: 0 });
            } else {
              items.push({ isTab: false, width: atom.width });
            }
          }
        }
        const width = hasTab
          ? resolveDrawingMlTabWidths(
            items,
            options.tabStops ?? [],
            options.tabStartPen?.(lineIndex) ?? (lineIndex === 0 ? options.firstLineIndent ?? 0 : 0),
            Infinity,
            noStopGap,
            options.defaultTabSize ?? 0,
          ).reduce((sum, value) => sum + value, 0)
          : closedSum + items[items.length - 1].width;
        if (i === end && width <= budget) return end;
        if (i < end && width <= budget) fit = i;
      }
      return fit;
    };
    /** Width of segment text [segFirst, stop), where segFirst is lineStart or a natural start. */
    const firstSegmentWidthFrom = (
      m: NonMonotoneModel, segFirst: number, lineStart: number, stop: number,
    ): number => segFirst === lineStart
      ? firstSegmentWidth(m, lineStart, stop)
      : firstSegmentWidth(m, segFirst, stop);

    let start = 0;
    if (end === 0) {
      lines.push({ segments: [], width: 0, ...(regionIndex + 1 < regions.length ? { endsWithBreak: true } : {}) });
      continue;
    }
    while (start < end) {
      const lineIndex = lines.length;
      const budget = options.maxWidth - (lineIndex === 0 ? options.firstLineIndent ?? 0 : 0);
      const widthAt = (stop: number): number =>
        plainText !== null && plainOffsets && plainStyle !== undefined
          ? options.measureText(plainText.slice(plainOffsets[start], plainOffsets[stop]), plainStyle)
          : makeSegments(start, stop, lineIndex).width;
      let fit = end;
      if (Number.isFinite(budget) && options.nonMonotoneMeasure) {
        // A later prefix can fit after an earlier overflow, so every candidate
        // is still checked and the greatest fitting one wins (the whole line
        // when it fits). Prefix widths are accumulated, not re-measured.
        fit = nonMonotoneFit(start, lineIndex, budget);
      } else if (Number.isFinite(budget)) {
        // Exponential search only measures prefixes close to the fit boundary.
        // In a one-grapheme box this avoids measuring the whole remaining word
        // on every visual line. The existing binary search still chooses the
        // greatest fitting prefix when advances are monotone.
        let lo = start;
        let hi = end;
        let step = 1;
        while (lo < end) {
          const probe = Math.min(end, start + step);
          if (widthAt(probe) <= budget) {
            lo = probe;
            if (probe === end) break;
            step *= 2;
          } else {
            hi = probe - 1;
            break;
          }
        }
        while (lo < hi) {
          const mid = Math.ceil((lo + hi) / 2);
          if (widthAt(mid) <= budget) lo = mid;
          else hi = mid - 1;
        }
        fit = lo;
      }
      if (fit >= end) {
        const line = makeSegments(start, end, lineIndex);
        if (regionIndex + 1 < regions.length) line.endsWithBreak = true;
        lines.push(line);
        break;
      }

      let split = 0;
      for (let i = start + 1; i <= fit; i++) if (breakAt[i]) split = i;
      if (split > start && isSingleExplicitTabCell(split)) {
        // Office keeps the tab and following Latin cell on the authored line
        // for an explicit stop, even when the cell passes the box edge. The
        // control C11 uses the default grid and wraps before its tab instead.
        split = end;
      }
      if (isSingleExplicitTabCell(start)) {
        // An authored stop makes the next Latin cell one visual unit even if
        // that cell contains a hyphen or another soft opportunity. PowerPoint
        // keeps this cell together in its explicit-stop paragraphs; C11 has
        // only the default grid and still breaks within the following word.
        split = start + 2;
        while (split < end && !isSpace(atoms[split]) && atoms[split].type !== 'tab') split++;
      }
      if (split === 0) {
        if (atoms[start].type === 'tab' && fit <= start + 1) {
          // A stop beyond the box still carries its next cell. A CJK cell
          // contributes at least one glyph; an indivisible Latin cell stays
          // together. The painter can clamp an explicit off-box stop.
          split = start + 1;
          if (split < end && atoms[split].type === 'text') {
            split++;
            while (split < end && !breakAt[split]) split++;
          }
        } else {
          split = Math.max(start + 1, fit); // overwide word: grapheme-safe emergency break
          // A mixed-font/style Latin word has no shaping-preserving emergency
          // break: retaining the old overflow avoids inventing a run seam as a
          // break. The matched C01 control has identical styles and does split.
          const nextOpportunity = nextBreak[start];
          if (nonTextOrCjk[nextOpportunity] === nonTextOrCjk[start]
              && styleSeams[nextOpportunity] > styleSeams[start + 1]) {
            split = nextOpportunity;
          }
        }
      }

      // Kinsoku (§17.15.1.58–.60) adjusts an in-run CJK boundary. The Office
      // C08 control is a counterexample at an authored run seam, so leave that
      // boundary intact instead of inferring a cross-run retraction rule.
      if (split < end && atoms[split - 1].run === atoms[split].run
          && (isCjk(atoms[split - 1]) || isCjk(atoms[split]))) {
        const left = atoms.slice(start, split).flatMap((atom) => atom.type === 'text' ? [...atom.text] : []);
        const right = atoms.slice(split).flatMap((atom) => atom.type === 'text' ? [...atom.text] : []);
        if (left.length > 1 && right.length > 0) {
          const adjusted = kinsokuAdjustedSplit([...left, ...right], left.length, DEFAULT_KINSOKU_RULES, 1);
          const retract = left.length - adjusted;
          if (retract > 0 && split - retract > start) split -= retract;
        }
      }

      // Keep authored spaces on the closed line. The controlled PDFs establish
      // the visible word break, but cannot identify the advance of invisible
      // trailing spaces; retaining them preserves that unresolved paint detail.
      lines.push(makeSegments(start, split, lineIndex));
      start = split;
      while (start < end && isSpace(atoms[start])) start++;
    }
  }
  return lines;
}
