import type { ParagraphLayoutRun, ParagraphTextBearingRun } from '../layout/text.js';
import type { LineLayoutEnvironment } from './model.js';
import { resolveFieldText } from './text-runs.js';
import type { FieldRun } from '../types.js';

export interface TextSequenceSource {
  readonly runIndex: number;
  /** UTF-16 offsets in the complete displayed sequence. */
  readonly start: number;
  readonly end: number;
}
export interface TextSequence {
  readonly run: Extract<ParagraphTextBearingRun, { type: 'text' }>;
  readonly sources: readonly TextSequenceSource[];
}

function visibleTextRun(run: ParagraphLayoutRun, environment: LineLayoutEnvironment) {
  if (run.type !== 'text' && run.type !== 'field') return undefined;
  const r: ParagraphTextBearingRun = run;
  if (environment.showTrackedChanges !== true
    && (r.revision?.kind === 'deletion' || r.revision?.kind === 'moveFrom')) return undefined;
  // These are authored units, rather than formatting-only seams: a ruby base,
  // note marker, fitText unit, noBreakHyphen owner or upright tate-chu-yoko cell.
  if (r.ruby || r.noteRef || r.fitTextVal != null || r.noBreakRanges?.length
    || r.noBreakBefore || r.noBreakAfter || r.eastAsianVert) return undefined;
  const text = run.type === 'text' ? run.text : resolveFieldText(run as FieldRun, environment);
  const { type: _type, ...properties } = run;
  const projected = { ...properties, type: 'text' as const, text,
    isLink: 'isLink' in run ? run.isLink : false,
    hyperlink: 'hyperlink' in run ? run.hyperlink : null };
  return projected as Extract<ParagraphTextBearingRun, { type: 'text' }>;
}

function formattingKey(run: Extract<ParagraphTextBearingRun, { type: 'text' }>, environment: LineLayoutEnvironment): string {
  const { text: _text, type: _type, isLink: _isLink, ...properties } = run;
  const record = properties as Record<string, unknown>;
  // A field's instruction/result and revision provenance are source ownership,
  // not text formatting. Markup-view revisions still own their painted style.
  delete record.fieldType;
  delete record.instruction;
  delete record.fallbackText;
  if (environment.showTrackedChanges !== true) delete record.revision;
  return JSON.stringify(record, (_key, value: unknown) =>
    value && typeof value === 'object' && !Array.isArray(value)
      ? Object.fromEntries(Object.entries(value).filter(([, v]) => v != null)
        .sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0))
      : value);
}

/**
 * Library invariant: formatting-only source seams never enter the glyph/token
 * stream. ECMA-376 §§17.3.2.25/.19 separate the content container from its
 * resolved run formatting/kerning threshold; this normalization does not infer
 * another Office kerning switch or fitting allowance. Every layout consumer
 * uses the same concatenated sequence, so a separator pair exists exactly when
 * that separator is in the measured range, including emergency prefixes.
 *
 * Source ownership remains a separate UTF-16 map for comments, revisions and
 * field results. Tabs and real font/property changes remain semantic boundaries
 * in the ordinary segment builder. One key per source run, one join per sequence
 * and one source sweep keep normalization linear in input size.
 */
export function acquireTextSequences(
  runs: readonly ParagraphLayoutRun[],
  environment: LineLayoutEnvironment,
  displayText: (text: string, run: Extract<ParagraphTextBearingRun, { type: 'text' }>) => string,
): ReadonlyMap<number, TextSequence> {
  const sequences = new Map<number, TextSequence>();
  // Evidence gap: U+3000 hanging and source-split fitting are not settled by
  // WORD_COMPRESSED_SPACE_LINE_FIT. Its prior-behavior gate is paragraph-wide,
  // so unrelated earlier seams must not change that paragraph's fitting either.
  if (runs.some(run => {
    if (run.type !== 'text' && run.type !== 'field') return false;
    // Final-view omissions are not acquisition input, including field results.
    if (environment.showTrackedChanges !== true
      && (run.revision?.kind === 'deletion' || run.revision?.kind === 'moveFrom')) return false;
    return (run.type === 'text' ? run.text : resolveFieldText(run as FieldRun, environment)).includes('\u3000');
  })) return sequences;
  let start = -1;
  let key: string | undefined;
  let first: Extract<ParagraphTextBearingRun, { type: 'text' }> | undefined;
  let texts: string[] = [];
  let sources: TextSequenceSource[] = [];
  let offset = 0;
  const finish = () => {
    if (first && sources.length > 1) sequences.set(start, {
      run: Object.freeze({ ...first, text: texts.join('') }),
      sources: Object.freeze(sources),
    });
    start = -1;
    first = undefined;
    texts = [];
    sources = [];
    offset = 0;
  };
  for (const [runIndex, run] of runs.entries()) {
    const visible = visibleTextRun(run, environment);
    const scopeKey = visible ? environment.layoutServices?.text.sourceScopeKey?.({
      text: displayText(visible.text, visible), fontSizePt: visible.fontSize,
      fonts: visible.fontSlots?.direct ?? { ascii: visible.fontFamily,
        highAnsi: visible.fontFamilyHighAnsi ?? visible.fontFamily,
        eastAsia: visible.fontFamilyEastAsia ?? visible.fontFamily,
        complexScript: visible.fontFamilyCs ?? visible.fontFamily },
      themeFonts: visible.fontSlots?.theme, themeFontPresence: visible.fontSlots?.themePresent,
      weight: (visible.rtl || visible.cs ? visible.boldCs : visible.bold) ? 700 : 400,
      style: (visible.rtl || visible.cs ? visible.italicCs : visible.italic) ? 'italic' : 'normal',
      complexScript: visible.rtl === true || visible.cs === true,
      fontHint: visible.fontHint, eastAsiaLanguage: visible.langEastAsia,
    }) : undefined;
    const nextKey = visible ? JSON.stringify([
      formattingKey(visible, environment), scopeKey === 'mixed' ? runIndex : scopeKey ?? null,
    ]) : undefined;
    if (!visible || nextKey !== key) finish();
    key = nextKey;
    if (!visible) continue;
    if (!first) { start = runIndex; first = visible; }
    texts.push(visible.text);
    const display = displayText(visible.text, visible);
    sources.push(Object.freeze({ runIndex, start: offset, end: offset + display.length }));
    offset += display.length;
  }
  finish();
  return sequences;
}
