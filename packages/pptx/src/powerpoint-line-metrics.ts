import { findReferenceFontMetrics } from '@silurus/ooxml-core';
import { excelDrawingMlLineRatios } from '@silurus/ooxml-core/internal/office-auto-line';

/**
 * PowerPoint's split of a text line at its baseline (observed behaviour; the
 * DrawingML text model in ECMA-376 §21.1.2 does not say where the baseline
 * sits inside a line).
 *
 * PowerPoint for Mac 16.113.2 PDF controls (#1610). Two decks with 432
 * boxes, every face verified from the embedded PDF fonts:
 * - The line box is 1.2 × the largest run size for every face.
 * - A run's ascent share of that box is usWinAscent / (usWinAscent +
 *   usWinDescent). A face that sets fsSelection USE_TYPO_METRICS uses
 *   (sTypoAscender + sTypoLineGap) / (sTypoAscender − sTypoDescender +
 *   sTypoLineGap) instead.
 *   - Faces that separate the tables: Yu Gothic (usWin 0.765, hhea 0.798),
 *     Baskerville Old Face, Stencil, and Gabriola (typo 0.814, usWin 0.647).
 *   - Unlike Excel, there is no Far East 1.3× leading and no dependence on
 *     the face's hhea table.
 * - Font resolution follows the face PowerPoint really used, not the name:
 *   - Office-bundled faces and the macOS Supplemental faces (Times New
 *     Roman, Georgia, Courier New) are used as named. A face present in both
 *     resolves to the Supplemental copy (Times New Roman's hhea lineGap 87 is
 *     the one embedded).
 *   - PowerPoint substitutes faces that only macOS /System/Library/Fonts
 *     provides:
 *     - Palatino → Palatino Linotype;
 *     - Helvetica and Helvetica Neue → Arial (embedded under the Helvetica
 *       name with Arial's glyphs and usWin metrics);
 *     - Avenir, Menlo and Hiragino Sans → the deck's theme font or a Far
 *       East fallback.
 *   - Only the first two mappings are pinned. Every other system-only face
 *     returns undefined, and the caller keeps its ordinary line model.
 */
const OBSERVED_SUBSTITUTES: Readonly<Record<string, string>> = {
  palatino: 'Palatino Linotype',
  helvetica: 'Arial',
  'helvetica neue': 'Arial',
};

/**
 * Resolved shares keyed by face/weight/style. Font names come from document
 * content, so the cache is a bounded LRU: at most SHARE_CACHE_LIMIT entries,
 * the least recently used evicted first. A miss only re-reads the static
 * reference table, so the cap trades a little lookup work for bounded memory.
 */
export const SHARE_CACHE_LIMIT = 256;
const shareCache = new Map<string, number | null>();

/** @internal Test hook: current number of cached face keys. */
export function powerPointShareCacheSize(): number {
  return shareCache.size;
}

export function powerPointAscentShare(
  family: string,
  bold: boolean,
  italic: boolean,
): number | undefined {
  const key = `${family.trim().toLocaleLowerCase('en-US')}|${bold ? 700 : 400}|${italic ? 'i' : 'n'}`;
  const cached = shareCache.get(key);
  if (cached !== undefined) {
    shareCache.delete(key);
    shareCache.set(key, cached);
    return cached ?? undefined;
  }
  const share = resolveShare(family, bold, italic);
  shareCache.set(key, share ?? null);
  if (shareCache.size > SHARE_CACHE_LIMIT) {
    const oldest = shareCache.keys().next().value;
    if (oldest !== undefined) shareCache.delete(oldest);
  }
  return share;
}

type Profile = ReturnType<typeof findReferenceFontMetrics>[number];

/** The profiles of the copy PowerPoint lays the face out with (see above). */
function chosenProfiles(family: string, bold: boolean, italic: boolean): Profile[] {
  const trimmed = family.trim();
  if (!trimmed) return [];
  const name = OBSERVED_SUBSTITUTES[trimmed.toLocaleLowerCase('en-US')] ?? trimmed;
  const style = italic ? 'italic' : 'normal';
  let profiles = findReferenceFontMetrics(name, { weight: bold ? 700 : 400, style });
  if (profiles.length === 0) {
    // A face name that is itself one cut of a family ("Calibri Light" is
    // weight 300) has no 400 or 700 profile; use that cut when the name
    // selects a single weight. #1630 controls (reference PDF export): Calibri
    // Light titles sat at the 1950 / 2500 usWin share, like Calibri. A bold
    // Calibri Light title embedded Calibri-Light itself (#1435 deck).
    const named = findReferenceFontMetrics(name, { style });
    if (named.length > 0 && named.every((p) => p.weight === named[0].weight)) profiles = named;
  }
  const supplemental = profiles.filter((p) => p.source === 'macos-supplemental');
  return supplemental.length > 0 ? supplemental : profiles.filter((p) => p.source === 'office-mac');
}

function resolveShare(family: string, bold: boolean, italic: boolean): number | undefined {
  const chosen = chosenProfiles(family, bold, italic);
  if (chosen.length === 0) return undefined;
  let share: number | undefined;
  for (const profile of chosen) {
    let next: number | undefined;
    if (profile.typoMetrics) {
      const [ascender, descender, lineGap] = profile.typoMetrics;
      const above = ascender + Math.max(0, lineGap);
      next = above / (above - descender);
    } else if (profile.win) {
      const [ascent, descent] = profile.win;
      next = ascent / (ascent + descent);
    }
    if (next === undefined || !Number.isFinite(next) || next <= 0 || next >= 1) return undefined;
    if (share !== undefined && Math.abs(share - next) > 1e-12) return undefined;
    share = next;
  }
  return share;
}

/**
 * The #1604 Excel natural line box (`excelDrawingMlLineRatios`) of the copy
 * PowerPoint uses, as em ratios. PowerPoint lays a body out with it when the
 * effective `a:bodyPr@compatLnSpc` is an explicit 0 (see
 * `PowerPointFaceMetrics.excel`). The copy is the same one the #1610 share
 * uses: a macOS Supplemental copy projects with the system tables (Times New
 * Roman: hhea ascender + lineGap 87, 1.150 em), an Office-bundled copy with
 * the Windows tables. Undefined when the copies disagree or a table is
 * missing; the body then keeps the ordinary line model.
 */
function resolveExcelBox(family: string, bold: boolean, italic: boolean): ExcelLineBox | undefined {
  let box: ExcelLineBox | undefined;
  for (const profile of chosenProfiles(family, bold, italic)) {
    if (profile.farEastCodePage == null || !profile.win) return undefined;
    const ratios = excelDrawingMlLineRatios({
      faceSource: profile.source === 'macos-supplemental' ? 'system' : 'office-bundle',
      unitsPerEm: profile.unitsPerEm,
      hhea: profile.hhea,
      win: profile.win,
      typoMetrics: profile.typoMetrics,
      farEastCodePage: profile.farEastCodePage,
    });
    if (!ratios) return undefined;
    if (box && (box.ascent !== ratios.ascentRatio || box.descent !== ratios.descentRatio)) return undefined;
    box = { ascent: ratios.ascentRatio, descent: ratios.descentRatio };
  }
  return box;
}

/** A face's natural line box as em ratios. */
export interface ExcelLineBox {
  ascent: number;
  descent: number;
}

/**
 * Everything PowerPoint's two line models need from one face. Instances are
 * interned per face/weight/style, so two segments compare equal exactly when
 * they size a line with the same face.
 */
export interface PowerPointFaceMetrics {
  /** #1610 ascent share of the 1.2 × size line box. */
  readonly share: number;
  /** #1604 natural box, used under an explicit compatLnSpc="0". */
  readonly excel: ExcelLineBox | undefined;
}

const faceCache = new Map<string, PowerPointFaceMetrics | null>();

/**
 * The line metrics of a face, or undefined when its share is unresolvable
 * (the caller then keeps the ordinary line model for the whole body). Bounded
 * LRU like the share cache.
 */
export function powerPointFaceMetrics(
  family: string,
  bold: boolean,
  italic: boolean,
): PowerPointFaceMetrics | undefined {
  const key = `${family.trim().toLocaleLowerCase('en-US')}|${bold ? 700 : 400}|${italic ? 'i' : 'n'}`;
  const cached = faceCache.get(key);
  if (cached !== undefined) {
    faceCache.delete(key);
    faceCache.set(key, cached);
    return cached ?? undefined;
  }
  const share = powerPointAscentShare(family, bold, italic);
  const metrics = share === undefined ? null
    : Object.freeze({ share, excel: resolveExcelBox(family, bold, italic) });
  faceCache.set(key, metrics);
  if (faceCache.size > SHARE_CACHE_LIMIT) {
    const oldest = faceCache.keys().next().value;
    if (oldest !== undefined) faceCache.delete(oldest);
  }
  return metrics ?? undefined;
}

/** One run's contribution to a line: its authored size and ascent share. */
export interface PowerPointLineRun {
  sizePx: number;
  share: number;
}

/**
 * The natural line box of one line. Each run claims 1.2 × its own size,
 * split by its share. The line unions the runs' ascent and descent parts and
 * rescales them into 1.2 × the largest size.
 *
 * Measured on mixed-face lines (Arial+Meiryo in both orders, Arial+MS Gothic,
 * Calibri+Yu Gothic, Arial+Gabriola) and mixed sizes (40+100 pt): all exact.
 * Taking the largest descent instead is off by up to 5 px (1/100 in), and the
 * largest ascent by 13 px.
 */
export function powerPointNaturalLine(runs: readonly PowerPointLineRun[]): { ascent: number; descent: number } {
  let maxSize = 0;
  let ascent = 0;
  let descent = 0;
  for (const run of runs) {
    maxSize = Math.max(maxSize, run.sizePx);
    ascent = Math.max(ascent, 1.2 * run.sizePx * run.share);
    descent = Math.max(descent, 1.2 * run.sizePx * (1 - run.share));
  }
  const height = 1.2 * maxSize;
  if (!(ascent + descent > 0)) return { ascent: height * 0.8, descent: height * 0.2 };
  const a = height * ascent / (ascent + descent);
  return { ascent: a, descent: height - a };
}

/**
 * PowerPoint rounds `spcPts` to whole points before laying out the line
 * (#1610 controls: 40.5 → 41, 45.25 → 45, 48.33 → 48, 49.25 → 49 and
 * 50.75 → 51 pt, each over eight lines). Font sizes stay exact: 10.5, 11.5,
 * 13.33, 40.5 and 55.5 pt keep a 1.2 × size line.
 */
export function powerPointExactLinePoints(points: number): number {
  return Math.floor(points + 0.5);
}

/**
 * The natural line box when the effective `a:bodyPr@compatLnSpc` is an
 * explicit 0 (ECMA-376 §21.1.2.1.1 gives no algorithm, only "decided in a
 * simplistic manner using the font scene").
 *
 * Observed with PowerPoint's reference (Windows-style) PDF export, #1619
 * controls: every left/right pair differing only in compatLnSpc 0/1, over
 * Arial, Calibri, Times New Roman, Gabriola, Meiryo, Yu Gothic, MS Gothic and
 * Aptos at 18-100 pt, lnSpc omitted / 60-150 % / 30-100 pt, spcBef/spcAft in
 * points and percent with and without spcFirstLastPara, anchors t/ctr/b and
 * mixed sizes and faces:
 * - compatLnSpc="1" renders exactly like an omitted compatLnSpc (the #1610
 *   model), whether it is authored on the slide, the layout or the master.
 * - An explicit 0 selects Excel's #1604 model instead: each run claims its
 *   own natural box (`PowerPointFaceMetrics.excel` × size), the line is the
 *   largest ascent over the largest descent (no rescale to 1.2 × size),
 *   lnSpc goes through the same `drawingMlSpacedLineBox` rule, and a
 *   percentage spcBef/spcAft is a fraction of that natural line.
 * All 546 baselines with fontAlgn omitted/auto/base, in both models, land on
 * the export's 1/100 in unit (3 sit within 0.005 unit of a rounding tie).
 * fontAlgn t/ctr/b are outside this rule (tracked separately).
 */
export function powerPointCompatOffNaturalLine(runs: readonly { sizePx: number; box: ExcelLineBox }[]): {
  ascent: number;
  descent: number;
} {
  let ascent = 0;
  let descent = 0;
  for (const run of runs) {
    ascent = Math.max(ascent, run.box.ascent * run.sizePx);
    descent = Math.max(descent, run.box.descent * run.sizePx);
  }
  return { ascent, descent };
}

/*
 * PDF exports quantize: each baseline lands on a whole device unit counted
 * from the anchored text top (1/100 in in the #1610 controls; 1/150 in in a
 * corpus export whose 12 pt lines alternate 13.92/14.88 pt). The unit belongs
 * to the export device, not to the slide, so layout keeps continuous
 * positions; the controls match to within half a unit.
 */
