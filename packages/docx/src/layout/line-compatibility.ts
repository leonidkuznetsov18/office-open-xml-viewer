import { defineCompatibilityRule } from './compatibility.js';
import { OFFICE_FAR_EAST_SINGLE_LINE_FACTOR, officeOpenTypeAutoLineRatios } from '@silurus/ooxml-core/internal/office-auto-line';
import type { LineSpacing, TabStop } from '../types.js';

export const WORD_TAB_DISPLACED_READING_FRAME = defineCompatibilityRule({
  id: 'word-float-tab-reading-frame',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/first-line-float-indent.test.ts#$kind $alignment ($count), rtl=$rtl, float=$float matches Word geometry',
  },
  description: 'Issue #1672 controlled Word exports cover ordinary left/right/center/decimal tabs, one/two tabs, both paragraph directions and either float edge, with no-float counterexamples. Eligibility stays margin-relative (ECMA-376 §17.3.1.37), but an authored target moves with a float-displaced leading line edge. Library policy retains this observation only where text fits the available float band. Word also overflows some tab cells into floats or past the margin; that unspecified fallback is intentionally unsupported. Out-of-band targets break the line and following text uses ordinary legal-break and emergency fitting, without an overflow allowance or a fabricated wide reference band.',
});

/** Projection of {@link WORD_TAB_DISPLACED_READING_FRAME}. The caller has already
 * selected the next eligible stop in margin coordinates. */
export function wordFloatTabStopPosition(
  stopPosition: number,
  custom: boolean,
  leadingShift: number,
): number {
  return custom ? stopPosition + Math.max(0, leadingShift) : stopPosition;
}

export const WORD_POSITIONAL_TAB_AVAILABLE_BAND = defineCompatibilityRule({
  id: 'word-positional-tab-available-band',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/first-line-float-indent.test.ts#$kind $alignment ($count), rtl=$rtl, float=$float matches Word geometry',
  },
  description: 'Issue #1672 Word exports cover all 36 positional-tab combinations of left/center/right alignment, margin/indent reference, LTR/RTL paragraph direction and no/left/right float. Word intersects the normative reference box (ECMA-376 §§17.3.3.23, 17.18.71, 17.18.73) with the available float band and aligns the following cell in reading order; no-float references remain unchanged. A target at the current pen is reachable without a line break; only a target behind it requires the next line. RTL must retain the positional descriptor through the bidi post-pass rather than resolving it as an ordinary tab.',
});

/** Reference-box projection of {@link WORD_POSITIONAL_TAB_AVAILABLE_BAND}. */
export function wordPositionalTabReferenceBox(
  referenceStart: number,
  referenceEnd: number,
  bandStart: number,
  bandEnd: number,
  narrowed: boolean,
): Readonly<{ start: number; end: number }> {
  return narrowed
    ? { start: Math.max(referenceStart, bandStart), end: Math.min(referenceEnd, bandEnd) }
    : { start: referenceStart, end: referenceEnd };
}

export const WORD_OPENTYPE_FEATURES_COMPAT_KERNING = defineCompatibilityRule({
  id: 'word-opentype-features-compat-kerning',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/run-char-metrics-render.test.ts#enables absent-threshold kerning only under enableOpenTypeFeatures',
  },
  description: '[MS-DOCX] §2.3.3 stores enableOpenTypeFeatures as a named compatibility setting. When enabled, an unqualified run enables OpenType kerning; an explicit or style-resolved w:kern threshold remains authoritative. Both line measurement and paint use the same resolved threshold.',
});

export const WORD_NUMBERING_MARKER_FIRST_LINE_UNION = defineCompatibilityRule({
  id: 'word-numbering-marker-first-line-union',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'numbering-marker-font-size-line-box-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'In a non-grid paragraph with 1.15 automatic spacing, a text marker participates in the first-line ascent/descent union. Relative to a marker-free control, 8, 14, and 20 pt markers added 0, 1.68, and 7.68 pt to the line advance; changing the marker alone shifted subsequent paragraph baselines by the same amount. Added automatic leading follows the body text single-line height rather than scaling the taller marker box. Exact spacing, grids, ruby, wrapping floats, picture markers, and continuation lines are outside the measured scope.',
});

export const WORD_EAST_ASIAN_GRID_LINE_ALLOCATION = defineCompatibilityRule({
  id: 'word-east-asian-grid-line-allocation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'word-font-metrics-resource-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Across eight East-Asian resources, six sizes, and no-grid/grid/useFELayout variants, Word allocated grid cells from the selected face line height, rounding up to whole pitches. A later controlled font-table intervention found that OS/2 code-page bits 17–20, rather than cmap/script alone, select the 1.3-times-hhea line box in Word for Mac. The grid-cell rule takes the resulting height as input; it does not itself classify the face. Mixed rFonts slots also prevent a family-wide Latin override.',
});

export const WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT = defineCompatibilityRule({
  id: 'word-table-cell-ignores-grid-right-indent-adjustment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'table-cell-adjust-right-indent-width-position-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In the observed linesAndChars matrix, paragraphs inside fixed-width table cells retain the same line breaks for omitted (default true) and explicit-false w:adjustRightInd across four boundary widths and both left/right cell positions. Scope this Word-only exception to table-cell containers; ordinary body paragraphs retain the ECMA-376 §17.3.1.1 adjustment.',
});

export const WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT = defineCompatibilityRule({
  id: 'word-snap-to-chars-east-asian-cell-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'snap-to-chars-east-asian-cell-fit-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For snapToChars, Word centers each East-Asian grapheme independently in the smallest whole number of character-pitch units that contains its natural advance. A grapheme that fits uses the one-unit placement described by [MS-OI29500] §2.1.534; an undersized authored pitch expands only that grapheme to additional units.',
});

export const WORD_SNAP_TO_CHARS_SCRIPT_BLOCK_ALLOCATION = defineCompatibilityRule({
  id: 'word-snap-to-chars-script-block-allocation',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.534',
  },
  description: 'Allocate snapToChars Latin text in contiguous blocks centered across the required grid units, complex-script blocks from their leading edge, and East-Asian graphemes independently by character cell.',
});

/** Word compatibility projection governed by
 * {@link WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT}. */
export function wordSnapToCharsEastAsianCellCount(
  naturalAdvancePt: number,
  pitchPt: number,
): number {
  if (!(pitchPt > 0) || !Number.isFinite(naturalAdvancePt)) return 1;
  return Math.max(1, Math.ceil(Math.max(0, naturalAdvancePt) / pitchPt - 1e-9));
}

/** Compatibility projection governed by
 * {@link WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT}. */
export function wordContainerAllowsGridRightIndentAdjustment(
  insideTableCell: boolean,
): boolean {
  return !insideTableCell;
}

export const WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT = defineCompatibilityRule({
  id: 'word-grid-right-indent-pitch-alignment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'grid-right-indent-character-pitch-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For body paragraphs whose ECMA-376 §17.3.1.1 adjustment is enabled on a linesAndChars character grid, Word reduces the physical line width to the greatest whole character-pitch multiple not exceeding the available width. The observed matrix covers exact and non-exact widths, zero and negative charSpace, explicit opt-out, line-only control, both physical indent sides, and the separately registered table-cell exception.',
});

/** Word compatibility projection governed by
 * {@link WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT}. */
export function wordGridRightIndentAdjustmentPt(
  availableWidthPt: number,
  pitchPt: number,
): number {
  if (!(pitchPt > 0) || !Number.isFinite(availableWidthPt) || availableWidthPt <= 0) {
    return 0;
  }
  const remainder = ((availableWidthPt % pitchPt) + pitchPt) % pitchPt;
  const epsilon = 1e-9;
  return remainder <= epsilon || pitchPt - remainder <= epsilon ? 0 : remainder;
}

export const WORD_HANGING_TAB_SAME_POSITION_PRECEDENCE = defineCompatibilityRule({
  id: 'word-hanging-tab-same-position-precedence',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'hanging-indent-authored-tab-collision-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When the implicit tab created by a hanging indent shares its coordinate with an authored center, end, or start stop, Word resolves one advancing stop at that coordinate using the authored alignment. An authored bar remains an independent drawing rule, so the implicit advancing stop survives beside it. If center/end alignment would place following text before the current pen, the tab contributes zero advance.',
});

/** Compatibility projection governed by
 * {@link WORD_HANGING_TAB_SAME_POSITION_PRECEDENCE}. */
export function wordAuthoredTabReplacesImplicitHangingStop(
  alignment: TabStop['alignment'],
): boolean {
  return alignment !== 'bar' && alignment !== 'clear';
}

export const WORD_RTL_DECIMAL_TAB_PHYSICAL_ALIGNMENT = defineCompatibilityRule({
  id: 'word-rtl-decimal-tab-physical-alignment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'rtl-decimal-tab-run-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For LTR numeric cells embedded in a bidi paragraph, Word aligns the physical left edge of the first halfwidth period to the decimal stop across source-run boundaries. When no period exists, it aligns the numeric cell\'s physical right edge to the stop.',
});

export const WORD_DECIMAL_TAB_SEPARATOR_RESOLUTION = defineCompatibilityRule({
  id: 'word-decimal-tab-separator-resolution',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.556',
  },
  description: 'Use the first explicit halfwidth period as the decimal-tab alignment point; when absent, use the implicit separator after the final digit of the first Unicode decimal-number sequence.',
});

export const WORD_USE_FE_LAYOUT_INHERITED_GRID_MINIMUM = defineCompatibilityRule({
  id: 'word-use-fe-layout-inherited-grid-minimum',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'use-fe-layout-visible-script-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With useFELayout enabled, a visible Latin line with a resolved eastAsia font axis participates in Far East grid metrics even when w:rFonts@hint is absent; inherited automatic spacing keeps the larger of its whole-cell design allocation and one grid pitch multiplied by the inherited spacing value.',
});

export const WORD_USE_FE_LAYOUT_EMPTY_MARK_GRID_ALLOCATION = defineCompatibilityRule({
  id: 'word-use-fe-layout-empty-mark-grid-allocation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'use-fe-layout-empty-mark-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With useFELayout enabled, a content-less paragraph mark participates in Far East whole-cell document-grid allocation even when the document contains no literal East Asian text. Its face-specific Far East design height governs the cell count; exact spacing and snapToGrid=false remain the document-grid overrides named by ECMA-376 §17.6.5. Observed Word output gives signed atLeast spacing a discontinuous boundary on an active grid: negative values use their absolute magnitude as the mark advance, zero keeps the ordinary atLeast-zero advance regardless of inheritance source, and positive values retain whole-cell allocation.',
});

/** Compatibility projection governed by the useFELayout empty-mark allocation
 * and {@link WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED}. The caller supplies the
 * ordinary line-spacing result and the mark's whole-cell grid allocation.
 * Exact spacing is the normative §17.6.5 override. Observed Word output gives
 * signed atLeast values an empty-mark-specific negative/zero/positive boundary. */
export function wordUseFeLayoutParagraphMarkGridAdvancePx(
  input: Readonly<{
    ordinaryAdvancePx: number;
    allocatedGridAdvancePx: number;
    atLeastZeroAdvancePx: number;
    lineSpacing: LineSpacing | null;
    gridAllocationActive: boolean;
    scale: number;
  }>,
): number {
  const {
    ordinaryAdvancePx,
    allocatedGridAdvancePx,
    atLeastZeroAdvancePx,
    lineSpacing,
    gridAllocationActive,
    scale,
  } = input;
  if (!gridAllocationActive) return ordinaryAdvancePx;
  if (lineSpacing?.rule === 'atLeast' && lineSpacing.value < 0) {
    return Math.abs(lineSpacing.value) * scale;
  }
  if (lineSpacing?.rule === 'atLeast' && lineSpacing.value === 0) {
    return atLeastZeroAdvancePx;
  }
  return lineSpacing?.rule === 'exact'
    ? ordinaryAdvancePx
    : Math.max(ordinaryAdvancePx, allocatedGridAdvancePx);
}

export const WORD_CONTIGUOUS_UNDERLINE_GEOMETRY = defineCompatibilityRule({
  id: 'word-contiguous-underline-geometry',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/paragraph.test.ts#keeps a solid underline continuous across floating-precision retained run seams',
  },
  description: 'Adjacent compatible underlined source runs share one safe baseline and continuous authored cadence while style, color, and thickness boundaries remain distinct.',
});

export const WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED = defineCompatibilityRule({
  id: 'word-grid-at-least-tall-line-unsnapped',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/line-box-height.test.ts#does not round tall East Asian content up to an additional grid cell',
  },
  description: 'An explicitly authored atLeast line on an active document grid keeps the maximum of its natural height, authored minimum, and one pitch instead of rounding tall content to another whole cell.',
});

export const WORD_DEGENERATE_LINE_SPACING_SINGLE = defineCompatibilityRule({
  id: 'word-degenerate-line-spacing-single',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-DOC] §2.9.146',
  },
  description: 'Preserve a non-collapsing single-line fallback for exact or automatic line spacing at or below zero, consistent with the native LSPD representation.',
});

export const WORD_AUTO_MULTIPLE_BASELINE_PIN = defineCompatibilityRule({
  id: 'word-auto-multiple-baseline-pin',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'auto-multiple-baseline-pin',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Paint a positive automatic line-spacing multiplier with its glyph baseline pinned inside the single design line, placing extra leading or compressed overflow toward block-end; this is draw-only and does not replace the centered trailing-mark pagination metric.',
});

export const WORD_MIXED_ANCHOR_VISIBLE_LINE_METRICS = defineCompatibilityRule({
  id: 'word-mixed-anchor-visible-line-metrics',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/anchor-host-metrics.test.ts#reserves host line height without using its zero-ink box for a visible run baseline',
  },
  description: 'A zero-ink drawing anchor host reserves its line and grid height while visible neighboring glyphs retain their own ascent, descent, and design-line baseline.',
});

export const WORD_JUSTIFICATION_LEADING_INDENT_EXCLUSION = defineCompatibilityRule({
  id: 'word-justification-leading-indent-exclusion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/text-distribute.test.ts#forwards (segs, slack, firstContentSi, lastDrawnSi) positionally',
  },
  description: 'Keep leading whitespace used as a first-line text indent fixed while distributing justified-line slack across content in a left-to-right line.',
});

export const WORD_JUSTIFIED_CANDIDATE_SEPARATOR_FIT = defineCompatibilityRule({
  id: 'word-justified-candidate-separator-fit',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/justify-shrink-overshoot.test.ts#counts a candidate trailing space when the prospective line will justify',
  },
  description: 'On a full paragraph-width line that will be fully justified, include the candidate word separator in its wrap-fit width; lines narrowed by DrawingML wrap exclusions retain collapsible line-end separator fit behavior.',
});

export const WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS = defineCompatibilityRule({
  id: 'word-overflow-punctuation-language-sets',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OE376] §2.1.56',
  },
  description: 'Apply the language-specific punctuation sets documented for Word in [MS-OE376] §2.1.56 to Chinese, Japanese, and Korean language runs, and let overflowPunct override kinsoku when both rules affect the same character. When an effective East Asian language is absent, content that actually selects the East Asian script path retains the union as a bounded fallback.',
});

export const WORD_OVERFLOW_PUNCTUATION_LATIN_PARENT_RUN = defineCompatibilityRule({
  id: 'word-overflow-punctuation-latin-parent-run',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'latin-ascii-punctuation-advance-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Although [MS-OE376] §2.1.56 presents concrete punctuation sets by CJK language, Office-produced Latin boundary controls admit `.` and `,` beyond the ordinary word-fit extent. Controls at 9 and 14 points in Calibri and Arial instead wrap `)` and `}` with the word at the measured advance boundary; 10-point controls also wrap `!`, `%`, `:`, `;`, `>`, `?`, and `]`. Complex-script segments with no explicit RTL-primary bidi language retain their separate observed fallback; an explicit `ar-SA` bidi language is its counterexample.',
});

export const WORD_FULL_WIDTH_CHARACTER_SPACING_SCOPE = defineCompatibilityRule({
  id: 'word-full-width-character-spacing-scope',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OE376] §2.1.562',
  },
  description: 'Interpret ST_CharacterSpacing as applying whitespace compression to full-width punctuation characters. This rule establishes only which characters are eligible; it does not define a universal compression amount.',
});

export const WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL = defineCompatibilityRule({
  id: 'word-japanese-punctuation-compression-cell',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'japanese-fullwidth-punctuation-compression-cell',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In the observed Japanese compatibility matrix, 、。 ，． and the closing forms 」』】）］｝ on a full ideographic-cell advance retain at least half of that cell. U+3017 and full-width !/? remain full-cell. A fontTable w:pitch value classifies the authored face for font selection; it is not a switch for document-level characterSpacingControl. Punctuation that the selected face already exposes on a smaller proportional advance is retained as measured rather than compressed a second time. Tight adjacent glyph ink can require a larger retained extent to prevent collision. This is an Office-observed compression amount, not a normative interpretation of ST_CharacterSpacing.',
});

export const WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY = defineCompatibilityRule({
  id: 'word-authored-character-spacing-pitch-priority',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'authored-character-spacing-punctuation-pitch',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When a run authors a positive w:spacing character pitch, Word preserves that expanded pitch instead of additionally applying the document-level punctuation whitespace compression. Omitted, zero, or overlapping run spacing leaves characterSpacingControl active.',
});

export const WORD_SOURCE_RUN_SPACE_SEQUENCE = defineCompatibilityRule({
  id: 'word-source-run-space-sequence',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'source-run-space-sequence-wrap-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'At a source-run boundary, Word keeps a space-only continuation attached when the preceding run already ends in a space. A single leading space in a distinct run without a preceding space remains a break opportunity. This isolates source-boundary compatibility from the ordinary UAX #14 LB7 handling within one authored run.',
});

export const WORD_BALANCED_CONSECUTIVE_SPACE_CELL = defineCompatibilityRule({
  id: 'word-balanced-consecutive-space-cell',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'single-double-byte-width-space-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With ECMA-376 §17.15.3.3 balanceSingleByteDoubleByteWidth enabled, Word retains one ordinary inter-word U+0020 at its proportional natural advance, while a sequence of two or more authored U+0020 spaces advances each space by half of the selected East-Asian ideographic cell. The observed matrix covers one, two, four, and eight spaces; same-run and source-run boundaries; proportional and fixed-pitch faces; linesAndChars with negative/zero charSpace; and a line-only grid.',
});

/** Compatibility projection governed by
 * {@link WORD_BALANCED_CONSECUTIVE_SPACE_CELL}. */
export function wordBalancedConsecutiveSpaceCellApplies(spaceCount: number): boolean {
  return Number.isInteger(spaceCount) && spaceCount >= 2;
}

/** Evidence-bounded grid scope governed by
 * {@link WORD_BALANCED_CONSECUTIVE_SPACE_CELL}. `snapToChars` has a separate
 * Microsoft-documented block/cell allocator and is outside this observation. */
export function wordBalancedSpaceCellAdjustmentApplies(
  gridType: string | null | undefined,
): boolean {
  return gridType !== 'snapToChars';
}

export const WORD_BALANCED_LINES_AND_CHARS_GRID_DELTA = defineCompatibilityRule({
  id: 'word-balanced-lines-and-chars-grid-delta',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'single-double-byte-width-grid-observation-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With balanceSingleByteDoubleByteWidth enabled on linesAndChars, Word applies half of the authored charSpace delta to ASCII SBCS text and to U+0020/U+3000 space characters, while applying the full delta to CJK ideographs and full-width ASCII forms. The Word-output evidence covers ASCII digits, letters, punctuation, spaces, CJK, full-width ASCII, mixed text, proportional/fixed-pitch faces, negative/zero/positive charSpace, and line-only controls. Non-ASCII high-ANSI and complex-script text are outside the observed matrix and retain the preexisting grid behavior.',
});

export const WORD_LATIN_INTERWORD_XAVG_FLOOR = defineCompatibilityRule({
  id: 'word-latin-interword-xavg-floor',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'latin-space-os2-xavg-one-twip-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For left-aligned homogeneous Latin words using one U+0020 separator and characterSpacingControl=compressPunctuation, Word reduces each natural inter-word space only as far as half the selected static face\'s positive OS/2 xAvgCharWidth. First-fit and beta-origin controls varied only U+0020 hmtx, then only xAvg, with fixed outlines and non-space advances at 8/11/16pt; a distinct Carlito outline provided a counterexample where natural U+0020 was narrower than the floor. Two-gap controls showed equal required-deficit allocation. A one-twip linesAndChars control preserved the xAvg-dependent deficit while moving both control and natural boundaries by the separate character-grid pitch. [MS-OE376] §2.1.472 documents lineWrapLikeWord6 as an explicit uncompressed-fit override. Mixed faces, explicit run spacing or scaling, justification, and snapToChars remain outside the observed scope.',
});

export const WORD_COMPRESSED_SPACE_LINE_FIT = defineCompatibilityRule({
  id: 'word-compressed-space-line-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'mixed-script-space-fit-1660-matrix',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS 27.0',
  },
  description: 'Issue #1660 controls (1773 Word-exported fixed-cell lines: 576 coarse, 595 fine one-twip, 602 hypothesis-targeted) establish U+0020 fitting on lines mixing East Asian and Latin text, which WORD_LATIN_INTERWORD_XAVG_FLOOR (homogeneous Latin lines, unchanged) never compresses. In compatibility modes 12 and 14 with characterSpacingControl compressPunctuation or compressPunctuationAndJapaneseKana, every U+0020 on such a line shrinks by the same amount (single, consecutive and source-run-split spaces alike) while ideographs, kana and other glyphs keep their natural advances; enableOpenTypeFeatures and explicit w:kern thresholds do not gate it, and left, both and distribute alignment behave alike. Each space keeps at least min(xAvgCharWidth / 2, font size / 4) of its own face: Arial and Times New Roman spaces follow half their OS/2 xAvgCharWidth, while BIZ UDGothic, Meiryo and Yu Gothic (xAvgCharWidth 0.84-0.96 em) and BIZ derivatives with only post.isFixedPitch or only the far-east code-page bits changed all stop at a quarter em, rejecting fixed-pitch, code-page and weighted-lowercase-average selectors. A candidate whose last character before trailing closing punctuation is East Asian is admitted only while its natural overflow, measured without that trailing punctuation, is at most half the font size: a final ideograph or half-width kana admits 0.5 em of overflow at 8.5/10.5/16 pt, a final kana followed by a full-width closing parenthesis keeps that limit before the parenthesis (which keeps half its cell at the line end), and a final Latin word has no limit beyond the space floors. characterSpacingControl omitted or doNotCompress and compatibility mode 15 keep natural spaces on mixed lines; lineWrapLikeWord6 ([MS-OE376] §2.1.472) and an omitted compatibility mode are outside the projection. Mode-15 justified mixed lines (Word admits 2.60pt of overflow over four 4.25pt spaces but not 2.65pt) are unexplained and keep natural fitting.',
});

/** Document gate of {@link WORD_COMPRESSED_SPACE_LINE_FIT}: an authored
 * compatibility mode below 15 (12 and 14 measured) with a compressing
 * characterSpacingControl. The line breaker further restricts it to lines
 * holding East Asian text. */
export function wordCompressedSpaceLineFitApplies(
  compatibilityMode: number | undefined,
  characterSpacingControl: string | undefined,
): boolean {
  return compatibilityMode !== undefined && compatibilityMode < 15
    && (characterSpacingControl === 'compressPunctuation'
      || characterSpacingControl === 'compressPunctuationAndJapaneseKana');
}

/** Per-space minimum advance of {@link WORD_COMPRESSED_SPACE_LINE_FIT} and
 * {@link WORD_LATIN_INTERWORD_XAVG_FLOOR}, in the caller's unit. */
export function wordCompressedSpaceFloor(fontSize: number, averageWidthRatio: number): number {
  return Math.min((fontSize * averageWidthRatio) / 2, fontSize / 4);
}

/** Natural-overflow limit of {@link WORD_COMPRESSED_SPACE_LINE_FIT} for a
 * candidate ending (before trailing closing punctuation) in an East Asian
 * character; undefined means no limit beyond the space floors. */
export function wordCompressedSpaceEastAsianOverflowLimit(
  lastCharacterIsEastAsian: boolean,
  fontSize: number,
): number | undefined {
  return lastCharacterIsEastAsian ? fontSize / 2 : undefined;
}

export const WORD_IDEOGRAPHIC_SPACE_LINE_END_ALLOWANCE = defineCompatibilityRule({
  id: 'word-ideographic-space-line-end-allowance',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'ideographic-space-line-end-count-and-run-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Word keeps a single U+3000 immediately following visible East-Asian text on that line when the visible glyph is force-fitted into a narrow table cell. A paragraph-final sequence of two or more U+3000 characters remains authored width-bearing content and may form blank continuation lines. The observed matrix covers single and trailing multiple spaces, linesAndChars with negative/positive charSpace, line-only grids, and snapToGrid opt-out.',
});

/** Compatibility projection governed by
 * {@link WORD_IDEOGRAPHIC_SPACE_LINE_END_ALLOWANCE}. */
export function wordIdeographicSpaceLineEndAllowanceCount(
  hasEastAsianVisiblePredecessor: boolean,
  consecutiveSpaceCount: number,
): 0 | 1 {
  return hasEastAsianVisiblePredecessor && consecutiveSpaceCount === 1 ? 1 : 0;
}

/** Compatibility projection governed by
 * {@link WORD_BALANCED_LINES_AND_CHARS_GRID_DELTA}. Script-slot acquisition
 * has already separated ordinary East-Asian and ASCII SBCS text; the explicit
 * space branch retains Word's observed U+3000 exception without reclassifying
 * other East-Asian glyphs. Non-ASCII high-ANSI/complex text stays outside the
 * observed projection. */
export function wordBalancedLinesAndCharsGridDeltaFactor(
  text: string,
  script: 'ascii' | 'highAnsi' | 'eastAsia' | 'complexScript',
): 0.5 | 1 | undefined {
  if (script === 'complexScript') return undefined;
  const spaceOnly = text.length > 0 && [...text].every(
    (character) => character === ' ' || character === '\u3000',
  );
  if (spaceOnly) return 0.5;
  if (script === 'eastAsia') return 1;
  return [...text].every((character) => (character.codePointAt(0) ?? 0x80) <= 0x7f)
    ? 0.5
    : undefined;
}

/** Compatibility projection governed by
 * {@link WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL}. */
export function wordJapanesePunctuationRetainedExtentPt(input: Readonly<{
  punctuationAdvancePt: number;
  punctuationInkEndPt: number;
  ideographicCellAdvancePt: number;
}>): number {
  const advancePt = Math.max(0, input.punctuationAdvancePt);
  const cellAdvancePt = Math.max(0, input.ideographicCellAdvancePt);
  if (advancePt < cellAdvancePt) return advancePt;
  return Math.min(
    advancePt,
    Math.max(
      0,
      input.punctuationInkEndPt,
      cellAdvancePt / 2,
    ),
  );
}

/** Compatibility projection governed by
 * {@link WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY}. */
export function wordDocumentCharacterCompressionApplies(
  authoredCharacterSpacingPt: number | undefined,
): boolean {
  return authoredCharacterSpacingPt === undefined || authoredCharacterSpacingPt <= 0;
}

/** Compatibility projection governed by {@link WORD_SOURCE_RUN_SPACE_SEQUENCE}. */
export function wordSourceRunSpaceContinuesSequence(
  previousText: string,
  currentText: string,
): boolean {
  return previousText.endsWith(' ') && currentText.startsWith(' ');
}

const WORD_OVERFLOW_PUNCTUATION = {
  ja: new Set([...',.’”、。」』】），．］｝｡､']),
  zhHans: new Set([...`!%),.:;>?]}¢°·ˇ’”‰′″℃∶、。〃〉》」』】〗〕〞﹚﹜﹞！＂％＇），．：；？］｝￠`]),
  zhHant: new Set([...`!),.:;?]}’”′、。〉》」』】〕〞﹚﹜﹞！），．：；？］｝`]),
  ko: new Set([...`!%),.:;?]}¢°’”′″℃〉》」』】〕！％），．：；？］｝￠`]),
} as const;
const ALL_WORD_OVERFLOW_PUNCTUATION = new Set([
  ...WORD_OVERFLOW_PUNCTUATION.ja,
  ...WORD_OVERFLOW_PUNCTUATION.zhHans,
  ...WORD_OVERFLOW_PUNCTUATION.zhHant,
  ...WORD_OVERFLOW_PUNCTUATION.ko,
]);
const LATIN_WORD_OVERFLOW_PUNCTUATION = new Set(['.', ',']);

export const RTL_PRIMARY_SUBTAGS = new Set([
  'ar', 'fa', 'ur', 'he', 'iw', 'yi', 'ji', 'ps', 'sd', 'ug', 'dv', 'syr', 'ckb',
]);

/** Compatibility projection governed by
 * {@link WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS}. */
export function wordIsOverflowPunctuation(
  character: string,
  language: string | undefined,
  parentRunHasEastAsianText = false,
  parentRunHasLatinText = false,
  parentRunHasComplexScriptText = false,
  bidiLanguage?: string,
): boolean {
  const normalized = language?.toLowerCase();
  if (normalized?.startsWith('ja')) return WORD_OVERFLOW_PUNCTUATION.ja.has(character);
  if (normalized?.startsWith('ko')) return WORD_OVERFLOW_PUNCTUATION.ko.has(character);
  if (normalized?.startsWith('zh')) {
    return (/(?:^|-)(?:tw|hk|mo)(?:-|$)|hant/u.test(normalized)
      ? WORD_OVERFLOW_PUNCTUATION.zhHant
      : WORD_OVERFLOW_PUNCTUATION.zhHans).has(character);
  }
  // ECMA-376 §17.3.1.21 is script-neutral. Although [MS-OE376] §2.1.56
  // describes Word's concrete sets as CJK-language behavior, Office-produced
  // boundary controls hang `.` and `,` in Latin parent runs; 9/14pt Calibri
  // and Arial counterexamples wrap `)` and `}` at the normal measured advance
  // boundary. Complex-script production controls with an absent bidi tag
  // hang `.`, `:`, `)` and `>`; an explicit RTL-primary tag such as
  // `ar-SA` is the counterexample. A CJK run carrying an inherited
  // non-CJK language also uses the union, because its actual script route is
  // more authoritative than that inherited language.
  const bidiPrimary = bidiLanguage?.split('-')[0].toLowerCase();
  const explicitRtlBidi = bidiPrimary != null && RTL_PRIMARY_SUBTAGS.has(bidiPrimary);
  const observedComplexFallback = parentRunHasComplexScriptText && !explicitRtlBidi;
  if (parentRunHasEastAsianText || observedComplexFallback) {
    return ALL_WORD_OVERFLOW_PUNCTUATION.has(character);
  }
  return parentRunHasLatinText && LATIN_WORD_OVERFLOW_PUNCTUATION.has(character);
}

/** Compatibility projection governed by {@link WORD_JUSTIFIED_CANDIDATE_SEPARATOR_FIT}. */
export function wordCandidateFitWidthPx(input: Readonly<{
  widthPx: number;
  trailingSpacePx: number;
  lineWillJustify: boolean;
  wrapNarrowed?: boolean;
}>): number {
  return input.lineWillJustify && input.wrapNarrowed !== true
    ? input.widthPx
    : input.widthPx - input.trailingSpacePx;
}

export const WORD_RUBY_PARAGRAPH_UNIFORM_LINE_ADVANCE = defineCompatibilityRule({
  id: 'word-ruby-paragraph-uniform-line-advance',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/paragraph-measure.test.ts#uses one uniform snapped advance for every line in a ruby paragraph',
  },
  description: 'Every line in a ruby-bearing paragraph uses the paragraph-wide maximum snapped line advance so its baseline rhythm remains uniform.',
});

export const WORD_FIT_TEXT_INTER_CHARACTER_EXPANSION = defineCompatibilityRule({
  id: 'word-fit-text-inter-character-expansion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/fit-text.test.ts#distributes (val − Σnatural)/(n−1) as the inter-character gap, no trailing gap',
  },
  description: 'Expand a multi-character fitText region to its authored width by distributing the residual evenly across interior character gaps.',
});

export const WORD_CJK_BOTH_INTER_CHARACTER_EXPANSION = defineCompatibilityRule({
  id: 'word-cjk-both-inter-character-expansion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/text-distribute.test.ts#§17.18.44: fills a wrapped pure-CJK line via inter-CJK pitch (expansion default)',
  },
  description: 'Treat inter-CJK boundaries as eligible inter-word gaps when expanding a non-final both-justified line that contains no spaces.',
});

export const WORD_THAI_DISTRIBUTE_CLUSTER_POLICY = defineCompatibilityRule({
  id: 'word-thai-distribute-cluster-policy',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/thai-distribute.test.ts#fills non-final lines to the right margin under thaiDistribute',
  },
  description: 'Expand non-final thaiDistribute lines at Thai grapheme-cluster boundaries while retaining a natural-width final line.',
});

export const WORD_NUMERIC_DECIMAL_TAB_INFERENCE = defineCompatibilityRule({
  id: 'word-numeric-decimal-tab-inference',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/decimal-tab-autoalign.test.ts#right-aligns numbers of different digit counts at the decimal tab',
  },
  description: 'Right-align an otherwise tab-less numeric paragraph at its leading decimal tab while leaving non-numeric and no-decimal-tab paragraphs unchanged.',
});

export const WORD_NUMBERING_MARKER_OVERFLOW_TAB_ADVANCE = defineCompatibilityRule({
  id: 'word-numbering-marker-overflow-tab-advance',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/numbered-marker-tab-advance.test.ts#advances the body past the marker to the next tab stop, not onto indentLeft',
  },
  description: 'When a numbering marker overruns its hanging-indent budget, advance the body to the next reachable tab stop beyond the marker edge.',
});

export const WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB = defineCompatibilityRule({
  id: 'word-numbering-suffix-coincident-list-tab',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/numbering-marker.test.ts#keeps a suffix tab on the list stop coincident with the marker end',
  },
  description: 'For the tab synthesized by a numbering suffix, accept an authored numeric list tab coincident with the shaped marker end instead of advancing to the next automatic tab stop.',
});

export const WORD_NUMBERING_MARKER_PARAGRAPH_MARK_FALLBACK = defineCompatibilityRule({
  id: 'word-numbering-marker-paragraph-mark-fallback',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'numbering-marker-paragraph-mark-formatting',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When numbering-level rPr omits a marker formatting axis, Word takes that axis from the effective paragraph-mark rPr rather than a content run. A numbering-level concrete value or explicit auto remains authoritative, and body and text-box stories use the same cascade.',
});

/** Compatibility projection governed by {@link WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB}. */
export function wordNumberingSuffixAcceptsCoincidentListTab(
  markerEndPt: number,
  stop: Readonly<{ pos: number; alignment: string }>,
): boolean {
  return stop.alignment === 'num' && Math.abs(stop.pos - markerEndPt) <= 1e-6;
}

export const WORD_TAB_STOP_PAGE_EDGE_CLAMP = defineCompatibilityRule({
  id: 'word-tab-stop-page-edge-clamp',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/rtl-tab-stops.test.ts#pins a page number to the left text margin when the stop is past it',
  },
  description: 'Clamp content assigned to a tab stop beyond the trailing text edge back onto that edge instead of placing ink outside the page content band.',
});

export const WORD_DICTIONARY_SEA_ATOMIC_CHUNK = defineCompatibilityRule({
  id: 'word-dictionary-sea-atomic-chunk',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/sea-justified-fit.test.ts#Rule 2: a no-space chunk that fits a full line moves whole instead of splitting',
  },
  description: 'Move a glued dictionary Southeast-Asian chunk to a fresh line whole when it fits that full line, using dictionary breaks only when the chunk itself is overlong.',
});

export const WORD_OVERLONG_TOKEN_EMERGENCY_BREAK = defineCompatibilityRule({
  id: 'word-overlong-token-emergency-break',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/run-inline-formatting.test.ts#breaks a no-space token wider than the line at the character level',
  },
  description: 'Emergency-break an overlong token at grapheme-safe character boundaries on an empty line so the complete token remains inside the content band.',
});

export const WORD_EXTERNAL_LINK_SYNTAX_BREAKS = defineCompatibilityRule({
  id: 'word-external-link-syntax-breaks',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'external-link-syntax-formatting-seam-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Treat readable separators in the path and query of displayed external URLs as line-break opportunities, while keeping the scheme and authority intact and preserving authored no-break hyphens and grapheme clusters.',
});

/** Compatibility projection governed by {@link WORD_EXTERNAL_LINK_SYNTAX_BREAKS}.
 * `graphemeBoundaries` and `authoredNoBreakOffsets` are UTF-16 offsets in the
 * complete displayed link token, not in an individual formatting run. */
export function wordExternalLinkSyntaxBreakOffsets(
  text: string,
  graphemeBoundaries: ReadonlySet<number>,
  authoredNoBreakOffsets: ReadonlySet<number>,
): readonly number[] {
  const scheme = /^[A-Za-z][A-Za-z0-9+.-]*:\/\//u.exec(text);
  if (!scheme) return [];
  const authorityStart = scheme[0].length;
  const authorityEndCandidate = text.slice(authorityStart).search(/[/?#]/u);
  const authorityEnd = authorityEndCandidate < 0
    ? text.length
    : authorityStart + authorityEndCandidate;
  const offsets: number[] = [];
  for (let index = authorityEnd; index < text.length; index += 1) {
    const character = text[index]!;
    const offset = index + 1;
    const readable =
      (character === '/' && index > authorityEnd)
      || character === '-'
      || character === '?'
      || character === '&';
    if (
      readable
      && graphemeBoundaries.has(offset)
      && !authoredNoBreakOffsets.has(offset)
    ) offsets.push(offset);
  }
  return offsets;
}

export const WORD_RUN_VERTICAL_ALIGN_BASELINE_SHIFT = defineCompatibilityRule({
  id: 'word-run-vertical-align-baseline-shift',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/run-char-metrics-render.test.ts#w:vertAlign raises superscript, lowers subscript, and leaves ordinary baselines unchanged',
  },
  description: 'Retain the established run-level baseline displacement for vertically aligned text: superscript rises by 0.35 of its authored font size and subscript falls by 0.15, while the separately authored w:position remains additive.',
});

export const WORD_UNIFORM_RUN_POSITION_LEADING = defineCompatibilityRule({
  id: 'word-uniform-run-position-leading',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'uniform-run-position-leading',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When every metric-bearing item on a line has the same non-zero w:position, Word places the allocated line box around the positioned glyphs. A line containing a differently-positioned item retains the full relative displacement; automatic line allocation is handled separately.',
});

/** Paint-relative baseline position governed by
 * {@link WORD_UNIFORM_RUN_POSITION_LEADING}. */
export function wordUniformRunPositionPaintPt(
  authoredPositionPt: number,
  commonLinePositionPt: number,
): number {
  return commonLinePositionPt === 0
    ? authoredPositionPt
    : authoredPositionPt - commonLinePositionPt / 2;
}

/** Compatibility projection governed by
 * {@link WORD_RUN_VERTICAL_ALIGN_BASELINE_SHIFT}. */
export function wordRunVerticalAlignRaisePt(
  verticalAlign: string | null | undefined,
  authoredFontSizePt: number,
): number {
  if (verticalAlign === 'super') return authoredFontSizePt * 0.35;
  if (verticalAlign === 'sub') return -authoredFontSizePt * 0.15;
  return 0;
}

export const WORD_FAR_EAST_SINGLE_LINE_FACTOR = OFFICE_FAR_EAST_SINGLE_LINE_FACTOR;

export const WORD_INLINE_PICTURE_AUTO_LEADING = defineCompatibilityRule({
  id: 'word-inline-picture-auto-leading',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'inline-picture-auto-leading-height-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'For an inline picture with automatic spacing at or above one line, add the authored leading of the selected text/paragraph-mark face to the natural picture/text baseline union, rather than multiplying the picture height. Image-only, text-only and alternating 28.35pt-picture controls at line=240/259 and exact=259 distinguish the two advances; earlier Word controls covered picture heights 5–255pt at auto multiples 1, 1.079 and 1.15. Exact spacing and non-picture lines retain their own allocation.',
});

/** Word-observed inline-picture projection of ECMA-376 §17.3.1.33. */
export function wordInlinePictureAutoLineHeightPx(
  naturalUnionPx: number,
  textSinglePx: number,
  multiple: number,
): number {
  void WORD_INLINE_PICTURE_AUTO_LEADING;
  return naturalUnionPx + textSinglePx * (multiple - 1);
}

/** Word for Mac 16.112.4 automatic line allocation observed with independently
 * varied synthetic fonts. With OS/2 code-page bits 17–20, the line box is
 * 1.3 × the hhea ascent/descent box and the added 0.3 is split equally above
 * and below. With those bits clear, signed hhea lineGap is placed above the
 * baseline. A Latin cmap alone did not change the code-page class. ECMA-376
 * §17.3.1.33 specifies the spacing multiplier, not this font-table choice.
 * Unsupported or malformed geometry leaves measurement in charge. */
export function wordOpenTypeAutoLineRatios(metrics: Readonly<{
  unitsPerEm: number;
  hheaAscent: number;
  hheaDescent: number;
  hheaLineGap: number;
  farEastCodePage: boolean;
}>): Readonly<{
  lineHeightRatio: number;
  designAscentRatio: number;
  designDescentRatio: number;
}> | null {
  return officeOpenTypeAutoLineRatios(metrics);
}

export function wordEastAsianGridLineCells(
  naturalHeightPx: number,
  pitchPx: number,
): number {
  return pitchPx > 0 ? Math.max(1, Math.ceil(naturalHeightPx / pitchPx)) : 1;
}

export function wordFarEastSingleLinePx(
  intendedSinglePx: number,
  emPx: number,
): number {
  return intendedSinglePx > 0
    ? intendedSinglePx
    : emPx * WORD_FAR_EAST_SINGLE_LINE_FACTOR;
}

/** Compatibility projection governed by
 * {@link WORD_USE_FE_LAYOUT_INHERITED_GRID_MINIMUM}. */
export function wordUseFeLayoutInheritedGridHeightPx(
  allocatedCellHeightPx: number,
  pitchPx: number,
  inheritedMultiple: number,
): number {
  return Math.max(allocatedCellHeightPx, pitchPx * inheritedMultiple);
}

export function wordGridAtLeastLineHeightPx(
  naturalPx: number,
  authoredMinimumPx: number,
  gridMinimumPx: number,
): number {
  return Math.max(naturalPx, authoredMinimumPx, gridMinimumPx);
}

export function wordDegenerateLineSpacingIsSingle(
  rule: string,
  value: number,
): boolean {
  return (rule === 'exact' || rule === 'auto') && value <= 0;
}

export function wordAutoMultipleCenterBoxPx(
  autoMultiple: boolean,
  compressedAuto: boolean,
  glyphNaturalPx: number,
  intendedSinglePx: number,
  lineHeightPx: number,
): number {
  return autoMultiple && !compressedAuto
    ? Math.max(glyphNaturalPx, intendedSinglePx)
    : lineHeightPx;
}

export function wordVisibleLineMetricPx(
  reservedMetricPx: number,
  visibleMetricPx: number | undefined,
): number {
  return visibleMetricPx ?? reservedMetricPx;
}

export function wordFirstJustifiedContentSegment(
  segments: readonly object[],
  bidi: boolean,
): number {
  if (bidi) return 0;
  for (let index = 0; index < segments.length; index += 1) {
    const segment = segments[index];
    const text = 'text' in segment && typeof segment.text === 'string'
      ? segment.text
      : undefined;
    if (text === undefined || /\S/.test(text)) return index;
  }
  return 0;
}

export function wordRubyUniformLineHeightPx(
  hasRuby: boolean,
  lineHeightsPx: readonly number[],
): number {
  return hasRuby ? Math.max(0, ...lineHeightsPx) : 0;
}

export const WORD_LATIN_DESIGN_GRID_CELLS = defineCompatibilityRule({
  id: 'word-latin-design-grid-cells',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'float-grid-picture-origin',
    application: 'Microsoft Word', version: '16.113.2', platform: 'macOS 27.0' },
  description: 'Issue #1674 modes 14/15 controls reserve 20/40/40pt for 10/20/30pt Arial single lines on a 20pt grid, including a preceding 20pt line. No-grid and snap-off controls retain natural advances. Apply whole-cell counting only to visible text with an admitted non-Far-East reference design profile. Far-East reference faces, native fallback boxes, empty marks, ruby, explicit multiples and exact/atLeast spacing retain their established paths. Later independent-font and anchor-free controls established WORD_SPECIFIED_TEXT_LINE_BOX for its separately gated ordinary Latin class; every other exact/atLeast class retains the established placement.',
});

export function wordLatinDesignGridSingleHeight(natural: number, pitch: number, admittedDesign: number): number {
  // The admitted face controls only its own whole-cell reserve. Other fonts,
  // inline objects and baseline displacements still own their natural extent
  // (§17.3.1.33); admission of one run must not shrink any peer's line box.
  return Math.max(natural, admittedDesign > 0
    ? Math.max(1, Math.ceil(admittedDesign / pitch)) * pitch : pitch);
}

export const WORD_SPECIFIED_TEXT_LINE_BOX = defineCompatibilityRule({
  id: 'word-specified-text-line-box',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'specified-spacing-anchor-free-host-font-grid-matrix',
    application: 'Microsoft Word',
    version: 'export build unrecorded (16.113.3 installed at measurement)',
    platform: 'macOS 27.0',
  },
  description: 'Issue #1674: 636 printing-PDF pages, modes 14/15, no grid and 20pt lines/linesAndChars grids, Arial/Times New Roman/MS Mincho/Yu Mincho/Verdana/Calibri Latin glyphs. Exact 12/20/24/30pt lines put the baseline four fifths down the authored line, independent of the 10pt visible face or zero/one/two hosts; 30pt visible-text and 30pt-host diagnostics distinguish font ascent, centering and host-height hypotheses. Off-grid atLeast 24/40pt retains the normal design descent and places added leading before the normal line box. Grid-atLeast remains unresolved and is excluded: 40pt minima keep anchor advances at 40pt but print successive text baselines 40.08pt apart; a centered normal-cell projection leaves an unexplained residual. Preserve the entire grid-atLeast class, including host allocation. In the admitted off-grid class a floating host does not enlarge a visible text line. Explicit auto=1 diagnostics on an active grid use the same normal grid box; off-grid diagnostics already agree with the established path, which is preserved. PDF coordinates/font sizes are quantized to 0.24pt; this is export precision, not a layout correction. The exact-height evidence interval is 12 through 30pt; smaller captions do not isolate baseline placement from accumulated printing cadence, and heights outside this interval retain the established path. Scope is homogeneous regular unpositioned undecorated Latin text and matching paragraph-mark face/size with an admitted authored reference profile and a native or positively loaded installed Office local() face. Registered local aliases retain the authored family from the resolver; Application-provided SFNT, CSS, embedded, Google and substitute resources remain outside the evidence. Float-exclusion-conditioned origins are also excluded: a narrow side-gap Office counterexample has a different physical partition, so this round does not isolate its baseline from the wrap origin. Admission is paragraph-wide: a mixed-script or decorated continuation paragraph cannot switch contracts on its ASCII-only undecorated physical lines. Mixed scripts/font axes/faces/sizes, differing paragraph marks, decorated/hyperlink text, transformed text, ruby, inline objects, numbering and resource/fallback faces are unmeasured and retain the established path.',
});

/** Word observation, not the normative centered-text rule: ECMA-376
 * §17.3.1.33 defines exact/atLeast units and describes bottom placement/clipping
 * for too-small lines and centering for too-large lines. [MS-OI29500]
 * §2.1.60 discusses style-hierarchy spacing but supplies no baseline formula.
 * The measured Word printing output instead preserves a fixed 4:1 exact baseline partition,
 * including oversized text that visibly overhangs without top clipping.
 * This partition is supported by the complete independent-font/host matrix,
 * not a fitted ascent or a sample-specific offset. §17.6.5 makes exact spacing
 * override the grid; non-exact normal boxes retain whole-cell grid leading.
 * Unsupported classes are declined by specifiedTextLineMetrics beside its gate. */
export function wordSpecifiedTextLineMetrics(input: Readonly<{
  rule: string; value: number; descentPt: number;
  singlePt: number; pitchPt: number | null;
}>): Readonly<{ advancePt: number; baselineOffsetPt: number }> {
  void WORD_SPECIFIED_TEXT_LINE_BOX;
  if (input.rule === 'exact') {
    return { advancePt: input.value, baselineOffsetPt: input.value * 4 / 5 };
  }
  const normalPt = input.pitchPt !== null && input.pitchPt > 0
    ? Math.ceil(input.singlePt / input.pitchPt) * input.pitchPt
    : input.singlePt;
  const advancePt = input.rule === 'atLeast' ? Math.max(input.value, normalPt) : normalPt;
  return {
    advancePt,
    baselineOffsetPt: advancePt - (normalPt - input.singlePt) / 2 - input.descentPt,
  };
}
