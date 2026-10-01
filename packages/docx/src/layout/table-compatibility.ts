import { defineCompatibilityRule } from './compatibility.js';
import type { ParagraphLayoutSource } from './text.js';
import type { TableColumnLayoutInput, LayoutRect } from './types.js';

export const WORD_ROTATED_CELL_AUTO_ROW_WRAP = defineCompatibilityRule({
  id: 'word-rotated-cell-auto-row-wrap',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'rotated-cell-row-height-direction-rule-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'In a fixed-width table, btLr and tbRl cells with 1, 4, or 8 glyphs keep the same automatic row height as a one-line horizontal neighbor by wrapping into additional columns. Two paragraphs and an explicit line break also keep that height when their columns fit. A 20pt top/bottom margin sum adds 20pt, revealing a line-box minimum even when the glyph advance is narrower; a 60pt atLeast minimum and a five-line horizontal neighbor govern their rows; exact remains authored. A horizontal text-direction control keeps its ordinary line. These Word PDF observations cover short text fitting across a 225pt cell in compatibility mode 14; larger content and other compatibility modes are not Office-verified by this matrix.',
});

export const WORD_AUTOFIT_EMPTY_PARAGRAPH_CONTENT_WIDTH = defineCompatibilityRule({
  id: 'word-autofit-empty-paragraph-content-width',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-empty-paragraph-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For table AutoFit content width, Word gives an empty unnumbered paragraph no intrinsic content width regardless of effective right, left, first-line, or hanging indentation. Cell margins still contribute, while whitespace, non-breaking space, visible text, and numbering remain content-bearing controls.',
});

export const WORD_AUTOFIT_OUTER_CELL_MARGIN_BAND = defineCompatibilityRule({
  id: 'word-autofit-outer-cell-margin-band',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-nowrap-identical-grid-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'For top-level ordinary AutoFit tables, Word permits a saved grid to extend beyond the text band by the resolved outer cell margins. Two-cell controls with 0, 2.7, and 5.4pt margins, dxa 250/300pt and auto cell widths, short/long text, and noWrap on/off establish that allowance. A Word-produced table with 5.4pt margins but a saved grid exactly at the text band retains that narrower grid on current Word PDF export. A nested table with saved grid overhang retains its containing-cell width rather than receiving the page-table allowance. Auto-width tables are now governed by WORD_AUTOFIT_LEADING_INDENT_BAND (compatibility-mode dependent); for a preferred-width table only a top-level outer-margin overhang already present in tblGrid increases the physical ceiling. Skipped outer grid tracks, leading-margin placement, vertical text, and floating tables remain outside this observation.',
});

export const WORD_AUTOFIT_LEADING_INDENT_BAND = defineCompatibilityRule({
  id: 'word-autofit-leading-indent-band',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-nowrap-page-boundary-compat-mode-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'For top-level left-justified auto-width AutoFit tables, the measured ceiling is text band - tblInd + outer cell margins in mode 14 (and with the mode omitted), and text band - tblInd in mode 15. The physical page is not a ceiling. Other explicit compatibility modes retain the established solver contract because they are unmeasured. The original controls cover a 468pt page, 0/18pt side margins, 5.4pt outer cell margins, indents of +5.4/+36/-36pt, wide and band-minus-indent grids, and noWrap on/off. Eight additional wrapping controls reverse the content distribution, vary an omitted first tcW versus a 100.55pt dxa first tcW, and vary grid totals of 426.6/432pt on a 432pt band with +5.4pt indent. Neither that cell preference nor the grid-total equality changes the resulting column starts or line partitions within a mode. The second cell remains auto throughout. An all-dxa multirow mode-15 table also agrees with this ceiling: Word PDF cell clipping bounds constrain its width to approximately 482pt, and the fitted 481.9pt width reproduces both inspected two-line paragraphs with the same Calibri face. Substituting a different face can move those line boundaries; retaining the old wider table would hide that font dependency. Four centered/right controls at that same positive indent and band-minus-indent grid match a 437.4pt mode-14 ceiling and a full-band 432pt mode-15 ceiling. In mode 14 those four controls cannot distinguish band-minus-indent-plus-margins from saved-grid-plus-margins. Other nonleading alignment/grid/indent combinations remain unverified and retain their established solver contract; the new nonleading ceiling is gated to grid-plus-indent equaling the band, where both mode-14 hypotheses agree. These borderless PDFs establish column starts and line partitions. Their integer-point clipping rectangles bound cell/table extents but do not establish an exact trailing table edge. Absolute leading-origin differences remain separate from width fitting. Floating, nested, vertical-text and preferred-table-width tables are outside this rule.',
});

export const WORD_AUTOFIT_CONTENT_COLUMN_GROWTH = defineCompatibilityRule({
  id: 'word-autofit-content-column-growth',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-nowrap-page-boundary-compat-mode-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'In the measured top-level auto-width AutoFit tables, a column without preferred cells grows beyond its initial tblGrid toward maximum content width, up to the occurrence ceiling. Two unpreferred tracks also release saved width beyond a short column maximum. Twelve simultaneous-growth controls in modes 14/15 start at 25/25pt and use breakable ii tokens in ratios 3:2, 1:1 and 5:1, with totals of 30 (nonbinding) and 90 (binding). Nonbinding columns reach their content maxima. Binding column widths match min_i + fraction * (max_i - min_i), where fraction = (ceiling - sum(min)) / sum(max - min), capped at 1. The unequal cases distinguish this from sharing deficits over saved widths, sharing absolute maxima, equal room and sequential growth. Intrinsic maxima must be unbounded by the text band. First-column differences are at most 0.228pt and both cells have identical Word line partitions. Only two nonspanning, wholly unpreferred tracks establish simultaneous sharing; larger grids, spans, and mixed preferred/unpreferred simultaneous growth retain the previous solver result. This is a measured content-interval rule, not an empirical scale factor or an observation of arbitrary multi-column growth.',
});

export const WORD_AUTOFIT_NOWRAP_AUTO_FORCED_FIT = defineCompatibilityRule({
  id: 'word-autofit-nowrap-auto-forced-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'autofit-nowrap-identical-grid-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'When a two-cell AutoFit table has one auto-width noWrap cell whose unbroken minimum alone exceeds the text band, Word scales the content-only widths to that band. It assigns the extra outer-margin width to each cell in proportion to the other cell content width. The observed matrix varies the ordinary cell text width from 5 to 85pt, outer margins from 0 to 5.4pt, and which side owns noWrap. Short noWrap text that fits the band is a counterexample and retains ordinary proportional fitting. More than two cells and spans are outside this compatibility claim.',
});

export function wordAutofitEmptyParagraphHasNoIntrinsicContent(
  paragraph: Pick<ParagraphLayoutSource, 'runs' | 'numbering'>,
): boolean {
  return paragraph.runs.length === 0 && paragraph.numbering == null;
}

export const WORD_EXACT_ROW_HEIGHT_BOTTOM_PADDING = defineCompatibilityRule({
  id: 'word-exact-row-height-bottom-padding',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.180(d)',
  },
  description: 'Word adds the largest bottom cell margin to an exact trHeight instead of treating that margin as part of the authored height.',
});

export const WORD_TABLE_BORDER_LAYER_CASCADE = defineCompatibilityRule({
  id: 'word-table-border-layer-cascade',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.169',
  },
  description: 'During per-side border acquisition, none falls through to a lower-precedence layer while nil remains authored and blocks fallback only on that side.',
});

export const WORD_SPACED_CELL_INSIDE_BORDER_CONFLICT = defineCompatibilityRule({
  id: 'word-spaced-cell-inside-border-conflict',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.136, 2.1.138',
  },
  description: 'With non-zero cell spacing, Word retains the narrow conditional tcBorders insideH/insideV conflict against the corresponding table inside border.',
});

export const WORD_TABLE_ORIGIN_COMPATIBILITY = defineCompatibilityRule({
  id: 'word-table-origin-compatibility',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-origin-compatibility-and-occupied-grid',
    application: 'Microsoft Word', version: '16.113.2', platform: 'macOS 27.0' },
  description: '528 ordinary-table controls cover modes 11/12/14/15, LTR/RTL, leading/center/end, omitted/zero/positive/negative indentation, explicit zero/asymmetric cell margins, cell overrides, first-row exceptions, fixed/AutoFit, nested contents, borderless and full-band tables. For top-level body tables with positive dxa cell preferences, explicit legacy horizontal margins and zero spacing (fixed auto/dxa table width or AutoFit dxa table width), modes 11/12/14 hang the first cell left margin: leading placement only with authored indentation and anchored to the first row, end placement per row. Center and end ignore indentation, as ECMA-376 §17.4.50 specifies (contrary to MS-OI29500 §2.1.155). Mode 15 removes the margin hang. Border rasterization is separate from nominal grid geometry. Spaced tables and wholly omitted horizontal margins remain unresolved and retain the previous origin and content-inset behavior; Nested table origins, logical start/end margin spellings, and other compatibility modes are outside the measured scope.',
});

export const WORD_FIXED_UNUSED_LEADING_GRID = defineCompatibilityRule({
  id: 'word-fixed-unused-leading-grid',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'table-origin-compatibility-and-occupied-grid',
    application: 'Microsoft Word', version: '16.113.2', platform: 'macOS 27.0' },
  description: 'The original fixed gridBefore controls in modes 11/12/14/15 and 48 follow-ups in modes 14/15 show that a leading track skipped by every row has zero layout width. Fixed auto table width retains the occupied 108/180pt tracks; dxa 324pt scales them to 121.5/202.5pt. A preceding row occupying the 36pt track preserves it and ordinary skipped-row placement. wBefore of 0/18/36pt has no effect in either class. This refines ECMA-376 §§17.4.15 and 17.18.87 only for top-level body nonspaced fixed tables with auto/dxa table width and positive dxa cell preferences; nested and nonbody stories, percentage table width, AutoFit, trailing skips, spans, and nonpositive or missing cell preferences remain outside the observation.',
});

/** The observation is deliberately restricted to the exported mode boundaries. */
export function wordMeasuredTableOriginMode(mode: number | undefined): boolean {
  return mode === 11 || mode === 12 || mode === 14 || mode === 15;
}

export const WORD_EXACT_ROW_VERTICAL_CLIP_ONLY = defineCompatibilityRule({
  id: 'word-exact-row-vertical-clip-only',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/table.test.ts#clips an overflowing merged owner when every row in its span is exact',
  },
  description: 'Preserve the established exact-row overflow behavior that clips the owned vertical interval without clipping nested table ink horizontally to the cell box.',
});

export const WORD_OVER_PAGE_CANT_SPLIT_CLIP = defineCompatibilityRule({
  id: 'word-over-page-cant-split-clip',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.120',
  },
  description: 'Word starts an over-page cantSplit row on a fresh page and clips its overflow instead of synthesizing a row continuation.',
});

export const WORD_OVER_PAGE_CELL_BREAK_OCCUPANCY = defineCompatibilityRule({
  id: 'word-over-page-cell-break-occupancy',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'over-page-cell-followed-by-authored-break',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'A table-cell paragraph taller than the body band counts its invisible continuation page before a following authored page break. A fitting paragraph does not. This holds with and without cantSplit; without the authored break, the next paragraph starts at the top of the continuation page. Tested at 500pt and 800pt against a 648pt body band, so farther overflow remains an inferred geometric extension.',
});

export const WORD_AUTHORED_ROW_HEIGHT_PAGE_BOUNDARY = defineCompatibilityRule({
  id: 'word-authored-row-height-page-boundary',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'ordinary-table-row-height-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'Word relocates an ordinary splittable exact-height row, or an atLeast row whose authored minimum governs its complete height, when that height exceeds the remaining page band and fits a fresh page. A shorter atLeast minimum permits content fragmentation; an auto row may fragment. Tested with cantSplit on/off, fitting/overflow bands, and keepLines/widow controls. Repeated headers and atLeast rows expanded by content are outside this observation.',
});

export const WORD_CELL_OWNED_ANCHOR_PAGE_CUT = defineCompatibilityRule({
  id: 'word-cell-owned-anchor-page-cut',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cell-owned-anchor-page-band-boundary',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'For an atLeast row without cantSplit containing a layoutInCell, allowOverlap wrapNone image, Word permits the image past its row border while the image stays in the page body band. Controlled preceding-spacing cases at 0pt and 40pt kept the row; at 80pt Word moved the complete row to the next page. A separate near-edge control confirmed that trailing empty cell paragraphs may still form an empty row continuation. Other anchor wrap/row-height combinations and over-page images are outside this observation.',
});

/** Compatibility choice governed by {@link WORD_CELL_OWNED_ANCHOR_PAGE_CUT}.
 * Geometry detection remains in the table paginator; this gate only chooses
 * the observed Word page cut when a fresh page offers more room. */
export function wordDefersCellOwnedAnchorPastPageBand(input: Readonly<{
  compatibility: 'word' | 'standard';
  availableHeightPt: number;
  freshPageHeightPt: number;
  epsilonPt: number;
}>): boolean {
  return input.compatibility === 'word'
    && input.availableHeightPt + input.epsilonPt < input.freshPageHeightPt;
}

export const WORD_PARALLEL_PARAGRAPH_ROW_CUT = defineCompatibilityRule({
  id: 'word-parallel-paragraph-row-cut',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'parallel-paragraph-row-cut-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When a page cut crosses a row containing parallel paragraph content, Word emits no cell content unless every unfinished paragraph cell can reach at least its first legal line or block boundary in that page band. The observed rule does not cover nested-table child boundaries.',
});

export const WORD_POSITIONED_TABLE_ADJACENCY_EXCLUSION = defineCompatibilityRule({
  id: 'word-positioned-table-adjacency-exclusion',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.149(a)',
  },
  description: 'Word excludes effectively positioned tables from the logical adjacent-table sequence before retained layout consumes the parser-owned sequence identity.',
});

export const WORD_TABLE_BORDER_WEIGHT_PRECEDENCE = defineCompatibilityRule({
  id: 'word-table-border-weight-precedence',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.169',
  },
  description: 'Use the documented Word border numbers for shared-cell conflict weight and force dotted and dashed borders to a complete weight of one.',
});

export const WORD_OMITTED_ROW_HEIGHT_RULE_AT_LEAST = defineCompatibilityRule({
  id: 'word-omitted-row-height-rule-at-least',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.180',
  },
  description: 'Treat an omitted trHeight hRule as atLeast while retaining an explicitly authored auto rule as authored input.',
});

export const WORD_AUTHORED_AUTO_ROW_HEIGHT_FLOOR = defineCompatibilityRule({
  id: 'word-authored-auto-row-height-floor',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/table-row-height.test.ts#auto with @val — @val is honored as a lower bound',
  },
  description: 'Preserve the established legacy-model behavior that an auto row with an authored height value uses that value as a lower bound.',
});

export const WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT = defineCompatibilityRule({
  id: 'word-collapsed-border-row-track-footprint',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'collapsed-border-row-track-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For automatic and at-least table rows with collapsed cell boundaries, Word includes half of the winning top and bottom rule widths in each row track. Exact rows retain their authored complete height, and cell spacing keeps independent edges out of the collapsed footprint.',
});

export const WORD_EFFECTIVE_FLOATING_TABLE_POSITIONING = defineCompatibilityRule({
  id: 'word-effective-floating-table-positioning',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.162',
  },
  description: 'Use parser-retained effective positioning status rather than lexical tblpPr presence to decide whether a table leaves ordinary flow.',
});

export const WORD_TABLE_CELL_SPACING_SCOPE_SHADOW = defineCompatibilityRule({
  id: 'word-table-cell-spacing-scope-shadow',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.152, 2.1.153, 2.1.154',
  },
  description: 'At each table-cell-spacing precedence scope, pct, auto, and nil resolve to zero and shadow lower scopes instead of being treated as absent.',
});

export const WORD_TABLE_MARGIN_SCOPE_SHADOW = defineCompatibilityRule({
  id: 'word-table-margin-scope-shadow',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.116, 2.1.125, 2.1.146, 2.1.177',
  },
  description: 'Preserve the documented scope-specific treatment of non-dxa table cell margins: leading/trailing defaults may resolve to zero while cell/exception and nil top/bottom values remain ignored.',
});

export const WORD_TABLE_CELL_ZERO_NIL_WIDTH_AUTO = defineCompatibilityRule({
  id: 'word-table-cell-zero-nil-width-auto',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.171(a-b)',
  },
  description: 'Interpret tcW with a zero w attribute, including dxa and pct, or nil type as auto. Intrinsic measurement and column constraints share the effective cell preference; noWrap uses an unbroken content minimum for these automatic cells instead of dxa width protection.',
});

export const WORD_FIRST_ROW_TABLE_EXCEPTION_SCOPE = defineCompatibilityRule({
  id: 'word-first-row-table-exception-scope',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.156, 2.1.158, 2.1.167',
  },
  description: 'Apply the supported first-row table-property exception facts at table scope. An authored first-row tblPrEx/tblW shadows the body width, including auto/nil/zero clearing it; width ceiling selection and column fitting share this effective preference.',
});

export const WORD_TRAILING_STRUCTURAL_CELL_MARKER = defineCompatibilityRule({
  id: 'word-trailing-structural-cell-marker',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/compatibility.test.ts#drops only an empty trailing paragraph after a non-paragraph cell block',
  },
  description: 'Exclude the required empty trailing cell paragraph from row-height and vertical-alignment measurements when it follows a visible non-paragraph block.',
});

export const WORD_CELL_VERTICAL_ALIGNMENT_INK_BLOCK = defineCompatibilityRule({
  id: 'word-cell-vertical-alignment-ink-block',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/cell-valign-leading-spacing.test.ts#inked block is vertically centred in the cell (midpoint = cell midpoint)',
  },
  description: 'Center or bottom-align the visible cell ink block without charging the first paragraph spaceBefore or final paragraph spaceAfter at the cell edges.',
});

export const WORD_VERTICAL_MERGE_TERMINAL_BORDER = defineCompatibilityRule({
  id: 'word-vertical-merge-terminal-border',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/cell-border-conflict-render.test.ts#uses the final continuation cell border at the bottom of a vertical merge',
  },
  description: 'Resolve the bottom edge of a vertically merged cell from its terminal continuation cell before applying shared-edge conflict rules.',
});

export const WORD_VERTICAL_SECTION_UPRIGHT_BLOCK_TABLE = defineCompatibilityRule({
  id: 'word-vertical-section-upright-block-table',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/vertical-table-upright.test.ts#the table advances the flow by its PHYSICAL WIDTH; body text stays vertical',
  },
  description: 'Paint a block table in an upright physical frame within a vertical section and charge its physical width as the body-flow advance.',
});

export function wordExactRowFloorPt(
  authoredHeightPt: number | null,
  bottomCellMarginsPt: readonly number[],
): number {
  return Math.max(0, authoredHeightPt ?? 0)
    + Math.max(0, ...bottomCellMarginsPt);
}

/** Compatibility projection governed by
 * {@link WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT}. */
export function wordCollapsedBorderRowTrackFootprintPt(
  topBoundaryWidthPt: number,
  bottomBoundaryWidthPt: number,
): number {
  return (Math.max(0, topBoundaryWidthPt) + Math.max(0, bottomBoundaryWidthPt)) / 2;
}

export function wordAuthoredBorderParticipates(
  authoredStyle: string | null | undefined,
): boolean {
  // `word-table-border-layer-cascade` owns this pre-conflict distinction.
  return authoredStyle !== null
    && authoredStyle !== undefined
    && authoredStyle !== 'none';
}

export function wordAlignedTableOriginPt(
  alignedPt: number,
  indentPt: number,
  bidiVisual: boolean,
): number {
  return bidiVisual ? alignedPt - indentPt : alignedPt + indentPt;
}

export function wordSpacedCellInsideBorderOverridesTable(input: Readonly<{
  spacingPt: number;
  directStyle: string | null | undefined;
  conditionalInsideStyle: string | null | undefined;
}>): boolean {
  return input.spacingPt > 0
    && !wordAuthoredBorderParticipates(input.directStyle)
    && wordAuthoredBorderParticipates(input.conditionalInsideStyle);
}

export function wordExactRowVerticalClipBounds(
  cellFlowBounds: LayoutRect,
  containingFlowBounds: LayoutRect,
): LayoutRect {
  // The exact trHeight constraint is block-axis-only. Keep the containing
  // flow's inline spill area, but extend it to any signed §17.4.50 tblInd
  // placement owned by this cell so a table shifted into the leading margin
  // is not clipped back to the ordinary text band.
  const xPt = Math.min(cellFlowBounds.xPt, containingFlowBounds.xPt);
  const rightPt = Math.max(
    cellFlowBounds.xPt + cellFlowBounds.widthPt,
    containingFlowBounds.xPt + containingFlowBounds.widthPt,
  );
  return Object.freeze({
    xPt,
    yPt: cellFlowBounds.yPt,
    widthPt: rightPt - xPt,
    heightPt: cellFlowBounds.heightPt,
  });
}

export function wordClipsOverPageCantSplitRow(input: Readonly<{
  compatibility: 'word' | 'standard';
  availableHeightPt: number;
  freshPageHeightPt: number;
  epsilonPt: number;
}>): boolean {
  return input.compatibility === 'word'
    && input.availableHeightPt + input.epsilonPt >= input.freshPageHeightPt;
}

/** Word observation under {@link WORD_AUTHORED_ROW_HEIGHT_PAGE_BOUNDARY}.
 * ECMA-376 §§17.4.6, 17.4.80 define cantSplit and row-height constraints but
 * do not prescribe this page-cut choice. Limit relocation to a first fragment
 * whose complete row fits a fresh page. The atLeast observation covered rows
 * whose content stayed below the authored minimum; content-expanded atLeast
 * rows and over-page rows are outside the tested boundary. */
export function wordRelocatesAuthoredHeightRowAtPageBoundary(input: Readonly<{
  compatibility: 'word' | 'standard';
  heightRule: 'auto' | 'atLeast' | 'exact';
  repeatedHeader: boolean;
  authoredHeightPt: number | null;
  availableHeightPt: number;
  wholeHeightPt: number;
  freshAvailableHeightPt: number;
  epsilonPt: number;
}>): boolean {
  return input.compatibility === 'word'
    && !input.repeatedHeader
    && (input.heightRule === 'exact' || input.heightRule === 'atLeast')
    && input.authoredHeightPt !== null
    && (input.heightRule === 'exact'
      || input.wholeHeightPt <= input.authoredHeightPt + input.epsilonPt)
    && input.authoredHeightPt > input.availableHeightPt + input.epsilonPt
    && input.wholeHeightPt <= input.freshAvailableHeightPt + input.epsilonPt;
}

/** Compatibility projection governed by {@link WORD_PARALLEL_PARAGRAPH_ROW_CUT}. */
export function wordRelocatesParallelParagraphRowCut(input: Readonly<{
  compatibility: 'word' | 'standard';
  hasUnfinishedParagraphWithoutProgress: boolean;
}>): boolean {
  return input.compatibility === 'word' && input.hasUnfinishedParagraphWithoutProgress;
}

export const WORD_TABLE_BORDER_STYLE_PRECEDENCE = Object.freeze([
  'single', 'thick', 'double', 'dotted', 'dashed', 'dotDash', 'dotDotDash', 'triple',
  'thinThickSmallGap', 'thickThinSmallGap', 'thinThickThinSmallGap', 'thinThickMediumGap',
  'thickThinMediumGap', 'thinThickThinMediumGap', 'thinThickLargeGap', 'thickThinLargeGap',
  'thinThickThinLargeGap', 'wave', 'doubleWave', 'dashSmallGap', 'dashDotStroked',
  'threeDEmboss', 'threeDEngrave', 'outset', 'inset',
] as const);

const WORD_BORDER_NUMBER: Readonly<Record<string, number>> = Object.freeze({
  single: 1,
  thick: 2,
  double: 3,
  dotDash: 8,
  dotDotDash: 9,
  triple: 10,
  thinThickSmallGap: 11,
  thickThinSmallGap: 12,
  thinThickThinSmallGap: 13,
  thinThickMediumGap: 14,
  thickThinMediumGap: 15,
  thinThickThinMediumGap: 16,
  thinThickLargeGap: 17,
  thickThinLargeGap: 18,
  thinThickThinLargeGap: 19,
  wave: 20,
  doubleWave: 21,
  dashSmallGap: 22,
  dashDotStroked: 23,
  threeDEmboss: 24,
  threeDEngrave: 25,
  outset: 26,
  inset: 27,
});

export function wordTableBorderWeight(
  style: string,
  widthPt: number,
): number {
  if (style === 'dotted' || style === 'dashed') return 1;
  return Math.max(0, widthPt) * 8 * (WORD_BORDER_NUMBER[style] ?? 0);
}

export function wordTableRowHeightRule(
  normalizedRule: 'exact' | 'atLeast' | 'auto',
  authored: boolean,
): 'exact' | 'atLeast' | 'auto' {
  return authored ? normalizedRule : 'atLeast';
}

export function wordAuthoredAutoRowHeightUsesFloor(
  rule: string | null | undefined,
  authoredHeight: number | null | undefined,
): boolean {
  return rule === 'auto'
    && authoredHeight !== null
    && authoredHeight !== undefined;
}

export function wordTableCellSpacingValuePt(
  kind: string,
  dxaValuePt: number | null,
): number | null {
  if (kind === 'pct' || kind === 'auto' || kind === 'nil') return 0;
  return dxaValuePt;
}

export function wordTableMarginValuePt(input: Readonly<{
  kind: string;
  dxaValuePt: number | null;
  scope: 'cell' | 'exception' | 'table' | 'style';
  edge: 'top' | 'bottom' | 'start' | 'end';
}>): number | null {
  if (input.kind === 'dxa') return input.dxaValuePt;
  if (input.scope === 'cell' || input.scope === 'exception') return null;
  if (input.edge === 'start' || input.edge === 'end') {
    if (input.kind === 'pct' || input.kind === 'auto' || input.kind === 'nil') return 0;
  }
  return null;
}

export function wordDropsTrailingStructuralCellMarker(input: Readonly<{
  contentLength: number;
  previousKind: string | undefined;
  lastKind: string | undefined;
  lastParagraphRunCount: number | undefined;
}>): boolean {
  return input.contentLength >= 2
    && input.lastKind === 'paragraph'
    && input.previousKind !== 'paragraph'
    && input.lastParagraphRunCount === 0;
}

/** ECMA-376 §17.4.50 and WORD_TABLE_ORIGIN_COMPATIBILITY: measured
 * center/end controls ignore tblInd for both placement and width fitting.
 * Margin hanging is a separate placement correction, never a width indent.
 * Unmeasured classes retain the established width/placement contract. */
export function wordTableWidthIndentPt(input: Readonly<{
  measured: boolean;
  justification: string | null | undefined;
  indentPt: number;
}>): number {
  return input.measured && (input.justification === 'center'
    || input.justification === 'right' || input.justification === 'end') ? 0 : input.indentPt;
}

/** WORD_TABLE_ORIGIN_COMPATIBILITY. An unresolved/unmeasured class retains
 * its established translation rather than inventing a margin default. */
export function wordTableOriginTranslationPt(input: Readonly<{
  mode: number | undefined;
  measured: boolean;
  justification: string | null | undefined;
  indentPt: number;
  indentAuthored: boolean;
  firstLeftMarginPt: number;
  rowLeftMarginPt: number;
}>): number {
  const indentPt = wordTableWidthIndentPt(input);
  if (!input.measured) return indentPt;
  if (input.justification === 'center') return indentPt;
  if (input.justification === 'right' || input.justification === 'end') {
    return indentPt + (input.mode === 15 ? 0 : input.rowLeftMarginPt);
  }
  return indentPt
    - (input.mode !== 15 && input.indentAuthored ? input.firstLeftMarginPt : 0);
}

/** Preserve grid indexes while dropping only the Office-measured unused track.
 * O(rows + columns + cells), with no parser mutation or extra content walk. */
export function wordFixedOccupiedGridInput(
  input: TableColumnLayoutInput,
  mode: number | undefined,
  measuredScope: boolean,
): TableColumnLayoutInput {
  if (!wordMeasuredTableOriginMode(mode) || !measuredScope
    || input.layout !== 'fixed' || input.rows.length === 0
    || !input.rows.every((row) => row.cells.length > 0
      && row.cells[0]?.columnStart === 1 && row.after === null
      && row.cells.every((cell) => cell.columnSpan === 1 && cell.preferredWidth?.kind === 'dxa' && cell.preferredWidth.value > 0))) {
    return input;
  }
  // A row occupying column zero is a measured counterexample. Only the single
  // universally skipped leading track was varied; spans, trailing skips,
  // missing preferences, and AutoFit retain their previous solver input.
  return {
    ...input,
    gridWidthsPt: input.gridWidthsPt.map((width, i) => i === 0 ? 0 : width),
    gridWidthKeys: input.gridWidthKeys?.map((key, i) => i === 0 ? null : key),
    rows: input.rows.map((row) => ({ ...row, before: null })),
  };
}
