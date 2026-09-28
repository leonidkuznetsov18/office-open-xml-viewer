import type { CompatibilityRule, DeepReadonly } from './types.js';

function requireText(value: string, field: string): void {
  if (value.trim() === '') throw new Error(`CompatibilityRule.${field} must not be empty`);
}

export function defineCompatibilityRule<const Rule extends CompatibilityRule>(
  rule: Rule,
): DeepReadonly<Rule> {
  requireText(rule.id, 'id');
  requireText(rule.description, 'description');
  if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(rule.id)) {
    throw new Error('CompatibilityRule.id must be a stable kebab-case identifier');
  }
  if (rule.evidence.kind === 'microsoft-note') {
    requireText(rule.evidence.reference, 'evidence.reference');
    if (!/^\[MS-[A-Z0-9]+\] §§?\d/.test(rule.evidence.reference)) {
      throw new Error('CompatibilityRule.evidence.reference must identify a Microsoft specification section');
    }
  } else if (rule.evidence.kind === 'regression-test') {
    requireText(rule.evidence.reference, 'evidence.reference');
    if (!/^packages\/docx\/src\/.+\.(?:test|spec)\.tsx?#[^#]+$/.test(
      rule.evidence.reference,
    )) {
      throw new Error('CompatibilityRule.evidence.reference must use DOCX path#test-title');
    }
  } else {
    requireText(rule.evidence.syntheticFixtureId, 'evidence.syntheticFixtureId');
    requireText(rule.evidence.application, 'evidence.application');
    requireText(rule.evidence.version, 'evidence.version');
    requireText(rule.evidence.platform, 'evidence.platform');
    if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(rule.evidence.syntheticFixtureId)) {
      throw new Error('CompatibilityRule.evidence.syntheticFixtureId must be kebab-case');
    }
  }
  Object.freeze(rule.evidence);
  return Object.freeze(rule) as DeepReadonly<Rule>;
}

export const WORD_SECTION_BTLR_TBRL_PAGE_FRAME = defineCompatibilityRule({
  id: 'word-section-btlr-tbrl-page-frame',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s',
  },
  description: 'Issue #988 comment 4950296007 records that, unlike the normative ECMA-376 Part 4 §14.11.7 equivalence to lr, Word uses the tbRl page frame for section-level btLr; this rule covers only the page frame, while glyph orientation is paint-owned.',
});

export const WORD_SQUARE_LINE_START_ONE_INCH = defineCompatibilityRule({
  id: 'word-square-line-start-one-inch',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/float-line-start-one-inch.test.ts#(e) the boundary is identical across scales (absolute pt width)',
  },
  description: 'Issue #676 records that Word starts a content line beside a square-wrapped object only when the free side gap is at least one inch; tight and through polygon openings and empty paragraph marks are outside this rule.',
});

export const WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT = defineCompatibilityRule({
  id: 'word-float-different-paragraph-displacement',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/floats.test.ts#keeps observed different-paragraph displacement on exclusion bounds',
  },
  description: 'Preserve the established Word-compatible policy that an overlap-permitted float is displaced by exclusion geometry from floats anchored in other paragraphs, while same-paragraph floats may overlap.',
});

export const WORD_PAGE_ANCHORED_TABLE_COLLISION_DEFERRAL = defineCompatibilityRule({
  id: 'word-page-anchored-table-collision-deferral',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/float-table-page-fit.test.ts#(g) DEFERS a page-anchored floating table when its raw band intersects an existing table float',
  },
  description: 'Preserve the established Word-compatible pagination behavior that defers an absolute page- or margin-anchored floating table when its authored object band intersects an existing floating-table text-exclusion band on the page.',
});

export const WORD_EMPTY_MARK_FLOAT_SIDE_GAP = defineCompatibilityRule({
  id: 'word-empty-mark-float-side-gap',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/float-line-start-one-inch.test.ts#keeps an anchor-host metric-only line on the paragraph-mark threshold',
  },
  description: 'An empty or anchor-only paragraph-mark line may start beside a square-wrapped object when the available side gap can hold the paragraph mark em; the one-inch content-line threshold does not apply.',
});

/** Word compatibility width from issue #676, in points. ECMA-376
 * §20.4.2.17 defines square wrapping but no minimum side-gap width. */
export const WORD_MIN_LINE_START_PT = 72;

/** One-twip tolerance for the inclusive one-inch boundary. It absorbs the
 * authored twip/EMU conversion deficit documented by the regression suite. */
export const LINE_START_GAP_EPS_PT = 0.05;

export function wordMinLineStartPx(scale: number): number {
  return (WORD_MIN_LINE_START_PT - LINE_START_GAP_EPS_PT) * scale;
}

export function wordEmptyMarkMinimumStartWidthPx(
  paragraphMarkEmPt: number,
  scale: number,
): number {
  return paragraphMarkEmPt * scale;
}
