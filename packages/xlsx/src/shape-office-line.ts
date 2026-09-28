import { findReferenceFontMetrics } from '@silurus/ooxml-core';
import { excelDrawingMlLineRatios } from '@silurus/ooxml-core/internal/office-auto-line';
import type { OfficeFontFallbackRequest, OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import type { ShapeText, ShapeTextRun } from './types.js';

type TextRun = Extract<ShapeTextRun, { type: 'text' }>;

/** Excel's natural line box of one shape-text run, as em ratios. */
export interface ShapeRunLineRatios {
  ascentRatio: number;
  descentRatio: number;
}

/** The single face that owns a run's line box, or undefined when the run
 * names no face or routes East Asian / complex-script text to another face.
 * The shape paint path selects a:latin; a distinct a:ea or a:cs face would
 * need script-run routing before its metrics could own the line. */
function soleFace(run: TextRun): string | undefined {
  const face = run.fontFace?.trim();
  if (!face) return undefined;
  const same = (other: string | undefined) => !other?.trim()
    || other.trim().toLocaleLowerCase('en-US') === face.toLocaleLowerCase('en-US');
  return same(run.fontFaceEa) && same(run.fontFaceCs) ? face : undefined;
}

/** Text runs of a shape body whose face could own Excel's line box. These are
 * the tuples the workbook preflights as exact local faces. */
export function shapeLineFontRuns(text: ShapeText): TextRun[] {
  const runs: TextRun[] = [];
  for (const paragraph of text.paragraphs) for (const run of paragraph.runs) {
    if (run.type === 'text' && soleFace(run)) runs.push(run);
  }
  return runs;
}

/** One key rule for workbook, worker, and synchronous shape paint. */
export function officeRequestKey(request: OfficeFontFallbackRequest): string {
  const family = request.family.trim().toLowerCase();
  const weight = request.weight ?? 400;
  const style = request.style ?? 'normal';
  return weight === 400 && style === 'normal' ? family : `${family}:${weight}:${style}`;
}

export function shapeOfficeRouteKey(run: TextRun): string {
  return officeRequestKey({ family: run.fontFace!, weight: run.bold ? 700 : 400,
    style: run.italic ? 'italic' : 'normal' });
}

/**
 * Excel's natural line box for one run (see `excelDrawingMlLineRatios`).
 *
 * The static catalog is reference geometry, not proof of the bytes behind
 * local(). Admit it only after the exact style has loaded. Excel for Mac
 * prefers a macOS system copy of a family over the copy bundled in Office
 * (#1604: its Times New Roman had the system hhea lineGap 87), so macOS
 * profiles win when present, and Office profiles are used otherwise. Every
 * profile in the chosen source must agree. Anything else returns undefined
 * and the caller keeps the ordinary Canvas line box: another source such as
 * a published open font, missing OS/2 data, disagreeing profiles, or a macOS
 * Far East face (not measured).
 */
export function shapeRunLineRatios(
  run: TextRun,
  route: OfficeFontFallbackRoute | undefined,
): ShapeRunLineRatios | undefined {
  const family = soleFace(run);
  if (!family || !route || route.source !== 'local' || route.metric.synthesized
    || !route.resourceIdentity.startsWith('office-local:')) return undefined;
  if (route.requestedFamily.toLocaleLowerCase('en-US') !== family.toLocaleLowerCase('en-US')
    || route.weight !== (run.bold ? 700 : 400)
    || route.style !== (run.italic ? 'italic' : 'normal')) return undefined;
  const profiles = findReferenceFontMetrics(family, { weight: route.weight, style: route.style });
  const system = profiles.filter((p) => p.source === 'macos-system' || p.source === 'macos-supplemental');
  const chosen = system.length > 0 ? system : profiles.filter((p) => p.source === 'office-mac');
  if (chosen.length === 0) return undefined;
  let ratios: ShapeRunLineRatios | undefined;
  for (const profile of chosen) {
    if (profile.farEastCodePage == null || !profile.win) return undefined;
    const projected = excelDrawingMlLineRatios({
      faceSource: system.length > 0 ? 'system' : 'office-bundle',
      unitsPerEm: profile.unitsPerEm,
      hhea: profile.hhea,
      win: profile.win,
      typoMetrics: profile.typoMetrics,
      farEastCodePage: profile.farEastCodePage,
    });
    if (!projected || (ratios && (projected.ascentRatio !== ratios.ascentRatio
      || projected.descentRatio !== ratios.descentRatio))) return undefined;
    ratios = projected;
  }
  return ratios;
}
