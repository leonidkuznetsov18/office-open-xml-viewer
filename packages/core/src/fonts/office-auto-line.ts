/**
 * Office-observed automatic single-line allocation from one static OpenType
 * face. Word for Mac controls isolated OS/2 code-page bits 17–20 and found a
 * 1.3× hhea glyph box for that class; other faces use signed hhea lineGap above
 * the baseline. Independent Excel DrawingML controls with omitted <a:lnSpc>
 * matched the same projection for Meiryo UI Bold and Arial Bold at 12/25 pt and
 * top/centre/bottom anchors. ECMA-376 defines those spacing/anchor attributes,
 * but does not select the font tables. Callers must gate the rule to their own
 * tested format and to a known font-metric source.
 */
export const OFFICE_FAR_EAST_SINGLE_LINE_FACTOR = 1.3;

export function officeOpenTypeAutoLineRatios(metrics: Readonly<{
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
  const { unitsPerEm, hheaAscent, hheaDescent, hheaLineGap, farEastCodePage } = metrics;
  if (!(Number.isFinite(unitsPerEm) && unitsPerEm > 0
    && Number.isFinite(hheaAscent) && hheaAscent >= 0
    && Number.isFinite(hheaDescent) && hheaDescent <= 0
    && Number.isFinite(hheaLineGap))) return null;
  const glyphBox = hheaAscent - hheaDescent;
  if (!(glyphBox > 0)) return null;
  const farEastHalfLeading = ((OFFICE_FAR_EAST_SINGLE_LINE_FACTOR - 1) / 2) * glyphBox;
  const ascent = farEastCodePage
    ? hheaAscent + farEastHalfLeading
    : hheaAscent + hheaLineGap;
  const descent = farEastCodePage
    ? -hheaDescent + farEastHalfLeading
    : -hheaDescent;
  if (!(ascent >= 0 && descent >= 0 && ascent + descent > 0)) return null;
  return Object.freeze({
    lineHeightRatio: (ascent + descent) / unitsPerEm,
    designAscentRatio: ascent / unitsPerEm,
    designDescentRatio: descent / unitsPerEm,
  });
}

/**
 * Excel's natural single line for DrawingML shape text, from one static
 * OpenType face (observed behaviour; ECMA-376 does not select font tables).
 *
 * Excel for Mac 16.113.2 PDF controls (#1604) exported 113 shape bodies over
 * Calibri, Arial, Times New Roman, Yu Gothic, Meiryo and MS Gothic at 8–72 pt,
 * with verified embedded faces. Every baseline, including the first one below
 * `tIns`, matched this projection within 0.16 pt:
 *
 * - The glyph box is the OS/2 usWin extent. Yu Gothic separates it from hhea
 *   (hhea box 1.102 em, usWin box 1.287 em); Excel's 1.673 em pitch is
 *   1.3 × the usWin box.
 * - A face in the Far East code-page class (OS/2 ulCodePageRange1 bits 17–20,
 *   the class Word for Mac selects by) gets a 1.3× box, with half of the
 *   added leading above the ascent and half below the descent (Yu Gothic,
 *   Meiryo, MS Gothic).
 * - Other faces add the Windows TEXTMETRIC external leading
 *   max(0, hhea.lineGap − (usWin box − hhea box)) above the ascent (Arial
 *   1.150 em, Times New Roman with lineGap 87 1.150 em, Calibri 1.221 em).
 *   Every measured Latin face has equal hhea and usWin boxes, so the
 *   external-leading clamp itself is the documented GDI definition, not a
 *   separately measured case.
 *
 * Callers must gate this to Excel shape text and a known font-metric source.
 */
export function excelDrawingMlLineRatios(metrics: Readonly<{
  unitsPerEm: number;
  winAscent: number;
  winDescent: number;
  hheaAscent: number;
  hheaDescent: number;
  hheaLineGap: number;
  farEastCodePage: boolean;
}>): Readonly<{ ascentRatio: number; descentRatio: number }> | null {
  const { unitsPerEm, winAscent, winDescent, hheaAscent, hheaDescent, hheaLineGap, farEastCodePage } = metrics;
  if (!(Number.isFinite(unitsPerEm) && unitsPerEm > 0
    && Number.isFinite(winAscent) && Number.isFinite(winDescent)
    && Number.isFinite(hheaAscent) && Number.isFinite(hheaDescent) && Number.isFinite(hheaLineGap))) return null;
  const box = winAscent + winDescent;
  if (!(box > 0) || winAscent < 0 || winDescent < 0) return null;
  let ascent: number;
  let descent: number;
  if (farEastCodePage) {
    const half = ((OFFICE_FAR_EAST_SINGLE_LINE_FACTOR - 1) / 2) * box;
    ascent = winAscent + half;
    descent = winDescent + half;
  } else {
    const externalLeading = Math.max(0, hheaLineGap - (box - (hheaAscent - hheaDescent)));
    ascent = winAscent + externalLeading;
    descent = winDescent;
  }
  return Object.freeze({ ascentRatio: ascent / unitsPerEm, descentRatio: descent / unitsPerEm });
}
