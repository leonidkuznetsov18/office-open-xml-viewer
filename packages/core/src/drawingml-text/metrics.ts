/** `a:bodyPr` text box after its four independent EMU insets. */
export interface DrawingMlInsets {
  lIns: number;
  rIns: number;
  tIns: number;
  bIns: number;
}

export interface DrawingMlTextRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** ECMA-376 §21.1.2.1.1: all four `bodyPr` insets affect the text rectangle. */
export function drawingMlTextRect(
  width: number,
  height: number,
  insets: DrawingMlInsets,
  pxPerEmu: number,
): DrawingMlTextRect {
  const left = insets.lIns * pxPerEmu;
  const right = insets.rIns * pxPerEmu;
  const top = insets.tIns * pxPerEmu;
  const bottom = insets.bIns * pxPerEmu;
  return {
    left,
    top,
    width: Math.max(0, width - left - right),
    height: Math.max(0, height - top - bottom),
  };
}

export type DrawingMlLineSpacing =
  | { type: 'pct'; val: number }
  | { type: 'pts'; val: number }
  | null
  | undefined;

/**
 * `a:lnSpc` (§21.1.2.2.5) and `a:normAutofit@lnSpcReduction`
 * (§21.1.2.1.3). Percentage spacing multiplies the natural line box;
 * point spacing replaces it. The stored normal-autofit reduction applies
 * only to percentage spacing (including the implicit 100%).
 */
export function drawingMlLineHeight(
  naturalHeight: number,
  spacing: DrawingMlLineSpacing,
  pxPerPt: number,
  reduction = 0,
): number {
  const base = spacing?.type === 'pts'
    ? spacing.val * pxPerPt
    : naturalHeight * (spacing?.type === 'pct' ? spacing.val / 100000 : 1);
  return spacing?.type === 'pts' ? base : base * (1 - reduction);
}

/** A line box split at its baseline: `ascent` above, `descent` below. */
export interface DrawingMlLineBox {
  ascent: number;
  descent: number;
}

/**
 * Where `a:lnSpc` (§21.1.2.2.5) puts the baseline inside the spaced line.
 * ECMA-376 gives the spaced height (see {@link drawingMlLineHeight}) but not
 * the baseline position. This is observed Excel behaviour (issue #1604 PDF
 * controls: six fonts, spcPct 80/150 % at 11 and 40 pt, spcPts 18/36/30 pt,
 * and mixed sizes). It matched every line within 0.16 pt except the one
 * boundary noted below.
 *
 * With the natural box `a`/`d` (L = a + d) and the spaced height H:
 * - Excel keeps a 3:1 split of the spaced line, lowered by
 *   k = max(0, 0.75·L − a) for faces whose natural ascent is below 75 %
 *   (Meiryo, Yu Gothic). A taller line gets `0.75·H − k` above the baseline.
 * - A shorter line keeps at least its natural descent: `max(H − d, 0.75·H − k)`.
 * - H = L therefore reproduces the natural box.
 *
 * The descent is the rest of H, so each baseline pitch is the previous
 * line's descent plus the next line's ascent, and a run of same-size lines
 * steps by exactly H.
 *
 * Known residual: a 14 pt Meiryo line at spcPts 30 (just above its 27.3 pt
 * natural height) sat 0.65 pt lower in Excel than this rule predicts.
 */
export function drawingMlSpacedLineBox(
  natural: DrawingMlLineBox,
  spacing: DrawingMlLineSpacing,
  pxPerPt: number,
  reduction = 0,
): DrawingMlLineBox {
  const L = natural.ascent + natural.descent;
  const H = drawingMlLineHeight(L, spacing, pxPerPt, reduction);
  if (H === L) return { ascent: natural.ascent, descent: natural.descent };
  const k = Math.max(0, 0.75 * L - natural.ascent);
  const ascent = H < L
    ? Math.max(H - natural.descent, 0.75 * H - k)
    : 0.75 * H - k;
  return { ascent, descent: H - ascent };
}

/**
 * `a:spcBef` (§21.1.2.2.10) / `a:spcAft` (§21.1.2.2.9) in px. spcPts is
 * absolute. For spcPct, Excel's base is the natural single-line height of
 * the adjacent line: the paragraph's first line for spcBef, its last line
 * for spcAft (issue #1604 controls: 50 % and 100 % at 24 pt, and 14/40 pt
 * mixes where the base followed the owning paragraph's 40 pt line). The
 * controls held lnSpc at 100 % whenever spcPct was used, so a percentage of
 * a spaced line is outside the measured scope.
 */
export function drawingMlParagraphSpacing(
  spacing: DrawingMlLineSpacing,
  naturalLineHeight: number,
  pxPerPt: number,
): number {
  if (!spacing) return 0;
  return spacing.type === 'pts'
    ? spacing.val * pxPerPt
    : naturalLineHeight * (spacing.val / 100000);
}

/** ECMA-376 §20.1.7.2 `a:bodyPr@anchor` for the line block. */
export function drawingMlBlockTop(
  anchor: string | undefined,
  rect: DrawingMlTextRect,
  blockHeight: number,
): number {
  if (anchor === 'ctr') return rect.top + (rect.height - blockHeight) / 2;
  if (anchor === 'b') return rect.top + Math.max(0, rect.height - blockHeight);
  return rect.top;
}
