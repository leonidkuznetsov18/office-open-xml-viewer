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
 * the baseline position. This is observed Excel behaviour (#1604 PDF
 * controls). They covered twelve faces, spcPct 80/150 %, and spcPts from
 * 15 pt to 56 pt on both sides of each natural height, including mixed sizes.
 * The rule fits every line of the controls to the export's rounding.
 *
 * Natural box: `a`/`d`, L = a + d. Spaced height: H.
 * Let k = max(0, d − 0.25·L); faces with a deep natural descent (Meiryo,
 * Yu Gothic) have k > 0.
 * - H < L: the descent becomes min(d, 0.25·H + k).
 * - H > L: while 0.25·H ≤ d, the natural descent is kept and all the extra
 *   space goes above. Past that point the descent becomes 0.25·H + k.
 *   Examples: Meiryo 14 pt (d 9.31 pt) keeps d up to spcPts 36 and gets
 *   12.5 pt at 40; Arial always gets 0.25·H.
 * - H = L: the natural box is kept.
 *
 * The ascent is the rest of H. Each baseline pitch is therefore the previous
 * line's descent plus the next line's ascent, and a run of same-size lines
 * steps by exactly H.
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
  const k = Math.max(0, natural.descent - 0.25 * L);
  const descent = H < L
    ? Math.min(natural.descent, 0.25 * H + k)
    : 0.25 * H <= natural.descent ? natural.descent : 0.25 * H + k;
  return { ascent: H - descent, descent };
}

/**
 * `a:spcBef` (§21.1.2.2.10) / `a:spcAft` (§21.1.2.2.9) in px. spcPts is
 * absolute. For spcPct, Excel's base is the natural single-line height of
 * the adjacent line: the paragraph's first line for spcBef, its last line
 * for spcAft (issue #1604 controls: 50 % and 100 % at 24 pt, 14/40 pt mixes
 * where the base followed the owning paragraph's 40 pt line). The base stays
 * the natural line when the paragraph's lnSpc is 80 %, 150 % or 36 pt.
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
