import type { GradientStop } from '../types/common';

// Office-implicit DrawingML gradient transfer rule (issue #1599).
//
// PowerPoint's electronic-distribution PDF export encodes every DrawingML
// gradient whose stop list is exactly two stops at 0% and 100% as a sampled
// 256-entry colour function, and every other stop list as piecewise-linear
// interpolation. All 410 sampled functions in the private controls (linear and
// path gradients; black/white and primary-colour pairs; single and reflected
// cycles) equal
//
//   w(t)   = (Φ((t - ½) / ¼) - Φ(-2)) / (Φ(2) - Φ(-2))
//   colour = ((1 - w)·c₀^γ + w·c₁^γ)^(1/γ),  γ = 2.2
//
// within 0.82 of 255 per channel. Φ is the standard normal CDF. This matches
// the legacy Office shade flags "sigma transfer" and "gamma correction after
// interpolation" ([MS-ODRAW] §2.2.50 MSOSHADETYPE); that document names them
// without defining the curves, so the constants above are the ones the
// exported functions determine. Other stop lists (e.g. 20%/80%, three or five
// stops) are linear in encoded RGB. ECMA-376 §20.1.8.36 defines no
// interpolation space.
//
// Scope: measured in PowerPoint 16.113 and applied as one shared DrawingML
// fill rule in every host; Word and Excel are not separately confirmed. Alpha is
// not a colour channel: it uses the transferred weight without gamma; the
// controls contain no two-stop alpha gradient to confirm this.

const SAMPLES = 1024;
/** Maximum deviation, in 8-bit levels per channel, of the emitted piecewise-
 * linear Canvas stops from the exact transfer curve. */
const TOLERANCE = 0.25;
const CACHE_LIMIT = 256;
const cache = new Map<string, GradientStop[]>();

/** Abramowitz & Stegun 7.1.26 (|error| < 1.5e-7). */
function erf(x: number): number {
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(x);
  const k = 1 / (1 + 0.3275911 * ax);
  const poly = k * (0.254829592 + k * (-0.284496736 + k * (1.421413741
    + k * (-1.453152027 + k * 1.061405429))));
  return sign * (1 - poly * Math.exp(-ax * ax));
}

const normalCdf = (z: number): number => 0.5 * (1 + erf(z / Math.SQRT2));
const SIGMA_LOW = normalCdf(-2);
const SIGMA_SPAN = normalCdf(2) - SIGMA_LOW;

/** The two-endpoint sigma transfer weight for gradient position t ∈ [0, 1]. */
export function officeTwoStopWeight(t: number): number {
  const clamped = Math.min(1, Math.max(0, t));
  return (normalCdf((clamped - 0.5) / 0.25) - SIGMA_LOW) / SIGMA_SPAN;
}

function channels(color: string): [number, number, number, number] {
  const hex = color.charCodeAt(0) === 35 ? color.slice(1) : color;
  const value = (offset: number) => Number.parseInt(hex.slice(offset, offset + 2), 16);
  return [value(0), value(2), value(4), hex.length >= 8 ? value(6) : 255];
}

const toHex = (value: number) => Math.min(255, Math.max(0, Math.round(value)))
  .toString(16).padStart(2, '0');

/** True for the stop list that Office interpolates with the transfer rule. */
export function usesOfficeTwoStopTransfer(stops: readonly GradientStop[]): boolean {
  if (stops.length !== 2) return false;
  const [first, second] = stops[0].position <= stops[1].position
    ? [stops[0], stops[1]] : [stops[1], stops[0]];
  return first.position === 0 && second.position === 1;
}

/**
 * Canvas colour stops reproducing Office interpolation. Stop lists outside the
 * two-endpoint rule are returned unchanged (encoded-RGB linear interpolation).
 * The two-endpoint curve is emitted as a simplified piecewise-linear list
 * (typically 15-45 stops) whose error is bounded by {@link TOLERANCE}.
 */
export function officeGradientStops(stops: readonly GradientStop[]): readonly GradientStop[] {
  if (!usesOfficeTwoStopTransfer(stops)) return stops;
  const [first, second] = stops[0].position <= stops[1].position
    ? [stops[0], stops[1]] : [stops[1], stops[0]];
  const key = `${first.color}|${second.color}`;
  const cached = cache.get(key);
  if (cached) return cached;
  const c0 = channels(first.color);
  const c1 = channels(second.color);
  if (![...c0, ...c1].every(Number.isFinite)) return stops;
  const gamma = 2.2;
  const lin0 = c0.slice(0, 3).map(v => (v / 255) ** gamma);
  const lin1 = c1.slice(0, 3).map(v => (v / 255) ** gamma);
  const curve: number[][] = [];
  for (let i = 0; i <= SAMPLES; i++) {
    const w = officeTwoStopWeight(i / SAMPLES);
    curve.push([
      ...lin0.map((v, k) => 255 * ((1 - w) * v + w * lin1[k]) ** (1 / gamma)),
      (1 - w) * c0[3] + w * c1[3],
    ]);
  }
  // Douglas-Peucker over the sampled curve, max-norm across channels.
  const keep = new Set<number>([0, SAMPLES]);
  const pending: Array<[number, number]> = [[0, SAMPLES]];
  while (pending.length > 0) {
    const [i, j] = pending.pop()!;
    let worst = -1; let worstError = TOLERANCE;
    for (let m = i + 1; m < j; m++) {
      const a = (m - i) / (j - i);
      for (let k = 0; k < 4; k++) {
        const error = Math.abs(curve[i][k] + a * (curve[j][k] - curve[i][k]) - curve[m][k]);
        if (error > worstError) { worstError = error; worst = m; }
      }
    }
    if (worst > 0) { keep.add(worst); pending.push([i, worst], [worst, j]); }
  }
  const expanded = [...keep].sort((a, b) => a - b).map(index => {
    const [r, g, b, alpha] = curve[index];
    return { position: index / SAMPLES, color: `${toHex(r)}${toHex(g)}${toHex(b)}${toHex(alpha)}` };
  });
  if (cache.size >= CACHE_LIMIT) cache.delete(cache.keys().next().value!);
  cache.set(key, expanded);
  return expanded;
}
