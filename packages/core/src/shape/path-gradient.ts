import type { GradientFill } from '../types/common';
import { createAuxCanvasForContext } from '../canvas/aux-canvas';

/** Append the fill-bearing outline; coordinates belong to the supplied box.
 * The caller owns beginPath. Decorative, unfilled paths must be excluded. */
export type FillOutline = (
  ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number,
) => void;

// Resource policy, not an OOXML distance rule: at most 512² pixels. The
// rectangular field needs one ramp fill/readback and one pixel pass per paint,
// independent of outline complexity; bounds use one streaming O(commands) pass.
// There are no contour replays, distance surfaces or retained geometry/cache.
const MAX_EDGE = 512;

function outlineBounds(outline: FillOutline, ctx: CanvasRenderingContext2D, w: number, h: number) {
  let left = 0; let top = 0; let right = w; let bottom = h;
  const include = (x: number, y: number) => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;
    left = Math.min(left, x); top = Math.min(top, y);
    right = Math.max(right, x); bottom = Math.max(bottom, y);
  };
  const pathMethods = new Set([
    'moveTo', 'lineTo', 'bezierCurveTo', 'quadraticCurveTo', 'ellipse',
    'arc', 'rect', 'closePath',
  ]);
  // Observe path coordinates without touching the caller's current path or
  // retaining a command list. Bezier control hulls and full ellipse extents
  // are conservative: overestimating coverage is safe; cropping is not.
  const recorder = new Proxy(ctx, {
    get(target, property) {
      if (typeof property !== 'string' || !pathMethods.has(property)) return Reflect.get(target, property);
      return (...a: number[]) => {
        if (property === 'ellipse' || property === 'arc') {
          const rx = Math.abs(a[2]); const ry = property === 'arc' ? rx : Math.abs(a[3]);
          const angle = property === 'arc' ? 0 : a[4];
          const ex = Math.hypot(rx * Math.cos(angle), ry * Math.sin(angle));
          const ey = Math.hypot(rx * Math.sin(angle), ry * Math.cos(angle));
          include(a[0] - ex, a[1] - ey); include(a[0] + ex, a[1] + ey);
        } else if (property === 'rect') {
          include(a[0], a[1]); include(a[0] + a[2], a[1] + a[3]);
        } else {
          for (let i = 0; i + 1 < a.length; i += 2) include(a[i], a[i + 1]);
        }
      };
    },
  });
  outline(recorder, 0, 0, w, h);
  return { x: left, y: top, w: right - left, h: bottom - top };
}

/** Rectangular paths (ECMA-376 §20.1.8.46, §20.1.10.38) use box isolines:
 * max(normalized distances from the focus edges), never a circular radius.
 * §20.1.8.31 defines the center shade as filling the focus rectangle, relative
 * to the tile. Positive offsets inset its edges; negative offsets outset them.
 *
 * Office-produced wide/tall/square controls establish centered/corner box
 * isolines, including three stops. Asymmetric area controls do NOT establish
 * a flat plateau: the area below follows ECMA semantics, not Office parity.
 * [MS-OE376] §2.1.1377 instead describes an inscribed center path and a
 * shape-relative focus; §2.1.1378 forces xy tile flip. Those Office deviations
 * are intentionally unsupported here: retain tile-relative focus/authored flip.
 * General shape-path interpolation is unresolved and is not handled here.
 */
export function resolveRectPathGradient(
  fill: GradientFill, ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  shapeRotationDeg: number, outline?: FillOutline,
): CanvasPattern | null {
  if (![x, y, w, h, shapeRotationDeg].every(Number.isFinite) || w <= 0 || h <= 0) return null;
  const matrix = ctx.getTransform();
  const deviceScale = Math.max(Math.hypot(matrix.a, matrix.b), Math.hypot(matrix.c, matrix.d));
  // The authored xfrm box need not contain a custom path or callout tail.
  // Observe in authored units and allocate for its conservative painted bounds,
  // while keeping fillToRect and rotation anchored to the original tile.
  const bounds = outline ? outlineBounds(outline, ctx, w, h) : { x: 0, y: 0, w, h };
  if (![bounds.w, bounds.h].every(Number.isFinite)) return null;
  // A duplicated edge texel prevents Canvas's no-repeat sampler from blending
  // opaque gradients with transparent pixels when the raster is enlarged.
  const scale = Math.min(deviceScale || 1, (MAX_EDGE - 2) / bounds.w, (MAX_EDGE - 2) / bounds.h);
  const bw = Math.max(1, Math.min(MAX_EDGE - 2, Math.ceil(bounds.w * scale)));
  const bh = Math.max(1, Math.min(MAX_EDGE - 2, Math.ceil(bounds.h * scale)));
  const sx = bw / bounds.w; const sy = bh / bounds.h;
  const surface = createAuxCanvasForContext(ctx, bw + 2, bh + 2);
  const target = surface?.getContext('2d');
  const ramp = createAuxCanvasForContext(ctx, 1024, 1);
  const rampCtx = ramp?.getContext('2d');
  if (!surface || !target || !ramp || !rampCtx) return null;

  const gradient = rampCtx.createLinearGradient(0, 0, 1024, 0);
  for (const stop of fill.stops) {
    const hex = stop.color.replace(/^#/, '');
    gradient.addColorStop(Math.max(0, Math.min(1, stop.position)), `#${hex}`);
  }
  rampCtx.fillStyle = gradient;
  rampCtx.fillRect(0, 0, 1024, 1);
  const colors = rampCtx.getImageData(0, 0, 1024, 1).data;
  const ordered = [...fill.stops].sort((a, b) => a.position - b.position);
  const endpointColors = [ordered[0], ordered[ordered.length - 1]].map(stop => {
    const hex = stop.color.replace(/^#/, '');
    return [parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16),
      parseInt(hex.slice(4, 6), 16), hex.length >= 8 ? parseInt(hex.slice(6, 8), 16) : 255];
  });

  // Omitted fillToRect means the tile (§20.1.8.31).
  const focus = fill.fillToRect ?? {};
  const l = focus.l ?? 0; const t = focus.t ?? 0;
  // Sum authored opposing insets first: a collapsed 20%/80% focus must stay
  // a point despite subtraction rounding (1 - .8 can be smaller than .2).
  const r = l + (1 - (l + (focus.r ?? 0)));
  const b = t + (1 - (t + (focus.b ?? 0)));
  if (![l, t, r, b].every(Number.isFinite) || l > r || t > b) return null;
  const pixels = target.createImageData(bw + 2, bh + 2);
  const angle = fill.rotWithShape === false ? shapeRotationDeg * Math.PI / 180 : 0;
  const cos = Math.cos(angle); const sin = Math.sin(angle);
  const edgeDistance = (p: number, low: number, high: number): number => {
    if (p < low) return low > 0 ? (low - p) / low : 0;
    if (p > high) return high < 1 ? (p - high) / (1 - high) : 0;
    return 0;
  };
  for (let row = -1; row <= bh; row++) {
    for (let col = -1; col <= bw; col++) {
      const offset = ((row + 1) * (bw + 2) + col + 1) * 4;
      const px = bounds.x + (Math.max(0, Math.min(bw - 1, col)) + .5) / sx - w / 2;
      const py = bounds.y + (Math.max(0, Math.min(bh - 1, row)) + .5) / sy - h / 2;
      const u = .5 + (cos * px - sin * py) / w;
      const v = .5 + (sin * px + cos * py) / h;
      const f = Math.max(edgeDistance(u, l, r), edgeDistance(v, t, b));
      const color = Math.min(1023, Math.max(0, Math.floor(f * 1024))) * 4;
      // Canvas samples the lookup ramp at pixel centers. Preserve exact stop
      // endpoints (including transparent RGB) instead of moving a focus into
      // the first half texel of the transition.
      const endpoint = f <= 0 ? endpointColors[0] : f >= 1 ? endpointColors[1] : undefined;
      if (endpoint) {
        pixels.data[offset] = endpoint[0];
        pixels.data[offset + 1] = endpoint[1];
        pixels.data[offset + 2] = endpoint[2];
        pixels.data[offset + 3] = endpoint[3];
      } else {
        pixels.data[offset] = colors[color];
        pixels.data[offset + 1] = colors[color + 1];
        pixels.data[offset + 2] = colors[color + 2];
        pixels.data[offset + 3] = colors[color + 3];
      }
    }
  }
  target.putImageData(pixels, 0, 0);
  const pattern = ctx.createPattern(surface, 'no-repeat');
  if (!pattern || typeof pattern.setTransform !== 'function') return null;
  pattern.setTransform({
    a: 1 / sx, b: 0, c: 0, d: 1 / sy,
    e: x + bounds.x - 1 / sx, f: y + bounds.y - 1 / sy,
  });
  return pattern;
}
