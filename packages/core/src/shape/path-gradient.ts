import type { GradientFill } from '../types/common';
import { createAuxCanvasForContext } from '../canvas/aux-canvas';
import { interiorDistances } from './gradient-distance';

/** Append the fill-bearing outline; coordinates belong to the supplied box.
 * The caller owns beginPath. Decorative, unfilled paths must be excluded. */
export type FillOutline = (
  ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number,
) => void;

// Resource policy, not an OOXML distance rule: one shape-local raster, at most
// 512² pixels, and an 8-bit contour field. No document-sized surface or cache.
const MAX_EDGE = 512;
const CONTOURS = 255;

type PathOperation = { method: string; args: unknown[] };

function recordOutline(outline: FillOutline, ctx: CanvasRenderingContext2D, w: number, h: number): PathOperation[] {
  const operations: PathOperation[] = [];
  const pathMethods = new Set([
    'moveTo', 'lineTo', 'bezierCurveTo', 'quadraticCurveTo', 'ellipse',
    'arc', 'rect', 'closePath',
  ]);
  const recorder = new Proxy(ctx, {
    get(target, property) {
      if (typeof property === 'string' && pathMethods.has(property)) {
        return (...args: unknown[]) => { operations.push({ method: property, args }); };
      }
      return Reflect.get(target, property);
    },
  });
  outline(recorder, 0, 0, w, h);
  return operations;
}

function outlineBounds(operations: PathOperation[], w: number, h: number) {
  let left = 0; let top = 0; let right = w; let bottom = h;
  const include = (x: number, y: number) => {
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;
    left = Math.min(left, x); top = Math.min(top, y);
    right = Math.max(right, x); bottom = Math.max(bottom, y);
  };
  for (const op of operations) {
    const a = op.args.map(Number);
    if (op.method === 'ellipse' || op.method === 'arc') {
      const rx = Math.abs(a[2]); const ry = op.method === 'arc' ? rx : Math.abs(a[3]);
      const angle = op.method === 'arc' ? 0 : a[4];
      const ex = Math.hypot(rx * Math.cos(angle), ry * Math.sin(angle));
      const ey = Math.hypot(rx * Math.sin(angle), ry * Math.cos(angle));
      include(a[0] - ex, a[1] - ey); include(a[0] + ex, a[1] + ey);
    } else if (op.method === 'rect') {
      include(a[0], a[1]); include(a[0] + a[2], a[1] + a[3]);
    } else {
      // A Bezier lies inside the hull of its control points. Including that
      // hull may overestimate bounds, but cannot crop authored outsets.
      for (let i = 0; i + 1 < a.length; i += 2) include(a[i], a[i + 1]);
    }
  }
  return { x: left, y: top, w: right - left, h: bottom - top };
}

/** ECMA-376 §20.1.8.31/46 and §20.1.10.38 specify a focus rectangle and a
 * rectangular or shape-following path, but not a contour interpolation mesh.
 * Library policy: interpolate outline coordinates affinely from the tile to
 * the focus rectangle, with stop 0 inside the focus and stop 1 at the boundary.
 * Rectangular contours therefore use max(normalized edge distances), never a
 * circular radius. Shape contours preserve the actual outline, including its
 * concavities. Intersect contracted contours with inward outline offsets whose
 * distance is proportional to the focus's clearance from the boundary.
 * Convex outlines keep their affine contours; inward offsets prevent concave
 * notch boundaries being repainted by contractions of another part of a shape.
 * If the focus is outside the silhouette, the maximum interior clearance is
 * the offset extent: an invisible focus cannot define an interior clearance.
 * Overlapping contractions belong to the innermost painted contour. The shape
 * supplies the final clip, and the authored focus takes precedence.
 *
 * PowerPoint controls confirm box isolines for wide/tall/square rectangles and
 * outline isolines for triangles/stars. Ellipse/arrow/concave custom controls
 * instead retain centered radial shading when focus changes; asymmetric rect
 * controls do not establish a flat focus plateau. Those observations are not a
 * uniform compatibility rule: honor the specified focus and outline here,
 * without geometry-name exceptions or an empirically fitted distance formula.
 * [MS-OE376] §2.1.1377 documents Office's inscribed-path center shade and
 * shape-relative focus, and §2.1.1378 its ignored flip attribute. This painter
 * follows ECMA's tile-relative focus and authored flip instead.
 */
export function resolvePathGradient(
  fill: GradientFill, ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  shapeRotationDeg: number, outline?: FillOutline,
): CanvasPattern | null {
  if (![x, y, w, h, shapeRotationDeg].every(Number.isFinite) || w <= 0 || h <= 0) return null;
  const matrix = ctx.getTransform();
  const deviceScale = Math.max(Math.hypot(matrix.a, matrix.b), Math.hypot(matrix.c, matrix.d));
  // The authored xfrm box need not contain a custom path or callout tail.
  // Record in authored units and allocate for its conservative painted bounds,
  // while keeping fillToRect and rotation anchored to the original tile.
  const operations = fill.path === 'shape' && outline ? recordOutline(outline, ctx, w, h) : undefined;
  const bounds = operations ? outlineBounds(operations, w, h) : { x: 0, y: 0, w, h };
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

  // Omitted fillToRect means the tile for rect (§20.1.8.31). Shape shading's
  // omitted focus uses the tile center, as observed on rect/ellipse/triangle/
  // arrow/star/custom controls; an explicit rectangle keeps all four edges.
  const focus = fill.fillToRect ?? (fill.path === 'shape'
    ? { l: .5, r: .5, t: .5, b: .5 } : {});
  const l = focus.l ?? 0; const t = focus.t ?? 0;
  // Sum authored opposing insets first: a collapsed 20%/80% focus must stay
  // a point despite subtraction rounding (1 - .8 can be smaller than .2).
  const r = l + (1 - (l + (focus.r ?? 0)));
  const b = t + (1 - (t + (focus.b ?? 0)));
  if (![l, t, r, b].every(Number.isFinite) || l > r || t > b) return null;
  let field: Uint8ClampedArray | undefined;
  let boundaryDistances: Float32Array | undefined;
  let clearance = 0;
  // Non-shape hosts (backgrounds, cells, text/line boxes, chart frames) have a
  // rectangular fill region. Shape hosts supply their fill-bearing geometry;
  // there is no parser/model reconstruction or geometry-name guessing here.
  if (operations) {
    // Record once: preset guides and custom paths must not be re-evaluated for
    // each contour (nor change their aspect-dependent geometry as they shrink).
    const path = typeof Path2D !== 'undefined' ? new Path2D() : undefined;
    if (path) {
      for (const op of operations) Reflect.apply(Reflect.get(path, op.method), path, op.args);
    }
    target.save();
    target.setTransform(sx, 0, 0, sy, 1 - bounds.x * sx, 1 - bounds.y * sy);
    const trace = () => {
      if (path) return;
      target.beginPath();
      for (const op of operations) {
        Reflect.apply(Reflect.get(target, op.method), target, op.args);
      }
    };
    trace();
    target.fillStyle = '#ffffff';
    if (path) target.fill(path); else target.fill();
    target.restore();
    boundaryDistances = interiorDistances(target.getImageData(0, 0, bw + 2, bh + 2).data, bw + 2, bh + 2);
    const fx = Math.max(0, Math.min(bw + 1, 1 + Math.floor(((l + r) / 2 * w - bounds.x) * sx)));
    const fy = Math.max(0, Math.min(bh + 1, 1 + Math.floor(((t + b) / 2 * h - bounds.y) * sy)));
    clearance = boundaryDistances[fy * (bw + 2) + fx];
    if (clearance === 0) {
      for (const distance of boundaryDistances) clearance = Math.max(clearance, distance);
    }
    target.fillStyle = '#ffffff';
    target.fillRect(0, 0, bw + 2, bh + 2);
    target.save();
    // Repaint opaque scalar values instead of nesting hundreds of clips or
    // compositing translucent stops. The distance pass protects notch edges;
    // colour/alpha are applied once below, after the field is complete.
    for (let level = CONTOURS - 1; level > 0; level--) {
      const f = level / CONTOURS;
      target.setTransform(sx * ((r - l) * (1 - f) + f), 0,
        0, sy * ((b - t) * (1 - f) + f),
        1 + sx * (l * w * (1 - f) - bounds.x), 1 + sy * (t * h * (1 - f) - bounds.y));
      trace();
      target.fillStyle = `rgb(${level},${level},${level})`;
      if (path) target.fill(path); else target.fill();
    }
    target.restore();
    field = target.getImageData(1, 1, bw, bh).data;
  }
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
      const insideFocus = u >= l && u <= r && v >= t && v <= b;
      const fieldX = (u * w - bounds.x) * sx; const fieldY = (v * h - bounds.y) * sy;
      const fieldCol = Math.min(bw - 1, Math.max(0, Math.floor(fieldX)));
      const fieldRow = Math.min(bh - 1, Math.max(0, Math.floor(fieldY)));
      const fieldOffset = (fieldRow * bw + fieldCol) * 4;
      const outside = fieldX < 0 || fieldX >= bw || fieldY < 0 || fieldY >= bh;
      let f = insideFocus ? 0 : field ? (outside ? 1 : field[fieldOffset] / CONTOURS)
        : Math.max(edgeDistance(u, l, r), edgeDistance(v, t, b));
      if (!insideFocus && boundaryDistances && clearance > 0 && !outside) {
        const bx = 1 + fieldCol; const by = 1 + fieldRow;
        f = Math.max(f, 1 - boundaryDistances[by * (bw + 2) + bx] / clearance);
      }
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
