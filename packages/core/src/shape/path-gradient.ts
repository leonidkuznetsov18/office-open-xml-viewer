import type { GradientFill, GradientStop } from '../types/common';
import { createAuxCanvasForContext } from '../canvas/aux-canvas';
import { officeGradientStops } from './gradient-transfer';

/** Append the fill-bearing outline; coordinates belong to the supplied box.
 * The caller owns beginPath. Decorative, unfilled paths must be excluded. */
export type FillOutline = (
  ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number,
) => void;

export interface ShadeBox { x: number; y: number; w: number; h: number }
type Point = [number, number];
type Matrix = [number, number, number, number, number, number];

// Resource policy, not an OOXML rule: the shade raster is at most 512² device
// pixels. Each paint flattens the outline once, buckets its edges by angle
// around the focus and makes one bounded pixel pass; a pixel tests only the
// edges spanning its angle. No geometry or raster is retained between paints.
const MAX_EDGE = 512;
const MARGIN = 2;
const BEZIER_SEGMENTS = 16;
const ARC_STEP = Math.PI / 32;

const apply = (m: Matrix, x: number, y: number): Point =>
  [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const multiply = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5],
];
function invert(m: Matrix): Matrix | null {
  const det = m[0] * m[3] - m[1] * m[2];
  if (!Number.isFinite(det) || Math.abs(det) < 1e-12) return null;
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det,
    (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}

/** Flatten the fill outline into closed polygons (fill subpaths are closed
 * implicitly). The recorder tracks the outline's own transforms and never
 * forwards drawing or state calls to the caller's context. */
export function flattenFillOutline(
  outline: FillOutline, ctx: CanvasRenderingContext2D, box: ShadeBox,
): Point[][] {
  const polygons: Point[][] = [];
  let current: Point[] = [];
  let matrix: Matrix = [1, 0, 0, 1, 0, 0];
  const stack: Matrix[] = [];
  const finish = () => { if (current.length >= 3) polygons.push(current); current = []; };
  const lastRaw = (): Point | undefined => {
    const last = current[current.length - 1];
    const inverse = last ? invert(matrix) : null;
    return last && inverse ? apply(inverse, last[0], last[1]) : undefined;
  };
  const point = (x: number, y: number) => {
    if (Number.isFinite(x) && Number.isFinite(y)) current.push(apply(matrix, x, y));
  };
  const ellipse = (cx: number, cy: number, rx: number, ry: number, rotation: number,
    start: number, end: number, ccw: boolean) => {
    const full = Math.PI * 2;
    let sweep = end - start;
    if (!ccw && sweep >= full) sweep = full;
    else if (ccw && -sweep >= full) sweep = -full;
    else if (!ccw) { sweep %= full; if (sweep < 0) sweep += full; }
    else { sweep %= full; if (sweep > 0) sweep -= full; }
    const steps = Math.max(1, Math.ceil(Math.abs(sweep) / ARC_STEP));
    const cos = Math.cos(rotation); const sin = Math.sin(rotation);
    for (let i = 0; i <= steps; i++) {
      const angle = start + sweep * i / steps;
      const ex = rx * Math.cos(angle); const ey = ry * Math.sin(angle);
      point(cx + ex * cos - ey * sin, cy + ex * sin + ey * cos);
    }
  };
  const curve = (order: 2 | 3, c: number[]) => {
    const p0 = lastRaw() ?? [c[0], c[1]];
    if (current.length === 0) point(c[0], c[1]);
    for (let i = 1; i <= BEZIER_SEGMENTS; i++) {
      const t = i / BEZIER_SEGMENTS; const u = 1 - t;
      if (order === 2) {
        point(u * u * p0[0] + 2 * u * t * c[0] + t * t * c[2],
          u * u * p0[1] + 2 * u * t * c[1] + t * t * c[3]);
      } else {
        point(u * u * u * p0[0] + 3 * u * u * t * c[0] + 3 * u * t * t * c[2] + t * t * t * c[4],
          u * u * u * p0[1] + 3 * u * u * t * c[1] + 3 * u * t * t * c[3] + t * t * t * c[5]);
      }
    }
  };
  const roundRect = (x: number, y: number, w: number, h: number, radii?: unknown) => {
    const list = Array.isArray(radii) ? radii : [radii ?? 0];
    const radius = (value: unknown): number => {
      const r = typeof value === 'number' ? value
        : typeof value === 'object' && value !== null ? Number((value as { x?: number }).x ?? 0) : 0;
      return Math.max(0, Math.min(Math.abs(w) / 2, Math.abs(h) / 2, Number.isFinite(r) ? r : 0));
    };
    const [tl, tr, br, bl] = list.length === 1 ? [list[0], list[0], list[0], list[0]]
      : list.length === 2 ? [list[0], list[1], list[0], list[1]]
        : list.length === 3 ? [list[0], list[1], list[2], list[1]] : list;
    finish();
    const corners: Array<[number, number, number, number]> = [
      [x + w - radius(tr), y + radius(tr), radius(tr), -Math.PI / 2],
      [x + w - radius(br), y + h - radius(br), radius(br), 0],
      [x + radius(bl), y + h - radius(bl), radius(bl), Math.PI / 2],
      [x + radius(tl), y + radius(tl), radius(tl), Math.PI],
    ];
    for (const [cx, cy, r, start] of corners) ellipse(cx, cy, r, r, 0, start, start + Math.PI / 2, false);
    finish();
    point(x, y);
  };
  const methods: Record<string, (...a: number[]) => void> = {
    beginPath: () => { polygons.length = 0; current = []; },
    moveTo: (x, y) => { finish(); point(x, y); },
    lineTo: (x, y) => point(x, y),
    closePath: () => { const first = current[0]; finish(); if (first) current.push(first); },
    bezierCurveTo: (...c) => curve(3, c),
    quadraticCurveTo: (...c) => curve(2, c),
    // Not used by DrawingML geometry builders; keeping the corner point keeps
    // the flattened outline conservative rather than dropping the segment.
    arcTo: (x1, y1) => point(x1, y1),
    arc: (cx, cy, r, start, end, ccw) => ellipse(cx, cy, r, r, 0, start, end, Boolean(ccw)),
    ellipse: (cx, cy, rx, ry, rotation, start, end, ccw) =>
      ellipse(cx, cy, rx, ry, rotation, start, end, Boolean(ccw)),
    rect: (x, y, w, h) => {
      finish(); point(x, y); point(x + w, y); point(x + w, y + h); point(x, y + h); finish(); point(x, y);
    },
    roundRect: roundRect as unknown as (...a: number[]) => void,
    save: () => { stack.push(matrix); },
    restore: () => { matrix = stack.pop() ?? matrix; },
    translate: (x, y) => { matrix = multiply(matrix, [1, 0, 0, 1, x, y]); },
    scale: (x, y) => { matrix = multiply(matrix, [x, 0, 0, y, 0, 0]); },
    rotate: (angle) => {
      const c = Math.cos(angle); const s = Math.sin(angle);
      matrix = multiply(matrix, [c, s, -s, c, 0, 0]);
    },
    transform: (a, b, c, d, e, f) => { matrix = multiply(matrix, [a, b, c, d, e, f]); },
    setTransform: (a, b, c, d, e, f) => { matrix = [a, b, c, d, e, f]; },
    resetTransform: () => { matrix = [1, 0, 0, 1, 0, 0]; },
  };
  const recorder = new Proxy(ctx, {
    get(target, property) {
      if (typeof property === 'string' && Object.prototype.hasOwnProperty.call(methods, property)) return methods[property];
      const value = Reflect.get(target, property);
      return typeof value === 'function' ? () => undefined : value;
    },
    set() { return true; },
  });
  outline(recorder, box.x, box.y, box.w, box.h);
  finish();
  return polygons;
}

/** Strictly star-shaped about `center`: every edge turns the same way around
 * it and the outline winds exactly once (one region, center strictly inside
 * its kernel). Holes, several loops and centers on the outline fail. */
export function isStrictlyStarShaped(polygons: Point[][], center: Point, box: ShadeBox): boolean {
  const size = Math.hypot(box.w, box.h);
  if (!(size > 0)) return false;
  const epsilon = size * 1e-9;
  let sign = 0; let turning = 0;
  for (const polygon of polygons) {
    for (let i = 0; i < polygon.length; i++) {
      const a = polygon[i]; const b = polygon[(i + 1) % polygon.length];
      const length = Math.hypot(b[0] - a[0], b[1] - a[1]);
      // Arc joins leave floating-point slivers; they carry no direction.
      if (length <= epsilon) continue;
      const ax = a[0] - center[0]; const ay = a[1] - center[1];
      const bx = b[0] - center[0]; const by = b[1] - center[1];
      const cross = ax * by - ay * bx;
      // The center must not lie on any edge's line (distance cross/length).
      if (Math.abs(cross) / length <= epsilon) return false;
      if (sign !== 0 && Math.sign(cross) !== sign) return false;
      sign = Math.sign(cross);
      turning += Math.atan2(cross, ax * bx + ay * by);
    }
  }
  return sign !== 0 && Math.abs(Math.abs(turning) - Math.PI * 2) < 1e-6;
}

/** Office focus point of a path shade, normalized to its frame. */
export function pathShadeFocus(fill: GradientFill): Point {
  const rect = fill.fillToRect;
  // Omitted fillToRect centers the shade. Otherwise the focus is the fixed
  // point of the frame→fillToRect scaling, l/(l+r) and t/(t+b). The identity
  // mapping (all zero) has no fixed point; Office then uses the frame origin.
  if (!rect) return [0.5, 0.5];
  const axis = (low = 0, high = 0) => {
    const sum = low + high;
    return Number.isFinite(sum) && sum !== 0 ? low / sum : 0;
  };
  return [axis(rect.l, rect.r), axis(rect.t, rect.b)];
}

function rgbaChannels(color: string): [number, number, number, number] {
  const hex = color.charCodeAt(0) === 35 ? color.slice(1) : color;
  const value = (offset: number) => Number.parseInt(hex.slice(offset, offset + 2), 16);
  return [value(0), value(2), value(4), hex.length >= 8 ? value(6) : 255];
}

const rectangle = (box: ShadeBox): Point[] =>
  [[box.x, box.y], [box.x + box.w, box.y], [box.x + box.w, box.y + box.h], [box.x, box.y + box.h]];

function boundsOf(points: Point[]): ShadeBox {
  let left = Infinity; let top = Infinity; let right = -Infinity; let bottom = -Infinity;
  for (const [x, y] of points) {
    left = Math.min(left, x); top = Math.min(top, y);
    right = Math.max(right, x); bottom = Math.max(bottom, y);
  }
  return { x: left, y: top, w: right - left, h: bottom - top };
}

const LUT_STEPS = 4095;

/** RGBA lookup over gradient position, interpolating the (Office-expanded)
 * stops linearly; positions outside the stop range take the end stops. */
function colorTable(stops: readonly GradientStop[]): Uint8ClampedArray {
  const table = new Uint8ClampedArray((LUT_STEPS + 1) * 4);
  const parsed = stops.map(stop => ({ position: stop.position, rgba: rgbaChannels(stop.color) }));
  let index = 0;
  for (let i = 0; i <= LUT_STEPS; i++) {
    const t = i / LUT_STEPS;
    while (index < parsed.length - 1 && parsed[index + 1].position < t) index++;
    const low = parsed[Math.max(0, Math.min(index, parsed.length - 1))];
    const high = parsed[Math.min(index + 1, parsed.length - 1)];
    const span = high.position - low.position;
    const a = t <= low.position ? 0 : t >= high.position || span <= 0 ? 1 : (t - low.position) / span;
    for (let k = 0; k < 4; k++) table[i * 4 + k] = low.rgba[k] + a * (high.rgba[k] - low.rgba[k]);
  }
  return table;
}

/**
 * Gradient position of the focus fan: the last edge (in path order) whose
 * triangle (focus, a, b) contains the point gives 1 - λ(focus), the homothety
 * ratio toward the focus; uncovered points are 1. Candidate edges are bucketed
 * by angle around the focus, so a point tests only edges spanning its angle.
 */
function fanField(polygons: Point[][], focus: Point): (x: number, y: number) => number {
  const edges: Array<{ ax: number; ay: number; bx: number; by: number; inverse: number }> = [];
  for (const polygon of polygons) {
    for (let i = 0; i < polygon.length; i++) {
      const a = polygon[i]; const b = polygon[(i + 1) % polygon.length];
      const ax = a[0] - focus[0]; const ay = a[1] - focus[1];
      const bx = b[0] - focus[0]; const by = b[1] - focus[1];
      const det = ax * by - ay * bx;
      // A triangle without area (focus on the edge's line) shades nothing.
      if (!(Math.abs(det) > 1e-9 * Math.max(1, Math.hypot(ax, ay) * Math.hypot(bx, by)))) continue;
      edges.push({ ax, ay, bx, by, inverse: 1 / det });
    }
  }
  const bins = Math.min(4096, Math.max(64, edges.length * 4));
  const binOf = (angle: number) => Math.min(bins - 1, Math.max(0,
    Math.floor(((angle + Math.PI) / (Math.PI * 2)) * bins)));
  const buckets: number[][] = Array.from({ length: bins }, () => []);
  edges.forEach((edge, index) => {
    // Each triangle spans less than π around the focus; walk its bins from a
    // to b the short way, inclusive of both partial end bins.
    const start = Math.atan2(edge.ay, edge.ax);
    let sweep = Math.atan2(edge.by, edge.bx) - start;
    if (sweep > Math.PI) sweep -= Math.PI * 2;
    if (sweep < -Math.PI) sweep += Math.PI * 2;
    let low = sweep >= 0 ? start : start + sweep;
    if (low < -Math.PI) low += Math.PI * 2;
    const from = binOf(low);
    const count = Math.ceil(Math.abs(sweep) / (Math.PI * 2) * bins) + 1;
    for (let k = 0; k <= count; k++) buckets[(from + k) % bins].push(index);
  });
  return (x: number, y: number) => {
    const px = x - focus[0]; const py = y - focus[1];
    if (px === 0 && py === 0) return 0;
    const bucket = buckets[binOf(Math.atan2(py, px))];
    for (let k = bucket.length - 1; k >= 0; k--) {
      const e = edges[bucket[k]];
      const la = (px * e.by - py * e.bx) * e.inverse;
      const lb = (e.ax * py - e.ay * px) * e.inverse;
      if (la >= -1e-12 && lb >= -1e-12 && la + lb <= 1 + 1e-12) return la + lb;
    }
    return 1;
  };
}

/**
 * Path shade for `path="rect"` and `path="shape"` (ECMA-376 §20.1.8.46). The
 * standard requires shape-following contours but defines no interpolation;
 * this is the Office model measured from PowerPoint's own rasters (#1599):
 *
 * - The field is a fan of triangles from the focus to every outline edge,
 *   painted in path order (a later triangle wins). Inside a triangle the
 *   gradient position is the homothety ratio toward the focus: 0 at the
 *   focus, 1 on that edge. For an outline star-shaped about the focus this is
 *   exactly "contours are scaled copies of the outline".
 * - `path="rect"` uses the frame rectangle as outline (box isolines) and
 *   `path="shape"` the fill-bearing geometry outline.
 * - The focus is {@link pathShadeFocus}; there is no flat first-stop area,
 *   unlike §20.1.8.31 and [MS-OE376] §2.1.1377.
 * - rotWithShape=false: the frame is the axis-aligned device bounding box of
 *   the transformed outline, unmirrored.
 *
 * Returns 'circle' for a shape outline that is not strictly star-shaped about
 * its box center: Office substitutes the circle path for those geometries.
 * Returns null when the host cannot allocate the shade raster.
 */
export function resolvePathShade(
  fill: GradientFill, ctx: CanvasRenderingContext2D,
  frame: ShadeBox, shapeBox: ShadeBox, outline?: FillOutline,
  focusOverride?: Point,
): CanvasPattern | 'circle' | null {
  if (![frame, shapeBox].every(box => [box.x, box.y, box.w, box.h].every(Number.isFinite)
    && box.w > 0 && box.h > 0)) return null;
  const outlinePolygons = outline ? flattenFillOutline(outline, ctx, shapeBox) : [];
  const coverage = outlinePolygons.length > 0 ? outlinePolygons : [rectangle(shapeBox)];
  if (fill.path === 'shape' && !isStrictlyStarShaped(
    coverage, [shapeBox.x + shapeBox.w / 2, shapeBox.y + shapeBox.h / 2], shapeBox,
  )) return 'circle';
  if (typeof ctx.getTransform !== 'function') return null;

  const t = ctx.getTransform();
  const user: Matrix = [t.a, t.b, t.c, t.d, t.e, t.f];
  const userInverse = user.every(Number.isFinite) ? invert(user) : null;
  if (!userInverse) return null;
  const fixed = fill.rotWithShape === false;
  // Field space is user space, or device space when the shade neither rotates
  // nor mirrors with the shape.
  const toField = (p: Point): Point => (fixed ? apply(user, p[0], p[1]) : p);
  const coverageField = coverage.map(polygon => polygon.map(toField));
  const fieldFrame = fixed
    ? boundsOf(outlinePolygons.length > 0 ? coverageField.flat() : rectangle(frame).map(toField))
    : frame;
  if (!(fieldFrame.w > 0 && fieldFrame.h > 0)) return null;
  const field = fill.path === 'shape' ? coverageField : [rectangle(fieldFrame)];
  const [fx, fy] = focusOverride ?? pathShadeFocus(fill);
  const focus: Point = [fieldFrame.x + fieldFrame.w * fx, fieldFrame.y + fieldFrame.h * fy];
  if (!focus.every(Number.isFinite)) return null;

  // Allocate for everything the fill may paint (callout tails, custom paths
  // outside the xfrm box) plus the frame itself.
  const toDevice = (p: Point): Point => (fixed ? p : apply(user, p[0], p[1]));
  const device = boundsOf([...coverageField.flat(), ...rectangle(fieldFrame)].map(toDevice));
  if (![device.x, device.y, device.w, device.h].every(Number.isFinite)) return null;
  const k = Math.min(1, (MAX_EDGE - 2 * MARGIN - 1) / Math.max(device.w, 1e-9),
    (MAX_EDGE - 2 * MARGIN - 1) / Math.max(device.h, 1e-9));
  // At full resolution the raster is aligned to the device pixel grid, so the
  // pattern is sampled without resampling blur.
  const originX = k === 1 ? Math.floor(device.x) : device.x;
  const originY = k === 1 ? Math.floor(device.y) : device.y;
  const inner = MAX_EDGE - 2 * MARGIN;
  const bw = Math.max(1, Math.min(inner, Math.ceil((device.x + device.w - originX) * k))) + 2 * MARGIN;
  const bh = Math.max(1, Math.min(inner, Math.ceil((device.y + device.h - originY) * k))) + 2 * MARGIN;
  const auxToDevice: Matrix = [1 / k, 0, 0, 1 / k, originX - MARGIN / k, originY - MARGIN / k];
  const deviceToAux = invert(auxToDevice);
  if (!deviceToAux) return null;
  const fieldToAux = fixed ? deviceToAux : multiply(deviceToAux, user);

  const stops = officeGradientStops([...fill.stops].sort((a, b) => a.position - b.position));
  const surface = createAuxCanvasForContext(ctx, bw, bh);
  const target = surface?.getContext('2d') as CanvasRenderingContext2D | null | undefined;
  if (!surface || !target || typeof target.createImageData !== 'function') return null;
  const lut = colorTable(stops);
  const shade = fanField(field.map(polygon => polygon.map(p => apply(fieldToAux, p[0], p[1]))),
    apply(fieldToAux, focus[0], focus[1]));
  const pixels = target.createImageData(bw, bh);
  const data = pixels.data;
  for (let row = 0, offset = 0; row < bh; row++) {
    for (let col = 0; col < bw; col++, offset += 4) {
      const entry = Math.round(Math.min(1, Math.max(0, shade(col + .5, row + .5))) * LUT_STEPS) * 4;
      data[offset] = lut[entry]; data[offset + 1] = lut[entry + 1];
      data[offset + 2] = lut[entry + 2]; data[offset + 3] = lut[entry + 3];
    }
  }
  target.putImageData(pixels, 0, 0);
  const pattern = ctx.createPattern(surface, 'no-repeat');
  if (!pattern || typeof pattern.setTransform !== 'function') return null;
  const auxToUser = multiply(userInverse, auxToDevice);
  pattern.setTransform({
    a: auxToUser[0], b: auxToUser[1], c: auxToUser[2], d: auxToUser[3], e: auxToUser[4], f: auxToUser[5],
  });
  return pattern;
}
