import { describe, expect, it, vi } from 'vitest';
import {
  paintDrawingMLShape, resolveFill,
  type GradientFill, type DrawingMLShapeGeometry,
} from '@silurus/ooxml-core';
import type { Presentation, ShapeElement } from '@silurus/ooxml-pptx';
import type { Styles, Worksheet } from '@silurus/ooxml-xlsx';
import { renderViewport } from '../../xlsx/src/renderer';
import { renderSlideNode } from './render';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
// Three stops keep interpolation linear, so pixel values read the geometric
// gradient position directly: shade ≈ 255·s.
const fill: GradientFill = {
  fillType: 'gradient', gradType: 'radial', path: 'rect', angle: 0,
  fillToRect: { l: .5, r: .5, t: .5, b: .5 },
  stops: [{ position: 0, color: '000000' }, { position: .5, color: '808080' },
    { position: 1, color: 'FFFFFF' }],
};
const twoStop = [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }];

/** Independent reference: triangle fan from the focus over the outline edges
 * in path order, the last containing triangle wins; s = 1 - λ(focus). */
function fanShade(polygon: number[][], focus: number[], p: number[]): number {
  let shade = 1;
  for (let i = 0; i < polygon.length; i++) {
    const a = polygon[i]; const b = polygon[(i + 1) % polygon.length];
    const det = (a[0] - focus[0]) * (b[1] - focus[1]) - (b[0] - focus[0]) * (a[1] - focus[1]);
    if (Math.abs(det) < 1e-12) continue;
    const la = ((p[0] - focus[0]) * (b[1] - focus[1]) - (b[0] - focus[0]) * (p[1] - focus[1])) / det;
    const lb = ((a[0] - focus[0]) * (p[1] - focus[1]) - (p[0] - focus[0]) * (a[1] - focus[1])) / det;
    if (la >= 0 && lb >= 0 && la + lb <= 1) shade = la + lb;
  }
  return shade;
}

describe.skipIf(!skia)('DrawingML path-gradient pixels', () => {
  function canvas(w = 240, h = 120) {
    const c = new (skia as NonNullable<typeof skia>).Canvas(w, h);
    return { c, ctx: c.getContext('2d') as unknown as CanvasRenderingContext2D };
  }
  function pixel(ctx: CanvasRenderingContext2D, x: number, y: number) {
    return [...ctx.getImageData(x, y, 1, 1).data];
  }
  function paint(recipe: GradientFill, ctx: CanvasRenderingContext2D, w: number, h: number) {
    ctx.fillStyle = resolveFill(recipe, ctx, 0, 0, w, h) as CanvasPattern;
    ctx.fillRect(0, 0, w, h);
  }

  it('shades rect paths with box isolines toward the focus', () => {
    const { ctx } = canvas();
    paint(fill, ctx, 240, 120);
    // Horizontal, vertical and diagonal half-box points share one isoline.
    for (const [x, y] of [[60, 60], [120, 30], [60, 30], [180, 90]]) {
      expect(Math.abs(pixel(ctx, x, y)[0] - 128)).toBeLessThanOrEqual(3);
    }
  });

  it('places the focus at the fillToRect fixed point with no flat area', () => {
    const asymmetric = canvas().ctx;
    paint({ ...fill, fillToRect: { l: .2, r: .4, t: .25, b: .25 } }, asymmetric, 240, 120);
    // Focus x = l/(l+r) = 1/3 → (80, 60). (60, 60) lies inside the ECMA focus
    // rectangle yet is a quarter of the way to the left edge.
    expect(pixel(asymmetric, 80, 60)[0]).toBeLessThan(6);
    expect(Math.abs(pixel(asymmetric, 60, 60)[0] - 64)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(asymmetric, 160, 60)[0] - 128)).toBeLessThanOrEqual(3);
    const zero = canvas().ctx;
    paint({ ...fill, fillToRect: { l: 0, t: 0, r: 0, b: 0 } }, zero, 240, 120);
    expect(pixel(zero, 2, 2)[0]).toBeLessThan(8);
    expect(Math.abs(pixel(zero, 120, 30)[0] - 128)).toBeLessThanOrEqual(3);
    const omitted = canvas().ctx;
    paint({ ...fill, fillToRect: undefined }, omitted, 240, 120);
    expect(pixel(omitted, 120, 60)[0]).toBeLessThan(6);
  });

  it('applies the Office two-stop transfer and keeps other stop lists linear', () => {
    const { ctx } = canvas(400, 30);
    const linear = (stops: GradientFill['stops'], y: number) => {
      ctx.fillStyle = resolveFill({ fillType: 'gradient', gradType: 'linear', angle: 0, stops },
        ctx, 0, 0, 400, 30) as CanvasGradient;
      ctx.fillRect(0, y, 400, 10);
    };
    linear(twoStop, 0);
    linear([{ position: 0, color: 'FF0000' }, { position: 1, color: '00FF00' }], 10);
    linear([{ position: .2, color: '000000' }, { position: .8, color: 'FFFFFF' }], 20);
    // Φ-sigma weights then gamma 2.2: w(.25) = .1424, w(.5) = .5.
    expect(Math.abs(pixel(ctx, 100, 5)[0] - 105)).toBeLessThanOrEqual(1);
    expect(Math.abs(pixel(ctx, 200, 5)[0] - 186)).toBeLessThanOrEqual(1);
    expect(Math.abs(pixel(ctx, 300, 5)[0] - 238)).toBeLessThanOrEqual(1);
    const mid = pixel(ctx, 200, 15);
    expect(Math.abs(mid[0] - 186)).toBeLessThanOrEqual(1);
    expect(Math.abs(mid[1] - 186)).toBeLessThanOrEqual(1);
    expect(Math.abs(pixel(ctx, 200, 25)[0] - 128)).toBeLessThanOrEqual(1);
  });

  it('follows star-shaped outlines with a focus fan in path order', () => {
    const outline = [[0, 0], [100, 40], [200, 0], [160, 60], [200, 120], [100, 80], [0, 120], [40, 60]];
    const custom: DrawingMLShapeGeometry = { kind: 'custom', subpaths: [[
      ...outline.map(([x, y], index) => ({ cmd: index ? 'lineTo' as const : 'moveTo' as const, x: x / 200, y: y / 120 })),
      { cmd: 'close' },
    ]] };
    for (const [fillToRect, focus] of [
      [undefined, [100, 60]], [{ l: 0, t: 0, r: 1, b: 1 }, [0, 0]], [{ l: .3, t: .5, r: .1, b: .5 }, [150, 60]],
    ] as const) {
      const { ctx } = canvas(200, 120);
      paintDrawingMLShape(ctx, {
        rect: { x: 0, y: 0, w: 200, h: 120 }, geometry: custom, stroke: null,
        fill: { ...fill, path: 'shape', fillToRect }, transform: { rotationDeg: 0, flipH: false, flipV: false },
      }, 1);
      for (const p of [[100, 60], [70, 55], [130, 70], [60, 30], [150, 95], [100, 50], [45, 80]]) {
        const expected = 255 * fanShade(outline, focus as unknown as number[], [p[0] + .5, p[1] + .5]);
        expect(Math.abs(pixel(ctx, p[0], p[1])[0] - expected), JSON.stringify({ focus, p })).toBeLessThanOrEqual(5);
      }
    }
  });

  it('substitutes the circle path for outlines not star-shaped about their center', () => {
    for (const name of ['rightArrow', 'donut', 'chevron']) {
      for (const fillToRect of [undefined, { l: .2, r: .4, t: .25, b: .25 }]) {
        const plan = (path: string) => ({
          rect: { x: 0, y: 0, w: 120, h: 120 }, stroke: null,
          geometry: { kind: 'preset' as const, name, adjustments: [] },
          fill: { ...fill, path, fillToRect }, transform: { rotationDeg: 0, flipH: false, flipV: false },
        });
        const shape = canvas(120, 120).ctx; const circle = canvas(120, 120).ctx;
        paintDrawingMLShape(shape, plan('shape'), 1);
        paintDrawingMLShape(circle, plan('circle'), 1);
        expect(Buffer.from(shape.getImageData(0, 0, 120, 120).data)
          .equals(Buffer.from(circle.getImageData(0, 0, 120, 120).data)), name).toBe(true);
      }
    }
    // A convex preset is shaded by its own outline instead.
    const ellipse = canvas(120, 120).ctx;
    paintDrawingMLShape(ellipse, {
      rect: { x: 0, y: 0, w: 120, h: 120 }, stroke: null,
      geometry: { kind: 'preset', name: 'ellipse', adjustments: [] },
      fill: { ...fill, path: 'shape' }, transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    for (const [x, y] of [[90, 60], [60, 30], [81, 81]]) {
      expect(Math.abs(pixel(ellipse, x, y)[0] - 128)).toBeLessThanOrEqual(4);
    }
  });

  it('repeats path tiles unmirrored with a box-relative rect focus', () => {
    const { ctx } = canvas();
    paint({ ...fill, tileRect: { r: .5 }, flip: 'x', fillToRect: { l: 0, r: 1, t: 0, b: 1 } }, ctx, 240, 120);
    expect(pixel(ctx, 3, 3)[0]).toBeLessThan(12);
    expect(pixel(ctx, 123, 3)[0]).toBeLessThan(12);
    expect(pixel(ctx, 117, 117)[0]).toBeGreaterThan(243);
    const inset = canvas().ctx;
    paint({ ...fill, tileRect: { l: .25, t: .25, r: .25, b: .25 },
      fillToRect: { l: .25, t: .25, r: .75, b: .75 } }, inset, 240, 120);
    // Box focus (60, 30) is the tile's top-left corner, not its quarter point.
    expect(pixel(inset, 62, 32)[0]).toBeLessThan(12);
    expect(Math.abs(pixel(inset, 120, 60)[0] - 128)).toBeLessThanOrEqual(3);
  });

  it('keeps out-of-box rect shading opaque without moving the authored gradient frame', () => {
    const { ctx } = canvas();
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 120, h: 120 },
      geometry: { kind: 'custom', subpaths: [[
        { cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 2, y: 1 },
        { cmd: 'lineTo', x: 0, y: 1 }, { cmd: 'close' },
      ]] },
      fill, stroke: null,
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(pixel(ctx, 180, 105)).toEqual([255, 255, 255, 255]);
    expect(pixel(ctx, 60, 60)[0]).toBeLessThan(6);
    expect(Math.abs(pixel(ctx, 90, 90)[0] - 128)).toBeLessThanOrEqual(3);
  });

  it('observes coverage without changing the caller path or canvas state', () => {
    const { ctx } = canvas();
    ctx.translate(10, 5);
    ctx.globalAlpha = .5;
    ctx.lineWidth = 7;
    ctx.beginPath();
    ctx.rect(20, 20, 10, 10);
    ctx.fillStyle = resolveFill(fill, ctx, 0, 0, 120, 120, 0, undefined, undefined,
      (target, x, y, w, h) => {
        target.save(); target.translate(1, 1); target.restore();
        target.moveTo(x, y);
        target.lineTo(x + w * 2, y + h);
        target.lineTo(x, y + h);
        target.closePath();
      }) as CanvasPattern;
    ctx.fill();
    expect(pixel(ctx, 34, 29)[3]).toBe(128);
    expect(pixel(ctx, 80, 60)[3]).toBe(0);
    expect(ctx.globalAlpha).toBe(.5);
    expect(ctx.lineWidth).toBe(7);
    expect(ctx.getTransform().e).toBe(10);
  });

  it('interpolates translucent stops once even where triangles overlap', () => {
    const { ctx } = canvas();
    paint({ ...fill, stops: [{ position: 0, color: 'FF0000FF' }, { position: .5, color: 'FF000080' },
      { position: 1, color: 'FF000000' }] }, ctx, 240, 120);
    expect(pixel(ctx, 120, 60)[3]).toBeGreaterThan(250);
    expect(Math.abs(pixel(ctx, 60, 60)[3] - 128)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(ctx, 120, 30)[3] - 128)).toBeLessThanOrEqual(3);
  });

  it('frames a non-rotating shade by the device bounds of the rotated, mirrored host', () => {
    const { ctx } = canvas(160, 160);
    const recipe = { ...fill, fillToRect: { l: 0, r: 0, t: 0, b: 0 }, rotWithShape: false };
    ctx.translate(80, 80); ctx.rotate(Math.PI / 4); ctx.scale(-1, 1); ctx.translate(-50, -50);
    ctx.fillStyle = resolveFill(recipe, ctx, 0, 0, 100, 100, 45) as CanvasPattern;
    ctx.fillRect(0, 0, 100, 100);
    // The rotated square spans 80±70.7 on both device axes; the focus is the
    // device bounding box's top-left corner, not the host's local origin.
    const d = 70.71;
    expect(pixel(ctx, 80, Math.round(80 - d + 3))[0]).toBeLessThan(130);
    expect(pixel(ctx, Math.round(80 + d - 3), 80)[0]).toBeGreaterThan(240);
    expect(Math.abs(pixel(ctx, 80, 80)[0] - 128)).toBeLessThanOrEqual(4);
  });

  it('bounds auxiliary allocation even for very large authored extents', () => {
    const { ctx } = canvas(2, 2);
    const allocations: number[][] = [];
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    vi.stubGlobal('OffscreenCanvas', class extends Canvas {
      constructor(w: number, h: number) {
        allocations.push([w, h]);
        expect(w).toBeLessThanOrEqual(512); expect(h).toBeLessThanOrEqual(512);
        super(w, h);
      }
    });
    try {
      paint(fill, ctx, 1e9, 1e9);
      expect(pixel(ctx, 0, 0)[3]).toBe(255);
      expect(allocations.some(([w, h]) => w > 500 && h > 500)).toBe(true);
    } finally { vi.unstubAllGlobals(); }
  });

  it('keeps repaint work to one outline pass and one bounded pixel write', () => {
    const { ctx } = canvas(512, 512);
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    const work = { surfaces: 0, readbacks: 0, writes: 0, outlines: 0 };
    vi.stubGlobal('OffscreenCanvas', class extends Canvas {
      constructor(w: number, h: number) { super(w, h); work.surfaces++; }
      getContext(type: '2d') {
        const target = super.getContext(type);
        for (const method of ['getImageData', 'putImageData'] as const) {
          const original = target[method].bind(target);
          vi.spyOn(target, method).mockImplementation((...args: unknown[]) => {
            if (method === 'getImageData') work.readbacks++; else work.writes++;
            return Reflect.apply(original, target, args);
          });
        }
        return target;
      }
    });
    const edges = 720;
    const outline = (target: CanvasRenderingContext2D, x: number, y: number, w: number, h: number) => {
      work.outlines++;
      for (let i = 0; i < edges; i++) {
        const angle = i / edges * Math.PI * 2; const r = i % 2 ? .5 : .3;
        target[i ? 'lineTo' : 'moveTo'](x + w * (.5 + r * Math.cos(angle)), y + h * (.5 + r * Math.sin(angle)));
      }
      target.closePath();
    };
    try {
      for (let i = 0; i < 3; i++) {
        ctx.fillStyle = resolveFill({ ...fill, path: 'shape' }, ctx, 0, 0, 512, 512, 0,
          undefined, undefined, outline) as CanvasPattern;
        ctx.fillRect(0, 0, 512, 512);
        expect(pixel(ctx, 256, 256)[0]).toBeLessThan(4);
      }
      expect(work).toEqual({ surfaces: 3, readbacks: 0, writes: 3, outlines: 3 });
    } finally { vi.unstubAllGlobals(); vi.restoreAllMocks(); }
  });

  it('wires rect fills and host-box shape-path strokes through the retained DOCX painter', () => {
    const { ctx } = canvas();
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 240, h: 120 },
      geometry: { kind: 'preset', name: 'triangle', adjustments: [] },
      fill, stroke: { color: '000000', width: 8, fill: { ...fill, path: 'shape' } },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(Math.abs(pixel(ctx, 120, 30)[0] - 128)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(ctx, 120, 117)[0] - 242)).toBeLessThanOrEqual(3);
  });

  it('wires PPTX rectangular shading through the complete slide painter', async () => {
    const { c, ctx } = canvas();
    const px = (n: number) => n * 9525;
    const shape = {
      type: 'shape', x: 0, y: 0, width: px(240), height: px(120),
      rotation: 0, flipH: false, flipV: false, geometry: 'triangle',
      fill, stroke: { color: '000000', width: px(8), fill: { ...fill, path: 'shape' } }, textBody: null,
      custGeom: null, shadow: null,
    } as ShapeElement;
    const presentation = {
      slideWidth: px(240), slideHeight: px(120),
      slides: [{ index: 0, slideNumber: 1, background: null, elements: [shape] }],
      defaultTextColor: null, majorFont: null, minorFont: null,
    } as Presentation;
    await renderSlideNode(c, presentation, 0, { width: 240, dpr: 1 });
    expect(Math.abs(pixel(ctx, 120, 30)[0] - 128)).toBeLessThanOrEqual(3);
    // Shape-path strokes have no fill silhouette: every format shades the
    // host box, whose rectangle is star-shaped (box isolines, s = .95).
    expect(Math.abs(pixel(ctx, 120, 117)[0] - 242)).toBeLessThanOrEqual(3);
  });

  it('wires XLSX custom coordinates and excludes unfilled decorative paths', () => {
    const { ctx } = canvas();
    const triangle = [
      { op: 'moveTo', x: 120, y: 0 }, { op: 'lineTo', x: 240, y: 120 },
      { op: 'lineTo', x: 0, y: 120 }, { op: 'close' },
    ];
    const worksheet = {
      name: 'Sheet1', isChartSheet: true, rows: [], colWidths: {}, rowHeights: {},
      freezeRows: 0, freezeCols: 0,
      defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [], conditionalFormats: [],
      images: [], charts: [], defaultFontFamily: 'Calibri', defaultFontSize: 11,
      shapeGroups: [{ fromCol: 0, fromRow: 0, fromColOff: 0, fromRowOff: 0,
        toCol: 1, toRow: 1, toColOff: 0, toRowOff: 0, editAs: 'oneCell',
        nativeExtCx: 240 * 9525, nativeExtCy: 120 * 9525,
        shapes: [{ x: 0, y: 0, w: 1, h: 1, rot: 0, strokeColor: '000000', strokeWidth: 8 * 9525,
          strokeFill: { ...fill, path: 'shape' },
          fill, geom: { type: 'custom', paths: [
            { w: 240, h: 120, commands: triangle },
            { w: 240, h: 120, fill: 'none', stroke: false, commands: [
              { op: 'moveTo', x: 0, y: 0 }, { op: 'lineTo', x: 240, y: 0 },
              { op: 'lineTo', x: 240, y: 120 }, { op: 'lineTo', x: 0, y: 120 }, { op: 'close' },
            ] },
          ] } }],
      }],
    } as Worksheet;
    const styles = { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles;
    renderViewport(ctx, worksheet, styles, { row: 1, col: 1, rows: 1, cols: 1 });
    expect(Math.abs(pixel(ctx, 120, 30)[0] - 128)).toBeLessThanOrEqual(3);
    // Shape-path strokes have no fill silhouette: every format shades the
    // host box, whose rectangle is star-shaped (box isolines, s = .95).
    expect(Math.abs(pixel(ctx, 120, 117)[0] - 242)).toBeLessThanOrEqual(3);
    expect(pixel(ctx, 20, 10)).toEqual([255, 255, 255, 255]);
  });
});
