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

  it('fills the inscribed focus area with the first stop and shades the band to the outline', () => {
    const area = canvas().ctx;
    paint({ ...fill, fillToRect: { l: .2, r: .4, t: .25, b: .25 } }, area, 240, 120);
    // Focus rectangle x 48..144, y 30..90 is flat first stop (§20.1.8.31).
    for (const [x, y] of [[50, 32], [80, 60], [141, 87]]) expect(pixel(area, x, y)[0]).toBeLessThan(4);
    // Halfway from the focus edge to the outline: s = .5 → 128.
    expect(Math.abs(pixel(area, 24, 60)[0] - 125)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(area, 192, 60)[0] - 129)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(area, 96, 105)[0] - 130)).toBeLessThanOrEqual(3);
    const zero = canvas().ctx;
    paint({ ...fill, fillToRect: { l: 0, t: 0, r: 0, b: 0 } }, zero, 240, 120);
    for (const [x, y] of [[2, 2], [120, 60], [237, 117]]) expect(pixel(zero, x, y)).toEqual([0, 0, 0, 255]);
    const segment = canvas().ctx;
    paint({ ...fill, fillToRect: { l: .2, r: .2, t: .5, b: .5 } }, segment, 240, 120);
    for (const x of [50, 120, 190]) expect(pixel(segment, x, 60)[0]).toBeLessThan(6);
    expect(Math.abs(pixel(segment, 120, 30)[0] - 128)).toBeLessThanOrEqual(3);
    const omitted = canvas().ctx;
    paint({ ...fill, fillToRect: undefined }, omitted, 240, 120);
    expect(pixel(omitted, 120, 60)[0]).toBeLessThan(6);
  });

  it('keeps Canvas interpolation for two-stop gradients', () => {
    const { ctx } = canvas(400, 10);
    ctx.fillStyle = resolveFill({ fillType: 'gradient', gradType: 'linear', angle: 0, stops: twoStop },
      ctx, 0, 0, 400, 10) as CanvasGradient;
    ctx.fillRect(0, 0, 400, 10);
    expect(Math.abs(pixel(ctx, 200, 5)[0] - 128)).toBeLessThanOrEqual(1);
  });

  it('follows star-shaped outlines with a focus fan in path order', () => {
    const outline = [[0, 0], [100, 40], [200, 0], [160, 60], [200, 120], [100, 80], [0, 120], [40, 60]];
    const custom: DrawingMLShapeGeometry = { kind: 'custom', subpaths: [[
      ...outline.map(([x, y], index) => ({ cmd: index ? 'lineTo' as const : 'moveTo' as const, x: x / 200, y: y / 120 })),
      { cmd: 'close' },
    ]] };
    for (const [fillToRect, focus] of [
      [undefined, [100, 60]], [{ l: 0, t: 0, r: 1, b: 1 }, [0, 0]], [{ l: .75, t: .5, r: .25, b: .5 }, [150, 60]],
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

  it('shades outlines not star-shaped about their center as a circle about the focus point', () => {
    const arrow: DrawingMLShapeGeometry = { kind: 'preset', name: 'rightArrow', adjustments: [] };
    for (const [fillToRect, focus] of [
      [{ l: 0, t: 0, r: 0, b: 0 }, [0, 0]], [{ l: 0, t: 0, r: 1, b: 1 }, [0, 0]],
      [{ l: .2, t: 0, r: 0, b: 0 }, [120, 0]], [{ l: .2, r: .2, t: .2, b: .2 }, [60, 60]], [undefined, [60, 60]],
    ] as const) {
      const { ctx } = canvas(120, 120);
      paintDrawingMLShape(ctx, {
        rect: { x: 0, y: 0, w: 120, h: 120 }, geometry: arrow, stroke: null,
        fill: { ...fill, path: 'shape', fillToRect }, transform: { rotationDeg: 0, flipH: false, flipV: false },
      }, 1);
      const radius = Math.max(...[[0, 0], [120, 0], [0, 120], [120, 120]]
        .map(([x, y]) => Math.hypot(x - focus[0], y - focus[1])));
      for (const [x, y] of [[10, 40], [50, 60], [90, 50], [72, 30], [100, 62]]) {
        const expected = 255 * Math.min(1, Math.hypot(x + .5 - focus[0], y + .5 - focus[1]) / radius);
        expect(Math.abs(pixel(ctx, x, y)[0] - expected), JSON.stringify({ fillToRect, x, y })).toBeLessThanOrEqual(3);
      }
    }
    // A convex ellipse instead keeps its inscribed focus area.
    const ellipse = canvas(120, 120).ctx;
    paintDrawingMLShape(ellipse, {
      rect: { x: 0, y: 0, w: 120, h: 120 }, stroke: null,
      geometry: { kind: 'preset', name: 'ellipse', adjustments: [] },
      fill: { ...fill, path: 'shape', fillToRect: { l: .2, r: .2, t: .2, b: .2 } },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(pixel(ellipse, 60, 60)[0]).toBeLessThan(4);
    expect(pixel(ellipse, 60 + 33, 60)[0]).toBeLessThan(4);
    // Contour radius .3 + .2s: radius .4 of the box is s = .5.
    expect(Math.abs(pixel(ellipse, 60 + 48, 60)[0] - 133)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(ellipse, 60, 60 - 48)[0] - 123)).toBeLessThanOrEqual(3);
  });

  it('keeps wide gradient strokes opaque outside the host box', () => {
    for (const path of ['rect', 'shape'] as const) {
      const { ctx } = canvas(200, 200);
      ctx.lineWidth = 20;
      ctx.strokeStyle = resolveFill({ ...fill, path }, ctx, 50, 50, 100, 100) as CanvasPattern;
      ctx.strokeRect(50, 50, 100, 100);
      // main painted these pixels opaque; the shade beyond the frame is the outer stop.
      expect(pixel(ctx, 41, 100)).toEqual([255, 255, 255, 255]);
      expect(pixel(ctx, 100, 158)).toEqual([255, 255, 255, 255]);
      expect(pixel(ctx, 158, 158)[3]).toBe(255);
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

  it('interpolates translucent stops once', () => {
    const { ctx } = canvas();
    paint({ ...fill, stops: [{ position: 0, color: 'FF0000FF' }, { position: .5, color: 'FF000080' },
      { position: 1, color: 'FF000000' }] }, ctx, 240, 120);
    expect(pixel(ctx, 120, 60)[3]).toBeGreaterThan(250);
    expect(Math.abs(pixel(ctx, 60, 60)[3] - 128)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(ctx, 120, 30)[3] - 128)).toBeLessThanOrEqual(3);
  });

  it('frames a non-rotating shade by the device bounds of the rotated, mirrored host', () => {
    const { ctx } = canvas(160, 160);
    const recipe = { ...fill, fillToRect: { l: 0, r: 1, t: 0, b: 1 }, rotWithShape: false };
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
        // The shade raster, or the 1024×1 stop ramp.
        expect(w * h).toBeLessThanOrEqual(512 * 512); expect(Math.max(w, h)).toBeLessThanOrEqual(1024);
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
      // Per paint: the shade raster plus the 1024×1 stop ramp and its readback.
      expect(work).toEqual({ surfaces: 6, readbacks: 3, writes: 3, outlines: 3 });
    } finally { vi.unstubAllGlobals(); vi.restoreAllMocks(); }
  });

  it('bounds corner-focus work on very complex star outlines', () => {
    const { ctx } = canvas(512, 512);
    for (const edges of [20000, 200000]) {
      let outlines = 0;
      const outline = (target: CanvasRenderingContext2D, x: number, y: number, w: number, h: number) => {
        outlines++;
        for (let i = 0; i < edges; i++) {
          const angle = i / edges * Math.PI * 2; const r = i % 2 ? .5 : .03;
          target[i ? 'lineTo' : 'moveTo'](x + w * (.5 + r * Math.cos(angle)), y + h * (.5 + r * Math.sin(angle)));
        }
        target.closePath();
      };
      const started = performance.now();
      ctx.fillStyle = resolveFill({ ...fill, path: 'shape', fillToRect: { l: 0, t: 0, r: 1, b: 1 } },
        ctx, 0, 0, 512, 512, 0, undefined, undefined, outline) as CanvasPattern;
      // O(edges × rows + pixels) with edges clustered to ≤ 32768: a generous
      // wall-clock guard against the former O(edges × pixels) fan scan.
      expect(performance.now() - started).toBeLessThan(2000);
      expect(outlines).toBe(1);
      ctx.fillRect(0, 0, 512, 512);
      expect(pixel(ctx, 2, 2)[0]).toBeLessThan(12);
    }
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
