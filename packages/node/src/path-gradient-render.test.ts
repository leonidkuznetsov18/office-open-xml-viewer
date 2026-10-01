import { describe, expect, it, vi } from 'vitest';
import {
  paintDrawingMLShape, resolveFill, buildCustomPath, buildPresetGeometryFillPath, buildShapePath,
  type GradientFill, type DrawingMLShapeGeometry, type PathCmd,
} from '@silurus/ooxml-core';
import type { Presentation, ShapeElement } from '@silurus/ooxml-pptx';
import type { Styles, Worksheet } from '@silurus/ooxml-xlsx';
import { renderViewport } from '../../xlsx/src/renderer';
import { renderSlideNode } from './render';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
const fill: GradientFill = {
  fillType: 'gradient', gradType: 'radial', path: 'rect', angle: 0,
  fillToRect: { l: .5, r: .5, t: .5, b: .5 },
  stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
};

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

  it('uses normalized box distance on a wide rectangle, including diagonal isolines', () => {
    const { ctx } = canvas();
    paint(fill, ctx, 240, 120);
    // The half-size box has the same shade on its horizontal, vertical and
    // diagonal points. A circular or elliptical approximation fails this.
    const shades = [[60, 60], [120, 30], [60, 30]].map(([x, y]) => pixel(ctx, x, y)[0]);
    shades.forEach(shade => expect(shade).toBeGreaterThan(124));
    shades.forEach(shade => expect(shade).toBeLessThan(132));
  });

  it('uses the ECMA focus rectangle and interpolates alpha once', () => {
    const { ctx } = canvas();
    paint({ ...fill, fillToRect: { l: .2, r: .4, t: .25, b: .25 },
      stops: [{ position: 0, color: 'FF000080' }, { position: 1, color: '0000FF00' }],
    }, ctx, 240, 120);
    expect(pixel(ctx, 75, 40)).toEqual([255, 0, 0, 128]);
    expect(pixel(ctx, 130, 80)).toEqual([255, 0, 0, 128]);
    expect(pixel(ctx, 192, 60)[3]).toBeGreaterThan(60);
    expect(pixel(ctx, 192, 60)[3]).toBeLessThan(67);
  });

  it('retains the legacy radial shape approximation across convex and concave geometry', () => {
    const vertices = [[0, 0], [1, 0], [1, 1], [.65, 1], [.65, .4], [.35, .4], [.35, 1], [0, 1]];
    const geometries: DrawingMLShapeGeometry[] = [
      ...['rect', 'ellipse', 'rightArrow', 'triangle', 'star5'].map(name => ({
        kind: 'preset' as const, name, adjustments: [],
      })),
      { kind: 'custom', subpaths: [[
        ...vertices.map(([x, y], index) => ({ cmd: index ? 'lineTo' as const : 'moveTo' as const, x, y })),
        { cmd: 'close' },
      ]] },
    ];
    for (const geometry of geometries) {
      for (const focus of [undefined, fill.fillToRect, { l: .2, r: .4, t: .25, b: .25 }]) {
        const actual = canvas(120, 120).ctx;
        const previous = canvas(120, 120).ctx;
        const plan = {
          rect: { x: 0, y: 0, w: 120, h: 120 }, geometry, stroke: null,
          transform: { rotationDeg: 0, flipH: false, flipV: false },
        };
        paintDrawingMLShape(actual, { ...plan, fill: { ...fill, path: 'shape', fillToRect: focus } }, 1);
        // Independent legacy contract: a native radial gradient focused at
        // the authored rectangle's center, reaching its farthest box corner.
        const cx = 60 * (1 + (focus?.l ?? 0) - (focus?.r ?? 0));
        const cy = 60 * (1 + (focus?.t ?? 0) - (focus?.b ?? 0));
        const radius = Math.max(...[[0, 0], [120, 0], [0, 120], [120, 120]]
          .map(([x, y]) => Math.hypot(x - cx, y - cy)));
        const gradient = previous.createRadialGradient(cx, cy, 0, cx, cy, radius);
        gradient.addColorStop(0, '#000000'); gradient.addColorStop(1, '#FFFFFF');
        previous.fillStyle = gradient;
        previous.beginPath();
        if (geometry.kind === 'preset') {
          if (!buildPresetGeometryFillPath(previous, geometry.name, 0, 0, 120, 120, [])) {
            buildShapePath(previous, geometry.name, 0, 0, 120, 120);
          }
        } else {
          buildCustomPath(previous, geometry.subpaths as PathCmd[][], 0, 0, 120, 120);
        }
        previous.fill();
        const actualBytes = Buffer.from(actual.getImageData(0, 0, 120, 120).data);
        const previousBytes = Buffer.from(previous.getImageData(0, 0, 120, 120).data);
        expect(actualBytes.equals(previousBytes), JSON.stringify({ geometry: geometry.kind === 'preset' ? geometry.name : 'custom', focus })).toBe(true);
      }
    }
  });

  it('repeats box tiles with reflection, retaining all authored stops and plateaus', () => {
    const { ctx } = canvas();
    paint({ ...fill, tileRect: { r: .5 }, flip: 'x',
      fillToRect: { l: 0, r: 1, t: 0, b: 1 },
      stops: [{ position: .2, color: 'FF0000' }, { position: .5, color: '00FF00' },
        { position: .8, color: '0000FF' }],
    }, ctx, 240, 120);
    expect(pixel(ctx, 12, 12)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 227, 12)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 110, 12)).toEqual([0, 0, 255, 255]);
    expect(pixel(ctx, 60, 12)[1]).toBeGreaterThan(245);
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
    // Enlarging coverage to 240px must leave the authored focus at (60,60).
    expect(pixel(ctx, 60, 60)[0]).toBeLessThan(5);
    expect(pixel(ctx, 90, 90)[0]).toBeGreaterThan(125);
    expect(pixel(ctx, 90, 90)[0]).toBeLessThan(133);
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
  });

  it('counter-rotates the shade within a reflected host without leaving transparent corners', () => {
    const { ctx } = canvas(160, 160);
    const recipe = { ...fill, fillToRect: { l: 0, r: 1, t: 0, b: 1 }, rotWithShape: false };
    ctx.translate(80, 80); ctx.rotate(Math.PI / 2); ctx.scale(-1, 1); ctx.translate(-80, -80);
    ctx.fillStyle = resolveFill(recipe, ctx, 0, 0, 160, 160, 90) as CanvasPattern;
    ctx.fillRect(0, 0, 160, 160);
    expect(pixel(ctx, 120, 40)[3]).toBe(255);
    expect(pixel(ctx, 120, 40)[0]).toBeGreaterThan(185);
    expect(pixel(ctx, 120, 40)[0]).toBeLessThan(196);
    expect(pixel(ctx, 1, 1)[3]).toBe(255);
  });

  it('bounds auxiliary allocation even for very large authored extents', () => {
    const { ctx } = canvas(2, 2);
    const allocations: number[][] = [];
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    vi.stubGlobal('OffscreenCanvas', class extends Canvas {
      constructor(w: number, h: number) {
        allocations.push([w, h]);
        expect(w).toBeLessThanOrEqual(1024); expect(h).toBeLessThanOrEqual(512);
        super(w, h);
      }
    });
    try {
      paint(fill, ctx, 1e9, 1e9);
      expect(pixel(ctx, 0, 0)[3]).toBe(255);
      expect(allocations).toContainEqual([512, 512]);
    } finally { vi.unstubAllGlobals(); }
  });

  it('keeps repeated rect shading to one ramp pass even for a complex outline', () => {
    const { ctx } = canvas(512, 512);
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    const work = { fills: 0, readbacks: 0, writes: 0 };
    vi.stubGlobal('OffscreenCanvas', class extends Canvas {
      getContext(type: '2d') {
        const target = super.getContext(type);
        for (const method of ['fill', 'fillRect'] as const) {
          const original = target[method].bind(target);
          vi.spyOn(target, method).mockImplementation((...args: unknown[]) => {
            work.fills++;
            return Reflect.apply(original, target, args);
          });
        }
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
    let commands = 0;
    const outline = (target: CanvasRenderingContext2D) => {
      target.moveTo(0, 0);
      for (let i = 0; i < 4000; i++) {
        commands++;
        target.lineTo(i % 512, Math.floor(i / 512));
      }
      target.closePath();
    };
    try {
      for (let i = 0; i < 5; i++) {
        ctx.fillStyle = resolveFill(fill, ctx, 0, 0, 512, 512, 0, undefined, undefined, outline) as CanvasPattern;
        ctx.fillRect(0, 0, 512, 512);
        expect(pixel(ctx, 256, 256)[3]).toBe(255);
        expect(pixel(ctx, 256, 256)[0]).toBeLessThan(4);
      }
      // One bounds traversal, one ramp fill/readback, one pixel write per
      // repaint. Replaying a complex outline for each contour violates this.
      expect(commands).toBe(5 * 4000);
      expect(work).toEqual({ fills: 5, readbacks: 5, writes: 5 });
    } finally { vi.unstubAllGlobals(); vi.restoreAllMocks(); }
  });

  it('wires rect fills and legacy shape-path strokes through the retained DOCX painter', () => {
    const { ctx } = canvas();
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 240, h: 120 },
      geometry: { kind: 'preset', name: 'triangle', adjustments: [] },
      fill, stroke: { color: '000000', width: 8, fill: { ...fill, path: 'shape' } },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(pixel(ctx, 120, 30)[0]).toBeGreaterThan(124);
    expect(pixel(ctx, 120, 30)[0]).toBeLessThan(133);
    expect(pixel(ctx, 120, 117)[0]).toBeGreaterThan(107);
    expect(pixel(ctx, 120, 117)[0]).toBeLessThan(112);
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
    expect(pixel(ctx, 120, 30)[0]).toBeGreaterThan(124);
    expect(pixel(ctx, 120, 30)[0]).toBeLessThan(133);
    // Shape-path strokes retain the host-box radial field in every format.
    expect(pixel(ctx, 120, 117)[0]).toBeGreaterThan(107);
    expect(pixel(ctx, 120, 117)[0]).toBeLessThan(112);
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
    expect(pixel(ctx, 120, 30)[0]).toBeGreaterThan(124);
    expect(pixel(ctx, 120, 30)[0]).toBeLessThan(133);
    // Shape-path strokes retain the host-box radial field in every format.
    expect(pixel(ctx, 120, 117)[0]).toBeGreaterThan(107);
    expect(pixel(ctx, 120, 117)[0]).toBeLessThan(112);
    expect(pixel(ctx, 20, 10)).toEqual([255, 255, 255, 255]);
  });
});
