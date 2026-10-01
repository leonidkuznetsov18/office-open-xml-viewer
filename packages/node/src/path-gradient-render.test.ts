import { describe, expect, it, vi } from 'vitest';
import { paintDrawingMLShape, resolveFill, type GradientFill } from '@silurus/ooxml-core';
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

  it('fills the complete asymmetric focus area and interpolates alpha once', () => {
    const { ctx } = canvas();
    paint({ ...fill, fillToRect: { l: .2, r: .4, t: .25, b: .25 },
      stops: [{ position: 0, color: 'FF000080' }, { position: 1, color: '0000FF00' }],
    }, ctx, 240, 120);
    expect(pixel(ctx, 75, 40)).toEqual([255, 0, 0, 128]);
    expect(pixel(ctx, 130, 80)).toEqual([255, 0, 0, 128]);
    expect(pixel(ctx, 192, 60)[3]).toBeGreaterThan(60);
    expect(pixel(ctx, 192, 60)[3]).toBeLessThan(67);
  });

  it('contracts a triangle outline rather than shading its bounding circle', () => {
    const { ctx } = canvas();
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 240, h: 120 },
      geometry: { kind: 'preset', name: 'triangle', adjustments: [] },
      fill: { ...fill, path: 'shape' }, stroke: null,
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    // Both points lie on the triangle contracted halfway toward (120,60).
    for (const [x, y] of [[120, 30], [75, 75]]) {
      const u = (x + .5 - 120) / 120;
      const v = (y + .5 - 60) / 60;
      const expected = 255 * Math.max(v, 2 * Math.abs(u) - v);
      // One contour quantization step plus antialiasing at the sloping edge.
      expect(Math.abs(pixel(ctx, x, y)[0] - expected)).toBeLessThanOrEqual(3);
    }
    expect(pixel(ctx, 20, 10)[3]).toBe(0);
  });

  it('keeps concave custom boundaries at the outer stop and leaves the notch empty', () => {
    const { ctx } = canvas(120, 120);
    const vertices = [[0, 0], [1, 0], [1, 1], [.65, 1], [.65, .4], [.35, .4], [.35, 1], [0, 1]];
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 120, h: 120 },
      geometry: { kind: 'custom', subpaths: [[
        ...vertices.map(([x, y], index) => ({ cmd: index ? 'lineTo' as const : 'moveTo' as const, x, y })),
        { cmd: 'close' },
      ]] },
      fill: { ...fill, path: 'shape', fillToRect: { l: .5, r: .5, t: .2, b: .8 } },
      stroke: null, transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(pixel(ctx, 60, 80)[3]).toBe(0);
    expect(pixel(ctx, 41, 80)[0]).toBeGreaterThan(245);
    expect(pixel(ctx, 60, 24)[0]).toBeLessThan(12);
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

  it('paints custom geometry outside the authored transform box without unbounded allocation', () => {
    const { ctx } = canvas();
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 120, h: 120 },
      geometry: { kind: 'custom', subpaths: [[
        { cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 2, y: 1 },
        { cmd: 'lineTo', x: 0, y: 1 }, { cmd: 'close' },
      ]] },
      fill: { ...fill, path: 'shape' }, stroke: null,
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(pixel(ctx, 180, 105)[3]).toBe(255);
    expect(pixel(ctx, 180, 105)[0]).toBeGreaterThan(150);
    expect(pixel(ctx, 180, 105)[0]).toBeLessThan(254);
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

  it('wires PPTX shape outlines through the complete slide painter', async () => {
    const { c, ctx } = canvas();
    const px = (n: number) => n * 9525;
    const shape = {
      type: 'shape', x: 0, y: 0, width: px(240), height: px(120),
      rotation: 0, flipH: false, flipV: false, geometry: 'triangle',
      fill: { ...fill, path: 'shape' }, stroke: null, textBody: null,
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
        shapes: [{ x: 0, y: 0, w: 1, h: 1, rot: 0, strokeWidth: 0,
          fill: { ...fill, path: 'shape' }, geom: { type: 'custom', paths: [
            { w: 240, h: 120, commands: triangle, stroke: false },
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
    expect(pixel(ctx, 20, 10)).toEqual([255, 255, 255, 255]);
  });
});
