import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { paintDrawingMLShape, type GradientFill } from '@silurus/ooxml-core';
import { paintDrawingLayout } from '../../docx/src/paint/canvas-drawing';
import type { DrawingLayout } from '../../docx/src/layout/types';
import { renderSlide } from '../../pptx/src/renderer';
import { renderViewport } from '../../xlsx/src/renderer';
import type { Slide } from '@silurus/ooxml-pptx';
import type { Styles, Worksheet } from '@silurus/ooxml-xlsx';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
const styles = { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles;

// Exact decoded-RGBA digests captured with skia-canvas from origin/main
// 71c3e209692b73173b5a8d8037182fb2e76b4d07. These are previous-renderer
// compatibility oracles, not Office fidelity claims. Each format paints its
// real shape path: plus has aspect-sensitive concave arms, gear6/gear9 cross
// the raster support boundary at tile aspect ratios, and star5 is supported.
const previous = {
  "docx": {
    "rect": [
      "f7011f9c8d4b512477b20ca2e778ed457cc49d1acd67519d71bea0a6140552ca",
      "f6dad83bec134e239e670fa6d2dd93b2c64619e21b12373596f6e043398fa372",
      "6934e98e6a97c329e643e27063ab654fb7ed889e1a0c1c4405af7759c2316ae2",
      "626890750f6d2a77ed6956131d6dd00b45e8d4425ab45ade08b21fc05a151e1b"
    ],
    "shape": [
      "7b935fb7f744457d7f7009a5245f9c1c4507d2ecf4014f7b55c83a120a59d7b5",
      "44651edfab562c12c67d532b7c0a3248c6535133dd410a9ad2b47c42f84b3e3c",
      "686744e1a47a78ebdb7b0845088285be019f8872282d7135a40057ec2a96feb3",
      "fcbe420f34ec11222c47f35faa43dbb3ceeb06245e89db95dff407e92c005b4e"
    ]
  },
  "pptx": {
    "rect": [
      "21eabc360c773a0a91bcd6d873219a2142a9a9ca30b399d87255877c0ec2dec8",
      "db8afc703e1cb212c74ea62acff2bcbaed86ad74f9001f761c3f7e3be3117c1b",
      "3d34488dfcca8022478916756bd9a03d2196efc52c4e90c7828e7c40c6533d40",
      "c01d6a8a55a314b234265bea4961af05a35088daecba9a1e1305036a2815919c"
    ],
    "shape": [
      "4824831bcfb5bf3ed18c6362982c245032a90d88e99bd42c73b0e736e52b3b40",
      "71f31a1b78315b83e3f6323e4bc8ef6e90d132eb28aaf9a99c05b8fd1fe59b73",
      "5ed8f5edb3d9de8dfd8a9c90a6c2bde0985b7219903ad60da1cf00fa5e783461",
      "f2ce902dba8f22c62710b001e3214a98eac0575cb61f4fe2d9a5b2482a70ddda"
    ]
  },
  "xlsx": {
    "rect": [
      "21eabc360c773a0a91bcd6d873219a2142a9a9ca30b399d87255877c0ec2dec8",
      "db8afc703e1cb212c74ea62acff2bcbaed86ad74f9001f761c3f7e3be3117c1b",
      "3d34488dfcca8022478916756bd9a03d2196efc52c4e90c7828e7c40c6533d40",
      "c01d6a8a55a314b234265bea4961af05a35088daecba9a1e1305036a2815919c"
    ],
    "shape": [
      "4824831bcfb5bf3ed18c6362982c245032a90d88e99bd42c73b0e736e52b3b40",
      "71f31a1b78315b83e3f6323e4bc8ef6e90d132eb28aaf9a99c05b8fd1fe59b73",
      "5ed8f5edb3d9de8dfd8a9c90a6c2bde0985b7219903ad60da1cf00fa5e783461",
      "f2ce902dba8f22c62710b001e3214a98eac0575cb61f4fe2d9a5b2482a70ddda"
    ]
  }
} as const;
const presets = ['plus', 'gear6', 'gear9', 'star5'] as const;

describe.skipIf(!skia)('tiled path gradients retain main pixels in every format', () => {
  it.each(['rect', 'shape'] as const)('%s retains main tiled connector decoration bytes', path => {
    const c = new (skia as NonNullable<typeof skia>).Canvas(400, 400);
    const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D;
    paintDrawingMLShape(ctx, {
      rect: { x: 200, y: 200, w: 50, h: 1 },
      geometry: { kind: 'preset', name: 'line', adjustments: [] }, fill: null,
      stroke: { color: '000000', width: 20, tailEnd: { type: 'triangle', w: 'lg', len: 'lg' },
        fill: { fillType: 'gradient', gradType: 'radial', path, angle: 0,
          tileRect: { r: .5 }, flip: 'xy', fillToRect: { l: .1, r: .5, t: .4, b: .1 },
          stops: [{ position: 0, color: '000000' }, { position: .5, color: '808080' },
            { position: 1, color: 'FFFFFF' }] } },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    const digest = createHash('sha256').update(ctx.getImageData(0, 0, 400, 400).data).digest('hex');
    expect(digest).toBe(path === 'rect' ? '37e09efc136fbbb8e793abc9abe794a5549141fa74a1d7fd3c9c5fde51257ade' : 'f488297a205043d3c52b674b33e6f44398497cf1f134410cbbde97378affa54c');
  });

  it.each(['docx', 'pptx', 'xlsx'] as const)('%s retains rect/shape tile bytes', async format => {
    for (const path of ['rect', 'shape'] as const) for (const [index, preset] of presets.entries()) {
      const c = new (skia as NonNullable<typeof skia>).Canvas(200, 120);
      const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D;
      const fill: GradientFill = {
        fillType: 'gradient', gradType: 'radial', path, angle: 0,
        tileRect: { r: .5 }, fillToRect: { l: .1, r: .5, t: .4, b: .1 },
        stops: [{ position: 0, color: '000000' }, { position: .5, color: '808080' },
          { position: 1, color: 'FFFFFF' }],
      };
      if (format === 'docx') {
        const bounds = { xPt: 0, yPt: 0, widthPt: 200, heightPt: 120 };
        const drawing: DrawingLayout = {
          kind: 'drawing', id: 'tiled-shape', source: { story: 'body', storyInstance: 'body', path: [0] },
          flowDomainId: 'body', flowBounds: bounds, inkBounds: bounds, advancePt: 120, ordinaryFlow: false,
          commands: [{ kind: 'drawingml-shape', plan: {
          rect: { x: 0, y: 0, w: 200, h: 120 },
          geometry: { kind: 'preset', name: preset, adjustments: [] },
          fill, stroke: null, transform: { rotationDeg: 0, flipH: false, flipV: false },
          } }],
        };
        paintDrawingLayout(drawing, { ctx, scale: 1, dpr: 1,
          resources: { paint: () => { throw new Error('shape must not paint a retained resource'); } } });
      } else if (format === 'pptx') {
        await renderSlide(c as unknown as HTMLCanvasElement, {
          index: 0, slideNumber: 1, background: null, elements: [{
            type: 'shape', x: 0, y: 0, width: 200 * 9525, height: 120 * 9525,
            rotation: 0, flipH: false, flipV: false, geometry: preset,
            fill, stroke: null, textBody: null, custGeom: null, shadow: null,
          }],
        } as Slide, 200 * 9525, 120 * 9525, { width: 200, dpr: 1 });
      } else {
        renderViewport(ctx, {
          name: 'Sheet1', isChartSheet: true, rows: [], colWidths: {}, rowHeights: {},
          freezeRows: 0, freezeCols: 0, defaultColWidth: 8.43, defaultRowHeight: 15,
          mergeCells: [], conditionalFormats: [], images: [], charts: [],
          defaultFontFamily: 'Calibri', defaultFontSize: 11,
          shapeGroups: [{ fromCol: 0, fromRow: 0, fromColOff: 0, fromRowOff: 0,
            toCol: 1, toRow: 1, toColOff: 0, toRowOff: 0, editAs: 'oneCell',
            nativeExtCx: 200 * 9525, nativeExtCy: 120 * 9525,
            shapes: [{ x: 0, y: 0, w: 1, h: 1, rot: 0, strokeWidth: 0,
              fill, geom: { type: 'preset', name: preset, adj: [] } }],
          }],
        } as Worksheet, styles, { row: 1, col: 1, rows: 1, cols: 1 });
      }
      const digest = createHash('sha256').update(ctx.getImageData(0, 0, 200, 120).data).digest('hex');
      expect(digest, `${format}/${path}/${preset}`).toBe(previous[format][path][index]);
    }
  });
});
