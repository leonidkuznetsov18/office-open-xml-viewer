import { createHash } from 'node:crypto';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { paintDrawingMLShape, resolveFill, trackPaintPath, withInheritedPatternScope,
  withPatternCoordinateSpace, type Stroke, type DrawingMLShapeGeometry } from '@silurus/ooxml-core';
import { withPatternPointScale } from '../../core/src/shape/paint';
import { renderSlide } from '../../pptx/src/renderer';
import { renderViewport } from '../../xlsx/src/renderer';
import type { Slide } from '@silurus/ooxml-pptx';
import type { Styles, Worksheet } from '@silurus/ooxml-xlsx';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
const stops = [{ position: 0, color: '112233' }, { position: 1, color: 'DDEEFF' }];
type DecorationFill = NonNullable<Stroke['fill']>;
const fills: Record<string, DecorationFill> = {
  pattern: { fillType: 'pattern', preset: 'pct50', fg: 'FF0000', bg: '0000FF' },
  solid: { fillType: 'solid', color: '5478AB' },
  linear: { fillType: 'gradient', gradType: 'linear', angle: 25, stops },
  circle: { fillType: 'gradient', gradType: 'radial', path: 'circle', angle: 0,
    fillToRect: { l: .1, r: .5, t: .4, b: .1 }, stops },
  tiledLinear: { fillType: 'gradient', gradType: 'linear', angle: 25, stops, tileRect: { r: .5 } },
  tiledCircle: { fillType: 'gradient', gradType: 'radial', path: 'circle', angle: 0,
    fillToRect: { l: .1, r: .5, t: .4, b: .1 }, stops, tileRect: { r: .5 } },
};

// Decoded-RGBA SHA-256 values from the production painters at origin/main
// 71c3e209692b73173b5a8d8037182fb2e76b4d07, with skia-canvas 2.0.2.
// These protect inherited paint frames, not Office fidelity. Connector tests
// include a filled triangle; custom paths include an open stroked arrow;
// rightArrow exercises ordinary shape fill/stroke helpers. XLSX has no line
// decorations, so its distinct integration covers only the rightArrow body.
const previous: Record<string, readonly string[]> = {
  "core-connector": [
    "569d78f3df89405a7c932bca6670903cd519945ff9d9ca0fdd96a5fb928fe491",
    "9e26c769606fbd41e7a65dad19ca6dc78aa8ec979947bbbdc7973918928e2f4f",
    "f640a5ccdd94dd5338076a2f89d8745f64a3dc5e173aaa47ac2e3b75900324aa",
    "8631ce89376329f116f2de39f944f410af51773d84af9e99683d8b1ff3d84296",
    "4255164e2bad6d90786f1192f6aa3c0394a462d7a11e0a5caf5ee9676e22b9dc",
    "7ccf37f47bfc293602c6cf5aca87493917064bb487d1383992a6883fe8799358"
  ],
  "core-custom": [
    "5d934b5afcd25a063e3c3f27d40097c6f6333f8ef3a3d524561a8386c95ba654",
    "1e413f9936ff9086631d9afee9a9e72733359b3ffa948284f2e18fe8673181f2",
    "b511c1bc22643ba855b95f9e216a59b898bc8b8c7bf91aba68dc8537524da317",
    "caef577f909b9b98dde7bcf882c27ef78b9cf6ddf98fd2208d340c9c709ba711",
    "1a5e123fca26a629c4cf3858c5ba9109ed1f70f6fb7f8433ecfe0f0828bf6427",
    "04d921f45de70a311592d95bcd18880d63ee0042ebc6075e8f244f4ff43153b5"
  ],
  "core-body": [
    "d77a8b1d7544dd6b147877e66b2c1c226f96dfa5a081dfee7ad25a2debe4af3c",
    "76a250caecc9f928fa54a89c8f344ecf345707f9ba7a37ce5c89fac2d9e6f787",
    "d3cd20a9b2af629be98ee9b874e42a67b3dd5f1bbc88461c19528b97f54ac16d",
    "8ce754b60fa0a029ea12e2be2c7d6c86b284724083faa1b3f3f8a72a3eb7668c",
    "e7950127e15fb53efa677600c3c2d8c5833f1b40d7cdf52470cdacd58cfbe003",
    "72e34e30b2897acad5cd5891f6479d016c65fecac11225a36bc2f67a7b9119dd"
  ],
  "pptx-connector": [
    "5129ac4f59e9dc4be7ed1ee189631d0f8d8bdf8432fa4672ee7408c4f6d5bcc2",
    "9cd11325826859a868646d739822d4605512e5c032b45383bf795146f7926fdb",
    "ba65e7575f61e428ff8fb56d7434652d3e853eeeca01e6db04bdfa2b4897367d",
    "4fee3ed47b4e4f33f37398fbee9b5a54e8aa39d3e8ebfe2bd7d5486581d667d4",
    "e863266cf04eb8ee08cce91e65376471ee149f87ed8bef4089c3946281b0071a",
    "905513af780cd882a464329fd9406832bb56aa16cd9990590ac14b98252a659a"
  ],
  "pptx-custom": [
    "0affc52b2d60429ae25c064bfd557beb20c64e9eab877dca4d33d7d26e65d564",
    "94a44684e5db3c6f681aff183dcd0e10dd2415dfabac8c1dbaf44951efe2257d",
    "406a03a24cca92ec887cf6045989e9d63f55c9e8e692295b4a6b5601984c6afa",
    "9de853bdeb20e5c94c693a4ae53208894901991e36562159167dc2920333e134",
    "b5ac821fe349337879398912792e9a0b7dcddcc8b6c5e98eda6d89b68e2b5aff",
    "e3b46b7225abf651b5215048f8261299b5e47bbdab75eafdbf8aa2671bf9b696"
  ],
  "xlsx-body": [
    "c2f7278ffe4802176056354cd1344f9fac7c5e5c092cbdd648b4932d84838ff9",
    "5dcd880c90c085ac8aa15ca16e6e8dc454ee67b18b2bfac62c2e938cf420a6bd",
    "a54e0948576f56493c30a36b85e74cae0cee293bb9ba3013bd95cd7e8bd6a43c",
    "00ecf875557da5c7cef576e3e45b91e60ed4f82f76c3b099f883b86fb64bf590",
    "7f5174b330873bfd5156396430382ea898ba93929a2effdbc839e9a0ef92cab0",
    "ba09b89043fe8e94de74b4ae66852c08f51c07c7ef12c8b2e0bf85d4749673a7"
  ]
};
const routes = ['core-connector', 'core-custom', 'core-body', 'pptx-connector', 'pptx-custom', 'xlsx-body'];

async function renderCase(route: string, fill: DecorationFill): Promise<string> {
  const c = new (skia as NonNullable<typeof skia>).Canvas(400, 400);
  const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D;
  const body = route.endsWith('body');
  const custom = route.endsWith('custom');
  const rect = body ? { x: 130, y: 140, w: 150, h: 100 } : { x: 200, y: 200, w: 50, h: 1 };
  const geometry: DrawingMLShapeGeometry = custom ? { kind: 'custom', subpaths: [[
    { cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 1, y: 1 },
  ]] } : { kind: 'preset', name: body ? 'rightArrow' : 'line', adjustments: [] };
  const stroke = { color: '000000', width: 20, fill,
    ...(body ? {} : { tailEnd: { type: custom ? 'arrow' as const : 'triangle' as const, w: 'lg' as const, len: 'lg' as const } }),
  };
  if (route.startsWith('core')) {
    // Retained DOCX drawing commands enter this same shared production painter.
    paintDrawingMLShape(ctx, { rect, geometry, fill: body ? fill : null, stroke,
      transform: { rotationDeg: 30, flipH: false, flipV: false } }, 1);
  } else if (route.startsWith('pptx')) {
    await renderSlide(c as unknown as HTMLCanvasElement, {
      index: 0, slideNumber: 1, background: null, elements: [{ type: 'shape',
        x: rect.x * 9525, y: rect.y * 9525, width: rect.w * 9525, height: rect.h * 9525,
        rotation: 30, flipH: false, flipV: false, geometry: 'line', fill: null,
        stroke: { ...stroke, width: stroke.width * 9525 }, textBody: null,
        custGeom: geometry.kind === 'custom' ? geometry.subpaths : null, shadow: null,
      }],
    } as Slide, 400 * 9525, 400 * 9525, { width: 400, dpr: 1 });
  } else {
    renderViewport(ctx, {
      name: 'Sheet1', isChartSheet: true, rows: [], colWidths: {}, rowHeights: {},
      freezeRows: 0, freezeCols: 0, defaultColWidth: 8.43, defaultRowHeight: 15,
      mergeCells: [], conditionalFormats: [], images: [], charts: [],
      defaultFontFamily: 'Calibri', defaultFontSize: 11,
      shapeGroups: [{ fromCol: 0, fromRow: 0, fromColOff: rect.x * 9525, fromRowOff: rect.y * 9525,
        toCol: 1, toRow: 1, toColOff: 0, toRowOff: 0, editAs: 'oneCell',
        nativeExtCx: rect.w * 9525, nativeExtCy: rect.h * 9525,
        shapes: [{ x: 0, y: 0, w: 1, h: 1, rot: 30, strokeWidth: 20,
          strokeColor: '000000', strokeFill: fill, fill,
          geom: { type: 'preset', name: 'rightArrow', adj: [] } }],
      }],
    } as Worksheet, { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles,
    { row: 1, col: 1, rows: 1, cols: 1 });
  }
  return createHash('sha256').update(ctx.getImageData(0, 0, 400, 400).data).digest('hex');
}

describe.skipIf(!skia)('non-path-gradient paints retain main frames on rotated decorations and shapes', () => {
  beforeAll(() => vi.stubGlobal('OffscreenCanvas', (skia as NonNullable<typeof skia>).Canvas));
  afterAll(() => vi.unstubAllGlobals());
  it.each(routes)('%s retains pattern/solid/linear/circle and tiled native pixels', async route => {
    for (const [index, [name, fill]] of Object.entries(fills).entries()) {
      expect(await renderCase(route, fill), `${route}/${name}`).toBe(previous[route][index]);
    }
  });

  it('preserves ordinary pattern scope through path tracking and effect-canvas inheritance', () => {
    const { Canvas } = skia as NonNullable<typeof skia>;
    const source = new Canvas(100, 100).getContext('2d') as unknown as CanvasRenderingContext2D;
    const paint = (tracked: boolean) => {
      const target = new Canvas(100, 100).getContext('2d') as unknown as CanvasRenderingContext2D;
      target.translate(30, 20);
      target.rotate(Math.PI / 6);
      withPatternPointScale(source, 2, () => withPatternCoordinateSpace(source,
        { a: 1.2, b: .2, c: -.3, d: .8, e: 4, f: 7 }, () => {
          withInheritedPatternScope(tracked ? trackPaintPath(source) : source, target, () => {
            target.fillStyle = resolveFill(fills.pattern, target, 0, 0, 50, 50) as CanvasPattern;
            target.fillRect(0, 0, 50, 50);
          }, { x: 3, y: 5 });
        }));
      return target.getImageData(0, 0, 100, 100).data;
    };
    expect(paint(true)).toEqual(paint(false));
  });
});
