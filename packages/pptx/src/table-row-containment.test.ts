import { describe, expect, it } from 'vitest';
import { renderTable } from './renderer';
import type { TableCell, TableElement, TextBody } from './types';

const EMU = 12700;
const previousRendererPath = (globalThis as typeof globalThis & {
  process?: { env: Record<string, string | undefined> };
}).process?.env.PPTX_PREVIOUS_RENDERER;
type Matrix = [number, number, number, number, number, number];

// Record glyph rectangles in the actual Canvas coordinate system, independent
// of selection boxes (which include terminal leading and are not ink bounds).
// The synthetic face has a one-em 10/2 ascent/descent box at 12pt,
// full-em Han advances and 8pt Latin advances. Metrics follow the font size
// and Canvas baseline, including the middle baseline used for upright eaVert.
// Both Latin and East Asian slots explicitly use this face; catalogue face
// metrics would describe a different font from this synthetic Canvas.
function recorder() {
  let matrix: Matrix = [1, 0, 0, 1, 0, 0];
  const saved: Matrix[] = [];
  const glyphs: { top: number; bottom: number; left: number; right: number; frame: number }[] = [];
  const frames: { x: number; y: number; height: number; width: number }[] = [];
  const ctx = new Proxy({
    canvas: { width: 2000, height: 2000 }, font: '', fillStyle: '',
    textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
    measureText(text: string) {
      const size = Number(/([\d.]+)px/.exec(this.font)?.[1] ?? 12) / 12;
      return { width: Array.from(text).reduce((advance, ch) => advance + (/\p{Script=Han}/u.test(ch) ? 12 : 8), 0) * size, actualBoundingBoxAscent: (this.textBaseline === 'middle' ? 6 : 10) * size,
        actualBoundingBoxDescent: (this.textBaseline === 'middle' ? 6 : 2) * size, fontBoundingBoxAscent: 10 * size, fontBoundingBoxDescent: 2 * size } as TextMetrics;
    },
    save() { saved.push([...matrix]); },
    restore() { matrix = saved.pop() as Matrix; },
    translate(x: number, y: number) {
      matrix[4] += matrix[0] * x + matrix[2] * y;
      matrix[5] += matrix[1] * x + matrix[3] * y;
    },
    rotate(a: number) {
      const [x, y, u, v, tx, ty] = matrix;
      matrix = [x * Math.cos(a) + u * Math.sin(a), y * Math.cos(a) + v * Math.sin(a),
        u * Math.cos(a) - x * Math.sin(a), v * Math.cos(a) - y * Math.sin(a), tx, ty];
    },
    fillRect(x: number, y: number, width: number, height: number) { frames.push({ x, y, width, height }); },
    fillText(text: string, x: number, y: number) {
      const size = Number(/([\d.]+)px/.exec(this.font)?.[1] ?? 12) / 12;
      const width = Array.from(text).reduce((advance, ch) => advance + (/\p{Script=Han}/u.test(ch) ? 12 : 8), 0) * size;
      if (this.textAlign === 'center') x -= width / 2;
      if (this.textBaseline === 'middle') y += 4 * size;
      const corners = [[x, y - 10 * size], [x + width, y - 10 * size], [x, y + 2 * size], [x + width, y + 2 * size]];
      const xs = corners.map(([a, b]) => matrix[0] * a + matrix[2] * b + matrix[4]);
      const ys = corners.map(([a, b]) => matrix[1] * a + matrix[3] * b + matrix[5]);
      glyphs.push({ top: Math.min(...ys), bottom: Math.max(...ys), left: Math.min(...xs), right: Math.max(...xs), frame: frames.length - 1 });
    },
    getLineDash() { return []; },
  }, { get(target, key) { return key in target ? Reflect.get(target, key) : () => {}; } });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, frames, glyphs };
}

function table(body: TextBody, height: number): TableElement {
  const cell = { textBody: body, fill: { fillType: 'solid', color: 'FFFFFF' },
    borderL: null, borderR: null, borderT: null, borderB: null,
    gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false } as TableCell;
  return { type: 'table', x: 0, y: 0, width: 2000 * EMU, height: height * EMU,
    rotation: 0, flipH: false, flipV: false, cols: [2000 * EMU],
    rows: [{ height: height * EMU, cells: [cell] }] };
}

function body(text: string, vert: string, numCol = 1, rtlCol = false): TextBody {
  return { verticalAnchor: 't', lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    vert, numCol, rtlCol, wrap: 'none', autoFit: 'none',
    paragraphs: [{ alignment: 'l', marL: 0, marR: 0, indent: 0,
      spaceBefore: null, spaceAfter: null, spaceLine: null, bullet: { type: 'none' },
      runs: [{ type: 'text', text, fontSize: 12, fontFamily: 'UnknownTestFace', fontFamilyEa: 'UnknownTestFace' }] }],
  } as unknown as TextBody;
}

function paint(t: TableElement, renderer = renderTable) {
  const rec = recorder();
  renderer(rec.ctx, t, 1 / EMU);
  return { frames: rec.frames, glyphs: rec.glyphs };
}

function contain(result: ReturnType<typeof paint>, label?: unknown) {
  for (const glyph of result.glyphs) {
    const frame = result.frames[glyph.frame];
    // Only arithmetic trigonometry noise is tolerated, not a pixel fit budget.
    expect(glyph.top, JSON.stringify(label)).toBeGreaterThanOrEqual(frame.y - 1e-9);
    expect(glyph.bottom, JSON.stringify(label)).toBeLessThanOrEqual(frame.y + frame.height + 1e-9);
    expect(glyph.left, JSON.stringify(label)).toBeGreaterThanOrEqual(frame.x - 1e-9);
    expect(glyph.right, JSON.stringify(label)).toBeLessThanOrEqual(frame.x + frame.width + 1e-9);
  }
}

describe('table rows contain paint on their final text frame', () => {
  it('contains the second RTL text column in a zero-minimum rotated row', () => {
    const result = paint(table(body('AAAAA', 'vert', 2, true), 0));
    contain(result);
    expect(result.frames[0].height).toBeCloseTo(80, 3);
  });

  it('contains the tab field and its trailing margin after repeated row growth', () => {
    const text = body('A\tA', 'vert');
    text.paragraphs[0].marL = 2 * EMU;
    text.paragraphs[0].marR = 3 * EMU;
    text.paragraphs[0].tabStops = [{ pos: 40 * EMU, algn: 'l' }];
    const result = paint(table(text, 30));
    contain(result);
    expect(Math.max(...result.glyphs.map(g => g.bottom)) + 3).toBeLessThanOrEqual(result.frames[0].height);
  });

  it('retains physical vertical insets in an empty stacked body', () => {
    const text = body('', 'wordArtVert');
    text.tIns = 2 * EMU; text.bIns = 3 * EMU;
    expect(paint(table(text, 0)).frames[0].height).toBe(5);
  });

  it('rechecks earlier cells when a neighbouring merge grows their row', () => {
    const input = table(body('A\tA', 'vert', 2, true), 30);
    input.cols = [2000 * EMU, 2000 * EMU]; input.width = 4000 * EMU;
    const neighbour = { ...input.rows[0].cells[0], textBody: body('AAAAA', 'vert', 4, true), rowSpan: 2 };
    input.rows[0].cells.push(neighbour);
    input.rows.push({ height: 30 * EMU, cells: [
      { ...input.rows[0].cells[0], textBody: body('AB', 'vert270', 2, true) },
      { ...neighbour, textBody: null, vMerge: true },
    ] });
    (input.rows[0].cells[0].textBody as TextBody).paragraphs[0].tabStops = [{ pos: 40 * EMU, algn: 'l' }];
    contain(paint(input));
  });

  it('resolves single-row minima before distributing a merged-cell deficit', () => {
    const input = table(body('A', 'horz'), 20);
    input.cols = [2000 * EMU, 2000 * EMU]; input.width = 4000 * EMU;
    const tall = body('A', 'horz');
    tall.paragraphs = Array.from({ length: 5 }, () => ({ ...tall.paragraphs[0] }));
    const merged = { ...input.rows[0].cells[0], textBody: tall, rowSpan: 2 };
    input.rows[0].cells.push(merged);
    input.rows.push({ height: 20 * EMU, cells: [
      { ...input.rows[0].cells[0], textBody: tall },
      { ...merged, textBody: null, vMerge: true },
    ] });
    const result = paint(input);
    expect(result.frames[0].height).toBe(20);
    expect(result.frames[2].height).toBeCloseTo(72, 8);
    contain(result);
  });

  it('contains seeded random text bodies across direction, columns, tabs, insets and autofit', async () => {
    const previous = previousRendererPath
      ? await import(/* @vite-ignore */ previousRendererPath) : undefined;
    let seed = 1657;
    const next = (n: number) => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return Math.floor(seed / 0x100000000 * n); };
    for (const vert of ['vert', 'vert270', 'eaVert', 'mongolianVert', 'wordArtVert', 'wordArtVertRtl', 'horz']) {
      for (const numCol of [1, 2, 4, 16]) for (const rtlCol of [false, true]) {
        for (const autoFit of ['none', 'norm', 'sp']) for (let sample = 0; sample < 4; sample++) {
          const text = body(Array.from({ length: 1 + next(8) }, () => ['A', 'B', '漢', ' '][next(4)]).join(''), vert, numCol, rtlCol);
          text.autoFit = autoFit;
          text.fontScale = autoFit === 'norm' ? 0.5 : undefined;
          text.wrap = sample % 2 ? 'square' : 'none';
          text.lIns = next(4) * EMU; text.rIns = next(4) * EMU;
          text.tIns = next(4) * EMU; text.bIns = next(4) * EMU;
          text.spcCol = next(3) * EMU;
          text.verticalAnchor = ['t', 'ctr', 'b'][next(3)];
          if (sample >= 2) {
            text.paragraphs[0].runs[0] = { ...text.paragraphs[0].runs[0], type: 'text', text: 'A\tA' } as TextBody['paragraphs'][number]['runs'][number];
            text.paragraphs[0].tabStops = [{ pos: (15 + next(30)) * EMU, algn: ['l', 'ctr', 'r'][next(3)] }];
          }
          text.paragraphs[0].marL = next(4) * EMU;
          text.paragraphs[0].marR = next(4) * EMU;
          // A wide physical frame isolates row-axis defects; unwrapped
          // horizontal text is allowed to overflow a narrow width by OOXML.
          if (sample === 1) text.paragraphs.push({ ...text.paragraphs[0], runs: [{ ...text.paragraphs[0].runs[0], type: 'text', text: 'AB' } as TextBody['paragraphs'][number]['runs'][number]] });
          try {
            contain(paint(table(text, sample === 0 ? 0 : 30)), { text, sample });
            // With an already sufficient frame and no parser/style change,
            // the complete randomized body must retain origin/main's paint.
            if (previous) {
              const sufficient = table(text, 4000);
              expect(paint(sufficient)).toEqual(paint(sufficient, previous.renderTable));
            }
          } catch (error) {
            throw new Error(JSON.stringify({ text, sample }), { cause: error });
          }
        }
      }
    }
  });

  it.skipIf(!previousRendererPath)('preserves previous-renderer paint when style and row growth are uninvolved', async () => {
    // The gate supplies an origin/main bundle from a clean detached checkout;
    // no copied renderer or private fixture is retained in the repository.
    const previous = await import(/* @vite-ignore */ previousRendererPath as string);
    for (const vert of ['vert', 'vert270', 'eaVert', 'mongolianVert', 'wordArtVert', 'wordArtVertRtl', 'horz']) {
      for (const numCol of [1, 2, 4]) for (const rtlCol of [false, true]) for (const tabs of [false, true]) {
        const text = body(tabs ? 'A\tA' : 'A A漢', vert, numCol, rtlCol);
        text.paragraphs[0].tabStops = [{ pos: 40 * EMU, algn: 'l' }];
        const input = table(text, 4000);
        expect(paint(input)).toEqual(paint(input, previous.renderTable));
      }
    }
  });
});
