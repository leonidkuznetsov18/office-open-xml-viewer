import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

const entry = fileURLToPath(new URL('../../packages/pptx/src/renderer.ts', import.meta.url));
const pattern = { fillType: 'pattern', preset: 'horz', fg: 'D21D54', bg: '12CED4' };
const slide = {
  index: 0, slideNumber: 1, background: null,
  elements: [
    {
      type: 'shape', x: 914400, y: 914400, width: 5486400, height: 1828800,
      rotation: 0, flipH: false, flipV: false, geometry: 'rect',
      fill: null, stroke: null,
      textBody: {
        verticalAnchor: 't',
        paragraphs: [{
          alignment: 'l', marL: 0, marR: 0, indent: 0,
          spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
          bullet: { type: 'none' }, defFontSize: null, defColor: null,
          defBold: null, defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true,
          runs: [{ type: 'text', text: 'HHHHHH', bold: true, italic: false,
            underline: false, strikethrough: false, fontSize: 72,
            color: null, patternFill: pattern, fontFamily: 'Arial' }],
        }],
        defaultFontSize: null, defaultBold: null, defaultItalic: null,
        lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
      },
    },
    {
      type: 'chart', x: 914400, y: 3200400, width: 5486400, height: 1828800,
      rotation: 0, flipH: false, flipV: false,
      chart: {
        chartType: 'clusteredBar', categories: [], series: [],
        authoredWithoutSeries: true, chartFill: pattern,
        showLegend: false, showDataLabels: false,
        catAxisHidden: true, valAxisHidden: true,
      },
    },
  ],
};

function countColors(data: Uint8ClampedArray) {
  let red = 0, cyan = 0;
  for (let i = 0; i < data.length; i += 4) {
    if (data[i + 3] < 220) continue;
    if (data[i] > 170 && data[i + 1] < 80 && data[i + 2] < 150) red++;
    if (data[i] < 80 && data[i + 1] > 150 && data[i + 2] > 150) cyan++;
  }
  return { red, cyan };
}

test('plain PPTX glyph pattern renders on main canvas and OffscreenCanvas worker', async ({ page }) => {
  test.setTimeout(90_000);
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  const script = bundle.output[0].code;
  await page.addScriptTag({ content: script });
  const main = await page.evaluate(async (s) => {
    const canvas = document.createElement('canvas');
    const renderer = (globalThis as typeof globalThis & { pptxRenderer: typeof import('../../packages/pptx/src/renderer.js') }).pptxRenderer;
    await renderer.renderSlide(canvas, s as never, 9144000, 6858000, { width: 960, dpr: 1 });
    return Array.from(canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data);
  }, slide);
  const mainColors = countColors(new Uint8ClampedArray(main));
  expect(mainColors.red).toBeGreaterThan(50);
  expect(mainColors.cyan).toBeGreaterThan(50);

  const worker = await page.evaluate(async ({ script, slide }) => {
    const source = `${script}\nself.onmessage = async (event) => { const canvas = new OffscreenCanvas(1, 1); await pptxRenderer.renderSlide(canvas, event.data, 9144000, 6858000, { width: 960, dpr: 1 }); self.postMessage(Array.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data)); };`;
    const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
    try {
      const w = new Worker(url);
      const result = await new Promise<number[]>((resolve, reject) => {
        w.onmessage = event => resolve(event.data);
        w.onerror = event => reject(new Error(event.message));
        w.postMessage(slide);
      });
      w.terminate();
      return result;
    } finally { URL.revokeObjectURL(url); }
  }, { script, slide });
  expect(countColors(new Uint8ClampedArray(worker))).toEqual(mainColors);
});
