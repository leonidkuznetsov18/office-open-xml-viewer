import { expect, test } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';

const packageRoot = fileURLToPath(new URL('../../packages/pptx/', import.meta.url));
const fixturePath = fileURLToPath(new URL('./fixtures/pptx-run-inheritance.pptx', import.meta.url));
let server: ViteDevServer;
let url: string;

test.beforeAll(async () => {
  server = await createServer({
    root: packageRoot,
    configFile: resolve(packageRoot, 'vite.config.ts'),
    server: { host: '127.0.0.1', port: 0, strictPort: false },
  });
  await server.listen();
  const address = server.httpServer?.address();
  if (!address || typeof address === 'string') throw new Error('Vite did not assign a port');
  url = `http://127.0.0.1:${address.port}/tests/visual/run-inheritance-fixture.html`;
});
test.afterAll(async () => { await server?.close(); });

test('partial paragraph defaults reach runs and fields; explicit noFill keeps highlight', async ({ page }) => {
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-ready', 'true');
  const bytes = [...await readFile(fixturePath)];
  const results = await page.evaluate(async input => {
    type Rendered = { width: number; height: number; bytes: number[] };
    const render = (window as typeof window & {
      renderRunInheritanceFixture: (bytes: number[], index: number, mode: string) => Promise<Rendered>;
    }).renderRunInheritanceFixture;
    const count = (image: Rendered, region: [number, number, number, number]) => {
      let red = 0, yellow = 0, black = 0;
      for (let y = region[1]; y < region[3]; y++) for (let x = region[0]; x < region[2]; x++) {
        const i = 4 * (y * image.width + x);
        const [r, g, b, a] = image.bytes.slice(i, i + 4);
        if (a < 220) continue;
        if (r > 170 && g < 90 && b > 45 && b < 180) red++;
        if (r > 215 && g > 170 && b < 90) yellow++;
        if (r < 50 && g < 50 && b < 50) black++;
      }
      return { red, yellow, black };
    };
    return Promise.all(['main', 'worker'].map(async mode => {
      const image = await render(input, 0, mode);
      return { mode,
        run: count(image, [40, 30, 950, 180]),
        field: count(image, [40, 190, 950, 360]),
        noFill: count(image, [40, 370, 950, 570]),
      };
    }));
  }, bytes);
  for (const result of results) {
    expect(result.run.red, result.mode).toBeGreaterThan(100);
    expect(result.field.red, result.mode).toBeGreaterThan(100);
    expect(result.noFill.yellow, result.mode).toBeGreaterThan(100);
    expect(result.noFill.red, result.mode).toBe(0);
    expect(result.noFill.black, result.mode).toBe(0);
  }
  expect(results[0].run).toEqual(results[1].run);
  expect(results[0].field).toEqual(results[1].field);
  expect(results[0].noFill).toEqual(results[1].noFill);
});
