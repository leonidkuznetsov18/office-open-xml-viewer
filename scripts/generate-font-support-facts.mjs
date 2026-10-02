/** Reference generation uses exactly the production bounded parser. Run once
 * per generator, then feed path/collection index requests as NDJSON on stdin.
 * Bundling stays in an owned OS temporary directory and is removed on exit. */
import { createRequire } from 'node:module';
const { build } = createRequire(new URL('../packages/core/package.json', import.meta.url))('esbuild');
import { mkdtemp, rm, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { createInterface } from 'node:readline';
const own = await mkdtemp(join(tmpdir(), 'ooxml-support-generator-'));
try {
  const output = join(own, 'facts.mjs');
  await build({ stdin: { contents: `export { parseOpenTypeResourceMetrics } from './packages/core/src/fonts/open-type-metrics.ts'; export { fontSupportFacts } from './packages/core/src/internal/font-support-facts.ts';`, resolveDir: process.cwd() }, bundle: true, format: 'esm', platform: 'node', outfile: output });
  const { parseOpenTypeResourceMetrics, fontSupportFacts } = await import(pathToFileURL(output));
  const input = createInterface({ input: process.stdin });
  let previousPath, previousBytes;
  for await (const line of input) {
    const { path, faceIndex } = JSON.parse(line);
    if (path !== previousPath) { previousBytes = new Uint8Array(await readFile(path)); previousPath = path; }
    const metric = parseOpenTypeResourceMetrics(previousBytes, faceIndex);
    process.stdout.write(JSON.stringify({ supportFacts: fontSupportFacts(metric) ?? null, unicodeRanges: metric?.unicodeRanges ?? null, unicodePossibleRanges: metric?.unicodePossibleRanges ?? null }) + '\n');
  }
} finally { await rm(own, { recursive: true, force: true }); }
