#!/usr/bin/env node
// The optional model-source implementation must stay outside every ordinary
// OOXML entry and worker. Resolve workspace package exports with TypeScript,
// then inspect the emitted static JS graph and decoded inline worker payloads.
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { resolve, dirname, relative, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parse } from '@babel/parser';
import { createTypeScriptResolver } from './check-core-legacy-boundary.mjs';

const resolveModule = createTypeScriptResolver();

function staticSpecifiers(ast) {
  return ast.program.body.flatMap((node) =>
    node.source && ['ImportDeclaration', 'ExportNamedDeclaration', 'ExportAllDeclaration'].includes(node.type)
      && node.importKind !== 'type' && node.exportKind !== 'type'
      && !(node.type === 'ImportDeclaration' && node.specifiers.length > 0
        && node.specifiers.every((specifier) => specifier.importKind === 'type'))
      ? [node.source.value] : []);
}

/** The real static source graph, including workspace package exports. */
export function eagerModules(entry) {
  const visited = new Set();
  const pending = [resolve(entry)];
  while (pending.length) {
    const file = pending.pop();
    if (visited.has(file)) continue;
    visited.add(file);
    if (file.includes('/src/wasm/') || !/\.[cm]?[jt]sx?$/.test(file)) continue;
    // Reparse on each audit. A test or editor can change a file between two
    // calls in one process; a path-only cache would silently preserve its old
    // import graph.
    const ast = parse(readFileSync(file, 'utf8'), { sourceType: 'module', plugins: ['typescript', 'jsx'] });
    const imports = staticSpecifiers(ast);
    const importer = relative(process.cwd(), file).replaceAll('\\', '/');
    for (const specifier of imports) {
      if (specifier.includes('?')) continue;
      let target;
      if (specifier.startsWith('.')) {
        const stem = resolve(dirname(file), specifier).replace(/\.js$/, '');
        target = [stem + '.ts', stem + '.tsx', stem + '.js', stem + '.mjs']
          .find((candidate) => existsSync(candidate));
      } else if (file.startsWith(process.cwd())) {
        target = resolveModule(importer, specifier);
      }
      if (target) pending.push(resolve(target));
    }
  }
  return visited;
}

export function assertLazySourceOwner(entry, owner) {
  const modules = eagerModules(entry);
  if (modules.has(resolve(owner))) throw new Error(`${entry} statically reaches ${owner}`);
}

function bundleGraph(entry) {
  const visited = new Set();
  const pending = [resolve(entry)];
  let bytes = 0;
  let joined = '';
  const codes = [];
  while (pending.length) {
    const file = pending.pop();
    if (visited.has(file)) continue;
    visited.add(file);
    const code = readFileSync(file, 'utf8');
    bytes += Buffer.byteLength(code);
    joined += '\n' + code;
    codes.push(code);
    const ast = parse(code, { sourceType: 'module' });
    for (const specifier of staticSpecifiers(ast)) {
      if (specifier.startsWith('.')) pending.push(resolve(dirname(file), specifier));
    }
  }
  return { bytes, files: visited.size, joined, codes };
}

function inlineWorkers(code) {
  const ast = parse(code, { sourceType: 'module' });
  const outputs = [];
  function walk(node) {
    if (!node || typeof node !== 'object') return;
    if (node.type === 'StringLiteral' && node.value.length > 10000) {
      const decoded = node.value.includes('self.') ? node.value : Buffer.from(node.value, 'base64').toString();
      if (decoded.includes('onmessage')) outputs.push(decoded);
    }
    for (const value of Object.values(node)) {
      if (Array.isArray(value)) value.forEach(walk);
      else if (value && typeof value === 'object' && value.type) walk(value);
    }
  }
  walk(ast);
  return outputs;
}

function assertNoSourceRuntime(code, name) {
  if (code.includes('ooxml-model-source-module/v1') || code.includes('model source view default')) {
    throw new Error(`${name} includes optional source runtime in its eager OOXML code`);
  }
}

// Baseline: the pre-feature OOXML production build at 776237df. The small
// entry allowance covers the modelSources presence dispatch and its Vite
// dynamic-chunk factoring; ordinary worker payloads have no allowance.
const OOXML_BUNDLE_BASELINE = Object.freeze({
  docx: { entry: 2_525_646, inline: 31_624, budget: 2_500 },
  xlsx: { entry: 1_835_670, inline: 39_902, budget: 2_500 },
  pptx: { entry: 1_824_262, inline: 59_554, budget: 2_100 },
  node: { entry: 2_575_997, budget: 3_600 },
});
const OOXML_RENDER_WORKERS = [1_416_412, 1_458_524, 2_046_946];

function assertBudget(actual, baseline, budget, label) {
  if (actual > baseline + budget) {
    throw new Error(`${label} exceeds OOXML dispatch budget: ${actual} > ${baseline} + ${budget}`);
  }
}

export function checkBuiltBundles(dist = 'dist', { packages = false } = {}) {
  for (const format of ['docx', 'xlsx', 'pptx', 'node']) {
    const graph = bundleGraph(join(dist, `${format}.mjs`));
    const baseline = OOXML_BUNDLE_BASELINE[format];
    assertBudget(graph.bytes, baseline.entry, baseline.budget, `${format} static entry`);
    assertNoSourceRuntime(graph.joined, `${format} static entry graph`);
    console.log(`${format} static JS: ${graph.bytes} bytes (${graph.bytes - baseline.entry} over main; budget ${baseline.budget}) across ${graph.files} files`);
    if (format !== 'node') {
      for (const payload of graph.codes.flatMap(inlineWorkers)) {
        assertNoSourceRuntime(payload, `${format} inline worker`);
        const bytes = Buffer.byteLength(payload);
        assertBudget(bytes, baseline.inline, 0, `${format} inline worker`);
        console.log(`${format} inline worker: ${bytes} decoded bytes`);
      }
      const sidecar = join(dist, `${format}-source-worker.mjs`);
      if (!existsSync(sidecar)) throw new Error(`Missing optional source sidecar ${sidecar}`);
      const sidecarImports = staticSpecifiers(parse(readFileSync(sidecar, 'utf8'), { sourceType: 'module' }));
      if (sidecarImports.length > 0) throw new Error(`${sidecar} is not self-contained`);
      const packageSidecar = join('packages', format, 'dist', `${format}-source-worker.mjs`);
      if (packages && !existsSync(packageSidecar)) {
        throw new Error(`Missing optional package source sidecar ${packageSidecar}`);
      }
      if (packages && existsSync(packageSidecar)
        && staticSpecifiers(parse(readFileSync(packageSidecar, 'utf8'), { sourceType: 'module' })).length > 0) {
        throw new Error(`${packageSidecar} is not self-contained`);
      }
      const packageEntry = join('packages', format, 'dist', 'index.mjs');
      if (packages) {
        if (!existsSync(packageEntry)) throw new Error(`Missing package entry ${packageEntry}`);
        const packageGraph = bundleGraph(packageEntry);
        assertNoSourceRuntime(packageGraph.joined, `${format} package static entry graph`);
        for (const payload of packageGraph.codes.flatMap(inlineWorkers)) {
          assertNoSourceRuntime(payload, `${format} package inline worker`);
        }
      }
    }
  }
  const ordinaryWorkers = [];
  for (const file of readdirSync(join(dist, 'assets')).filter((name) => /^render-worker-.*\.js$/.test(name))) {
    if (!file.startsWith('render-worker-source-')) {
      ordinaryWorkers.push(Buffer.byteLength(readFileSync(join(dist, 'assets', file))));
    }
    assertNoSourceRuntime(readFileSync(join(dist, 'assets', file), 'utf8'), file);
  }
  if (ordinaryWorkers.length !== OOXML_RENDER_WORKERS.length) {
    throw new Error(`Expected ${OOXML_RENDER_WORKERS.length} ordinary render workers, found ${ordinaryWorkers.length}`);
  }
  ordinaryWorkers.sort((a, b) => a - b);
  OOXML_RENDER_WORKERS.forEach((baseline, index) =>
    assertBudget(ordinaryWorkers[index], baseline, 0, `ordinary render worker ${index}`));
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const forbidden = [
    'packages/core/src/source/model-source.ts',
    ...['docx', 'xlsx', 'pptx'].map((format) => `packages/${format}/src/internal/worker-${format === 'docx' ? 'document' : format === 'xlsx' ? 'worksheet' : 'presentation'}-source.ts`),
    ...['docx', 'xlsx'].map((format) => `packages/${format}/src/internal/node-model-source-acquisition.ts`),
    'packages/pptx/src/internal/node-session-acquisition.ts',
    ...['docx', 'xlsx', 'pptx'].map((format) => `packages/${format}/src/internal/model-source-session.ts`),
    'packages/xlsx/src/internal/host-layout-measure.ts',
    'packages/node/src/model-source.ts',
    ...['docx', 'xlsx', 'pptx'].map((format) => `packages/node/src/${format}-model-source.ts`),
    'packages/docx/src/internal/document-model-source.ts',
    'packages/xlsx/src/internal/workbook-model-source.ts',
    'packages/pptx/src/internal/presentation-model-source.ts',
    ...['docx', 'xlsx', 'pptx'].flatMap((format) => [
      `packages/${format}/src/worker-source.ts`,
      `packages/${format}/src/render-worker-source.ts`,
    ]),
  ];
  for (const entry of [
    'packages/core/src/index.ts',
    ...['docx', 'xlsx', 'pptx'].flatMap((format) => [
      `packages/${format}/src/index.ts`,
      `packages/${format}/src/worker.ts`,
      `packages/${format}/src/render-worker.ts`,
    ]),
    'packages/node/src/index.ts',
  ]) {
    const graph = eagerModules(entry);
    for (const module of forbidden) {
      if (graph.has(resolve(module))) throw new Error(`${entry} statically reaches ${module}`);
    }
  }
  if (existsSync('dist/docx.mjs')) checkBuiltBundles('dist', { packages: process.argv.includes('--packages') });
  console.log('OOXML entries and workers keep model-source runtime behind dynamic loads.');
}
