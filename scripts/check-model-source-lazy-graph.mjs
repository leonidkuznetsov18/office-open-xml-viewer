#!/usr/bin/env node
// The optional model-source implementation must stay outside every ordinary
// OOXML entry and worker. Resolve workspace package exports with TypeScript,
// then inspect the emitted static JS graph and decoded inline worker payloads.
// This guards accidental eager coupling by maintainers or agents; it is not a
// security boundary against deliberately adversarial JavaScript.
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { resolve, dirname, relative, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parse } from '@babel/parser';
import { execFileSync } from 'node:child_process';
import { createTypeScriptResolver } from './check-core-legacy-boundary.mjs';

const resolveModule = createTypeScriptResolver();

// The only entry-to-source transitions are these opt-in dispatch functions.
// Keep this list exact: a new caller must explain why it can run only after
// modelSources (or a worker source descriptor) is present.
export const SOURCE_DISPATCH = new Map([
  ['packages/core/src/source/model-source.ts', ['openModelSourceModule']],
  ['packages/docx/src/document.ts', ['load']],
  ['packages/xlsx/src/workbook.ts', ['load']],
  ['packages/pptx/src/presentation.ts', ['load']],
  ['packages/node/src/docx.ts', ['openDocxDocument', 'materializeDocxDocument']],
  ['packages/node/src/xlsx.ts', ['openXlsxWorkbook']],
  ['packages/node/src/pptx.ts', ['openPptxPresentationImpl']],
  ['packages/node/src/docx-model-source.ts', ['acquireDocxInput']],
  ['packages/node/src/xlsx-model-source.ts', ['acquireXlsxInput']],
  ['packages/node/src/pptx-model-source.ts', ['acquirePptxInput']],
  ...['docx', 'xlsx', 'pptx'].flatMap((format) => [
    [`packages/${format}/src/worker-source.ts`, ['self.onmessage']],
    [`packages/${format}/src/render-worker-source.ts`,
      format === 'pptx' ? ['executeArchiveFromNew'] : ['self.onmessage']],
  ]),
]);

export function checkSourceDispatchImports(files) {
  for (const { path, text } of files) {
    if (/\.(?:test|spec|stories|probe)\.[cm]?[jt]sx?$/.test(path)) continue;
    const ast = parse(text, { sourceType: 'module', plugins: ['typescript', 'jsx'] });
    const allowed = SOURCE_DISPATCH.get(path) ?? [];
    function walk(node, functionName) {
      if (!node || typeof node !== 'object' || !node.type) return;
      let current = functionName;
      if (node.type === 'FunctionDeclaration') current = node.id?.name;
      if (['ClassMethod', 'ObjectMethod'].includes(node.type)) current = node.key?.name;
      if (['ArrowFunctionExpression', 'FunctionExpression'].includes(node.type)) {
        // Preserve the containing named dispatch for its local callback/IIFE.
        current = functionName;
      }
      if ((node.type === 'CallExpression' && node.callee.type === 'Import')
        || node.type === 'ImportExpression') {
        const target = node.arguments?.[0] ?? node.source;
        const specifier = target?.value;
        const sourceImport = typeof specifier === 'string'
          ? /model-source/.test(specifier)
          : text.slice(target?.start ?? 0, target?.end ?? 0).includes('sourceOwnerUrl')
            || text.slice(target?.start ?? 0, target?.end ?? 0).includes('sourceModule.moduleUrl');
        if (sourceImport && !allowed.includes(current)) {
          throw new Error(`${path}:${node.loc?.start.line} model-source import outside an allowed dispatch function (${current ?? 'top level'})`);
        }
      }
      // Assignment to self.onmessage is a named worker dispatch boundary.
      if (node.type === 'AssignmentExpression' && node.left.type === 'MemberExpression'
        && node.left.object.name === 'self' && node.left.property.name === 'onmessage') {
        walk(node.right, 'self.onmessage');
        return;
      }
      for (const value of Object.values(node)) {
        if (Array.isArray(value)) value.forEach((child) => walk(child, current));
        else if (value && typeof value === 'object' && value.type) walk(value, current);
      }
    }
    walk(ast, undefined);
  }
}

function trackedSourceFiles() {
  return execFileSync('git', ['ls-files', 'packages/core/src', 'packages/docx/src',
    'packages/xlsx/src', 'packages/pptx/src', 'packages/node/src'], { encoding: 'utf8' })
    .split('\n').filter((path) => /\.[cm]?[jt]sx?$/.test(path) && existsSync(path))
    .map((path) => ({ path, text: readFileSync(path, 'utf8') }));
}

function staticSpecifiers(ast) {
  return ast.program.body.flatMap((node) =>
    node.source && ['ImportDeclaration', 'ExportNamedDeclaration', 'ExportAllDeclaration'].includes(node.type)
      && node.importKind !== 'type' && node.exportKind !== 'type'
      && !(node.type === 'ImportDeclaration' && node.specifiers.length > 0
        && node.specifiers.every((specifier) => specifier.importKind === 'type'))
      ? [node.source.value] : []);
}

export function assertNoTopLevelModelImport(code, name) {
  const ast = parse(code, { sourceType: 'module' });
  function walk(node, functionDepth) {
    if (!node || typeof node !== 'object' || !node.type) return;
    const target = node.type === 'ImportExpression' ? node.source
      : node.type === 'CallExpression' && node.callee.type === 'Import' ? node.arguments[0]
        : undefined;
    if (functionDepth === 0 && typeof target?.value === 'string'
      && target.value.includes('model-source')) {
      throw new Error(`${name} imports a model-source chunk at top level`);
    }
    const nested = functionDepth + (['FunctionDeclaration', 'FunctionExpression',
      'ArrowFunctionExpression', 'ClassMethod', 'ObjectMethod'].includes(node.type) ? 1 : 0);
    for (const value of Object.values(node)) {
      if (Array.isArray(value)) value.forEach((child) => walk(child, nested));
      else if (value && typeof value === 'object' && value.type) walk(value, nested);
    }
  }
  walk(ast, 0);
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
    assertNoTopLevelModelImport(code, file);
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
// Rebased DOCX and Node entries against the clean 470743cb build for #1557
// (merged by PR #1586):
// sliced layout, stepwise finalization, and viewer load ownership add 4,728
// DOCX bytes; shared layout validation/freezing adds 398 Node bytes. The
// optional model-source runtime remains outside both eager entry graphs.
// The ordinary DOCX render worker grows by 289 bytes from those same layout
// changes. PR #1590's projection consolidation then brings the measured main
// graph at aec306b6 to 2,533,239 DOCX, 2,579,716 Node, and 2,047,181 worker
// bytes (+200, -20, and -54 respectively versus the #1557 guard values).
// Issue #1591 then makes validation paths lazy, brands only frozen plain-data
// roots, shares verified occurrence-independent data, and bounds text caches.
// Against clean aec306b6, that grew the DOCX worker from 2,047,181 to
// 2,047,720 bytes (+539) while keeping its zero-allowance guard.
// The combined 546160b1 + #1591 production build measures 2,563,637 DOCX,
// 1,845,156 XLSX, 1,839,481 PPTX, and 2,598,519 Node static bytes. This
// rebases exact measured entries after the DOCX heap work and merged XLSX
// formatting changes; the dispatch allowances remain unchanged.
// #1562 then moves DrawingML text phases into core, shares the bidi segment
// kernel, and bounds the negative-tracking wrap fit. Merged with f9c333d4,
// static bytes change by +98 DOCX, +13,789 XLSX, +6,502 PPTX, and +5,149 Node.
// These are measured post-merge baselines; the allowances remain unchanged.
const OOXML_BUNDLE_BASELINE = Object.freeze({
  // #1566's explicit-state line-breaker and table measurement add 18,002
  // DOCX bytes against aec306b6 (2,533,239 -> 2,551,241) in 36 chunks:
  // 18,953 and 2,501 rendered module bytes respectively, offset by minification.
  // That main (036ddd31) also includes #1586 sliced layout and #1590 projection
  // consolidation. #1561 adds another 11,595
  // static bytes after moving scroll/find behavior into core collaborators:
  // 44,570 bytes in new shared/adapter modules offset 28,748 removed bytes
  // from the old viewer/find modules; other graph changes account for the rest.
  // The static chunk count remains 36 and the dispatch allowance is unchanged.
  docx: { entry: 2_563_735, inline: 31_624, budget: 2_800 },
  // XLSX entry +8,512 bytes versus 776237df: worksheet LRU/leases and
  // viewer state restoration. The optional model-source runtime stays lazy.
  xlsx: { entry: 1_858_945, inline: 39_902, budget: 2_500 },
  // PPTX #1561 adds 13,299 static bytes against main (036ddd31): 48,632
  // bytes in shared/adapter modules offset 30,988 removed viewer/find bytes,
  // with the remaining graph changes preserving the 38 static chunks.
  // The dispatch allowance is unchanged.
  pptx: { entry: 1_845_983, inline: 59_554, budget: 2_100 },
  node: { entry: 2_603_668, budget: 3_600 },
});
// The XLSX render worker adds 536 bytes for explicit worksheet eviction and
// 51 bytes for table-style font color precedence. The DOCX worker includes
// PRs #1586 and #1590 plus the #1566 line-breaker split (+19,183 bytes
// against aec306b6). Rounding Excel serials to the nearest millisecond adds
// 24 bytes to every worker, and per-section date/time format detection with
// text-section exclusion another 618 to the XLSX worker. Workers retain zero
// allowance.
// Issue #1591 adds 539 measured bytes to the combined DOCX worker. The #1562
// shared DrawingML layout adds 9,338 XLSX, 4,116 PPTX, and 93 DOCX worker
// bytes (measured).
const OOXML_RENDER_WORKERS = [1_426_979, 1_462_664, 2_067_020];

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
  checkSourceDispatchImports(trackedSourceFiles());
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
