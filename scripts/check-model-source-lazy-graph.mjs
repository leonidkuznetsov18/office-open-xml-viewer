#!/usr/bin/env node
// The optional model-source implementation must stay outside every ordinary
// OOXML entry and worker. Resolve workspace package exports with TypeScript,
// then inspect the emitted static JS graph and decoded inline worker payloads.
// This guards accidental eager coupling by maintainers or agents; it is not a
// security boundary against deliberately adversarial JavaScript.
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { resolve, dirname, relative, join } from 'node:path';
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

// These are the small presence branches retained in ordinary browser and Node
// entries. Optional source owners and worker-source sidecars are separate chunks,
// so their implementation bytes are never part of this dispatch budget.
const EAGER_DISPATCH = new Map([
  ['packages/docx/src/document.ts', 1],
  ['packages/xlsx/src/workbook.ts', 1],
  ['packages/pptx/src/presentation.ts', 1],
  ['packages/node/src/docx.ts', 2],
  ['packages/node/src/xlsx.ts', 1],
  ['packages/node/src/pptx.ts', 1],
]);
const DISPATCH_SHIM_BUDGET = 512;
const EMITTED_DISPATCH_BUDGET = 256;

export function checkDispatchShimBudget(files) {
  for (const { path, text } of files) {
    const expected = EAGER_DISPATCH.get(path);
    if (expected === undefined) throw new Error(`Unknown eager dispatch owner ${path}`);
    const ast = parse(text, { sourceType: 'module', plugins: ['typescript', 'jsx'] });
    const branches = [];
    function walk(node) {
      if (!node || typeof node !== 'object' || !node.type) return;
      if (node.type === 'IfStatement' && text.slice(node.test.start, node.test.end).includes('modelSources')) {
        branches.push(node);
      }
      for (const value of Object.values(node)) {
        if (Array.isArray(value)) value.forEach(walk);
        else if (value && typeof value === 'object' && value.type) walk(value);
      }
    }
    walk(ast);
    if (branches.length !== expected) {
      throw new Error(`${path} expected ${expected} modelSources dispatch branches, found ${branches.length}`);
    }
    for (const branch of branches) {
      const shim = text.slice(branch.start, branch.consequent.end);
      if (!/import\([^)]*model-source/.test(shim)) {
        throw new Error(`${path} modelSources dispatch has no selected-source import`);
      }
      const bytes = Buffer.byteLength(shim);
      if (bytes > DISPATCH_SHIM_BUDGET) {
        throw new Error(`${path} dispatch shim exceeds ${DISPATCH_SHIM_BUDGET}-byte budget: ${bytes}`);
      }
      console.log(`${path} dispatch shim: ${bytes}/${DISPATCH_SHIM_BUDGET} source bytes`);
    }
  }
}

export function checkSourceDispatchImports(files) {
  for (const { path, text } of files) {
    if (/\.(?:test|spec|stories|probe)\.[cm]?[jt]sx?$/.test(path)) continue;
    const ast = parse(text, { sourceType: 'module', plugins: ['typescript', 'jsx'] });
    const allowed = SOURCE_DISPATCH.get(path) ?? [];
    function walk(node, functionName, presenceGuard = false) {
      if (!node || typeof node !== 'object' || !node.type) return;
      let current = functionName;
      if (node.type === 'FunctionDeclaration') current = node.id?.name;
      if (['ClassMethod', 'ObjectMethod'].includes(node.type)) current = node.key?.name;
      if (['ArrowFunctionExpression', 'FunctionExpression'].includes(node.type)) {
        // Preserve the containing named dispatch for its local callback/IIFE.
        current = functionName;
      }
      if (node.type === 'IfStatement') {
        walk(node.test, current, presenceGuard);
        walk(node.consequent, current,
          presenceGuard || text.slice(node.test.start, node.test.end).includes('modelSources'));
        walk(node.alternate, current, presenceGuard);
        return;
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
        if (sourceImport && EAGER_DISPATCH.has(path) && !presenceGuard) {
          throw new Error(`${path}:${node.loc?.start.line} model-source import outside modelSources presence dispatch`);
        }
      }
      // Assignment to self.onmessage is a named worker dispatch boundary.
      if (node.type === 'AssignmentExpression' && node.left.type === 'MemberExpression'
        && node.left.object.name === 'self' && node.left.property.name === 'onmessage') {
        walk(node.right, 'self.onmessage', presenceGuard);
        return;
      }
      for (const value of Object.values(node)) {
        if (Array.isArray(value)) value.forEach((child) => walk(child, current, presenceGuard));
        else if (value && typeof value === 'object' && value.type) walk(value, current, presenceGuard);
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

// Measure the emitted presence branch, including Vite's dynamic import
// factoring. Unrelated OOXML code in the same chunk never enters this count.
export function checkBuiltDispatchBudget(codes, expected, label) {
  const shims = [];
  for (const code of codes) {
    const ast = parse(code, { sourceType: 'module' });
    function walk(node) {
      if (!node || typeof node !== 'object' || !node.type) return;
      if (node.type === 'IfStatement'
        && code.slice(node.test.start, node.test.end).includes('modelSources')) {
        const shim = code.slice(node.start, node.consequent.end);
        if (/import\([^)]*model-source/.test(shim)) shims.push(Buffer.byteLength(shim));
      }
      for (const value of Object.values(node)) {
        if (Array.isArray(value)) value.forEach(walk);
        else if (value && typeof value === 'object' && value.type) walk(value);
      }
    }
    walk(ast);
  }
  if (shims.length !== expected) {
    throw new Error(`${label} expected ${expected} emitted modelSources dispatch shims, found ${shims.length}`);
  }
  for (const bytes of shims) {
    if (bytes > EMITTED_DISPATCH_BUDGET) {
      throw new Error(`${label} emitted dispatch shim exceeds ${EMITTED_DISPATCH_BUDGET}-byte budget: ${bytes}`);
    }
    console.log(`${label} emitted dispatch shim: ${bytes}/${EMITTED_DISPATCH_BUDGET} bytes`);
  }
}

export function assertNoSourceRuntime(code, name) {
  if (code.includes('ooxml-model-source-module/v1') || code.includes('model source view default')) {
    throw new Error(`${name} includes optional source runtime in its eager OOXML code`);
  }
}

export function checkBuiltBundles(dist = 'dist', { packages = false } = {}) {
  for (const format of ['docx', 'xlsx', 'pptx', 'node']) {
    const graph = bundleGraph(join(dist, `${format}.mjs`));
    assertNoSourceRuntime(graph.joined, `${format} static entry graph`);
    checkBuiltDispatchBudget(graph.codes, format === 'node' ? 4 : 1, format);
    console.log(`${format} static JS: ${graph.bytes} bytes across ${graph.files} files`);
    if (format !== 'node') {
      for (const payload of graph.codes.flatMap(inlineWorkers)) {
        assertNoSourceRuntime(payload, `${format} inline worker`);
        console.log(`${format} inline worker: ${Buffer.byteLength(payload)} decoded bytes`);
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
  let ordinaryWorkers = 0;
  for (const file of readdirSync(join(dist, 'assets')).filter((name) => /^render-worker-.*\.js$/.test(name))) {
    if (!file.startsWith('render-worker-source-')) ordinaryWorkers++;
    assertNoSourceRuntime(readFileSync(join(dist, 'assets', file), 'utf8'), file);
  }
  if (ordinaryWorkers !== 3) {
    throw new Error(`Expected 3 ordinary render workers, found ${ordinaryWorkers}`);
  }
}

export function assertEagerSourceEntries(entries, forbidden = []) {
  const forbiddenPaths = new Set(forbidden.map((module) => resolve(module)));
  for (const entry of entries) {
    for (const module of eagerModules(entry)) {
      // Match future model-source owners as well as today's explicit list.
      // A type-only import is absent from eagerModules, as it should be.
      if (forbiddenPaths.has(module) || /(?:^|\/)(?:[^/]*model-source[^/]*|worker-source|render-worker-source)\.[cm]?[jt]sx?$/.test(module)) {
        throw new Error(`${entry} statically reaches ${module}`);
      }
    }
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  checkSourceDispatchImports(trackedSourceFiles());
  checkDispatchShimBudget([...EAGER_DISPATCH.keys()].map((path) => ({ path, text: readFileSync(path, 'utf8') })));
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
  assertEagerSourceEntries([
    'packages/core/src/index.ts',
    ...['docx', 'xlsx', 'pptx'].flatMap((format) => [
      `packages/${format}/src/index.ts`,
      `packages/${format}/src/worker.ts`,
      `packages/${format}/src/render-worker.ts`,
    ]),
    'packages/node/src/index.ts',
  ], forbidden);
  if (existsSync('dist/docx.mjs')) checkBuiltBundles('dist', { packages: process.argv.includes('--packages') });
  console.log('OOXML entries and workers keep model-source runtime behind dynamic loads.');
}
