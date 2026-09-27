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
const parsedSource = new Map();

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
    let imports = parsedSource.get(file);
    if (!imports) {
      const ast = parse(readFileSync(file, 'utf8'), { sourceType: 'module', plugins: ['typescript', 'jsx'] });
      imports = staticSpecifiers(ast);
      parsedSource.set(file, imports);
    }
    const importer = relative(process.cwd(), file).replaceAll('\\', '/');
    for (const specifier of imports) {
      if (specifier.includes('?')) continue;
      const target = resolveModule(importer, specifier);
      if (target && target.startsWith(process.cwd())) pending.push(resolve(target));
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

export function checkBuiltBundles(dist = 'dist') {
  for (const format of ['docx', 'xlsx', 'pptx', 'node']) {
    const graph = bundleGraph(join(dist, `${format}.mjs`));
    assertNoSourceRuntime(graph.joined, `${format} static entry graph`);
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
      if (existsSync(join('packages', format, 'dist')) && !existsSync(packageSidecar)) {
        throw new Error(`Missing optional package source sidecar ${packageSidecar}`);
      }
      if (existsSync(packageSidecar)
        && staticSpecifiers(parse(readFileSync(packageSidecar, 'utf8'), { sourceType: 'module' })).length > 0) {
        throw new Error(`${packageSidecar} is not self-contained`);
      }
      const packageEntry = join('packages', format, 'dist', 'index.mjs');
      if (existsSync(packageEntry)) {
        const packageGraph = bundleGraph(packageEntry);
        assertNoSourceRuntime(packageGraph.joined, `${format} package static entry graph`);
        for (const payload of packageGraph.codes.flatMap(inlineWorkers)) {
          assertNoSourceRuntime(payload, `${format} package inline worker`);
        }
      }
    }
  }
  for (const file of readdirSync(join(dist, 'assets')).filter((name) => /^render-worker-.*\.js$/.test(name))) {
    assertNoSourceRuntime(readFileSync(join(dist, 'assets', file), 'utf8'), file);
  }
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
  if (existsSync('dist/docx.mjs')) checkBuiltBundles();
  console.log('OOXML entries and workers keep model-source runtime behind dynamic loads.');
}
