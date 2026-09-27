import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import test from 'node:test';
import { assertLazySourceOwner, assertNoTopLevelModelImport, assertEagerSourceEntries, assertNoSourceRuntime, checkBuiltDispatchBudget, checkDispatchShimBudget, checkSourceDispatchImports } from './check-model-source-lazy-graph.mjs';

test('a static owner import in the worker graph fails, while a selected-source import stays lazy', () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-source-graph-'));
  try {
    const internal = join(root, 'internal');
    mkdirSync(internal);
    const owner = join(internal, 'worker-document-source.ts');
    const worker = join(root, 'worker.ts');
    writeFileSync(owner, 'export class Owner {}\n');
    writeFileSync(worker, "if (source) await import('./internal/worker-document-source.js');\n");
    assert.doesNotThrow(() => assertLazySourceOwner(worker, owner));
    writeFileSync(worker, "import { Owner } from './internal/worker-document-source.js';\nif (source) await import('./internal/worker-document-source.js');\n");
    assert.throws(() => assertLazySourceOwner(worker, owner), /statically reaches/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('rejects a source import moved outside the modelSources dispatch', () => {
  const path = 'packages/docx/src/document.ts';
  const dispatch = "class DocxDocument { static async load(opts) { if (opts.modelSources) return import('./internal/document-model-source.js'); } }";
  assert.doesNotThrow(() => checkSourceDispatchImports([{ path, text: dispatch }]));
  assert.throws(() => checkSourceDispatchImports([{
    path, text: `${dispatch}\nvoid import('./internal/document-model-source.js');`,
  }]), /outside an allowed dispatch function/);
  assert.throws(() => checkSourceDispatchImports([{
    path, text: "class DocxDocument { static async load(opts) { return import('./internal/document-model-source.js'); } }",
  }]), /outside modelSources presence dispatch/);
});

test('rejects an unconditional model-source import in a built entry', () => {
  assert.throws(() => assertNoTopLevelModelImport(
    "void import('./document-model-source-CYLDKzzV.js');", 'dist/docx.mjs',
  ), /top level/);
  assert.doesNotThrow(() => assertNoTopLevelModelImport(
    "async function load(opts) { if (opts.modelSources) await import('./document-model-source-CYLDKzzV.js'); }",
    'dist/docx.mjs',
  ));
});

test('each eager entry type rejects a transitive static model-source import', () => {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-source-entry-'));
  try {
    const owner = join(root, 'internal', 'document-model-source.ts');
    mkdirSync(join(root, 'internal'));
    writeFileSync(owner, 'export const source = true;\n');
    for (const name of ['core-index', 'package-index', 'worker', 'render-worker', 'node-index']) {
      const entry = join(root, `${name}.ts`);
      const bridge = join(root, `${name}-bridge.ts`);
      writeFileSync(entry, `import './${name}-bridge.js';\n`);
      writeFileSync(bridge, "import './internal/document-model-source.js';\n");
      assert.throws(() => assertEagerSourceEntries([entry]), /statically reaches/);
      writeFileSync(bridge, "void import('./internal/document-model-source.js');\n");
      assert.doesNotThrow(() => assertEagerSourceEntries([entry]));
    }
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('a model-source marker leaking into eager emitted code fails', () => {
  assert.throws(() => assertNoSourceRuntime('const marker="ooxml-model-source-module/v1";', 'entry'), /optional source runtime/);
});

test('the dispatch budget measures only the selected-source branch', () => {
  const path = 'packages/xlsx/src/workbook.ts';
  const dispatch = "async function load(opts) { if (opts.modelSources !== undefined) { return import('./internal/workbook-model-source.js'); } }";
  const ordinary = 'const ordinary = "' + 'x'.repeat(2000) + '";\n';
  assert.doesNotThrow(() => checkDispatchShimBudget([{ path, text: ordinary + dispatch }]));
  const grown = dispatch.replace('return import', `const padding = '${'x'.repeat(600)}'; return import`);
  assert.throws(() => checkDispatchShimBudget([{ path, text: ordinary + grown }]), /dispatch shim.*budget/);
  const emitted = "async function load(o) { if (o.modelSources !== void 0) { return import('./workbook-model-source-hash.js'); } }";
  assert.doesNotThrow(() => checkBuiltDispatchBudget([ordinary + emitted], 1, 'xlsx'));
  const emittedGrown = emitted.replace('return import', `const padding = '${'x'.repeat(300)}'; return import`);
  assert.throws(() => checkBuiltDispatchBudget([ordinary + emittedGrown], 1, 'xlsx'), /emitted dispatch shim.*budget/);
});
