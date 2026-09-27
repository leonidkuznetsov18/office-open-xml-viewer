import assert from 'node:assert/strict';
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import test from 'node:test';
import { assertLazySourceOwner, assertNoTopLevelModelImport, checkSourceDispatchImports } from './check-model-source-lazy-graph.mjs';

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
});

test('rejects Astra’s unconditional model-source import in a built entry', () => {
  assert.throws(() => assertNoTopLevelModelImport(
    "void import('./document-model-source-CYLDKzzV.js');", 'dist/docx.mjs',
  ), /top level/);
  assert.doesNotThrow(() => assertNoTopLevelModelImport(
    "async function load(opts) { if (opts.modelSources) await import('./document-model-source-CYLDKzzV.js'); }",
    'dist/docx.mjs',
  ));
});
