import assert from 'node:assert/strict';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';
import { PNG } from 'pngjs';
import {
  clearSelfVrtCandidateOutput,
  harnessBootstrapDiffViolations,
  pngPixelsEqual,
  selfVrtBaselineRoot,
  selfVrtInputPath,
} from './private-corpus.mjs';

test('private corpus self-VRT compares decoded pixels, not encoder bytes', () => {
  const image = new PNG({ width: 2, height: 1 });
  image.data.set([255, 0, 0, 255, 0, 0, 255, 255]);
  const fast = PNG.sync.write(image, { deflateLevel: 0 });
  const compact = PNG.sync.write(image, { deflateLevel: 9 });

  assert.equal(fast.equals(compact), false);
  assert.equal(pngPixelsEqual(fast, compact), true);
});

test('private corpus self-VRT rejects a one-channel pixel change', () => {
  const left = new PNG({ width: 1, height: 1 });
  left.data.set([1, 2, 3, 255]);
  const right = new PNG({ width: 1, height: 1 });
  right.data.set([1, 2, 4, 255]);

  assert.equal(pngPixelsEqual(PNG.sync.write(left), PNG.sync.write(right)), false);
});

test('candidate capture discards stale pages without following local evidence symlinks', () => {
  const root = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-private-vrt-output-'));
  try {
    const directory = join(root, 'docx', 'case');
    mkdirSync(directory, { recursive: true });
    writeFileSync(join(directory, 'page-29.png'), 'stale');
    writeFileSync(join(directory, 'notes.json'), 'retain');
    clearSelfVrtCandidateOutput({
      corpus: 'private', stem: 'docx/case', itemKind: 'page', outputRoot: root,
    });
    assert.equal(existsSync(join(directory, 'page-29.png')), false);
    assert.equal(readFileSync(join(directory, 'notes.json'), 'utf8'), 'retain');

    symlinkSync(directory, join(root, 'docx', 'linked'), 'dir');
    assert.throws(() => clearSelfVrtCandidateOutput({
      corpus: 'private', stem: 'docx/linked', itemKind: 'page', outputRoot: root,
    }), /symlinked self-VRT output/);
    assert.equal(readFileSync(join(directory, 'notes.json'), 'utf8'), 'retain');
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('harness bootstrap allows only test-renderer alias lines in Vite configs', () => {
  const header = [
    'diff --git a/packages/pptx/vite.config.ts b/packages/pptx/vite.config.ts',
    '--- a/packages/pptx/vite.config.ts',
    '+++ b/packages/pptx/vite.config.ts',
    '@@ -20,0 +21 @@',
  ];
  const alias = "+      '@ooxml-test-chart-ex-renderer': resolve(dirname, '../../src/chart-ex.ts'),";
  assert.deepEqual(harnessBootstrapDiffViolations([...header, alias].join('\n')), []);
  assert.deepEqual(
    harnessBootstrapDiffViolations([
      ...header,
      alias,
      "+  define: { __OOXML_MODEL_SOURCES__: 'false' },",
    ].join('\n')),
    ["+  define: { __OOXML_MODEL_SOURCES__: 'false' },"],
  );
});

test('self-VRT corpora keep demo and private stems in their own namespaces', () => {
  assert.equal(selfVrtInputPath({ corpus: 'demo', file: 'demo/sample-1.docx' }), 'demo/sample-1.docx');
  assert.equal(selfVrtInputPath({ corpus: 'private', file: 'docx/case.docx' }), 'private/docx/case.docx');
  assert.throws(() => clearSelfVrtCandidateOutput({
    corpus: 'demo', stem: 'docx/case', itemKind: 'page',
  }), /invalid self-VRT output identity/);
  assert.throws(() => clearSelfVrtCandidateOutput({
    corpus: 'private', stem: 'demo/sample-1', itemKind: 'page',
  }), /invalid self-VRT output identity/);
  assert.throws(() => selfVrtInputPath({ corpus: 'public', file: 'x' }), /unknown self-VRT corpus/);
});

test('comparison runs can read the previous renderer from another checkout', () => {
  const previous = process.env.VRT_BASELINE_CHECKOUT;
  try {
    delete process.env.VRT_BASELINE_CHECKOUT;
    assert.equal(selfVrtBaselineRoot(), 'tests/visual/baseline');

    const checkout = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-baseline-checkout-'));
    try {
      process.env.VRT_BASELINE_CHECKOUT = checkout;
      // This test runs from the repository root, so the package-relative
      // prefix is empty.
      assert.equal(selfVrtBaselineRoot(), resolve(checkout, 'tests/visual/baseline'));
      assert.throws(() => selfVrtBaselineRoot({ snapshot: true }), /capture snapshots in the baseline checkout/);
      process.env.VRT_BASELINE_CHECKOUT = process.cwd();
      assert.throws(() => selfVrtBaselineRoot(), /not this one/);
    } finally {
      rmSync(checkout, { recursive: true, force: true });
    }
  } finally {
    if (previous === undefined) delete process.env.VRT_BASELINE_CHECKOUT;
    else process.env.VRT_BASELINE_CHECKOUT = previous;
  }
});
