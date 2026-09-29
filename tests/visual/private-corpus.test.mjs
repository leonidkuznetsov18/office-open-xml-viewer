import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, realpathSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { PNG } from 'pngjs';
import {
  clearSelfVrtCandidateOutput,
  harnessBootstrapDiffViolations,
  pngPixelsEqual,
  selfVrtBaselineRoot,
  selfVrtInputPath,
  verifyBaselineCheckout,
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

function git(cwd, ...args) {
  return execFileSync('git', args, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }).trim();
}

/** A throwaway repository with two commits; HEAD is left at the first. */
function baselineRepository() {
  const root = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-baseline-checkout-'));
  git(root, 'init', '-q');
  git(root, 'config', 'user.email', 'vrt@example.invalid');
  git(root, 'config', 'user.name', 'vrt');
  writeFileSync(join(root, '.gitignore'), 'tests/visual/baseline/\n');
  writeFileSync(join(root, 'renderer.js'), 'export const version = 1;\n');
  git(root, 'add', '.');
  git(root, 'commit', '-q', '-m', 'first');
  const first = git(root, 'rev-parse', 'HEAD');
  writeFileSync(join(root, 'renderer.js'), 'export const version = 2;\n');
  git(root, 'commit', '-q', '-am', 'second');
  const second = git(root, 'rev-parse', 'HEAD');
  git(root, 'checkout', '-q', '--detach', first);
  return { root, first, second };
}

function withEnv(values, run) {
  const previous = Object.fromEntries(Object.keys(values).map((key) => [key, process.env[key]]));
  try {
    for (const [key, value] of Object.entries(values)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    return run();
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
}

test('a baseline checkout must be a clean Git checkout at the baseline revision', () => {
  const { root, first, second } = baselineRepository();
  const plain = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-baseline-plain-'));
  try {
    withEnv({ VRT_ALLOW_HARNESS_CHANGES: undefined }, () => {
      // The previous renderer's own checkout, with ignored baseline images.
      mkdirSync(join(root, 'tests/visual/baseline'), { recursive: true });
      writeFileSync(join(root, 'tests/visual/baseline/manifest.json'), '{}');
      verifyBaselineCheckout({ checkout: root, revision: first });

      // A plain directory holding copied images is not a renderer checkout.
      mkdirSync(join(plain, 'tests/visual/baseline'), { recursive: true });
      assert.throws(
        () => verifyBaselineCheckout({ checkout: plain, revision: first }),
        /not a Git checkout/,
      );
      // Nor is a directory inside a checkout.
      assert.throws(
        () => verifyBaselineCheckout({ checkout: join(root, 'tests'), revision: first }),
        /must be the root of a Git checkout/,
      );
      // The checkout must be at the baseline revision.
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: second }),
        /checkout mismatch/,
      );
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: 'HEAD' }),
        /full commit SHA/,
      );

      // A tracked modification means the images need not come from that revision.
      writeFileSync(join(root, 'renderer.js'), 'export const version = 3;\n');
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: first }),
        /clean renderer checkout; changed paths at .*: renderer\.js/,
      );
      git(root, 'checkout', '-q', '--', 'renderer.js');
      // So does an untracked source file.
      writeFileSync(join(root, 'patch.js'), 'export {};\n');
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: first }),
        /clean renderer checkout; changed paths at .*: patch\.js/,
      );
      rmSync(join(root, 'patch.js'));
      verifyBaselineCheckout({ checkout: root, revision: first });
    });

    // The harness bootstrap never excuses a renderer change.
    withEnv({ VRT_ALLOW_HARNESS_CHANGES: '1' }, () => {
      writeFileSync(join(root, 'renderer.js'), 'export const version = 3;\n');
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: first }),
        /clean renderer checkout/,
      );
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
    rmSync(plain, { recursive: true, force: true });
  }
});

test('comparison runs read only a verified previous-renderer checkout', () => {
  const { root, first } = baselineRepository();
  const plain = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-baseline-plain-'));
  // This test runs from the repository root, whose HEAD is a real commit.
  const head = git(process.cwd(), 'rev-parse', 'HEAD');
  try {
    withEnv({ VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: head }, () => {
      assert.equal(selfVrtBaselineRoot({ snapshot: true }), 'tests/visual/baseline');
      assert.throws(() => selfVrtBaselineRoot(), /VRT_BASELINE_CHECKOUT is required/);
    });
    withEnv({ VRT_BASELINE_CHECKOUT: plain, VRT_BASELINE_REVISION: head }, () => {
      assert.throws(() => selfVrtBaselineRoot({ snapshot: true }), /capture snapshots in the baseline checkout/);
      assert.throws(() => selfVrtBaselineRoot(), /not a Git checkout/);
    });
    withEnv({ VRT_BASELINE_CHECKOUT: root, VRT_BASELINE_REVISION: head }, () => {
      // A real checkout, but not at the revision the manifests are bound to.
      assert.notEqual(first, head);
      assert.throws(() => selfVrtBaselineRoot(), /checkout mismatch/);
    });
    withEnv({ VRT_BASELINE_CHECKOUT: process.cwd(), VRT_BASELINE_REVISION: head }, () => {
      assert.throws(() => selfVrtBaselineRoot(), /not this one/);
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
    rmSync(plain, { recursive: true, force: true });
  }
});
