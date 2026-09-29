import { existsSync, lstatSync, mkdirSync, readFileSync, readdirSync, unlinkSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { join, parse, relative, resolve, sep } from 'node:path';
import { PNG } from 'pngjs';

const SCHEMA_VERSION = 1;

// Bootstrap-only exception for running the current VRT harness against an old
// renderer commit that predates the harness. No renderer/parser source is
// allowed in this set. Future baselines should be clean and need no exception.
// The package Vite configs are listed only because they hold the
// `@ooxml-test-*` aliases through which the fixtures load optional renderers
// (for example ChartEx); `harnessBootstrapDiffViolations` enforces that a
// bootstrap diff there adds or removes nothing but such alias lines.
const VRT_HARNESS_PATHS = new Set([
  'package.json',
  'packages/docx/package.json',
  'packages/docx/playwright.config.ts',
  'packages/docx/tests/visual/fixture.html',
  'packages/docx/vite.config.ts',
  'packages/docx/tests/visual/stable-canvas-render.mjs',
  'packages/docx/tests/visual/visual.spec.ts',
  'packages/xlsx/package.json',
  'packages/xlsx/playwright.config.ts',
  'packages/xlsx/tests/visual/fixture.html',
  'packages/xlsx/vite.config.ts',
  'packages/xlsx/tests/visual/visual.spec.ts',
  'packages/pptx/package.json',
  'packages/pptx/playwright.config.ts',
  'packages/pptx/tests/visual/fixture.html',
  'packages/pptx/vite.config.ts',
  'packages/pptx/tests/visual/visual.spec.ts',
  'tests/visual/private-corpus.mjs',
]);

const TEST_RENDERER_ALIAS_LINE =
  /^[+-]\s*'@ooxml-test-[a-z0-9-]+': resolve\((?:__)?dirname, '\.\.\/\.\.\/src\/[a-z0-9-]+\.ts'\),$/;

/** Lines of a `git diff -U0` for an allowlisted Vite config that are not a
 * test-renderer alias addition/removal. Any such line makes the harness
 * bootstrap unsafe, because the config also drives the package build. */
export function harnessBootstrapDiffViolations(diff) {
  return diff.split('\n').filter((line) =>
    (line.startsWith('+') || line.startsWith('-'))
    && !line.startsWith('+++')
    && !line.startsWith('---')
    && !TEST_RENDERER_ALIAS_LINE.test(line));
}

function gitRevision(revision) {
  return execFileSync('git', ['rev-parse', `${revision}^{commit}`], {
    encoding: 'utf8',
  }).trim();
}

function baselineRevision(snapshot = false) {
  const revision = process.env.VRT_BASELINE_REVISION?.trim();
  if (!revision) {
    throw new Error('VRT_BASELINE_REVISION is required for private corpus self-VRT');
  }
  const resolved = gitRevision(revision);
  if (snapshot) {
    const checkout = gitRevision('HEAD');
    if (checkout !== resolved) {
      throw new Error(
        `private self-VRT snapshot checkout mismatch: expected ${resolved}, running ${checkout}`,
      );
    }
    const root = execFileSync('git', ['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();
    const tracked = execFileSync(
      'git',
      ['diff', '--name-only', 'HEAD'],
      { cwd: root, encoding: 'utf8' },
    ).trim().split('\n').filter(Boolean);
    const untracked = execFileSync(
      'git',
      ['ls-files', '--others', '--exclude-standard'],
      { cwd: root, encoding: 'utf8' },
    ).trim().split('\n').filter(Boolean).filter((path) =>
      !/(^|\/)node_modules(?:\/|$)/.test(path)
      && !/^packages\/(docx|xlsx|pptx)\/public\/private(?:\/|$)/.test(path));
    const changed = [...new Set([...tracked, ...untracked])];
    if (changed.length > 0) {
      const harnessBootstrap = process.env.VRT_ALLOW_HARNESS_CHANGES === '1'
        && changed.every((path) => VRT_HARNESS_PATHS.has(path));
      if (!harnessBootstrap) {
        throw new Error(
          'private self-VRT snapshot requires a clean renderer checkout; changed paths: '
          + changed.join(', '),
        );
      }
      for (const path of changed.filter((changedPath) => changedPath.endsWith('vite.config.ts'))) {
        // An untracked config has no HEAD version to diff against, so its
        // whole content would escape the alias-only bound.
        if (untracked.includes(path)) {
          throw new Error(
            `private self-VRT harness bootstrap cannot add an untracked ${path}`,
          );
        }
        const violations = harnessBootstrapDiffViolations(execFileSync(
          'git', ['diff', '-U0', 'HEAD', '--', path], { cwd: root, encoding: 'utf8' },
        ));
        if (violations.length > 0) {
          throw new Error(
            `private self-VRT harness bootstrap may only change test renderer aliases in ${path}: `
            + violations.join(' | '),
          );
        }
      }
    }
  }
  return resolved;
}

// Self-VRT corpora. `demo` is the tracked public demo set under `public/demo/`;
// `private` is the local, gitignored corpus under `public/private/<format>/`.
// Both use the same exact-pixel previous-renderer oracle and the same manifest
// binding; they differ only in where the inputs live and where their images go.
const CORPORA = {
  demo: {
    publicPrefix: '',
    outputPrefix: '',
    manifestDirectory: 'demo',
    stemPattern: /^demo\/[^/\\]+$/,
  },
  private: {
    publicPrefix: 'private/',
    outputPrefix: 'private-corpus/',
    manifestDirectory: 'private-corpus',
    stemPattern: /^(docx|xlsx|pptx)\/[^/\\]+$/,
  },
};

function corpusConfig(corpus) {
  const config = CORPORA[corpus];
  if (!config) throw new Error(`unknown self-VRT corpus: ${corpus}`);
  return config;
}

/** Sorted input files of a corpus, relative to that corpus's public prefix
 * (`demo/sample-1.pptx`, `pptx/deck.pptx`). */
export function selfVrtCorpusFiles({ corpus, format }) {
  const directory = corpus === 'demo' ? 'demo' : format;
  const { publicPrefix } = corpusConfig(corpus);
  return readdirSync(`public/${publicPrefix}${directory}`)
    .filter((file) => file.endsWith(`.${format}`) && !file.startsWith('~$'))
    .map((file) => `${directory}/${file}`)
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

/** URL path of a corpus input, relative to the package's public root. */
export function selfVrtInputPath({ corpus, file }) {
  return `${corpusConfig(corpus).publicPrefix}${file}`;
}

/** Directory holding previous-renderer images. Snapshots are always written
 * into the checkout that renders them. A comparison run reads the same
 * package's baseline from `VRT_BASELINE_CHECKOUT` (the previous-renderer
 * worktree) when set, so candidate and baseline never share a checkout. */
export function selfVrtBaselineRoot({ snapshot = false } = {}) {
  const checkout = process.env.VRT_BASELINE_CHECKOUT?.trim();
  if (!checkout) return 'tests/visual/baseline';
  if (snapshot) {
    throw new Error(
      'VRT_BASELINE_CHECKOUT is for comparison runs; capture snapshots in the baseline checkout itself',
    );
  }
  const root = execFileSync('git', ['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();
  const baseline = resolve(checkout, relative(root, process.cwd()), 'tests/visual/baseline');
  if (resolve(baseline) === resolve('tests/visual/baseline')) {
    throw new Error('VRT_BASELINE_CHECKOUT must name the previous-renderer checkout, not this one');
  }
  return baseline;
}

function corpusFiles(corpus, files) {
  return files.map((name) => ({
    name,
    sha256: createHash('sha256')
      .update(readFileSync(`public/${selfVrtInputPath({ corpus, file: name })}`))
      .digest('hex'),
  }));
}

function readJson(path) {
  if (!existsSync(path)) throw new Error(`missing previous-renderer manifest: ${path}`);
  return JSON.parse(readFileSync(path, 'utf8'));
}

function assertExactManifest(actual, expected, path) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(
      `previous-renderer manifest mismatch at ${path}\n`
      + `expected ${JSON.stringify(expected)}\nreceived ${JSON.stringify(actual)}`,
    );
  }
}

/** Fail closed on an empty/stale corpus and bind every baseline to an explicit
 * merge-base revision. This prevents a candidate server or unrelated old
 * snapshot from silently becoming its own regression oracle. */
export function prepareSelfVrtCorpus({ corpus, format, files, snapshot }) {
  if (files.length === 0) {
    throw new Error(`${format} ${corpus} corpus is empty; zero-test self-VRT is not coverage`);
  }
  const root = `${selfVrtBaselineRoot({ snapshot })}/${corpusConfig(corpus).manifestDirectory}`;
  const path = `${root}/manifest.json`;
  const manifest = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    files: corpusFiles(corpus, files),
  };
  if (snapshot) {
    mkdirSync(root, { recursive: true });
    writeFileSync(path, `${JSON.stringify(manifest, null, 2)}\n`);
  } else {
    assertExactManifest(readJson(path), manifest, path);
  }
}

/** A prior candidate run may have rendered more pages than this one. Remove
 * only this input's generated screenshots before capture, so downstream
 * reviews cannot mistake stale page-N files for the current renderer output.
 * Baselines and non-image evidence are never touched. */
export function clearSelfVrtCandidateOutput({
  corpus,
  stem,
  itemKind,
  outputRoot = `tests/visual/screenshots/${corpusConfig(corpus).outputPrefix}`,
}) {
  if (!corpusConfig(corpus).stemPattern.test(stem)
    || ['.', '..'].includes(stem.split('/')[1])
    || !/^(page|sheet|slide)$/.test(itemKind)) {
    throw new Error('invalid self-VRT output identity');
  }
  const directory = resolve(outputRoot, stem);
  // Generated-output directories can be local symlinks. Never follow one when
  // removing stale files: it may point into another worktree or dependencies.
  let ancestor = parse(directory).root;
  for (const component of directory.slice(ancestor.length).split(sep).filter(Boolean)) {
    ancestor = join(ancestor, component);
    if (existsSync(ancestor) && lstatSync(ancestor).isSymbolicLink()) {
      throw new Error(`refusing to clear symlinked self-VRT output: ${ancestor}`);
    }
  }
  if (!existsSync(directory)) return;
  const item = new RegExp(`^${itemKind}-\\d+\\.png$`);
  for (const file of readdirSync(directory)) {
    if (item.test(file)) unlinkSync(`${directory}/${file}`);
  }
}

/** Verify that the baseline contains exactly the complete page/sheet/slide set
 * reported by the previous renderer. An item-count reduction or stale extra PNG
 * is therefore a hard failure instead of an ignored file. */
export function verifySelfVrtItemManifest({
  corpus,
  format,
  stem,
  itemKind,
  itemCount,
  snapshot,
}) {
  const directory = `${selfVrtBaselineRoot({ snapshot })}/${corpusConfig(corpus).outputPrefix}${stem}`;
  const path = `${directory}/manifest.json`;
  const items = Array.from({ length: itemCount }, (_, index) => `${itemKind}-${index + 1}.png`);
  const manifest = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    itemKind,
    itemCount,
    items,
  };
  if (snapshot) {
    mkdirSync(directory, { recursive: true });
    writeFileSync(path, `${JSON.stringify(manifest, null, 2)}\n`);
    return;
  }
  assertExactManifest(readJson(path), manifest, path);
  const actualItems = readdirSync(directory)
    .filter((file) => new RegExp(`^${itemKind}-\\d+\\.png$`).test(file))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  if (JSON.stringify(actualItems) !== JSON.stringify(items)) {
    throw new Error(
      `previous-renderer item set mismatch at ${directory}\n`
      + `expected ${JSON.stringify(items)}\nreceived ${JSON.stringify(actualItems)}`,
    );
  }
}

/** Persist the candidate artifact and return a diagnostic instead of throwing,
 * so one changed item cannot prevent later items in the same document from
 * being rendered and compared. */
export function captureOrCompareSelfVrtItem({
  corpus,
  stem,
  itemKind,
  itemIndex,
  actual,
  snapshot,
}) {
  const key = `${itemKind}-${itemIndex + 1}.png`;
  const { outputPrefix } = corpusConfig(corpus);
  const outputDirectory = `tests/visual/${snapshot ? 'baseline' : 'screenshots'}/${outputPrefix}${stem}`;
  mkdirSync(outputDirectory, { recursive: true });
  writeFileSync(`${outputDirectory}/${key}`, actual);
  if (snapshot) return null;
  const baselinePath = `${selfVrtBaselineRoot()}/${outputPrefix}${stem}/${key}`;
  if (!existsSync(baselinePath)) return `missing previous-renderer baseline: ${baselinePath}`;
  return pngPixelsEqual(actual, readFileSync(baselinePath))
    ? null
    : `${stem} ${key} differs from the previous renderer`;
}

/** PNG byte streams may differ in encoder metadata/compression while decoding
 * to the same canvas. Self-VRT compares the actual rendered RGBA pixels; width
 * or height changes remain regressions. */
export function pngPixelsEqual(leftBuffer, rightBuffer) {
  const left = PNG.sync.read(leftBuffer);
  const right = PNG.sync.read(rightBuffer);
  return left.width === right.width
    && left.height === right.height
    && left.data.equals(right.data);
}
