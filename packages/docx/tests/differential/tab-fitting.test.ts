/// <reference types="node" />
import { execFileSync } from 'node:child_process';
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { beforeAll, expect, it } from 'vitest';
import init, { DocxArchive } from '../../src/wasm/docx_parser.js';
import { normalizeInternalDocumentModel } from '../../src/parser-model.js';
import { layoutDocument } from '../../src/document-layout.js';
import { createLayoutServices } from '../../src/layout-runtime.js';
import type { ParagraphLayout } from '../../src/layout/types.js';
import { documentBytes, measureContext, matrix, fittingMatrix } from '../../src/test-support/tab-fitting.test-support.js';

// Run with DOCX_TAB_BASELINE_CHECKOUT pointing to a clean detached origin/main
// checkout: vitest run --config packages/docx/tests/differential/tab-fitting.config.ts.
// Compare observable line partitions and placement geometry, never private
// breaker state. The two layout graphs use the same parsed facts and metrics.
let baselineLayout: typeof layoutDocument;
let baselineServices: typeof createLayoutServices;
beforeAll(async () => {
  const base = process.env.DOCX_TAB_BASELINE_CHECKOUT;
  if (!base) throw new Error('DOCX_TAB_BASELINE_CHECKOUT must name a clean origin/main checkout');
  const gitEnv = Object.fromEntries(Object.entries(process.env).filter(([key]) => !/^GIT_/i.test(key)));
  const git = (cwd: string, args: string[]) => execFileSync('git', ['-C', cwd, ...args], { encoding: 'utf8', env: gitEnv }).trim();
  expect(git(base, ['rev-parse', 'HEAD'])).toBe(git(process.cwd(), ['rev-parse', 'origin/main']));
  expect(git(base, ['status', '--porcelain', '--untracked-files=no'])).toBe('');
  baselineLayout = (await import(/* @vite-ignore */ resolve(base, 'packages/docx/src/document-layout.ts'))).layoutDocument;
  baselineServices = (await import(/* @vite-ignore */ resolve(base, 'packages/docx/src/layout-runtime.ts'))).createLayoutServices;
  await init({ module_or_path: await readFile(new URL('../../src/wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function compare(bytes: Uint8Array) {
  const archive = new DocxArchive(bytes);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const paragraph = (layout: ReturnType<typeof layoutDocument>) => {
    const node = layout.pages[0]?.layers.body.find((item) => item.kind === 'paragraph');
    if (node?.kind !== 'paragraph') throw new Error('Missing paragraph');
    return node;
  };
  const candidate = paragraph(layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), { currentDateMs: 0 }));
  const baseline = paragraph(baselineLayout(model, baselineServices(model, { measureContext: measureContext() }), { currentDateMs: 0 }));
  return { candidate, baseline };
}

function outOfBand(paragraph: ParagraphLayout, indent: string, rtl: boolean) {
  const twips = (name: string) => Number(new RegExp(`w:${name}="(\\d+)"`).exec(indent)?.[1] ?? 0) / 20;
  return paragraph.lines.some((line, index) => line.placements.some((node) => {
    if (!node.bounds || (node.kind === 'text' && !node.text.trim())) return false;
    const first = index === 0 ? twips('firstLine') - twips('hanging') : 0;
    const start = 72 + (rtl ? twips('right') : twips('left') + first);
    const end = 540 - (rtl ? twips('left') + first : twips('right'));
    const trailing = node.kind === 'text' ? (node.text.length - node.text.trimEnd().length) * 5 : 0;
    return node.bounds.xPt < start - 1e-9 || node.bounds.xPt + node.bounds.widthPt - trailing > end + 1e-9;
  }));
}

const geometry = (paragraph: ParagraphLayout) => paragraph.lines.map((line) => ({
  bounds: line.bounds,
  placements: line.placements.map((node) => ({ kind: node.kind, bounds: node.bounds, ...('text' in node ? { text: node.text } : {}) })),
}));

const matrixCases = [...matrix, ...fittingMatrix].filter((entry) => entry.float === 'none').map((entry) => ({
  indent: 'w:left="720" w:hanging="720"', rtl: entry.rtl,
  alignment: entry.alignment, count: entry.count,
  text: 'content' in entry ? entry.content : entry.alignment === 'decimal' ? '12.3' : 'word',
  positional: 'relativeTo' in entry, relativeTo: 'relativeTo' in entry ? entry.relativeTo : undefined,
  automatic: entry.kind === 'automatic',
}));
const generatedCases = ['', 'w:left="720" w:hanging="720"', 'w:left="720" w:right="1440"'].flatMap((indent) =>
  [false, true].flatMap((rtl) => [1, 2, 3].flatMap((count) =>
    ['left', 'start', 'right', 'end', 'center', 'decimal', 'bar', 'clear', 'num'].flatMap((alignment) =>
      [720, 2400, 5600, 11000].flatMap((stop) => ['', 'prefix '].flatMap((prefix) =>
        ['word', 'word '.repeat(15).trim(), '漢'.repeat(50), 'ภาษาไทย'.repeat(10)].map((text) =>
          ({ indent, rtl, count, alignment, stop, prefix, text, positional: false, relativeTo: undefined, automatic: false }))))))));

it('preserves every in-band no-float matrix and generated tab layout from origin/main', async () => {
  let exceptions = 0;
  let identical = 0;
  let overflow = 0;
  for (const entry of [...matrixCases, ...generatedCases]) {
    const { candidate, baseline } = compare(documentBytes(entry.indent, entry.rtl, 0, { ...entry, noFloat: true }));
    // Exactly the five short Word-backed RTL positional controls documented in
    // #1672: margin-left, both centers, both rights. Long cells are not exempt.
    const exception = entry.positional && entry.rtl && entry.text === 'word'
      && (entry.alignment !== 'left' || entry.relativeTo === 'margin');
    const content = candidate.lines.flatMap((line) => line.placements).filter((node) => node.kind === 'text').map((node) => node.text).join('');
    const prefix = 'prefix' in entry ? entry.prefix : '';
    // Placements are visual-order cells; RTL can reverse their logical order.
    // Compare scalar inventories to detect dropped or duplicated content.
    const inventory = (text: string) => [...text.replace(/\s/g, '')].sort();
    expect(inventory(content), JSON.stringify(entry)).toEqual(inventory(`${prefix}${entry.text}`));
    expect(outOfBand(candidate, entry.indent, entry.rtl), JSON.stringify(entry)).toBe(false);
    if (exception) { exceptions += 1; continue; }
    if (outOfBand(baseline, entry.indent, entry.rtl)) { overflow += 1; continue; }
    expect(geometry(candidate), JSON.stringify(entry)).toEqual(geometry(baseline));
    identical += 1;
  }
  expect(matrixCases).toHaveLength(178);
  expect(exceptions).toBe(5);
  if (process.env.DOCX_TAB_DIFFERENTIAL_REPORT) {
    await writeFile(process.env.DOCX_TAB_DIFFERENTIAL_REPORT, JSON.stringify({ matrix: matrixCases.length, generated: generatedCases.length, identical, overflow, exceptions }, null, 2));
  }
}, 180_000);
