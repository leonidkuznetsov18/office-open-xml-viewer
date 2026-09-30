/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { CONFORMANCE_CASES } from '../conformance/cases.js';
import { generateConformanceParts, storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';

function documentBytes(indent: string, rtl: boolean, xPt: number): Uint8Array {
  const seed = CONFORMANCE_CASES.find(({ axes }) => axes.story === 'body'
    && axes.container === 'paragraph' && axes.object === 'floating');
  if (!seed) throw new Error('Missing floating conformance seed');
  const parts = new Map(generateConformanceParts({
    ...seed,
    expected: { ...seed.expected, targetText: 'word '.repeat(30).trim() },
    axes: { ...seed.axes, direction: rtl ? 'rtl' : 'ltr', paragraph: 'single',
      styleSource: 'direct', fontSource: 'direct', spacing: 'exact', anchorReference: 'margin' },
  }));
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let xml = decoder.decode(parts.get('word/document.xml'));
  xml = xml.replace('</w:pPr>', `<w:ind ${indent}/></w:pPr>`)
    .replace('<wp:posOffset>914400</wp:posOffset>', `<wp:posOffset>${xPt * 12700}</wp:posOffset>`)
    .replaceAll('cx="457200"', 'cx="2540000"')
    .replaceAll('cy="274320"', 'cy="1270000"');
  parts.set('word/document.xml', encoder.encode(xml));
  return storeZip(parts);
}

const measureContext = () => ({
  font: '', letterSpacing: '0px', fontKerning: 'normal',
  measureText: (text: string) => ({ width: [...text].length * 5,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
}) as unknown as CanvasRenderingContext2D;

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('first-line indents beside floats through the DOCX parser', () => {
  it.each([
    { name: 'hanging beside a left float', indent: 'w:left="720" w:hanging="720"', rtl: false, floatX: 0, expectedStart: 272, continuationStart: 272 },
    { name: 'positive first-line indent beside a left float', indent: 'w:firstLine="720"', rtl: false, floatX: 0, expectedStart: 308, continuationStart: 272 },
    { name: 'hanging beyond a nonblocking left float', indent: 'w:left="4800" w:hanging="360"', rtl: false, floatX: 0, expectedStart: 294, continuationStart: 312 },
    { name: 'hanging beside a right float in RTL', indent: 'w:left="720" w:hanging="720"', rtl: true, floatX: 268, expectedEnd: 340 },
  ])('$name', ({ indent, rtl, floatX, expectedStart, expectedEnd, continuationStart }) => {
    const archive = new DocxArchive(documentBytes(indent, rtl, floatX));
    let model;
    try {
      model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
    } finally { archive.free(); }
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), { currentDateMs: 0 });
    const paragraph = layout.pages[0]?.layers.body.find((node) => node.kind === 'paragraph');
    if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
    const line = paragraph.lines[0];
    expect(line?.bounds.yPt).toBe(72);
    // A correction only to placement would leave the breaker's old, widened
    // hanging-line budget and let long text cross the opposite window edge.
    expect(paragraph.lines.length).toBeGreaterThan(1);
    expect(line?.bounds.widthPt).toBeLessThanOrEqual(268);
    if (expectedStart !== undefined) expect(line?.bounds.xPt).toBeCloseTo(expectedStart, 8);
    if (expectedEnd !== undefined) expect((line?.bounds.xPt ?? 0) + (line?.bounds.widthPt ?? 0)).toBeCloseTo(expectedEnd, 8);
    if (continuationStart !== undefined) expect(paragraph.lines[1]?.bounds.xPt).toBeCloseTo(continuationStart, 8);
  });
});
