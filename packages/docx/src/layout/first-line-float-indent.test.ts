/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { CONFORMANCE_CASES } from '../conformance/cases.js';
import { generateConformanceParts, storeZip } from '../conformance/generate.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';

function documentBytes(
  indent: string,
  rtl: boolean,
  xPt: number,
  tab?: { alignment: string; count: number; text: string; positional?: boolean },
): Uint8Array {
  const seed = CONFORMANCE_CASES.find(({ axes }) => axes.story === 'body'
    && axes.container === 'paragraph' && axes.object === 'floating');
  if (!seed) throw new Error('Missing floating conformance seed');
  const parts = new Map(generateConformanceParts({
    ...seed,
    expected: { ...seed.expected, targetText: tab?.text ?? 'word '.repeat(30).trim() },
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
  if (tab) {
    const tabXml = tab.positional
      ? `<w:ptab w:alignment="${tab.alignment}" w:relativeTo="margin" w:leader="none"/>`
      : '<w:tab/>'.repeat(tab.count);
    xml = xml.replace('<w:t', `${tabXml}<w:t`)
      .replace('</w:pPr>', `<w:tabs><w:tab w:val="${tab.alignment}" w:pos="5600"/></w:tabs></w:pPr>`);
    // The float's anchor must precede the tab cell, rather than count as
    // trailing content when the aligned tab measures that cell.
    const drawingRun = xml.match(/<w:r>\s*<w:drawing>[\s\S]*?<\/w:drawing>\s*<\/w:r>/)?.[0];
    if (!drawingRun) throw new Error('Missing floating drawing run');
    xml = xml.replace(drawingRun, '').replace('</w:pPr>', `</w:pPr>${drawingRun}`);
  }
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

function layoutParagraph(bytes: Uint8Array) {
  const archive = new DocxArchive(bytes);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext() }), { currentDateMs: 0 });
  const paragraph = layout.pages[0]?.layers.body.find((node) => node.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
  return paragraph;
}

describe('first-line indents beside floats through the DOCX parser', () => {
  it.each([
    // First tab reaches the custom stop at margin + 280; the second reaches
    // the next automatic stop at margin + 288 (the default interval is 36).
    { alignment: 'right', count: 2, text: 'word', expectedStart: 360 },
    { alignment: 'left', count: 1, text: 'word', expectedStart: 352 },
    { alignment: 'right', count: 1, text: 'word', expectedStart: 332 },
    { alignment: 'center', count: 1, text: 'word', expectedStart: 342 },
    { alignment: 'decimal', count: 1, text: '12.3', expectedStart: 342 },
    { alignment: 'center', count: 1, text: 'word', expectedStart: 296, positional: true },
    { alignment: 'right', count: 1, text: 'word', expectedStart: 520, positional: true },
  ])('resolves LTR $alignment tabs ($count, positional=$positional) from the float window', (tab) => {
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', false, 0, tab));
    const text = paragraph.lines[0]?.placements.find((node) => node.kind === 'text');
    if (text?.kind !== 'text') throw new Error('Missing tabbed text');
    expect(text.bounds.xPt).toBeCloseTo(tab.expectedStart, 8);
    expect(text.bounds.xPt).toBeGreaterThanOrEqual(272);
    expect(text.bounds.xPt + text.bounds.widthPt).toBeLessThanOrEqual(540);
  });

  it.each([
    { alignment: 'right', count: 2, text: 'word', expectedStart: 232 },
    { alignment: 'left', count: 1, text: 'word', expectedStart: 240 },
    { alignment: 'right', count: 1, text: 'word', expectedStart: 260 },
    { alignment: 'center', count: 1, text: 'word', expectedStart: 250 },
    { alignment: 'decimal', count: 1, text: '12.3', expectedStart: 250 },
  ])('resolves RTL $alignment tabs ($count) from the mirrored float window', (tab) => {
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', true, 268, tab));
    const text = paragraph.lines[0]?.placements.find((node) => node.kind === 'text');
    if (text?.kind !== 'text') throw new Error('Missing tabbed text');
    expect(text.bounds.xPt).toBeCloseTo(tab.expectedStart, 8);
    expect(text.bounds.xPt).toBeGreaterThanOrEqual(72);
    expect(text.bounds.xPt + text.bounds.widthPt).toBeLessThanOrEqual(340);
  });

  it.each([
    { name: 'hanging beside a left float', indent: 'w:left="720" w:hanging="720"', rtl: false, floatX: 0, expectedStart: 272, continuationStart: 272 },
    { name: 'positive first-line indent beside a left float', indent: 'w:firstLine="720"', rtl: false, floatX: 0, expectedStart: 308, continuationStart: 272 },
    { name: 'hanging beyond a nonblocking left float', indent: 'w:left="4800" w:hanging="360"', rtl: false, floatX: 0, expectedStart: 294, continuationStart: 312 },
    { name: 'hanging beside a right float in RTL', indent: 'w:left="720" w:hanging="720"', rtl: true, floatX: 268, expectedEnd: 340 },
  ])('$name', ({ indent, rtl, floatX, expectedStart, expectedEnd, continuationStart }) => {
    const paragraph = layoutParagraph(documentBytes(indent, rtl, floatX));
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
