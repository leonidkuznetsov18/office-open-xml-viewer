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
  tab?: { alignment: string; count: number; text: string; positional?: boolean; relativeTo?: 'margin' | 'indent'; noFloat?: boolean; fontSizePt?: number; automatic?: boolean; prefix?: string },
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
      ? `<w:ptab w:alignment="${tab.alignment}" w:relativeTo="${tab.relativeTo ?? 'margin'}" w:leader="none"/>`
      : '<w:tab/>'.repeat(tab.count);
    xml = xml.replace('<w:t', `${tab.prefix ? `<w:t>${tab.prefix}</w:t>` : ""}${tabXml}<w:t`)
      .replace('</w:pPr>', `<w:tabs><w:tab w:val="${tab.alignment}" w:pos="5600"/></w:tabs></w:pPr>`);
    // CT_PPrBase requires tabs before bidi/spacing/ind (§A.1).
    xml = xml.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (_match, properties: string) => {
      const tabs = properties.match(/<w:tabs>[\s\S]*?<\/w:tabs>/)?.[0] ?? '';
      return `<w:pPr>${tabs}${properties.replace(tabs, '')}</w:pPr>`;
    });
    // The float's anchor must precede the tab cell, rather than count as
    // trailing content when the aligned tab measures that cell.
    const drawingRun = xml.match(/<w:r>\s*<w:drawing>[\s\S]*?<\/w:drawing>\s*<\/w:r>/)?.[0];
    if (!drawingRun) throw new Error('Missing floating drawing run');
    xml = xml.replace(drawingRun, '').replace('</w:pPr>', `</w:pPr>${drawingRun}`);
  }
  if (tab?.fontSizePt) xml = xml.replace(/<w:sz(?:Cs)? w:val="\d+"\/>/g, (tag) => tag.replace(/\d+/, String((tab.fontSizePt ?? 12) * 2)));
  if (tab?.automatic) xml = xml.replace(/<w:tabs>[\s\S]*?<\/w:tabs>/g, '');
  if (tab?.noFloat) xml = xml.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, '');
  parts.set('word/document.xml', encoder.encode(xml));
  return storeZip(parts);
}

const measureContext = (advancePt = 5) => ({
  font: '', letterSpacing: '0px', fontKerning: 'normal',
  measureText: (text: string) => ({ width: [...text].length * advancePt,
    actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
}) as unknown as CanvasRenderingContext2D;

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function layoutParagraph(bytes: Uint8Array, advancePt = 5) {
  const archive = new DocxArchive(bytes);
  let model;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally { archive.free(); }
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(advancePt) }), { currentDateMs: 0 });
  const paragraph = layout.pages[0]?.layers.body.find((node) => node.kind === 'paragraph');
  if (paragraph?.kind !== 'paragraph') throw new Error('Missing paragraph');
  return paragraph;
}

describe('first-line indents beside floats through the DOCX parser', () => {
  // Independent Word positions for the valid-image 84-case matrix. The
  // control has a 24 pt cell; this deterministic measurer has a 20 pt cell,
  // so leading/trailing/center positions differ by 4/0/2 pt in RTL (0/4/2 LTR).
  // No-float ordinary tabs retain main's geometry. Five no-float RTL positional
  // cases instead use Word: margin-left, both centers, and both right targets.
  // Overlap and ink beyond the margin are intentional measured outcomes.
  const floatStarts = [
    [552, 108, 484, 40], [566, 352, 240, 26],
    [532, 108, 484, 60], [566, 332, 260, 26],
    [542, 108, 484, 50], [566, 342, 250, 26],
    [542, 108, 484, 50], [566, 342, 250, 26],
    [272, 72, 520, 320], [272, 108, 484, 320],
    [396, 196, 396, 196], [396, 214, 378, 196],
    [520, 320, 272, 72], [520, 320, 272, 72],
  ] as const;
  const tabCases = [
    { alignment: 'left', count: 1, ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
    { alignment: 'left', count: 2, ltrX: 352, ltrWidth: 300, rtlX: 240, rtlWidth: 300 },
    { alignment: 'right', count: 1, ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
    { alignment: 'right', count: 2, ltrX: 332, ltrWidth: 280, rtlX: 260, rtlWidth: 280 },
    { alignment: 'center', count: 1, ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
    { alignment: 'center', count: 2, ltrX: 342, ltrWidth: 290, rtlX: 250, rtlWidth: 290 },
    { alignment: 'decimal', count: 1, ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
    { alignment: 'decimal', count: 2, ltrX: 342, ltrWidth: 290, rtlX: 250, rtlWidth: 290 },
    { alignment: 'left', count: 1, relativeTo: 'margin', ltrX: 72, ltrWidth: 20, rtlX: 520, rtlWidth: 20 },
    { alignment: 'left', count: 1, relativeTo: 'indent', ltrX: 108, ltrWidth: 56, rtlX: 484, rtlWidth: 56 },
    { alignment: 'center', count: 1, relativeTo: 'margin', ltrX: 296, ltrWidth: 244, rtlX: 296, rtlWidth: 244 },
    { alignment: 'center', count: 1, relativeTo: 'indent', ltrX: 314, ltrWidth: 262, rtlX: 278, rtlWidth: 262 },
    { alignment: 'right', count: 1, relativeTo: 'margin', ltrX: 520, ltrWidth: 468, rtlX: 72, rtlWidth: 468 },
    { alignment: 'right', count: 1, relativeTo: 'indent', ltrX: 520, ltrWidth: 468, rtlX: 72, rtlWidth: 468 },
  ] as const;
  const matrix = tabCases.flatMap((tab, tabIndex) => [false, true].flatMap((rtl) =>
    (['none', 'left', 'right'] as const).map((float) => ({
      ...tab, tabIndex, rtl, float, kind: 'relativeTo' in tab ? `positional-${tab.relativeTo}` : 'ordinary',
    }))));

  it.each(matrix)('$kind $alignment ($count), rtl=$rtl, float=$float matches Word geometry', (entry) => {
    const relativeTo = 'relativeTo' in entry ? entry.relativeTo : undefined;
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', entry.rtl,
      entry.float === 'right' ? 268 : 0, {
        alignment: entry.alignment, count: entry.count,
        text: entry.alignment === 'decimal' ? '12.3' : 'word',
        positional: relativeTo !== undefined, relativeTo, noFloat: entry.float === 'none',
      }));
    const texts = paragraph.lines.flatMap((line) => line.placements).filter((node) => node.kind === 'text');
    expect(texts).toHaveLength(1);
    if (entry.float === 'none') {
      const xPt = entry.rtl ? entry.rtlX : entry.ltrX;
      const widthPt = entry.rtl ? entry.rtlWidth : entry.ltrWidth;
      expect(paragraph.lines.map((line) => line.bounds)).toEqual([
        { xPt: entry.rtl ? xPt : 72, yPt: 72, widthPt, heightPt: 12 },
      ]);
      expect(texts.map((text) => text.bounds)).toEqual([
        { xPt, yPt: 72, widthPt: 20, heightPt: 12 },
      ]);
    } else {
      const index = (entry.rtl ? 2 : 0) + (entry.float === 'right' ? 1 : 0);
      expect(texts.map((text) => text.bounds)).toEqual([
        { xPt: floatStarts[entry.tabIndex][index], yPt: 72, widthPt: 20, heightPt: 12 },
      ]);
    }
  });

  it.each([false, true])('fits a margin-aligned cell independently of the paragraph right indent, positional=%s', (positional) => {
    const paragraph = layoutParagraph(documentBytes('w:right="7200"', false, 0,
      { alignment: 'right', count: 1, text: 'word '.repeat(10).trim(), positional, noFloat: true }));
    expect(paragraph.lines).toHaveLength(1);
    const texts = paragraph.lines[0].placements.filter((node) => node.kind === 'text');
    const last = texts.at(-1);
    expect(last ? last.bounds.xPt + last.bounds.widthPt : undefined).toBe(positional ? 540 : 352);
  });

  it('wraps content after an automatic tab in a narrowed float window', () => {
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', false, 268,
      { alignment: 'left', count: 1, text: 'word '.repeat(100).trim(), automatic: true }), 6);
    expect(paragraph.lines).toHaveLength(12);
    for (const line of paragraph.lines) {
      for (const placement of line.placements) {
        if (placement.kind === 'text') expect(placement.bounds.xPt + placement.bounds.widthPt).toBeLessThanOrEqual(540);
      }
    }
  });

  it.each(['margin', 'indent'].flatMap((relativeTo) => [false, true].map((long) => ({ relativeTo: relativeTo as 'margin' | 'indent', long }))))('moves an RTL positional tab past its $relativeTo target to the next line, long=$long', ({ relativeTo, long }) => {
    const paragraph = layoutParagraph(documentBytes('', true, 0,
      { alignment: 'left', count: 1, text: long ? 'word '.repeat(100).trim() : 'word', positional: true, relativeTo, noFloat: true, prefix: 'prefix' }), 6);
    const text = paragraph.lines.flatMap((line) => line.placements).find((node) => node.kind === 'text' && node.text.startsWith('word'));
    expect(text?.kind === 'text' ? text.bounds.yPt : undefined).toBe(84);
    if (long) expect(paragraph.lines.length).toBeGreaterThan(2);
    else expect(paragraph.lines).toHaveLength(2);
  });

  it.each(matrix)('long $kind $alignment ($count), rtl=$rtl, float=$float retains legal line breaks', (entry) => {
    const relativeTo = 'relativeTo' in entry ? entry.relativeTo : undefined;
    const content = (entry.alignment === 'decimal' ? '12.3 ' : 'word ').repeat(100).trim();
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', entry.rtl,
      entry.float === 'right' ? 268 : 0, {
        alignment: entry.alignment, count: entry.count, text: content,
        positional: relativeTo !== undefined, relativeTo, noFloat: entry.float === 'none',
      }));
    expect(paragraph.lines.length).toBeGreaterThan(1);
    const textLines = paragraph.lines.map((line) => line.placements.filter((node) => node.kind === 'text'));
    expect(textLines.flat().map((text) => text.text).join('').replace(/\s/g, ''))
      .toBe(content.replace(/\s/g, ''));
    // A displaced tab may itself pass the margin, as the short Word controls
    // demonstrate. That does not grant its entire following cell infinite room.
    for (const texts of textLines) {
      expect(texts.reduce((total, text) => total + text.bounds.widthPt, 0)).toBeLessThanOrEqual(468);
    }
    for (const texts of textLines.slice(1)) {
      for (const text of texts) {
        expect(text.bounds.xPt).toBeGreaterThanOrEqual(72);
        expect(text.bounds.xPt + text.bounds.widthPt).toBeLessThanOrEqual(540);
      }
    }
  });

  it.each([10, 20])('keeps the measured overflow tab advance at %s pt font size', (fontSizePt) => {
    const paragraph = layoutParagraph(documentBytes('w:left="720" w:hanging="720"', false, 0,
      { alignment: 'left', count: 2, text: 'word', fontSizePt }), fontSizePt * 0.6);
    const text = paragraph.lines.flatMap((line) => line.placements).find((node) => node.kind === 'text');
    expect(paragraph.lines).toHaveLength(1);
    expect(text?.bounds.xPt).toBe(566);
    expect(text?.bounds.yPt).toBe(72);
    expect(text?.bounds.widthPt).toBe(fontSizePt * 2.4);
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
