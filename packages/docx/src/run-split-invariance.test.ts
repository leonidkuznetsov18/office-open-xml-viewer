import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService } from './layout/text.js';
import { readFile } from 'node:fs/promises';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';
import { acquireShapeTextBoxLayout } from './layout/paragraph.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import type { ShapeRun } from './types.js';
import { beforeAll, describe, expect, it } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunsForPage } from './text-run-projection.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, DocxTextRun, SectionProps } from './types.js';
import type { DocumentLayout, ParagraphLayout } from './layout/types.js';

const borders = { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
function run(text: string, extra: Partial<DocxTextRun> = {}) {
  return { type: 'text' as const, text, fontSize: 18, fontFamily: 'Arial', kerning: 8,
    bold: false, italic: false, underline: false, strikethrough: false, color: null,
    isLink: false, background: null, vertAlign: null, hyperlink: null, ...extra };
}
function paragraph(parts: string[], extra: Partial<DocxTextRun>, alignment: DocParagraph['alignment']): DocParagraph {
  return { runs: parts.map(text => run(text, extra)), alignment, indentLeft: 0, indentRight: 0,
    indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops: [], defaultFontSize: extra.fontSize ?? 18, defaultFontFamily: extra.fontFamily ?? 'Arial',
    widowControl: false };
}
function model(p: DocParagraph, width: number, container: 'paragraph' | 'fixed' | 'autofit'): DocxDocumentModel {
  const body: BodyElement[] = [{ type: 'paragraph', ...p }];
  if (container !== 'paragraph') body.splice(0, 1, { type: 'table', layout: container,
    colWidths: [width], widthPt: width, borders, cellMarginTop: 0, cellMarginRight: 0,
    cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left', rows: [{ cells: [{
      content: [{ type: 'paragraph', ...p }], colSpan: 1, vMerge: null, borders,
      background: null, vAlign: 'top', widthPt: width, widthPct: undefined,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    }], rowHeight: null, rowHeightRule: 'auto', isHeader: false }] });
  return { body, settings: { compatibilityMode: 15 }, section: { pageWidth: width,
    pageHeight: 1000, marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false } as SectionProps,
    headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
    fontFamilyClasses: {}, footnotes: [] } as unknown as DocxDocumentModel;
}
function context(): CanvasRenderingContext2D {
  // Independent pair metrics, including a pair spanning an ordinary space.
  // The Arial T/i scalars and T-space pair reproduce the reviewer's 10.8pt cell.
  const ctx = { font: '18px Arial', fontKerning: 'auto', letterSpacing: '0px',
    measureText(text: string) {
      const size = Number(/([\d.]+)px/u.exec(ctx.font)?.[1] ?? 18);
      const scalar = (c: string) => c === 'T' ? 10.9951171875 : c === ' ' ? 5.0009765625
        : c === 'i' ? 3.9990234375 : 10;
      const pair = ctx.fontKerning === 'normal'
        ? (text.match(/T /gu)?.length ?? 0) * .3251953125 + (text.match(/AV/gu)?.length ?? 0) * 2 : 0;
      const width = ([...text].reduce((sum, c) => sum + scalar(c), 0) - pair) * size / 18;
      return { width, fontBoundingBoxAscent: size * .8, fontBoundingBoxDescent: size * .2,
        actualBoundingBoxLeft: 0, actualBoundingBoxRight: width,
        actualBoundingBoxAscent: size * .8, actualBoundingBoxDescent: size * .2 } as TextMetrics;
    } };
  return ctx as unknown as CanvasRenderingContext2D;
}
function acquire(parts: string[], width: number, container: 'paragraph' | 'fixed' | 'autofit',
  extra: Partial<DocxTextRun> = {}, alignment: DocParagraph['alignment'] = 'right') {
  const doc = model(paragraph(parts, extra, alignment), width, container);
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
}
function intrinsic(runs: DocParagraph['runs']) {
  const p = { ...paragraph([], {}, 'right'), runs };
  const doc = model(p, 100, 'paragraph');
  const measure = context();
  return measureParagraphIntrinsicWidths(p, contextPt, 100,
    { context: measure, fontFamilyClasses: {} }, {
      pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
      documentHasEastAsianText: false, compatibilityMode: 15,
      layoutServices: createLayoutServices(doc, { measureContext: measure }),
    });
}
function paragraphs(layout: DocumentLayout): ParagraphLayout[] {
  const result: ParagraphLayout[] = [];
  const seen = new WeakSet<object>();
  const visit = (value: unknown): void => {
    if (!value || typeof value !== 'object' || seen.has(value)) return;
    seen.add(value);
    if ('kind' in value && value.kind === 'paragraph') {
      const p = value as ParagraphLayout;
      result.push(p);
      p.textBoxes.forEach(visit);
      return;
    }
    for (const [key, child] of Object.entries(value)) if (key !== 'source') {
      if (Array.isArray(child)) child.forEach(visit); else visit(child);
    }
  };
  layout.pages.forEach(visit);
  return result;
}
function geometry(layout: DocumentLayout | readonly ParagraphLayout[]) {
  return (Array.isArray(layout) ? layout as ParagraphLayout[] : paragraphs(layout as DocumentLayout)).map(p => ({ bounds: p.flowBounds, lines: p.lines.map(l => ({
    range: l.range, bounds: l.bounds, advance: l.advancePt, baseline: l.baselinePt,
    text: l.placements.flatMap(s => s.kind === 'text' ? [s.text] : []).join(''),
    glyphs: l.placements.flatMap(s => s.kind === 'text' ? s.clusters.map(c => ({
      range: c.range, x: s.origin.xPt + c.offset.xPt, advance: c.advancePt,
    })) : []),
    paint: l.placements.flatMap(s => s.kind === 'text' ? s.paintOps.map(op => ({
      text: op.text, x: s.origin.xPt + op.offset.xPt, y: s.origin.yPt + op.offset.yPt,
      range: op.range, kerning: op.kerning,
    })) : []),
  })) }));
}
function partitions(text: string): string[][] {
  const result = [[text], [...text]];
  for (let i = 1; i < text.length; i++) result.push([text.slice(0, i), text.slice(i)]);
  // Stable pseudo-random partitions: reproducible failures and no fixture matrix.
  let state = 1703;
  for (let n = 0; n < 8; n++) {
    const parts: string[] = []; let start = 0;
    for (let i = 1; i < text.length; i++) {
      state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
      if (state % 3 === 0) { parts.push(text.slice(start, i)); start = i; }
    }
    parts.push(text.slice(start)); result.push(parts);
  }
  return [...new Map(result.map(parts => [JSON.stringify(parts), parts])).values()];
}

describe('formatting-only run boundaries are transparent to text acquisition', () => {
  it.each(['paragraph', 'fixed'] as const)('rejects the narrow T-space prefix consistently in a %s', container => {
    const expected = geometry(acquire(['T i'], 10.8, container));
    expect(expected[0]?.lines.map(l => l.text)).toEqual(['T', ' i']);
    expect(expected[0]?.lines[0]?.glyphs[0]?.x).toBeCloseTo(-.1951171875, 8);
    for (const parts of partitions('T i')) expect(geometry(acquire(parts, 10.8, container)), parts.join('|')).toEqual(expected);
    const accepted = geometry(acquire(['T i'], 16, container));
    expect(accepted[0]?.lines.map(l => l.text)).toEqual(['T ', 'i']);
    for (const parts of partitions('T i')) expect(geometry(acquire(parts, 16, container))).toEqual(accepted);
  });

  it.each([
    { text: 'AVAT i-AV T i', fontFamily: 'Arial', fontSize: 18, kerning: 8 },
    { text: ' T i AV-T  i ', fontFamily: 'Times New Roman', fontSize: 10, kerning: 10 },
    { text: 'AV T i\tAV-T i', fontFamily: 'Georgia', fontSize: 24, kerning: 24.5 },
    { text: 'AV-T i T  i', fontFamily: 'Arial', fontSize: 18, kerning: 0 },
  ])('preserves widths, wraps, alignment and retained paint for $fontFamily / $kerning', ({ text, ...format }) => {
    const widths = intrinsic([run(text, format)]);
    for (const parts of partitions(text)) expect(intrinsic(parts.map(text => run(text, format)))).toEqual(widths);
    for (const container of ['paragraph', 'fixed', 'autofit'] as const) {
      for (const alignment of ['left', 'right', 'center', 'both'] as const) {
        const expected = geometry(acquire([text], 35, container, format, alignment));
        for (const parts of partitions(text)) {
          expect(geometry(acquire(parts, 35, container, format, alignment)), `${container}/${alignment}/${parts.join('|')}`).toEqual(expected);
        }
      }
    }
  });

  it('keeps original source ownership after shaping a sequence', () => {
    const layout = acquire(['A', 'V T', ' i'], 100, 'paragraph');
    const runs = textRunsForPage(layout, 0, { scale: 1 });
    expect(runs.map(r => r.text).join('')).toBe('AV T i');
    for (const [index, text] of ['A', 'V T', ' i'].entries()) {
      expect(runs.filter(r => r.sourceRunIndex === index).map(r => r.text).join('')).toBe(text);
    }
  });
});


beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

function parsedRuns(parts: string[], wrapper: string): DocParagraph['runs'] {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const prop = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="36"/><w:kern w:val="16"/></w:rPr>';
  const runs = parts.map((text, index) => `<w:r w:rsidR="0000000${index % 8}">${prop}<w:t xml:space="preserve">${text}</w:t></w:r>`);
  const wrap = (run: string, index: number) => {
    if (wrapper === 'smart-tags') return `<w:smartTag w:uri="urn:test" w:element="word">${run}</w:smartTag>`;
    if (wrapper === 'revisions') return `<w:ins w:id="${index}" w:author="Reviewer">${run}</w:ins>`;
    if (wrapper === 'hyperlinks') return `<w:hyperlink w:anchor="Destination">${run}</w:hyperlink>`;
    if (wrapper === 'simple-fields') return `<w:fldSimple w:instr="AUTHOR">${run}</w:fldSimple>`;
    if (wrapper === 'bookmarks') return `<w:bookmarkStart w:id="${index}" w:name="B${index}"/>${run}<w:bookmarkEnd w:id="${index}"/>`;
    if (wrapper === 'proofing') return `<w:proofErr w:type="spellStart"/>${run}<w:proofErr w:type="spellEnd"/>`;
    if (wrapper === 'complex-fields') return `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>AUTHOR</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${run}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    return run;
  };
  const wrapped = runs.map(wrap).join('');
  const content = wrapper === 'textbox'
    ? `<w:r><w:pict><v:shape id="box" type="#_x0000_t202" style="width:10.8pt;height:100pt"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:pPr><w:jc w:val="right"/></w:pPr>${wrapped}</w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>`
    : wrapped;
  const archive = new DocxArchive(storeZip(new Map([
    ['[Content_Types].xml', new TextEncoder().encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', new TextEncoder().encode(`<w:document xmlns:w="${W}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body><w:p>${content}</w:p><w:sectPr/></w:body></w:document>`)],
  ])));
  try {
    const doc = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
    const p = doc.body.find(b => b.type === 'paragraph');
    if (!p || p.type !== 'paragraph') throw new Error('Expected parsed paragraph');
    return p.runs;
  } finally { archive.free(); }
}

function acquireRuns(runs: DocParagraph['runs'], container: 'paragraph' | 'fixed' | 'autofit') {
  const p = { ...paragraph([], {}, 'right'), runs };
  const doc = model(p, 10.8, container);
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
}

describe('parsed structural splits feed the same production text sequence', () => {
  it.each(['rsid', 'smart-tags', 'revisions', 'hyperlinks', 'simple-fields', 'complex-fields', 'bookmarks', 'proofing', 'textbox'])(
    '%s preserves glyph geometry across split result runs', wrapper => {
      const whole = parsedRuns(['T i'], wrapper);
      if (wrapper === 'textbox') {
        const boxes = geometry(acquireRuns(whole, 'paragraph'));
        expect(boxes).toHaveLength(2);
        expect(boxes[1]?.lines.map(line => line.text)).toEqual(['T', ' i']);
      }
      for (const container of ['paragraph', 'fixed', 'autofit'] as const) {
        const expected = geometry(acquireRuns(whole, container));
        for (const parts of partitions('T i')) {
          expect(geometry(acquireRuns(parsedRuns(parts, wrapper), container)), `${container}/${parts.join('|')}`).toEqual(expected);
        }
      }
    },
  );
});

const contextPt: ParagraphLayoutContext = {
    lineGrid: { active: false, pitchPt: null },
    characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
    rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
    physicalIndentLeftPt: 0, physicalIndentRightPt: 0, firstIndentPt: 0,
    lineSpacing: null, spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: false,
    isJustified: false, stretchLastLine: false, tabStops: [], hasRuby: false, hasEastAsianText: false,
    kinsoku: { enabled: true, lineStartForbidden: new Set(), lineEndForbidden: new Set() }, defaultTabPt: 36,
  };

it('uses the same sequence in the shape text-box adapter', () => {

  const box = (parts: string[]) => {
    const doc = model(paragraph([], {}, 'right'), 100, 'paragraph');
    const measure = context();
    const layout = acquireShapeTextBoxLayout({
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textBlocks: [{ text: parts.join(''), fontSizePt: 18, alignment: 'right',
        runs: parts.map(text => ({ text, fontSizePt: 18, fontFamily: 'Arial' })) }],
    } as ShapeRun, { xPt: 0, yPt: 0, widthPt: 10.8, heightPt: 100 }, {
      id: 'box', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      flowDomainId: 'body', context: contextPt, measurer: { context: measure, fontFamilyClasses: {} },
      environment: { pageIndex: 0, totalPages: 1, documentHasEastAsianText: false,
        pageWritingMode: 'horizontal-tb', compatibilityMode: 15,
        layoutServices: createLayoutServices(doc, { measureContext: measure }) },
    });
    if (!layout) throw new Error('Expected text box');
    return geometry(layout.story.blocks.filter((b): b is ParagraphLayout => b.kind === 'paragraph'));
  };
  const expected = box(['T i']);
  for (const parts of partitions('T i')) expect(box(parts)).toEqual(expected);
});


it.each(['font', 'weight', 'threshold'] as const)('retains a real %s boundary in both narrow-cell and intrinsic paths', boundary => {
  const runs = [run('T'), run(' i', boundary === 'font' ? { fontFamily: 'Georgia' }
    : boundary === 'weight' ? { bold: true } : { kerning: 18.5 })];
  const fixed = acquireRuns(runs, 'fixed');
  expect(geometry(fixed)[0]?.lines[0]?.text).toBe('T');
  expect(geometry(fixed)[0]?.lines[0]?.glyphs[0]?.advance).toBeCloseTo(10.9951171875, 8);
  expect(intrinsic(runs)).toEqual({ minWidthPt: 10.9951171875, maxWidthPt: 19.9951171875 });
});

it('keeps source-addressed RTL geometry and later run ranges after sequence shaping', () => {
  const p = paragraph(['A', 'V T', ' i'], { rtl: true }, 'right');
  p.runs.push(run(' final', { bold: true, rtl: true }));
  const doc = model(p, 200, 'paragraph');
  const layout = layoutDocument(doc, createLayoutServices(doc, { measureContext: context() }));
  const runs = textRunsForPage(layout, 0, { scale: 1 });
  for (const text of ['A', 'V T', ' i', ' final']) {
    const index = ['A', 'V T', ' i', ' final'].indexOf(text);
    const owned = runs.filter(r => r.sourceRunIndex === index);
    expect(owned.map(r => r.text).join('')).toBe(text);
    expect(owned.every(r => r.w > 0)).toBe(true);
  }
  expect(paragraphs(layout)[0]?.lines[0]?.range.end).toBe('AV T i final'.length);
});


it('preserves independently resolved substitute faces at an otherwise identical source seam', () => {
  const p = paragraph(['مرحبا', '12'], { fontFamily: 'Scoped Face' }, 'left');
  const doc = model(p, 200, 'paragraph');
  const fonts = createFontResolver([
    { requestedFamily: 'Scoped Face', resolvedFamily: 'Arabic Substitute', source: 'substitute', script: 'arabic' },
  ], { scriptScopedFamilies: { 'scoped face': { script: 'arabic', substituteFamilies: ['Arabic Substitute'] } } });
  const services = createLayoutServices(doc, { measureContext: context() });
  const text = createTextLayoutService({ fonts, measurer: { fingerprint: 'split-script-proof',
    measure: request => ({ advancePt: [...request.text].length * 10, ascentPt: 8, descentPt: 2 }),
  } });
  const layout = layoutDocument(doc, { ...services, text });
  const runs = paragraphs(layout)[0]?.lines[0]?.placements.filter(s => s.kind === 'text') ?? [];
  const digits = runs.find(s => s.kind === 'text' && s.text === '12');
  const arabic = runs.find(s => s.kind === 'text' && s.text === 'مرحبا');
  expect(arabic?.fontRoute.familyList).toContain('Arabic Substitute');
  expect(digits?.fontRoute.familyList).not.toContain('Arabic Substitute');
  expect(digits?.sourceRunIndex).toBe(1);
});
