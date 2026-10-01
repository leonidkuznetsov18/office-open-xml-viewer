/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

// Synthetic equivalent of a private journal footer: an A4 page with 936 twip
// margins (501.7 pt text width), a paragraph right indent of 8108 twips
// (405.4 pt, leaving a 96.3 pt indent band) and a right stop at 9923 twips
// (496.15 pt), i.e. past the indent but inside the text margin.
function footerDocument(): Uint8Array {
  const footer = `<w:ftr xmlns:w="${W}"><w:p><w:pPr>`
    + '<w:tabs><w:tab w:val="right" w:pos="9923"/></w:tabs><w:ind w:right="8108"/>'
    + '<w:rPr><w:sz w:val="16"/></w:rPr></w:pPr>'
    + '<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>DOI: 10.2478/x</w:t></w:r>'
    + '<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:tab/></w:r>'
    + '<w:r><w:rPr><w:i/><w:sz w:val="16"/></w:rPr><w:t xml:space="preserve"> *Corresponding author: a@b.edu (F. Surname) </w:t></w:r>'
    + '<w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:tab/></w:r>'
    + '</w:p></w:ftr>';
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>`
    + '<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/>'
    + '<w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1418" w:right="936" w:bottom="1418" w:left="936" w:header="431" w:footer="431" w:gutter="0"/>'
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
      + '</Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/_rels/document.xml.rels', encode(`<Relationships xmlns="${PKG}"><Relationship Id="rIdFooter" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>')],
    ['word/document.xml', encode(document)],
    ['word/footer1.xml', encode(footer)],
  ]));
}

describe('footer right tab past the paragraph right indent', () => {
  // ECMA-376 §17.3.1.37 positions custom stops relative to the page margins;
  // the right indent does not bound them. Word keeps such a footer's aligned
  // cell on the first line; containing it in the 96.3 pt indent band instead
  // grew the footer from two lines to five and pushed body text to later pages.
  it('aligns the cell at its margin-relative stop and keeps two lines', () => {
    const archive = new DocxArchive(footerDocument());
    let model;
    try {
      model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
    } finally { archive.free(); }
    const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(4) }), { currentDateMs: 0 });
    const paragraph = layout.pages[0]?.layers.footer.find((node) => node.kind === 'paragraph');
    if (paragraph?.kind !== 'paragraph') throw new Error('Missing footer paragraph');
    // Line 1: DOI, the right tab and its cell; line 2: the trailing tab.
    expect(paragraph.lines).toHaveLength(2);
    const texts = paragraph.lines[0].placements.filter((node) => node.kind === 'text');
    expect(texts.map((text) => text.text).join('').replace(/\s+/g, ' ').trim())
      .toBe('DOI: 10.2478/x *Corresponding author: a@b.edu (F. Surname)');
    const cell = texts.filter((text) => text.text.trim().length > 0).at(-1)!;
    // The stop is at 46.8 + 496.15 pt; the right-aligned cell (with its 4 pt
    // trailing space) ends there, outside the 46.8 + 96.3 pt indent band.
    expect(cell.text.endsWith(') ')).toBe(true);
    expect(cell.bounds.xPt + cell.bounds.widthPt).toBeCloseTo(46.8 + 496.15, 6);
  });
});
