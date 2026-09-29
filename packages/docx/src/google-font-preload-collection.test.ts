/// <reference types="node" />
import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { createLayoutServices } from './layout-runtime.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { renderDocumentToCanvas } from './renderer.js';
import { docxFontPreloadNames } from './google-fonts.js';
import type { DocxDocumentModel } from './types.js';

// useGoogleFonts substitutes (Calibri → Carlito, Cambria → Caladea) are fetched
// only for names in docxFontPreloadNames, and the layout font inventory routes
// only those names to a loaded substitute. Each case parses a synthetic DOCX
// with NO theme part (so theme names cannot mask the gap) and asserts that the
// family production shaping requests is both collected and painted with the
// loaded substitute.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const encoder = new TextEncoder();

function docx(body: string, numbering?: string): Uint8Array {
  const parts = new Map<string, Uint8Array>([
    ['[Content_Types].xml', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + (numbering ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : '')
      + '</Types>')],
    ['_rels/.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
      + '</Relationships>')],
    ['word/_rels/document.xml.rels', encoder.encode('<?xml version="1.0" encoding="UTF-8"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + (numbering ? '<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' : '')
      + '</Relationships>')],
    ['word/document.xml', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:document xmlns:w="${W}" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>${body}`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr></w:body></w:document>')],
  ]);
  if (numbering) {
    parts.set('word/numbering.xml', encoder.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">${numbering}</w:abstractNum>`
      + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'));
  }
  return storeZip(parts);
}

const arialRun = (text: string) =>
  `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/></w:rPr><w:t>${text}</w:t></w:r>`;
const numberedParagraph = `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${arialRun('Item')}</w:p>`;

function parse(bytes: Uint8Array): DocxDocumentModel {
  const archive = new DocxArchive(bytes);
  try {
    return normalizeInternalDocumentModel(
      JSON.parse(new TextDecoder().decode(archive.parse())),
    ).document;
  } finally {
    archive.free();
  }
}

function recordingCanvas() {
  let font = '10px serif';
  const calls: { text: string; font: string }[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const noop = () => {};
  const ctx = new Proxy({
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * px() * 0.5,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
    }) as TextMetrics,
    fillText(text: string) { calls.push({ text, font }); },
    createLinearGradient: () => ({ addColorStop: noop }),
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
  } as Record<string | symbol, unknown>, {
    get(target, key) {
      if (key in target) return target[key];
      return noop;
    },
    set(target, key, value) {
      target[key] = value;
      return true;
    },
  });
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

const loaded = (family: string) =>
  ({ family, weight: '400', style: 'normal', status: 'loaded' }) as FontFace;

async function paintedFamily(model: DocxDocumentModel, text: string): Promise<string | undefined> {
  const { canvas, calls } = recordingCanvas();
  await renderDocumentToCanvas(model, canvas, 0, {
    dpr: 1,
    width: 612,
    layoutServices: createLayoutServices(model, {
      useGoogleFonts: true,
      googleFaces: [loaded('Carlito'), loaded('Caladea')],
      measureContext: canvas.getContext('2d') as CanvasRenderingContext2D,
    }),
  });
  const call = calls.find((entry) => entry.text.includes(text));
  return call && /px\s+"?([^",]+)"?/.exec(call.font)?.[1];
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('Google Fonts preload collects every rendered substitute family', () => {
  it('collects and paints a directly authored Cambria body run with Caladea', async () => {
    const model = parse(docx(
      '<w:p><w:r><w:rPr><w:rFonts w:ascii="Cambria" w:hAnsi="Cambria"/></w:rPr><w:t>SerifBody</w:t></w:r></w:p>',
    ));
    expect(model.majorFont ?? null).toBeNull();
    expect(docxFontPreloadNames(model)).toContain('Cambria');
    expect(docxFontPreloadNames(model)).not.toContain('Calibri');
    expect(await paintedFamily(model, 'SerifBody')).toBe('Caladea');
  });

  it('collects Calibri from a table nested in a text box story', async () => {
    const model = parse(docx(
      `<w:p>${arialRun('Anchor')}<w:r><w:pict>`
      + '<v:shape id="box" type="#_x0000_t202" style="position:relative;width:300pt;height:120pt" filled="f" stroked="f">'
      + '<v:textbox><w:txbxContent>'
      + '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
      + '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
      + '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr><w:t>CellText</w:t></w:r></w:p>'
      + '</w:tc></w:tr></w:tbl><w:p/>'
      + '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>',
    ));
    expect(docxFontPreloadNames(model)).toContain('Calibri');
    expect(await paintedFamily(model, 'CellText')).toBe('Carlito');
  });

  it('collects the highAnsi Calibri slot of a bullet whose ascii slot is Arial', async () => {
    const model = parse(docx(numberedParagraph, '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
      + '<w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:suff w:val="space"/>'
      + '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Calibri"/></w:rPr></w:lvl>'));
    expect(docxFontPreloadNames(model)).toContain('Calibri');
    expect(await paintedFamily(model, '•')).toBe('Carlito');
  });

  it('collects the complex-script Calibri slot of a cs number whose ascii slot is Arial', async () => {
    const model = parse(docx(numberedParagraph, '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
      + '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="space"/>'
      + '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Calibri"/><w:cs/></w:rPr></w:lvl>'));
    expect(docxFontPreloadNames(model)).toContain('Calibri');
    expect(await paintedFamily(model, '1.')).toBe('Carlito');
  });
});
