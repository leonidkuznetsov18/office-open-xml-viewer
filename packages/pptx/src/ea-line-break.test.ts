import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph, renderTable, type PptxTextRunInfo } from './renderer.js';
import type { Paragraph, TableCell, TableElement, TextBody } from './types.js';

// ECMA-376 §21.1.2.2.7 a:pPr@eaLnBrk: false forbids breaking an East Asian
// word; the word moves whole to the next line and overflows when wider than
// the box. Every glyph measures 10 px in these probes.
function measuringContext(): CanvasRenderingContext2D {
  let font = '';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({
      width: [...text].length * 10,
      actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2,
      fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2,
    } as TextMetrics),
    canvas: { width: 1000, height: 1000 },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setTransform() {}, transform() {},
    setLineDash() {}, getLineDash() { return []; }, drawImage() {},
    fillText() {}, strokeText() {},
    fillStyle: '', strokeStyle: '', lineWidth: 1, globalAlpha: 1,
    textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr', letterSpacing: '0px',
  } as unknown as CanvasRenderingContext2D;
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000',
    fontFamily: 'Arial', fontFamilyEa: 'Meiryo',
  };
}

function paragraph(runs: TextRunData[], eaLnBrk: boolean): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [],
    eaLnBrk, runs,
  } as Paragraph;
}

function lines(runs: TextRunData[], width: number, eaLnBrk: boolean): string[] {
  return layoutParagraph(measuringContext(), paragraph(runs, eaLnBrk), width, 20, '000000', 1, 0)
    .map((line) => line.segments.map((segment) => segment.text).join(''));
}

describe('pptx eaLnBrk (§21.1.2.2.7)', () => {
  it('keeps an overwide East Asian word whole when eaLnBrk is false', () => {
    expect(lines([run('日本語')], 10, false)).toEqual(['日本語']);
    expect(lines([run('日本語')], 10, true)).toEqual(['日', '本', '語']);
  });

  it('moves the whole East Asian word after a space instead of splitting it', () => {
    expect(lines([run('ab 日本語です')], 60, false)).toEqual(['ab ', '日本語です']);
    expect(lines([run('ab 日本語です')], 60, true)).toEqual(['ab 日本語', 'です']);
  });

  it('honours eaLnBrk in table cell text', () => {
    const EMU = 12_700;
    const body = (eaLnBrk: boolean) => ({
      verticalAnchor: 't', paragraphs: [{ ...paragraph([run('日本語')], eaLnBrk) }],
      defaultFontSize: null, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
    }) as unknown as TextBody;
    const cellTexts = (eaLnBrk: boolean): string[] => {
      const cell = {
        textBody: body(eaLnBrk), fill: null,
        borderL: null, borderR: null, borderT: null, borderB: null,
        diagonalTL: null, diagonalTR: null,
        gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false,
      } as TableCell;
      const table: TableElement = {
        type: 'table', x: 0, y: 0, width: 12 * EMU, height: 90 * EMU,
        rotation: 0, flipH: false, flipV: false,
        cols: [12 * EMU], rows: [{ height: 90 * EMU, cells: [cell] }],
      };
      const runs: PptxTextRunInfo[] = [];
      renderTable(measuringContext(), table, 1 / EMU, undefined,
        { themeMajorFont: null, themeMinorFont: null, dpr: 1 }, (info) => runs.push(info));
      return runs.map((info) => info.text);
    };
    expect(cellTexts(false)).toEqual(['日本語']);
    expect(cellTexts(true)).toEqual(['日', '本', '語']);
  });
});
