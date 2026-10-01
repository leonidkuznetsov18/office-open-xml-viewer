import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { layoutParagraph, naturalWidthExceedsBbox, paragraphInputRuns, renderTextBody } from './renderer.js';
import type { Paragraph, TextBody } from './types.js';

const SCALE = 1 / 12700;
const RC = { themeMajorFont: null, themeMinorFont: null, dpr: 1 };
const FACES = { fontFamily: 'Corbel', fontFamilyEa: 'Meiryo UI', fontFamilyCs: 'Microsoft Sans Serif' };
function run(text: string, lang?: string, extra: Partial<TextRunData> = {}): TextRunData {
  return { type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: 20, color: '000000', ...FACES, lang, ...extra };
}
function paragraph(runs: Paragraph['runs']): Paragraph {
  return { alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null,
    spaceLine: null, lvl: 0, bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true, runs };
}
function body(runs: Paragraph['runs'], vert = 'horz'): TextBody {
  return { paragraphs: [paragraph(runs)], verticalAnchor: 't', defaultFontSize: 20,
    defaultBold: null, defaultItalic: null, lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert, autoFit: 'none' };
}
function segments(runs: TextRunData[]) {
  return paragraphInputRuns(paragraph(runs), 20, '#000000', SCALE, false, false, 1, undefined, RC)
    .input.flatMap((item) => item.type === 'text' ? [{ text: item.text, face: item.style.faceFamily, font: item.style.font }] : []);
}
function context() {
  const calls: { text: string; font: string }[] = [];
  let font = '';
  const ctx = {
    canvas: { style: {} }, get font() { return font; }, set font(value: string) { font = value; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px', direction: 'ltr',
    textAlign: 'left', textBaseline: 'alphabetic',
    measureText(text: string) {
      const advance = font.includes('"Meiryo UI"') ? 20 : font.includes('"Microsoft Sans Serif"') ? 15 : 5;
      return { width: [...text].reduce((sum, ch) => sum + (ch === '¥' ? 20 : advance), 0),
        actualBoundingBoxAscent: 16, actualBoundingBoxDescent: 4,
        fontBoundingBoxAscent: 16, fontBoundingBoxDescent: 4 };
    },
    fillText(text: string) { calls.push({ text, font }); },
    save() {}, restore() {}, translate() {}, rotate() {}, scale() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, clip() {}, rect() {}, fillRect() {}, setLineDash() {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

describe('PPTX language-dependent font slots through the renderer', () => {
  it('routes measured punctuation and symbol boundaries independently of line-break classes', () => {
    const result = segments([run('A§¨°±´×÷⁇⁈⓾⓿─▟■☙♰⚀❟❨❶B', 'en-US')]);
    expect(result.map(({ text, face }) => [text, face])).toEqual([
      ['A', 'Corbel'], ['§¨°±´×÷⁇⁈⓾', 'Meiryo UI'], ['⓿─▟', 'Corbel'],
      ['■', 'Meiryo UI'], ['☙♰⚀❟❨', 'Corbel'], ['❶', 'Meiryo UI'], ['B', 'Corbel'],
    ]);
  });

  it('uses the East Asian slot for seven curly quotes but keeps U+201F Latin', () => {
    for (const lang of ['ja-JP', 'ko-KR', 'zh-CN', 'zh-TW']) {
      expect(segments([run('‘’‚‛“”„‟', lang)]).map(({ text, face }) => [text, face]))
        .toEqual([['‘’‚‛“”„', 'Meiryo UI'], ['‟', 'Corbel']]);
    }
    expect(segments([run('‘’‚‛“”„‟', 'en-US')])[0].face).toBe('Corbel');
    expect(segments([run('“”', 'en-US', { altLang: 'ja-JP' })])[0].face).toBe('Corbel');
  });

  it('selects cs for all measured European digits without a scalar-only punctuation override', () => {
    for (const lang of ['he-IL', 'ar-SA', 'th-TH', 'hi-IN', 'fa-IR', 'ur-PK', 'yi-001', 'syr-SY', 'ug-CN', 'ar-EG', 'he', 'ur-IN']) {
      expect(segments([run('A0123456789B', lang)]).map(({ text, face }) => [text, face]))
        .toEqual([['A', 'Corbel'], ['0123456789', 'Microsoft Sans Serif'], ['B', 'Corbel']]);
    }
    expect(segments([run('0123456789', 'en-US', { altLang: 'he-IL' })])[0].face).toBe('Corbel');
    expect(segments([run('A»B', 'he-IL')])[0].face).toBe('Corbel');
    expect(segments([run('»', 'th-TH')])[0].face).toBe('Corbel');
    expect(segments([run('×÷⁇⁈“”', 'ar-SA')])[0].face).toBe('Corbel');
  });

  it('routes measured Myanmar extensions through layout and painting while keeping cluster limits', () => {
    const text = '\u1000\ua9e0\uaa60\u1000\ua9e5';
    expect(segments([run(text, 'my-MM')]).map(({ text, face }) => [text, face]))
      .toEqual([['\u1000', 'Microsoft Sans Serif'], ['\ua9e0\uaa60', 'Meiryo UI'],
        ['\u1000\ua9e5', 'Microsoft Sans Serif']]);
    const { ctx, calls } = context();
    renderTextBody(ctx, body([run(text, 'my-MM')]), 0, 0, 300, 100, SCALE);
    expect(calls.find((c) => c.text.includes('\ua9e0'))?.font).toContain('"Meiryo UI"');
    expect(calls.find((c) => c.text.includes('\u1000'))?.font).toContain('"Microsoft Sans Serif"');
  });

  it('measures and wraps in the selected cs face, including the shape-autofit probe', () => {
    const en = paragraph([run('1 2 3', 'en-US')]);
    const he = paragraph([run('1 2 3', 'he-IL')]);
    expect(layoutParagraph(context().ctx, en, 35, 20, '#000', SCALE, 0)).toHaveLength(1);
    expect(layoutParagraph(context().ctx, he, 35, 20, '#000', SCALE, 0).length).toBeGreaterThan(1);
    expect(naturalWidthExceedsBbox(context().ctx, body(en.runs), 35, 0, 0, SCALE, RC)).toBe(false);
    expect(naturalWidthExceedsBbox(context().ctx, body(he.runs), 35, 0, 0, SCALE, RC)).toBe(true);
  });

  it('includes tracking at font-slot seams in the autofit width', () => {
    expect(naturalWidthExceedsBbox(context().ctx,
      body([run('A§B', 'en-US', { letterSpacing: 3 })]), 35, 0, 0, SCALE, RC)).toBe(true);
  });

  it('uses the existing application-default tiers for newly selected ea punctuation', () => {
    expect(segments([run('§°±×÷“”', 'ja-JP', { fontFamilyEa: undefined })])[0].face).toBe('MS Gothic');
    expect(segments([run('§°±×÷“”', 'ja-JP', { fontFamily: 'Perpetua', fontFamilyEa: undefined })])[0].face)
      .toBe('MS Mincho');
  });

  it('maps Japanese backslash before measuring, without changing its Latin slot', () => {
    expect(segments([run('A\\B', 'ja-JP')]).map(({ text, face }) => [text, face])).toEqual([['A¥B', 'Corbel']]);
    expect(segments([run('A\\B', 'en-US')])[0].text).toBe('A\\B');
    expect(naturalWidthExceedsBbox(context().ctx, body([run('A\\B', 'ja-JP')]), 25, 0, 0, SCALE, RC)).toBe(true);
  });

  it('keeps graphemes and formatting across a language-changing run seam', () => {
    expect(segments([run('“', 'ja-JP'), run('\u0301B', 'en-US')]).map(({ text, face }) => [text, face]))
      .toEqual([['“\u0301', 'Meiryo UI'], ['B', 'Corbel']]);
  });

  it('uses the symbol slot throughout U+F0xx and maps known glyphs before layout', () => {
    const result = segments([run('\uf000\uf0b7', 'ja-JP', { fontFamilySym: 'Symbol' })]);
    expect(result[0].face).toBe('Symbol');
    expect(result.map((s) => s.text).join('')).toBe('\uf000•');
    expect(result.at(-1)?.face).toBe('sans-serif');
  });

  it.each(['wordArtVert', 'wordArtVertRtl', 'eaVert', 'vert', 'vert270'])(
    'paints %s text with the same language-selected faces and Japanese glyph mapping', (vert) => {
      const { ctx, calls } = context();
      renderTextBody(ctx, body([run('“§\\', 'ja-JP'), run('123', 'he-IL')], vert), 0, 0, 300, 500, SCALE);
      const fontFor = (ch: string) => calls.find((c) => c.text.includes(ch))?.font;
      expect(fontFor('“')).toContain('"Meiryo UI"');
      expect(fontFor('§')).toContain('"Meiryo UI"');
      expect(fontFor('¥')).toContain('"Corbel"');
      expect(fontFor('1')).toContain('"Microsoft Sans Serif"');
    },
  );
});
