import { afterEach, describe, expect, it, vi } from 'vitest';
import { excludeEmbeddedFontFamilies, loadEmbeddedFonts, uncoveredOfficeFontRequests } from './embedded-fonts.js';
import { paragraphInputRuns } from './renderer.js';
import type { Paragraph } from './types.js';
import type { PptxEmbeddedFontRef } from './worker-protocol';

const globals = globalThis as Record<string, unknown>;
const original = { document: globals.document, self: globals.self, FontFace: globals.FontFace };

afterEach(() => {
  globals.document = original.document;
  globals.self = original.self;
  globals.FontFace = original.FontFace;
  vi.restoreAllMocks();
});

function installFontFaceSet(failLoad = false) {
  const added: Array<{ family: string; source: ArrayBuffer; descriptors: FontFaceDescriptors }> = [];
  class FakeFontFace {
    constructor(
      public family: string,
      public source: ArrayBuffer,
      public descriptors: FontFaceDescriptors,
    ) {}
    get weight() { return this.descriptors.weight; }
    get style() { return this.descriptors.style; }
    load() {
      return failLoad
        ? Promise.reject(new Error('load failed'))
        : Promise.resolve(this);
    }
  }
  globals.FontFace = FakeFontFace;
  globals.document = { fonts: { add: (face: typeof added[number]) => added.push(face), ready: Promise.resolve() } };
  delete globals.self;
  return added;
}

const bytes = () => new Uint8Array([0, 1, 0, 0, 1]);

// A table-directory fixture with a real Unicode format-12 cmap and OS/2
// design metrics. FontFace registration is the only browser boundary mocked.
function cjkResource(codePoint = 0x6f22): Uint8Array {
  const bytes = new Uint8Array(284);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, 4);
  for (const [index, tag, offset, length] of [
    [0, 'head', 76, 54], [1, 'hhea', 130, 36],
    [2, 'OS/2', 166, 78], [3, 'cmap', 244, 40],
  ] as const) {
    for (let i = 0; i < 4; i++) bytes[12 + index * 16 + i] = tag.charCodeAt(i);
    view.setUint32(20 + index * 16, offset);
    view.setUint32(24 + index * 16, length);
  }
  view.setUint16(94, 1000);
  view.setInt16(134, 700);
  view.setInt16(136, -300);
  view.setUint16(240, 700);
  view.setUint16(242, 300);
  view.setUint16(246, 1);
  view.setUint16(248, 3);
  view.setUint16(250, 10);
  view.setUint32(252, 12);
  view.setUint16(256, 12);
  view.setUint32(260, 28);
  view.setUint32(268, 1);
  view.setUint32(272, codePoint); // Default repertoire: 漢 is covered, 한 is missing.
  view.setUint32(276, codePoint);
  view.setUint32(280, 1);
  return bytes;
}

// The full-Unicode cmap adds a BMP glyph absent from the BMP-only map,
// as permitted by OpenType cmap “Encoding records and encodings”.
function supersetCjkResource(): Uint8Array {
  const bytes = new Uint8Array(336);
  bytes.set(cjkResource().subarray(0, 244));
  const view = new DataView(bytes.buffer);
  view.setUint32(72, 92); // cmap table length
  view.setUint16(246, 2);
  for (const [index, encoding, offset] of [[0, 1, 20], [1, 10, 52]]) {
    const record = 248 + index * 8;
    view.setUint16(record, 3);
    view.setUint16(record + 2, encoding);
    view.setUint32(record + 4, offset);
  }
  const bmp = 264;
  view.setUint16(bmp, 4);
  view.setUint16(bmp + 2, 32);
  view.setUint16(bmp + 6, 4);
  view.setUint16(bmp + 14, 0x41);
  view.setUint16(bmp + 16, 0xffff);
  view.setUint16(bmp + 20, 0x41);
  view.setUint16(bmp + 22, 0xffff);
  view.setInt16(bmp + 24, 1 - 0x41);
  view.setInt16(bmp + 26, 1);
  const full = 296;
  view.setUint16(full, 12);
  view.setUint32(full + 4, 40);
  view.setUint32(full + 12, 2);
  for (const [index, cp] of [[0, 0x41], [1, 0x6f22]]) {
    view.setUint32(full + 16 + index * 12, cp);
    view.setUint32(full + 20 + index * 12, cp);
    view.setUint32(full + 24 + index * 12, index + 1);
  }
  return bytes;
}

describe('loadEmbeddedFonts (ECMA-376 §19.2.1.9 / §15.2.13)', () => {
  it('uses the embedded cmap and metrics for CJK, including synthetic styles and same-name replacements', async () => {
    installFontFaceSet();
    for (const fontName of ['Uncatalogued CJK', 'DengXian']) {
      const loaded = await loadEmbeddedFonts([{
        fontName, style: 'regular', partPath: 'ppt/fonts/cjk.fntdata', contentType: 'application/x-font-ttf',
      }], async () => cjkResource());
      const rc = {
        themeMajorFont: null, themeMinorFont: null, dpr: 1,
        embeddedFontAliases: loaded.aliases, embeddedFontAuthoredFamilies: loaded.authoredFamilies,
        embeddedFontTuples: loaded.tuples, embeddedFontMetrics: loaded.metrics,
      };
      for (const [bold, italic] of [[false, false], [true, true]]) {
        const para = { runs: [{ type: 'text', text: '漢한', fontFamily: 'Arial', fontFamilyCs: fontName,
          lang: 'en-US', fontSize: 22, bold, italic }], tabStops: [] } as unknown as Paragraph;
        const segments = paragraphInputRuns(para, 22, '#000', 1 / 12700, bold, italic, 1, undefined, rc)
          .input.filter((item) => item.type === 'text');
        const han = segments.find((item) => item.text.includes('漢'));
        const hangul = segments.find((item) => item.text.includes('한'));
        expect(han?.style.faceFamily).toBe(loaded.aliases.get(fontName.toLowerCase()));
        expect(han?.style.lineMetric?.share).toBe(0.7);
        // The missing glyph continues through the actual painting stack; an
        // uncatalogued family's sans chain and Deng's Far-East chain differ.
        expect(hangul?.style.faceFamily).toBe(fontName === 'DengXian' ? 'Batang' : 'Malgun Gothic');
        expect(hangul?.style.lineMetric).toBeDefined();
        expect(hangul?.style.lineMetric?.share).not.toBe(0.7);
      }
    }
  });

  it('stops CJK fallback attribution when embedded BMP and full-Unicode cmaps disagree', async () => {
    installFontFaceSet();
    for (const fontName of ['Deck CJK', 'DengXian']) {
      const loaded = await loadEmbeddedFonts([{
        fontName, style: 'regular', partPath: 'font', contentType: 'application/x-font-ttf',
      }], async () => supersetCjkResource());
      const rc = { themeMajorFont: null, themeMinorFont: null, dpr: 1,
        embeddedFontAliases: loaded.aliases, embeddedFontAuthoredFamilies: loaded.authoredFamilies,
        embeddedFontTuples: loaded.tuples, embeddedFontMetrics: loaded.metrics };
      const para = { runs: [{ type: 'text', text: '漢', fontFamily: 'Avenir', fontFamilyCs: fontName,
        lang: 'en-US', fontSize: 22 }], tabStops: [] } as unknown as Paragraph;
      const item = paragraphInputRuns(para, 22, '#000', 1 / 12700, false, false, 1, undefined, rc)
        .input.find((item) => item.type === 'text');
      expect(item?.style.font).toContain(`"${loaded.aliases.get(fontName.toLowerCase())}"`);
      expect(item?.style.faceFamily).toBeUndefined();
      expect(item?.style.lineMetric).toBeUndefined();
    }
  });

  it('matches style before weight for a missing embedded bold-italic cut', async () => {
    installFontFaceSet();
    const loaded = await loadEmbeddedFonts([
      { fontName: 'Deck Serif', style: 'bold', partPath: 'bold', contentType: 'application/x-font-ttf' },
      { fontName: 'Deck Serif', style: 'italic', partPath: 'italic', contentType: 'application/x-font-ttf' },
    ], async (path) => cjkResource(path === 'italic' ? 0xa7 : 0x41));
    const rc = { themeMajorFont: null, themeMinorFont: null, dpr: 1,
      embeddedFontAliases: loaded.aliases, embeddedFontAuthoredFamilies: loaded.authoredFamilies,
      embeddedFontTuples: loaded.tuples, embeddedFontMetrics: loaded.metrics };
    const para = { runs: [{ type: 'text', text: '§', fontFamily: 'Avenir', fontFamilyCs: 'Deck Serif',
      lang: 'en-US', fontSize: 22, bold: true, italic: true }], tabStops: [] } as unknown as Paragraph;
    const item = paragraphInputRuns(para, 22, '#000', 1 / 12700, true, true, 1, undefined, rc)
      .input.find((item) => item.type === 'text');
    expect(item?.style.faceFamily).toBe(loaded.aliases.get('deck serif'));
    expect(item?.style.lineMetric?.share).toBe(0.7);
    // A registered italic cut with unreadable metrics still owns the glyph;
    // the readable bold cut cannot stand in for the face Canvas selected.
    const unreadable = await loadEmbeddedFonts([
      { fontName: 'Deck Serif', style: 'bold', partPath: 'bold', contentType: 'application/x-font-ttf' },
      { fontName: 'Deck Serif', style: 'italic', partPath: 'italic', contentType: 'application/x-font-ttf' },
    ], async (path) => path === 'italic' ? bytes() : cjkResource(0x41));
    const unknownRc = { ...rc, embeddedFontAliases: unreadable.aliases,
      embeddedFontAuthoredFamilies: unreadable.authoredFamilies,
      embeddedFontTuples: unreadable.tuples, embeddedFontMetrics: unreadable.metrics };
    const unknown = paragraphInputRuns(para, 22, '#000', 1 / 12700, true, true, 1, undefined, unknownRc)
      .input.find((item) => item.type === 'text');
    expect(unknown?.style.lineMetric).toBeUndefined();
  });

  it('maps all four PresentationML slots to CSS weight and style', async () => {
    const added = installFontFaceSet();
    const refs: PptxEmbeddedFontRef[] = ['regular', 'bold', 'italic', 'boldItalic'].map(
      (style, index) => ({
        fontName: 'Deck Sans',
        style: style as PptxEmbeddedFontRef['style'],
        partPath: `ppt/fonts/font${index + 1}.fntdata`,
        contentType: 'application/x-font-ttf',
      }),
    );
    const loaded = await loadEmbeddedFonts(refs, async () => bytes());
    expect(added.map((face) => `${face.descriptors.weight}/${face.descriptors.style}`).sort()).toEqual([
      'bold/italic', 'bold/normal', 'normal/italic', 'normal/normal',
    ]);
    expect(new Set(added.map((face) => face.family)).size).toBe(1);
    expect(loaded.aliases.get('deck sans')).toBe(added[0].family);
    expect(loaded.authoredFamilies.get(added[0].family)).toBe('deck sans');
    expect(loaded.tuples).toEqual(new Set([
      'deck sans:400:normal', 'deck sans:700:normal',
      'deck sans:400:italic', 'deck sans:700:italic',
    ]));
  });

  it('keeps raw PPTX bytes and skips an unreadable part without aborting siblings', async () => {
    const added = installFontFaceSet();
    const refs: PptxEmbeddedFontRef[] = [
      { fontName: 'Good', style: 'regular', partPath: 'ppt/fonts/good.fntdata', contentType: 'application/x-font-ttf' },
      { fontName: 'Missing', style: 'regular', partPath: 'ppt/fonts/missing.fntdata', contentType: 'application/x-fontdata' },
    ];
    const loaded = await loadEmbeddedFonts(refs, async (path) => {
      if (path.includes('missing')) throw new Error('missing');
      return bytes();
    });
    expect(added).toHaveLength(1);
    expect(loaded.aliases.has('good')).toBe(true);
    expect(loaded.aliases.has('missing')).toBe(false);
    expect(Array.from(new Uint8Array(added[0].source))).toEqual(Array.from(bytes()));
  });

  it('does not fetch when there are no embedded fonts', async () => {
    installFontFaceSet();
    const fetchFont = vi.fn(async () => bytes());
    await loadEmbeddedFonts([], fetchFont);
    expect(fetchFont).not.toHaveBeenCalled();
  });

  it('bounds concurrent extraction to two font parts', async () => {
    installFontFaceSet();
    let active = 0;
    let peak = 0;
    const refs: PptxEmbeddedFontRef[] = Array.from({ length: 5 }, (_, index) => ({
      fontName: `Deck Font ${index}`,
      style: 'regular',
      partPath: `ppt/fonts/font${index}.fntdata`,
      contentType: 'application/x-font-ttf',
    }));
    await loadEmbeddedFonts(refs, async () => {
      active++;
      peak = Math.max(peak, active);
      await Promise.resolve();
      active--;
      return bytes();
    });
    expect(peak).toBe(2);
  });

  it('isolates the same authored family across concurrently open presentations', async () => {
    const added = installFontFaceSet();
    const refs: PptxEmbeddedFontRef[] = [{
      fontName: 'Shared Family', style: 'regular', partPath: 'ppt/fonts/font1.fntdata',
      contentType: 'application/x-font-ttf',
    }];
    const first = await loadEmbeddedFonts(refs, async () => bytes());
    const second = await loadEmbeddedFonts(refs, async () => new Uint8Array([0, 1, 0, 0, 2]));
    expect(added).toHaveLength(2);
    expect(first.aliases.get('shared family')).not.toBe(second.aliases.get('shared family'));
    expect(new Set(added.map((face) => face.family)).size).toBe(2);
  });

  it('keeps successfully loaded embedded families ahead of optional Google-font substitutes', async () => {
    installFontFaceSet();
    const loaded = await loadEmbeddedFonts([{
      fontName: 'Calibri', style: 'regular', partPath: 'ppt/fonts/font1.fntdata',
      contentType: 'application/x-font-ttf',
    }], async () => bytes());
    expect(excludeEmbeddedFontFamilies(['Aptos', 'calibri', null], loaded.aliases)).toEqual(['Aptos', null]);
  });

  it('keeps the web substitute eligible when an embedded face fails to load', async () => {
    installFontFaceSet(true);
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    const loaded = await loadEmbeddedFonts([{
      fontName: 'Calibri', style: 'regular', partPath: 'ppt/fonts/font1.fntdata',
      contentType: 'application/x-font-ttf',
    }], async () => bytes());
    expect(loaded.faces).toEqual([]);
    expect(excludeEmbeddedFontFamilies(['calibri'], loaded.aliases)).toEqual(['calibri']);
  });

  it('records only successfully registered style tuples', async () => {
    installFontFaceSet();
    const refs: PptxEmbeddedFontRef[] = [
      { fontName: 'Calibri', style: 'regular', partPath: 'regular.ttf', contentType: 'application/x-font-ttf' },
      { fontName: 'Calibri', style: 'bold', partPath: 'bold.ttf', contentType: 'application/x-font-ttf' },
    ];
    const loaded = await loadEmbeddedFonts(refs, async (path) => {
      if (path === 'bold.ttf') throw new Error('missing');
      return bytes();
    });
    expect(loaded.tuples).toEqual(new Set(['calibri:400:normal']));
    expect(uncoveredOfficeFontRequests([
      { family: 'Calibri', weight: 400, style: 'normal' },
      { family: 'Calibri', weight: 700, style: 'normal' },
    ], loaded.tuples)).toEqual([
      { family: 'Calibri', weight: 700, style: 'normal' },
    ]);
  });
});
