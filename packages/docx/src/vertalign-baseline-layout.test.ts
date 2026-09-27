import { expect, it } from 'vitest';
import { buildSegments } from './line-layout.js';
import type { DocRun } from './types.js';

it('uses an authored baseline to clear a superscript style before measuring text', () => {
  const run = {
    type: 'text', text: 'ordinary', fontSize: 10, fontFamily: 'Arial',
    bold: false, italic: false, underline: false, strikethrough: false,
    color: null, isLink: false, background: null, vertAlign: 'super',
    typographyInput: {
      verticalAlign: { status: 'valid', raw: 'baseline', value: 'baseline' },
    },
  } as unknown as DocRun;
  const segments = buildSegments([run], { pageIndex: 0, totalPages: 1 });
  const text = segments.find((segment) => 'text' in segment);
  expect(text).toBeDefined();
  expect(text && 'text' in text && text.vertAlign).toBeNull();
});
