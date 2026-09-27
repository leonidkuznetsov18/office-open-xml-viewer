import { describe, expect, it } from 'vitest';
import { breakDrawingMlText, type DrawingMlInputRun } from './index.js';

// Synthetic advances keep the test independent of the host Canvas fonts. The
// assertions are the line choices in the matched Office wrap controls (C00–C11).
const measure = (value: string): number =>
  [...value].reduce((width, ch) => width + (/\p{Script=Han}/u.test(ch) || ch === '、' ? 20 : 10), 0);

function lines(parts: readonly string[], width: number, defaultTabSize = 72): string[] {
  const runs: DrawingMlInputRun<string>[] = parts.map((text) => ({ type: 'text', text, style: 'same' }));
  return breakDrawingMlText(runs, {
    maxWidth: width,
    measureText: measure,
    sameStyle: (a, b) => a === b,
    defaultTabSize,
  }).map((line) => line.segments.map((segment) => segment.type === 'text'
    ? segment.text : segment.type === 'tab' ? '\t' : '').join('').replace(/ +$/u, ''));
}

describe('matched PowerPoint and Excel DrawingML wrap controls', () => {
  it('wraps a Latin word at the same grapheme with or without a run seam (C00/C01)', () => {
    expect(lines(['abcdef'], 45)).toEqual(['abcd', 'ef']);
    expect(lines(['abc', 'def'], 45)).toEqual(['abcd', 'ef']);
  });

  it('uses the space, hyphen, and CJK/Latin opportunities (C02–C04/C07)', () => {
    expect(lines(['abc def'], 50)).toEqual(['abc', 'def']);
    expect(lines(['日本語Power'], 70)).toEqual(['日本語', 'Power']);
    expect(lines(['non-managed'], 70)).toEqual(['non-', 'managed']);
    expect(lines(['abc ', '$', '100'], 55)).toEqual(['abc', '$100']);
  });

  it('does not reserve a visual continuation for terminal ordinary spaces (C05/C06)', () => {
    expect(lines(['abc  '], 35)).toEqual(['abc']);
    expect(lines(['abc', '  '], 35)).toEqual(['abc']);
  });

  it('retains the advance of spaces before an authored line break for alignment', () => {
    const result = breakDrawingMlText<string>([
      { type: 'text', text: 'Analyze & ', style: 'same' },
      { type: 'break' },
      { type: 'text', text: 'Control', style: 'same' },
    ], { maxWidth: 200, measureText: measure });
    expect(result.map((line) => ({
      text: line.segments.map((segment) => segment.type === 'text' ? segment.text : '').join(''),
      width: line.width,
    }))).toEqual([
      { text: 'Analyze & ', width: 100 },
      { text: 'Control', width: 70 },
    ]);
  });

  it('keeps the observed authored punctuation seam and NBSP behavior (C08/C10)', () => {
    expect(lines(['日本語', '、次'], 60)).toEqual(['日本語', '、次']);
    expect(lines(['abc\u00a0def'], 50)).toEqual(['abc\u00a0d', 'ef']);
  });

  it('breaks an overwide word at grapheme boundaries (C09)', () => {
    expect(lines(['supercalifragilistic'], 80)).toEqual(['supercal', 'ifragili', 'stic']);
  });

  it('carries an overflowing tab to the next line and seats one glyph after its stop (C11)', () => {
    expect(lines(['abc\tdef'], 85)).toEqual(['abc', '\td', 'ef']);
  });

});
