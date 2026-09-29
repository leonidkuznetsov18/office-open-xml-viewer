import { describe, expect, it } from 'vitest';
import { GOOGLE_FONT_SUBSTITUTES } from './google-fonts.js';
import { fontSubstituteScriptCoversText, substituteEntryCoversText } from './substitute-script.js';

describe('script-scoped visual substitutes', () => {
  it('scopes only the Arabic visual substitutes, never metric or same-name faces', () => {
    const scoped = Object.entries(GOOGLE_FONT_SUBSTITUTES)
      .filter(([, entry]) => entry.script !== undefined).map(([key]) => key).sort();
    expect(scoped).toEqual([
      'arabic typesetting', 'sakkal majalla', 'simplified arabic', 'traditional arabic', 'univers next arabic',
    ]);
  });

  it('distinguishes an exclusive Arabic span from a span that merely contains Arabic', () => {
    expect(fontSubstituteScriptCoversText('arabic', 'مرحبا بكم', 'exclusive')).toBe(true);
    expect(fontSubstituteScriptCoversText('arabic', 'مرحبا 2026', 'exclusive')).toBe(false);
    expect(fontSubstituteScriptCoversText('arabic', 'مرحبا 2026', 'any')).toBe(true);
    expect(fontSubstituteScriptCoversText('arabic', 'Leader 2026', 'any')).toBe(false);
    expect(fontSubstituteScriptCoversText('arabic', '   ', 'exclusive')).toBe(false);
    // Hebrew is complex script but not Arabic script.
    expect(fontSubstituteScriptCoversText('arabic', 'שלום', 'any')).toBe(false);
  });

  it('lets unscoped entries cover any text', () => {
    expect(substituteEntryCoversText(GOOGLE_FONT_SUBSTITUTES.calibri, 'Latin', 'exclusive')).toBe(true);
    expect(substituteEntryCoversText(GOOGLE_FONT_SUBSTITUTES['sakkal majalla'], 'Latin', 'any')).toBe(false);
  });
});
