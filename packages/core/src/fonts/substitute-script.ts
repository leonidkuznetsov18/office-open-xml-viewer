/**
 * Script scope of a *visual* web-font substitute.
 *
 * Some registry entries map an Office face to a substitute that only
 * stands in for one script. For example, Sakkal Majalla maps to Noto Naskh
 * Arabic. The substitute also carries Latin glyphs, but those glyphs, and the
 * substitute's much taller line box, have nothing to do with the authored face:
 * the authored face has its own Latin glyphs. A scoped substitute may
 * therefore paint and measure only characters of its script. Every other
 * character requested from the same family falls back exactly as if no
 * substitute existed.
 *
 * This is library font-substitution policy (ECMA-376 §17.8.2 leaves
 * substitution implementation-defined), not an Office layout rule. It is shared
 * by every format that consults {@link GOOGLE_FONT_SUBSTITUTES}.
 */
export type FontSubstituteScript = 'arabic';

/** Arabic-script blocks: Arabic, Supplement, Extended-A/B, Presentation
 * Forms-A/B and Arabic Mathematical Alphabetic Symbols. Hebrew, Syriac and the
 * other complex-script blocks are deliberately excluded: an Arabic substitute
 * has no glyphs for them. */
function isArabicScriptCodePoint(cp: number): boolean {
  return (cp >= 0x0600 && cp <= 0x06ff)
    || (cp >= 0x0750 && cp <= 0x077f)
    || (cp >= 0x0870 && cp <= 0x08ff)
    || (cp >= 0xfb50 && cp <= 0xfdff)
    || (cp >= 0xfe70 && cp <= 0xfeff)
    || (cp >= 0x1ee00 && cp <= 0x1eeff);
}

/** Whether one code point belongs to `script`. */
export function isFontSubstituteScriptCodePoint(script: FontSubstituteScript, cp: number): boolean {
  switch (script) {
    case 'arabic': return isArabicScriptCodePoint(cp);
  }
}

/** Script-neutral separators inherit the script of the text around them:
 * white space and the joiner/bidi controls (ZWNJ, ZWJ, LRM, RLM, ALM). A space
 * between two Arabic words stays with the Arabic text. Digits and punctuation
 * are Latin-shaped in the substitute, so they are not neutral. */
export function isFontSubstituteScriptNeutralCodePoint(cp: number): boolean {
  return cp === 0x0020 || cp === 0x00a0 || cp === 0x0009
    || (cp >= 0x2000 && cp <= 0x200f) || cp === 0x202f || cp === 0x061c;
}

/**
 * Whether a scoped substitute may supply `text`.
 * - `'exclusive'`: every non-neutral character belongs to the script, and at
 *   least one does. Use this for a span that can mix scripts, such as an
 *   ECMA-376 §17.3.2.26 ascii-slot span.
 * - `'any'`: at least one character belongs to the script. Use this for a
 *   complex-script span that the script owns as a whole, including its neutral
 *   digits and punctuation.
 */
export function fontSubstituteScriptCoversText(
  script: FontSubstituteScript,
  text: string,
  mode: 'exclusive' | 'any',
): boolean {
  let covered = false;
  for (const character of text) {
    const cp = character.codePointAt(0) ?? 0;
    if (isFontSubstituteScriptCodePoint(script, cp)) covered = true;
    else if (mode === 'exclusive' && !isFontSubstituteScriptNeutralCodePoint(cp)) return false;
  }
  return covered;
}

/** Whether a registry entry may supply `text`. Unscoped entries always may. */
export function substituteEntryCoversText(
  entry: Readonly<{ script?: FontSubstituteScript }> | undefined,
  text: string,
  mode: 'exclusive' | 'any',
): boolean {
  return !entry?.script || fontSubstituteScriptCoversText(entry.script, text, mode);
}
