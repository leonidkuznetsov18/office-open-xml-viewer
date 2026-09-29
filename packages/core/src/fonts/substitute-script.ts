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

const FORMAT_CONTROL = /^\p{Cf}$/u;
const ARABIC_SCRIPT = /^\p{Script=Arabic}$/u;
const LETTER_OR_MARK = /^[\p{L}\p{M}]$/u;

/** Arabic-script characters that a substitute actually draws: every visible
 * Unicode Script=Arabic character (letters, marks, Arabic-Indic digits and
 * Arabic punctuation). Invisible format controls are excluded even when their
 * Script property is Arabic (ALM U+061C, the U+0600-U+0605 number signs,
 * U+06DD, U+08E2). U+FEFF is Script=Common and is excluded anyway. Hebrew,
 * Syriac and the other complex-script blocks are excluded: an Arabic
 * substitute has no glyphs for them. */
function isArabicScriptCodePoint(cp: number): boolean {
  const character = String.fromCodePoint(cp);
  return ARABIC_SCRIPT.test(character) && !FORMAT_CONTROL.test(character);
}

/** Whether one code point is a visible character of `script` that the
 * substitute paints. */
export function isFontSubstituteScriptCodePoint(script: FontSubstituteScript, cp: number): boolean {
  switch (script) {
    case 'arabic': return isArabicScriptCodePoint(cp);
  }
}

/** Whether one code point proves that text is in `script`: a letter or a
 * combining mark of that script. Digits, punctuation and invisible controls
 * never enable a scoped substitute on their own. */
function provesFontSubstituteScript(script: FontSubstituteScript, cp: number): boolean {
  return isFontSubstituteScriptCodePoint(script, cp) && LETTER_OR_MARK.test(String.fromCodePoint(cp));
}

/** Script-neutral characters inherit the script of the text around them:
 * white space and every invisible format control (ZWNJ, ZWJ, LRM, RLM, ALM,
 * U+FEFF, the bidi embeddings and isolates). A space between two Arabic words
 * stays with the Arabic text. Latin digits and punctuation are not neutral:
 * the substitute would draw them with its own Latin-style glyphs. */
export function isFontSubstituteScriptNeutralCodePoint(cp: number): boolean {
  const character = String.fromCodePoint(cp);
  return /^\s$/u.test(character) || FORMAT_CONTROL.test(character);
}

/**
 * Whether a scoped substitute may supply `text`.
 * - `'exclusive'`: every non-neutral character belongs to the script, and at
 *   least one is a letter or mark of the script. Use this for a span that can mix scripts, such as an
 *   ECMA-376 §17.3.2.26 ascii-slot span.
 * - `'any'`: at least one character is a letter or mark of the script. Use this for a
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
    if (provesFontSubstituteScript(script, cp)) covered = true;
    else if (mode === 'exclusive' && !isFontSubstituteScriptCodePoint(script, cp)
      && !isFontSubstituteScriptNeutralCodePoint(cp)) return false;
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
