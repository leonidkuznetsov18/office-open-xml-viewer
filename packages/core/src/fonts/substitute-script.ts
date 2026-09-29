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
import { graphemeClusterOffsets } from '../text/sea-break.js';

export type FontSubstituteScript = 'arabic';

const FORMAT_CONTROL = /^\p{Cf}$/u;
const WHITE_SPACE = /^\s$/u;
const MARK = /^\p{M}$/u;
const ARABIC_EXTENSIONS = /^\p{Script_Extensions=Arabic}$/u;
const ARABIC_LETTER = /^(?=\p{Script=Arabic})\p{L}$/u;

/** Class of one grapheme cluster for a script-scoped substitute.
 * - `'script'`: the cluster's base belongs to the script by UAX #24
 *   Script_Extensions. This includes Arabic letters, Arabic combining marks
 *   (Script=Inherited, scx=Arab: fatha, sukun, tanwin, superscript alef...),
 *   tatweel, the Arabic comma and Arabic-Indic digits.
 * - `'neutral'`: white space, invisible format controls (ZWNJ, ZWJ, LRM, RLM,
 *   ALM, U+FEFF, the bidi controls), or a lone combining mark with no script of
 *   its own. It inherits the script of the preceding text, as UAX #24
 *   prescribes for Inherited characters.
 * - `'other'`: anything else, such as Latin letters, Latin digits and
 *   punctuation, which the substitute would draw with its own Latin-style glyphs.
 * The base is the first code point that is not white space or a format control.
 * Classifying whole clusters means a mark is never split from its base. */
export function fontSubstituteScriptClusterClass(
  script: FontSubstituteScript,
  cluster: string,
): 'script' | 'neutral' | 'other' {
  for (const character of cluster) {
    if (WHITE_SPACE.test(character) || FORMAT_CONTROL.test(character)) continue;
    switch (script) {
      case 'arabic':
        if (ARABIC_EXTENSIONS.test(character)) return 'script';
        return MARK.test(character) ? 'neutral' : 'other';
    }
  }
  return 'neutral';
}

/** Whether a `'script'` cluster proves the text is in that script: its base is
 * a letter of the script, or a combining mark whose Script_Extensions include
 * it (a lone vowel sign). Tatweel, Arabic punctuation and Arabic-Indic digits
 * belong to Arabic text but never enable the substitute on their own. */
function clusterProvesScript(script: FontSubstituteScript, cluster: string): boolean {
  for (const character of cluster) {
    if (WHITE_SPACE.test(character) || FORMAT_CONTROL.test(character)) continue;
    switch (script) {
      case 'arabic':
        return ARABIC_LETTER.test(character) || (MARK.test(character) && ARABIC_EXTENSIONS.test(character));
    }
  }
  return false;
}

function graphemeClusters(text: string): string[] {
  const offsets = [...new Set([0, ...graphemeClusterOffsets(text), text.length])].sort((x, y) => x - y);
  return offsets.slice(0, -1).map((start, index) => text.slice(start, offsets[index + 1]));
}

/**
 * Whether a scoped substitute may supply `text`, judged by grapheme cluster.
 * - `'exclusive'`: every cluster is `'script'` or `'neutral'`, and at least
 *   one proves the script. Use this for a span that can mix scripts, such as an
 *   ECMA-376 §17.3.2.26 ascii-slot span.
 * - `'any'`: at least one cluster proves the script. Use this for a
 *   complex-script span that the script owns as a whole, including its neutral
 *   digits and punctuation.
 */
export function fontSubstituteScriptCoversText(
  script: FontSubstituteScript,
  text: string,
  mode: 'exclusive' | 'any',
): boolean {
  let covered = false;
  for (const cluster of graphemeClusters(text)) {
    if (clusterProvesScript(script, cluster)) covered = true;
    else if (mode === 'exclusive' && fontSubstituteScriptClusterClass(script, cluster) === 'other') return false;
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
