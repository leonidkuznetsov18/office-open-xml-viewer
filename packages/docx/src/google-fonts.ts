import { ScriptPreloadAccumulator } from '@silurus/ooxml-core/internal/script-preload-accumulator';
import type { CjkLang, OfficeFontFallbackRequest } from '@silurus/ooxml-core';
import {
  classifyCjkFont,
  cjkLangFromLanguage,
  scriptPreloadNamesForText,
} from '@silurus/ooxml-core';
import type {
  DocxDocumentModel,
} from './types.js';
import { docxRenderedTextUsages } from './document-content.js';
import { DOCX_GOOGLE_FONTS } from './google-font-registry.js';

// Kept in a parser-model-free module: layout font routing reads the registry,
// while this module traverses parser-owned facts to collect rendered families.
export { DOCX_GOOGLE_FONTS } from './google-font-registry.js';

function* docxTextRuns(doc: DocxDocumentModel): Generator<string> {
  for (const usage of docxRenderedTextUsages(doc)) yield usage.text;
}

/**
 * The font-family names to preload for a document: the theme major/minor fonts,
 * every family rendered text resolves to, plus only the script-fallback Noto
 * faces whose script the document's TEXT actually contains
 * ({@link scriptPreloadNamesForText}).
 *
 * Rendered families come from {@link docxRenderedTextUsages}: all four
 * §17.3.2.26 slots (ascii, hAnsi, eastAsia, cs) resolved through the shaper's
 * own slot rule, across body, tables, headers/footers, notes, complete text-box
 * stories and numbering markers. The loader requests only names that have a
 * {@link DOCX_GOOGLE_FONTS} entry, and the layout font inventory routes only
 * preloaded names to their loaded substitute, so a rendered Cambria run needs
 * Cambria here for Caladea to be both fetched and selected. Rendered names
 * without an entry are omitted (they would be inert). The document font table
 * alone never adds a name.
 *
 * The renderer's font fallback chains still END with the full Noto set, but
 * eagerly fetching the multi-MB CJK families for a document that has no CJK
 * glyphs would block first paint for nothing; an un-preloaded face loads
 * lazily if it ever proves needed.
 *
 * Single source of truth shared by the main-thread `load()` and the render
 * worker. Both derive the set from the SAME parsed {@link DocxDocumentModel}, so
 * they preload an identical set — worker/main rendering must stay
 * pixel-equivalent. (Fonts must also be loaded before pagination, which measures
 * text; both callers await this before paginating.)
 */
export function docxFontPreloadNames(
  doc: DocxDocumentModel,
  fallback?: CjkLang,
): (string | null | undefined)[] {
  const cjkLang =
    classifyCjkFont(doc.majorFont) ?? classifyCjkFont(doc.minorFont) ?? fallback ?? null;
  const scripts = new ScriptPreloadAccumulator(cjkLang);
  const languageNames = new Set<string>();
  const renderedFamilies = new Map<string, string>();
  const themeKeys = new Set([doc.majorFont, doc.minorFont]
    .map((family) => family?.trim().toLocaleLowerCase('en-US')));
  for (const usage of docxRenderedTextUsages(doc)) {
    for (const family of usage.fontFamilies) {
      const name = family?.trim();
      const key = name?.toLocaleLowerCase('en-US');
      if (name && key && key in DOCX_GOOGLE_FONTS && !themeKeys.has(key)
        && !renderedFamilies.has(key)) {
        renderedFamilies.set(key, name);
      }
    }
    const region = cjkLangFromLanguage(usage.eastAsiaLanguage);
    if (region) {
      for (const name of scriptPreloadNamesForText([usage.text], region, true)) languageNames.add(name);
    } else {
      scripts.addText([usage.text]);
    }
  }
  return [doc.majorFont, doc.minorFont, ...renderedFamilies.values(),
    ...new Set([...scripts.names(), ...languageNames])];
}

/** Probe exact local style tuples used by rendered text. The shared loader
 * declines uncatalogued names, and the document font table alone never queues
 * a face that rendered content does not use. No font bytes are packaged. */
export function docxOfficeFontFallbackRequests(
  doc: DocxDocumentModel,
): OfficeFontFallbackRequest[] {
  const tuples = new Map<string, OfficeFontFallbackRequest>();
  const add = (family: string | null | undefined, bold = false, italic = false) => {
    const name = family?.trim();
    if (!name) return;
    const weight = bold ? 700 : 400;
    const style = italic ? 'italic' : 'normal';
    tuples.set(`${name.toLocaleLowerCase('en-US')}:${weight}:${style}`, { family: name, weight, style });
  };
  // Theme faces can govern paragraph marks and inherited runs even when a
  // particular authored run does not repeat its font name.
  add(doc.majorFont);
  add(doc.minorFont);
  for (const usage of docxRenderedTextUsages(doc)) {
    for (const family of usage.fontFamilies) add(family, usage.bold, usage.italic);
    if (usage.text && (usage.latinFontFamily === null ||
      (usage.latinFontFamily === undefined && !usage.fontFamilies.some(Boolean)))) {
      // A Latin slot can inherit the theme even when another script slot has an
      // authored face. The inherited bold/italic axes need their own resource.
      add(doc.minorFont, usage.bold, usage.italic);
    }
  }
  return [...tuples.values()];
}


export function docxScriptCjkLanguage(doc: DocxDocumentModel): CjkLang | null {
  const scripts = new ScriptPreloadAccumulator(null);
  scripts.addText(docxTextRuns(doc));
  return scripts.scriptCjkLanguage();
}
