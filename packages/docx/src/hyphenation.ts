import french from 'hyphen/fr/index.js';

/** ECMA-376 §17.15.1.10 requests automatic hyphenation but does not define a
 * dictionary. The bundled French patterns provide bounded discretionary
 * candidates for French language tags; other languages remain unaltered until
 * a matching dictionary is available. Word's French controls show that an
 * enabled document breaks words only when the remaining band exceeds its
 * authored §17.15.1.43 hyphenation zone. */
export function frenchHyphenationOffsets(text: string, language: string | undefined): number[] {
  if (language?.split('-')[0]?.toLowerCase() !== 'fr' || text.length < 5 || text.length > 128) {
    return [];
  }
  const word = /^\p{L}+/u.exec(text)?.[0];
  // Punctuation after the word belongs to the same layout token but does not
  // participate in dictionary lookup. A later letter would mean that this is
  // a compound token whose language boundary has not been established.
  if (!word || word.length < 5 || /\p{L}/u.test(text.slice(word.length))) return [];
  const separated = french.hyphenateSync(word);
  const offsets: number[] = [];
  let offset = 0;
  for (const part of separated.split('\u00ad')) {
    offset += part.length;
    if (offset > 1 && offset < word.length - 1) offsets.push(offset);
  }
  return offsets;
}
