import {
  parseOpenTypeResourceMetrics,
  registerEmbeddedFonts,
  unregisterEmbeddedFonts,
  type EmbeddedFontFace,
  type OfficeFontFallbackRequest,
} from '@silurus/ooxml-core';
import type { PptxEmbeddedFontRef } from './worker-protocol';
import { powerPointResourceFaceMetrics, type PowerPointFaceMetrics } from './powerpoint-line-metrics.js';

export interface LoadedPptxEmbeddedFonts {
  readonly faces: FontFace[];
  /** Lower-cased authored family → presentation-scoped FontFace family. */
  readonly aliases: ReadonlyMap<string, string>;
  /** Presentation-scoped FontFace family → lower-cased authored family. */
  readonly authoredFamilies: ReadonlyMap<string, string>;
  /** Successfully registered authored family/style slots (§19.2.1.9). */
  readonly tuples: ReadonlySet<string>;
  /** Line metrics of each registered face, keyed like `tuples`, from the
   * font part's own OS/2 tables, with bounded scalar cmap coverage retained
   * for drawing-resource attribution. A part whose tables do not parse (EOT, an
   * unsupported collection) has no entry: its lines keep the metric model of
   * their other faces (#1689). */
  readonly metrics: ReadonlyMap<string, PowerPointFaceMetrics>;
}

/** Only an actually registered PresentationML face occupies its style slot. */
export function uncoveredOfficeFontRequests(
  requests: readonly OfficeFontFallbackRequest[],
  embeddedTuples: ReadonlySet<string>,
): OfficeFontFallbackRequest[] {
  return requests.filter((request) =>
    !embeddedTuples.has(`${request.family.toLowerCase()}:${request.weight}:${request.style}`));
}

let nextFontScope = 1;

function normalizedFamily(value: string): string {
  return value.trim().toLowerCase();
}

/**
 * Load PresentationML font parts and register them before text measurement.
 * PPTX font parts are raw sfnt or EOT (ECMA-376 Part 1 §15.2.13), never the
 * WordprocessingML ODTTF obfuscation format.
 */
export async function loadEmbeddedFonts(
  refs: readonly PptxEmbeddedFontRef[],
  fetchFontBytes: (partPath: string) => Promise<Uint8Array>,
): Promise<LoadedPptxEmbeddedFonts> {
  if (refs.length === 0) return {
    faces: [],
    aliases: new Map(),
    authoredFamilies: new Map(),
    tuples: new Set(),
    metrics: new Map(),
  };
  const scope = nextFontScope++;
  const candidateAliases = new Map<string, string>();
  for (const ref of refs) {
    const key = normalizedFamily(ref.fontName);
    if (!candidateAliases.has(key)) {
      candidateAliases.set(key, `__ooxml_pptx_${scope}_${candidateAliases.size + 1}`);
    }
  }
  // Retain no more than two unregistered WASM/transfer buffers at once. Each
  // batch is copied into FontFace storage before the next extraction begins.
  const loaded: FontFace[] = [];
  const held = new Set<FontFace>();
  // Parsed before registration hands the bytes to FontFace storage; only a
  // registered face's entry is kept below. One small record per font part.
  const parsed = new Map<string, PowerPointFaceMetrics>();
  const slotKey = (family: string, weight: string, style: string) =>
    `${normalizedFamily(family)}:${weight === 'bold' || weight === '700' ? 700 : 400}:${style === 'italic' ? 'italic' : 'normal'}`;
  const batchSize = 2;
  for (let offset = 0; offset < refs.length; offset += batchSize) {
    const faces = await Promise.all(refs.slice(offset, offset + batchSize).map(
      async (ref): Promise<EmbeddedFontFace | null> => {
        try {
          return {
            family: candidateAliases.get(normalizedFamily(ref.fontName)) as string,
            bytes: await fetchFontBytes(ref.partPath),
            odttf: false,
            weight: ref.style === 'bold' || ref.style === 'boldItalic' ? 'bold' : 'normal',
            style: ref.style === 'italic' || ref.style === 'boldItalic' ? 'italic' : 'normal',
          };
        } catch {
          return null;
        }
      },
    ));
    const loadable = faces.filter((face): face is EmbeddedFontFace => face !== null);
    if (loadable.length === 0) continue;
    for (const face of loadable) {
      const key = slotKey(face.family, String(face.weight ?? 'normal'), String(face.style ?? 'normal'));
      if (parsed.has(key)) continue;
      const tables = parseOpenTypeResourceMetrics(face.bytes);
      const metrics = tables ? powerPointResourceFaceMetrics(tables) : undefined;
      if (metrics) parsed.set(key, metrics);
    }
    for (const face of await registerEmbeddedFonts(loadable)) {
      if (held.has(face)) {
        // A content-identical face retained by an earlier batch needs no second
        // holder from this presentation. Balance that registry retain now.
        unregisterEmbeddedFonts([face]);
      } else {
        held.add(face);
        loaded.push(face);
      }
    }
  }
  const loadedAliases = new Set(loaded.map((face) => normalizedFamily(face.family)));
  const aliases = new Map(
    [...candidateAliases].filter(([, alias]) => loadedAliases.has(normalizedFamily(alias))),
  );
  const authoredFamilies = new Map(
    [...aliases].map(([authored, alias]) => [alias, authored]),
  );
  // An embedded regular face does not satisfy the bold or italic slots. Keep
  // this distinct from the family alias so a missing Calibri slot may use its
  // measured Office fallback without replacing the successful embedded face.
  const tuples = new Set(loaded.map((face) => {
    const authored = authoredFamilies.get(face.family) as string;
    return `${authored}:${face.weight === 'bold' || face.weight === '700' ? 700 : 400}:${face.style === 'italic' ? 'italic' : 'normal'}`;
  }));
  const metrics = new Map<string, PowerPointFaceMetrics>();
  for (const face of loaded) {
    const entry = parsed.get(slotKey(face.family, face.weight, face.style));
    if (!entry) continue;
    const authored = authoredFamilies.get(face.family) as string;
    metrics.set(`${authored}:${face.weight === 'bold' || face.weight === '700' ? 700 : 400}:${face.style === 'italic' ? 'italic' : 'normal'}`, entry);
  }
  return { faces: loaded, aliases, authoredFamilies, tuples, metrics };
}

/** Do not register a web substitute for a family successfully loaded from the deck. */
export function excludeEmbeddedFontFamilies(
  names: readonly (string | null)[],
  loadedAliases: ReadonlyMap<string, string>,
): (string | null)[] {
  return names.filter((name) => name === null || !loadedAliases.has(name.trim().toLowerCase()));
}
