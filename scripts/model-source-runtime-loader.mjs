// Node loader hook used only by the distribution regression check. Recording
// resolution catches a fire-and-forget import even when the caller never awaits it.
import { appendFileSync } from 'node:fs';

export async function resolve(specifier, context, nextResolve) {
  if (/model-source|source-worker/.test(specifier) && !specifier.includes('/scripts/')) {
    appendFileSync(process.env.OOXML_SOURCE_REQUEST_LOG, `${specifier}\n`);
  }
  return nextResolve(specifier, context);
}
