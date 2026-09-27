import { readFileSync } from 'node:fs';
import {
  materializeDocxDocument,
  materializeXlsxWorkbookIndex,
  materializePptxPresentation,
} from '../dist/node.mjs';

const sample = (format) => readFileSync(new URL(
  `../packages/${format}/public/demo/sample-1.${format}`, import.meta.url,
));

await materializeDocxDocument(sample('docx'));
await materializeXlsxWorkbookIndex(sample('xlsx'));
await materializePptxPresentation(sample('pptx'));
// Let any unawaited import posted during the ordinary load reach the loader.
await new Promise((resolve) => setTimeout(resolve, 50));
