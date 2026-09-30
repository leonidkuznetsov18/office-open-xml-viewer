import { defineConfig } from 'vitest/config';

export default defineConfig({
  define: { __OOXML_MODEL_SOURCES__: 'true' },
  test: { include: ['packages/docx/tests/differential/tab-fitting.test.ts'] },
});
