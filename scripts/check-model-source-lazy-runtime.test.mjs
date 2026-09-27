import assert from 'node:assert/strict';
import test from 'node:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { runSourceRequestProbe } from './check-model-source-lazy-runtime.mjs';

test('synchronous getBuiltinModule/createRequire cannot bypass the runtime proof', () => {
  const temp = mkdtempSync(join(tmpdir(), 'ooxml-source-cjs-probe-'));
  try {
    writeFileSync(join(temp, 'model-source-probe.cjs'), 'module.exports = 1;\n');
    const script = join(temp, 'probe.mjs');
    writeFileSync(script,
      "process.getBuiltinModule('module').createRequire(import.meta.url)('./model-source-probe.cjs');\n");
    assert.throws(() => runSourceRequestProbe({ script }),
      /Default OOXML load requested optional source code:[\s\S]*model-source-probe\.cjs/);
  } finally {
    rmSync(temp, { recursive: true, force: true });
  }
});
