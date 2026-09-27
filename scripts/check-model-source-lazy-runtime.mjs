#!/usr/bin/env node
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';

const temp = mkdtempSync(join(tmpdir(), 'ooxml-source-requests-'));
const log = join(temp, 'requests.log');
try {
  writeFileSync(log, '');
  const result = spawnSync(process.execPath, [
    '--loader', './scripts/model-source-runtime-loader.mjs',
    './scripts/model-source-runtime-probe.mjs',
  ], {
    cwd: process.cwd(), encoding: 'utf8', timeout: 120_000,
    env: { ...process.env, OOXML_SOURCE_REQUEST_LOG: log },
  });
  if (result.status !== 0) {
    throw new Error(`Built Node OOXML probe failed: ${result.stderr || result.stdout}`);
  }
  const requests = readFileSync(log, 'utf8').trim();
  if (requests) throw new Error(`Default OOXML load requested optional source code:\n${requests}`);
  console.log('Built Node OOXML loads requested no model-source modules.');
} finally {
  rmSync(temp, { recursive: true, force: true });
}
