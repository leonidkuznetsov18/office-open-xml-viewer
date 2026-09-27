#!/usr/bin/env node
// A source owner is used only after a parse request selects a model source.
// Keep it out of the static dependency graph of both OOXML worker entries.
import { existsSync, readFileSync } from 'node:fs';
import { resolve, dirname, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parse } from '@babel/parser';

function resolveLocal(from, specifier) {
  if (!specifier.startsWith('.')) return undefined;
  const base = resolve(dirname(from), specifier);
  const roots = extname(base) === '.js' ? [base.slice(0, -3)] : [base];
  for (const root of roots) {
    for (const candidate of [root, `${root}.ts`, `${root}.tsx`, `${root}.js`, `${root}/index.ts`]) {
      if (existsSync(candidate)) return candidate;
    }
  }
  throw new Error(`Cannot resolve ${specifier} from ${from}`);
}

export function eagerModules(entry) {
  const visited = new Set();
  const pending = [resolve(entry)];
  while (pending.length) {
    const file = pending.pop();
    if (visited.has(file)) continue;
    visited.add(file);
    const ast = parse(readFileSync(file, 'utf8'), {
      sourceType: 'module', plugins: ['typescript', 'jsx'],
    });
    for (const statement of ast.program.body) {
      if (statement.type !== 'ImportDeclaration'
        && statement.type !== 'ExportNamedDeclaration'
        && statement.type !== 'ExportAllDeclaration') continue;
      if (!statement.source || statement.importKind === 'type' || statement.exportKind === 'type') continue;
      if (statement.type === 'ImportDeclaration' && statement.specifiers.length > 0
        && statement.specifiers.every((specifier) => specifier.importKind === 'type')) continue;
      const next = resolveLocal(file, statement.source.value);
      if (next) pending.push(next);
    }
  }
  return visited;
}

export function assertLazySourceOwner(entry, owner) {
  if (eagerModules(entry).has(resolve(owner))) {
    throw new Error(`${entry} statically imports ${owner}`);
  }
  const source = readFileSync(entry, 'utf8');
  if (!source.includes(`import('./internal/${owner.split('/').at(-1).replace(/\.ts$/, '')}`)) {
    throw new Error(`${entry} does not import its source owner lazily`);
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  for (const [format, owner] of [
    ['docx', 'worker-document-source'],
    ['xlsx', 'worker-worksheet-source'],
    ['pptx', 'worker-presentation-source'],
  ]) {
    for (const worker of ['worker.ts', 'render-worker.ts']) {
      const base = `packages/${format}/src`;
      assertLazySourceOwner(`${base}/${worker}`, `${base}/internal/${owner}.ts`);
    }
  }
  console.log('All six OOXML workers load source owners only on the selected-source path.');
}
