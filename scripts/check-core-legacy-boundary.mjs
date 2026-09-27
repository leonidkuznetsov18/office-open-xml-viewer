#!/usr/bin/env node
// Keep legacy Office support out of the OOXML packages.
//
// The legacy DOC/XLS/PPT readers are an opt-in module (packages/legacy-converter)
// that plugs into the renderers only through the format-generic ModelSource
// contract. Core, DOCX, XLSX, PPTX and the Node facade must therefore carry no
// legacy-specific names, branches, defaults, dependencies or imports: anything
// they need must be expressible as a generic contract or capability.
//
// Usage: node scripts/check-core-legacy-boundary.mjs [--ref <git-ref>]
// Without --ref the working tree's tracked files are checked.
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import nodePath from 'node:path';
import { parse } from '@babel/parser';
import ts from 'typescript-compiler-api';

export const GUARDED_ROOTS = [
  'packages/core/',
  'packages/docx/',
  'packages/xlsx/',
  'packages/pptx/',
  'packages/node/',
];

// Generated or local-only content inside the guarded roots.
const IGNORED = [
  /\/src\/wasm\//,
  /\/public\//,
  /\/tests\/visual\/(baseline|screenshots|diffs|references|report)\//,
];

const CHECKED_EXTENSIONS = /\.(?:[cm]?[jt]sx?|rs|json|toml|md|astro|html)$/;

/**
 * Each rule names a legacy-only concept. `legacy-binary-format` (the OOXML
 * loaders' typed rejection of a CFB container) and the container sniffer's
 * stream-name table predate the legacy module and stay allowed.
 */
export const RULES = [
  { id: 'legacy-package', pattern: /ooxml-legacy-converter|legacy-converter|legacy_office_converter/ },
  { id: 'legacy-format-name', pattern: /legacy[-_ ]?(?:doc|xls|ppt)(?![a-z])/i },
  { id: 'legacy-office-api', pattern: /LegacyOffice|legacyConversion|legacy-office|legacy_office/ },
  { id: 'legacy-sniffer', pattern: /sniffLegacyOfficeFormat|LegacyCfbFormat/ },
  { id: 'direct-format-name', pattern: /\bdirect[-_]?(?:doc|xls|ppt)(?![a-z])/i },
  { id: 'native-legacy-source', pattern: /native(?:Doc|Xls|Ppt)\b|nativeSource\b/ },
  { id: 'legacy-revision-view', pattern: /sourceRevisionView|sourceRevisionMarkup|revision_markup_in_print/ },
  { id: 'legacy-xls-measurement', pattern: /measureLegacy|xls-font|XLS_FONT_|configure_mdw|measurement_request/ },
  // Owner decision: DOCX w:lang/@w:val (langDefault) is not modeled.
  { id: 'docx-lang-default', pattern: /langDefault|lang_default/ },
];

// The only computed imports in production source are the six worker-realm
// sidecars. The main thread supplies these absolute URLs after it selects a
// model source; document bytes cannot choose them. Keep this list exact so a
// new computed import or require cannot evade module resolution.
export const ALLOWED_COMPUTED_IMPORTS = new Map([
  ['packages/core/src/math/engine-runtime.ts', ['src']],
  ['packages/core/src/source/model-source.ts', ['module.moduleUrl']],
  ['packages/docx/src/worker-source.ts', ['req.sourceOwnerUrl']],
  ['packages/docx/src/render-worker-source.ts', ['req.sourceOwnerUrl']],
  ['packages/xlsx/src/worker-source.ts', ['req.sourceOwnerUrl']],
  ['packages/xlsx/src/render-worker-source.ts', ['req.sourceOwnerUrl']],
  ['packages/pptx/src/worker-source.ts', ['request.sourceOwnerUrl']],
  ['packages/pptx/src/render-worker-source.ts', ['request.sourceOwnerUrl']],
  // Local benchmark CLIs load the caller-specified build they are measuring.
  ['packages/node/src/bench-handle.mjs', ['jsPath']],
  ['packages/node/src/bench-parse.mjs', ['resolve(HERE, relJs)']],
  // Tests load a selected fixture or module after setting up mocks. These
  // files are never package entries, but the allowance is still path/expression
  // specific so a new computed import is reviewed.
  ['packages/core/src/source/model-source.test.ts', ['url']],
  ['packages/node/src/docx-find-highlight.test.ts', ['FIND_PATH', 'HIGHLIGHT_PATH']],
  ['packages/node/src/docx-vertical-tr-ink-overlap.probe.test.ts', ['PLAYWRIGHT', 'ESBUILD']],
  ['packages/node/src/pptx-find-highlight.test.ts', ['RENDERER_PATH', 'FIND_PATH', 'HIGHLIGHT_PATH']],
  ['packages/node/src/verify-chart-regressions.probe.test.ts', ['CORE_RENDERER']],
  ['packages/node/src/xlsx-blip-duotone-alpha.probe.test.ts', ['ORCH_PATH']],
  ['packages/node/src/xlsx-border-crisp.probe.test.ts', ['ORCH_PATH']],
  ['packages/node/src/xlsx-find.test.ts', ['FIND_PATH', 'NUMFMT_PATH']],
  ['packages/node/src/xlsx-merge-border-zorder.probe.test.ts', ['ORCH_PATH']],
]);

// The Node WASM locator needs resolution only. It never loads a package and
// its exact call shape is checked below; no other require factory is allowed.
export const ALLOWED_INDIRECT_REQUIRE = new Map([
  ['packages/node/src/wasm-loader.ts', 'createRequire(metaUrl).resolve(workspaceSpecifier)'],
]);

function expressionName(node) {
  if (node?.type === 'Identifier') return node.name;
  if (node?.type === 'MemberExpression' && !node.computed && node.property.type === 'Identifier') {
    const owner = expressionName(node.object);
    return owner && `${owner}.${node.property.name}`;
  }
  if (node?.type === 'CallExpression' && node.callee.type === 'Identifier'
    && node.callee.name === 'resolve' && node.arguments.length === 2
    && node.arguments.every((argument) => argument.type === 'Identifier')) {
    return `resolve(${node.arguments.map((argument) => argument.name).join(', ')})`;
  }
  return undefined;
}

function literalValue(node) {
  if (node?.type === 'StringLiteral') return node.value;
  if (node?.type === 'TemplateLiteral' && node.expressions.length === 0) {
    return node.quasis[0].value.cooked;
  }
  return undefined;
}

function isLegacyModule(path) {
  return /ooxml-legacy-converter|(?:^|[\\/])legacy-converter(?:[\\/]|$)/.test(path);
}

/** Use the same compiler options and module resolver as TypeScript's program. */
export function createTypeScriptResolver(root = process.cwd()) {
  const configs = new Map();
  const resolver = (importer, specifier) => {
    const absolute = nodePath.resolve(root, importer);
    const packageName = importer.startsWith('packages/') ? importer.split('/')[1] : undefined;
    const configPath = packageName
      ? nodePath.resolve(root, 'packages', packageName, 'tsconfig.json')
      : nodePath.resolve(root, 'tsconfig.json');
    let options = configs.get(configPath);
    if (!options) {
      if (existsSync(configPath)) {
        const config = ts.readConfigFile(configPath, ts.sys.readFile);
        if (config.error) throw new Error(ts.flattenDiagnosticMessageText(config.error.messageText, '\n'));
        options = ts.parseJsonConfigFileContent(config.config, ts.sys, nodePath.dirname(configPath)).options;
      } else {
        // A reached file can live outside the guarded package. Resolve its
        // imports with the importing program's already loaded options.
        options = configs.values().next().value;
        if (!options) throw new Error(`No TypeScript program for ${importer}`);
      }
      configs.set(configPath, options);
    }
    return ts.resolveModuleName(specifier, absolute, options, ts.sys).resolvedModule?.resolvedFileName;
  };
  resolver.root = nodePath.resolve(root);
  return resolver;
}

export function findViolations(files, { resolveModule } = {}) {
  const violations = [];
  const queue = [...files];
  const seen = new Set();
  const root = resolveModule?.root ?? process.cwd();
  while (queue.length > 0) {
    const { path, text } = queue.shift();
    if (seen.has(path)) continue;
    seen.add(path);
    const lines = text.split('\n');
    if (isGuardedPath(path)) lines.forEach((line, index) => {
      for (const rule of RULES) {
        if (rule.pattern.test(line)) {
          violations.push({ path, line: index + 1, rule: rule.id, text: line.trim().slice(0, 160) });
        }
      }
    });
    if (/\.[cm]?[jt]sx?$/.test(path)) {
      const source = parse(text, { sourceType: 'unambiguous', errorRecovery: true, plugins: ['typescript', 'jsx'] });
      const allowedExpressions = ALLOWED_COMPUTED_IMPORTS.get(path) ?? [];
      const usedExpressions = new Set();
      const indirectAllowed = ALLOWED_INDIRECT_REQUIRE.get(path);
      const inspect = (node, parent, grandparent) => {
        const line = node.loc?.start.line ?? 1;
        const rejectIndirect = () => violations.push({
          path, line, rule: 'indirect-require', text: lines[line - 1].trim().slice(0, 160),
        });
        if (node.type === 'ImportSpecifier' && node.imported.name === 'createRequire'
          && (indirectAllowed === undefined || node.local.name !== 'createRequire')) rejectIndirect();
        if (node.type === 'ObjectProperty' && node.key.type === 'Identifier'
          && node.key.name === 'createRequire' && parent?.type === 'ObjectPattern') rejectIndirect();
        if (['MemberExpression', 'OptionalMemberExpression'].includes(node.type) && (
          (node.object.type === 'Identifier' && node.object.name === 'require')
          || (node.property.type === 'Identifier' && node.property.name === 'createRequire')
          || (node.property.type === 'StringLiteral' && node.property.value === 'createRequire')
          || ((node.object.type === 'Identifier'
            && ['module', 'globalThis'].includes(node.object.name))
            && ((node.property.type === 'Identifier' && node.property.name === 'require')
              || (node.property.type === 'StringLiteral' && node.property.value === 'require')))
        )) rejectIndirect();
        if (node.type === 'CallExpression' && node.callee.type === 'Identifier'
          && node.callee.name === 'createRequire') {
          const approved = indirectAllowed !== undefined
            && parent?.type === 'MemberExpression' && parent.object === node
            && parent.property.type === 'Identifier' && parent.property.name === 'resolve'
            && grandparent?.type === 'CallExpression' && grandparent.callee === parent
            && node.arguments.length === 1 && expressionName(node.arguments[0]) === 'metaUrl'
            && grandparent.arguments.length === 1
            && expressionName(grandparent.arguments[0]) === 'workspaceSpecifier';
          if (!approved) rejectIndirect();
        }
        if (node.type === 'OptionalCallExpression'
          && node.callee.type === 'Identifier' && node.callee.name === 'require') rejectIndirect();
        if (node.type === 'Identifier' && node.name === 'createRequire'
          && parent?.type !== 'ImportSpecifier'
          && !(parent?.type === 'CallExpression' && parent.callee === node)
          && !(parent?.type === 'MemberExpression' && parent.property === node)) rejectIndirect();
        if (node.type === 'Identifier' && node.name === 'require'
          && !(parent?.type === 'CallExpression' && parent.callee === node)
          && !(parent?.type === 'MemberExpression' && (parent.object === node || parent.property === node))
          && !(parent?.type === 'ClassMethod' && parent.key === node)
          && !(parent?.type === 'OptionalCallExpression' && parent.callee === node)) rejectIndirect();
        let literal;
        let dynamic = false;
        if (['ImportDeclaration', 'ExportNamedDeclaration', 'ExportAllDeclaration'].includes(node.type)) {
          literal = node.source;
        } else if (node.type === 'TSImportEqualsDeclaration'
          && node.moduleReference?.type === 'TSExternalModuleReference') {
          literal = node.moduleReference.expression;
        } else if (node.type === 'TSImportType') {
          literal = node.argument?.type === 'TSLiteralType' ? node.argument.literal : node.argument;
        } else if (node.type === 'CallExpression'
          && (node.callee.type === 'Import'
            || (node.callee.type === 'Identifier' && node.callee.name === 'require'))) {
          [literal] = node.arguments;
          dynamic = true;
        }
        const specifier = literalValue(literal);
        if (dynamic && specifier === undefined) {
          const expression = expressionName(literal);
          const approved = expression !== undefined
            && node.callee.type === 'Import'
            && node.arguments.length === 1
            && allowedExpressions.includes(expression)
            && !usedExpressions.has(expression);
          if (approved) usedExpressions.add(expression);
          if (!approved) {
            violations.push({ path, line: node.loc.start.line, rule: 'computed-module', text: text.split('\n')[node.loc.start.line - 1].trim().slice(0, 160) });
          }
        }
        if (typeof specifier === 'string') {
          // The parser has already decoded string escapes here. Normalize
          // relative paths, then also check TypeScript's actual resolution for
          // aliases, package exports, and type-only imports.
          const resolved = specifier.startsWith('.')
            ? nodePath.posix.normalize(nodePath.posix.join(nodePath.posix.dirname(path), specifier))
            : specifier;
          const compilerResolved = resolveModule?.(path, specifier);
          if (isLegacyModule(resolved) || (compilerResolved && isLegacyModule(compilerResolved))) {
            const line = literal.loc.start.line;
            if (!violations.some((entry) => entry.path === path && entry.line === line && entry.rule === 'legacy-package')) {
              violations.push({ path, line, rule: 'legacy-package', text: specifier.slice(0, 160) });
            }
          }
          if (compilerResolved && !isLegacyModule(compilerResolved)
            && /\.[cm]?[jt]sx?$/.test(compilerResolved)
            && !compilerResolved.includes(`${nodePath.sep}node_modules${nodePath.sep}`)
            && (compilerResolved === root || compilerResolved.startsWith(root + nodePath.sep))) {
            const child = nodePath.relative(root, compilerResolved).replaceAll('\\', '/');
            if (!seen.has(child) && existsSync(compilerResolved)) {
              queue.push({ path: child, text: readFileSync(compilerResolved, 'utf8') });
            }
          }
        }
        for (const value of Object.values(node)) {
          if (Array.isArray(value)) {
            for (const child of value) {
              if (child && typeof child === 'object' && typeof child.type === 'string') inspect(child, node, parent);
            }
          } else if (value && typeof value === 'object' && typeof value.type === 'string') {
            inspect(value, node, parent);
          }
        }
      };
      inspect(source);
    }
  }
  return violations;
}

export function isGuardedPath(path) {
  return GUARDED_ROOTS.some((root) => path.startsWith(root))
    && CHECKED_EXTENSIONS.test(path)
    && !IGNORED.some((pattern) => pattern.test(path));
}

function trackedFiles(ref) {
  const git = (args) => execFileSync('git', args, { encoding: 'utf8', maxBuffer: 256 * 1024 * 1024 });
  const paths = (ref
    ? git(['ls-tree', '-r', '--name-only', ref, '--', ...GUARDED_ROOTS])
    : git(['ls-files', '--', ...GUARDED_ROOTS]))
    .split('\n')
    .filter(Boolean)
    .filter(isGuardedPath)
    // A tracked file deleted in the working tree is not part of the tree checked.
    .filter((path) => ref || existsSync(path));
  return paths.map((path) => ({
    path,
    text: ref ? git(['show', `${ref}:${path}`]) : readFileSync(path, 'utf8'),
  }));
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const refIndex = process.argv.indexOf('--ref');
  const ref = refIndex >= 0 ? process.argv[refIndex + 1] : undefined;
  // --ref checks lexical/AST rules at that revision. The live compiler graph
  // is checked for the working tree, where its tsconfig and dependencies exist.
  const violations = findViolations(trackedFiles(ref), ref ? {} : { resolveModule: createTypeScriptResolver() });
  if (violations.length > 0) {
    for (const violation of violations.slice(0, 200)) {
      console.error(`${violation.path}:${violation.line} [${violation.rule}] ${violation.text}`);
    }
    console.error(`\n${violations.length} legacy-specific reference(s) in the OOXML packages.`);
    console.error('Express the need as a format-generic contract (see core source/model-source.ts).');
    process.exit(1);
  }
  console.log('OOXML packages carry no legacy-specific names, branches or imports.');
}
