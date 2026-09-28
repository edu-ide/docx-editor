// Emits packages/core/dist/**/*.d.ts: one declaration file per published subpath, with the
// declarations those subpaths share split into chunks.
//
// Two steps, because the one-step bundler cost several times the memory. tsup's `dts: true`
// hands TypeScript sources to rollup-plugin-dts, which keeps the whole program alive while
// it emits and re-parses every module through rollup. For core that ran out of Node's
// default heap (ERR_WORKER_OUT_OF_MEMORY), so every caller had to know to raise it. Here
// TypeScript emits the declarations once, and rollup-plugin-dts then bundles `.d.ts` files
// only, which needs no TypeScript program at all.
//
// The subpath list is not written down again here. `exports` in package.json says which
// declaration file each subpath ships, `./dist/<name>.d.ts`, and that file compiles from the
// tsup entry of the same name: `src/<name>.ts` or `src/<name>/index.ts`, the rule
// scripts/lib/api-extractor-runner.mjs resolves sources by. exports-map.test.ts holds the
// tsup entries to it.

import { existsSync, mkdtempSync, readdirSync, readFileSync, realpathSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, isAbsolute, join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { rollup } from 'rollup';
import dts from 'rollup-plugin-dts';
import ts from 'typescript';

const core = join(dirname(fileURLToPath(import.meta.url)), '..', 'packages', 'core');
const src = join(core, 'src');
const dist = join(core, 'dist');
const PACKAGE = '@eigenpal/docx-editor-core';

/** Every published subpath: its rollup entry name and its source file under `root`. */
export function publishedEntries(manifest, root = core) {
  const entries = [];
  for (const [subpath, target] of Object.entries(manifest.exports)) {
    const types = typeof target === 'object' && target !== null ? target.types : undefined;
    if (typeof types !== 'string') {
      // Only a top-level `types` condition is read. A nested one would ship no declarations.
      if (JSON.stringify(target).includes('"types"'))
        throw new Error(`${subpath} nests its types condition; put \`types\` at the top level.`);
      continue;
    }
    const specifier = subpath === '.' ? PACKAGE : `${PACKAGE}${subpath.slice(1)}`;
    const name = types.replace(/^\.\/dist\//, '').replace(/\.d\.ts$/, '');
    const source = [`${name}.ts`, join(name, 'index.ts')]
      .map((file) => join(root, 'src', file))
      .find((file) => existsSync(file));
    if (!source) {
      throw new Error(
        `${specifier} ships ${types} but neither src/${name}.ts nor src/${name}/index.ts exists.`
      );
    }
    entries.push({ specifier, name, source });
  }
  if (entries.length === 0) throw new Error('package.json exports name no declaration files.');
  return entries;
}

/** Emit declarations for everything the entries reach, into `outDir`. */
function emitDeclarations(entries, outDir, parsedOptions) {
  const options = {
    ...parsedOptions,
    noEmit: false,
    declaration: true,
    emitDeclarationOnly: true,
    declarationMap: false,
    // As tsup did: a type error anywhere in the program fails the build, not only an
    // error the declaration emitter itself reports.
    noEmitOnError: true,
    rootDir: src,
    outDir,
  };
  // Rooted at the entries, not at tsconfig's `include`: test files and anything no subpath
  // reaches stay out of the program.
  const program = ts.createProgram(
    entries.map((entry) => entry.source),
    options
  );
  const { diagnostics, emitSkipped } = program.emit(undefined, undefined, undefined, true);
  const errors = diagnostics.filter((d) => d.category === ts.DiagnosticCategory.Error);
  if (emitSkipped || errors.length > 0) {
    throw new Error(`Declaration emit failed:\n${ts.formatDiagnostics(errors, formatHost)}`);
  }
}

const formatHost = {
  getCanonicalFileName: (fileName) => fileName,
  getCurrentDirectory: () => core,
  getNewLine: () => '\n',
};

/** `@scope/name/sub` → `@scope/name`, `name/sub` → `name`. */
export function packageName(specifier) {
  const parts = specifier.split('/');
  return specifier.startsWith('@') ? parts.slice(0, 2).join('/') : parts[0];
}

/** Remove the declarations of an earlier run: chunk names carry hashes, so none is overwritten. */
function removeDeclarations(dir) {
  if (!existsSync(dir)) return;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) removeDeclarations(path);
    else if (entry.name.endsWith('.d.ts')) rmSync(path);
  }
}

/**
 * The emitted declaration files an import of `target` can mean, most specific first. An
 * extension keeps its module kind: `.mts` and `.mjs` map to `.d.mts`, `.cts` and `.cjs`
 * to `.d.cts`. A path without one can name a file or a directory index.
 */
export function declarationCandidates(target) {
  const match = /(?:\.d)?\.(m|c)?(?:ts|tsx|js|jsx)$/.exec(target);
  if (!match) return [`${target}.d.ts`, join(target, 'index.d.ts')];
  return [`${target.slice(0, match.index)}.d.${match[1] ?? ''}ts`];
}

/** The emitted declaration file for a source file. */
function declarationFor(outDir, sourceFile) {
  return join(outDir, relative(src, sourceFile)).replace(/\.([cm]?)tsx?$/, '.d.$1ts');
}

/**
 * Resolve the imports inside the emitted declarations.
 *
 * TypeScript keeps each specifier as written, so relative imports resolve to the emitted
 * files and bundle like any other module, as does an import of one of the package's own
 * published subpaths. Every other bare specifier stays an import, and must be a dependency
 * or peer dependency.
 */
function declarationResolver(entries, outDir) {
  const manifest = JSON.parse(readFileSync(join(core, 'package.json'), 'utf8'));
  const declared = new Set([
    ...Object.keys(manifest.dependencies ?? {}),
    ...Object.keys(manifest.peerDependencies ?? {}),
  ]);
  const selfImports = new Map(
    entries.map((entry) => [entry.specifier, declarationFor(outDir, entry.source)])
  );
  return {
    name: 'core-declaration-resolver',
    resolveId(source, importer) {
      const self = selfImports.get(source);
      if (self) return self;
      if (source.startsWith('.') && importer) {
        const target = resolve(dirname(importer), source);
        for (const candidate of declarationCandidates(target)) {
          if (existsSync(candidate)) return candidate;
        }
        throw new Error(`Cannot resolve ${source} from ${relative(outDir, importer)}.`);
      }
      if (!isAbsolute(source)) {
        const name = packageName(source);
        // A core subpath that is not published would leave consumers an import they cannot
        // resolve.
        if (name === PACKAGE)
          throw new Error(`The declarations import ${source}, which package.json does not export.`);
        if (!declared.has(name))
          throw new Error(
            `The declarations import ${source}, but ${name} is not a dependency or peer ` +
              'dependency of the package, so consumers would not have its types.'
          );
        return { id: source, external: true };
      }
      return null;
    },
  };
}

async function bundleDeclarations(entries, outDir) {
  const bundle = await rollup({
    input: Object.fromEntries(
      entries.map((entry) => [entry.name, declarationFor(outDir, entry.source)])
    ),
    plugins: [declarationResolver(entries, outDir), dts()],
  });
  try {
    removeDeclarations(dist);
    await bundle.write({
      dir: dist,
      format: 'es',
      entryFileNames: '[name].d.ts',
      // A chunk takes its name from its first module, `document.d.ts`, so drop the `.d`.
      chunkFileNames: (chunk) => `${chunk.name.replace(/\.d$/, '')}-[hash].d.ts`,
    });
  } finally {
    await bundle.close();
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(realpathSync(process.argv[1])).href) {
  // The effective compiler options, read the way TypeScript reads them: comments, `extends`,
  // errors.
  const { config, error } = ts.readConfigFile(join(core, 'tsconfig.json'), ts.sys.readFile);
  if (error) throw new Error(ts.formatDiagnostic(error, formatHost));
  const parsed = ts.parseJsonConfigFileContent(config, ts.sys, core);
  if (parsed.errors.length > 0)
    throw new Error(
      `tsconfig.json is invalid:\n${ts.formatDiagnostics(parsed.errors, formatHost)}`
    );
  const entries = publishedEntries(JSON.parse(readFileSync(join(core, 'package.json'), 'utf8')));
  const outDir = mkdtempSync(join(tmpdir(), 'docx-core-declarations-'));
  try {
    emitDeclarations(entries, outDir, parsed.options);
    await bundleDeclarations(entries, outDir);
    const missing = entries.filter((entry) => !existsSync(join(dist, `${entry.name}.d.ts`)));
    if (missing.length > 0)
      throw new Error(`No declarations written for ${missing.map((e) => e.specifier).join(', ')}.`);
  } finally {
    rmSync(outDir, { recursive: true, force: true });
  }
}
