/**
 * The package declaration builds (upstream 4fafb7ca, #992): core's `.d.ts` come from
 * scripts/build-core-declarations.mjs, and the React and Vue adapters read their siblings'
 * built declarations instead of compiling the sibling sources again.
 */
import { afterEach, describe, expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import {
  declarationCandidates,
  packageName,
  publishedEntries,
} from '../../../../scripts/build-core-declarations.mjs';
import { siblingsOf, staleness } from '../../../../scripts/check-built-siblings.mjs';
import { declarationCompilerOptions } from '../../../../scripts/declaration-options.mjs';

const packages = resolve(import.meta.dir, '..', '..', '..');
const configOf = (name: string, file: string) => pathToFileURL(join(packages, name, file));

describe('core declaration entries', () => {
  test('each export with a top-level types condition becomes one declaration entry', () => {
    const entries = publishedEntries({
      exports: {
        '.': { types: './dist/core.d.ts', import: './dist/core.mjs' },
        './prosemirror': { types: './dist/prosemirror/index.d.ts' },
        './prosemirror/extensions/nodes/TableExtension': {
          types: './dist/prosemirror/extensions/nodes/TableExtension.d.ts',
        },
        './prosemirror/editor.css': './dist/prosemirror/editor.css',
      },
    });
    expect(entries.map(({ specifier, name }) => [specifier, name])).toEqual([
      ['@eigenpal/docx-editor-core', 'core'],
      ['@eigenpal/docx-editor-core/prosemirror', 'prosemirror/index'],
      [
        '@eigenpal/docx-editor-core/prosemirror/extensions/nodes/TableExtension',
        'prosemirror/extensions/nodes/TableExtension',
      ],
    ]);
    expect(entries.map(({ source }) => source.slice(packages.length))).toEqual([
      '/core/src/core.ts',
      '/core/src/prosemirror/index.ts',
      '/core/src/prosemirror/extensions/nodes/TableExtension/index.ts',
    ]);
  });

  test('a subpath without a source or with nested types fails instead of shipping no types', () => {
    expect(() =>
      publishedEntries({ exports: { './layout': { types: './dist/layout.d.ts' } } })
    ).toThrow('neither src/layout.ts nor src/layout/index.ts exists');
    expect(() =>
      publishedEntries({
        exports: { './docx': { import: { types: './dist/docx/index.d.ts', default: './x.mjs' } } },
      })
    ).toThrow('nests its types condition');
  });

  test('a bare import is checked against the package that owns it', () => {
    expect(packageName('prosemirror-model')).toBe('prosemirror-model');
    expect(packageName('xml-js/lib/types')).toBe('xml-js');
    expect(packageName('@eigenpal/docx-editor-i18n/de')).toBe('@eigenpal/docx-editor-i18n');
  });

  test('an import keeps its module kind when it resolves to a declaration file', () => {
    expect(declarationCandidates('/o/foo.ts')).toEqual(['/o/foo.d.ts']);
    expect(declarationCandidates('/o/foo.js')).toEqual(['/o/foo.d.ts']);
    expect(declarationCandidates('/o/foo.mts')).toEqual(['/o/foo.d.mts']);
    expect(declarationCandidates('/o/foo.cjs')).toEqual(['/o/foo.d.cts']);
    expect(declarationCandidates('/o/foo')).toEqual(['/o/foo.d.ts', '/o/foo/index.d.ts']);
  });
});

describe('sibling declarations', () => {
  test('declaration builds drop sibling source paths and keep local aliases', () => {
    const react = declarationCompilerOptions(configOf('react', 'tsup.config.ts'));
    expect(Object.keys(react.paths)).toEqual(['@/*']);
    const vue = declarationCompilerOptions(configOf('vue', 'vite.config.ts'), {
      declarationMap: false,
    });
    expect(vue).toEqual({ declarationMap: false, paths: {} });
    // agents compiles core into its own output rather than depending on it.
    const agents = declarationCompilerOptions(configOf('agents', 'tsup.config.ts'));
    expect(Object.keys(agents.paths)).toContain('@eigenpal/docx-editor-core');
  });

  test('a package checks every sibling it reads, including those behind a sibling', () => {
    const siblings = (name: string) => [...siblingsOf(join(packages, name)).keys()].sort();
    const adapterSiblings = [
      '@eigenpal/docx-editor-agents',
      '@eigenpal/docx-editor-core',
      '@eigenpal/docx-editor-i18n',
    ];
    expect(siblings('react')).toEqual(adapterSiblings);
    expect(siblings('vue')).toEqual(adapterSiblings);
    expect(siblings('nuxt')).toEqual([...adapterSiblings, '@eigenpal/docx-editor-vue']);
    expect(siblings('core')).toEqual([]);
    expect(siblings('agents')).toEqual([]);
  });

  describe('a sibling build', () => {
    const dirs: string[] = [];
    afterEach(() => {
      for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
    });

    /** A package whose source and declaration were last written at the given times. */
    function sibling(sourceTime: number, builtTime?: number): string {
      const dir = mkdtempSync(join(tmpdir(), 'docx-sibling-'));
      dirs.push(dir);
      mkdirSync(join(dir, 'src'));
      writeFileSync(join(dir, 'package.json'), '{}');
      writeFileSync(join(dir, 'src', 'index.ts'), 'export const a = 1;');
      for (const path of [join(dir, 'package.json'), join(dir, 'src', 'index.ts')]) {
        utimesSync(path, sourceTime, sourceTime);
      }
      if (builtTime !== undefined) {
        mkdirSync(join(dir, 'dist'));
        writeFileSync(join(dir, 'dist', 'index.d.ts'), 'export declare const a = 1;');
        utimesSync(join(dir, 'dist', 'index.d.ts'), builtTime, builtTime);
      }
      return dir;
    }

    test('is current when its declarations are newer than its source', () => {
      expect(staleness('pkg', sibling(1000, 2000))).toBeNull();
    });

    test('names what is missing or stale', () => {
      expect(staleness('pkg', sibling(1000))).toBe('pkg has no built declarations');
      expect(staleness('pkg', sibling(2000, 1000))).toBe(
        "pkg's built declarations are older than its source"
      );
    });
  });
});
