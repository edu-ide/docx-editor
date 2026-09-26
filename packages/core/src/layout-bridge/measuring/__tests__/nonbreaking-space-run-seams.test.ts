/**
 * A no-break character glues its neighbours whether or not a run boundary sits
 * beside it (upstream ae1afe03, #976).
 *
 * U+00A0, U+202F, U+2007 and the word joiners U+2060 / U+FEFF never offer a
 * line break inside one run. A generator that writes "Sec." + a run holding
 * only U+00A0 + "12" must lay out exactly like the one-run text. The cross-run
 * glue estimate used to measure the continuation token without its trailing
 * space while the word loop measured it with the space, so on a line where
 * only the untrimmed token overflowed, the line broke at the run seam — before
 * or after the glue. Ordinary spaces, tabs and hyphens keep their seam breaks.
 */
import { describe, expect, test } from 'bun:test';

// Deterministic canvas stub, installed only when no other suite file has set up
// a document. Thresholds derive from measureTextWidth so the test holds under
// whichever monospace stub is active (see footnote-ref-glue.test.ts).
if (typeof document === 'undefined') {
  (globalThis as Record<string, unknown>).document = {
    createElement: () => ({
      getContext: () => ({
        font: '',
        measureText: (text: string) => ({ width: text.length * 6 }),
      }),
    }),
  };
}

import { measureParagraph } from '../measureParagraph';
import { measureTextWidth } from '../measureContainer';
import type { ParagraphBlock, Run } from '../../../layout-engine/types';

const CHAR = measureTextWidth('x', { fontFamily: 'Calibri', fontSize: 11 });

const NBSP = ' ';
const NNBSP = ' ';
const FIGURE = ' ';
const BOM = '﻿';
const WORD_JOINER = '⁠';
const GLUE = [NBSP, NNBSP, FIGURE, BOM, WORD_JOINER];
const name = (ch: string) => `U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`;

/** Lay out `texts` as one text run each and return each line's text. */
function linesOf(texts: string[], chars: number): string[] {
  const runs = texts.map((text) => ({ kind: 'text', text }) as Run);
  const block = { kind: 'paragraph', id: 'p', runs } as ParagraphBlock;
  return measureParagraph(block, chars * CHAR).lines.map((line) => {
    let text = '';
    for (let r = line.fromRun; r <= line.toRun; r++) {
      const runText = texts[r] ?? '';
      const from = r === line.fromRun ? line.fromChar : 0;
      const to = r === line.toRun ? line.toChar : runText.length;
      text += runText.slice(from, to);
    }
    return text;
  });
}

/** Seam layouts of `before + glue + after`. */
function seamShapes(before: string, glue: string, after: string): Record<string, string[]> {
  return {
    'glue in its own run': [before, glue, after],
    'glue ends the first run': [before + glue, after],
    'glue starts the second run': [before, glue + after],
    'doubled glue in one-character runs': [before, glue, glue, after],
  };
}

function brokenAtGlue(lines: readonly string[], glue: string): boolean {
  return lines.some(
    (line, index) =>
      index + 1 < lines.length && (line.endsWith(glue) || lines[index + 1]!.startsWith(glue))
  );
}

describe('no-break characters at run seams', () => {
  for (const glue of GLUE) {
    test(`${name(glue)} wraps the same across runs as inside one run`, () => {
      for (const [shape, texts] of Object.entries(seamShapes('aa bbbb', glue, 'cc dd ee'))) {
        // From nine characters, where "bbbb", the glue and "cc " fit a line of their own.
        for (let chars = 9; chars <= 16; chars++) {
          const oneRun = linesOf([texts.join('')], chars);
          const split = linesOf(texts, chars);
          expect({ shape, chars, lines: split }).toEqual({ shape, chars, lines: oneRun });
          expect(brokenAtGlue(split, glue)).toBe(false);
        }
      }
    });
  }

  test('a double U+00A0 after a label stays with the label and the next word', () => {
    // The shape a numbered legal clause takes: "(a)", two no-break spaces, then the text.
    const lines = linesOf(['xxxxx (a)', NBSP, NBSP, 'Next word'], 10);
    expect(lines).toEqual(['xxxxx ', `(a)${NBSP}${NBSP}Next `, 'word']);
    expect(linesOf(['xxxxx (a)', NBSP + NBSP, 'Next word'], 10)).toEqual(lines);
    expect(linesOf([`xxxxx (a)${NBSP}${NBSP}Next word`], 10)).toEqual(lines);
  });

  test('a glued group that does not fit moves to the next line as a whole', () => {
    const expected = ['aaaa ', `bbbb${NBSP}cc`];
    expect(linesOf(['aaaa bbbb', NBSP, 'cc'], 10)).toEqual(expected);
    expect(linesOf(['aaaa bbbb' + NBSP, 'cc'], 10)).toEqual(expected);
    expect(linesOf(['aaaa bbbb', NBSP + 'cc'], 10)).toEqual(expected);
  });

  test('a format-only split inside a word wraps like the unsplit word', () => {
    for (let chars = 9; chars <= 16; chars++) {
      expect(linesOf(['aa bbbb', 'x', 'cc dd ee'], chars)).toEqual(
        linesOf(['aa bbbbxcc dd ee'], chars)
      );
    }
  });
});

describe('seam break opportunities that must remain', () => {
  test('an ordinary space in its own run still breaks', () => {
    expect(linesOf(['aaaa bbbb', ' ', 'cc'], 10)).toEqual(['aaaa bbbb ', 'cc']);
  });

  test('an ordinary space after the glue still breaks', () => {
    expect(linesOf([`aaaa bbb${NBSP}`, ' ', 'cc'], 10)).toEqual([`aaaa bbb${NBSP} `, 'cc']);
    expect(linesOf([`aaaa bbb${NBSP} cc`], 10)).toEqual([`aaaa bbb${NBSP} `, 'cc']);
  });

  test('a tab after the glue still lets the next text open a line', () => {
    expect(linesOf([`aa${NBSP}bb`, '\t', 'cccccc'], 10)).toEqual([`aa${NBSP}bb\t`, 'cccccc']);
  });

  test('a hyphen at the end of a run still breaks', () => {
    expect(linesOf(['aaaa bbb-', 'cccc'], 10)).toEqual(linesOf(['aaaa bbb-cccc'], 10));
  });
});
