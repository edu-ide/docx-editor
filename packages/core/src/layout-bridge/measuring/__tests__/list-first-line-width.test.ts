/**
 * A list paragraph's first line is measured from where its text starts.
 *
 * The painter draws the list marker as an inline-block slot at the start of
 * the first line — at `left + firstLine`, or at `left - hanging` — and the
 * text after it. The measurer subtracted that slot only on the firstLine path,
 * so a hanging list's first line got `hanging` px more text than the painter
 * has room for, and a full first line ran past the right indent.
 */
import { describe, expect, test } from 'bun:test';

// Deterministic canvas stub, installed only when no other suite file has set up
// a document; widths below derive from measureTextWidth (see footnote-ref-glue.test.ts).
if (typeof document === 'undefined') {
  (globalThis as Record<string, unknown>).document = {
    createElement: () => ({
      getContext: () => ({
        font: '',
        measureText: (text: string) => ({ width: text.length * 8 }),
      }),
    }),
  };
}

import { measureParagraph } from '../measureParagraph';
import { measureTextWidth } from '../measureContainer';
import type { ParagraphBlock } from '../../../layout-engine/types';

const CHAR = measureTextWidth('x', { fontFamily: 'Calibri', fontSize: 11 });
const WIDTH = 50 * CHAR;
const TEXT = Array(60).fill('aaa').join(' ');

function paragraph(attrs: ParagraphBlock['attrs']): ParagraphBlock {
  return { kind: 'paragraph', id: 'p', runs: [{ kind: 'text', text: TEXT }], attrs };
}

/** Characters each line holds. */
function lineChars(block: ParagraphBlock): number[] {
  return measureParagraph(block, WIDTH).lines.map((line) => line.toChar - line.fromChar);
}

describe('first-line width of a list paragraph', () => {
  const indent = { left: 12 * CHAR, hanging: 6 * CHAR };

  test('a hanging list gives its first line the body width: text starts at the indent', () => {
    const lines = lineChars(paragraph({ listMarker: '1.', indent }));
    // 38 characters of body width hold nine "aaa " words.
    expect(lines[0]).toBe(36);
    expect(lines[0]).toBe(lines[1]);
    for (const line of measureParagraph(paragraph({ listMarker: '1.', indent }), WIDTH).lines) {
      expect(line.width).toBeLessThanOrEqual(WIDTH - indent.left);
    }
  });

  test('a hanging paragraph without a marker keeps its wider first line', () => {
    const lines = lineChars(paragraph({ indent }));
    // 38 + 6 characters: eleven words on the first line.
    expect(lines[0]).toBe(44);
    expect(lines[1]).toBe(36);
  });

  test('a hidden marker takes no slot, so the first line keeps the hang', () => {
    const lines = lineChars(paragraph({ listMarker: '1.', listMarkerHidden: true, indent }));
    expect(lines[0]).toBe(44);
  });

  test('a firstLine list still loses the marker slot from its first line', () => {
    const lines = lineChars(
      paragraph({ listMarker: '1.', listMarkerSuffix: 'space', indent: { firstLine: 4 * CHAR } })
    );
    // 50 - 4 (firstLine) - 3 ("1. ") = 43 characters: ten words.
    expect(lines[0]).toBe(40);
    expect(lines[1]).toBe(48);
  });
});
