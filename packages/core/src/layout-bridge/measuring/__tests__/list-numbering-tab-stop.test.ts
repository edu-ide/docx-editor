/**
 * The `w:suff="tab"` advance of a hanging list whose marker ends before the
 * text indent (§17.9.25, upstream bf776f2d #979).
 *
 * The hanging indent is an implicit stop for that tab, but not the only one:
 * an authored stop between the marker and the indent is nearer, so the first
 * line's text starts there and the line gets the extra width. Continuation
 * lines keep the indent.
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

import { getListMarkerInlineWidth } from '../listMarkerWidth';
import { measureParagraph } from '../measureParagraph';
import { measureTextWidth } from '../measureContainer';
import type { ParagraphBlock, TabStop } from '../../../layout-engine/types';

const CHAR = measureTextWidth('x', { fontFamily: 'Calibri', fontSize: 11 });
/** Pixels to twips (tab stop positions are twips). */
const tw = (px: number) => px * 15;

// Marker slot at 6 characters (left 12, hanging 6), so `1.` ends at 8 characters.
const LEFT = 12 * CHAR;
const HANGING = 6 * CHAR;
const MARKER_START = LEFT - HANGING;
const MARKER_END = MARKER_START + 2 * CHAR;
const BETWEEN = 10 * CHAR;

function item(
  tabs: TabStop[] = [],
  extra: Partial<NonNullable<ParagraphBlock['attrs']>> = {},
  text = 'Text'
): ParagraphBlock {
  return {
    kind: 'paragraph',
    id: 'p',
    runs: [{ kind: 'text', text }],
    attrs: { listMarker: '1.', indent: { left: LEFT, hanging: HANGING }, tabs, ...extra },
  };
}

const stop = (px: number, val: TabStop['val'] = 'start'): TabStop => ({ val, pos: tw(px) });

/** Where the first line's text starts. */
const textStart = (block: ParagraphBlock) => MARKER_START + getListMarkerInlineWidth(block);

describe('a numbering tab stops at an authored stop before the text indent', () => {
  test('without a stop, the text starts at the indent', () => {
    expect(textStart(item())).toBe(LEFT);
  });

  test('a stop between the marker and the indent wins', () => {
    expect(textStart(item([stop(BETWEEN)]))).toBe(BETWEEN);
  });

  test('the nearest of several stops wins, in any authored order', () => {
    expect(textStart(item([stop(11 * CHAR), stop(BETWEEN)]))).toBe(BETWEEN);
  });

  test('a stop the marker already covers is skipped', () => {
    expect(textStart(item([stop(MARKER_START + CHAR)]))).toBe(LEFT);
    expect(textStart(item([stop(MARKER_START + CHAR), stop(BETWEEN)]))).toBe(BETWEEN);
    // A stop exactly at the marker end is not PAST it.
    expect(textStart(item([stop(MARKER_END)]))).toBe(LEFT);
  });

  test('clear and bar stops are not destinations', () => {
    expect(textStart(item([stop(BETWEEN, 'clear')]))).toBe(LEFT);
    expect(textStart(item([stop(BETWEEN, 'bar')]))).toBe(LEFT);
  });

  test('default-interval stops before the indent do not compete', () => {
    // A one-character grid puts default stops between the marker and the indent.
    expect(textStart(item([], { defaultTabStopTwips: tw(CHAR) }))).toBe(LEFT);
  });

  test('a stop at or past the indent does not pull the text right', () => {
    expect(textStart(item([stop(LEFT)]))).toBe(LEFT);
    expect(textStart(item([stop(18 * CHAR)]))).toBe(LEFT);
  });

  test('space and nothing suffixes ignore tab stops', () => {
    const space = measureTextWidth(' ', { fontFamily: 'Calibri', fontSize: 11 });
    expect(textStart(item([stop(BETWEEN)], { listMarkerSuffix: 'space' }))).toBe(
      MARKER_END + space
    );
    expect(textStart(item([stop(BETWEEN)], { listMarkerSuffix: 'nothing' }))).toBe(MARKER_END);
  });

  test('a marker wider than its slot keeps its natural width', () => {
    // Seven characters in a six-character slot: the slot is a min-width.
    const wide: ParagraphBlock = { ...item([stop(BETWEEN)]) };
    wide.attrs = { ...wide.attrs, listMarker: '(xviii)' };
    expect(getListMarkerInlineWidth(wide)).toBe(HANGING);
  });

  test('the first line wraps against the width the stop gives it', () => {
    // 50 characters wide, body 38 from the indent: nine "aaa " words. From the stop,
    // two characters earlier, the first line has 40: ten words.
    const text = Array(30).fill('aaa').join(' ');
    const words = (block: ParagraphBlock) => {
      const [first, second] = measureParagraph(block, 50 * CHAR).lines;
      return [first!.toChar - first!.fromChar, second!.toChar - second!.fromChar];
    };
    expect(words(item([], {}, text))).toEqual([36, 36]);
    expect(words(item([stop(BETWEEN)], {}, text))).toEqual([40, 36]);
  });
});
