/**
 * A forced break ends a `w:keepNext` group (upstream ab460dc9, #985).
 *
 * When the paragraph a heading is kept with starts a new page (`pageBreakBefore`),
 * or a page or column break follows the heading, nothing after the break can share
 * the heading's page. Pricing the next paragraph's first line into the group moved
 * the heading to a new page, and the break then opened another: the heading was
 * stranded alone. The group now ends at the break, and the break discards the last
 * member's trailing spacing.
 */
import { describe, expect, test } from 'bun:test';

import { layoutDocument } from '../index';
import { computeKeepNextChains } from '../keep-together';
import type { FlowBlock, Measure } from '../types';
import { makeLayoutOptions, makeLine, makeParagraphBlock, makeParagraphMeasure } from './helpers';

/** The default page's 864px content box holds eight 100px lines. */
const LINE = 100;

interface Spec {
  lines?: number;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  spacingAfter?: number;
  /** A standalone page break block instead of a paragraph. */
  pageBreak?: true;
}

function lay(specs: Spec[]) {
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  let pos = 1;
  specs.forEach((spec, id) => {
    if (spec.pageBreak) {
      blocks.push({ kind: 'pageBreak', id });
      measures.push({ kind: 'pageBreak' });
      return;
    }
    const lines = spec.lines ?? 1;
    const block = makeParagraphBlock(id, 'x'.repeat(lines), pos, {
      keepNext: spec.keepNext,
      pageBreakBefore: spec.pageBreakBefore,
    });
    if (spec.keepLines) block.attrs = { ...block.attrs, keepLines: true };
    if (spec.spacingAfter) block.attrs = { ...block.attrs, spacing: { after: spec.spacingAfter } };
    pos += lines + 2;
    blocks.push(block);
    const measure = makeParagraphMeasure(
      Array.from({ length: lines }, (_, line) => makeLine(0, line, 0, line + 1, 10, LINE))
    );
    measure.totalHeight += spec.spacingAfter ?? 0;
    measures.push(measure);
  });
  return { blocks, layout: layoutDocument(blocks, measures, makeLayoutOptions()) };
}

/** Lines placed on each page. */
function linesPerPage(specs: Spec[]): number[] {
  return lay(specs).layout.pages.map((page) =>
    page.fragments.reduce(
      (sum, fragment) =>
        sum + (fragment.kind === 'paragraph' ? fragment.toLine - fragment.fromLine : 0),
      0
    )
  );
}

const fillers = (count: number): Spec[] => Array.from({ length: count }, () => ({}));
const heading: Spec = { keepNext: true };

describe('a forced break ends the keepNext group', () => {
  test('a successor with page break before leaves the heading on its page', () => {
    // Seven fillers and the heading fill eight lines; the successor starts a new page anyway.
    expect(linesPerPage([...fillers(7), heading, { lines: 3, pageBreakBefore: true }])).toEqual([
      8, 3,
    ]);
  });

  test('keepLines on that successor adds no extra heading page', () => {
    expect(
      linesPerPage([...fillers(7), heading, { lines: 4, pageBreakBefore: true, keepLines: true }])
    ).toEqual([8, 4]);
  });

  test('the break discards the heading trailing spacing from the group', () => {
    const spaced: Spec = { keepNext: true, spacingAfter: 80 };
    expect(linesPerPage([...fillers(7), spaced, { lines: 3, pageBreakBefore: true }])).toEqual([
      8, 3,
    ]);
  });

  test('a page break before an internal chain member ends the preceding group', () => {
    expect(
      linesPerPage([
        ...fillers(7),
        heading,
        { keepNext: true, pageBreakBefore: true },
        { lines: 3 },
      ])
    ).toEqual([8, 4]);
  });

  test('a page break after the heading ends its group', () => {
    const spaced: Spec = { keepNext: true, spacingAfter: 80 };
    expect(linesPerPage([...fillers(7), spaced, { pageBreak: true }, { lines: 3 }])).toEqual([
      8, 3,
    ]);
  });

  test('an ordinary successor still takes the heading to the next page', () => {
    expect(linesPerPage([...fillers(7), heading, { lines: 3 }])).toEqual([7, 4]);
  });
});

describe('computeKeepNextChains', () => {
  test('marks a chain that ends at a forced break and gives it no anchor', () => {
    const { blocks } = lay([heading, { pageBreakBefore: true }, heading, {}]);
    const chains = computeKeepNextChains(blocks);
    expect(chains.get(0)).toMatchObject({ memberIndices: [0], anchorIndex: -1, endsAtBreak: true });
    expect(chains.get(2)).toMatchObject({ memberIndices: [2], anchorIndex: 3 });
    expect(chains.get(2)?.endsAtBreak).toBeUndefined();
  });
});
