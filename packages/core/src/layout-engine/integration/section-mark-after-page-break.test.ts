/**
 * A manual page break followed by an empty section mark advances one sheet,
 * not two (upstream d6c75d2c, #981).
 *
 * When the section after the mark opens its own sheet anyway, the break and
 * the empty mark stay on the sheet the break closes and the section break
 * alone moves on. Mark-only continuous sections in between pass the question
 * on; a continuing section with content keeps the break.
 */
import { describe, expect, test } from 'bun:test';

import { layoutDocument } from '../index';
import { findBreakSheetJoins } from '../section-mark-break';
import type { FlowBlock, Layout, Measure, ParagraphBlock } from '../types';
import { parseDocumentBody } from '../../docx/documentParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import { getMargins, getPageSize } from '../../layout-bridge/sectionGeometry';
import { makeLayoutOptions, makeLine, makeParagraphMeasure } from './helpers';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const LINE = 24;
/** Letter, one-inch margins: an 864px content box holds 36 lines. */
const LINES_PER_PAGE = 36;

const sect = (type = '') =>
  `<w:sectPr>${type ? `<w:type w:val="${type}"/>` : ''}<w:pgSz w:w="12240" w:h="15840"/>` +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0"/>' +
  '</w:sectPr>';
const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const textThenBreak = '<w:p><w:r><w:t>closing</w:t></w:r><w:r><w:br w:type="page"/></w:r></w:p>';
const mark = (sectPr: string, pPr = '', content = '') =>
  `<w:p><w:pPr>${pPr}${sectPr}</w:pPr>${content}</w:p>`;

interface Laid {
  layout: Layout;
  blocks: FlowBlock[];
}

function lay(bodyXml: string): Laid {
  const body = parseDocumentBody(
    `<w:document xmlns:w="${W}"><w:body>${bodyXml}</w:body></w:document>`
  );
  const blocks = toFlowBlocks(toProseDoc({ package: { document: body } }), {});
  const measures: Measure[] = blocks.map((block) =>
    block.kind === 'paragraph'
      ? makeParagraphMeasure([makeLine(0, 0, 0, 1, 100, LINE)])
      : ({ kind: block.kind } as Measure)
  );
  const final = body.finalSectionProperties;
  const first = body.sections?.[0]?.properties ?? final;
  const layout = layoutDocument(
    blocks,
    measures,
    makeLayoutOptions({
      pageSize: getPageSize(first),
      margins: getMargins(first),
      finalPageSize: getPageSize(final),
      finalMargins: getMargins(final),
      bodyBreakType: final?.sectionStart as 'nextPage' | 'continuous' | undefined,
    })
  );
  return { layout, blocks };
}

const textOf = (block: FlowBlock | undefined) =>
  block?.kind === 'paragraph'
    ? block.runs.map((run) => (run.kind === 'text' ? run.text : '')).join('')
    : '';

/** Visible text per page. */
function pageTexts({ layout, blocks }: Laid): string[] {
  return layout.pages.map((page) =>
    page.fragments
      .map((fragment) => textOf(blocks.find((block) => block.id === fragment.blockId)))
      .filter((text) => text !== '')
      .join(' ')
  );
}

/** Page index and fragment of the `nth` paragraph block. */
function paragraphAt({ layout, blocks }: Laid, nth: number) {
  const block = blocks.filter((candidate) => candidate.kind === 'paragraph')[nth]!;
  for (const [page, record] of layout.pages.entries()) {
    const fragment = record.fragments.find((candidate) => candidate.blockId === block.id);
    if (fragment) return { page, fragment };
  }
  throw new Error(`paragraph ${nth} not placed`);
}

describe('a page break followed by an empty section mark', () => {
  test('advances one sheet before a next-page section', () => {
    const laid = lay(paragraph('one') + pageBreak + mark(sect()) + paragraph('two') + sect());
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    // The mark keeps a fragment for the caret on the sheet the break closed.
    const joined = paragraphAt(laid, 2);
    expect(joined.page).toBe(0);
    expect(joined.fragment.pmStart).toBeDefined();
    // The next section starts at the top of its own sheet.
    const two = paragraphAt(laid, 3);
    expect(two.page).toBe(1);
    expect(two.fragment.y).toBe(laid.layout.pages[1]!.margins.top);
  });

  test('advances one sheet when text precedes the break in the same paragraph', () => {
    const laid = lay(paragraph('one') + textThenBreak + mark(sect()) + paragraph('two') + sect());
    expect(pageTexts(laid)).toEqual(['one closing', 'two']);
    expect(paragraphAt(laid, 2).page).toBe(0);
  });

  test('ignores keep and page-break-before properties on the mark', () => {
    const keeps = '<w:keepNext/><w:keepLines/><w:pageBreakBefore/>';
    const laid = lay(
      paragraph('one') + pageBreak + mark(sect(), keeps) + paragraph('two') + sect()
    );
    expect(pageTexts(laid)).toEqual(['one', 'two']);
  });

  test('advances one sheet when a mark-only section opens the next sheet', () => {
    const laid = lay(
      paragraph('one') +
        pageBreak +
        mark(sect()) +
        mark(sect()) +
        paragraph('two') +
        sect('continuous')
    );
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    expect(paragraphAt(laid, 2).page).toBe(0);
    expect(paragraphAt(laid, 3).page).toBe(1);
  });

  test('advances one sheet across an intervening mark-only continuous section', () => {
    // The final section says nextPage explicitly: without a type the layout loop falls back
    // to the previous break's `continuous`.
    const laid = lay(
      paragraph('one') +
        pageBreak +
        mark(sect()) +
        mark(sect('continuous')) +
        paragraph('two') +
        sect('nextPage')
    );
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    expect(paragraphAt(laid, 2).page).toBe(0);
    expect(paragraphAt(laid, 3).page).toBe(0);
    const two = paragraphAt(laid, 4);
    expect(two.fragment.y).toBe(laid.layout.pages[1]!.margins.top);
  });

  test('keeps the second sheet when the break sits at the bottom of a full page', () => {
    const fill = Array.from({ length: LINES_PER_PAGE }, (_, index) => paragraph(`f${index}`));
    const laid = lay(fill.join('') + pageBreak + mark(sect()) + paragraph('two') + sect());
    expect(laid.layout.pages).toHaveLength(3);
    expect(pageTexts(laid)[2]).toBe('two');
    // The break paragraph and its mark share the sheet the break line opened.
    expect(paragraphAt(laid, LINES_PER_PAGE).page).toBe(1);
    expect(paragraphAt(laid, LINES_PER_PAGE + 1).page).toBe(1);
  });
});

describe('explicit breaks and content that still take a sheet', () => {
  test('two page breaks before the mark keep one empty sheet', () => {
    const laid = lay(
      paragraph('one') + pageBreak + pageBreak + mark(sect()) + paragraph('two') + sect()
    );
    expect(pageTexts(laid)).toEqual(['one', '', 'two']);
  });

  test('a mark paragraph with text takes its own sheet', () => {
    const marked = mark(sect(), '', '<w:r><w:t>marked</w:t></w:r>');
    const laid = lay(paragraph('one') + pageBreak + marked + paragraph('two') + sect());
    expect(pageTexts(laid)).toEqual(['one', 'marked', 'two']);
  });

  test('an empty paragraph between the break and the mark takes its own sheet', () => {
    const laid = lay(
      paragraph('one') + pageBreak + '<w:p/>' + mark(sect()) + paragraph('two') + sect()
    );
    expect(pageTexts(laid)).toEqual(['one', '', 'two']);
  });

  test('a trailing break at the end of the document keeps its empty sheet', () => {
    expect(lay(paragraph('one') + pageBreak + '<w:p/>' + sect()).layout.pages).toHaveLength(2);
  });

  test('a continuous section with content still starts after the break', () => {
    const laid = lay(
      paragraph('one') + pageBreak + mark(sect()) + paragraph('two') + sect('continuous')
    );
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    expect(paragraphAt(laid, 3).page).toBe(1);
  });

  test('a mark-only continuous section before a continuing section keeps the break', () => {
    const laid = lay(
      paragraph('one') +
        pageBreak +
        mark(sect()) +
        mark(sect('continuous')) +
        paragraph('two') +
        sect('continuous')
    );
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    expect(paragraphAt(laid, 4).page).toBe(1);
  });
});

describe('which marks join the break sheet', () => {
  const empty = (id: number): ParagraphBlock => ({ kind: 'paragraph', id, runs: [] });
  const config = {
    pageSize: { w: 816, h: 1056 },
    margins: { top: 96, right: 96, bottom: 96, left: 96 },
  };

  test('a long run of mark-only continuous sections resolves in one sweep', () => {
    // break, mark | 4000 mark-only continuous sections | final nextPage section.
    const count = 4000;
    const blocks: FlowBlock[] = [
      { kind: 'pageBreak', id: 0 },
      empty(1),
      { kind: 'sectionBreak', id: 2 },
    ];
    for (let index = 0; index < count; index += 1) {
      blocks.push(empty(blocks.length), { kind: 'sectionBreak', id: blocks.length + 1 });
    }
    const breakIndices = blocks.flatMap((block, index) =>
      block.kind === 'sectionBreak' ? [index] : []
    );
    const types = [
      undefined,
      ...Array<'continuous'>(count).fill('continuous'),
      'nextPage' as const,
    ];
    const configs = Array(breakIndices.length + 1).fill(config);
    const joins = findBreakSheetJoins(blocks, breakIndices, types, configs);
    expect([...joins.breaks]).toEqual([0]);
    expect([...joins.marks]).toEqual([1]);
    // The same run ending in a continuous section with content keeps the break.
    blocks.push(empty(blocks.length));
    blocks[blocks.length - 1] = {
      ...empty(blocks.length - 1),
      runs: [{ kind: 'text', text: 'x' }],
    };
    const kept = findBreakSheetJoins(
      blocks,
      breakIndices,
      [...types.slice(0, -1), 'continuous'],
      configs
    );
    expect(kept.breaks.size).toBe(0);
  });
});
