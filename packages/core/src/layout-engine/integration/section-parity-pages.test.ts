/**
 * Blank parity pages between sections (upstream 6794f4d3, #978).
 *
 * An `oddPage` / `evenPage` section (ECMA-376 §17.18.77) starts on the next
 * page whose number has that parity. When the page after the break has the
 * other parity, layout leaves it blank (`Page.parityBlank`) and starts the
 * section on the page after it. The paginator's `forcePageBreak` is
 * idempotent on an empty page, so the old second call never added that page
 * and the section started on a page of the wrong parity.
 */
import { describe, expect, test } from 'bun:test';

import { layoutDocument } from '../index';
import type { FlowBlock, Layout, Measure, SectionBreakBlock } from '../types';
import { parseDocumentBody } from '../../docx/documentParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import { makeParagraphBlock, makeLine, makeParagraphMeasure, makeLayoutOptions } from './helpers';

type BreakType = SectionBreakBlock['type'];
type PageSize = { w: number; h: number };
/** A section-ending break carries its own section's start type and page size. */
type Item = string | { pageBreak: true } | { section: BreakType | undefined; pageSize?: PageSize };

const pageBreak = { pageBreak: true } as const;
const sectionEnd = (type?: BreakType, pageSize?: PageSize) => ({ section: type, pageSize });

function lay(items: Item[], bodyBreakType?: BreakType, finalPageSize?: PageSize) {
  const blocks: FlowBlock[] = [];
  const measures: Measure[] = [];
  let pos = 1;
  items.forEach((item, id) => {
    if (typeof item === 'string') {
      blocks.push(makeParagraphBlock(id, item, pos));
      measures.push(makeParagraphMeasure([makeLine(0, 0, 0, item.length, 100, 24)]));
      pos += item.length + 2;
    } else if ('pageBreak' in item) {
      blocks.push({ kind: 'pageBreak', id });
      measures.push({ kind: 'pageBreak' });
    } else {
      blocks.push({ kind: 'sectionBreak', id, type: item.section, pageSize: item.pageSize });
      measures.push({ kind: 'sectionBreak' });
    }
  });
  const layout = layoutDocument(
    blocks,
    measures,
    makeLayoutOptions({ bodyBreakType, ...(finalPageSize ? { finalPageSize } : {}) })
  );
  return { layout, blocks };
}

/** Body text per page, with `|` for a blank parity page. */
function pagesOf({ layout, blocks }: { layout: Layout; blocks: FlowBlock[] }): string[] {
  return layout.pages.map((page) => {
    if (page.parityBlank) return '|';
    return page.fragments
      .map((fragment) => {
        const block = blocks.find((candidate) => candidate.id === fragment.blockId);
        return block?.kind === 'paragraph' && block.runs[0]?.kind === 'text'
          ? block.runs[0].text
          : '';
      })
      .join('');
  });
}

describe('oddPage and evenPage sections start on a page number of that parity', () => {
  test('oddPage after one page inserts a blank page that takes number 2', () => {
    const result = lay(['First', sectionEnd(), 'Second'], 'oddPage');
    expect(pagesOf(result)).toEqual(['First', '|', 'Second']);
    const blank = result.layout.pages[1]!;
    expect(blank.number).toBe(2);
    expect(blank.fragments).toHaveLength(0);
    expect(result.layout.pages[2]!.number).toBe(3);
    expect(result.layout.pages[2]!.parityBlank).toBeUndefined();
  });

  test('evenPage after one page needs no blank page', () => {
    expect(pagesOf(lay(['First', sectionEnd(), 'Second'], 'evenPage'))).toEqual([
      'First',
      'Second',
    ]);
  });

  test('oddPage after two pages needs no blank page', () => {
    const result = lay(['One', pageBreak, 'Two', sectionEnd(), 'Second'], 'oddPage');
    expect(pagesOf(result)).toEqual(['One', 'Two', 'Second']);
  });

  test('evenPage after two pages inserts a blank page', () => {
    const result = lay(['One', pageBreak, 'Two', sectionEnd(), 'Second'], 'evenPage');
    expect(pagesOf(result)).toEqual(['One', 'Two', '|', 'Second']);
    expect(result.layout.pages[3]!.number).toBe(4);
  });

  test('an empty final oddPage section still gets its blank page first', () => {
    const result = lay(['First', sectionEnd()], 'oddPage');
    expect(pagesOf(result)).toEqual(['First', '|', '']);
    expect(result.layout.pages[2]!.parityBlank).toBeUndefined();
  });

  test('the parity of a middle section comes from its own break type', () => {
    // Section 2 starts on an odd page (blank page 2), section 3 on an even page (page 4).
    const result = lay(
      ['First', sectionEnd(), 'Second', sectionEnd('oddPage'), 'Third'],
      'evenPage'
    );
    expect(pagesOf(result)).toEqual(['First', '|', 'Second', 'Third']);
  });

  test('the blank page takes the geometry of the section it precedes', () => {
    const landscape = { w: 1056, h: 816 };
    // Section 2 is a landscape oddPage section; the landscape final section continues on
    // its sheet.
    const result = lay(
      ['First', sectionEnd(), 'Second', sectionEnd('oddPage', landscape), 'Tail'],
      'continuous',
      landscape
    );
    expect(pagesOf(result)).toEqual(['First', '|', 'SecondTail']);
    expect(result.layout.pages[0]!.size).toEqual({ w: 816, h: 1056 });
    expect(result.layout.pages[1]!.size).toEqual(landscape);
    expect(result.layout.pages[2]!.size).toEqual(landscape);
  });

  test('a page break or pageBreakBefore after the break adds no further page', () => {
    const withBreak = lay(['First', sectionEnd(), pageBreak, 'Second'], 'oddPage');
    expect(pagesOf(withBreak)).toEqual(['First', '|', 'Second']);
  });

  test('continuous and nextPage sections never get a blank page', () => {
    expect(pagesOf(lay(['First', sectionEnd(), 'Second'], 'continuous'))).toEqual(['FirstSecond']);
    expect(pagesOf(lay(['First', sectionEnd(), 'Second'], 'nextPage'))).toEqual([
      'First',
      'Second',
    ]);
  });
});

describe('parity from a parsed document', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const sectPr = (type?: string) =>
    `<w:sectPr>${type ? `<w:type w:val="${type}"/>` : ''}<w:pgSz w:w="12240" w:h="15840"/>` +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

  function layParsed(bodyXml: string): string[] {
    const body = parseDocumentBody(
      `<w:document xmlns:w="${W}"><w:body>${bodyXml}</w:body></w:document>`
    );
    const pmDoc = toProseDoc({ package: { document: body } });
    const blocks = toFlowBlocks(pmDoc, {});
    const measures: Measure[] = blocks.map((block) =>
      block.kind === 'paragraph'
        ? makeParagraphMeasure([makeLine(0, 0, 0, 1, 100, 24)])
        : ({ kind: block.kind } as Measure)
    );
    const layout = layoutDocument(
      blocks,
      measures,
      makeLayoutOptions({
        bodyBreakType: body.finalSectionProperties?.sectionStart as BreakType,
      })
    );
    return pagesOf({ layout, blocks });
  }

  test('a w:type oddPage section after one page starts on page 3', () => {
    const pages = layParsed(
      `<w:p><w:pPr>${sectPr()}</w:pPr><w:r><w:t>First</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>Second</w:t></w:r></w:p>${sectPr('oddPage')}`
    );
    expect(pages).toEqual(['First', '|', 'Second']);
  });

  test('a w:type evenPage section after one page starts on page 2', () => {
    const pages = layParsed(
      `<w:p><w:pPr>${sectPr()}</w:pPr><w:r><w:t>First</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>Second</w:t></w:r></w:p>${sectPr('evenPage')}`
    );
    expect(pages).toEqual(['First', 'Second']);
  });
});
