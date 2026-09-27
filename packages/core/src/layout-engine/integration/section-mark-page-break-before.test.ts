/**
 * An empty section mark that follows content in its own section ignores its own
 * `w:pageBreakBefore`, from direct formatting or its style, so it never opens a blank
 * sheet before the next section (upstream 9afb832b, #983).
 *
 * The mark ends the section on the sheet its content reached; the next section's start
 * type alone decides where that section begins. A mark that is its section's only block
 * is that section's content and still breaks, and so does a mark that holds text, a
 * space, a tab, hidden text, a field, or a page-break run.
 */
import { describe, expect, test } from 'bun:test';

import { pageBreaksBefore } from '../section-mark-break';
import { schema } from '../../prosemirror/schema';
import {
  LINES_PER_PAGE,
  body,
  lay,
  layPmDoc,
  mark,
  pageBreak,
  pageTexts,
  paragraph,
  paragraphAt,
  read,
  sect,
  style,
} from './section-mark-helpers';

describe('an empty section mark with its own page break before', () => {
  test('stays after the content of its section before a next-page section', () => {
    const laid = lay(paragraph('one') + mark(sect()) + paragraph('two') + sect());
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    // The mark keeps its line on its section's sheet, right under the content.
    const one = paragraphAt(laid, 0).fragment;
    const kept = paragraphAt(laid, 1);
    expect(kept.page).toBe(0);
    expect(kept.fragment.y).toBe(one.y + one.height);
    const two = paragraphAt(laid, 2);
    expect(two.page).toBe(1);
    expect(two.fragment.y).toBe(laid.layout.pages[1]!.margins.top);
  });

  test('ignores a direct page break before the same way', () => {
    const laid = lay(
      paragraph('one') + mark(sect(), '<w:pageBreakBefore/>') + paragraph('two') + sect()
    );
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    expect(paragraphAt(laid, 1).page).toBe(0);
  });

  test('stays on the sheet before a continuous section', () => {
    const laid = lay(paragraph('one') + mark(sect()) + paragraph('two') + sect('continuous'));
    expect(pageTexts(laid)).toEqual(['one two']);
  });

  test('stays when the next section starts on an odd or even page', () => {
    // Page 2 is even: an odd-page section still inserts its parity sheet.
    const odd = lay(paragraph('one') + mark(sect()) + paragraph('two') + sect('oddPage'));
    expect(odd.layout.pages).toHaveLength(3);
    expect(odd.layout.pages[1]!.parityBlank).toBe(true);
    expect(paragraphAt(odd, 1).page).toBe(0);
    expect(paragraphAt(odd, 2).page).toBe(2);
    const even = lay(paragraph('one') + mark(sect()) + paragraph('two') + sect('evenPage'));
    expect(pageTexts(even)).toEqual(['one', 'two']);
  });

  test('stays when the next section changes the page size', () => {
    const wide = '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>';
    for (const type of ['', 'continuous']) {
      const laid = lay(paragraph('one') + mark(sect()) + paragraph('two') + sect(type, wide));
      expect(pageTexts(laid)).toEqual(['one', 'two']);
      expect(paragraphAt(laid, 1).page).toBe(0);
    }
  });

  test('stays before a mark-only continuous section and a next-page section', () => {
    // The final section says nextPage explicitly: without a type the layout loop falls
    // back to the previous break's `continuous`.
    const laid = lay(
      paragraph('one') +
        mark(sect()) +
        mark(sect('continuous'), '') +
        paragraph('two') +
        sect('nextPage')
    );
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    expect(paragraphAt(laid, 1).page).toBe(0);
    expect(paragraphAt(laid, 2).page).toBe(0);
  });

  test('ignores the break when the mark holds only bookmarks, markers, or empty runs', () => {
    for (const content of [
      '<w:bookmarkStart w:id="0" w:name="b"/><w:bookmarkEnd w:id="0"/>',
      '<w:r><w:rPr><w:b/></w:rPr></w:r>',
      '<w:r><w:t/></w:r>',
      '<w:r><w:t xml:space="preserve"></w:t></w:r>',
      '<w:proofErr w:type="gramStart"/><w:proofErr w:type="gramEnd"/>',
      '<w:permStart w:id="1" w:edGrp="everyone"/><w:permEnd w:id="1"/>',
      '<w:commentRangeStart w:id="0"/>',
      '<w:commentRangeEnd w:id="0"/>',
      '<w:r><w:lastRenderedPageBreak/></w:r>',
      '<w:proofErr w:type="spellStart"/><w:r><w:rPr><w:b/></w:rPr>' +
        '<w:lastRenderedPageBreak/><w:t/></w:r><w:proofErr w:type="spellEnd"/>',
    ]) {
      const laid = lay(
        paragraph('one') + mark(sect(), style('Break'), content) + paragraph('two') + sect()
      );
      expect(pageTexts(laid)).toEqual(['one', 'two']);
      expect(paragraphAt(laid, 1).page).toBe(0);
    }
  });

  test('ignores the break when the mark has a border, shading, or list numbering', () => {
    for (const pPr of [
      `${style('Break')}<w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="000000"/></w:pBdr>`,
      `${style('Break')}<w:shd w:val="clear" w:color="auto" w:fill="CCCCCC"/>`,
      style('Listed'),
    ]) {
      const laid = lay(paragraph('one') + mark(sect(), pPr) + paragraph('two') + sect());
      expect(laid.layout.pages).toHaveLength(2);
      expect(paragraphAt(laid, 1).page).toBe(0);
    }
  });

  test('ignores the break after a table, an empty paragraph, or a paragraph that breaks', () => {
    const table =
      '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    const afterTable = lay(table + mark(sect()) + paragraph('two') + sect());
    expect(afterTable.layout.pages).toHaveLength(2);
    expect(paragraphAt(afterTable, 0).page).toBe(0);
    const afterEmpty = lay(paragraph('one') + '<w:p/>' + mark(sect()) + paragraph('two') + sect());
    expect(pageTexts(afterEmpty)).toEqual(['one', 'two']);
    // The earlier paragraph's own break still opens its sheet; the mark follows it there.
    const broken = lay(
      paragraph('one') +
        paragraph('broken', '<w:pageBreakBefore/>') +
        mark(sect()) +
        paragraph('two') +
        sect()
    );
    expect(pageTexts(broken)).toEqual(['one', 'broken', 'two']);
    expect(paragraphAt(broken, 2).page).toBe(1);
  });
});

describe('section marks whose page break before still applies', () => {
  test('a mark that is its section only block breaks before a continued sheet', () => {
    const laid = lay(
      paragraph('one') +
        mark(sect(), '') +
        mark(sect('continuous')) +
        paragraph('two') +
        sect('nextPage')
    );
    expect(laid.layout.pages).toHaveLength(3);
    expect(paragraphAt(laid, 2).page).toBe(1);
    expect(pageTexts(laid)[2]).toBe('two');
  });

  test('two marks in a row: the second is its own next-page section', () => {
    const laid = lay(paragraph('one') + mark(sect()) + mark(sect()) + paragraph('two') + sect());
    expect(laid.layout.pages).toHaveLength(3);
    expect(paragraphAt(laid, 1).page).toBe(0);
    expect(paragraphAt(laid, 2).page).toBe(1);
  });

  test('a mark with text, a space, a tab, hidden text, or a field breaks', () => {
    const marked = mark(sect(), style('Break'), '<w:r><w:t>marked</w:t></w:r>');
    expect(pageTexts(lay(paragraph('one') + marked + paragraph('two') + sect()))).toEqual([
      'one',
      'marked',
      'two',
    ]);
    for (const content of [
      '<w:r><w:t xml:space="preserve"> </w:t></w:r>',
      '<w:r><w:tab/></w:r>',
      '<w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r>',
      '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>',
      '<w:proofErr w:type="spellStart"/><w:r><w:t>marked</w:t></w:r><w:proofErr w:type="spellEnd"/>',
      '<w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve"> </w:t></w:r>',
      '<w:r><w:t/><w:tab/></w:r>',
      '<w:permStart w:id="1" w:edGrp="everyone"/>' +
        '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:permEnd w:id="1"/>',
    ]) {
      const laid = lay(
        paragraph('one') + mark(sect(), style('Break'), content) + paragraph('two') + sect()
      );
      expect(laid.layout.pages).toHaveLength(3);
      expect(paragraphAt(laid, 1).page).toBe(1);
    }
  });

  test('a mark with inert markers that is its section only block breaks', () => {
    const laid = lay(
      paragraph('one') +
        mark(sect(), '') +
        mark(sect('continuous'), style('Break'), '<w:proofErr w:type="gramStart"/>') +
        paragraph('two') +
        sect('nextPage')
    );
    expect(laid.layout.pages).toHaveLength(3);
    expect(paragraphAt(laid, 2).page).toBe(1);
  });

  test('an empty paragraph that carries no section mark breaks', () => {
    const laid = lay(
      paragraph('one') + `<w:p><w:pPr>${style('Break')}</w:pPr></w:p>` + paragraph('two') + sect()
    );
    expect(laid.layout.pages).toHaveLength(2);
    expect(paragraphAt(laid, 1).page).toBe(1);
  });

  test('a page break before the mark still advances one sheet', () => {
    const laid = lay(paragraph('one') + pageBreak + mark(sect()) + paragraph('two') + sect());
    expect(pageTexts(laid)).toEqual(['one', 'two']);
    expect(paragraphAt(laid, 2).page).toBe(0);
  });
});

describe('a page-break run is content, not a paragraph property', () => {
  test('a mark that opens with a page-break run keeps that break', () => {
    // This rule lets go of the mark's `w:pageBreakBefore` only. Where the run itself leaves
    // the mark is for the page-break rules to decide.
    const opening = mark(sect(), style('Break'), '<w:r><w:br w:type="page"/></w:r>');
    const { blocks } = lay(paragraph('one') + opening + paragraph('two') + sect());
    const breakIndices = blocks.flatMap((block, index) =>
      block.kind === 'sectionBreak' ? [index] : []
    );
    const markIndex = breakIndices[0]! - 1;
    expect(blocks[markIndex]).toMatchObject({
      attrs: { pageBreakBefore: true, pageBreakBeforeSource: 'leadingBreak' },
    });
    expect(pageBreaksBefore(blocks, breakIndices)(markIndex)).toBe(true);
  });

  test('a paragraph that opens with a page break after a full page opens one sheet', () => {
    // Upstream e633def8 (#982): the break takes the paragraph to the next sheet, no further.
    const fill = Array.from({ length: LINES_PER_PAGE }, (_, index) => paragraph(`f${index}`));
    const opening = '<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>after</w:t></w:r></w:p>';
    const laid = lay(fill.join('') + opening + sect());
    expect(laid.layout.pages).toHaveLength(2);
    expect(pageTexts(laid)[1]).toBe('after');
  });

  test('a mark whose break source was never recorded keeps its break', () => {
    // Collaborative documents persist ProseMirror JSON: a paragraph stored before
    // `pageBreakBeforeSource` existed comes back with the attr at its default.
    const { document, pmDoc } = read(
      body(paragraph('one') + mark(sect()) + paragraph('two') + sect())
    );
    const json = pmDoc.toJSON() as { content: Array<{ attrs: Record<string, unknown> }> };
    delete json.content[1]!.attrs.pageBreakBeforeSource;
    const laid = layPmDoc(schema.nodeFromJSON(json), document.package.document);
    expect(pageTexts(laid)).toEqual(['one', '', 'two']);
  });
});
