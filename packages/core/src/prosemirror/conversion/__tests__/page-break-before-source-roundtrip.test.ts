/**
 * `pageBreakBeforeSource` records what a paragraph's `pageBreakBefore` folds in, so saving
 * writes back the page break the document had: the paragraph's own `w:pageBreakBefore` as
 * direct formatting, nothing for one its style sets, and a page-break run for one the
 * paragraph opened with.
 *
 * The layout needs the same distinction for upstream 9afb832b (#983): an empty section mark
 * after its section's content ignores its page break property, but a page-break run is
 * content.
 */
import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../../../docx/documentParser';
import { parseStyleDefinitions, parseStyles } from '../../../docx/styleParser';
import { serializeParagraph } from '../../../docx/serializer/paragraphSerializer';
import type { Document, Paragraph } from '../../../types/document';
import { schema } from '../../schema';
import { fromProseDoc } from '../fromProseDoc';
import { toProseDoc } from '../toProseDoc';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** `Break` inherits `w:pageBreakBefore` from `Base` through `basedOn`. */
const STYLES =
  `<w:styles xmlns:w="${W}">` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Base"><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:pageBreakBefore/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Break"><w:basedOn w:val="Base"/></w:style>' +
  '</w:styles>';

const SECT =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  '</w:sectPr>';
const BREAK_RUN = '<w:r><w:br w:type="page"/></w:r>';
const TEXT_RUN = '<w:r><w:t>x</w:t></w:r>';

/** Parse one body paragraph with the styles, convert it to PM and back, and serialize it. */
function roundTrip(paragraphXml: string) {
  const styles = parseStyleDefinitions(STYLES, null);
  const document: Document = {
    package: {
      document: parseDocumentBody(
        `<w:document xmlns:w="${W}"><w:body>${paragraphXml}<w:sectPr/></w:body></w:document>`,
        parseStyles(STYLES, null)
      ),
      styles,
    },
  };
  const pmDoc = toProseDoc(document, { styles });
  const saved = fromProseDoc(pmDoc, document).package.document.content[0] as Paragraph;
  return { attrs: pmDoc.child(0).attrs, xml: serializeParagraph(saved) };
}

describe('a page break before round-trips as it was written', () => {
  test('a direct w:pageBreakBefore stays direct formatting', () => {
    const xml = `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${TEXT_RUN}</w:p>`;
    const saved = roundTrip(xml);
    expect(saved.attrs.pageBreakBefore).toBe(true);
    expect(saved.attrs.pageBreakBeforeSource).toBe('direct');
    expect(saved.xml).toBe(xml);
  });

  test('a w:pageBreakBefore from the style is not written as direct formatting', () => {
    const xml = `<w:p><w:pPr><w:pStyle w:val="Break"/></w:pPr>${TEXT_RUN}</w:p>`;
    const saved = roundTrip(xml);
    expect(saved.attrs.pageBreakBefore).toBe(true);
    expect(saved.attrs.pageBreakBeforeSource).toBe('style');
    expect(saved.xml).toBe(xml);
  });

  test('a leading page-break run stays a run', () => {
    const xml = `<w:p>${BREAK_RUN}${TEXT_RUN}</w:p>`;
    const saved = roundTrip(xml);
    expect(saved.attrs.pageBreakBefore).toBe(true);
    expect(saved.attrs.pageBreakBeforeSource).toBe('leadingBreak');
    expect(saved.xml).toBe(xml);
  });

  test('a direct w:pageBreakBefore and a leading page-break run both stay', () => {
    const xml = `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${BREAK_RUN}${TEXT_RUN}</w:p>`;
    const saved = roundTrip(xml);
    expect(saved.attrs.pageBreakBeforeSource).toBe('leadingBreak');
    expect(saved.xml).toBe(xml);
  });

  test('a style page break and a leading page-break run both stay', () => {
    const xml = `<w:p><w:pPr><w:pStyle w:val="Break"/></w:pPr>${BREAK_RUN}${TEXT_RUN}</w:p>`;
    expect(roundTrip(xml).xml).toBe(xml);
  });

  test('a paragraph holding only a page-break run keeps the run and its properties', () => {
    for (const xml of [
      `<w:p>${BREAK_RUN}</w:p>`,
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:br w:type="page"/></w:r></w:p>',
    ]) {
      expect(roundTrip(xml).xml).toBe(xml);
    }
  });

  test('a direct w:val="0" that switches the style page break off stays', () => {
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="Break"/><w:pageBreakBefore w:val="0"/></w:pPr>' +
      `${TEXT_RUN}</w:p>`;
    const saved = roundTrip(xml);
    expect(saved.attrs.pageBreakBefore).toBe(false);
    expect(saved.xml).toBe(xml);
  });

  test('a section mark keeps each form', () => {
    for (const xml of [
      `<w:p><w:pPr><w:pageBreakBefore/>${SECT}</w:pPr></w:p>`,
      `<w:p><w:pPr><w:pStyle w:val="Break"/>${SECT}</w:pPr></w:p>`,
      `<w:p><w:pPr>${SECT}</w:pPr>${BREAK_RUN}</w:p>`,
      `<w:p><w:pPr><w:pageBreakBefore/>${SECT}</w:pPr>${BREAK_RUN}</w:p>`,
    ]) {
      expect(roundTrip(xml).xml).toBe(xml);
    }
  });

  test('the rendered page break Word records after a leading page break stays after it', () => {
    const xml = `<w:p>${BREAK_RUN}<w:r><w:lastRenderedPageBreak/><w:t>x</w:t></w:r></w:p>`;
    const saved = roundTrip(xml);
    expect(saved.attrs.renderedPageBreakBefore).toBe(true);
    expect(saved.xml).toBe(xml);
  });

  test('a paragraph from before the source was recorded saves its break as before', () => {
    // Collaborative documents persist PM JSON; a paragraph written before the attr existed
    // comes back with `pageBreakBeforeSource` at its default and saves as direct formatting.
    const legacy = schema.nodeFromJSON({
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          attrs: { pageBreakBefore: true },
          content: [{ type: 'text', text: 'x' }],
        },
      ],
    });
    expect(legacy.child(0).attrs.pageBreakBeforeSource).toBeNull();
    const saved = fromProseDoc(legacy).package.document.content[0] as Paragraph;
    expect(serializeParagraph(saved)).toBe(
      `<w:p><w:pPr><w:pageBreakBefore/></w:pPr>${TEXT_RUN}</w:p>`
    );
  });
});
