/**
 * A paragraph mark's `w:pPr/w:rPr` joins the paragraph's run cascade at the
 * character level (upstream 2eea4deb, #987): the paragraph style, then the
 * character style the mark's `w:rStyle` names (with its `basedOn` chain), then
 * the mark's direct properties. That cascade sizes an empty paragraph, gives an
 * empty list paragraph's marker its font, and is what typing starts from.
 *
 * The mark used to be resolved from the document defaults up, so any mark
 * `w:rPr` (a `w:lang`, a `w:b`, a style without a size) put the defaults'
 * size back over the paragraph style's.
 */
import { describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../../docx/documentParser';
import { parseStyleDefinitions, parseStyles } from '../../docx/styleParser';
import { parseNumbering } from '../../docx/numberingParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../toFlowBlocks';
import { resolveListMarkerFont } from '../measuring/listMarkerWidth';
import type { ParagraphBlock } from '../../layout-engine/types';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** 12pt Normal and 18pt Heading; `Big` is a 24pt character style, `Child` inherits it. */
const styles = (docDefaults = '') =>
  `<w:styles xmlns:w="${W}">${docDefaults}` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:sz w:val="24"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading"><w:basedOn w:val="Normal"/>' +
  '<w:rPr><w:sz w:val="36"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:default="1" w:styleId="DefaultFont"/>' +
  '<w:style w:type="character" w:styleId="Big"><w:basedOn w:val="DefaultFont"/>' +
  '<w:rPr><w:sz w:val="48"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Child"><w:basedOn w:val="Big"/>' +
  '<w:rPr><w:color w:val="C00000"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Small"><w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Colored"><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>' +
  '</w:styles>';
const DOC_DEFAULTS =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>';

const empty = (mark: string, pPr = '') => `<w:p><w:pPr>${pPr}<w:rPr>${mark}</w:rPr></w:pPr></w:p>`;

function paragraphs(body: string, stylesXml = styles()): ParagraphBlock[] {
  const definitions = parseStyleDefinitions(stylesXml, null);
  const document = parseDocumentBody(
    `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`,
    parseStyles(stylesXml, null)
  );
  const pmDoc = toProseDoc({ package: { document, styles: definitions } }, { styles: definitions });
  return toFlowBlocks(pmDoc, {}).filter(
    (block): block is ParagraphBlock => block.kind === 'paragraph'
  );
}

const sizes = (body: string, stylesXml?: string) =>
  paragraphs(body, stylesXml).map((block) => block.attrs?.defaultFontSize);

describe('an empty paragraph is sized by its mark', () => {
  test('a named character style sizes it the same as a direct size', () => {
    expect(
      sizes(empty('<w:rStyle w:val="Big"/>') + empty('<w:sz w:val="48"/>') + empty(''))
    ).toEqual([24, 24, 12]);
  });

  test('a smaller style shrinks it; basedOn inherits; a direct size wins', () => {
    expect(
      sizes(
        empty('<w:rStyle w:val="Small"/>') +
          empty('<w:rStyle w:val="Child"/>') +
          empty('<w:rStyle w:val="Big"/><w:sz w:val="32"/>')
      )
    ).toEqual([10, 24, 16]);
  });

  test('mark properties never put the document defaults back over the paragraph style', () => {
    const heading = '<w:pStyle w:val="Heading"/>';
    expect(
      sizes(
        empty('<w:rStyle w:val="Colored"/>', heading) +
          empty('<w:b/>', heading) +
          empty('<w:i/><w:color w:val="00FF00"/>', heading) +
          empty('', heading) +
          empty('<w:rStyle w:val="Small"/>', heading),
        styles(DOC_DEFAULTS)
      )
    ).toEqual([18, 18, 18, 18, 10]);
  });

  test('an id that names no character style changes nothing', () => {
    expect(
      sizes(empty('<w:rStyle w:val="Missing"/>') + empty('<w:rStyle w:val="Heading"/>'))
    ).toEqual([12, 12]);
  });
});

describe('what else reads the mark', () => {
  test('an empty list paragraph marker takes the size', () => {
    const numbering = parseNumbering(
      `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">` +
        '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
        '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>'
    );
    const definitions = parseStyleDefinitions(styles(DOC_DEFAULTS), null);
    const document = parseDocumentBody(
      `<w:document xmlns:w="${W}"><w:body>` +
        empty(
          '<w:b/>',
          '<w:pStyle w:val="Heading"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        ) +
        '</w:body></w:document>',
      parseStyles(styles(DOC_DEFAULTS), null),
      null,
      numbering
    );
    const pmDoc = toProseDoc(
      { package: { document, styles: definitions } },
      { styles: definitions }
    );
    const block = toFlowBlocks(pmDoc, {}).find((b): b is ParagraphBlock => b.kind === 'paragraph')!;
    expect(block.attrs?.listMarker).toBe('1.');
    // With no text run, the marker takes the paragraph mark's font: the Heading size.
    expect(resolveListMarkerFont(block).fontSize).toBe(18);
  });

  test('text typed into the empty paragraph starts from the character style', () => {
    const definitions = parseStyleDefinitions(styles(), null);
    const document = parseDocumentBody(
      `<w:document xmlns:w="${W}"><w:body>${empty('<w:rStyle w:val="Big"/>')}</w:body></w:document>`,
      parseStyles(styles(), null)
    );
    const pmDoc = toProseDoc(
      { package: { document, styles: definitions } },
      { styles: definitions }
    );
    const formatting = pmDoc.child(0).attrs.defaultTextFormatting as {
      styleId?: string;
      fontSize?: number;
    };
    expect(formatting.styleId).toBe('Big');
    expect(formatting.fontSize).toBe(48);
  });
});
