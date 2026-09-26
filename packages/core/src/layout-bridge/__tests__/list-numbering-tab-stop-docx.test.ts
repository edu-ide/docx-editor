/**
 * A parsed numbered paragraph reaches the numbering-tab rule with the stops
 * Word gives it (upstream bf776f2d #979): a direct stop, a legacy `num` stop
 * and a paragraph-style stop all pull the first line to the stop before the
 * indent, and a direct `clear` removes the style's stop.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { parseDocumentBody } from '../../docx/documentParser';
import { parseNumbering } from '../../docx/numberingParser';
import { parseStyleDefinitions, parseStyles } from '../../docx/styleParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../toFlowBlocks';
import { getListMarkerInlineWidth } from '../measuring/listMarkerWidth';
import type { ParagraphBlock } from '../../layout-engine/types';

let originalGetContext: typeof HTMLCanvasElement.prototype.getContext | undefined;

beforeAll(() => {
  GlobalRegistrator.register();
  originalGetContext = HTMLCanvasElement.prototype.getContext;
  HTMLCanvasElement.prototype.getContext = function getContext(type: string) {
    if (type !== '2d') return null;
    return {
      font: '',
      measureText: (text: string) => ({ width: text.length * 8 }),
    } as unknown as CanvasRenderingContext2D;
  } as typeof HTMLCanvasElement.prototype.getContext;
});

afterAll(() => {
  if (originalGetContext) HTMLCanvasElement.prototype.getContext = originalGetContext;
  GlobalRegistrator.unregister();
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Level 0: marker slot at 720tw (left 1440, hanging 720); `1.` ends well before 1200tw.
const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">` +
  '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="720"/></w:pPr></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>';

const STYLES =
  `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Stops">` +
  '<w:name w:val="Stops"/><w:pPr><w:tabs><w:tab w:val="left" w:pos="1200"/></w:tabs></w:pPr>' +
  '</w:style></w:styles>';

const px = (twips: number) => (twips / 1440) * 96;

/** Where the first line's text of a numbered paragraph with `pPr` extras starts. */
function textStart(pPr = ''): number {
  const body = parseDocumentBody(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:pPr>${pPr}<w:numPr><w:ilvl w:val="0"/>` +
      '<w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>',
    parseStyles(STYLES, null),
    null,
    parseNumbering(NUMBERING)
  );
  const styles = parseStyleDefinitions(STYLES, null);
  const pmDoc = toProseDoc({ package: { document: body, styles } }, { styles });
  const block = toFlowBlocks(pmDoc, {}).find((b): b is ParagraphBlock => b.kind === 'paragraph')!;
  const indent = block.attrs!.indent!;
  return indent.left! - indent.hanging! + getListMarkerInlineWidth(block);
}

const tabs = (...stops: string[]) => `<w:tabs>${stops.join('')}</w:tabs>`;
const stop = (pos: number, val = 'left') => `<w:tab w:val="${val}" w:pos="${pos}"/>`;

describe('numbering tab stops from a parsed document', () => {
  test('without a stop, the text starts at the indent', () => {
    expect(textStart()).toBe(px(1440));
  });

  test('a direct stop between the marker and the indent wins', () => {
    expect(textStart(tabs(stop(1200)))).toBe(px(1200));
  });

  test('a legacy numbering stop wins the same way', () => {
    expect(textStart(tabs(stop(1200, 'num')))).toBe(px(1200));
  });

  test('a paragraph-style stop is inherited, and a direct clear removes it', () => {
    expect(textStart('<w:pStyle w:val="Stops"/>')).toBe(px(1200));
    expect(textStart(`<w:pStyle w:val="Stops"/>${tabs(stop(1200, 'clear'))}`)).toBe(px(1440));
  });
});
