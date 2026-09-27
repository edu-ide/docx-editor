/**
 * Lays out DOCX body XML the way the editor does: parse it with the styles and
 * numbering below, convert it to ProseMirror and to flow blocks, and lay it out
 * with one 24px line per paragraph and per table row. Shared by the section-mark
 * page-break-before suites.
 */

import type { Node as PMNode } from 'prosemirror-model';

import { layoutDocument } from '../index';
import type {
  FlowBlock,
  Layout,
  LayoutOptions,
  Measure,
  ParagraphMeasure,
  TableBlock,
} from '../types';
import { parseDocumentBody } from '../../docx/documentParser';
import { parseNumbering } from '../../docx/numberingParser';
import { parseStyleDefinitions, parseStyles } from '../../docx/styleParser';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import { getMargins, getPageSize } from '../../layout-bridge/sectionGeometry';
import type { Document, DocumentBody } from '../../types/document';
import { makeLayoutOptions, makeLine, makeParagraphMeasure } from './helpers';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const LINE = 24;
/** Letter, one-inch margins: an 864px content box holds 36 lines. */
export const LINES_PER_PAGE = 36;

/** `Break` inherits the page break and keep through `basedOn`; `Listed` adds numbering. */
const STYLES =
  `<w:styles xmlns:w="${W}">` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:pageBreakBefore/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Break"><w:name w:val="Break"/><w:basedOn w:val="Base"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Listed"><w:basedOn w:val="Break"/>' +
  '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>' +
  '</w:styles>';
const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
  '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

const LETTER = '<w:pgSz w:w="12240" w:h="15840"/>';
export const sect = (type = '', size = LETTER) =>
  `<w:sectPr>${type ? `<w:type w:val="${type}"/>` : ''}${size}` +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0"/>' +
  '</w:sectPr>';
export const paragraph = (text: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
export const style = (id: string) => `<w:pStyle w:val="${id}"/>`;
/** A section-mark paragraph; by default its `Break` style breaks the page before it. */
export const mark = (sectPr: string, pPr = style('Break'), content = '') =>
  `<w:p><w:pPr>${pPr}${sectPr}</w:pPr>${content}</w:p>`;
export const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
export const body = (bodyXml: string) =>
  `<w:document xmlns:w="${W}"><w:body>${bodyXml}</w:body></w:document>`;

/** A document.xml read with the styles and numbering above, and its ProseMirror doc. */
export function read(documentXml: string): { document: Document; pmDoc: PMNode } {
  const styles = parseStyleDefinitions(STYLES, null);
  const document: Document = {
    package: {
      document: parseDocumentBody(
        documentXml,
        parseStyles(STYLES, null),
        null,
        parseNumbering(NUMBERING)
      ),
      styles,
    },
  };
  return { document, pmDoc: toProseDoc(document, { styles }) };
}

export interface Laid {
  layout: Layout;
  blocks: FlowBlock[];
}

const lineMeasure = (): ParagraphMeasure => makeParagraphMeasure([makeLine(0, 0, 0, 1, 100, LINE)]);

function tableMeasure(block: TableBlock): Measure {
  return {
    kind: 'table',
    columnWidths: block.columnWidths ?? [100],
    totalWidth: 100,
    totalHeight: block.rows.length * LINE,
    rows: block.rows.map((row) => ({
      height: LINE,
      cells: row.cells.map(() => ({ blocks: [lineMeasure()], width: 100, height: LINE })),
    })),
  };
}

function measure(block: FlowBlock): Measure {
  if (block.kind === 'paragraph') return lineMeasure();
  if (block.kind === 'table') return tableMeasure(block);
  return { kind: block.kind } as Measure;
}

/** Lay out a ProseMirror doc of `body`'s sections. */
export function layPmDoc(pmDoc: PMNode, body: DocumentBody): Laid {
  const blocks = toFlowBlocks(pmDoc, {});
  const final = body.finalSectionProperties;
  const first = body.sections?.[0]?.properties ?? final;
  const layout = layoutDocument(
    blocks,
    blocks.map(measure),
    makeLayoutOptions({
      pageSize: getPageSize(first),
      margins: getMargins(first),
      finalPageSize: getPageSize(final),
      finalMargins: getMargins(final),
      bodyBreakType: final?.sectionStart as LayoutOptions['bodyBreakType'],
    })
  );
  return { layout, blocks };
}

export function lay(bodyXml: string): Laid {
  const { document, pmDoc } = read(body(bodyXml));
  return layPmDoc(pmDoc, document.package.document);
}

const textOf = (block: FlowBlock | undefined) =>
  block?.kind === 'paragraph'
    ? block.runs.map((run) => (run.kind === 'text' ? run.text : '')).join('')
    : '';

/** Visible text per page. */
export function pageTexts({ layout, blocks }: Laid): string[] {
  return layout.pages.map((page) =>
    page.fragments
      .map((fragment) => textOf(blocks.find((block) => block.id === fragment.blockId)))
      .filter((text) => text !== '')
      .join(' ')
  );
}

/** Page index and fragment of the `nth` top-level paragraph. */
export function paragraphAt({ layout, blocks }: Laid, nth: number) {
  const block = blocks.filter((candidate) => candidate.kind === 'paragraph')[nth]!;
  for (const [page, record] of layout.pages.entries()) {
    const fragment = record.fragments.find((candidate) => candidate.blockId === block.id);
    if (fragment) return { page, fragment };
  }
  throw new Error(`paragraph ${nth} not placed`);
}
