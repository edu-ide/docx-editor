/**
 * Page breaks before a paragraph, Document → ProseMirror direction.
 *
 * A paragraph that opens with `<w:br w:type="page"/>` becomes `pageBreakBefore`
 * on that paragraph; a later hard break becomes a standalone PM `pageBreak`
 * block (the detector consumed by the top-level orchestrator).
 * `pageBreakBeforeSource` records what `pageBreakBefore` came from, so
 * fromProseDoc writes it back the same way (see ../fromProseDoc/pageBreaks.ts).
 */

import type { ParagraphAttrs } from '../../schema/nodes';
import type { Hyperlink, Paragraph, Run, RunContent } from '../../../types/document';

/**
 * Set `pageBreakBefore` and where it comes from: a page-break run the paragraph
 * opens with, else the paragraph's own `w:pageBreakBefore`, else its style's
 * (`styleValue`, resolved by the caller). A paragraph with none keeps the defaults.
 */
export function assignPageBreakBefore(
  attrs: ParagraphAttrs,
  paragraph: Paragraph,
  styleValue: boolean | undefined
): void {
  const direct = paragraph.formatting?.pageBreakBefore;
  if (paragraphStartsWithPageBreak(paragraph)) {
    attrs.pageBreakBefore = true;
    attrs.pageBreakBeforeSource = 'leadingBreak';
  } else if (direct != null) {
    attrs.pageBreakBefore = direct;
    attrs.pageBreakBeforeSource = 'direct';
  } else if (styleValue != null) {
    attrs.pageBreakBefore = styleValue;
    attrs.pageBreakBeforeSource = 'style';
  }
}

type ParagraphContentToken = 'pageBreak' | 'visible';

function isVisibleRunContent(content: RunContent): boolean {
  if (content.type === 'text') return content.text.length > 0;
  return true;
}

function collectRunContentTokens(contents: RunContent[], tokens: ParagraphContentToken[]): void {
  for (const content of contents) {
    if (content.type === 'break' && content.breakType === 'page') {
      tokens.push('pageBreak');
    } else if (isVisibleRunContent(content)) {
      tokens.push('visible');
    }
  }
}

function collectRunOrHyperlinkTokens(
  items: readonly (Run | Hyperlink)[],
  tokens: ParagraphContentToken[]
): void {
  for (const item of items) {
    if (item.type === 'run') {
      collectRunContentTokens(item.content, tokens);
    } else {
      collectRunOrHyperlinkTokens(
        item.children.filter((child): child is Run => child.type === 'run'),
        tokens
      );
    }
  }
}

function collectParagraphContentTokens(
  items: readonly Paragraph['content'][number][],
  tokens: ParagraphContentToken[]
): void {
  for (const item of items) {
    switch (item.type) {
      case 'run':
        collectRunContentTokens(item.content, tokens);
        break;
      case 'hyperlink':
        collectRunOrHyperlinkTokens(
          item.children.filter((child): child is Run => child.type === 'run'),
          tokens
        );
        break;
      case 'simpleField':
        collectRunOrHyperlinkTokens(item.content, tokens);
        break;
      case 'complexField':
        collectRunOrHyperlinkTokens([...item.fieldCode, ...item.fieldResult], tokens);
        break;
      case 'inlineSdt':
        collectParagraphContentTokens(item.content as Paragraph['content'], tokens);
        break;
      case 'insertion':
      case 'deletion':
      case 'moveFrom':
      case 'moveTo':
        collectRunOrHyperlinkTokens(item.content, tokens);
        break;
      case 'mathEquation':
        tokens.push('visible');
        break;
    }
  }
}

function paragraphContentTokens(paragraph: Paragraph): ParagraphContentToken[] {
  const tokens: ParagraphContentToken[] = [];
  collectParagraphContentTokens(paragraph.content, tokens);
  return tokens;
}

export function paragraphStartsWithPageBreak(paragraph: Paragraph): boolean {
  return paragraphContentTokens(paragraph)[0] === 'pageBreak';
}

/**
 * Returns true when `<w:br w:type="page"/>` appears after the leading
 * position in a paragraph.
 *
 * A leading hard page break can be represented as `pageBreakBefore` on the
 * same paragraph, preserving the DOCX paragraph count through the PM round
 * trip. Later hard breaks still need a standalone PM `pageBreak` block so
 * layout keeps forcing a page boundary.
 */
export function paragraphHasNonLeadingPageBreak(paragraph: Paragraph): boolean {
  let consumedLeadingPageBreak = false;
  let sawVisibleContent = false;

  for (const token of paragraphContentTokens(paragraph)) {
    if (token === 'pageBreak') {
      if (sawVisibleContent || consumedLeadingPageBreak) {
        return true;
      }
      consumedLeadingPageBreak = true;
    } else {
      sawVisibleContent = true;
    }
  }

  return false;
}
