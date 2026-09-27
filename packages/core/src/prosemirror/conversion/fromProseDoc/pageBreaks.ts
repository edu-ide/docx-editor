/**
 * Page breaks before a paragraph, ProseMirror → Document direction.
 *
 * `pageBreakBefore` folds in what `pageBreakBeforeSource` names: the paragraph's
 * own `w:pageBreakBefore` (`direct`), its style's (`style`), or a
 * `<w:br w:type="page"/>` run the paragraph opened with (`leadingBreak`, see
 * ../toProseDoc/pageBreaks.ts). Each is written back as it came: direct
 * formatting, nothing (the style still carries it), or the run. A paragraph
 * without a source (a node from before the attr existed) keeps the older
 * behaviour and writes its value as direct formatting.
 */

import type { BreakContent, Paragraph } from '../../../types/document';
import type { ParagraphAttrs } from '../../schema/nodes';

/**
 * The `w:pageBreakBefore` a paragraph states directly. `parsed` is the value
 * its `w:pPr` had when the document was read, if it had one.
 */
export function statedPageBreakBefore(
  attrs: ParagraphAttrs,
  parsed: boolean | undefined
): boolean | undefined {
  switch (attrs.pageBreakBeforeSource) {
    case 'direct':
      return attrs.pageBreakBefore ?? undefined;
    case 'style':
      return undefined;
    case 'leadingBreak':
      return parsed;
    default:
      return attrs.pageBreakBefore !== (parsed || undefined)
        ? attrs.pageBreakBefore || undefined
        : parsed;
  }
}

/** Paragraph content that marks a position without taking any. */
const MARKERS = new Set(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd']);

/**
 * Write back the page-break run the paragraph opened with, and Word's rendered
 * page-break marker. toProseDoc keeps the run as an empty run boundary, with its
 * run properties, while the paragraph's runs are unchanged; the break goes back
 * into that run. Otherwise it opens the paragraph's content as a run of its own.
 */
export function restorePageBreaks(paragraph: Paragraph, attrs: ParagraphAttrs): void {
  // Preserve `<w:lastRenderedPageBreak/>` so a save+reload doesn't silently
  // drop the break Word recorded for paginating this paragraph.
  if (attrs.renderedPageBreakBefore) {
    paragraph.renderedPageBreakBefore = true;
  }
  if (attrs.pageBreakBeforeSource !== 'leadingBreak') return;

  const pageBreak: BreakContent = { type: 'break', breakType: 'page' };
  const content = paragraph.content;
  const at = content.findIndex((item) => !MARKERS.has(item.type));
  const first = content[at];
  if (first?.type === 'run' && first.content.length === 0) {
    content[at] = { ...first, content: [pageBreak] };
  } else {
    content.splice(at < 0 ? content.length : at, 0, { type: 'run', content: [pageBreak] });
  }
}
