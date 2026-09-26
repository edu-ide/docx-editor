/**
 * A manual page break followed by an empty section mark advances one sheet,
 * not two (upstream d6c75d2c, #981).
 *
 * The fork keeps a page-break-only paragraph as an empty paragraph with
 * `pageBreakBefore`, and a break after text as a standalone `pageBreak` block,
 * so the break opens the next page and the empty section-mark paragraph after
 * it lands there; a section that starts on its own sheet then opens a third
 * page. Word keeps the break and the empty mark on the sheet the break closes
 * when the next section opens a sheet anyway: the section break alone moves on.
 * Mark-only continuous sections in between change nothing, so the question
 * passes through them; a continuing section with content keeps the break.
 */

import type { SectionLayoutConfig } from './index';
import type { Paginator } from './paginator';
import { getParagraphFragmentPmRange } from './paragraphFragmentRange';
import type {
  FlowBlock,
  ParagraphBlock,
  ParagraphFragment,
  ParagraphMeasure,
  SectionBreakBlock,
} from './types';

/** A paragraph that paints nothing: no text, tab, picture, field, break, list marker or rule. */
function paintsNothing(block: FlowBlock | undefined): block is ParagraphBlock {
  if (block?.kind !== 'paragraph') return false;
  const attrs = block.attrs;
  if (attrs?.listMarker && !attrs.listMarkerHidden) return false;
  const rules = attrs?.borders;
  if (
    rules &&
    (rules.top ?? rules.bottom ?? rules.left ?? rules.right ?? rules.between ?? rules.bar)
  ) {
    return false;
  }
  return block.runs.every((run) => run.kind === 'text' && run.text.length === 0);
}

function sizeChanges(from: SectionLayoutConfig, to: SectionLayoutConfig): boolean {
  return (
    Math.round(from.pageSize.w) !== Math.round(to.pageSize.w) ||
    Math.round(from.pageSize.h) !== Math.round(to.pageSize.h)
  );
}

/** Where page-break blocks hand their break to the section break after an empty mark. */
export interface BreakSheetJoins {
  /** The break (a `pageBreak` block or a break-only paragraph) that no longer forces a page. */
  readonly breaks: ReadonlySet<number>;
  /** The empty section-mark paragraph that stays on the break's sheet, out of the flow. */
  readonly marks: ReadonlySet<number>;
}

/**
 * Find the page break + empty section mark pairs whose section break opens a sheet anyway.
 *
 * `types` and `configs` are the layout loop's own `sectionBreakTypes` and section configs, so
 * "opens a sheet" is decided exactly as `handleSectionBreak` will decide it: the next
 * section's type (falling back to this break's type), and a continuous section that changes
 * the page size. One backward sweep, so a run of mark-only sections costs one step each.
 */
export function findBreakSheetJoins(
  blocks: FlowBlock[],
  breakIndices: number[],
  types: (SectionBreakBlock['type'] | undefined)[],
  configs: SectionLayoutConfig[]
): BreakSheetJoins {
  const sectionStart = (k: number) => (k === 0 ? 0 : breakIndices[k - 1] + 1);
  const markOnly = (k: number) =>
    breakIndices[k] - sectionStart(k) === 1 && paintsNothing(blocks[breakIndices[k] - 1]);
  const opens = new Array<boolean>(breakIndices.length).fill(false);
  for (let k = breakIndices.length - 1; k >= 0; k -= 1) {
    const type = types[k + 1] ?? types[k] ?? 'nextPage';
    opens[k] =
      type !== 'continuous' ||
      sizeChanges(configs[k], configs[k + 1] ?? configs[k]) ||
      (k + 1 < breakIndices.length && markOnly(k + 1) && opens[k + 1]);
  }
  const breaks = new Set<number>();
  const marks = new Set<number>();
  breakIndices.forEach((sectionBreak, k) => {
    const mark = sectionBreak - 1;
    const pageBreak = sectionBreak - 2;
    if (!opens[k] || pageBreak < sectionStart(k) || !paintsNothing(blocks[mark])) return;
    const before = blocks[pageBreak];
    if (before.kind === 'pageBreak' || (paintsNothing(before) && before.attrs?.pageBreakBefore)) {
      breaks.add(pageBreak);
      marks.add(mark);
    }
  });
  return { breaks, marks };
}

/**
 * Place a joined empty mark on the current sheet without advancing the flow: it keeps a
 * fragment for the caret and selection, inside the content band, and moves nothing.
 */
export function placeJoinedMark(
  block: ParagraphBlock,
  measure: ParagraphMeasure,
  paginator: Paginator
): void {
  const state = paginator.getCurrentState();
  const height = measure.lines[0]?.lineHeight ?? 0;
  const range = getParagraphFragmentPmRange(block, measure, 0, measure.lines.length);
  const fragment: ParagraphFragment = {
    kind: 'paragraph',
    blockId: block.id,
    x: paginator.getColumnX(state.columnIndex),
    y: Math.max(state.topMargin, Math.min(state.cursorY, state.contentBottom - height)),
    width: paginator.getContentWidth(),
    height,
    fromLine: 0,
    toLine: measure.lines.length,
    pmStart: range.pmStart,
    pmEnd: range.pmEnd,
  };
  state.page.fragments.push(fragment);
}
