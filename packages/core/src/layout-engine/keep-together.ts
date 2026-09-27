/**
 * Keep Together Logic - Handle keepNext and keepLines paragraph properties
 *
 * DOCX paragraphs can have keepNext (keep with next paragraph) and keepLines
 * (keep all lines together) properties that affect pagination.
 */

import type { FlowBlock, ParagraphBlock, Measure, ParagraphMeasure } from './types';

/**
 * A chain of consecutive keepNext paragraphs.
 */
export type KeepNextChain = {
  /** Index of the first paragraph in the chain. */
  startIndex: number;
  /** Index of the last paragraph in the chain. */
  endIndex: number;
  /** All paragraph indices in the chain. */
  memberIndices: number[];
  /** Index of the anchor paragraph (first non-keepNext after chain), or -1 if none. */
  anchorIndex: number;
  /**
   * The chain ends at a forced break — a page or column break, or a paragraph with
   * `pageBreakBefore` — so nothing after it shares the page, and the last member's
   * trailing spacing is discarded there (upstream ab460dc9, #985).
   */
  endsAtBreak?: boolean;
};

/**
 * Pre-scan blocks to find all keepNext chains.
 *
 * A keepNext chain is a sequence of consecutive paragraphs with keepNext=true,
 * followed by an anchor paragraph (the first non-keepNext paragraph).
 * The entire chain must stay on the same page as the anchor's first line.
 *
 * `breaksBefore` says whether the block at an index starts a new page for its
 * own `pageBreakBefore`. The layout passes its own answer, under which an empty
 * section mark after its section's content does not (`pageBreaksBefore`).
 *
 * Returns a map from chain start index to chain info.
 */
export function computeKeepNextChains(
  blocks: FlowBlock[],
  breaksBefore: (index: number) => boolean = (index) => hasPageBreakBefore(blocks[index])
): Map<number, KeepNextChain> {
  const chains = new Map<number, KeepNextChain>();
  const processed = new Set<number>();

  for (let i = 0; i < blocks.length; i++) {
    // Skip already-processed blocks (mid-chain members)
    if (processed.has(i)) continue;

    const block = blocks[i];
    // Only paragraphs can have keepNext
    if (block.kind !== 'paragraph') continue;

    const para = block as ParagraphBlock;
    // Skip paragraphs without keepNext
    if (!para.attrs?.keepNext) continue;

    // Found a keepNext paragraph - scan forward to find full chain
    const memberIndices: number[] = [i];
    let endIndex = i;
    let endsAtBreak = false;

    for (let j = i + 1; j < blocks.length; j++) {
      const nextBlock = blocks[j];

      // A forced break ends the keep group: what follows starts on another page or
      // column, so a member or anchor past it would strand the chain on a page alone.
      if (nextBlock.kind === 'pageBreak' || nextBlock.kind === 'columnBreak' || breaksBefore(j)) {
        endsAtBreak = true;
        break;
      }

      // Section breaks terminate the chain
      if (nextBlock.kind === 'sectionBreak') {
        break;
      }

      // Non-paragraphs terminate the chain
      if (nextBlock.kind !== 'paragraph') {
        break;
      }

      const nextPara = nextBlock as ParagraphBlock;
      if (nextPara.attrs?.keepNext) {
        // Continue the chain
        memberIndices.push(j);
        endIndex = j;
        processed.add(j);
      } else {
        // Found the anchor - stop here
        break;
      }
    }

    // Find the anchor (first paragraph after the chain)
    const potentialAnchor = endIndex + 1;
    let anchorIndex = -1;

    if (!endsAtBreak && potentialAnchor < blocks.length) {
      const anchorBlock = blocks[potentialAnchor];
      // Anchor must not be a break
      if (
        anchorBlock.kind !== 'sectionBreak' &&
        anchorBlock.kind !== 'pageBreak' &&
        anchorBlock.kind !== 'columnBreak'
      ) {
        anchorIndex = potentialAnchor;
      }
    }

    // Record the chain
    chains.set(i, {
      startIndex: i,
      endIndex,
      memberIndices,
      anchorIndex,
      ...(endsAtBreak ? { endsAtBreak } : {}),
    });
  }

  return chains;
}

/**
 * Calculate the total height needed to keep a chain together.
 *
 * Includes all chain members plus the first line of the anchor paragraph.
 */
export function calculateChainHeight(
  chain: KeepNextChain,
  blocks: FlowBlock[],
  measures: Measure[]
): number {
  let totalHeight = 0;

  // Sum heights of all chain members
  for (const memberIndex of chain.memberIndices) {
    const block = blocks[memberIndex];
    const measure = measures[memberIndex];

    if (block.kind !== 'paragraph' || measure.kind !== 'paragraph') continue;

    const para = block as ParagraphBlock;
    const paraMeasure = measure as ParagraphMeasure;

    // Add spacing before (simplified - could be more sophisticated with collapse)
    const spacingBefore = para.attrs?.spacing?.before ?? 0;
    totalHeight += spacingBefore;

    // Add paragraph height
    totalHeight += paraMeasure.totalHeight;

    // Add spacing after. A forced break after the chain discards the last member's
    // trailing spacing, including the part `totalHeight` carries.
    const spacingAfter = para.attrs?.spacing?.after ?? 0;
    if (chain.endsAtBreak && memberIndex === chain.endIndex) totalHeight -= spacingAfter;
    else totalHeight += spacingAfter;
  }

  // Add first line height of anchor (if any)
  if (chain.anchorIndex !== -1) {
    const anchorMeasure = measures[chain.anchorIndex];
    if (anchorMeasure?.kind === 'paragraph') {
      const anchorPara = anchorMeasure as ParagraphMeasure;
      if (anchorPara.lines.length > 0) {
        // Add just the first line height
        totalHeight += anchorPara.lines[0].lineHeight;
      }
    }
  }

  return totalHeight;
}

/**
 * Get the set of indices that are mid-chain (not chain starters).
 * These should skip the keepNext check since their chain starter already decided.
 */
export function getMidChainIndices(chains: Map<number, KeepNextChain>): Set<number> {
  const midChain = new Set<number>();

  for (const chain of chains.values()) {
    // All members except the first are mid-chain
    for (let i = 1; i < chain.memberIndices.length; i++) {
      midChain.add(chain.memberIndices[i]);
    }
  }

  return midChain;
}

/**
 * Check if a paragraph has keepLines property (all lines must stay together).
 */
export function hasKeepLines(block: FlowBlock): boolean {
  if (block.kind !== 'paragraph') return false;
  const para = block as ParagraphBlock;
  return para.attrs?.keepLines === true;
}

/**
 * Check if a paragraph should start on a new page (pageBreakBefore).
 */
export function hasPageBreakBefore(block: FlowBlock): boolean {
  if (block.kind !== 'paragraph') return false;
  const para = block as ParagraphBlock;
  return para.attrs?.pageBreakBefore === true;
}
