/**
 * Comment + tracked-change ID allocation.
 *
 * Comments (`w:comment` ids) and tracked changes (`w:ins`/`w:del` revision ids)
 * share a single OOXML ID space — a duplicate ID between the two corrupts the
 * round-trip. Allocation is therefore one monotonic, no-reuse counter, exposed
 * as an **instance-scoped** factory rather than module-global state so two
 * editor instances on one page never share (or collide on) a counter.
 *
 * Kept separate from the comment/tracked-change transaction builders
 * (`commentOps.ts`) so the allocator can be owned independently — the editor
 * engine seeds and threads it without dragging in the PM-text-lookup graph.
 */

import type { EditorView } from 'prosemirror-view';
import type { Comment } from '../types/content';

/** Sentinel ID for a comment that hasn't been persisted yet (anchored to selection). */
export const PENDING_COMMENT_ID = -1;

export interface CommentIdAllocator {
  /** Allocate the next ID and advance the counter. */
  next(): number;
  /**
   * On document load, bump the counter above the highest ID found in the
   * loaded comments and tracked-change marks so subsequent allocations don't
   * collide with already-present IDs.
   */
  seedAbove(maxId: number): void;
}

export interface CommentIdAllocatorOptions {
  /**
   * First ID to allocate inside a sharded numeric space. Defaults to 1, which
   * preserves the classic 1,2,3... allocator.
   */
  shardOffset?: number;
  /**
   * Distance between IDs allocated by this instance. Defaults to 1. Realtime
   * collaboration adapters can give each peer a different offset with the same
   * stride so simultaneous local allocations do not produce the same number.
   */
  shardStride?: number;
}

function positiveInteger(value: number | undefined, fallback: number): number {
  if (value == null || !Number.isFinite(value)) return fallback;
  return Math.max(1, Math.floor(value));
}

/**
 * Create an instance-scoped monotonic comment/revision ID allocator. IDs are
 * never reused (deleting a comment does not free its ID), and the counter is
 * private to this allocator — multiple editors get independent ID spaces.
 */
export function createCommentIdAllocator(
  options: CommentIdAllocatorOptions = {}
): CommentIdAllocator {
  const stride = positiveInteger(options.shardStride, 1);
  const offset = positiveInteger(options.shardOffset, 1);
  let nextId = offset;
  return {
    next: () => {
      const id = nextId;
      nextId += stride;
      return id;
    },
    seedAbove(maxId: number) {
      if (maxId >= nextId) {
        nextId += (Math.floor((maxId - nextId) / stride) + 1) * stride;
      }
    },
  };
}

/**
 * Seed an allocator above every comment/revision ID currently in the document
 * — comment objects (including replies, which carry no mark) plus
 * tracked-change `revisionId` marks. Because `seedAbove` only ever raises the
 * counter, this is safe to call on load (React) or before each allocation
 * (Vue): new IDs never collide with or reuse an existing one, and the comment
 * and revision ID spaces stay unified.
 */
export function seedCommentAllocator(
  allocator: CommentIdAllocator,
  comments: Comment[] | undefined,
  view: EditorView | null
): void {
  let max = 0;
  for (const comment of comments ?? []) max = Math.max(max, comment.id);
  if (view) {
    view.state.doc.descendants((node) => {
      for (const mark of node.marks) {
        if (mark.attrs.revisionId != null) max = Math.max(max, mark.attrs.revisionId as number);
      }
    });
  }
  allocator.seedAbove(max);
}
