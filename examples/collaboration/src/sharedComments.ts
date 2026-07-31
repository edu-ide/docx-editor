import type * as Y from 'yjs';
import type { Comment } from '@eigenpal/docx-editor-core/types/content';

function commentKey(comment: Comment): string {
  return comment.parentId == null ? `root:${comment.id}` : `reply:${comment.parentId}:${comment.id}`;
}

export function commentsFromMap(comments: Y.Map<Comment>): Comment[] {
  return Array.from(comments.values()).sort((a, b) => {
    if (a.id !== b.id) return a.id - b.id;
    return (a.parentId ?? -1) - (b.parentId ?? -1);
  });
}

export function applyCommentSnapshotToMap(
  ydoc: Y.Doc,
  comments: Y.Map<Comment>,
  previous: Comment[],
  next: Comment[]
): void {
  const nextKeys = new Set(next.map(commentKey));

  ydoc.transact(() => {
    for (const oldComment of previous) {
      const key = commentKey(oldComment);
      if (!nextKeys.has(key)) {
        comments.delete(key);
      }
    }

    for (const nextComment of next) {
      comments.set(commentKey(nextComment), nextComment);
    }
  });
}
