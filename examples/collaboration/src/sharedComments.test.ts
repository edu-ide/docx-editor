import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { Comment } from '@eigenpal/docx-editor-core/types/content';
import { applyCommentSnapshotToMap, commentsFromMap } from './sharedComments';

function comment(id: number, author: string, parentId?: number): Comment {
  return {
    id,
    author,
    parentId,
    content: [],
  };
}

describe('shared comment map merge', () => {
  test('preserves remote comments that were not in the local snapshot', () => {
    const ydoc = new Y.Doc();
    const yComments = ydoc.getMap<Comment>('comments');

    const localPrevious = [comment(1, 'Ada')];
    applyCommentSnapshotToMap(ydoc, yComments, [], localPrevious);

    ydoc.transact(() => {
      yComments.set('root:2', comment(2, 'Grace'));
    });

    applyCommentSnapshotToMap(ydoc, yComments, localPrevious, [
      comment(1, 'Ada Lovelace'),
    ]);

    expect(commentsFromMap(yComments)).toEqual([
      comment(1, 'Ada Lovelace'),
      comment(2, 'Grace'),
    ]);
  });

  test('deletes comments only when they existed in the local previous snapshot', () => {
    const ydoc = new Y.Doc();
    const yComments = ydoc.getMap<Comment>('comments');
    const localPrevious = [comment(1, 'Ada'), comment(3, 'Linus', 1)];

    applyCommentSnapshotToMap(ydoc, yComments, [], localPrevious);
    applyCommentSnapshotToMap(ydoc, yComments, localPrevious, [comment(1, 'Ada')]);

    expect(commentsFromMap(yComments)).toEqual([comment(1, 'Ada')]);
  });
});
