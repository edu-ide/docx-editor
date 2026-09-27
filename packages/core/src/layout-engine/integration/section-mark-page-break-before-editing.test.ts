/**
 * An empty section mark whose style breaks the page before it (upstream 9afb832b, #983):
 * kept after its section's content, still a place to type, and derived again after every
 * edit, undo, and save + reopen.
 */
import { describe, expect, test } from 'bun:test';
import { EditorState, type Transaction } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';

import { serializeDocument } from '../../docx/serializer/documentSerializer';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import type { Document } from '../../types/document';
import { body, layPmDoc, mark, paragraph, read, sect } from './section-mark-helpers';

const bodyWithMark = (markContent = '') =>
  body(paragraph('First') + mark(sect(), undefined, markContent) + paragraph('Second') + sect());

interface Editor {
  readonly document: Document;
  readonly state: EditorState;
  dispatch(tr: Transaction): void;
}

function open(documentXml: string): Editor {
  const { document, pmDoc } = read(documentXml);
  let state = EditorState.create({ doc: pmDoc, plugins: [history()] });
  return {
    document,
    get state() {
      return state;
    },
    dispatch(tr) {
      state = state.apply(tr);
    },
  };
}

const lay = (editor: Editor) => layPmDoc(editor.state.doc, editor.document.package.document);

/** The sheet the mark, the second paragraph, lands on. */
function markPage(editor: Editor): number {
  const { layout, blocks } = lay(editor);
  const markBlock = blocks.filter((block) => block.kind === 'paragraph')[1]!;
  return layout.pages.findIndex((sheet) =>
    sheet.fragments.some((fragment) => fragment.blockId === markBlock.id)
  );
}

/** Every sheet as its fragments' boxes and words: comparable across sessions. */
function geometry(editor: Editor) {
  const { layout, blocks } = lay(editor);
  return layout.pages.map((sheet) =>
    sheet.fragments.map((fragment) => {
      const block = blocks.find((candidate) => candidate.id === fragment.blockId);
      const words =
        block?.kind === 'paragraph'
          ? block.runs.map((run) => (run.kind === 'text' ? run.text : '')).join('')
          : block?.kind;
      return { box: [fragment.x, fragment.y, fragment.width, fragment.height], words };
    })
  );
}

/** The position inside the mark paragraph. */
const inMark = (editor: Editor) => editor.state.doc.child(0).nodeSize + 1;

const save = (editor: Editor) => serializeDocument(fromProseDoc(editor.state.doc, editor.document));

describe('an empty section mark with a page break before it, while editing', () => {
  test('stays after the section content', () => {
    const editor = open(bodyWithMark());
    expect(lay(editor).layout.pages).toHaveLength(2);
    expect(markPage(editor)).toBe(0);
  });

  test('typing into the mark applies its page break, and undo removes that sheet', () => {
    const editor = open(bodyWithMark());
    editor.dispatch(editor.state.tr.insertText('x', inMark(editor)));
    expect(lay(editor).layout.pages).toHaveLength(3);
    expect(markPage(editor)).toBe(1);
    undo(editor.state, editor.dispatch);
    expect(lay(editor).layout.pages).toHaveLength(2);
    expect(markPage(editor)).toBe(0);
  });

  test('typing into the mark and deleting it again leaves no sheet, before or after a save', () => {
    const inert =
      '<w:proofErr w:type="gramStart"/><w:r><w:t></w:t></w:r><w:proofErr w:type="gramEnd"/>';
    const editor = open(bodyWithMark(inert));
    expect(lay(editor).layout.pages).toHaveLength(2);
    editor.dispatch(editor.state.tr.insertText('x', inMark(editor)));
    expect(lay(editor).layout.pages).toHaveLength(3);
    editor.dispatch(editor.state.tr.delete(inMark(editor), inMark(editor) + 1));
    expect(lay(editor).layout.pages).toHaveLength(2);
    expect(markPage(editor)).toBe(0);
    const reopened = open(save(editor));
    expect(lay(reopened).layout.pages).toHaveLength(2);
    expect(markPage(reopened)).toBe(0);
  });

  test('edited layout matches a reopened copy of the saved document', () => {
    const editor = open(bodyWithMark());
    editor.dispatch(editor.state.tr.insertText('x', inMark(editor)));
    undo(editor.state, editor.dispatch);
    const xml = save(editor);
    // The mark paragraph keeps its style, and the style's page break stays the style's.
    expect(xml.match(/<w:p[ >]/g)).toHaveLength(3);
    expect(xml.match(/<w:sectPr[ >]/g)).toHaveLength(2);
    expect(xml).toContain('<w:pStyle w:val="Break"/>');
    expect(xml).not.toContain('<w:pageBreakBefore');
    const reopened = open(xml);
    expect(geometry(reopened)).toEqual(geometry(editor));
    expect(markPage(reopened)).toBe(0);
  });
});
