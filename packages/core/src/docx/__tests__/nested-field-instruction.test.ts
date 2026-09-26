/**
 * A field nested inside another field's instruction is input to that field and
 * is never displayed (upstream 390c1772, #986).
 *
 * `IF { STYLEREF Heading } <> "x" ...` compares a nested field's result; only
 * the outer field's cached result displays. The paragraph parser used to pull a
 * nested `w:fldSimple` out as a displayed field of its own, and a nested
 * `begin` reset the outer field. Nested fields now stay in the outer field
 * code, and save writes them back inside the instruction.
 */
import { describe, expect, test } from 'bun:test';
import { EditorState } from 'prosemirror-state';
import { parseDocumentBody } from '../documentParser';
import { serializeDocument } from '../serializer/documentSerializer';
import { toProseDoc } from '../../prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../../prosemirror/conversion/fromProseDoc';
import { toFlowBlocks } from '../../layout-bridge/toFlowBlocks';
import type { ComplexField, Document } from '../../types/document';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const BEGIN = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
const SEPARATE = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const instr = (value: string) =>
  `<w:r><w:instrText xml:space="preserve">${value}</w:instrText></w:r>`;
const run = (content: string) => `<w:r>${content}</w:r>`;
const text = (value: string) => run(`<w:t xml:space="preserve">${value}</w:t>`);
const simple = (instruction: string, result: string) =>
  `<w:fldSimple w:instr="${instruction}">${result}</w:fldSimple>`;
const styleRef = (result: string) => simple(' STYLEREF Heading ', result);
const complexStyleRef = (result: string) =>
  BEGIN + instr(' STYLEREF Heading ') + SEPARATE + result + END;
/** `IF {first} <> "x" {second}` with a cached outer result. */
const conditional = (first: string, second: string, result: string) =>
  BEGIN +
  instr(' IF ') +
  first +
  instr(' &lt;&gt; "x" ') +
  second +
  instr(' ') +
  SEPARATE +
  result +
  END;

function documentOf(paragraph: string): Document {
  const body = parseDocumentBody(
    `<w:document xmlns:w="${W}"><w:body><w:p>${paragraph}</w:p></w:body></w:document>`
  );
  return { package: { document: body } };
}

function contentOf(doc: Document) {
  const block = doc.package.document.content[0];
  if (block?.type !== 'paragraph') throw new Error('no paragraph');
  return block.content;
}

/** The text the paragraph displays, fields shown by their cached result. */
function displayed(doc: Document): string {
  const block = toFlowBlocks(toProseDoc(doc), {}).find((b) => b.kind === 'paragraph');
  if (block?.kind !== 'paragraph') throw new Error('no paragraph block');
  return block.runs
    .map((r) => (r.kind === 'text' ? r.text : r.kind === 'field' ? (r.fallback ?? '') : ''))
    .join('');
}

/** document.xml after the editor round trip (parse -> PM -> model -> XML). */
function savedThroughEditor(doc: Document): string {
  return serializeDocument(fromProseDoc(toProseDoc(doc), doc));
}

describe('fields nested inside a complex field instruction', () => {
  test('nested simple fields do not display; the outer cached result does', () => {
    const doc = documentOf(
      text('A') +
        conditional(styleRef(run('<w:cr/>')), styleRef(run('<w:cr/>')), text('R')) +
        text('Z')
    );
    expect(contentOf(doc).map((item) => item.type)).toEqual(['run', 'complexField', 'run']);
    expect(displayed(doc)).toBe('ARZ');
  });

  test('text, tab, and break results stay out', () => {
    const rich = run('<w:t>nested</w:t><w:tab/><w:br/><w:cr/>');
    expect(displayed(documentOf(conditional(styleRef(rich), styleRef(rich), text('shown'))))).toBe(
      'shown'
    );
  });

  test('empty nested caches and an empty outer cache display nothing', () => {
    expect(
      displayed(documentOf(text('A') + conditional(styleRef(''), styleRef(''), '') + text('Z')))
    ).toBe('AZ');
  });

  test('a nested complex field keeps the outer field whole', () => {
    const doc = documentOf(
      text('A') + conditional(complexStyleRef(text('nested')), '', text('R')) + text('Z')
    );
    const items = contentOf(doc);
    expect(items.map((item) => item.type)).toEqual(['run', 'complexField', 'run']);
    const field = items[1] as ComplexField;
    expect(field.fieldType).toBe('IF');
    expect(displayed(doc)).toBe('ARZ');
  });

  test('a nested field never becomes part of the outer instruction', () => {
    const doc = documentOf(
      BEGIN +
        instr(' HYPERLINK "') +
        complexStyleRef(text('https://nested.example/')) +
        styleRef(text('also-nested')) +
        instr('" ') +
        SEPARATE +
        text('link') +
        END
    );
    const field = contentOf(doc)[0] as ComplexField;
    expect(field.instruction).toBe('HYPERLINK ""');
    expect(field.instruction).not.toContain('nested');
  });

  test('fields inside a saved result still display', () => {
    const doc = documentOf(
      text('A') +
        BEGIN +
        instr(' IF 1 = 1 "x" ') +
        SEPARATE +
        styleRef(text('inner')) +
        END +
        text('Z')
    );
    expect(displayed(doc)).toContain('inner');
  });
});

describe('saving a field with nested instruction fields', () => {
  const paragraph =
    text('A') +
    conditional(styleRef(text('n1')), complexStyleRef(text('n2')), text('R')) +
    text('Z');

  test('the editor round trip writes the nested fields back inside the instruction', () => {
    const xml = savedThroughEditor(documentOf(paragraph));
    const body = xml.slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'));
    // Outer begin, IF, a nested field (from the fldSimple), <> "x", a nested field, separate.
    const order = [...body.matchAll(/fldCharType="(\w+)"|<w:instrText[^>]*>([^<]*)</g)].map(
      (match) => match[1] ?? match[2]!.trim()
    );
    expect(order).toEqual([
      'begin',
      'IF',
      'begin',
      'STYLEREF Heading',
      'separate',
      'end',
      '&lt;&gt; &quot;x&quot;',
      'begin',
      'STYLEREF Heading',
      'separate',
      'end',
      '',
      'separate',
      'end',
    ]);
    // It reopens as the same paragraph.
    const reopened = parseDocumentBody(xml);
    const block = reopened.content[0];
    expect(block?.type === 'paragraph' && block.content.map((item) => item.type)).toEqual([
      'run',
      'complexField',
      'run',
    ]);
  });

  test('the model serializer keeps the nested fields in place too', () => {
    const xml = serializeDocument(documentOf(paragraph));
    const ifAt = xml.indexOf('> IF <');
    expect(ifAt).toBeGreaterThan(-1);
    expect(xml.indexOf('n1')).toBeGreaterThan(ifAt);
    expect(xml.indexOf('n2')).toBeLessThan(xml.lastIndexOf('fldCharType="separate"'));
    expect(xml.lastIndexOf('fldCharType="separate"')).toBeLessThan(xml.indexOf('>R<'));
  });

  test('an ordinary complex field carries no stored field code', () => {
    const pm = toProseDoc(documentOf(BEGIN + instr(' PAGE ') + SEPARATE + text('1') + END));
    let fieldCode: unknown = 'unset';
    pm.descendants((node) => {
      if (node.type.name === 'field') fieldCode = node.attrs.fieldCode;
    });
    expect(fieldCode).toBeNull();
  });

  test('stored field code is dropped once the instruction changes', () => {
    const doc = documentOf(paragraph);
    const state = EditorState.create({ doc: toProseDoc(doc) });
    let fieldPos = -1;
    state.doc.descendants((node, pos) => {
      if (node.type.name === 'field') fieldPos = pos;
    });
    const field = state.doc.nodeAt(fieldPos)!;
    const tr = state.tr.setNodeMarkup(fieldPos, undefined, { ...field.attrs, instruction: 'PAGE' });
    const xml = serializeDocument(fromProseDoc(tr.doc, doc));
    expect(xml).not.toContain('STYLEREF');
    expect(xml).toContain('>PAGE<');
  });
});
