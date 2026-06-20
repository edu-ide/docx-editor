import { describe, expect, test } from 'bun:test';
import { assertNoStructuralLoss, countStructuralTags, diffStructuralTags } from './structureGate';

const wrap = (body: string) =>
  '<?xml version="1.0"?>' +
  '<w:document xmlns:w="urn:w" xmlns:mc="urn:mc" xmlns:a="urn:a" xmlns:r="urn:r">' +
  `<w:body>${body}</w:body></w:document>`;

describe('countStructuralTags', () => {
  test('counts w:p without over-counting w:pPr / w:pict', () => {
    const c = countStructuralTags(wrap('<w:p><w:pPr/></w:p><w:p/><w:pict/>'));
    expect(c['w:p']).toBe(2);
    expect(c['w:pict']).toBe(1);
  });

  test('counts self-closing and open structural tags, not closing tags', () => {
    const c = countStructuralTags(wrap('<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><a:blip r:embed="rId5"/>'));
    expect(c['w:tbl']).toBe(1);
    expect(c['w:tr']).toBe(1);
    expect(c['w:tc']).toBe(1);
    expect(c['a:blip']).toBe(1);
    expect(c['w:p']).toBe(1); // the </w:p>, </w:tc>... closers must not add to the count
  });
});

describe('assertNoStructuralLoss', () => {
  test('passes when fragile passthrough content survives an edit', () => {
    const before = wrap('<w:p/><w:txbxContent><w:p/></w:txbxContent>');
    const after = wrap('<w:p/><w:p/><w:txbxContent><w:p/></w:txbxContent>'); // text edited, text box kept
    expect(() => assertNoStructuralLoss(before, after)).not.toThrow();
  });

  test('throws when text box / pict / AlternateContent are silently dropped (the eigenpal -58% bug)', () => {
    const before = wrap(
      '<w:p/><mc:AlternateContent><w:pict/></mc:AlternateContent><w:txbxContent><w:p/></w:txbxContent>'
    );
    const after = wrap('<w:p/>'); // unmodeled content vanished on repack
    expect(() => assertNoStructuralLoss(before, after)).toThrow(/silently lost/);
  });

  test('does NOT throw when only editable paragraphs are deleted (no false positive)', () => {
    const before = wrap('<w:p/><w:p/><w:p/>');
    const after = wrap('<w:p/>'); // user intentionally deleted two paragraphs
    expect(() => assertNoStructuralLoss(before, after)).not.toThrow();
  });

  test('strict mode also blocks editable-structure loss', () => {
    const before = wrap('<w:p/><w:p/>');
    const after = wrap('<w:p/>');
    expect(() => assertNoStructuralLoss(before, after, { strict: true })).toThrow();
  });
});

describe('diffStructuralTags', () => {
  test('reports per-tag deltas; editable-tag loss is not fragile', () => {
    const before = wrap('<w:p/><w:tbl><w:tr><w:tc/></w:tr></w:tbl>');
    const after = wrap('<w:p/><w:p/>'); // gained a paragraph, lost a table
    const diff = diffStructuralTags(before, after);
    expect(diff.tags['w:p'].delta).toBe(1);
    expect(diff.tags['w:tbl'].delta).toBe(-1);
    expect(diff.lost).toContain('w:tbl');
    expect(diff.fragileLost).not.toContain('w:tbl');
    expect(diff.ok).toBe(true); // a table is editable, so its loss does not fail the gate
  });
});
