import { describe, expect, test } from 'bun:test';
import type { Document, Paragraph } from '../../../types/document';
import { serializeParagraph } from '../../../docx/serializer/paragraphSerializer';
import { fromProseDoc } from '../fromProseDoc';
import { toProseDoc } from '../toProseDoc';

function docOf(...content: Paragraph[]): Document {
  return { package: { document: { content } } };
}

function textParagraph(text: string): Paragraph {
  return {
    type: 'paragraph',
    content: [{ type: 'run', content: [{ type: 'text', text }] }],
  };
}

function childTypes(doc: ReturnType<typeof toProseDoc>): string[] {
  const types: string[] = [];
  doc.forEach((node) => {
    types.push(node.type.name);
  });
  return types;
}

describe('issue #830 leading hard page break round-trip', () => {
  test('does not export a leading hard page break as an extra empty paragraph', () => {
    const leadingBreakParagraph: Paragraph = {
      type: 'paragraph',
      renderedPageBreakBefore: true,
      formatting: { styleId: 'CenterSingle' },
      content: [
        { type: 'run', content: [{ type: 'break', breakType: 'page' }] },
        { type: 'run', content: [{ type: 'text', text: 'After hard break' }] },
      ],
    };
    const input = docOf(textParagraph('Before'), leadingBreakParagraph);

    const pmDoc = toProseDoc(input);
    expect(childTypes(pmDoc)).toEqual(['paragraph', 'paragraph']);
    expect(pmDoc.child(1).attrs.pageBreakBefore).toBe(true);
    expect(pmDoc.child(1).attrs.pageBreakBeforeSource).toBe('leadingBreak');
    expect(pmDoc.child(1).attrs.renderedPageBreakBefore).toBe(true);

    const roundTripped = fromProseDoc(pmDoc, input);
    expect(roundTripped.package.document.content).toHaveLength(2);

    const outputParagraph = roundTripped.package.document.content[1];
    expect(outputParagraph?.type).toBe('paragraph');
    if (outputParagraph?.type !== 'paragraph') {
      throw new Error('Expected second body block to remain a paragraph');
    }

    // The break goes back as the run it was, not as a `w:pageBreakBefore`
    // property, and the lastRenderedPageBreak marker stays on the line after it.
    const xml = serializeParagraph(outputParagraph);
    expect(xml).not.toContain('<w:pageBreakBefore');
    expect(xml).toContain(
      '<w:r><w:br w:type="page"/></w:r><w:r><w:lastRenderedPageBreak/><w:t>After hard break</w:t></w:r>'
    );
  });

  test('preserves a break-only paragraph with no direct formatting', () => {
    const leadingBreakParagraph: Paragraph = {
      type: 'paragraph',
      content: [{ type: 'run', content: [{ type: 'break', breakType: 'page' }] }],
    };

    const pmDoc = toProseDoc(docOf(leadingBreakParagraph));
    expect(childTypes(pmDoc)).toEqual(['paragraph']);
    expect(pmDoc.child(0).attrs.pageBreakBefore).toBe(true);

    const roundTripped = fromProseDoc(pmDoc, docOf(leadingBreakParagraph));
    const outputParagraph = roundTripped.package.document.content[0];
    expect(outputParagraph?.type).toBe('paragraph');
    if (outputParagraph?.type !== 'paragraph') {
      throw new Error('Expected body block to remain a paragraph');
    }

    expect(serializeParagraph(outputParagraph)).toBe('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
  });

  test('keeps non-leading hard page breaks as explicit PM page break blocks', () => {
    const midParagraph: Paragraph = {
      type: 'paragraph',
      content: [
        { type: 'run', content: [{ type: 'text', text: 'Before break' }] },
        { type: 'run', content: [{ type: 'break', breakType: 'page' }] },
        { type: 'run', content: [{ type: 'text', text: 'After break' }] },
      ],
    };

    const pmDoc = toProseDoc(docOf(midParagraph));
    expect(childTypes(pmDoc)).toEqual(['paragraph', 'pageBreak']);
  });
});
